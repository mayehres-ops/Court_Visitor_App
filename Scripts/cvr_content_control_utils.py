"""
Utility functions for filling Word document content controls.
Used by both CVR generation (Step 8) and CVR completion (Step 10).

This module provides a standardized way to fill content controls that handles:
- Text content controls (wdContentControlRichText = 1, wdContentControlText = 2)
- Checkbox content controls (wdContentControlCheckBox = 3)
- Date content controls (wdContentControlDate = 5)
- Dropdown content controls (wdContentControlDropdownList = 6, wdContentControlComboBox = 7)
"""

from datetime import datetime
import pandas as pd
import re


def format_date_value(value):
    """
    Format a value as MM/DD/YYYY if it's a date.
    Returns the formatted date string or the original value if not a date.
    """
    # If it's already a pandas Timestamp or datetime
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.strftime("%m/%d/%Y")

    # If it's a string that looks like a date
    if isinstance(value, str):
        # Try common date formats
        date_patterns = [
            r'^\d{4}-\d{2}-\d{2}',  # YYYY-MM-DD
            r'^\d{1,2}/\d{1,2}/\d{4}',  # MM/DD/YYYY or M/D/YYYY
            r'^\d{1,2}-\d{1,2}-\d{4}',  # MM-DD-YYYY
        ]

        for pattern in date_patterns:
            if re.match(pattern, value.strip()):
                try:
                    # Try to parse with pandas
                    parsed_date = pd.to_datetime(value)
                    return parsed_date.strftime("%m/%d/%Y")
                except:
                    pass

    # Return original value if not a date
    return value


def clean_placeholder_text(value):
    """
    Remove Word template placeholder text from values.
    Filters out common placeholder phrases that might have been accidentally
    copied from Word templates into the Excel data.

    Args:
        value: String value to clean

    Returns:
        Cleaned value, or empty string if the value was just placeholder text
    """
    if not isinstance(value, str):
        return value

    # Common Word placeholder patterns
    placeholders = [
        r'Click or tap here to enter text\.?',
        r'Click here to enter text\.?',
        r'Type here\.?',
        r'Enter text here\.?',
        r'\[.*?Type here.*?\]',
        r'\[.*?Click.*?here.*?\]',
    ]

    cleaned = value
    for pattern in placeholders:
        cleaned = re.sub(pattern, '', cleaned, flags=re.IGNORECASE)

    # Clean up extra spaces
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()

    # If cleaning removed everything, return empty string
    return cleaned if cleaned else ''


def fill_content_control(cc, value, verbose=False):
    """
    Fill a single Word content control with the given value.
    FIXED: Unlock controls before filling, handle XML mapping, correct checkbox type.

    Args:
        cc: Word ContentControl COM object
        value: String value to fill (or 'yes'/'no' for checkboxes)
        verbose: If True, print detailed debugging info

    Returns:
        True if successfully filled, False otherwise
    """
    try:
        # Clean placeholder text from value first
        value = clean_placeholder_text(value)

        control_type = cc.Type
        name = (cc.Title or cc.Tag or "unnamed").strip()

        if verbose:
            type_names = {
                1: "RichText",
                2: "Text",
                3: "BuildingBlockGallery",
                4: "ComboBox",
                5: "DropdownList",
                6: "Picture",
                7: "Date",
                8: "Checkbox"  # CORRECT: Type 8 = wdContentControlCheckBox
            }
            print(f"    Filling '{name}' (Type: {type_names.get(control_type, control_type)}) with: {value}")

        # UNLOCK the content control before editing (critical fix!)
        try:
            cc.LockContentControl = False
            if hasattr(cc, 'LockContents'):
                cc.LockContents = False
        except:
            pass  # Some controls don't have these properties

        # Try XML mapping first (if mapped)
        try:
            if hasattr(cc, 'XMLMapping') and cc.XMLMapping and cc.XMLMapping.IsMapped:
                cc.XMLMapping.SetString(str(value) if value else "")
                return True
        except:
            pass  # Fall through to standard methods

        # Checkbox content control (Type 8 = wdContentControlCheckBox)
        # Note: Type 3 is wdContentControlBuildingBlockGallery, Type 5 is wdContentControlDate
        if control_type == 8:  # wdContentControlCheckBox
            try:
                if isinstance(value, bool):
                    cc.Checked = value
                elif isinstance(value, str):
                    # Accept various checkbox values
                    cc.Checked = value.lower() in ['yes', 'y', 'true', '1', 'checked', 'on', 'x']
                else:
                    cc.Checked = bool(value)
                return True
            except:
                # If Checked doesn't work, try Range.Text with X
                try:
                    if value and str(value).lower() in ['yes', 'y', 'true', '1', 'x']:
                        cc.Range.Text = 'X'
                    else:
                        cc.Range.Text = ''
                    return True
                except:
                    pass

        # Date content control (Type 5)
        elif control_type == 5:
            # For dates, format as MM/DD/YYYY
            formatted_value = format_date_value(value)
            cc.Range.Text = str(formatted_value)
            return True

        # Dropdown/ComboBox (Types 4, 5)
        elif control_type in [4, 5]:
            # Try to select the matching dropdown item
            str_value = str(value).strip()
            try:
                for item in cc.DropdownListEntries:
                    if item.Text.strip().lower() == str_value.lower():
                        cc.DropdownListEntries.Item(item.Index).Select()
                        return True
            except:
                pass
            # If no match found, try to set text directly (ComboBox allows this)
            if control_type == 4:  # ComboBox
                cc.Range.Text = str_value
                return True
            return False

        # Text content controls (Types 1, 2) or any other type
        else:
            # Format dates even in text controls
            formatted_value = format_date_value(value)
            cc.Range.Text = str(formatted_value)
            return True

    except Exception as e:
        if verbose:
            print(f"    ERROR filling '{name}': {e}")
        return False


def fill_content_controls_from_dict(doc, data_dict, column_aliases=None, verbose=False):
    """
    Fill all content controls in a Word document from a dictionary.

    Args:
        doc: Word Document COM object
        data_dict: Dictionary mapping field names to values
        column_aliases: Optional dict mapping alternative names to canonical names
        verbose: If True, print detailed debugging info

    Returns:
        Number of content controls successfully filled
    """
    if column_aliases is None:
        column_aliases = {}

    filled_count = 0

    # Build a map of multi-select fields (for checkboxes like access_phone, access_tv, etc.)
    multi_select_map = {}
    for key, val in data_dict.items():
        # Check if value contains commas (multi-select answer)
        if ',' in str(val):
            # Split and normalize each option
            options = [opt.strip().lower() for opt in str(val).split(',')]
            multi_select_map[key.lower()] = options

    try:
        # First, build a list of all control names in the document for debugging
        all_control_names = set()
        for cc in doc.ContentControls:
            name = (cc.Title or cc.Tag or "").strip()
            if name:
                all_control_names.add(name.lower())

        if verbose:
            print(f"\n  DEBUG: Found {len(all_control_names)} controls in document")
            # Show ALL cond_ controls
            cond_controls = [n for n in sorted(all_control_names) if n.startswith('cond_')]
            if cond_controls:
                print(f"  DEBUG: ALL controls starting with 'cond_': {len(cond_controls)}")
                for c in cond_controls:
                    print(f"    - {c}")

            # Show text controls (non-checkbox, non-cond)
            text_controls = [n for n in sorted(all_control_names)
                           if not n.endswith('_yes') and not n.endswith('_no')
                           and not n.startswith('cond_') and not n.startswith('residence_')
                           and not n.startswith('access_') and not n.startswith('activities_')]
            if text_controls:
                print(f"\n  DEBUG: Possible text controls: {len(text_controls)}")
                for c in text_controls[:10]:
                    print(f"    - {c}")
                if len(text_controls) > 10:
                    print(f"    ... and {len(text_controls) - 10} more")

        for cc in doc.ContentControls:
            # Get the content control name (Title or Tag)
            name = (cc.Title or cc.Tag or "").strip()
            if not name:
                continue

            # Look up value by name (case-insensitive)
            value = None
            name_lower = name.lower()

            # Check if this is a yes/no checkbox (ends with _yes or _no)
            is_yesno_checkbox = name_lower.endswith('_yes') or name_lower.endswith('_no')

            if is_yesno_checkbox:
                # Extract base name (e.g., "ownbed" from "ownbed_yes")
                base_name = name_lower[:-4] if name_lower.endswith('_yes') else name_lower[:-3]
                is_yes = name_lower.endswith('_yes')

                # Look up the base field value
                for key, val in data_dict.items():
                    if key.lower() == base_name:
                        answer = str(val).strip().lower()
                        # Check if this checkbox should be marked
                        if is_yes and answer in ['yes', 'y', 'true', '1']:
                            value = 'X'  # or '✓' or whatever mark you want
                        elif not is_yes and answer in ['no', 'n', 'false', '0']:
                            value = 'X'
                        else:
                            value = ''  # Leave blank
                        break
            # Check if this is a multi-select checkbox (e.g., access_phone, access_tv)
            elif '_' in name_lower:
                # Try to match against multi-select options
                # e.g., "access_phone" should match if "Telephone" is in the multi-select answer
                parts = name_lower.split('_')
                if len(parts) >= 2:
                    prefix = parts[0]  # e.g., "access"
                    option = '_'.join(parts[1:])  # e.g., "phone", "tv"

                    # Map common abbreviations to full words
                    option_map = {
                        'phone': 'telephone',
                        'tv': 'television',
                        'daycare': 'day care',
                    }
                    full_option = option_map.get(option, option)

                    # Check if this option is in any multi-select field
                    for key, selected_options in multi_select_map.items():
                        if full_option in selected_options or option in selected_options:
                            value = 'X'
                            break

                    # Check if this is a single-select option (e.g., residence_nursing)
                    if value is None and prefix in ['residence']:
                        # Match residence_nursing to "Nursing home/Assisted Living"
                        residence_map = {
                            'own': ['own home', 'apartment'],
                            'guardian': ['guardian', 'guardian\'s home', 'lives in your home'],
                            'relative': ['relative', 'other relative'],
                            'nursing': ['nursing', 'assisted living'],
                            'group': ['group home'],
                            'hospital': ['hospital', 'medical facility'],
                            'state': ['state school', 'state supported'],
                            'other': ['other']
                        }

                        # Look for matching answer in data
                        for key, val in data_dict.items():
                            answer_lower = str(val).strip().lower()
                            # Check if this residence type matches
                            if option in residence_map:
                                for keyword in residence_map[option]:
                                    if keyword in answer_lower:
                                        value = 'X'
                                        break
                            if value:
                                break

                    # If not found in multi-select or single-select, might still be a regular field
                    if value is None:
                        # Try direct match
                        for key, val in data_dict.items():
                            if key.lower() == name_lower:
                                value = val
                                break
            else:
                # Try direct match first
                for key, val in data_dict.items():
                    if key.lower() == name_lower:
                        value = val
                        break

                # Try with cond_ prefix removed (e.g., cond_needs_help → needs_help)
                if value is None and name_lower.startswith('cond_'):
                    base_name = name_lower[5:]  # Remove 'cond_' prefix
                    for key, val in data_dict.items():
                        if key.lower() == base_name:
                            value = val
                            break

                # Try aliases if no direct match
                if value is None and name_lower in column_aliases:
                    alias = column_aliases[name_lower]
                    for key, val in data_dict.items():
                        if key.lower() == alias.lower():
                            value = val
                            break

            # Fill the control if we found a value
            if value is not None and value != "":
                if fill_content_control(cc, value, verbose=verbose):
                    filled_count += 1

    except Exception as e:
        if verbose:
            print(f"  ERROR iterating content controls: {e}")

    return filled_count


def fill_content_controls_from_row(doc, row, header_map, column_aliases=None, verbose=False):
    """
    Fill all content controls in a Word document from a pandas DataFrame row.

    Args:
        doc: Word Document COM object
        row: pandas Series (DataFrame row)
        header_map: Dict mapping lowercase column names to actual column names
        column_aliases: Optional dict mapping alternative names to canonical names
        verbose: If True, print detailed debugging info

    Returns:
        Number of content controls successfully filled
    """
    if column_aliases is None:
        column_aliases = {}

    def get_value_from_row(row, header_map, field_name):
        """
        Get a value from the row by field name (case-insensitive).
        Returns the value if found and not empty, empty string if column exists but is empty,
        or None if column doesn't exist.
        """
        field_lower = field_name.lower()

        # Try direct match
        if field_lower in header_map:
            col = header_map[field_lower]
            val = row.get(col)
            if val is not None and str(val).strip() and str(val).lower() not in ['nan', 'nat', 'none']:
                return str(val).strip()
            else:
                # Column exists but is empty - return empty string to clear placeholder
                return ""

        return None

    filled_count = 0

    try:
        for cc in doc.ContentControls:
            # Get the content control name
            name = (cc.Title or cc.Tag or "").strip()
            if not name:
                continue

            # Get value from row
            value = get_value_from_row(row, header_map, name)

            # Try aliases if no direct match
            if value is None:
                name_lower = name.lower()
                if name_lower in column_aliases:
                    alias = column_aliases[name_lower]
                    value = get_value_from_row(row, header_map, alias)

            # Fill the control if we found a value (including empty string to clear placeholders)
            if value is not None:
                if fill_content_control(cc, value, verbose=verbose):
                    filled_count += 1

    except Exception as e:
        if verbose:
            print(f"  ERROR iterating content controls: {e}")

    return filled_count


# Predefined column aliases commonly used in GuardianShip app
DEFAULT_COLUMN_ALIASES = {
    "wlast": "wardlast",
    "wfirst": "wardfirst",
    "wmiddle": "wardmiddle",
    "datearpfiled": "DateARPfiled",
    "datefiled": "DateARPfiled",
    "cause": "causeno",
    "cause no": "causeno",
    "cause number": "causeno",
}
