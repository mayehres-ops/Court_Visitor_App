"""
Step 10: Auto-fill CVR documents with Google Form responses

This script:
1. Reads form responses from Google Sheets
2. Matches responses to CVR files using cause number and ward name
3. Auto-fills ONLY the fields NOT already filled by Step 8
4. Handles checkboxes, text, and multi-select properly
"""

import os
import sys
import json
import pandas as pd
import win32com.client
from pathlib import Path

# Add Scripts directory to path for imports
sys.path.insert(0, str(Path(__file__).parent / "Scripts"))

try:
    from cvr_content_control_utils import fill_content_controls_from_dict
except ImportError:
    print("ERROR: Could not import cvr_content_control_utils")
    print("Make sure C:\\GoogleSync\\GuardianShip_App\\Scripts\\cvr_content_control_utils.py exists")
    sys.exit(1)

# Google API imports
try:
    from google.auth.transport.requests import Request
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    from google.oauth2 import service_account
    import pickle
except ImportError:
    print("ERROR: Google API libraries not installed")
    print("Run: pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client")
    sys.exit(1)


class GoogleFormCVRAutofill:
    def __init__(self, config_dir=None):
        self.config_dir = Path(config_dir or r"C:\GoogleSync\GuardianShip_App\Config")
        self.credentials = None
        self.sheets_service = None
        self.field_mapping = self.load_field_mapping()

    def load_field_mapping(self):
        """Load field mapping configuration"""
        mapping_path = self.config_dir / "cvr_google_form_mapping.json"
        try:
            with open(mapping_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            print(f"ERROR: Mapping file not found: {mapping_path}")
            return {}
        except json.JSONDecodeError as e:
            print(f"ERROR: Invalid JSON in mapping file: {e}")
            return {}

    def authenticate_service_account(self):
        """Authenticate using service account"""
        creds_path = self.config_dir / "API" / "google_service_account.json"

        if not creds_path.exists():
            print(f"Service account credentials not found: {creds_path}")
            return False

        try:
            credentials = service_account.Credentials.from_service_account_file(
                str(creds_path),
                scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
            )
            self.credentials = credentials
            self.sheets_service = build('sheets', 'v4', credentials=credentials)
            print("[OK] Authenticated with service account")
            return True
        except Exception as e:
            print(f"Service account authentication failed: {e}")
            return False

    def authenticate_oauth(self):
        """Authenticate using OAuth (fallback)"""
        token_path = self.config_dir / "API" / "google_token.pickle"
        creds_path = self.config_dir / "API" / "gmail_token.json"

        creds = None

        # Try to load existing token
        if token_path.exists():
            try:
                with open(token_path, 'rb') as token:
                    creds = pickle.load(token)
            except Exception as e:
                print(f"Could not load OAuth token: {e}")

        # Refresh or create new token
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())
                except Exception as e:
                    # Token refresh failed (expired/revoked) - delete and re-authenticate
                    print(f"Could not refresh OAuth token: {e}")
                    print("Deleting expired token and starting fresh OAuth flow...")
                    if token_path.exists():
                        token_path.unlink()
                    creds = None  # Force re-auth below

            if not creds:
                if not creds_path.exists():
                    print(f"OAuth credentials not found: {creds_path}")
                    return False

                try:
                    flow = InstalledAppFlow.from_client_secrets_file(
                        str(creds_path),
                        scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
                    )
                    creds = flow.run_local_server(port=0)
                except Exception as e:
                    print(f"OAuth authentication failed: {e}")
                    return False

            # Save token
            with open(token_path, 'wb') as token:
                pickle.dump(creds, token)

        self.credentials = creds
        self.sheets_service = build('sheets', 'v4', credentials=creds)
        print("[OK] Authenticated with OAuth")
        return True

    def authenticate(self):
        """Authenticate with Google Sheets API"""
        print("Authenticating with Google Sheets API...")

        # Try service account first
        if self.authenticate_service_account():
            return True

        # Fall back to OAuth
        print("Service account failed, trying OAuth...")
        return self.authenticate_oauth()

    def get_form_responses(self, spreadsheet_id):
        """Get form responses from Google Sheets"""
        try:
            # Get sheet metadata
            sheet = self.sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
            sheet_name = sheet['sheets'][0]['properties']['title']

            # Read all data
            result = self.sheets_service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f"{sheet_name}!A:AZ"
            ).execute()

            values = result.get('values', [])
            if not values:
                print("No data found in spreadsheet")
                return pd.DataFrame()

            # Convert to DataFrame
            headers = values[0]
            data_rows = values[1:]

            # Pad rows to match header length
            padded_rows = []
            for row in data_rows:
                padded_row = row + [''] * (len(headers) - len(row))
                padded_rows.append(padded_row)

            df = pd.DataFrame(padded_rows, columns=headers)
            print(f"[OK] Retrieved {len(df)} form responses")

            return df

        except Exception as e:
            print(f"ERROR getting form responses: {e}")
            return pd.DataFrame()

    def match_response_to_excel(self, response, excel_df):
        """Match a form response to an Excel row using cause number OR ward name OR supplemented_by"""
        # Get matching field names from config
        matching_config = self.field_mapping.get('_matching_fields', {})
        cause_field = matching_config.get('cause_number', 'What is the cause number of this case? Please only provide the last 8 numbers like this: xx-xxxxxx')
        first_field = matching_config.get('ward_first', 'What is the first name of the person name under your care?')
        last_field = matching_config.get('ward_last', 'What is the last name of the person name under your care?')
        supplemented_first_field = matching_config.get('supplemented_by_first', 'What is your first name?')
        supplemented_last_field = matching_config.get('supplemented_by_last', 'What is your Last name?')

        # Get values from form response
        form_cause = str(response.get(cause_field, '')).strip()
        form_ward_first = str(response.get(first_field, '')).strip().lower()
        form_ward_last = str(response.get(last_field, '')).strip().lower()
        form_supp_first = str(response.get(supplemented_first_field, '')).strip().lower()
        form_supp_last = str(response.get(supplemented_last_field, '')).strip().lower()

        print(f"  Looking for match: cause='{form_cause}', ward='{form_ward_first} {form_ward_last}', supp_by='{form_supp_first} {form_supp_last}'")

        # Find matching row in Excel - try multiple matching strategies
        for idx, row in excel_df.iterrows():
            excel_cause = str(row.get('causeno', '')).strip()
            excel_ward_first = str(row.get('wardfirst', '')).strip().lower()
            excel_ward_last = str(row.get('wardlast', '')).strip().lower()
            excel_guardian_first = str(row.get('guardfirst', '')).strip().lower()
            excel_guardian_last = str(row.get('guardlast', '')).strip().lower()

            # Strategy 1: Match by cause number (most reliable)
            if form_cause and excel_cause and excel_cause == form_cause:
                print(f"  [MATCH] Found by cause number: {excel_cause}")
                return row

            # Strategy 2: Match by ward name (first + last)
            if (form_ward_first and form_ward_last and
                excel_ward_first and excel_ward_last and
                excel_ward_first == form_ward_first and excel_ward_last == form_ward_last):
                print(f"  [MATCH] Found by ward name: {form_ward_first} {form_ward_last}")
                return row

            # Strategy 3: Match by supplemented_by name (person filling out form, likely the guardian)
            if (form_supp_first and form_supp_last and
                excel_guardian_first and excel_guardian_last and
                excel_guardian_first == form_supp_first and excel_guardian_last == form_supp_last):
                print(f"  [MATCH] Found by guardian name: {form_supp_first} {form_supp_last}")
                return row

        return None

    def build_cvr_data_dict(self, form_response):
        """Build a dictionary of CVR field names to values from form response"""
        data_dict = {}

        for question, config in self.field_mapping.items():
            # Skip meta fields
            if question.startswith('_'):
                continue

            cvr_control = config.get('cvr_control')
            field_type = config.get('type')

            # Get response value
            value = form_response.get(question, '')
            if not value or str(value).strip() == '':
                continue

            value_str = str(value).strip()

            # Convert based on type
            if field_type == 'yesno':
                # Convert Yes/No to boolean for checkboxes, or keep as text
                data_dict[cvr_control] = value_str
            elif field_type == 'checkbox_list':
                # Multi-select checkboxes - parse comma-separated values
                # and map to individual CVR controls
                mapping = config.get('mapping', {})
                if mapping:
                    # Split by comma and check each item against mapping
                    selected_items = [item.strip() for item in value_str.split(',')]
                    for form_option, cvr_checkbox in mapping.items():
                        # Check if this option was selected
                        if any(form_option.lower() in item.lower() for item in selected_items):
                            data_dict[cvr_checkbox] = 'X'
                else:
                    # No mapping, store as-is
                    data_dict[cvr_control] = value_str
            elif field_type == 'choice':
                # Single-choice question - map to CVR checkbox
                mapping = config.get('mapping', {})
                if mapping:
                    # Check which option was selected
                    for form_option, cvr_checkbox in mapping.items():
                        if form_option.lower() in value_str.lower():
                            data_dict[cvr_checkbox] = 'X'
                            break
                else:
                    data_dict[cvr_control] = value_str
            elif field_type == 'checkbox':
                # Single checkbox
                data_dict[cvr_control] = value_str
            else:
                # Text, longtext
                data_dict[cvr_control] = value_str

        # Special case: Combine first and last name for supplemented_by if not already set
        if 'supplemented_by' not in data_dict or not data_dict['supplemented_by']:
            first_name = form_response.get("What is your first name?", '').strip()
            last_name = form_response.get("What is your Last name?", '').strip()

            if first_name or last_name:
                # Combine first and last name
                full_name = ' '.join([first_name, last_name]).strip()
                if full_name:
                    data_dict['supplemented_by'] = full_name

        return data_dict

    def find_cvr_file(self, excel_row, cvr_folder):
        """Find the CVR file for a given case"""
        cause = str(excel_row.get('causeno', '')).strip()
        ward_first = str(excel_row.get('wardfirst', '')).strip()
        ward_last = str(excel_row.get('wardlast', '')).strip()

        if not cause:
            return None

        # Search for CVR file
        cvr_folder = Path(cvr_folder)

        for root, dirs, files in os.walk(cvr_folder):
            for file in files:
                if file.endswith('.docx') and not file.startswith('~$'):
                    # Check if filename contains cause number
                    if cause in file and 'Court Visitor Report' in file:
                        return Path(root) / file

        return None

    def fill_cvr_document(self, cvr_path, form_data_dict, verbose=True):
        """Fill a CVR document with form data"""
        import pythoncom

        try:
            print(f"\nOpening CVR: {cvr_path.name}")

            # Initialize COM
            pythoncom.CoInitialize()

            # Open Word - try early binding first for better reliability
            import os
            try:
                word = win32com.client.gencache.EnsureDispatch('Word.Application')
                print(f"  [OK] Word started (early binding)")
            except Exception as e:
                print(f"  Early binding failed, using late binding...")
                word = win32com.client.Dispatch('Word.Application')
                print(f"  [OK] Word started (late binding)")

            try:
                print(f"  Opening document: {cvr_path.name}")
                abs_path = os.path.abspath(str(cvr_path))
                print(f"  Full path: {abs_path}")

                # Open with explicit FileName parameter
                doc = word.Documents.Open(FileName=abs_path, ReadOnly=False, ConfirmConversions=False)
                print(f"  [OK] Document opened")

                # UNPROTECT the document so we can edit it
                protection_type = doc.ProtectionType
                if protection_type != -1:  # -1 = wdNoProtection
                    print(f"  Document is protected (type: {protection_type}), unprotecting...")
                    try:
                        doc.Unprotect("")  # Empty password
                        print(f"  [OK] Document unprotected")
                    except Exception as e:
                        # Try without password arg
                        try:
                            doc.Unprotect()
                            print(f"  [OK] Document unprotected")
                        except Exception as e2:
                            print(f"  [ERROR] Could not unprotect document: {e2}")

                # Turn off Design Mode if it's on (critical - hides values!)
                try:
                    # Toggle Design Mode off (two calls ensure it's off)
                    word.CommandBars.ExecuteMso("DeveloperDesignMode")
                    word.CommandBars.ExecuteMso("DeveloperDesignMode")
                    print(f"  [OK] Design Mode toggled off")
                except:
                    pass  # CommandBars might not be available

                # Get list of already-filled controls (from Step 8)
                filled_controls = set()
                checkbox_placeholders = ['☐', '☑', '☒', '□', '■', '▢', '▣', 'X', 'x', ' ', '']
                for cc in doc.ContentControls:
                    name = (cc.Title or cc.Tag or '').strip()
                    if name and cc.Range.Text and cc.Range.Text.strip():
                        # Check if it's not just placeholder text or checkbox placeholder
                        text = cc.Range.Text.strip()
                        # Skip placeholder text
                        if text in ['Click or tap here to enter text.', 'Click here to enter text.']:
                            continue
                        # Skip single-character checkbox placeholders
                        if len(text) == 1 and text in checkbox_placeholders:
                            continue
                        # This control has real data from Step 8
                        if text:
                            filled_controls.add(name.lower())

                if verbose and filled_controls:
                    print(f"  Skipping {len(filled_controls)} already-filled controls from Step 8")

                # Filter out already-filled fields
                fields_to_fill = {k: v for k, v in form_data_dict.items() if k.lower() not in filled_controls}

                if verbose:
                    print(f"  Filling {len(fields_to_fill)} fields from Google Form:")
                    for name, value in list(fields_to_fill.items())[:5]:
                        print(f"    - {name}: {value[:50]}..." if len(value) > 50 else f"    - {name}: {value}")
                    if len(fields_to_fill) > 5:
                        print(f"    ... and {len(fields_to_fill) - 5} more")

                # Track which fields get filled
                fields_before = set(fields_to_fill.keys())

                # Fill the document
                filled_count = fill_content_controls_from_dict(doc, fields_to_fill, verbose=True)

                print(f"  [OK] Filled {filled_count} content controls")

                # Debug: Show which fields were NOT filled
                if filled_count < len(fields_to_fill):
                    unfilled_count = len(fields_to_fill) - filled_count
                    print(f"\n  [WARN] {unfilled_count} field(s) were not filled:")
                    print(f"  (These might not exist in the CVR template, or have different names)")

                    # Show unfilled field names
                    # Since we can't easily track which were filled, show all fields and mark status
                    print(f"\n  Fields attempted:")
                    for field_name in sorted(fields_to_fill.keys()):
                        print(f"    - {field_name}")
                    print(f"\n  Please check which of these should exist in the CVR template.")

                # RE-PROTECT the document if it was originally protected
                if protection_type != -1:  # Was protected before
                    try:
                        # Re-protect with same type (2 = wdAllowOnlyFormFields allows form filling)
                        doc.Protect(Type=2, NoReset=True)
                        print(f"  [OK] Document re-protected")
                    except Exception as e:
                        print(f"  [WARN] Could not re-protect document: {e}")

                # Save and close
                doc.Save()
                doc.Close()

                return True

            finally:
                try:
                    word.Quit()
                except:
                    pass
                try:
                    pythoncom.CoUninitialize()
                except:
                    pass

        except Exception as e:
            print(f"  ERROR filling CVR: {e}")
            try:
                pythoncom.CoUninitialize()
            except:
                pass
            return False

    def process_all_responses(self, spreadsheet_id, excel_path, cvr_folder):
        """Process all form responses and fill CVRs"""
        print("=" * 80)
        print("STEP 10: AUTO-FILL CVR FROM GOOGLE FORM RESPONSES")
        print("=" * 80)

        # Authenticate
        if not self.authenticate():
            print("\n[FAIL] Authentication failed. Please check your API credentials.")
            return False

        # Get form responses
        print("\nFetching form responses...")
        form_df = self.get_form_responses(spreadsheet_id)

        if form_df.empty:
            print("[FAIL] No form responses found")
            return False

        print(f"[OK] Found {len(form_df)} form response(s)")

        # Load Excel data
        print(f"\nLoading Excel data from: {excel_path}")
        try:
            excel_df = pd.read_excel(excel_path)
            print(f"[OK] Loaded {len(excel_df)} cases from Excel")
        except Exception as e:
            print(f"[FAIL] Could not load Excel file: {e}")
            return False

        # Process each form response
        print(f"\nProcessing {len(form_df)} form response(s)...")
        print("=" * 80)

        success_count = 0
        skipped_count = 0
        error_count = 0

        for idx, response in form_df.iterrows():
            try:
                # Match to Excel
                excel_row = self.match_response_to_excel(response, excel_df)

                if excel_row is None:
                    # Try to get cause from correct field name in config
                    matching_config = self.field_mapping.get('_matching_fields', {})
                    cause_field = matching_config.get('cause_number', 'What is the cause number of this case? Please only provide the last 8 numbers like this: xx-xxxxxx')
                    cause = response.get(cause_field, 'unknown')
                    print(f"\n[WARN] Response {idx + 1}: No matching case found (cause={cause})")
                    skipped_count += 1
                    continue

                cause = excel_row.get('causeno', '')
                ward_name = f"{excel_row.get('wardfirst', '')} {excel_row.get('wardlast', '')}"
                print(f"\nResponse {idx + 1}: {ward_name} ({cause})")

                # Find CVR file
                cvr_file = self.find_cvr_file(excel_row, cvr_folder)

                if not cvr_file:
                    print(f"  [WARN] CVR file not found")
                    skipped_count += 1
                    continue

                # Build data dictionary
                form_data = self.build_cvr_data_dict(response)

                if not form_data:
                    print(f"  [WARN] No form data to fill")
                    skipped_count += 1
                    continue

                # Fill CVR
                if self.fill_cvr_document(cvr_file, form_data, verbose=True):
                    success_count += 1
                else:
                    error_count += 1

            except Exception as e:
                print(f"\n[ERROR] Error processing response {idx + 1}: {e}")
                error_count += 1

        # Summary
        print("\n" + "=" * 80)
        print("SUMMARY")
        print("=" * 80)
        print(f"[OK] Successfully filled: {success_count} CVR(s)")
        if skipped_count > 0:
            print(f"[WARN] Skipped: {skipped_count}")
        if error_count > 0:
            print(f"[ERROR] Errors: {error_count}")
        print("=" * 80)

        return success_count > 0


def main():
    """Main entry point"""
    # Configuration
    SPREADSHEET_ID = "1O9Sv5M8SEdD_bbxew28QScKOazCYivUTvTxpMfZl1HI"
    EXCEL_PATH = r"C:\GoogleSync\GuardianShip_App\App Data\ward_guardian_info.xlsx"
    CVR_FOLDER = r"C:\GoogleSync\GuardianShip_App\New Clients"

    # Create processor
    processor = GoogleFormCVRAutofill()

    # Process all responses
    result = processor.process_all_responses(SPREADSHEET_ID, EXCEL_PATH, CVR_FOLDER)

    # Exit with proper error code so GUI can detect failure
    if not result:
        print("\n[FAIL] Google Forms CVR Integration FAILED - see errors above")
        sys.exit(1)
    else:
        print("\n[OK] Google Forms CVR Integration SUCCESS")
        sys.exit(0)


if __name__ == "__main__":
    main()
