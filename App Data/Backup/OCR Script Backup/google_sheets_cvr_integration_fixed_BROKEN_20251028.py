import os
import pickle
import pandas as pd
import win32com.client
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import json

class GoogleSheetsCVRIntegration:
    def __init__(self):
        self.credentials = None
        self.sheets_service = None
        self.drive_service = None
        self.field_mapping_config = self.load_field_mapping_config()
    
    def load_field_mapping_config(self):
        """Load field mapping configuration from JSON file"""
        try:
            # Try multiple possible paths
            config_paths = [
                r"C:\GoogleSync\GuardianShip_App\Config\cvr_google_form_mapping.json",
                "cvr_field_mapping_config.json",
                "cvr_google_form_mapping.json"
            ]

            for config_path in config_paths:
                if os.path.exists(config_path):
                    with open(config_path, 'r') as f:
                        config = json.load(f)
                    print(f"Field mapping configuration loaded from: {config_path}")
                    return config

            print("ERROR: Field mapping configuration file not found")
            print(f"Searched paths: {config_paths}")
            return None
        except Exception as e:
            print(f"ERROR: Loading field mapping configuration: {e}")
            return None
    
    def authenticate(self):
        """Authenticate with Google APIs"""
        try:
            # Define required scopes
            SCOPES = [
                'https://www.googleapis.com/auth/spreadsheets.readonly',
                'https://www.googleapis.com/auth/drive.readonly',
                'https://www.googleapis.com/auth/gmail.readonly'
            ]
            
            # Check for existing credentials
            # Step 10 uses OAuth (like Gmail/Calendar), NOT service account
            oauth_client_path = r"C:\configlocal\API\client_secret_sheets.json"
            token_path = r"C:\configlocal\API\sheets_token.pickle"
            
            # Try to load existing token first
            if os.path.exists(token_path):
                try:
                    with open(token_path, 'rb') as token:
                        self.credentials = pickle.load(token)
                    print("Loaded existing token")
                except:
                    print("Could not load existing token, will re-authenticate")
                    self.credentials = None
            else:
                self.credentials = None
            
            # If no valid credentials, authenticate
            if not self.credentials or not self.credentials.valid:
                if self.credentials and self.credentials.expired and self.credentials.refresh_token:
                    print("Refreshing expired token...")
                    try:
                        self.credentials.refresh(Request())
                        print("Token refreshed successfully")
                    except Exception as e:
                        print(f"Token refresh failed: {e}, will re-authenticate")
                        self.credentials = None
                
                if not self.credentials or not self.credentials.valid:
                    print("No valid credentials, starting OAuth flow...")
                    if os.path.exists(oauth_client_path):
                        from google_auth_oauthlib.flow import InstalledAppFlow
                        flow = InstalledAppFlow.from_client_secrets_file(oauth_client_path, SCOPES)
                        self.credentials = flow.run_local_server(port=0)
                    else:
                        print(f"ERROR: OAuth client secret not found at: {oauth_client_path}")
                        print("Step 10 requires OAuth credentials (client_secret_sheets.json)")
                        print("This is different from the service account JSON used in Step 1")
                        raise Exception("Google Sheets OAuth credentials not found. Please set up OAuth client secret.")
                
                # Save the token for future use
                with open(token_path, 'wb') as token:
                    pickle.dump(self.credentials, token)
                print("Token saved for future use")
            
            # Build services
            self.sheets_service = build('sheets', 'v4', credentials=self.credentials)
            self.drive_service = build('drive', 'v3', credentials=self.credentials)
            
            return True
            
        except Exception as e:
            print(f"Authentication error: {str(e)}")
            return False
    
    def get_form_responses(self, spreadsheet_id):
        """Get form responses from Google Sheets"""
        try:
            # Get the sheet data
            sheet = self.sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
            sheet_name = sheet['sheets'][0]['properties']['title']
            
            # Read the data
            result = self.sheets_service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f"{sheet_name}!A:Z"
            ).execute()
            
            values = result.get('values', [])
            if not values:
                print("No data found in the sheet")
                return []
            
            # Convert to DataFrame
            headers = values[0]
            data = values[1:]
            df = pd.DataFrame(data, columns=headers)

            print(f"Retrieved {len(df)} form responses")
            print(f"Google Sheet columns: {headers}")
            if len(df) > 0:
                print(f"First response sample: {dict(df.iloc[0])}")
            return df
            
        except Exception as e:
            print(f"Error getting form responses: {str(e)}")
            return []
    
    def match_responses_to_cases(self, form_responses, excel_file=r"C:\GoogleSync\GuardianShip_App\App Data\ward_guardian_info.xlsx"):
        """Match form responses to cases in Excel file using Ward Name and Guardian Name"""
        try:
            # Read Excel file
            excel_df = pd.read_excel(excel_file)
            
            # Create a mapping dictionary
            matched_cases = {}
            
            for _, response in form_responses.iterrows():
                # Get form response data using EXACT Google Sheets column headers
                # Cause number - use the exact question text from the form
                cause_number = str(response.get('What is the cause number of this case? Please only provide the last 8 numbers like this: xx-xxxxxx', '')).strip()

                # Person filling out form (supplemented_by, NOT guardian)
                supplemented_first = str(response.get('What is your first name?', '')).strip()

                # Ward name - not collected in separate fields, so we'll match by cause number only
                ward_name = ''  # Not available in this form structure

                print(f"Looking for match: Cause='{cause_number}', Filled out by='{supplemented_first}'")
                
                # Find matching case in Excel - prefer cause number matching first
                matching_case = None
                for _, case in excel_df.iterrows():
                    # Get Excel cause number
                    excel_cause = str(case.get('causeno', '')).strip()

                    # Try to match by cause number first (most reliable)
                    if cause_number and excel_cause:
                        if cause_number.lower() == excel_cause.lower():
                            print(f"  [MATCH] Found case by cause number: {cause_number}")
                            matching_case = case
                            break

                    # No cause number match, continue to next Excel row
                    continue
                
                if matching_case is not None:
                    # Use ward name as key
                    case_key = ward_name or f"Unknown_{len(matched_cases)}"
                    matched_cases[case_key] = {
                        'excel_data': matching_case,
                        'form_response': response
                    }
                    print(f"Matched case: {case_key}")
                else:
                    print(f"No match found for: Cause='{cause_number}', Filled out by='{supplemented_first}'")
            
            return matched_cases
            
        except Exception as e:
            print(f"Error matching responses to cases: {str(e)}")
            return {}

    def _build_control_index(self, doc):
        """
        Build an index of all content controls in the document for fast lookup.

        Args:
            doc: Word Document object

        Returns:
            Dictionary mapping control names (lowercase) to ContentControl objects
        """
        control_index = {}
        try:
            for cc in doc.ContentControls:
                cc_name = (cc.Title or cc.Tag or '').strip()
                if cc_name:
                    # Store with lowercase key for case-insensitive lookup
                    control_index[cc_name.lower()] = cc
        except Exception as e:
            print(f"Warning: Error building control index: {e}")

        return control_index

    def _fill_control_from_index(self, control_index, control_name, value):
        """
        Fill a content control using a pre-built index.

        FIXED: Uses correct checkbox type (8 not 5), unlocks controls, handles XML mapping

        Args:
            control_index: Dictionary of control name -> ContentControl object
            control_name: Name of the content control (Title or Tag)
            value: Value to fill (str or bool for checkboxes)

        Returns:
            True if control was found and filled, False otherwise
        """
        try:
            # Look up control in index (case-insensitive)
            cc = control_index.get(control_name.lower())

            if not cc:
                return False

            # UNLOCK the control before modifying (critical fix!)
            try:
                cc.LockContentControl = False
                if cc.LockContents:
                    cc.LockContents = False
            except:
                pass  # Some controls don't have these properties

            # Handle XML-mapped controls (write to CustomXML node if mapped)
            try:
                if cc.XMLMapping and cc.XMLMapping.IsMapped:
                    node = cc.XMLMapping.CustomXMLNode
                    if node is not None:
                        node.Text = "" if value is None else str(value)
                        return True
            except:
                pass  # Fall through to standard write

            # Checkbox control - CORRECT TYPE IS 8, NOT 5!
            # Type 5 = dropdown, Type 8 = checkbox
            if cc.Type == 8:  # wdContentControlCheckBox
                if value:
                    cc.Checked = True
                return True
            # Text/RichText controls (Type 1, 2)
            elif cc.Type in (1, 2):
                cc.Range.Text = str(value) if value else ""
                return True
            # Dropdown/ComboBox (Type 4, 5)
            elif cc.Type in (4, 5):
                cc.Range.Text = str(value) if value else ""
                return True
            # Date control (Type 6)
            elif cc.Type == 6:
                cc.Range.Text = str(value) if value else ""
                return True
            # Fallback for other types
            else:
                cc.Range.Text = str(value) if value else ""
                return True

        except Exception as e:
            print(f"    Error filling control '{control_name}': {e}")
            return False

    def populate_cvr_document(self, cvr_path, case_data):
        """Populate CVR document with form data - simplified approach"""
        try:
            # Get form response data
            form_data = case_data['form_response']
            excel_data = case_data['excel_data']
            
            print("Preparing to populate CVR document...")
            print(f"Form data available: {len(form_data)} fields")
            
            # Create field mapping from config
            # Config structure: { "Google Form Question": { "cvr_control": "control_name", "type": "yesno/text/..." } }
            field_mapping = {}

            for google_question, mapping_info in self.field_mapping_config.items():
                # Skip internal fields like _matching_fields
                if google_question.startswith('_'):
                    continue

                # Check if this Google Form question exists in the form data
                if google_question in form_data:
                    cvr_control = mapping_info.get('cvr_control', '')
                    field_type = mapping_info.get('type', 'text')
                    form_value = form_data[google_question]

                    # Store the mapping
                    if cvr_control:
                        field_mapping[cvr_control] = {
                            'value': form_value,
                            'type': field_type,
                            'mapping_info': mapping_info
                        }
            
            # Debug: Show what fields we're trying to populate
            print(f"Field mapping contains {len(field_mapping)} fields:")
            for field_name, field_data in field_mapping.items():
                value = field_data.get('value', '') if isinstance(field_data, dict) else field_data
                print(f"  - {field_name}: {value}")
            
            # Open Word application - try to get existing instance first, then create new
            try:
                word_app = win32com.client.GetActiveObject("Word.Application")
                print("Using existing Word instance")
            except:
                word_app = win32com.client.Dispatch("Word.Application")
                print("Created new Word instance")

            # Try to set visibility and screen updating (wrap in try/except to handle errors)
            try:
                word_app.Visible = True  # Keep visible so user can see progress
                word_app.ScreenUpdating = True
            except Exception as e:
                print(f"Note: Could not set Word visibility ({e}). Continuing anyway...")

            # Open the document in edit mode (ReadOnly=False)
            try:
                full_path = os.path.abspath(cvr_path)
                print(f"Opening document: {full_path}")
                doc = word_app.Documents.Open(full_path, ReadOnly=False)
                print(f"Document opened successfully")
            except Exception as e:
                print(f"Error opening document: {e}")
                raise

            # Turn OFF Design Mode if it's on (critical - hides values!)
            try:
                if word_app.CommandBars.GetPressedMso("DeveloperDesignMode"):
                    word_app.CommandBars.ExecuteMso("DeveloperDesignMode")
                    print("Design Mode was ON - toggled OFF")
            except:
                pass  # Ignore if CommandBars not available

            # Build control index for fast lookups (instead of iterating for each field)
            print("Building content control index...")
            control_index = self._build_control_index(doc)
            print(f"Found {len(control_index)} named controls in document")

            # Fill content controls with Google Form data
            print("\n=== FILLING CONTENT CONTROLS ===")
            filled_count = 0
            skipped_count = 0
            error_count = 0

            for control_name, field_data in field_mapping.items():
                try:
                    # Get value and type
                    if isinstance(field_data, dict):
                        value = field_data.get('value', '')
                        field_type = field_data.get('type', 'text')
                        mapping_info = field_data.get('mapping_info', {})
                    else:
                        value = field_data
                        field_type = 'text'
                        mapping_info = {}

                    # Skip empty values
                    if not value or str(value).strip() == '':
                        skipped_count += 1
                        continue

                    # Find and fill content control(s)
                    if field_type == 'text':
                        # Simple text field - fill directly
                        filled = self._fill_control_from_index(control_index, control_name, str(value))
                        if filled:
                            print(f"  [OK] {control_name}: {value}")
                            filled_count += 1
                        else:
                            print(f"  [NOT FOUND] Control not found: {control_name}")
                            error_count += 1

                    elif field_type == 'yesno':
                        # Yes/No checkbox - TWO separate checkboxes with _yes and _no suffixes
                        value_lower = str(value).lower()
                        if 'yes' in value_lower:
                            # Check the "_yes" checkbox
                            filled = self._fill_control_from_index(control_index, f"{control_name}_yes", True)
                            if filled:
                                print(f"  [OK] {control_name}_yes: checked")
                                filled_count += 1
                            else:
                                print(f"  [NOT FOUND] Control not found: {control_name}_yes")
                                error_count += 1
                        elif 'no' in value_lower:
                            # Check the "_no" checkbox
                            filled = self._fill_control_from_index(control_index, f"{control_name}_no", True)
                            if filled:
                                print(f"  [OK] {control_name}_no: checked")
                                filled_count += 1
                            else:
                                print(f"  [NOT FOUND] Control not found: {control_name}_no")
                                error_count += 1
                        else:
                            print(f"  [UNCLEAR] Unclear yes/no value for {control_name}: {value}")
                            skipped_count += 1

                    elif field_type == 'checkbox_list':
                        # Multiple checkboxes - value is comma-separated list
                        selected_items = [item.strip() for item in str(value).split(',')]
                        checkbox_mapping = mapping_info.get('mapping', {})

                        for form_option, control_name_suffix in checkbox_mapping.items():
                            if form_option in selected_items:
                                filled = self._fill_control_from_index(control_index, control_name_suffix, True)
                                if filled:
                                    print(f"  [OK] {control_name_suffix}: checked")
                                    filled_count += 1

                    elif field_type == 'choice':
                        # Radio button choice - check one option
                        choice_mapping = mapping_info.get('mapping', {})

                        for form_option, control_name_option in choice_mapping.items():
                            if str(value).strip() == form_option:
                                filled = self._fill_control_from_index(control_index, control_name_option, True)
                                if filled:
                                    print(f"  [OK] {control_name}: {form_option}")
                                    filled_count += 1
                                break

                except Exception as e:
                    print(f"  [ERROR] Error filling {control_name}: {e}")
                    error_count += 1

            print(f"\n=== SUMMARY ===")
            print(f"  Filled: {filled_count} controls")
            print(f"  Skipped (empty): {skipped_count} controls")
            print(f"  Errors: {error_count} controls")

            # Protect the document now that all fields are filled (allow form filling only, no password)
            try:
                doc.Protect(Type=2, NoReset=True)  # 2 = wdAllowOnlyFormFields
                print("\nDocument protected (allows form filling)")
            except Exception as e:
                print(f"\nNote: Could not protect document ({e})")

            # Re-enable screen updating and make Word visible
            try:
                word_app.ScreenUpdating = True
                word_app.Visible = True
                print(f"Document is now visible for your review.")
            except Exception as e:
                print(f"Note: Document is open. Could not set visibility ({e}).")

            return filled_count > 0
            
        except Exception as e:
            print(f"Error populating CVR document: {str(e)}")
            return False
    
    def process_cvr_completion(self, spreadsheet_id, cvr_folder=r"C:\GoogleSync\GuardianShip_App\New Clients"):
        """Process CVR completion with Google Form data"""
        try:
            print("Starting CVR completion process...")
            
            # Authenticate
            if not self.authenticate():
                return False
            
            # Get form responses
            form_responses = self.get_form_responses(spreadsheet_id)
            if len(form_responses) == 0:
                print("No form responses found")
                return False
            
            # Match responses to cases
            matched_cases = self.match_responses_to_cases(form_responses)
            if not matched_cases:
                print("No cases matched with form responses")
                return False
            
            print(f"Found {len(matched_cases)} matched cases")
            
            # Process each matched case
            processed_count = 0
            for case_key, case_data in matched_cases.items():
                try:
                    print(f"Processing CVR for case {case_key}")
                    
                    # Get cause number from Excel data
                    cause_number = str(case_data['excel_data'].get('causeno', '')).strip()
                    print(f"Looking for CVR with cause number: '{cause_number}'")
                    print(f"Case key: '{case_key}'")
                    
                    # Find CVR document for this case
                    cvr_files = []
                    print(f"Searching in folder: {cvr_folder}")
                    for root, dirs, files in os.walk(cvr_folder):
                        for file in files:
                            if (file.endswith('.docx') and 
                                ('cvr' in file.lower() or 'court visitor report' in file.lower()) and
                                not file.startswith('~$')):  # Exclude temporary files
                                # Check if this CVR belongs to this case by cause number or case key
                                print(f"Checking file: {file} for cause: {cause_number}")
                                if cause_number in file or case_key.lower() in file.lower():
                                    cvr_files.append(os.path.join(root, file))
                                    print(f"Found potential CVR: {file}")
                                else:
                                    print(f"  - Cause number {cause_number} not found in {file}")
                                    print(f"  - Case key {case_key.lower()} not found in {file.lower()}")
                    
                    if not cvr_files:
                        print(f"No CVR document found for case {case_key}")
                        continue
                    
                    # Sort by modification time (newest first)
                    cvr_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
                    cvr_path = cvr_files[0]
                    print(f"Using newest CVR: {os.path.basename(cvr_path)}")
                    
                    # Populate the CVR document
                    if self.populate_cvr_document(cvr_path, case_data):
                        print(f"Successfully processed CVR for case {case_key}")
                        processed_count += 1
                    else:
                        print(f"Failed to process CVR for case {case_key}")
                
                except Exception as e:
                    print(f"Error processing case {case_key}: {str(e)}")
                    continue

            print(f"CVR completion process finished. Processed {processed_count} cases.")

            # ALWAYS open the most recent CVR for manual review/completion
            print("\nOpening most recent CVR for review...")
            all_cvr_files = []
            for root, dirs, files in os.walk(cvr_folder):
                for file in files:
                    if (file.endswith('.docx') and
                        ('cvr' in file.lower() or 'court visitor report' in file.lower()) and
                        not file.startswith('~$')):
                        all_cvr_files.append(os.path.join(root, file))

            if all_cvr_files:
                # Sort by modification time (newest first)
                all_cvr_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
                most_recent_cvr = all_cvr_files[0]
                print(f"Opening: {os.path.basename(most_recent_cvr)}")
                try:
                    import subprocess
                    subprocess.Popen(['start', '', most_recent_cvr], shell=True)
                except Exception as e:
                    print(f"Error opening CVR: {str(e)}")
            else:
                print("No CVR files found to open")

            return processed_count > 0
            
        except Exception as e:
            print(f"Error in CVR completion process: {str(e)}")
            return False

if __name__ == "__main__":
    import sys
    integration = GoogleSheetsCVRIntegration()
    result = integration.process_cvr_completion("1O9Sv5M8SEdD_bbxew28QScKOazCYivUTvTxpMfZl1HI")
    print(f"Result: {result}")

    # Exit with proper error code so GUI can detect failure
    if not result:
        print("\n[FAIL] CVR completion FAILED - see errors above")
        sys.exit(1)
    else:
        print("\n[OK] CVR completion SUCCESS")
        sys.exit(0)









