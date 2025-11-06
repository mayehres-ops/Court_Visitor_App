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
            print("✓ Authenticated with service account")
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
                    print(f"Could not refresh OAuth token: {e}")
                    return False
            else:
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
        print("✓ Authenticated with OAuth")
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
            print(f"✓ Retrieved {len(df)} form responses")

            return df

        except Exception as e:
            print(f"ERROR getting form responses: {e}")
            return pd.DataFrame()

    def match_response_to_excel(self, response, excel_df):
        """Match a form response to an Excel row"""
        # Get matching field names from config
        matching_config = self.field_mapping.get('_matching_fields', {})
        cause_field = matching_config.get('cause_number', 'Case/Cause Number (found on court documents)')
        first_field = matching_config.get('ward_first', 'Ward First Name')
        middle_field = matching_config.get('ward_middle', 'Ward Middle Name')
        last_field = matching_config.get('ward_last', 'Ward Last Name')

        # Get values from form response
        form_cause = str(response.get(cause_field, '')).strip()
        form_first = str(response.get(first_field, '')).strip().lower()
        form_middle = str(response.get(middle_field, '')).strip().lower()
        form_last = str(response.get(last_field, '')).strip().lower()

        if not form_cause:
            return None

        # Find matching row in Excel
        for idx, row in excel_df.iterrows():
            excel_cause = str(row.get('causeno', '')).strip()
            excel_first = str(row.get('wardfirst', '')).strip().lower()
            excel_middle = str(row.get('wardmiddle', '')).strip().lower()
            excel_last = str(row.get('wardlast', '')).strip().lower()

            # Match on cause number (required) + name components
            if excel_cause == form_cause:
                # Check name match (at least first and last must match)
                if excel_first == form_first and excel_last == form_last:
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
            elif field_type == 'checkbox':
                # Multi-select checkboxes come as comma-separated
                # Keep as comma-separated text for now
                data_dict[cvr_control] = value_str
            else:
                # Text, longtext, choice
                data_dict[cvr_control] = value_str

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
        try:
            print(f"\nOpening CVR: {cvr_path.name}")

            # Open Word
            word = win32com.client.Dispatch('Word.Application')
            word.Visible = False  # Keep invisible during filling

            try:
                doc = word.Documents.Open(str(cvr_path.absolute()))

                # Get list of already-filled controls (from Step 8)
                filled_controls = set()
                for cc in doc.ContentControls:
                    name = (cc.Title or cc.Tag or '').strip()
                    if name and cc.Range.Text and cc.Range.Text.strip():
                        # Check if it's not just placeholder text
                        text = cc.Range.Text.strip()
                        if text and text not in ['Click or tap here to enter text.', 'Click here to enter text.']:
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

                # Fill the document
                filled_count = fill_content_controls_from_dict(doc, fields_to_fill, verbose=False)

                print(f"  ✓ Filled {filled_count} content controls")

                # Save and close
                doc.Save()
                doc.Close()

                return True

            finally:
                word.Quit()

        except Exception as e:
            print(f"  ERROR filling CVR: {e}")
            return False

    def process_all_responses(self, spreadsheet_id, excel_path, cvr_folder):
        """Process all form responses and fill CVRs"""
        print("=" * 80)
        print("STEP 10: AUTO-FILL CVR FROM GOOGLE FORM RESPONSES")
        print("=" * 80)

        # Authenticate
        if not self.authenticate():
            print("\n❌ Authentication failed. Please check your API credentials.")
            return False

        # Get form responses
        print("\nFetching form responses...")
        form_df = self.get_form_responses(spreadsheet_id)

        if form_df.empty:
            print("❌ No form responses found")
            return False

        print(f"✓ Found {len(form_df)} form response(s)")

        # Load Excel data
        print(f"\nLoading Excel data from: {excel_path}")
        try:
            excel_df = pd.read_excel(excel_path)
            print(f"✓ Loaded {len(excel_df)} cases from Excel")
        except Exception as e:
            print(f"❌ Could not load Excel file: {e}")
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
                    cause = response.get('Case/Cause Number (found on court documents)', 'unknown')
                    print(f"\n⚠ Response {idx + 1}: No matching case found for cause {cause}")
                    skipped_count += 1
                    continue

                cause = excel_row.get('causeno', '')
                ward_name = f"{excel_row.get('wardfirst', '')} {excel_row.get('wardlast', '')}"
                print(f"\nResponse {idx + 1}: {ward_name} ({cause})")

                # Find CVR file
                cvr_file = self.find_cvr_file(excel_row, cvr_folder)

                if not cvr_file:
                    print(f"  ⚠ CVR file not found")
                    skipped_count += 1
                    continue

                # Build data dictionary
                form_data = self.build_cvr_data_dict(response)

                if not form_data:
                    print(f"  ⚠ No form data to fill")
                    skipped_count += 1
                    continue

                # Fill CVR
                if self.fill_cvr_document(cvr_file, form_data, verbose=True):
                    success_count += 1
                else:
                    error_count += 1

            except Exception as e:
                print(f"\n❌ Error processing response {idx + 1}: {e}")
                error_count += 1

        # Summary
        print("\n" + "=" * 80)
        print("SUMMARY")
        print("=" * 80)
        print(f"✓ Successfully filled: {success_count} CVR(s)")
        if skipped_count > 0:
            print(f"⚠ Skipped: {skipped_count}")
        if error_count > 0:
            print(f"❌ Errors: {error_count}")
        print("=" * 80)

        return success_count > 0


def main():
    """Main entry point"""
    # Configuration
    SPREADSHEET_ID = "1O9Sv5M8SEdD_bbxew28QScKOazCYivUTvTxpMfZl1HI"
    EXCEL_PATH = r"C:\GoogleSync\GuardianShip_App\App Data\ward_guardian_info.xlsx"
    CVR_FOLDER = r"C:\GoogleSync\GuardianShip_App\New Files"

    # Create processor
    processor = GoogleFormCVRAutofill()

    # Process all responses
    processor.process_all_responses(SPREADSHEET_ID, EXCEL_PATH, CVR_FOLDER)


if __name__ == "__main__":
    main()
