"""
Clinical Audit System - Data Processing Backend
Updated with Gemini 2.5 models and Dynamic Schema Support
Version 2.0
"""

import pandas as pd
import json
import time
import logging
import sys
from datetime import datetime
from typing import Dict, List, Any, Optional, Tuple
from pathlib import Path

# Google imports
import google.generativeai as genai
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('audit_processing.log'),
        logging.StreamHandler()
    ]
)

class GoogleSheetsManager:
    """Manages Google Sheets operations"""
    
    def __init__(self, credentials_file: str, spreadsheet_id: str):
        """Initialize Google Sheets connection"""
        self.spreadsheet_id = spreadsheet_id
        
        try:
            self.creds = service_account.Credentials.from_service_account_file(
                credentials_file,
                scopes=['https://www.googleapis.com/auth/spreadsheets']
            )
            self.service = build('sheets', 'v4', credentials=self.creds)
            self.sheet = self.service.spreadsheets()
            logging.info("Google Sheets connection established")
            
        except FileNotFoundError:
            logging.error(f"Credentials file not found: {credentials_file}")
            raise
        except Exception as e:
            logging.error(f"Failed to connect to Google Sheets: {e}")
            raise
    
    def ensure_sheets_exist(self) -> bool:
        """Create required sheets if they don't exist"""
        try:
            result = self.sheet.get(spreadsheetId=self.spreadsheet_id).execute()
            existing_sheets = [s['properties']['title'] for s in result.get('sheets', [])]
            
            required_sheets = ['Data', 'Validator', 'Lead Review', 'Config', 'Audit Log']
            
            requests = []
            for sheet_name in required_sheets:
                if sheet_name not in existing_sheets:
                    requests.append({
                        'addSheet': {
                            'properties': {
                                'title': sheet_name,
                                'gridProperties': {
                                    'rowCount': 1000,
                                    'columnCount': 40  # Increased for AI source columns
                                }
                            }
                        }
                    })
            
            if requests:
                body = {'requests': requests}
                self.sheet.batchUpdate(
                    spreadsheetId=self.spreadsheet_id,
                    body=body
                ).execute()
                logging.info(f"Created sheets: {[s for s in required_sheets if s not in existing_sheets]}")
            
            return True
            
        except HttpError as error:
            logging.error(f"Error creating sheets: {error}")
            return False
    
    def setup_data_sheet_headers(self) -> bool:
        """Set up headers for the Data sheet - ONLY if empty"""
        
        # Check if headers already exist
        try:
            result = self.sheet.values().get(
                spreadsheetId=self.spreadsheet_id,
                range='Data!A1:AK1'
            ).execute()
            
            existing_values = result.get('values', [])
            
            if existing_values and existing_values[0]:
                logging.info("Headers already exist in Data sheet - skipping setup")
                logging.info(f"Found {len(existing_values[0])} existing columns")
                return True  # Headers already exist
                
        except HttpError:
            pass  # Sheet might not exist yet
        
        # Only create headers if sheet is completely empty
        headers = [[
            'claim_id', 'note_preview', 'full_note_text', 'processing_timestamp',
            'processing_status', 'diagnosis_codes', 'cpt_codes', 'medical_necessity',
            'treatment_goals', 'interventions', 'progress_noted', 'plan_documented',
            'therapist_signature', 'date_of_service', 'session_duration',
            'compliance_issues', 'addendum_needed', 'addendum_reason',
            'validator_decision', 'validator_notes', 'validation_timestamp',
            'validator_email', 'lead_review_status', 'final_addendum',
            'lead_reviewer', 'lead_review_timestamp', 'ai_confidence_score',
            # AI Source columns
            'ai_source_diagnosis_codes', 'ai_source_cpt_codes', 'ai_source_medical_necessity',
            'ai_source_treatment_goals', 'ai_source_interventions', 'ai_source_progress_noted',
            'ai_source_plan_documented', 'ai_source_therapist_signature',
            'ai_source_date_of_service', 'ai_source_session_duration'
        ]]
        
        try:
            body = {'values': headers}
            self.sheet.values().update(
                spreadsheetId=self.spreadsheet_id,
                range='Data!A1',
                valueInputOption='RAW',
                body=body
            ).execute()
            
            logging.info(f"Data sheet headers created with {len(headers[0])} columns")
            return True
            
        except HttpError as error:
            logging.error(f"Error setting up headers: {error}")
            return False
    
    def _get_sheet_id(self, sheet_name: str) -> Optional[int]:
        """Get sheet ID by name"""
        try:
            result = self.sheet.get(spreadsheetId=self.spreadsheet_id).execute()
            for sheet in result.get('sheets', []):
                if sheet['properties']['title'] == sheet_name:
                    return sheet['properties']['sheetId']
            return 0
        except:
            return 0
    
    def append_row(self, sheet_name: str, row_data: List[Any]) -> bool:
        """Append a single row to a sheet"""
        try:
            body = {'values': [row_data]}
            self.sheet.values().append(
                spreadsheetId=self.spreadsheet_id,
                range=f'{sheet_name}!A1',
                valueInputOption='RAW',
                insertDataOption='INSERT_ROWS',
                body=body
            ).execute()
            return True
        except HttpError as error:
            logging.error(f"Error appending row: {error}")
            return False
    
    def batch_append_rows(self, sheet_name: str, rows_data: List[List[Any]]) -> bool:
        """Append multiple rows at once for better performance"""
        try:
            body = {'values': rows_data}
            self.sheet.values().append(
                spreadsheetId=self.spreadsheet_id,
                range=f'{sheet_name}!A1',
                valueInputOption='RAW',
                insertDataOption='INSERT_ROWS',
                body=body
            ).execute()
            logging.info(f"Batch uploaded {len(rows_data)} rows to {sheet_name}")
            return True
        except HttpError as error:
            logging.error(f"Error batch appending rows: {error}")
            return False
    
    def trigger_schema_detection(self) -> bool:
        """Trigger the Apps Script schema detection after upload"""
        try:
            # The onEdit trigger should handle this automatically
            # But we log it for clarity
            logging.info("Data upload complete - Apps Script will auto-detect schema changes")
            return True
            
        except Exception as e:
            logging.error(f"Error triggering schema detection: {e}")
            return False

class ClinicalNoteProcessor:
    """Processes clinical notes through Gemini AI"""
    
    def __init__(self, api_key: str, model_name: str = "gemini-2.5-flash",  # BACK TO YOUR ORIGINAL
                sheets_manager: Optional[GoogleSheetsManager] = None):
        self.sheets_manager = sheets_manager
        self.model_name = model_name
        
        # Configure Gemini
        genai.configure(api_key=api_key)
        
        # Initialize model
        try:
            self.model = genai.GenerativeModel(model_name)
            logging.info(f"Successfully initialized model: {model_name}")
            print(f"✓ Using Gemini model: {model_name}\n")
        except Exception as e:
            logging.error(f"Failed to initialize {model_name}: {e}")
            raise
        
        # Ensure folders exist
        self._ensure_folders()
        
        # Processing stats
        self.stats = {
            'processed': 0,
            'failed': 0,
            'api_calls': 0,
            'sheets_updates': 0
        }
    
    def _ensure_folders(self):
        """Create necessary folders"""
        for folder in ['./input', './output', './processed', './errors']:
            Path(folder).mkdir(parents=True, exist_ok=True)
    
    def create_audit_prompt(self, note_text: str) -> str:
        """Create structured prompt for Gemini AI with source tracking"""
        return f"""
Analyze this clinical therapy note and extract audit information.

For each field, provide the SOURCE (specific text excerpt, section name, or line reference) where you found the information.

Return ONLY a valid JSON object with this exact structure:
{{
    "diagnosis_codes": ["list of ICD-10 codes found, e.g., F32.1"],
    "diagnosis_codes_source": "exact text excerpt or location where codes were found",
    
    "cpt_codes": ["list of CPT codes found, e.g., 90834"],
    "cpt_codes_source": "exact text excerpt or location where codes were found",
    
    "medical_necessity": "Yes/No/Unclear",
    "medical_necessity_source": "specific text excerpt demonstrating medical necessity",
    
    "treatment_goals": ["list of treatment goals mentioned"],
    "treatment_goals_source": "section or text excerpt containing goals",
    
    "interventions": ["list of interventions performed"],
    "interventions_source": "section or text excerpt describing interventions",
    
    "progress_noted": "Yes/No",
    "progress_noted_source": "text excerpt showing progress documentation",
    
    "plan_documented": "Yes/No",
    "plan_documented_source": "text excerpt showing treatment plan",
    
    "therapist_signature": "Present/Missing/Unclear",
    "therapist_signature_source": "location or text showing signature",
    
    "date_of_service": "extracted date or Not Found",
    "date_of_service_source": "location where date appears",
    
    "session_duration": "duration in minutes or Not Documented",
    "session_duration_source": "text excerpt showing duration",
    
    "compliance_issues": ["list of any compliance or documentation issues"],
    "addendum_needed": "Yes/No",
    "addendum_reason": "detailed reason if Yes, empty string if No",
    "confidence_score": 0.85
}}

CLINICAL NOTE:
{note_text}

Return ONLY the JSON object, no markdown formatting, no additional text.
"""
    
    def call_gemini_api(self, note_text: str, retry_count: int = 0) -> Optional[Dict]:
        """Call Gemini API with retry logic"""
        try:
            self.stats['api_calls'] += 1
            
            prompt = self.create_audit_prompt(note_text)
            
            # Generate content
            response = self.model.generate_content(prompt)
            
            # Extract text from response
            response_text = response.text.strip()
            
            # Clean markdown formatting
            if '```json' in response_text:
                response_text = response_text.split('```json')[1].split('```')[0].strip()
            elif '```' in response_text:
                parts = response_text.split('```')
                for part in parts:
                    if '{' in part and '}' in part:
                        response_text = part.strip()
                        break
            
            # Remove any remaining 'json' prefix
            if response_text.startswith('json'):
                response_text = response_text[4:].strip()
            
            # Parse JSON
            parsed = json.loads(response_text)
            
            # Ensure confidence score exists
            if 'confidence_score' not in parsed:
                parsed['confidence_score'] = 0.85
            
            # Ensure all source fields exist
            source_fields = [
                'diagnosis_codes_source', 'cpt_codes_source', 'medical_necessity_source',
                'treatment_goals_source', 'interventions_source', 'progress_noted_source',
                'plan_documented_source', 'therapist_signature_source',
                'date_of_service_source', 'session_duration_source'
            ]
            
            for field in source_fields:
                if field not in parsed:
                    parsed[field] = 'Not specified'
            
            logging.info(f"Successfully parsed AI response with sources")
            return parsed
            
        except json.JSONDecodeError as e:
            logging.error(f"JSON parse error (attempt {retry_count + 1}): {e}")
            logging.error(f"Response was: {response_text[:300]}")
            
            if retry_count < 2:
                time.sleep(2)
                return self.call_gemini_api(note_text, retry_count + 1)
            
        except Exception as e:
            logging.error(f"API error (attempt {retry_count + 1}): {e}")
            
            if retry_count < 2:
                time.sleep(2)
                return self.call_gemini_api(note_text, retry_count + 1)
        
        return None
    
    def flatten_to_row(self, claim_id: str, note_text: str, ai_response: Optional[Dict]) -> List[Any]:
        """Convert data to spreadsheet row format with AI sources"""
        row = [
            claim_id,
            note_text[:200] + '...' if len(note_text) > 200 else note_text,
            note_text,
            datetime.now().isoformat(),
            'Success' if ai_response else 'Failed'
        ]
        
        if ai_response:
            row.extend([
                '|'.join(ai_response.get('diagnosis_codes', [])),
                '|'.join(ai_response.get('cpt_codes', [])),
                ai_response.get('medical_necessity', 'Unknown'),
                '|'.join(ai_response.get('treatment_goals', [])),
                '|'.join(ai_response.get('interventions', [])),
                ai_response.get('progress_noted', 'Unknown'),
                ai_response.get('plan_documented', 'Unknown'),
                ai_response.get('therapist_signature', 'Unknown'),
                ai_response.get('date_of_service', 'Not Found'),
                str(ai_response.get('session_duration', 'Not Documented')),
                '|'.join(ai_response.get('compliance_issues', [])),
                ai_response.get('addendum_needed', 'Unknown'),
                ai_response.get('addendum_reason', ''),
                '', '', '', '', 'Pending', '', '', '',
                str(ai_response.get('confidence_score', 0.0)),
                # AI Source columns
                ai_response.get('diagnosis_codes_source', ''),
                ai_response.get('cpt_codes_source', ''),
                ai_response.get('medical_necessity_source', ''),
                ai_response.get('treatment_goals_source', ''),
                ai_response.get('interventions_source', ''),
                ai_response.get('progress_noted_source', ''),
                ai_response.get('plan_documented_source', ''),
                ai_response.get('therapist_signature_source', ''),
                ai_response.get('date_of_service_source', ''),
                ai_response.get('session_duration_source', '')
            ])
        else:
            row.extend([
                '', '', 'Error', '', '', 'Error', 'Error', 'Error',
                'Error', 'Error', 'Processing Failed', 'Yes',
                'Failed to process note with AI', '', '', '', '',
                'Error', '', '', '', '0.0',
                # Empty sources for failed processing
                '', '', '', '', '', '', '', '', '', ''
            ])
        
        return row
    
    def detect_csv_format(self, file_path: str) -> Tuple[pd.DataFrame, str, str]:
        """Auto-detect CSV format and columns"""
        df = pd.read_csv(file_path)
        logging.info(f"CSV has {len(df)} rows, columns: {df.columns.tolist()}")
        
        # Detect ID column
        id_patterns = ['id', 'claim', 'case', 'record', 'patient', 'encounter']
        id_col = None
        
        for col in df.columns:
            if any(pattern in col.lower() for pattern in id_patterns):
                id_col = col
                break
        
        if not id_col:
            df['generated_id'] = [f'AUDIT-{datetime.now().strftime("%Y%m%d")}-{i:04d}' 
                                  for i in range(1, len(df) + 1)]
            id_col = 'generated_id'
            logging.info(f"Generated ID column: {id_col}")
        else:
            logging.info(f"Using ID column: {id_col}")
        
        # Detect text column (largest average length)
        text_cols = df.select_dtypes(include=['object']).columns.tolist()
        text_col = None
        max_length = 0
        
        for col in text_cols:
            if col != id_col:
                avg_len = df[col].astype(str).str.len().mean()
                if avg_len > max_length:
                    max_length = avg_len
                    text_col = col
        
        if not text_col and text_cols:
            text_col = text_cols[0]
        
        logging.info(f"Using text column: {text_col} (avg length: {max_length:.0f})")
        
        return df, id_col, text_col
    
    def process_csv_file(self, file_path: str, batch_upload: bool = True) -> Tuple[bool, str]:
        """Process CSV file and update Google Sheets"""
        try:
            logging.info(f"Starting processing: {file_path}")
            
            df, id_col, text_col = self.detect_csv_format(file_path)
            
            if not text_col:
                raise ValueError("Could not identify text column")
            
            output_rows = []
            total = len(df)
            
            print(f"\nProcessing {total} records...")
            print("=" * 60)
            
            for idx, row in df.iterrows():
                claim_id = str(row[id_col])
                note_text = str(row[text_col])
                
                print(f"\n[{idx + 1}/{total}] {claim_id}")
                logging.info(f"Processing {idx + 1}/{total}: {claim_id}")
                
                ai_response = self.call_gemini_api(note_text)
                
                if ai_response:
                    self.stats['processed'] += 1
                    print(f"  ✓ Success - Confidence: {ai_response.get('confidence_score', 'N/A')}")
                    if ai_response.get('addendum_needed') == 'Yes':
                        print(f"    ⚠ Addendum needed: {ai_response.get('addendum_reason', '')[:60]}...")
                else:
                    self.stats['failed'] += 1
                    print(f"  ✗ Failed to process")
                
                row_data = self.flatten_to_row(claim_id, note_text, ai_response)
                output_rows.append(row_data)
                
                # Single row upload if not batching
                if not batch_upload and self.sheets_manager:
                    if self.sheets_manager.append_row('Data', row_data):
                        self.stats['sheets_updates'] += 1
                        print(f"  → Added to Google Sheets")
                
                # Rate limiting
                if idx < total - 1:
                    time.sleep(1)  # Adjust based on your quota
            
            print("\n" + "=" * 60)
            
            # Batch upload all rows at once
            if batch_upload and self.sheets_manager and output_rows:
                print(f"\nUploading {len(output_rows)} rows to Google Sheets...")
                if self.sheets_manager.batch_append_rows('Data', output_rows):
                    self.stats['sheets_updates'] = len(output_rows)
                    print(f"  ✓ Successfully uploaded all data")
                    
                    # Trigger schema detection
                    self.sheets_manager.trigger_schema_detection()
                    print(f"  → Schema detection will run automatically")
                else:
                    print(f"  ✗ Failed to upload data")
            
            # Save CSV backup
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_file = f"./output/processed_{timestamp}.csv"
            
            columns = [
                'claim_id', 'note_preview', 'full_note_text', 'processing_timestamp',
                'processing_status', 'diagnosis_codes', 'cpt_codes', 'medical_necessity',
                'treatment_goals', 'interventions', 'progress_noted', 'plan_documented',
                'therapist_signature', 'date_of_service', 'session_duration',
                'compliance_issues', 'addendum_needed', 'addendum_reason',
                'validator_decision', 'validator_notes', 'validation_timestamp',
                'validator_email', 'lead_review_status', 'final_addendum',
                'lead_reviewer', 'lead_review_timestamp', 'ai_confidence_score',
                # AI Source columns
                'ai_source_diagnosis_codes', 'ai_source_cpt_codes', 'ai_source_medical_necessity',
                'ai_source_treatment_goals', 'ai_source_interventions', 'ai_source_progress_noted',
                'ai_source_plan_documented', 'ai_source_therapist_signature',
                'ai_source_date_of_service', 'ai_source_session_duration'
            ]
            
            output_df = pd.DataFrame(output_rows, columns=columns)
            output_df.to_csv(output_file, index=False)
            
            # Move to processed folder
            processed_path = f"./processed/{Path(file_path).name}"
            Path(file_path).rename(processed_path)
            
            logging.info(f"Processing complete: {output_file}")
            
            return True, output_file
            
        except Exception as e:
            logging.error(f"Processing failed: {e}")
            
            error_file = f"./errors/error_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            with open(error_file, 'w') as f:
                f.write(f"Error processing {file_path}\n")
                f.write(f"Error: {str(e)}\n")
                f.write(f"Stats: {self.stats}\n")
            
            return False, str(e)

def load_config() -> dict:
    """Load configuration"""
    config_file = Path("config.json")
    
    if config_file.exists():
        with open(config_file) as f:
            return json.load(f)
    else:
        default = {
            "gemini_api_key": "YOUR_API_KEY_HERE",
            "gemini_model": "gemini-2.0-flash-exp",
            "google_sheets_enabled": True,
            "google_credentials_file": "credentials.json",
            "spreadsheet_id": "YOUR_SPREADSHEET_ID",
            "batch_upload": True,
            "rate_limit_seconds": 1
        }
        
        with open(config_file, 'w') as f:
            json.dump(default, f, indent=4)
        
        print("Created config.json - please update with your settings")
        sys.exit(1)

def list_available_models(api_key: str):
    """List all available Gemini models"""
    genai.configure(api_key=api_key)
    
    print("\n" + "="*60)
    print("AVAILABLE GEMINI MODELS FOR YOUR API KEY")
    print("="*60)
    
    try:
        models = genai.list_models()
        
        compatible_models = []
        
        for model in models:
            if 'generateContent' in model.supported_generation_methods:
                model_name = model.name.split('/')[-1]
                compatible_models.append(model_name)
                
                print(f"\n✓ {model_name}")
                print(f"  Display Name: {model.display_name}")
                if hasattr(model, 'input_token_limit'):
                    print(f"  Input Token Limit: {model.input_token_limit:,}")
                if hasattr(model, 'output_token_limit'):
                    print(f"  Output Token Limit: {model.output_token_limit:,}")
        
        print("\n" + "="*60)
        print(f"Found {len(compatible_models)} compatible model(s)")
        print("="*60)
        
        if compatible_models:
            print("\nCopy one of these model names to your config.json:")
            for model in compatible_models:
                print(f'  "gemini_model": "{model}"')
        
        return compatible_models
        
    except Exception as e:
        print(f"❌ Error listing models: {e}")
        return []


def main():
    """Main entry point"""
    print("""
    ╔═══════════════════════════════════════════╗
    ║    Clinical Audit Processing System       ║
    ║    Version 2.0 - Dynamic Schema Support   ║
    ║    Gemini 2.0 Flash / 2.5 Pro             ║
    ╚═══════════════════════════════════════════╝
    """)
    
    config = load_config()
    

    if len(sys.argv) > 1 and sys.argv[1] == '--list-models':
        list_available_models(config['gemini_api_key'])
        return

    # Initialize Google Sheets if enabled
    sheets_manager = None
    if config.get('google_sheets_enabled') and config.get('spreadsheet_id'):
        try:
            sheets_manager = GoogleSheetsManager(
                config['google_credentials_file'],
                config['spreadsheet_id']
            )
            
            sheets_manager.ensure_sheets_exist()
            sheets_manager.setup_data_sheet_headers()
            
            print("✓ Google Sheets connected")
            print("✓ Headers checked (existing headers preserved)\n")
            
        except Exception as e:
            print(f"⚠ Google Sheets connection failed: {e}")
            print("Continuing with CSV output only...\n")
    
    # Initialize processor
    model_name = config.get('gemini_model', 'gemini-2.0-flash-exp')
    processor = ClinicalNoteProcessor(
        config['gemini_api_key'],
        model_name,
        sheets_manager
    )
    
    # Check for input files
    csv_files = list(Path('./input').glob("*.csv"))
    
    if csv_files:
        print(f"Found {len(csv_files)} CSV file(s):")
        for f in csv_files:
            print(f"  - {f.name}")
        
        proceed = input("\nProcess all files? (y/n): ")
        
        if proceed.lower() == 'y':
            batch_upload = config.get('batch_upload', True)
            
            for csv_file in csv_files:
                print(f"\n{'='*60}")
                print(f"Processing: {csv_file.name}")
                print('='*60)
                
                success, result = processor.process_csv_file(str(csv_file), batch_upload)
                
                if success:
                    print(f"\n✓ Completed: {csv_file.name}")
                    print(f"  Output: {result}")
                else:
                    print(f"\n✗ Failed: {csv_file.name}")
                    print(f"  Error: {result}")
            
            print(f"\n{'='*60}")
            print("FINAL STATISTICS")
            print('='*60)
            print(f"Total Processed: {processor.stats['processed']}")
            print(f"Failed: {processor.stats['failed']}")
            print(f"API Calls: {processor.stats['api_calls']}")
            print(f"Sheets Updates: {processor.stats['sheets_updates']}")
            print('='*60)
            
            if sheets_manager:
                print("\n✓ Data uploaded to Google Sheets")
                print("✓ Apps Script will auto-detect schema changes")
                print("✓ Dropdowns will auto-refresh on next validation\n")
    else:
        print("No CSV files in ./input folder")
        print("Add CSV files and run again")

if __name__ == "__main__":
    main()