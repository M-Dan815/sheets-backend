"""
Clinical Audit System - Data Processing Backend
Version 3.0 - Schema-Aware with Scoring Support (1-5 and Binary Fields)
Updated with Gemini 2.5 models and Dynamic Schema Support
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

# ============================================================================
# SCHEMA MANAGEMENT - WITH SCORING SUPPORT
# ============================================================================

class SchemaManager:
    """Manages schema.json as single source of truth with scoring support"""
    
    def __init__(self, schema_path: str = 'schema.json'):
        self.schema_path = schema_path
        self.schema = self.load_schema()
        self.fields = self.schema['fields']
        self.validation_fields = self.schema.get('validation_fields', [])
        self.validation_field_config = self.schema.get('validation_field_config', {})
        self.core_fields = self.schema.get('core_fields', [])
        self.scoring_legend = self.schema.get('scoring_legend', {})
        
    def load_schema(self) -> dict:
        """Load schema from JSON file"""
        try:
            with open(self.schema_path, 'r') as f:
                schema = json.load(f)
            logging.info(f"✅ Schema loaded: v{schema.get('version', 'unknown')} ({len(schema['fields'])} fields)")
            
            # Log field type distribution
            if 'validation_field_config' in schema:
                scored = sum(1 for v in schema['validation_field_config'].values() if v.get('type') == 'score')
                binary = sum(1 for v in schema['validation_field_config'].values() if v.get('type') == 'binary')
                logging.info(f"  - Scored fields: {scored}, Binary fields: {binary}")
            
            return schema
        except FileNotFoundError:
            logging.warning(f"⚠️ Schema file not found: {self.schema_path}")
            return self.create_default_schema()
        except json.JSONDecodeError as e:
            logging.error(f"❌ Invalid JSON in schema file: {e}")
            raise
    
    def create_default_schema(self) -> dict:
        """Create default schema.json if it doesn't exist"""
        default_schema = {
            "version": "3.0",
            "last_updated": datetime.now().isoformat(),
            "description": "Clinical Audit System - Auto-generated Schema with Scoring",
            "fields": [
                "claim_id", "note_preview", "full_note_text", "processing_timestamp",
                "assigned_to_validator", "diagnosis_codes", "cpt_codes", "medical_necessity",
                "treatment_goals", "interventions", "progress_noted", "plan_documented",
                "therapist_signature", "date_of_service", "session_duration",
                "compliance_issues", "addendum_needed", "addendum_reason",
                "validator_decision", "validator_notes", "validation_timestamp",
                "validator_email", "assigned_to_lead", "lead_review_status", "final_addendum",
                "lead_reviewer", "lead_review_timestamp", "ai_confidence_score",
                "ai_source_diagnosis_codes", "ai_source_cpt_codes", "ai_source_medical_necessity",
                "ai_source_treatment_goals", "ai_source_interventions", "ai_source_progress_noted",
                "ai_source_plan_documented", "ai_source_therapist_signature",
                "ai_source_date_of_service", "ai_source_session_duration"
            ],
            "validation_fields": [
                "diagnosis_codes", "cpt_codes", "medical_necessity",
                "treatment_goals", "interventions", "progress_noted",
                "plan_documented", "therapist_signature",
                "date_of_service", "session_duration"
            ],
            "validation_field_config": {
                "diagnosis_codes": {
                    "type": "score",
                    "range": [1, 5],
                    "description": "Quality and accuracy of diagnosis coding (1=Poor, 5=Excellent)"
                },
                "cpt_codes": {
                    "type": "binary",
                    "options": ["Pass", "Fail"],
                    "description": "CPT codes present and accurate"
                },
                "medical_necessity": {
                    "type": "score",
                    "range": [1, 5],
                    "description": "Medical necessity justification quality (1=Poor, 5=Excellent)"
                },
                "treatment_goals": {
                    "type": "score",
                    "range": [1, 5],
                    "description": "Treatment goals clarity and measurability (1=Poor, 5=Excellent)"
                },
                "interventions": {
                    "type": "score",
                    "range": [1, 5],
                    "description": "Intervention documentation quality (1=Poor, 5=Excellent)"
                },
                "progress_noted": {
                    "type": "binary",
                    "options": ["Yes", "No"],
                    "description": "Progress documentation present"
                },
                "plan_documented": {
                    "type": "binary",
                    "options": ["Yes", "No"],
                    "description": "Treatment plan documented"
                },
                "therapist_signature": {
                    "type": "binary",
                    "options": ["Present", "Missing"],
                    "description": "Therapist signature present"
                },
                "date_of_service": {
                    "type": "binary",
                    "options": ["Present", "Missing"],
                    "description": "Date of service documented"
                },
                "session_duration": {
                    "type": "binary",
                    "options": ["Present", "Missing"],
                    "description": "Session duration documented"
                }
            },
            "core_fields": [
                "claim_id", "note_preview", "full_note_text", "processing_timestamp",
                "assigned_to_validator", "validator_decision", "validator_notes",
                "validation_timestamp", "validator_email", "assigned_to_lead",
                "lead_review_status", "final_addendum", "lead_reviewer",
                "lead_review_timestamp"
            ],
            "scoring_legend": {
                "1": "Poor - Major issues, requires significant correction",
                "2": "Below Average - Notable deficiencies, needs improvement",
                "3": "Average - Meets minimum requirements, minor issues",
                "4": "Good - Well documented, minor improvements possible",
                "5": "Excellent - Comprehensive and accurate documentation"
            }
        }
        
        # Save default schema
        with open(self.schema_path, 'w') as f:
            json.dump(default_schema, f, indent=2)
        
        logging.info(f"✅ Created default schema.json with scoring support")
        print(f"\n✅ Created schema.json with {len(default_schema['fields'])} fields")
        print(f"   - Scored fields: 5 (1-5 rating)")
        print(f"   - Binary fields: 5 (Pass/Fail or Yes/No)\n")
        
        return default_schema
    
    def get_field_index(self, field_name: str) -> Optional[int]:
        """Get index of a field in the schema"""
        try:
            return self.fields.index(field_name)
        except ValueError:
            return None
    
    def get_field_type(self, field_name: str) -> str:
        """Get field type (score or binary)"""
        if field_name in self.validation_field_config:
            return self.validation_field_config[field_name].get('type', 'binary')
        return 'binary'
    
    def get_field_options(self, field_name: str) -> list:
        """Get valid options for a field"""
        if field_name in self.validation_field_config:
            config = self.validation_field_config[field_name]
            if config.get('type') == 'score':
                range_vals = config.get('range', [1, 5])
                return list(range(range_vals[0], range_vals[1] + 1))
            else:
                return config.get('options', ['Pass', 'Fail'])
        return ['Pass', 'Fail']
    
    def validate_against_sheets(self, sheet_headers: List[str]) -> Tuple[bool, List[str], List[str]]:
        """Compare schema with Google Sheets headers"""
        schema_set = set(self.fields)
        sheets_set = set(sheet_headers)
        
        missing_in_sheets = list(schema_set - sheets_set)
        extra_in_sheets = list(sheets_set - schema_set)
        
        matches = (self.fields == sheet_headers)
        
        return matches, missing_in_sheets, extra_in_sheets

# ============================================================================
# GOOGLE SHEETS MANAGER
# ============================================================================

class GoogleSheetsManager:
    """Manages Google Sheets operations with schema awareness"""
    
    def __init__(self, credentials_file: str, spreadsheet_id: str, schema_manager: SchemaManager):
        """Initialize Google Sheets connection"""
        self.spreadsheet_id = spreadsheet_id
        self.schema_manager = schema_manager
        
        try:
            self.creds = service_account.Credentials.from_service_account_file(
                credentials_file,
                scopes=['https://www.googleapis.com/auth/spreadsheets']
            )
            self.service = build('sheets', 'v4', credentials=self.creds)
            self.sheet = self.service.spreadsheets()
            logging.info("✅ Google Sheets connection established")
            
        except FileNotFoundError:
            logging.error(f"❌ Credentials file not found: {credentials_file}")
            raise
        except Exception as e:
            logging.error(f"❌ Failed to connect to Google Sheets: {e}")
            raise
    
    def ensure_sheets_exist(self) -> bool:
        """Create required sheets if they don't exist"""
        try:
            result = self.sheet.get(spreadsheetId=self.spreadsheet_id).execute()
            existing_sheets = [s['properties']['title'] for s in result.get('sheets', [])]
            
            required_sheets = ['Data', 'Config', 'Audit Log', 'Schema Metadata']
            
            requests = []
            for sheet_name in required_sheets:
                if sheet_name not in existing_sheets:
                    requests.append({
                        'addSheet': {
                            'properties': {
                                'title': sheet_name,
                                'gridProperties': {
                                    'rowCount': 1000,
                                    'columnCount': len(self.schema_manager.fields) + 10
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
                logging.info(f"✅ Created sheets: {[s for s in required_sheets if s not in existing_sheets]}")
            
            return True
            
        except HttpError as error:
            logging.error(f"❌ Error creating sheets: {error}")
            return False
    
    def get_existing_headers(self) -> List[str]:
        """Get current headers from Data sheet"""
        try:
            result = self.sheet.values().get(
                spreadsheetId=self.spreadsheet_id,
                range='Data!A1:ZZ1'
            ).execute()
            
            existing_values = result.get('values', [])
            
            if existing_values and existing_values[0]:
                return existing_values[0]
            
            return []
            
        except HttpError:
            return []
    
    def setup_data_sheet_headers(self, force_update: bool = False) -> bool:
        """Set up headers using schema.json"""
        
        existing_headers = self.get_existing_headers()
        
        if existing_headers and not force_update:
            # Validate schema match
            matches, missing, extra = self.schema_manager.validate_against_sheets(existing_headers)
            
            if matches:
                logging.info(f"✅ Headers match schema.json ({len(existing_headers)} columns)")
                return True
            
            # Schema mismatch - migrate
            logging.warning(f"⚠️ Schema mismatch detected!")
            logging.warning(f"  Missing in Sheets: {missing}")
            logging.warning(f"  Extra in Sheets: {extra}")
            
            return self.migrate_schema(existing_headers)
        
        # No headers exist - create from schema
        try:
            headers = [self.schema_manager.fields]
            
            body = {'values': headers}
            self.sheet.values().update(
                spreadsheetId=self.spreadsheet_id,
                range='Data!A1',
                valueInputOption='RAW',
                body=body
            ).execute()
            
            logging.info(f"✅ Data sheet headers created from schema.json ({len(headers[0])} columns)")
            return True
            
        except HttpError as error:
            logging.error(f"❌ Error setting up headers: {error}")
            return False
    
    def migrate_schema(self, old_headers: List[str]) -> bool:
        """Migrate existing data to new schema"""
        try:
            logging.info("🔄 Migrating data to new schema...")
            
            # Read all existing data
            result = self.sheet.values().get(
                spreadsheetId=self.spreadsheet_id,
                range='Data!A:ZZ'
            ).execute()
            
            all_data = result.get('values', [])
            
            if len(all_data) <= 1:
                # Only headers, no data to migrate
                return self.setup_data_sheet_headers(force_update=True)
            
            existing_data = all_data[1:]  # Skip header row
            
            # Create mapping from old to new
            migrated_data = []
            
            for old_row in existing_data:
                new_row = [''] * len(self.schema_manager.fields)
                
                for old_idx, old_header in enumerate(old_headers):
                    if old_header in self.schema_manager.fields:
                        new_idx = self.schema_manager.fields.index(old_header)
                        if old_idx < len(old_row):
                            new_row[new_idx] = old_row[old_idx]
                
                migrated_data.append(new_row)
            
            # Clear sheet
            self.sheet.values().clear(
                spreadsheetId=self.spreadsheet_id,
                range='Data!A:ZZ'
            ).execute()
            
            # Write new headers
            headers = [self.schema_manager.fields]
            body = {'values': headers}
            self.sheet.values().update(
                spreadsheetId=self.spreadsheet_id,
                range='Data!A1',
                valueInputOption='RAW',
                body=body
            ).execute()
            
            # Write migrated data
            if migrated_data:
                body = {'values': migrated_data}
                self.sheet.values().append(
                    spreadsheetId=self.spreadsheet_id,
                    range='Data!A2',
                    valueInputOption='RAW',
                    insertDataOption='INSERT_ROWS',
                    body=body
                ).execute()
            
            logging.info(f"✅ Schema migrated: {len(migrated_data)} rows preserved")
            print(f"\n✅ Schema migrated successfully!")
            print(f"  - Preserved: {len(migrated_data)} rows")
            print(f"  - New schema: {len(self.schema_manager.fields)} columns")
            
            return True
            
        except Exception as e:
            logging.error(f"❌ Schema migration failed: {e}")
            return False
    
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
            logging.error(f"❌ Error appending row: {error}")
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
            logging.info(f"✅ Batch uploaded {len(rows_data)} rows to {sheet_name}")
            return True
        except HttpError as error:
            logging.error(f"❌ Error batch appending rows: {error}")
            return False
    
    def get_existing_claim_ids(self) -> set:
        """Get set of existing claim IDs to avoid duplicates"""
        try:
            claim_id_idx = self.schema_manager.get_field_index('claim_id')
            if claim_id_idx is None:
                return set()
            
            col_letter = self._index_to_column_letter(claim_id_idx + 1)
            
            result = self.sheet.values().get(
                spreadsheetId=self.spreadsheet_id,
                range=f'Data!{col_letter}2:{col_letter}'
            ).execute()
            
            values = result.get('values', [])
            return set([row[0] for row in values if row])
            
        except Exception as e:
            logging.warning(f"⚠️ Could not fetch existing claim IDs: {e}")
            return set()
    
    def _index_to_column_letter(self, col_idx: int) -> str:
        """Convert column index to letter (1 -> A, 27 -> AA)"""
        letter = ''
        while col_idx > 0:
            col_idx -= 1
            letter = chr(col_idx % 26 + 65) + letter
            col_idx //= 26
        return letter
    
    def trigger_schema_detection(self) -> bool:
        """Notify that schema detection should run"""
        try:
            logging.info("✅ Data upload complete - Apps Script will auto-detect schema")
            return True
        except Exception as e:
            logging.error(f"❌ Error triggering schema detection: {e}")
            return False

# ============================================================================
# CLINICAL NOTE PROCESSOR - WITH SCORING SUPPORT
# ============================================================================

class ClinicalNoteProcessor:
    """Processes clinical notes through Gemini AI with scoring support"""
    
    def __init__(self, api_key: str, model_name: str = "gemini-2.0-flash-exp",
                 schema_manager: SchemaManager = None,
                 sheets_manager: Optional[GoogleSheetsManager] = None):
        self.schema_manager = schema_manager
        self.sheets_manager = sheets_manager
        self.model_name = model_name
        
        # Configure Gemini
        genai.configure(api_key=api_key)
        
        # Initialize model
        try:
            self.model = genai.GenerativeModel(model_name)
            logging.info(f"✅ Successfully initialized model: {model_name}")
            print(f"✅ Using Gemini model: {model_name}\n")
        except Exception as e:
            logging.error(f"❌ Failed to initialize {model_name}: {e}")
            raise
        
        # Ensure folders exist
        self._ensure_folders()
        
        # Processing stats
        self.stats = {
            'processed': 0,
            'failed': 0,
            'skipped_duplicates': 0,
            'api_calls': 0,
            'sheets_updates': 0
        }
    
    def _ensure_folders(self):
        """Create necessary folders"""
        for folder in ['./input', './output', './processed', './errors']:
            Path(folder).mkdir(parents=True, exist_ok=True)
    
    def create_audit_prompt(self, note_text: str) -> str:
        """Create structured prompt for Gemini AI with source tracking and scoring"""
        
        # Build validation fields list from schema
        validation_fields = self.schema_manager.validation_fields if self.schema_manager else []
        
        # Build field examples based on type
        fields_json = {}
        scoring_instructions = []
        
        for field in validation_fields:
            field_type = self.schema_manager.get_field_type(field)
            field_config = self.schema_manager.validation_field_config.get(field, {})
            
            if field_type == 'score':
                range_vals = field_config.get('range', [1, 5])
                fields_json[field] = f"score from {range_vals[0]} to {range_vals[1]}"
                scoring_instructions.append(
                    f"  - {field}: Rate {range_vals[0]}-{range_vals[1]} ({field_config.get('description', '')})"
                )
            else:
                options = field_config.get('options', ['Yes', 'No'])
                fields_json[field] = f"one of: {', '.join(options)}"
                scoring_instructions.append(
                    f"  - {field}: Choose from: {', '.join(options)}"
                )
            
            fields_json[f"{field}_source"] = "exact text excerpt showing where found"
        
        fields_json.update({
            "compliance_issues": ["list of any issues"],
            "addendum_needed": "Yes/No",
            "addendum_reason": "detailed reason if Yes, empty if No",
            "confidence_score": 0.85
        })
        
        scoring_legend = "\n".join([f"{k}: {v}" for k, v in self.schema_manager.scoring_legend.items()])
        
        return f"""
Analyze this clinical therapy note and extract audit information with SCORING.

SCORING LEGEND (for scored fields):
{scoring_legend}

FIELD REQUIREMENTS:
{chr(10).join(scoring_instructions)}

For each field:
1. Provide the extracted VALUE (score or binary choice)
2. Provide the SOURCE (specific text excerpt showing where found)

Return ONLY a valid JSON object with this structure:
{json.dumps(fields_json, indent=2)}

SCORING GUIDELINES:
- Score 1-2: Document has major deficiencies, requires significant correction
- Score 3: Meets minimum requirements but has notable gaps
- Score 4-5: Well documented, comprehensive information

CLINICAL NOTE:
{note_text}

Return ONLY the JSON object, no markdown formatting, no additional text.
"""
    
    def validate_ai_response(self, ai_response: Dict) -> Dict:
        """Validate AI response values against schema rules"""
        
        for field in self.schema_manager.validation_fields:
            if field not in ai_response:
                continue
            
            field_type = self.schema_manager.get_field_type(field)
            field_value = ai_response[field]
            
            if field_type == 'score':
                # Validate score is within range
                config = self.schema_manager.validation_field_config.get(field, {})
                range_vals = config.get('range', [1, 5])
                
                try:
                    score = int(field_value) if not isinstance(field_value, int) else field_value
                    if score < range_vals[0] or score > range_vals[1]:
                        logging.warning(f"⚠️ Score {score} for {field} outside range {range_vals}, defaulting to 3")
                        ai_response[field] = 3
                    else:
                        ai_response[field] = score
                except (ValueError, TypeError):
                    logging.warning(f"⚠️ Invalid score for {field}: {field_value}, defaulting to 3")
                    ai_response[field] = 3
            
            else:
                # Validate binary choice is in allowed options
                options = self.schema_manager.get_field_options(field)
                if field_value not in options:
                    logging.warning(f"⚠️ Value '{field_value}' for {field} not in {options}, defaulting to {options[0]}")
                    ai_response[field] = options[0]
        
        return ai_response
    
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
            if self.schema_manager:
                for field in self.schema_manager.validation_fields:
                    source_field = f'{field}_source'
                    if source_field not in parsed:
                        parsed[source_field] = 'Not specified'
            
            # VALIDATE SCORES AND BINARY VALUES
            parsed = self.validate_ai_response(parsed)
            
            logging.info(f"✅ Successfully parsed and validated AI response")
            return parsed
            
        except json.JSONDecodeError as e:
            logging.error(f"❌ JSON parse error (attempt {retry_count + 1}): {e}")
            logging.error(f"Response was: {response_text[:300]}")
            
            if retry_count < 2:
                time.sleep(2)
                return self.call_gemini_api(note_text, retry_count + 1)
            
        except Exception as e:
            logging.error(f"❌ API error (attempt {retry_count + 1}): {e}")
            
            if retry_count < 2:
                time.sleep(2)
                return self.call_gemini_api(note_text, retry_count + 1)
        
        return None
    
    def flatten_to_row(self, claim_id: str, note_text: str, ai_response: Optional[Dict], 
                      assigned_to: Optional[str] = None) -> List[Any]:
        """Convert data to spreadsheet row format using schema.json field order"""
        
        # Build a dictionary with all values
        row_data = {}
        
        # Core fields
        row_data['claim_id'] = claim_id
        row_data['note_preview'] = note_text[:200] + '...' if len(note_text) > 200 else note_text
        row_data['full_note_text'] = note_text
        row_data['processing_timestamp'] = datetime.now().isoformat()
        row_data['assigned_to_validator'] = assigned_to or ''
        
        if ai_response:
            # Extract all fields from AI response
            for field in self.schema_manager.validation_fields:
                value = ai_response.get(field, '')
                # Convert lists to pipe-separated strings
                if isinstance(value, list):
                    value = '|'.join(str(v) for v in value)
                row_data[field] = value
                
                # AI source fields
                source_field = f'ai_source_{field}'
                row_data[source_field] = ai_response.get(source_field, '')
            
            # Metadata
            row_data['ai_confidence_score'] = str(ai_response.get('confidence_score', 0.0))
            row_data['compliance_issues'] = '|'.join(ai_response.get('compliance_issues', []))
            row_data['addendum_needed'] = ai_response.get('addendum_needed', 'Unknown')
            row_data['addendum_reason'] = ai_response.get('addendum_reason', '')
            
            # Workflow fields (empty for now)
            row_data['validator_decision'] = ''
            row_data['validator_notes'] = ''
            row_data['validation_timestamp'] = ''
            row_data['validator_email'] = ''
            row_data['assigned_to_lead'] = ''
            row_data['lead_review_status'] = 'Pending'
            row_data['final_addendum'] = ''
            row_data['lead_reviewer'] = ''
            row_data['lead_review_timestamp'] = ''
            
        else:
            # Failed processing - fill with error markers
            for field in self.schema_manager.validation_fields:
                row_data[field] = 'Error'
                row_data[f'ai_source_{field}'] = ''
            
            row_data['ai_confidence_score'] = '0.0'
            row_data['compliance_issues'] = 'Processing Failed'
            row_data['addendum_needed'] = 'Yes'
            row_data['addendum_reason'] = 'Failed to process note with AI'
            row_data['validator_decision'] = ''
            row_data['validator_notes'] = ''
            row_data['validation_timestamp'] = ''
            row_data['validator_email'] = ''
            row_data['assigned_to_lead'] = ''
            row_data['lead_review_status'] = 'Error'
            row_data['final_addendum'] = ''
            row_data['lead_reviewer'] = ''
            row_data['lead_review_timestamp'] = ''
        
        # Convert to ordered list matching schema
        row = [row_data.get(field, '') for field in self.schema_manager.fields]
        
        return row
    
    def detect_csv_format(self, file_path: str) -> Tuple[pd.DataFrame, str, str]:
        """Auto-detect CSV format and columns"""
        df = pd.read_csv(file_path)
        logging.info(f"📄 CSV has {len(df)} rows, columns: {df.columns.tolist()}")
        
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
            logging.info(f"✅ Generated ID column: {id_col}")
        else:
            logging.info(f"✅ Using ID column: {id_col}")
        
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
        
        logging.info(f"✅ Using text column: {text_col} (avg length: {max_length:.0f})")
        
        return df, id_col, text_col
    
    def process_csv_file(self, file_path: str, batch_upload: bool = True) -> Tuple[bool, str]:
        """Process CSV file and update Google Sheets"""
        try:
            logging.info(f"🚀 Starting processing: {file_path}")
            
            df, id_col, text_col = self.detect_csv_format(file_path)
            
            if not text_col:
                raise ValueError("Could not identify text column")
            
            # Get existing claim IDs to avoid duplicates
            existing_claim_ids = set()
            if self.sheets_manager:
                existing_claim_ids = self.sheets_manager.get_existing_claim_ids()
                if existing_claim_ids:
                    logging.info(f"📋 Found {len(existing_claim_ids)} existing claims in sheet")
            
            output_rows = []
            total = len(df)
            
            print(f"\n{'='*60}")
            print(f"Processing {total} records...")
            print('='*60)
            
            for idx, row in df.iterrows():
                claim_id = str(row[id_col])
                note_text = str(row[text_col])
                
                # Check for duplicates
                if claim_id in existing_claim_ids:
                    self.stats['skipped_duplicates'] += 1
                    print(f"\n[{idx + 1}/{total}] {claim_id}")
                    print(f"  ⏭️  Skipped (duplicate)")
                    logging.info(f"Skipped duplicate: {claim_id}")
                    continue
                
                print(f"\n[{idx + 1}/{total}] {claim_id}")
                logging.info(f"Processing {idx + 1}/{total}: {claim_id}")
                
                ai_response = self.call_gemini_api(note_text)
                
                if ai_response:
                    self.stats['processed'] += 1
                    print(f"  ✅ Success - Confidence: {ai_response.get('confidence_score', 'N/A')}")
                    
                    # Show scores for scored fields
                    scored_fields = [f for f in self.schema_manager.validation_fields 
                                   if self.schema_manager.get_field_type(f) == 'score']
                    
                    if scored_fields:
                        print(f"    📊 Scores:")
                        for field in scored_fields[:3]:  # Show first 3
                            score = ai_response.get(field, 'N/A')
                            print(f"       - {field.replace('_', ' ').title()}: {score}/5")
                    
                    if ai_response.get('addendum_needed') == 'Yes':
                        print(f"    ⚠️  Addendum needed: {ai_response.get('addendum_reason', '')[:60]}...")
                else:
                    self.stats['failed'] += 1
                    print(f"  ❌ Failed to process")
                
                row_data = self.flatten_to_row(claim_id, note_text, ai_response)
                output_rows.append(row_data)
                
                # Single row upload if not batching
                if not batch_upload and self.sheets_manager:
                    if self.sheets_manager.append_row('Data', row_data):
                        self.stats['sheets_updates'] += 1
                        print(f"  → Added to Google Sheets")
                
                # Rate limiting
                if idx < total - 1:
                    time.sleep(1)
            
            print("\n" + "="*60)
            
            # Batch upload all rows at once
            if batch_upload and self.sheets_manager and output_rows:
                print(f"\n📤 Uploading {len(output_rows)} rows to Google Sheets...")
                if self.sheets_manager.batch_append_rows('Data', output_rows):
                    self.stats['sheets_updates'] = len(output_rows)
                    print(f"  ✅ Successfully uploaded all data")
                    
                    # Trigger schema detection
                    self.sheets_manager.trigger_schema_detection()
                    print(f"  → Schema auto-detection triggered in Apps Script")
                else:
                    print(f"  ❌ Failed to upload data")
            
            # Save CSV backup
            if output_rows:
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                output_file = f"./output/processed_{timestamp}.csv"
                
                output_df = pd.DataFrame(output_rows, columns=self.schema_manager.fields)
                output_df.to_csv(output_file, index=False)
                
                logging.info(f"✅ Saved backup: {output_file}")
            
            # Move to processed folder
            processed_path = f"./processed/{Path(file_path).name}"
            Path(file_path).rename(processed_path)
            
            logging.info(f"✅ Processing complete")
            
            return True, output_file if output_rows else "No new records processed"
            
        except Exception as e:
            logging.error(f"❌ Processing failed: {e}")
            
            error_file = f"./errors/error_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            with open(error_file, 'w') as f:
                f.write(f"Error processing {file_path}\n")
                f.write(f"Error: {str(e)}\n")
                f.write(f"Stats: {self.stats}\n")
            
            return False, str(e)

# ============================================================================
# CONFIGURATION & UTILITIES
# ============================================================================

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
            "schema_file": "schema.json",
            "batch_upload": True,
            "rate_limit_seconds": 1
        }
        
        with open(config_file, 'w') as f:
            json.dump(default, f, indent=4)
        
        print("✅ Created config.json - please update with your settings")
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

def export_schema_json(schema_manager: SchemaManager):
    """Print schema.json for manual Google Sheets import"""
    print("\n" + "="*60)
    print("SCHEMA.JSON FOR GOOGLE SHEETS IMPORT")
    print("="*60)
    print("\nCopy the JSON below and paste into Google Sheets:")
    print("🔧 Audit System > 📥 Import Schema (JSON)")
    print("\n" + "-"*60)
    print(json.dumps(schema_manager.schema, indent=2))
    print("-"*60 + "\n")

# ============================================================================
# MAIN ENTRY POINT
# ============================================================================

def main():
    """Main entry point"""
    print("""
    ╔═══════════════════════════════════════════╗
    ║    Clinical Audit Processing System       ║
    ║    Version 3.0 - Schema-Aware + Scoring   ║
    ║    Gemini AI + schema.json Integration    ║
    ╚═══════════════════════════════════════════╝
    """)
    
    config = load_config()
    
    # Handle special commands
    if len(sys.argv) > 1:
        if sys.argv[1] == '--list-models':
            list_available_models(config['gemini_api_key'])
            return
        elif sys.argv[1] == '--export-schema':
            schema_manager = SchemaManager(config.get('schema_file', 'schema.json'))
            export_schema_json(schema_manager)
            return
        elif sys.argv[1] == '--sync-schema':
            print("\n🔄 Schema Synchronization Mode (No CSV Processing)\n")
            schema_manager = SchemaManager(config.get('schema_file', 'schema.json'))
            
            if config.get('google_sheets_enabled'):
                sheets_manager = GoogleSheetsManager(
                    config['google_credentials_file'],
                    config['spreadsheet_id'],
                    schema_manager
                )
                sheets_manager.ensure_sheets_exist()
                sheets_manager.setup_data_sheet_headers()
                print("\n✅ Schema synchronized with Google Sheets\n")
            else:
                print("❌ Google Sheets not enabled in config.json\n")
            return
    
    # ========================================================================
    # NORMAL PROCESSING MODE - WITH SCHEMA AUTO-DETECTION
    # ========================================================================
    
    print("\n" + "="*60)
    print("STEP 1: LOADING SCHEMA")
    print("="*60)
    
    # Load schema
    schema_manager = SchemaManager(config.get('schema_file', 'schema.json'))
    print(f"✅ Loaded schema.json")
    print(f"  - Version: {schema_manager.schema.get('version', 'unknown')}")
    print(f"  - Total fields: {len(schema_manager.fields)}")
    print(f"  - Validation fields: {len(schema_manager.validation_fields)}")
    
    # Count scored vs binary fields
    scored = sum(1 for v in schema_manager.validation_field_config.values() if v.get('type') == 'score')
    binary = sum(1 for v in schema_manager.validation_field_config.values() if v.get('type') == 'binary')
    print(f"  - Scored fields (1-5): {scored}")
    print(f"  - Binary fields: {binary}")
    
    # Initialize Google Sheets if enabled
    sheets_manager = None
    if config.get('google_sheets_enabled') and config.get('spreadsheet_id'):
        try:
            print("\n" + "="*60)
            print("STEP 2: CONNECTING TO GOOGLE SHEETS")
            print("="*60)
            
            sheets_manager = GoogleSheetsManager(
                config['google_credentials_file'],
                config['spreadsheet_id'],
                schema_manager
            )
            
            print("✅ Google Sheets connected")
            
            # This is where schema validation and auto-migration happens
            print("\n🔍 Validating schema against Google Sheets...")
            
            existing_headers = sheets_manager.get_existing_headers()
            
            if existing_headers:
                matches, missing, extra = schema_manager.validate_against_sheets(existing_headers)
                
                if matches:
                    print(f"✅ Schema matches perfectly!")
                    print(f"  - {len(existing_headers)} columns in sync")
                else:
                    print(f"⚠️  Schema mismatch detected!")
                    if missing:
                        print(f"  - Missing in Sheets: {missing[:3]}{'...' if len(missing) > 3 else ''}")
                    if extra:
                        print(f"  - Extra in Sheets: {extra[:3]}{'...' if len(extra) > 3 else ''}")
                    print(f"\n🔄 Auto-migrating schema...")
                    
                    # This auto-migrates
                    sheets_manager.setup_data_sheet_headers()
                    
                    print(f"✅ Schema migration complete!")
            else:
                print("📝 No existing headers - initializing from schema.json...")
                sheets_manager.ensure_sheets_exist()
                sheets_manager.setup_data_sheet_headers()
                print(f"✅ Headers created ({len(schema_manager.fields)} columns)")
            
        except Exception as e:
            print(f"\n⚠️  Google Sheets connection failed: {e}")
            print("Continuing with CSV output only...\n")
    else:
        print("\n⚠️  Google Sheets disabled in config.json")
        print("Will save to CSV files only\n")
    
    print("\n" + "="*60)
    print("STEP 3: PROCESSING CSV FILES")
    print("="*60)
    
    # Initialize processor
    model_name = config.get('gemini_model', 'gemini-2.0-flash-exp')
    processor = ClinicalNoteProcessor(
        config['gemini_api_key'],
        model_name,
        schema_manager,
        sheets_manager
    )
    
    # Check for input files
    csv_files = list(Path('./input').glob("*.csv"))
    
    if csv_files:
        print(f"\nFound {len(csv_files)} CSV file(s):")
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
                    print(f"\n✅ Completed: {csv_file.name}")
                    print(f"  Output: {result}")
                else:
                    print(f"\n❌ Failed: {csv_file.name}")
                    print(f"  Error: {result}")
            
            print(f"\n{'='*60}")
            print("FINAL STATISTICS")
            print('='*60)
            print(f"Total Processed: {processor.stats['processed']}")
            print(f"Failed: {processor.stats['failed']}")
            print(f"Skipped (Duplicates): {processor.stats['skipped_duplicates']}")
            print(f"API Calls: {processor.stats['api_calls']}")
            print(f"Sheets Updates: {processor.stats['sheets_updates']}")
            print('='*60)
            
            if sheets_manager:
                print("\n✅ Data uploaded to Google Sheets")
                print("✅ Apps Script will auto-detect schema changes")
                print("✅ User tabs will auto-refresh\n")
    else:
        print("\n❌ No CSV files in ./input folder")
        print("Add CSV files and run again\n")
        
        print("Available Commands:")
        print("="*60)
        print("  python process_claims.py")
        print("    → Process CSV files (auto-detects schema changes)")
        print("")
        print("  python process_claims.py --sync-schema")
        print("    → Sync schema.json to Sheets (no CSV processing)")
        print("")
        print("  python process_claims.py --export-schema")
        print("    → Show schema.json for manual import")
        print("")
        print("  python process_claims.py --list-models")
        print("    → List available Gemini models")
        print("="*60 + "\n")

if __name__ == "__main__":
    main()