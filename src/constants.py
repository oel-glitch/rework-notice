"""
Application constants and configuration.
Centralized location for all hardcoded values and magic numbers.
"""

from pathlib import Path
from typing import Final

# Application metadata
APP_NAME: Final[str] = "PDF Parser OATI"
APP_VERSION: Final[str] = "3.0.0"
APP_AUTHOR: Final[str] = "OATI Moscow"

# Directory paths
OUTPUT_DIR: Final[Path] = Path("output")
DATABASE_DIR: Final[Path] = Path("database")
TEMPLATES_DIR: Final[Path] = Path("templates")
CONFIG_DIR: Final[Path] = Path("config")

# Database
DATABASE_FILE: Final[str] = "app.db"
DATABASE_PATH: Final[Path] = DATABASE_DIR / DATABASE_FILE

# File extensions
PDF_EXTENSION: Final[str] = ".pdf"
DOCX_EXTENSION: Final[str] = ".docx"
JSON_EXTENSION: Final[str] = ".json"

# Encoding
DEFAULT_ENCODING: Final[str] = "utf-8"

# Date formats
DATE_FORMAT_FULL: Final[str] = "%d.%m.%Y %H:%M:%S"
DATE_FORMAT_SHORT: Final[str] = "%d.%m.%Y"

# Logging
LOG_FORMAT: Final[str] = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
LOG_DATE_FORMAT: Final[str] = "%Y-%m-%d %H:%M:%S"

# PDF Processing
PDF_TEXT_ENCODING: Final[str] = "utf-8"
MAX_PDF_SIZE_MB: Final[int] = 50

# Name declension - Russian grammatical cases
DATIVE_CASE: Final[str] = "datv"
NOMINATIVE_CASE: Final[str] = "nomn"
ACCUSATIVE_CASE: Final[str] = "accs"
GENITIVE_CASE: Final[str] = "gent"
INSTRUMENTAL_CASE: Final[str] = "ablt"
PREPOSITIONAL_CASE: Final[str] = "loct"

# Gender - Pymorphy3 tags
GENDER_MALE: Final[str] = "masc"
GENDER_FEMALE: Final[str] = "femn"

# Gender - Human-readable values
GENDER_MALE_VALUE: Final[str] = "male"
GENDER_FEMALE_VALUE: Final[str] = "female"

# Name type tags for pymorphy3
TAG_SURNAME: Final[str] = "Surn"
TAG_FIRST_NAME: Final[str] = "Name"
TAG_PATRONYMIC: Final[str] = "Patr"
TAG_NOUN: Final[str] = "NOUN"
TAG_ADJECTIVE: Final[str] = "ADJF"

# UI Messages
MSG_NO_FILES_SELECTED: Final[str] = "No PDF files selected for processing"
MSG_PROCESSING_STARTED: Final[str] = "Processing started"
MSG_PROCESSING_COMPLETED: Final[str] = "Processing completed successfully"
MSG_PROCESSING_ERROR: Final[str] = "Error during processing"
MSG_DATABASE_ERROR: Final[str] = "Database operation failed"
MSG_TEMPLATE_NOT_FOUND: Final[str] = "Template file not found"

# Status prefixes for logging
STATUS_INFO: Final[str] = "[INFO]"
STATUS_SUCCESS: Final[str] = "[SUCCESS]"
STATUS_WARNING: Final[str] = "[WARNING]"
STATUS_ERROR: Final[str] = "[ERROR]"

# Thread names
THREAD_PDF_PROCESSING: Final[str] = "PDFProcessingThread"

# Window geometry
WINDOW_MIN_WIDTH: Final[int] = 800
WINDOW_MIN_HEIGHT: Final[int] = 600

# Database table names
TABLE_DEPARTMENTS: Final[str] = "departments"
TABLE_DIRECTORS: Final[str] = "directors"
TABLE_INSPECTORS: Final[str] = "inspectors"
TABLE_INSPECTOR_CHIEFS: Final[str] = "inspectors_chiefs"
TABLE_PROCESSING_HISTORY: Final[str] = "processing_history"
TABLE_PERSONS: Final[str] = "persons"
TABLE_NAME_OVERRIDES: Final[str] = "name_overrides"
TABLE_MANUAL_DECLENSIONS: Final[str] = "manual_declensions"

# Database column names - Common
COL_ID: Final[str] = "id"
COL_IS_ACTIVE: Final[str] = "is_active"
COL_LAST_NAME: Final[str] = "last_name"
COL_FIRST_NAME: Final[str] = "first_name"
COL_MIDDLE_NAME: Final[str] = "middle_name"
COL_ROLE: Final[str] = "role"

# Database column names - Departments
COL_NAME: Final[str] = "name"
COL_SHORT_NAME: Final[str] = "short_name"
COL_INN: Final[str] = "inn"
COL_OGRN: Final[str] = "ogrn"
COL_ADDRESS: Final[str] = "address"

# Database column names - Directors
COL_DEPARTMENT: Final[str] = "department"

# Database column names - Inspector Chiefs
COL_LAST_NAME_NOMINATIVE: Final[str] = "last_name_nominative"
COL_LAST_NAME_DATIVE: Final[str] = "last_name_dative"
COL_FIRST_INITIAL: Final[str] = "first_initial"
COL_MIDDLE_INITIAL: Final[str] = "middle_initial"
COL_ORGANIZATION: Final[str] = "organization"

# Database column names - Processing History
COL_PDF_FILE: Final[str] = "pdf_file"
COL_CITIZEN_NAME: Final[str] = "citizen_name"
COL_OATI_NUMBER: Final[str] = "oati_number"
COL_PORTAL_ID: Final[str] = "portal_id"
COL_PROCESSED_DATE: Final[str] = "processed_date"
COL_OUTPUT_FILE: Final[str] = "output_file"
COL_STATUS: Final[str] = "status"

# Database column names - Name Overrides
COL_WORD_NOMINATIVE: Final[str] = "word_nominative"
COL_WORD_VALUE: Final[str] = "word_value"
COL_WORD_TYPE: Final[str] = "word_type"
COL_GENDER: Final[str] = "gender"
COL_CASE: Final[str] = "grammatical_case"

# Database column names - Manual Declensions
COL_PERSON_ID: Final[str] = "person_id"
COL_CASE_TYPE: Final[str] = "case_type"
COL_CUSTOM_VALUE: Final[str] = "custom_value"

# Database default values
DEFAULT_ACTIVE: Final[int] = 1
DEFAULT_INACTIVE: Final[int] = 0
DEFAULT_STATUS_SUCCESS: Final[str] = "success"
DEFAULT_HISTORY_LIMIT: Final[int] = 100

# Person roles
ROLE_DIRECTOR: Final[str] = "director"
ROLE_INSPECTOR: Final[str] = "inspector"
