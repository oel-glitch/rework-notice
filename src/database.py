"""
Database management module for PDF Parser OATI application.

This module provides a comprehensive interface for managing SQLite database operations,
including departments, directors, inspectors, inspector chiefs, and processing history.
"""

import json
import logging
import os
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from src.constants import (
    COL_ADDRESS,
    COL_CASE,
    COL_CASE_TYPE,
    COL_CITIZEN_NAME,
    COL_CUSTOM_VALUE,
    COL_DEPARTMENT,
    COL_FIRST_INITIAL,
    COL_FIRST_NAME,
    COL_GENDER,
    COL_ID,
    COL_INN,
    COL_IS_ACTIVE,
    COL_LAST_NAME,
    COL_LAST_NAME_DATIVE,
    COL_LAST_NAME_NOMINATIVE,
    COL_MIDDLE_INITIAL,
    COL_MIDDLE_NAME,
    COL_NAME,
    COL_OATI_NUMBER,
    COL_OGRN,
    COL_ORGANIZATION,
    COL_OUTPUT_FILE,
    COL_PDF_FILE,
    COL_PERSON_ID,
    COL_PORTAL_ID,
    COL_PROCESSED_DATE,
    COL_ROLE,
    COL_SHORT_NAME,
    COL_STATUS,
    COL_WORD_NOMINATIVE,
    COL_WORD_TYPE,
    COL_WORD_VALUE,
    DATABASE_PATH,
    DEFAULT_ACTIVE,
    DEFAULT_HISTORY_LIMIT,
    DEFAULT_INACTIVE,
    DEFAULT_STATUS_SUCCESS,
    ROLE_DIRECTOR,
    ROLE_INSPECTOR,
    TABLE_DEPARTMENTS,
    TABLE_DIRECTORS,
    TABLE_INSPECTOR_CHIEFS,
    TABLE_INSPECTORS,
    TABLE_MANUAL_DECLENSIONS,
    TABLE_NAME_OVERRIDES,
    TABLE_PERSONS,
    TABLE_PROCESSING_HISTORY,
)

logger = logging.getLogger(__name__)


class Database:
    """
    Database manager for the PDF Parser OATI application.
    
    This class handles all SQLite database operations including:
    - Database initialization and schema creation
    - CRUD operations for departments, directors, inspectors, and inspector chiefs
    - Processing history tracking
    - Name-based lookups and matching
    
    Attributes:
        db_path: Path to the SQLite database file.
    """
    
    def __init__(self, db_path: str = str(DATABASE_PATH)) -> None:
        """
        Initialize the Database instance and create schema if needed.
        
        Args:
            db_path: Path to the SQLite database file. Defaults to the path
                defined in constants.
        """
        self.db_path: str = db_path
        Path(db_path).parent.mkdir(parents=True, exist_ok=True)
        self.init_database()
        logger.info(f"Database initialized at {db_path}")
    
    def _get_connection(self) -> sqlite3.Connection:
        """
        Create and return a new database connection with UTF-8 encoding support.
        
        Returns:
            A SQLite database connection object configured for UTF-8.
            
        Raises:
            sqlite3.Error: If connection to database fails.
        """
        try:
            conn = sqlite3.connect(self.db_path)
            conn.text_factory = str
            return conn
        except sqlite3.Error as e:
            logger.error(f"Failed to connect to database at {self.db_path}: {e}")
            raise
    
    def init_database(self) -> None:
        """
        Initialize database schema and populate with default data.
        
        Creates all required tables if they don't exist and inserts
        default departments, directors, inspectors, and inspector chiefs.
        Automatically migrates data to unified persons table.
        
        Raises:
            sqlite3.Error: If database initialization fails.
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            self._create_tables(cursor)
            self._migrate_departments_schema(cursor)
            self._populate_default_data(cursor)
            
            conn.commit()
            logger.info("Database schema initialized successfully")
        except sqlite3.Error as e:
            logger.error(f"Database initialization failed: {e}")
            raise
        finally:
            conn.close()
        
        # Migrate data from old tables to unified persons table
        try:
            stats = self.migrate_to_persons_table()
            logger.info(f"Auto-migration to persons table: {stats['total_persons']} persons loaded")
        except Exception as e:
            logger.error(f"Auto-migration failed (non-fatal): {e}")
    
    def _create_tables(self, cursor: sqlite3.Cursor) -> None:
        """
        Create all required database tables.
        
        Args:
            cursor: SQLite cursor for executing SQL statements.
        """
        cursor.execute(f'''
            CREATE TABLE IF NOT EXISTS {TABLE_DEPARTMENTS} (
                {COL_ID} INTEGER PRIMARY KEY AUTOINCREMENT,
                {COL_NAME} TEXT NOT NULL UNIQUE,
                {COL_SHORT_NAME} TEXT,
                {COL_IS_ACTIVE} INTEGER DEFAULT {DEFAULT_ACTIVE}
            )
        ''')
        
        cursor.execute(f'''
            CREATE TABLE IF NOT EXISTS {TABLE_PROCESSING_HISTORY} (
                {COL_ID} INTEGER PRIMARY KEY AUTOINCREMENT,
                {COL_PDF_FILE} TEXT NOT NULL,
                {COL_CITIZEN_NAME} TEXT,
                {COL_OATI_NUMBER} TEXT,
                {COL_PORTAL_ID} TEXT,
                {COL_PROCESSED_DATE} TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                {COL_OUTPUT_FILE} TEXT,
                {COL_STATUS} TEXT
            )
        ''')
        
        cursor.execute(f'''
            CREATE TABLE IF NOT EXISTS {TABLE_DIRECTORS} (
                {COL_ID} INTEGER PRIMARY KEY AUTOINCREMENT,
                {COL_LAST_NAME} TEXT NOT NULL,
                {COL_FIRST_NAME} TEXT NOT NULL,
                {COL_MIDDLE_NAME} TEXT NOT NULL,
                {COL_DEPARTMENT} TEXT NOT NULL,
                {COL_IS_ACTIVE} INTEGER DEFAULT {DEFAULT_ACTIVE}
            )
        ''')
        
        cursor.execute(f'''
            CREATE TABLE IF NOT EXISTS {TABLE_INSPECTORS} (
                {COL_ID} INTEGER PRIMARY KEY AUTOINCREMENT,
                {COL_LAST_NAME} TEXT NOT NULL,
                {COL_FIRST_NAME} TEXT NOT NULL,
                {COL_MIDDLE_NAME} TEXT NOT NULL,
                {COL_IS_ACTIVE} INTEGER DEFAULT {DEFAULT_ACTIVE}
            )
        ''')
        
        cursor.execute(f'''
            CREATE TABLE IF NOT EXISTS {TABLE_INSPECTOR_CHIEFS} (
                {COL_ID} INTEGER PRIMARY KEY AUTOINCREMENT,
                {COL_LAST_NAME_NOMINATIVE} TEXT NOT NULL,
                {COL_LAST_NAME_DATIVE} TEXT NOT NULL,
                {COL_FIRST_INITIAL} TEXT NOT NULL,
                {COL_MIDDLE_INITIAL} TEXT NOT NULL,
                {COL_ORGANIZATION} TEXT NOT NULL,
                {COL_IS_ACTIVE} INTEGER DEFAULT {DEFAULT_ACTIVE}
            )
        ''')
        
        cursor.execute(f'''
            CREATE TABLE IF NOT EXISTS {TABLE_PERSONS} (
                {COL_ID} INTEGER PRIMARY KEY AUTOINCREMENT,
                {COL_LAST_NAME} TEXT NOT NULL,
                {COL_FIRST_NAME} TEXT NOT NULL,
                {COL_MIDDLE_NAME} TEXT NOT NULL,
                {COL_ROLE} TEXT NOT NULL CHECK({COL_ROLE} IN ('{ROLE_DIRECTOR}', '{ROLE_INSPECTOR}')),
                {COL_DEPARTMENT} TEXT,
                {COL_LAST_NAME_DATIVE} TEXT,
                {COL_IS_ACTIVE} INTEGER DEFAULT {DEFAULT_ACTIVE},
                UNIQUE({COL_LAST_NAME}, {COL_FIRST_NAME}, {COL_MIDDLE_NAME})
            )
        ''')
        
        cursor.execute(f'''
            CREATE TABLE IF NOT EXISTS {TABLE_NAME_OVERRIDES} (
                {COL_ID} INTEGER PRIMARY KEY AUTOINCREMENT,
                {COL_WORD_NOMINATIVE} TEXT NOT NULL,
                {COL_WORD_VALUE} TEXT NOT NULL,
                {COL_WORD_TYPE} TEXT NOT NULL CHECK({COL_WORD_TYPE} IN ('surname', 'firstname', 'patronymic')),
                {COL_GENDER} TEXT NOT NULL CHECK({COL_GENDER} IN ('male', 'female')),
                {COL_CASE} TEXT NOT NULL CHECK({COL_CASE} IN ('nomn', 'gent', 'datv', 'accs', 'ablt', 'loct')),
                {COL_IS_ACTIVE} INTEGER DEFAULT {DEFAULT_ACTIVE},
                UNIQUE({COL_WORD_NOMINATIVE}, {COL_WORD_TYPE}, {COL_GENDER}, {COL_CASE})
            )
        ''')
        
        cursor.execute(f'''
            CREATE INDEX IF NOT EXISTS idx_name_overrides_lookup
            ON {TABLE_NAME_OVERRIDES} ({COL_WORD_NOMINATIVE}, {COL_WORD_TYPE}, {COL_GENDER}, {COL_CASE}, {COL_IS_ACTIVE})
        ''')
        
        cursor.execute(f'''
            CREATE TABLE IF NOT EXISTS {TABLE_MANUAL_DECLENSIONS} (
                {COL_ID} INTEGER PRIMARY KEY AUTOINCREMENT,
                {COL_PERSON_ID} INTEGER NOT NULL,
                {COL_CASE_TYPE} TEXT NOT NULL CHECK({COL_CASE_TYPE} IN ('nomn', 'gent', 'datv', 'accs')),
                {COL_CUSTOM_VALUE} TEXT NOT NULL,
                {COL_IS_ACTIVE} INTEGER DEFAULT {DEFAULT_ACTIVE},
                FOREIGN KEY ({COL_PERSON_ID}) REFERENCES {TABLE_PERSONS}({COL_ID}) ON DELETE CASCADE,
                UNIQUE({COL_PERSON_ID}, {COL_CASE_TYPE})
            )
        ''')
        
        cursor.execute(f'''
            CREATE INDEX IF NOT EXISTS idx_manual_declensions_lookup
            ON {TABLE_MANUAL_DECLENSIONS} ({COL_PERSON_ID}, {COL_CASE_TYPE}, {COL_IS_ACTIVE})
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS shipments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tracking_number TEXT NOT NULL UNIQUE,
                recipient_name TEXT,
                address TEXT,
                status TEXT,
                sent_date TEXT,
                delivery_date TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS shipment_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                shipment_id INTEGER NOT NULL,
                event_date TEXT NOT NULL,
                event_type TEXT NOT NULL,
                location TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (shipment_id) REFERENCES shipments(id) ON DELETE CASCADE
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS scanned_documents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_path TEXT NOT NULL,
                scan_date TEXT NOT NULL,
                document_type TEXT,
                citizen_id INTEGER,
                pages_count INTEGER,
                ocr_completed INTEGER DEFAULT 0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS mosedo_workflows (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                steps_json TEXT NOT NULL,
                description TEXT,
                created_date TEXT NOT NULL,
                is_active INTEGER DEFAULT 1
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS mosedo_jobs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                workflow_id INTEGER NOT NULL,
                status TEXT NOT NULL,
                current_step INTEGER DEFAULT 0,
                total_steps INTEGER,
                log TEXT,
                started_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                completed_at TIMESTAMP,
                FOREIGN KEY (workflow_id) REFERENCES mosedo_workflows(id) ON DELETE CASCADE
            )
        ''')
        
        logger.debug("Database tables created successfully")
    
    def _migrate_departments_schema(self, cursor: sqlite3.Cursor) -> None:
        """
        Migrate departments table to include INN, OGRN, and Address columns.
        
        Safely adds new columns if they don't exist (for legacy databases).
        
        Args:
            cursor: SQLite cursor for executing SQL statements.
        """
        cursor.execute(f"PRAGMA table_info({TABLE_DEPARTMENTS})")
        columns = {row[1] for row in cursor.fetchall()}
        
        if COL_INN not in columns:
            cursor.execute(f'''
                ALTER TABLE {TABLE_DEPARTMENTS}
                ADD COLUMN {COL_INN} TEXT DEFAULT ''
            ''')
            logger.info(f"Added {COL_INN} column to {TABLE_DEPARTMENTS}")
        
        if COL_OGRN not in columns:
            cursor.execute(f'''
                ALTER TABLE {TABLE_DEPARTMENTS}
                ADD COLUMN {COL_OGRN} TEXT DEFAULT ''
            ''')
            logger.info(f"Added {COL_OGRN} column to {TABLE_DEPARTMENTS}")
        
        if COL_ADDRESS not in columns:
            cursor.execute(f'''
                ALTER TABLE {TABLE_DEPARTMENTS}
                ADD COLUMN {COL_ADDRESS} TEXT DEFAULT ''
            ''')
            logger.info(f"Added {COL_ADDRESS} column to {TABLE_DEPARTMENTS}")
    
    def _populate_default_data(self, cursor: sqlite3.Cursor) -> None:
        """
        Populate database with default data from JSON configuration files.
        
        Loads and inserts default departments, directors, and inspector chiefs
        from JSON files if they don't already exist in the database.
        
        Args:
            cursor: SQLite cursor for executing SQL statements.
        """
        default_departments: List[Tuple[str, str]] = [
            ("Государственная жилищная инспекция города Москвы", "ГЖИ"),
            ("Департамент природопользования и охраны окружающей среды города Москвы", "Депприроды"),
            ("Комитет государственного строительного надзора города Москвы", "Мосгосстройнадзор"),
            ("Объединение административно-технических инспекций города Москвы", "ОАТИ"),
        ]
        
        for dept_name, short_name in default_departments:
            cursor.execute(f'''
                INSERT OR IGNORE INTO {TABLE_DEPARTMENTS} ({COL_NAME}, {COL_SHORT_NAME})
                VALUES (?, ?)
            ''', (dept_name, short_name))
        
        try:
            directors_json_path = os.path.join('config', 'directors.json')
            if os.path.exists(directors_json_path):
                with open(directors_json_path, 'r', encoding='utf-8') as f:
                    directors_config = json.load(f)
                    for director in directors_config.get('directors', []):
                        last_name = director['last_name']
                        initials = director['initials']
                        department = director['department']
                        first_initial, middle_initial = self._parse_initials(initials)
                        cursor.execute(f'''
                            INSERT OR IGNORE INTO {TABLE_DIRECTORS} 
                            ({COL_LAST_NAME}, {COL_FIRST_NAME}, {COL_MIDDLE_NAME}, {COL_DEPARTMENT})
                            VALUES (?, ?, ?, ?)
                        ''', (last_name, first_initial, middle_initial, department))
                logger.debug(f"Loaded {len(directors_config.get('directors', []))} directors from JSON")
        except Exception as e:
            logger.error(f"Failed to load directors from JSON: {e}")
        
        try:
            inspectors_chiefs_json_path = os.path.join('config', 'inspectors_chiefs.json')
            if os.path.exists(inspectors_chiefs_json_path):
                with open(inspectors_chiefs_json_path, 'r', encoding='utf-8') as f:
                    inspectors_config = json.load(f)
                    for inspector in inspectors_config.get('inspectors_chiefs', []):
                        nom = inspector['last_name_nominative']
                        dat = inspector['last_name_dative']
                        f_init = inspector['first_initial']
                        m_init = inspector['middle_initial']
                        org = inspector['organization']
                        cursor.execute(f'''
                            INSERT OR IGNORE INTO {TABLE_INSPECTOR_CHIEFS} 
                            ({COL_LAST_NAME_NOMINATIVE}, {COL_LAST_NAME_DATIVE}, {COL_FIRST_INITIAL}, 
                             {COL_MIDDLE_INITIAL}, {COL_ORGANIZATION})
                            VALUES (?, ?, ?, ?, ?)
                        ''', (nom, dat, f_init, m_init, org))
                        
                        first_initial = f_init
                        middle_initial = m_init
                        cursor.execute(f'''
                            INSERT OR IGNORE INTO {TABLE_INSPECTORS} 
                            ({COL_LAST_NAME}, {COL_FIRST_NAME}, {COL_MIDDLE_NAME})
                            VALUES (?, ?, ?)
                        ''', (nom, first_initial, middle_initial))
                logger.debug(f"Loaded {len(inspectors_config.get('inspectors_chiefs', []))} inspector chiefs from JSON")
        except Exception as e:
            logger.error(f"Failed to load inspector chiefs from JSON: {e}")
        
        try:
            name_overrides_json_path = os.path.join('config', 'name_overrides.json')
            if os.path.exists(name_overrides_json_path):
                with open(name_overrides_json_path, 'r', encoding='utf-8') as f:
                    overrides_config = json.load(f)
                    for override in overrides_config.get('overrides', []):
                        cursor.execute(f'''
                            INSERT OR IGNORE INTO {TABLE_NAME_OVERRIDES} 
                            ({COL_WORD_NOMINATIVE}, {COL_WORD_VALUE}, {COL_WORD_TYPE}, 
                             {COL_GENDER}, {COL_CASE})
                            VALUES (?, ?, ?, ?, ?)
                        ''', (
                            override['word_nominative'].lower(),
                            override['word_value'],
                            override['word_type'],
                            override['gender'],
                            override['case']
                        ))
                logger.debug(f"Loaded {len(overrides_config.get('overrides', []))} name overrides from JSON")
        except Exception as e:
            logger.error(f"Failed to load name overrides from JSON: {e}")
        
        logger.debug("Default data populated successfully from JSON files")
    
    def _parse_initials(self, initials: str) -> Tuple[str, str]:
        """
        Parse initials string into first and middle initials.
        
        Args:
            initials: String containing initials separated by dots (e.g., "А.И.").
            
        Returns:
            Tuple of (first_initial, middle_initial).
        """
        parts = initials.split('.')
        first_initial = parts[0] if len(parts) > 0 else ""
        middle_initial = parts[1] if len(parts) > 1 else ""
        return first_initial, middle_initial
    
    def get_department_by_name(self, name: str) -> Optional[Tuple]:
        """
        Retrieve a department by its full name.
        
        Args:
            name: Full name of the department.
            
        Returns:
            Tuple (id, name, short_name, inn, ogrn, address, is_active) or None if not found.
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute(f'''
                SELECT {COL_ID}, {COL_NAME}, {COL_SHORT_NAME}, {COL_INN}, {COL_OGRN}, {COL_ADDRESS}, {COL_IS_ACTIVE}
                FROM {TABLE_DEPARTMENTS}
                WHERE {COL_NAME} = ?
                LIMIT 1
            ''', (name,))
            result = cursor.fetchone()
            return result
        except sqlite3.Error as e:
            logger.error(f"Failed to get department by name '{name}': {e}")
            return None
        finally:
            conn.close()
    
    def get_department_by_inn(self, inn: str) -> Optional[Tuple]:
        """
        Retrieve a department by its INN (Tax ID).
        
        Args:
            inn: INN (Tax Identification Number).
            
        Returns:
            Tuple (id, name, short_name, inn, ogrn, address, is_active) or None if not found.
        """
        if not inn or not inn.strip():
            return None
        
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute(f'''
                SELECT {COL_ID}, {COL_NAME}, {COL_SHORT_NAME}, {COL_INN}, {COL_OGRN}, {COL_ADDRESS}, {COL_IS_ACTIVE}
                FROM {TABLE_DEPARTMENTS}
                WHERE {COL_INN} = ? AND {COL_INN} != ''
                LIMIT 1
            ''', (inn.strip(),))
            result = cursor.fetchone()
            return result
        except sqlite3.Error as e:
            logger.error(f"Failed to get department by INN '{inn}': {e}")
            return None
        finally:
            conn.close()
    
    def get_department_by_ogrn(self, ogrn: str) -> Optional[Tuple]:
        """
        Retrieve a department by its OGRN (Primary State Registration Number).
        
        Args:
            ogrn: OGRN (Primary State Registration Number).
            
        Returns:
            Tuple (id, name, short_name, inn, ogrn, address, is_active) or None if not found.
        """
        if not ogrn or not ogrn.strip():
            return None
        
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute(f'''
                SELECT {COL_ID}, {COL_NAME}, {COL_SHORT_NAME}, {COL_INN}, {COL_OGRN}, {COL_ADDRESS}, {COL_IS_ACTIVE}
                FROM {TABLE_DEPARTMENTS}
                WHERE {COL_OGRN} = ? AND {COL_OGRN} != ''
                LIMIT 1
            ''', (ogrn.strip(),))
            result = cursor.fetchone()
            return result
        except sqlite3.Error as e:
            logger.error(f"Failed to get department by OGRN '{ogrn}': {e}")
            return None
        finally:
            conn.close()
    
    def find_departments_by_director_surname(self, surname_dative: str) -> List[Dict[str, Any]]:
        """
        Find departments by matching director's surname in dative case.
        
        This method is useful when parsing resolutions that mention directors
        by surname (e.g., "Урожаевой Ю.В." → Department of Nature Use).
        It converts the dative surname back to nominative and searches for matches.
        
        Args:
            surname_dative: Director's surname in dative case (e.g., "Урожаевой", "Слободчикову")
            
        Returns:
            List of dictionaries with keys: id, name, short_name, inn, ogrn, address, director_full_name
        """
        import pymorphy3
        
        if not surname_dative or not surname_dative.strip():
            return []
        
        morph = pymorphy3.MorphAnalyzer()
        
        # Parse surname and try to convert to nominative case
        parsed = morph.parse(surname_dative)
        nominative_variants = set()
        
        for p in parsed:
            if 'Surn' in p.tag or 'Name' in p.tag:
                normal_form = p.normal_form
                nominative_variants.add(normal_form.capitalize())
                
                # Also add the inflected form in case it's already nominative
                nominative_variants.add(surname_dative.capitalize())
                
                # Try to explicitly inflect to nominative
                try:
                    nominative_inflected = p.inflect({'nomn'})
                    if nominative_inflected:
                        nominative_variants.add(nominative_inflected.word.capitalize())
                except:
                    pass
        
        if not nominative_variants:
            # If morphological analysis failed, try direct match
            nominative_variants = {surname_dative.capitalize()}
        
        logger.debug(f"Searching for directors with surname variants: {nominative_variants}")
        
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            results = []
            
            for variant in nominative_variants:
                cursor.execute(f'''
                    SELECT p.{COL_DEPARTMENT}, p.{COL_LAST_NAME}, p.{COL_FIRST_NAME}, p.{COL_MIDDLE_NAME},
                           d.{COL_ID}, d.{COL_NAME}, d.{COL_SHORT_NAME}, d.{COL_INN}, d.{COL_OGRN}, d.{COL_ADDRESS}
                    FROM {TABLE_PERSONS} p
                    LEFT JOIN {TABLE_DEPARTMENTS} d ON p.{COL_DEPARTMENT} = d.{COL_NAME}
                    WHERE p.{COL_LAST_NAME} LIKE ? AND p.{COL_ROLE} = ?
                    AND p.{COL_IS_ACTIVE} = ?
                ''', (f"%{variant}%", ROLE_DIRECTOR, DEFAULT_ACTIVE))
                
                for row in cursor.fetchall():
                    dept_name, last_name, first_name, middle_name, dept_id, full_name, short_name, inn, ogrn, address = row
                    
                    # Skip if department not found in departments table
                    if not dept_id:
                        logger.warning(f"Director {last_name} has department '{dept_name}' but it's not in departments table")
                        continue
                    
                    result = {
                        'id': dept_id,
                        'name': full_name,
                        'short_name': short_name,
                        'inn': inn,
                        'ogrn': ogrn,
                        'address': address,
                        'director_full_name': f"{last_name} {first_name} {middle_name}",
                        'matched_surname': variant
                    }
                    
                    # Avoid duplicates
                    if result not in results:
                        results.append(result)
                        logger.info(f"Found department by director surname: {full_name} (director: {last_name})")
            
            return results
            
        except sqlite3.Error as e:
            logger.error(f"Failed to find departments by director surname '{surname_dative}': {e}")
            return []
        finally:
            conn.close()
    
    def add_department(self, name: str, short_name: str = "", inn: str = "", ogrn: str = "", address: str = "") -> bool:
        """
        Add a new department to the database.
        
        Args:
            name: Full name of the department.
            short_name: Abbreviated name of the department.
            
        Returns:
            True if department was added successfully, False if it already exists
            or if an error occurred.
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute(f'''
                INSERT INTO {TABLE_DEPARTMENTS} ({COL_NAME}, {COL_SHORT_NAME}, {COL_INN}, {COL_OGRN}, {COL_ADDRESS})
                VALUES (?, ?, ?, ?, ?)
            ''', (name, short_name, inn, ogrn, address))
            conn.commit()
            logger.info(f"Department added: {name}")
            return True
        except sqlite3.IntegrityError:
            logger.warning(f"Department already exists: {name}")
            return False
        except sqlite3.Error as e:
            logger.error(f"Failed to add department '{name}': {e}")
            return False
        finally:
            conn.close()
    
    def get_all_departments(self, active_only: bool = True) -> List[Dict[str, Any]]:
        """
        Retrieve all departments from the database.
        
        Args:
            active_only: If True, only return active departments. Defaults to True.
            
        Returns:
            List of dictionaries containing department data with keys:
            'id', 'name', 'short_name', 'is_active'.
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            query = f'''
                SELECT {COL_ID}, {COL_NAME}, {COL_SHORT_NAME}, {COL_IS_ACTIVE} 
                FROM {TABLE_DEPARTMENTS}
            '''
            if active_only:
                query += f' WHERE {COL_IS_ACTIVE} = {DEFAULT_ACTIVE}'
            
            cursor.execute(query)
            rows = cursor.fetchall()
            
            return [
                {
                    'id': row[0],
                    'name': row[1],
                    'short_name': row[2],
                    'is_active': bool(row[3])
                }
                for row in rows
            ]
        except sqlite3.Error as e:
            logger.error(f"Failed to retrieve departments: {e}")
            return []
        finally:
            conn.close()
    
    def update_department(self, dept_id: int, name: str, short_name: str = "", inn: str = "", ogrn: str = "", address: str = "") -> bool:
        """
        Update an existing department.
        
        Args:
            dept_id: ID of the department to update.
            name: New full name of the department.
            short_name: New abbreviated name of the department.
            inn: Tax Identification Number (INN).
            ogrn: Primary State Registration Number (OGRN).
            address: Legal address of the department.
            
        Returns:
            True if update was successful, False otherwise.
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute(f'''
                UPDATE {TABLE_DEPARTMENTS}
                SET {COL_NAME} = ?, {COL_SHORT_NAME} = ?, {COL_INN} = ?, {COL_OGRN} = ?, {COL_ADDRESS} = ?
                WHERE {COL_ID} = ?
            ''', (name, short_name, inn, ogrn, address, dept_id))
            conn.commit()
            
            if cursor.rowcount > 0:
                logger.info(f"Department updated: ID {dept_id}")
                return True
            else:
                logger.warning(f"Department not found: ID {dept_id}")
                return False
        except sqlite3.Error as e:
            logger.error(f"Failed to update department ID {dept_id}: {e}")
            return False
        finally:
            conn.close()
    
    def delete_department(self, dept_id: int) -> bool:
        """
        Soft delete a department by marking it as inactive.
        
        Args:
            dept_id: ID of the department to delete.
            
        Returns:
            True if deletion was successful, False otherwise.
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute(f'''
                UPDATE {TABLE_DEPARTMENTS} 
                SET {COL_IS_ACTIVE} = {DEFAULT_INACTIVE} 
                WHERE {COL_ID} = ?
            ''', (dept_id,))
            conn.commit()
            
            if cursor.rowcount > 0:
                logger.info(f"Department deleted (soft): ID {dept_id}")
                return True
            else:
                logger.warning(f"Department not found: ID {dept_id}")
                return False
        except sqlite3.Error as e:
            logger.error(f"Failed to delete department ID {dept_id}: {e}")
            return False
        finally:
            conn.close()
    
    def add_processing_record(
        self,
        pdf_file: str,
        citizen_name: str,
        oati_number: str,
        portal_id: str,
        output_file: str,
        status: str = DEFAULT_STATUS_SUCCESS
    ) -> bool:
        """
        Add a processing history record to the database.
        
        Args:
            pdf_file: Name of the processed PDF file.
            citizen_name: Name of the citizen from the document.
            oati_number: OATI reference number.
            portal_id: Portal identification number.
            output_file: Name of the generated output file.
            status: Processing status. Defaults to 'success'.
            
        Returns:
            True if record was added successfully, False otherwise.
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute(f'''
                INSERT INTO {TABLE_PROCESSING_HISTORY} 
                ({COL_PDF_FILE}, {COL_CITIZEN_NAME}, {COL_OATI_NUMBER}, 
                 {COL_PORTAL_ID}, {COL_OUTPUT_FILE}, {COL_STATUS})
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (pdf_file, citizen_name, oati_number, portal_id, output_file, status))
            conn.commit()
            logger.info(f"Processing record added: {pdf_file}")
            return True
        except sqlite3.Error as e:
            logger.error(f"Failed to add processing record for '{pdf_file}': {e}")
            return False
        finally:
            conn.close()
    
    def get_processing_history(self, limit: int = DEFAULT_HISTORY_LIMIT) -> List[Dict[str, Any]]:
        """
        Retrieve processing history records.
        
        Args:
            limit: Maximum number of records to retrieve. Defaults to 100.
            
        Returns:
            List of dictionaries containing processing history data with keys:
            'id', 'pdf_file', 'citizen_name', 'oati_number', 'portal_id',
            'processed_date', 'output_file', 'status'.
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute(f'''
                SELECT {COL_ID}, {COL_PDF_FILE}, {COL_CITIZEN_NAME}, {COL_OATI_NUMBER}, 
                       {COL_PORTAL_ID}, {COL_PROCESSED_DATE}, {COL_OUTPUT_FILE}, {COL_STATUS}
                FROM {TABLE_PROCESSING_HISTORY}
                ORDER BY {COL_PROCESSED_DATE} DESC
                LIMIT ?
            ''', (limit,))
            rows = cursor.fetchall()
            
            return [
                {
                    'id': row[0],
                    'pdf_file': row[1],
                    'citizen_name': row[2],
                    'oati_number': row[3],
                    'portal_id': row[4],
                    'processed_date': row[5],
                    'output_file': row[6],
                    'status': row[7]
                }
                for row in rows
            ]
        except sqlite3.Error as e:
            logger.error(f"Failed to retrieve processing history: {e}")
            return []
        finally:
            conn.close()
    
    def get_all_directors(self) -> List[Dict[str, Any]]:
        """
        Retrieve all active directors from the database.
        
        Returns:
            List of dictionaries containing director data with keys:
            'id', 'last_name', 'first_name', 'middle_name', 'department', 'is_active'.
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute(f'''
                SELECT {COL_ID}, {COL_LAST_NAME}, {COL_FIRST_NAME}, {COL_MIDDLE_NAME}, 
                       {COL_DEPARTMENT}, {COL_IS_ACTIVE}
                FROM {TABLE_DIRECTORS}
                WHERE {COL_IS_ACTIVE} = {DEFAULT_ACTIVE}
            ''')
            rows = cursor.fetchall()
            
            return [
                {
                    'id': row[0],
                    'last_name': row[1],
                    'first_name': row[2],
                    'middle_name': row[3],
                    'department': row[4],
                    'is_active': bool(row[5])
                }
                for row in rows
            ]
        except sqlite3.Error as e:
            logger.error(f"Failed to retrieve directors: {e}")
            return []
        finally:
            conn.close()
    
    def get_all_inspectors(self) -> List[Dict[str, Any]]:
        """
        Retrieve all active inspectors from the database.
        
        Returns:
            List of dictionaries containing inspector data with keys:
            'id', 'last_name', 'first_name', 'middle_name', 'is_active'.
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute(f'''
                SELECT {COL_ID}, {COL_LAST_NAME}, {COL_FIRST_NAME}, {COL_MIDDLE_NAME}, 
                       {COL_IS_ACTIVE}
                FROM {TABLE_INSPECTORS}
                WHERE {COL_IS_ACTIVE} = {DEFAULT_ACTIVE}
            ''')
            rows = cursor.fetchall()
            
            return [
                {
                    'id': row[0],
                    'last_name': row[1],
                    'first_name': row[2],
                    'middle_name': row[3],
                    'is_active': bool(row[4])
                }
                for row in rows
            ]
        except sqlite3.Error as e:
            logger.error(f"Failed to retrieve inspectors: {e}")
            return []
        finally:
            conn.close()
    
    def find_director_by_name(
        self,
        last_name: str,
        first_initial: str,
        middle_initial: str
    ) -> Optional[Dict[str, Any]]:
        """
        Find a director by their name components.
        
        Args:
            last_name: Last name of the director.
            first_initial: First initial of the director.
            middle_initial: Middle initial of the director.
            
        Returns:
            Dictionary containing director data if found, None otherwise.
            Dictionary keys: 'id', 'last_name', 'first_name', 'middle_name', 'department'.
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute(f'''
                SELECT {COL_ID}, {COL_LAST_NAME}, {COL_FIRST_NAME}, {COL_MIDDLE_NAME}, 
                       {COL_DEPARTMENT}
                FROM {TABLE_DIRECTORS}
                WHERE LOWER({COL_LAST_NAME}) = LOWER(?) 
                AND LOWER({COL_FIRST_NAME}) = LOWER(?) 
                AND LOWER({COL_MIDDLE_NAME}) = LOWER(?)
                AND {COL_IS_ACTIVE} = {DEFAULT_ACTIVE}
            ''', (last_name, first_initial, middle_initial))
            row = cursor.fetchone()
            
            if row:
                return {
                    'id': row[0],
                    'last_name': row[1],
                    'first_name': row[2],
                    'middle_name': row[3],
                    'department': row[4]
                }
            return None
        except sqlite3.Error as e:
            logger.error(f"Failed to find director '{last_name} {first_initial}.{middle_initial}.': {e}")
            return None
        finally:
            conn.close()
    
    def add_director(
        self,
        last_name: str,
        first_name: str,
        middle_name: str,
        department: str
    ) -> bool:
        """
        Add a new director to the database (wrapper method).
        
        This method is a legacy wrapper that delegates to add_person().
        Internally stores directors in the unified persons table.
        
        Args:
            last_name: Last name of the director.
            first_name: First name of the director.
            middle_name: Middle name of the director.
            department: Department name.
            
        Returns:
            True if director was added successfully, False otherwise.
        """
        return self.add_person(
            last_name=last_name,
            first_name=first_name,
            middle_name=middle_name,
            role=ROLE_DIRECTOR,
            department=department,
            last_name_dative=None
        )
    
    def update_director(
        self,
        director_id: int,
        last_name: str,
        first_name: str,
        middle_name: str,
        department: str
    ) -> bool:
        """
        Update an existing director (wrapper method).
        
        This method is a legacy wrapper that delegates to update_person().
        Internally updates directors in the unified persons table.
        
        Args:
            director_id: ID of the director to update.
            last_name: Last name of the director.
            first_name: First name of the director.
            middle_name: Middle name of the director.
            department: Department name.
            
        Returns:
            True if update was successful, False otherwise.
        """
        return self.update_person(
            person_id=director_id,
            last_name=last_name,
            first_name=first_name,
            middle_name=middle_name,
            role=ROLE_DIRECTOR,
            department=department,
            last_name_dative=None
        )
    
    def delete_director(self, director_id: int) -> bool:
        """
        Soft delete a director by marking them as inactive (wrapper method).
        
        This method is a legacy wrapper that delegates to delete_person().
        Internally deletes directors from the unified persons table.
        
        Args:
            director_id: ID of the director to delete.
            
        Returns:
            True if deletion was successful, False otherwise.
        """
        return self.delete_person(person_id=director_id)
    
    def find_inspector_by_name(
        self,
        last_name: str,
        first_initial: str,
        middle_initial: str
    ) -> Optional[Dict[str, Any]]:
        """
        Find an inspector by their name components.
        
        Args:
            last_name: Last name of the inspector.
            first_initial: First initial of the inspector.
            middle_initial: Middle initial of the inspector.
            
        Returns:
            Dictionary containing inspector data if found, None otherwise.
            Dictionary keys: 'id', 'last_name', 'first_name', 'middle_name'.
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute(f'''
                SELECT {COL_ID}, {COL_LAST_NAME}, {COL_FIRST_NAME}, {COL_MIDDLE_NAME}
                FROM {TABLE_INSPECTORS}
                WHERE LOWER({COL_LAST_NAME}) = LOWER(?) 
                AND LOWER({COL_FIRST_NAME}) = LOWER(?) 
                AND LOWER({COL_MIDDLE_NAME}) = LOWER(?)
                AND {COL_IS_ACTIVE} = {DEFAULT_ACTIVE}
            ''', (last_name, first_initial, middle_initial))
            row = cursor.fetchone()
            
            if row:
                return {
                    'id': row[0],
                    'last_name': row[1],
                    'first_name': row[2],
                    'middle_name': row[3]
                }
            return None
        except sqlite3.Error as e:
            logger.error(f"Failed to find inspector '{last_name} {first_initial}.{middle_initial}.': {e}")
            return None
        finally:
            conn.close()
    
    def add_inspector_chief(
        self,
        last_name_nominative: str,
        last_name_dative: str,
        first_initial: str,
        middle_initial: str,
        organization: str
    ) -> bool:
        """
        Add a new inspector chief to the database (wrapper method).
        
        This method is a legacy wrapper that delegates to add_person().
        Internally stores inspectors in the unified persons table.
        
        Args:
            last_name_nominative: Last name in nominative case.
            last_name_dative: Last name in dative case.
            first_initial: First name initial.
            middle_initial: Middle name initial.
            organization: Organization name.
            
        Returns:
            True if inspector chief was added successfully, False otherwise.
        """
        return self.add_person(
            last_name=last_name_nominative,
            first_name=first_initial,
            middle_name=middle_initial,
            role=ROLE_INSPECTOR,
            department=organization,
            last_name_dative=last_name_dative
        )
    
    def get_all_inspectors_chiefs(self) -> List[Dict[str, Any]]:
        """
        Retrieve all active inspector chiefs from the database.
        
        Returns:
            List of dictionaries containing inspector chief data with keys:
            'id', 'last_name_nominative', 'last_name_dative', 'first_initial',
            'middle_initial', 'organization', 'is_active'.
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute(f'''
                SELECT {COL_ID}, {COL_LAST_NAME_NOMINATIVE}, {COL_LAST_NAME_DATIVE}, 
                       {COL_FIRST_INITIAL}, {COL_MIDDLE_INITIAL}, {COL_ORGANIZATION}, 
                       {COL_IS_ACTIVE}
                FROM {TABLE_INSPECTOR_CHIEFS}
                WHERE {COL_IS_ACTIVE} = {DEFAULT_ACTIVE}
            ''')
            rows = cursor.fetchall()
            
            return [
                {
                    'id': row[0],
                    'last_name_nominative': row[1],
                    'last_name_dative': row[2],
                    'first_initial': row[3],
                    'middle_initial': row[4],
                    'organization': row[5],
                    'is_active': bool(row[6])
                }
                for row in rows
            ]
        except sqlite3.Error as e:
            logger.error(f"Failed to retrieve inspector chiefs: {e}")
            return []
        finally:
            conn.close()
    
    def update_inspector_chief(
        self,
        chief_id: int,
        last_name_nominative: str,
        last_name_dative: str,
        first_initial: str,
        middle_initial: str,
        organization: str
    ) -> bool:
        """
        Update an existing inspector chief (wrapper method).
        
        This method is a legacy wrapper that delegates to update_person().
        Internally updates inspectors in the unified persons table.
        
        Args:
            chief_id: ID of the inspector chief to update.
            last_name_nominative: Last name in nominative case.
            last_name_dative: Last name in dative case.
            first_initial: First name initial.
            middle_initial: Middle name initial.
            organization: Organization name.
            
        Returns:
            True if update was successful, False otherwise.
        """
        return self.update_person(
            person_id=chief_id,
            last_name=last_name_nominative,
            first_name=first_initial,
            middle_name=middle_initial,
            role=ROLE_INSPECTOR,
            department=organization,
            last_name_dative=last_name_dative
        )
    
    def delete_inspector_chief(self, chief_id: int) -> bool:
        """
        Soft delete an inspector chief by marking them as inactive (wrapper method).
        
        This method is a legacy wrapper that delegates to delete_person().
        Internally deletes inspectors from the unified persons table.
        
        Args:
            chief_id: ID of the inspector chief to delete.
            
        Returns:
            True if deletion was successful, False otherwise.
        """
        return self.delete_person(person_id=chief_id)
    
    def find_inspector_chief_by_dative_name(
        self,
        dative_last_name: str,
        first_initial: str,
        middle_initial: str
    ) -> Optional[Dict[str, Any]]:
        """
        Find an inspector chief by their name in dative case.
        
        Args:
            dative_last_name: Last name in dative case.
            first_initial: First name initial.
            middle_initial: Middle name initial.
            
        Returns:
            Dictionary containing inspector chief data if found, None otherwise.
            Dictionary keys: 'id', 'last_name_nominative', 'last_name_dative',
            'first_initial', 'middle_initial', 'organization'.
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute(f'''
                SELECT {COL_ID}, {COL_LAST_NAME_NOMINATIVE}, {COL_LAST_NAME_DATIVE}, 
                       {COL_FIRST_INITIAL}, {COL_MIDDLE_INITIAL}, {COL_ORGANIZATION}
                FROM {TABLE_INSPECTOR_CHIEFS}
                WHERE LOWER({COL_LAST_NAME_DATIVE}) = LOWER(?) 
                AND LOWER({COL_FIRST_INITIAL}) = LOWER(?) 
                AND LOWER({COL_MIDDLE_INITIAL}) = LOWER(?)
                AND {COL_IS_ACTIVE} = {DEFAULT_ACTIVE}
            ''', (dative_last_name, first_initial, middle_initial))
            row = cursor.fetchone()
            
            if row:
                return {
                    'id': row[0],
                    'last_name_nominative': row[1],
                    'last_name_dative': row[2],
                    'first_initial': row[3],
                    'middle_initial': row[4],
                    'organization': row[5]
                }
            return None
        except sqlite3.Error as e:
            logger.error(f"Failed to find inspector chief by dative name '{dative_last_name}': {e}")
            return None
        finally:
            conn.close()
    
    def get_all_persons(
        self,
        role: Optional[str] = None,
        active_only: bool = True
    ) -> List[Dict[str, Any]]:
        """
        Retrieve all persons from the database.
        
        Args:
            role: If specified, filter by role ('director' or 'inspector').
            active_only: If True, only return active persons. Defaults to True.
            
        Returns:
            List of dictionaries containing person data with keys:
            'id', 'last_name', 'first_name', 'middle_name', 'role', 
            'department', 'last_name_dative', 'is_active'.
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            query = f'''
                SELECT {COL_ID}, {COL_LAST_NAME}, {COL_FIRST_NAME}, {COL_MIDDLE_NAME}, 
                       {COL_ROLE}, {COL_DEPARTMENT}, {COL_LAST_NAME_DATIVE}, {COL_IS_ACTIVE}
                FROM {TABLE_PERSONS}
            '''
            
            conditions = []
            params = []
            
            if active_only:
                conditions.append(f'{COL_IS_ACTIVE} = {DEFAULT_ACTIVE}')
            
            if role:
                conditions.append(f'{COL_ROLE} = ?')
                params.append(role)
            
            if conditions:
                query += ' WHERE ' + ' AND '.join(conditions)
            
            cursor.execute(query, params)
            rows = cursor.fetchall()
            
            return [
                {
                    'id': row[0],
                    'last_name': row[1],
                    'first_name': row[2],
                    'middle_name': row[3],
                    'role': row[4],
                    'department': row[5],
                    'last_name_dative': row[6],
                    'is_active': bool(row[7])
                }
                for row in rows
            ]
        except sqlite3.Error as e:
            logger.error(f"Failed to retrieve persons: {e}")
            return []
        finally:
            conn.close()
    
    def add_person(
        self,
        last_name: str,
        first_name: str,
        middle_name: str,
        role: str,
        department: Optional[str] = None,
        last_name_dative: Optional[str] = None
    ) -> bool:
        """
        Add a new person to the database.
        
        Args:
            last_name: Last name of the person.
            first_name: First name of the person.
            middle_name: Middle name of the person.
            role: Role of the person ('director' or 'inspector').
            department: Department name (required for directors, should be None for inspectors).
            last_name_dative: Last name in dative case (required for inspectors, should be None for directors).
            
        Returns:
            True if person was added successfully, False if person already exists
            or if an error occurred.
        """
        if role not in [ROLE_DIRECTOR, ROLE_INSPECTOR]:
            logger.error(f"Invalid role '{role}'. Must be '{ROLE_DIRECTOR}' or '{ROLE_INSPECTOR}'")
            return False
        
        if role == ROLE_DIRECTOR and not department:
            logger.error("Department is required for directors")
            return False
        
        if role == ROLE_INSPECTOR and not last_name_dative:
            logger.error("Last name in dative case is required for inspectors")
            return False
        
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute(f'''
                INSERT INTO {TABLE_PERSONS} 
                ({COL_LAST_NAME}, {COL_FIRST_NAME}, {COL_MIDDLE_NAME}, {COL_ROLE}, 
                 {COL_DEPARTMENT}, {COL_LAST_NAME_DATIVE})
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (last_name, first_name, middle_name, role, department, last_name_dative))
            conn.commit()
            logger.info(f"Person added: {last_name} {first_name} {middle_name} ({role})")
            return True
        except sqlite3.IntegrityError:
            logger.warning(f"Person already exists: {last_name} {first_name} {middle_name}")
            return False
        except sqlite3.Error as e:
            logger.error(f"Failed to add person '{last_name} {first_name} {middle_name}': {e}")
            return False
        finally:
            conn.close()
    
    def update_person(
        self,
        person_id: int,
        last_name: str,
        first_name: str,
        middle_name: str,
        role: str,
        department: Optional[str] = None,
        last_name_dative: Optional[str] = None
    ) -> bool:
        """
        Update an existing person in the database.
        
        Args:
            person_id: ID of the person to update.
            last_name: Last name of the person.
            first_name: First name of the person.
            middle_name: Middle name of the person.
            role: Role of the person ('director' or 'inspector').
            department: Department name (required for directors, should be None for inspectors).
            last_name_dative: Last name in dative case (required for inspectors, should be None for directors).
            
        Returns:
            True if update was successful, False otherwise.
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute(f'''
                UPDATE {TABLE_PERSONS}
                SET {COL_LAST_NAME} = ?, {COL_FIRST_NAME} = ?, {COL_MIDDLE_NAME} = ?, 
                    {COL_ROLE} = ?, {COL_DEPARTMENT} = ?, {COL_LAST_NAME_DATIVE} = ?
                WHERE {COL_ID} = ?
            ''', (last_name, first_name, middle_name, role, department, last_name_dative, person_id))
            conn.commit()
            
            if cursor.rowcount > 0:
                logger.info(f"Person updated: ID {person_id}")
                return True
            else:
                logger.warning(f"Person not found: ID {person_id}")
                return False
        except sqlite3.Error as e:
            logger.error(f"Failed to update person ID {person_id}: {e}")
            return False
        finally:
            conn.close()
    
    def delete_person(self, person_id: int) -> bool:
        """
        Soft delete a person by marking them as inactive.
        
        Args:
            person_id: ID of the person to delete.
            
        Returns:
            True if deletion was successful, False otherwise.
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute(f'''
                UPDATE {TABLE_PERSONS}
                SET {COL_IS_ACTIVE} = {DEFAULT_INACTIVE}
                WHERE {COL_ID} = ?
            ''', (person_id,))
            conn.commit()
            
            if cursor.rowcount > 0:
                logger.info(f"Person deleted (soft): ID {person_id}")
                return True
            else:
                logger.warning(f"Person not found: ID {person_id}")
                return False
        except sqlite3.Error as e:
            logger.error(f"Failed to delete person ID {person_id}: {e}")
            return False
        finally:
            conn.close()
    
    def match_recipient_from_resolution(
        self,
        dative_last_name: str,
        first_initial: str,
        middle_initial: str,
        decliner: Any
    ) -> Optional[Dict[str, Any]]:
        """
        Match a recipient from a resolution document by converting dative to nominative case.
        
        This method attempts to find a matching person in the persons table by first
        trying to match inspectors by dative last name, then converting to nominative
        and searching for directors.
        
        Args:
            dative_last_name: Last name in dative case.
            first_initial: First name initial.
            middle_initial: Middle name initial.
            decliner: Name declension utility object with dative_to_nominative method.
            
        Returns:
            Dictionary with 'type' ('director' or 'inspector') and 'data' keys
            if a match is found, None otherwise.
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            # First try direct dative match for inspectors in persons table
            cursor.execute(f'''
                SELECT {COL_ID}, {COL_LAST_NAME}, {COL_FIRST_NAME}, {COL_MIDDLE_NAME}, 
                       {COL_ROLE}, {COL_DEPARTMENT}, {COL_LAST_NAME_DATIVE}
                FROM {TABLE_PERSONS}
                WHERE LOWER({COL_LAST_NAME_DATIVE}) = LOWER(?) 
                AND LOWER({COL_FIRST_NAME}) = LOWER(?) 
                AND LOWER({COL_MIDDLE_NAME}) = LOWER(?)
                AND {COL_ROLE} = '{ROLE_INSPECTOR}'
                AND {COL_IS_ACTIVE} = {DEFAULT_ACTIVE}
            ''', (dative_last_name, first_initial, middle_initial))
            row = cursor.fetchone()
            
            if row:
                person_data = {
                    'id': row[0],
                    'last_name': row[1],
                    'first_name': row[2],
                    'middle_name': row[3],
                    'role': row[4],
                    'department': row[5],
                    'last_name_dative': row[6]
                }
                logger.debug(f"Matched inspector (dative): {dative_last_name} {first_initial}.{middle_initial}.")
                return {'type': 'inspector', 'data': person_data}
            
            # Convert to nominative and try directors
            nominative_last, f_init, m_init = decliner.dative_to_nominative(
                dative_last_name, first_initial, middle_initial
            )
            
            cursor.execute(f'''
                SELECT {COL_ID}, {COL_LAST_NAME}, {COL_FIRST_NAME}, {COL_MIDDLE_NAME}, 
                       {COL_ROLE}, {COL_DEPARTMENT}, {COL_LAST_NAME_DATIVE}
                FROM {TABLE_PERSONS}
                WHERE LOWER({COL_LAST_NAME}) = LOWER(?) 
                AND LOWER({COL_FIRST_NAME}) = LOWER(?) 
                AND LOWER({COL_MIDDLE_NAME}) = LOWER(?)
                AND {COL_ROLE} = '{ROLE_DIRECTOR}'
                AND {COL_IS_ACTIVE} = {DEFAULT_ACTIVE}
            ''', (nominative_last, f_init, m_init))
            row = cursor.fetchone()
            
            if row:
                person_data = {
                    'id': row[0],
                    'last_name': row[1],
                    'first_name': row[2],
                    'middle_name': row[3],
                    'role': row[4],
                    'department': row[5],
                    'last_name_dative': row[6]
                }
                logger.debug(f"Matched director: {nominative_last} {f_init}.{m_init}.")
                return {'type': 'director', 'data': person_data}
            
            logger.debug(f"No match found for: {dative_last_name} -> {nominative_last} {f_init}.{m_init}.")
            return None
        except Exception as e:
            logger.error(f"Failed to match recipient from resolution: {e}")
            return None
        finally:
            conn.close()
    
    def migrate_to_persons_table(self) -> Dict[str, int]:
        """
        Migrate data from old directors and inspectors_chiefs tables to the new persons table.
        
        This method consolidates directors and inspector chiefs into a unified persons table.
        In case of duplicates (same full name), inspectors take priority over directors.
        
        Returns:
            Dictionary with migration statistics: 
            {'directors_migrated', 'inspectors_migrated', 'duplicates_resolved', 'total_persons'}
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            # Track migration statistics
            stats = {
                'directors_migrated': 0,
                'inspectors_migrated': 0,
                'duplicates_resolved': 0,
                'total_persons': 0
            }
            
            # Dictionary to track persons by (last_name, first_name, middle_name)
            persons_map = {}
            
            # Step 1: Get all directors
            cursor.execute(f'''
                SELECT DISTINCT {COL_LAST_NAME}, {COL_FIRST_NAME}, {COL_MIDDLE_NAME}, {COL_DEPARTMENT}
                FROM {TABLE_DIRECTORS}
                WHERE {COL_IS_ACTIVE} = {DEFAULT_ACTIVE}
            ''')
            directors = cursor.fetchall()
            
            for last_name, first_name, middle_name, department in directors:
                key = (last_name, first_name, middle_name)
                persons_map[key] = {
                    'role': ROLE_DIRECTOR,
                    'department': department,
                    'last_name_dative': None
                }
                stats['directors_migrated'] += 1
            
            # Step 2: Get all inspectors (overwrite directors if duplicate)
            cursor.execute(f'''
                SELECT DISTINCT {COL_LAST_NAME_NOMINATIVE}, {COL_FIRST_INITIAL}, {COL_MIDDLE_INITIAL}, 
                                {COL_LAST_NAME_DATIVE}, {COL_ORGANIZATION}
                FROM {TABLE_INSPECTOR_CHIEFS}
                WHERE {COL_IS_ACTIVE} = {DEFAULT_ACTIVE}
            ''')
            inspectors = cursor.fetchall()
            
            for last_name_nom, first_init, middle_init, last_name_dat, organization in inspectors:
                # Use initials as first/middle names for inspectors
                key = (last_name_nom, first_init, middle_init)
                
                # Check if already exists as director
                if key in persons_map:
                    stats['duplicates_resolved'] += 1
                    logger.info(f"Duplicate found: {last_name_nom} {first_init}.{middle_init}. - Using inspector role (priority)")
                
                # Inspector takes priority - store organization in department field
                persons_map[key] = {
                    'role': ROLE_INSPECTOR,
                    'department': organization,  # Store organization for inspectors
                    'last_name_dative': last_name_dat
                }
                stats['inspectors_migrated'] += 1
            
            # Step 3: Insert all persons into new table
            for (last_name, first_name, middle_name), data in persons_map.items():
                try:
                    cursor.execute(f'''
                        INSERT OR IGNORE INTO {TABLE_PERSONS}
                        ({COL_LAST_NAME}, {COL_FIRST_NAME}, {COL_MIDDLE_NAME}, {COL_ROLE}, 
                         {COL_DEPARTMENT}, {COL_LAST_NAME_DATIVE}, {COL_IS_ACTIVE})
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        last_name, 
                        first_name, 
                        middle_name, 
                        data['role'], 
                        data['department'], 
                        data['last_name_dative'], 
                        DEFAULT_ACTIVE
                    ))
                    stats['total_persons'] += 1
                except sqlite3.IntegrityError as e:
                    logger.warning(f"Duplicate person skipped: {last_name} {first_name} {middle_name} - {e}")
            
            conn.commit()
            
            logger.info(f"Migration completed: {stats}")
            return stats
            
        except sqlite3.Error as e:
            logger.error(f"Migration failed: {e}")
            raise
        finally:
            conn.close()
    
    def get_name_override(
        self,
        word_nominative: str,
        word_type: str,
        gender: str,
        case: str
    ) -> Optional[str]:
        """
        Get override for name declension from database.
        
        Returns the declined word value if override exists, None otherwise.
        Normalizes word_nominative to lowercase before lookup.
        
        Args:
            word_nominative: The word in nominative case (will be normalized to lowercase).
            word_type: Type of word ('surname', 'firstname', or 'patronymic').
            gender: Gender ('male' or 'female').
            case: Grammatical case ('nomn', 'gent', 'datv', 'accs', 'ablt', 'loct').
            
        Returns:
            The declined word value if override exists, None otherwise.
            
        Examples:
            >>> db = Database()
            >>> db.get_name_override("гомер", "surname", "male", "datv")
            'Гомеру'
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            cursor.execute(f'''
                SELECT {COL_WORD_VALUE} 
                FROM {TABLE_NAME_OVERRIDES}
                WHERE {COL_WORD_NOMINATIVE} = ? 
                  AND {COL_WORD_TYPE} = ?
                  AND {COL_GENDER} = ?
                  AND {COL_CASE} = ?
                  AND {COL_IS_ACTIVE} = {DEFAULT_ACTIVE}
            ''', (word_nominative.lower(), word_type, gender, case))
            
            result = cursor.fetchone()
            
            if result:
                logger.debug(
                    f"Found name override: {word_nominative} ({word_type}, {gender}, {case}) -> {result[0]}"
                )
                return result[0]
            
            return None
            
        except sqlite3.Error as e:
            logger.error(f"Failed to get name override for '{word_nominative}': {e}")
            return None
        finally:
            conn.close()
    
    def get_all_name_overrides(self, active_only: bool = True) -> List[Dict[str, Any]]:
        """
        Retrieve all name overrides from the database.
        
        Args:
            active_only: If True, only return active overrides. Defaults to True.
            
        Returns:
            List of dictionaries containing override data with keys:
            'word_nominative', 'word_value', 'word_type', 'gender', 'case', 'is_active'.
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            query = f'''
                SELECT {COL_WORD_NOMINATIVE}, {COL_WORD_VALUE}, {COL_WORD_TYPE}, 
                       {COL_GENDER}, {COL_CASE}, {COL_IS_ACTIVE}
                FROM {TABLE_NAME_OVERRIDES}
            '''
            if active_only:
                query += f' WHERE {COL_IS_ACTIVE} = {DEFAULT_ACTIVE}'
            
            cursor.execute(query)
            rows = cursor.fetchall()
            
            return [
                {
                    'word_nominative': row[0],
                    'word_value': row[1],
                    'word_type': row[2],
                    'gender': row[3],
                    'case': row[4],
                    'is_active': bool(row[5])
                }
                for row in rows
            ]
        except sqlite3.Error as e:
            logger.error(f"Failed to retrieve name overrides: {e}")
            return []
        finally:
            conn.close()
    
    def set_manual_declension(
        self,
        person_id: int,
        case_type: str,
        custom_value: str
    ) -> bool:
        """
        Set or update a manual declension for a person in a specific case.
        
        Args:
            person_id: ID of the person in the persons table.
            case_type: Grammatical case ('nomn', 'gent', 'datv', 'accs').
            custom_value: The manually entered declined form.
            
        Returns:
            True if the declension was set successfully, False otherwise.
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            cursor.execute(f'''
                INSERT INTO {TABLE_MANUAL_DECLENSIONS} 
                ({COL_PERSON_ID}, {COL_CASE_TYPE}, {COL_CUSTOM_VALUE})
                VALUES (?, ?, ?)
                ON CONFLICT({COL_PERSON_ID}, {COL_CASE_TYPE}) 
                DO UPDATE SET {COL_CUSTOM_VALUE} = excluded.{COL_CUSTOM_VALUE}
            ''', (person_id, case_type, custom_value))
            
            conn.commit()
            logger.info(f"Set manual declension for person_id={person_id}, case={case_type}: {custom_value}")
            return True
            
        except sqlite3.Error as e:
            logger.error(f"Failed to set manual declension: {e}")
            return False
        finally:
            conn.close()
    
    def get_manual_declension(
        self,
        person_id: int,
        case_type: str
    ) -> Optional[str]:
        """
        Get a manual declension for a person in a specific case.
        
        Args:
            person_id: ID of the person in the persons table.
            case_type: Grammatical case ('nomn', 'gent', 'datv', 'accs').
            
        Returns:
            The custom declined form if it exists, None otherwise.
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            cursor.execute(f'''
                SELECT {COL_CUSTOM_VALUE}
                FROM {TABLE_MANUAL_DECLENSIONS}
                WHERE {COL_PERSON_ID} = ? 
                  AND {COL_CASE_TYPE} = ?
                  AND {COL_IS_ACTIVE} = {DEFAULT_ACTIVE}
            ''', (person_id, case_type))
            
            result = cursor.fetchone()
            return result[0] if result else None
            
        except sqlite3.Error as e:
            logger.error(f"Failed to get manual declension: {e}")
            return None
        finally:
            conn.close()
    
    def get_all_manual_declensions(
        self,
        person_id: int
    ) -> Dict[str, str]:
        """
        Get all manual declensions for a person.
        
        Args:
            person_id: ID of the person in the persons table.
            
        Returns:
            Dictionary mapping case types to their custom values.
            Example: {'nomn': 'Иванов', 'datv': 'Иванову', ...}
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            cursor.execute(f'''
                SELECT {COL_CASE_TYPE}, {COL_CUSTOM_VALUE}
                FROM {TABLE_MANUAL_DECLENSIONS}
                WHERE {COL_PERSON_ID} = ?
                  AND {COL_IS_ACTIVE} = {DEFAULT_ACTIVE}
            ''', (person_id,))
            
            rows = cursor.fetchall()
            return {row[0]: row[1] for row in rows}
            
        except sqlite3.Error as e:
            logger.error(f"Failed to get manual declensions: {e}")
            return {}
        finally:
            conn.close()
    
    def delete_manual_declension(
        self,
        person_id: int,
        case_type: Optional[str] = None
    ) -> bool:
        """
        Delete manual declension(s) for a person.
        
        Args:
            person_id: ID of the person in the persons table.
            case_type: If specified, delete only this case. If None, delete all.
            
        Returns:
            True if deletion was successful, False otherwise.
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            if case_type:
                cursor.execute(f'''
                    DELETE FROM {TABLE_MANUAL_DECLENSIONS}
                    WHERE {COL_PERSON_ID} = ? AND {COL_CASE_TYPE} = ?
                ''', (person_id, case_type))
            else:
                cursor.execute(f'''
                    DELETE FROM {TABLE_MANUAL_DECLENSIONS}
                    WHERE {COL_PERSON_ID} = ?
                ''', (person_id,))
            
            conn.commit()
            logger.info(f"Deleted manual declensions for person_id={person_id}, case={case_type or 'all'}")
            return True
            
        except sqlite3.Error as e:
            logger.error(f"Failed to delete manual declension: {e}")
            return False
        finally:
            conn.close()
    
    def add_shipment(
        self,
        tracking_number: str,
        recipient_name: str,
        address: str,
        status: str = 'Отправлено',
        sent_date: Optional[str] = None
    ) -> Optional[int]:
        """
        Add a new shipment to the database.
        
        Args:
            tracking_number: Russia Post tracking number (ШПИ)
            recipient_name: Full name of recipient
            address: Delivery address
            status: Current shipment status
            sent_date: Date when shipment was sent (format: DD.MM.YYYY)
        
        Returns:
            ID of created shipment or None if failed
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            cursor.execute('''
                INSERT INTO shipments (tracking_number, recipient_name, address, status, sent_date)
                VALUES (?, ?, ?, ?, ?)
            ''', (tracking_number, recipient_name, address, status, sent_date))
            
            conn.commit()
            shipment_id = cursor.lastrowid
            logger.info(f"Added shipment {tracking_number} with ID {shipment_id}")
            return shipment_id
            
        except sqlite3.IntegrityError:
            logger.error(f"Shipment with tracking number {tracking_number} already exists")
            return None
        except sqlite3.Error as e:
            logger.error(f"Failed to add shipment: {e}")
            return None
        finally:
            conn.close()
    
    def get_shipment_by_tracking(self, tracking_number: str) -> Optional[Dict]:
        """
        Get shipment by tracking number.
        
        Args:
            tracking_number: Russia Post tracking number
        
        Returns:
            Dictionary with shipment data or None if not found
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            cursor.execute('''
                SELECT id, tracking_number, recipient_name, address, status, sent_date, delivery_date
                FROM shipments
                WHERE tracking_number = ?
            ''', (tracking_number,))
            
            row = cursor.fetchone()
            if row:
                return {
                    'id': row[0],
                    'tracking_number': row[1],
                    'recipient_name': row[2],
                    'address': row[3],
                    'status': row[4],
                    'sent_date': row[5],
                    'delivery_date': row[6]
                }
            return None
            
        except sqlite3.Error as e:
            logger.error(f"Failed to get shipment: {e}")
            return None
        finally:
            conn.close()
    
    def get_all_shipments(self, status_filter: Optional[str] = None) -> List[Dict]:
        """
        Get all shipments, optionally filtered by status.
        
        Args:
            status_filter: Optional status to filter by
        
        Returns:
            List of shipment dictionaries
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            if status_filter:
                cursor.execute('''
                    SELECT id, tracking_number, recipient_name, address, status, sent_date, delivery_date
                    FROM shipments
                    WHERE status = ?
                    ORDER BY sent_date DESC
                ''', (status_filter,))
            else:
                cursor.execute('''
                    SELECT id, tracking_number, recipient_name, address, status, sent_date, delivery_date
                    FROM shipments
                    ORDER BY sent_date DESC
                ''')
            
            rows = cursor.fetchall()
            return [{
                'id': row[0],
                'tracking_number': row[1],
                'recipient_name': row[2],
                'address': row[3],
                'status': row[4],
                'sent_date': row[5],
                'delivery_date': row[6]
            } for row in rows]
            
        except sqlite3.Error as e:
            logger.error(f"Failed to get shipments: {e}")
            return []
        finally:
            conn.close()
    
    def update_shipment_status(
        self,
        tracking_number: str,
        status: str,
        delivery_date: Optional[str] = None
    ) -> bool:
        """
        Update shipment status.
        
        Args:
            tracking_number: Russia Post tracking number
            status: New status
            delivery_date: Optional delivery date (format: DD.MM.YYYY)
        
        Returns:
            True if successful, False otherwise
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            if delivery_date:
                cursor.execute('''
                    UPDATE shipments
                    SET status = ?, delivery_date = ?
                    WHERE tracking_number = ?
                ''', (status, delivery_date, tracking_number))
            else:
                cursor.execute('''
                    UPDATE shipments
                    SET status = ?
                    WHERE tracking_number = ?
                ''', (status, tracking_number))
            
            conn.commit()
            logger.info(f"Updated shipment {tracking_number} to status: {status}")
            return True
            
        except sqlite3.Error as e:
            logger.error(f"Failed to update shipment status: {e}")
            return False
        finally:
            conn.close()
    
    def add_shipment_event(
        self,
        tracking_number: str,
        event_date: str,
        event_type: str,
        location: str
    ) -> bool:
        """
        Add tracking event for a shipment.
        
        Args:
            tracking_number: Russia Post tracking number
            event_date: Date and time of event (format: DD.MM.YYYY HH:MM)
            event_type: Type of event (e.g., "Принято", "В пути")
            location: Location where event occurred
        
        Returns:
            True if successful, False otherwise
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            cursor.execute('SELECT id FROM shipments WHERE tracking_number = ?', (tracking_number,))
            row = cursor.fetchone()
            if not row:
                logger.error(f"Shipment {tracking_number} not found")
                return False
            
            shipment_id = row[0]
            
            cursor.execute('''
                INSERT INTO shipment_events (shipment_id, event_date, event_type, location)
                VALUES (?, ?, ?, ?)
            ''', (shipment_id, event_date, event_type, location))
            
            conn.commit()
            logger.info(f"Added event for shipment {tracking_number}: {event_type}")
            return True
            
        except sqlite3.Error as e:
            logger.error(f"Failed to add shipment event: {e}")
            return False
        finally:
            conn.close()
    
    def get_shipment_events(self, tracking_number: str) -> List[Dict]:
        """
        Get all tracking events for a shipment.
        
        Args:
            tracking_number: Russia Post tracking number
        
        Returns:
            List of event dictionaries
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            cursor.execute('''
                SELECT se.event_date, se.event_type, se.location
                FROM shipment_events se
                JOIN shipments s ON se.shipment_id = s.id
                WHERE s.tracking_number = ?
                ORDER BY se.event_date DESC
            ''', (tracking_number,))
            
            rows = cursor.fetchall()
            return [{
                'date': row[0],
                'type': row[1],
                'location': row[2]
            } for row in rows]
            
        except sqlite3.Error as e:
            logger.error(f"Failed to get shipment events: {e}")
            return []
        finally:
            conn.close()
    
    # ====================================================================================
    # MOSEDO Workflows CRUD Operations
    # ====================================================================================
    
    def add_workflow(
        self, 
        name: str, 
        steps_json: str,
        description: Optional[str] = None
    ) -> Optional[int]:
        """
        Add a new MOSEDO workflow to the database.
        
        Args:
            name: Workflow name (must be unique)
            steps_json: JSON string of workflow steps
            description: Optional workflow description
        
        Returns:
            Workflow ID if successful, None otherwise
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            created_date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            cursor.execute('''
                INSERT INTO mosedo_workflows (name, steps_json, description, created_date, is_active)
                VALUES (?, ?, ?, ?, 1)
            ''', (name, steps_json, description, created_date))
            
            workflow_id = cursor.lastrowid
            conn.commit()
            
            logger.info(f"✓ Workflow added: {name} (ID: {workflow_id})")
            return workflow_id
            
        except sqlite3.IntegrityError:
            logger.error(f"✗ Workflow name already exists: {name}")
            return None
        except sqlite3.Error as e:
            logger.error(f"✗ Failed to add workflow: {e}")
            return None
        finally:
            conn.close()
    
    def get_workflow(self, workflow_id: int) -> Optional[Dict]:
        """
        Get a workflow by ID.
        
        Args:
            workflow_id: Workflow ID
        
        Returns:
            Workflow dictionary if found, None otherwise
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            cursor.execute('''
                SELECT id, name, steps_json, description, created_date, is_active
                FROM mosedo_workflows
                WHERE id = ?
            ''', (workflow_id,))
            
            row = cursor.fetchone()
            if not row:
                logger.warning(f"Workflow {workflow_id} not found")
                return None
            
            return {
                'id': row[0],
                'name': row[1],
                'steps_json': row[2],
                'description': row[3],
                'created_date': row[4],
                'is_active': bool(row[5])
            }
            
        except sqlite3.Error as e:
            logger.error(f"Failed to get workflow {workflow_id}: {e}")
            return None
        finally:
            conn.close()
    
    def get_all_workflows(self, active_only: bool = True) -> List[Dict]:
        """
        Get all workflows from database.
        
        Args:
            active_only: Only return active workflows (default True)
        
        Returns:
            List of workflow dictionaries
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            if active_only:
                cursor.execute('''
                    SELECT id, name, steps_json, description, created_date, is_active
                    FROM mosedo_workflows
                    WHERE is_active = 1
                    ORDER BY created_date DESC
                ''')
            else:
                cursor.execute('''
                    SELECT id, name, steps_json, description, created_date, is_active
                    FROM mosedo_workflows
                    ORDER BY created_date DESC
                ''')
            
            rows = cursor.fetchall()
            workflows = [{
                'id': row[0],
                'name': row[1],
                'steps_json': row[2],
                'description': row[3],
                'created_date': row[4],
                'is_active': bool(row[5])
            } for row in rows]
            
            logger.info(f"Retrieved {len(workflows)} workflows (active_only={active_only})")
            return workflows
            
        except sqlite3.Error as e:
            logger.error(f"Failed to get workflows: {e}")
            return []
        finally:
            conn.close()
    
    def update_workflow(
        self,
        workflow_id: int,
        name: Optional[str] = None,
        steps_json: Optional[str] = None,
        description: Optional[str] = None,
        is_active: Optional[bool] = None
    ) -> bool:
        """
        Update a workflow.
        
        Args:
            workflow_id: Workflow ID
            name: New name (optional)
            steps_json: New steps JSON (optional)
            description: New description (optional)
            is_active: Active status (optional)
        
        Returns:
            True if updated successfully
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            # Build update query dynamically based on provided parameters
            updates = []
            params = []
            
            if name is not None:
                updates.append("name = ?")
                params.append(name)
            if steps_json is not None:
                updates.append("steps_json = ?")
                params.append(steps_json)
            if description is not None:
                updates.append("description = ?")
                params.append(description)
            if is_active is not None:
                updates.append("is_active = ?")
                params.append(1 if is_active else 0)
            
            if not updates:
                logger.warning("No fields to update")
                return False
            
            params.append(workflow_id)
            query = f"UPDATE mosedo_workflows SET {', '.join(updates)} WHERE id = ?"
            
            cursor.execute(query, params)
            conn.commit()
            
            if cursor.rowcount > 0:
                logger.info(f"✓ Workflow {workflow_id} updated")
                return True
            else:
                logger.warning(f"Workflow {workflow_id} not found")
                return False
            
        except sqlite3.IntegrityError:
            logger.error(f"✗ Workflow name already exists: {name}")
            return False
        except sqlite3.Error as e:
            logger.error(f"✗ Failed to update workflow: {e}")
            return False
        finally:
            conn.close()
    
    def delete_workflow(self, workflow_id: int) -> bool:
        """
        Delete a workflow (soft delete - sets is_active to 0).
        
        Args:
            workflow_id: Workflow ID
        
        Returns:
            True if deleted successfully
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            cursor.execute('''
                UPDATE mosedo_workflows
                SET is_active = 0
                WHERE id = ?
            ''', (workflow_id,))
            
            conn.commit()
            
            if cursor.rowcount > 0:
                logger.info(f"✓ Workflow {workflow_id} deactivated")
                return True
            else:
                logger.warning(f"Workflow {workflow_id} not found")
                return False
            
        except sqlite3.Error as e:
            logger.error(f"✗ Failed to delete workflow: {e}")
            return False
        finally:
            conn.close()


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    db = Database()
    
    logger.info("=== Department List ===")
    departments = db.get_all_departments()
    for dept in departments:
        logger.info(f"{dept['id']}: {dept['name']} ({dept['short_name']})")
