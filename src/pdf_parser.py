"""
PDF Parser module for extracting citizen information from OATI documents.

This module provides functionality to parse PDF documents containing citizen
appeals and extract structured information including personal data, departments,
resolutions, and recipients.
"""

import re
import logging
import PyPDF2
import fitz
import pymorphy3
from pathlib import Path
from typing import Dict, Optional, List

try:
    from src.config_loader import ConfigLoader
except ModuleNotFoundError:
    from config_loader import ConfigLoader

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)


class PDFParser:
    """
    Parser for extracting structured data from OATI PDF documents.
    
    This class handles PDF text extraction, citizen information parsing,
    department extraction, and recipient identification from resolution documents.
    """
    
    def __init__(self):
        """Initialize PDF parser with text storage and morphological analyzer."""
        self.text = ""
        self.morph = pymorphy3.MorphAnalyzer()
        self.config_loader = ConfigLoader()
        
        # Load configurations
        try:
            self.ignored_recipients_config = self.config_loader.load('ignored_recipients')
            self.keywords_config = self.config_loader.load('keywords')
            logger.info("Configuration files loaded successfully")
        except Exception as e:
            logger.error(f"Failed to load configuration files: {str(e)}")
            raise
        
    def compress_pdf(self, input_path: str, output_path: str = None) -> str:
        """
        Compress PDF file to reduce file size.
        
        Args:
            input_path: Path to input PDF file
            output_path: Path for compressed output file (optional)
            
        Returns:
            Path to compressed PDF file
            
        Raises:
            Exception: If compression fails
        """
        if output_path is None:
            output_path = input_path.replace('.pdf', '_compressed.pdf')
        
        try:
            doc = fitz.open(input_path)
            doc.save(output_path, garbage=4, deflate=True, clean=True)
            doc.close()
            logger.info(f"PDF compressed successfully: {output_path}")
            return output_path
        except Exception as e:
            logger.error(f"Failed to compress PDF {input_path}: {str(e)}")
            raise
    
    def extract_text(self, pdf_path: str) -> str:
        """
        Extract text content from PDF file.
        
        Args:
            pdf_path: Path to PDF file
            
        Returns:
            Extracted text content
            
        Raises:
            FileNotFoundError: If PDF file does not exist
            ValueError: If PDF file is empty or corrupted
            Exception: If text extraction fails
        """
        # Validate input
        if not pdf_path or not isinstance(pdf_path, str):
            raise ValueError("Invalid PDF path provided")
        
        pdf_file = Path(pdf_path)
        if not pdf_file.exists():
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")
        
        if pdf_file.stat().st_size == 0:
            raise ValueError(f"PDF file is empty: {pdf_path}")
        
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                
                # Check if PDF has pages
                if len(pdf_reader.pages) == 0:
                    raise ValueError(f"PDF file has no pages: {pdf_path}")
                
                text = ""
                for page_num, page in enumerate(pdf_reader.pages, 1):
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + "\n"
                    logger.debug(f"Extracted text from page {page_num}")
                
                if not text.strip():
                    raise ValueError(f"No text could be extracted from PDF: {pdf_path}")
                
                self.text = text
                logger.info(f"Successfully extracted text from PDF: {pdf_path}")
                return text
        except PyPDF2.errors.PdfReadError as e:
            error_msg = f"PDF file is corrupted or encrypted: {str(e)}"
            logger.error(error_msg)
            raise ValueError(error_msg)
        except Exception as e:
            error_msg = f"Error reading PDF file: {str(e)}"
            logger.error(error_msg)
            raise Exception(error_msg)
    
    def extract_citizen_info(self, text: str = None) -> Dict[str, str]:
        """
        Extract citizen personal information from PDF text.
        
        Parses text to extract full name, email, portal ID, OATI number,
        date, resolution text, and law part references.
        
        Args:
            text: Text to parse (optional, uses stored text if not provided)
            
        Returns:
            Dictionary containing extracted citizen information with keys:
            - full_name: Complete name in format "Last First Middle"
            - last_name: Surname
            - first_name: Given name
            - middle_name: Patronymic
            - email: Email address
            - portal_id: Portal identifier
            - oati_number: OATI document number
            - oati_date: Document date
            - resolution: Resolution text
            - law_part: Referenced law article part
            - departments: List of mentioned departments
        """
        if text is None:
            text = self.text
            
        result = {
            'full_name': '',
            'last_name': '',
            'first_name': '',
            'middle_name': '',
            'email': '',
            'portal_id': '',
            'oati_number': '',
            'oati_date': '',
            'resolution': '',
            'law_part': '',
            'departments': [],
        }
        
        # Extract full name (FIO pattern) - handles newlines after "ФИО:"
        fio_pattern = r'ФИО:[\s\n]+([А-ЯЁ][А-ЯЁа-яё]+)\s+([А-ЯЁ][А-ЯЁа-яё]+)\s+([А-ЯЁ][А-ЯЁа-яё]+)'
        fio_match = re.search(fio_pattern, text)
        if fio_match:
            last_name = fio_match.group(1).capitalize()
            first_name = fio_match.group(2).capitalize()
            middle_name = fio_match.group(3).capitalize()
            
            result['last_name'] = last_name
            result['first_name'] = first_name
            result['middle_name'] = middle_name
            result['full_name'] = f"{last_name} {first_name} {middle_name}"
            logger.debug(f"Extracted citizen name: {result['full_name']}")
        else:
            logger.warning("Could not extract citizen name from text")
        
        # Extract email address - handles newlines
        email_pattern = r'Электронный адрес:[\s\n]+([\w\.-]+@[\w\.-]+\.\w+)'
        email_match = re.search(email_pattern, text)
        if email_match:
            result['email'] = email_match.group(1)
            logger.debug(f"Extracted email: {result['email']}")
        
        # Extract portal ID
        portal_id_pattern = r'идентификатор:\s*(\d+)'
        portal_id_match = re.search(portal_id_pattern, text)
        if portal_id_match:
            result['portal_id'] = portal_id_match.group(1)
            logger.debug(f"Extracted portal ID: {result['portal_id']}")
        
        # Extract OATI number and date - multiple patterns with priority order
        # PRIORITY 1: "Документ зарегистрирован № 01-21-П-8715/25 от 06.10.2025 (ОАТИ)" (document card - MOST RELIABLE)
        oati_pattern1 = r'Документ зарегистрирован\s*№\s*(01-21-П-\d+(?:/\d+)?(?:-\d+)?)\s+от\s+(\d{2}\.\d{2}\.\d{4})\s+\(ОАТИ\)'
        # PRIORITY 2: "№ 01-21-П-8715/25 от 06.10.2025" (standard format with "от")
        oati_pattern2 = r'№\s*(01-21-П-\d+(?:/\d+)?(?:-\d+)?)\s+от\s+(\d{2}\.\d{2}\.\d{4})'
        # PRIORITY 3 (FALLBACK): "К № 01-21-П-8715/25 06.10.2025" (resolution header - less reliable, no "от")
        oati_pattern3 = r'К\s*№\s*(01-21-П-\d+(?:/\d+)?(?:-\d+)?)\s+(\d{2}\.\d{2}\.\d{4})'
        
        # Try patterns in priority order: document card → standard → resolution
        oati_match = re.search(oati_pattern1, text)
        date_source = None
        if oati_match:
            date_source = "document card (карточка документа)"
        else:
            oati_match = re.search(oati_pattern2, text)
            if oati_match:
                date_source = "standard format (стандартный формат)"
            else:
                oati_match = re.search(oati_pattern3, text)
                if oati_match:
                    date_source = "resolution header (резолюция) - LESS RELIABLE"
        
        if oati_match:
            raw_number = oati_match.group(1)
            result['oati_number'] = self.normalize_document_number(raw_number)
            result['oati_date'] = oati_match.group(2)
            logger.debug(f"Extracted OATI number: {result['oati_number']}, date: {result['oati_date']} from {date_source}")
        else:
            logger.warning("Could not extract OATI number and date from any pattern")
        
        # Extract resolution text
        resolution_pattern = r'(В соответствии с ч\.\s*\d+\s*ст\.8.*?Федерации[^\n]*)'
        resolution_match = re.search(resolution_pattern, text, re.DOTALL)
        if resolution_match:
            result['resolution'] = resolution_match.group(1).strip()
            
            # DEPRECATED: law_part is now calculated automatically in main.py
            # based on recipient classification (see RecipientsClassification.calculate_law_part())
            # Leaving law_part empty to avoid stale values
            logger.debug("Law part will be calculated automatically based on recipients")
        
        logger.info("Citizen information extraction completed")
        return result
    
    @staticmethod
    def normalize_document_number(doc_number: str) -> str:
        """
        Normalize document number by removing extra spaces around delimiters.
        
        Converts: "01-21-П- 10123/25" -> "01-21-П-10123/25"
        Converts: "01 - 21- П -8902 / 25-1" -> "01-21-П-8902/25-1"
        
        Supports optional suffix: "01-21-П-8902/25-1"
        
        Args:
            doc_number: Document number string (may contain extra spaces)
            
        Returns:
            Normalized document number without extra spaces
        """
        if not doc_number:
            return doc_number
        
        # Remove all whitespace
        normalized = re.sub(r'\s+', '', doc_number)
        
        # Validate format: XX-XX-П-XXXXX/XX or XX-XX-П-XXXXX or XX-XX-П-XXXXX/XX-X
        # Matches the extraction pattern: 01-21-П-\d+(?:/\d+)?(?:-\d+)?
        valid_pattern = r'^\d{2}-\d{2}-П-\d+(?:/\d+)?(?:-\d+)?$'
        if re.match(valid_pattern, normalized):
            logger.debug(f"Document number normalized: '{doc_number}' -> '{normalized}'")
            return normalized
        else:
            logger.warning(f"Document number format invalid after normalization: '{normalized}'")
            return doc_number  # Return original if validation fails
    
    def extract_departments(self, text: str = None, database = None) -> List[str]:
        """
        Extract department names from PDF text, primarily from resolution section.
        
        ENHANCED STRATEGY (v3.9.0):
        1. Extract ALL director surnames from resolutions (Волков, Урожаева, etc.)
        2. Match surnames to departments via database lookup (100% accurate)
        3. Fallback: regex patterns for full department names if DB matching fails
        
        This approach ensures ALL departments are found, even when only director
        names are mentioned (common in ОАТИ resolutions).
        
        Args:
            text: Text to parse (optional, uses stored text if not provided)
            database: Database instance for director surname matching
            
        Returns:
            List of extracted department names
        """
        if text is None:
            text = self.text
        
        departments = []
        
        # STEP 1: Try director surname matching first (most reliable method)
        if database:
            logger.info("STEP 1: Extracting director surnames from resolutions for DB matching")
            surnames = self.extract_surnames_from_resolutions(text)
            
            if surnames:
                logger.info(f"Found {len(surnames)} director surnames: {surnames}")
                
                for surname in surnames:
                    dept_results = database.find_departments_by_director_surname(surname)
                    for dept in dept_results:
                        dept_name = dept.get('name') or dept.get('department')
                        if dept_name and dept_name not in departments:
                            departments.append(dept_name)
                            logger.info(f"✓ Matched surname '{surname}' → {dept_name}")
                
                if departments:
                    logger.info(f"✓ Director matching found {len(departments)} departments")
                    return departments
                else:
                    logger.warning("Director matching found no departments in database")
            else:
                logger.info("No director surnames extracted, proceeding to regex fallback")
        
        # STEP 2: Fallback to regex patterns for full department names
        logger.info("STEP 2: Using regex patterns for full department names")
        
        # Try to extract from resolution section first for better accuracy
        resolution_patterns = [
            r'(В соответствии с ч\.\s*\d+\s*ст\.8.*?направлен[оы].*?(?:Федерации|компетенции)[^\n]*(?:\n(?!\s*Вопрос)[^\n]*)*)',
            r'(В соответствии с ч\.\s*\d+\s*ст\.\s*8.*?(?:направить|для рассмотрения).*?(?:\n(?!\s*Вопрос)[^\n]*){1,5})',
        ]
        
        resolution_match = None
        for pattern in resolution_patterns:
            resolution_match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
            if resolution_match:
                break
        
        # Use resolution text if found, otherwise use entire text
        if resolution_match:
            search_text = resolution_match.group(0)
            logger.debug(f"Found resolution section ({len(search_text)} chars)")
        else:
            search_text = text
            logger.warning("Resolution section NOT found, searching entire document")
        
        text_normalized = ' '.join(search_text.split())
        
        # Department name patterns (ordered by specificity)
        dept_patterns = [
            (r'(Государственн(?:ая|ой) жилищн(?:ая|ой) инспекци[ияюе][^\.]*город[ае] Москв[ые])', 'GZhI pattern'),
            (r'(Департамент[ае]? [^\.]+?(?:город[ае] Москв[ые]|Москв[ые]))', 'Department pattern'),
            (r'(Комитет[ае]? [^\.]+?(?:город[ае] Москв[ые]|Москв[ые]))', 'Committee pattern'),
            (r'(префектур[ауеы] [А-ЯЁа-яё\-]+\s+административного округа[^\.]*?(?:город[ае] Москв[ые]|Москв[ые]))', 'Prefecture pattern'),
            (r'(ГЖИ(?:\s+город[ае]\s+Москв[ые])?)', 'GZhI short pattern'),
        ]
        
        for pattern, pattern_name in dept_patterns:
            matches = re.findall(pattern, text_normalized, re.IGNORECASE)
            logger.debug(f"{pattern_name}: found {len(matches)} match(es)")
            
            for match in matches:
                dept = re.sub(r'\s+', ' ', match.strip())
                dept = re.sub(r'\s*-\s*', '-', dept)
                dept = re.sub(r'\s*\.\s*$', '', dept)
                dept = re.sub(r'\s+города Москвы$', ' города Москвы', dept)
                
                if dept and dept not in departments and len(dept) > 15:
                    departments.append(dept)
                    logger.info(f"Extracted department via regex: {dept}")
        
        if not departments:
            logger.warning("NO departments extracted by any method!")
        else:
            logger.info(f"Successfully extracted {len(departments)} departments total")
        
        return departments
    
    def extract_surnames_from_resolutions(self, text: str = None) -> List[str]:
        """
        Extract surnames from resolution section when department names are not found.
        
        This method extracts surnames in dative case (e.g., "Урожаевой", "Слободчикову")
        from resolution text. These surnames can then be matched against the database
        to identify departments through their directors.
        
        Args:
            text: Text to parse (optional, uses stored text if not provided)
            
        Returns:
            List of extracted surnames in dative case
        """
        if text is None:
            text = self.text
        
        # Extract resolution blocks
        resolution_pattern = r'(В соответствии с ч\.\s*\d+\s*ст\.8.*?(?:Федерации|ОАТИ|направить|поручить)[^\n]*(?:\n(?!\s*Вопрос)[^\n]*)*)'
        resolution_blocks = re.findall(resolution_pattern, text, re.DOTALL | re.IGNORECASE)
        
        if not resolution_blocks:
            # Try to find resolution header format
            header_pattern = r'((?:Вопрос \d+:.*?)?(?:[А-ЯЁ][а-яёА-ЯЁ-]+\s+[А-ЯЁ]\.\s*[А-ЯЁ]\..*?(?:В соответствии|направляем)))'
            resolution_blocks = re.findall(header_pattern, text, re.DOTALL | re.IGNORECASE)
        
        if not resolution_blocks:
            logger.warning("No resolution blocks found for surname extraction")
            return []
        
        surnames = []
        
        # Pattern for surname with initials in dative case
        # Examples: "Урожаевой Ю.В.", "Слободчикову А.О."
        name_pattern = r'([А-ЯЁ][а-яёА-ЯЁ-]+(?:ой|ому|ову|еву|ину|ыну|ко|енко|их|ых))\s+([А-ЯЁ])\.\s*([А-ЯЁ])\.'
        
        for block in resolution_blocks:
            for match in re.finditer(name_pattern, block):
                surname = match.group(1)
                
                # Validate as surname using morphological analysis
                is_valid_surname = False
                parsed = self.morph.parse(surname.lower())
                for p in parsed:
                    if 'Surn' in p.tag or 'Name' in p.tag:
                        is_valid_surname = True
                        break
                
                # Additional validation for dative case endings
                if not is_valid_surname:
                    surname_lower = surname.lower()
                    is_dative_ending = surname_lower.endswith((
                        'ому', 'ову', 'еву', 'ину', 'ыну', 'кому', 'скому', 'цкому', 'зькому',
                        'ой', 'овой', 'евой', 'ской', 'иной',
                        'ью', 'ей', 'ну', 'ру', 'лю', 'рю', 'цу', 'чу', 'шу', 'щу', 'жу', 'ху',
                        'ко', 'енко', 'ых', 'их'
                    ))
                    
                    if is_dative_ending and surname[0].isupper():
                        is_valid_surname = True
                
                if is_valid_surname and surname not in surnames:
                    surnames.append(surname)
                    logger.debug(f"Extracted surname from resolution: {surname}")
        
        logger.info(f"Extracted {len(surnames)} unique surnames from resolutions: {surnames}")
        return surnames
    
    def extract_portal_source(self, text: str = None) -> str:
        """
        Identify the portal source of the citizen appeal.
        
        Determines whether the appeal came from 'Nash Gorod' portal,
        'mos.ru' portal, or was submitted directly to OATI.
        
        Args:
            text: Text to parse (optional, uses stored text if not provided)
            
        Returns:
            Formatted portal source description
        """
        if text is None:
            text = self.text
        
        # Get portal source patterns from config
        nash_gorod_patterns = self.keywords_config.get('portal_sources', {}).get('nash_gorod', [])
        mos_ru_patterns = self.keywords_config.get('portal_sources', {}).get('mos_ru', [])
        
        # Check for Nash Gorod portal
        for pattern in nash_gorod_patterns:
            if re.search(pattern, text, re.IGNORECASE):
                logger.debug("Detected portal source: Nash Gorod")
                return "поступившее с портала Правительства Москвы «Наш город»"
        
        # Check for mos.ru portal
        for pattern in mos_ru_patterns:
            if re.search(pattern, text, re.IGNORECASE):
                logger.debug("Detected portal source: mos.ru")
                return "поступившее с портала Мэра и Правительства Москвы"
        
        logger.debug("No specific portal detected, defaulting to OATI")
        return "поступившее в ОАТИ"
    
    def extract_recipients_from_resolution(self, text: str = None) -> List[Dict[str, str]]:
        """
        Extract recipient names from resolution section.
        
        Parses resolution text to identify recipients by name patterns,
        filtering out executors, service roles, and ignored names based
        on configuration.
        
        Args:
            text: Text to parse (optional, uses stored text if not provided)
            
        Returns:
            List of dictionaries containing recipient information with keys:
            - last_name: Recipient surname
            - first_initial: First name initial
            - middle_initial: Middle name initial
            - full_name: Formatted full name with initials
        """
        if text is None:
            text = self.text
        
        # Extract resolution blocks
        resolution_pattern = r'(В соответствии с ч\.\s*\d+\s*ст\.8.*?(?:Федерации|ОАТИ|направить|поручить)[^\n]*(?:\n(?!\s*Вопрос)[^\n]*)*)'
        resolution_blocks = re.findall(resolution_pattern, text, re.DOTALL | re.IGNORECASE)
        
        if not resolution_blocks:
            logger.warning("No resolution blocks found in text")
            return []
        
        # Load ignored names from config
        ignored_names = self.ignored_recipients_config.get('ignored_executor_names', [])
        ignore_prefixes = self.ignored_recipients_config.get('ignore_prefixes', [])
        service_roles = self.ignored_recipients_config.get('service_roles', [])
        
        recipients = []
        
        for block in resolution_blocks:
            # Get context before the block
            lines_before = text[:text.find(block)].split('\n')[-5:]
            search_text = '\n'.join(lines_before) + '\n' + block
            
            # Name pattern: "Lastname I.I."
            name_pattern = r'([А-ЯЁ][а-яёА-ЯЁ-]+)\s+([А-ЯЁ])\.\s*([А-ЯЁ])\.'
            
            for match in re.finditer(name_pattern, search_text):
                last_name = match.group(1)
                first_initial = match.group(2)
                middle_initial = match.group(3)
                
                # Validate as surname using morphological analysis
                is_valid_surname = False
                parsed = self.morph.parse(last_name.lower())
                for p in parsed:
                    if 'Surn' in p.tag:
                        is_valid_surname = True
                        break
                
                # Additional validation for dative case endings
                if not is_valid_surname:
                    last_name_lower = last_name.lower()
                    is_dative_ending = last_name_lower.endswith((
                        'ому', 'ову', 'еву', 'ину', 'ыну', 'кому', 'скому', 'цкому', 'зькому',
                        'ой', 'овой', 'евой', 'ской', 'иной',
                        'ью', 'ей', 'ну', 'ру', 'лю', 'рю', 'цу', 'чу', 'шу', 'щу', 'жу', 'ху'
                    ))
                    is_invariant = last_name_lower.endswith(('ко', 'о', 'ых', 'их', 'е', 'енко'))
                    
                    if not (is_dative_ending or is_invariant):
                        continue
                
                # Check if name is in ignored list
                is_ignored = any(last_name.lower().startswith(ignored.lower()) for ignored in ignored_names)
                
                # Get preceding context
                preceding_text_context = search_text[max(0, match.start()-30):match.start()]
                
                # Check for ignore prefixes (e.g., '+')
                has_ignore_prefix = any(prefix in preceding_text_context[-5:] for prefix in ignore_prefixes)
                
                # Check for service role markers
                is_service_role = any(
                    re.search(role, preceding_text_context, re.IGNORECASE) 
                    for role in service_roles
                )
                
                # Filter out ignored recipients
                if not is_ignored and not has_ignore_prefix and not is_service_role and last_name[0].isupper():
                    recipients.append({
                        'last_name': last_name,
                        'first_initial': first_initial,
                        'middle_initial': middle_initial,
                        'full_name': f"{last_name} {first_initial}.{middle_initial}."
                    })
                    logger.debug(f"Found recipient: {last_name} {first_initial}.{middle_initial}.")
        
        # Remove duplicates
        unique_recipients = []
        seen = set()
        for recipient in recipients:
            key = f"{recipient['last_name']}_{recipient['first_initial']}_{recipient['middle_initial']}"
            if key not in seen:
                seen.add(key)
                unique_recipients.append(recipient)
        
        logger.info(f"Extracted {len(unique_recipients)} unique recipients from resolution")
        return unique_recipients
    
    def check_union_response_required(self, text: str = None, database=None) -> bool:
        """
        Determine if union response to citizen is required (ENHANCED v3.9.0).
        
        Union responds when ОАТИ inspector is assigned to handle the case:
        1. Find inspector surnames (role='inspector') from database (Алёшина, Кичиков, Блинов, etc.)
        2. Check if ANY inspector surname appears in resolutions
        3. Verify ОАТИ handling phrases ("рассмотреть в рамках компетенции ОАТИ" OR deadline)
        4. Exclude if "НПА" keyword present (regulatory act questions)
        
        Args:
            text: Text to parse (optional, uses stored text if not provided)
            database: Database instance for fetching inspector list
            
        Returns:
            True if union should respond to citizen, False otherwise
        """
        if text is None:
            text = self.text
        
        if not database:
            logger.warning("Union response check skipped: no database provided")
            return False
        
        # STEP 1: Get all inspectors from database (role='inspector')
        inspectors = database.get_all_persons(role='inspector')
        if not inspectors:
            logger.warning("No inspectors found in database")
            return False
        
        inspector_surnames = [p['last_name'] for p in inspectors]
        logger.info(f"Loaded {len(inspector_surnames)} inspectors for checking: {inspector_surnames[:5]}...")
        
        # STEP 2: Check for exclude keywords (НПА, etc.)
        exclude_keywords = self.keywords_config.get('union_response', {}).get('exclude_keywords', ['НПА'])
        for keyword in exclude_keywords:
            if re.search(rf'\b{re.escape(keyword)}\b', text, re.IGNORECASE):
                logger.debug(f"Union response not required: found exclude keyword '{keyword}'")
                return False
        
        # STEP 3: Search for inspector surnames in resolutions
        found_inspectors = []
        for surname in inspector_surnames:
            # Pattern: surname with initials (e.g., "Алёшиной О.В." or "Алешиной О.В.")
            pattern = rf'\b{re.escape(surname)}(?:ой|е|у|ым|ом)?\s+[А-ЯЁ]\.\s*[А-ЯЁ]\.'
            if re.search(pattern, text, re.IGNORECASE):
                found_inspectors.append(surname)
                logger.info(f"✓ Found inspector in resolution: {surname}")
        
        if not found_inspectors:
            logger.debug("Union response not required: no inspectors found in resolutions")
            return False
        
        # STEP 4: Verify ОАТИ handling phrases or deadlines
        oati_phrases = [
            r'рассмотреть\s+в\s+рамках\s+компетенции\s+ОАТИ',
            r'рассмотреть\s+в\s+установленном\s+порядке',
            r'компетенци(?:и|ей)\s+ОАТИ',
            r'срок[:\s]+\d{2}\.\d{2}\.\d{4}',  # Deadline pattern
            r'до\s+\d{2}\.\d{2}\.\d{4}'        # Alternative deadline
        ]
        
        for phrase_pattern in oati_phrases:
            # Check if phrase appears near ANY found inspector
            for inspector in found_inspectors:
                context_pattern = rf'{re.escape(inspector)}.{{0,300}}?{phrase_pattern}'
                if re.search(context_pattern, text, re.IGNORECASE | re.DOTALL):
                    logger.info(f"✓ Union response REQUIRED: {inspector} + ОАТИ handling phrase found")
                    return True
                
                # Check reverse: phrase before inspector
                reverse_pattern = rf'{phrase_pattern}.{{0,300}}?{re.escape(inspector)}'
                if re.search(reverse_pattern, text, re.IGNORECASE | re.DOTALL):
                    logger.info(f"✓ Union response REQUIRED: ОАТИ phrase + {inspector} found")
                    return True
        
        logger.debug(f"Union response not required: found inspectors {found_inspectors} but no ОАТИ handling phrases")
        return False
    
    def extract_question_text(self, text: str = None) -> str:
        """
        Extract question text from resolution where deadline is mentioned for inspector chief.
        
        This extracts the specific question/issue that the Union (OATI) should address,
        which is mentioned in the resolution section near the deadline for inspector chief.
        
        Args:
            text: Text to parse (optional, uses stored text if not provided)
            
        Returns:
            Question text from resolution, or empty string if not found
        """
        if text is None:
            text = self.text
        
        try:
            # Multiple patterns to find question text
            patterns = [
                # Pattern 1: "вопрос [text] в срок"
                r'вопрос\s+([^\.]+?)\s+(?:в\s+)?срок(?:ом)?',
                # Pattern 2: "вопрос [text] - срок"
                r'вопрос\s+([^\.]+?)\s*[-–—]\s*срок',
                # Pattern 3: "по вопросу [text]" followed by deadline keywords
                r'по\s+вопрос[уе]\s+([^\.]+?)(?:\s+(?:в\s+)?срок|\.)',
            ]
            
            for pattern in patterns:
                match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
                if match:
                    question_text = match.group(1).strip()
                    # Clean up whitespace
                    question_text = re.sub(r'\s+', ' ', question_text)
                    # Remove any trailing names or initials
                    question_text = re.sub(r'\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.[А-ЯЁ]\..*$', '', question_text)
                    # Remove trailing conjunctions and prepositions
                    question_text = re.sub(r'\s+(и|в|на|по|с|для|от)$', '', question_text, flags=re.IGNORECASE)
                    
                    if len(question_text) > 5:  # Validate minimum length
                        logger.debug(f"Extracted question from resolution: {question_text[:50]}...")
                        return question_text.strip()
            
            logger.warning("Could not extract question text from resolution")
            return ""
            
        except Exception as e:
            logger.warning(f"Error extracting question text: {e}")
            return ""
    
    def extract_questions(self, text: str = None) -> List[Dict[str, str]]:
        """
        Extract numbered questions from document.
        
        Parses text to find numbered questions and their content,
        cleaning up formatting and removing trailing names.
        
        Args:
            text: Text to parse (optional, uses stored text if not provided)
            
        Returns:
            List of dictionaries containing question data with keys:
            - number: Question number
            - text: Question content
        """
        if text is None:
            text = self.text
        
        questions = []
        
        # Pattern to match "Vopros N: content"
        question_pattern = r'Вопрос\s+(\d+):\s*(.+?)(?=Вопрос\s+\d+:|$)'
        
        for match in re.finditer(question_pattern, text, re.DOTALL):
            question_num = match.group(1)
            question_text = match.group(2).strip()
            
            # Normalize whitespace
            question_text = re.sub(r'\s+', ' ', question_text)
            # Remove trailing names (e.g., "Lastname I.I.")
            question_text = re.sub(r'\s+([А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.[А-ЯЁ]\.).*$', '', question_text)
            question_text = question_text.strip()
            
            # Filter out short or empty questions
            if question_text and len(question_text) > 15:
                questions.append({
                    'number': question_num,
                    'text': question_text
                })
                logger.debug(f"Extracted question {question_num}: {question_text[:50]}...")
        
        logger.info(f"Extracted {len(questions)} questions from document")
        return questions
    
    def parse_pdf(self, pdf_path: str, database=None) -> Dict[str, any]:
        """
        Parse PDF file and extract all relevant information.
        
        Main entry point for PDF parsing. Extracts text and all structured
        information including citizen data and departments.
        
        Args:
            pdf_path: Path to PDF file to parse
            database: Optional Database instance for fallback department lookup by director surnames
            
        Returns:
            Dictionary containing all extracted information
        """
        logger.info(f"Starting PDF parsing: {pdf_path}")
        
        text = self.extract_text(pdf_path)
        citizen_info = self.extract_citizen_info(text)
        
        # Extract departments using enhanced two-tier strategy:
        # Tier 1: Director surname matching via database (most accurate)
        # Tier 2: Regex patterns for full department names (fallback)
        departments = self.extract_departments(text, database=database)
        
        citizen_info['departments'] = departments
        
        logger.info(f"PDF parsing completed: {pdf_path}")
        return citizen_info


if __name__ == "__main__":
    # Test code for development
    parser = PDFParser()
    test_pdf = "attached_assets/13.10.2025_01-21-П-9141_25_Обращение_граждан_Ларин_А.С._1760536596725.pdf"
    
    logger.info("Starting PDF parsing test")
    result = parser.parse_pdf(test_pdf)
    
    logger.info("=== Extracted Data ===")
    logger.info(f"Full Name: {result['full_name']}")
    logger.info(f"Email: {result['email']}")
    logger.info(f"Portal ID: {result['portal_id']}")
    logger.info(f"OATI Number: {result['oati_number']}")
    logger.info(f"Date: {result['oati_date']}")
    logger.info(f"Law Part: Part {result['law_part']}")
    logger.info(f"Departments: {result['departments']}")
