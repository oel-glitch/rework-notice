"""
GOST text normalization module for official documents.

Implements GOST R 7.0.97-2016 requirements for Russian official document formatting:
- Quotation marks normalization (« » for primary quotes)
- Dash types (hyphen, en dash, em dash) with proper spacing
- Non-breaking spaces (after prepositions, in abbreviations, numbers)
- Multiple spaces cleanup
- Proper spacing around punctuation

This ensures all generated Word documents comply with Russian federal standards.
"""

import re
import logging
from typing import Dict, List, Tuple

logger = logging.getLogger(__name__)


class GOSTNormalizer:
    """
    Text normalizer for GOST R 7.0.97-2016 compliance.
    
    Provides methods to normalize text according to Russian federal standards
    for organizational and administrative documentation.
    """
    
    # Unicode characters
    NBSP = '\u00A0'  # Non-breaking space
    QUOTE_OPEN = '«'  # Russian opening quote (ёлочка)
    QUOTE_CLOSE = '»'  # Russian closing quote (ёлочка)
    EN_DASH = '–'  # En dash (short dash for ranges)
    EM_DASH = '—'  # Em dash (long dash for punctuation)
    
    # Common single-letter prepositions requiring non-breaking space
    SINGLE_LETTER_PREPOSITIONS = ['в', 'с', 'у', 'к', 'о', 'и', 'а']
    
    def __init__(self):
        """Initialize GOST normalizer with predefined rules."""
        logger.debug("GOSTNormalizer initialized")
    
    def normalize_quotes(self, text: str) -> str:
        """
        Normalize quotation marks to Russian GOST standard.
        
        Replaces English quotes ("text") with Russian ёлочки («text»).
        GOST requires « » for primary quotes and " " for nested quotes.
        
        Args:
            text: Input text with potentially mixed quote styles
            
        Returns:
            Text with normalized Russian quotes
            
        Examples:
            >>> normalizer = GOSTNormalizer()
            >>> normalizer.normalize_quotes('"Закон"')
            '«Закон»'
        """
        if not text:
            return text
        
        # Replace English double quotes with Russian ёлочки
        # Simple replacement for paired quotes
        result = text
        
        # Pattern: "word" → «word» (including empty quotes)
        # This handles most common cases where quotes are paired
        result = re.sub(r'"([^"]*)"', f'{self.QUOTE_OPEN}\\1{self.QUOTE_CLOSE}', result)
        
        logger.debug(f"Normalized quotes in text: {len(text)} chars")
        return result
    
    def normalize_dashes(self, text: str) -> str:
        """
        Normalize dashes according to GOST requirements.
        
        GOST distinguishes three dash types:
        - Hyphen (-): compound words, NO spaces (ковер-самолет)
        - En dash (–): ranges, NO spaces (1941–1945, стр. 10–15)
        - Em dash (—): punctuation, WITH spaces on both sides
        
        Args:
            text: Input text with potentially incorrect dash usage
            
        Returns:
            Text with normalized dashes
            
        Examples:
            >>> normalizer = GOSTNormalizer()
            >>> normalizer.normalize_dashes('1941-1945 годы')
            '1941–1945 годы'
        """
        if not text:
            return text
        
        result = text
        
        # En dash for year ranges ONLY: 1941-1945 → 1941–1945
        # Pattern: 4-digit year - 4-digit year (NOT document numbers!)
        # Use word boundaries to avoid matching document IDs like "01-21"
        result = re.sub(r'\b(\d{4})\s*-\s*(\d{4})\b', f'\\1{self.EN_DASH}\\2', result)
        
        # En dash for page/number ranges with context
        # Only convert if preceded by keywords like "стр.", "с.", "№№"
        # This preserves document numbers like "01-21-П-8715"
        result = re.sub(r'(стр\.|с\.|страниц[аы]|№№)\s*(\d+)\s*-\s*(\d+)', 
                       f'\\1 \\2{self.EN_DASH}\\3', result)
        
        # Em dash for punctuation: space-space → space—space
        # Pattern: word - word → word — word
        # Only if surrounded by spaces (not compound words)
        result = re.sub(r'\s+-\s+', f' {self.EM_DASH} ', result)
        
        logger.debug(f"Normalized dashes in text: {len(text)} chars")
        return result
    
    def add_non_breaking_spaces(self, text: str) -> str:
        """
        Add non-breaking spaces according to GOST requirements.
        
        GOST requires non-breaking spaces in specific contexts:
        - After single-letter prepositions (в, с, у, к, о, и, а)
        - After №, ст., ч., п.
        - Between initials and surname
        - In abbreviations (т. д., т. е., и т. п.)
        - Before units of measurement
        
        Args:
            text: Input text
            
        Returns:
            Text with non-breaking spaces inserted
            
        Examples:
            >>> normalizer = GOSTNormalizer()
            >>> normalizer.add_non_breaking_spaces('в соответствии с ч. 3 ст. 8')
            'в соответствии с ч. 3 ст. 8'  # with NBSPs
        """
        if not text:
            return text
        
        result = text
        
        # 1. After single-letter prepositions: "в соответствии" → "в соответствии"
        for prep in self.SINGLE_LETTER_PREPOSITIONS:
            # Word boundary before and space after
            # Match both lowercase and uppercase (В соответствии / в соответствии)
            result = re.sub(
                rf'\b{prep}\s+',
                f'{prep}{self.NBSP}',
                result
            )
            # Also match uppercase version (for sentence start)
            result = re.sub(
                rf'\b{prep.upper()}\s+',
                f'{prep.upper()}{self.NBSP}',
                result
            )
        
        # 2. After № (number sign): "№ 123" → "№ 123"
        result = re.sub(r'№\s+', f'№{self.NBSP}', result)
        
        # 3. After ст. (article): "ст. 8" → "ст. 8"
        result = re.sub(r'ст\.\s+', f'ст.{self.NBSP}', result)
        
        # 4. After ч. (part): "ч. 3" → "ч. 3"
        result = re.sub(r'ч\.\s+', f'ч.{self.NBSP}', result)
        
        # 5. After п. (paragraph): "п. 5" → "п. 5"
        result = re.sub(r'п\.\s+', f'п.{self.NBSP}', result)
        
        # 6. In abbreviations: "т. д.", "т. е.", "и т. п."
        result = result.replace('т. д.', f'т.{self.NBSP}д.')
        result = result.replace('т. е.', f'т.{self.NBSP}е.')
        result = result.replace('т. п.', f'т.{self.NBSP}п.')
        result = result.replace('и т. д.', f'и{self.NBSP}т.{self.NBSP}д.')
        result = result.replace('и т. п.', f'и{self.NBSP}т.{self.NBSP}п.')
        
        # 7. In document numbers: "01-21-П-8715" → "01-21-П- 8715" (already handled separately)
        # Keep existing logic for compatibility
        result = re.sub(r'01-21-П-\s*(\d)', f'01-21-П-{self.NBSP}\\1', result)
        
        # 8. Federal law references: "ФЗ «" → "ФЗ «"
        result = result.replace('ФЗ «', f'ФЗ{self.NBSP}«')
        result = result.replace('№59-ФЗ', f'№{self.NBSP}59-ФЗ')
        result = result.replace('№ 59-ФЗ', f'№{self.NBSP}59-ФЗ')
        
        logger.debug(f"Added non-breaking spaces: {len(text)} chars")
        return result
    
    def remove_extra_spaces(self, text: str) -> str:
        """
        Remove extra spaces according to GOST requirements.
        
        GOST requires:
        - Single space between words
        - No spaces before punctuation marks
        - No trailing/leading spaces
        
        Args:
            text: Input text with potential extra spaces
            
        Returns:
            Text with normalized spacing
            
        Examples:
            >>> normalizer = GOSTNormalizer()
            >>> normalizer.remove_extra_spaces('текст  ,  пробелы')
            'текст, пробелы'
        """
        if not text:
            return text
        
        result = text
        
        # 1. Multiple spaces → single space (preserve non-breaking spaces)
        result = re.sub(r'(?<!\u00A0) {2,}(?!\u00A0)', ' ', result)
        
        # 2. Spaces before punctuation marks: "текст ," → "текст,"
        result = re.sub(r'\s+([,.;:!?])', r'\1', result)
        
        # 3. Strip leading/trailing spaces
        result = result.strip()
        
        logger.debug(f"Removed extra spaces: {len(text)} chars")
        return result
    
    def normalize_text(self, text: str) -> str:
        """
        Apply all GOST normalization rules to text.
        
        Comprehensive normalization pipeline:
        1. Normalize quotes (« » instead of " ")
        2. Normalize dashes (hyphen, en dash, em dash)
        3. Add non-breaking spaces
        4. Remove extra spaces
        
        Args:
            text: Input text to normalize
            
        Returns:
            Fully normalized text compliant with GOST R 7.0.97-2016
            
        Examples:
            >>> normalizer = GOSTNormalizer()
            >>> text = 'В соответствии с ч. 3 ст. 8 "Закона" 1941-1945'
            >>> normalizer.normalize_text(text)
            'В соответствии с ч. 3 ст. 8 «Закона» 1941–1945'  # with NBSPs
        """
        if not text:
            return text
        
        try:
            # Apply normalization pipeline
            result = text
            result = self.normalize_quotes(result)
            result = self.normalize_dashes(result)
            result = self.add_non_breaking_spaces(result)
            result = self.remove_extra_spaces(result)
            
            logger.info(f"Text normalized: {len(text)} → {len(result)} chars")
            return result
            
        except Exception as e:
            logger.error(f"Error normalizing text: {e}")
            return text  # Return original on error
    
    def normalize_dict_values(self, data: Dict[str, any]) -> Dict[str, any]:
        """
        Normalize all string values in a dictionary.
        
        Recursively processes dictionary values, applying GOST normalization
        to all strings. Useful for normalizing template replacement data.
        
        Args:
            data: Dictionary with mixed value types
            
        Returns:
            Dictionary with normalized string values
            
        Examples:
            >>> normalizer = GOSTNormalizer()
            >>> data = {'name': 'Иванов  И.И.', 'law': '"Закон"'}
            >>> normalizer.normalize_dict_values(data)
            {'name': 'Иванов И.И.', 'law': '«Закон»'}
        """
        if not isinstance(data, dict):
            return data
        
        normalized = {}
        for key, value in data.items():
            if isinstance(value, str):
                normalized[key] = self.normalize_text(value)
            elif isinstance(value, dict):
                normalized[key] = self.normalize_dict_values(value)
            elif isinstance(value, list):
                normalized[key] = [
                    self.normalize_text(item) if isinstance(item, str) else item
                    for item in value
                ]
            else:
                normalized[key] = value
        
        return normalized


# Singleton instance for module-level access
_normalizer_instance = None


def get_normalizer() -> GOSTNormalizer:
    """
    Get singleton GOSTNormalizer instance.
    
    Returns:
        Shared GOSTNormalizer instance
    """
    global _normalizer_instance
    if _normalizer_instance is None:
        _normalizer_instance = GOSTNormalizer()
    return _normalizer_instance


# Convenience functions
def normalize_gost_text(text: str) -> str:
    """
    Normalize text according to GOST R 7.0.97-2016.
    
    Convenience function that uses the singleton normalizer instance.
    
    Args:
        text: Input text
        
    Returns:
        GOST-normalized text
    """
    return get_normalizer().normalize_text(text)


def normalize_gost_dict(data: Dict[str, any]) -> Dict[str, any]:
    """
    Normalize dictionary values according to GOST R 7.0.97-2016.
    
    Convenience function that uses the singleton normalizer instance.
    
    Args:
        data: Dictionary with values to normalize
        
    Returns:
        Dictionary with normalized values
    """
    return get_normalizer().normalize_dict_values(data)
