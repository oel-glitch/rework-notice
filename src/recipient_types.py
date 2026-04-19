"""
Recipient types and classification module for Word document generation.

This module provides typed structures for classifying and handling different
types of document recipients (departments, prefectures, OATI inspectors).
"""

import logging
from dataclasses import dataclass
from enum import Enum
from typing import List


logger = logging.getLogger(__name__)


class RecipientRole(str, Enum):
    """
    Enum representing the role/type of document recipient.
    
    Recipients can be:
    - DEPARTMENT: Government departments (Департамент транспорта, ГЖИ, Комитет)
    - PREFECTURE: District prefectures (Префектура ЦАО, Префектура ВАО)
    - OATI: OATI organization itself (Объединение)
    """
    DEPARTMENT = "department"
    PREFECTURE = "prefecture"
    OATI = "oati"


@dataclass
class RecipientInfo:
    """
    Structured information about a document recipient.
    
    Attributes:
        name: Full official name of the recipient organization.
        role: Type of recipient (DEPARTMENT, PREFECTURE, or OATI).
        original_text: Original text from which this recipient was extracted.
    """
    name: str
    role: RecipientRole
    original_text: str = ""
    
    def __str__(self) -> str:
        """String representation showing name and role."""
        return f"{self.name} ({self.role.value})"


@dataclass
class RecipientsClassification:
    """
    Classification result containing categorized recipients.
    
    Attributes:
        departments: List of department recipients.
        prefectures: List of prefecture recipients.
        oati: List of OATI recipients (typically 0 or 1).
        all_recipients: Combined list of all recipients.
        has_prefecture: Quick check if any prefecture exists.
        has_oati: Quick check if OATI recipient exists.
        recipient_count: Total number of recipients.
        has_multiple_departments: Check if multiple departments exist.
        dominant_role: Most common recipient role type.
    """
    departments: List[RecipientInfo]
    prefectures: List[RecipientInfo]
    oati: List[RecipientInfo]
    
    @property
    def all_recipients(self) -> List[RecipientInfo]:
        """Get all recipients combined."""
        return self.departments + self.prefectures + self.oati
    
    @property
    def has_prefecture(self) -> bool:
        """Check if any prefecture recipient exists."""
        return len(self.prefectures) > 0
    
    @property
    def has_oati(self) -> bool:
        """Check if OATI recipient exists."""
        return len(self.oati) > 0
    
    @property
    def recipient_count(self) -> int:
        """Total count of all recipients."""
        return len(self.all_recipients)
    
    @property
    def has_multiple_departments(self) -> bool:
        """Check if multiple department recipients exist (excluding prefectures/OATI)."""
        return len(self.departments) > 1
    
    @property
    def has_multiple_recipients(self) -> bool:
        """Check if total recipients count is greater than 1."""
        return self.recipient_count > 1
    
    @property
    def dominant_role(self) -> RecipientRole:
        """
        Determine the dominant (most common) recipient role.
        
        Priority order if counts are equal:
        1. PREFECTURE (highest priority)
        2. OATI
        3. DEPARTMENT (default)
        
        Returns:
            RecipientRole that appears most frequently.
        """
        counts = {
            RecipientRole.DEPARTMENT: len(self.departments),
            RecipientRole.PREFECTURE: len(self.prefectures),
            RecipientRole.OATI: len(self.oati)
        }
        
        # If prefecture exists, prioritize it
        if self.has_prefecture:
            return RecipientRole.PREFECTURE
        
        # If OATI exists, prioritize it over departments
        if self.has_oati:
            return RecipientRole.OATI
        
        # Default to department
        return RecipientRole.DEPARTMENT
    
    def calculate_law_part(self) -> str:
        """
        Автоматически вычисляет часть закона (ч. 3 или ч. 4) на основе количества получателей.
        
        Бизнес-правила (ИСПРАВЛЕНО 2025-11-14):
        - ч. 3: Ровно 1 получатель (департамент, префектура или ОАТИ)
        - ч. 4: Более 1 получателя (любая комбинация)
        
        ВАЖНО: Тип получателя (префектура/департамент/ОАТИ) НЕ влияет на выбор части закона.
        Имеет значение ТОЛЬКО количество адресатов, которые должны дать ответ.
        
        Returns:
            "3" для ч. 3, "4" для ч. 4
            
        Examples:
            >>> # 1 департамент → ч. 3
            >>> recipients = RecipientsClassification([dept1], [], [])
            >>> recipients.calculate_law_part()  # "3"
            
            >>> # 1 префектура → ч. 3 (ИСПРАВЛЕНО!)
            >>> recipients = RecipientsClassification([], [prefecture], [])
            >>> recipients.calculate_law_part()  # "3"
            
            >>> # 2 департамента → ч. 4
            >>> recipients = RecipientsClassification([dept1, dept2], [], [])
            >>> recipients.calculate_law_part()  # "4"
            
            >>> # 1 департамент + 1 ОАТИ → ч. 4
            >>> recipients = RecipientsClassification([dept1], [], [oati])
            >>> recipients.calculate_law_part()  # "4"
            
            >>> # 1 префектура + 1 департамент → ч. 4
            >>> recipients = RecipientsClassification([dept1], [prefecture], [])
            >>> recipients.calculate_law_part()  # "4"
        """
        # Подсчитать ОБЩЕЕ количество получателей (департаменты + префектуры + ОАТИ)
        total_count = self.recipient_count
        
        # ПРОСТОЕ ПРАВИЛО: > 1 получателя → ч. 4, иначе → ч. 3
        if total_count > 1:
            logger.debug(f"Law part: ч. 4 ({total_count} получателей)")
            return "4"
        elif total_count == 1:
            logger.debug(f"Law part: ч. 3 (1 получатель)")
            return "3"
        else:
            # Fallback: нет получателей - ч. 3 с предупреждением
            logger.warning("No recipients found, defaulting to law part ч. 3")
            return "3"
    
    def get_combination_key(self) -> str:
        """
        Generate a unique key representing the combination of recipient types.
        
        Used for template selection matrix. Examples:
        - "single_department": Exactly 1 department, no prefecture/OATI
        - "multiple_departments": 2+ departments, no prefecture/OATI
        - "prefecture": Prefecture only (no other recipients)
        - "prefecture_single_department": Prefecture + exactly 1 department
        - "prefecture_multiple": Prefecture + 2+ departments
        - "oati": OATI only
        - "single_department_with_oati": 1 department + OATI
        - "multiple_departments_with_oati": 2+ departments + OATI
        - "prefecture_with_oati": Prefecture + OATI (with any number of departments)
        
        Returns:
            String key for template selection matrix.
        """
        # CRITICAL: Check OATI presence FIRST to avoid incorrect prefecture classification
        if self.has_oati:
            dept_count = len(self.departments)
            
            # Prefecture + OATI combinations (takes priority)
            if self.has_prefecture:
                return "prefecture_with_oati"
            
            # OATI with departments
            if dept_count == 0:
                return "oati"  # Only OATI
            elif dept_count == 1:
                return "single_department_with_oati"  # 1 department + OATI
            else:
                return "multiple_departments_with_oati"  # 2+ departments + OATI
        
        # No OATI - check prefecture combinations
        if self.has_prefecture:
            dept_count = len(self.departments)
            
            if dept_count == 0:
                return "prefecture"  # Prefecture only
            elif dept_count == 1:
                return "prefecture_single_department"  # Prefecture + 1 department
            else:
                return "prefecture_multiple"  # Prefecture + 2+ departments
        
        # No OATI, no prefecture - just departments
        if self.has_multiple_departments:
            return "multiple_departments"
        
        if len(self.departments) == 1:
            return "single_department"
        
        # Fallback for empty recipients
        logger.warning("No recipients found, using fallback combination key 'single_department'")
        return "single_department"
    
    def __str__(self) -> str:
        """String representation showing counts and breakdown."""
        return (
            f"Recipients(total={self.recipient_count}, "
            f"departments={len(self.departments)}, "
            f"prefectures={len(self.prefectures)}, "
            f"oati={len(self.oati)}, "
            f"combination='{self.get_combination_key()}')"
        )


def classify_recipient(recipient_name: str) -> RecipientInfo:
    """
    Classify a single recipient by analyzing its name pattern.
    
    Uses pattern matching to determine if a recipient is a department,
    prefecture, or OATI organization.
    
    Args:
        recipient_name: Official name of the recipient organization.
        
    Returns:
        RecipientInfo object with classified role.
        
    Examples:
        >>> classify_recipient("Префектура Центрального административного округа города Москвы")
        RecipientInfo(name='...', role=RecipientRole.PREFECTURE)
        >>> classify_recipient("Департамент транспорта города Москвы")
        RecipientInfo(name='...', role=RecipientRole.DEPARTMENT)
    """
    name_lower = recipient_name.lower()
    
    # Check for prefecture patterns
    if 'префектур' in name_lower:
        logger.debug(f"Classified as PREFECTURE: {recipient_name}")
        return RecipientInfo(
            name=recipient_name,
            role=RecipientRole.PREFECTURE,
            original_text=recipient_name
        )
    
    # Check for OATI patterns (Объединение)
    # Note: OATI is usually the sender, not recipient, but included for completeness
    if 'объединение' in name_lower and ('оати' in name_lower or 'административно-технических' in name_lower):
        logger.debug(f"Classified as OATI: {recipient_name}")
        return RecipientInfo(
            name=recipient_name,
            role=RecipientRole.OATI,
            original_text=recipient_name
        )
    
    # Default to DEPARTMENT for all other cases
    # This includes: Департамент, Комитет, ГЖИ, Управление, etc.
    logger.debug(f"Classified as DEPARTMENT: {recipient_name}")
    return RecipientInfo(
        name=recipient_name,
        role=RecipientRole.DEPARTMENT,
        original_text=recipient_name
    )


def classify_recipients(recipient_names: List[str]) -> RecipientsClassification:
    """
    Classify a list of recipient names into typed categories.
    
    Takes a list of recipient organization names and categorizes them
    into departments, prefectures, and OATI organizations.
    
    Args:
        recipient_names: List of recipient organization names.
        
    Returns:
        RecipientsClassification object with categorized recipients.
        
    Examples:
        >>> names = [
        ...     "Департамент транспорта города Москвы",
        ...     "Префектура ЦАО города Москвы"
        ... ]
        >>> result = classify_recipients(names)
        >>> print(result)
        Recipients(total=2, departments=1, prefectures=1, oati=0)
    """
    departments = []
    prefectures = []
    oati_list = []
    
    for name in recipient_names:
        if not name or not name.strip():
            logger.warning("Skipping empty recipient name")
            continue
            
        recipient = classify_recipient(name.strip())
        
        if recipient.role == RecipientRole.DEPARTMENT:
            departments.append(recipient)
        elif recipient.role == RecipientRole.PREFECTURE:
            prefectures.append(recipient)
        elif recipient.role == RecipientRole.OATI:
            oati_list.append(recipient)
    
    classification = RecipientsClassification(
        departments=departments,
        prefectures=prefectures,
        oati=oati_list
    )
    
    logger.info(f"Classified recipients: {classification}")
    
    return classification
