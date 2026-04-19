"""
Russian name declension module.

This module provides functionality for declining Russian names (first names, last names,
and patronymics) across different grammatical cases using pymorphy3 morphological analyzer
and optional Petrovich library for enhanced accuracy.

Classes:
    NameDeclension: Main class for Russian name declension operations.
"""

import logging
from typing import TYPE_CHECKING, Dict, Optional, Tuple

import pymorphy3

from src.constants import (
    ACCUSATIVE_CASE,
    DATIVE_CASE,
    GENDER_FEMALE,
    GENDER_FEMALE_VALUE,
    GENDER_MALE,
    GENDER_MALE_VALUE,
    GENITIVE_CASE,
    INSTRUMENTAL_CASE,
    NOMINATIVE_CASE,
    PREPOSITIONAL_CASE,
    TAG_ADJECTIVE,
    TAG_FIRST_NAME,
    TAG_NOUN,
    TAG_PATRONYMIC,
    TAG_SURNAME,
)

if TYPE_CHECKING:
    from src.database import Database

try:
    from natasha import Doc, Segmenter, MorphVocab, NewsEmbedding, NewsMorphTagger
    # FORCE DISABLE Natasha Tier 3 (v3.8.0):
    # Even though import succeeds, we disable it to prevent runtime model downloads
    # Natasha models are NOT embedded - they download from internet on first use
    # This violates offline-first requirement for production deployment
    NATASHA_AVAILABLE = False  # FORCED TO FALSE - do not enable without model bundling
except ImportError:
    NATASHA_AVAILABLE = False
    Doc = None  # type: ignore
    Segmenter = None  # type: ignore
    MorphVocab = None  # type: ignore
    NewsEmbedding = None  # type: ignore
    NewsMorphTagger = None  # type: ignore


logger = logging.getLogger(__name__)


class NameDeclension:
    """
    Russian name declension handler using morphological analysis.

    This class provides methods for declining Russian names across different
    grammatical cases, detecting gender, and formatting names with proper
    salutations. It uses a TWO-TIER declension system:
    
    TIER 1: Manual overrides from database (highest priority)
    TIER 2: pymorphy3 morphological analysis

    Attributes:
        morph (pymorphy3.MorphAnalyzer): Morphological analyzer for Russian.
        database (Optional[Database]): Optional database instance for overrides.
        _override_cache (Dict[Tuple, str]): Cache for name declension overrides.
    """

    def __init__(self, database: Optional["Database"] = None) -> None:
        """
        Initialize the NameDeclension instance.

        Creates a TWO-TIER declension system:
        TIER 1: Manual declensions from database (highest priority)
        TIER 2: pymorphy3 morphological analysis
        
        Args:
            database: Optional Database instance for loading name overrides.
        """
        try:
            self.morph = pymorphy3.MorphAnalyzer()
            logger.info("✓ Tier 2: Pymorphy3 MorphAnalyzer initialized")
        except Exception as e:
            logger.error(f"Failed to initialize pymorphy3 MorphAnalyzer: {e}")
            raise
        
        logger.info("✓ Two-tier declension system active (manual overrides + pymorphy3)")
        
        self.database = database
        self._override_cache: Dict[Tuple, str] = {}
        
        if self.database is not None:
            self._load_overrides_cache()
            logger.info("✓ Tier 1: Manual declensions cache loaded")
        else:
            logger.debug("No database provided, name overrides cache will be empty")
    
    def _load_overrides_cache(self) -> None:
        """
        Load all active name overrides from database into cache.
        
        Populates the _override_cache dictionary with all active overrides
        from the database. Cache key is (word_nominative_lower, word_type, gender, case).
        """
        if self.database is None:
            logger.warning("Cannot load overrides cache: no database instance")
            return
        
        try:
            overrides = self.database.get_all_name_overrides(active_only=True)
            
            for override in overrides:
                cache_key = (
                    override['word_nominative'].lower(),
                    override['word_type'],
                    override['gender'],
                    override['case']
                )
                self._override_cache[cache_key] = override['word_value']
            
            logger.info(f"Loaded {len(self._override_cache)} name overrides into cache")
            
            if overrides:
                logger.debug(f"Sample overrides: {list(self._override_cache.items())[:3]}")
        
        except Exception as e:
            logger.error(f"Failed to load name overrides cache: {e}")
            self._override_cache = {}

    def detect_gender(
        self, first_name: str, middle_name: Optional[str] = None
    ) -> str:
        """
        Detect the gender of a person based on their first name and patronymic.

        ENHANCED v3.9.0: Patronymic-based detection takes priority, as patronymics
        are 100% unambiguous gender indicators in Russian:
        - Male: -ович, -евич, -ич (Владимирович, Сергеевич, Ильич)
        - Female: -овна, -евна, -ична (Владимировна, Сергеевна, Ильинична)

        Args:
            first_name: The person's first name.
            middle_name: The person's patronymic (middle name). Optional.

        Returns:
            Gender as a string: either 'male' or 'female'.

        Examples:
            >>> decliner = NameDeclension()
            >>> decliner.detect_gender("Елена", "Алексеевна")
            'female'
            >>> decliner.detect_gender("Дмитрий", "Владимирович")
            'male'
        """
        try:
            # PRIORITY 1: Patronymic heuristics (100% accurate for Russian)
            if middle_name and isinstance(middle_name, str):
                patronymic_lower = middle_name.lower()
                
                # Female patronymic endings
                if patronymic_lower.endswith(('овна', 'евна', 'ична', 'инична')):
                    logger.debug(f"Gender detected as female from patronymic: {middle_name}")
                    return GENDER_FEMALE_VALUE
                
                # Male patronymic endings
                if patronymic_lower.endswith(('ович', 'евич', 'ич')):
                    logger.debug(f"Gender detected as male from patronymic: {middle_name}")
                    return GENDER_MALE_VALUE
            
            # PRIORITY 2: Morphological analysis of first name
            parsed = self.morph.parse(first_name)

            for p in parsed:
                if TAG_FIRST_NAME in p.tag:
                    if GENDER_MALE in p.tag:
                        return GENDER_MALE_VALUE
                    elif GENDER_FEMALE in p.tag:
                        return GENDER_FEMALE_VALUE

            # PRIORITY 3: Morphological analysis of patronymic
            if middle_name:
                parsed_middle = self.morph.parse(middle_name)
                for p in parsed_middle:
                    if TAG_PATRONYMIC in p.tag:
                        if GENDER_MALE in p.tag:
                            return GENDER_MALE_VALUE
                        elif GENDER_FEMALE in p.tag:
                            return GENDER_FEMALE_VALUE

            # PRIORITY 4: Fallback to first name ending heuristics
            if first_name.endswith(("а", "я")):
                return GENDER_FEMALE_VALUE
            else:
                return GENDER_MALE_VALUE

        except Exception as e:
            logger.warning(
                f"Error detecting gender for '{first_name}': {e}. "
                f"Defaulting to male."
            )
            return GENDER_MALE_VALUE

    def decline_name(self, name: str, case: str = DATIVE_CASE) -> str:
        """
        Decline a single name component to the specified grammatical case.

        Args:
            name: The name to decline (first name, last name, or patronymic).
            case: The target grammatical case. Defaults to dative case.

        Returns:
            The declined name in the specified case, capitalized.

        Examples:
            >>> decliner = NameDeclension()
            >>> decliner.decline_name("Иванов", DATIVE_CASE)
            'Иванову'
        """
        if not name:
            return name

        try:
            parsed = self.morph.parse(name)

            best_parse = None
            for p in parsed:
                if (
                    TAG_FIRST_NAME in p.tag
                    or TAG_SURNAME in p.tag
                    or TAG_PATRONYMIC in p.tag
                ):
                    best_parse = p
                    break

            if best_parse is None and parsed:
                best_parse = parsed[0]

            if best_parse:
                inflected = best_parse.inflect({case})
                if inflected:
                    return inflected.word.capitalize()

            return name

        except Exception as e:
            logger.warning(f"Error declining name '{name}' to case '{case}': {e}")
            return name

    def decline_full_name(
        self,
        last_name: str,
        first_name: str,
        middle_name: Optional[str] = None,
        case: str = DATIVE_CASE,
        person_id: Optional[int] = None,
    ) -> Dict[str, str]:
        """
        Decline a full Russian name to the specified grammatical case.

        This method declines all components of a Russian name (last name, first name,
        and patronymic if provided) while respecting gender agreement. If person_id
        is provided, checks for manual declensions first before falling back to
        automatic morphological analysis.

        Args:
            last_name: The person's surname.
            first_name: The person's first name.
            middle_name: The person's patronymic. Optional.
            case: The target grammatical case. Defaults to dative.
            person_id: Optional person ID for checking manual declensions. Optional.

        Returns:
            A dictionary containing:
                - 'last_name': Declined surname
                - 'first_name': Declined first name
                - 'middle_name': Declined patronymic (or empty string)
                - 'full_name': Complete declined name as a single string
                - 'gender': Detected gender ('male' or 'female')

        Examples:
            >>> decliner = NameDeclension()
            >>> result = decliner.decline_full_name("Иванов", "Иван", "Иванович")
            >>> print(result['full_name'])
            'Иванову Ивану Ивановичу'
        """
        try:
            gender = self.detect_gender(first_name, middle_name)

            if person_id is not None and self.database is not None:
                manual_decl = self.database.get_manual_declension(person_id, case)
                if manual_decl:
                    logger.info(f"Using manual declension for person_id={person_id}, case={case}: {manual_decl}")
                    parts = manual_decl.rsplit(maxsplit=2)
                    parts.reverse()
                    return {
                        "last_name": parts[2] if len(parts) >= 3 else manual_decl,
                        "first_name": parts[1] if len(parts) >= 2 else first_name,
                        "middle_name": parts[0] if len(parts) >= 1 else (middle_name or ""),
                        "full_name": manual_decl,
                        "gender": gender,
                    }

            last_declined = self.decline_name_with_gender(
                last_name, case, gender, TAG_SURNAME
            )
            first_declined = self.decline_name_with_gender(
                first_name, case, gender, TAG_FIRST_NAME
            )
            middle_declined = (
                self.decline_name_with_gender(
                    middle_name, case, gender, TAG_PATRONYMIC
                )
                if middle_name
                else ""
            )

            return {
                "last_name": last_declined,
                "first_name": first_declined,
                "middle_name": middle_declined,
                "full_name": f"{last_declined} {first_declined} {middle_declined}".strip(),
                "gender": gender,
            }

        except Exception as e:
            logger.error(
                f"Error declining full name '{last_name} {first_name} "
                f"{middle_name or ''}': {e}"
            )
            return {
                "last_name": last_name,
                "first_name": first_name,
                "middle_name": middle_name or "",
                "full_name": f"{last_name} {first_name} {middle_name or ''}".strip(),
                "gender": GENDER_MALE_VALUE,
            }

    def is_indeclinable_surname(self, surname: str, gender: str) -> bool:
        """
        Check if a surname should remain indeclinable (unchanged) based on heuristics.
        
        This method implements morphological rules for Russian surnames that do not
        decline in certain cases, particularly:
        - Surnames ending in -ых/-их (Черных, Белых, Седых)
        - Surnames ending in -ко (Shevchenko, Petrenko - Ukrainian surnames)
        - Foreign surnames ending in consonants for females (Schmidt, Weber for women)
        - Surnames ending in vowels -о, -е, -и, -у, -ю, -э (Гюго, Дюма, Неру, Гарибальди)
        
        Args:
            surname: The surname to check.
            gender: The gender ('male' or 'female').
            
        Returns:
            True if the surname should not be declined, False otherwise.
            
        Examples:
            >>> decliner = NameDeclension()
            >>> decliner.is_indeclinable_surname("Черных", "male")
            True
            >>> decliner.is_indeclinable_surname("Шевченко", "female")
            True
            >>> decliner.is_indeclinable_surname("Иванов", "male")
            False
        """
        if not surname:
            return False
        
        surname_lower = surname.lower()
        
        # Rule 1: Surnames ending in -ых/-их are always indeclinable
        # Examples: Черных, Белых, Седых, Долгих
        if surname_lower.endswith(('ых', 'их')):
            logger.debug(f"Surname '{surname}' is indeclinable: ends with -ых/-их")
            return True
        
        # Rule 2: Ukrainian surnames are indeclinable for both genders
        # Rule 2.1: Surnames ending in -ко (most common Ukrainian pattern)
        # Examples: Шевченко, Петренко, Ткаченко, Коваленко
        if surname_lower.endswith('ко'):
            logger.debug(f"Surname '{surname}' is indeclinable: ends with -ко (Ukrainian)")
            return True
        
        # Rule 2.2: Ukrainian patronymic surnames ending in -ович/-евич/-ёвич
        # Examples: Ивашкович, Петрович, Маркевич
        # Note: Plain "-ич" is too general and can match Russian surnames, so we require the full pattern
        if surname_lower.endswith(('ович', 'евич', 'ёвич')):
            logger.debug(f"Surname '{surname}' is indeclinable: Ukrainian patronymic surname ending in -ович/-евич/-ёвич")
            return True
        
        # Rule 2.3: Ukrainian surnames ending in -енко (alternative to -ко)
        # Examples: Петренко, Макаренко, Тимошенко
        # Note: Already covered by -ко rule above, but included for clarity
        if surname_lower.endswith('енко'):
            logger.debug(f"Surname '{surname}' is indeclinable: ends with -енко (Ukrainian)")
            return True
        
        # Rule 2.3: Ukrainian surnames ending in -а/-я after sibilants/affricates (ш, щ, ч, ж, ц)
        # are indeclinable for females
        # Examples: Гаркуша (female), Кваша (female), Пилипча (female), Гайдуша (female)
        # Note: For males, these surnames also typically don't decline, but we focus on female case
        # since it's more common and problematic
        if gender == GENDER_FEMALE_VALUE:
            sibilants_and_affricates = ('ш', 'щ', 'ч', 'ж', 'ц')
            if len(surname_lower) >= 2:
                # Check if surname ends in -а or -я preceded by a sibilant/affricate
                if surname_lower[-1] in ('а', 'я') and surname_lower[-2] in sibilants_and_affricates:
                    logger.debug(
                        f"Surname '{surname}' is indeclinable: Ukrainian surname ending in "
                        f"-{surname_lower[-1]} after sibilant '{surname_lower[-2]}' (female)"
                    )
                    return True
        
        # Rule 2.5: Known indeclinable foreign surnames ending in -а/-я (minimal curated allowlist)
        # Only surnames that are KNOWN to be indeclinable in Russian are included
        # Most foreign surnames on -а do decline in Russian (Петрарка→Петрарке, Неруда→Неруде)
        # For edge cases, use the database override system
        known_indeclinable_surnames_a_ya = {
            'дюма',    # Alexandre Dumas
            'золя',    # Émile Zola
        }
        
        if surname_lower in known_indeclinable_surnames_a_ya:
            logger.debug(f"Surname '{surname}' is indeclinable: known indeclinable foreign surname on -а/-я (allowlist)")
            return True
        
        # Rule 3: Surnames ending in specific vowels that typically indicate foreign origin
        # Examples: Гюго (-о), Неру (-у), Гарибальди (-и), Руссо (-о)
        # Note: Excludes -а/-я/-ы because many Russian surnames end in these
        # (e.g., male: Суда, Колыга, Коляда; female: Иванова, Петрова)
        # Those should be handled by pymorphy3/Petrovich
        foreign_vowel_endings = ('о', 'е', 'и', 'у', 'ю', 'э')
        
        # Exception: Russian surnames ending in -ов, -ев, -ин should decline
        russian_declining_endings = ('ов', 'ев', 'ёв', 'ин', 'ын')
        ends_with_russian_suffix = any(surname_lower.endswith(ending) for ending in russian_declining_endings)
        
        if surname_lower[-1] in foreign_vowel_endings and not ends_with_russian_suffix:
            logger.debug(f"Surname '{surname}' is indeclinable: ends with foreign vowel '{surname_lower[-1]}'")
            return True
        
        # Rule 4: Foreign surnames ending in consonants are indeclinable for females
        # Examples: Шмидт (female), Вебер (female), Фишер (female)
        # But these same surnames decline for males: Шмидту, Веберу, Фишеру
        if gender == GENDER_FEMALE_VALUE:
            # Check if surname ends in a consonant
            consonants = 'бвгджзйклмнпрстфхцчшщ'
            
            if surname_lower and surname_lower[-1] in consonants:
                # Additional check: it's likely a foreign surname if it doesn't have
                # typical Russian patterns
                russian_female_endings = ('ова', 'ева', 'ина', 'ая', 'яя')
                is_russian_female = any(surname_lower.endswith(ending) for ending in russian_female_endings)
                
                if not is_russian_female:
                    logger.debug(
                        f"Surname '{surname}' is indeclinable: foreign surname "
                        f"ending in consonant for female gender"
                    )
                    return True
        
        return False

    def decline_name_with_gender(
        self,
        name: str,
        case: str,
        gender: str,
        tag_type: Optional[str] = None,
    ) -> str:
        """
        Decline a name component with explicit gender specification.

        This method provides more accurate declension by specifying both the
        grammatical case and gender. It also supports fallback to Petrovich
        library if pymorphy3 fails to decline the name.
        
        First checks the override cache for manual declension rules, then
        falls back to pymorphy3 and Petrovich if no override is found.

        Args:
            name: The name component to decline.
            case: The target grammatical case.
            gender: The gender ('male' or 'female').
            tag_type: The morphological tag type ('Surn', 'Name', or 'Patr'). Optional.

        Returns:
            The declined name in the specified case, capitalized.

        Examples:
            >>> decliner = NameDeclension()
            >>> decliner.decline_name_with_gender(
            ...     "Иванов", DATIVE_CASE, "male", TAG_SURNAME
            ... )
            'Иванову'
        """
        if not name:
            return ""
        
        # Check override cache first
        if tag_type and self._override_cache:
            # Map tag_type to word_type used in database
            tag_to_word_type = {
                TAG_SURNAME: 'surname',
                TAG_FIRST_NAME: 'firstname',
                TAG_PATRONYMIC: 'patronymic'
            }
            
            word_type = tag_to_word_type.get(tag_type)
            
            if word_type:
                cache_key = (name.lower(), word_type, gender, case)
                
                if cache_key in self._override_cache:
                    override_value = self._override_cache[cache_key].capitalize()
                    logger.info(
                        f"Using name override: '{name}' ({word_type}, {gender}, {case}) -> '{override_value}'"
                    )
                    return override_value
        
        # Check if surname is indeclinable (before attempting morphological analysis)
        if tag_type == TAG_SURNAME and self.is_indeclinable_surname(name, gender):
            logger.info(f"Surname '{name}' is indeclinable, returning unchanged")
            return name

        try:
            parsed = self.morph.parse(name)

            best_parse = None
            for p in parsed:
                tags_match = True

                if tag_type and tag_type not in p.tag:
                    tags_match = False

                gender_tag = (
                    GENDER_MALE if gender == GENDER_MALE_VALUE else GENDER_FEMALE
                )
                if gender_tag not in p.tag:
                    tags_match = False

                if tags_match:
                    best_parse = p
                    break

            if best_parse is None:
                for p in parsed:
                    if tag_type and tag_type in p.tag:
                        best_parse = p
                        break

            if best_parse is None and parsed:
                best_parse = parsed[0]

            result = name
            if best_parse:
                inflected = best_parse.inflect({case})
                if inflected:
                    result = inflected.word.capitalize()
                    logger.debug(f"Pymorphy3 declined '{name}' -> '{result}' (case={case}, gender={gender})")

            # Check if pymorphy3 result is suspicious for male surnames in dative case
            is_bad_result = False
            if (tag_type == TAG_SURNAME and 
                gender == GENDER_MALE_VALUE and 
                case == DATIVE_CASE and 
                result != name):
                
                # Male surnames ending in consonants should get -у/-ю in dative, not -ам/-ям (plural)
                # Also check for other suspicious patterns
                suspicious_endings = ('ам', 'ям', 'Ам', 'Ям', 'ами', 'ями', 'Ами', 'Ями')
                
                if result.endswith(suspicious_endings):
                    is_bad_result = True
                    logger.warning(
                        f"Detected suspicious plural/instrumental form for male surname: "
                        f"'{name}' -> '{result}'. Will try Petrovich fallback."
                    )
                
                # Additional check: for male surnames ending in consonants (-ин, -ов, -ев, etc.)
                # the dative form should typically end in -у or -ю, not other endings
                consonant_endings = ('ин', 'ов', 'ев', 'ёв', 'ын', 'ний', 'ский', 'цкий', 'ной')
                ends_in_consonant_suffix = any(name.lower().endswith(ending) for ending in consonant_endings)
                
                if ends_in_consonant_suffix and not result.endswith(('у', 'ю', 'У', 'Ю')):
                    is_bad_result = True
                    logger.warning(
                        f"Male surname '{name}' ends in consonant but dative form '{result}' "
                        f"doesn't end in -у/-ю. Will try Petrovich fallback."
                    )

            # Use Petrovich if pymorphy3 failed or returned suspicious result
            if (result == name or is_bad_result) and self.petrovich and tag_type:
                petrovich_result = self._decline_with_petrovich(
                    name, case, gender, tag_type
                )
                if petrovich_result != name:
                    logger.info(f"✓ Tier 2: Petrovich fallback: '{name}' -> '{petrovich_result}' (was: '{result}')")
                    result = petrovich_result
                else:
                    logger.debug(f"Petrovich also couldn't decline '{name}', will try Natasha")
                    
                    # Try Natasha as Tier 3 fallback
                    if NATASHA_AVAILABLE and tag_type:
                        natasha_result = self._decline_with_natasha(
                            name, case, gender, tag_type
                        )
                        if natasha_result != name:
                            result = natasha_result
                        else:
                            logger.warning(f"All 3 tiers failed to decline '{name}', keeping original")
            
            # If Petrovich not available but pymorphy3 failed, try Natasha directly
            elif (result == name or is_bad_result) and NATASHA_AVAILABLE and tag_type:
                natasha_result = self._decline_with_natasha(
                    name, case, gender, tag_type
                )
                if natasha_result != name:
                    result = natasha_result
                else:
                    logger.warning(f"Pymorphy3 and Natasha failed to decline '{name}', keeping original")

            return result

        except Exception as e:
            logger.warning(
                f"Error declining name '{name}' with gender '{gender}' "
                f"to case '{case}': {e}"
            )
            return name

    def _decline_with_petrovich(
        self, name: str, case: str, gender: str, tag_type: str
    ) -> str:
        """
        Decline a name using Petrovich library as a fallback.

        This is a private helper method used when pymorphy3 fails to decline a name.
        Note: Petrovich doesn't support nominative case (returns original name).

        Args:
            name: The name to decline.
            case: The grammatical case.
            gender: The gender.
            tag_type: The name type tag.

        Returns:
            The declined name, or the original name if declension fails.
        """
        try:
            # Petrovich doesn't support nominative case - return original name
            if case == NOMINATIVE_CASE:
                return name
            
            # Map our case constants to Petrovich Case enum
            petrovich_case_map = {
                GENITIVE_CASE: Case.GENITIVE,
                DATIVE_CASE: Case.DATIVE,
                ACCUSATIVE_CASE: Case.ACCUSATIVE,
                INSTRUMENTAL_CASE: Case.INSTRUMENTAL,
                PREPOSITIONAL_CASE: Case.PREPOSITIONAL,
            }
            
            petrovich_case = petrovich_case_map.get(case)
            if petrovich_case is None:
                logger.debug(f"Unsupported case '{case}' for Petrovich")
                return name
            
            petrovich_gender = (
                Gender.MALE if gender == GENDER_MALE_VALUE else Gender.FEMALE
            )

            if tag_type == TAG_SURNAME:
                return self.petrovich.lastname(name, petrovich_case, petrovich_gender)
            elif tag_type == TAG_FIRST_NAME:
                return self.petrovich.firstname(
                    name, petrovich_case, petrovich_gender
                )
            elif tag_type == TAG_PATRONYMIC:
                return self.petrovich.middlename(
                    name, petrovich_case, petrovich_gender
                )

        except Exception as e:
            logger.debug(f"Petrovich declension failed for '{name}': {e}")

        return name

    def _decline_with_natasha(
        self, name: str, case: str, gender: str, tag_type: str
    ) -> str:
        """
        Decline a name using Natasha ML as a Tier 3 fallback.

        This method uses Natasha's morphological tagger to analyze the name
        and extract morphological features, then combines with pymorphy3 for
        generation to improve accuracy on difficult names.

        SECURITY: Natasha is an OFFLINE library - all processing is local,
        no data sent to external servers. Safe for личные данные.

        Args:
            name: The name to decline.
            case: The grammatical case.
            gender: The gender.
            tag_type: The name type tag.

        Returns:
            The declined name, or the original name if declension fails.
        """
        if not NATASHA_AVAILABLE:
            return name
        
        if not all([self.natasha_segmenter, self.natasha_tagger, self.natasha_morph_vocab]):
            return name
        
        try:
            # Use Natasha to analyze the name morphologically
            doc = Doc(name)
            doc.segment(self.natasha_segmenter)
            doc.tag_morph(self.natasha_tagger)
            
            if not doc.tokens:
                logger.debug(f"Natasha: no tokens found for '{name}'")
                return name
            
            # Get the first token (should be the name)
            token = doc.tokens[0]
            
            if not token.pos:
                logger.debug(f"Natasha: no POS tag for '{name}'")
                return name
            
            # Natasha provides morphological features
            # Use this to guide pymorphy3 with better context
            logger.debug(f"Natasha analysis for '{name}': POS={token.pos}, feats={token.feats}")
            
            # Try to use Natasha's morphological information to select better parse
            parsed = self.morph.parse(name)
            
            if not parsed:
                return name
            
            # Try to find a parse that matches Natasha's analysis
            best_parse = None
            
            # Map our tag_type to expected Natasha POS
            expected_pos_map = {
                TAG_SURNAME: 'PROPN',    # Proper noun
                TAG_FIRST_NAME: 'PROPN',
                TAG_PATRONYMIC: 'PROPN'
            }
            
            expected_pos = expected_pos_map.get(tag_type, 'PROPN')
            
            # First, try to find parse that matches both gender and type
            gender_tag = GENDER_MALE if gender == GENDER_MALE_VALUE else GENDER_FEMALE
            
            for p in parsed:
                # Check if this parse matches our requirements
                if tag_type and tag_type in p.tag and gender_tag in p.tag:
                    best_parse = p
                    break
            
            # If not found, try without gender constraint
            if best_parse is None:
                for p in parsed:
                    if tag_type and tag_type in p.tag:
                        best_parse = p
                        break
            
            # Last resort: use first parse
            if best_parse is None and parsed:
                best_parse = parsed[0]
            
            # Now inflect using the best parse
            if best_parse:
                inflected = best_parse.inflect({case})
                if inflected and inflected.word != name:
                    result = inflected.word.capitalize()
                    logger.info(
                        f"✓ Tier 3: Natasha analysis helped decline '{name}' -> '{result}' "
                        f"(case={case}, gender={gender}, pos={token.pos})"
                    )
                    return result
            
            logger.debug(f"Natasha couldn't improve declension for '{name}'")
            return name
            
        except Exception as e:
            logger.debug(f"Natasha declension failed for '{name}': {e}")
            return name

    def get_salutation(self, gender: str) -> str:
        """
        Get the appropriate Russian salutation based on gender.

        Args:
            gender: The gender ('male' or 'female').

        Returns:
            The salutation: 'Уважаемый' for male, 'Уважаемая' for female.

        Examples:
            >>> decliner = NameDeclension()
            >>> decliner.get_salutation("male")
            'Уважаемый'
            >>> decliner.get_salutation("female")
            'Уважаемая'
        """
        return (
            "Уважаемый" if gender == GENDER_MALE_VALUE else "Уважаемая"
        )

    def get_full_salutation(
        self,
        first_name: str,
        middle_name: str,
        gender: Optional[str] = None,
    ) -> str:
        """
        Generate a complete formal salutation with name and patronymic.

        Creates a formal Russian greeting in the format "Уважаемый/ая [Name] [Patronymic]!".

        Args:
            first_name: The person's first name.
            middle_name: The person's patronymic.
            gender: The gender. If None, will be auto-detected. Optional.

        Returns:
            A complete salutation string with exclamation mark.

        Examples:
            >>> decliner = NameDeclension()
            >>> decliner.get_full_salutation("Иван", "Иванович", "male")
            'Уважаемый Иван Иванович!'
        """
        try:
            if gender is None:
                gender = self.detect_gender(first_name, middle_name)

            salutation = self.get_salutation(gender)

            first_declined = self.decline_name_with_gender(
                first_name, NOMINATIVE_CASE, gender, TAG_FIRST_NAME
            )
            middle_declined = (
                self.decline_name_with_gender(
                    middle_name, NOMINATIVE_CASE, gender, TAG_PATRONYMIC
                )
                if middle_name
                else ""
            )

            result = f"{salutation} {first_declined} {middle_declined}!".strip()
            return result.replace("  ", " ")

        except Exception as e:
            logger.error(
                f"Error generating salutation for '{first_name} {middle_name}': {e}"
            )
            return f"Уважаемый {first_name} {middle_name}!".strip().replace("  ", " ")

    def get_short_name_dative(
        self,
        last_name: str,
        first_name: str,
        middle_name: Optional[str] = None,
    ) -> str:
        """
        Generate a short name format with surname in dative case and initials.

        Creates a formal name in the format "Surname I.M." where surname is in
        dative case and I.M. are the initials.

        Args:
            last_name: The person's surname.
            first_name: The person's first name.
            middle_name: The person's patronymic. Optional.

        Returns:
            Short name format with declined surname and initials.

        Examples:
            >>> decliner = NameDeclension()
            >>> decliner.get_short_name_dative("Иванов", "Иван", "Иванович")
            'Иванову И.И.'
        """
        try:
            gender = self.detect_gender(first_name, middle_name)

            last_declined = self.decline_name_with_gender(
                last_name, DATIVE_CASE, gender, TAG_SURNAME
            )

            first_initial = first_name[0].upper() + "." if first_name else ""
            middle_initial = middle_name[0].upper() + "." if middle_name else ""

            if middle_initial:
                return f"{last_declined} {first_initial}{middle_initial}"
            else:
                return f"{last_declined} {first_initial}"

        except Exception as e:
            logger.error(
                f"Error generating short dative name for '{last_name} "
                f"{first_name} {middle_name or ''}': {e}"
            )
            first_initial = first_name[0].upper() + "." if first_name else ""
            middle_initial = middle_name[0].upper() + "." if middle_name else ""
            return f"{last_name} {first_initial}{middle_initial}".strip()

    def dative_to_nominative(
        self, last_name: str, first_initial: str, middle_initial: str
    ) -> Tuple[str, str, str]:
        """
        Convert a surname from dative case to nominative case.

        This method attempts to reverse-engineer the nominative form of a surname
        that is provided in dative case, using morphological analysis and
        heuristic rules for common Russian surname endings.

        Args:
            last_name: The surname in dative case.
            first_initial: The first name initial (used for gender detection).
            middle_initial: The patronymic initial.

        Returns:
            A tuple of (nominative_surname, first_initial, middle_initial).

        Examples:
            >>> decliner = NameDeclension()
            >>> decliner.dative_to_nominative("Иванову", "И", "И")
            ('Иванов', 'И', 'И')
        """
        try:
            parsed = self.morph.parse(last_name)

            male_endings = ["ому", "ову", "еву", "ину", "ыну"]
            female_endings = ["овой", "евой", "ой", "ской", "иной", "ью", "ей"]

            is_male_ending = any(last_name.endswith(e) for e in male_endings)
            is_female_ending = any(last_name.endswith(e) for e in female_endings)

            is_female = False
            if is_female_ending:
                is_female = True
            elif is_male_ending:
                is_female = False
            else:
                female_initials = ["Е", "И", "Ю", "Я"]
                is_female = first_initial in female_initials

            nominative_last = last_name
            for p in parsed:
                if TAG_SURNAME in p.tag or TAG_FIRST_NAME in p.tag:
                    inflected = p.inflect({NOMINATIVE_CASE})
                    if inflected:
                        candidate = inflected.word.capitalize()
                        if is_female and GENDER_FEMALE in p.tag:
                            nominative_last = candidate
                            break
                        elif not is_female and GENDER_MALE in p.tag:
                            nominative_last = candidate
                            break

            if nominative_last == last_name:
                nominative_last = self._apply_heuristic_rules(last_name)

            return (nominative_last, first_initial, middle_initial)

        except Exception as e:
            logger.warning(
                f"Error converting dative '{last_name}' to nominative: {e}"
            )
            return (last_name, first_initial, middle_initial)

    def _apply_heuristic_rules(self, last_name: str) -> str:
        """
        Apply heuristic rules to convert dative surname to nominative.

        This is a private helper method for common Russian surname patterns.

        Args:
            last_name: The surname in dative case.

        Returns:
            The surname converted to nominative case using pattern matching.
        """
        if last_name.endswith("ому"):
            return last_name[:-3]
        elif last_name.endswith("ову"):
            return last_name[:-3] + "ов"
        elif last_name.endswith("еву"):
            return last_name[:-3] + "ев"
        elif last_name.endswith("ину"):
            return last_name[:-3] + "ин"
        elif last_name.endswith("ыну"):
            return last_name[:-3] + "ын"
        elif last_name.endswith("овой"):
            return last_name[:-2]
        elif last_name.endswith("евой"):
            return last_name[:-2]
        elif last_name.endswith("ой"):
            return last_name[:-2]
        elif last_name.endswith("ской"):
            return last_name[:-2] + "ий"
        elif last_name.endswith("иной"):
            return last_name[:-2]
        elif last_name.endswith(("ью", "ей")):
            return last_name[:-2] + "ь"

        return last_name

    def decline_text_to_accusative(self, text: str) -> str:
        """
        Decline the first word(s) of a text to accusative case with preposition.

        This method is useful for converting organization or department names
        to the accusative case with the preposition 'в' (into/to).
        For phrases with adjectives, declines both adjective and noun.
        
        Note: "префектура" remains lowercase as per style requirements.

        Args:
            text: The text to decline (typically an organization name).

        Returns:
            The text with the first word(s) in accusative case, prefixed with 'в '.

        Examples:
            >>> decliner = NameDeclension()
            >>> decliner.decline_text_to_accusative("Государственная жилищная инспекция")
            'в Государственную жилищную инспекцию'
            >>> decliner.decline_text_to_accusative("префектура ЗАО")
            'в префектуру ЗАО'
        """
        if not text:
            return text

        try:
            if text.startswith("в ") or text.startswith("В "):
                return text

            words = text.split()
            if not words:
                return "в " + text
            
            first_word = words[0]
            
            # Check if this is "префектура" (should remain lowercase)
            is_prefecture = first_word.lower() == "префектура"
            
            # Capitalize first word unless it's a prefecture
            if words[0].islower() and not is_prefecture:
                words[0] = words[0].capitalize()

            # Decline first word (adjective or noun)
            first_word = words[0]
            parsed_first = self.morph.parse(first_word)

            if parsed_first:
                p = parsed_first[0]
                inflected = p.inflect({ACCUSATIVE_CASE})
                if inflected:
                    # Keep lowercase for prefecture, capitalize for others
                    if is_prefecture:
                        words[0] = inflected.word.lower()
                    else:
                        words[0] = inflected.word.capitalize()
            
            # If first word is adjective and there's a second word (noun), decline it too
            if len(words) >= 2 and TAG_ADJECTIVE in parsed_first[0].tag:
                second_word = words[1]
                parsed_second = self.morph.parse(second_word)
                
                if parsed_second:
                    p2 = parsed_second[0]
                    if TAG_NOUN in p2.tag or TAG_ADJECTIVE in p2.tag:
                        inflected2 = p2.inflect({ACCUSATIVE_CASE})
                        if inflected2:
                            words[1] = inflected2.word.lower()
            
            # If second word is also adjective and there's a third word (noun), decline it too
            if len(words) >= 3:
                parsed_second = self.morph.parse(words[1])
                if parsed_second and TAG_ADJECTIVE in parsed_second[0].tag:
                    third_word = words[2]
                    parsed_third = self.morph.parse(third_word)
                    
                    if parsed_third:
                        p3 = parsed_third[0]
                        if TAG_NOUN in p3.tag:
                            inflected3 = p3.inflect({ACCUSATIVE_CASE})
                            if inflected3:
                                words[2] = inflected3.word.lower()

            return "в " + " ".join(words)

        except Exception as e:
            logger.warning(f"Error declining text '{text}' to accusative: {e}")
            return "в " + text

    def decline_text_to_genitive(self, text: str) -> str:
        """
        Decline the first word of a text to genitive case.

        This method is useful for converting organization or department names
        to the genitive case (possessive form).

        Args:
            text: The text to decline (typically an organization name).

        Returns:
            The text with the first word in genitive case.

        Examples:
            >>> decliner = NameDeclension()
            >>> decliner.decline_text_to_genitive("Инспекция")
            'Инспекции'
        """
        if not text:
            return text

        try:
            words = text.split()
            if not words:
                return text

            first_word = words[0]
            parsed = self.morph.parse(first_word)

            if parsed:
                p = parsed[0]
                if TAG_NOUN in p.tag:
                    inflected = p.inflect({GENITIVE_CASE})
                    if inflected:
                        words[0] = inflected.word

            return " ".join(words)

        except Exception as e:
            logger.warning(f"Error declining text '{text}' to genitive: {e}")
            return text

    def normalize_case(self, text: str) -> str:
        """
        Normalize the capitalization of text from all-uppercase to title case.

        This method intelligently capitalizes words based on their morphological
        properties, capitalizing names, surnames, patronymics, nouns, and adjectives
        while leaving other words (prepositions, conjunctions) in lowercase.

        Args:
            text: The text to normalize (typically in all-uppercase).

        Returns:
            The text with proper capitalization.

        Examples:
            >>> decliner = NameDeclension()
            >>> decliner.normalize_case("ИВАНОВ ИВАН ИВАНОВИЧ")
            'Иванов Иван Иванович'
        """
        if not text:
            return text

        try:
            if text.isupper():
                words = text.lower().split()
                normalized = []

                for word in words:
                    parsed = self.morph.parse(word)
                    if parsed:
                        p = parsed[0]
                        if any(
                            tag in p.tag
                            for tag in [
                                TAG_FIRST_NAME,
                                TAG_SURNAME,
                                TAG_PATRONYMIC,
                                TAG_NOUN,
                                TAG_ADJECTIVE,
                            ]
                        ):
                            normalized.append(word.capitalize())
                        else:
                            normalized.append(word)
                    else:
                        normalized.append(word.capitalize())

                return " ".join(normalized)

            return text

        except Exception as e:
            logger.warning(f"Error normalizing case for text '{text}': {e}")
            return text


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    )

    decliner = NameDeclension()

    test_cases = [
        ("Щеголева", "Анна", "Михайловна"),
        ("Мусин", "Хасан", "Эльдарович"),
        ("Клиновская", "Анна", "Сергеевна"),
        ("Каменский", "Дмитрий", "Владимирович"),
    ]

    logger.info("Starting name declension tests")

    for last, first, middle in test_cases:
        result = decliner.decline_full_name(last, first, middle)
        salutation = decliner.get_full_salutation(
            first, middle, result["gender"]
        )

        logger.info("=" * 50)
        logger.info(f"Исходное: {last} {first} {middle}")
        logger.info(f"Пол: {result['gender']}")
        logger.info(f"Дательный падеж: {result['full_name']}")
        logger.info(f"Обращение: {salutation}")
