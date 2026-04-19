"""
Тесты модуля склонения имен
"""
import pytest
from src.name_declension import NameDeclension


class TestNameDeclension:
    
    @pytest.fixture
    def decliner(self):
        return NameDeclension()
    
    def test_female_name_declension(self, decliner):
        """Тест склонения женского имени"""
        result = decliner.decline_full_name("Щеголева", "Анна", "Михайловна")
        
        assert result['gender'] == 'female'
        salutation = decliner.get_salutation(result['gender'])
        assert salutation == 'Уважаемая'
        assert 'Щеголевой' in result['full_name']
        assert 'Анне' in result['full_name']
        assert 'Михайловне' in result['full_name']
    
    def test_male_name_declension(self, decliner):
        """Тест склонения мужского имени"""
        result = decliner.decline_full_name("Иванов", "Петр", "Сергеевич")
        
        assert result['gender'] == 'male'
        salutation = decliner.get_salutation(result['gender'])
        assert salutation == 'Уважаемый'
        assert 'Иванову' in result['full_name']
        assert 'Петру' in result['full_name']
        assert 'Сергеевичу' in result['full_name']
    
    def test_different_male_names(self, decliner):
        """Тест различных мужских имен"""
        test_cases = [
            ("Смирнов", "Алексей", "Владимирович"),
            ("Козлов", "Дмитрий", "Иванович"),
            ("Соколов", "Андрей", "Петрович")
        ]
        
        for last, first, middle in test_cases:
            result = decliner.decline_full_name(last, first, middle)
            assert result['gender'] == 'male'
            salutation = decliner.get_salutation(result['gender'])
            assert salutation == 'Уважаемый'
    
    def test_different_female_names(self, decliner):
        """Тест различных женских имен"""
        test_cases = [
            ("Петрова", "Мария", "Александровна"),
            ("Новикова", "Елена", "Дмитриевна"),
            ("Волкова", "Ольга", "Сергеевна")
        ]
        
        for last, first, middle in test_cases:
            result = decliner.decline_full_name(last, first, middle)
            assert result['gender'] == 'female'
            salutation = decliner.get_salutation(result['gender'])
            assert salutation == 'Уважаемая'
    
    def test_empty_name(self, decliner):
        """Тест с пустым именем"""
        result = decliner.decline_full_name("", "", "")
        
        assert result is not None
        assert 'full_name' in result
    
    def test_name_with_yo(self, decliner):
        """Тест имени с буквой ё"""
        result = decliner.decline_full_name("Семёнов", "Артём", "Фёдорович")
        
        assert result['gender'] == 'male'
        salutation = decliner.get_salutation(result['gender'])
        assert salutation == 'Уважаемый'
    
    def test_dative_to_nominative_female(self, decliner):
        """Тест преобразования дательный→именительный (женские фамилии)"""
        test_cases = [
            ('Урожаевой', 'Ю', 'В', 'Урожаева'),
            ('Фисковой', 'Е', 'А', 'Фискова'),
            ('Щеголевой', 'А', 'М', 'Щеголева'),
        ]
        
        for dative_last, f, m, expected_nom in test_cases:
            result = decliner.dative_to_nominative(dative_last, f, m)
            assert result[0] == expected_nom
            assert result[1] == f
            assert result[2] == m
    
    def test_dative_to_nominative_male(self, decliner):
        """Тест преобразования дательный→именительный (мужские фамилии)"""
        test_cases = [
            ('Слободчикову', 'А', 'О', 'Слободчиков'),
            ('Кичикову', 'Б', 'Б', 'Кичиков'),
            ('Иванову', 'И', 'И', 'Иванов'),
        ]
        
        for dative_last, f, m, expected_nom in test_cases:
            result = decliner.dative_to_nominative(dative_last, f, m)
            assert result[0] == expected_nom
    
    def test_normalize_case_uppercase(self, decliner):
        """Тест нормализации БОЛЬШИХ БУКВ"""
        test_cases = [
            ('ДУВАНКОВА МАРИЯ АЛЕКСАНДРОВНА', 'Дуванкова Мария Александровна'),
            ('ЩЕГОЛЕВА АННА МИХАЙЛОВНА', 'Щеголева Анна Михайловна'),
            ('ТКАЧЕНКО ВЛАДИМИР ВЛАДИМИРОВИЧ', 'Ткаченко Владимир Владимирович'),
        ]
        
        for uppercase_text, expected in test_cases:
            result = decliner.normalize_case(uppercase_text)
            assert result == expected
    
    def test_normalize_case_mixed(self, decliner):
        """Тест нормализации смешанного регистра"""
        mixed_text = "Иванов Иван Иванович"
        result = decliner.normalize_case(mixed_text)
        assert result == mixed_text
    
    def test_decline_text_to_accusative(self, decliner):
        """Тест склонения департаментов в винительный падеж"""
        test_cases = [
            ('Департамент транспорта города Москвы', 'в Департамент транспорта города Москвы'),
            ('префектура Северного административного округа города Москвы', 'в Префектуру Северного административного округа города Москвы'),
        ]
        
        for dept, expected in test_cases:
            result = decliner.decline_text_to_accusative(dept)
            assert result.startswith('в ')
            assert 'Департамент' in result or 'Префектуру' in result
    
    def test_decline_text_to_genitive(self, decliner):
        """Тест склонения вопросов в родительный падеж"""
        test_cases = [
            ('Соблюдение условий производства работ', 'соблюдения'),
            ('Соблюдение условий по уровню шума', 'соблюдения'),
        ]
        
        for question, expected_start in test_cases:
            result = decliner.decline_text_to_genitive(question)
            assert result.lower().startswith(expected_start)
    
    def test_get_short_name_dative(self, decliner):
        """Тест получения короткого ФИО с инициалами в дательном падеже"""
        test_cases = [
            ('Иванова', 'Екатерина', 'Александровна', 'Ивановой Е.А.'),
            ('Сорокина', 'Виктория', 'Евгеньевна', 'Сорокиной В.Е.'),
            ('Кондратьева', 'Анна', 'Вячеславовна', 'Кондратьевой А.В.'),
            ('Ложкова', 'Мария', 'Евгеньевна', 'Ложковой М.Е.'),
        ]
        
        for last, first, middle, expected in test_cases:
            result = decliner.get_short_name_dative(last, first, middle)
            assert result == expected
