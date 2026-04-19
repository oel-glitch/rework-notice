"""
Тесты модуля парсинга PDF
"""
import pytest
from src.pdf_parser import PDFParser


class TestPDFParser:
    
    @pytest.fixture
    def parser(self):
        return PDFParser()
    
    def test_parser_initialization(self, parser):
        """Тест инициализации парсера"""
        assert parser is not None
    
    def test_extract_citizen_info_fio(self, parser):
        """Тест извлечения ФИО через extract_citizen_info"""
        text = "ФИО: Иванов Петр Сергеевич"
        result = parser.extract_citizen_info(text)
        assert result['last_name'] == "Иванов"
        assert result['first_name'] == "Петр"
        assert result['middle_name'] == "Сергеевич"
        assert result['full_name'] == "Иванов Петр Сергеевич"
    
    def test_extract_citizen_info_with_variations(self, parser):
        """Тест извлечения ФИО с вариациями формата"""
        test_cases = [
            ("ФИО: Петрова Мария Александровна", "Петрова", "Мария", "Александровна"),
            ("ФИО: Смирнов Алексей Владимирович", "Смирнов", "Алексей", "Владимирович"),
        ]
        
        for text, expected_last, expected_first, expected_middle in test_cases:
            result = parser.extract_citizen_info(text)
            assert result['last_name'] == expected_last
            assert result['first_name'] == expected_first
            assert result['middle_name'] == expected_middle
    
    def test_extract_citizen_info_email(self, parser):
        """Тест извлечения email через extract_citizen_info"""
        text = "Электронный адрес: test@example.com"
        result = parser.extract_citizen_info(text)
        assert result['email'] == "test@example.com"
    
    def test_extract_citizen_info_email_variations(self, parser):
        """Тест извлечения email с вариациями"""
        test_cases = [
            ("Электронный адрес: user@domain.ru", "user@domain.ru"),
            ("Электронный адрес: contact@company.org", "contact@company.org")
        ]
        
        for text, expected_email in test_cases:
            result = parser.extract_citizen_info(text)
            assert result['email'] == expected_email
    
    def test_extract_citizen_info_oati_number(self, parser):
        """Тест извлечения номера ОАТИ через extract_citizen_info"""
        text = "№ 01-21-П-9141/25 от 15.10.2025"
        result = parser.extract_citizen_info(text)
        assert result['oati_number'] == "01-21-П-9141/25"
        assert result['oati_date'] == "15.10.2025"
    
    def test_extract_citizen_info_law_part(self, parser):
        """Тест извлечения части закона из резолюции"""
        test_cases = [
            ("В соответствии с ч. 3 ст.8 Федерального закона от 02.05.2006 №59-ФЗ «О порядке рассмотрения обращений граждан Российской Федерации»", "3"),
            ("В соответствии с ч. 4 ст.8 Федерального закона от 02.05.2006 №59-ФЗ «О порядке рассмотрения обращений граждан Российской Федерации» направить", "4"),
        ]
        
        for text, expected_part in test_cases:
            result = parser.extract_citizen_info(text)
            assert result['law_part'] == expected_part
    
    def test_extract_departments(self, parser):
        """Тест извлечения департаментов"""
        text = """
        Резолюция: направить в Департамент транспорта города Москвы,
        префектуру ЦАО города Москвы
        """
        result = parser.extract_departments(text)
        assert isinstance(result, list)
        assert len(result) > 0
    
    def test_extract_citizen_info_empty_text(self, parser):
        """Тест с пустым текстом"""
        result = parser.extract_citizen_info("")
        assert result['last_name'] == ""
        assert result['first_name'] == ""
        assert result['middle_name'] == ""
        assert result['email'] == ""
        assert result['oati_number'] == ""
    
    def test_full_text_extraction(self, parser, test_data_dir):
        """Тест извлечения из полного текста (end-to-end)"""
        test_file = test_data_dir / "test_sample.txt"
        if test_file.exists():
            with open(test_file, 'r', encoding='utf-8') as f:
                text = f.read()
            
            result = parser.extract_citizen_info(text)
            
            assert result['last_name'] == 'Щеголева'
            assert result['first_name'] == 'Анна'
            assert result['middle_name'] == 'Михайловна'
            assert result['email'] == 'test@example.com'
            assert '9141' in result['oati_number']
            assert result['law_part'] == '3'
            
            departments = parser.extract_departments(text)
            assert len(departments) > 0
    
    def test_extract_portal_source_nash_gorod(self, parser):
        """Тест определения портала 'Наш город'"""
        text = "Сообщение с портала Наш Город №12345678"
        result = parser.extract_portal_source(text)
        assert result == "поступившее с портала Правительства Москвы «Наш город»"
    
    def test_extract_portal_source_mos_ru(self, parser):
        """Тест определения портала 'mos.ru'"""
        text = "Обращение № 12345678, поступившее на портал mos.ru"
        result = parser.extract_portal_source(text)
        assert result == "поступившее с портала Мэра и Правительства Москвы"
    
    def test_extract_portal_source_default(self, parser):
        """Тест дефолтного значения портала"""
        text = "Обращение без указания портала"
        result = parser.extract_portal_source(text)
        assert result == "поступившее в ОАТИ"
    
    def test_extract_recipients_from_resolution(self, parser):
        """Тест извлечения получателей из резолюции"""
        text = """
        В соответствии с ч. 4 ст.8 Федерального закона направить
        Урожаевой Ю.В. и Слободчикову А.О. для рассмотрения.
        """
        recipients = parser.extract_recipients_from_resolution(text)
        assert len(recipients) >= 1
        assert any(r['last_name'] == 'Урожаевой' for r in recipients)
    
    def test_extract_recipients_uppercase(self, parser):
        """Тест извлечения получателей в UPPERCASE"""
        text = """
        В соответствии с ч. 4 ст.8 направить КИЧИКОВУ Б.Б. и УРОЖАЕВОЙ Ю.В.
        """
        recipients = parser.extract_recipients_from_resolution(text)
        assert len(recipients) == 2
        assert any('КИЧИКОВУ' in r['last_name'] for r in recipients)
    
    def test_extract_recipients_invariant_surnames(self, parser):
        """Тест извлечения несклоняемых фамилий"""
        text = """
        В соответствии с ч. 4 ст.8 направить ТКАЧЕНКО В.В. и ИГНАТЮК О.О.
        """
        recipients = parser.extract_recipients_from_resolution(text)
        assert len(recipients) == 2
        assert any('ТКАЧЕНКО' in r['last_name'] for r in recipients)
        assert any('ИГНАТЮК' in r['last_name'] for r in recipients)
    
    def test_extract_recipients_filters_service_personnel(self, parser):
        """Тест фильтрации служебных лиц"""
        text = """
        В соответствии с ч. 4 ст.8 направить
        Аравину А.А., Урожаевой Ю.В., +Ларину И.И.
        """
        recipients = parser.extract_recipients_from_resolution(text)
        assert not any('Аравин' in r['last_name'] for r in recipients)
        assert not any('Ларин' in r['last_name'] for r in recipients)
        assert any('Урожаевой' in r['last_name'] for r in recipients)
    
    def test_extract_recipients_filters_executors(self, parser):
        """Тест фильтрации исполнителей"""
        text = """
        В соответствии с ч. 4 ст.8 направить Урожаевой Ю.В.
        Исп: Драгушина В.С.
        """
        recipients = parser.extract_recipients_from_resolution(text)
        assert not any('Драгушина' in r['last_name'] for r in recipients)
    
    def test_extract_questions(self, parser):
        """Тест извлечения вопросов ОАТИ"""
        text = """
        Вопрос 1: Соблюдение условий по уровню шума и вибрации
        Вопрос 2: Соблюдение условий проведения работ
        """
        questions = parser.extract_questions(text)
        assert len(questions) == 2
        assert questions[0]['number'] == '1'
        assert 'шума' in questions[0]['text'].lower()
        assert questions[1]['number'] == '2'
    
    def test_extract_questions_removes_names(self, parser):
        """Тест удаления ФИО из вопросов"""
        text = """
        Вопрос 1: Соблюдение условий производства работ Иванов И.И.
        """
        questions = parser.extract_questions(text)
        assert len(questions) == 1
        assert 'Иванов' not in questions[0]['text']
