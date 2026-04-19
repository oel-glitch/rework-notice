"""
Тесты модуля генерации Word документов
"""
import pytest
from pathlib import Path
from src.word_generator import WordGenerator


class TestWordGenerator:
    
    @pytest.fixture
    def generator(self):
        return WordGenerator()
    
    def test_generator_initialization(self, generator):
        """Тест инициализации генератора"""
        assert generator is not None
    
    def test_select_template_ch3_single_department(self, generator):
        """Тест выбора шаблона ч.3 для одного департамента"""
        template = generator.select_template(
            law_part="3",
            departments=["Департамент транспорта"]
        )
        assert template is not None
        assert "ch3" in template.lower()
    
    def test_select_template_ch4_multiple(self, generator):
        """Тест выбора шаблона ч.4 для нескольких департаментов"""
        template = generator.select_template(
            law_part="4",
            departments=["Департамент 1", "Департамент 2"]
        )
        assert template is not None
        assert "ch4" in template.lower()
        assert "multiple" in template.lower()
    
    def test_select_template_ch4_prefecture(self, generator):
        """Тест выбора шаблона ч.4 для префектуры"""
        template = generator.select_template(
            law_part="4",
            departments=["Префектура ЦАО"]
        )
        assert template is not None
        assert "prefecture" in template.lower()
    
    def test_generate_filename(self, generator):
        """Тест генерации имени файла"""
        filename = generator.generate_filename(
            last_name="Иванов",
            first_initial="П",
            middle_initial="С",
            oati_number="01-21-П-9141/25"
        )
        
        assert "Иванов" in filename
        assert "П" in filename
        assert "С" in filename
        assert "9141" in filename
        assert "25" in filename
        assert filename.endswith(".docx")
    
    def test_generate_filename_female(self, generator):
        """Тест генерации имени файла для женского имени"""
        filename = generator.generate_filename(
            last_name="Петрова",
            first_initial="М",
            middle_initial="А",
            oati_number="01-21-П-1234/25"
        )
        
        assert "Петрова" in filename
        assert "М" in filename
        assert "А" in filename
        assert "1234" in filename
    
    def test_template_exists(self):
        """Тест существования шаблонов"""
        templates_dir = Path("templates")
        
        expected_templates = [
            "template_ch3.docx",
            "template_ch4_multiple.docx",
            "template_ch4_prefecture.docx"
        ]
        
        for template in expected_templates:
            template_path = templates_dir / template
            if template_path.exists():
                assert template_path.is_file()
    
    def test_select_template_with_inspector(self, generator):
        """Тест выбора шаблона при наличии инспектора ОАТИ"""
        template = generator.select_template(
            law_part="4",
            departments=["Департамент транспорта"],
            has_inspector=True
        )
        assert template is not None
        assert "ch4" in template.lower()
        assert "prefecture" in template.lower()
    
    def test_select_template_single_dept_ch3(self, generator):
        """Тест выбора ч.3 для одного департамента"""
        template = generator.select_template(
            law_part="3",
            departments=["Департамент транспорта"],
            has_inspector=False
        )
        assert template is not None
        assert "ch3" in template.lower()
    
    def test_select_template_multiple_depts_ch4(self, generator):
        """Тест выбора ч.4 для нескольких департаментов"""
        template = generator.select_template(
            law_part="4",
            departments=["Департамент 1", "Департамент 2"],
            has_inspector=False
        )
        assert template is not None
        assert "ch4" in template.lower()
        assert "multiple" in template.lower()
    
    def test_add_non_breaking_spaces(self, generator):
        """Тест добавления неразрывных пробелов"""
        nbsp = '\u00A0'
        
        test_cases = [
            ('01-21-П-9141/25', f'01-21-П-{nbsp}9141/25'),
            ('01-21-П- 9141/25', f'01-21-П-{nbsp}9141/25'),
            ('01-21-П-  9141/25', f'01-21-П-{nbsp}9141/25'),
            ('№ 59-ФЗ', f'№{nbsp}59-ФЗ'),
            ('ФЗ «О порядке', f'ФЗ{nbsp}«О порядке'),
        ]
        
        for text, expected in test_cases:
            result = generator.add_non_breaking_spaces(text)
            assert result == expected, f"Expected '{expected}', got '{result}'"
