"""
Интеграционные тесты полного цикла обработки
"""
import pytest
from pathlib import Path
from src.pdf_parser import PDFParser
from src.name_declension import NameDeclension
from src.word_generator import WordGenerator
from src.database import Database


class TestIntegration:
    
    @pytest.fixture
    def setup_components(self, temp_db_path):
        """Настройка всех компонентов"""
        return {
            'parser': PDFParser(),
            'decliner': NameDeclension(),
            'generator': WordGenerator(),
            'db': Database(str(temp_db_path))
        }
    
    def test_full_pipeline_with_mock_data(self, setup_components):
        """Тест полного цикла с моковыми данными"""
        components = setup_components
        
        mock_pdf_data = {
            'last_name': 'Щеголева',
            'first_name': 'Анна',
            'middle_name': 'Михайловна',
            'email': 'test@example.com',
            'oati_number': '01-21-П-9141/25',
            'portal_id': '12345',
            'date': '15.10.2025',
            'law_part': '3',
            'departments': ['Департамент транспорта']
        }
        
        declined = components['decliner'].decline_full_name(
            mock_pdf_data['last_name'],
            mock_pdf_data['first_name'],
            mock_pdf_data['middle_name']
        )
        
        assert declined['gender'] == 'female'
        assert 'Щеголевой' in declined['full_name']
        
        template = components['generator'].select_template(
            law_part=mock_pdf_data['law_part'],
            departments=mock_pdf_data['departments']
        )
        assert template is not None
        
        filename = components['generator'].generate_filename(
            last_name=mock_pdf_data['last_name'],
            first_initial=mock_pdf_data['first_name'][0],
            middle_initial=mock_pdf_data['middle_name'][0],
            oati_number=mock_pdf_data['oati_number']
        )
        assert '9141' in filename
        
        components['db'].add_processing_record(
            pdf_file="test.pdf",
            citizen_name=f"{mock_pdf_data['last_name']} {mock_pdf_data['first_name']} {mock_pdf_data['middle_name']}",
            oati_number=mock_pdf_data['oati_number'],
            portal_id=mock_pdf_data['portal_id'],
            output_file=filename,
            status="success"
        )
    
    def test_male_name_pipeline(self, setup_components):
        """Тест цикла для мужского имени"""
        components = setup_components
        
        declined = components['decliner'].decline_full_name(
            'Иванов', 'Петр', 'Сергеевич'
        )
        
        assert declined['gender'] == 'male'
        salutation = components['decliner'].get_salutation(declined['gender'])
        assert salutation == 'Уважаемый'
        
        filename = components['generator'].generate_filename(
            last_name='Иванов',
            first_initial='П',
            middle_initial='С',
            oati_number='01-21-П-1234/25'
        )
        assert 'Иванов' in filename
        assert '1234' in filename
    
    def test_multiple_departments_pipeline(self, setup_components):
        """Тест цикла для нескольких департаментов"""
        components = setup_components
        
        departments = [
            'Департамент транспорта',
            'Департамент жилищно-коммунального хозяйства'
        ]
        
        template = components['generator'].select_template(
            law_part='4',
            departments=departments
        )
        
        assert template is not None
        assert 'multiple' in template.lower()
    
    def test_prefecture_pipeline(self, setup_components):
        """Тест цикла для префектуры"""
        components = setup_components
        
        template = components['generator'].select_template(
            law_part='4',
            departments=['Префектура ЦАО']
        )
        
        assert template is not None
        assert 'prefecture' in template.lower()
    
    def test_database_integration(self, setup_components):
        """Тест интеграции с БД"""
        db = setup_components['db']
        
        db.add_department('Тестовый департамент', 'ТД')
        
        departments = db.get_all_departments(active_only=True)
        assert any(d['name'] == 'Тестовый департамент' for d in departments)
