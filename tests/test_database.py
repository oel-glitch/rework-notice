"""
Тесты модуля базы данных
"""
import pytest
from src.database import Database


class TestDatabase:
    
    @pytest.fixture
    def db(self, temp_db_path):
        """Создать тестовую БД"""
        return Database(str(temp_db_path))
    
    def test_database_initialization(self, db):
        """Тест инициализации БД"""
        assert db is not None
        departments = db.get_all_departments()
        assert isinstance(departments, list)
    
    def test_add_department(self, db):
        """Тест добавления департамента"""
        result = db.add_department(
            "Тестовый департамент", 
            "ТД"
        )
        assert result is True
        
        departments = db.get_all_departments(active_only=False)
        assert any(d['name'] == "Тестовый департамент" for d in departments)
    
    def test_update_department(self, db):
        """Тест обновления департамента"""
        db.add_department("Старое название", "СН")
        
        departments = db.get_all_departments(active_only=False)
        dept_id = next(d['id'] for d in departments if d['name'] == "Старое название")
        
        result = db.update_department(dept_id, "Новое название", "НН")
        assert result is True
        
        updated_deps = db.get_all_departments(active_only=False)
        assert any(d['name'] == "Новое название" for d in updated_deps)
        assert not any(d['name'] == "Старое название" for d in updated_deps)
    
    def test_delete_department(self, db):
        """Тест удаления департамента"""
        db.add_department("Для удаления", "УД")
        
        departments = db.get_all_departments(active_only=False)
        dept_id = next(d['id'] for d in departments if d['name'] == "Для удаления")
        
        result = db.delete_department(dept_id)
        assert result is True
        
        updated_deps = db.get_all_departments(active_only=True)
        assert not any(d['name'] == "Для удаления" for d in updated_deps)
    
    def test_get_active_departments(self, db):
        """Тест получения активных департаментов"""
        db.add_department("Активный департамент", "АД")
        
        active_deps = db.get_all_departments(active_only=True)
        assert isinstance(active_deps, list)
        assert any(d['name'] == "Активный департамент" for d in active_deps)
    
    def test_add_processing_record(self, db):
        """Тест добавления записи об обработке"""
        db.add_processing_record(
            pdf_file="test.pdf",
            citizen_name="Иванов Иван Иванович",
            oati_number="01-21-П-1234/25",
            portal_id="12345",
            output_file="output.docx",
            status="success"
        )
    
    def test_duplicate_department(self, db):
        """Тест добавления дубликата"""
        db.add_department("Дубликат", "ДБ")
        result = db.add_department("Дубликат", "ДБ")
        
        assert result is False
    
    def test_match_recipient_director(self, db):
        """Тест сопоставления получателя-директора"""
        from src.name_declension import NameDeclension
        decliner = NameDeclension()
        
        result = db.match_recipient_from_resolution('Урожаевой', 'Ю', 'В', decliner)
        assert result is not None
        assert result['type'] == 'director'
        assert 'Департамент природопользования' in result['data']['department']
    
    def test_match_recipient_inspector(self, db):
        """Тест сопоставления получателя-инспектора ОАТИ"""
        from src.name_declension import NameDeclension
        decliner = NameDeclension()
        
        result = db.match_recipient_from_resolution('Кичикову', 'Б', 'Б', decliner)
        assert result is not None
        assert result['type'] == 'inspector'
    
    def test_match_recipient_not_found(self, db):
        """Тест несуществующего получателя"""
        from src.name_declension import NameDeclension
        decliner = NameDeclension()
        
        result = db.match_recipient_from_resolution('Несуществующему', 'Н', 'Н', decliner)
        assert result is None
    
    def test_department_inn_ogrn_support(self, db):
        """Тест добавления департамента с ИНН и ОГРН"""
        result = db.add_department(
            "Организация с ИНН",
            "ОИ",
            "1234567890",
            "1234567890123",
            "г. Москва"
        )
        assert result is True
        
        dept = db.get_department_by_inn("1234567890")
        assert dept is not None
        assert dept[1] == "Организация с ИНН"
        assert dept[2] == "ОИ"
        assert dept[3] == "1234567890"
        assert dept[4] == "1234567890123"
    
    def test_get_department_by_ogrn(self, db):
        """Тест поиска департамента по ОГРН"""
        db.add_department("Организация ОГРН", "ОО", "", "9999999999999", "")
        
        dept = db.get_department_by_ogrn("9999999999999")
        assert dept is not None
        assert dept[1] == "Организация ОГРН"
    
    def test_manual_declension_storage(self, db):
        """Тест сохранения ручных склонений"""
        person_id = db.add_person("Иванов", "Иван", "Иванович", "director", "Департамент тестов")
        assert person_id is not None
        
        db.set_manual_declension(person_id, "gent", "Иванова Ивана Ивановича")
        
        declension = db.get_manual_declension(person_id, "gent")
        assert declension == "Иванова Ивана Ивановича"
    
    def test_manual_declension_override(self, db):
        """Тест переопределения ручных склонений"""
        person_id = db.add_person("Петров", "Петр", "Петрович", "inspector", None, "Петрову")
        assert person_id is not None
        
        db.set_manual_declension(person_id, "datv", "Первое значение")
        assert db.get_manual_declension(person_id, "datv") == "Первое значение"
        
        db.set_manual_declension(person_id, "datv", "Второе значение")
        assert db.get_manual_declension(person_id, "datv") == "Второе значение"
    
    def test_delete_manual_declensions(self, db):
        """Тест удаления ручных склонений"""
        person_id = db.add_person("Сидоров", "Сидор", "Сидорович", "director", "Тестовый департамент")
        assert person_id is not None
        
        db.set_manual_declension(person_id, "gent", "Сидорова")
        db.set_manual_declension(person_id, "datv", "Сидорову")
        db.set_manual_declension(person_id, "accs", "Сидорова")
        
        db.delete_manual_declension(person_id, "gent")
        db.delete_manual_declension(person_id, "datv")
        db.delete_manual_declension(person_id, "accs")
        
        assert db.get_manual_declension(person_id, "gent") is None
        assert db.get_manual_declension(person_id, "datv") is None
        assert db.get_manual_declension(person_id, "accs") is None
    
    def test_persons_table_crud(self, db):
        """Тест CRUD операций с таблицей persons"""
        person_id = db.add_person("Тестов", "Тест", "Тестович", "director", "Департамент CRUD")
        assert person_id is not None
        
        persons = db.get_all_persons(role="director")
        person_match = [p for p in persons if p['last_name'] == "Тестов"]
        assert len(person_match) > 0
        
        result = db.update_person(person_id, "Обновленный", "Тест", "Тестович", "inspector", None, "Обновленному")
        assert result is True
        
        updated = db.get_all_persons(role="inspector")
        updated_match = [p for p in updated if p['last_name'] == "Обновленный"]
        assert len(updated_match) > 0
