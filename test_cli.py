#!/usr/bin/env python3

import sys
import os
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from src.pdf_parser import PDFParser
from src.name_declension import NameDeclension
from src.database import Database
from src.word_generator import WordGenerator


def main():
    print("=" * 60)
    print("  PDF Парсер и Генератор Word Документов - Демо")
    print("=" * 60)
    print()
    
    test_pdf = "attached_assets/13.10.2025_01-21-П-9141_25_Обращение_граждан_Ларин_А.С._1760536596725.pdf"
    
    if not os.path.exists(test_pdf):
        print(f"❌ Тестовый PDF файл не найден: {test_pdf}")
        return
    
    print("📄 Парсинг PDF файла...")
    parser = PDFParser()
    
    try:
        citizen_data = parser.parse_pdf(test_pdf)
        
        print("\n✅ Данные извлечены успешно:")
        print(f"   ФИО: {citizen_data['full_name']}")
        print(f"   Email: {citizen_data['email']}")
        print(f"   ID портала: {citizen_data['portal_id']}")
        print(f"   Номер ОАТИ: {citizen_data['oati_number']}")
        print(f"   Дата: {citizen_data['oati_date']}")
        print(f"   Часть закона: ч.{citizen_data['law_part']}")
        print(f"   Департаменты: {', '.join(citizen_data['departments'])}")
        
        print("\n🔄 Склонение имени...")
        decliner = NameDeclension()
        declined = decliner.decline_full_name(
            citizen_data['last_name'],
            citizen_data['first_name'],
            citizen_data['middle_name']
        )
        
        short_name_dative = decliner.get_short_name_dative(
            citizen_data['last_name'],
            citizen_data['first_name'],
            citizen_data['middle_name']
        )
        
        salutation = decliner.get_full_salutation(
            citizen_data['first_name'],
            citizen_data['middle_name'],
            declined['gender']
        )
        
        print(f"\n✅ Склонение выполнено:")
        print(f"   Пол: {declined['gender']}")
        print(f"   Полное ФИО (дательный): {declined['full_name']}")
        print(f"   Короткое ФИО (дательный): {short_name_dative}")
        print(f"   Обращение: {salutation}")
        
        print("\n📝 Проверка базы данных...")
        db = Database()
        departments = db.get_all_departments()
        print(f"✅ В базе данных {len(departments)} департаментов")
        
        print("\n📄 Формирование параметров для шаблона...")
        portal_source = parser.extract_portal_source()
        
        departments_list = ''
        if citizen_data['law_part'] == '4' and citizen_data['departments']:
            declined_depts = [decliner.decline_text_to_accusative(dept).replace('в ', '', 1) for dept in citizen_data['departments']]
            departments_list = ', '.join(declined_depts)
        elif citizen_data['law_part'] == '3':
            departments_list = decliner.decline_text_to_accusative('Объединение административно-технических инспекций города Москвы').replace('в ', '', 1)
        
        print(f"   Источник портала: {portal_source}")
        print(f"   Список департаментов: {departments_list}")
        
        print("\n📄 Генерация Word документа...")
        generator = WordGenerator()
        
        declined_with_short = declined.copy()
        declined_with_short['full_name'] = short_name_dative
        
        output_file = generator.process_citizen_document(
            citizen_data,
            declined_with_short,
            salutation,
            citizen_data['departments'],
            portal_source=portal_source,
            departments_list=departments_list
        )
        
        print(f"\n✅ Документ успешно создан:")
        print(f"   {output_file}")
        
        db.add_processing_record(
            test_pdf,
            citizen_data['full_name'],
            citizen_data['oati_number'],
            citizen_data['portal_id'],
            output_file,
            'success'
        )
        
        print("\n" + "=" * 60)
        print("  ✨ Демонстрация завершена успешно!")
        print("=" * 60)
        print("\nДля использования GUI приложения запустите: python main.py")
        
    except Exception as e:
        print(f"\n❌ Ошибка: {str(e)}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
