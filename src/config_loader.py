"""
Configuration loader module for PDF Parser application.
Handles loading and validation of JSON configuration files.
"""

import json
import os
from pathlib import Path
from typing import Dict, Any, Optional


# Embedded fallback configurations for when files are not found
# (e.g., when running as PyInstaller .exe)
EMBEDDED_CONFIGS = {
    "ignored_recipients": {
        "description": "Configuration for filtering recipients from resolution parsing",
        "ignore_prefixes": ["+"],
        "ignore_keywords": ["в дело", "в работу", "на контроль"],
        "service_roles": ["исполнитель", "ответственный исполнитель"],
        "ignored_executor_names": [
            "Воронков", "Воронкова", "Ларин", "Ларина",
            "Аравин", "Аравина", "Фисков", "Фискова"
        ]
    },
    "keywords": {
        "description": "Keywords for parsing and logic determination",
        "union_response": {
            "exclude_keywords": ["НПА", "нпа"],
            "deadline_keywords": ["срок", "сроком", "в срок до"],
            "proximity_threshold": 200
        },
        "resolution_markers": {
            "resolution_start": ["РЕЗОЛЮЦИЯ", "Резолюция"],
            "date_patterns": ["от", "дата"],
            "recipient_patterns": ["кому", "адресат"]
        },
        "portal_sources": {
            "nash_gorod": ["Наш город", "наш город", "nashgorod"],
            "mos_ru": ["mos.ru", "МОС.РУ", "Портал МОС.РУ"]
        }
    },
    "templates_mapping": {
        "description": "Template selection logic based on document type and law part",
        "law_parts": {
            "ch3": {
                "condition": "single_department_without_inspector",
                "template_pattern": "template_ch3*.docx",
                "union_included": False
            },
            "ch4_prefecture": {
                "condition": "prefecture_recipient",
                "template_pattern": "template_ch4_prefecture*.docx",
                "union_included": True,
                "union_suffix": " (dalee Obedinenie)"
            },
            "ch4_multiple": {
                "condition": "multiple_departments_or_inspector",
                "template_pattern": "template_ch4_multiple*.docx",
                "union_included": True,
                "union_suffix": " (dalee Obedinenie)"
            }
        },
        "placeholders": {
            "applicant_name": "{{APPLICANT_NAME}}",
            "applicant_address": "{{APPLICANT_ADDRESS}}",
            "complaint_date": "{{COMPLAINT_DATE}}",
            "complaint_number": "{{COMPLAINT_NUMBER}}",
            "complaint_text": "{{COMPLAINT_TEXT}}",
            "recipient_name": "{{RECIPIENT_NAME}}",
            "recipient_position": "{{RECIPIENT_POSITION}}",
            "department_name": "{{DEPARTMENT_NAME}}",
            "portal_source": "{{PORTAL_SOURCE}}",
            "union_suffix": "{{UNION_SUFFIX}}"
        }
    },
    "ui_settings": {
        "description": "UI configuration for the application",
        "theme": "windows_xp",
        "colors": {
            "background": "#ECE9D8",
            "primary": "#0054E3",
            "secondary": "#D4D0C8",
            "accent": "#B0B0B0",
            "text": "#000000",
            "border": "#808080",
            "button": "#D4D0C8",
            "button_active": "#C0C0C0"
        },
        "fonts": {
            "main": {"family": "Tahoma", "size": 9, "weight": "normal"},
            "header": {"family": "Tahoma", "size": 12, "weight": "bold"},
            "button": {"family": "Tahoma", "size": 9, "weight": "bold"},
            "console": {"family": "Consolas", "size": 9, "weight": "normal"}
        },
        "window": {
            "width": 1100,
            "height": 750,
            "title": "PDF Parser and Word Document Generator - OATI"
        },
        "button_labels": {
            "select_pdf": "Select PDF Files",
            "process": "Process Documents",
            "clear": "Clear List",
            "directors": "Manage Directors",
            "inspector_chiefs": "Manage Inspector Chiefs"
        }
    },
    "inspectors_chiefs": {
        "inspectors_chiefs": [
            {"last_name_nominative": "Авдеев", "last_name_dative": "Авдееву", "first_initial": "В", "middle_initial": "Ю", "organization": "Объединение административно-технических инспекций города Москвы"},
            {"last_name_nominative": "Александров", "last_name_dative": "Александрову", "first_initial": "А", "middle_initial": "С", "organization": "Объединение административно-технических инспекций города Москвы"},
            {"last_name_nominative": "Алёшина", "last_name_dative": "Алёшиной", "first_initial": "О", "middle_initial": "В", "organization": "Объединение административно-технических инспекций города Москвы"},
            {"last_name_nominative": "Блинов", "last_name_dative": "Блинову", "first_initial": "Р", "middle_initial": "С", "organization": "Объединение административно-технических инспекций города Москвы"},
            {"last_name_nominative": "Кичиков", "last_name_dative": "Кичикову", "first_initial": "Б", "middle_initial": "Б", "organization": "Объединение административно-технических инспекций города Москвы"},
            {"last_name_nominative": "Кравченко", "last_name_dative": "Кравченко", "first_initial": "И", "middle_initial": "С", "organization": "Объединение административно-технических инспекций города Москвы"},
            {"last_name_nominative": "Лобанов", "last_name_dative": "Лобанову", "first_initial": "С", "middle_initial": "Д", "organization": "Объединение административно-технических инспекций города Москвы"},
            {"last_name_nominative": "Мишин", "last_name_dative": "Мишину", "first_initial": "И", "middle_initial": "А", "organization": "Объединение административно-технических инспекций города Москвы"},
            {"last_name_nominative": "Тоненьков", "last_name_dative": "Тоненькову", "first_initial": "И", "middle_initial": "А", "organization": "Объединение административно-технических инспекций города Москвы"},
            {"last_name_nominative": "Цатурян", "last_name_dative": "Цатуряну", "first_initial": "Г", "middle_initial": "Б", "organization": "Объединение административно-технических инспекций города Москвы"},
            {"last_name_nominative": "Шинин", "last_name_dative": "Шинину", "first_initial": "С", "middle_initial": "Ю", "organization": "Объединение административно-технических инспекций города Москвы"}
        ]
    },
    "directors": {
        "directors": [
            {"last_name": "Хрипун", "initials": "А.И.", "department": "Департамент здравоохранения города Москвы"},
            {"last_name": "Каклюгина", "initials": "И.А.", "department": "Департамент образования и науки города Москвы"},
            {"last_name": "Кибовский", "initials": "А.В.", "department": "Департамент культуры города Москвы"},
            {"last_name": "Кондаранцев", "initials": "А.А.", "department": "Департамент спорта города Москвы"},
            {"last_name": "Ликсутов", "initials": "М.С.", "department": "Департамент транспорта и развития дорожно-транспортной инфраструктуры города Москвы"},
            {"last_name": "Багреева", "initials": "Е.А.", "department": "Департамент экономической политики и развития города Москвы"},
            {"last_name": "Соловьёва", "initials": "Е.А.", "department": "Департамент городского имущества города Москвы"},
            {"last_name": "Лысенко", "initials": "Э.А.", "department": "Департамент информационных технологий города Москвы"},
            {"last_name": "Фурсин", "initials": "А.О.", "department": "Департамент строительства города Москвы"},
            {"last_name": "Лысенко", "initials": "Э.С.", "department": "Департамент информационных технологий города Москвы"},
            {"last_name": "Алёшин", "initials": "Н.В.", "department": "префектура Восточного административного округа города Москвы"},
            {"last_name": "Александров", "initials": "А.О.", "department": "префектура Западного административного округа города Москвы"},
            {"last_name": "Изутдинов", "initials": "Г.И.", "department": "префектура Северного административного округа города Москвы"},
            {"last_name": "Мельников", "initials": "С.А.", "department": "префектура Северо-Восточного административного округа города Москвы"},
            {"last_name": "Пашков", "initials": "В.В.", "department": "префектура Северо-Западного административного округа города Москвы"},
            {"last_name": "Говердовский", "initials": "В.В.", "department": "префектура Центрального административного округа города Москвы"},
            {"last_name": "Волков", "initials": "О.А.", "department": "префектура Юго-Западного административного округа города Москвы"},
            {"last_name": "Челышев", "initials": "Ю.А.", "department": "префектура Южного административного округа города Москвы"},
            {"last_name": "Цыбин", "initials": "А.П.", "department": "префектура Юго-Восточного административного округа города Москвы"},
            {"last_name": "Набокин", "initials": "А.А.", "department": "префектура Троицкого и Новомосковского административных округов города Москвы"},
            {"last_name": "Смирнов", "initials": "А.Н.", "department": "префектура Зеленоградского административного округа города Москвы"},
            {"last_name": "Рябов", "initials": "И.Н.", "department": "Государственная инспекция по контролю за использованием объектов недвижимости города Москвы"},
            {"last_name": "Погребняк", "initials": "А.В.", "department": "Московская административная дорожная инспекция города Москвы"},
            {"last_name": "Кузьмин", "initials": "А.В.", "department": "Комитет по архитектуре и градостроительству города Москвы"},
            {"last_name": "Локтев", "initials": "О.А.", "department": "Комитет государственного строительного надзора города Москвы"},
            {"last_name": "Громова", "initials": "Е.В.", "department": "Комитет общественных связей и молодёжной политики города Москвы"},
            {"last_name": "Пятова", "initials": "А.В.", "department": "Комитет по ценовой политике в строительстве и государственной экспертизе проектов города Москвы"},
            {"last_name": "Урожаева", "initials": "Ю.В.", "department": "Департамент природопользования и охраны окружающей среды города Москвы"},
            {"last_name": "Кичиков", "initials": "О.В.", "department": "Государственная жилищная инспекция города Москвы"},
            {"last_name": "Чигликов", "initials": "Р.Р.", "department": "Департамент жилищно-коммунального хозяйства города Москвы"},
            {"last_name": "Слободчиков", "initials": "А.О.", "department": "Комитет государственного строительного надзора города Москвы"}
        ]
    }
}


class ConfigLoader:
    """Loads and manages application configuration from JSON files."""
    
    def __init__(self, config_dir: str = "config"):
        """
        Initialize configuration loader.
        
        Args:
            config_dir: Directory containing configuration files
        """
        self.config_dir = Path(config_dir)
        self._configs: Dict[str, Dict[str, Any]] = {}
        
    def load(self, config_name: str) -> Dict[str, Any]:
        """
        Load configuration from JSON file, with embedded fallback.
        
        Args:
            config_name: Name of configuration file (without .json extension)
            
        Returns:
            Dictionary containing configuration data
            
        Raises:
            FileNotFoundError: If configuration file and embedded config both don't exist
            json.JSONDecodeError: If configuration file is invalid JSON
        """
        if config_name in self._configs:
            return self._configs[config_name]
            
        config_path = self.config_dir / f"{config_name}.json"
        
        # Try to load from file first
        if config_path.exists():
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    config_data = json.load(f)
                self._configs[config_name] = config_data
                return config_data
            except (json.JSONDecodeError, IOError) as e:
                print(f"Warning: Failed to load {config_path}: {e}")
                # Fall through to embedded config
        
        # Use embedded configuration as fallback
        if config_name in EMBEDDED_CONFIGS:
            print(f"Using embedded configuration for: {config_name}")
            config_data = EMBEDDED_CONFIGS[config_name]
            self._configs[config_name] = config_data
            return config_data
        
        # If neither file nor embedded config exists
        raise FileNotFoundError(
            f"Configuration '{config_name}' not found in files or embedded configs"
        )
    
    def get(self, config_name: str, key_path: str, default: Any = None) -> Any:
        """
        Get configuration value by key path.
        
        Args:
            config_name: Name of configuration file
            key_path: Dot-separated path to configuration key (e.g., 'colors.primary')
            default: Default value if key not found
            
        Returns:
            Configuration value or default
        """
        config = self.load(config_name)
        
        keys = key_path.split('.')
        value = config
        
        for key in keys:
            if isinstance(value, dict) and key in value:
                value = value[key]
            else:
                return default
                
        return value
    
    def reload(self, config_name: Optional[str] = None) -> None:
        """
        Reload configuration from disk.
        
        Args:
            config_name: Specific configuration to reload, or None to reload all
        """
        if config_name:
            self._configs.pop(config_name, None)
        else:
            self._configs.clear()


# Global configuration loader instance
config_loader = ConfigLoader()
