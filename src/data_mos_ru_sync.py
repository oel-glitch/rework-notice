import requests
import json
import logging
from typing import Any, Dict, List, Optional, Tuple
from datetime import datetime

logger = logging.getLogger(__name__)

API_BASE_URL = "https://apidata.mos.ru/v1"

DATASET_REGISTRY_INSTITUTIONS = "7710152113-reestr-uchrejdeniy-goroda-moskvy"
DATASET_BUDGET_PARTICIPANTS = "7710152113-moskovskiy-reestr-uchastnikov-i-neuchastnikov-byudjetnogo-protsessa"

MAX_ROWS_PER_REQUEST = 500


class DataMosRuAPI:
    def __init__(self, api_key: Optional[str] = None):
        self.api_key = api_key
        self.session = requests.Session()
        if api_key:
            self.session.params = {'api_key': api_key}
    
    def set_api_key(self, api_key: str):
        self.api_key = api_key
        self.session.params = {'api_key': api_key}
    
    def get_dataset_info(self, dataset_id: str) -> Optional[Dict]:
        url = f"{API_BASE_URL}/datasets/{dataset_id}"
        try:
            response = self.session.get(url, timeout=30)
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            logger.error(f"Ошибка получения информации о датасете {dataset_id}: {e}")
            return None
    
    def get_dataset_rows(self, dataset_id: str, top: int = 500, skip: int = 0, 
                        orderby: Optional[str] = None) -> Optional[List[Dict]]:
        url = f"{API_BASE_URL}/datasets/{dataset_id}/rows"
        params: Dict[str, Any] = {
            '$top': min(top, MAX_ROWS_PER_REQUEST),
            '$skip': skip
        }
        if orderby:
            params['$orderby'] = orderby
        
        try:
            response = self.session.get(url, params=params, timeout=60)
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            logger.error(f"Ошибка получения данных датасета {dataset_id}: {e}")
            return None
    
    def get_all_dataset_rows(self, dataset_id: str, 
                            progress_callback=None) -> List[Dict]:
        """
        Load all dataset rows into memory.
        
        Note: For very large datasets (>10k rows), consider using 
        iter_dataset_rows() to process data in chunks without loading
        everything into memory at once.
        
        Args:
            dataset_id: Dataset identifier
            progress_callback: Optional callback(loaded, total) for progress updates
            
        Returns:
            List of all rows
        """
        all_rows = []
        skip = 0
        
        info = self.get_dataset_info(dataset_id)
        total_items = info.get('ItemsCount', 0) if info else 0
        
        logger.info(f"Начало загрузки датасета {dataset_id}, всего записей: {total_items}")
        
        while True:
            rows = self.get_dataset_rows(dataset_id, top=MAX_ROWS_PER_REQUEST, skip=skip)
            
            if not rows:
                break
            
            all_rows.extend(rows)
            skip += len(rows)
            
            if progress_callback:
                progress_callback(len(all_rows), total_items)
            
            logger.info(f"Загружено {len(all_rows)} из {total_items} записей")
            
            if len(rows) < MAX_ROWS_PER_REQUEST:
                break
        
        logger.info(f"Загрузка завершена: {len(all_rows)} записей")
        return all_rows
    
    def iter_dataset_rows(self, dataset_id: str, 
                          progress_callback=None):
        """
        Iterate over dataset rows in chunks for memory-efficient processing.
        
        Use this method for very large datasets (>10k rows) to avoid loading
        everything into memory at once. Yields chunks of rows as they're fetched.
        
        Args:
            dataset_id: Dataset identifier
            progress_callback: Optional callback(loaded, total) for progress updates
            
        Yields:
            Chunks of rows (each chunk up to MAX_ROWS_PER_REQUEST rows)
            
        Example:
            >>> api = DataMosRuAPI(api_key)
            >>> total_processed = 0
            >>> for chunk in api.iter_dataset_rows('dataset-id'):
            ...     process_chunk(chunk)
            ...     total_processed += len(chunk)
        """
        skip = 0
        loaded = 0
        
        info = self.get_dataset_info(dataset_id)
        total_items = info.get('ItemsCount', 0) if info else 0
        
        logger.info(f"Начало потоковой загрузки датасета {dataset_id}, всего записей: {total_items}")
        
        while True:
            rows = self.get_dataset_rows(dataset_id, top=MAX_ROWS_PER_REQUEST, skip=skip)
            
            if not rows:
                break
            
            skip += len(rows)
            loaded += len(rows)
            
            if progress_callback:
                progress_callback(loaded, total_items)
            
            logger.info(f"Загружено {loaded} из {total_items} записей")
            
            yield rows  # Yield chunk instead of accumulating
            
            if len(rows) < MAX_ROWS_PER_REQUEST:
                break
        
        logger.info(f"Потоковая загрузка завершена: {loaded} записей")
    
    def get_moscow_institutions(self, progress_callback=None) -> List[Dict]:
        raw_rows = self.get_all_dataset_rows(DATASET_REGISTRY_INSTITUTIONS, progress_callback)
        return self._parse_institution_rows(raw_rows)
    
    def _parse_institution_rows(self, raw_rows: List[Dict]) -> List[Dict]:
        institutions = []
        
        for row in raw_rows:
            cells = row.get('Cells', {})
            
            institution = {
                'row_number': row.get('Number'),
                'name': cells.get('FullName', cells.get('ShortName', '')),
                'short_name': cells.get('ShortName', ''),
                'inn': cells.get('INN', ''),
                'ogrn': cells.get('OGRN', ''),
                'address': cells.get('Address', cells.get('LegalAddress', '')),
                'department': cells.get('Department', ''),
                'type': cells.get('Type', ''),
                'is_active': cells.get('IsActive', 1)
            }
            
            institutions.append(institution)
        
        return institutions
    
    def test_connection(self) -> Tuple[bool, str]:
        if not self.api_key:
            return False, "API ключ не установлен"
        
        try:
            info = self.get_dataset_info(DATASET_REGISTRY_INSTITUTIONS)
            if info:
                return True, f"Соединение успешно. Датасет: {info.get('Caption', 'N/A')}"
            else:
                return False, "Не удалось получить информацию о датасете"
        except Exception as e:
            return False, f"Ошибка соединения: {str(e)}"


def sync_organizations_from_data_mos_ru(db_instance, api_key: str, 
                                        progress_callback=None,
                                        log_callback=None) -> Tuple[int, int, int]:
    def log(message: str):
        logger.info(message)
        if log_callback:
            log_callback(message)
    
    log("Начало синхронизации с data.mos.ru...")
    
    api = DataMosRuAPI(api_key)
    
    success, message = api.test_connection()
    if not success:
        log(f"ОШИБКА: {message}")
        return 0, 0, 0
    
    log(message)
    
    institutions = api.get_moscow_institutions(progress_callback)
    log(f"Получено {len(institutions)} учреждений из data.mos.ru")
    
    added_count = 0
    updated_count = 0
    skipped_count = 0
    
    for inst in institutions:
        name = inst['name'].strip()
        short_name = inst['short_name'].strip()
        inn = inst['inn'].strip()
        ogrn = inst['ogrn'].strip()
        address = inst['address'].strip()
        
        if not name:
            skipped_count += 1
            continue
        
        existing = None
        
        if inn:
            existing = db_instance.get_department_by_inn(inn)
            if existing:
                log(f"Найдено по ИНН: {name}")
        
        if not existing and ogrn:
            existing = db_instance.get_department_by_ogrn(ogrn)
            if existing:
                log(f"Найдено по ОГРН: {name}")
        
        if not existing:
            existing = db_instance.get_department_by_name(name)
            if existing:
                log(f"Найдено по названию: {name}")
        
        if existing:
            dept_id, existing_name, existing_short, existing_inn, existing_ogrn, existing_addr, _ = existing
            
            needs_update = (
                existing_name != name or
                existing_short != short_name or
                existing_inn != inn or
                existing_ogrn != ogrn or
                existing_addr != address
            )
            
            if needs_update:
                db_instance.update_department(dept_id, name, short_name, inn, ogrn, address)
                updated_count += 1
                log(f"Обновлено: {name}")
            else:
                skipped_count += 1
        else:
            db_instance.add_department(name, short_name, inn, ogrn, address)
            added_count += 1
            log(f"Добавлено: {name}")
    
    log(f"""
Синхронизация завершена:
- Добавлено: {added_count}
- Обновлено: {updated_count}
- Пропущено: {skipped_count}
- Всего: {len(institutions)}
""")
    
    return added_count, updated_count, skipped_count
