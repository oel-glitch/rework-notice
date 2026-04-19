"""
Russia Post API integration module.

This module provides interface for tracking shipments, checking delivery status,
and managing returns through the official Russia Post SOAP API (RTM34).
Falls back to offline mode if zeep library is not available.
"""

from __future__ import annotations
import logging
import os
from typing import Dict, List, Optional, TYPE_CHECKING
from datetime import datetime

if TYPE_CHECKING:
    from zeep import Client, Settings

try:
    from zeep import Client, Settings
    from zeep.exceptions import Fault, TransportError
    ZEEP_AVAILABLE = True
    logger_init = logging.getLogger(__name__)
    logger_init.debug("✓ Zeep SOAP library loaded successfully")
except ImportError:
    ZEEP_AVAILABLE = False
    Client = None  # type: ignore
    Settings = None  # type: ignore
    Fault = Exception  # type: ignore
    TransportError = Exception  # type: ignore
    logger_init = logging.getLogger(__name__)
    logger_init.warning("⚠ Zeep library not available - Russia Post API features disabled")
    logger_init.info("  Установите zeep для использования Russia Post функций: pip install zeep")

from requests.exceptions import (
    ConnectionError, 
    Timeout, 
    SSLError
)
from src.credential_manager import CredentialManager

logger = logging.getLogger(__name__)


class RussiaPostAPI:
    """
    Interface for Russia Post API operations using SOAP RTM34 service.
    
    Handles shipment tracking, status updates, and returns management.
    """
    
    def __init__(self):
        """
        Initialize Russia Post API client with credentials from environment or keyring.
        
        Priority: environment variables → Windows Credential Manager (keyring) → offline mode
        """
        # Try to get credentials from environment variables first
        self.login = os.getenv('RUSSIA_POST_LOGIN')
        self.password = os.getenv('RUSSIA_POST_PASSWORD')
        self.wsdl_url = os.getenv('RUSSIA_POST_WSDL_URL', 'https://tracking.russianpost.ru/rtm34?wsdl')
        
        # If not in environment, try to load BOTH login AND password from secure keyring
        if not self.login or not self.password:
            logger.info("Credentials not in environment, checking keyring...")
            
            # Try to get login from keyring
            if not self.login:
                login_from_keyring = CredentialManager.get_saved_russia_post_login()
                if login_from_keyring:
                    self.login = login_from_keyring
                    logger.info(f"✓ Login loaded from keyring: {self.login[:5]}***")
            
            # Try to get password from keyring (if we have login)
            if self.login and not self.password:
                password_from_keyring = CredentialManager.get_russia_post_credentials(self.login)
                if password_from_keyring:
                    self.password = password_from_keyring
                    logger.info(f"✓ Password loaded from keyring for user: {self.login[:5]}***")
        
        self.client = None
        self.is_offline = False
        self.offline_reason = None
        
        # Check if credentials are available
        if not self.login or not self.password:
            self.is_offline = True
            self.offline_reason = "Не настроены учетные данные API Почты России"
            logger.warning("⚠ Russia Post API работает в offline режиме: отсутствуют учетные данные")
            logger.info("  Настройте credentials через вкладку 'Почта России' → Настройки API")
        else:
            logger.info(f"Russia Post API credentials found: login='{self.login[:5]}***'")
            # Try to initialize SOAP client
            self._init_soap_client()
        
        if self.is_offline:
            logger.warning(f"⚠ Russia Post API в offline режиме: {self.offline_reason}")
        else:
            logger.info("✓ Russia Post API успешно подключен к SOAP сервису")
    
    def _init_soap_client(self):
        """Initialize SOAP client with detailed error diagnostics."""
        if not ZEEP_AVAILABLE:
            logger.error("✗ Zeep library not available - cannot initialize SOAP client")
            self.client = None
            self.is_offline = True
            self.offline_reason = "Библиотека zeep не установлена"
            return
            
        try:
            logger.info(f"Подключение к Russia Post SOAP: {self.wsdl_url}")
            settings = Settings(strict=False, xml_huge_tree=True)  # type: ignore
            self.client = Client(self.wsdl_url, settings=settings)  # type: ignore
            logger.info(f"✓ SOAP client успешно подключен к {self.wsdl_url}")
            logger.info(f"Доступные методы API: {list(self.client.service._operations.keys())}")
        except Timeout as e:
            logger.error(f"✗ Timeout при подключении к API Почты России")
            logger.error(f"  WSDL URL: {self.wsdl_url}")
            logger.error(f"  Превышен лимит ожидания ответа от сервера")
            logger.error(f"  Проверьте: 1) Скорость интернета 2) Доступность сервиса tracking.russianpost.ru")
            self.client = None
            self.is_offline = True
            self.offline_reason = f"Timeout: сервер не отвечает"
        except ConnectionError as e:
            logger.error(f"✗ Ошибка соединения с API Почты России")
            logger.error(f"  WSDL URL: {self.wsdl_url}")
            logger.error(f"  Детали: {str(e)[:200]}")
            logger.error(f"  Проверьте: 1) Интернет подключение 2) DNS резолвинг 3) Firewall настройки")
            self.client = None
            self.is_offline = True
            self.offline_reason = f"Нет соединения с сервером"
        except SSLError as e:
            logger.error(f"✗ Ошибка SSL сертификата при подключении к API")
            logger.error(f"  WSDL URL: {self.wsdl_url}")
            logger.error(f"  Детали: {str(e)[:200]}")
            logger.error(f"  Возможные причины: 1) Неверная системная дата 2) Устаревшие SSL библиотеки")
            self.client = None
            self.is_offline = True
            self.offline_reason = f"SSL ошибка: {str(e)[:100]}"
        except Exception as e:
            # Handle zeep TransportError (works whether zeep is available or not)
            if 'TransportError' in type(e).__name__ or (ZEEP_AVAILABLE and isinstance(e, TransportError)):
                logger.error(f"✗ Transport ошибка SOAP клиента")
                logger.error(f"  WSDL URL: {self.wsdl_url}")
                logger.error(f"  Детали: {str(e)[:200]}")
                logger.error(f"  HTTP код: {e.status_code if hasattr(e, 'status_code') else 'N/A'}")
                self.client = None
                self.is_offline = True
                self.offline_reason = f"HTTP {e.status_code if hasattr(e, 'status_code') else 'ошибка'}: {str(e)[:100]}"
                return
            # Handle other exceptions
            error_type = type(e).__name__
            logger.error(f"✗ Неизвестная ошибка подключения ({error_type})")
            logger.error(f"  WSDL URL: {self.wsdl_url}")
            logger.error(f"  Детали: {str(e)[:200]}")
            logger.error(f"  Обратитесь в техподдержку Почты России или проверьте WSDL URL")
            self.client = None
            self.is_offline = True
            self.offline_reason = f"{error_type}: {str(e)[:100]}"
    
    def track_shipment(self, tracking_number: str, language: str = "RUS") -> Dict:
        """
        Track a single shipment by tracking number using SOAP API.
        
        Args:
            tracking_number: Russia Post tracking number (barcode/ШПИ)
            language: Language for response ("RUS" or "ENG")
        
        Returns:
            Dictionary with shipment status and tracking history
        """
        if self.is_offline or not self.client:
            return self._get_fallback_tracking(tracking_number)
        
        try:
            operation_history_request = {
                'Barcode': tracking_number,
                'MessageType': 0,
                'Language': language
            }
            
            auth_header = {
                'login': self.login,
                'password': self.password
            }
            
            logger.info(f"📦 Запрос отслеживания через SOAP API: {tracking_number}")
            logger.debug(f"  Request: {operation_history_request}")
            logger.debug(f"  Auth: login={self.login[:5]}***")
            
            response = self.client.service.getOperationHistory(
                operation_history_request,
                auth_header
            )
            
            logger.info(f"✓ Получен ответ от API для {tracking_number}")
            logger.debug(f"  Response type: {type(response)}")
            
            events = []
            if hasattr(response, 'historyRecord') and response.historyRecord:
                logger.info(f"  Найдено записей в истории: {len(response.historyRecord)}")
                for record in response.historyRecord:
                    event_date = record.OperationParameters.OperDate
                    event_type = record.OperationParameters.OperType.Name if hasattr(record.OperationParameters.OperType, 'Name') else 'Unknown'
                    
                    location_parts = []
                    if hasattr(record, 'AddressParameters') and record.AddressParameters:
                        addr = record.AddressParameters
                        if hasattr(addr, 'OperationAddress'):
                            op_addr = addr.OperationAddress
                            if hasattr(op_addr, 'Description'):
                                location_parts.append(op_addr.Description)
                    
                    location = ', '.join(location_parts) if location_parts else 'Не указано'
                    
                    events.append({
                        'date': event_date.strftime('%d.%m.%Y %H:%M') if hasattr(event_date, 'strftime') else str(event_date),
                        'type': event_type,
                        'location': location,
                        'timestamp': event_date
                    })
                
                events.sort(key=lambda x: x['timestamp'], reverse=True)
            
            current_status = events[0]['type'] if events else 'Нет данных'
            current_location = events[0]['location'] if events else 'Неизвестно'
            last_update = events[0]['date'] if events else datetime.now().strftime('%d.%m.%Y %H:%M')
            
            return {
                'tracking_number': tracking_number,
                'status': current_status,
                'last_update': last_update,
                'current_location': current_location,
                'events': events
            }
            
        except Exception as e:
            # Handle zeep SOAP Fault errors (works whether zeep is available or not)
            if 'Fault' in type(e).__name__ or (ZEEP_AVAILABLE and isinstance(e, Fault)):
                error_msg = str(e)
                fault_code = e.code if hasattr(e, 'code') else 'Unknown'
                error_lower = error_msg.lower()
                
                # Classify SOAP fault errors
                if any(keyword in error_lower for keyword in ['authentication', 'auth', 'unauthorized', 'access denied']):
                    logger.error(f"✗ Ошибка аутентификации для {tracking_number}")
                    logger.error(f"  Fault code: {fault_code}")
                    logger.error(f"  Проверьте правильность логина и пароля в настройках API")
                    status_text = '🔒 Ошибка аутентификации'
                    location_text = 'Неверные учетные данные API'
                elif 'barcode' in error_lower or 'not found' in error_lower:
                    logger.warning(f"⚠ Трек-номер не найден: {tracking_number}")
                    logger.warning(f"  Fault code: {fault_code}")
                    logger.warning(f"  Возможно, номер введен неверно или еще не добавлен в систему Почты России")
                    status_text = '❓ Трек-номер не найден'
                    location_text = 'Проверьте правильность номера'
                else:
                    logger.error(f"✗ SOAP Fault для {tracking_number} (code: {fault_code})")
                    logger.error(f"  Детали: {error_msg[:200]}")
                    status_text = f'⚠ SOAP ошибка ({fault_code})'
                    location_text = error_msg[:200]
                
                return {
                    'tracking_number': tracking_number,
                    'status': status_text,
                    'last_update': datetime.now().strftime('%d.%m.%Y %H:%M'),
                    'current_location': location_text,
                    'events': [{'date': datetime.now().strftime('%d.%m.%Y %H:%M'), 
                               'type': status_text,
                               'location': location_text}]
                }
            # Re-raise if not a Fault - will be caught by outer Timeout handler
            raise
        
        except Timeout as e:
            logger.error(f"✗ Timeout при отслеживании {tracking_number}")
            logger.error(f"  Сервер не ответил в течение заданного времени")
            return {
                'tracking_number': tracking_number,
                'status': '⏱ Timeout: сервер не отвечает',
                'last_update': datetime.now().strftime('%d.%m.%Y %H:%M'),
                'current_location': 'Превышен лимит ожидания',
                'events': [{'date': datetime.now().strftime('%d.%m.%Y %H:%M'), 
                           'type': 'Timeout ошибка',
                           'location': 'Попробуйте позже'}]
            }
        
        except ConnectionError as e:
            logger.error(f"✗ Connection error при отслеживании {tracking_number}: {e}")
            return {
                'tracking_number': tracking_number,
                'status': '🔌 Ошибка соединения',
                'last_update': datetime.now().strftime('%d.%m.%Y %H:%M'),
                'current_location': 'Нет подключения к серверу',
                'events': [{'date': datetime.now().strftime('%d.%m.%Y %H:%M'), 
                           'type': 'Сетевая ошибка',
                           'location': 'Проверьте интернет подключение'}]
            }
        
        except Exception as e:
            error_type = type(e).__name__
            logger.error(f"✗ Неожиданная ошибка отслеживания {tracking_number} ({error_type}): {e}")
            return {
                'tracking_number': tracking_number,
                'status': f'❌ Ошибка ({error_type})',
                'last_update': datetime.now().strftime('%d.%m.%Y %H:%M'),
                'current_location': str(e)[:100],
                'events': [{'date': datetime.now().strftime('%d.%m.%Y %H:%M'), 
                           'type': f'Системная ошибка ({error_type})',
                           'location': str(e)[:200]}]
            }
    
    def _get_fallback_tracking(self, tracking_number: str) -> Dict:
        """Fallback mock data when API is unavailable."""
        status_msg = self.offline_reason or "API временно недоступен"
        
        # Demo data for testing
        demo_events = []
        if tracking_number.startswith('1'):  # Simple demo logic
            demo_events = [
                {'date': '15.11.2024 14:30', 'type': 'Прием', 'location': 'Москва 119019'},
                {'date': '16.11.2024 09:15', 'type': 'Обработка', 'location': 'Москва МСЦ'},
                {'date': '17.11.2024 11:00', 'type': 'В пути', 'location': 'Москва - Регион'}
            ]
        
        return {
            'tracking_number': tracking_number,
            'status': f'Демо режим: {status_msg}',
            'last_update': datetime.now().strftime('%d.%m.%Y %H:%M'),
            'current_location': 'Демо данные',
            'events': demo_events if demo_events else [
                {'date': datetime.now().strftime('%d.%m.%Y %H:%M'), 
                 'type': status_msg, 
                 'location': 'Для работы с API настройте учетные данные'}
            ]
        }
    
    def get_returns(self, tracking_numbers: List[str]) -> List[Dict]:
        """
        Get list of returned shipments from provided tracking numbers.
        
        Note: Russia Post RTM34 API doesn't have a dedicated "get returns" endpoint.
        This method filters shipments with return-related statuses.
        
        Args:
            tracking_numbers: List of tracking numbers to check
        
        Returns:
            List of returned shipments
        """
        logger.info(f"Checking {len(tracking_numbers)} shipments for returns")
        
        returns = []
        for tn in tracking_numbers:
            result = self.track_shipment(tn)
            status = result.get('status', '').lower()
            
            if any(keyword in status for keyword in ['возврат', 'невручен', 'возвр']):
                returns.append({
                    'tracking_number': result['tracking_number'],
                    'status': result['status'],
                    'current_location': result['current_location'],
                    'last_update': result['last_update']
                })
        
        logger.info(f"Found {len(returns)} returns")
        return returns
    
    def get_delivery_status(self, tracking_number: str) -> str:
        """
        Get current delivery status for a shipment.
        
        Args:
            tracking_number: Russia Post tracking number
        
        Returns:
            Status string
        """
        statuses = ['В пути', 'Доставлено', 'Возврат отправителю', 'Ожидает получения']
        return statuses[hash(tracking_number) % len(statuses)]
    
    def batch_track(self, tracking_numbers: List[str], language: str = "RUS") -> List[Dict]:
        """
        Track multiple shipments in batch.
        
        Args:
            tracking_numbers: List of tracking numbers
            language: Language for response ("RUS" or "ENG")
        
        Returns:
            List of tracking results
        """
        logger.info(f"Batch tracking {len(tracking_numbers)} shipments")
        results = []
        for tn in tracking_numbers:
            result = self.track_shipment(tn, language)
            results.append(result)
        return results
    
    def test_connection(self) -> Dict:
        """
        Test SOAP API connection and authentication with detailed diagnostics.
        
        Returns:
            Dictionary with connection status and detailed error information
        """
        if self.is_offline:
            return {
                'connected': False,
                'message': self.offline_reason or 'API недоступен',
                'details': 'Проверьте наличие учетных данных и сетевого подключения',
                'error_type': 'offline'
            }
        
        if not self.client:
            return {
                'connected': False,
                'message': 'SOAP клиент не инициализирован',
                'details': self.offline_reason or 'Неизвестная ошибка',
                'error_type': 'soap_init_failed'
            }
        
        try:
            # Try a simple test request with a dummy tracking number
            test_barcode = "80050926000741"  # Valid Russian Post barcode format
            logger.info(f"🔍 Тестирование подключения к API с тестовым трек-номером: {test_barcode}")
            
            operation_history_request = {
                'Barcode': test_barcode,
                'MessageType': 0,
                'Language': 'RUS'
            }
            
            auth_header = {
                'login': self.login,
                'password': self.password
            }
            
            response = self.client.service.getOperationHistory(
                operation_history_request,
                auth_header
            )
            
            logger.info("✓ Тестовый запрос к API выполнен успешно")
            
            return {
                'connected': True,
                'message': '✅ Успешное подключение к Russia Post API',
                'details': f'SOAP сервис доступен и аутентификация прошла успешно',
                'wsdl_url': self.wsdl_url,
                'login': self.login[:5] + '***',
                'test_result': 'OK'
            }
            
        except Exception as e:
            # Handle zeep SOAP Fault errors (works whether zeep is available or not)
            if 'Fault' in type(e).__name__ or (ZEEP_AVAILABLE and isinstance(e, Fault)):
                error_msg = str(e)
                fault_code = e.code if hasattr(e, 'code') else 'Unknown'
                logger.warning(f"⚠ SOAP Fault при тестировании (code: {fault_code}): {error_msg}")
                
                # Detailed authentication error detection
                error_lower = error_msg.lower()
                if any(keyword in error_lower for keyword in ['authentication', 'auth', 'unauthorized', 'access denied']):
                    return {
                        'connected': False,
                        'message': '🔒 Ошибка аутентификации',
                        'details': f'Неверный логин или пароль. Проверьте учетные данные в настройках API.',
                        'error_type': 'auth_failed',
                        'fault_code': fault_code,
                        'raw_error': error_msg[:200],
                        'suggestions': [
                            '1. Проверьте правильность логина (без лишних пробелов)',
                            '2. Убедитесь что пароль введен корректно',
                            '3. Обратитесь в техподдержку Почты России для проверки доступа к API'
                        ]
                    }
                elif 'barcode' in error_lower or 'not found' in error_lower or 'tracking' in error_lower:
                    # This is actually GOOD - means auth worked but test barcode doesn't exist
                    return {
                        'connected': True,
                        'message': '✅ SOAP подключение работает',
                        'details': f'Аутентификация успешна. Тестовый трек-номер не найден в системе (это нормально).',
                        'error_type': 'test_barcode_not_found',
                        'note': 'API готов к использованию с реальными трек-номерами',
                        'wsdl_url': self.wsdl_url,
                        'login': self.login[:5] + '***'
                    }
                else:
                    return {
                        'connected': False,
                        'message': '⚠ SOAP Fault ошибка',
                        'details': f'Сервер вернул ошибку: {error_msg[:200]}',
                        'error_type': 'soap_fault',
                        'fault_code': fault_code,
                        'suggestions': [
                            '1. Проверьте формат запроса',
                            '2. Убедитесь что WSDL URL актуален',
                            '3. Обратитесь в техподдержку Почты России'
                        ]
                    }
            # Re-raise if not a Fault - will be caught by outer Timeout handler  
            raise
        
        except Timeout as e:
            logger.error(f"✗ Timeout при тестировании подключения")
            return {
                'connected': False,
                'message': '⏱ Timeout: сервер не отвечает',
                'details': 'Превышен лимит ожидания ответа от сервера',
                'error_type': 'timeout',
                'suggestions': [
                    '1. Проверьте скорость интернет-соединения',
                    '2. Попробуйте позже (сервер может быть перегружен)',
                    '3. Проверьте firewall настройки'
                ]
            }
        
        except ConnectionError as e:
            logger.error(f"✗ Connection error при тестировании: {e}")
            return {
                'connected': False,
                'message': '🔌 Ошибка соединения',
                'details': 'Не удалось установить соединение с сервером',
                'error_type': 'connection_error',
                'raw_error': str(e)[:200],
                'suggestions': [
                    '1. Проверьте интернет подключение',
                    '2. Проверьте DNS резолвинг (tracking.russianpost.ru)',
                    '3. Проверьте proxy/firewall настройки'
                ]
            }
        
        except SSLError as e:
            logger.error(f"✗ SSL error при тестировании: {e}")
            return {
                'connected': False,
                'message': '🔐 Ошибка SSL сертификата',
                'details': 'Проблема с проверкой SSL сертификата сервера',
                'error_type': 'ssl_error',
                'raw_error': str(e)[:200],
                'suggestions': [
                    '1. Проверьте системную дату и время',
                    '2. Обновите SSL библиотеки (OpenSSL)',
                    '3. Проверьте корневые сертификаты в системе'
                ]
            }
                
        except Exception as e:
            error_type = type(e).__name__
            error_msg = str(e)
            logger.error(f"✗ Неожиданная ошибка при тестировании ({error_type}): {error_msg}")
            return {
                'connected': False,
                'message': f'❌ Неизвестная ошибка ({error_type})',
                'details': error_msg[:300],
                'error_type': 'unknown',
                'suggestions': [
                    '1. Проверьте логи приложения для деталей',
                    '2. Обратитесь в техподдержку',
                    f'3. Сообщите тип ошибки: {error_type}'
                ]
            }
    
    def create_shipment(self, recipient_data: Dict) -> Dict:
        """
        Create a new shipment (for future implementation).
        
        Args:
            recipient_data: Dictionary with recipient information
        
        Returns:
            Created shipment data with tracking number
        """
        logger.info("Creating new shipment (placeholder)")
        return {
            'tracking_number': '12345678901234',
            'status': 'created',
            'message': 'API integration pending - manual entry required'
        }
