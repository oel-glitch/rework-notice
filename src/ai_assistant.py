"""
AI Assistant module for OATI PDF Parser.

Provides AI-powered template selection and text generation using YandexGPT API.
Implements offline-first approach with graceful degradation.
"""

import logging
import os
import json
from typing import Dict, Optional, Tuple, List
from datetime import datetime

# YandexGPT API imports
try:
    import requests
    REQUESTS_AVAILABLE = True
except ImportError:
    REQUESTS_AVAILABLE = False
    logging.warning("requests library not available - AI assistant will work in offline mode")

from src.credential_manager import CredentialManager
from src.config_loader import config_loader


logger = logging.getLogger(__name__)


class AIAssistant:
    """
    AI-powered assistant for document generation using YandexGPT.
    
    Features:
    - Template selection based on recipient configuration
    - Union paragraph text generation
    - Offline-first with graceful degradation
    - Secure credential management
    """
    
    def __init__(self):
        """
        Initialize AI Assistant.
        
        Attempts to load credentials from environment or Windows Credential Manager.
        Falls back to offline mode if credentials unavailable or API unreachable.
        """
        self.api_key: Optional[str] = None
        self.folder_id: Optional[str] = None
        self.is_online: bool = False
        self.offline_reason: Optional[str] = None
        self.config: Dict = {}
        
        # Load AI configuration
        try:
            self.config = config_loader.load('ai_settings')
            logger.info("✓ AI configuration loaded")
        except Exception as e:
            logger.warning(f"⚠ Failed to load AI config: {e}. Using defaults.")
            self.config = self._get_default_config()
        
        # Check if AI is enabled in config
        if not self.config.get('enabled', False):
            self.is_online = False
            self.offline_reason = "AI отключен в настройках"
            logger.info("AI assistant disabled in configuration")
            return
        
        # Try to get credentials from environment first
        self.api_key = os.getenv('YANDEXGPT_API_KEY')
        self.folder_id = os.getenv('YANDEXGPT_FOLDER_ID')
        
        # If not in environment, try to load from keyring
        if not self.api_key or not self.folder_id:
            logger.info("Credentials not in environment, checking keyring...")
            
            try:
                saved_credentials = CredentialManager.get_yandexgpt_credentials()
                if saved_credentials:
                    self.api_key = saved_credentials.get('api_key')
                    self.folder_id = saved_credentials.get('folder_id')
                    logger.info("✓ YandexGPT credentials loaded from keyring")
            except Exception as e:
                logger.warning(f"⚠ Failed to load credentials from keyring: {e}")
        
        # Check if we have credentials
        if not self.api_key or not self.folder_id:
            self.is_online = False
            self.offline_reason = "Не настроены учетные данные YandexGPT API"
            logger.warning("⚠ AI Assistant в offline режиме: отсутствуют учетные данные")
            return
        
        # Check if requests library is available
        if not REQUESTS_AVAILABLE:
            self.is_online = False
            self.offline_reason = "Библиотека requests не установлена"
            logger.warning("⚠ AI Assistant в offline режиме: отсутствует библиотека requests")
            return
        
        # Try to test connection
        logger.info("Attempting to initialize YandexGPT connection...")
        if self._test_connection():
            self.is_online = True
            logger.info("✓ AI Assistant успешно подключен к YandexGPT")
        else:
            logger.warning("⚠ AI Assistant в offline режиме после неудачного подключения")
    
    def _get_default_config(self) -> Dict:
        """Get default AI configuration if config file unavailable."""
        return {
            "enabled": False,
            "provider": "yandexgpt",
            "model": "yandexgpt-lite",
            "temperature": 0.3,
            "max_tokens": 500,
            "timeout": 10
        }
    
    def _test_connection(self) -> bool:
        """
        Test connection to YandexGPT API.
        
        Returns:
            bool: True if connection successful, False otherwise.
        """
        if not REQUESTS_AVAILABLE:
            self.offline_reason = "Библиотека requests не установлена"
            return False
        
        try:
            url = f"https://llm.api.cloud.yandex.net/foundationModels/v1/completion"
            headers = {
                "Authorization": f"Api-Key {self.api_key}",
                "Content-Type": "application/json"
            }
            
            test_payload = {
                "modelUri": f"gpt://{self.folder_id}/yandexgpt-lite",
                "completionOptions": {
                    "temperature": 0.1,
                    "maxTokens": 10
                },
                "messages": [
                    {
                        "role": "user",
                        "text": "test"
                    }
                ]
            }
            
            timeout = self.config.get('timeout', 10)
            response = requests.post(url, headers=headers, json=test_payload, timeout=timeout)
            
            if response.status_code == 200:
                logger.info("✓ YandexGPT API connection test successful")
                return True
            else:
                self.offline_reason = f"API вернул код {response.status_code}"
                logger.error(f"✗ YandexGPT API test failed: {response.status_code}")
                logger.error(f"Response: {response.text[:200]}")
                return False
                
        except requests.Timeout:
            self.offline_reason = "Timeout при подключении к API"
            logger.error("✗ Timeout connecting to YandexGPT API")
            return False
        except requests.ConnectionError:
            self.offline_reason = "Ошибка соединения с API"
            logger.error("✗ Connection error to YandexGPT API")
            return False
        except Exception as e:
            self.offline_reason = f"Ошибка: {str(e)[:100]}"
            logger.error(f"✗ Error testing YandexGPT connection: {e}")
            return False
    
    def test_connection(self) -> Dict:
        """
        Public method to test AI connection and return status.
        
        Returns:
            dict: Status information with 'connected', 'message', 'details' keys.
        """
        if not self.config.get('enabled', False):
            return {
                'connected': False,
                'message': 'AI отключен в настройках',
                'details': 'Включите AI в config/ai_settings.json'
            }
        
        if not self.api_key or not self.folder_id:
            return {
                'connected': False,
                'message': 'Не настроены учетные данные',
                'details': 'Введите API Key и Folder ID в настройках'
            }
        
        if self._test_connection():
            return {
                'connected': True,
                'message': 'Подключение успешно',
                'details': f'YandexGPT API доступен (model: {self.config.get("model", "yandexgpt-lite")})'
            }
        else:
            return {
                'connected': False,
                'message': 'Ошибка подключения',
                'details': self.offline_reason or 'Неизвестная ошибка'
            }
    
    def _call_yandexgpt(self, prompt: str, system_prompt: Optional[str] = None) -> Optional[str]:
        """
        Call YandexGPT API with given prompt.
        
        Args:
            prompt: User prompt text.
            system_prompt: Optional system instruction.
            
        Returns:
            Generated text or None if failed.
        """
        if not self.is_online or not REQUESTS_AVAILABLE:
            logger.warning("AI is offline, cannot call YandexGPT")
            return None
        
        try:
            url = "https://llm.api.cloud.yandex.net/foundationModels/v1/completion"
            headers = {
                "Authorization": f"Api-Key {self.api_key}",
                "Content-Type": "application/json"
            }
            
            messages = []
            if system_prompt:
                messages.append({"role": "system", "text": system_prompt})
            messages.append({"role": "user", "text": prompt})
            
            model = self.config.get('model', 'yandexgpt-lite')
            payload = {
                "modelUri": f"gpt://{self.folder_id}/{model}",
                "completionOptions": {
                    "temperature": self.config.get('temperature', 0.3),
                    "maxTokens": self.config.get('max_tokens', 500)
                },
                "messages": messages
            }
            
            timeout = self.config.get('timeout', 10)
            response = requests.post(url, headers=headers, json=payload, timeout=timeout)
            
            if response.status_code == 200:
                result = response.json()
                alternatives = result.get('result', {}).get('alternatives', [])
                if alternatives:
                    generated_text = alternatives[0].get('message', {}).get('text', '')
                    logger.info(f"✓ YandexGPT response received ({len(generated_text)} chars)")
                    return generated_text
                else:
                    logger.error("✗ No alternatives in YandexGPT response")
                    return None
            else:
                logger.error(f"✗ YandexGPT API error: {response.status_code}")
                logger.error(f"Response: {response.text[:300]}")
                return None
                
        except Exception as e:
            logger.error(f"✗ Error calling YandexGPT: {e}")
            return None
    
    def select_template_with_ai(
        self, 
        num_departments: int,
        has_oati: bool,
        department_names: List[str]
    ) -> Optional[str]:
        """
        Use AI to select appropriate template based on situation.
        
        Args:
            num_departments: Number of departments (excluding OATI).
            has_oati: Whether OATI is involved in handling the appeal.
            department_names: List of department names (for context).
            
        Returns:
            Template name (e.g., 'template_ch3', 'template_ch4_multiple', etc.) or None if offline.
        """
        if not self.is_online:
            logger.info("AI offline, cannot select template")
            return None
        
        # Get prompt from config
        prompts = self.config.get('prompts', {})
        system_prompt = prompts.get('template_selection_system', '')
        user_prompt_template = prompts.get('template_selection_user', '')
        
        if not user_prompt_template:
            logger.error("Template selection prompt not found in config")
            return None
        
        # Format user prompt
        dept_list = ', '.join(department_names) if department_names else 'нет департаментов'
        user_prompt = user_prompt_template.format(
            num_departments=num_departments,
            has_oati='Да' if has_oati else 'Нет',
            department_names=dept_list
        )
        
        logger.info(f"Requesting AI template selection: {num_departments} depts, OATI={has_oati}")
        
        response = self._call_yandexgpt(user_prompt, system_prompt)
        
        if response:
            # Parse response to extract template name
            response_clean = response.strip().lower()
            
            # Map AI response to actual template names
            if 'ch3' in response_clean or 'template_ch3' in response_clean:
                return 'template_ch3'
            elif 'oati_handles' in response_clean or 'ch4_oati' in response_clean:
                return 'template_ch4_oati_handles'
            elif 'multiple' in response_clean or 'ch4_multiple' in response_clean:
                return 'template_ch4_multiple'
            else:
                logger.warning(f"Could not parse template from AI response: {response[:100]}")
                return None
        
        return None
    
    def generate_union_paragraph(
        self,
        question_text: str,
        question_declined: str
    ) -> Optional[str]:
        """
        Generate union paragraph text using AI.
        
        Args:
            question_text: Original question text from resolution.
            question_declined: Question text in accusative case.
            
        Returns:
            Generated paragraph text or None if offline/failed.
        """
        if not self.is_online:
            logger.info("AI offline, cannot generate union paragraph")
            return None
        
        # Get prompts from config
        prompts = self.config.get('prompts', {})
        system_prompt = prompts.get('union_paragraph_system', '')
        user_prompt_template = prompts.get('union_paragraph_user', '')
        
        if not user_prompt_template:
            logger.error("Union paragraph generation prompt not found in config")
            return None
        
        # Format user prompt
        user_prompt = user_prompt_template.format(
            question_text=question_text,
            question_declined=question_declined
        )
        
        logger.info(f"Requesting AI union paragraph generation for: '{question_text[:50]}...'")
        
        response = self._call_yandexgpt(user_prompt, system_prompt)
        
        if response:
            # Clean up response
            paragraph = response.strip()
            
            # Ensure proper formatting (should start with newlines and contain required text)
            if not paragraph.startswith('\n'):
                paragraph = '\n\n' + paragraph
            
            logger.info(f"✓ Generated union paragraph ({len(paragraph)} chars)")
            return paragraph
        
        logger.warning("Failed to generate union paragraph with AI")
        return None
    
    def get_status(self) -> Dict:
        """
        Get current AI assistant status.
        
        Returns:
            dict: Status information including online state, reason, config.
        """
        return {
            'enabled': self.config.get('enabled', False),
            'is_online': self.is_online,
            'offline_reason': self.offline_reason,
            'provider': self.config.get('provider', 'yandexgpt'),
            'model': self.config.get('model', 'yandexgpt-lite'),
            'has_credentials': bool(self.api_key and self.folder_id)
        }
