"""
Credential Manager for secure storage of API credentials.

Uses Windows Credential Manager via keyring library for secure password storage.
Falls back to local encoded JSON file if keyring is not available.
"""

import logging
import json
import base64
from pathlib import Path
from typing import Optional, Tuple

logger = logging.getLogger(__name__)

try:
    import keyring
    KEYRING_AVAILABLE = True
    logger.debug("✓ Keyring library loaded successfully")
except ImportError:
    KEYRING_AVAILABLE = False
    keyring = None
    logger.warning("⚠ Keyring library not available - credential storage features disabled")
    logger.info("  Установите keyring для использования Russia Post и MOSEDO функций: pip install keyring")


class LocalCredentialStorage:
    """
    Fallback credential storage using local JSON file with base64 encoding.
    
    This is NOT secure encryption, just obfuscation. Use only as fallback when
    Windows Credential Manager is unavailable.
    """
    
    CREDENTIALS_FILE = Path("config") / ".credentials.json"
    
    @staticmethod
    def _encode(text: str) -> str:
        """Encode text using base64."""
        return base64.b64encode(text.encode('utf-8')).decode('utf-8')
    
    @staticmethod
    def _decode(encoded: str) -> str:
        """Decode base64 text."""
        return base64.b64decode(encoded.encode('utf-8')).decode('utf-8')
    
    @staticmethod
    def save(service: str, key: str, value: str) -> bool:
        """Save credential to local file."""
        try:
            LocalCredentialStorage.CREDENTIALS_FILE.parent.mkdir(parents=True, exist_ok=True)
            
            # Load existing credentials
            if LocalCredentialStorage.CREDENTIALS_FILE.exists():
                with open(LocalCredentialStorage.CREDENTIALS_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
            else:
                data = {}
            
            # Ensure service exists in data
            if service not in data:
                data[service] = {}
            
            # Encode and save
            data[service][key] = LocalCredentialStorage._encode(value)
            
            with open(LocalCredentialStorage.CREDENTIALS_FILE, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2)
            
            logger.info(f"✓ Credential saved to local storage: {service}/{key}")
            return True
        except Exception as e:
            logger.error(f"✗ Failed to save to local storage: {e}")
            return False
    
    @staticmethod
    def get(service: str, key: str) -> Optional[str]:
        """Get credential from local file."""
        try:
            if not LocalCredentialStorage.CREDENTIALS_FILE.exists():
                return None
            
            with open(LocalCredentialStorage.CREDENTIALS_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            if service in data and key in data[service]:
                return LocalCredentialStorage._decode(data[service][key])
            return None
        except Exception as e:
            logger.warning(f"⚠ Failed to read from local storage: {e}")
            return None
    
    @staticmethod
    def delete(service: str, key: str) -> bool:
        """Delete credential from local file."""
        try:
            if not LocalCredentialStorage.CREDENTIALS_FILE.exists():
                return True
            
            with open(LocalCredentialStorage.CREDENTIALS_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            if service in data and key in data[service]:
                del data[service][key]
                
                # Remove service if empty
                if not data[service]:
                    del data[service]
                
                with open(LocalCredentialStorage.CREDENTIALS_FILE, 'w', encoding='utf-8') as f:
                    json.dump(data, f, indent=2)
                
                logger.info(f"✓ Credential deleted from local storage: {service}/{key}")
            return True
        except Exception as e:
            logger.error(f"✗ Failed to delete from local storage: {e}")
            return False


class CredentialManager:
    """
    Manage secure storage of API credentials using OS keyring.
    
    Uses Windows Credential Manager on Windows for secure credential storage.
    """
    
    # Service identifiers for different APIs
    RUSSIA_POST_SERVICE = "OATI_RussiaPost_API"
    RUSSIA_POST_LOGIN_KEY = "OATI_RussiaPost_Login"  # Special key for storing login
    MOSEDO_SERVICE = "OATI_MOSEDO"
    MOSEDO_LOGIN_KEY = "OATI_MOSEDO_Login"  # Special key for storing login
    YANDEXGPT_SERVICE = "OATI_YandexGPT_API"
    YANDEXGPT_API_KEY = "OATI_YandexGPT_APIKey"  # Special key for API Key
    YANDEXGPT_FOLDER_ID = "OATI_YandexGPT_FolderID"  # Special key for Folder ID
    
    @staticmethod
    def save_russia_post_credentials(login: str, password: str) -> bool:
        """
        Save Russia Post API credentials securely.
        
        Saves both login AND password to keyring for complete offline support.
        
        Args:
            login: API login/username
            password: API password
            
        Returns:
            True if saved successfully, False otherwise
        """
        if not KEYRING_AVAILABLE:
            logger.warning("⚠ Keyring недоступен - невозможно сохранить credentials")
            logger.info("  Используйте переменные окружения RUSSIA_POST_LOGIN и RUSSIA_POST_PASSWORD")
            return False
            
        try:
            # Save password with login as username
            keyring.set_password(  # type: ignore
                CredentialManager.RUSSIA_POST_SERVICE,
                login,
                password
            )
            # Also save login itself for retrieval without env vars
            keyring.set_password(  # type: ignore
                CredentialManager.RUSSIA_POST_SERVICE,
                CredentialManager.RUSSIA_POST_LOGIN_KEY,
                login
            )
            logger.info(f"✓ Russia Post credentials (login + password) saved for user: {login[:5]}***")
            return True
        except Exception as e:
            logger.error(f"✗ Failed to save Russia Post credentials: {e}")
            return False
    
    @staticmethod
    def get_russia_post_credentials(login: str) -> Optional[str]:
        """
        Retrieve Russia Post API password for given login.
        
        Args:
            login: API login/username
            
        Returns:
            Password if found, None otherwise
        """
        if not KEYRING_AVAILABLE:
            logger.debug(f"⚠ Keyring недоступен - credentials не найдены для user: {login}")
            return None
            
        try:
            password = keyring.get_password(  # type: ignore
                CredentialManager.RUSSIA_POST_SERVICE,
                login
            )
            if password:
                logger.info(f"✓ Russia Post credentials retrieved for user: {login[:5]}***")
            else:
                logger.info(f"⚠ No Russia Post credentials found for user: {login}")
            return password
        except Exception as e:
            logger.warning(f"⚠ Keyring backend unavailable: {str(e)[:100]}")
            logger.info("  На Replit/Linux keyring работает с ограничениями. Используйте environment variables.")
            return None
    
    @staticmethod
    def delete_russia_post_credentials(login: str) -> bool:
        """
        Delete Russia Post API credentials (both login and password).
        
        Args:
            login: API login/username
            
        Returns:
            True if deleted successfully, False otherwise
        """
        if not KEYRING_AVAILABLE:
            logger.warning("⚠ Keyring недоступен - нечего удалять")
            return False
            
        success = True
        try:
            # Delete password
            keyring.delete_password(  # type: ignore
                CredentialManager.RUSSIA_POST_SERVICE,
                login
            )
            logger.info(f"✓ Russia Post password deleted for user: {login}")
        except Exception as e:
            if "PasswordDeleteError" in type(e).__name__ or "not found" in str(e).lower():
                logger.warning(f"⚠ No password found to delete for user: {login}")
            else:
                logger.error(f"✗ Failed to delete Russia Post password: {e}")
            success = False
        
        try:
            # Delete saved login
            keyring.delete_password(  # type: ignore
                CredentialManager.RUSSIA_POST_SERVICE,
                CredentialManager.RUSSIA_POST_LOGIN_KEY
            )
            logger.info(f"✓ Russia Post login deleted from keyring")
        except Exception as e:
            if "PasswordDeleteError" in type(e).__name__ or "not found" in str(e).lower():
                logger.warning(f"⚠ No login found to delete")
            else:
                logger.error(f"✗ Failed to delete Russia Post login: {e}")
            success = False
        
        return success
    
    @staticmethod
    def get_saved_russia_post_login() -> Optional[str]:
        """
        Get the saved Russia Post login from keyring.
        
        Returns:
            Saved login if found, None otherwise
        """
        if not KEYRING_AVAILABLE:
            logger.debug("⚠ Keyring недоступен - saved login не найден")
            return None
            
        try:
            login = keyring.get_password(  # type: ignore
                CredentialManager.RUSSIA_POST_SERVICE,
                CredentialManager.RUSSIA_POST_LOGIN_KEY
            )
            if login:
                logger.info(f"✓ Russia Post login retrieved from keyring: {login[:5]}***")
            return login
        except Exception as e:
            logger.warning(f"⚠ Keyring backend unavailable for login retrieval: {str(e)[:100]}")
            return None
    
    @staticmethod
    def has_russia_post_credentials(login: str) -> bool:
        """
        Check if Russia Post credentials exist for given login.
        
        Args:
            login: API login/username
            
        Returns:
            True if credentials exist, False otherwise
        """
        password = CredentialManager.get_russia_post_credentials(login)
        return password is not None
    
    @staticmethod
    def save_mosedo_credentials(username: str, password: str) -> bool:
        """
        Save MOSEDO credentials securely.
        
        Saves both username AND password to keyring for complete offline support.
        
        Args:
            username: MOSEDO username
            password: MOSEDO password
            
        Returns:
            True if saved successfully, False otherwise
        """
        if not KEYRING_AVAILABLE:
            logger.warning("⚠ Keyring недоступен - невозможно сохранить MOSEDO credentials")
            logger.info("  Используйте переменные окружения MOSEDO_USERNAME и MOSEDO_PASSWORD")
            return False
            
        try:
            # Save password with username as key
            keyring.set_password(  # type: ignore
                CredentialManager.MOSEDO_SERVICE,
                username,
                password
            )
            # Also save username itself for retrieval without env vars
            keyring.set_password(  # type: ignore
                CredentialManager.MOSEDO_SERVICE,
                CredentialManager.MOSEDO_LOGIN_KEY,
                username
            )
            logger.info(f"✓ MOSEDO credentials (username + password) saved for user: {username[:5]}***")
            return True
        except Exception as e:
            logger.error(f"✗ Failed to save MOSEDO credentials: {e}")
            return False
    
    @staticmethod
    def get_mosedo_credentials(username: str) -> Optional[str]:
        """
        Retrieve MOSEDO password for given username.
        
        Args:
            username: MOSEDO username
            
        Returns:
            Password if found, None otherwise
        """
        if not KEYRING_AVAILABLE:
            logger.debug(f"⚠ Keyring недоступен - MOSEDO credentials не найдены для user: {username}")
            return None
            
        try:
            password = keyring.get_password(  # type: ignore
                CredentialManager.MOSEDO_SERVICE,
                username
            )
            if password:
                logger.info(f"✓ MOSEDO credentials retrieved for user: {username[:5]}***")
            else:
                logger.warning(f"⚠ No MOSEDO credentials found for user: {username}")
            return password
        except Exception as e:
            logger.error(f"✗ Failed to retrieve MOSEDO credentials: {e}")
            return None
    
    @staticmethod
    def delete_mosedo_credentials(username: str) -> bool:
        """
        Delete MOSEDO credentials from secure storage (both username and password).
        
        Args:
            username: MOSEDO username
            
        Returns:
            True if deleted successfully, False otherwise
        """
        if not KEYRING_AVAILABLE:
            logger.warning("⚠ Keyring недоступен - нечего удалять")
            return False
            
        success = True
        try:
            # Delete password
            keyring.delete_password(  # type: ignore
                CredentialManager.MOSEDO_SERVICE,
                username
            )
            logger.info(f"✓ MOSEDO password deleted for user: {username}")
        except Exception as e:
            if "PasswordDeleteError" in type(e).__name__ or "not found" in str(e).lower():
                logger.warning(f"⚠ No MOSEDO password found to delete for user: {username}")
            else:
                logger.error(f"✗ Failed to delete MOSEDO password: {e}")
            success = False
        
        try:
            # Delete saved username
            keyring.delete_password(  # type: ignore
                CredentialManager.MOSEDO_SERVICE,
                CredentialManager.MOSEDO_LOGIN_KEY
            )
            logger.info(f"✓ MOSEDO username deleted from keyring")
        except Exception as e:
            if "PasswordDeleteError" in type(e).__name__ or "not found" in str(e).lower():
                logger.warning(f"⚠ No MOSEDO username found to delete")
            else:
                logger.error(f"✗ Failed to delete MOSEDO username: {e}")
            success = False
        
        return success
    
    @staticmethod
    def get_saved_mosedo_login() -> Optional[str]:
        """
        Get the saved MOSEDO username from keyring.
        
        Returns:
            Saved username if found, None otherwise
        """
        if not KEYRING_AVAILABLE:
            logger.debug("⚠ Keyring недоступен - MOSEDO username не найден")
            return None
            
        try:
            username = keyring.get_password(  # type: ignore
                CredentialManager.MOSEDO_SERVICE,
                CredentialManager.MOSEDO_LOGIN_KEY
            )
            if username:
                logger.info(f"✓ MOSEDO username retrieved from keyring: {username[:5]}***")
            return username
        except Exception as e:
            logger.warning(f"⚠ Keyring backend unavailable for MOSEDO username retrieval: {str(e)[:100]}")
            return None
    
    @staticmethod
    def save_yandexgpt_credentials(api_key: str, folder_id: str) -> bool:
        """
        Save YandexGPT API credentials securely using Windows Credential Manager.
        
        Requires Windows Credential Manager access. If unavailable, returns False.
        Users should run the application with administrator privileges.
        
        Args:
            api_key: YandexGPT API Key
            folder_id: Yandex Cloud Folder ID
            
        Returns:
            True if saved successfully, False otherwise
        """
        if not KEYRING_AVAILABLE:
            logger.error("✗ Keyring library not available")
            logger.info("  Install keyring: pip install keyring")
            return False
        
        try:
            # Save API Key
            keyring.set_password(  # type: ignore
                CredentialManager.YANDEXGPT_SERVICE,
                CredentialManager.YANDEXGPT_API_KEY,
                api_key
            )
            # Save Folder ID
            keyring.set_password(  # type: ignore
                CredentialManager.YANDEXGPT_SERVICE,
                CredentialManager.YANDEXGPT_FOLDER_ID,
                folder_id
            )
            logger.info(f"✓ YandexGPT credentials saved to Windows Credential Manager")
            return True
        except Exception as e:
            logger.error(f"✗ Windows Credential Manager access denied: {e}")
            logger.error("  Run the application as Administrator to save credentials")
            return False
    
    @staticmethod
    def get_yandexgpt_credentials() -> Optional[dict]:
        """
        Retrieve YandexGPT API credentials from Windows Credential Manager.
        
        Returns:
            Dict with 'api_key' and 'folder_id' if found, None otherwise
        """
        if not KEYRING_AVAILABLE:
            logger.debug("⚠ Keyring not available - credentials not found")
            return None
        
        try:
            api_key = keyring.get_password(  # type: ignore
                CredentialManager.YANDEXGPT_SERVICE,
                CredentialManager.YANDEXGPT_API_KEY
            )
            folder_id = keyring.get_password(  # type: ignore
                CredentialManager.YANDEXGPT_SERVICE,
                CredentialManager.YANDEXGPT_FOLDER_ID
            )
            
            if api_key and folder_id:
                logger.info(f"✓ YandexGPT credentials retrieved from Windows Credential Manager")
                return {'api_key': api_key, 'folder_id': folder_id}
            else:
                logger.debug(f"⚠ YandexGPT credentials not found in Credential Manager")
                return None
        except Exception as e:
            logger.warning(f"⚠ Failed to retrieve credentials: {str(e)[:100]}")
            return None
    
    @staticmethod
    def delete_yandexgpt_credentials() -> bool:
        """
        Delete YandexGPT API credentials from secure storage.
        
        Deletes from both Windows Credential Manager and local storage.
        
        Returns:
            True if deleted successfully, False otherwise
        """
        success = True
        
        # Try deleting from keyring
        if KEYRING_AVAILABLE:
            try:
                # Delete API Key
                keyring.delete_password(  # type: ignore
                    CredentialManager.YANDEXGPT_SERVICE,
                    CredentialManager.YANDEXGPT_API_KEY
                )
                logger.info(f"✓ YandexGPT API Key deleted from keyring")
            except Exception as e:
                if "PasswordDeleteError" not in type(e).__name__ and "not found" not in str(e).lower():
                    logger.warning(f"⚠ Keyring delete failed: {e}")
            
            try:
                # Delete Folder ID
                keyring.delete_password(  # type: ignore
                    CredentialManager.YANDEXGPT_SERVICE,
                    CredentialManager.YANDEXGPT_FOLDER_ID
                )
                logger.info(f"✓ YandexGPT Folder ID deleted from keyring")
            except Exception as e:
                if "PasswordDeleteError" not in type(e).__name__ and "not found" not in str(e).lower():
                    logger.warning(f"⚠ Keyring delete failed: {e}")
        
        # Also delete from local storage
        try:
            LocalCredentialStorage.delete(
                CredentialManager.YANDEXGPT_SERVICE,
                CredentialManager.YANDEXGPT_API_KEY
            )
            LocalCredentialStorage.delete(
                CredentialManager.YANDEXGPT_SERVICE,
                CredentialManager.YANDEXGPT_FOLDER_ID
            )
            logger.info(f"✓ YandexGPT credentials deleted from local storage")
        except Exception as e:
            logger.warning(f"⚠ Local storage delete failed: {e}")
        
        return success
    
    @staticmethod
    def has_yandexgpt_credentials() -> bool:
        """
        Check if YandexGPT credentials exist.
        
        Returns:
            True if credentials exist, False otherwise
        """
        credentials = CredentialManager.get_yandexgpt_credentials()
        return credentials is not None
