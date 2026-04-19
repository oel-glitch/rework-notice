"""
Scanner integration module for Canon ImageFormula DR-M260.

This module provides interface for scanning documents, OCR text recognition,
and integration with the PDF processing pipeline.
"""

import logging
from typing import Dict, List, Optional
from pathlib import Path
from datetime import datetime

logger = logging.getLogger(__name__)


class ScannerService:
    """
    Service for scanner operations and document digitization.
    
    Handles scanning, OCR, and document management for Canon DR-M260 scanner.
    """
    
    def __init__(self, scanner_model: str = "Canon DR-M260"):
        """
        Initialize scanner service.
        
        Args:
            scanner_model: Scanner model identifier
        """
        self.scanner_model = scanner_model
        self.scan_queue = []
        self.demo_mode = True
        logger.info(f"✓ Scanner service initialized for {scanner_model} (ДЕМО РЕЖИМ)")
        logger.info("📋 Для продакшена требуется: установка WIA/TWAIN драйверов и Tesseract OCR")
    
    def start_scan(self, settings: Optional[Dict] = None) -> Dict:
        """
        Start scanning operation with specified settings.
        
        Args:
            settings: Scan settings (resolution, color mode, duplex)
        
        Returns:
            Scan operation result
        """
        if settings is None:
            settings = {
                'resolution': 300,
                'color_mode': 'Color',
                'duplex': True
            }
        
        logger.info(f"Starting scan with settings: {settings}")
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        scan_result = {
            'status': 'demo',
            'file_path': f'scans/scan_{timestamp}.pdf',
            'pages_scanned': 5,
            'timestamp': timestamp,
            'settings': settings,
            'message': 'ДЕМО: Тестовый результат сканирования. Для реальной работы установите драйверы Canon DR-M260'
        }
        
        self.scan_queue.append(scan_result)
        return scan_result
    
    def process_batch(self, batch_name: str) -> List[Dict]:
        """
        Process batch of scanned documents.
        
        Args:
            batch_name: Name for the batch
        
        Returns:
            List of processed documents
        """
        logger.info(f"Processing batch: {batch_name}")
        
        return [
            {
                'id': 1,
                'filename': 'scan_001.pdf',
                'pages': 3,
                'ocr_status': 'completed',
                'document_type': 'Обращение гражданина'
            },
            {
                'id': 2,
                'filename': 'scan_002.pdf',
                'pages': 2,
                'ocr_status': 'completed',
                'document_type': 'Ответ юриста'
            }
        ]
    
    def ocr_document(self, file_path: str, language: str = 'rus') -> Dict:
        """
        Perform OCR on scanned document.
        
        Args:
            file_path: Path to scanned document
            language: OCR language (default: Russian)
        
        Returns:
            OCR result with extracted text
        """
        logger.info(f"Performing OCR on {file_path} (language: {language})")
        
        demo_text = """ДЕМО ТЕКСТ:
        Уважаемая администрация ОАТИ!
        Прошу принять меры по факту незаконной парковки...
        
        ДЛЯ РЕАЛЬНОЙ РАБОТЫ: Установите Tesseract OCR и pytesseract"""
        
        return {
            'file_path': file_path,
            'text': demo_text,
            'confidence': 0.95,
            'language': language,
            'status': 'demo_mode'
        }
    
    def get_scan_queue(self) -> List[Dict]:
        """
        Get current scan queue.
        
        Returns:
            List of scanned documents waiting for processing
        """
        return self.scan_queue
    
    def clear_queue(self) -> None:
        """Clear the scan queue."""
        self.scan_queue = []
        logger.info("Scan queue cleared")
    
    def is_scanner_connected(self) -> bool:
        """
        Check if scanner is connected and ready.
        
        Returns:
            True if scanner is available
        """
        if self.demo_mode:
            logger.info("ДЕМО: Сканер не подключен (тестовый режим)")
            return False
        return False
    
    def get_scanner_status(self) -> Dict:
        """
        Get current scanner status.
        
        Returns:
            Scanner status information
        """
        return {
            'model': self.scanner_model,
            'connected': False,
            'ready': False,
            'demo_mode': True,
            'message': '🔧 ДЕМО РЕЖИМ: Для подключения реального сканера Canon DR-M260 требуется:\n' +
                      '1. Установить драйверы Canon (WIA/TWAIN)\n' +
                      '2. Установить Tesseract OCR для распознавания текста\n' +
                      '3. Подключить сканер через USB и настроить в системе'
        }
    
    def check_system_requirements(self) -> Dict:
        """
        Check system requirements for scanner operation.
        
        Returns:
            Dictionary with check results
        """
        import shutil
        import sys
        
        checks = {
            'tesseract_installed': False,
            'twain_available': False,
            'wia_available': False,
            'pillow_available': False,
            'pytesseract_available': False,
            'overall_status': 'not_ready'
        }
        
        # Check Tesseract OCR
        tesseract_path = shutil.which('tesseract')
        checks['tesseract_installed'] = tesseract_path is not None
        if tesseract_path:
            logger.info(f"✓ Tesseract найден: {tesseract_path}")
        else:
            logger.warning("✗ Tesseract OCR не установлен")
        
        # Check PIL/Pillow
        try:
            import PIL
            checks['pillow_available'] = True
            logger.info("✓ Pillow установлен")
        except ImportError:
            logger.warning("✗ Pillow не установлен")
        
        # Check pytesseract
        try:
            import pytesseract  # type: ignore
            checks['pytesseract_available'] = True
            logger.info("✓ pytesseract установлен")
        except ImportError:
            logger.warning("✗ pytesseract не установлен")
        
        # Check TWAIN (Windows only)
        if sys.platform == 'win32':
            try:
                import twain
                checks['twain_available'] = True
                logger.info("✓ TWAIN драйвер доступен")
            except ImportError:
                logger.warning("✗ TWAIN драйвер не установлен")
            
            # Check WIA
            try:
                import win32com.client
                checks['wia_available'] = True
                logger.info("✓ WIA доступен (pywin32)")
            except ImportError:
                logger.warning("✗ WIA не доступен (pywin32 не установлен)")
        
        # Determine overall status
        if checks['tesseract_installed'] and checks['pillow_available']:
            if sys.platform == 'win32' and (checks['twain_available'] or checks['wia_available']):
                checks['overall_status'] = 'ready'
            elif sys.platform != 'win32':
                checks['overall_status'] = 'partial'  # Linux/Mac needs different approach
        
        return checks
