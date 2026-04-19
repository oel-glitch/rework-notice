"""
MOSEDO Automation module for workflow recording and playback.

This module provides browser automation for Moscow Electronic Document Management System (MOSEDO).
It records user actions (clicks, text input, navigation) and replays them for batch processing.
Falls back to offline mode if Selenium is not available.
"""

from __future__ import annotations
import logging
import json
import time
from typing import Dict, List, Optional, Tuple, TYPE_CHECKING
from datetime import datetime

if TYPE_CHECKING:
    from selenium import webdriver
    from selenium.webdriver.remote.webelement import WebElement

try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.remote.webelement import WebElement
    from selenium.common.exceptions import (
        TimeoutException,
        NoSuchElementException,
        WebDriverException
    )
    SELENIUM_AVAILABLE = True
    logger_init = logging.getLogger(__name__)
    logger_init.debug("✓ Selenium library loaded successfully")
except ImportError:
    SELENIUM_AVAILABLE = False
    webdriver = None  # type: ignore
    WebElement = None  # type: ignore
    By = None  # type: ignore
    WebDriverWait = None  # type: ignore
    EC = None  # type: ignore
    TimeoutException = Exception  # type: ignore
    NoSuchElementException = Exception  # type: ignore
    WebDriverException = Exception  # type: ignore
    logger_init = logging.getLogger(__name__)
    logger_init.warning("⚠ Selenium library not available - MOSEDO automation features disabled")
    logger_init.info("  Установите selenium для использования MOSEDO функций: pip install selenium")

logger = logging.getLogger(__name__)


class WorkflowStep:
    """
    Represents a single step in a workflow.
    
    Attributes:
        action: Type of action (click, type, navigate, wait)
        selector: CSS selector for the target element
        value: Optional value (for text input actions)
        description: Human-readable description
        timestamp: When the step was recorded
    """
    
    def __init__(
        self, 
        action: str, 
        selector: Optional[str] = None,
        value: Optional[str] = None,
        description: Optional[str] = None
    ):
        self.action = action
        self.selector = selector
        self.value = value
        self.description = description or ""
        self.timestamp = datetime.now()
    
    def to_dict(self) -> Dict:
        """Convert step to dictionary for JSON serialization."""
        return {
            'action': self.action,
            'selector': self.selector,
            'value': self.value,
            'description': self.description,
            'timestamp': self.timestamp.isoformat()
        }
    
    @classmethod
    def from_dict(cls, data: Dict) -> 'WorkflowStep':
        """Create WorkflowStep from dictionary."""
        step = cls(
            action=data['action'],
            selector=data.get('selector'),
            value=data.get('value'),
            description=data.get('description', '')
        )
        if 'timestamp' in data:
            step.timestamp = datetime.fromisoformat(data['timestamp'])
        return step
    
    def __repr__(self) -> str:
        if self.action == 'type':
            return f"<Step: {self.action} '{self.value}' to {self.selector}>"
        elif self.action == 'click':
            return f"<Step: {self.action} on {self.selector}>"
        elif self.action == 'navigate':
            return f"<Step: {self.action} to {self.value}>"
        else:
            return f"<Step: {self.action}>"


class WorkflowRecorder:
    """
    Records user actions in the browser for workflow creation.
    
    This class uses Selenium WebDriver with injected JavaScript to automatically
    capture user interactions (clicks, input, keyboard) with the web interface.
    """
    
    def __init__(self):
        self.driver: Optional[webdriver.Chrome] = None
        self.is_recording = False
        self.current_workflow: List[WorkflowStep] = []
        self.workflow_name = ""
        self.polling_active = False
        self.last_event_count = 0
        logger.info("WorkflowRecorder initialized")
    
    def _load_recorder_script(self) -> str:
        """Load the browser recorder JavaScript code."""
        import os
        script_path = os.path.join(os.path.dirname(__file__), 'browser_recorder.js')
        try:
            with open(script_path, 'r', encoding='utf-8') as f:
                return f.read()
        except FileNotFoundError:
            logger.error(f"✗ Recorder script not found at: {script_path}")
            return ""
    
    def _inject_recorder_script(self):
        """Inject the browser recorder JavaScript into the current page."""
        if not self.driver:
            return False
        
        try:
            script = self._load_recorder_script()
            if not script:
                return False
            
            # Inject the script
            self.driver.execute_script(script)
            logger.info("✓ Browser recorder script injected")
            
            # Wait a bit for initialization
            time.sleep(0.5)
            
            # Verify injection
            is_recording = self.driver.execute_script("return window.isRecording === true;")
            if is_recording:
                logger.info("✓ Recorder is active in browser")
                return True
            else:
                logger.error("✗ Recorder script injected but not active")
                return False
                
        except Exception as e:
            logger.error(f"✗ Failed to inject recorder script: {e}")
            return False
    
    def _poll_events(self):
        """Poll events from the browser and convert to WorkflowSteps."""
        if not self.driver or not self.is_recording:
            return
        
        try:
            # Get recorded actions from browser
            actions = self.driver.execute_script("return window.recordedActions || [];")
            
            # Process only new events
            new_events = actions[self.last_event_count:]
            
            for event in new_events:
                self._process_event(event)
            
            self.last_event_count = len(actions)
            
            # Check if recording stopped (via hotkey)
            is_recording = self.driver.execute_script("return window.isRecording === true;")
            if not is_recording and self.is_recording:
                logger.info("✓ Recording stopped via browser hotkey (Ctrl+Shift+S)")
                self.is_recording = False
                self.polling_active = False
                
        except Exception as e:
            logger.debug(f"Error polling events: {e}")
    
    def _process_event(self, event: Dict):
        """Convert a browser event to WorkflowStep."""
        try:
            event_type = event.get('type', '')
            
            if event_type == 'click':
                step = WorkflowStep(
                    action='click',
                    selector=event.get('selector', ''),
                    description=f"Click on {event.get('tagName', 'element')}: {event.get('text', '')[:50]}"
                )
                self.current_workflow.append(step)
                logger.info(f"  Step {len(self.current_workflow)}: Click {event.get('selector', '')}")
                
            elif event_type == 'input':
                step = WorkflowStep(
                    action='type',
                    selector=event.get('selector', ''),
                    value=event.get('value', ''),
                    description=f"Type into {event.get('tagName', 'field')}"
                )
                self.current_workflow.append(step)
                logger.info(f"  Step {len(self.current_workflow)}: Type '{event.get('value', '')[:20]}...' into {event.get('selector', '')}")
                
            elif event_type == 'keypress':
                step = WorkflowStep(
                    action='keypress',
                    selector=event.get('selector', ''),
                    value=event.get('key', ''),
                    description=f"Press {event.get('key', '')} key"
                )
                self.current_workflow.append(step)
                logger.info(f"  Step {len(self.current_workflow)}: Press {event.get('key', '')}")
                
            elif event_type == 'navigate':
                step = WorkflowStep(
                    action='navigate',
                    value=event.get('url', ''),
                    description=f"Navigate to {event.get('url', '')}"
                )
                self.current_workflow.append(step)
                logger.info(f"  Step {len(self.current_workflow)}: Navigate to {event.get('url', '')}")
                
        except Exception as e:
            logger.error(f"Error processing event: {e}")
    
    def start_recording(self, workflow_name: str, start_url: str) -> bool:
        """
        Start recording a new workflow with automatic action capture.
        
        Args:
            workflow_name: Name for the workflow
            start_url: URL to navigate to
            
        Returns:
            True if recording started successfully
        """
        if not SELENIUM_AVAILABLE:
            logger.error("✗ Selenium not available - cannot start recording")
            return False
        
        try:
            logger.info(f"Starting workflow recording: {workflow_name}")
            logger.info(f"  Start URL: {start_url}")
            
            # Initialize Chrome driver
            options = webdriver.ChromeOptions()
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            # Don't use headless mode for recording (user needs to see and interact)
            
            self.driver = webdriver.Chrome(options=options)
            self.driver.maximize_window()
            
            # Navigate to start URL
            logger.info("  Navigating to start URL...")
            self.driver.get(start_url)
            
            # Wait for page load
            time.sleep(1)
            
            # Inject recorder script
            logger.info("  Injecting browser recorder script...")
            if not self._inject_recorder_script():
                raise Exception("Failed to inject recorder script")
            
            self.is_recording = True
            self.workflow_name = workflow_name
            self.current_workflow = []
            self.last_event_count = 0
            self.polling_active = True
            
            # Start polling events in background
            import threading
            def poll_loop():
                while self.polling_active and self.is_recording:
                    self._poll_events()
                    time.sleep(0.5)  # Poll every 500ms
            
            threading.Thread(target=poll_loop, daemon=True).start()
            
            logger.info(f"✓ Recording started: {workflow_name}")
            logger.info("  🔴 Browser is recording your actions automatically")
            logger.info("  Press Ctrl+Shift+S in browser to stop recording")
            logger.info("  All clicks, inputs, and navigation will be captured")
            
            return True
            
        except Exception as e:
            logger.error(f"✗ Failed to start recording: {e}")
            self.is_recording = False
            self.polling_active = False
            if self.driver:
                self.driver.quit()
                self.driver = None
            return False
    
    def stop_recording(self) -> Tuple[bool, List[WorkflowStep]]:
        """
        Stop recording and return captured workflow.
        
        Returns:
            Tuple of (success, workflow_steps)
        """
        if not self.is_recording:
            logger.warning("No recording in progress")
            return False, []
        
        try:
            logger.info(f"Stopping workflow recording: {self.workflow_name}")
            
            # Stop polling
            self.polling_active = False
            
            # Get final events from browser
            if self.driver:
                try:
                    time.sleep(0.5)  # Give time for last events
                    self._poll_events()  # Final poll
                except:
                    pass
            
            logger.info(f"  Recorded steps: {len(self.current_workflow)}")
            
            steps = self.current_workflow.copy()
            
            # Cleanup
            if self.driver:
                try:
                    self.driver.quit()
                except:
                    pass
                self.driver = None
            
            self.is_recording = False
            self.current_workflow = []
            self.last_event_count = 0
            
            logger.info(f"✓ Recording stopped: {self.workflow_name}")
            logger.info(f"  Total steps recorded: {len(steps)}")
            
            return True, steps
            
        except Exception as e:
            logger.error(f"✗ Error stopping recording: {e}")
            return False, []
    
    def get_current_steps(self) -> List[WorkflowStep]:
        """
        Get currently recorded steps (for real-time UI updates).
        
        Returns:
            List of WorkflowStep objects
        """
        return self.current_workflow.copy()
    
    def add_step(
        self, 
        action: str, 
        selector: Optional[str] = None,
        value: Optional[str] = None,
        description: str = ""
    ):
        """
        Manually add a step to the current workflow.
        
        Args:
            action: Type of action
            selector: CSS selector
            value: Optional value
            description: Description
        """
        if not self.is_recording:
            logger.warning("Cannot add step: no recording in progress")
            return
        
        step = WorkflowStep(action, selector, value, description)
        self.current_workflow.append(step)
        logger.debug(f"  Step {len(self.current_workflow)}: {step}")
    
    def get_element_selector(self, element: WebElement) -> str:
        """
        Generate CSS selector for a web element.
        
        Priority: id > name > class > tag[attribute]
        
        Args:
            element: Selenium WebElement
            
        Returns:
            CSS selector string
        """
        # Try ID first
        elem_id = element.get_attribute('id')
        if elem_id:
            return f"#{elem_id}"
        
        # Try name
        elem_name = element.get_attribute('name')
        if elem_name:
            return f"[name='{elem_name}']"
        
        # Try class
        elem_class = element.get_attribute('class')
        if elem_class:
            classes = elem_class.strip().split()
            if classes:
                return f".{classes[0]}"
        
        # Fallback: tag name
        tag = element.tag_name
        return tag


class WorkflowPlayer:
    """
    Replays recorded workflows with error handling and wait logic.
    """
    
    def __init__(self):
        self.driver: Optional[webdriver.Chrome] = None
        self.default_timeout = 10  # seconds
        logger.info("WorkflowPlayer initialized")
    
    def play_workflow(
        self, 
        steps: List[WorkflowStep],
        headless: bool = False
    ) -> Tuple[bool, str]:
        """
        Execute a workflow from recorded steps.
        
        Args:
            steps: List of workflow steps
            headless: Run browser in headless mode
            
        Returns:
            Tuple of (success, error_message)
        """
        if not SELENIUM_AVAILABLE:
            logger.error("✗ Selenium not available - cannot play workflow")
            return False, "Selenium library not installed"
        
        try:
            logger.info(f"Starting workflow playback: {len(steps)} steps")
            
            # Initialize Chrome driver
            options = webdriver.ChromeOptions()
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            if headless:
                options.add_argument('--headless')
            
            self.driver = webdriver.Chrome(options=options)
            self.driver.maximize_window()
            
            # Execute each step
            for i, step in enumerate(steps, 1):
                logger.info(f"  Step {i}/{len(steps)}: {step.action}")
                
                success, error = self._execute_step(step)
                if not success:
                    logger.error(f"✗ Step {i} failed: {error}")
                    return False, f"Step {i} failed: {error}"
                
                # Small delay between steps for stability
                time.sleep(0.5)
            
            logger.info(f"✓ Workflow completed successfully: {len(steps)} steps")
            return True, ""
            
        except Exception as e:
            error_msg = f"Workflow execution error: {e}"
            logger.error(f"✗ {error_msg}")
            return False, error_msg
        
        finally:
            if self.driver:
                self.driver.quit()
                self.driver = None
    
    def _execute_step(self, step: WorkflowStep) -> Tuple[bool, str]:
        """
        Execute a single workflow step.
        
        Returns:
            Tuple of (success, error_message)
        """
        try:
            if step.action == 'navigate':
                self.driver.get(step.value)
                logger.debug(f"    Navigated to {step.value}")
                return True, ""
            
            elif step.action == 'click':
                element = self._wait_for_element(step.selector)
                if not element:
                    return False, f"Element not found: {step.selector}"
                element.click()
                logger.debug(f"    Clicked: {step.selector}")
                return True, ""
            
            elif step.action == 'type':
                element = self._wait_for_element(step.selector)
                if not element:
                    return False, f"Element not found: {step.selector}"
                element.clear()
                element.send_keys(step.value)
                logger.debug(f"    Typed '{step.value}' into {step.selector}")
                return True, ""
            
            elif step.action == 'wait':
                wait_time = float(step.value) if step.value else 1.0
                time.sleep(wait_time)
                logger.debug(f"    Waited {wait_time}s")
                return True, ""
            
            else:
                return False, f"Unknown action: {step.action}"
                
        except Exception as e:
            return False, str(e)
    
    def _wait_for_element(
        self, 
        selector: str, 
        timeout: Optional[int] = None
    ) -> Optional[WebElement]:
        """
        Wait for element to be present and clickable.
        
        Args:
            selector: CSS selector
            timeout: Wait timeout in seconds
            
        Returns:
            WebElement if found, None otherwise
        """
        if timeout is None:
            timeout = self.default_timeout
        
        try:
            wait = WebDriverWait(self.driver, timeout)  # type: ignore
            element = wait.until(  # type: ignore
                EC.element_to_be_clickable((By.CSS_SELECTOR, selector))  # type: ignore
            )
            return element
        except Exception as e:
            # Check if it's a timeout (works whether selenium is available or not)
            if 'Timeout' in type(e).__name__ or 'timeout' in str(e).lower():
                logger.warning(f"Timeout waiting for element: {selector}")
                return None
            else:
                logger.error(f"Error finding element {selector}: {e}")
                return None


class MOSEDOAutomation:
    """
    Main class for MOSEDO workflow automation.
    
    Manages workflow recording, storage, and playback.
    """
    
    def __init__(self, database=None):
        """
        Initialize MOSEDO automation.
        
        Args:
            database: Database instance for workflow storage
        """
        self.database = database
        self.recorder = WorkflowRecorder() if SELENIUM_AVAILABLE else None
        self.player = WorkflowPlayer() if SELENIUM_AVAILABLE else None
        self.workflows: Dict[int, Dict] = {}
        self.last_recorded_steps: Optional[List[WorkflowStep]] = None  # Backup for failed saves
        self.last_workflow_name: Optional[str] = None
        
        if SELENIUM_AVAILABLE:
            logger.info("✓ MOSEDO Automation initialized with Selenium")
        else:
            logger.warning("⚠ MOSEDO Automation initialized in DEMO mode (Selenium not available)")
            logger.warning("  Install selenium package for full functionality")
    
    def create_workflow(self, name: str, start_url: str) -> Tuple[bool, str]:
        """
        Start creating a new workflow by recording.
        
        Args:
            name: Workflow name
            start_url: Starting URL
            
        Returns:
            Tuple of (success, message)
        """
        if not SELENIUM_AVAILABLE:
            return False, "Selenium not available - install selenium package"
        
        if self.recorder.is_recording:
            return False, "Already recording a workflow"
        
        success = self.recorder.start_recording(name, start_url)
        if success:
            return True, f"Recording started: {name}"
        else:
            return False, "Failed to start recording"
    
    def save_workflow(self, workflow_id: Optional[int] = None, description: Optional[str] = None) -> Tuple[bool, str]:
        """
        Save the currently recorded workflow to database.
        
        Preserves recorded steps in memory if database save fails.
        
        Args:
            workflow_id: Optional ID for updating existing workflow
            description: Optional workflow description
            
        Returns:
            Tuple of (success, message)
        """
        if not SELENIUM_AVAILABLE or not self.recorder:
            return False, "Selenium not available"
        
        if not self.recorder.is_recording:
            # Check if we have cached steps from previous failed save
            if self.last_recorded_steps:
                logger.info(f"Using cached steps from previous recording: {self.last_workflow_name}")
                steps = self.last_recorded_steps
                workflow_name = self.last_workflow_name
            else:
                return False, "No recording in progress and no cached steps"
        else:
            # Stop recording and get steps
            success, steps = self.recorder.stop_recording()
            if not success or not steps:
                return False, "No steps recorded"
            
            workflow_name = self.recorder.workflow_name
            
            # Cache steps in case database save fails
            self.last_recorded_steps = steps
            self.last_workflow_name = workflow_name
        
        # Convert steps to JSON
        steps_json = json.dumps([step.to_dict() for step in steps], ensure_ascii=False)
        
        # Save to database
        if not self.database:
            logger.warning("Database not available - workflow cached in memory")
            return True, f"Workflow recorded with {len(steps)} steps (cached - no database). Call save_workflow() again when database available."
        
        try:
            if workflow_id:
                # Update existing workflow
                updated = self.database.update_workflow(
                    workflow_id=workflow_id,
                    steps_json=steps_json,
                    description=description
                )
                if updated:
                    logger.info(f"✓ Workflow {workflow_id} updated: {workflow_name} ({len(steps)} steps)")
                    # Clear cache on successful save
                    self.last_recorded_steps = None
                    self.last_workflow_name = None
                    return True, f"Workflow updated with {len(steps)} steps"
                else:
                    logger.warning(f"Workflow {workflow_id} update failed - steps cached for retry")
                    return False, f"Failed to update workflow {workflow_id}. Steps cached - check logs for details. Retry save_workflow()."
            else:
                # Create new workflow
                new_id = self.database.add_workflow(
                    name=workflow_name,
                    steps_json=steps_json,
                    description=description
                )
                if new_id:
                    logger.info(f"✓ Workflow created (ID: {new_id}): {workflow_name} ({len(steps)} steps)")
                    # Clear cache on successful save
                    self.last_recorded_steps = None
                    self.last_workflow_name = None
                    return True, f"Workflow saved with ID {new_id} ({len(steps)} steps)"
                else:
                    logger.warning(f"Workflow '{workflow_name}' save failed (likely name collision) - steps cached for retry")
                    return False, f"Failed to save workflow '{workflow_name}'. Possible reasons: name collision, database error. Steps cached - try different name or check logs."
        
        except Exception as e:
            logger.error(f"✗ Error saving workflow: {e}")
            return False, f"Error saving workflow: {e}. Steps cached - retry save_workflow()."
    
    def load_workflow(self, workflow_id: int) -> Optional[List[WorkflowStep]]:
        """
        Load a workflow from database with validation.
        
        Args:
            workflow_id: Workflow ID
            
        Returns:
            List of WorkflowSteps if found and valid, None otherwise
        """
        if not self.database:
            logger.error("Database not available - cannot load workflow")
            return None
        
        try:
            workflow_data = self.database.get_workflow(workflow_id)
            if not workflow_data:
                logger.warning(f"Workflow {workflow_id} not found in database")
                return None
            
            # Parse steps from JSON
            steps_data = json.loads(workflow_data['steps_json'])
            if not isinstance(steps_data, list):
                logger.error(f"✗ Workflow {workflow_id} has invalid steps format (not a list)")
                return None
            
            # Validate and instantiate steps
            steps = []
            for i, step_dict in enumerate(steps_data):
                # Validate required fields
                if not isinstance(step_dict, dict):
                    logger.warning(f"  Step {i+1} is not a dict, skipping")
                    continue
                
                if 'action' not in step_dict:
                    logger.warning(f"  Step {i+1} missing 'action' field, skipping")
                    continue
                
                # Validate action-specific requirements
                action = step_dict['action']
                if action in ['click', 'type'] and not step_dict.get('selector'):
                    logger.warning(f"  Step {i+1} ({action}) missing required 'selector', skipping")
                    continue
                
                if action == 'type' and not step_dict.get('value'):
                    logger.warning(f"  Step {i+1} (type) missing required 'value', skipping")
                    continue
                
                if action == 'navigate' and not step_dict.get('value'):
                    logger.warning(f"  Step {i+1} (navigate) missing required 'value' (URL), skipping")
                    continue
                
                # Create step
                try:
                    step = WorkflowStep.from_dict(step_dict)
                    steps.append(step)
                except Exception as e:
                    logger.warning(f"  Step {i+1} failed to instantiate: {e}, skipping")
                    continue
            
            if not steps:
                logger.error(f"✗ Workflow {workflow_id} has no valid steps after validation")
                return None
            
            logger.info(f"✓ Loaded workflow {workflow_id}: {workflow_data['name']} ({len(steps)} valid steps)")
            if len(steps) < len(steps_data):
                logger.warning(f"  ⚠ Skipped {len(steps_data) - len(steps)} invalid steps")
            
            return steps
            
        except json.JSONDecodeError as e:
            logger.error(f"✗ Invalid JSON in workflow {workflow_id}: {e}")
            return None
        except KeyError as e:
            logger.error(f"✗ Missing required field in workflow {workflow_id}: {e}")
            return None
        except Exception as e:
            logger.error(f"✗ Error loading workflow {workflow_id}: {e}")
            return None
    
    def execute_workflow(
        self, 
        workflow_id: int, 
        headless: bool = False
    ) -> Tuple[bool, str]:
        """
        Execute a saved workflow.
        
        Args:
            workflow_id: Workflow ID to execute
            headless: Run in headless mode
            
        Returns:
            Tuple of (success, message)
        """
        if not SELENIUM_AVAILABLE or not self.player:
            return False, "Selenium not available"
        
        steps = self.load_workflow(workflow_id)
        if not steps:
            return False, f"Workflow {workflow_id} not found"
        
        logger.info(f"Executing workflow {workflow_id}")
        return self.player.play_workflow(steps, headless)


# Legacy class for backwards compatibility (демо режим)
class MosedoRobot:
    """
    Automation robot for MOSEDO document registration workflows.
    
    Records user actions and replays them for batch processing of documents.
    """
    
    def __init__(self):
        """Initialize MOSEDO automation robot."""
        self.workflows = {}
        self.current_recording = None
        self.is_recording = False
        self.demo_mode = True
        logger.info("✓ MOSEDO Robot initialized (ДЕМО РЕЖИМ)")
        logger.info("📋 Для продакшена требуется: PyAutoGUI или Selenium WebDriver для автоматизации")
    
    def record_workflow(self, workflow_name: str) -> Dict:
        """
        Start recording a new workflow.
        
        Args:
            workflow_name: Name for the workflow
        
        Returns:
            Recording status
        """
        logger.info(f"Starting workflow recording: {workflow_name}")
        
        self.current_recording = {
            'name': workflow_name,
            'steps': [],
            'created_at': datetime.now().isoformat(),
            'status': 'recording'
        }
        self.is_recording = True
        
        return {
            'status': 'recording_started',
            'workflow_name': workflow_name,
            'message': 'ДЕМО: В реальной версии робот записывает ваши действия в МОСЭДО. Требуется PyAutoGUI.'
        }
    
    def stop_recording(self) -> Dict:
        """
        Stop current workflow recording and save it.
        
        Returns:
            Saved workflow data
        """
        if not self.is_recording or not self.current_recording:
            return {'status': 'error', 'message': 'No active recording'}
        
        logger.info(f"Stopping workflow recording: {self.current_recording['name']}")
        
        workflow_name = self.current_recording['name']
        self.workflows[workflow_name] = self.current_recording
        self.current_recording['status'] = 'saved'
        
        result = self.current_recording.copy()
        self.is_recording = False
        self.current_recording = None
        
        return {
            'status': 'recording_saved',
            'workflow': result,
            'message': f'Workflow "{workflow_name}" saved with {len(result["steps"])} steps'
        }
    
    def replay_workflow(self, workflow_name: str, document_ids: List[str]) -> Dict:
        """
        Replay a saved workflow for multiple documents.
        
        Args:
            workflow_name: Name of workflow to replay
            document_ids: List of MOSEDO document IDs to process
        
        Returns:
            Batch processing result
        """
        if workflow_name not in self.workflows:
            return {'status': 'error', 'message': f'Workflow "{workflow_name}" not found'}
        
        logger.info(f"Replaying workflow '{workflow_name}' for {len(document_ids)} documents")
        
        workflow = self.workflows[workflow_name]
        
        results = []
        for i, doc_id in enumerate(document_ids, 1):
            results.append({
                'document_id': doc_id,
                'status': 'demo',
                'progress': f'{i}/{len(document_ids)}',
                'message': 'ДЕМО: В реальной версии документ регистрируется автоматически'
            })
        
        return {
            'status': 'batch_started',
            'workflow_name': workflow_name,
            'total_documents': len(document_ids),
            'results': results,
            'message': '🤖 ДЕМО РЕЖИМ: Для автоматизации МОСЭДО требуется:\n' +
                      '1. PyAutoGUI для управления мышью и клавиатурой\n' + 
                      '2. Настройка координат элементов интерфейса МОСЭДО\n' +
                      '3. Учетные данные для входа в систему'
        }
    
    def process_document_batch(self, workflow_name: str, documents: List[Dict]) -> List[Dict]:
        """
        Process batch of documents using saved workflow.
        
        Args:
            workflow_name: Workflow to use
            documents: List of documents with their data
        
        Returns:
            Processing results for each document
        """
        logger.info(f"Processing {len(documents)} documents with workflow '{workflow_name}'")
        
        results = []
        for doc in documents:
            results.append({
                'document': doc,
                'status': 'completed',
                'registered_number': f"01-{doc.get('id', '000')}-25",
                'timestamp': datetime.now().strftime('%d.%m.%Y %H:%M')
            })
        
        return results
    
    def get_workflows(self) -> List[Dict]:
        """
        Get list of saved workflows.
        
        Returns:
            List of all workflows
        """
        return [
            {
                'name': name,
                'created_at': wf['created_at'],
                'steps_count': len(wf['steps']),
                'status': wf['status']
            }
            for name, wf in self.workflows.items()
        ]
    
    def create_sample_workflow(self) -> Dict:
        """
        Create a sample workflow for demonstration.
        
        Returns:
            Sample workflow data
        """
        sample = {
            'name': 'Регистрация исходящих уведомлений',
            'steps': [
                {'action': 'click', 'target': 'Кнопка "Зарегистрировать"', 'delay': 0.5},
                {'action': 'select', 'target': 'Номенклатура', 'value': '01-21-П', 'delay': 0.3},
                {'action': 'input', 'target': 'Номер документа', 'value': '{номер}', 'delay': 0.2},
                {'action': 'click', 'target': 'Enter', 'delay': 0.5},
                {'action': 'click', 'target': 'Отправить по email', 'delay': 0.5},
                {'action': 'select', 'target': 'Email получателя', 'value': '{email}', 'delay': 0.3},
                {'action': 'input', 'target': 'Тема письма', 'value': 'Уведомление', 'delay': 0.2},
                {'action': 'click', 'target': 'Отправить', 'delay': 0.5},
                {'action': 'navigate', 'target': 'Связанный входящий', 'delay': 0.5},
                {'action': 'click', 'target': 'Резолюция +Фисковой Е.А.', 'delay': 0.3},
                {'action': 'click', 'target': 'Исполнение', 'delay': 0.5},
                {'action': 'input', 'target': 'Номер исходящего', 'value': '{номер_исх}', 'delay': 0.2},
                {'action': 'click', 'target': 'Сохранить', 'delay': 0.5}
            ],
            'created_at': datetime.now().isoformat(),
            'status': 'ready'
        }
        
        self.workflows[sample['name']] = sample
        logger.info("Sample workflow created")
        
        return sample
