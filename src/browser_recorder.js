/**
 * Browser Action Recorder
 * Automatically captures user interactions (clicks, input, keyboard) in the browser
 * and stores them for workflow automation.
 */

(function() {
    'use strict';
    
    // Global storage for recorded actions
    window.recordedActions = window.recordedActions || [];
    window.isRecording = true;
    
    /**
     * Generate optimal CSS/XPath selector for an element
     * Priority: id > name > data-* > unique class > xpath
     */
    function getOptimalSelector(element) {
        // Priority 1: ID (most reliable)
        if (element.id) {
            return `#${element.id}`;
        }
        
        // Priority 2: Name attribute (for form elements)
        if (element.name) {
            return `[name="${element.name}"]`;
        }
        
        // Priority 3: data-* attributes
        for (let attr of element.attributes) {
            if (attr.name.startsWith('data-') && attr.value) {
                return `[${attr.name}="${attr.value}"]`;
            }
        }
        
        // Priority 4: Unique class combination
        if (element.className && typeof element.className === 'string') {
            const classes = element.className.trim().split(/\s+/).filter(c => c);
            if (classes.length > 0) {
                const classSelector = '.' + classes.join('.');
                // Check if unique
                if (document.querySelectorAll(classSelector).length === 1) {
                    return classSelector;
                }
            }
        }
        
        // Priority 5: Tag + unique text content (for buttons, links)
        const tagName = element.tagName.toLowerCase();
        if (['a', 'button', 'span'].includes(tagName)) {
            const text = element.textContent.trim().substring(0, 30);
            if (text) {
                const selector = `${tagName}:contains("${text}")`;
                // Note: :contains is jQuery, fallback to xpath
                return getXPath(element);
            }
        }
        
        // Priority 6: XPath as fallback
        return getXPath(element);
    }
    
    /**
     * Generate XPath for element
     */
    function getXPath(element) {
        if (element.id) {
            return `//*[@id="${element.id}"]`;
        }
        
        const parts = [];
        while (element && element.nodeType === Node.ELEMENT_NODE) {
            let index = 0;
            let sibling = element.previousSibling;
            
            while (sibling) {
                if (sibling.nodeType === Node.ELEMENT_NODE && 
                    sibling.nodeName === element.nodeName) {
                    index++;
                }
                sibling = sibling.previousSibling;
            }
            
            const tagName = element.nodeName.toLowerCase();
            const pathIndex = index > 0 ? `[${index + 1}]` : '';
            parts.unshift(tagName + pathIndex);
            
            element = element.parentNode;
        }
        
        return parts.length ? '/' + parts.join('/') : '';
    }
    
    /**
     * Check if element should be ignored
     */
    function shouldIgnoreElement(element) {
        const ignoredTags = ['html', 'body', 'head', 'script', 'style'];
        const tagName = element.tagName.toLowerCase();
        
        // Ignore document-level elements
        if (ignoredTags.includes(tagName)) {
            return true;
        }
        
        // Ignore our own recording indicator
        if (element.id === 'workflow-recorder-indicator') {
            return true;
        }
        
        return false;
    }
    
    /**
     * Highlight element when clicked (visual feedback)
     */
    function highlightElement(element) {
        const originalOutline = element.style.outline;
        element.style.outline = '3px solid #00ff00';
        element.style.outlineOffset = '2px';
        
        setTimeout(() => {
            element.style.outline = originalOutline;
        }, 500);
    }
    
    /**
     * Click event listener
     */
    function handleClick(event) {
        if (!window.isRecording) return;
        
        const element = event.target;
        if (shouldIgnoreElement(element)) return;
        
        const selector = getOptimalSelector(element);
        const action = {
            type: 'click',
            selector: selector,
            tagName: element.tagName.toLowerCase(),
            text: element.textContent.trim().substring(0, 50),
            timestamp: Date.now(),
            url: window.location.href
        };
        
        window.recordedActions.push(action);
        highlightElement(element);
        
        console.log('[RECORDER] Click:', action);
    }
    
    /**
     * Input event listener (text input)
     */
    function handleInput(event) {
        if (!window.isRecording) return;
        
        const element = event.target;
        if (shouldIgnoreElement(element)) return;
        
        // Only record input for form elements
        if (!['input', 'textarea', 'select'].includes(element.tagName.toLowerCase())) {
            return;
        }
        
        const selector = getOptimalSelector(element);
        const action = {
            type: 'input',
            selector: selector,
            value: element.value,
            tagName: element.tagName.toLowerCase(),
            timestamp: Date.now(),
            url: window.location.href
        };
        
        // Debounce: remove previous input event for same selector
        window.recordedActions = window.recordedActions.filter(a => 
            !(a.type === 'input' && a.selector === selector)
        );
        
        window.recordedActions.push(action);
        
        console.log('[RECORDER] Input:', action);
    }
    
    /**
     * Keypress event listener (Enter key for form submission)
     */
    function handleKeyPress(event) {
        if (!window.isRecording) return;
        
        // Record Enter key in input fields (form submission)
        if (event.key === 'Enter' && event.target.tagName.toLowerCase() === 'input') {
            const element = event.target;
            const selector = getOptimalSelector(element);
            
            const action = {
                type: 'keypress',
                selector: selector,
                key: 'Enter',
                timestamp: Date.now(),
                url: window.location.href
            };
            
            window.recordedActions.push(action);
            console.log('[RECORDER] Keypress Enter:', action);
        }
        
        // Check for stop recording hotkey: Ctrl+Shift+S
        if (event.ctrlKey && event.shiftKey && event.key === 'S') {
            event.preventDefault();
            stopRecording();
        }
    }
    
    /**
     * Navigation event listener
     */
    function handleNavigation() {
        if (!window.isRecording) return;
        
        const action = {
            type: 'navigate',
            url: window.location.href,
            timestamp: Date.now()
        };
        
        window.recordedActions.push(action);
        console.log('[RECORDER] Navigation:', action);
    }
    
    /**
     * Stop recording
     */
    function stopRecording() {
        window.isRecording = false;
        console.log('[RECORDER] Recording stopped by hotkey');
        
        // Update indicator
        const indicator = document.getElementById('workflow-recorder-indicator');
        if (indicator) {
            indicator.textContent = '⏹ Recording Stopped (Ctrl+Shift+S)';
            indicator.style.backgroundColor = '#e74c3c';
        }
        
        // Alert user
        alert('Recording stopped! Close the browser window to save the workflow.');
    }
    
    /**
     * Create recording indicator
     */
    function createRecordingIndicator() {
        const indicator = document.createElement('div');
        indicator.id = 'workflow-recorder-indicator';
        indicator.innerHTML = `
            <div style="
                position: fixed;
                top: 10px;
                right: 10px;
                background: #e74c3c;
                color: white;
                padding: 12px 20px;
                border-radius: 8px;
                font-family: Arial, sans-serif;
                font-size: 14px;
                font-weight: bold;
                z-index: 999999;
                box-shadow: 0 4px 12px rgba(0,0,0,0.3);
                cursor: pointer;
                animation: pulse 2s infinite;
            ">
                🔴 Recording... (Ctrl+Shift+S to stop)
            </div>
            <style>
                @keyframes pulse {
                    0%, 100% { opacity: 1; }
                    50% { opacity: 0.7; }
                }
            </style>
        `;
        
        indicator.onclick = stopRecording;
        document.body.appendChild(indicator);
    }
    
    /**
     * Initialize recorder
     */
    function initRecorder() {
        console.log('[RECORDER] Initializing browser action recorder...');
        
        // Create visual indicator
        createRecordingIndicator();
        
        // Attach event listeners
        document.addEventListener('click', handleClick, true);
        document.addEventListener('input', handleInput, true);
        document.addEventListener('keydown', handleKeyPress, true);
        
        // Navigation detection
        let lastUrl = window.location.href;
        setInterval(() => {
            if (window.location.href !== lastUrl) {
                lastUrl = window.location.href;
                handleNavigation();
            }
        }, 500);
        
        console.log('[RECORDER] ✓ Recorder initialized and listening');
        console.log('[RECORDER] Press Ctrl+Shift+S to stop recording');
    }
    
    // Auto-initialize when script loads
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initRecorder);
    } else {
        initRecorder();
    }
    
    // Expose stop function globally
    window.stopWorkflowRecording = stopRecording;
    
})();
