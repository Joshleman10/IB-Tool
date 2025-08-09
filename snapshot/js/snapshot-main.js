// ============================================================================
// MAIN ORCHESTRATOR & INITIALIZATION - Fixed version
// snapshot-main.js
// ============================================================================

// Store reference to original function before override
let originalDisplayKPIs = null;

// ============================================================================
// SAFE KPI CALCULATION WRAPPER
// ============================================================================

function safeCalculateKPIs() {
  if (typeof calculateKPIs === 'function') {
    calculateKPIs();
  } else {
    console.error('calculateKPIs function not available');
    // Retry after a short delay
    setTimeout(() => {
      if (typeof calculateKPIs === 'function') {
        calculateKPIs();
      } else {
        console.error('calculateKPIs still not available after retry');
      }
    }, 100);
  }
}

// Make the safe function globally available
window.safeCalculateKPIs = safeCalculateKPIs;

// ============================================================================
// ENHANCED FUNCTION AVAILABILITY CHECKER
// ============================================================================

function waitForFunctions(functionNames, callback, maxRetries = 10, retryDelay = 100) {
  let retries = 0;
  
  function check() {
    const allAvailable = functionNames.every(funcName => {
      const isAvailable = typeof window[funcName] === 'function';
      if (!isAvailable) {
        console.log(`Waiting for function: ${funcName} (attempt ${retries + 1})`);
      }
      return isAvailable;
    });
    
    if (allAvailable) {
      console.log('All required functions are now available');
      callback();
    } else if (retries < maxRetries) {
      retries++;
      setTimeout(check, retryDelay);
    } else {
      console.error('Timeout waiting for functions:', functionNames.filter(funcName => 
        typeof window[funcName] !== 'function'
      ));
      // Call callback anyway to continue initialization
      callback();
    }
  }
  
  check();
}

function initializeEventListeners() {
  // File input change event - defensive check with enhanced retry
  const fileInput = document.getElementById('fileInput');
  if (fileInput) {
    // Wait for the handleFile function to be available
    waitForFunctions(['handleFile'], () => {
      if (typeof handleFile === 'function') {
        // Remove any existing listeners first
        fileInput.removeEventListener('change', handleFile);
        fileInput.addEventListener('change', handleFile);
        console.log('File input event listener initialized successfully');
      } else {
        console.error('handleFile function still not available after waiting');
      }
    });
  } else {
    console.warn('File input element not found');
  }
  
  // KPI card click events for future drill-down capability
  const kpiCards = document.querySelectorAll('.kpi-card');
  kpiCards.forEach(card => {
    card.addEventListener('click', () => {
      console.log('KPI card clicked:', card.querySelector('.kpi-title')?.textContent);
      // Future enhancement: drill-down capability
    });
  });
  
  // Add paste button listener if available
  const pasteButton = document.getElementById('pasteButton');
  if (pasteButton) {
    waitForFunctions(['handlePasteClick'], () => {
      if (typeof handlePasteClick === 'function') {
        pasteButton.removeEventListener('click', handlePasteClick);
        pasteButton.addEventListener('click', handlePasteClick);
        console.log('Paste button event listener initialized successfully');
      }
    });
  }
  
  // Add clear data button listener if available
  const clearDataButton = document.getElementById('clearDataButton');
  if (clearDataButton) {
    waitForFunctions(['clearAllData'], () => {
      if (typeof clearAllData === 'function') {
        clearDataButton.removeEventListener('click', clearAllData);
        clearDataButton.addEventListener('click', clearAllData);
        console.log('Clear data button event listener initialized successfully');
      }
    });
  }
}

// Enhanced insights display function
function enhancedDisplayKPIs() {
  if (originalDisplayKPIs) {
    originalDisplayKPIs();
  }
  setTimeout(() => {
    if (typeof displayInsights === 'function') {
      displayInsights();
    }
  }, 100); // Small delay to ensure DOM is updated
}

// ============================================================================
// ENHANCED INITIALIZATION WITH BETTER ERROR HANDLING
// ============================================================================

function initializeSnapshotPage() {
  console.log('Starting Inbound Snapshot initialization...');
  
  // ENHANCED FUNCTION AVAILABILITY CHECK
  const requiredFunctions = [
    'calculateKPIs',
    'displayKPIs', 
    'parseLaborData',
    'handleFile',
    'handleSnapshotPageLoad'
  ];
  
  console.log('Function availability check:');
  requiredFunctions.forEach(funcName => {
    const isAvailable = typeof window[funcName] === 'function';
    console.log(`- ${funcName}: ${isAvailable ? 'function' : 'undefined'}`);
  });
  
  // Store reference to original displayKPIs function
  if (typeof displayKPIs === 'function') {
    originalDisplayKPIs = displayKPIs;
    // Override the displayKPIs function with enhanced version
    window.displayKPIs = enhancedDisplayKPIs;
    console.log('displayKPIs function enhanced with insights');
  } else {
    console.warn('displayKPIs function not available for enhancement');
  }
  
  // Initialize event listeners (with function waiting)
  initializeEventListeners();
  
  // Handle page load lifecycle
  if (typeof handleSnapshotPageLoad === 'function') {
    handleSnapshotPageLoad();
  } else {
    console.warn('handleSnapshotPageLoad function not available - will retry');
    // Try to load saved data directly if the function isn't available
    setTimeout(() => {
      if (typeof loadSavedData === 'function') {
        loadSavedData();
      }
    }, 200);
  }
  
  console.log('Snapshot initialization process started');
}

// ============================================================================
// SINGLE DOMContentLoaded HANDLER WITH ENHANCED ERROR HANDLING
// ============================================================================

document.addEventListener('DOMContentLoaded', () => {
  console.log('Inbound Snapshot page DOM loaded');
  
  // Small delay to ensure all scripts are loaded
  setTimeout(() => {
    initializeSnapshotPage();
    console.log('Snapshot initialization complete');
  }, 50);
});

// ============================================================================
// GLOBAL ERROR HANDLER FOR DEBUGGING
// ============================================================================

window.addEventListener('error', (event) => {
  if (event.filename && event.filename.includes('snapshot')) {
    console.error('Snapshot script error:', {
      message: event.message,
      filename: event.filename,
      lineno: event.lineno,
      colno: event.colno
    });
  }
});

// ============================================================================
// ADDITIONAL SAFETY MEASURES
// ============================================================================

// Ensure critical functions are available globally
window.addEventListener('load', () => {
  // Final check after everything should be loaded
  const criticalFunctions = ['handleFile', 'parseLaborData', 'safeCalculateKPIs'];
  const missingFunctions = criticalFunctions.filter(funcName => 
    typeof window[funcName] !== 'function'
  );
  
  if (missingFunctions.length > 0) {
    console.error('Critical functions still missing after page load:', missingFunctions);
    console.log('Available functions:', Object.keys(window).filter(key => 
      typeof window[key] === 'function' && key.includes('snapshot') || 
      ['handleFile', 'parseLaborData', 'calculateKPIs', 'displayKPIs'].includes(key)
    ));
  } else {
    console.log('All critical functions are available');
  }
});