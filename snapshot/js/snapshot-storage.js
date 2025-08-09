// ============================================================================
// LIGHTWEIGHT STORAGE - File Handle Approach Compatible
// snapshot-storage.js - Updated to work with new file handle system
// ============================================================================

// Page refresh detection (consistent with PPA pattern)
window.addEventListener('beforeunload', () => {
    localStorage.setItem('spa_isLeaving', 'true');
});

function saveDataToStorage() {
    // Only save lightweight data - file reference is handled in snapshot-data.js
    if (laborData) {
        localStorage.setItem(SNAPSHOT_STORAGE_KEYS.laborData, JSON.stringify(laborData));
    }

    if (kpiResults) {
        localStorage.setItem(SNAPSHOT_STORAGE_KEYS.kpiResults, JSON.stringify(kpiResults));
    }
    
    console.log('✅ Lightweight data saved to storage');
}

function loadSavedData() {
    try {
        console.log('🔄 Loading saved snapshot data (file handle approach)...');
        
        // Load file reference and Excel KPIs using the new system
        const fileRef = getStoredFileReference && getStoredFileReference();
        if (fileRef) {
            // Restore file reference variables
            if (typeof selectedFileName !== 'undefined') {
                selectedFileName = fileRef.fileName;
                fileLastModified = fileRef.lastModified;
            } else {
                window.selectedFileName = fileRef.fileName;
                window.fileLastModified = fileRef.lastModified;
            }
            
            console.log(`📁 File reference restored: ${fileRef.fileName}`);
            console.log(`📊 File contains ${fileRef.totalRows?.toLocaleString() || 'unknown'} rows`);
            console.log(`💾 Storage used: ~${JSON.stringify(fileRef).length} bytes (metadata only)`);
            
            updateDataStatus('excel', `Referenced ✅ (${fileRef.fileName})`);
        }
        
        // Load processed Excel KPIs
        const savedExcelKPIs = localStorage.getItem(SNAPSHOT_STORAGE_KEYS.excelData);
        if (savedExcelKPIs) {
            excelData = JSON.parse(savedExcelKPIs);
            console.log('✅ Excel KPIs restored');
        }

        // Load Labor data
        const savedLabor = localStorage.getItem(SNAPSHOT_STORAGE_KEYS.laborData);
        if (savedLabor) {
            laborData = JSON.parse(savedLabor);
            updateDataStatus('labor', 'Loaded ✅');
            console.log('✅ Labor data restored');
        }

        // Load and display KPI results
        const savedKPIs = localStorage.getItem(SNAPSHOT_STORAGE_KEYS.kpiResults);
        if (savedKPIs) {
            kpiResults = JSON.parse(savedKPIs);
            
            // Display everything immediately
            if (typeof displayKPIs === 'function') {
                displayKPIs();
            }
            if (typeof displayDetailedAnalysis === 'function') {
                displayDetailedAnalysis();
            }
            if (typeof displayInsights === 'function') {
                displayInsights();
            }
            
            console.log('✅ KPI results restored and displayed');
        } else if (excelData || laborData) {
            if (typeof safeCalculateKPIs === 'function') {
                safeCalculateKPIs();
            }
        }

        // Show success message
        if (fileRef || laborData) {
            showDataRestorationSuccess(fileRef);
        }

    } catch (error) {
        console.warn('Failed to restore snapshot data:', error);
        showDataRestorationError();
    }
}

function clearAllData() {
    console.log('🗑️ Clearing all snapshot data...');
    
    // Clear localStorage
    Object.values(SNAPSHOT_STORAGE_KEYS).forEach(key => {
        try {
            localStorage.removeItem(key);
        } catch (error) {
            console.warn(`Failed to remove ${key}:`, error);
        }
    });

    // Reset global variables
    excelData = null;
    laborData = null;
    kpiResults = null;
    
    // Reset file handle variables (use new approach)
    if (typeof rawExcelDataCache !== 'undefined') {
        rawExcelDataCache = null;
    }
    if (typeof selectedFileHandle !== 'undefined') {
        selectedFileHandle = null;
    }
    if (typeof selectedFileName !== 'undefined') {
        selectedFileName = null;
        fileLastModified = null;
    }

    // Reset UI
    updateDataStatus('excel', 'Not Loaded');
    updateDataStatus('labor', 'Not Loaded');

    const elementsToHide = ['kpiDashboard', 'analysisSection', 'actionButtons', 'fileInfo', 'insightsSection'];
    elementsToHide.forEach(id => {
        const element = document.getElementById(id);
        if (element) element.style.display = 'none';
    });

    const fileInput = document.getElementById('fileInput');
    if (fileInput) fileInput.value = '';

    console.log('✅ All snapshot data and file references cleared');
}

function handleSnapshotPageLoad() {
    const wasLeaving = localStorage.getItem('spa_isLeaving');

    if (!wasLeaving) {
        console.log('🔄 Page refresh detected - clearing snapshot data');
        clearAllData();
    } else {
        console.log('🚀 SPA navigation detected - restoring snapshot data');
        loadSavedData();
    }

    localStorage.removeItem('spa_isLeaving');
}

function showDataRestorationSuccess(fileRef) {
    const statusDiv = document.createElement('div');
    statusDiv.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
        padding: 12px 16px;
        border-radius: 8px;
        z-index: 1000;
        font-weight: 500;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        max-width: 400px;
    `;
    
    const message = fileRef 
        ? `✅ Snapshot data restored!<br><small>📁 ${fileRef.fileName} (${fileRef.totalRows?.toLocaleString() || 'unknown'} rows)</small>`
        : '✅ Snapshot data restored!';
        
    statusDiv.innerHTML = message;
    
    document.body.appendChild(statusDiv);
    
    setTimeout(() => {
        if (statusDiv.parentNode) {
            statusDiv.remove();
        }
    }, 4000);
}

function showDataRestorationError() {
    const statusDiv = document.createElement('div');
    statusDiv.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background: #f8d7da;
        border: 1px solid #f5c6cb;
        color: #721c24;
        padding: 12px 16px;
        border-radius: 8px;
        z-index: 1000;
        font-weight: 500;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
    `;
    statusDiv.innerHTML = '⚠️ Could not restore data - please reload files';
    
    document.body.appendChild(statusDiv);
    
    setTimeout(() => {
        if (statusDiv.parentNode) {
            statusDiv.remove();
        }
    }, 5000);
}

function exportResults() {
    if (!kpiResults) {
        alert('No results to export. Please load and process data first.');
        return;
    }

    const fileRef = getStoredFileReference && getStoredFileReference();
    
    const exportData = {
        timestamp: new Date().toISOString(),
        kpiResults: kpiResults,
        fileReference: fileRef,
        summary: {
            laborLoaded: !!kpiResults.labor,
            excelLoaded: !!kpiResults.excel,
            combinedAnalysis: !!kpiResults.combined,
            fileReferenceApproach: true
        }
    };

    const dataStr = JSON.stringify(exportData, null, 2);
    const dataBlob = new Blob([dataStr], { type: 'application/json' });

    const link = document.createElement('a');
    link.href = URL.createObjectURL(dataBlob);
    link.download = `inbound_snapshot_analysis_${new Date().toISOString().split('T')[0]}.json`;
    link.click();
}