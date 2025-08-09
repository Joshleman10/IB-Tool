// ============================================================================
// FILE HANDLE APPROACH - COMPLETE UPDATED VERSION
// snapshot-data.js - Lightweight storage with file reference system
// ============================================================================

// Global variables for file handle management
if (typeof rawExcelDataCache === 'undefined') {
    var rawExcelDataCache = null; // In-memory cache for current session
}
if (typeof selectedFileHandle === 'undefined') {
    var selectedFileHandle = null;
}
if (typeof selectedFileName === 'undefined') {
    var selectedFileName = null;
}
if (typeof fileLastModified === 'undefined') {
    var fileLastModified = null;
}

// ============================================================================
// MAIN FILE HANDLING FUNCTION
// ============================================================================

function handleFile(event) {
    const file = event.target.files[0];
    if (!file) return;

    displayFileInfo(file);

    if (typeof XLSX === 'undefined') {
        console.error('XLSX library not loaded');
        alert('Excel processing library not loaded. Please refresh the page and try again.');
        return;
    }

    // Store file information
    selectedFileName = file.name;
    fileLastModified = file.lastModified;

    console.log('=== FILE HANDLE APPROACH ===');
    console.log('File name:', selectedFileName);
    console.log('File size:', (file.size / (1024 * 1024)).toFixed(2) + ' MB');

    // Process file with new lightweight approach
    processExcelFileWithFileHandle(file);
}

// ============================================================================
// EXCEL FILE PROCESSING WITH FILE HANDLE APPROACH
// ============================================================================

async function processExcelFileWithFileHandle(file) {
    const reader = new FileReader();

    reader.onload = async function (e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array" });

            // Process the first sheet
            const firstSheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[firstSheetName];
            const rows = XLSX.utils.sheet_to_json(sheet);
            
            console.log('Excel data processed:', {
                sheets: workbook.SheetNames.length,
                firstSheetRows: rows.length,
                fileName: file.name
            });

            // Cache data in memory for current session
            rawExcelDataCache = rows;
            
            // Try to establish file handle for future access
            if ('showOpenFilePicker' in window) {
                try {
                    await establishFileHandle(file);
                } catch (handleError) {
                    console.log('Could not establish file handle, but continuing with processing');
                }
            }
            
            // Process KPIs (lightweight)
            const processedKPIs = calculateExcelKPIs({ [firstSheetName]: rows });
            excelData = processedKPIs;
            
            // Save only lightweight metadata to localStorage
            saveFileReferenceToStorage({
                fileName: selectedFileName,
                lastModified: fileLastModified,
                totalRows: rows.length,
                processedKPIs: processedKPIs,
                timestamp: new Date().toISOString(),
                fileSize: file.size,
                hasFileAccess: true
            });

            updateDataStatus('excel', `Loaded ✅ (${rows.length.toLocaleString()} rows)`);
            console.log('✅ Excel processing complete - file reference approach active');
            console.log('💾 Storage used: ~1KB (metadata only)');

            safeCalculateKPIs();
            
        } catch (error) {
            console.error('Error processing Excel file:', error);
            alert('Error processing Excel file. Please check the file format and try again.');
            updateDataStatus('excel', 'Error ❌');
        }
    };

    reader.onerror = function () {
        console.error('Error reading file');
        alert('Error reading file. Please try again.');
        updateDataStatus('excel', 'Error ❌');
    };

    reader.readAsArrayBuffer(file);
}

// ============================================================================
// FILE HANDLE MANAGEMENT FUNCTIONS
// ============================================================================

// Establish file handle for future access
async function establishFileHandle(originalFile) {
    try {
        if ('showOpenFilePicker' in window) {
            // Store file access capability for this session
            selectedFileHandle = { 
                available: true, 
                fileName: originalFile.name,
                lastModified: originalFile.lastModified
            };
            console.log('✅ File access established for current session');
            return true;
        }
    } catch (error) {
        console.log('File handle establishment failed:', error);
        return false;
    }
}

// Save file reference and metadata (not the actual data)
function saveFileReferenceToStorage(fileData) {
    try {
        localStorage.setItem(SNAPSHOT_STORAGE_KEYS.fileReference, JSON.stringify(fileData));
        if (fileData.processedKPIs) {
            localStorage.setItem(SNAPSHOT_STORAGE_KEYS.excelData, JSON.stringify(fileData.processedKPIs));
        }
        
        const storageSize = JSON.stringify(fileData).length;
        console.log(`✅ File reference saved (${storageSize} bytes - lightweight metadata only)`);
    } catch (error) {
        console.error('Error saving file reference:', error);
    }
}

// Get stored file reference
function getStoredFileReference() {
    try {
        const stored = localStorage.getItem(SNAPSHOT_STORAGE_KEYS.fileReference);
        return stored ? JSON.parse(stored) : null;
    } catch (error) {
        console.warn('Error loading file reference:', error);
        return null;
    }
}

// ============================================================================
// EXCEL DATA ACCESS FUNCTIONS
// ============================================================================

// Main function to get Excel data for analysis
async function getExcelDataForAnalysis() {
    // First check in-memory cache (fastest)
    if (rawExcelDataCache && rawExcelDataCache.length > 0) {
        console.log('✅ Using cached Excel data');
        return rawExcelDataCache;
    }
    
    // Check if we have file reference info
    const fileRef = getStoredFileReference();
    if (!fileRef || !fileRef.fileName) {
        console.log('No file reference available');
        return null;
    }
    
    // For detailed analysis, we need to re-read the file
    console.log(`📖 Need to re-read file for detailed analysis: ${fileRef.fileName}`);
    
    // Show user-friendly prompt for file access
    const shouldReselect = await showFileAccessPrompt(fileRef);
    if (shouldReselect) {
        return await requestFileReselection(fileRef);
    }
    
    return null;
}

// Synchronous version (uses cache only)
function getExcelDataForAnalysisSync() {
    if (rawExcelDataCache && rawExcelDataCache.length > 0) {
        return rawExcelDataCache;
    }
    
    console.log('No cached Excel data available - use getExcelDataForAnalysis() for file re-reading');
    return null;
}

// ============================================================================
// USER INTERFACE FOR FILE ACCESS
// ============================================================================

// Show user-friendly file access prompt
async function showFileAccessPrompt(fileRef) {
    return new Promise((resolve) => {
        // Create a nice modal-style prompt
        const modal = document.createElement('div');
        modal.style.cssText = `
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0,0,0,0.5);
            display: flex;
            justify-content: center;
            align-items: center;
            z-index: 2000;
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
        `;
        
        const dialog = document.createElement('div');
        dialog.style.cssText = `
            background: white;
            padding: 24px;
            border-radius: 12px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.15);
            max-width: 400px;
            text-align: center;
        `;
        
        dialog.innerHTML = `
            <h3 style="margin: 0 0 16px 0; color: #333;">📊 Access Analysis Data</h3>
            <p style="margin: 0 0 20px 0; color: #666; line-height: 1.4;">
                To view detailed VAS and CART analysis, please reselect your Excel file:
            </p>
            <p style="margin: 0 0 20px 0; font-weight: 500; color: #333;">
                📁 ${fileRef.fileName}
            </p>
            <p style="margin: 0 0 24px 0; font-size: 14px; color: #888;">
                Your file contains ${fileRef.totalRows?.toLocaleString() || 'unknown'} rows of data
            </p>
            <div style="display: flex; gap: 12px; justify-content: center;">
                <button id="selectFileBtn" style="
                    background: #007bff;
                    color: white;
                    border: none;
                    padding: 10px 20px;
                    border-radius: 6px;
                    cursor: pointer;
                    font-weight: 500;
                ">📁 Select File</button>
                <button id="cancelBtn" style="
                    background: #6c757d;
                    color: white;
                    border: none;
                    padding: 10px 20px;
                    border-radius: 6px;
                    cursor: pointer;
                    font-weight: 500;
                ">Cancel</button>
            </div>
        `;
        
        modal.appendChild(dialog);
        document.body.appendChild(modal);
        
        // Add event listeners
        document.getElementById('selectFileBtn').onclick = () => {
            document.body.removeChild(modal);
            resolve(true);
        };
        
        document.getElementById('cancelBtn').onclick = () => {
            document.body.removeChild(modal);
            resolve(false);
        };
        
        // Close on background click
        modal.onclick = (e) => {
            if (e.target === modal) {
                document.body.removeChild(modal);
                resolve(false);
            }
        };
    });
}

// Handle file reselection
async function requestFileReselection(fileRef) {
    try {
        if ('showOpenFilePicker' in window) {
            const [fileHandle] = await window.showOpenFilePicker({
                types: [{
                    description: 'Excel files',
                    accept: {
                        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
                        'application/vnd.ms-excel': ['.xls']
                    }
                }],
                excludeAcceptAllOption: true,
                multiple: false
            });
            
            const file = await fileHandle.getFile();
            
            // Verify it's the same file (or at least same name)
            if (file.name !== fileRef.fileName) {
                const proceed = confirm(
                    `You selected "${file.name}" but the original file was "${fileRef.fileName}". ` +
                    `Do you want to proceed with the new file?`
                );
                if (!proceed) return null;
            }
            
            // Process the reselected file
            const data = await readFileAsArrayBuffer(file);
            const workbook = XLSX.read(data, { type: "array" });
            const firstSheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[firstSheetName];
            const rows = XLSX.utils.sheet_to_json(sheet);
            
            // Update cache and reference
            rawExcelDataCache = rows;
            selectedFileName = file.name;
            selectedFileHandle = fileHandle;
            
            // Update stored reference
            saveFileReferenceToStorage({
                ...fileRef,
                fileName: file.name,
                lastModified: file.lastModified,
                totalRows: rows.length,
                timestamp: new Date().toISOString(),
                hasFileAccess: true
            });
            
            console.log(`✅ File reselected and ${rows.length} rows loaded`);
            return rows;
            
        } else {
            // Fallback for older browsers
            const fileInput = document.getElementById('fileInput');
            if (fileInput) {
                fileInput.click();
            }
            return null;
        }
        
    } catch (error) {
        console.error('File reselection failed:', error);
        return null;
    }
}

// Helper function to read file as ArrayBuffer
function readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve(new Uint8Array(e.target.result));
        reader.onerror = () => reject(reader.error);
        reader.readAsArrayBuffer(file);
    });
}

// ============================================================================
// TRANSACTION DATA HELPERS
// ============================================================================

// Helper function to get specific transaction types (optimized for performance)
async function getTransactionData(transactionType, fromLocationFilter = null) {
    const allData = await getExcelDataForAnalysis();
    if (!allData) return [];
    
    return allData.filter(row => {
        const rowTransactionType = String(row["Transaction Type"]).trim();
        const rowFromLocation = String(row["From Location"] || '').trim();
        
        // Check transaction type match
        const typeMatch = rowTransactionType === String(transactionType);
        
        // Check location filter if provided
        if (fromLocationFilter && typeMatch) {
            if (typeof fromLocationFilter === 'string') {
                return rowFromLocation.startsWith(fromLocationFilter);
            } else if (Array.isArray(fromLocationFilter)) {
                return fromLocationFilter.some(filter => rowFromLocation.startsWith(filter));
            }
        }
        
        return typeMatch;
    });
}

// Specific helpers for common transaction types
async function getType152Data() {
    return await getTransactionData('152');
}

async function getType151Data() {
    return await getTransactionData('151');
}

async function getVASData() {
    return await getTransactionData('600');
}

async function getCARTData() {
    return await getTransactionData('212', 'MOVEXX');
}

// ============================================================================
// DATA PERSISTENCE FUNCTIONS
// ============================================================================

// Load saved data function for file reference approach
function loadSavedData() {
    try {
        console.log('🔄 Loading saved snapshot data (file handle approach)...');
        
        // Load file reference
        const fileRef = getStoredFileReference();
        if (fileRef) {
            selectedFileName = fileRef.fileName;
            fileLastModified = fileRef.lastModified;
            
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

// Clear function updated for file reference approach
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
    rawExcelDataCache = null;
    selectedFileHandle = null;
    selectedFileName = null;
    fileLastModified = null;

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

// ============================================================================
// UI FEEDBACK FUNCTIONS
// ============================================================================

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

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

function displayFileInfo(file) {
    const fileInfoEl = document.getElementById('fileInfo');
    const fileNameEl = document.getElementById('fileName');
    const fileSizeEl = document.getElementById('fileSize');
    const fileDateEl = document.getElementById('fileDate');

    if (fileInfoEl && fileNameEl && fileSizeEl && fileDateEl) {
        fileInfoEl.style.display = 'block';
        fileNameEl.textContent = file.name;

        const sizeInMB = (file.size / (1024 * 1024)).toFixed(2);
        const sizeInKB = (file.size / 1024).toFixed(1);
        const sizeDisplay = file.size > 1024 * 1024 ? `${sizeInMB} MB` : `${sizeInKB} KB`;
        fileSizeEl.textContent = sizeDisplay;

        const fileDate = new Date(file.lastModified);
        const options = {
            year: 'numeric',
            month: 'short',
            day: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
        };
        fileDateEl.textContent = `Modified: ${fileDate.toLocaleDateString(undefined, options)}`;
    }
}

// ============================================================================
// LABOR DATA FUNCTIONS (unchanged from your original)
// ============================================================================

function handlePasteClick() {
    const processingStatus = document.getElementById('processingStatus');

    // Show processing status
    if (processingStatus) {
        processingStatus.style.display = 'block';
        processingStatus.innerHTML = '<span class="processing-text">📋 Ready to paste - use Ctrl+V (Cmd+V on Mac)</span>';
    }

    // Try to read from clipboard directly
    if (navigator.clipboard && navigator.clipboard.readText) {
        navigator.clipboard.readText()
            .then(pastedData => {
                if (pastedData && pastedData.trim().length > 10) {
                    if (processingStatus) {
                        processingStatus.innerHTML = '<span class="processing-text">🔄 Processing pasted data...</span>';
                    }

                    setTimeout(() => {
                        processPastedDataFromClipboard(pastedData);
                    }, 300);
                } else {
                    showPasteError(processingStatus, 'No valid data found in clipboard');
                }
            })
            .catch(err => {
                // Fallback to textarea method
                useFallbackPasteMethod(processingStatus);
            });
    } else {
        // Fallback for older browsers
        useFallbackPasteMethod(processingStatus);
    }
}

function useFallbackPasteMethod(processingStatus) {
    const hiddenTextarea = document.getElementById('hiddenPasteArea');

    if (hiddenTextarea) {
        hiddenTextarea.focus();
        hiddenTextarea.value = '';

        const handlePaste = (e) => {
            e.preventDefault();
            const pastedData = e.clipboardData.getData('text');

            if (pastedData && pastedData.trim().length > 10) {
                if (processingStatus) {
                    processingStatus.innerHTML = '<span class="processing-text">🔄 Processing pasted data...</span>';
                }

                setTimeout(() => {
                    processPastedDataFromClipboard(pastedData);
                }, 300);
            } else {
                showPasteError(processingStatus, 'No valid data found in clipboard');
            }

            hiddenTextarea.removeEventListener('paste', handlePaste);
        };

        hiddenTextarea.addEventListener('paste', handlePaste);

        setTimeout(() => {
            if (processingStatus && processingStatus.innerHTML.includes('Ready to paste')) {
                processingStatus.style.display = 'none';
            }
            hiddenTextarea.removeEventListener('paste', handlePaste);
        }, 10000);
    }
}

function processPastedDataFromClipboard(rawData) {
    const processingStatus = document.getElementById('processingStatus');

    if (!rawData || !rawData.trim()) {
        showPasteError(processingStatus, 'No data to process');
        return;
    }

    try {
        laborData = parseLaborData(rawData.trim());
        localStorage.setItem(SNAPSHOT_STORAGE_KEYS.laborData, JSON.stringify(laborData));

        updateDataStatus('labor', 'Loaded ✅');

        if (processingStatus) {
            processingStatus.innerHTML = '<span class="processing-text success">✅ Data processed successfully!</span>';
            setTimeout(() => {
                processingStatus.style.display = 'none';
            }, 2000);
        }

        console.log('Labor data processed:', laborData);
        safeCalculateKPIs();
    } catch (error) {
        console.error('Error processing pasted data:', error);
        showPasteError(processingStatus, 'Error processing data. Please check format.', 3000);
        updateDataStatus('labor', 'Error ❌');
    }
}

function showPasteError(processingStatus, message, timeout = 2000) {
    if (processingStatus) {
        processingStatus.innerHTML = `<span class="processing-text error">❌ ${message}</span>`;
        setTimeout(() => {
            processingStatus.style.display = 'none';
        }, timeout);
    }
}

function parseLaborData(rawData) {
    const lines = rawData.split('\n').map(line => line.trim()).filter(line => line);
    const parsed = {
        departments: [],
        areas: [],
        functions: [],
        metadata: {}
    };

    let currentSection = null;
    let currentDepartment = null;
    let isTableData = false;

    // Enhanced parsing to capture all the detailed function data
    lines.forEach((line, index) => {
        // Skip header/navigation lines
        if (line.includes('Chewy Labor Management') ||
            line.includes('Welcome,') ||
            line.includes('Reports') ||
            line.includes('Actions') ||
            line.includes('Labor Area Summary') ||
            line.includes('Report an Issue') ||
            line.includes('FAQ/User Guide')) {
            return;
        }

        // Extract metadata with more precision
        if (line.includes('Last Updated:')) {
            parsed.metadata.lastUpdated = line.split('Last Updated:')[1]?.trim();
            return;
        }

        if (line.includes('FC:')) {
            const fcMatch = line.match(/FC:\s*([A-Z0-9]+)/);
            if (fcMatch) {
                parsed.metadata.fc = fcMatch[1];
            }
            return;
        }

        if (line.includes('Start Date:')) {
            parsed.metadata.startDate = line.split('Start Date:')[1]?.trim();
            return;
        }

        if (line.includes('End Date:')) {
            parsed.metadata.endDate = line.split('End Date:')[1]?.trim();
            return;
        }

        if (line.includes('FC Timezone:')) {
            parsed.metadata.timezone = line.split('FC Timezone:')[1]?.trim();
            return;
        }

        // Identify main sections
        if (line === 'Labor Department Totals') {
            currentSection = 'departments';
            isTableData = true;
            return;
        } else if (line === 'Labor Area Totals') {
            currentSection = 'areas';
            isTableData = true;
            return;
        }

        // Detect department subsections (like Putaway, Receiving, etc.)
        const departmentSections = [
            'Putaway', 'Receiving', 'Support', 'Unallocated', 'VAS',
            'Customer Returns', 'Inventory Control', 'Outbound',
            'Administrative', 'Consolidation', 'Fresh', 'Returns',
            'Replenishment'
        ];

        if (departmentSections.some(dept => line.includes(dept) && !line.includes('/'))) {
            // Check if this is a standalone department section header
            const exactMatch = departmentSections.find(dept =>
                line.trim() === dept || line.trim() === dept + ' Totals'
            );
            if (exactMatch) {
                currentDepartment = exactMatch;
                currentSection = 'functions';
                isTableData = true;
                return;
            }
        }

        // Skip table headers
        if (line.includes('Labor Function') ||
            line.includes('Total Hours') ||
            line.includes('Labor Department') ||
            line.includes('Labor Dept / Area')) {
            return;
        }

        // Process department totals
        if (currentSection === 'departments' && isTableData) {
            const parts = line.split(/\s{2,}|\t/);

            if (parts.length >= 6 && !line.includes('Totals')) {
                const item = {
                    name: parts[0],
                    totalHours: parseFloat(parts[1]) || 0,
                    totalUnits: parseInt(parts[2]) || 0,
                    uph: parseFloat(parts[3]) || 0,
                    totalTransactions: parseInt(parts[4]) || 0,
                    tph: parseFloat(parts[5]) || 0
                };
                parsed.departments.push(item);
            }
        }

        // Process area data
        if (currentSection === 'areas' && isTableData) {
            const parts = line.split(/\s{2,}|\t/);

            if (parts.length >= 6 && parts[0].includes('/') && !line.includes('Totals')) {
                const [dept, area] = parts[0].split(' / ', 2);
                const item = {
                    department: dept,
                    name: area || dept,
                    fullName: parts[0],
                    totalHours: parseFloat(parts[1]) || 0,
                    totalUnits: parseInt(parts[2]) || 0,
                    uph: parseFloat(parts[3]) || 0,
                    totalTransactions: parseInt(parts[4]) || 0,
                    tph: parseFloat(parts[5]) || 0
                };
                parsed.areas.push(item);
            }
        }

        // Process function data (the detailed breakdowns we need)
        if (currentSection === 'functions' && isTableData && currentDepartment) {
            const parts = line.split(/\s{2,}|\t/);

            if (parts.length >= 6 && !line.includes('Totals')) {
                const item = {
                    department: currentDepartment,
                    name: parts[0],
                    totalHours: parseFloat(parts[1]) || 0,
                    totalUnits: parseInt(parts[2]) || 0,
                    uph: parseFloat(parts[3]) || 0,
                    totalTransactions: parseInt(parts[4]) || 0,
                    tph: parseFloat(parts[5]) || 0
                };
                parsed.functions.push(item);
            }
        }

        // Reset section on totals or when encountering new sections
        if (line.includes('Totals') &&
            !line.includes('Labor Department Totals') &&
            !line.includes('Labor Area Totals')) {

            // Only reset if we've processed some data
            if (currentSection === 'functions') {
                currentDepartment = null;
            }
            isTableData = false;
        }

        // Reset everything for Summary section
        if (line === 'Summary') {
            currentSection = null;
            currentDepartment = null;
            isTableData = false;
        }
    });

    // Post-processing: ensure we have the key functions we need
    console.log('=== PARSING DEBUG ===');
    console.log('Raw lines processed:', lines.length);
    console.log('Departments found:', parsed.departments.length);
    console.log('Areas found:', parsed.areas.length);
    console.log('Functions found:', parsed.functions.length);
    console.log('All parsed functions:', parsed.functions.map(f => ({ name: f.name, hours: f.totalHours, units: f.totalUnits, transactions: f.totalTransactions })));
    console.log('Functions with hours > 0:', parsed.functions.filter(f => f.totalHours > 0));
    console.log('Total functions found:', parsed.functions.length);

    return parsed;
}

function validateLaborData(laborData) {
    const expectedDirectFunctions = [
        'Non- PIT Manual Putaway', 'Putaway - Unknown', 'Reach-Truck Putaway',
        'Break Down Receiving', 'Cart Receiving', 'Container Receiver',
        'Full Pallet Receiving', 'Parcel Receiving', 'Receiving',
        'Small Pack Debundle', 'VAS Execute'
    ];

    const expectedIndirectFunctions = [
        'Breakdown Receiver', 'Container Unload', 'Pallet Rework',
        'Pallet Wrangler – Dock Stocker IB', 'Pallet Wrangler – Reach IB',
        'Pallet Wrapper', 'Parcel Receiver', 'Unloader',
        'Inbound Lead', 'Inbound Training', 'Problem Solver',
        'Vendor Compliance', 'On-Clock Unallocated'
    ];

    const foundFunctions = laborData.functions.map(f => f.name);
    const missingDirect = expectedDirectFunctions.filter(func =>
        !foundFunctions.includes(func));
    const missingIndirect = expectedIndirectFunctions.filter(func =>
        !foundFunctions.includes(func));

    console.log('Missing direct functions:', missingDirect);
    console.log('Missing indirect functions:', missingIndirect);

    return {
        isValid: missingDirect.length === 0 && missingIndirect.length === 0,
        missingDirect,
        missingIndirect,
        foundFunctions
    };
}

// ============================================================================
// GLOBAL EXPORTS
// ============================================================================

// Make sure these functions are globally available for other files
window.getExcelDataForAnalysis = getExcelDataForAnalysis;
window.getExcelDataForAnalysisSync = getExcelDataForAnalysisSync;
window.getTransactionData = getTransactionData;
window.getType152Data = getType152Data;
window.getType151Data = getType151Data;
window.getVASData = getVASData;
window.getCARTData = getCARTData;
window.loadSavedData = loadSavedData;
window.clearAllData = clearAllData;
window.showDataRestorationSuccess = showDataRestorationSuccess;
window.showDataRestorationError = showDataRestorationError;