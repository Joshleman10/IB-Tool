let itemMasterData = {};
let transactionData = [];
let airLocationSet = new Set();
let totalAirTransactions = 0;
let totalGroundTransactions = 0;
let isComparisonMode = false;
let period1Data = null;
let period2Data = null;

// Storage key for persistence
const AIR_GROUND_STORAGE_KEY = 'airVsGround_transactionData';

// Page refresh detection (same pattern as other pages)
window.addEventListener('beforeunload', () => {
  localStorage.setItem('spa_isLeaving', 'true');
});

// Load item master
fetch('item_master.json')
  .then(response => response.json())
  .then(data => {
    itemMasterData = data['Sheet1'].reduce((map, item) => {
      // Handle both string and number formats
      const key = Number(item.ITEM_NUMBER);
      map[key] = item;
      return map;
    }, {});
    console.log('Item master data loaded successfully');
    console.log('Sample items in master data:', Object.keys(itemMasterData).slice(0, 10));
    console.log('Total items in master data:', Object.keys(itemMasterData).length);
  })
  .catch(error => {
    console.error('Error loading item master data:', error);
  });

function handleFile(event) {
  const file = event.target.files[0];
  
  if (!file) return;

  // Display file information
  displayFileInfo(file);

  // Check if XLSX library is loaded
  if (typeof XLSX === 'undefined') {
    console.error('XLSX library not loaded');
    alert('Excel processing library not loaded. Please refresh the page and try again.');
    return;
  }

  const reader = new FileReader();
  
  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(firstSheet, { defval: "" });

      transactionData = rows.map(row => ({
        date: row["Start Date"],
        location: row["To Location"] || "",
        item: row["Item Number"],
        quantity: Number(row["Quantity"] || 0),
        timeToExecute: Number(row["Time to Execute"] || 0),
        employeeId: row["Employee ID"] || ""
      }));

      // Save to localStorage
      localStorage.setItem(AIR_GROUND_STORAGE_KEY, JSON.stringify(transactionData));
      console.log('Air vs Ground data saved to localStorage');

      // Reset to full data view
      resetToFullData();
      
      // Show date range controls and clear button
      document.getElementById("dateRangeControls").style.display = "block";
      document.getElementById("clearButton").style.display = "inline-block";
      
      // Set up date range inputs with min/max values
      setupDateRangeInputs();
      
    } catch (error) {
      console.error('Error processing Excel file:', error);
      alert('Error processing Excel file. Please check the file format and try again.');
    }
  };

  reader.onerror = function() {
    console.error('Error reading file');
    alert('Error reading file. Please try again.');
  };

  reader.readAsArrayBuffer(file);
}

function displayFileInfo(file) {
  const fileInfoEl = document.getElementById('fileInfo');
  const fileNameEl = document.getElementById('fileName');
  const fileSizeEl = document.getElementById('fileSize');
  const fileDateEl = document.getElementById('fileDate');
  
  if (fileInfoEl && fileNameEl && fileSizeEl && fileDateEl) {
    // Show file info section
    fileInfoEl.style.display = 'block';
    
    // Display file name
    fileNameEl.textContent = file.name;
    
    // Display file size
    const sizeInMB = (file.size / (1024 * 1024)).toFixed(2);
    const sizeInKB = (file.size / 1024).toFixed(1);
    const sizeDisplay = file.size > 1024 * 1024 ? `${sizeInMB} MB` : `${sizeInKB} KB`;
    fileSizeEl.textContent = sizeDisplay;
    
    // Display file last modified date
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

function setupDateRangeInputs() {
  const dates = transactionData
    .map(t => new Date(t.date))
    .filter(d => !isNaN(d.getTime()))
    .sort((a, b) => a - b);

  if (dates.length === 0) return;

  const minDate = dates[0];
  const maxDate = dates[dates.length - 1];
  const midDate = new Date(minDate.getTime() + (maxDate.getTime() - minDate.getTime()) / 2);

  // Format dates for input fields
  const formatDate = (date) => date.toISOString().split('T')[0];

  // Set min/max constraints
  const inputs = ['period1Start', 'period1End', 'period2Start', 'period2End'];
  inputs.forEach(id => {
    const input = document.getElementById(id);
    input.min = formatDate(minDate);
    input.max = formatDate(maxDate);
  });

  // Set default values (split data in half)
  document.getElementById('period1Start').value = formatDate(minDate);
  document.getElementById('period1End').value = formatDate(midDate);
  document.getElementById('period2Start').value = formatDate(new Date(midDate.getTime() + 86400000)); // +1 day
  document.getElementById('period2End').value = formatDate(maxDate);
}

function displayDateRange() {
  const dates = transactionData
    .map(t => new Date(t.date))
    .filter(d => !isNaN(d.getTime()))
    .sort((a, b) => a - b);

  const dateRangeEl = document.getElementById("dateRange");
  if (dateRangeEl) {
    if (dates.length > 0) {
      const options = { year: 'numeric', month: 'short', day: 'numeric' };
      const minDate = dates[0].toLocaleDateString(undefined, options);
      const maxDate = dates[dates.length - 1].toLocaleDateString(undefined, options);
      dateRangeEl.textContent = `Date Range: ${minDate} – ${maxDate}`;
    } else {
      dateRangeEl.textContent = "Date Range: Not available";
    }
  }
}

function filterDataByDateRange(data, startDate, endDate) {
  const start = new Date(startDate);
  const end = new Date(endDate);
  
  return data.filter(entry => {
    const entryDate = new Date(entry.date);
    return entryDate >= start && entryDate <= end;
  });
}

// IMPROVED CART TRANSACTION LOGIC WITH BETTER ERROR HANDLING
function calculateMetricsForData(data) {
  let airTotal = 0, airCount = 0;
  let groundTotal = 0, groundCount = 0;
  const allSmallVolume = [], cartToAir = [];
  let modTransactionCount = 0;
  
  // Enhanced debugging arrays
  const debugInfo = {
    itemsNotFound: [],
    invalidVolumes: [],
    processingErrors: [],
    itemLookupIssues: [],
    volumeCalculations: []
  };

  console.log('Starting calculateMetricsForData with', data.length, 'transactions');
  console.log('Item master data available:', Object.keys(itemMasterData).length, 'items');

  data.forEach((entry, index) => {
    try {
      const minutes = Math.round((entry.timeToExecute / 60) * 100) / 100;
      const loc = (entry.location || "").trim().toUpperCase();
      const suffix = loc.split("-")[2] || "";
      
      const formatRegex = /^\d{2}-\d{2}-[A-Z0-9]{3}$/i;
      if (formatRegex.test(loc)) modTransactionCount++;

      // Air vs Ground classification
      if (/[ABC]/i.test(suffix)) {
        groundTotal += minutes;
        groundCount++;
      } else {
        airTotal += minutes;
        airCount++;
      }

      // ENHANCED VOLUME ANALYSIS FOR NEW JSON FORMAT
      if (formatRegex.test(loc)) {
        // Validate item number exists
        const itemNumber = entry.item;
        if (!itemNumber || itemNumber === "" || itemNumber === null || itemNumber === undefined) {
          debugInfo.processingErrors.push({
            index, 
            issue: 'Missing or null item number', 
            entry: { location: loc, item: itemNumber, quantity: entry.quantity }
          });
          return;
        }
        
        // Try multiple ways to match the item number
        let item;
        let itemData;
        
        // Method 1: Direct number conversion
        try {
          item = Number(itemNumber);
          if (isNaN(item) || item <= 0) {
            throw new Error('Invalid number conversion');
          }
          itemData = itemMasterData[item];
        } catch (e) {
          debugInfo.itemLookupIssues.push({
            index,
            method: 'Direct number conversion',
            itemNumber,
            error: e.message
          });
        }
        
        // Method 2: Try as string key if direct number failed
        if (!itemData) {
          const itemStr = String(itemNumber);
          itemData = itemMasterData[itemStr];
          if (itemData) {
            item = itemStr;
            debugInfo.itemLookupIssues.push({
              index,
              method: 'String lookup succeeded',
              itemNumber,
              foundAs: itemStr
            });
          }
        }
        
        // Method 3: Try Math.floor conversion (original logic)
        if (!itemData) {
          try {
            item = Math.floor(Number(itemNumber));
            if (!isNaN(item) && item > 0) {
              itemData = itemMasterData[item];
              if (itemData) {
                debugInfo.itemLookupIssues.push({
                  index,
                  method: 'Math.floor conversion succeeded',
                  itemNumber,
                  convertedTo: item
                });
              }
            }
          } catch (e) {
            debugInfo.itemLookupIssues.push({
              index,
              method: 'Math.floor conversion failed',
              itemNumber,
              error: e.message
            });
          }
        }
        
        // If still no item found, log and skip
        if (!itemData) {
          debugInfo.itemsNotFound.push({
            index,
            itemNumber: itemNumber,
            triedAsNumber: Number(itemNumber),
            triedAsString: String(itemNumber),
            triedAsFloor: Math.floor(Number(itemNumber)),
            location: loc,
            availableKeys: Object.keys(itemMasterData).slice(0, 5) // Show first 5 keys for reference
          });
          return;
        }
        
        // Validate quantity
        const quantity = Number(entry.quantity);
        if (!quantity || isNaN(quantity) || quantity <= 0) {
          debugInfo.processingErrors.push({
            index, 
            issue: 'Invalid quantity', 
            entry: { location: loc, item: itemNumber, quantity: entry.quantity }
          });
          return;
        }

        // Get and validate volume from master data
        const cubicVolume = Number(itemData.CUBIC_VOL);
        
        if (isNaN(cubicVolume) || cubicVolume <= 0) {
          debugInfo.invalidVolumes.push({
            index,
            itemNumber,
            rawCubicVol: itemData.CUBIC_VOL,
            convertedCubicVol: cubicVolume,
            location: loc,
            itemDataStructure: Object.keys(itemData)
          });
          return;
        }
        
        // Calculate total volume
        const totalVolume = cubicVolume * quantity;
        
        // Log a few successful calculations for verification
        if (debugInfo.volumeCalculations.length < 5) {
          debugInfo.volumeCalculations.push({
            index,
            itemNumber,
            quantity,
            cubicVolume,
            totalVolume,
            isCartSize: totalVolume < 5000,
            location: loc,
            suffix
          });
        }
        
        // Cart transaction criteria: volume < 5000
        if (totalVolume < 5000) {
          const cartEntry = {
            ...entry,
            itemData: {
              itemNumber: itemNumber,
              matchedItem: item,
              cubicVolume: cubicVolume,
              quantity: quantity,
              totalVolume: totalVolume
            },
            location: loc,
            suffix: suffix
          };
          
          allSmallVolume.push(cartEntry);
          
          // If not ground (A, B, C suffix), then it's cart to air
          if (!/[ABC]/i.test(suffix)) {
            cartToAir.push(cartEntry);
          }
        }
      }
    } catch (error) {
      debugInfo.processingErrors.push({
        index,
        issue: 'Unexpected processing error',
        error: error.message,
        stack: error.stack,
        entry: entry
      });
    }
  });

  // Calculate final metrics
  const total = airCount + groundCount;
  const cartToGround = allSmallVolume.length - cartToAir.length;

  const metrics = {
    airCount,
    groundCount,
    total,
    avgAir: airCount ? (airTotal / airCount).toFixed(2) : "0.00",
    avgGround: groundCount ? (groundTotal / groundCount).toFixed(2) : "0.00",
    airPercent: total > 0 ? (airCount / total * 100).toFixed(2) : "0.00",
    groundPercent: total > 0 ? (groundCount / total * 100).toFixed(2) : "0.00",
    smallVolTotal: allSmallVolume.length,
    cartToAir: cartToAir.length,
    cartToGround,
    cartAirPercent: allSmallVolume.length > 0 ? (cartToAir.length / allSmallVolume.length * 100).toFixed(2) : "0.00",
    modTransactionCount,
    debugInfo
  };

  // Enhanced console logging for debugging
  console.log('=== CART TRANSACTION ANALYSIS COMPLETE ===');
  console.log('Basic Stats:', {
    totalTransactions: data.length,
    modFormatTransactions: modTransactionCount,
    cartTransactions: allSmallVolume.length,
    cartToAir: cartToAir.length,
    cartToGround: cartToGround
  });
  
  if (debugInfo.itemsNotFound.length > 0) {
    console.warn(`⚠️  ${debugInfo.itemsNotFound.length} items not found in master data:`);
    console.table(debugInfo.itemsNotFound.slice(0, 10));
    console.log('Sample of available master data keys:', Object.keys(itemMasterData).slice(0, 20));
  }
  
  if (debugInfo.invalidVolumes.length > 0) {
    console.warn(`⚠️  ${debugInfo.invalidVolumes.length} items with invalid volume data:`);
    console.table(debugInfo.invalidVolumes.slice(0, 5));
  }
  
  if (debugInfo.processingErrors.length > 0) {
    console.warn(`⚠️  ${debugInfo.processingErrors.length} processing errors:`);
    console.table(debugInfo.processingErrors.slice(0, 5));
  }
  
  if (debugInfo.itemLookupIssues.length > 0) {
    console.log(`ℹ️  ${debugInfo.itemLookupIssues.length} item lookup method variations used`);
    console.table(debugInfo.itemLookupIssues.slice(0, 10));
  }
  
  if (debugInfo.volumeCalculations.length > 0) {
    console.log('✅ Sample successful volume calculations:');
    console.table(debugInfo.volumeCalculations);
  }
  
  console.log('Final Cart Metrics:', {
    totalCartEligible: allSmallVolume.length,
    cartToAirCount: cartToAir.length,
    cartToGroundCount: cartToGround,
    cartAirPercentage: metrics.cartAirPercent + '%'
  });

  return metrics;
}

// Enhanced updateMetricsDisplay function with comparison support
function updateMetricsDisplay(metrics, suffix = '') {
  const updateElement = (id, text) => {
    const el = document.getElementById(id + suffix);
    if (el) el.textContent = text;
  };

  updateElement("totalAirTrans", `Total Air Transactions: ${metrics.airCount}`);
  updateElement("totalGroundTrans", `Total Ground Transactions: ${metrics.groundCount}`);
  updateElement("avgAirTime", `Avg Air Time (min): ${metrics.avgAir}`);
  updateElement("avgGroundTime", `Avg Ground Time (min): ${metrics.avgGround}`);
  updateElement("airPercent", `Air %: ${metrics.airPercent}%`);
  updateElement("groundPercent", `Ground %: ${metrics.groundPercent}%`);
  updateElement("smallVolTotal", `Potential Cart Total: ${metrics.smallVolTotal}`);
  updateElement("cartAirTotal", `Cart to Air Total: ${metrics.cartToAir}`);
  updateElement("cartGroundTotal", `Cart to Ground Total: ${metrics.cartToGround}`);
  updateElement("cartAirPercent", `Cart to Air %: ${metrics.cartAirPercent}%`); // This was missing in comparison mode
  updateElement("total212", `Total MOD Transactions: ${metrics.modTransactionCount}`);
}

function createChartsForData(data, chartSuffix = '') {
  const dailySummary = {};
  data.forEach(entry => {
    const date = entry.date;
    if (!dailySummary[date]) dailySummary[date] = { air: 0, ground: 0 };
    const suffix = (entry.location || '').split("-")[2] || "";
    if (/[ABC]/i.test(suffix)) dailySummary[date].ground++;
    else dailySummary[date].air++;
  });

  const dates = Object.keys(dailySummary).sort();
  const airCounts = dates.map(d => dailySummary[d].air);
  const groundCounts = dates.map(d => dailySummary[d].ground);
  const airPercents = dates.map(d => {
    const total = dailySummary[d].air + dailySummary[d].ground;
    return total ? (dailySummary[d].air / total * 100).toFixed(2) : 0;
  });
  const groundPercents = dates.map(d => {
    const total = dailySummary[d].air + dailySummary[d].ground;
    return total ? (dailySummary[d].ground / total * 100).toFixed(2) : 0;
  });

  // Check if Chart.js is loaded
  if (typeof Chart === 'undefined') {
    console.error('Chart.js not loaded');
    return;
  }

  // Create Daily Chart
  const dailyChartEl = document.getElementById('dailyChart' + chartSuffix);
  if (dailyChartEl) {
    if (dailyChartEl.chart) dailyChartEl.chart.destroy();
    
    dailyChartEl.chart = new Chart(dailyChartEl.getContext('2d'), {
      type: 'line',
      data: {
        labels: dates,
        datasets: [
          {
            label: 'Air Transactions',
            data: airCounts,
            borderColor: '#3f51b5',
            fill: false,
            tension: 0.3
          },
          {
            label: 'Ground Transactions',
            data: groundCounts,
            borderColor: '#4caf50',
            fill: false,
            tension: 0.3
          }
        ]
      },
      options: {
        responsive: true,
        plugins: {
          title: {
            display: !chartSuffix,
            text: 'Daily Air vs Ground Transactions',
            font: { size: 18 }
          },
          legend: { position: 'top' }
        }
      }
    });
  }

  // Create Percentage Chart
  const percentChartEl = document.getElementById('percentChart' + chartSuffix);
  if (percentChartEl) {
    if (percentChartEl.chart) percentChartEl.chart.destroy();
    
    percentChartEl.chart = new Chart(percentChartEl.getContext('2d'), {
      type: 'line',
      data: {
        labels: dates,
        datasets: [
          {
            label: '% Ground',
            data: groundPercents,
            borderColor: '#388e3c',
            backgroundColor: 'rgba(76, 175, 80, 0.2)',
            fill: true,
            tension: 0.3
          },
          {
            label: '% Air',
            data: airPercents,
            borderColor: '#303f9f',
            backgroundColor: 'rgba(63, 81, 181, 0.2)',
            fill: true,
            tension: 0.3
          }
        ]
      },
      options: {
        responsive: true,
        plugins: {
          title: {
            display: !chartSuffix,
            text: 'Percentage of Ground vs Air Work by Day',
            font: { size: 18 }
          },
          legend: { position: 'top' }
        },
        scales: {
          y: {
            min: 0,
            max: 100,
            title: { display: true, text: '%' }
          }
        }
      }
    });
  }
}

function comparePeriods() {
  const period1Start = document.getElementById('period1Start').value;
  const period1End = document.getElementById('period1End').value;
  const period2Start = document.getElementById('period2Start').value;
  const period2End = document.getElementById('period2End').value;

  if (!period1Start || !period1End || !period2Start || !period2End) {
    alert('Please select all date ranges before comparing.');
    return;
  }

  // Validate date ranges
  if (new Date(period1Start) >= new Date(period1End)) {
    alert('Period 1 start date must be before end date.');
    return;
  }
  if (new Date(period2Start) >= new Date(period2End)) {
    alert('Period 2 start date must be before end date.');
    return;
  }

  // Filter data for each period
  period1Data = filterDataByDateRange(transactionData, period1Start, period1End);
  period2Data = filterDataByDateRange(transactionData, period2Start, period2End);

  if (period1Data.length === 0) {
    alert('No data found for Period 1. Please adjust the date range.');
    return;
  }
  if (period2Data.length === 0) {
    alert('No data found for Period 2. Please adjust the date range.');
    return;
  }

  // Switch to comparison mode
  isComparisonMode = true;
  document.getElementById('fullDataCharts').style.display = 'none';
  document.getElementById('fullDataMetrics').style.display = 'none';
  document.getElementById('comparisonCharts').style.display = 'block';
  document.getElementById('comparisonMetrics').style.display = 'block';

  // Create charts for each period
  createChartsForData(period1Data, '1');
  createChartsForData(period2Data, '2');

  // Calculate and display metrics for each period
  const metrics1 = calculateMetricsForData(period1Data);
  const metrics2 = calculateMetricsForData(period2Data);
  
  // Update basic metrics display
  updateMetricsDisplay(metrics1, '1');
  updateMetricsDisplay(metrics2, '2');

  // Calculate and display percentage changes
  displayPercentageChanges(metrics1, metrics2);

  // Update date range display
  const formatDate = (dateStr) => new Date(dateStr).toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
  document.getElementById("dateRange").textContent = 
    `Comparison: ${formatDate(period1Start)} - ${formatDate(period1End)} vs ${formatDate(period2Start)} - ${formatDate(period2End)}`;

  console.log('Period comparison completed with percentage changes');
}

function resetToFullData() {
  isComparisonMode = false;
  period1Data = null;
  period2Data = null;

  // Show full data views
  document.getElementById('fullDataCharts').style.display = 'block';
  document.getElementById('fullDataMetrics').style.display = 'block';
  document.getElementById('comparisonCharts').style.display = 'none';
  document.getElementById('comparisonMetrics').style.display = 'none';

  // Update with full dataset
  displayDateRange();
  document.getElementById("inputData").value = transactionData.map(t => t.location).join("\n");
  createChartsForData(transactionData);
  filterData();
  calculateTimeAverages();
}

function calculateTimeAverages() {
  const metrics = calculateMetricsForData(transactionData);
  updateMetricsDisplay(metrics);
  
  totalAirTransactions = metrics.airCount;
  totalGroundTransactions = metrics.groundCount;
}

function filterData() {
  const rawData = document.getElementById("inputData").value;
  const lines = rawData.split(/\r?\n/).map(line => line.trim()).filter(line => line !== "");

  const formatRegex = /^\d{2}-\d{2}-[A-Z0-9]{3}$/i;
  const validLines = lines.filter(line => formatRegex.test(line));
  airLocationSet = new Set(
    validLines.filter(loc => !/[ABC]/i.test((loc.split("-")[2] || "")))
  );

  const metrics = calculateMetricsForData(transactionData);
  updateMetricsDisplay(metrics);
}

function clearAirGroundData() {
  localStorage.removeItem(AIR_GROUND_STORAGE_KEY);
  
  // Reset data
  transactionData = [];
  airLocationSet = new Set();
  totalAirTransactions = 0;
  totalGroundTransactions = 0;
  period1Data = null;
  period2Data = null;
  isComparisonMode = false;
  
  // Clear input
  const inputDataEl = document.getElementById("inputData");
  if (inputDataEl) inputDataEl.value = "";
  
  // Reset file input
  const fileInput = document.getElementById("fileInput");
  if (fileInput) fileInput.value = "";
  
  // Hide file info
  const fileInfoEl = document.getElementById('fileInfo');
  if (fileInfoEl) fileInfoEl.style.display = 'none';
  
  // Hide controls and clear button
  document.getElementById("dateRangeControls").style.display = "none";
  document.getElementById("clearButton").style.display = "none";
  
  // Reset to full data view
  document.getElementById('fullDataCharts').style.display = 'block';
  document.getElementById('fullDataMetrics').style.display = 'block';
  document.getElementById('comparisonCharts').style.display = 'none';
  document.getElementById('comparisonMetrics').style.display = 'none';
  
  // Destroy all charts
  const chartIds = ['dailyChart', 'percentChart', 'dailyChart1', 'percentChart1', 'dailyChart2', 'percentChart2'];
  chartIds.forEach(id => {
    const chartEl = document.getElementById(id);
    if (chartEl && chartEl.chart) {
      chartEl.chart.destroy();
      chartEl.chart = null;
    }
  });
  
  // Clear all display elements
  const elementsToReset = [
    "dateRange", "avgAirTime", "avgGroundTime", "totalAirTrans", "totalGroundTrans",
    "groundPercent", "airPercent", "smallVolTotal", "cartGroundTotal", 
    "cartAirPercent", "cartAirTotal", "total212"
  ];
  
  elementsToReset.forEach(id => {
    const el = document.getElementById(id);
    if (el) el.textContent = "";
    
    // Also clear comparison versions
    ['1', '2'].forEach(suffix => {
      const compEl = document.getElementById(id + suffix);
      if (compEl) compEl.textContent = "";
    });
  });
  
  document.getElementById("dateRange").textContent = "Date Range: Not loaded";
  console.log('Air vs Ground data cleared');
}

// ADDITIONAL HELPER FUNCTION FOR DETAILED CART ANALYSIS
function analyzeCartTransactions(data) {
  console.log('Starting detailed cart transaction analysis...');
  
  const cartAnalysis = {
    validCartTransactions: [],
    rejectedTransactions: [],
    volumeDistribution: {},
    topCartItems: {},
    employeeCartActivity: {}
  };
  
  const formatRegex = /^\d{2}-\d{2}-[A-Z0-9]{3}$/i;
  
  data.forEach((entry, index) => {
    const loc = (entry.location || "").trim().toUpperCase();
    
    // Only process MOD format transactions
    if (!formatRegex.test(loc)) {
      cartAnalysis.rejectedTransactions.push({
        index,
        reason: 'Invalid location format',
        location: loc
      });
      return;
    }
    
    // Parse item number
    const itemNumber = entry.item;
    if (!itemNumber) {
      cartAnalysis.rejectedTransactions.push({
        index,
        reason: 'Missing item number',
        entry
      });
      return;
    }
    
    const item = Math.floor(Number(itemNumber));
    if (isNaN(item) || !itemMasterData[item]) {
      cartAnalysis.rejectedTransactions.push({
        index,
        reason: 'Item not found in master data',
        itemNumber,
        convertedItem: item
      });
      return;
    }
    
    const quantity = Number(entry.quantity);
    if (!quantity || quantity <= 0) {
      cartAnalysis.rejectedTransactions.push({
        index,
        reason: 'Invalid quantity',
        quantity: entry.quantity
      });
      return;
    }
    
    const cubicVolume = Number(itemMasterData[item].CUBIC_VOL);
    if (isNaN(cubicVolume) || cubicVolume <= 0) {
      cartAnalysis.rejectedTransactions.push({
        index,
        reason: 'Invalid volume data',
        volume: itemMasterData[item].CUBIC_VOL
      });
      return;
    }
    
    const totalVolume = cubicVolume * quantity;
    
    // Cart transaction threshold
    if (totalVolume < 5000) {
      const suffix = loc.split("-")[2] || "";
      const isAir = !/[ABC]/i.test(suffix);
      
      cartAnalysis.validCartTransactions.push({
        index,
        itemNumber,
        quantity,
        cubicVolume,
        totalVolume,
        location: loc,
        suffix,
        isAir,
        employee: entry.employeeId || 'Unknown',
        date: entry.date
      });
      
      // Track volume distribution
      const volumeRange = getVolumeRange(totalVolume);
      cartAnalysis.volumeDistribution[volumeRange] = (cartAnalysis.volumeDistribution[volumeRange] || 0) + 1;
      
      // Track top items
      cartAnalysis.topCartItems[itemNumber] = (cartAnalysis.topCartItems[itemNumber] || 0) + 1;
      
      // Track employee activity
      const employee = entry.employeeId || 'Unknown';
      if (!cartAnalysis.employeeCartActivity[employee]) {
        cartAnalysis.employeeCartActivity[employee] = { air: 0, ground: 0 };
      }
      if (isAir) {
        cartAnalysis.employeeCartActivity[employee].air++;
      } else {
        cartAnalysis.employeeCartActivity[employee].ground++;
      }
    }
  });
  
  console.log('Cart Analysis Complete:', {
    validCartTransactions: cartAnalysis.validCartTransactions.length,
    rejectedTransactions: cartAnalysis.rejectedTransactions.length,
    volumeDistribution: cartAnalysis.volumeDistribution
  });
  
  return cartAnalysis;
}

function getVolumeRange(volume) {
  if (volume < 100) return '0-100';
  if (volume < 500) return '100-500';
  if (volume < 1000) return '500-1000';
  if (volume < 2500) return '1000-2500';
  return '2500-5000';
}

// PDF Download Report Function with comparison support
// Complete PDF Download Report Function with 3-column comparison support
// Complete PDF Download Report Function with 3-column comparison support
async function downloadReport() {
  if (transactionData.length === 0) {
    alert('No data available to generate report. Please upload a file first.');
    return;
  }

  try {
    const { jsPDF } = window.jspdf || await import('https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js');
    
    const doc = new jsPDF();
    const pageWidth = doc.internal.pageSize.width;
    const pageHeight = doc.internal.pageSize.height;
    const margin = 20;
    let yPosition = margin;

    const checkPageBreak = (spaceNeeded) => {
      if (yPosition + spaceNeeded > pageHeight - margin) {
        doc.addPage();
        yPosition = margin;
        return true;
      }
      return false;
    };

    // Simplified Header
    doc.setFontSize(20);
    doc.setFont(undefined, 'bold');
    doc.setTextColor(47, 62, 77);
    doc.text('Air vs Ground Analysis Report', pageWidth / 2, yPosition, { align: 'center' });
    yPosition += 15;

    const dateRangeEl = document.getElementById("dateRange");
    if (dateRangeEl && dateRangeEl.textContent !== "Date Range: Not loaded") {
      doc.setFontSize(10);
      doc.setFont(undefined, 'normal');
      doc.setTextColor(108, 117, 125);
      doc.text(dateRangeEl.textContent, pageWidth / 2, yPosition, { align: 'center' });
      yPosition += 20;
    } else {
      yPosition += 10;
    }

    if (isComparisonMode) {
      // Charts First - Enhanced Comparison Mode Report
      doc.setFontSize(16);
      doc.setFont(undefined, 'bold');
      doc.setTextColor(47, 62, 77);
      doc.text('Period Comparison Analysis', margin, yPosition);
      yPosition += 15;

      // Capture comparison charts
      const chartIds = ['dailyChart1', 'percentChart1', 'dailyChart2', 'percentChart2'];
      const chartTitles = ['Period 1 - Daily Trends', 'Period 1 - Percentages', 'Period 2 - Daily Trends', 'Period 2 - Percentages'];
      
      for (let i = 0; i < chartIds.length; i++) {
        const chartEl = document.getElementById(chartIds[i]);
        if (chartEl) {
          try {
            const canvas = await html2canvas(chartEl, {
              backgroundColor: '#ffffff',
              scale: 2,
              useCORS: true
            });
            
            const imgData = canvas.toDataURL('image/png');
            const chartWidth = 80;
            const chartHeight = 60;
            
            checkPageBreak(chartHeight + 20);
            
            doc.setFontSize(10);
            doc.setFont(undefined, 'bold');
            doc.setTextColor(i < 2 ? 30 : 229, i < 2 ? 136 : 53, i < 2 ? 229 : 53);
            doc.text(chartTitles[i], margin + (i % 2) * 90, yPosition);
            
            doc.addImage(imgData, 'PNG', margin + (i % 2) * 90, yPosition + 5, chartWidth, chartHeight);
            
            if (i % 2 === 1) yPosition += chartHeight + 20;
          } catch (error) {
            console.warn(`Could not capture chart ${chartIds[i]}:`, error);
          }
        }
      }

      // Performance Metrics Section - 3 Column Layout
      checkPageBreak(120);
      doc.setFontSize(16);
      doc.setFont(undefined, 'bold');
      doc.setTextColor(47, 62, 77);
      doc.text('Performance Metrics Comparison', margin, yPosition);
      yPosition += 15;

      const colWidth = (pageWidth - 2 * margin) / 3;
      const col1X = margin;
      const col2X = margin + colWidth;
      const col3X = margin + (2 * colWidth);

      // Period 1 column
      doc.setFontSize(14);
      doc.setFont(undefined, 'bold');
      doc.setTextColor(30, 136, 229);
      doc.text('Period 1 Baseline', col1X, yPosition);
      
      // Period 2 column
      doc.setFontSize(14);
      doc.setFont(undefined, 'bold');
      doc.setTextColor(229, 53, 53);
      doc.text('Period 2 Results', col2X, yPosition);
      
      // Change column
      doc.setFontSize(14);
      doc.setFont(undefined, 'bold');
      doc.setTextColor(76, 175, 80);
      doc.text('% Change', col3X, yPosition);
      
      yPosition += 15;

      // Get metrics for PDF (extract clean text only, avoiding HTML formatting)
      const getCleanMetrics = (suffix = '') => {
        const elementIds = [
          "totalAirTrans" + suffix,
          "totalGroundTrans" + suffix, 
          "avgAirTime" + suffix,
          "avgGroundTime" + suffix,
          "airPercent" + suffix,
          "groundPercent" + suffix,
          "cartAirPercent" + suffix,
          "smallVolTotal" + suffix
        ];
        
        return elementIds.map(id => {
          const el = document.getElementById(id);
          if (!el) return '';
          
          // Get text content only, which ignores HTML spans
          let text = el.textContent || el.innerText || '';
          
          // Remove any arrows and percentage changes that might be in the text
          text = text.replace(/[↗↘]/g, '').replace(/\s*[+\-]\d+\.\d+%/g, '').trim();
          
          return text;
        }).filter(item => item && item.length > 0);
      };

      const metrics1 = getCleanMetrics('1');
      const metrics2 = getCleanMetrics('2');
      
      // Calculate changes for PDF using the actual metrics data
      if (period1Data && period2Data) {
        const m1 = calculateMetricsForData(period1Data);
        const m2 = calculateMetricsForData(period2Data);
        
        const changes = [
          calculatePercentChange(m1.airCount, m2.airCount),
          calculatePercentChange(m1.groundCount, m2.groundCount),
          calculatePercentChange(parseFloat(m1.avgAir), parseFloat(m2.avgAir)),
          calculatePercentChange(parseFloat(m1.avgGround), parseFloat(m2.avgGround)),
          calculatePercentChange(parseFloat(m1.airPercent), parseFloat(m2.airPercent)),
          calculatePercentChange(parseFloat(m1.groundPercent), parseFloat(m2.groundPercent)),
          calculatePercentChange(parseFloat(m1.cartAirPercent), parseFloat(m2.cartAirPercent)),
          calculatePercentChange(m1.smallVolTotal, m2.smallVolTotal)
        ];

        doc.setFontSize(9);
        doc.setFont(undefined, 'normal');
        doc.setTextColor(0, 0, 0);
        
        const maxRows = Math.max(metrics1.length, metrics2.length, changes.length);
        for (let i = 0; i < maxRows; i++) {
          // Period 1 - plain black text
          if (metrics1[i]) {
            doc.setTextColor(0, 0, 0);
            doc.text(metrics1[i], col1X, yPosition);
          }
          
          // Period 2 - plain black text
          if (metrics2[i]) {
            doc.setTextColor(0, 0, 0);
            doc.text(metrics2[i], col2X, yPosition);
          }
          
          // Changes - color coded, clean format
          if (changes[i]) {
            const changeValue = parseFloat(changes[i]);
            if (!isNaN(changeValue)) {
              const changeColor = changeValue >= 0 ? [76, 175, 80] : [244, 67, 54]; // Green for positive, Red for negative
              doc.setTextColor(...changeColor);
              doc.text(changes[i] + '%', col3X, yPosition);
              doc.setTextColor(0, 0, 0); // Reset to black
            }
          }
          
          yPosition += 10;
        }
      }

    } else {
      // Single period report
      doc.setFontSize(16);
      doc.setFont(undefined, 'bold');
      doc.setTextColor(47, 62, 77);
      doc.text('Visual Analysis', margin, yPosition);
      yPosition += 15;

      // Capture main charts
      const chartIds = ['dailyChart', 'percentChart'];
      const chartTitles = ['Daily Transaction Trends', 'Percentage Distribution Over Time'];
      
      for (const [index, chartId] of chartIds.entries()) {
        const chartEl = document.getElementById(chartId);
        if (chartEl) {
          try {
            const canvas = await html2canvas(chartEl, {
              backgroundColor: '#ffffff',
              scale: 2,
              useCORS: true
            });
            
            const imgData = canvas.toDataURL('image/png');
            const chartWidth = 160;
            const chartHeight = 100;
            
            checkPageBreak(chartHeight + 20);
            
            doc.setFontSize(12);
            doc.setFont(undefined, 'bold');
            doc.text(chartTitles[index], margin, yPosition);
            yPosition += 10;
            
            doc.addImage(imgData, 'PNG', margin, yPosition, chartWidth, chartHeight);
            yPosition += chartHeight + 15;
            
          } catch (error) {
            console.warn(`Could not capture chart ${chartId}:`, error);
          }
        }
      }

      // Metrics Section for single period
      checkPageBreak(100);
      doc.setFontSize(16);
      doc.setFont(undefined, 'bold');
      doc.setTextColor(47, 62, 77);
      doc.text('Performance Metrics', margin, yPosition);
      yPosition += 15;

      const colWidth = (pageWidth - 2 * margin) / 3;
      const col1X = margin;
      const col2X = margin + colWidth;
      const col3X = margin + (2 * colWidth);

      // Get metrics from current display
      const getMetrics = () => ({
        volume: [
          document.getElementById("smallVolTotal")?.textContent || '',
          document.getElementById("total212")?.textContent || ''
        ].filter(item => item),
        air: [
          document.getElementById("airPercent")?.textContent || '',
          document.getElementById("cartAirPercent")?.textContent || '',
          document.getElementById("avgAirTime")?.textContent || '',
          document.getElementById("totalAirTrans")?.textContent || ''
        ].filter(item => item),
        ground: [
          document.getElementById("groundPercent")?.textContent || '',
          document.getElementById("cartGroundTotal")?.textContent || '',
          document.getElementById("avgGroundTime")?.textContent || '',
          document.getElementById("totalGroundTrans")?.textContent || ''
        ].filter(item => item)
      });

      const metrics = getMetrics();

      // Volume & Distribution
      doc.setFontSize(14);
      doc.setFont(undefined, 'bold');
      doc.setTextColor(142, 36, 170);
      doc.text('Volume & Distribution', col1X, yPosition);
      let col1Y = yPosition + 12;
      doc.setFontSize(10);
      doc.setFont(undefined, 'normal');
      doc.setTextColor(0, 0, 0);
      metrics.volume.forEach(item => {
        doc.text(item, col1X, col1Y);
        col1Y += 8;
      });

      // Air Metrics
      doc.setFontSize(14);
      doc.setFont(undefined, 'bold');
      doc.setTextColor(30, 136, 229);
      doc.text('Air Metrics', col2X, yPosition);
      let col2Y = yPosition + 12;
      doc.setFontSize(10);
      doc.setFont(undefined, 'normal');
      doc.setTextColor(0, 0, 0);
      metrics.air.forEach(item => {
        doc.text(item, col2X, col2Y);
        col2Y += 8;
      });

      // Ground Metrics
      doc.setFontSize(14);
      doc.setFont(undefined, 'bold');
      doc.setTextColor(67, 160, 71);
      doc.text('Ground Metrics', col3X, yPosition);
      let col3Y = yPosition + 12;
      doc.setFontSize(10);
      doc.setFont(undefined, 'normal');
      doc.setTextColor(0, 0, 0);
      metrics.ground.forEach(item => {
        doc.text(item, col3X, col3Y);
        col3Y += 8;
      });

      yPosition = Math.max(col1Y, col2Y, col3Y) + 20;
    }

    // Footer
    const pageCount = doc.internal.getNumberOfPages();
    for (let i = 1; i <= pageCount; i++) {
      doc.setPage(i);
      doc.setFontSize(8);
      doc.setFont(undefined, 'normal');
      doc.setTextColor(108, 117, 125);
      doc.text(`Air vs Ground Analysis Report - Page ${i} of ${pageCount}`, pageWidth / 2, pageHeight - 10, { align: 'center' });
      doc.text(`Generated by Excel Cart Analyzer v1.5`, pageWidth - margin, pageHeight - 10, { align: 'right' });
    }

    // Generate dynamic filename based on date ranges
    let fileName;
    if (isComparisonMode) {
      const period1Start = document.getElementById('period1Start').value;
      const period1End = document.getElementById('period1End').value;
      const period2Start = document.getElementById('period2Start').value;
      const period2End = document.getElementById('period2End').value;
      
      const formatDateForFile = (dateStr) => {
        const date = new Date(dateStr);
        return date.toISOString().split('T')[0]; // YYYY-MM-DD format
      };
      
      const p1Start = formatDateForFile(period1Start);
      const p1End = formatDateForFile(period1End);
      const p2Start = formatDateForFile(period2Start);
      const p2End = formatDateForFile(period2End);
      
      fileName = `air-vs-ground-comparison_${p1Start}-to-${p1End}_vs_${p2Start}-to-${p2End}.pdf`;
    } else {
      // For single analysis, get the date range from the data
      const dates = transactionData
        .map(t => new Date(t.date))
        .filter(d => !isNaN(d.getTime()))
        .sort((a, b) => a - b);
      
      if (dates.length > 0) {
        const startDate = dates[0].toISOString().split('T')[0];
        const endDate = dates[dates.length - 1].toISOString().split('T')[0];
        fileName = `air-vs-ground-analysis_${startDate}-to-${endDate}.pdf`;
      } else {
        fileName = 'air-vs-ground-analysis-report.pdf';
      }
    }

    doc.save(fileName);
    console.log('Enhanced Air vs Ground PDF report generated successfully:', fileName);
    
  } catch (error) {
    console.error('Error generating PDF report:', error);
    alert('Error generating PDF report. Please try again.');
  }
}

// Handle page load - check for refresh vs SPA navigation
function handleAirGroundPageLoad() {
  const wasLeaving = localStorage.getItem('spa_isLeaving');
  
  if (!wasLeaving) {
    // This was a refresh, clear Air vs Ground data
    console.log('Page refresh detected - clearing Air vs Ground data');
    clearAirGroundData();
  } else {
    // Normal SPA navigation, load saved data
    console.log('SPA navigation detected - loading saved Air vs Ground data');
    const saved = localStorage.getItem(AIR_GROUND_STORAGE_KEY);
    if (saved) {
      try {
        transactionData = JSON.parse(saved);
        resetToFullData();
        
        // Show controls and clear button
        document.getElementById("dateRangeControls").style.display = "block";
        document.getElementById("clearButton").style.display = "inline-block";
        
        // Set up date range inputs
        setupDateRangeInputs();
        
        console.log('Air vs Ground data restored from localStorage');
      } catch (e) {
        console.warn("Failed to restore Air vs Ground data from localStorage:", e);
        clearAirGroundData();
      }
    }
  }
  
  // Clear the leaving flag
  localStorage.removeItem('spa_isLeaving');
}

function displayPercentageChanges(metrics1, metrics2) {
  const calculatePercentChange = (oldValue, newValue) => {
    if (oldValue === 0 && newValue === 0) return "0.00";
    if (oldValue === 0) return "+∞";
    const change = ((newValue - oldValue) / oldValue) * 100;
    return (change >= 0 ? "+" : "") + change.toFixed(2);
  };

  // Extract numeric values from metrics
  const getValue = (metrics, key) => {
    switch(key) {
      case 'airCount': return metrics.airCount;
      case 'groundCount': return metrics.groundCount;
      case 'avgAir': return parseFloat(metrics.avgAir);
      case 'avgGround': return parseFloat(metrics.avgGround);
      case 'airPercent': return parseFloat(metrics.airPercent);
      case 'groundPercent': return parseFloat(metrics.groundPercent);
      case 'smallVolTotal': return metrics.smallVolTotal;
      case 'cartToAir': return metrics.cartToAir;
      case 'cartToGround': return metrics.cartToGround;
      case 'cartAirPercent': return parseFloat(metrics.cartAirPercent);
      case 'modTransactionCount': return metrics.modTransactionCount;
      default: return 0;
    }
  };

  // Calculate percentage changes for all metrics
  const changes = {
    airCount: calculatePercentChange(getValue(metrics1, 'airCount'), getValue(metrics2, 'airCount')),
    groundCount: calculatePercentChange(getValue(metrics1, 'groundCount'), getValue(metrics2, 'groundCount')),
    avgAir: calculatePercentChange(getValue(metrics1, 'avgAir'), getValue(metrics2, 'avgAir')),
    avgGround: calculatePercentChange(getValue(metrics1, 'avgGround'), getValue(metrics2, 'avgGround')),
    airPercent: calculatePercentChange(getValue(metrics1, 'airPercent'), getValue(metrics2, 'airPercent')),
    groundPercent: calculatePercentChange(getValue(metrics1, 'groundPercent'), getValue(metrics2, 'groundPercent')),
    smallVolTotal: calculatePercentChange(getValue(metrics1, 'smallVolTotal'), getValue(metrics2, 'smallVolTotal')),
    cartToAir: calculatePercentChange(getValue(metrics1, 'cartToAir'), getValue(metrics2, 'cartToAir')),
    cartToGround: calculatePercentChange(getValue(metrics1, 'cartToGround'), getValue(metrics2, 'cartToGround')),
    cartAirPercent: calculatePercentChange(getValue(metrics1, 'cartAirPercent'), getValue(metrics2, 'cartAirPercent')),
    modTransactionCount: calculatePercentChange(getValue(metrics1, 'modTransactionCount'), getValue(metrics2, 'modTransactionCount'))
  };

  // Update Period 2 display with percentage changes
  const updateElementWithChange = (id, baseText, changePercent) => {
    const el = document.getElementById(id + '2');
    if (el) {
      const changeColor = parseFloat(changePercent) >= 0 ? '#4caf50' : '#f44336';
      const changeSymbol = parseFloat(changePercent) >= 0 ? '↗' : '↘';
      el.innerHTML = `${baseText} <span style="color: ${changeColor}; font-weight: bold; margin-left: 8px;">${changeSymbol} ${changePercent}%</span>`;
    }
  };

  // Apply changes to Period 2 metrics
  updateElementWithChange("totalAirTrans", `Total Air Transactions: ${getValue(metrics2, 'airCount')}`, changes.airCount);
  updateElementWithChange("totalGroundTrans", `Total Ground Transactions: ${getValue(metrics2, 'groundCount')}`, changes.groundCount);
  updateElementWithChange("avgAirTime", `Avg Air Time (min): ${metrics2.avgAir}`, changes.avgAir);
  updateElementWithChange("avgGroundTime", `Avg Ground Time (min): ${metrics2.avgGround}`, changes.avgGround);
  updateElementWithChange("airPercent", `Air %: ${metrics2.airPercent}%`, changes.airPercent);
  updateElementWithChange("groundPercent", `Ground %: ${metrics2.groundPercent}%`, changes.groundPercent);
  updateElementWithChange("smallVolTotal", `Potential Cart Total: ${getValue(metrics2, 'smallVolTotal')}`, changes.smallVolTotal);
  updateElementWithChange("cartAirTotal", `Cart to Air Total: ${getValue(metrics2, 'cartToAir')}`, changes.cartToAir);
  updateElementWithChange("cartGroundTotal", `Cart to Ground Total: ${getValue(metrics2, 'cartToGround')}`, changes.cartToGround);
  updateElementWithChange("cartAirPercent", `Cart to Air %: ${metrics2.cartAirPercent}%`, changes.cartAirPercent);
  updateElementWithChange("total212", `Total MOD Transactions: ${getValue(metrics2, 'modTransactionCount')}`, changes.modTransactionCount);

  // Log the changes for debugging
  console.log('Percentage Changes from Period 1 to Period 2:', changes);
}

function calculatePercentChange(oldValue, newValue) {
  if (oldValue === 0 && newValue === 0) return "0.00";
  if (oldValue === 0) return "+∞";
  const change = ((newValue - oldValue) / oldValue) * 100;
  return (change >= 0 ? "+" : "") + change.toFixed(2);
}

// Load on page ready
window.addEventListener("DOMContentLoaded", handleAirGroundPageLoad);