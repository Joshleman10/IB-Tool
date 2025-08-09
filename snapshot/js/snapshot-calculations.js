// ============================================================================
// KPI CALCULATIONS
// snapshot-calculations.js
// ============================================================================

// ============================================================================
// ENHANCED LABOR KPI CALCULATION
// Replace the calculateLaborKPIs function in snapshot-calculations.js
// ============================================================================
function calculateKPIs() {
  if (!excelData && !laborData) {
    console.log('No data available for KPI calculation');
    return;
  }

  kpiResults = {
    labor: laborData ? calculateLaborKPIs(laborData) : null,
    excel: excelData ? excelData : null,
    combined: (laborData && excelData) ? calculateCombinedKPIs(laborData, excelData) : null
  };

  saveDataToStorage();
  
  displayKPIs();
  displayDetailedAnalysis();
  displayInsights();
}

function calculateLaborKPIs(data) {
  const kpis = {};
  
  if (data.departments && data.departments.length > 0) {
    // Find the inbound department (CRET is already separate in the data)
    const inboundDept = data.departments.find(dept => {
      const deptName = dept.name.toLowerCase();
      return deptName.includes('inbound');
    });
    
    if (!inboundDept) {
      console.warn('No Inbound department found');
      return { error: 'No Inbound department found' };
    }
    
    console.log('=== INBOUND LABOR ANALYSIS ===');
    console.log('Selected department:', inboundDept.name);
    console.log('Hours:', inboundDept.totalHours);
    console.log('Units:', inboundDept.totalUnits);
    console.log('Transactions:', inboundDept.totalTransactions);
    
    kpis.inboundDepartment = inboundDept;
    kpis.totalHours = inboundDept.totalHours;
    kpis.totalUnits = inboundDept.totalUnits;
    kpis.totalTransactions = inboundDept.totalTransactions;
    kpis.overallUPH = inboundDept.uph;
    kpis.overallTPH = inboundDept.tph;
    
    // Store the functions array from the original laborData
    kpis.functions = data.functions || [];
    
    // Filter for inbound-specific functions (excluding returns/CRET functions)
    const inboundFunctions = (data.functions || []).filter(func => {
      const deptName = func.department?.toLowerCase() || '';
      const funcName = func.name?.toLowerCase() || '';
      
      // Include inbound-related departments
      const isInboundDept = deptName.includes('receiving') || 
                           deptName.includes('putaway') || 
                           deptName.includes('vas') ||
                           deptName.includes('inbound');
      
      // Exclude returns/CRET functions
      const isReturnsFunction = funcName.includes('return') ||
                               funcName.includes('cret') ||
                               deptName.includes('customer returns');
      
      return isInboundDept && !isReturnsFunction;
    });
    
    kpis.inboundFunctions = inboundFunctions;
    
    console.log('Inbound functions found (excluding returns):', inboundFunctions.length);
    console.log('Functions with hours > 0:', inboundFunctions.filter(f => f.totalHours > 0).length);
    
    // Find inbound areas if available (excluding returns)
    const inboundAreas = data.areas?.filter(area => {
      const areaName = area.name?.toLowerCase() || '';
      const isInboundArea = INBOUND_AREA_KEYWORDS.some(keyword => 
        areaName.includes(keyword));
      const isReturnsArea = areaName.includes('return') || areaName.includes('cret');
      
      return isInboundArea && !isReturnsArea;
    }) || [];
    
    if (inboundAreas.length > 0) {
      kpis.inboundAreas = inboundAreas.map(area => ({
        ...area,
        hoursPercent: inboundDept.totalHours > 0 ? 
          ((area.totalHours / inboundDept.totalHours) * 100).toFixed(1) : 0
      }));
    }
    
    // Log final structure
    console.log('Final KPIs structure:', {
      hasInboundDept: !!kpis.inboundDepartment,
      deptName: kpis.inboundDepartment?.name,
      totalHours: kpis.totalHours,
      functionsCount: kpis.functions?.length || 0,
      inboundFunctionsCount: kpis.inboundFunctions?.length || 0,
      areasCount: kpis.inboundAreas?.length || 0
    });
  }
  
  return kpis;
}

function calculateExcelKPIs(data) {
  const kpis = {};
  
  const firstSheet = Object.keys(data)[0];
  if (firstSheet && data[firstSheet]) {
    const transactions = data[firstSheet];
    
    // Filter for Type 152 transactions only (Inbound Order Receipt - Put)
    const type152Transactions = transactions.filter(row => {
      const transactionType = row["Transaction Type"];
      return transactionType === 152 || transactionType === '152';
    });
    
    console.log('=== TYPE 152 ANALYSIS ===');
    console.log('Total rows in sheet:', transactions.length);
    console.log('Type 152 transactions found:', type152Transactions.length);
    
    // Calculate the 5 critical KPIs
    
    // 1. Total Transactions (Type 152 count)
    kpis.totalTransactions = type152Transactions.length;
    
    // 2. Total Volume (sum of Quantity for Type 152)
    kpis.totalVolume = type152Transactions.reduce((sum, row) => {
      const quantity = parseFloat(row["Quantity"]) || 0;
      return sum + quantity;
    }, 0);
    
    // 3. UPT (Units Per Transaction) = Total Volume / Total Transactions
    kpis.UPT = kpis.totalTransactions > 0 ? kpis.totalVolume / kpis.totalTransactions : 0;
    
    // Store Type 152 specific data for combined calculations
    kpis.type152Data = {
      transactions: kpis.totalTransactions,
      volume: kpis.totalVolume,
      upt: kpis.UPT
    };
    
    // Additional summary for debugging/validation
    kpis.dataSummary = {
      totalRecords: transactions.length,
      type152Count: type152Transactions.length,
      type152Percentage: transactions.length > 0 ? 
        ((type152Transactions.length / transactions.length) * 100).toFixed(1) : 0
    };
    
    console.log('=== CALCULATED KPIs ===');
    console.log('Total Transactions (152):', kpis.totalTransactions);
    console.log('Total Volume (152):', kpis.totalVolume);
    console.log('UPT:', kpis.UPT.toFixed(2));
    
    // Log some sample quantities to verify calculation
    const sampleQuantities = type152Transactions.slice(0, 5).map(row => ({
      transType: row["Transaction Type"],
      quantity: row["Quantity"],
      parsed: parseFloat(row["Quantity"]) || 0
    }));
    console.log('Sample quantities:', sampleQuantities);
  }
  
  return kpis;
}

function calculateCombinedKPIs(laborData, excelData) {
  const combined = {};
  
  // Get Inbound department data (CRET is already separate in the data structure)
  const inboundDept = laborData.departments?.find(dept => 
    dept.name.toLowerCase().includes('inbound'));
  
  if (!inboundDept) {
    console.warn('No Inbound department found');
    return { error: 'No Inbound department found' };
  }
  
  console.log('=== INBOUND LABOR DATA ===');
  console.log('Department:', inboundDept.name);
  console.log('Total Hours:', inboundDept.totalHours);
  
  // Store core data
  combined.inboundDepartment = {
    name: inboundDept.name,
    totalHours: inboundDept.totalHours,
    laborUPH: inboundDept.uph,
    laborTPH: inboundDept.tph
  };
  
  // Calculate the 5 critical KPIs using Type 152 data
  if (excelData && excelData.type152Data) {
    const type152 = excelData.type152Data;
    const laborHours = inboundDept.totalHours;
    
    console.log('=== TYPE 152 DATA ===');
    console.log('Transactions:', type152.transactions);
    console.log('Volume:', type152.volume);
    console.log('Labor Hours:', laborHours);
    
    // Calculate the 5 critical KPIs
    combined.criticalKPIs = {
      // 1. Total Volume (from Type 152)
      totalVolume: type152.volume,
      
      // 2. Total Transactions (from Type 152)
      totalTransactions: type152.transactions,
      
      // 3. TPH = Total Volume / Total Labor Hours (IB excluding CRET)
      TPH: laborHours > 0 ? type152.volume / laborHours : 0,
      
      // 4. TPLH = Total Transactions / Total Labor Hours (IB excluding CRET)  
      TPLH: laborHours > 0 ? type152.transactions / laborHours : 0,
      
      // 5. UPT = Total Volume / Total Transactions
      UPT: type152.transactions > 0 ? type152.volume / type152.transactions : 0
    };
    
    console.log('=== CRITICAL KPIs CALCULATED ===');
    console.log('Total Volume:', combined.criticalKPIs.totalVolume);
    console.log('Total Transactions:', combined.criticalKPIs.totalTransactions);
    console.log('TPH:', combined.criticalKPIs.TPH.toFixed(2));
    console.log('TPLH:', combined.criticalKPIs.TPLH.toFixed(2));
    console.log('UPT:', combined.criticalKPIs.UPT.toFixed(2));
  } else {
    console.warn('No Type 152 data available for KPI calculation');
  }
  
  return combined;
}

function calculateEfficiencyInsights(inboundDept, excelData, type152Count) {
  const insights = [];
  
  if (type152Count > 0) {
    const actualTPLH152 = type152Count / inboundDept.totalHours;
    insights.push({
      type: 'receiving_put',
      metric: 'Type 152 TPLH (Inbound Put)',
      laborHours: inboundDept.totalHours.toFixed(1),
      transactions: type152Count,
      actualTPLH: actualTPLH152.toFixed(2),
      laborTPH: inboundDept.tph.toFixed(1),
      variance: inboundDept.tph > 0 ? 
        (((actualTPLH152 - inboundDept.tph) / inboundDept.tph) * 100).toFixed(1) + '%' : 'N/A'
    });
  }
  
  // Check for 151 vs 152 ratio using the summary data
  if (excelData.inboundTransactionSummary) {
    const type151Count = excelData.inboundTransactionSummary.type151Count;
    const type152Count = excelData.inboundTransactionSummary.type152Count;
    
    if (type151Count > 0 && type152Count > 0) {
      const ratio = (type152Count / type151Count).toFixed(2);
      const isBalanced = Math.abs(type152Count - type151Count) / 
        Math.max(type151Count, type152Count) < 0.05;
      
      insights.push({
        type: 'transaction_ratio',
        metric: 'Receipt vs Put Ratio',
        type151Count: type151Count,
        type152Count: type152Count,
        ratio: ratio,
        status: isBalanced ? 'balanced' : 'imbalanced'
      });
    }
  }
  
  return insights;
}

function safecalculateKPIs() {
  if (!excelData && !laborData) {
    console.log('No data available for KPI calculation');
    return;
  }

  kpiResults = {
    labor: laborData ? calculateLaborKPIs(laborData) : null,
    excel: excelData ? excelData : null,
    combined: null
  };

  if (laborData && excelData) {
    kpiResults.combined = calculateCombinedKPIs(laborData, excelData);
  }

  saveDataToStorage();
  displayKPIs();
  displayDetailedAnalysis();
  displayInsights();

  console.log('=== FINAL KPI RESULTS ===');
  console.log('KPI Results:', kpiResults);
}

// 🔽 This is CRITICAL
window.safeCalculateKPIs = safecalculateKPIs;
