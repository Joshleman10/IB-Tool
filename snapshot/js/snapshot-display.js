function displayKPIs() {
  const kpiGrid = document.getElementById('kpiGrid');
  const kpiDashboard = document.getElementById('kpiDashboard');
  
  if (!kpiResults || !kpiGrid) return;
  
  kpiDashboard.style.display = 'block';
  
  let kpiCards = [];
  
  // Check for errors first
  if (kpiResults.combined && kpiResults.combined.error) {
    kpiGrid.innerHTML = `
      <div class="kpi-card error">
        <div class="kpi-title">⚠️ Data Error</div>
        <div class="kpi-value">${kpiResults.combined.error}</div>
        <div class="kpi-subtitle">Check data format</div>
      </div>
    `;
    return;
  }
  
  // Display the 6 Critical KPIs (Total IB Hours + 5 performance metrics)
  if (kpiResults.combined && kpiResults.combined.criticalKPIs && kpiResults.combined.inboundDepartment) {
    const kpis = kpiResults.combined.criticalKPIs;
    const inboundDept = kpiResults.combined.inboundDepartment;
    
    // Create the 6 priority KPI cards
    kpiCards = [
      // 1. Total IB Hours (FIRST - foundational metric)
      createFocusedKPICard(
        'Total IB Hours', 
        inboundDept.totalHours.toFixed(2), 
        'inbound labor hours',
        'labor'
      ),
      
      // 2. Total Volume
      createFocusedKPICard(
        'Total Volume', 
        kpis.totalVolume.toLocaleString(), 
        'units from type 152 transactions',
        'productivity'
      ),
      
      // 3. Total Transactions  
      createFocusedKPICard(
        'Total Transactions', 
        kpis.totalTransactions.toLocaleString(), 
        'type 152 put transactions',
        'combined'
      ),
      
      // 4. TPH
      createFocusedKPICard(
        'TPH', 
        kpis.TPH.toFixed(1), 
        'units per labor hour',
        'efficiency'
      ),
      
      // 5. TPLH
      createFocusedKPICard(
        'TPLH', 
        kpis.TPLH.toFixed(2), 
        'transactions per labor hour',
        'data'
      ),
      
      // 6. UPT
      createFocusedKPICard(
        'UPT', 
        kpis.UPT.toFixed(1), 
        'units per transaction',
        'productivity'
      )
    ];
  } else {
    // Show data loading status
    kpiCards = [
      createFocusedKPICard(
        'Status', 
        'Need Both Datasets', 
        'upload excel & paste labor data',
        'data'
      )
    ];
  }
  
  kpiGrid.innerHTML = kpiCards.join('');
  
  // Show analysis sections
  document.getElementById('analysisSection').style.display = 'block';
  document.getElementById('actionButtons').style.display = 'block';
}

function createFocusedKPICard(title, value, subtitle, type) {
  return `
    <div class="kpi-card ${type}">
      <div class="kpi-title">${title}</div>
      <div class="kpi-value">${value}</div>
      <div class="kpi-subtitle">${subtitle}</div>
    </div>
  `;
}

// Helper functions for KPI styling
function getVarianceTypeClass(variance) {
  const absVariance = Math.abs(variance);
  if (absVariance <= 5) return 'efficiency'; // Good variance
  if (absVariance <= 15) return 'data'; // Acceptable variance
  return 'combined'; // High variance
}

function getRatioTypeClass(ratio) {
  if (ratio >= 0.95 && ratio <= 1.05) return 'efficiency'; // Balanced
  if (ratio >= 0.8 && ratio <= 1.2) return 'productivity'; // Reasonable
  return 'combined'; // Imbalanced
}

function createKPICard(title, value, subtitle, type) {
  return `
    <div class="kpi-card ${type}">
      <div class="kpi-title">${title}</div>
      <div class="kpi-value">${value}</div>
      <div class="kpi-subtitle">${subtitle}</div>
    </div>
  `;
}

function createKPICard(title, value, subtitle, type) {
  return `
    <div class="kpi-card ${type}">
      <div class="kpi-title">${title}</div>
      <div class="kpi-value">${value}</div>
      <div class="kpi-subtitle">${subtitle}</div>
    </div>
  `;
}

function displayDetailedAnalysis() {
  if (kpiResults.labor) {
    displayLaborAnalysis();
  }
  if (kpiResults.excel) {
    displayProductivityAnalysis();
    displayEfficiencyAnalysis();
  }
  if (kpiResults.combined) {
    displayCombinedAnalysis();
  }
}

// Purpose: Refactor the Labor Analysis tab to display direct vs indirect labor breakdown based on specific mappings

// ============================================================================
// ENHANCED LABOR ANALYSIS DISPLAY
// Replaces the existing displayLaborAnalysis function in snapshot-display.js
// ============================================================================

// ============================================================================
// DEBUG VERSION - Enhanced Labor Analysis Display
// Replace your displayLaborAnalysis function with this debug version
// ============================================================================

// ============================================================================
// DEBUG VERSION - Enhanced Labor Analysis Display
// Replace your displayLaborAnalysis function with this debug version
// ============================================================================

function displayLaborAnalysis() {
  const laborBreakdown = document.getElementById('laborBreakdown');
  if (!laborBreakdown || !kpiResults.labor) return;

  // Get functions data
  let allFunctions = [];
  if (kpiResults.labor.functions && kpiResults.labor.functions.length > 0) {
    allFunctions = kpiResults.labor.functions;
  } else if (laborData && laborData.functions && laborData.functions.length > 0) {
    allFunctions = laborData.functions;
  } else if (kpiResults.labor.inboundFunctions && kpiResults.labor.inboundFunctions.length > 0) {
    allFunctions = kpiResults.labor.inboundFunctions;
  }
  
  if (allFunctions.length === 0) {
    laborBreakdown.innerHTML = `
      <div class="error-message" style="text-align: center; padding: 2rem; background: #f8d7da; border: 1px solid #f5c6cb; border-radius: 8px; color: #721c24;">
        <h3>⚠️ No Function Data Available</h3>
        <p>Unable to load labor function details. Please verify the labor data format.</p>
      </div>
    `;
    return;
  }

  console.log('=== REFINED LABOR ANALYSIS ===');
  console.log('Available functions:', allFunctions.length);
  allFunctions.forEach(func => {
    if (func.totalHours > 0) {
      console.log(`"${func.name}": ${func.totalHours}h, ${func.totalUnits} units, ${func.totalTransactions} trans`);
    }
  });

  // DIRECT LABOR CATEGORIES - Your Specific Requirements
  const directLaborCategories = {
    reachTruckPutaway: [
      'Reach-Truck Putaway'
    ],
    cartPutaway: [
      'Non- PIT Manual Putaway'  // You want this tracked as "Cart Putaways"
    ],
    putawayUnknown: [
      'Putaway - Unknown'  // Keep separate as requested
    ],
    receiving: [
      'Break Down Receiving',
      'Cart Receiving',
      'Container Receiver',
      'Full Pallet Receiving', 
      'Parcel Receiving',
      'Receiving',
      'Small Pack Debundle'  // Moving this to direct as it's operational work
    ]
  };

  // INDIRECT LABOR CATEGORIES - Your Specific Requirements
  const indirectLaborCategories = {
    receivingSupport: [
      'Unloader',
      'Container Unload',  // Added "Container" functions
      'Breakdown Receiver',
      'Pallet Rework',
      'Pallet Wrangler – Dock Stocker IB',
      'Pallet Wrapper',  // Added as requested
      'Pallet Loader'    // Added as requested
    ],
    leadership: [
      'Inbound Lead',
      'Problem Solver',    // Added as requested
      'Vendor Compliance', // Added as requested
      'Inbound Training'
    ],
    vasIndirect: [
      'VAS',              // All VAS functions as requested
      'VAS Execute',
      'Heat Shrink',
      'DeBundling'
    ],
    cartPutawayIndirect: [
      'Cart Putaway'      // Added as requested
    ],
    unallocated: [
      'On-Clock Unallocated'
    ]
  };

  // CUSTOMER RETURNS CATEGORIES - Separate tracking
  const customerReturnsCategories = {
    customerReturns: [
      'Customer Returns Unloader',
      'Inbound Returns Processor',
      'Customer Returns - Ground',
      'Customer Returns Problem Solver'
    ]
  };

  // Helper function to sum metrics for given function names
  function sumFunctionMetrics(functionNames, categoryName = '') {
    const result = functionNames.reduce((acc, name) => {
      const match = allFunctions.find(f => f.name === name || f.name.includes(name));
      if (match && match.totalHours > 0) {
        console.log(`✓ Found ${categoryName}: ${name} = ${match.totalHours}h`);
        acc.hours += match.totalHours || 0;
        acc.units += match.totalUnits || 0;
        acc.transactions += match.totalTransactions || 0;
        acc.functions.push({
          name: match.name,
          hours: match.totalHours || 0,
          units: match.totalUnits || 0,
          transactions: match.totalTransactions || 0,
          uph: match.uph || 0,
          tph: match.tph || 0
        });
      } else {
        console.log(`✗ Function NOT found: ${name}`);
      }
      return acc;
    }, { hours: 0, units: 0, transactions: 0, functions: [] });
    
    console.log(`${categoryName} Total:`, {
      hours: result.hours,
      functions: result.functions.length
    });
    
    return result;
  }

  // Calculate Direct Labor - Your Specific Buckets
  console.log('=== CALCULATING DIRECT LABOR ===');
  const reachTruckPutaway = sumFunctionMetrics(directLaborCategories.reachTruckPutaway, 'Reach Truck Putaway');
  const cartPutaway = sumFunctionMetrics(directLaborCategories.cartPutaway, 'Cart Putaway');
  const putawayUnknown = sumFunctionMetrics(directLaborCategories.putawayUnknown, 'Putaway Unknown');
  const receivingDirect = sumFunctionMetrics(directLaborCategories.receiving, 'Direct Receiving');

  // Calculate Indirect Labor - Your Specific Requirements
  console.log('=== CALCULATING INDIRECT LABOR ===');
  const receivingSupport = sumFunctionMetrics(indirectLaborCategories.receivingSupport, 'Receiving Support');
  const leadership = sumFunctionMetrics(indirectLaborCategories.leadership, 'Leadership');
  const vasIndirect = sumFunctionMetrics(indirectLaborCategories.vasIndirect, 'VAS Functions');
  const cartPutawayIndirect = sumFunctionMetrics(indirectLaborCategories.cartPutawayIndirect, 'Cart Putaway Indirect');
  const unallocated = sumFunctionMetrics(indirectLaborCategories.unallocated, 'Unallocated');

  // Calculate Customer Returns - Separate Section
  console.log('=== CALCULATING CUSTOMER RETURNS ===');
  const customerReturns = sumFunctionMetrics(customerReturnsCategories.customerReturns, 'Customer Returns');

  // Calculate totals (Customer Returns separate from IB operations)
  const totalDirectHours = reachTruckPutaway.hours + cartPutaway.hours + putawayUnknown.hours + receivingDirect.hours;
  const totalIndirectHours = receivingSupport.hours + leadership.hours + vasIndirect.hours + cartPutawayIndirect.hours + unallocated.hours;
  const totalCRETHours = customerReturns.hours;
  const totalDirectUnits = reachTruckPutaway.units + cartPutaway.units + putawayUnknown.units + receivingDirect.units;
  const totalDirectTransactions = reachTruckPutaway.transactions + cartPutaway.transactions + putawayUnknown.transactions + receivingDirect.transactions;

  // Get total IB hours for percentage calculation (excluding CRET)
  const totalIBHours = kpiResults.labor.inboundDepartment?.totalHours || 0;
  const directPercentage = totalIBHours > 0 ? ((totalDirectHours / totalIBHours) * 100).toFixed(1) : 0;
  const indirectPercentage = totalIBHours > 0 ? ((totalIndirectHours / totalIBHours) * 100).toFixed(1) : 0;
  const accountedHours = totalDirectHours + totalIndirectHours;
  const coveragePercentage = totalIBHours > 0 ? ((accountedHours / totalIBHours) * 100).toFixed(1) : 0;

  // Customer Returns calculations
  const cretTPH = totalCRETHours > 0 ? customerReturns.transactions / totalCRETHours : 0;

  // Calculate rates
  const overallDirectUPH = totalDirectHours > 0 ? totalDirectUnits / totalDirectHours : 0;
  const overallDirectTPH = totalDirectHours > 0 ? totalDirectTransactions / totalDirectHours : 0;

  console.log('=== FINAL TOTALS ===');
  console.log('Total Direct Hours:', totalDirectHours);
  console.log('Total Indirect Hours:', totalIndirectHours);
  console.log('Total CRET Hours:', totalCRETHours);
  console.log('Coverage:', coveragePercentage + '%');

  // Check for uncategorized functions
  const categorizedFunctionNames = [
    ...directLaborCategories.reachTruckPutaway,
    ...directLaborCategories.cartPutaway,
    ...directLaborCategories.putawayUnknown,
    ...directLaborCategories.receiving,
    ...indirectLaborCategories.receivingSupport,
    ...indirectLaborCategories.leadership,
    ...indirectLaborCategories.vasIndirect,
    ...indirectLaborCategories.cartPutawayIndirect,
    ...indirectLaborCategories.unallocated,
    ...customerReturnsCategories.customerReturns
  ];

  const uncategorizedFunctions = allFunctions.filter(func => 
    func.totalHours > 0 && 
    !categorizedFunctionNames.some(catName => 
      func.name === catName || func.name.includes(catName)
    )
  );

  if (uncategorizedFunctions.length > 0) {
    console.log('⚠️ UNCATEGORIZED FUNCTIONS:');
    uncategorizedFunctions.forEach(func => {
      console.log(`"${func.name}": ${func.totalHours}h - NEEDS CATEGORIZATION`);
    });
  }

  // Helper function to create function breakdown HTML
  function createFunctionBreakdown(categoryData, categoryName, categoryClass, showDetails = true) {
    if (categoryData.functions.length === 0) return '';
    
    return `
      <div class="breakdown-item ${categoryClass}">
        <div class="breakdown-header">
          <div class="breakdown-name">${categoryName}</div>
          <div class="breakdown-hours">${categoryData.hours.toFixed(1)}h</div>
        </div>
        <div class="breakdown-metrics">
          <div class="metric">
            <div class="metric-label">Units</div>
            <div class="metric-value">${categoryData.units.toLocaleString()}</div>
          </div>
          <div class="metric">
            <div class="metric-label">Transactions</div>
            <div class="metric-value">${categoryData.transactions.toLocaleString()}</div>
          </div>
          <div class="metric">
            <div class="metric-label">UPH</div>
            <div class="metric-value">${categoryData.hours > 0 ? (categoryData.units / categoryData.hours).toFixed(1) : '0.0'}</div>
          </div>
          <div class="metric">
            <div class="metric-label">TPH</div>
            <div class="metric-value">${categoryData.hours > 0 ? (categoryData.transactions / categoryData.hours).toFixed(1) : '0.0'}</div>
          </div>
        </div>
        ${showDetails ? `
        <div class="function-details">
          ${categoryData.functions.map(func => `
            <div class="function-row">
              <span class="function-name">${func.name}</span>
              <span class="function-hours">${func.hours.toFixed(1)}h</span>
              <span class="function-metrics">${func.units.toLocaleString()}u | ${func.transactions.toLocaleString()}t</span>
            </div>
          `).join('')}
        </div>
        ` : ''}
      </div>
    `;
  }

  // Build the HTML
  let html = `
    <div class="labor-analysis-container">
      
      <!-- Key Metrics Overview - 3 Equal Sections Side by Side -->
      <div class="labor-summary" style="grid-template-columns: 1fr 1fr 1fr;">
        <div class="summary-card direct">
          <h4>📊 Direct Labor</h4>
          <div class="summary-metrics">
            <div class="metric">
              <span class="metric-label">Hours:</span>
              <span class="metric-value">${totalDirectHours.toFixed(1)}</span>
            </div>
            <div class="metric">
              <span class="metric-label">% of Total:</span>
              <span class="metric-value">${directPercentage}%</span>
            </div>
            <div class="metric">
              <span class="metric-label">UPH:</span>
              <span class="metric-value">${overallDirectUPH.toFixed(1)}</span>
            </div>
            <div class="metric">
              <span class="metric-label">TPH:</span>
              <span class="metric-value">${overallDirectTPH.toFixed(1)}</span>
            </div>
          </div>
        </div>
        
        <div class="summary-card indirect">
          <h4>⚙️ Indirect Labor</h4>
          <div class="summary-metrics">
            <div class="metric">
              <span class="metric-label">Hours:</span>
              <span class="metric-value">${totalIndirectHours.toFixed(1)}</span>
            </div>
            <div class="metric">
              <span class="metric-label">% of Total:</span>
              <span class="metric-value">${indirectPercentage}%</span>
            </div>
            <div class="metric">
              <span class="metric-label">Support Ratio:</span>
              <span class="metric-value">${totalDirectHours > 0 ? (totalIndirectHours / totalDirectHours).toFixed(2) : '0.00'}</span>
            </div>
          </div>
        </div>
        
        <div class="summary-card" style="border-left-color: #28a745;">
          <h4>🎯 Coverage Analysis</h4>
          <div class="summary-metrics">
            <div class="metric">
              <span class="metric-label">Total IB Hours:</span>
              <span class="metric-value">${totalIBHours.toFixed(1)}</span>
            </div>
            <div class="metric">
              <span class="metric-label">Categorized Hours:</span>
              <span class="metric-value">${accountedHours.toFixed(1)}</span>
            </div>
            <div class="metric">
              <span class="metric-label">Coverage:</span>
              <span class="metric-value">${coveragePercentage}%</span>
            </div>
          </div>
        </div>
      </div>
      
      <!-- Direct Labor Operations - Your Specific Buckets -->
      <div class="labor-section">
        <h3>📊 Direct Labor Operations</h3>
        <div class="breakdown-grid">
          ${createFunctionBreakdown(reachTruckPutaway, '🚛 Reach Truck Putaway', 'direct')}
          ${createFunctionBreakdown(cartPutaway, '🛒 Cart Putaway', 'direct')}
          ${putawayUnknown.hours > 0 ? createFunctionBreakdown(putawayUnknown, '❓ Putaway - Unknown', 'direct') : ''}
          ${createFunctionBreakdown(receivingDirect, '📥 Receiving Operations', 'direct')}
        </div>
      </div>
      
      <!-- Indirect Labor Support - Your Specific Categories -->
      <div class="labor-section">
        <h3>⚙️ Indirect Labor & Support</h3>
        <div class="breakdown-grid">
          ${createFunctionBreakdown(receivingSupport, '🔧 Receiving Support', 'indirect')}
          ${createFunctionBreakdown(leadership, '👥 Leadership & Problem Solving', 'indirect')}
          ${vasIndirect.hours > 0 ? createFunctionBreakdown(vasIndirect, '🛠️ VAS Functions', 'vas-special') : ''}
          ${cartPutawayIndirect.hours > 0 ? createFunctionBreakdown(cartPutawayIndirect, '🛒 Cart Putaway Support', 'indirect') : ''}
          ${unallocated.hours > 0 ? createFunctionBreakdown(unallocated, '❓ Unallocated Time', 'indirect', false) : ''}
        </div>
      </div>
      
      ${totalCRETHours > 0 ? `
      <!-- Customer Returns Section - Separate from IB Operations -->
      <div class="labor-section" style="border-left-color: #ffc107;">
        <h3>🔄 Customer Returns (CRET)</h3>
        <div class="labor-summary" style="margin-bottom: 1rem;">
        </div>
        <div class="breakdown-grid">
          ${createFunctionBreakdown(customerReturns, '🔄 Customer Returns Operations', 'returns')}
        </div>
      </div>
      ` : ''}
      
      ${uncategorizedFunctions.length > 0 ? `
      <!-- Uncategorized Functions Alert -->
      <div class="labor-section" style="border-left-color: #dc3545;">
        <h3>⚠️ Uncategorized Functions</h3>
        <div class="alert-box" style="background: #f8d7da; padding: 1rem; border-radius: 8px; color: #721c24;">
          <p><strong>The following functions need categorization:</strong></p>
          <ul style="margin: 0.5rem 0;">
            ${uncategorizedFunctions.map(func => 
              `<li>"${func.name}" - ${func.totalHours.toFixed(1)} hours</li>`
            ).join('')}
          </ul>
          <p style="margin-top: 1rem;"><em>Please let me know how to categorize these functions.</em></p>
        </div>
      </div>
      ` : ''}
      
    </div>
  `;

  laborBreakdown.innerHTML = html;
}

// ============================================================================
// VAS/CART/CRET TAB IMPLEMENTATION
// Replace displayProductivityAnalysis() function in snapshot-display.js
// ============================================================================

// VAS Target UPH Configuration
const VAS_TARGET_UPH = {
  'SM.SHRNKWRP': 150,
  'LG.SHRNKWRP': 115,
  'IBCASEPACK': 180,
  'IBREMOUCA': 180,
  'IBTAPEBAG': 60,
  'IBCOVERCSUPC': 210,
  'IBLARGETAPE': 90,
  'IBTAPETOP': 100,
  'IBINSPECT': 60,
  'IB-QA-INSPECT': 60,
  'IBITEMLABEL': 50,
  'IBBUBBLEWRAP': 40,
  'IBASSEMBLE': 48,
  'IBPLASTIC': 120
};

// Performance thresholds
const PERFORMANCE_THRESHOLDS_CONFIG = {
  excellent: 105, // >=105% of target
  acceptable: 95, // 95-105% of target
  poor: 95       // <95% of target
};

async function displayProductivityAnalysis() {
    const productivityMetrics = document.getElementById('productivityMetrics');
    if (!productivityMetrics) {
        console.error('productivityMetrics container not found!');
        return;
    }
    
    let html = '<div class="vas-cart-cret-container">';
    // START DIRECTLY with VAS Analysis (your enhanced version)
    html += `
        <div class="vas-analysis-section">
            <h3>🛠️ VAS Detailed Analysis</h3>
            <div id="vasAnalysisContent">Loading VAS analysis...</div>
        </div>
    `;
    
    // CART Analysis section
    html += `
        <div class="cart-analysis-section">
            <h3>🛒 CART Operations</h3>
            <div id="cartAnalysisContent">Loading CART analysis...</div>
        </div>
    `;
    
    // CRET Analysis section
    html += `
        <div class="cret-analysis-section">
            <h3>🔄 Customer Returns (CRET)</h3>
            ${generateCRETAnalysis()}
        </div>
    `;
    
    // Performance Key section
    html += `
        <div class="performance-key-section">
            <h3>📋 Performance Key & Definitions</h3>
            ${generatePerformanceKey()}
        </div>
    `;
    
    html += '</div>';
    productivityMetrics.innerHTML = html;
    
    // Load VAS and CART analysis asynchronously (your enhanced versions)
    try {
        const vasContent = await generateVASAnalysis(); // Your enhanced VAS with pie chart
        const vasContainer = document.getElementById('vasAnalysisContent');
        if (vasContainer) {
            vasContainer.innerHTML = vasContent;
        }
    } catch (error) {
        console.error('Error loading VAS analysis:', error);
        const vasContainer = document.getElementById('vasAnalysisContent');
        if (vasContainer) {
            vasContainer.innerHTML = '<p>Error loading VAS analysis. Please try refreshing.</p>';
        }
    }
    
    try {
        const cartContent = await generateCARTAnalysis();
        const cartContainer = document.getElementById('cartAnalysisContent');
        if (cartContainer) {
            cartContainer.innerHTML = cartContent;
        }
    } catch (error) {
        console.error('Error loading CART analysis:', error);
        const cartContainer = document.getElementById('cartAnalysisContent');
        if (cartContainer) {
            cartContainer.innerHTML = '<p>Error loading CART analysis. Please try refreshing.</p>';
        }
    }
}

// ============================================================================
// SUMMARY CARDS GENERATION
// ============================================================================
function generateSummaryCards() {
  let summaryHtml = '';
  
  // Get labor summary data
  if (kpiResults.labor) {
    const laborData = calculateLaborSummaryForVASTab();
    
    summaryHtml += `
      <div class="summary-card direct">
        <h4>📊 Direct Labor</h4>
        <div class="summary-metrics">
          <div class="metric">
            <span class="metric-label">Hours:</span>
            <span class="metric-value">${laborData.directHours.toFixed(1)}</span>
          </div>
          <div class="metric">
            <span class="metric-label">% of Total:</span>
            <span class="metric-value">${laborData.directPercentage}%</span>
          </div>
        </div>
      </div>
      
      <div class="summary-card indirect">
        <h4>⚙️ Indirect Labor</h4>
        <div class="summary-metrics">
          <div class="metric">
            <span class="metric-label">Hours:</span>
            <span class="metric-value">${laborData.indirectHours.toFixed(1)}</span>
          </div>
          <div class="metric">
            <span class="metric-label">% of Total:</span>
            <span class="metric-value">${laborData.indirectPercentage}%</span>
          </div>
        </div>
      </div>
    `;
  }
  
  // Add VAS/CART/CRET summary
  if (kpiResults.excel) {
    const operationalData = calculateOperationalSummary();
    
    summaryHtml += `
      <div class="summary-card vas-special">
        <h4>🛠️ VAS Operations</h4>
        <div class="summary-metrics">
          <div class="metric">
            <span class="metric-label">Total Transactions:</span>
            <span class="metric-value">${operationalData.vasTransactions}</span>
          </div>
          <div class="metric">
            <span class="metric-label">Total Units:</span>
            <span class="metric-value">${operationalData.vasUnits.toLocaleString()}</span>
          </div>
        </div>
      </div>
      
      <div class="summary-card cart">
        <h4>🛒 CART Operations</h4>
        <div class="summary-metrics">
          <div class="metric">
            <span class="metric-label">Total Transactions:</span>
            <span class="metric-value">${operationalData.cartTransactions}</span>
          </div>
          <div class="metric">
            <span class="metric-label">Total Units:</span>
            <span class="metric-value">${operationalData.cartUnits.toLocaleString()}</span>
          </div>
        </div>
      </div>
    `;
  }
  
  return summaryHtml;
}

// ============================================================================
// VAS ANALYSIS GENERATION
// ============================================================================
async function generateVASAnalysis() {
    console.log('🔍 Generating enhanced VAS analysis...');
    
    // Try to get Excel data using the async version that can prompt for file reselection
    const excelData = await getExcelDataForAnalysis();
    
    if (!excelData || excelData.length === 0) {
        const fileRef = getStoredFileReference && getStoredFileReference();
        const fileName = fileRef ? fileRef.fileName : 'Excel file';
        
        return `
            <div class="analysis-unavailable" style="text-align: center; padding: 2rem; background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 8px; margin: 1rem 0;">
                <h4 style="color: #495057; margin-bottom: 1rem;">📊 VAS Analysis Ready</h4>
                <p style="color: #6c757d; margin-bottom: 1rem;">
                    To view detailed VAS analysis, please reselect your Excel file:
                </p>
                <p style="font-weight: 600; color: #495057; margin-bottom: 1.5rem;">
                    📁 ${fileName}
                </p>
                <button onclick="selectFileForAnalysis()" style="
                    background: #007bff;
                    color: white;
                    border: none;
                    padding: 10px 20px;
                    border-radius: 6px;
                    cursor: pointer;
                    font-weight: 500;
                    font-size: 14px;
                ">📁 Select File for Analysis</button>
                <p style="font-size: 12px; color: #868e96; margin-top: 1rem;">
                    Your file reference is preserved - just reselect to enable full analysis
                </p>
            </div>
        `;
    }
    
    console.log(`✅ Excel data available for VAS analysis: ${excelData.length} rows`);
    
    // Filter for VAS transactions (Transaction Type = 600)
    const vasTransactions = excelData.filter(row => {
        const transactionType = row["Transaction Type"];
        return transactionType === 600 || transactionType === '600';
    });
    
    if (vasTransactions.length === 0) {
        return `
            <div style="text-align: center; padding: 2rem; background: #fff3cd; border: 1px solid #ffeaa7; border-radius: 8px;">
                <h4 style="color: #856404;">⚠️ No VAS Data Found</h4>
                <p style="color: #856404;">No VAS transactions (Type 600) found in the selected file.</p>
            </div>
        `;
    }
    
    console.log(`📊 Found ${vasTransactions.length} VAS transactions`);
    
    // Group VAS transactions by From Location
    const vasTaskAnalysis = analyzeVASByTask(vasTransactions);
    
    // Calculate comprehensive VAS metrics
    const vasMetrics = calculateComprehensiveVASMetrics(vasTaskAnalysis);
    
    let html = `
        <div class="vas-breakdown">
            <!-- Enhanced Operations Summary -->
            <div class="vas-summary" style="
                background: linear-gradient(135deg, #e8f5e8 0%, #d4edda 100%); 
                padding: 2rem; 
                border-radius: 12px;
                margin-bottom: 2rem;
                border: 1px solid #c3e6cb;
            ">
                <h4 style="margin: 0 0 1.5rem 0; color: #155724; font-size: 1.25rem; text-align: center;">
                    📊 VAS Operations Summary
                </h4>
                
                <!-- 5-column metrics grid -->
                <div style="display: grid; grid-template-columns: repeat(5, 1fr); gap: 1.5rem;">
                    <!-- Total Transactions -->
                    <div style="text-align: center;">
                        <div style="font-size: 2rem; font-weight: 800; color: #155724; line-height: 1;">
                            ${vasTransactions.length}
                        </div>
                        <div style="font-size: 0.875rem; color: #6c757d; font-weight: 600; margin-top: 0.25rem;">
                            Total Transactions
                        </div>
                    </div>
                    
                    <!-- Total Units -->
                    <div style="text-align: center;">
                        <div style="font-size: 2rem; font-weight: 800; color: #155724; line-height: 1;">
                            ${vasTaskAnalysis.totalUnits.toLocaleString()}
                        </div>
                        <div style="font-size: 0.875rem; color: #6c757d; font-weight: 600; margin-top: 0.25rem;">
                            Total Units
                        </div>
                    </div>
                    
                    <!-- Total Hours -->
                    <div style="text-align: center;">
                        <div style="font-size: 2rem; font-weight: 800; color: #155724; line-height: 1;">
                            ${vasTaskAnalysis.totalHours.toFixed(1)}
                        </div>
                        <div style="font-size: 0.875rem; color: #6c757d; font-weight: 600; margin-top: 0.25rem;">
                            Total Hours
                        </div>
                    </div>
                    
                    <!-- Actual TPH -->
                    <div style="text-align: center;">
                        <div style="font-size: 2rem; font-weight: 800; color: #155724; line-height: 1;">
                            ${vasMetrics.actualTPH.toFixed(1)}
                        </div>
                        <div style="font-size: 0.875rem; color: #6c757d; font-weight: 600; margin-top: 0.25rem;">
                            Actual TPH
                        </div>
                    </div>
                    
                    <!-- Actual TPLH -->
                    <div style="text-align: center;">
                        <div style="font-size: 2rem; font-weight: 800; color: #155724; line-height: 1;">
                            ${vasMetrics.actualTPLH.toFixed(1)}
                        </div>
                        <div style="font-size: 0.875rem; color: #6c757d; font-weight: 600; margin-top: 0.25rem;">
                            Actual TPLH
                        </div>
                    </div>
                </div>
                
                <!-- Target Performance Section -->
                <div style="
                    margin-top: 2rem; 
                    padding: 1.5rem; 
                    background: rgba(255,255,255,0.8); 
                    border-radius: 8px;
                    border: 1px solid rgba(21,87,36,0.1);
                ">
                    <h5 style="margin: 0 0 1rem 0; color: #155724; font-size: 1rem; text-align: center;">
                        🎯 Target Performance Analysis
                    </h5>
                    
                    <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 1.5rem;">
                        <!-- Suggested VAS TPH -->
                        <div style="text-align: center;">
                            <div style="font-size: 1.5rem; font-weight: 700; color: ${vasMetrics.performanceColor}; line-height: 1;">
                                ${vasMetrics.suggestedTPH.toFixed(1)}
                            </div>
                            <div style="font-size: 0.875rem; color: #6c757d; font-weight: 600; margin-top: 0.25rem;">
                                Suggested TPH
                            </div>
                            <div style="font-size: 0.75rem; color: #868e96; margin-top: 0.125rem;">
                                (Based on mix targets)
                            </div>
                        </div>
                        
                        <!-- Performance vs Target -->
                        <div style="text-align: center;">
                            <div style="font-size: 1.5rem; font-weight: 700; color: ${vasMetrics.performanceColor}; line-height: 1;">
                                ${vasMetrics.overallPerformance.toFixed(1)}%
                            </div>
                            <div style="font-size: 0.875rem; color: #6c757d; font-weight: 600; margin-top: 0.25rem;">
                                Performance vs Target
                            </div>
                            <div style="font-size: 0.75rem; color: #868e96; margin-top: 0.125rem;">
                                (${vasMetrics.performanceIndicator})
                            </div>
                        </div>
                        
                        <!-- Time Variance -->
                        <div style="text-align: center;">
                            <div style="font-size: 1.5rem; font-weight: 700; color: ${vasMetrics.varianceColor}; line-height: 1;">
                                ${vasMetrics.timeVariance >= 0 ? '+' : ''}${vasMetrics.timeVariance.toFixed(1)}h
                            </div>
                            <div style="font-size: 0.875rem; color: #6c757d; font-weight: 600; margin-top: 0.25rem;">
                                Time Variance
                            </div>
                            <div style="font-size: 0.75rem; color: #868e96; margin-top: 0.125rem;">
                                (Actual vs Expected)
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- Larger VAS Distribution Chart - Full Width Section -->
            <div class="vas-distribution-section" style="
                background: white; 
                padding: 2rem; 
                border-radius: 12px; 
                border: 1px solid #dee2e6;
                margin-bottom: 2rem;
                box-shadow: 0 2px 8px rgba(0,0,0,0.08);
            ">
                <h4 style="margin: 0 0 1.5rem 0; color: #495057; font-size: 1.1rem; text-align: center;">
                    🥧 VAS Task Distribution
                </h4>
                
                <div style="display: flex; align-items: center; justify-content: center; gap: 3rem;">
                    <!-- Large SVG Pie Chart -->
                    <div style="flex-shrink: 0;">
                        <svg viewBox="0 0 100 100" style="width: 240px; height: 240px;">
                            ${generateLargeVASPieChart(vasTaskAnalysis.tasks)}
                        </svg>
                    </div>
                    
                    <!-- Comprehensive Legend -->
                    <div style="flex: 1; max-width: 400px;">
                        ${generateVASPieChartLegend(vasTaskAnalysis.tasks)}
                    </div>
                </div>
            </div>
            
            <!-- Enhanced VAS task cards with performance indicators -->
            <div class="vas-tasks-section">
                <h4 style="margin: 0 0 1.5rem 0; color: #495057; font-size: 1.1rem;">
                    🛠️ VAS Task Performance Details
                </h4>
                
                <div class="vas-tasks-grid" style="display: grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); gap: 1rem;">
    `;
    
    // Display each VAS task with enhanced styling
    vasTaskAnalysis.tasks.forEach(task => {
        const targetUPH = VAS_TARGET_UPH && VAS_TARGET_UPH[task.location] ? VAS_TARGET_UPH[task.location] : null;
        const performanceData = calculateVASPerformance(task, targetUPH);
        
        html += `
            <div class="vas-task-card" style="
                border: 1px solid #dee2e6; 
                border-radius: 12px; 
                padding: 1.25rem; 
                background: white;
                box-shadow: 0 2px 8px rgba(0,0,0,0.08);
                transition: all 0.3s ease;
                position: relative;
                overflow: hidden;
            " onmouseover="this.style.transform='translateY(-2px)'; this.style.boxShadow='0 4px 16px rgba(0,0,0,0.12)'" 
               onmouseout="this.style.transform='translateY(0)'; this.style.boxShadow='0 2px 8px rgba(0,0,0,0.08)'">
                
                <!-- Task header with performance indicator -->
                <div class="task-header" style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 1rem;">
                    <div class="task-name" style="
                        font-weight: 600; 
                        color: #495057; 
                        font-size: 1rem;
                        line-height: 1.2;
                        flex: 1;
                        margin-right: 1rem;
                    ">${task.location}</div>
                    
                    <!-- Enhanced performance badge -->
                    <div class="task-performance" style="
                        padding: 0.375rem 0.75rem; 
                        border-radius: 20px; 
                        font-size: 0.875rem; 
                        font-weight: 700;
                        background: ${performanceData.backgroundColor};
                        color: ${performanceData.textColor};
                        border: 2px solid ${performanceData.borderColor};
                        min-width: 80px;
                        text-align: center;
                        white-space: nowrap;
                    ">
                        ${performanceData.displayText}%
                    </div>
                </div>
                
                <!-- Progress bar for visual performance indicator -->
                ${targetUPH ? `
                <div style="margin-bottom: 1rem;">
                    <div style="display: flex; justify-content: space-between; font-size: 0.75rem; color: #6c757d; margin-bottom: 0.25rem;">
                        <span>Performance vs Target</span>
                        <span>${task.actualUPH.toFixed(1)} / ${targetUPH} UPH</span>
                    </div>
                    <div style="
                        width: 100%; 
                        height: 6px; 
                        background: #e9ecef; 
                        border-radius: 3px; 
                        overflow: hidden;
                    ">
                        <div style="
                            width: ${Math.min(task.performancePercentage, 150)}%; 
                            height: 100%; 
                            background: ${performanceData.progressColor}; 
                            border-radius: 3px;
                            transition: width 0.8s ease;
                        "></div>
                    </div>
                </div>
                ` : ''}
                
                <!-- Task metrics in clean grid -->
                <div class="task-metrics" style="
                    display: grid; 
                    grid-template-columns: 1fr 1fr; 
                    gap: 0.75rem; 
                    font-size: 0.875rem;
                ">
                    <div class="metric">
                        <span class="metric-label" style="color: #6c757d; display: block; font-size: 0.75rem; margin-bottom: 0.25rem;">Units</span>
                        <span class="metric-value" style="font-weight: 600; color: #495057; font-size: 1rem;">${task.units.toLocaleString()}</span>
                    </div>
                    <div class="metric">
                        <span class="metric-label" style="color: #6c757d; display: block; font-size: 0.75rem; margin-bottom: 0.25rem;">% of Total VAS</span>
                        <span class="metric-value" style="font-weight: 600; color: #495057; font-size: 1rem;">${task.percentageOfTotal.toFixed(1)}%</span>
                    </div>
                    <div class="metric">
                        <span class="metric-label" style="color: #6c757d; display: block; font-size: 0.75rem; margin-bottom: 0.25rem;">Hours</span>
                        <span class="metric-value" style="font-weight: 600; color: #495057; font-size: 1rem;">${task.hours.toFixed(2)}</span>
                    </div>
                    <div class="metric">
                        <span class="metric-label" style="color: #6c757d; display: block; font-size: 0.75rem; margin-bottom: 0.25rem;">Transactions</span>
                        <span class="metric-value" style="font-weight: 600; color: #495057; font-size: 1rem;">${task.transactions}</span>
                    </div>
                </div>
            </div>
        `;
    });
    
    html += `
                </div>
            </div>
        </div>
    `;
    
    return html;
}

// ============================================================================
// CART ANALYSIS GENERATION
// ============================================================================
const CART_TARGET_TPLH = 19; // Transactions Per Labor Hour target

async function generateCARTAnalysis() {
    console.log('🔍 Generating CART analysis by username...');
    
    const excelData = await getExcelDataForAnalysis();
    
    if (!excelData || excelData.length === 0) {
        const fileRef = getStoredFileReference && getStoredFileReference();
        const fileName = fileRef ? fileRef.fileName : 'Excel file';
        
        return `
            <div class="analysis-unavailable" style="text-align: center; padding: 2rem; background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 8px; margin: 1rem 0;">
                <h4 style="color: #495057; margin-bottom: 1rem;">🛒 CART Analysis Ready</h4>
                <p style="color: #6c757d; margin-bottom: 1rem;">
                    To view detailed CART analysis, please reselect your Excel file:
                </p>
                <p style="font-weight: 600; color: #495057; margin-bottom: 1.5rem;">
                    📁 ${fileName}
                </p>
                <button onclick="selectFileForAnalysis()" style="
                    background: #007bff;
                    color: white;
                    border: none;
                    padding: 10px 20px;
                    border-radius: 6px;
                    cursor: pointer;
                    font-weight: 500;
                    font-size: 14px;
                ">📁 Select File for Analysis</button>
                <p style="font-size: 12px; color: #868e96; margin-top: 1rem;">
                    Your file reference is preserved - just reselect to enable full analysis
                </p>
            </div>
        `;
    }
    
    console.log(`✅ Excel data available for CART analysis: ${excelData.length} rows`);
    
    // Filter for CART transactions (Transaction Type = 212 AND From Location starts with MOVEXX)
    const cartTransactions = excelData.filter(row => {
        const transactionType = row["Transaction Type"];
        const fromLocation = row["From Location"] || '';
        return (transactionType === 212 || transactionType === '212') && 
               fromLocation.startsWith('MOVEXX');
    });
    
    if (cartTransactions.length === 0) {
        return `
            <div style="text-align: center; padding: 2rem; background: #fff3cd; border: 1px solid #ffeaa7; border-radius: 8px;">
                <h4 style="color: #856404;">⚠️ No CART Data Found</h4>
                <p style="color: #856404;">No CART transactions (Type 212 with MOVEXX locations) found in the selected file.</p>
            </div>
        `;
    }
    
    console.log(`🛒 Found ${cartTransactions.length} CART transactions`);
    
    const cartAnalysis = analyzeCARTTransactionsByUser(cartTransactions);
    
    // Calculate overall TPLH performance
    const overallTPLH = cartAnalysis.totalHours > 0 ? cartTransactions.length / cartAnalysis.totalHours : 0;
    const tplhPerformance = CART_TARGET_TPLH > 0 ? (overallTPLH / CART_TARGET_TPLH) * 100 : 0;
    const tplhVariance = overallTPLH - CART_TARGET_TPLH;
    
    // Determine performance colors
    const tplhPerformanceColor = tplhPerformance >= 100 ? '#28a745' : 
                                 tplhPerformance >= 90 ? '#ffc107' : '#dc3545';
    const tplhVarianceColor = tplhVariance >= 0 ? '#28a745' : '#dc3545';
    
    return `
        <div class="cart-breakdown">
            <!-- Enhanced CART Summary with TPLH -->
            <div class="cart-summary" style="
                background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%); 
                padding: 1.5rem; 
                border-radius: 8px; 
                margin-bottom: 1.5rem;
                border-left: 4px solid #2196f3;
            ">
                <h4 style="margin: 0 0 1rem 0; color: #1565c0; font-size: 1.1rem;">
                    📊 CART Operations Summary
                </h4>
                
                <!-- Enhanced metrics grid with TPLH -->
                <div style="display: grid; grid-template-columns: repeat(6, 1fr); gap: 1rem; margin-bottom: 1rem;">
                    <!-- Total Users -->
                    <div style="text-align: center;">
                        <div style="font-size: 1.5rem; font-weight: 700; color: #1565c0; line-height: 1;">
                            ${cartAnalysis.users.length}
                        </div>
                        <div style="font-size: 0.8rem; color: #6c757d; font-weight: 600; margin-top: 0.25rem;">
                            Active Users
                        </div>
                    </div>
                    
                    <!-- Total Transactions -->
                    <div style="text-align: center;">
                        <div style="font-size: 1.5rem; font-weight: 700; color: #1565c0; line-height: 1;">
                            ${cartTransactions.length}
                        </div>
                        <div style="font-size: 0.8rem; color: #6c757d; font-weight: 600; margin-top: 0.25rem;">
                            Total Transactions
                        </div>
                    </div>
                    
                    <!-- Total Units -->
                    <div style="text-align: center;">
                        <div style="font-size: 1.5rem; font-weight: 700; color: #1565c0; line-height: 1;">
                            ${cartAnalysis.totalUnits.toLocaleString()}
                        </div>
                        <div style="font-size: 0.8rem; color: #6c757d; font-weight: 600; margin-top: 0.25rem;">
                            Total Units
                        </div>
                    </div>
                    
                    <!-- Total Hours -->
                    <div style="text-align: center;">
                        <div style="font-size: 1.5rem; font-weight: 700; color: #1565c0; line-height: 1;">
                            ${cartAnalysis.totalHours.toFixed(2)}
                        </div>
                        <div style="font-size: 0.8rem; color: #6c757d; font-weight: 600; margin-top: 0.25rem;">
                            Total Hours
                        </div>
                    </div>
                    
                    <!-- Average UPH -->
                    <div style="text-align: center;">
                        <div style="font-size: 1.5rem; font-weight: 700; color: #1565c0; line-height: 1;">
                            ${cartAnalysis.averageUPH.toFixed(1)}
                        </div>
                        <div style="font-size: 0.8rem; color: #6c757d; font-weight: 600; margin-top: 0.25rem;">
                            Average UPH
                        </div>
                    </div>
                    
                    <!-- Actual TPLH -->
                    <div style="text-align: center;">
                        <div style="font-size: 1.5rem; font-weight: 700; color: ${tplhPerformanceColor}; line-height: 1;">
                            ${overallTPLH.toFixed(1)}
                        </div>
                        <div style="font-size: 0.8rem; color: #6c757d; font-weight: 600; margin-top: 0.25rem;">
                            Overall TPLH
                        </div>
                    </div>
                </div>
                
                <!-- TPLH Performance Section -->
                <div style="
                    background: rgba(255,255,255,0.8); 
                    padding: 1rem; 
                    border-radius: 6px;
                    display: grid; 
                    grid-template-columns: repeat(3, 1fr); 
                    gap: 1rem;
                ">
                    <!-- Target TPLH -->
                    <div style="text-align: center;">
                        <div style="font-size: 1.25rem; font-weight: 600; color: #495057; line-height: 1;">
                            ${CART_TARGET_TPLH}
                        </div>
                        <div style="font-size: 0.75rem; color: #6c757d; font-weight: 600;">
                            Target TPLH
                        </div>
                    </div>
                    
                    <!-- Performance vs Target -->
                    <div style="text-align: center;">
                        <div style="font-size: 1.25rem; font-weight: 600; color: ${tplhPerformanceColor}; line-height: 1;">
                            ${tplhPerformance.toFixed(1)}%
                        </div>
                        <div style="font-size: 0.75rem; color: #6c757d; font-weight: 600;">
                            Performance vs Target
                        </div>
                    </div>
                    
                    <!-- TPLH Variance -->
                    <div style="text-align: center;">
                        <div style="font-size: 1.25rem; font-weight: 600; color: ${tplhVarianceColor}; line-height: 1;">
                            ${tplhVariance >= 0 ? '+' : ''}${tplhVariance.toFixed(1)}
                        </div>
                        <div style="font-size: 0.75rem; color: #6c757d; font-weight: 600;">
                            TPLH Variance
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- Individual User Performance Cards -->
            <div class="cart-users-section">
                <h4 style="margin: 0 0 1rem 0; color: #495057; font-size: 1rem;">
                    👤 Individual CART Performance
                </h4>
                
                <div class="cart-users-grid" style="display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 1rem;">
                    ${cartAnalysis.users.map(user => {
                        // Calculate TPLH for this user
                        const userTPLH = user.hours > 0 ? user.transactions / user.hours : 0;
                        const userTPLHPerformance = CART_TARGET_TPLH > 0 ? (userTPLH / CART_TARGET_TPLH) * 100 : 0;
                        const userTPLHVariance = userTPLH - CART_TARGET_TPLH;
                        
                        // Performance colors and rating for this user
                        let performanceColor, performanceRating, cardBorderColor;
                        if (userTPLHPerformance >= 100) {
                            performanceColor = '#28a745';
                            performanceRating = 'Above Target';
                            cardBorderColor = '#28a745';
                        } else if (userTPLHPerformance >= 90) {
                            performanceColor = '#ffc107';
                            performanceRating = 'Near Target';
                            cardBorderColor = '#ffc107';
                        } else {
                            performanceColor = '#dc3545';
                            performanceRating = 'Below Target';
                            cardBorderColor = '#dc3545';
                        }
                        
                        const varianceColor = userTPLHVariance >= 0 ? '#28a745' : '#dc3545';
                        
                        return `
                            <div class="cart-user-card" style="
                                border: 2px solid ${cardBorderColor}; 
                                border-radius: 12px; 
                                padding: 1.25rem; 
                                background: white;
                                box-shadow: 0 3px 8px rgba(0,0,0,0.08);
                                transition: all 0.3s ease;
                                position: relative;
                            " onmouseover="this.style.transform='translateY(-3px)'; this.style.boxShadow='0 6px 16px rgba(0,0,0,0.15)'" 
                               onmouseout="this.style.transform='translateY(0)'; this.style.boxShadow='0 3px 8px rgba(0,0,0,0.08)'">
                                
                                <!-- User Header with Performance Badge -->
                                <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 1rem;">
                                    <div>
                                        <div class="user-name" style="
                                            font-weight: 700; 
                                            color: #495057; 
                                            font-size: 1.1rem;
                                            margin-bottom: 0.25rem;
                                        ">${user.username}</div>
                                        <div style="font-size: 0.8rem; color: #6c757d; font-weight: 500;">
                                            ${user.locations.length} MOVEXX Location${user.locations.length !== 1 ? 's' : ''}
                                        </div>
                                    </div>
                                    
                                    <div style="
                                        padding: 0.5rem 0.75rem; 
                                        border-radius: 20px; 
                                        font-size: 0.875rem; 
                                        font-weight: 700;
                                        background: ${performanceColor === '#28a745' ? '#d4edda' : 
                                                     performanceColor === '#ffc107' ? '#fff3cd' : '#f8d7da'};
                                        color: ${performanceColor};
                                        border: 2px solid ${performanceColor};
                                        text-align: center;
                                        min-width: 70px;
                                    ">
                                        ${userTPLHPerformance.toFixed(0)}%
                                    </div>
                                </div>
                                
                                <!-- User Metrics Grid -->
                                <div style="
                                    display: grid; 
                                    grid-template-columns: 1fr 1fr; 
                                    gap: 0.75rem; 
                                    margin-bottom: 1rem;
                                ">
                                    <div>
                                        <span style="color: #6c757d; display: block; font-size: 0.75rem; margin-bottom: 0.25rem;">Units</span>
                                        <span style="font-weight: 600; color: #495057; font-size: 1rem;">${user.units.toLocaleString()}</span>
                                    </div>
                                    <div>
                                        <span style="color: #6c757d; display: block; font-size: 0.75rem; margin-bottom: 0.25rem;">Transactions</span>
                                        <span style="font-weight: 600; color: #495057; font-size: 1rem;">${user.transactions}</span>
                                    </div>
                                    <div>
                                        <span style="color: #6c757d; display: block; font-size: 0.75rem; margin-bottom: 0.25rem;">UPH</span>
                                        <span style="font-weight: 600; color: #495057; font-size: 1rem;">${user.uph.toFixed(1)}</span>
                                    </div>
                                    <div>
                                        <span style="color: #6c757d; display: block; font-size: 0.75rem; margin-bottom: 0.25rem;">Hours</span>
                                        <span style="font-weight: 600; color: #495057; font-size: 1rem;">${user.hours.toFixed(2)}</span>
                                    </div>
                                </div>
                                
                                <!-- TPLH Performance Section -->
                                <div style="
                                    background: #f8f9fa; 
                                    padding: 1rem; 
                                    border-radius: 8px;
                                    border-left: 4px solid ${performanceColor};
                                ">
                                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.5rem;">
                                        <div style="font-size: 0.875rem; font-weight: 600; color: #495057;">
                                            TPLH Performance
                                        </div>
                                        <div style="font-size: 0.75rem; color: ${performanceColor}; font-weight: 600;">
                                            ${performanceRating}
                                        </div>
                                    </div>
                                    
                                    <div style="display: flex; justify-content: space-between; align-items: center;">
                                        <div>
                                            <div style="font-size: 0.75rem; color: #6c757d; margin-bottom: 0.25rem;">
                                                Actual / Target
                                            </div>
                                            <div style="font-size: 1.1rem; font-weight: 700; color: ${performanceColor};">
                                                ${userTPLH.toFixed(1)} / ${CART_TARGET_TPLH}
                                            </div>
                                        </div>
                                        <div style="text-align: right;">
                                            <div style="font-size: 0.75rem; color: #6c757d; margin-bottom: 0.25rem;">
                                                Variance
                                            </div>
                                            <div style="font-size: 1.1rem; font-weight: 700; color: ${varianceColor};">
                                                ${userTPLHVariance >= 0 ? '+' : ''}${userTPLHVariance.toFixed(1)}
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                
                                <!-- Locations worked -->
                                <div style="margin-top: 0.75rem;">
                                    <div style="font-size: 0.75rem; color: #6c757d; margin-bottom: 0.5rem;">
                                        Locations: ${user.locations.join(', ')}
                                    </div>
                                </div>
                            </div>
                        `;
                    }).join('')}
                </div>
            </div>
        </div>
    `;
}

// ADD this global function for the file selection button:
function selectFileForAnalysis() {
    const fileInput = document.getElementById('fileInput');
    if (fileInput) {
        fileInput.click();
    } else {
        alert('File input not found. Please refresh the page and try again.');
    }
}

// ============================================================================
// CRET ANALYSIS GENERATION
// ============================================================================
function generateCRETAnalysis() {
    // Get CRET data from labor analysis
    if (!kpiResults.labor || !laborData) {
        return `
            <div class="analysis-placeholder" style="text-align: center; padding: 2rem; background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 8px;">
                <h4 style="color: #495057; margin-bottom: 1rem;">🔄 CRET Analysis</h4>
                <p style="color: #6c757d;">
                    Customer Returns (CRET) analysis will be available when labor data is loaded.
                </p>
            </div>
        `;
    }
    
    const cretData = calculateCRETMetrics();
    
    if (cretData.totalHours === 0) {
        return `
            <div class="analysis-placeholder" style="text-align: center; padding: 2rem; background: #fff3cd; border: 1px solid #ffeaa7; border-radius: 8px;">
                <h4 style="color: #856404;">📋 No CRET Data Found</h4>
                <p style="color: #856404;">
                    No Customer Returns data found in the current labor analysis.
                </p>
            </div>
        `;
    }
    
    // ONLY RETURN THE CLEAN CARD SECTION - NO REDUNDANT METRICS
    return `
        <div class="cret-breakdown">
            <!-- This matches your existing card design -->
            <div class="cret-functions">
                <h4>🔄 Customer Returns Operations</h4>
                <div class="cret-functions-grid">
                    ${cretData.functions.map(func => `
                        <div class="cret-function-row">
                            <span class="function-name">${func.name}</span>
                            <span class="function-hours">${func.hours.toFixed(1)}h</span>
                            <span class="function-transactions">${func.transactions.toLocaleString()} trans</span>
                        </div>
                    `).join('')}
                </div>
            </div>
        </div>
    `;
}


// ============================================================================
// PERFORMANCE KEY GENERATION
// ============================================================================
function generatePerformanceKey() {
  return `
    <div class="performance-key">
      <div class="key-section">
        <h4>🎯 Performance Thresholds</h4>
        <div class="threshold-grid">
          <div class="threshold-item excellent">
            <div class="threshold-color"></div>
            <div class="threshold-text">
              <strong>🟢 Excellent:</strong> ≥105% of target UPH
            </div>
          </div>
          <div class="threshold-item acceptable">
            <div class="threshold-color"></div>
            <div class="threshold-text">
              <strong>🟡 Acceptable:</strong> 95-105% of target UPH
            </div>
          </div>
          <div class="threshold-item poor">
            <div class="threshold-color"></div>
            <div class="threshold-text">
              <strong>🔴 Needs Attention:</strong> &lt;95% of target UPH
            </div>
          </div>
        </div>
      </div>
      
      <div class="key-section">
        <h4>📋 Definitions</h4>
        <div class="definitions-grid">
          <div class="definition-item">
            <strong>VAS:</strong> Value-Added Services (Transaction Code 600)
          </div>
          <div class="definition-item">
            <strong>CART:</strong> Cart Operations (Transaction Code 212 with MOVEXX locations)
          </div>
          <div class="definition-item">
            <strong>CRET:</strong> Customer Returns Operations
          </div>
          <div class="definition-item">
            <strong>UPH:</strong> Units Per Hour (calculated from actual transaction data)
          </div>
          <div class="definition-item">
            <strong>Performance %:</strong> (Actual UPH ÷ Target UPH) × 100
          </div>
        </div>
      </div>
    </div>
  `;
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

function getPerformanceClass(performancePercentage) {
    if (performancePercentage >= 105) {
        return 'performance-excellent';
    } else if (performancePercentage >= 95) {
        return 'performance-acceptable';
    } else {
        return 'performance-poor';
    }
}

function analyzeVASByTask(vasTransactions) {
    const taskGroups = {};
    let totalUnits = 0;
    let totalHours = 0;
    
    vasTransactions.forEach(row => {
        const fromLocation = row["From Location"] || 'Unknown';
        const units = parseFloat(row["Quantity"]) || 0;
        const timeToExecute = parseFloat(row["Time to Execute"]) || 0;
        const hours = timeToExecute / 3600; // Convert seconds to hours
        
        if (!taskGroups[fromLocation]) {
            taskGroups[fromLocation] = {
                location: fromLocation,
                units: 0,
                hours: 0,
                transactions: 0
            };
        }
        
        taskGroups[fromLocation].units += units;
        taskGroups[fromLocation].hours += hours;
        taskGroups[fromLocation].transactions += 1;
        
        totalUnits += units;
        totalHours += hours;
    });
    
    // Calculate metrics for each task
    const tasks = Object.values(taskGroups).map(task => {
        const actualUPH = task.hours > 0 ? task.units / task.hours : 0;
        const targetUPH = VAS_TARGET_UPH && VAS_TARGET_UPH[task.location] ? VAS_TARGET_UPH[task.location] : 0;
        const performancePercentage = targetUPH > 0 ? (actualUPH / targetUPH) * 100 : 0;
        const percentageOfTotal = totalUnits > 0 ? (task.units / totalUnits) * 100 : 0;
        
        return {
            ...task,
            actualUPH,
            performancePercentage,
            percentageOfTotal
        };
    });
    
    // Sort by units descending
    tasks.sort((a, b) => b.units - a.units);
    
    return {
        tasks,
        totalUnits,
        totalHours
    };
}

function analyzeCARTTransactionsEnhanced(cartTransactions) {
    const locationGroups = {};
    let totalUnits = 0;
    let totalHours = 0;
    
    cartTransactions.forEach(row => {
        const fromLocation = row["From Location"] || 'Unknown';
        const units = parseFloat(row["Quantity"]) || 0;
        const timeToExecute = parseFloat(row["Time to Execute"]) || 0;
        const hours = timeToExecute / 3600; // Convert seconds to hours
        
        if (!locationGroups[fromLocation]) {
            locationGroups[fromLocation] = {
                name: fromLocation,
                units: 0,
                hours: 0,
                transactions: 0
            };
        }
        
        locationGroups[fromLocation].units += units;
        locationGroups[fromLocation].hours += hours;
        locationGroups[fromLocation].transactions += 1;
        
        totalUnits += units;
        totalHours += hours;
    });
    
    // Calculate enhanced metrics for each location
    const locations = Object.values(locationGroups).map(location => ({
        ...location,
        uph: location.hours > 0 ? location.units / location.hours : 0,
        tplh: location.hours > 0 ? location.transactions / location.hours : 0
    }));
    
    // Sort by transactions descending (most active locations first)
    locations.sort((a, b) => b.transactions - a.transactions);
    
    const averageUPH = totalHours > 0 ? totalUnits / totalHours : 0;
    
    return {
        locations,
        totalUnits,
        totalHours,
        averageUPH
    };
}

// Additional helper functions would go here for:
// - calculateLaborSummaryForVASTab()
// - calculateOperationalSummary() 
// - getExcelDataForAnalysis()
// - calculateCRETMetrics()

// These would extract the necessary data from existing kpiResults and laborData

function displayEfficiencyAnalysis() {
  const efficiencyAnalysis = document.getElementById('efficiencyAnalysis');
  if (!efficiencyAnalysis) return;

  let html = '<div class="efficiency-grid">';

  // Inbound Department Focus
  if (kpiResults.combined && kpiResults.combined.inboundDepartment) {
    const inbound = kpiResults.combined.inboundDepartment;

    html += `
      <div class="efficiency-summary">
        <h4>🏭 Inbound Department Performance</h4>
        <div class="efficiency-metrics">
          <div class="efficiency-metric">
            <span class="metric-label">Labor UPH:</span>
            <span class="metric-value">${inbound.laborUPH.toFixed(1)}</span>
          </div>
          <div class="efficiency-metric">
            <span class="metric-label">Labor TPH:</span>
            <span class="metric-value">${inbound.laborTPH.toFixed(1)}</span>
          </div>
          <div class="efficiency-metric">
            <span class="metric-label">Actual TPLH:</span>
            <span class="metric-value">
  ${kpiResults.combined?.criticalKPIs?.TPLH?.toFixed(2) ?? 'N/A'}
</span>

          </div>
          <div class="efficiency-metric">
            <span class="metric-label">Actual TPH:</span>
            <span class="metric-value">
  ${kpiResults.combined?.criticalKPIs?.TPH?.toFixed(1) ?? 'N/A'}
</span>
          </div>
        </div>
      </div>
    `;
  }

  // Inbound Areas Ranking (if available)
  if (kpiResults.labor && kpiResults.labor.inboundAreas) {
    const sortedAreas = [...kpiResults.labor.inboundAreas]
      .filter(area => area.totalHours > 0)
      .sort((a, b) => b.uph - a.uph);

    if (sortedAreas.length > 0) {
      html += `
        <h4>🎯 Inbound Area Efficiency Ranking (by UPH)</h4>
        <div class="ranking-list">
      `;

      sortedAreas.forEach((area, index) => {
        const rank = index + 1;
        const rankClass = rank <= 2 ? 'top-rank' : 'normal-rank';

        html += `
          <div class="rank-item ${rankClass}">
            <div class="rank-number">${rank}</div>
            <div class="rank-details">
              <div class="rank-name">${area.name}</div>
              <div class="rank-metrics">
                <span>UPH: ${area.uph.toFixed(1)}</span> | 
                <span>Hours: ${area.totalHours.toFixed(1)}</span> | 
                <span>Units: ${area.totalUnits.toLocaleString()}</span>
              </div>
            </div>
          </div>
        `;
      });

      html += '</div>';
    }
  }

  html += '</div>';
  efficiencyAnalysis.innerHTML = html;
}

function displayCombinedAnalysis() {
  const combinedInsights = document.getElementById('combinedInsights');
  if (!combinedInsights || !kpiResults.combined) return;

  let html = '<div class="combined-grid">';

  if (kpiResults.combined.inboundDepartment) {
    const inbound = kpiResults.combined.inboundDepartment;
    const laborTPH = inbound.laborTPH;
    const actualTPLH = kpiResults.combined?.criticalKPIs?.TPLH;
    const actualTPH = kpiResults.combined?.criticalKPIs?.TPH;

    html += `
      <div class="insight-card">
        <h4>🎯 Inbound Performance Summary</h4>
        <div class="insight-content">
          <div class="performance-comparison">
            <div class="performance-metric">
              <span class="metric-label">Labor Hours:</span>
              <span class="metric-value">${inbound.totalHours.toFixed(1)}</span>
            </div>
            <div class="performance-metric">
              <span class="metric-label">Actual TPLH:</span>
              <div class="metric-value">${typeof actualTPLH === 'number' ? actualTPLH.toFixed(2) : 'N/A'}</div>
            </div>
            <div class="performance-metric">
              <span class="metric-label">Actual TPH:</span>
              <span class="metric-value performance-highlight">${typeof actualTPH === 'number' ? actualTPH.toFixed(1) : 'N/A'}</span>
            </div>
            <div class="performance-metric">
              <span class="metric-label">Labor Reported TPH:</span>
              <span class="metric-value">${typeof laborTPH === 'number' ? laborTPH.toFixed(1) : 'N/A'}</span>
            </div>
          </div>
        </div>
      </div>
    `;
  }

  html += '</div>';
  combinedInsights.innerHTML = html;
}

// Helper function for performance classification
function getPerformanceClass(variance) {
  if (variance <= 5) return 'performance-excellent';
  if (variance <= 15) return 'performance-good';
  return 'performance-needs-attention';
}

function switchTab(tabName) {
  // Update tab buttons
  document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
  event.target.classList.add('active');

  // Update tab content
  document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
  document.getElementById(tabName + 'Tab').classList.add('active');
}
// ============================================================================
// MISSING HELPER FUNCTIONS - Add these to snapshot-display.js
// ============================================================================

// Function to calculate labor summary for VAS tab
// ============================================================================
// MISSING HELPER FUNCTIONS - Add these to snapshot-display.js
// ============================================================================

// Function to calculate labor summary for VAS tab
function calculateLaborSummaryForVASTab() {
  if (!kpiResults.labor || !kpiResults.labor.inboundDepartment) {
    return {
      directHours: 0,
      indirectHours: 0,
      directPercentage: 0,
      indirectPercentage: 0
    };
  }

  // Use the same logic from displayLaborAnalysis but simplified
  const allFunctions = kpiResults.labor.functions || laborData?.functions || [];
  
  // Direct labor categories
  const directCategories = [
    'Reach-Truck Putaway',
    'Non- PIT Manual Putaway',
    'Putaway - Unknown',
    'Break Down Receiving',
    'Cart Receiving',
    'Container Receiver',
    'Full Pallet Receiving',
    'Parcel Receiving',
    'Receiving',
    'Small Pack Debundle'
  ];

  // Indirect labor categories
  const indirectCategories = [
    'Unloader',
    'Container Unload',
    'Breakdown Receiver',
    'Pallet Rework',
    'Pallet Wrangler – Dock Stocker IB',
    'Pallet Wrapper',
    'Pallet Loader',
    'Inbound Lead',
    'Problem Solver',
    'Vendor Compliance',
    'Inbound Training',
    'VAS',
    'VAS Execute',
    'Heat Shrink',
    'DeBundling',
    'Cart Putaway',
    'On-Clock Unallocated'
  ];

  let directHours = 0;
  let indirectHours = 0;

  allFunctions.forEach(func => {
    const hours = func.totalHours || 0;
    if (directCategories.some(cat => func.name === cat || func.name.includes(cat))) {
      directHours += hours;
    } else if (indirectCategories.some(cat => func.name === cat || func.name.includes(cat))) {
      indirectHours += hours;
    }
  });

  const totalIBHours = kpiResults.labor.inboundDepartment.totalHours;
  const directPercentage = totalIBHours > 0 ? ((directHours / totalIBHours) * 100).toFixed(1) : 0;
  const indirectPercentage = totalIBHours > 0 ? ((indirectHours / totalIBHours) * 100).toFixed(1) : 0;

  return {
    directHours,
    indirectHours,
    directPercentage,
    indirectPercentage
  };
}

// Function to calculate operational summary (VAS/CART counts)
function calculateOperationalSummary() {
  const excelData = getExcelDataForAnalysis();
  
  if (!excelData) {
    return {
      vasTransactions: 0,
      vasUnits: 0,
      cartTransactions: 0,
      cartUnits: 0
    };
  }

  // VAS transactions (Type 600)
  const vasTransactions = excelData.filter(row => {
    const transactionType = row["Transaction Type"];
    return transactionType === 600 || transactionType === '600';
  });

  // CART transactions (Type 212 with MOVEXX)
  const cartTransactions = excelData.filter(row => {
    const transactionType = row["Transaction Type"];
    const fromLocation = row["From Location"] || '';
    return (transactionType === 212 || transactionType === '212') && 
           fromLocation.startsWith('MOVEXX');
  });

  const vasUnits = vasTransactions.reduce((sum, row) => {
    return sum + (parseFloat(row["Quantity"]) || 0);
  }, 0);

  const cartUnits = cartTransactions.reduce((sum, row) => {
    return sum + (parseFloat(row["Quantity"]) || 0);
  }, 0);

  return {
    vasTransactions: vasTransactions.length,
    vasUnits,
    cartTransactions: cartTransactions.length,
    cartUnits
  };
}

// Function to get Excel data for analysis
// Function to get Excel data for analysis (updated for file reference approach)
function getExcelDataForAnalysis() {
    // Use the new function from snapshot-data.js
    if (typeof getExcelDataForAnalysisSync === 'function') {
        return getExcelDataForAnalysisSync();
    }
    
    console.warn('getExcelDataForAnalysisSync not available');
    return null;
}

// Function to calculate CRET metrics
function calculateCRETMetrics() {
  if (!kpiResults.labor || !laborData) {
    return {
      totalHours: 0,
      totalTransactions: 0,
      tph: 0,
      functions: []
    };
  }

  const allFunctions = kpiResults.labor.functions || laborData.functions || [];
  
  // Customer Returns function names
  const cretFunctionNames = [
    'Customer Returns Unloader',
    'Inbound Returns Processor',
    'Customer Returns - Ground',
    'Customer Returns Problem Solver'
  ];

  let totalHours = 0;
  let totalTransactions = 0;
  const functions = [];

  cretFunctionNames.forEach(funcName => {
    const match = allFunctions.find(f => f.name === funcName || f.name.includes(funcName));
    if (match && match.totalHours > 0) {
      totalHours += match.totalHours;
      totalTransactions += match.totalTransactions || 0;
      functions.push({
        name: match.name,
        hours: match.totalHours,
        transactions: match.totalTransactions || 0
      });
    }
  });

  const tph = totalHours > 0 ? totalTransactions / totalHours : 0;

  return {
    totalHours,
    totalTransactions,
    tph,
    functions
  };
}
// ============================================================================
// REFRESH FUNCTIONS FOR FILE REFERENCE APPROACH
// ============================================================================

async function refreshVASAnalysis() {
    console.log('Refreshing VAS analysis...');
    
    // Use the new file handle approach
    if (typeof getExcelDataForAnalysis === 'function') {
        try {
            const data = await window.getExcelDataForAnalysis(); // Use the async version from snapshot-data.js
            if (data && data.length > 0) {
                console.log('✅ Excel data refreshed, updating VAS analysis');
                displayProductivityAnalysis(); // Refresh the entire tab
            } else {
                console.log('⚠️ No Excel data available - showing file selection prompt');
                showAnalysisDataPrompt('VAS analysis');
            }
        } catch (error) {
            console.warn('❌ Failed to refresh Excel data:', error);
            showAnalysisDataPrompt('VAS analysis');
        }
    } else {
        console.warn('getExcelDataForAnalysis function not available');
        showAnalysisDataPrompt('VAS analysis');
    }
}

async function refreshCARTAnalysis() {
    console.log('Refreshing CART analysis...');
    
    // Use the new file handle approach
    if (typeof getExcelDataForAnalysis === 'function') {
        try {
            const data = await window.getExcelDataForAnalysis(); // Use the async version from snapshot-data.js
            if (data && data.length > 0) {
                console.log('✅ Excel data refreshed, updating CART analysis');
                displayProductivityAnalysis(); // Refresh the entire tab
            } else {
                console.log('⚠️ No Excel data available - showing file selection prompt');
                showAnalysisDataPrompt('CART analysis');
            }
        } catch (error) {
            console.warn('❌ Failed to refresh Excel data:', error);
            showAnalysisDataPrompt('CART analysis');
        }
    } else {
        console.warn('getExcelDataForAnalysis function not available');
        showAnalysisDataPrompt('CART analysis');
    }
}
function showAnalysisDataPrompt(analysisType) {
    const fileRef = getStoredFileReference && getStoredFileReference();
    const fileName = fileRef ? fileRef.fileName : 'Excel file';
    
    const message = `To view ${analysisType}, please reselect your Excel file:\n\n📁 ${fileName}\n\nThis will enable full analysis capabilities.`;
    
    if (confirm(message)) {
        const fileInput = document.getElementById('fileInput');
        if (fileInput) {
            fileInput.click();
        }
    }
}

// Helper function to calculate VAS performance with colors and indicators
function calculateVASPerformance(task, targetUPH) {
    if (!targetUPH || targetUPH === 'Unknown') {
        return {
            displayText: `${task.actualUPH.toFixed(1)} UPH`,
            backgroundColor: '#f8f9fa',
            textColor: '#6c757d',
            borderColor: '#dee2e6',
            progressColor: '#6c757d'
        };
    }
    
    const performanceRatio = task.actualUPH / targetUPH;
    const performancePercentage = (performanceRatio * 100);
    const difference = task.actualUPH - targetUPH;
    const sign = difference >= 0 ? '+' : '';
    
    let backgroundColor, textColor, borderColor, progressColor;
    
    if (performancePercentage >= 100) {
        // Above target - Green
        backgroundColor = '#d4edda';
        textColor = '#155724';
        borderColor = '#28a745';
        progressColor = '#28a745';
    } else {
        // Below target - Red
        backgroundColor = '#f8d7da';
        textColor = '#721c24';
        borderColor = '#dc3545';
        progressColor = '#dc3545';
    }
    
    return {
        displayText: `${sign}${difference.toFixed(1)}`,
        backgroundColor,
        textColor,
        borderColor,
        progressColor,
        performancePercentage
    };
}

// Helper function to generate CSS-based pie chart
function generateVASPieChart(tasks) {
    if (!tasks || tasks.length === 0) return '<p>No data for chart</p>';
    
    // Sort tasks by percentage for better visualization
    const sortedTasks = [...tasks].sort((a, b) => b.percentageOfTotal - a.percentageOfTotal);
    
    // Colors for pie segments
    const colors = [
        '#007bff', '#28a745', '#ffc107', '#dc3545', '#6f42c1', 
        '#fd7e14', '#20c997', '#e83e8c', '#6c757d', '#17a2b8'
    ];
    
    let currentAngle = 0;
    let pieSegments = '';
    let legend = '';
    
    sortedTasks.forEach((task, index) => {
        const percentage = task.percentageOfTotal;
        const angle = (percentage / 100) * 360;
        const color = colors[index % colors.length];
        
        // Create pie segment
        const x1 = 50 + 40 * Math.cos((currentAngle - 90) * Math.PI / 180);
        const y1 = 50 + 40 * Math.sin((currentAngle - 90) * Math.PI / 180);
        const x2 = 50 + 40 * Math.cos((currentAngle + angle - 90) * Math.PI / 180);
        const y2 = 50 + 40 * Math.sin((currentAngle + angle - 90) * Math.PI / 180);
        
        const largeArcFlag = angle > 180 ? 1 : 0;
        
        pieSegments += `
            <path d="M 50 50 L ${x1} ${y1} A 40 40 0 ${largeArcFlag} 1 ${x2} ${y2} Z" 
                  fill="${color}" 
                  stroke="white" 
                  stroke-width="1"
                  opacity="0.8">
            </path>
        `;
        
        // Create legend item
        legend += `
            <div style="display: flex; align-items: center; margin-bottom: 0.25rem; font-size: 0.75rem;">
                <div style="width: 12px; height: 12px; background: ${color}; border-radius: 2px; margin-right: 0.5rem;"></div>
                <span style="color: #495057; font-weight: 500;">${task.location.substring(0, 12)}${task.location.length > 12 ? '...' : ''}</span>
                <span style="margin-left: auto; color: #6c757d;">${percentage.toFixed(1)}%</span>
            </div>
        `;
        
        currentAngle += angle;
    });
    
    return `
        <div style="display: flex; align-items: center; height: 100%;">
            <!-- SVG Pie Chart -->
            <div style="flex: 1;">
                <svg viewBox="0 0 100 100" style="width: 120px; height: 120px; margin: 0 auto; display: block;">
                    ${pieSegments}
                </svg>
            </div>
            
            <!-- Legend -->
            <div style="flex: 1; padding-left: 1rem; max-height: 120px; overflow-y: auto;">
                ${legend}
            </div>
        </div>
    `;
}

// Calculate comprehensive VAS metrics including suggested TPH
function calculateComprehensiveVASMetrics(vasTaskAnalysis) {
    const { tasks, totalUnits, totalHours } = vasTaskAnalysis;
    const totalTransactions = tasks.reduce((sum, task) => sum + task.transactions, 0);
    
    // Calculate actual metrics
    const actualTPH = totalHours > 0 ? totalUnits / totalHours : 0;
    const actualTPLH = totalHours > 0 ? totalTransactions / totalHours : 0;
    
    // Calculate suggested TPH based on mix of VAS tasks
    let totalExpectedTime = 0;
    let totalKnownUnits = 0;
    
    tasks.forEach(task => {
        const targetUPH = VAS_TARGET_UPH && VAS_TARGET_UPH[task.location] ? VAS_TARGET_UPH[task.location] : null;
        if (targetUPH && targetUPH !== 'Unknown') {
            const expectedTime = task.units / targetUPH; // Expected hours for this task
            totalExpectedTime += expectedTime;
            totalKnownUnits += task.units;
        }
    });
    
    // Calculate suggested TPH (what the TPH should be based on the mix)
    const suggestedTPH = totalExpectedTime > 0 ? totalKnownUnits / totalExpectedTime : 0;
    
    // Calculate overall performance
    const overallPerformance = suggestedTPH > 0 ? (actualTPH / suggestedTPH) * 100 : 0;
    
    // Calculate time variance (how much extra/less time was taken)
    const timeVariance = totalHours - totalExpectedTime;
    
    // Determine colors based on performance
    let performanceColor, varianceColor, performanceIndicator;
    
    if (overallPerformance >= 100) {
        performanceColor = '#28a745';
        performanceIndicator = 'Above Target';
    } else if (overallPerformance >= 90) {
        performanceColor = '#ffc107';
        performanceIndicator = 'Near Target';
    } else {
        performanceColor = '#dc3545';
        performanceIndicator = 'Below Target';
    }
    
    varianceColor = timeVariance <= 0 ? '#28a745' : '#dc3545';
    
    return {
        actualTPH,
        actualTPLH,
        suggestedTPH,
        overallPerformance,
        timeVariance,
        performanceColor,
        varianceColor,
        performanceIndicator
    };
}

// Generate larger pie chart SVG paths
function generateLargeVASPieChart(tasks) {
    if (!tasks || tasks.length === 0) return '<text x="50" y="50" text-anchor="middle" fill="#6c757d">No data</text>';
    
    // Sort tasks by percentage for better visualization
    const sortedTasks = [...tasks].sort((a, b) => b.percentageOfTotal - a.percentageOfTotal);
    
    // Colors for pie segments (more vibrant for larger chart)
    const colors = [
        '#007bff', '#28a745', '#ffc107', '#dc3545', '#6f42c1', 
        '#fd7e14', '#20c997', '#e83e8c', '#6c757d', '#17a2b8'
    ];
    
    let currentAngle = 0;
    let pieSegments = '';
    
    sortedTasks.forEach((task, index) => {
        const percentage = task.percentageOfTotal;
        const angle = (percentage / 100) * 360;
        const color = colors[index % colors.length];
        
        // Create pie segment with larger radius for bigger chart
        const x1 = 50 + 45 * Math.cos((currentAngle - 90) * Math.PI / 180);
        const y1 = 50 + 45 * Math.sin((currentAngle - 90) * Math.PI / 180);
        const x2 = 50 + 45 * Math.cos((currentAngle + angle - 90) * Math.PI / 180);
        const y2 = 50 + 45 * Math.sin((currentAngle + angle - 90) * Math.PI / 180);
        
        const largeArcFlag = angle > 180 ? 1 : 0;
        
        pieSegments += `
            <path d="M 50 50 L ${x1} ${y1} A 45 45 0 ${largeArcFlag} 1 ${x2} ${y2} Z" 
                  fill="${color}" 
                  stroke="white" 
                  stroke-width="2"
                  opacity="0.9"
                  style="transition: opacity 0.3s ease;">
                <title>${task.location}: ${percentage.toFixed(1)}%</title>
            </path>
        `;
        
        currentAngle += angle;
    });
    
    return pieSegments;
}

// Generate comprehensive legend for larger pie chart
function generateVASPieChartLegend(tasks) {
    if (!tasks || tasks.length === 0) return '<p>No data for legend</p>';
    
    const sortedTasks = [...tasks].sort((a, b) => b.percentageOfTotal - a.percentageOfTotal);
    
    const colors = [
        '#007bff', '#28a745', '#ffc107', '#dc3545', '#6f42c1', 
        '#fd7e14', '#20c997', '#e83e8c', '#6c757d', '#17a2b8'
    ];
    
    let legend = '<div style="display: grid; gap: 0.5rem;">';
    
    sortedTasks.forEach((task, index) => {
        const color = colors[index % colors.length];
        const targetUPH = VAS_TARGET_UPH && VAS_TARGET_UPH[task.location] ? VAS_TARGET_UPH[task.location] : 'Unknown';
        
        legend += `
            <div style="
                display: grid; 
                grid-template-columns: 16px 1fr auto auto; 
                gap: 0.75rem; 
                align-items: center; 
                padding: 0.5rem; 
                background: #f8f9fa; 
                border-radius: 6px;
                font-size: 0.875rem;
            ">
                <div style="width: 16px; height: 16px; background: ${color}; border-radius: 3px;"></div>
                <span style="color: #495057; font-weight: 500; text-align: left;">${task.location}</span>
                <span style="color: #6c757d; font-weight: 600;">${task.percentageOfTotal.toFixed(1)}%</span>
                <span style="color: #6c757d; font-size: 0.8rem;">${task.units.toLocaleString()} units</span>
            </div>
        `;
    });
    
    legend += '</div>';
    return legend;
}

function analyzeCARTTransactionsByUser(cartTransactions) {
    const userGroups = {};
    let totalUnits = 0;
    let totalHours = 0;
    
    cartTransactions.forEach(row => {
        // Get username from column G (index 6 in 0-based array)
        // Note: Excel columns are typically: A=0, B=1, C=2, D=3, E=4, F=5, G=6
        const username = row[Object.keys(row)[6]] || 'Unknown User'; // Column G
        const fromLocation = row["From Location"] || 'Unknown';
        const units = parseFloat(row["Quantity"]) || 0;
        const timeToExecute = parseFloat(row["Time to Execute"]) || 0;
        const hours = timeToExecute / 3600; // Convert seconds to hours
        
        if (!userGroups[username]) {
            userGroups[username] = {
                username: username,
                units: 0,
                hours: 0,
                transactions: 0,
                locations: new Set() // Track unique locations for this user
            };
        }
        
        userGroups[username].units += units;
        userGroups[username].hours += hours;
        userGroups[username].transactions += 1;
        userGroups[username].locations.add(fromLocation);
        
        totalUnits += units;
        totalHours += hours;
    });
    
    // Calculate enhanced metrics for each user
    const users = Object.values(userGroups).map(user => ({
        ...user,
        uph: user.hours > 0 ? user.units / user.hours : 0,
        tplh: user.hours > 0 ? user.transactions / user.hours : 0,
        locations: Array.from(user.locations).sort() // Convert Set to sorted Array
    }));
    
    // Sort by TPLH performance (best performers first)
    users.sort((a, b) => b.tplh - a.tplh);
    
    const averageUPH = totalHours > 0 ? totalUnits / totalHours : 0;
    
    return {
        users,
        totalUnits,
        totalHours,
        averageUPH
    };
}

// Make refresh functions globally available
window.selectFileForAnalysis = selectFileForAnalysis;
window.refreshVASAnalysis = refreshVASAnalysis;
window.refreshCARTAnalysis = refreshCARTAnalysis;

