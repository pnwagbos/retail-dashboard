/* --- 4. script.js Requirements (Consolidated and Fixed) --- */

// --- A. Initialization & UI Control (DOM Elements) ---
const DOM = {
    excelFile: document.getElementById('excelFile'),
    runAnalysisBtn: document.getElementById('runAnalysisBtn'),
    resetAppBtn: document.getElementById('resetAppBtn'),
    // Guide Dropdown Elements
    dataGuideBtn: document.getElementById('dataGuideBtn'),
    guideDropdown: document.getElementById('dataGuideDropdown'),
    requiredColumnsTable: document.getElementById('requiredColumnsTable'),
    kpiGuideTable: document.getElementById('kpiGuideTable'),
    // NEW Palette Dropdown Elements
    paletteBtn: document.getElementById('paletteBtn'),
    paletteDropdown: document.getElementById('paletteDropdown'),
    // Filter Elements
    startDateFilter: document.getElementById('startDateFilter'),
    endDateFilter: document.getElementById('endDateFilter'),
    categoryFilter: document.getElementById('categoryFilter'),
    productFilter: document.getElementById('productFilter'),
    applyFilterBtn: document.getElementById('applyFilterBtn'),
    clearFilterBtn: document.getElementById('clearFilterBtn'),
    // DYNAMIC PROXY INPUTS
    sqFootageInput: document.getElementById('sqFootageInput'), 
    trafficInput: document.getElementById('trafficInput'), 
    // Status Elements
    loader: document.getElementById('loader'),
    statusMessage: document.getElementById('statusMessage'),
    runtimeDisplay: document.getElementById('runtimeDisplay'),
    // Dashboard Elements
    dashboardContainer: document.getElementById('dashboardContainer'),
    metricsContainer: document.getElementById('metrics'),
    inventoryMetricsContainer: document.getElementById('inventoryMetrics'),
    salesTrendsChart: document.getElementById('salesTrendsChart'),
    categorySalesChart: document.getElementById('categorySalesChart'),
    topProductsChart: document.getElementById('topProductsChart'),
    topProductsTable: document.getElementById('topProductsTable').querySelector('tbody'),
    reorderTable: document.getElementById('reorderTable').querySelector('tbody'),
    productProfitTable: document.getElementById('productProfitTable').querySelector('tbody'),
    // NEW ABC TABLE
    abcAnalysisTable: document.getElementById('abcAnalysisTable').querySelector('tbody'), 
    downloadSampleLink: document.getElementById('download-sample-link'),
    // Chart Control Elements
    chartToggleBtns: document.querySelectorAll('.chart-toggle-btn'),
};

let rawData = [];
let analysisResults = null;
let chartInstances = {}; // Store Chart.js instances

// Filter State
let currentFilters = {
    startDate: null,
    endDate: null,
    categories: [],
    products: []
};

// Chart State
let currentChartTypes = {
    salesTrendsChart: 'bar' // Default is bar (monthly comparison)
};

// NEW: Sort State for Tables
// Stores the active sorting column, direction, and the source data key for each table.
let currentSortState = {
    'topProductsTable': { key: 'Revenue', direction: 'desc', sourceKey: 'allProducts' },
    'reorderTable': { key: 'StockLevel', direction: 'asc', sourceKey: 'allProducts' },
    'productProfitTable': { key: 'Margin', direction: 'desc', sourceKey: 'allProducts' },
    'abcAnalysisTable': { key: 'Revenue', direction: 'desc', sourceKey: 'abcProducts' },
};


// --- START: 10,000-Row Sample Data Generator ---

/**
 * Generates a large synthetic retail sales dataset with 10,000 rows.
 * @param {number} numRows The number of rows to generate.
 * @returns {Array<Object>} The generated sample data array.
 */
const generateLargeSampleData = (numRows) => {
    const data = [];
    const categories = {
        Electronics: {
            basePrice: 500, products: ['Laptop Pro', 'Smartphone X', 'Wireless Earbuds', 'Smart Watch', '4K Monitor', 'Gaming Console', 'Portable Speaker', 'Drone', 'E-Reader', 'External SSD']
        },
        Apparel: {
            basePrice: 40, products: ['T-Shirt (S)', 'Jeans (M)', 'Winter Jacket', 'Running Shoes', 'Formal Dress', 'Baseball Cap', 'Socks (Pack)', 'Sunglasses', 'Casual Hoodie', 'Leather Belt']
        },
        Home_Goods: {
            basePrice: 80, products: ['Coffee Maker', 'Blender', 'Smart Bulb (4-Pack)', 'Vacuum Cleaner', 'Frying Pan Set', 'Area Rug', 'Wall Clock', 'Digital Scale', 'Glassware Set', 'Bedding Set']
        },
        Furniture: {
            basePrice: 450, products: ['Desk Chair', 'Standing Desk', 'Side Table', 'Bookshelf', 'Sofa Bed', 'Dining Table', 'Outdoor Patio Set', 'Lamp', 'TV Stand', 'Storage Ottoman']
        },
        Books: {
            basePrice: 20, products: ['Bestseller Novel', 'Cookbook', 'Self-Help Guide', 'Sci-Fi Classic', 'Art History Book', 'Business Strategy', 'Childrens Book', 'Fantasy Series', 'Poetry Anthology', 'Travel Guide']
        }
    };
    
    let orderIDCounter = 10000;
    const startDate = new Date('2023-10-01');
    const endDate = new Date('2024-10-01');
    const dateRangeMs = endDate.getTime() - startDate.getTime();

    // Initial Stock levels (must be high for low stock logic to work later)
    const initialStock = {};

    for (let i = 0; i < numRows; i++) {
        const categoryKey = Object.keys(categories)[Math.floor(Math.random() * Object.keys(categories).length)];
        const categoryData = categories[categoryKey];
        const product = categoryData.products[Math.floor(Math.random() * categoryData.products.length)];
        const basePrice = categoryData.basePrice;

        // Generate sales date within the one-year range
        const randomTime = startDate.getTime() + Math.floor(Math.random() * dateRangeMs);
        const orderDate = new Date(randomTime);

        // Define sales and cost
        const sales = basePrice * (1 + (Math.random() * 0.4 - 0.2)); // Price variance +/- 20%
        const costFactor = 0.5 + Math.random() * 0.25; // Margin between 25% and 50%
        const cost = sales * costFactor;
        
        const quantity = Math.floor(Math.random() * 3) + 1; // 1 to 3 items per transaction, mostly 1

        // Inventory simulation: StockLevel and ReorderPoint are typically static per product
        if (!initialStock[product]) {
            initialStock[product] = {
                stock: 200 + Math.floor(Math.random() * 100), // High initial stock
                reorderPoint: Math.floor(basePrice / 10) + 5 // ROP is scaled by price
            };
            // Force a few products to be low stock initially for the reorder alert table
            if (i % 2000 === 0) { // Every 2000th product will be low stock
                 initialStock[product].stock = 5;
            }
        }
        
        // This is a simplified inventory snapshot model where stock is a lookup, not calculated
        const stockLevel = initialStock[product].stock; 
        const reorderPoint = initialStock[product].reorderPoint;
        
        // Simulating multiple items in one transaction (OrderID)
        if (Math.random() < 0.3) {
            orderIDCounter++;
        }

        data.push({
            Product: product,
            Category: categoryKey.replace('_', ' '),
            // Format as DD/MM/YYYY to test the robust date parser
            OrderDate: orderDate.toLocaleDateString('en-GB'), 
            Sales: parseFloat(sales.toFixed(2)),
            Cost: parseFloat(cost.toFixed(2)),
            StockLevel: stockLevel,
            ReorderPoint: reorderPoint,
            Quantity: quantity,
            OrderID: `ORD${orderIDCounter}`,
        });
    }

    // Sort by date to make trends look better
    data.sort((a, b) => new Date(a.OrderDate.split('/').reverse().join('-')) - new Date(b.OrderDate.split('/').reverse().join('-')));
    
    return data;
};


// The small sample data array is replaced by the generator call
const sampleRetailData = generateLargeSampleData(10000); 

// --- END: 10,000-Row Sample Data Generator ---

// --- Mutable Variables for new metrics requiring external data (DYNAMIC PROXIES) ---
let STORE_SQUARE_FOOTAGE = 5000; // Default Value for SPSF calculation
let TOTAL_FOOT_TRAFFIC = 5000;    // Default Value for Conversion Rate (CR) calculation


// --- Static Data for Guide (UPDATED to include all suggested metrics) ---
const REQUIRED_COLUMNS = [
    { key: 'Product', desc: 'Unique item name', tooltip: 'Used for aggregation and product-level analysis (e.g., Top 10).' },
    { key: 'Category', desc: 'Product grouping', tooltip: 'Used to segment sales and profit (e.g., Sales by Category chart).' },
    { key: 'OrderDate', desc: 'Date of sale', tooltip: 'Crucial for calculating sales trends over time and time-series analysis.' },
    { key: 'Sales', desc: 'Selling price', tooltip: 'Used to calculate Total Revenue and Average Order Value.' },
    { key: 'Cost', desc: 'Cost price', tooltip: 'Used to calculate Profit (Sales - Cost) and Profit Margin.' },
    { key: 'StockLevel', desc: 'Current units in stock', tooltip: 'The most recent inventory count. Used for Low Stock alerts.' },
    { key: 'ReorderPoint', desc: 'Stock level trigger', tooltip: 'The minimum safe level. Used to flag Low Stock alerts (StockLevel < ReorderPoint).' },
    { key: 'Quantity', desc: 'Units sold in the transaction', tooltip: 'Used for total units sold and calculating average cost/revenue per unit.' }
];

const KPI_GUIDE = [
    { kpi: 'Total Revenue', formula: 'SUM(Sales)', definition: 'The total value of goods sold.', useCase: 'Measures overall business performance and scale.' },
    { kpi: 'Total Profit', formula: 'SUM(Sales - Cost)', definition: 'The money left after subtracting cost of goods sold.', useCase: 'The most direct measure of business health.' },
    { kpi: 'AOV (Avg. Order Value / ATV)', formula: 'Total Revenue / Total Orders', definition: 'The average amount spent each time a customer places an order.', useCase: 'Helps determine customer spending habits and pricing strategy efficiency.' },
    { kpi: 'Profit Margin', formula: 'Total Profit / Total Revenue', definition: 'The percentage of revenue that turns into profit.', useCase: 'Measures the efficiency and profitability of product pricing.' },
    { kpi: 'Inventory Turnover Rate (ITR)', formula: 'Total COGS / Avg. Inventory Cost', definition: 'Measures how many times inventory is sold and replaced over a period.', useCase: 'Assesses inventory management efficiency; higher is usually better.' },
    { kpi: 'GMROI (Gross Margin Return on Investment)', formula: 'Gross Profit / Avg. Inventory Cost', definition: 'Calculates the dollar return for every dollar invested in inventory.', useCase: 'The key metric for merchandisers to evaluate product profitability and buying efficiency.' },
    { kpi: 'Customer Lifetime Value (CLV)', formula: 'Avg. Customer Revenue (Proxy)', definition: 'The average revenue generated by a unique customer/order group.', useCase: 'Determines the sustainable budget for customer acquisition (CAC).' },
    { kpi: 'Sales Per Square Foot (SPSF)', formula: 'Total Revenue / Selling Space (sq ft)', definition: 'Revenue generated for each square foot of selling space.', useCase: 'Evaluates the productivity and profitability of physical retail space.' },
    { kpi: 'Conversion Rate (CR)', formula: 'Total Transactions / Total Traffic', definition: 'The percentage of visitors who complete a purchase.', useCase: 'Measures the efficiency of store layout or website design.' },
    { kpi: 'Sell-Through Rate (STR)', formula: 'Units Sold / (Units Sold + Remaining Units)', definition: 'The percentage of a product batch sold within a given period.', useCase: 'Critical for short-term markdown and ordering decisions.' },
    { kpi: 'ABC Segment (A, B, C)', formula: 'Cumulative Revenue % (80/15/5)', definition: 'Categorizes products based on their revenue contribution (A=Top 80%, B=Next 15%, C=Remaining 5%).', useCase: 'Prioritizes inventory and merchandising efforts.' },
    { kpi: 'Reorder Alert', formula: 'StockLevel < ReorderPoint', definition: 'A flag that triggers when current stock drops below the defined minimum.', useCase: 'Critical for avoiding stockouts.' }
];


// --- Utility Functions (Including downloadCSV) ---

/** Formats a number as currency. */
const formatCurrency = (value, short = false) => {
    return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: short ? 0 : 2 }).format(value);
};

/** Formats a number as a percentage. */
const formatPercentage = (value) => {
    return new Intl.NumberFormat('en-US', { style: 'percent', minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(value);
};

/** Formats a number with thousands separators. */
const formatNumber = (value) => {
    return new Intl.NumberFormat('en-US').format(Math.round(value));
};

/** Parses date strings, specifically handling DD/MM/YYYY. */
const parseDate = (dateStr) => {
    let date = new Date(dateStr);
    if (!isNaN(date)) return date;

    if (typeof dateStr === 'string') {
        const parts = dateStr.split(/[-\/.]/);
        if (parts.length === 3) {
            // Check for potential DD/MM/YYYY format assumption (DD/MM/YYYY)
            const d = parseInt(parts[0], 10);
            const m = parseInt(parts[1], 10);
            const y = parseInt(parts[2], 10);
            // Re-order to YYYY-MM-DD for standard JavaScript Date parsing
            const dateISO = `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
            date = new Date(dateISO);
            if (!isNaN(date)) return date;
        }
    }
    return new Date('Invalid Date');
};

/** Formats a Date object to YYYY-MM-DD for use in input[type=date] */
const formatDateForInput = (date) => {
    if (!date || isNaN(date.getTime())) return '';
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
};

// Data Validation Function
/**
 * Validates the raw data structure against required columns.
 * @param {Array<Object>} data The raw data array.
 * @returns {Array<string>} An array of error messages.
 */
const validateDataStructure = (data) => {
    if (!data || data.length === 0) {
        return ["Data file is empty or could not be parsed."];
    }
    
    const errors = [];
    const firstRow = data[0];
    const presentKeys = Object.keys(firstRow).map(k => k.trim());
    
    const REQUIRED_KEYS = [
        { keys: ['Product', 'SKU'], type: 'string', name: 'Product' },
        { keys: ['Category', 'Dept'], type: 'string', name: 'Category' },
        { keys: ['OrderDate', 'Date', 'TransactionDate'], type: 'date', name: 'OrderDate' },
        { keys: ['Sales', 'Revenue', 'Price'], type: 'number', name: 'Sales' },
        { keys: ['Cost', 'CostPrice', 'COGS'], type: 'number', name: 'Cost' },
        { keys: ['StockLevel', 'Stock', 'Inventory'], type: 'number', name: 'StockLevel' },
        { keys: ['ReorderPoint', 'ROP'], type: 'number', name: 'ReorderPoint' },
        { keys: ['Quantity', 'Units'], type: 'number', name: 'Quantity' },
        { keys: ['OrderID', 'TransactionID'], type: 'string', name: 'OrderID' }, // Added OrderID to required check
    ];

    REQUIRED_KEYS.forEach(req => {
        const found = req.keys.some(k => presentKeys.includes(k.trim()));
        if (!found) {
            errors.push(`Missing required field: ${req.name}. Please ensure your data has one of these column headers: ${req.keys.join(', ')}.`);
        }
    });

    if (errors.length === 0) {
        // Sample data type check (first 10 rows for date consistency)
        const dateKeys = REQUIRED_KEYS.find(r => r.name === 'OrderDate').keys;
        const dateKey = dateKeys.find(k => presentKeys.includes(k.trim()));
        const checkCount = Math.min(10, data.length);
        
        for (let i = 0; i < checkCount; i++) {
            const item = data[i];
            if (dateKey && isNaN(parseDate(item[dateKey]).getTime())) {
                 errors.push(`Data type error: 'OrderDate' value in row ${i+2} ('${item[dateKey]}') is invalid. Check formatting (e.g., DD/MM/YYYY).`);
                 break;
            }
        }
    }
    
    return errors;
};


/** Creates and triggers a download for a CSV file. */
const downloadCSV = (data, filename) => {
    if (data.length === 0) return;
    const header = Object.keys(data[0]);
    const csv = [
        header.join(','),
        ...data.map(row => header.map(fieldName => JSON.stringify(row[fieldName] ?? '')).join(','))
    ].join('\n');

    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    if (link.download !== undefined) { 
        const url = URL.createObjectURL(blob);
        link.setAttribute("href", url);
        link.setAttribute("download", filename);
        link.style.visibility = 'hidden';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }
}


/** Populates the tables in the new dropdown panel, including tooltips. */
const displayDataGuide = () => {
    // Required Columns Table
    DOM.requiredColumnsTable.innerHTML = `<thead><tr><th>Field</th><th>Description</th></tr></thead><tbody>` +
        REQUIRED_COLUMNS.map(c => `
            <tr class="tooltip-container">
                <td>${c.key}</td>
                <td>${c.desc}</td>
                <div class="tooltip-text">
                    <strong>Logic/Use:</strong> ${c.tooltip}
                </div>
            </tr>
        `).join('') + `</tbody>`;

    // KPI Guide Table
    DOM.kpiGuideTable.innerHTML = `<thead><tr><th>KPI</th><th>Formula</th></tr></thead><tbody>` +
        KPI_GUIDE.map(k => `
            <tr class="tooltip-container">
                <td>${k.kpi}</td>
                <td>${k.formula}</td>
                <div class="tooltip-text">
                    <strong>Definition:</strong> ${k.definition}<br>
                    <strong>Use Case:</strong> ${k.useCase}
                </div>
            </tr>
        `).join('') + `</tbody>`;
};

/** Toggles the new Data Guide dropdown panel's open/closed state. */
const toggleDataGuide = () => {
    DOM.guideDropdown.classList.toggle('open');
    const isExpanded = DOM.guideDropdown.classList.contains('open');
    DOM.dataGuideBtn.setAttribute('aria-expanded', isExpanded);
    // Close the Palette dropdown
    DOM.paletteDropdown.classList.remove('open');
    DOM.paletteBtn.setAttribute('aria-expanded', false);
};

/** Toggles the Palette dropdown panel's open/closed state. */
const togglePaletteGuide = () => {
    DOM.paletteDropdown.classList.toggle('open');
    const isExpanded = DOM.paletteDropdown.classList.contains('open');
    DOM.paletteBtn.setAttribute('aria-expanded', isExpanded);
    // Close the Data Guide dropdown
    DOM.guideDropdown.classList.remove('open');
    DOM.dataGuideBtn.setAttribute('aria-expanded', false);
};

/**
 * Extracts unique categories and products from the raw data and populates the filter options.
 */
const populateFilterOptions = (data) => {
    const categories = new Set();
    const products = new Set();
    let minDate = null;
    let maxDate = null;

    data.forEach(item => {
        // FIX: Ensure Category and Product names are trimmed before adding to the Set
        if (item.Category) categories.add(String(item.Category).trim());
        if (item.Product) products.add(String(item.Product).trim());

        // Date has already been parsed and mapped in startAnalysisAndRender
        const date = item.OrderDate;
        if (date && !isNaN(date.getTime())) {
            if (!minDate || date < minDate) minDate = date;
            if (!maxDate || date > maxDate) maxDate = date;
        }
    });

    // Date Filters
    if (!currentFilters.startDate) DOM.startDateFilter.value = minDate ? formatDateForInput(minDate) : '';
    if (!currentFilters.endDate) DOM.endDateFilter.value = maxDate ? formatDateForInput(maxDate) : '';

    // Category Filter
    DOM.categoryFilter.innerHTML = [...categories].sort().map(c => `<option value="${c}">${c}</option>`).join('');
    // Product Filter
    DOM.productFilter.innerHTML = [...products].sort().map(p => `<option value="${p}">${p}</option>`).join('');
    
    // Restore previous selections
    if (currentFilters.categories.length > 0) {
        Array.from(DOM.categoryFilter.options).forEach(option => {
            option.selected = currentFilters.categories.includes(option.value);
        });
    }
    if (currentFilters.products.length > 0) {
        Array.from(DOM.productFilter.options).forEach(option => {
            option.selected = currentFilters.products.includes(option.value);
        });
    }
};

/**
 * Reads the numerical input fields and updates the global proxy variables.
 */
const updateProxyInputs = () => {
    // Read and validate Square Footage
    const sqFtValue = parseFloat(DOM.sqFootageInput.value);
    if (!isNaN(sqFtValue) && sqFtValue > 0) {
        STORE_SQUARE_FOOTAGE = sqFtValue;
    } else {
        STORE_SQUARE_FOOTAGE = 1; // Prevent division by zero, use 1 if invalid
        DOM.sqFootageInput.value = 1; 
    }

    // Read and validate Foot Traffic
    const trafficValue = parseInt(DOM.trafficInput.value);
    if (!isNaN(trafficValue) && trafficValue >= 0) {
        TOTAL_FOOT_TRAFFIC = trafficValue;
    } else {
        TOTAL_FOOT_TRAFFIC = 0; // Use 0 if invalid
        DOM.trafficInput.value = 0;
    }
}

// UPDATED: applyFilters function to handle chart clicks
/**
 * Filters the MAPPED data based on currentFilters and re-renders the dashboard.
 * @param {Object} [newFilter={}] Optional filter object from chart click {type: 'category', value: 'Electronics'}
 */
const applyFilters = (newFilter = {}) => {
    DOM.statusMessage.textContent = "Applying filters...";
    
    // 1. Capture current filter state (from inputs)
    const start = DOM.startDateFilter.value ? parseDate(DOM.startDateFilter.value) : null;
    const end = DOM.endDateFilter.value ? parseDate(DOM.endDateFilter.value) : null;
    
    // Use TRIMMED values from the options list
    let categories = Array.from(DOM.categoryFilter.selectedOptions).map(opt => opt.value.trim());
    let products = Array.from(DOM.productFilter.selectedOptions).map(opt => opt.value.trim());

    // 2. Apply optional new filter (from chart click)
    if (newFilter.type === 'category' && newFilter.value) {
        const trimmedNewValue = String(newFilter.value).trim();
        // Toggle the category: If already selected AND it's the only one, clear. If not, select ONLY this one.
        if (categories.includes(trimmedNewValue) && categories.length === 1) {
            categories = []; // Clear filter if it's the only one selected
        } else {
            categories = [trimmedNewValue]; // Select only the clicked category
        }
    }
    
    // 3. Update global state
    currentFilters = { startDate: start, endDate: end, categories: categories, products: products };
    
    // 4. Update UI to reflect new state
    Array.from(DOM.categoryFilter.options).forEach(option => {
        option.selected = currentFilters.categories.includes(option.value.trim());
    });
    
    // 5. Trigger the analysis pipeline
    if (rawData.length > 0) {
        startAnalysisAndRender(rawData, true); 
    } else {
        DOM.statusMessage.textContent = "Please upload a data file first.";
    }
};

/**
 * Resets all filter inputs and re-renders the full dashboard.
 */
const clearFilters = () => {
    // Clear the filter state
    currentFilters = { startDate: null, endDate: null, categories: [], products: [] };
    
    // Reset inputs
    DOM.startDateFilter.value = '';
    DOM.endDateFilter.value = '';
    
    // Reset selections on multi-selects
    Array.from(DOM.categoryFilter.options).forEach(option => option.selected = false);
    Array.from(DOM.productFilter.options).forEach(option => option.selected = false);

    // Reset Proxy Inputs to defaults
    DOM.sqFootageInput.value = 5000;
    DOM.trafficInput.value = 5000;
    
    // Repopulate options to restore date range to max possible and trigger analysis
    if (rawData.length > 0) {
        // Pass a version of rawData where OrderDate is already a Date object
        const dataWithDates = rawData.map(item => ({...item, OrderDate: parseDate(item.OrderDate)}));
        populateFilterOptions(dataWithDates); 
        startAnalysisAndRender(rawData, true);
    } else {
        DOM.statusMessage.textContent = "Filters cleared.";
    }
};

/**
 * Toggles the type of the Sales Trends Chart and re-renders that chart.
 */
const toggleChartType = (chartId, newType) => {
    if (currentChartTypes[chartId] === newType) return; // Already active
    
    // Update active state in UI
    DOM.chartToggleBtns.forEach(btn => {
        if (btn.dataset.chart === chartId) {
            btn.classList.toggle('active', btn.dataset.type === newType);
        }
    });

    // Update global state
    currentChartTypes[chartId] = newType;
    
    // Re-render the charts with the new type
    if (analysisResults) {
        // Only re-render the charts, not the entire pipeline
        renderCharts(analysisResults.coreMetrics, analysisResults.productData);
    }
};


/**
 * Handles file upload, triggers initial analysis, and populates filters.
 */
const handleFileUpload = (event) => {
    const file = event.target.files[0];
    if (file) {
        DOM.runAnalysisBtn.disabled = true;
        DOM.runAnalysisBtn.textContent = 'Processing...';

        resetApplication(false); 
        
        startAnalysisAndRender(file);
    } else {
        DOM.runAnalysisBtn.disabled = false;
        DOM.runAnalysisBtn.textContent = 'Clear Data & New File';
    }
};

/**
 * Executes the entire analysis pipeline with the provided data (either file or sample).
 * @param {File | Object} source - The File object or the array of sample data/full rawData.
 * @param {boolean} isFilterRun - True if this is a filter-triggered re-render.
 */
const startAnalysisAndRender = async (source, isFilterRun = false) => {
    // UI State: Start
    DOM.loader.classList.remove('hidden');
    DOM.runAnalysisBtn.disabled = true;
    DOM.dashboardContainer.classList.add('hidden');
    DOM.runtimeDisplay.textContent = "";

    const startTime = performance.now();
    let dataToProcess = [];

    try {
        if (source instanceof File) {
            DOM.statusMessage.textContent = "Reading file...";
            dataToProcess = await processFile(source);
            
            // NEW: Validate data structure immediately after processing file
            const validationErrors = validateDataStructure(dataToProcess);
            if (validationErrors.length > 0) {
                // Throw combined error message
                throw new Error("\n" + validationErrors.join('\n'));
            }
            
            rawData = dataToProcess; 
            
            // Populate filters only on initial file read
            populateFilterOptions(dataToProcess.map(item => ({...item, OrderDate: parseDate(item.OrderDate)})));
        } else if (Array.isArray(source)) {
            dataToProcess = source;
            // If it's the sample data load, populate filters
            if (!isFilterRun) {
                DOM.statusMessage.textContent = "Loading sample data dashboard...";
                
                // NEW: Validate sample data structure
                const validationErrors = validateDataStructure(dataToProcess);
                if (validationErrors.length > 0) {
                    throw new Error("Sample Data Validation Failed: \n" + validationErrors.join('\n'));
                }
                
                rawData = dataToProcess; 
                // Populate filters requires OrderDate to be a Date object
                populateFilterOptions(dataToProcess.map(item => ({...item, OrderDate: parseDate(item.OrderDate)})));
            } else {
                DOM.statusMessage.textContent = "Re-analyzing filtered data...";
            }
        } else {
            throw new Error("Invalid data source provided.");
        }
        
        // --- NEW STEP: Read User Proxy Inputs before analysis ---
        updateProxyInputs(); 
        // --------------------------------------------------------

        // Apply filters to the raw data *before* staged analysis starts
        const filteredData = filterData(dataToProcess);
        
        if (filteredData.length === 0) {
            throw new Error("No data remains after applying current filters. Please clear filters and try again.");
        }

        DOM.statusMessage.textContent = "Data processed. Starting multi-stage analysis...";
        await stagedAnalysis(filteredData);

        // UI State: Success
        DOM.dashboardContainer.classList.remove('hidden');
        DOM.statusMessage.textContent = "Analysis complete! Dashboard updated.";
    } catch (error) {
        // UI State: Error
        DOM.statusMessage.textContent = `Analysis failed. Please fix errors below:\n${error.message}`;
        DOM.dashboardContainer.classList.add('hidden'); // Hide dashboard on error
        console.error("Analysis Error:", error);
    } finally {
        // UI State: End
        DOM.loader.classList.add('hidden');
        const endTime = performance.now();
        DOM.runtimeDisplay.textContent = `Runtime: ${(endTime - startTime).toFixed(2)}ms`;
        
        DOM.runAnalysisBtn.textContent = 'Clear Data & New File';
        DOM.runAnalysisBtn.disabled = false;
    }
};

/**
 * Filters the raw data based on the current global filter state.
 */
const filterData = (data) => {
    let filtered = data;
    const filters = currentFilters;

    if (filters.startDate) {
        const startTimestamp = filters.startDate.getTime();
        // The data passed here is the original raw data (before stagedAnalysis maps it)
        filtered = filtered.filter(item => parseDate(item.OrderDate).getTime() >= startTimestamp);
    }
    if (filters.endDate) {
        const endTimestamp = filters.endDate.getTime();
        const inclusiveEnd = new Date(filters.endDate);
        inclusiveEnd.setDate(inclusiveEnd.getDate() + 1);
        filtered = filtered.filter(item => parseDate(item.OrderDate).getTime() < inclusiveEnd.getTime());
    }

    if (filters.categories.length > 0) {
        // FIX: Trim item.Category for robust matching
        filtered = filtered.filter(item => {
            const itemCategory = String(item.Category || '').trim();
            return filters.categories.includes(itemCategory);
        });
    }

    if (filters.products.length > 0) {
        // FIX: Trim item.Product for robust matching
        filtered = filtered.filter(item => {
            const itemProduct = String(item.Product || '').trim();
            return filters.products.includes(itemProduct);
        });
    }
    
    return filtered;
};


/**
 * Main button handler: REPURPOSED to clear data and reset the file input.
 */
const runAnalysisBtn = () => {
    resetApplication(true);
    DOM.statusMessage.textContent = 'Dashboard cleared. Ready for new file upload.';
    DOM.runAnalysisBtn.textContent = 'Clear Data & New File'; 
    DOM.runAnalysisBtn.disabled = true;
};

/** Processes the uploaded Excel file using SheetJS. */
const processFile = (file) => {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                // Use header:1 to get the first row for better column name handling in validation
                const workbook = XLSX.read(data, { type: 'array' });
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                // Use raw: true to keep date/number types as they appear if possible
                const json = XLSX.utils.sheet_to_json(worksheet, { raw: true }); 
                resolve(json);
            } catch (error) {
                reject(new Error("Failed to parse Excel file. It may be corrupted or not a valid XLSX/CSV format."));
            }
        };

        reader.onerror = (error) => reject(new Error("File reading failed."));
        reader.readAsArrayBuffer(file);
    });
};

/** * Resets the application state and clears UI elements. (Modified to handle filter reset)
 * @param {boolean} fullReset - True if clearing file data too (i.e., not just a filter clear).
 */
const resetApplication = (fullReset = true) => {
    // Clear data and state
    if (fullReset) {
        rawData = [];
        DOM.excelFile.value = '';
        
        // Reset Proxy Variables to default (in memory)
        STORE_SQUARE_FOOTAGE = 5000;
        TOTAL_FOOT_TRAFFIC = 5000;
        
        // Reset Proxy Input Fields (UI)
        DOM.sqFootageInput.value = 5000;
        DOM.trafficInput.value = 5000;
    }
    
    analysisResults = null;

    // Reset UI
    DOM.dashboardContainer.classList.add('hidden');
    DOM.runtimeDisplay.textContent = '';
    
    // Clear metrics and tables
    DOM.metricsContainer.innerHTML = '';
    DOM.inventoryMetricsContainer.innerHTML = '';
    DOM.topProductsTable.innerHTML = '';
    DOM.reorderTable.innerHTML = '';
    DOM.productProfitTable.innerHTML = '';
    DOM.abcAnalysisTable.innerHTML = ''; // NEW

    // Close the Data Guide and Palette dropdowns if open
    DOM.guideDropdown.classList.remove('open'); 
    DOM.dataGuideBtn.setAttribute('aria-expanded', false);
    DOM.paletteDropdown.classList.remove('open'); 
    DOM.paletteBtn.setAttribute('aria-expanded', false);

    // Destroy existing charts
    Object.values(chartInstances).forEach(chart => chart.destroy());
    chartInstances = {};

    // Reset Filters and Toggles (calls clearFilters which resets the input fields)
    clearFilters();
    currentChartTypes = { salesTrendsChart: 'bar' };
    DOM.chartToggleBtns.forEach(btn => btn.classList.remove('active'));
    document.querySelector('.chart-toggle-btn[data-type="bar"]').classList.add('active');

    // Reset Sort Icons and state
    document.querySelectorAll('.data-table thead th').forEach(th => {
        th.classList.remove('sort-asc', 'sort-desc');
        const icon = th.querySelector('i');
        if (icon) icon.className = 'fas fa-sort';
    });
    // Reset Sort State to default views
    currentSortState = {
        'topProductsTable': { key: 'Revenue', direction: 'desc', sourceKey: 'allProducts' },
        'reorderTable': { key: 'StockLevel', direction: 'asc', sourceKey: 'allProducts' },
        'productProfitTable': { key: 'Margin', direction: 'desc', sourceKey: 'allProducts' },
        'abcAnalysisTable': { key: 'Revenue', direction: 'desc', sourceKey: 'abcProducts' },
    };
};

/** Changes the active CSS color palette. */
const changePalette = (paletteName) => {
    document.body.className = paletteName;
    // Close the dropdown after selection
    DOM.paletteDropdown.classList.remove('open');
    DOM.paletteBtn.setAttribute('aria-expanded', false);
};

// --- B. Data Processing & Analysis (Multi-Stage) ---
/** Multi-stage analysis pipeline using setTimeout to allow UI updates. */
const stagedAnalysis = (data) => {
    return new Promise((resolve) => {
        let mappedData = [];

        // 1. Defensive Data Mapping (Initial Filter & Clean)
        const mapStage = () => {
            DOM.statusMessage.textContent = "Stage 0/3: Mapping and cleaning data...";
            
            // Define all possible keys for robust mapping/inference
            const KEY_MAP = {
                Product: ['Product', 'SKU'],
                Category: ['Category', 'Dept'],
                OrderDate: ['OrderDate', 'Order Date', 'Date', 'TransactionDate'],
                Sales: ['Sales', 'Revenue', 'Price'],
                Cost: ['Cost', 'CostPrice', 'COGS'],
                StockLevel: ['StockLevel', 'Stock', 'Inventory'],
                ReorderPoint: ['ReorderPoint', 'ROP'],
                Quantity: ['Quantity', 'Units'],
                OrderID: ['OrderID', 'InvoiceID', 'TransactionID'],
            };

            const findKey = (item, keys) => keys.find(k => item.hasOwnProperty(k));

            mappedData = data.map(item => {
                // Find the actual keys used in the dataset
                const productKey = findKey(item, KEY_MAP.Product);
                const categoryKey = findKey(item, KEY_MAP.Category);
                const dateKey = findKey(item, KEY_MAP.OrderDate);
                const salesKey = findKey(item, KEY_MAP.Sales);
                const costKey = findKey(item, KEY_MAP.Cost);
                const stockKey = findKey(item, KEY_MAP.StockLevel);
                const reorderKey = findKey(item, KEY_MAP.ReorderPoint);
                const quantityKey = findKey(item, KEY_MAP.Quantity);
                const orderIDKey = findKey(item, KEY_MAP.OrderID);

                // Validation check (should be mostly redundant if validateDataStructure passed, but good for safety)
                if (!salesKey || !costKey || !dateKey || !productKey || !quantityKey || !stockKey || !reorderKey) {
                    return null; 
                }

                const sales = parseFloat(item[salesKey]) || 0;
                const cost = parseFloat(item[costKey]) || 0;
                const quantity = parseInt(item[quantityKey]) || 0;

                return {
                    // Core Data (Trim non-numeric/date fields for later matching)
                    Product: String(item[productKey] || 'Unknown Product').trim(),
                    Category: String(item[categoryKey] || 'Other').trim(),
                    StockLevel: parseInt(item[stockKey]) || 0,
                    ReorderPoint: parseInt(item[reorderKey]) || 0,
                    // Calculated Metrics
                    Sales: sales,
                    Cost: cost,
                    Profit: sales - cost,
                    Quantity: quantity,
                    OrderDate: parseDate(item[dateKey]), 
                    OrderID: String(item[orderIDKey] || item.OrderID || 'N/A').trim(),
                };
            }).filter(item => item && !isNaN(item.OrderDate.getTime()));

            // Go to Stage 1
            setTimeout(stage1_CoreMetrics, 0);
        };
        
        let results = {};

        // 2. Stage 1 (Core Metrics)
        const stage1_CoreMetrics = () => {
            DOM.statusMessage.textContent = "Stage 1/3: Calculating core financial and customer metrics...";
            let totalRevenue = 0;
            let totalProfit = 0;
            const uniqueOrders = new Set();
            const customerTransactions = {}; // Use OrderID as proxy for CustomerID
            const salesByMonth = {};
            const salesByDay = {}; 
            const salesByCategory = {};

            mappedData.forEach(item => {
                totalRevenue += item.Sales;
                totalProfit += item.Profit;

                // Use OrderID as a customer proxy
                const customerProxyID = item.OrderID || 'N/A'; 
                if (!customerTransactions[customerProxyID]) {
                    customerTransactions[customerProxyID] = { revenue: 0, orders: new Set() };
                }
                customerTransactions[customerProxyID].revenue += item.Sales;
                customerTransactions[customerProxyID].orders.add(item.OrderID); // Track orders by this 'customer'

                if (item.OrderID && item.OrderID !== 'N/A') uniqueOrders.add(item.OrderID);
                else uniqueOrders.add(`${item.OrderDate.toISOString()}_${item.Product}`);

                // Sales Trends 
                const monthKey = item.OrderDate.toISOString().substring(0, 7);
                salesByMonth[monthKey] = (salesByMonth[monthKey] || 0) + item.Sales;
                const dayKey = item.OrderDate.toISOString().split('T')[0];
                salesByDay[dayKey] = (salesByDay[dayKey] || 0) + item.Sales;
                const category = item.Category;
                salesByCategory[category] = (salesByCategory[category] || 0) + item.Sales;
            });

            const totalOrders = uniqueOrders.size;
            const totalCustomersProxy = Object.keys(customerTransactions).length;
            const avgOrderValue = totalOrders > 0 ? totalRevenue / totalOrders : 0; // Also Average Transaction Value (ATV)
            const profitMargin = totalRevenue > 0 ? totalProfit / totalRevenue : 0;
            
            // NEW CUSTOMER/OPERATIONAL METRICS (NOW USING DYNAMIC GLOBALS)
            const salesPerSqFt = totalRevenue / STORE_SQUARE_FOOTAGE; // Uses Dynamic STORE_SQUARE_FOOTAGE
            const conversionRate = totalOrders > 0 && TOTAL_FOOT_TRAFFIC > 0 ? totalOrders / TOTAL_FOOT_TRAFFIC : 0; // Uses Dynamic TOTAL_FOOT_TRAFFIC
            
            // CLV Proxy: Avg. Customer Revenue 
            const totalCustomerRevenue = Object.values(customerTransactions).reduce((sum, c) => sum + c.revenue, 0);
            const clvProxy = totalCustomersProxy > 0 ? totalCustomerRevenue / totalCustomersProxy : 0;
            
            // CRR is N/A without multi-period data but included as a key concept.

            results.coreMetrics = {
                TotalRevenue: totalRevenue,
                TotalProfit: totalProfit,
                AOV: avgOrderValue, 
                ProfitMargin: profitMargin,
                TotalOrders: totalOrders,
                UniqueCustomers: totalCustomersProxy, // Proxy for unique customers
                SalesByMonth: salesByMonth,
                SalesByDay: salesByDay, 
                SalesByCategory: salesByCategory,
                // NEW METRICS
                CLV: clvProxy,
                SPSF: salesPerSqFt,
                ConversionRate: conversionRate,
            };

            // Go to Stage 2
            setTimeout(stage2_InventoryAndProducts, 0);
        };

        // NEW: ABC Analysis Function
        const calculateABCAnalysis = (productList) => {
            const sortedProducts = [...productList].sort((a, b) => b.Revenue - a.Revenue);
            const totalRevenue = sortedProducts.reduce((sum, p) => sum + p.Revenue, 0);
            let cumulativeRevenue = 0;

            const abcProducts = sortedProducts.map(p => {
                cumulativeRevenue += p.Revenue;
                const cumulativePercent = totalRevenue > 0 ? cumulativeRevenue / totalRevenue : 0;
                
                let segment = 'C';
                if (cumulativePercent <= 0.80) {
                    segment = 'A';
                } else if (cumulativePercent <= 0.95) { // 80% to 95%
                    segment = 'B';
                }
                
                return {
                    ...p,
                    CumulativeRevenue: cumulativeRevenue,
                    CumulativePercent: cumulativePercent,
                    Segment: segment,
                };
            });
            return abcProducts;
        };


        // 3. Stage 2 (Inventory/Products)
        const stage2_InventoryAndProducts = () => {
            DOM.statusMessage.textContent = "Stage 2/3: Analyzing inventory and product performance...";

            const productSummary = {};
            let totalStockUnits = 0;
            let totalCOGS = 0; 
            let endingInventoryCost = 0; 

            mappedData.forEach(item => {
                const product = item.Product;
                const costOfSale = item.Cost * item.Quantity;
                totalCOGS += costOfSale; 

                if (!productSummary[product]) {
                    productSummary[product] = {
                        Revenue: 0,
                        Profit: 0,
                        QuantitySold: 0,
                        TotalCOGS: 0, 
                        StockLevel: 0,
                        ReorderPoint: 0,
                        Category: item.Category, // Store category for ABC table
                        Margin: 0, // Placeholder
                        SellThroughRate: 0, // Placeholder
                    };
                }

                productSummary[product].Revenue += item.Sales;
                productSummary[product].Profit += item.Profit;
                productSummary[product].QuantitySold += item.Quantity;
                productSummary[product].TotalCOGS += costOfSale;
                
                // IMPORTANT: In a transactional model, the last OrderDate's StockLevel/ROP is used as the current snapshot.
                productSummary[product].StockLevel = item.StockLevel;
                productSummary[product].ReorderPoint = item.ReorderPoint;

                // Category is assumed consistent per product, but we take the last seen value
                productSummary[product].Category = item.Category; 
            });
            
            // Calculate Ending Inventory Cost (EIC) and final product list
            const productList = Object.entries(productSummary).map(([product, data]) => {
                const margin = data.Revenue > 0 ? data.Profit / data.Revenue : 0;
                
                // Calculate Avg Unit Cost (used for EIC)
                const avgUnitCost = data.QuantitySold > 0 ? data.TotalCOGS / data.QuantitySold : 0;
                
                // Calculate Product's contribution to Ending Inventory Cost (EIC)
                const productEIC = data.StockLevel * avgUnitCost;
                endingInventoryCost += productEIC; 

                totalStockUnits += data.StockLevel;

                // NEW: Sell-Through Rate (STR) Proxy: Units Sold / (Units Sold + Ending Stock)
                const totalAvailableUnits = data.QuantitySold + data.StockLevel;
                const sellThroughRate = totalAvailableUnits > 0 ? data.QuantitySold / totalAvailableUnits : 0;

                return {
                    Product: product,
                    ...data,
                    Margin: margin,
                    IsLowStock: data.StockLevel < data.ReorderPoint,
                    AvgUnitCost: avgUnitCost,
                    SellThroughRate: sellThroughRate,
                };
            });
            
            // NEW: ABC Analysis
            const abcProducts = calculateABCAnalysis(productList);

            // Calculate ITR and GMROI
            const ITR = endingInventoryCost > 0 ? totalCOGS / endingInventoryCost : 0;
            const GMROI = endingInventoryCost > 0 ? results.coreMetrics.TotalProfit / endingInventoryCost : 0;

            const topProductsByRevenue = [...productList]
                .sort((a, b) => b.Revenue - a.Revenue)
                .slice(0, 10);
            
            const topProductsByProfit = [...productList]
                .sort((a, b) => b.Profit - a.Profit)
                .slice(0, 10);

            const reorderAlerts = productList.filter(p => p.IsLowStock)
                .sort((a, b) => a.StockLevel - b.StockLevel);

            results.inventoryMetrics = {
                TotalStockUnits: totalStockUnits,
                TotalProducts: productList.length,
                InventoryTurnoverRate: ITR,
                GMROI: GMROI,
                EndingInventoryCost: endingInventoryCost, 
                AvgSellThroughRate: productList.length > 0 ? productList.reduce((sum, p) => sum + p.SellThroughRate, 0) / productList.length : 0,
            };

            results.productData = {
                topProductsByRevenue: topProductsByRevenue,
                topProductsByProfit: topProductsByProfit,
                reorderAlerts: reorderAlerts,
                allProducts: productList,
                abcProducts: abcProducts,
            };

            analysisResults = results; // Save final results

            // Go to Rendering Stage
            setTimeout(renderDashboard, 0);
        };

        // 4. Rendering Stage
        const renderDashboard = () => {
            DOM.statusMessage.textContent = "Stage 3/3: Rendering dashboard charts and tables...";
            renderMetrics(analysisResults.coreMetrics, analysisResults.inventoryMetrics);
            renderCharts(analysisResults.coreMetrics, analysisResults.productData);
            renderTables(analysisResults.productData);

            // Analysis is complete
            resolve(analysisResults);
        };

        // Start the pipeline
        mapStage();
    });
};

// --- C. Rendering ---
/** Renders all primary and inventory KPI metrics. */
const renderMetrics = (core, inventory) => {
    // 9 CORE FINANCIAL / CUSTOMER / OPERATIONAL Metrics
    const allCoreMetrics = [
        // Core Financial (6)
        { icon: 'fas fa-dollar-sign', title: 'Total Revenue', value: formatCurrency(core.TotalRevenue) },
        { icon: 'fas fa-hand-holding-usd', title: 'Total Profit', value: formatCurrency(core.TotalProfit) },
        { icon: 'fas fa-percent', title: 'Profit Margin', value: formatPercentage(core.ProfitMargin) },
        { icon: 'fas fa-calculator', title: 'Avg. Order Value (ATV)', value: formatCurrency(core.AOV) },
        { icon: 'fas fa-shopping-basket', title: 'Total Orders', value: formatNumber(core.TotalOrders) },
        { icon: 'fas fa-users', title: 'Total Customers (Proxy)', value: formatNumber(core.UniqueCustomers) },
        // New Customer/Operational (3)
        { icon: 'fas fa-heart', title: 'CLV (Customer Avg Value)*', value: formatCurrency(core.CLV) },
        { icon: 'fas fa-store', title: `Sales Per Sq Ft (${STORE_SQUARE_FOOTAGE} sq ft)`, value: formatCurrency(core.SPSF) },
        { icon: 'fas fa-handshake', title: `Conversion Rate (${TOTAL_FOOT_TRAFFIC} Traffic)`, value: formatPercentage(core.ConversionRate) },
    ].map(m => `
        <div class="metric-card">
            <span class="icon"><i class="${m.icon}"></i></span>
            <div class="value">${m.value}</div>
            <div class="title">${m.title}</div>
        </div>
    `).join('');

    DOM.metricsContainer.innerHTML = allCoreMetrics;

    // 6 INVENTORY Metrics
    const inventoryMetricsHtml = [
        { icon: 'fas fa-cubes', title: 'Total Stock Units', value: formatNumber(inventory.TotalStockUnits) },
        { icon: 'fas fa-box-open', title: 'Total Product SKUs', value: formatNumber(inventory.TotalProducts) },
        { icon: 'fas fa-sync-alt', title: 'Inventory Turnover Rate (ITR)', value: inventory.InventoryTurnoverRate.toFixed(2) + 'x' }, 
        { icon: 'fas fa-chart-line', title: 'GMROI (Return on $1 Cost)', value: formatCurrency(inventory.GMROI) }, 
        { icon: 'fas fa-box-full', title: 'Ending Inventory Cost (EIC)', value: formatCurrency(inventory.EndingInventoryCost) }, 
        { icon: 'fas fa-percentage', title: 'Avg. Sell-Through Rate (STR)', value: formatPercentage(inventory.AvgSellThroughRate) }, 
    ].map(m => `
        <div class="metric-card">
            <span class="icon"><i class="${m.icon}"></i></span>
            <div class="value">${m.value}</div>
            <div class="title">${m.title}</div>
        </div>
    `).join('');

    DOM.inventoryMetricsContainer.innerHTML = inventoryMetricsHtml;
};


/** Renders the Chart.js graphs. */
const renderCharts = (core, productData) => {
    // Destroy previous chart instances
    Object.values(chartInstances).forEach(chart => chart.destroy());

    // --- Chart 1: Sales Trends Over Time (Dynamic Type) ---
    const chartType = currentChartTypes.salesTrendsChart;
    const salesLabels = Object.keys(chartType === 'line' ? core.SalesByDay : core.SalesByMonth).sort();
    const salesValues = salesLabels.map(key => chartType === 'line' ? core.SalesByDay[key] : core.SalesByMonth[key]);

    // NEW: Custom label formatter for X-axis (shows Month/Year or Day/Month/Year)
    const xLabels = salesLabels.map(key => {
        const date = new Date(key);
        if (chartType === 'line') {
            // Daily: Show Month/Day/Year for clear time visibility
            return date.toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' }); 
        } else {
            // Monthly: Show Month/Year
            return date.toLocaleDateString('en-US', { year: 'numeric', month: 'short' }); 
        }
    });

    const salesTrendsConfig = {
        type: chartType, // DYNAMIC TYPE
        data: {
            labels: xLabels,
            datasets: [{
                label: (chartType === 'line' ? 'Daily Revenue Trend' : 'Monthly Revenue Comparison'),
                data: salesValues,
                borderColor: getComputedStyle(document.body).getPropertyValue('--primary-color'),
                backgroundColor: (chartType === 'line' ? 'rgba(0, 121, 107, 0.1)' : getComputedStyle(document.body).getPropertyValue('--primary-color')),
                borderWidth: chartType === 'line' ? 2 : 1,
                pointRadius: chartType === 'line' ? 3 : 0, // Show dots for line chart
                tension: chartType === 'line' ? 0.4 : 0.1,
                fill: chartType === 'line', // Fill area under line chart
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            // START: FIX FOR AXIS LABEL CUT-OFF
            layout: {
                padding: {
                    bottom: 30 /* Add explicit bottom padding for X-axis labels */
                }
            },
            // END: FIX FOR AXIS LABEL CUT-OFF
            scales: {
                x: { grid: { display: false } },
                y: { beginAtZero: true, ticks: { callback: (val) => formatCurrency(val, true) } }
            },
            plugins: { 
                legend: { display: false },
                // NEW: Tooltip for full time data visibility
                tooltip: { 
                    callbacks: {
                        // Show full date in tooltip title
                        title: (context) => {
                            // Get the original YYYY-MM-DD or YYYY-MM key
                            const originalKey = salesLabels[context[0].dataIndex]; 
                            const date = new Date(originalKey);
                            return date.toLocaleDateString('en-US', { weekday: 'short', year: 'numeric', month: 'long', day: 'numeric' });
                        },
                        label: (context) => `Revenue: ${formatCurrency(context.parsed.y)}`
                    }
                }
            }
        }
    };
    chartInstances.salesTrends = new Chart(DOM.salesTrendsChart, salesTrendsConfig);


    // --- Chart 2: Sales by Category (Horizontal Bar Chart) ---
    const categoryDataKeys = Object.keys(core.SalesByCategory);
    const categoryDataValues = categoryDataKeys.map(key => core.SalesByCategory[key]);
    const fixedColors = ['#00796B', '#FFB300', '#1976D2', '#D84315', '#4CAF50', '#7B1FA2', '#FBC02D', '#5D4037'];

    const categorySalesConfig = {
        type: 'bar',
        data: {
            labels: categoryDataKeys,
            datasets: [{
                label: 'Total Revenue',
                data: categoryDataValues,
                backgroundColor: fixedColors.slice(0, categoryDataKeys.length),
            }]
        },
        options: {
            indexAxis: 'y', 
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: { beginAtZero: true, ticks: { callback: (val) => formatCurrency(val, true) } },
                y: { grid: { display: false } }
            },
            plugins: {
                legend: { display: false },
                tooltip: { callbacks: { label: (context) => `${context.label}: ${formatCurrency(context.parsed.x)}` } }
            },
            // NEW INTERACTIVITY: Click to filter by category
            onClick: (e) => {
                const activePoints = chartInstances.categorySales.getElementsAtEventForMode(e, 'nearest', { intersect: true }, true);
                if (activePoints.length > 0) {
                    const firstPoint = activePoints[0];
                    const label = chartInstances.categorySales.data.labels[firstPoint.index];
                    // Trigger filter with the clicked category
                    applyFilters({ type: 'category', value: label }); 
                }
            }
        }
    };
    chartInstances.categorySales = new Chart(DOM.categorySalesChart, categorySalesConfig);

    // --- Chart 3: Top 10 Products by Profit (Vertical Bar Chart) ---
    const topProfitData = productData.topProductsByProfit;
    const profitLabels = topProfitData.map(p => p.Product);
    const profitValues = topProfitData.map(p => p.Profit);
    const accentColor = getComputedStyle(document.body).getPropertyValue('--accent-color');

    const topProductsConfig = {
        type: 'bar',
        data: {
            labels: profitLabels,
            datasets: [{
                label: 'Total Profit',
                data: profitValues,
                backgroundColor: accentColor,
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: { grid: { display: false } },
                y: { beginAtZero: true, ticks: { callback: (val) => formatCurrency(val, true) } }
            },
            plugins: {
                legend: { display: false },
                tooltip: { callbacks: { label: (context) => `${context.label}: ${formatCurrency(context.parsed.y)}` } }
            }
        }
    };
    chartInstances.topProducts = new Chart(DOM.topProductsChart, topProductsConfig);

};

// NEW: Sorting Logic (Applied to a copy of the list before rendering)

/**
 * Sorts a data list based on a key and direction.
 * @param {Array<Object>} list The list to sort (a copy is returned).
 * @param {string} key The object property key.
 * @param {string} type The data type ('number' or 'string').
 * @param {string} direction The sort direction ('asc' or 'desc').
 * @returns {Array<Object>} The sorted list.
 */
const sortList = (list, key, type, direction) => {
    return [...list].sort((a, b) => {
        let aVal = a[key];
        let bVal = b[key];
        
        if (type === 'number') {
            aVal = parseFloat(aVal) || 0;
            bVal = parseFloat(bVal) || 0;
            return direction === 'asc' ? aVal - bVal : bVal - aVal;
        } else { // String or other
            aVal = String(aVal || '').toLowerCase();
            bVal = String(bVal || '').toLowerCase();
            if (aVal < bVal) return direction === 'asc' ? -1 : 1;
            if (aVal > bVal) return direction === 'asc' ? 1 : -1;
            return 0;
        }
    });
};

/**
 * Handles the table sorting click event, updates sort state, and re-renders the table.
 */
const handleTableSort = (tableId, key, type) => {
    if (!analysisResults || !analysisResults.productData) return;

    // 1. Determine new sort direction
    const currentState = currentSortState[tableId] || { key: key, direction: 'asc' };
    let newDirection = 'asc';

    if (currentState.key === key) {
        newDirection = currentState.direction === 'asc' ? 'desc' : 'asc';
    } else {
        // Default to descending for numbers/ascending for strings when key changes
        newDirection = type === 'number' ? 'desc' : 'asc';
    }

    // 2. Update state
    currentSortState[tableId] = { key: key, direction: newDirection, sourceKey: currentState.sourceKey };

    // 3. Update UI to reflect sort icons
    const tableElement = document.getElementById(tableId);
    tableElement.querySelectorAll('th').forEach(th => {
        th.classList.remove('sort-asc', 'sort-desc');
        const icon = th.querySelector('i');
        if (icon) icon.className = 'fas fa-sort'; // Reset icon
    });

    const activeHeader = tableElement.querySelector(`th[data-sort-key="${key}"]`);
    if (activeHeader) {
        activeHeader.classList.add(`sort-${newDirection}`);
        activeHeader.querySelector('i').className = newDirection === 'asc' ? 'fas fa-sort-up' : 'fas fa-sort-down';
    }
    
    // 4. Re-render the tables
    renderTables(analysisResults.productData);
};


/** Renders the Top Products, Reorder Alerts, and Product Profitability tables. */
const renderTables = (productData) => {

    // --- Data Derivation (Always re-derive default views) ---
    const defaultTopProducts = [...productData.allProducts]
        .sort((a, b) => b.Revenue - a.Revenue)
        .slice(0, 10);

    const defaultReorderAlerts = productData.allProducts.filter(p => p.IsLowStock)
        .sort((a, b) => a.StockLevel - b.StockLevel);

    const defaultProfitability = [...productData.allProducts]
        .sort((a, b) => b.Margin - a.Margin)
        .slice(0, 20);


    // --- 1. Top Products Table (by Revenue) ---
    let topProductsList = defaultTopProducts;
    const topSort = currentSortState['topProductsTable'];
    // If user clicked a column in this table, sort the full product list by that column and take top 10 rows
    if (topSort && (topSort.key !== 'Revenue' || topSort.direction !== 'desc')) {
        const sortedAll = sortList(productData.allProducts, topSort.key, document.querySelector(`#topProductsTable th[data-sort-key="${topSort.key}"]`)?.dataset.sortType || 'number', topSort.direction);
        topProductsList = sortedAll.slice(0, 10);
    }
    
    DOM.topProductsTable.innerHTML = topProductsList.map((p, index) => `
        <tr>
            <td>${index + 1}</td>
            <td>${p.Product}</td>
            <td>${formatCurrency(p.Revenue)}</td>
            <td>${formatCurrency(p.Profit)}</td>
        </tr>
    `).join('');


    // --- 2. Reorder Alerts Table (Low Stock) ---
    let reorderAlertsList = defaultReorderAlerts;
    const reorderSort = currentSortState['reorderTable'];
    // If user clicked a column in this table, sort the filtered list by that column
    if (reorderSort && (reorderSort.key !== 'StockLevel' || reorderSort.direction !== 'asc')) {
        reorderAlertsList = sortList(defaultReorderAlerts, reorderSort.key, document.querySelector(`#reorderTable th[data-sort-key="${reorderSort.key}"]`)?.dataset.sortType || 'number', reorderSort.direction);
    }
    
    DOM.reorderTable.innerHTML = reorderAlertsList.map(p => {
        const statusClass = p.IsLowStock ? 'low-stock' : 'reorder-safe';
        const statusText = p.IsLowStock ? 'LOW STOCK' : 'OK';
        return `
            <tr>
                <td>${p.Product}</td>
                <td>${formatNumber(p.StockLevel)}</td>
                <td>${formatNumber(p.ReorderPoint)}</td>
                <td class="${statusClass}">${statusText}</td>
            </tr>
        `;
    }).join('') || '<tr><td colspan="4" style="text-align:center;">No low stock alerts detected! 🎉</td></tr>';


    // --- 3. Product Profitability Table (All Products, sorted by Margin/User Sort) ---
    let profitabilityList = defaultProfitability;
    const profitSort = currentSortState['productProfitTable'];
    // If sorted by a non-default column, sort the full list. Show top 20 or all (if < 20).
    if (profitSort && (profitSort.key !== 'Margin' || profitSort.direction !== 'desc')) {
        profitabilityList = sortList(productData.allProducts, profitSort.key, document.querySelector(`#productProfitTable th[data-sort-key="${profitSort.key}"]`)?.dataset.sortType || 'number', profitSort.direction).slice(0, 20);
    } else {
        profitabilityList = defaultProfitability;
    }
    
    DOM.productProfitTable.innerHTML = profitabilityList.map(p => `
        <tr>
            <td>${p.Product}</td>
            <td>${formatPercentage(p.Margin)}</td>
            <td>${formatCurrency(p.Profit)}</td>
        </tr>
    `).join('');
    
    // --- 4. ABC Analysis Table ---
    let abcList = [...productData.abcProducts];
    const abcSort = currentSortState['abcAnalysisTable'];
    // If sorted by a non-default column, sort the ABC list.
    if (abcSort && (abcSort.key !== 'Revenue' || abcSort.direction !== 'desc')) {
        abcList = sortList(productData.abcProducts, abcSort.key, document.querySelector(`#abcAnalysisTable th[data-sort-key="${abcSort.key}"]`)?.dataset.sortType || 'number', abcSort.direction);
    }
    
    DOM.abcAnalysisTable.innerHTML = abcList.map((p, index) => `
        <tr class="segment-${p.Segment.toLowerCase()}">
            <td>${index + 1}</td>
            <td>${p.Product}</td>
            <td>${p.Category}</td>
            <td>${formatCurrency(p.Revenue)}</td>
            <td>${formatPercentage(p.CumulativePercent)}</td>
            <td>${p.Segment}</td>
        </tr>
    `).join('');

    // The sort icon update logic is moved to the end of handleTableSort, 
    // but we ensure defaults are set if no sort has occurred.
};

// --- Initialization and Event Listeners ---
document.addEventListener('DOMContentLoaded', () => {
    // Populate guide tables
    displayDataGuide();
    
    // Load Sample Data on page load
    startAnalysisAndRender(sampleRetailData);

    // Event Listeners
    DOM.excelFile.addEventListener('change', handleFileUpload);
    DOM.runAnalysisBtn.addEventListener('click', runAnalysisBtn);
    DOM.resetAppBtn.addEventListener('click', () => resetApplication(true)); // Full reset on Reset button

    // Filter Listeners
    DOM.applyFilterBtn.addEventListener('click', () => applyFilters()); // Pass no argument for manual button click
    DOM.clearFilterBtn.addEventListener('click', clearFilters);
    
    // Chart Toggle Listeners
    DOM.chartToggleBtns.forEach(btn => {
        btn.addEventListener('click', (e) => {
            const chartId = e.currentTarget.dataset.chart;
            const newType = e.currentTarget.dataset.type;
            toggleChartType(chartId, newType);
        });
    });

    // NEW: Table Sorting Listeners
    document.querySelectorAll('.data-table thead th[data-sort-key]').forEach(th => {
        th.addEventListener('click', (e) => {
            const thElement = e.currentTarget;
            const tableId = thElement.closest('table').id;
            const key = thElement.dataset.sortKey;
            const type = thElement.dataset.sortType;
            handleTableSort(tableId, key, type);
        });
    });

    // Data Guide Toggle Listener
    DOM.dataGuideBtn.addEventListener('click', toggleDataGuide);
    
    // NEW: Palette Toggle Listener
    DOM.paletteBtn.addEventListener('click', togglePaletteGuide);

    // Color Palette Changer
    document.querySelectorAll('.palette-option').forEach(btn => {
        btn.addEventListener('click', (e) => {
            changePalette(e.target.dataset.palette);
        });
    });

    // NEW: Click-outside listener to close any open dropdowns
    document.addEventListener('click', (e) => {
        // Data Guide Click Check
        const isGuideBtn = DOM.dataGuideBtn.contains(e.target);
        const isGuidePanel = DOM.guideDropdown.contains(e.target);
        
        // Palette Dropdown Click Check
        const isPaletteBtn = DOM.paletteBtn.contains(e.target);
        const isPalettePanel = DOM.paletteDropdown.contains(e.target);

        // Close Data Guide if click is outside the button and the panel
        if (!isGuideBtn && !isGuidePanel) {
            DOM.guideDropdown.classList.remove('open');
            DOM.dataGuideBtn.setAttribute('aria-expanded', false);
        }
        
        // Close Palette Guide if click is outside the button and the panel
        if (!isPaletteBtn && !isPalettePanel) {
            DOM.paletteDropdown.classList.remove('open');
            DOM.paletteBtn.setAttribute('aria-expanded', false);
        }
    });

    // Download Sample Link Listener
    if (DOM.downloadSampleLink) {
        DOM.downloadSampleLink.addEventListener('click', (e) => {
            e.preventDefault();
            downloadCSV(sampleRetailData, 'retail_dashboard_sample.csv');
        });
    }

    // Set initial button states
    DOM.runAnalysisBtn.textContent = 'Clear Data & New File';
    DOM.runAnalysisBtn.disabled = false;

    // Set initial default palette
    changePalette('palette-teal');
});