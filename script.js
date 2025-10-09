/* --- 4. script.js Requirements --- */

// --- A. Initialization & UI Control (DOM Elements) ---
const DOM = {
    excelFile: document.getElementById('excelFile'),
    runAnalysisBtn: document.getElementById('runAnalysisBtn'),
    resetAppBtn: document.getElementById('resetAppBtn'),
    guideSidebar: document.getElementById('guideSidebar'),
    toggleSidebarBtn: document.getElementById('toggleSidebarBtn'),
    requiredColumnsTable: document.getElementById('requiredColumnsTable'),
    kpiGuideTable: document.getElementById('kpiGuideTable'),
    loader: document.getElementById('loader'),
    statusMessage: document.getElementById('statusMessage'),
    runtimeDisplay: document.getElementById('runtimeDisplay'),
    dashboardContainer: document.getElementById('dashboardContainer'),
    metricsContainer: document.getElementById('metrics'),
    inventoryMetricsContainer: document.getElementById('inventoryMetrics'),
    salesTrendsChart: document.getElementById('salesTrendsChart'),
    categorySalesChart: document.getElementById('categorySalesChart'),
    topProductsTable: document.getElementById('topProductsTable').querySelector('tbody'),
    reorderTable: document.getElementById('reorderTable').querySelector('tbody'),
    productProfitTable: document.getElementById('productProfitTable').querySelector('tbody'),
};

let rawData = [];
let analysisResults = null;
let chartInstances = {}; // Store Chart.js instances

// --- Static Data for Sidebar Guide ---
const REQUIRED_COLUMNS = [
    { key: 'Product', example: 'Laptop A', desc: 'Unique item name' },
    { key: 'Category', example: 'Electronics', desc: 'Product grouping' },
    { key: 'OrderDate', example: '01/15/2024', desc: 'Date of sale' },
    { key: 'Sales', example: '1200.50', desc: 'Selling price' },
    { key: 'Cost', example: '800.00', desc: 'Cost price' },
    { key: 'StockLevel', example: '45', desc: 'Current units in stock' },
    { key: 'ReorderPoint', example: '10', desc: 'Stock level trigger' },
    { key: 'Quantity', example: '1', desc: 'Units sold in the transaction' }
];

const KPI_GUIDE = [
    { kpi: 'AOV', formula: 'Total Revenue / Total Orders' },
    { kpi: 'Margin', formula: 'Total Profit / Total Revenue' },
    { kpi: 'Reorder Alert', formula: 'StockLevel < ReorderPoint' }
];

// --- Utility Functions ---

/** Formats a number as currency. */
const formatCurrency = (value) => {
    return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(value);
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
    // Attempt to parse standard formats first
    let date = new Date(dateStr);
    if (!isNaN(date)) return date;

    // Handle DD/MM/YYYY or MM/DD/YYYY ambiguity by assuming DD/MM/YYYY if slashes/hyphens are present
    if (typeof dateStr === 'string') {
        const parts = dateStr.split(/[-\/.]/);
        if (parts.length === 3) {
            // Assume DD/MM/YYYY (Day-Month-Year)
            // Note: Date constructor is (Year, MonthIndex, Day)
            const d = parseInt(parts[0], 10);
            const m = parseInt(parts[1], 10);
            const y = parseInt(parts[2], 10);
            return new Date(y, m - 1, d);
        }
    }
    return new Date('Invalid Date');
};


/** Populates the tables in the retractable sidebar. */
const displayDataGuide = () => {
    // Required Columns Table
    DOM.requiredColumnsTable.innerHTML = `<thead><tr><th>Key Field</th><th>Example</th><th>Description</th></tr></thead><tbody>` +
        REQUIRED_COLUMNS.map(c => `<tr><td>${c.key}</td><td>${c.example}</td><td>${c.desc}</td></tr>`).join('') +
        `</tbody>`;

    // KPI Guide Table
    DOM.kpiGuideTable.innerHTML = `<thead><tr><th>KPI</th><th>Calculation</th></tr></thead><tbody>` +
        KPI_GUIDE.map(k => `<tr><td>${k.kpi}</td><td>${k.formula}</td></tr>`).join('') +
        `</tbody>`;

    // Initialize sidebar state
    DOM.guideSidebar.classList.add('open');
};

/** Toggles the sidebar's open/closed state. */
const toggleSidebar = () => {
    DOM.guideSidebar.classList.toggle('open');
    DOM.toggleSidebarBtn.querySelector('i').classList.toggle('fa-chevron-left');
    DOM.toggleSidebarBtn.querySelector('i').classList.toggle('fa-chevron-right');
};

/** Handles file upload and enables the analysis button. */
const handleFileUpload = (event) => {
    const file = event.target.files[0];
    if (file) {
        DOM.runAnalysisBtn.disabled = false;
        DOM.statusMessage.textContent = `File selected: ${file.name}. Ready to run analysis.`;
    } else {
        DOM.runAnalysisBtn.disabled = true;
        DOM.statusMessage.textContent = 'Awaiting file upload...';
    }
};

/**
 * Main button handler: displays loader, hides guide, starts analysis pipeline.
 */
const runAnalysisBtn = async () => {
    const file = DOM.excelFile.files[0];
    if (!file) {
        DOM.statusMessage.textContent = "Please select an Excel file first.";
        return;
    }

    // UI State: Start
    DOM.loader.classList.remove('hidden');
    DOM.runAnalysisBtn.disabled = true;
    DOM.dashboardContainer.classList.add('hidden');
    DOM.statusMessage.textContent = "Reading file...";
    DOM.runtimeDisplay.textContent = "";
    DOM.guideSidebar.classList.remove('open'); // Hide sidebar

    const startTime = performance.now();

    try {
        const data = await processFile(file);
        rawData = data;

        DOM.statusMessage.textContent = "File parsed successfully. Starting multi-stage analysis...";
        await stagedAnalysis(rawData);

        // UI State: Success
        DOM.dashboardContainer.classList.remove('hidden');
        DOM.statusMessage.textContent = "Analysis complete! Dashboard updated.";
    } catch (error) {
        // UI State: Error
        DOM.statusMessage.textContent = `Analysis failed: ${error.message}. Please check your file format.`;
        console.error("Analysis Error:", error);
    } finally {
        // UI State: End
        DOM.loader.classList.add('hidden');
        DOM.runAnalysisBtn.disabled = false;
        const endTime = performance.now();
        DOM.runtimeDisplay.textContent = `Runtime: ${(endTime - startTime).toFixed(2)}ms`;
    }
};

/** Processes the uploaded Excel file using SheetJS. */
const processFile = (file) => {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                // Convert sheet to array of JSON objects
                const json = XLSX.utils.sheet_to_json(worksheet);
                resolve(json);
            } catch (error) {
                reject(new Error("Failed to parse Excel file."));
            }
        };

        reader.onerror = (error) => reject(new Error("File reading failed."));
        reader.readAsArrayBuffer(file);
    });
};

/** Resets the application to its initial state. */
const resetApplication = () => {
    // Clear data and state
    rawData = [];
    analysisResults = null;
    DOM.excelFile.value = '';
    DOM.runAnalysisBtn.disabled = true;

    // Reset UI
    DOM.dashboardContainer.classList.add('hidden');
    DOM.statusMessage.textContent = 'Awaiting file upload...';
    DOM.runtimeDisplay.textContent = '';
    DOM.metricsContainer.innerHTML = '';
    DOM.inventoryMetricsContainer.innerHTML = '';
    DOM.topProductsTable.innerHTML = '';
    DOM.reorderTable.innerHTML = '';
    DOM.productProfitTable.innerHTML = '';
    DOM.guideSidebar.classList.add('open');

    // Destroy existing charts
    Object.values(chartInstances).forEach(chart => chart.destroy());
    chartInstances = {};
};

/** Changes the active CSS color palette. */
const changePalette = (paletteName) => {
    document.body.className = paletteName;
};

// --- B. Data Processing & Analysis (Multi-Stage) ---

/** Multi-stage analysis pipeline using setTimeout to allow UI updates. */
const stagedAnalysis = (data) => {
    return new Promise((resolve) => {
        let mappedData = [];
        let results = {};

        // 1. Defensive Data Mapping (Initial Filter & Clean)
        const mapStage = () => {
            DOM.statusMessage.textContent = "Stage 0/3: Mapping and cleaning data...";
            // Define expected/alternate keys for robustness
            const KEY_MAP = {
                'Sales': ['Sales', 'Revenue'],
                'OrderDate': ['Order Date', 'Date', 'OrderDate'],
                'Cost': ['Cost', 'Cost Price', 'CostPrice'],
                'ReorderPoint': ['Reorder Point', 'ReorderPoint'],
            };

            const findKey = (item, keys) => keys.find(k => item.hasOwnProperty(k));

            mappedData = data.map(item => {
                const salesKey = findKey(item, KEY_MAP.Sales);
                const costKey = findKey(item, KEY_MAP.Cost);
                const dateKey = findKey(item, KEY_MAP.OrderDate);
                const reorderKey = findKey(item, KEY_MAP.ReorderPoint);

                if (!salesKey || !costKey || !dateKey || !item.Product || !item.Quantity || !item.StockLevel) {
                    return null; // Skip invalid rows
                }

                const sales = parseFloat(item[salesKey]) || 0;
                const cost = parseFloat(item[costKey]) || 0;
                const quantity = parseInt(item.Quantity) || 0;

                return {
                    // Core Data
                    Product: String(item.Product || 'Unknown Product'),
                    Category: String(item.Category || 'Other'),
                    StockLevel: parseInt(item.StockLevel) || 0,
                    ReorderPoint: parseInt(item[reorderKey]) || 0,
                    // Calculated Metrics
                    Sales: sales,
                    Cost: cost,
                    Profit: sales - cost,
                    Quantity: quantity,
                    OrderDate: parseDate(item[dateKey]),
                    OrderID: item.OrderID || 'N/A', // Assuming OrderID exists or fallback
                };
            }).filter(item => item && !isNaN(item.OrderDate.getTime())); // Filter out nulls and invalid dates

            // Go to Stage 1
            setTimeout(stage1_CoreMetrics, 0);
        };

        // 2. Stage 1 (Core Metrics)
        const stage1_CoreMetrics = () => {
            DOM.statusMessage.textContent = "Stage 1/3: Calculating core financial and sales metrics...";
            let totalRevenue = 0;
            let totalProfit = 0;
            const uniqueOrders = new Set();
            const uniqueCustomers = new Set(); // Assumes data has a CustomerID field (or uses OrderID as proxy if not)
            const salesByDate = {};
            const salesByCategory = {};

            mappedData.forEach(item => {
                totalRevenue += item.Sales;
                totalProfit += item.Profit;

                // For AOV: Tally unique orders
                if (item.OrderID && item.OrderID !== 'N/A') uniqueOrders.add(item.OrderID);
                else uniqueOrders.add(`${item.OrderDate.toISOString()}_${item.Product}`); // Fallback

                // For Customer Count (assuming Product is a proxy if no CustomerID)
                uniqueCustomers.add(item.OrderID || item.Product); // Simple unique count proxy

                // Sales Trends (Daily)
                const dateKey = item.OrderDate.toISOString().split('T')[0]; // YYYY-MM-DD
                salesByDate[dateKey] = (salesByDate[dateKey] || 0) + item.Sales;

                // Category Sales
                const category = item.Category;
                salesByCategory[category] = (salesByCategory[category] || 0) + item.Sales;
            });

            const totalOrders = uniqueOrders.size;
            const avgOrderValue = totalOrders > 0 ? totalRevenue / totalOrders : 0;
            const profitMargin = totalRevenue > 0 ? totalProfit / totalRevenue : 0;

            results.coreMetrics = {
                TotalRevenue: totalRevenue,
                TotalProfit: totalProfit,
                AOV: avgOrderValue,
                ProfitMargin: profitMargin,
                TotalOrders: totalOrders,
                UniqueCustomers: uniqueCustomers.size,
                SalesByDate: salesByDate,
                SalesByCategory: salesByCategory,
            };

            // Go to Stage 2
            setTimeout(stage2_InventoryAndProducts, 0);
        };

        // 3. Stage 2 (Inventory/Products)
        const stage2_InventoryAndProducts = () => {
            DOM.statusMessage.textContent = "Stage 2/3: Analyzing inventory and product performance...";

            const productSummary = {};
            let totalStockValue = 0;
            let totalStockUnits = 0;

            mappedData.forEach(item => {
                const product = item.Product;

                // Aggregate sales metrics
                if (!productSummary[product]) {
                    productSummary[product] = {
                        Revenue: 0,
                        Profit: 0,
                        QuantitySold: 0,
                        TotalCost: 0,
                        StockLevel: 0, // Will be updated by the last row for this product (assumes StockLevel is current)
                        ReorderPoint: 0,
                    };
                }

                productSummary[product].Revenue += item.Sales;
                productSummary[product].Profit += item.Profit;
                productSummary[product].QuantitySold += item.Quantity;
                productSummary[product].TotalCost += item.Cost * item.Quantity; // Total cost per unit

                // Inventory Metrics (using the last encountered value for the current stock/reorder point)
                productSummary[product].StockLevel = item.StockLevel;
                productSummary[product].ReorderPoint = item.ReorderPoint;
            });

            // Calculate Inventory KPIs and Final Product Summary
            const productList = Object.entries(productSummary).map(([product, data]) => {
                const margin = data.Revenue > 0 ? data.Profit / data.Revenue : 0;
                totalStockValue += data.StockLevel * (data.Revenue / data.QuantitySold || 0); // Approx. stock value
                totalStockUnits += data.StockLevel;
                return {
                    Product: product,
                    ...data,
                    Margin: margin,
                    IsLowStock: data.StockLevel < data.ReorderPoint,
                };
            });

            // Top Products
            const topProducts = [...productList]
                .sort((a, b) => b.Revenue - a.Revenue)
                .slice(0, 10);

            // Low Stock Alerts
            const reorderAlerts = productList.filter(p => p.IsLowStock)
                .sort((a, b) => a.StockLevel - b.StockLevel); // Sort by lowest stock first

            results.inventoryMetrics = {
                TotalStockUnits: totalStockUnits,
                AvgStockLevel: totalStockUnits / productList.length,
                TotalStockValue: totalStockValue,
                TotalProducts: productList.length,
            };

            results.productData = {
                topProducts: topProducts,
                reorderAlerts: reorderAlerts,
                allProducts: productList,
            };

            analysisResults = results; // Save final results

            // Go to Rendering Stage
            setTimeout(renderDashboard, 0);
        };

        // 4. Rendering Stage
        const renderDashboard = () => {
            DOM.statusMessage.textContent = "Stage 3/3: Rendering dashboard charts and tables...";
            renderMetrics(analysisResults.coreMetrics, analysisResults.inventoryMetrics);
            renderCharts(analysisResults.coreMetrics);
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
    // 1. Core Metrics
    const coreMetricsHtml = [
        { icon: 'fas fa-dollar-sign', title: 'Total Revenue', value: formatCurrency(core.TotalRevenue) },
        { icon: 'fas fa-hand-holding-usd', title: 'Total Profit', value: formatCurrency(core.TotalProfit) },
        { icon: 'fas fa-percent', title: 'Profit Margin', value: formatPercentage(core.ProfitMargin) },
        { icon: 'fas fa-calculator', title: 'Avg. Order Value (AOV)', value: formatCurrency(core.AOV) },
        { icon: 'fas fa-shopping-basket', title: 'Total Orders', value: formatNumber(core.TotalOrders) },
        { icon: 'fas fa-users', title: 'Unique Customers (Proxy)', value: formatNumber(core.UniqueCustomers) },
    ].map(m => `
        <div class="metric-card">
            <span class="icon"><i class="${m.icon}"></i></span>
            <div class="value">${m.value}</div>
            <div class="title">${m.title}</div>
        </div>
    `).join('');

    DOM.metricsContainer.innerHTML = coreMetricsHtml;

    // 2. Inventory Metrics
    const inventoryMetricsHtml = [
        { icon: 'fas fa-cubes', title: 'Total Stock Units', value: formatNumber(inventory.TotalStockUnits) },
        { icon: 'fas fa-box-open', title: 'Total Product SKUs', value: formatNumber(inventory.TotalProducts) },
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
const renderCharts = (core) => {
    // Destroy previous chart instances
    Object.values(chartInstances).forEach(chart => chart.destroy());

    // --- Chart 1: Sales Trends Over Time (Line Chart) ---
    const salesDataKeys = Object.keys(core.SalesByDate).sort();
    const salesDataValues = salesDataKeys.map(key => core.SalesByDate[key]);

    const salesTrendsConfig = {
        type: 'line',
        data: {
            labels: salesDataKeys,
            datasets: [{
                label: 'Daily Revenue',
                data: salesDataValues,
                borderColor: getComputedStyle(document.body).getPropertyValue('--primary-color'),
                backgroundColor: 'rgba(0, 121, 107, 0.1)',
                borderWidth: 2,
                pointRadius: 0,
                tension: 0.1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: { display: false }, // Hide x-axis labels if too crowded
                y: { beginAtZero: true, ticks: { callback: (val) => formatCurrency(val, true) } }
            },
            plugins: { legend: { display: false } }
        }
    };
    chartInstances.salesTrends = new Chart(DOM.salesTrendsChart, salesTrendsConfig);


    // --- Chart 2: Sales by Category (Doughnut Chart) ---
    const categoryDataKeys = Object.keys(core.SalesByCategory);
    const categoryDataValues = categoryDataKeys.map(key => core.SalesByCategory[key]);

    // Simple fixed color array for consistency
    const fixedColors = ['#00796B', '#FFB300', '#1976D2', '#D84315', '#4CAF50', '#7B1FA2', '#FBC02D', '#5D4037'];

    const categorySalesConfig = {
        type: 'doughnut',
        data: {
            labels: categoryDataKeys,
            datasets: [{
                data: categoryDataValues,
                backgroundColor: fixedColors.slice(0, categoryDataKeys.length),
                hoverOffset: 4
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'right' },
                tooltip: { callbacks: { label: (context) => `${context.label}: ${formatCurrency(context.parsed)}` } }
            }
        }
    };
    chartInstances.categorySales = new Chart(DOM.categorySalesChart, categorySalesConfig);
};

/** Renders the Top Products, Reorder Alerts, and Product Profitability tables. */
const renderTables = (productData) => {
    // 1. Top Products Table
    DOM.topProductsTable.innerHTML = productData.topProducts.map((p, index) => `
        <tr>
            <td>${index + 1}</td>
            <td>${p.Product}</td>
            <td>${formatCurrency(p.Revenue)}</td>
            <td>${formatCurrency(p.Profit)}</td>
        </tr>
    `).join('');

    // 2. Reorder Alerts Table (Low Stock)
    DOM.reorderTable.innerHTML = productData.reorderAlerts.map(p => {
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

    // 3. Product Profitability Table (All Products, sorted by Margin)
    const profitabilityData = [...productData.allProducts]
        .sort((a, b) => b.Margin - a.Margin)
        .slice(0, 20); // Show top 20 for brevity

    DOM.productProfitTable.innerHTML = profitabilityData.map(p => `
        <tr>
            <td>${p.Product}</td>
            <td>${formatPercentage(p.Margin)}</td>
            <td>${formatCurrency(p.Profit)}</td>
        </tr>
    `).join('');
};

// --- Initialization and Event Listeners ---
document.addEventListener('DOMContentLoaded', () => {
    // Populate sidebar guide tables
    displayDataGuide();

    // Event Listeners
    DOM.excelFile.addEventListener('change', handleFileUpload);
    DOM.runAnalysisBtn.addEventListener('click', runAnalysisBtn);
    DOM.resetAppBtn.addEventListener('click', resetApplication);
    DOM.toggleSidebarBtn.addEventListener('click', toggleSidebar);

    // Color Palette Changer
    document.querySelectorAll('.palette-option').forEach(btn => {
        btn.addEventListener('click', (e) => {
            changePalette(e.target.dataset.palette);
        });
    });

    // Set initial default palette
    changePalette('palette-teal');
});