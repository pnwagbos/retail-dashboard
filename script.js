document.addEventListener("DOMContentLoaded", () => {
  displayDataGuide();
  document.getElementById("runAnalysisBtn").addEventListener("click", runAnalysisBtn);
  document.getElementById("colorPaletteDropdown").addEventListener("change", changePalette);
});

function displayDataGuide() {
  const requiredColumns = {
    "OrderDate": "Date when the order was placed",
    "Cost": "Cost of the product sold",
    "ReorderPoint": "Threshold to trigger restocking",
    "StockLevel": "Current inventory level",
    "Product": "Name of the product",
    "Revenue": "Total revenue from the sale"
  };

  const kpis = {
    "Total Sales": "Sum of all revenue generated",
    "Total Profit": "Revenue minus cost",
    "AOV": "Average Order Value = Total Sales / Orders",
    "Profit Margin": "Profit as a percentage of sales",
    "Top Products": "Products with highest revenue",
    "Low Stock Alerts": "Products below reorder threshold"
  };

  const requiredTable = document.getElementById("requiredColumnsTable");
  requiredTable.innerHTML = Object.entries(requiredColumns)
    .map(([col, tip]) => `<tr><td data-tooltip="${tip}">${col}</td></tr>`).join("");

  const kpiTable = document.getElementById("kpiGuideTable");
  kpiTable.innerHTML = Object.entries(kpis)
    .map(([kpi, tip]) => `<tr><td data-tooltip="${tip}">${kpi}</td></tr>`).join("");
}

function changePalette(e) {
  document.body.className = e.target.value;
}

function runAnalysisBtn() {
  const fileInput = document.getElementById("excelFile");
  if (!fileInput.files.length) return alert("Please upload an Excel file.");

  document.getElementById("loader").classList.remove("hidden");
  document.getElementById("statusMessage").textContent = "Processing file...";
  setTimeout(() => processFileAndRunAnalysis(fileInput.files[0]), 0);
}

function processFileAndRunAnalysis(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    const workbook = XLSX.read(e.target.result, { type: "binary" });
    const sheetName = workbook.SheetNames[0];
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
    stagedAnalysis(data);
  };
  reader.readAsBinaryString(file);
}

function parseDate(dateStr) {
  const [day, month, year] = dateStr.split("/");
  return new Date(`${year}-${month}-${day}`);
}

function stagedAnalysis(data) {
  setTimeout(() => {
    const results = {
      totalSales: 0,
      totalProfit: 0,
      aov: 0,
      margin: 0,
      topProducts: {},
      lowStock: [],
      salesByDate: {},
      categorySales: {}
    };

    data.forEach(row => {
      const revenue = parseFloat(row.Revenue || 0);
      const cost = parseFloat(row.Cost || row["Cost price"] || 0);
      const profit = revenue - cost;
      results.totalSales += revenue;
      results.totalProfit += profit;

      const product = row.Product;
      results.topProducts[product] = (results.topProducts[product] || 0) + revenue;

      if (parseFloat(row.StockLevel || 0) < parseFloat(row.ReorderPoint || 0)) {
        results.lowStock.push(row);
      }

      const date = parseDate(row.OrderDate);
      const dateKey = date.toISOString().split("T")[0];
      results.salesByDate[dateKey] = (results.salesByDate[dateKey] || 0) + revenue;

      const category = row.Category || "Uncategorized";
      results.categorySales[category] = (results.categorySales[category] || 0) + revenue;
    });

    results.aov = results.totalSales / data.length;
    results.margin = results.totalProfit / results.totalSales;

    renderDashboard(results);
  }, 0);
}

function renderDashboard(results) {
  document.getElementById("loader").classList.add("hidden");
  document.getElementById("dashboardContainer").classList.remove("hidden");

  renderMetrics(results);
  renderCharts(results);
  renderTables(results);
}

function renderMetrics(results) {
  const metricsHTML = `
    <div>Total Sales: ${formatCurrency(results.totalSales)}</div>
    <div>Total Profit: ${formatCurrency(results.totalProfit)}</div>
    <div>AOV: ${formatCurrency(results.aov)}</div>
    <div>Profit Margin: ${formatPercentage(results.margin)}</div>
  `;
  document.getElementById("metrics").innerHTML = metricsHTML;
}

function renderCharts(results) {
  const salesCtx = document.getElementById("salesTrendsChart").getContext("2d");
  new Chart(salesCtx, {
    type: "line",
    data: {
      labels: Object.keys(results.salesByDate),
      datasets: [{
        label: "Sales Over Time",
        data: Object.values(results.salesByDate),
        borderColor: "#fff",
        backgroundColor: "rgba(255,255,255,0.2)"
      }]
    }
  });

  const categoryCtx = document.getElementById("categorySalesChart").getContext("2d");
  new Chart(categoryCtx, {
    type: "bar",
    data: {
      labels: Object.keys(results.categorySales),
      datasets: [{
        label: "Sales by Category",
        data: Object.values(results.categorySales),
        backgroundColor: "rgba(255,255,255,0.5)"
      }]
    }
  });
}

function renderTables(results) {
  const topProducts = Object.entries(results.topProducts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10);

  document.getElementById("topProductsTable").innerHTML = `
    <tr><th>Product</th><th>Revenue</th></tr>
    ${topProducts.map(([product, revenue]) => `<tr><td>${product}</td><td>${formatCurrency(revenue)}</td></tr>`).join("")}
  `;

  document.getElementById("reorderTable").innerHTML = `
    <tr><th>Product</th><th>Stock Level</th><th>Reorder Point</th></tr>
    ${results.lowStock.map(row => `<tr><td>${row.Product}</td><td>${row.StockLevel}</td><td>${row.ReorderPoint}</td></tr>`).join("")}
  `;
}

function formatCurrency(num) {
  return `\$${num.toFixed(2)}`;
}

function formatPercentage(num) {
  return `${(num * 100).toFixed(2)}%`;
}

function formatNumber(num) {
  return num.toLocaleString();
}