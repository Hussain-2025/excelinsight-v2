// Universal App State
const AppState = {
  apiBaseUrl: (window.ENV && window.ENV.API_URL) ? window.ENV.API_URL : 'http://localhost:5050/api',
  auth: {
    token: localStorage.getItem('excelinsight_token') || null,
    user: JSON.parse(localStorage.getItem('excelinsight_user') || 'null')
  },
  fileName: 'Sales_Performance_Q3.xlsx',
  activeSheet: 'Sales',
  sheets: {},
  currentRows: [],
  originalRows: [],
  filteredRows: [],
  columns: [],
  colMeta: {},
  grid: {
    page: 1,
    pageSize: 50,
    sortCol: null,
    sortDir: 'asc',
    searchQuery: '',
    columnFilter: 'all'
  },
  categoryChartInstance: null,
  studioChartInstance: null,
  threeJs: {
    renderer: null,
    scene: null,
    camera: null,
    controls: null,
    gridHelper: null,
    animationId: null,
    autoRotate: true,
    rotationAngle: 0,
    theme: 'cyber',
    showGrid: true,
    meshes: [],
    raycaster: null,
    mouse: null,
    activeHoverMesh: null,
    currentType: '3d-column',
    isFullscreen: false
  },
  threeJsDedicated: {
    renderer: null,
    scene: null,
    camera: null,
    controls: null,
    gridHelper: null,
    animationId: null,
    autoRotate: true,
    rotationSpeed: 1.0,
    rotationAngle: 0,
    theme: 'cyber',
    showGrid: true,
    wireframe: false,
    heightScale: 1.0,
    meshes: [],
    activeHoverMesh: null,
    currentType: '3d-column',
    isFullscreen: false
  },
  lastAiInsight: '',
  automl: {
    taskType: 'regression',
    targetColumn: null,
    featureColumns: [],
    dateColumn: null,
    horizonSteps: 6,
    primaryMetric: 'mae',
    splitRatio: 0.8,
    models: [],
    bestModel: null,
    featureImportance: [],
    testPredictions: [],
    charts: {
      xai: null,
      comparison: null,
      validation: null,
      forecast: null
    },
    activePipeline: null
  }
};

// ---------------------------------------------------------------------------
// Shared value helpers: one definition of "blank" and "number" for every feature
// ---------------------------------------------------------------------------
function isBlankValue(v) {
  return v === null || v === undefined || (typeof v === 'string' && v.trim() === '');
}

function toNumberOrNull(v) {
  if (isBlankValue(v) || typeof v === 'boolean') return null;
  const n = Number(typeof v === 'string' ? v.trim() : v);
  return Number.isFinite(n) ? n : null;
}

function numericValues(rows, col) {
  const out = [];
  rows.forEach(r => { const n = toNumberOrNull(r[col]); if (n !== null) out.push(n); });
  return out;
}

function groupKeyOf(v) {
  return isBlankValue(v) ? '(Blank)' : String(v).trim();
}

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

// Aggregate a list of numbers. Blank cells are never part of `nums`.
function aggregateNumbers(nums, agg, rowCount) {
  if (agg === 'count') return rowCount ?? nums.length;
  if (nums.length === 0) return null;
  if (agg === 'sum') return nums.reduce((a, b) => a + b, 0);
  if (agg === 'avg') return nums.reduce((a, b) => a + b, 0) / nums.length;
  if (agg === 'max') return Math.max(...nums);
  if (agg === 'min') return Math.min(...nums);
  if (agg === 'median') {
    const s = [...nums].sort((a, b) => a - b);
    const m = Math.floor(s.length / 2);
    return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
  }
  return null;
}

function formatNumber(v, maxDigits = 2) {
  if (v === null || v === undefined || !Number.isFinite(v)) return '–';
  return v.toLocaleString(undefined, { maximumFractionDigits: maxDigits });
}

// Small, safe markdown renderer for AI text (headings, bullets, numbered lists, bold, italics)
function renderMarkdownLite(text) {
  if (!text) return '';
  const inline = s => escapeHtml(s)
    .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
    .replace(/(^|[^*])\*(?!\s)(.+?)\*(?!\*)/g, '$1<em>$2</em>')
    .replace(/`([^`]+)`/g, '<code>$1</code>');
  const lines = String(text).replace(/\r/g, '').split('\n');
  let html = '';
  let list = null;
  const close = () => { if (list) { html += `</${list}>`; list = null; } };
  lines.forEach(raw => {
    const line = raw.trimEnd();
    let m;
    if (!line.trim()) { close(); return; }
    if ((m = line.match(/^\s*#{1,6}\s+(.*)$/))) { close(); html += `<h4 class="md-h">${inline(m[1])}</h4>`; return; }
    if ((m = line.match(/^\s*[-*•]\s+(.*)$/))) { if (list !== 'ul') { close(); html += '<ul class="md-ul">'; list = 'ul'; } html += `<li>${inline(m[1])}</li>`; return; }
    if ((m = line.match(/^\s*\d+[.)]\s+(.*)$/))) { if (list !== 'ol') { close(); html += '<ol class="md-ol">'; list = 'ol'; } html += `<li>${inline(m[1])}</li>`; return; }
    close();
    html += `<p class="md-p">${inline(line)}</p>`;
  });
  close();
  return html;
}

// If a model reply arrives as raw JSON text, pull out the "answer" field
function extractAiAnswer(text) {
  if (typeof text !== 'string') return text;
  const s = text.trim().replace(/^```json/i, '').replace(/```$/, '').trim();
  if (!s.startsWith('{')) return text;
  try { const o = JSON.parse(s); if (o && o.answer) return o.answer; } catch (e) { /* truncated */ }
  const m = s.match(/"answer"\s*:\s*"((?:[^"\\]|\\.)*)/);
  if (m) { try { return JSON.parse(`"${m[1]}"`); } catch (e) { return m[1]; } }
  return text;
}

// Exact group totals over ALL rows, sent to the AI so it never sums a 15-row sample
function buildDataAggregates() {
  const rows = AppState.currentRows || [];
  const numCols = AppState.columns.filter(c => AppState.colMeta[c]?.type === 'numeric').slice(0, 8);
  const catCols = AppState.columns.filter(c => {
    const m = AppState.colMeta[c];
    return m?.type === 'categorical' && m.distinctCount <= 30 && m.distinctCount < Math.max(2, rows.length * 0.9);
  }).slice(0, 6);
  const columnTotals = {};
  numCols.forEach(c => {
    const nums = numericValues(rows, c);
    columnTotals[c] = { sum: aggregateNumbers(nums, 'sum'), mean: aggregateNumbers(nums, 'avg'), min: aggregateNumbers(nums, 'min'), max: aggregateNumbers(nums, 'max'), filled: nums.length };
  });
  const groupTotals = {};
  catCols.forEach(cat => {
    groupTotals[cat] = {};
    numCols.forEach(num => {
      const sums = {};
      rows.forEach(r => {
        const n = toNumberOrNull(r[num]);
        if (n === null) return;
        const k = groupKeyOf(r[cat]);
        sums[k] = Number(((sums[k] || 0) + n).toFixed(4));
      });
      groupTotals[cat][num] = sums;
    });
  });
  return { rowCount: rows.length, columnTotals, groupTotals };
}

// Preloaded Datasets
const SAMPLE_DATASETS = {
  sales: {
    fileName: 'Sales_Performance_Q3.xlsx',
    sheets: {
      'Sales': [
        { "OrderID": "SO-1001", "Region": "North America", "Category": "Electronics", "SubCategory": "Laptops", "Sales": 10500, "Profit": 2100, "Units": 14 },
        { "OrderID": "SO-1002", "Region": "Europe", "Category": "Furniture", "SubCategory": "Chairs", "Sales": 3000, "Profit": 450, "Units": 25 },
        { "OrderID": "SO-1003", "Region": "Asia Pacific", "Category": "Office Supplies", "SubCategory": "Paper", "Sales": 1200, "Profit": 320, "Units": 80 },
        { "OrderID": "SO-1004", "Region": "North America", "Category": "Furniture", "SubCategory": "Desks", "Sales": 3600, "Profit": 720, "Units": 8 },
        { "OrderID": "SO-1005", "Region": "Europe", "Category": "Electronics", "SubCategory": "Monitors", "Sales": 6600, "Profit": 1650, "Units": 30 },
        { "OrderID": "SO-1006", "Region": "Latin America", "Category": "Office Supplies", "SubCategory": "Binders", "Sales": 810, "Profit": 190, "Units": 45 },
        { "OrderID": "SO-1007", "Region": "Asia Pacific", "Category": "Electronics", "SubCategory": "Smartphones", "Sales": 13200, "Profit": 3300, "Units": 22 },
        { "OrderID": "SO-1008", "Region": "Middle East", "Category": "Furniture", "SubCategory": "Bookcases", "Sales": 3720, "Profit": 600, "Units": 12 },
        { "OrderID": "SO-1009", "Region": "North America", "Category": "Office Supplies", "SubCategory": "Storage", "Sales": 1400, "Profit": 280, "Units": 35 },
        { "OrderID": "SO-1010", "Region": "Europe", "Category": "Electronics", "SubCategory": "Accessories", "Sales": 2250, "Profit": 700, "Units": 90 },
        { "OrderID": "SO-1011", "Region": "Asia Pacific", "Category": "Furniture", "SubCategory": "Tables", "Sales": 3300, "Profit": 420, "Units": 6 },
        { "OrderID": "SO-1012", "Region": "Latin America", "Category": "Electronics", "SubCategory": "Laptops", "Sales": 7200, "Profit": 1400, "Units": 10 },
        { "OrderID": "SO-1013", "Region": "Middle East", "Category": "Electronics", "SubCategory": "Smartphones", "Sales": 8900, "Profit": 1900, "Units": 15 },
        { "OrderID": "SO-1014", "Region": "North America", "Category": "Furniture", "SubCategory": "Chairs", "Sales": 4500, "Profit": 850, "Units": 20 },
        { "OrderID": "SO-1015", "Region": "Europe", "Category": "Office Supplies", "SubCategory": "Paper", "Sales": 950, "Profit": 210, "Units": 50 },
        { "OrderID": "SO-1016", "Region": "Asia Pacific", "Category": "Office Supplies", "SubCategory": "Binders", "Sales": 1100, "Profit": 240, "Units": 60 },
        { "OrderID": "SO-1017", "Region": "Latin America", "Category": "Furniture", "SubCategory": "Desks", "Sales": 2900, "Profit": -45, "Units": 7 },
        { "OrderID": "SO-1018", "Region": "Middle East", "Category": "Office Supplies", "SubCategory": "Storage", "Sales": 1650, "Profit": 310, "Units": 28 }
      ],
      'Targets': [
        { "Region": "North America", "Target": 25000 },
        { "Region": "Europe", "Target": 18000 },
        { "Region": "Asia Pacific", "Target": 20000 }
      ]
    }
  },
  students: {
    fileName: 'Dummy_Student_and_Employee_Data.xlsx',
    sheets: {
      'Students': [
        { "StudentID": "STU-001", "Name": "Aarav Sharma", "Gender": "Male", "Class": "12-B", "Math": 52, "Physics": 46, "Chemistry": 92, "English": 62 },
        { "StudentID": "STU-002", "Name": "Aditya Sharma", "Gender": "Female", "Class": "10-B", "Math": 53, "Physics": 92, "Chemistry": 51, "English": 88 },
        { "StudentID": "STU-003", "Name": "Arjun Sharma", "Gender": "Male", "Class": "12-A", "Math": 50, "Physics": 82, "Chemistry": 72, "English": 47 },
        { "StudentID": "STU-004", "Name": "Kabir Sharma", "Gender": "Female", "Class": "10-A", "Math": 58, "Physics": 59, "Chemistry": 65, "English": 70 },
        { "StudentID": "STU-005", "Name": "Diya Patel", "Gender": "Female", "Class": "10-A", "Math": 88, "Physics": 78, "Chemistry": 85, "English": 90 },
        { "StudentID": "STU-006", "Name": "Ishaan Verma", "Gender": "Male", "Class": "11-A", "Math": 70, "Physics": 75, "Chemistry": 68, "English": 72 },
        { "StudentID": "STU-007", "Name": "Kavya Nair", "Gender": "Female", "Class": "11-B", "Math": 65, "Physics": 70, "Chemistry": 80, "English": 70 },
        { "StudentID": "STU-008", "Name": "Rohan Gupta", "Gender": "Male", "Class": "10-A", "Math": 92, "Physics": 85, "Chemistry": 89, "English": 78 },
        { "StudentID": "STU-009", "Name": "Ananya Iyer", "Gender": "Female", "Class": "12-A", "Math": 78, "Physics": 84, "Chemistry": 80, "English": 82 },
        { "StudentID": "STU-010", "Name": "Manish Kumar", "Gender": "Male", "Class": "10-B", "Math": 60, "Physics": 58, "Chemistry": 62, "English": 64 },
        { "StudentID": "STU-011", "Name": "Pooja Reddy", "Gender": "Female", "Class": "12-B", "Math": 85, "Physics": 80, "Chemistry": 76, "English": 84 },
        { "StudentID": "STU-012", "Name": "Varun Joshi", "Gender": "Male", "Class": "11-A", "Math": 45, "Physics": 52, "Chemistry": 50, "English": 58 },
        { "StudentID": "STU-013", "Name": "Sneha Kulkarni", "Gender": "Female", "Class": "11-B", "Math": 72, "Physics": 76, "Chemistry": 74, "English": 68 },
        { "StudentID": "STU-014", "Name": "Karan Mehra", "Gender": "Male", "Class": "10-A", "Math": 58, "Physics": 64, "Chemistry": 60, "English": 62 },
        { "StudentID": "STU-015", "Name": "Meera Rao", "Gender": "Female", "Class": "12-B", "Math": 90, "Physics": 88, "Chemistry": 92, "English": 85 },
        { "StudentID": "STU-016", "Name": "Nikhil Bose", "Gender": "Male", "Class": "10-B", "Math": 74, "Physics": 70, "Chemistry": 68, "English": 72 },
        { "StudentID": "STU-017", "Name": "Priya Sen", "Gender": "Female", "Class": "12-A", "Math": 48, "Physics": 50, "Chemistry": 52, "English": 60 },
        { "StudentID": "STU-018", "Name": "Siddharth Das", "Gender": "Male", "Class": "11-A", "Math": 82, "Physics": 80, "Chemistry": 85, "English": 79 }
      ],
      'Employees': [
        { "EmpID": "EMP-101", "FullName": "Rajesh Kumar", "Department": "Engineering", "Location": "Delhi", "MonthlySalary": 85000, "ExperienceYears": 6, "Status": "Active" },
        { "EmpID": "EMP-102", "FullName": "Sunita Verma", "Department": "Finance", "Location": "Mumbai", "MonthlySalary": 72000, "ExperienceYears": 4, "Status": "Active" },
        { "EmpID": "EMP-103", "FullName": "Amit Singh", "Department": "Human Resources", "Location": "Delhi ", "MonthlySalary": 62000, "ExperienceYears": 3, "Status": "Active" },
        { "EmpID": "EMP-104", "FullName": "Priya Nair", "Department": "Engineering", "Location": "Bangalore", "MonthlySalary": 96000, "ExperienceYears": 8, "Status": "Active" },
        { "EmpID": "EMP-105", "FullName": "Vikram Patel", "Department": "Marketing", "Location": "Delhi", "MonthlySalary": 58000, "ExperienceYears": 2, "Status": "On Leave" },
        { "EmpID": "EMP-106", "FullName": "Ananya Roy", "Department": "Finance", "Location": "Bangalore", "MonthlySalary": 79000, "ExperienceYears": 5, "Status": "Active" }
      ]
    }
  },
  employees: {
    fileName: 'HR_Payroll_Directory.xlsx',
    sheets: {
      'Staff': [
        { "EmpID": "EMP-101", "FullName": "Rajesh Kumar", "Department": "Engineering", "Location": "Delhi", "MonthlySalary": 85000, "ExperienceYears": 6, "Status": "Active" },
        { "EmpID": "EMP-102", "FullName": "Sunita Verma", "Department": "Finance", "Location": "Mumbai", "MonthlySalary": 72000, "ExperienceYears": 4, "Status": "Active" },
        { "EmpID": "EMP-103", "FullName": "Amit Singh", "Department": "Human Resources", "Location": "Delhi ", "MonthlySalary": 62000, "ExperienceYears": 3, "Status": "Active" },
        { "EmpID": "EMP-104", "FullName": "Priya Nair", "Department": "Engineering", "Location": "Bangalore", "MonthlySalary": 96000, "ExperienceYears": 8, "Status": "Active" },
        { "EmpID": "EMP-105", "FullName": "Vikram Patel", "Department": "Marketing", "Location": "Delhi", "MonthlySalary": 58000, "ExperienceYears": 2, "Status": "On Leave" },
        { "EmpID": "EMP-106", "FullName": "Ananya Roy", "Department": "Finance", "Location": "Bangalore", "MonthlySalary": 79000, "ExperienceYears": 5, "Status": "Active" }
      ]
    }
  },
  churn: {
    fileName: 'Customer_Churn_Retention.xlsx',
    sheets: {
      'Customers': [
        { "CustomerID": "C-101", "TenureMonths": 2, "MonthlySpend": 85.5, "SupportTickets": 4, "ContractType": "Month-to-Month", "PaymentMethod": "Electronic", "Churned": "Yes" },
        { "CustomerID": "C-102", "TenureMonths": 34, "MonthlySpend": 45.0, "SupportTickets": 0, "ContractType": "One-Year", "PaymentMethod": "Credit Card", "Churned": "No" },
        { "CustomerID": "C-103", "TenureMonths": 12, "MonthlySpend": 72.0, "SupportTickets": 2, "ContractType": "Month-to-Month", "PaymentMethod": "Electronic", "Churned": "Yes" },
        { "CustomerID": "C-104", "TenureMonths": 48, "MonthlySpend": 99.0, "SupportTickets": 1, "ContractType": "Two-Year", "PaymentMethod": "Bank Transfer", "Churned": "No" },
        { "CustomerID": "C-105", "TenureMonths": 5, "MonthlySpend": 64.0, "SupportTickets": 3, "ContractType": "Month-to-Month", "PaymentMethod": "Electronic", "Churned": "Yes" },
        { "CustomerID": "C-106", "TenureMonths": 22, "MonthlySpend": 52.5, "SupportTickets": 1, "ContractType": "One-Year", "PaymentMethod": "Credit Card", "Churned": "No" },
        { "CustomerID": "C-107", "TenureMonths": 1, "MonthlySpend": 91.0, "SupportTickets": 5, "ContractType": "Month-to-Month", "PaymentMethod": "Electronic", "Churned": "Yes" },
        { "CustomerID": "C-108", "TenureMonths": 60, "MonthlySpend": 115.0, "SupportTickets": 0, "ContractType": "Two-Year", "PaymentMethod": "Bank Transfer", "Churned": "No" },
        { "CustomerID": "C-109", "TenureMonths": 8, "MonthlySpend": 78.0, "SupportTickets": 2, "ContractType": "Month-to-Month", "PaymentMethod": "Electronic", "Churned": "Yes" },
        { "CustomerID": "C-110", "TenureMonths": 29, "MonthlySpend": 61.0, "SupportTickets": 1, "ContractType": "One-Year", "PaymentMethod": "Credit Card", "Churned": "No" },
        { "CustomerID": "C-111", "TenureMonths": 3, "MonthlySpend": 88.0, "SupportTickets": 4, "ContractType": "Month-to-Month", "PaymentMethod": "Electronic", "Churned": "Yes" },
        { "CustomerID": "C-112", "TenureMonths": 41, "MonthlySpend": 55.0, "SupportTickets": 0, "ContractType": "Two-Year", "PaymentMethod": "Bank Transfer", "Churned": "No" },
        { "CustomerID": "C-113", "TenureMonths": 16, "MonthlySpend": 70.0, "SupportTickets": 2, "ContractType": "Month-to-Month", "PaymentMethod": "Credit Card", "Churned": "No" },
        { "CustomerID": "C-114", "TenureMonths": 7, "MonthlySpend": 95.0, "SupportTickets": 3, "ContractType": "Month-to-Month", "PaymentMethod": "Electronic", "Churned": "Yes" },
        { "CustomerID": "C-115", "TenureMonths": 50, "MonthlySpend": 105.0, "SupportTickets": 1, "ContractType": "Two-Year", "PaymentMethod": "Bank Transfer", "Churned": "No" }
      ]
    }
  },
  forecast: {
    fileName: 'Monthly_Revenue_Demand_TimeSeries.xlsx',
    sheets: {
      'RevenueDemand': [
        { "Period": "2023-01", "Revenue": 42000, "AdSpend": 5000, "Orders": 820 },
        { "Period": "2023-02", "Revenue": 44500, "AdSpend": 5200, "Orders": 860 },
        { "Period": "2023-03", "Revenue": 47200, "AdSpend": 5500, "Orders": 910 },
        { "Period": "2023-04", "Revenue": 49000, "AdSpend": 5800, "Orders": 940 },
        { "Period": "2023-05", "Revenue": 52300, "AdSpend": 6100, "Orders": 1010 },
        { "Period": "2023-06", "Revenue": 56000, "AdSpend": 6600, "Orders": 1080 },
        { "Period": "2023-07", "Revenue": 59100, "AdSpend": 7000, "Orders": 1140 },
        { "Period": "2023-08", "Revenue": 61800, "AdSpend": 7300, "Orders": 1190 },
        { "Period": "2023-09", "Revenue": 64500, "AdSpend": 7600, "Orders": 1240 },
        { "Period": "2023-10", "Revenue": 68000, "AdSpend": 8000, "Orders": 1310 },
        { "Period": "2023-11", "Revenue": 74200, "AdSpend": 9200, "Orders": 1450 },
        { "Period": "2023-12", "Revenue": 81000, "AdSpend": 10500, "Orders": 1600 },
        { "Period": "2024-01", "Revenue": 71500, "AdSpend": 8200, "Orders": 1380 },
        { "Period": "2024-02", "Revenue": 74000, "AdSpend": 8500, "Orders": 1420 },
        { "Period": "2024-03", "Revenue": 78200, "AdSpend": 8900, "Orders": 1500 },
        { "Period": "2024-04", "Revenue": 82000, "AdSpend": 9300, "Orders": 1570 },
        { "Period": "2024-05", "Revenue": 86500, "AdSpend": 9800, "Orders": 1650 },
        { "Period": "2024-06", "Revenue": 91000, "AdSpend": 10400, "Orders": 1730 }
      ]
    }
  }
};
window.addEventListener('DOMContentLoaded', () => {
  lucide.createIcons();
  checkAuthState();
  loadSampleDataset('sales'); // Matches screenshot default (OrderID, Region, Category, Sales, Profit)

  // Silently wake up Render backend to eliminate cold-start lag
  fetch(`${AppState.apiBaseUrl}/health`).catch(() => {});

  // Dismiss dropdowns when clicking outside
  document.addEventListener('click', (e) => {
    const sBtn = document.getElementById('sampleDropdownBtn');
    const sMenu = document.getElementById('sampleDropdownMenu');
    if (sBtn && sMenu && !sBtn.contains(e.target) && !sMenu.contains(e.target)) {
      sMenu.classList.add('hidden');
    }

    const eBtn = document.getElementById('exportDropdownBtn');
    const eMenu = document.getElementById('exportDropdownMenu');
    if (eBtn && eMenu && !eBtn.contains(e.target) && !eMenu.contains(e.target)) {
      eMenu.classList.add('hidden');
    }
  });

  // Deep-linking via URL hash (e.g. #automl, #3dchart, #anomalies, #hygiene)
  const applyHashTab = () => {
    const hash = window.location.hash.replace('#', '');
    if (hash && document.getElementById(`tabBtn-${hash}`)) {
      switchTab(hash);
    }
  };
  applyHashTab();
  window.addEventListener('hashchange', applyHashTab);

  // Tabs Sidescroller Mouse Wheel & Dynamic Navigation
  const tabsNav = document.getElementById('platformTabsNav');
  if (tabsNav) {
    tabsNav.addEventListener('wheel', (e) => {
      if (e.deltaY !== 0) {
        e.preventDefault();
        tabsNav.scrollLeft += e.deltaY;
        updateTabsNavScrollButtons();
      }
    }, { passive: false });
    window.addEventListener('resize', updateTabsNavScrollButtons);
    setTimeout(updateTabsNavScrollButtons, 300);
  }
});

function toggleSampleMenu(e) {
  if (e && e.stopPropagation) e.stopPropagation();
  const menu = document.getElementById('sampleDropdownMenu');
  const exportMenu = document.getElementById('exportDropdownMenu');
  if (exportMenu) exportMenu.classList.add('hidden');
  if (menu) {
    menu.classList.toggle('hidden');
    lucide.createIcons();
  }
}

function toggleExportMenu(e) {
  if (e && e.stopPropagation) e.stopPropagation();
  const menu = document.getElementById('exportDropdownMenu');
  const sampleMenu = document.getElementById('sampleDropdownMenu');
  if (sampleMenu) sampleMenu.classList.add('hidden');
  if (menu) {
    menu.classList.toggle('hidden');
    lucide.createIcons();
  }
}

function loadSampleDataset(key) {
  const sample = SAMPLE_DATASETS[key];
  if (!sample) return;
  AppState.fileName = sample.fileName;
  AppState.sheets = JSON.parse(JSON.stringify(sample.sheets));
  AppState.activeSheet = Object.keys(sample.sheets)[0];
  loadActiveSheet();
  showToast(`Loaded ${sample.fileName}`);
  document.getElementById('sampleDropdownMenu').classList.add('hidden');
}

function switchSheet(sheetName) {
  if (!AppState.sheets[sheetName]) return;
  AppState.activeSheet = sheetName;
  loadActiveSheet();
}

function loadActiveSheet() {
  const rows = AppState.sheets[AppState.activeSheet] || [];
  AppState.originalRows = JSON.parse(JSON.stringify(rows));
  AppState.currentRows = JSON.parse(JSON.stringify(rows));
  AppState.filteredRows = [...AppState.currentRows];

  analyzeColumns();

  // Update banner indicators
  document.getElementById('bannerFileName').textContent = AppState.fileName;
  document.getElementById('bannerRowCount').textContent = `${AppState.currentRows.length} Rows`;
  document.getElementById('bannerColCount').textContent = `${AppState.columns.length} Columns`;

  renderSheetPills();
  populateDropdowns();

  // Reset grid pagination and filters
  AppState.grid.page = 1;
  AppState.grid.searchQuery = '';
  AppState.grid.columnFilter = 'all';
  AppState.grid.sortCol = null;
  document.getElementById('gridSearchInput').value = '';

  // Render tab panels
  renderGridData();
  renderColumnProfiles();
  runCategoryAggregation();
  renderStudioChart();
  updateKpiCards();
  populatePivotDropdowns();
  renderPivotMatrix();
  populateDiffDropdowns();
  populateAnomalyDropdown();
  AppState.automl.bestModel = null;
  AppState.automl.models = [];
  AppState.automl.testPredictions = [];
  populateAutoMlSelectors();
  updateSmartHygieneMetrics();

  const sheetBadgeName = document.getElementById('view3dActiveSheetName');
  if (sheetBadgeName) sheetBadgeName.textContent = AppState.activeSheet;
  const tab3d = document.getElementById('tabContent-3dchart');
  if (tab3d && !tab3d.classList.contains('hidden')) {
    renderDedicated3DView();
  }
  const tabMl = document.getElementById('tabContent-automl');
  if (tabMl && !tabMl.classList.contains('hidden')) {
    resizeAutoMlCharts();
  }

  lucide.createIcons();
}

function renderSheetPills() {
  const container = document.getElementById('sheetPillContainer');
  container.innerHTML = '';
  const sheetNames = Object.keys(AppState.sheets);

  sheetNames.forEach(name => {
    const btn = document.createElement('button');
    const isActive = name === AppState.activeSheet;
    btn.onclick = () => switchSheet(name);
    btn.className = `px-3.5 py-1 rounded-full text-xs font-semibold transition shrink-0 ${
      isActive
        ? 'bg-indigo-600 text-white shadow-xs'
        : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
    }`;
    btn.textContent = name;
    container.appendChild(btn);
  });
}

function analyzeColumns() {
  if (AppState.currentRows.length === 0) {
    AppState.columns = [];
    AppState.colMeta = {};
    return;
  }

  AppState.columns = Object.keys(AppState.currentRows[0]);
  AppState.colMeta = {};

  AppState.columns.forEach(col => {
    let nonNull = 0;
    let numCount = 0;
    let sum = 0;
    let min = Infinity;
    let max = -Infinity;
    const freqMap = {};

    AppState.currentRows.forEach(row => {
      const val = row[col];
      if (!isBlankValue(val)) {
        nonNull++;
        const strVal = String(val).trim();
        freqMap[strVal] = (freqMap[strVal] || 0) + 1;

        const n = toNumberOrNull(val);
        if (n !== null) {
          numCount++;
          sum += n;
          if (n < min) min = n;
          if (n > max) max = n;
        }
      }
    });

    // Column typing
    const isNumeric = nonNull > 0 && (numCount / nonNull) > 0.6;
    const distinctKeys = Object.keys(freqMap);
    let modeKey = '-';
    let modeCount = 0;

    distinctKeys.forEach(k => {
      if (freqMap[k] > modeCount) {
        modeCount = freqMap[k];
        modeKey = k;
      }
    });

    AppState.colMeta[col] = {
      type: isNumeric ? 'numeric' : 'categorical',
      nonNull,
      nullCount: AppState.currentRows.length - nonNull,
      percentFilled: Math.round((nonNull / (AppState.currentRows.length || 1)) * 100),
      distinctCount: distinctKeys.length,
      sum: isNumeric ? sum : null,
      avg: isNumeric && numCount > 0 ? (sum / numCount) : null,
      min: isNumeric && numCount > 0 ? min : null,
      max: isNumeric && numCount > 0 ? max : null,
      mode: modeKey,
      modeCount
    };
  });
}

function populateDropdowns() {
  // 1. Grid Column Filter
  const gridColSelect = document.getElementById('gridColumnFilterSelect');
  gridColSelect.innerHTML = '<option value="all">All Columns</option>';
  AppState.columns.forEach(col => {
    gridColSelect.innerHTML += `<option value="${col}">${col}</option>`;
  });

  // 2. Category Aggregation Selectors
  const catGroup = document.getElementById('categoryGroupSelect');
  const catMetric = document.getElementById('categoryMetricSelect');
  catGroup.innerHTML = '';
  catMetric.innerHTML = '';

  AppState.columns.forEach(col => {
    catGroup.innerHTML += `<option value="${col}">${col}</option>`;
  });

  const numCols = AppState.columns.filter(c => AppState.colMeta[c]?.type === 'numeric');
  (numCols.length > 0 ? numCols : AppState.columns).forEach(col => {
    catMetric.innerHTML += `<option value="${col}">${col}</option>`;
  });

  // Defaults
  const catCol = AppState.columns.find(c => AppState.colMeta[c]?.type === 'categorical');
  if (catCol) catGroup.value = catCol;
  if (numCols.length > 0) catMetric.value = numCols[0];

  // 3. Studio Chart Selectors
  const studioX = document.getElementById('studioXCol');
  const studioY = document.getElementById('studioYCol');
  studioX.innerHTML = '';
  studioY.innerHTML = '';

  AppState.columns.forEach(c => studioX.innerHTML += `<option value="${c}">${c}</option>`);
  (numCols.length > 0 ? numCols : AppState.columns).forEach(c => studioY.innerHTML += `<option value="${c}">${c}</option>`);

  if (catCol) studioX.value = catCol;
  if (numCols.length > 0) studioY.value = numCols[0];

  // 4. Dedicated 3D View Selectors
  const view3dX = document.getElementById('view3dXCol');
  const view3dY = document.getElementById('view3dYCol');
  if (view3dX && view3dY) {
    view3dX.innerHTML = '';
    view3dY.innerHTML = '';
    AppState.columns.forEach(c => view3dX.innerHTML += `<option value="${c}">${c}</option>`);
    (numCols.length > 0 ? numCols : AppState.columns).forEach(c => view3dY.innerHTML += `<option value="${c}">${c}</option>`);
    if (catCol) view3dX.value = catCol;
    if (numCols.length > 0) view3dY.value = numCols[0];
  }
}

// Exact recreation of the Column Profiles & Data Health cards shown in screenshot
function renderColumnProfiles() {
  const container = document.getElementById('columnProfilesContainer');
  container.innerHTML = '';

  AppState.columns.forEach(col => {
    const meta = AppState.colMeta[col];
    const isNumeric = meta.type === 'numeric';

    // Format numbers cleanly (e.g. 3100 -> 3.1k)
    const formatVal = (val) => {
      if (val === null || val === undefined || isNaN(val)) return '-';
      if (Math.abs(val) >= 1000000) return (val / 1000000).toFixed(1) + 'M';
      if (Math.abs(val) >= 1000) return (val / 1000).toFixed(1) + 'k';
      return Number.isInteger(val) ? val : val.toFixed(1);
    };

    const card = document.createElement('div');
    card.className = 'bg-white rounded-2xl border border-slate-200/80 p-4 sm:p-5 shadow-[0_1px_3px_rgba(0,0,0,0.02)] space-y-3';

    // Pill badge colors matching screenshot
    const badgeHtml = isNumeric
      ? `<span class="text-[11px] font-semibold px-2.5 py-0.5 rounded-full bg-amber-50 text-amber-600 border border-amber-200/70">numeric</span>`
      : `<span class="text-[11px] font-semibold px-2.5 py-0.5 rounded-full bg-indigo-50 text-indigo-600 border border-indigo-200/70">categorical</span>`;

    let metricsHtml = '';
    if (isNumeric) {
      metricsHtml = `
        <div class="flex justify-between items-center text-xs font-mono text-slate-500">
          <span>Filled:</span>
          <span class="font-medium text-slate-800">${meta.nonNull} (${meta.percentFilled}%)</span>
        </div>
        <div class="flex justify-between items-center text-xs font-mono text-slate-500">
          <span>Unique:</span>
          <span class="font-medium text-slate-800">${meta.distinctCount}</span>
        </div>
        <div class="flex justify-between items-center text-xs font-mono text-slate-500">
          <span>Min:</span>
          <span class="font-medium text-slate-800">${formatVal(meta.min)}</span>
        </div>
        <div class="flex justify-between items-center text-xs font-mono text-slate-500">
          <span>Max:</span>
          <span class="font-medium text-slate-800">${formatVal(meta.max)}</span>
        </div>
        <div class="flex justify-between items-center text-xs font-mono text-slate-500">
          <span>Mean:</span>
          <span class="font-medium text-indigo-600">${formatVal(meta.avg)}</span>
        </div>
        <div class="flex justify-between items-center text-xs font-mono text-slate-500">
          <span>Sum:</span>
          <span class="font-medium text-slate-800">${formatVal(meta.sum)}</span>
        </div>
      `;
    } else {
      metricsHtml = `
        <div class="flex justify-between items-center text-xs font-mono text-slate-500">
          <span>Filled:</span>
          <span class="font-medium text-slate-800">${meta.nonNull} (${meta.percentFilled}%)</span>
        </div>
        <div class="flex justify-between items-center text-xs font-mono text-slate-500">
          <span>Unique:</span>
          <span class="font-medium text-slate-800">${meta.distinctCount}</span>
        </div>
        <div class="flex justify-between items-center text-xs font-mono text-slate-500">
          <span>Top Value:</span>
          <span class="font-medium text-slate-800 truncate max-w-[150px]">${meta.mode}</span>
        </div>
        <div class="flex justify-between items-center text-xs font-mono text-slate-500">
          <span>Frequency:</span>
          <span class="font-medium text-slate-800">${meta.modeCount}x</span>
        </div>
      `;
    }

    card.innerHTML = `
      <div class="flex items-center justify-between pb-1 border-b border-slate-100">
        <h4 class="font-bold text-sm text-slate-900 truncate" title="${col}">${col}</h4>
        ${badgeHtml}
      </div>

      <div class="space-y-1.5 pt-1">
        ${metricsHtml}
      </div>

      <!-- Completeness Bar -->
      <div class="w-full bg-slate-100 h-1.5 rounded-full overflow-hidden mt-2">
        <div class="h-full ${meta.percentFilled === 100 ? 'bg-emerald-500' : 'bg-amber-500'}" style="width: ${meta.percentFilled}%"></div>
      </div>
    `;

    container.appendChild(card);
  });
}

// Quick Data Transformations Handlers
function rowSignature(row) {
  // Two rows are duplicates when every cell matches after trimming
  return JSON.stringify(AppState.columns.map(c => isBlankValue(row[c]) ? '' : String(row[c]).trim()));
}

function hygieneRemoveDuplicates() {
  const initial = AppState.currentRows.length;
  const seen = new Set();
  AppState.currentRows = AppState.currentRows.filter(row => {
    const s = rowSignature(row);
    if (seen.has(s)) return false;
    seen.add(s);
    return true;
  });
  const removed = initial - AppState.currentRows.length;
  refreshPostTransformation(removed ? `Removed ${removed} duplicate row(s)` : 'No duplicate rows found');
}

function hygieneStripNullRows() {
  const initial = AppState.currentRows.length;
  AppState.currentRows = AppState.currentRows.filter(row => AppState.columns.some(c => !isBlankValue(row[c])));
  const removed = initial - AppState.currentRows.length;
  refreshPostTransformation(removed ? `Removed ${removed} empty row(s)` : 'No empty rows found');
}

// Trim spaces and merge case variants ("saket", " Saket ") into the most common spelling
function standardizeTextValues(rows) {
  let trimmed = 0;
  let merged = 0;
  const textCols = AppState.columns.filter(c => AppState.colMeta[c]?.type !== 'numeric');

  rows.forEach(row => {
    Object.keys(row).forEach(k => {
      if (typeof row[k] === 'string') {
        const clean = row[k].trim().replace(/\s+/g, ' ');
        if (clean !== row[k]) { row[k] = clean; trimmed++; }
      }
    });
  });

  textCols.forEach(col => {
    const variants = {};
    rows.forEach(row => {
      const v = row[col];
      if (typeof v !== 'string' || !v) return;
      const key = v.toLowerCase();
      variants[key] = variants[key] || {};
      variants[key][v] = (variants[key][v] || 0) + 1;
    });
    const canonical = {};
    Object.entries(variants).forEach(([key, forms]) => {
      const names = Object.keys(forms);
      if (names.length < 2) return;
      // Most frequent spelling wins; ties go to the one that starts with a capital letter
      names.sort((a, b) => forms[b] - forms[a] || (/^[A-Z]/.test(b) ? 1 : 0) - (/^[A-Z]/.test(a) ? 1 : 0));
      canonical[key] = names[0];
    });
    rows.forEach(row => {
      const v = row[col];
      if (typeof v === 'string' && canonical[v.toLowerCase()] && canonical[v.toLowerCase()] !== v) {
        row[col] = canonical[v.toLowerCase()];
        merged++;
      }
    });
  });
  return { trimmed, merged };
}

function hygieneTrimWhitespace() {
  const { trimmed, merged } = standardizeTextValues(AppState.currentRows);
  refreshPostTransformation(`Trimmed ${trimmed} cell(s) and merged ${merged} case variant(s)`);
}

function hygieneRestoreOriginal() {
  AppState.currentRows = JSON.parse(JSON.stringify(AppState.originalRows));
  refreshPostTransformation('Restored the original sheet');
}

// Re-run every view that depends on the data, keeping the user's dropdown choices
function refreshPostTransformation(message, opts = {}) {
  const keep = ['categoryGroupSelect', 'categoryMetricSelect', 'studioXCol', 'studioYCol', 'view3dXCol', 'view3dYCol',
    'pivotRowSelect', 'pivotColSelect', 'pivotValSelect', 'anomalyColSelect', 'gridColumnFilterSelect']
    .map(id => [id, document.getElementById(id)?.value]);
  analyzeColumns();
  populateDropdowns();
  populatePivotDropdowns();
  keep.forEach(([id, val]) => {
    const el = document.getElementById(id);
    if (el && val && [...el.options].some(o => o.value === val)) el.value = val;
  });
  document.getElementById('bannerRowCount').textContent = `${AppState.currentRows.length} Rows`;
  document.getElementById('bannerColCount').textContent = `${AppState.columns.length} Columns`;
  updateKpiCards();
  if (!opts.keepGrid) renderGridData();
  renderColumnProfiles();
  runCategoryAggregation();
  renderStudioChart();
  renderPivotMatrix();
  runAnomalyDetection();
  updateSmartHygieneMetrics();
  const tab3d = document.getElementById('tabContent-3dchart');
  if (tab3d && !tab3d.classList.contains('hidden')) renderDedicated3DView();
  if (message) showToast(message);
}

// Grid Rendering
function renderGridData() {
  const colFilter = document.getElementById('gridColumnFilterSelect').value;
  const search = (document.getElementById('gridSearchInput').value || '').toLowerCase().trim();

  let list = [...AppState.currentRows];
  if (search) {
    list = list.filter(row => {
      if (colFilter === 'all') {
        return AppState.columns.some(col => String(row[col] ?? '').toLowerCase().includes(search));
      } else {
        return String(row[colFilter] ?? '').toLowerCase().includes(search);
      }
    });
  }

  if (AppState.grid.sortCol) {
    const col = AppState.grid.sortCol;
    const dir = AppState.grid.sortDir === 'asc' ? 1 : -1;
    const isNum = AppState.colMeta[col]?.type === 'numeric';

    list.sort((a, b) => {
      const valA = a[col];
      const valB = b[col];
      // Blank cells always sort last, whichever direction
      const aBlank = isBlankValue(valA), bBlank = isBlankValue(valB);
      if (aBlank || bBlank) return aBlank === bBlank ? 0 : (aBlank ? 1 : -1);
      if (isNum) return ((toNumberOrNull(valA) ?? 0) - (toNumberOrNull(valB) ?? 0)) * dir;
      return String(valA).localeCompare(String(valB), undefined, { numeric: true, sensitivity: 'base' }) * dir;
    });
  }

  AppState.filteredRows = list;

  document.getElementById('gridShowingCountLabel').textContent =
    `Showing ${list.length} of ${AppState.currentRows.length}`;

  renderGridHeaders();
  renderGridRows();
  updateGridPagination();
}

function renderGridHeaders() {
  const tr = document.getElementById('gridHeaderRow');
  tr.innerHTML = '';

  const thIdx = document.createElement('th');
  thIdx.className = 'py-3 px-3.5 w-12 text-slate-400 font-semibold border-r border-slate-100/80';
  thIdx.textContent = '#';
  tr.appendChild(thIdx);

  AppState.columns.forEach(col => {
    const meta = AppState.colMeta[col];
    const isNumeric = meta?.type === 'numeric';
    const isSorted = AppState.grid.sortCol === col;
    const th = document.createElement('th');
    th.className = 'py-3 px-3.5 font-semibold text-slate-700 whitespace-nowrap cursor-pointer hover:bg-slate-100 transition select-none';
    th.onclick = () => sortGridBy(col);

    let sortIcon = '<i data-lucide="chevrons-up-down" class="w-3 h-3 text-slate-300 inline ml-1"></i>';
    if (isSorted) {
      sortIcon = AppState.grid.sortDir === 'asc'
        ? '<i data-lucide="chevron-up" class="w-3 h-3 text-indigo-600 inline ml-1"></i>'
        : '<i data-lucide="chevron-down" class="w-3 h-3 text-indigo-600 inline ml-1"></i>';
    }

    const typeBadge = isNumeric
      ? '<span class="text-[9.5px] font-mono font-bold uppercase tracking-wider px-1.5 py-0.5 rounded bg-amber-50 text-amber-700 border border-amber-200/60">123</span>'
      : '<span class="text-[9.5px] font-mono font-bold uppercase tracking-wider px-1.5 py-0.5 rounded bg-blue-50 text-blue-700 border border-blue-200/60">ABC</span>';

    th.innerHTML = `
      <div class="flex items-center justify-between gap-3">
        <span>${col}</span>
        <div class="flex items-center gap-1.5">
          ${typeBadge}
          ${sortIcon}
        </div>
      </div>
    `;
    tr.appendChild(th);
  });

  lucide.createIcons();
}

function renderGridRows() {
  const tbody = document.getElementById('gridTableBody');
  tbody.innerHTML = '';

  const { page, pageSize } = AppState.grid;
  const start = (page - 1) * pageSize;
  const end = start + pageSize;
  const pageRows = AppState.filteredRows.slice(start, end);

  if (pageRows.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="${AppState.columns.length + 1}" class="py-12 text-center text-slate-400">
          No matching records found.
        </td>
      </tr>
    `;
    return;
  }

  pageRows.forEach((row, i) => {
    const rowNumber = start + i + 1;
    const tr = document.createElement('tr');
    tr.className = 'hover:bg-indigo-50/30 transition-colors border-b border-slate-100';

    let html = `<td class="py-2.5 px-3.5 text-slate-400 font-mono text-[11px] border-r border-slate-100/80">${rowNumber}</td>`;

    AppState.columns.forEach(col => {
      const val = row[col];
      const isNum = AppState.colMeta[col]?.type === 'numeric';
      const displayVal = isBlankValue(val)
        ? '<span class="cell-empty" title="Empty cell">—</span>'
        : escapeHtml(val);

      html += `<td data-row-idx="${start + i}" data-col="${col}" ondblclick="makeCellEditable(this)" title="Double click to edit" class="py-2.5 px-3.5 whitespace-nowrap text-slate-700 cursor-cell hover:bg-amber-50/70 transition-colors ${isNum ? 'font-mono text-right' : ''}">${displayVal}</td>`;
    });

    tr.innerHTML = html;
    tbody.appendChild(tr);
  });
}

function sortGridBy(col) {
  if (AppState.grid.sortCol === col) {
    AppState.grid.sortDir = AppState.grid.sortDir === 'asc' ? 'desc' : 'asc';
  } else {
    AppState.grid.sortCol = col;
    AppState.grid.sortDir = 'asc';
  }
  renderGridData();
}

function handleGridSearch(val) {
  AppState.grid.searchQuery = val;
  AppState.grid.page = 1;
  renderGridData();
}

function resetGridSearch() {
  document.getElementById('gridSearchInput').value = '';
  document.getElementById('gridColumnFilterSelect').value = 'all';
  AppState.grid.sortCol = null;
  AppState.grid.page = 1;
  renderGridData();
}

function updateGridPagination() {
  const total = AppState.filteredRows.length;
  const totalPages = Math.ceil(total / AppState.grid.pageSize) || 1;
  document.getElementById('gridPageIndicator').textContent = `Page ${AppState.grid.page} of ${totalPages}`;
  document.getElementById('btnGridPrev').disabled = AppState.grid.page <= 1;
  document.getElementById('btnGridNext').disabled = AppState.grid.page >= totalPages;
}

function navigateGridPage(delta) {
  const totalPages = Math.ceil(AppState.filteredRows.length / AppState.grid.pageSize) || 1;
  const newPage = AppState.grid.page + delta;
  if (newPage >= 1 && newPage <= totalPages) {
    AppState.grid.page = newPage;
    renderGridRows();
    updateGridPagination();
  }
}

function changeGridPageSize(size) {
  AppState.grid.pageSize = Number(size);
  AppState.grid.page = 1;
  renderGridData();
}

// Category Summary & Averages Engine
function runCategoryAggregation() {
  const groupCol = document.getElementById('categoryGroupSelect').value;
  const metricCol = document.getElementById('categoryMetricSelect').value;

  if (!groupCol || !metricCol) return;

  const groupMap = {};
  const allNums = [];

  AppState.currentRows.forEach(row => {
    const key = groupKeyOf(row[groupCol]);
    if (!groupMap[key]) groupMap[key] = { count: 0, nums: [] };
    groupMap[key].count++;
    const n = toNumberOrNull(row[metricCol]);
    if (n !== null) { groupMap[key].nums.push(n); allNums.push(n); }
  });

  const grandTotalSum = aggregateNumbers(allNums, 'sum') || 0;
  const categories = Object.keys(groupMap).sort((a, b) => a === '(Blank)' ? 1 : b === '(Blank)' ? -1 : a.localeCompare(b));
  const tbody = document.getElementById('categoryTableBody');
  const tfoot = document.getElementById('categoryTableFoot');
  tbody.innerHTML = '';
  tfoot.innerHTML = '';

  const chartLabels = [];
  const chartAverages = [];

  categories.forEach(cat => {
    const item = groupMap[cat];
    const sum = aggregateNumbers(item.nums, 'sum');
    const avg = aggregateNumbers(item.nums, 'avg');
    const pctOfTotal = grandTotalSum !== 0 && sum !== null ? (sum / grandTotalSum) * 100 : 0;

    chartLabels.push(cat);
    chartAverages.push(avg === null ? null : parseFloat(avg.toFixed(2)));

    const tr = document.createElement('tr');
    tr.className = 'hover:bg-slate-50 transition border-b border-slate-100';
    const blanks = item.count - item.nums.length;
    tr.innerHTML = `
      <td class="py-2.5 px-3 font-sans font-semibold text-slate-900">${escapeHtml(cat)}</td>
      <td class="py-2.5 px-3">${item.count.toLocaleString()}${blanks ? ` <span class="text-[10px] text-slate-400" title="Rows with no ${escapeHtml(metricCol)} value are left out of the average">(${blanks} blank)</span>` : ''}</td>
      <td class="py-2.5 px-3 font-semibold">${formatNumber(sum, 1)}</td>
      <td class="py-2.5 px-3 font-bold text-indigo-600">${formatNumber(avg)}</td>
      <td class="py-2.5 px-3 text-slate-500">${formatNumber(aggregateNumbers(item.nums, 'min'))}</td>
      <td class="py-2.5 px-3 text-slate-500">${formatNumber(aggregateNumbers(item.nums, 'max'))}</td>
      <td class="py-2.5 px-3 text-right font-medium">${pctOfTotal.toFixed(1)}%</td>
    `;
    tbody.appendChild(tr);
  });

  const grandTotalCount = AppState.currentRows.length;
  tfoot.innerHTML = `
    <tr>
      <td class="py-2.5 px-3">Grand Total (${categories.length} groups)</td>
      <td class="py-2.5 px-3">${grandTotalCount.toLocaleString()}</td>
      <td class="py-2.5 px-3">${formatNumber(grandTotalSum, 1)}</td>
      <td class="py-2.5 px-3 text-indigo-700">${formatNumber(aggregateNumbers(allNums, 'avg'))}</td>
      <td class="py-2.5 px-3">${formatNumber(aggregateNumbers(allNums, 'min'))}</td>
      <td class="py-2.5 px-3">${formatNumber(aggregateNumbers(allNums, 'max'))}</td>
      <td class="py-2.5 px-3 text-right">100.0%</td>
    </tr>
  `;

  document.getElementById('categoryChartTitle').textContent = `Average ${metricCol} by ${groupCol}`;

  const ctx = document.getElementById('categoryChartCanvas').getContext('2d');
  if (AppState.categoryChartInstance) AppState.categoryChartInstance.destroy();

  AppState.categoryChartInstance = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: chartLabels,
      datasets: [{
        label: `Average ${metricCol}`,
        data: chartAverages,
        backgroundColor: '#0d44e8',
        borderRadius: 6,
        barThickness: categories.length > 8 ? undefined : 28
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (ctx) => ` Average ${metricCol}: ${ctx.raw}`
          }
        }
      },
      scales: {
        y: { ticks: { font: { family: 'Inter', size: 10 } }, grid: { color: '#f1f5f9' } },
        x: { grid: { display: false }, ticks: { font: { family: 'Inter', size: 10 } } }
      }
    }
  });
}

// Visual Mode Switcher (2D vs 3D)
function setVisualMode(mode) {
  const btn2D = document.getElementById('visualMode2DBtn');
  const btn3D = document.getElementById('visualMode3DBtn');
  const select = document.getElementById('studioChartType');
  if (!select) return;

  if (mode === '3d') {
    if (btn2D && btn3D) {
      btn2D.className = 'flex-1 py-1.5 px-2 rounded-lg text-xs font-semibold flex items-center justify-center gap-1.5 transition text-slate-600 hover:text-slate-900';
      btn3D.className = 'flex-1 py-1.5 px-2 rounded-lg text-xs font-semibold flex items-center justify-center gap-1.5 transition bg-white text-indigo-700 shadow-xs';
    }
    if (!select.value.startsWith('3d-')) {
      select.value = '3d-column';
    }
  } else {
    if (btn2D && btn3D) {
      btn3D.className = 'flex-1 py-1.5 px-2 rounded-lg text-xs font-semibold flex items-center justify-center gap-1.5 transition text-slate-600 hover:text-slate-900';
      btn2D.className = 'flex-1 py-1.5 px-2 rounded-lg text-xs font-semibold flex items-center justify-center gap-1.5 transition bg-white text-indigo-700 shadow-xs';
    }
    if (select.value.startsWith('3d-')) {
      select.value = 'bar';
    }
  }
  renderStudioChart();
  lucide.createIcons();
}

// Chart Studio with 2D Chart.js and 3D Three.js
function renderStudioChart() {
  const xCol = document.getElementById('studioXCol').value;
  const yCol = document.getElementById('studioYCol').value;
  const type = document.getElementById('studioChartType').value;
  const agg = document.getElementById('studioAgg').value;

  if (!xCol || !yCol) return;

  const canvas = document.getElementById('studioChartCanvas');
  const threeContainer = document.getElementById('threeJsCanvasContainer');
  const threeOptions = document.getElementById('threeDOptionsSection');
  const btn2D = document.getElementById('visualMode2DBtn');
  const btn3D = document.getElementById('visualMode3DBtn');

  // Handle 3D Three.js Visualization
  if (type.startsWith('3d-')) {
    if (btn2D && btn3D) {
      btn2D.className = 'flex-1 py-1.5 px-2 rounded-lg text-xs font-semibold flex items-center justify-center gap-1.5 transition text-slate-600 hover:text-slate-900';
      btn3D.className = 'flex-1 py-1.5 px-2 rounded-lg text-xs font-semibold flex items-center justify-center gap-1.5 transition bg-white text-indigo-700 shadow-xs';
    }
    canvas.classList.add('hidden');
    threeContainer.classList.remove('hidden');
    if (threeOptions) threeOptions.classList.remove('hidden');
    if (AppState.studioChartInstance) {
      AppState.studioChartInstance.destroy();
      AppState.studioChartInstance = null;
    }
    renderThreeJsScene(type, xCol, yCol, agg);

    const typeTitles = {
      '3d-column': '3D Column Grid',
      '3d-scatter': '3D Bubble & Scatter Universe',
      '3d-cylinder': '3D Radial Cylinder Donut',
      '3d-surface': '3D Surface Elevation Mesh',
      '3d-matrix': '3D Stepped Grid Waterfall',
      '3d-pyramid': '3D Pyramid & Funnel',
      '3d-pie': '3D Floating Donut Ring'
    };
    const titleText = typeTitles[type] || '3D WebGL Visualizer';
    document.getElementById('studioChartTitle').textContent = `${titleText}: ${agg.toUpperCase()} of ${yCol} by ${xCol}`;
    document.getElementById('studioChartDesc').textContent = 'Rendered in interactive WebGL with real-time lighting, orbital controls & data inspection HUD';
    return;
  }

  // Restore 2D Canvas
  if (btn2D && btn3D) {
    btn3D.className = 'flex-1 py-1.5 px-2 rounded-lg text-xs font-semibold flex items-center justify-center gap-1.5 transition text-slate-600 hover:text-slate-900';
    btn2D.className = 'flex-1 py-1.5 px-2 rounded-lg text-xs font-semibold flex items-center justify-center gap-1.5 transition bg-white text-indigo-700 shadow-xs';
  }
  threeContainer.classList.add('hidden');
  if (threeOptions) threeOptions.classList.add('hidden');
  canvas.classList.remove('hidden');
  cleanupThreeJs();

  const grouped = {};
  AppState.currentRows.forEach(row => {
    const k = groupKeyOf(row[xCol]);
    if (!grouped[k]) grouped[k] = { count: 0, nums: [] };
    grouped[k].count++;
    const v = toNumberOrNull(row[yCol]);
    if (v !== null) grouped[k].nums.push(v);
  });

  const labels = Object.keys(grouped);
  const data = labels.map(k => {
    const val = aggregateNumbers(grouped[k].nums, agg, grouped[k].count);
    return val === null ? null : parseFloat(val.toFixed(2));
  });

  const ctx = canvas.getContext('2d');
  if (AppState.studioChartInstance) AppState.studioChartInstance.destroy();

  const palette = ['#0d44e8', '#5a84fb', '#0f172a', '#94a3b8', '#10b981', '#f59e0b', '#8eadff', '#334155', '#14b8a6', '#cbd5e1'];

  let chartConfigType = type;
  let chartDatasets = [];

  if (type === 'area') {
    chartConfigType = 'line';
    chartDatasets = [{
      label: `${agg.toUpperCase()}(${yCol})`,
      data: data,
      fill: true,
      backgroundColor: 'rgba(13, 68, 232, 0.14)',
      borderColor: '#0d44e8',
      tension: 0.35,
      borderWidth: 2
    }];
  } else if (type === 'scatter') {
    chartConfigType = 'scatter';
    const scatterPoints = AppState.currentRows
      .map(r => ({ x: toNumberOrNull(r[xCol]), y: toNumberOrNull(r[yCol]) }))
      .filter(p => p.x !== null && p.y !== null);

    chartDatasets = [{
      label: `${yCol} vs ${xCol}`,
      data: scatterPoints,
      backgroundColor: '#0d44e8',
      borderColor: '#0a36c0',
      pointRadius: 5
    }];
  } else if (type === 'radar') {
    chartConfigType = 'radar';
    chartDatasets = [{
      label: `${agg.toUpperCase()}(${yCol})`,
      data: data,
      backgroundColor: 'rgba(13, 68, 232, 0.14)',
      borderColor: '#0d44e8',
      borderWidth: 2
    }];
  } else {
    chartDatasets = [{
      label: `${agg.toUpperCase()}(${yCol})`,
      data: data,
      backgroundColor: ['doughnut', 'pie'].includes(type) ? palette : '#0d44e8',
      borderWidth: 1.5,
      borderColor: '#ffffff',
      borderRadius: type === 'bar' ? 6 : 0
    }];
  }

  AppState.studioChartInstance = new Chart(ctx, {
    type: chartConfigType,
    data: {
      labels: type === 'scatter' ? undefined : labels,
      datasets: chartDatasets
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          display: ['doughnut', 'pie', 'radar'].includes(type),
          position: 'bottom',
          labels: { boxWidth: 10, font: { family: 'Inter', size: 10 } }
        }
      },
      scales: ['doughnut', 'pie', 'radar'].includes(type) ? {} : {
        y: { grid: { color: '#f1f5f9' }, ticks: { font: { family: 'Inter', size: 10 } } },
        x: { grid: { display: false }, ticks: { font: { family: 'Inter', size: 10 } } }
      }
    }
  });

  document.getElementById('studioChartTitle').textContent = `${agg.toUpperCase()} of ${yCol} by ${xCol}`;
  document.getElementById('studioChartDesc').textContent = `Plotted across ${labels.length} segments with dynamic calculations`;
}

function downloadStudioChart() {
  const chartType = document.getElementById('studioChartType')?.value || '';
  if (chartType.startsWith('3d-') && AppState.threeJs.renderer && AppState.threeJs.scene && AppState.threeJs.camera) {
    AppState.threeJs.renderer.render(AppState.threeJs.scene, AppState.threeJs.camera);
    const url = AppState.threeJs.renderer.domElement.toDataURL('image/png');
    const a = document.createElement('a');
    a.href = url;
    a.download = `DataVista_3D_${chartType}_${Date.now()}.png`;
    a.click();
    showToast('3D WebGL chart snapshot downloaded as PNG');
    return;
  }
  if (!AppState.studioChartInstance) return;
  const url = AppState.studioChartInstance.toBase64Image();
  const a = document.createElement('a');
  a.href = url;
  a.download = `DataVista_Chart_${Date.now()}.png`;
  a.click();
  showToast('Chart snapshot downloaded as PNG');
}

// File Upload Handler (SheetJS)
function handleFileUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  AppState.fileName = file.name;
  const reader = new FileReader();

  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      AppState.sheets = {};
      workbook.SheetNames.forEach(sheetName => {
        AppState.sheets[sheetName] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
      });

      AppState.activeSheet = workbook.SheetNames[0];
      loadActiveSheet();
      showToast(`Uploaded ${file.name}`);
    } catch (err) {
      console.error(err);
      showToast('Failed to parse uploaded Excel file', true);
    }
  };

  reader.readAsArrayBuffer(file);
  event.target.value = '';
}

// Export Options
function exportCurrentData(format) {
  const ext = format === 'csv' ? 'csv' : format === 'xlsx' ? 'xlsx' : 'json';
  if (ext === 'json') {
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(AppState.filteredRows, null, 2));
    const a = document.createElement('a');
    a.href = dataStr;
    a.download = `DataVista_${AppState.activeSheet}_${Date.now()}.json`;
    a.click();
  } else {
    const ws = XLSX.utils.json_to_sheet(AppState.filteredRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, AppState.activeSheet);
    XLSX.writeFile(wb, `DataVista_${AppState.activeSheet}_${Date.now()}.${ext}`);
  }
  showToast(`Exported ${AppState.filteredRows.length} records to .${ext}`);
  document.getElementById('exportDropdownMenu').classList.add('hidden');
}

// Tab Navigation Switcher
function switchTab(tabId) {
  if (tabId === 'admin' && AppState.auth.user?.role !== 'admin') {
    showToast('The admin panel is only available to admin accounts', true);
    return;
  }
  document.querySelectorAll('.tab-panel').forEach(el => el.classList.add('hidden'));
  document.querySelectorAll('.tab-nav-btn').forEach(btn => {
    btn.classList.remove('text-indigo-600', 'border-indigo-600');
    btn.classList.add('text-slate-500', 'border-transparent');
  });

  const targetPane = document.getElementById(`tabContent-${tabId}`);
  const targetBtn = document.getElementById(`tabBtn-${tabId}`);

  if (targetPane) targetPane.classList.remove('hidden');
  if (targetBtn) {
    targetBtn.classList.add('text-indigo-600', 'border-indigo-600');
    targetBtn.classList.remove('text-slate-500', 'border-transparent');
    try {
      targetBtn.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'center' });
    } catch (_) {}
  }

  if (tabId === 'category' && AppState.categoryChartInstance) {
    AppState.categoryChartInstance.resize();
  }
  if (tabId === 'charts') {
    if (AppState.studioChartInstance) {
      AppState.studioChartInstance.resize();
    }
    const chartType = document.getElementById('studioChartType')?.value || '';
    if (chartType.startsWith('3d-')) {
      setTimeout(() => resizeThreeJs(), 60);
    }
  }
  if (tabId === '3dchart') {
    renderDedicated3DView();
    setTimeout(() => resizeDedicated3D(), 60);
  }
  if (tabId === 'automl') {
    if (!document.getElementById('mlTargetColSelect')?.value) {
      populateAutoMlSelectors();
    }
    setTimeout(() => {
      resizeAutoMlCharts();
      if (!AppState.automl.bestModel && AppState.currentRows && AppState.currentRows.length > 3 && !AppState.automl.isTraining) {
        runAutoMlTraining();
      }
    }, 80);
  }
  if (tabId === 'history' && AppState.auth.token) {
    fetchUploadHistory();
  }
  if (tabId === 'admin') {
    fetchAdminData();
  }

  setTimeout(updateTabsNavScrollButtons, 120);
  lucide.createIcons();
}

function scrollTabsNav(offset) {
  const nav = document.getElementById('platformTabsNav');
  if (nav) {
    nav.scrollBy({ left: offset, behavior: 'smooth' });
    setTimeout(updateTabsNavScrollButtons, 220);
  }
}

function updateTabsNavScrollButtons() {
  const nav = document.getElementById('platformTabsNav');
  const leftBtn = document.getElementById('tabNavScrollLeftBtn');
  const rightBtn = document.getElementById('tabNavScrollRightBtn');
  if (!nav || !leftBtn || !rightBtn) return;

  const maxScroll = nav.scrollWidth - nav.clientWidth;
  if (maxScroll <= 2) {
    leftBtn.classList.add('opacity-30', 'cursor-not-allowed');
    rightBtn.classList.add('opacity-30', 'cursor-not-allowed');
    leftBtn.disabled = true;
    rightBtn.disabled = true;
    return;
  }

  if (nav.scrollLeft > 6) {
    leftBtn.classList.remove('opacity-30', 'cursor-not-allowed');
    leftBtn.disabled = false;
  } else {
    leftBtn.classList.add('opacity-30', 'cursor-not-allowed');
    leftBtn.disabled = true;
  }

  if (nav.scrollLeft < maxScroll - 6) {
    rightBtn.classList.remove('opacity-30', 'cursor-not-allowed');
    rightBtn.disabled = false;
  } else {
    rightBtn.classList.add('opacity-30', 'cursor-not-allowed');
    rightBtn.disabled = true;
  }
}

function showToast(msg, isError = false) {
  const toast = document.getElementById('toastNotification');
  const msgSpan = document.getElementById('toastMessage');
  const icon = document.getElementById('toastIcon');

  msgSpan.textContent = msg;
  if (isError) {
    icon.setAttribute('data-lucide', 'alert-circle');
    icon.className = 'w-4 h-4 text-rose-400';
  } else {
    icon.setAttribute('data-lucide', 'check-circle');
    icon.className = 'w-4 h-4 text-emerald-400';
  }

  toast.classList.remove('opacity-0', 'translate-y-20', 'pointer-events-none');
  lucide.createIcons();

  setTimeout(() => {
    toast.classList.add('opacity-0', 'translate-y-20', 'pointer-events-none');
  }, 3000);
}

// -------------------------------------------------------------
// 🌟 1. Real-Time Animated KPI Cards
// -------------------------------------------------------------
function updateKpiCards() {
  const totalRowsEl = document.getElementById('kpiTotalRows');
  const totalColsEl = document.getElementById('kpiTotalCols');
  const m1Label = document.getElementById('kpiMetric1Label');
  const m1Val = document.getElementById('kpiMetric1Val');
  const m2Label = document.getElementById('kpiMetric2Label');
  const m2Val = document.getElementById('kpiMetric2Val');

  if (!totalRowsEl) return;

  totalRowsEl.textContent = AppState.currentRows.length.toLocaleString();
  totalColsEl.textContent = AppState.columns.length.toLocaleString();

  // Pick the column people most likely care about (sales, revenue, profit...), not just the first number
  const numCols = AppState.columns.filter(c => AppState.colMeta[c]?.type === 'numeric');
  const priority = ['sales', 'revenue', 'amount', 'total', 'profit', 'salary', 'spend', 'price', 'cost', 'score', 'marks', 'orders', 'units', 'qty', 'quantity'];
  const rank = c => { const cl = c.toLowerCase(); const i = priority.findIndex(k => cl.includes(k)); return i === -1 ? priority.length : i; };
  const ranked = [...numCols].sort((a, b) => rank(a) - rank(b) || numCols.indexOf(a) - numCols.indexOf(b));

  if (ranked.length > 0) {
    const col1 = ranked[0];
    const meta1 = AppState.colMeta[col1];
    m1Label.textContent = `Total ${col1}`;
    m1Val.textContent = formatNumber(meta1?.sum ?? 0, meta1 && Math.abs(meta1.sum) >= 1000 ? 0 : 2);

    const col2 = ranked[1] || col1;
    const meta2 = AppState.colMeta[col2];
    m2Label.textContent = `Avg ${col2}`;
    m2Val.textContent = formatNumber(meta2?.avg ?? null, 1);
  } else {
    m1Label.textContent = 'Categories';
    m1Val.textContent = AppState.columns.length;
    m2Label.textContent = 'Active Sheet';
    m2Val.textContent = AppState.activeSheet;
  }
}

// -------------------------------------------------------------
// 🤖 2. Google Gemini AI Smart Executive Summary
// -------------------------------------------------------------
async function requestGeminiInsight() {
  const btn = document.getElementById('btnGenerateAiInsight');
  const contentEl = document.getElementById('aiInsightContent');
  const badgeEl = document.getElementById('aiInsightBadge');
  const isGuest = !AppState.auth.token;

  btn.disabled = true;
  btn.innerHTML = `<i data-lucide="loader-2" class="w-4 h-4 animate-spin"></i><span>Analyzing…</span>`;
  lucide.createIcons();

  contentEl.innerHTML = `<div class="flex items-center gap-2 text-slate-500 py-3"><i data-lucide="sparkles" class="w-4 h-4 animate-pulse text-indigo-600"></i> Writing the executive summary…</div>`;
  lucide.createIcons();

  try {
    const headers = { 'Content-Type': 'application/json' };
    if (AppState.auth.token) {
      headers['Authorization'] = `Bearer ${AppState.auth.token}`;
    }

    const res = await fetch(`${AppState.apiBaseUrl}/ai/insights`, {
      method: 'POST',
      headers,
      body: JSON.stringify({
        fileName: AppState.fileName,
        totalRows: AppState.currentRows.length,
        totalColumns: AppState.columns.length,
        columnStats: AppState.colMeta,
        sampleRows: AppState.currentRows.slice(0, 10),
        aggregates: buildDataAggregates()
      })
    });

    const data = await res.json();
    if (!res.ok || !data.success) {
      throw new Error(data.message || 'AI request failed');
    }

    const insightText = extractAiAnswer(data.insight || '');
    AppState.lastAiInsight = insightText;
    badgeEl.textContent = `${data.source === 'statistical-engine' ? 'Offline statistical summary' : `Model: ${data.source}`}${isGuest ? ' · Guest' : ''}`;
    contentEl.innerHTML = `<div class="ai-md">${renderMarkdownLite(insightText)}</div>`;
    
    if (isGuest) {
      showToast('AI insights generated! (Sign in to save this analysis to your account history)');
    } else {
      showToast('Gemini AI executive insights generated successfully!');
    }
  } catch (err) {
    console.error(err);
    contentEl.innerHTML = `<div class="text-rose-600 text-xs py-2">Could not reach the AI service: ${escapeHtml(err.message)}. The server may be waking up; try again in a few seconds.</div>`;
    showToast(err.message, true);
  } finally {
    btn.disabled = false;
    btn.innerHTML = `<i data-lucide="wand-2" class="w-4 h-4"></i><span>Generate Insights</span>`;
    lucide.createIcons();
  }
}

function copyAiInsight() {
  const text = AppState.lastAiInsight || document.getElementById('aiInsightContent').innerText;
  navigator.clipboard.writeText(text);
  showToast('Copied AI insights to clipboard');
}

// -------------------------------------------------------------
// 🌟 3. Enterprise Three.js 3D Visualizer Engine
// -------------------------------------------------------------
function cleanupThreeJs() {
  if (AppState.threeJs.animationId) {
    cancelAnimationFrame(AppState.threeJs.animationId);
    AppState.threeJs.animationId = null;
  }
  if (AppState.threeJs.controls) {
    if (typeof AppState.threeJs.controls.dispose === 'function') {
      AppState.threeJs.controls.dispose();
    }
    AppState.threeJs.controls = null;
  }
  if (AppState.threeJs.scene) {
    AppState.threeJs.scene.traverse((obj) => {
      if (obj.geometry) obj.geometry.dispose();
      if (obj.material) {
        if (Array.isArray(obj.material)) obj.material.forEach(m => m.dispose());
        else obj.material.dispose();
      }
    });
  }
  const container = document.getElementById('threeJsCanvasContainer');
  if (container) {
    const canvases = container.querySelectorAll('canvas');
    canvases.forEach(c => c.remove());
  }
  const tooltip = document.getElementById('threeDTooltip');
  if (tooltip) tooltip.classList.add('hidden');
  AppState.threeJs.meshes = [];
  AppState.threeJs.activeHoverMesh = null;
  AppState.threeJs.scene = null;
  AppState.threeJs.camera = null;
  AppState.threeJs.renderer = null;
}

function resizeThreeJs() {
  const container = document.getElementById('threeJsCanvasContainer');
  if (!container || !AppState.threeJs.renderer || !AppState.threeJs.camera) return;
  const width = container.clientWidth || 700;
  const height = container.clientHeight || 400;
  AppState.threeJs.camera.aspect = width / height;
  AppState.threeJs.camera.updateProjectionMatrix();
  AppState.threeJs.renderer.setSize(width, height);
}

function renderThreeJsScene(type, xCol, yCol, agg) {
  cleanupThreeJs();
  AppState.threeJs.currentType = type;

  const container = document.getElementById('threeJsCanvasContainer');
  if (!container) return;

  const width = container.clientWidth || 700;
  const height = container.clientHeight || 400;

  // Aggregate Data
  const grouped = {};
  let totalSum = 0;
  AppState.currentRows.forEach(row => {
    const k = String(row[xCol] || '(Blank)').trim();
    const v = Number(row[yCol]) || 0;
    if (!grouped[k]) grouped[k] = { sum: 0, count: 0, min: Infinity, max: -Infinity };
    grouped[k].sum += v;
    grouped[k].count++;
    if (v < grouped[k].min) grouped[k].min = v;
    if (v > grouped[k].max) grouped[k].max = v;
    totalSum += v;
  });

  const categories = Object.keys(grouped).slice(0, 20);
  if (categories.length === 0) return;

  let maxVal = 1;
  let minVal = Infinity;
  const parsedItems = categories.map(k => {
    const item = grouped[k];
    let val = item.sum;
    if (agg === 'avg') val = item.sum / (item.count || 1);
    else if (agg === 'count') val = item.count;
    else if (agg === 'max') val = item.max === -Infinity ? 0 : item.max;
    else if (agg === 'min') val = item.min === Infinity ? 0 : item.min;

    if (val > maxVal) maxVal = val;
    if (val < minVal) minVal = val;
    return { category: k, val: val, count: item.count, sum: item.sum };
  });

  // Color palettes & theme settings
  const themes = {
    cyber: {
      bg: 0x090d16,
      gridPrimary: 0x4f46e5,
      gridSecondary: 0x1e293b,
      lightColor: 0x818cf8,
      rimColor: 0x06b6d4,
      colors: [0x6366f1, 0x06b6d4, 0x10b981, 0xf59e0b, 0xec4899, 0x8b5cf6, 0x3b82f6, 0x14b8a6]
    },
    sunset: {
      bg: 0x12071f,
      gridPrimary: 0xd946ef,
      gridSecondary: 0x2e1065,
      lightColor: 0xf43f5e,
      rimColor: 0xfbbf24,
      colors: [0xf43f5e, 0xf97316, 0xfbbf24, 0xd946ef, 0xa855f7, 0xe11d48, 0xf472b6, 0xfb923c]
    },
    emerald: {
      bg: 0x031713,
      gridPrimary: 0x10b981,
      gridSecondary: 0x064e3b,
      lightColor: 0x34d399,
      rimColor: 0x06b6d4,
      colors: [0x10b981, 0x059669, 0x34d399, 0x06b6d4, 0x14b8a6, 0x6ee7b7, 0x2dd4bf, 0x0284c7]
    },
    minimal: {
      bg: 0x0f172a,
      gridPrimary: 0x64748b,
      gridSecondary: 0x1e293b,
      lightColor: 0xe2e8f0,
      rimColor: 0x38bdf8,
      colors: [0x38bdf8, 0x60a5fa, 0x818cf8, 0xa78bfa, 0xc084fc, 0x94a3b8, 0x475569, 0x93c5fd]
    }
  };

  const activeTheme = themes[AppState.threeJs.theme] || themes.cyber;

  // Scene
  const scene = new THREE.Scene();
  scene.background = new THREE.Color(activeTheme.bg);

  // Camera
  const camera = new THREE.PerspectiveCamera(45, width / height, 0.1, 1000);
  camera.position.set(16, 18, 24);
  camera.lookAt(0, 3, 0);

  // Renderer with preserveDrawingBuffer enabled for high-res PNG export
  const renderer = new THREE.WebGLRenderer({ antialias: true, alpha: true, preserveDrawingBuffer: true });
  renderer.setSize(width, height);
  renderer.setPixelRatio(Math.min(window.devicePixelRatio || 1, 2));
  renderer.shadowMap.enabled = true;
  renderer.shadowMap.type = THREE.PCFSoftShadowMap;
  container.appendChild(renderer.domElement);

  // Lights
  const ambientLight = new THREE.AmbientLight(0xffffff, 0.75);
  scene.add(ambientLight);

  const mainLight = new THREE.DirectionalLight(activeTheme.lightColor, 1.3);
  mainLight.position.set(15, 25, 20);
  mainLight.castShadow = true;
  mainLight.shadow.mapSize.width = 1024;
  mainLight.shadow.mapSize.height = 1024;
  scene.add(mainLight);

  const rimLight = new THREE.PointLight(activeTheme.rimColor, 1.2, 80);
  rimLight.position.set(-18, 14, -14);
  scene.add(rimLight);

  // Grid Floor
  const gridHelper = new THREE.GridHelper(30, 30, activeTheme.gridPrimary, activeTheme.gridSecondary);
  gridHelper.position.y = 0;
  gridHelper.visible = AppState.threeJs.showGrid;
  scene.add(gridHelper);
  AppState.threeJs.gridHelper = gridHelper;

  // Meshes collection for raycasting
  const interactiveMeshes = [];
  const palette = activeTheme.colors;

  // -----------------------------------------------------------
  // Build Visuals according to Type
  // -----------------------------------------------------------
  if (type === '3d-column') {
    // 🌟 3D Column Grid with pedestals and glowing caps
    const spacing = 2.0;
    const startX = -((parsedItems.length - 1) * spacing) / 2;

    parsedItems.forEach((item, idx) => {
      const heightVal = Math.max(0.5, (item.val / (maxVal || 1)) * 9);
      const color = palette[idx % palette.length];

      // Base Pedestal
      const baseGeo = new THREE.CylinderGeometry(0.85, 0.95, 0.25, 16);
      const baseMat = new THREE.MeshStandardMaterial({
        color: 0x1e293b,
        metalness: 0.8,
        roughness: 0.2
      });
      const baseMesh = new THREE.Mesh(baseGeo, baseMat);
      baseMesh.position.set(startX + idx * spacing, 0.125, 0);
      baseMesh.receiveShadow = true;
      scene.add(baseMesh);

      // Main Column
      const colGeo = new THREE.BoxGeometry(1.2, heightVal, 1.2);
      const colMat = new THREE.MeshPhongMaterial({
        color: color,
        shininess: 90,
        specular: 0xffffff,
        emissive: 0x000000
      });
      const colMesh = new THREE.Mesh(colGeo, colMat);
      colMesh.position.set(startX + idx * spacing, heightVal / 2 + 0.25, 0);
      colMesh.castShadow = true;
      colMesh.receiveShadow = true;

      // Glow Cap
      const capGeo = new THREE.BoxGeometry(1.22, 0.15, 1.22);
      const capMat = new THREE.MeshBasicMaterial({ color: 0xffffff });
      const capMesh = new THREE.Mesh(capGeo, capMat);
      capMesh.position.set(0, heightVal / 2 + 0.08, 0);
      colMesh.add(capMesh);

      // Data association
      const percent = totalSum > 0 ? ((item.val / totalSum) * 100).toFixed(1) : '0';
      colMesh.userData = {
        label: item.category,
        val: item.val,
        percent: percent,
        rank: idx + 1,
        total: parsedItems.length,
        originalColor: color
      };

      scene.add(colMesh);
      interactiveMeshes.push(colMesh);
    });

  } else if (type === '3d-scatter') {
    // 🌐 3D Bubble & Scatter Universe
    const count = parsedItems.length;
    parsedItems.forEach((item, idx) => {
      const radius = 0.5 + (item.val / (maxVal || 1)) * 0.9;
      const heightY = 1.0 + (item.val / (maxVal || 1)) * 8;
      const posX = ((idx - count / 2) / (count / 2 || 1)) * 11;
      const posZ = Math.sin(idx * 0.8) * 6;
      const color = palette[idx % palette.length];

      // Stem Line
      const points = [new THREE.Vector3(posX, 0, posZ), new THREE.Vector3(posX, heightY, posZ)];
      const lineGeo = new THREE.BufferGeometry().setFromPoints(points);
      const lineMat = new THREE.LineBasicMaterial({ color: activeTheme.gridPrimary, transparent: true, opacity: 0.5 });
      const line = new THREE.Line(lineGeo, lineMat);
      scene.add(line);

      // Sphere
      const sphereGeo = new THREE.SphereGeometry(radius, 32, 32);
      const sphereMat = new THREE.MeshPhysicalMaterial({
        color: color,
        metalness: 0.1,
        roughness: 0.15,
        clearcoat: 0.8,
        clearcoatRoughness: 0.1,
        emissive: 0x000000
      });
      const sphereMesh = new THREE.Mesh(sphereGeo, sphereMat);
      sphereMesh.position.set(posX, heightY, posZ);
      sphereMesh.castShadow = true;

      // Glow halo ring
      const ringGeo = new THREE.RingGeometry(radius + 0.15, radius + 0.25, 32);
      const ringMat = new THREE.MeshBasicMaterial({ color: color, side: THREE.DoubleSide, transparent: true, opacity: 0.45 });
      const ringMesh = new THREE.Mesh(ringGeo, ringMat);
      ringMesh.rotation.x = Math.PI / 2;
      sphereMesh.add(ringMesh);

      const percent = totalSum > 0 ? ((item.val / totalSum) * 100).toFixed(1) : '0';
      sphereMesh.userData = {
        label: item.category,
        val: item.val,
        percent: percent,
        rank: idx + 1,
        total: parsedItems.length,
        originalColor: color
      };

      scene.add(sphereMesh);
      interactiveMeshes.push(sphereMesh);
    });

  } else if (type === '3d-cylinder') {
    // 🍩 3D Radial Cylinder Donut Ring
    const count = parsedItems.length;
    const ringRadius = 7.5;
    const angleStep = (Math.PI * 2) / Math.max(count, 1);

    parsedItems.forEach((item, idx) => {
      const angle = idx * angleStep;
      const heightVal = Math.max(0.6, (item.val / (maxVal || 1)) * 7);
      const posX = Math.cos(angle) * ringRadius;
      const posZ = Math.sin(angle) * ringRadius;
      const color = palette[idx % palette.length];

      const cylGeo = new THREE.CylinderGeometry(0.8, 0.8, heightVal, 24);
      const cylMat = new THREE.MeshPhongMaterial({
        color: color,
        shininess: 95,
        specular: 0xffffff,
        emissive: 0x000000
      });
      const cylMesh = new THREE.Mesh(cylGeo, cylMat);
      cylMesh.position.set(posX, heightVal / 2, posZ);
      cylMesh.castShadow = true;
      cylMesh.receiveShadow = true;

      // Top glowing indicator ring
      const topRingGeo = new THREE.CylinderGeometry(0.82, 0.82, 0.1, 24);
      const topRingMat = new THREE.MeshBasicMaterial({ color: 0xffffff });
      const topRing = new THREE.Mesh(topRingGeo, topRingMat);
      topRing.position.set(0, heightVal / 2, 0);
      cylMesh.add(topRing);

      const percent = totalSum > 0 ? ((item.val / totalSum) * 100).toFixed(1) : '0';
      cylMesh.userData = {
        label: item.category,
        val: item.val,
        percent: percent,
        rank: idx + 1,
        total: parsedItems.length,
        originalColor: color,
        radialAngle: angle,
        baseX: posX,
        baseZ: posZ
      };

      scene.add(cylMesh);
      interactiveMeshes.push(cylMesh);
    });

  } else if (type === '3d-surface') {
    // 🏔️ 3D Surface Elevation Mesh
    const count = parsedItems.length;
    const segments = count - 1;
    const surfWidth = 18;
    const surfDepth = 8;
    const planeGeo = new THREE.PlaneGeometry(surfWidth, surfDepth, Math.max(segments, 1), 8);
    planeGeo.rotateX(-Math.PI / 2);

    const posAttr = planeGeo.attributes.position;
    for (let i = 0; i < posAttr.count; i++) {
      const x = posAttr.getX(i);
      const normX = (x + surfWidth / 2) / surfWidth;
      const segIdx = Math.min(Math.floor(normX * count), count - 1);
      const itemVal = parsedItems[segIdx]?.val || 0;
      const targetH = (itemVal / (maxVal || 1)) * 6.5;
      const z = posAttr.getZ(i);
      const zFactor = 1 - Math.abs(z) / (surfDepth / 2);
      posAttr.setY(i, Math.max(0.2, targetH * Math.max(0.3, zFactor)));
    }
    planeGeo.computeVertexNormals();

    const surfaceMat = new THREE.MeshStandardMaterial({
      color: activeTheme.colors[0],
      metalness: 0.2,
      roughness: 0.35,
      side: THREE.DoubleSide
    });
    const surfaceMesh = new THREE.Mesh(planeGeo, surfaceMat);
    surfaceMesh.receiveShadow = true;
    surfaceMesh.castShadow = true;
    scene.add(surfaceMesh);

    // Wireframe overlay for cyber elevation aesthetic
    const wireMat = new THREE.MeshBasicMaterial({ color: activeTheme.rimColor, wireframe: true, transparent: true, opacity: 0.35 });
    const wireMesh = new THREE.Mesh(planeGeo, wireMat);
    scene.add(wireMesh);

    // Beacons on each category peak
    parsedItems.forEach((item, idx) => {
      const posX = -surfWidth / 2 + (idx / Math.max(count - 1, 1)) * surfWidth;
      const peakH = Math.max(0.4, (item.val / (maxVal || 1)) * 6.5);
      const pinGeo = new THREE.SphereGeometry(0.35, 16, 16);
      const pinMat = new THREE.MeshBasicMaterial({ color: palette[idx % palette.length] });
      const pinMesh = new THREE.Mesh(pinGeo, pinMat);
      pinMesh.position.set(posX, peakH + 0.3, 0);

      const percent = totalSum > 0 ? ((item.val / totalSum) * 100).toFixed(1) : '0';
      pinMesh.userData = {
        label: item.category,
        val: item.val,
        percent: percent,
        rank: idx + 1,
        total: parsedItems.length,
        originalColor: palette[idx % palette.length]
      };
      scene.add(pinMesh);
      interactiveMeshes.push(pinMesh);
    });

  } else if (type === '3d-matrix') {
    // 📊 3D Stepped Grid Waterfall
    const colsCount = Math.min(parsedItems.length, 12);
    const spacingX = 2.0;
    const spacingZ = 2.4;
    const startX = -((colsCount - 1) * spacingX) / 2;

    for (let r = 0; r < 2; r++) {
      for (let c = 0; c < colsCount; c++) {
        const item = parsedItems[c];
        if (!item) continue;
        const multiplier = r === 0 ? 1 : 0.65;
        const calcVal = item.val * multiplier;
        const heightVal = Math.max(0.4, (calcVal / (maxVal || 1)) * 7.5);
        const color = r === 0 ? palette[c % palette.length] : 0x475569;

        const boxGeo = new THREE.BoxGeometry(1.2, heightVal, 1.2);
        const boxMat = new THREE.MeshPhongMaterial({
          color: color,
          shininess: 85,
          specular: 0xffffff,
          emissive: 0x000000
        });
        const boxMesh = new THREE.Mesh(boxGeo, boxMat);
        boxMesh.position.set(startX + c * spacingX, heightVal / 2, (r === 0 ? 1 : -1) * spacingZ);
        boxMesh.castShadow = true;
        boxMesh.receiveShadow = true;

        const percent = totalSum > 0 ? ((calcVal / totalSum) * 100).toFixed(1) : '0';
        boxMesh.userData = {
          label: `${item.category} (${r === 0 ? 'Primary' : 'Benchmark'})`,
          val: calcVal,
          percent: percent,
          rank: c + 1,
          total: colsCount,
          originalColor: color
        };

        scene.add(boxMesh);
        interactiveMeshes.push(boxMesh);
      }
    }

  } else if (type === '3d-pyramid') {
    // 🔺 3D Pyramid & Funnel
    const count = parsedItems.length;
    const spacing = 2.2;
    const startX = -((count - 1) * spacing) / 2;

    parsedItems.forEach((item, idx) => {
      const heightVal = Math.max(0.6, (item.val / (maxVal || 1)) * 8.5);
      const color = palette[idx % palette.length];

      const pyrGeo = new THREE.CylinderGeometry(0.05, 0.9, heightVal, 4);
      const pyrMat = new THREE.MeshPhongMaterial({
        color: color,
        shininess: 90,
        specular: 0xffffff,
        emissive: 0x000000
      });
      const pyrMesh = new THREE.Mesh(pyrGeo, pyrMat);
      pyrMesh.position.set(startX + idx * spacing, heightVal / 2, 0);
      pyrMesh.rotation.y = Math.PI / 4;
      pyrMesh.castShadow = true;
      pyrMesh.receiveShadow = true;

      const apexGeo = new THREE.SphereGeometry(0.18, 12, 12);
      const apexMat = new THREE.MeshBasicMaterial({ color: 0xffffff });
      const apexMesh = new THREE.Mesh(apexGeo, apexMat);
      apexMesh.position.set(0, heightVal / 2 + 0.1, 0);
      pyrMesh.add(apexMesh);

      const percent = totalSum > 0 ? ((item.val / totalSum) * 100).toFixed(1) : '0';
      pyrMesh.userData = {
        label: item.category,
        val: item.val,
        percent: percent,
        rank: idx + 1,
        total: parsedItems.length,
        originalColor: color
      };

      scene.add(pyrMesh);
      interactiveMeshes.push(pyrMesh);
    });

  } else if (type === '3d-pie') {
    // 🛸 3D Floating Donut Ring / Pie Wedges
    const count = parsedItems.length;
    const ringRadius = 6.5;
    const angleStep = (Math.PI * 2) / Math.max(count, 1);

    parsedItems.forEach((item, idx) => {
      const angle = idx * angleStep;
      const heightVal = Math.max(0.4, (item.val / (maxVal || 1)) * 6.0);
      const posX = Math.cos(angle) * ringRadius;
      const posZ = Math.sin(angle) * ringRadius;
      const color = palette[idx % palette.length];

      const wedgeGeo = new THREE.CylinderGeometry(0.85, 0.85, heightVal, 16);
      const wedgeMat = new THREE.MeshStandardMaterial({
        color: color,
        metalness: 0.3,
        roughness: 0.25,
        emissive: 0x000000
      });
      const wedgeMesh = new THREE.Mesh(wedgeGeo, wedgeMat);
      wedgeMesh.position.set(posX, heightVal / 2 + 0.5, posZ);
      wedgeMesh.castShadow = true;
      wedgeMesh.receiveShadow = true;

      const percent = totalSum > 0 ? ((item.val / totalSum) * 100).toFixed(1) : '0';
      wedgeMesh.userData = {
        label: item.category,
        val: item.val,
        percent: percent,
        rank: idx + 1,
        total: parsedItems.length,
        originalColor: color,
        baseY: heightVal / 2 + 0.5
      };

      scene.add(wedgeMesh);
      interactiveMeshes.push(wedgeMesh);
    });
  }

  AppState.threeJs.meshes = interactiveMeshes;

  // -----------------------------------------------------------
  // Orbit Controls & Interactive Fallback
  // -----------------------------------------------------------
  let controls = null;
  if (typeof THREE.OrbitControls === 'function') {
    controls = new THREE.OrbitControls(camera, renderer.domElement);
    controls.enableDamping = true;
    controls.dampingFactor = 0.05;
    controls.maxPolarAngle = Math.PI / 2 - 0.04;
    controls.minDistance = 6;
    controls.maxDistance = 75;
    controls.target.set(0, 3, 0);
    AppState.threeJs.controls = controls;
  } else {
    // Custom fallback pointer drag/zoom orbit
    let isDragging = false;
    let prevMouseX = 0;
    let prevMouseY = 0;
    let theta = Math.PI / 4;
    let phi = Math.PI / 3;
    let radius = 28;

    const updateCameraPos = () => {
      camera.position.x = radius * Math.sin(phi) * Math.sin(theta);
      camera.position.y = radius * Math.cos(phi);
      camera.position.z = radius * Math.sin(phi) * Math.cos(theta);
      camera.lookAt(0, 3, 0);
    };

    renderer.domElement.addEventListener('pointerdown', (e) => {
      isDragging = true;
      prevMouseX = e.clientX;
      prevMouseY = e.clientY;
    });

    window.addEventListener('pointerup', () => { isDragging = false; });

    renderer.domElement.addEventListener('pointermove', (e) => {
      if (!isDragging) return;
      const dx = e.clientX - prevMouseX;
      const dy = e.clientY - prevMouseY;
      prevMouseX = e.clientX;
      prevMouseY = e.clientY;

      theta -= dx * 0.008;
      phi = Math.max(0.1, Math.min(Math.PI / 2 - 0.05, phi - dy * 0.008));
      updateCameraPos();
    });

    renderer.domElement.addEventListener('wheel', (e) => {
      e.preventDefault();
      radius = Math.max(8, Math.min(65, radius + e.deltaY * 0.02));
      updateCameraPos();
    }, { passive: false });
  }

  // -----------------------------------------------------------
  // Raycaster & Mouse Hover Data HUD
  // -----------------------------------------------------------
  const raycaster = new THREE.Raycaster();
  const mouse = new THREE.Vector2();
  const tooltip = document.getElementById('threeDTooltip');
  const tipTitle = document.getElementById('threeDTooltipTitle');
  const tipVal = document.getElementById('threeDTooltipValue');
  const tipShare = document.getElementById('threeDTooltipShare');

  function onPointerMove(e) {
    const rect = renderer.domElement.getBoundingClientRect();
    mouse.x = ((e.clientX - rect.left) / rect.width) * 2 - 1;
    mouse.y = -((e.clientY - rect.top) / rect.height) * 2 + 1;

    raycaster.setFromCamera(mouse, camera);
    const intersects = raycaster.intersectObjects(AppState.threeJs.meshes);

    if (intersects.length > 0) {
      const hit = intersects[0].object;
      if (AppState.threeJs.activeHoverMesh !== hit) {
        // Reset previous
        if (AppState.threeJs.activeHoverMesh && AppState.threeJs.activeHoverMesh.material) {
          if (AppState.threeJs.activeHoverMesh.material.emissive) {
            AppState.threeJs.activeHoverMesh.material.emissive.setHex(0x000000);
          }
        }
        AppState.threeJs.activeHoverMesh = hit;

        // Highlight hit mesh
        if (hit.material && hit.material.emissive) {
          hit.material.emissive.setHex(0x38bdf8);
        }
      }

      // Update Tooltip
      if (hit.userData && tooltip) {
        tipTitle.textContent = hit.userData.label || 'Item';
        tipVal.textContent = Number(hit.userData.val || 0).toLocaleString(undefined, { maximumFractionDigits: 1 });
        tipShare.textContent = `${hit.userData.percent}% of total • Rank #${hit.userData.rank} of ${hit.userData.total}`;

        const tipX = Math.min(rect.width - 160, Math.max(10, e.clientX - rect.left + 14));
        const tipY = Math.max(10, e.clientY - rect.top - 50);
        tooltip.style.left = `${tipX}px`;
        tooltip.style.top = `${tipY}px`;
        tooltip.classList.remove('hidden');
      }
      container.style.cursor = 'pointer';
    } else {
      if (AppState.threeJs.activeHoverMesh) {
        if (AppState.threeJs.activeHoverMesh.material && AppState.threeJs.activeHoverMesh.material.emissive) {
          AppState.threeJs.activeHoverMesh.material.emissive.setHex(0x000000);
        }
        AppState.threeJs.activeHoverMesh = null;
      }
      if (tooltip) tooltip.classList.add('hidden');
      container.style.cursor = 'default';
    }
  }

  function onPointerLeave() {
    if (AppState.threeJs.activeHoverMesh) {
      if (AppState.threeJs.activeHoverMesh.material && AppState.threeJs.activeHoverMesh.material.emissive) {
        AppState.threeJs.activeHoverMesh.material.emissive.setHex(0x000000);
      }
      AppState.threeJs.activeHoverMesh = null;
    }
    if (tooltip) tooltip.classList.add('hidden');
  }

  renderer.domElement.addEventListener('pointermove', onPointerMove);
  renderer.domElement.addEventListener('pointerleave', onPointerLeave);

  // -----------------------------------------------------------
  // Smooth Animation Loop
  // -----------------------------------------------------------
  let rotationAngle = 0;
  function animate() {
    AppState.threeJs.animationId = requestAnimationFrame(animate);

    if (AppState.threeJs.autoRotate) {
      if (controls) {
        controls.autoRotate = true;
        controls.autoRotateSpeed = 1.6;
      } else {
        rotationAngle += 0.005;
        const currentDist = camera.position.distanceTo(new THREE.Vector3(0, 3, 0)) || 28;
        camera.position.x = Math.sin(rotationAngle) * currentDist * 0.8;
        camera.position.z = Math.cos(rotationAngle) * currentDist * 0.8;
        camera.lookAt(0, 3, 0);
      }
    } else {
      if (controls) controls.autoRotate = false;
    }

    if (controls) controls.update();
    renderer.render(scene, camera);
  }

  animate();

  AppState.threeJs.renderer = renderer;
  AppState.threeJs.scene = scene;
  AppState.threeJs.camera = camera;
}

// -------------------------------------------------------------
// 3D Visualizer Control Helpers & Presets
// -------------------------------------------------------------
function toggle3DAutoRotate() {
  AppState.threeJs.autoRotate = !AppState.threeJs.autoRotate;
  const label = document.getElementById('threeDAutoRotateLabel');
  const hudText = document.getElementById('hudRotateText');
  const status = AppState.threeJs.autoRotate ? 'Active' : 'Paused';
  if (label) label.textContent = status;
  if (hudText) hudText.textContent = AppState.threeJs.autoRotate ? 'Pause' : 'Play';
  showToast(`3D Auto-Rotation: ${status}`);
}

function set3DCameraView(mode) {
  const camera = AppState.threeJs.camera;
  if (!camera) return;

  if (mode === 'isometric') {
    camera.position.set(16, 18, 24);
    camera.lookAt(0, 3, 0);
  } else if (mode === 'top') {
    camera.position.set(0, 34, 0.01);
    camera.lookAt(0, 0, 0);
  } else if (mode === 'front') {
    camera.position.set(0, 8, 30);
    camera.lookAt(0, 3, 0);
  }

  if (AppState.threeJs.controls) {
    AppState.threeJs.controls.target.set(0, 3, 0);
    AppState.threeJs.controls.update();
  }
  showToast(`Switched to ${mode.toUpperCase()} 3D view`);
}

function reset3DCamera() {
  set3DCameraView('isometric');
}

function toggle3DFloorGrid() {
  AppState.threeJs.showGrid = !AppState.threeJs.showGrid;
  if (AppState.threeJs.gridHelper) {
    AppState.threeJs.gridHelper.visible = AppState.threeJs.showGrid;
  }
  const label = document.getElementById('threeDGridToggleLabel');
  if (label) label.textContent = AppState.threeJs.showGrid ? 'Enabled' : 'Hidden';
  showToast(`3D Floor Grid: ${AppState.threeJs.showGrid ? 'Enabled' : 'Hidden'}`);
}

function update3DTheme(themeName) {
  AppState.threeJs.theme = themeName;
  const xCol = document.getElementById('studioXCol')?.value;
  const yCol = document.getElementById('studioYCol')?.value;
  const type = document.getElementById('studioChartType')?.value || '3d-column';
  const agg = document.getElementById('studioAgg')?.value || 'sum';
  renderThreeJsScene(type, xCol, yCol, agg);
  showToast(`Applied 3D theme: ${themeName}`);
}

function toggle3DFullscreen() {
  const container = document.getElementById('threeJsCanvasContainer');
  if (!container) return;
  AppState.threeJs.isFullscreen = !AppState.threeJs.isFullscreen;
  container.classList.toggle('fullscreen-3d-active', AppState.threeJs.isFullscreen);

  setTimeout(() => {
    resizeThreeJs();
  }, 100);

  showToast(AppState.threeJs.isFullscreen ? 'Entered 3D fullscreen mode (Press Esc or click again to exit)' : 'Exited 3D fullscreen mode');
}

// Escape key listener to exit fullscreen
window.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && AppState.threeJs.isFullscreen) {
    toggle3DFullscreen();
  }
});

// =============================================================
// 🌟 Dedicated 3D Spatial Analytics & WebGL Studio Engine
// =============================================================

function cleanupDedicated3D() {
  if (AppState.threeJsDedicated.animationId) {
    cancelAnimationFrame(AppState.threeJsDedicated.animationId);
    AppState.threeJsDedicated.animationId = null;
  }
  if (AppState.threeJsDedicated.controls) {
    if (typeof AppState.threeJsDedicated.controls.dispose === 'function') {
      AppState.threeJsDedicated.controls.dispose();
    }
    AppState.threeJsDedicated.controls = null;
  }
  if (AppState.threeJsDedicated.scene) {
    AppState.threeJsDedicated.scene.traverse((obj) => {
      if (obj.geometry) obj.geometry.dispose();
      if (obj.material) {
        if (Array.isArray(obj.material)) obj.material.forEach(m => m.dispose());
        else obj.material.dispose();
      }
    });
  }
  const container = document.getElementById('dedicatedThreeJsCanvasContainer');
  if (container) {
    const canvases = container.querySelectorAll('canvas');
    canvases.forEach(c => c.remove());
  }
  const tooltip = document.getElementById('view3dTooltip');
  if (tooltip) tooltip.classList.add('hidden');
  AppState.threeJsDedicated.meshes = [];
  AppState.threeJsDedicated.activeHoverMesh = null;
  AppState.threeJsDedicated.scene = null;
  AppState.threeJsDedicated.camera = null;
  AppState.threeJsDedicated.renderer = null;
}

function resizeDedicated3D() {
  const container = document.getElementById('dedicatedThreeJsCanvasContainer');
  if (!container || !AppState.threeJsDedicated.renderer || !AppState.threeJsDedicated.camera) return;
  const width = container.clientWidth || 800;
  const height = container.clientHeight || 520;
  AppState.threeJsDedicated.camera.aspect = width / height;
  AppState.threeJsDedicated.camera.updateProjectionMatrix();
  AppState.threeJsDedicated.renderer.setSize(width, height);
}

function openInDedicated3DView() {
  const xCol = document.getElementById('studioXCol')?.value;
  const yCol = document.getElementById('studioYCol')?.value;
  const agg = document.getElementById('studioAgg')?.value || 'sum';
  const type = document.getElementById('studioChartType')?.value || '3d-column';

  switchTab('3dchart');

  if (xCol) {
    const viewX = document.getElementById('view3dXCol');
    if (viewX) viewX.value = xCol;
  }
  if (yCol) {
    const viewY = document.getElementById('view3dYCol');
    if (viewY) viewY.value = yCol;
  }
  if (agg) {
    const viewAgg = document.getElementById('view3dAgg');
    if (viewAgg) viewAgg.value = agg;
  }
  if (type && type.startsWith('3d-')) {
    const viewType = document.getElementById('view3dChartType');
    if (viewType) viewType.value = type;
  }

  renderDedicated3DView();
  showToast('Switched to Dedicated 3D Spatial Analytics View');
}

function renderDedicated3DView() {
  cleanupDedicated3D();

  const container = document.getElementById('dedicatedThreeJsCanvasContainer');
  if (!container) return;

  const xCol = document.getElementById('view3dXCol')?.value || AppState.columns[0];
  const yCol = document.getElementById('view3dYCol')?.value || AppState.columns[1] || AppState.columns[0];
  const type = document.getElementById('view3dChartType')?.value || '3d-column';
  const agg = document.getElementById('view3dAgg')?.value || 'sum';

  if (!xCol || !yCol) return;

  AppState.threeJsDedicated.currentType = type;

  // Aggregate Data
  const grouped = {};
  let totalSum = 0;
  AppState.currentRows.forEach(row => {
    const k = String(row[xCol] || '(Blank)').trim();
    const v = Number(row[yCol]) || 0;
    if (!grouped[k]) grouped[k] = { sum: 0, count: 0, min: Infinity, max: -Infinity };
    grouped[k].sum += v;
    grouped[k].count++;
    if (v < grouped[k].min) grouped[k].min = v;
    if (v > grouped[k].max) grouped[k].max = v;
    totalSum += v;
  });

  const categories = Object.keys(grouped).slice(0, 24);
  if (categories.length === 0) return;

  let maxVal = 0.0001;
  let minVal = Infinity;
  let peakCategory = '';
  let minCategory = '';

  const parsedItems = categories.map(k => {
    const item = grouped[k];
    let val = item.sum;
    if (agg === 'avg') val = item.sum / (item.count || 1);
    else if (agg === 'count') val = item.count;
    else if (agg === 'max') val = item.max === -Infinity ? 0 : item.max;
    else if (agg === 'min') val = item.min === Infinity ? 0 : item.min;

    if (val > maxVal) { maxVal = val; peakCategory = k; }
    if (val < minVal) { minVal = val; minCategory = k; }
    return { category: k, val: val, count: item.count, sum: item.sum };
  });

  if (minVal === Infinity) minVal = 0;
  const overallAvg = totalSum / (AppState.currentRows.length || 1);

  // Update Under-Stage KPI Cards
  const kpiTotal = document.getElementById('view3dKpiTotal');
  const kpiMax = document.getElementById('view3dKpiMax');
  const kpiMaxLabel = document.getElementById('view3dKpiMaxLabel');
  const kpiMin = document.getElementById('view3dKpiMin');
  const kpiMinLabel = document.getElementById('view3dKpiMinLabel');
  const kpiAvg = document.getElementById('view3dKpiAvg');
  const kpiCount = document.getElementById('view3dKpiCount');

  if (kpiTotal) kpiTotal.textContent = totalSum.toLocaleString(undefined, { maximumFractionDigits: 1 });
  if (kpiMax) kpiMax.textContent = maxVal.toLocaleString(undefined, { maximumFractionDigits: 1 });
  if (kpiMaxLabel) kpiMaxLabel.textContent = `Peak: ${peakCategory || '-'}`;
  if (kpiMin) kpiMin.textContent = minVal.toLocaleString(undefined, { maximumFractionDigits: 1 });
  if (kpiMinLabel) kpiMinLabel.textContent = `Low: ${minCategory || '-'}`;
  if (kpiAvg) kpiAvg.textContent = overallAvg.toLocaleString(undefined, { maximumFractionDigits: 1 });
  if (kpiCount) kpiCount.textContent = `${parsedItems.length} Segments`;

  // Update Stage Header
  const typeTitles = {
    '3d-column': '3D Column Grid',
    '3d-scatter': '3D Bubble Universe',
    '3d-cylinder': '3D Radial Cylinder Donut',
    '3d-surface': '3D Surface Elevation Mesh',
    '3d-matrix': '3D Stepped Grid Waterfall',
    '3d-pyramid': '3D Pyramid & Funnel',
    '3d-pie': '3D Floating Donut Ring'
  };
  const titleText = typeTitles[type] || '3D WebGL Visualization';
  const stageTitle = document.getElementById('view3dStageTitle');
  const stageDesc = document.getElementById('view3dStageDesc');
  if (stageTitle) stageTitle.textContent = `${titleText} — ${agg.toUpperCase()} of ${yCol} by ${xCol}`;
  if (stageDesc) stageDesc.textContent = `Plotted across ${parsedItems.length} categories • Max: ${maxVal.toLocaleString(undefined, { maximumFractionDigits: 1 })} • Orbit Controls Active`;

  const width = container.clientWidth || 800;
  const height = container.clientHeight || 520;

  // Theme configurations
  const themes = {
    cyber: {
      bg: 0x090d16,
      gridPrimary: 0x4f46e5,
      gridSecondary: 0x1e293b,
      lightColor: 0x818cf8,
      rimColor: 0x06b6d4,
      colors: [0x6366f1, 0x06b6d4, 0x10b981, 0xf59e0b, 0xec4899, 0x8b5cf6, 0x3b82f6, 0x14b8a6]
    },
    sunset: {
      bg: 0x12071f,
      gridPrimary: 0xd946ef,
      gridSecondary: 0x2e1065,
      lightColor: 0xf43f5e,
      rimColor: 0xfbbf24,
      colors: [0xf43f5e, 0xf97316, 0xfbbf24, 0xd946ef, 0xa855f7, 0xe11d48, 0xf472b6, 0xfb923c]
    },
    emerald: {
      bg: 0x031713,
      gridPrimary: 0x10b981,
      gridSecondary: 0x064e3b,
      lightColor: 0x34d399,
      rimColor: 0x06b6d4,
      colors: [0x10b981, 0x059669, 0x34d399, 0x06b6d4, 0x14b8a6, 0x6ee7b7, 0x2dd4bf, 0x0284c7]
    },
    minimal: {
      bg: 0x0f172a,
      gridPrimary: 0x64748b,
      gridSecondary: 0x1e293b,
      lightColor: 0xe2e8f0,
      rimColor: 0x38bdf8,
      colors: [0x38bdf8, 0x60a5fa, 0x818cf8, 0xa78bfa, 0xc084fc, 0x94a3b8, 0x475569, 0x93c5fd]
    },
    nebula: {
      bg: 0x0d061c,
      gridPrimary: 0x8b5cf6,
      gridSecondary: 0x3b0764,
      lightColor: 0xa855f7,
      rimColor: 0xec4899,
      colors: [0x8b5cf6, 0xec4899, 0xc084fc, 0xf43f5e, 0x6366f1, 0x06b6d4, 0xd946ef, 0x38bdf8]
    },
    solar: {
      bg: 0x170b05,
      gridPrimary: 0xf97316,
      gridSecondary: 0x431407,
      lightColor: 0xfbbf24,
      rimColor: 0xef4444,
      colors: [0xf97316, 0xf59e0b, 0xfbbf24, 0xe11d48, 0xd97706, 0xb45309, 0xf43f5e, 0xfb923c]
    }
  };

  const activeTheme = themes[AppState.threeJsDedicated.theme] || themes.cyber;

  // Scene & Camera
  const scene = new THREE.Scene();
  scene.background = new THREE.Color(activeTheme.bg);

  const camera = new THREE.PerspectiveCamera(45, width / height, 0.1, 1000);
  camera.position.set(16, 18, 24);
  camera.lookAt(0, 3, 0);

  // Renderer
  const renderer = new THREE.WebGLRenderer({ antialias: true, alpha: true, preserveDrawingBuffer: true });
  renderer.setSize(width, height);
  renderer.setPixelRatio(Math.min(window.devicePixelRatio || 1, 2));
  renderer.shadowMap.enabled = true;
  renderer.shadowMap.type = THREE.PCFSoftShadowMap;
  container.appendChild(renderer.domElement);

  // Lighting
  const ambientLight = new THREE.AmbientLight(0xffffff, 0.85);
  scene.add(ambientLight);

  const keyLight = new THREE.DirectionalLight(activeTheme.lightColor, 1.4);
  keyLight.position.set(16, 26, 22);
  keyLight.castShadow = true;
  keyLight.shadow.mapSize.width = 1024;
  keyLight.shadow.mapSize.height = 1024;
  scene.add(keyLight);

  const rimLight = new THREE.PointLight(activeTheme.rimColor, 1.3, 85);
  rimLight.position.set(-18, 16, -14);
  scene.add(rimLight);

  // Floor Grid
  const gridHelper = new THREE.GridHelper(32, 32, activeTheme.gridPrimary, activeTheme.gridSecondary);
  gridHelper.position.y = 0;
  gridHelper.visible = AppState.threeJsDedicated.showGrid;
  scene.add(gridHelper);
  AppState.threeJsDedicated.gridHelper = gridHelper;

  const interactiveMeshes = [];
  const palette = activeTheme.colors;
  const isWireframe = AppState.threeJsDedicated.wireframe;
  const hScale = AppState.threeJsDedicated.heightScale || 1.0;

  // -----------------------------------------------------------
  // Build 3D Geometries
  // -----------------------------------------------------------
  if (type === '3d-column') {
    const spacing = 2.0;
    const startX = -((parsedItems.length - 1) * spacing) / 2;

    parsedItems.forEach((item, idx) => {
      const heightVal = Math.max(0.5, (item.val / (maxVal || 1)) * 9 * hScale);
      const color = palette[idx % palette.length];

      // Base pedestal
      const baseGeo = new THREE.CylinderGeometry(0.85, 0.95, 0.25, 16);
      const baseMat = new THREE.MeshStandardMaterial({
        color: 0x1e293b,
        metalness: 0.8,
        roughness: 0.2,
        wireframe: isWireframe
      });
      const baseMesh = new THREE.Mesh(baseGeo, baseMat);
      baseMesh.position.set(startX + idx * spacing, 0.125, 0);
      baseMesh.receiveShadow = true;
      scene.add(baseMesh);

      // Column
      const colGeo = new THREE.BoxGeometry(1.2, heightVal, 1.2);
      const colMat = new THREE.MeshPhongMaterial({
        color: color,
        shininess: 90,
        specular: 0xffffff,
        emissive: 0x000000,
        wireframe: isWireframe
      });
      const colMesh = new THREE.Mesh(colGeo, colMat);
      colMesh.position.set(startX + idx * spacing, heightVal / 2 + 0.25, 0);
      colMesh.castShadow = true;
      colMesh.receiveShadow = true;

      // Glow Cap
      const capGeo = new THREE.BoxGeometry(1.22, 0.15, 1.22);
      const capMat = new THREE.MeshBasicMaterial({ color: 0xffffff, wireframe: isWireframe });
      const capMesh = new THREE.Mesh(capGeo, capMat);
      capMesh.position.set(0, heightVal / 2 + 0.08, 0);
      colMesh.add(capMesh);

      const percent = totalSum > 0 ? ((item.val / totalSum) * 100).toFixed(1) : '0';
      colMesh.userData = {
        label: item.category,
        val: item.val,
        percent: percent,
        rank: idx + 1,
        total: parsedItems.length,
        originalColor: color
      };

      scene.add(colMesh);
      interactiveMeshes.push(colMesh);
    });

  } else if (type === '3d-scatter') {
    const count = parsedItems.length;
    parsedItems.forEach((item, idx) => {
      const radius = 0.5 + (item.val / (maxVal || 1)) * 0.9;
      const heightY = (1.0 + (item.val / (maxVal || 1)) * 8) * hScale;
      const posX = ((idx - count / 2) / (count / 2 || 1)) * 12;
      const posZ = Math.sin(idx * 0.8) * 6;
      const color = palette[idx % palette.length];

      // Stem Line
      const points = [new THREE.Vector3(posX, 0, posZ), new THREE.Vector3(posX, heightY, posZ)];
      const lineGeo = new THREE.BufferGeometry().setFromPoints(points);
      const lineMat = new THREE.LineBasicMaterial({ color: activeTheme.gridPrimary, transparent: true, opacity: 0.5 });
      const line = new THREE.Line(lineGeo, lineMat);
      scene.add(line);

      // Sphere
      const sphereGeo = new THREE.SphereGeometry(radius, 32, 32);
      const sphereMat = new THREE.MeshPhysicalMaterial({
        color: color,
        metalness: 0.1,
        roughness: 0.15,
        clearcoat: 0.8,
        clearcoatRoughness: 0.1,
        emissive: 0x000000,
        wireframe: isWireframe
      });
      const sphereMesh = new THREE.Mesh(sphereGeo, sphereMat);
      sphereMesh.position.set(posX, heightY, posZ);
      sphereMesh.castShadow = true;

      // Glow halo ring
      const ringGeo = new THREE.RingGeometry(radius + 0.15, radius + 0.25, 32);
      const ringMat = new THREE.MeshBasicMaterial({ color: color, side: THREE.DoubleSide, transparent: true, opacity: 0.45 });
      const ringMesh = new THREE.Mesh(ringGeo, ringMat);
      ringMesh.rotation.x = Math.PI / 2;
      sphereMesh.add(ringMesh);

      const percent = totalSum > 0 ? ((item.val / totalSum) * 100).toFixed(1) : '0';
      sphereMesh.userData = {
        label: item.category,
        val: item.val,
        percent: percent,
        rank: idx + 1,
        total: parsedItems.length,
        originalColor: color
      };

      scene.add(sphereMesh);
      interactiveMeshes.push(sphereMesh);
    });

  } else if (type === '3d-cylinder') {
    const count = parsedItems.length;
    const ringRadius = 7.8;
    const angleStep = (Math.PI * 2) / Math.max(count, 1);

    parsedItems.forEach((item, idx) => {
      const angle = idx * angleStep;
      const heightVal = Math.max(0.6, (item.val / (maxVal || 1)) * 7.5 * hScale);
      const posX = Math.cos(angle) * ringRadius;
      const posZ = Math.sin(angle) * ringRadius;
      const color = palette[idx % palette.length];

      const cylGeo = new THREE.CylinderGeometry(0.8, 0.8, heightVal, 24);
      const cylMat = new THREE.MeshPhongMaterial({
        color: color,
        shininess: 95,
        specular: 0xffffff,
        emissive: 0x000000,
        wireframe: isWireframe
      });
      const cylMesh = new THREE.Mesh(cylGeo, cylMat);
      cylMesh.position.set(posX, heightVal / 2, posZ);
      cylMesh.castShadow = true;
      cylMesh.receiveShadow = true;

      const topRingGeo = new THREE.CylinderGeometry(0.82, 0.82, 0.1, 24);
      const topRingMat = new THREE.MeshBasicMaterial({ color: 0xffffff, wireframe: isWireframe });
      const topRing = new THREE.Mesh(topRingGeo, topRingMat);
      topRing.position.set(0, heightVal / 2, 0);
      cylMesh.add(topRing);

      const percent = totalSum > 0 ? ((item.val / totalSum) * 100).toFixed(1) : '0';
      cylMesh.userData = {
        label: item.category,
        val: item.val,
        percent: percent,
        rank: idx + 1,
        total: parsedItems.length,
        originalColor: color
      };

      scene.add(cylMesh);
      interactiveMeshes.push(cylMesh);
    });

  } else if (type === '3d-surface') {
    const count = parsedItems.length;
    const segments = count - 1;
    const surfWidth = 19;
    const surfDepth = 8;
    const planeGeo = new THREE.PlaneGeometry(surfWidth, surfDepth, Math.max(segments, 1), 8);
    planeGeo.rotateX(-Math.PI / 2);

    const posAttr = planeGeo.attributes.position;
    for (let i = 0; i < posAttr.count; i++) {
      const x = posAttr.getX(i);
      const normX = (x + surfWidth / 2) / surfWidth;
      const segIdx = Math.min(Math.floor(normX * count), count - 1);
      const itemVal = parsedItems[segIdx]?.val || 0;
      const targetH = (itemVal / (maxVal || 1)) * 6.5 * hScale;
      const z = posAttr.getZ(i);
      const zFactor = 1 - Math.abs(z) / (surfDepth / 2);
      posAttr.setY(i, Math.max(0.2, targetH * Math.max(0.3, zFactor)));
    }
    planeGeo.computeVertexNormals();

    const surfaceMat = new THREE.MeshStandardMaterial({
      color: activeTheme.colors[0],
      metalness: 0.2,
      roughness: 0.35,
      side: THREE.DoubleSide,
      wireframe: isWireframe
    });
    const surfaceMesh = new THREE.Mesh(planeGeo, surfaceMat);
    surfaceMesh.receiveShadow = true;
    surfaceMesh.castShadow = true;
    scene.add(surfaceMesh);

    // Wireframe overlay
    const wireMat = new THREE.MeshBasicMaterial({ color: activeTheme.rimColor, wireframe: true, transparent: true, opacity: 0.35 });
    const wireMesh = new THREE.Mesh(planeGeo, wireMat);
    scene.add(wireMesh);

    // Peak beacons
    parsedItems.forEach((item, idx) => {
      const posX = -surfWidth / 2 + (idx / Math.max(count - 1, 1)) * surfWidth;
      const peakH = Math.max(0.4, (item.val / (maxVal || 1)) * 6.5 * hScale);
      const pinGeo = new THREE.SphereGeometry(0.35, 16, 16);
      const pinMat = new THREE.MeshBasicMaterial({ color: palette[idx % palette.length] });
      const pinMesh = new THREE.Mesh(pinGeo, pinMat);
      pinMesh.position.set(posX, peakH + 0.3, 0);

      const percent = totalSum > 0 ? ((item.val / totalSum) * 100).toFixed(1) : '0';
      pinMesh.userData = {
        label: item.category,
        val: item.val,
        percent: percent,
        rank: idx + 1,
        total: parsedItems.length,
        originalColor: palette[idx % palette.length]
      };
      scene.add(pinMesh);
      interactiveMeshes.push(pinMesh);
    });

  } else if (type === '3d-matrix') {
    const colsCount = Math.min(parsedItems.length, 12);
    const spacingX = 2.0;
    const spacingZ = 2.4;
    const startX = -((colsCount - 1) * spacingX) / 2;

    for (let r = 0; r < 2; r++) {
      for (let c = 0; c < colsCount; c++) {
        const item = parsedItems[c];
        if (!item) continue;
        const multiplier = r === 0 ? 1 : 0.65;
        const calcVal = item.val * multiplier;
        const heightVal = Math.max(0.4, (calcVal / (maxVal || 1)) * 7.5 * hScale);
        const color = r === 0 ? palette[c % palette.length] : 0x475569;

        const boxGeo = new THREE.BoxGeometry(1.2, heightVal, 1.2);
        const boxMat = new THREE.MeshPhongMaterial({
          color: color,
          shininess: 85,
          specular: 0xffffff,
          emissive: 0x000000,
          wireframe: isWireframe
        });
        const boxMesh = new THREE.Mesh(boxGeo, boxMat);
        boxMesh.position.set(startX + c * spacingX, heightVal / 2, (r === 0 ? 1 : -1) * spacingZ);
        boxMesh.castShadow = true;
        boxMesh.receiveShadow = true;

        const percent = totalSum > 0 ? ((calcVal / totalSum) * 100).toFixed(1) : '0';
        boxMesh.userData = {
          label: `${item.category} (${r === 0 ? 'Primary' : 'Benchmark'})`,
          val: calcVal,
          percent: percent,
          rank: c + 1,
          total: colsCount,
          originalColor: color
        };

        scene.add(boxMesh);
        interactiveMeshes.push(boxMesh);
      }
    }

  } else if (type === '3d-pyramid') {
    const count = parsedItems.length;
    const spacing = 2.2;
    const startX = -((count - 1) * spacing) / 2;

    parsedItems.forEach((item, idx) => {
      const heightVal = Math.max(0.6, (item.val / (maxVal || 1)) * 8.5 * hScale);
      const color = palette[idx % palette.length];

      const pyrGeo = new THREE.CylinderGeometry(0.05, 0.9, heightVal, 4);
      const pyrMat = new THREE.MeshPhongMaterial({
        color: color,
        shininess: 90,
        specular: 0xffffff,
        emissive: 0x000000,
        wireframe: isWireframe
      });
      const pyrMesh = new THREE.Mesh(pyrGeo, pyrMat);
      pyrMesh.position.set(startX + idx * spacing, heightVal / 2, 0);
      pyrMesh.rotation.y = Math.PI / 4;
      pyrMesh.castShadow = true;
      pyrMesh.receiveShadow = true;

      const apexGeo = new THREE.SphereGeometry(0.18, 12, 12);
      const apexMat = new THREE.MeshBasicMaterial({ color: 0xffffff });
      const apexMesh = new THREE.Mesh(apexGeo, apexMat);
      apexMesh.position.set(0, heightVal / 2 + 0.1, 0);
      pyrMesh.add(apexMesh);

      const percent = totalSum > 0 ? ((item.val / totalSum) * 100).toFixed(1) : '0';
      pyrMesh.userData = {
        label: item.category,
        val: item.val,
        percent: percent,
        rank: idx + 1,
        total: parsedItems.length,
        originalColor: color
      };

      scene.add(pyrMesh);
      interactiveMeshes.push(pyrMesh);
    });

  } else if (type === '3d-pie') {
    const count = parsedItems.length;
    const ringRadius = 6.5;
    const angleStep = (Math.PI * 2) / Math.max(count, 1);

    parsedItems.forEach((item, idx) => {
      const angle = idx * angleStep;
      const heightVal = Math.max(0.4, (item.val / (maxVal || 1)) * 6.0 * hScale);
      const posX = Math.cos(angle) * ringRadius;
      const posZ = Math.sin(angle) * ringRadius;
      const color = palette[idx % palette.length];

      const wedgeGeo = new THREE.CylinderGeometry(0.85, 0.85, heightVal, 16);
      const wedgeMat = new THREE.MeshStandardMaterial({
        color: color,
        metalness: 0.3,
        roughness: 0.25,
        emissive: 0x000000,
        wireframe: isWireframe
      });
      const wedgeMesh = new THREE.Mesh(wedgeGeo, wedgeMat);
      wedgeMesh.position.set(posX, heightVal / 2 + 0.5, posZ);
      wedgeMesh.castShadow = true;
      wedgeMesh.receiveShadow = true;

      const percent = totalSum > 0 ? ((item.val / totalSum) * 100).toFixed(1) : '0';
      wedgeMesh.userData = {
        label: item.category,
        val: item.val,
        percent: percent,
        rank: idx + 1,
        total: parsedItems.length,
        originalColor: color,
        baseY: heightVal / 2 + 0.5
      };

      scene.add(wedgeMesh);
      interactiveMeshes.push(wedgeMesh);
    });
  }

  AppState.threeJsDedicated.meshes = interactiveMeshes;

  // Controls (OrbitControls or custom fallback)
  let controls = null;
  if (typeof THREE.OrbitControls === 'function') {
    controls = new THREE.OrbitControls(camera, renderer.domElement);
    controls.enableDamping = true;
    controls.dampingFactor = 0.05;
    controls.maxPolarAngle = Math.PI / 2 - 0.04;
    controls.minDistance = 6;
    controls.maxDistance = 80;
    controls.target.set(0, 3, 0);
    AppState.threeJsDedicated.controls = controls;
  } else {
    // Custom pointer orbit fallback
    let isDragging = false;
    let prevMouseX = 0;
    let prevMouseY = 0;
    let theta = Math.PI / 4;
    let phi = Math.PI / 3;
    let radius = 28;

    const updateCameraPos = () => {
      camera.position.x = radius * Math.sin(phi) * Math.sin(theta);
      camera.position.y = radius * Math.cos(phi);
      camera.position.z = radius * Math.sin(phi) * Math.cos(theta);
      camera.lookAt(0, 3, 0);
    };

    renderer.domElement.addEventListener('pointerdown', (e) => {
      isDragging = true;
      prevMouseX = e.clientX;
      prevMouseY = e.clientY;
    });

    window.addEventListener('pointerup', () => { isDragging = false; });

    renderer.domElement.addEventListener('pointermove', (e) => {
      if (!isDragging) return;
      const dx = e.clientX - prevMouseX;
      const dy = e.clientY - prevMouseY;
      prevMouseX = e.clientX;
      prevMouseY = e.clientY;

      theta -= dx * 0.008;
      phi = Math.max(0.1, Math.min(Math.PI / 2 - 0.05, phi - dy * 0.008));
      updateCameraPos();
    });

    renderer.domElement.addEventListener('wheel', (e) => {
      e.preventDefault();
      radius = Math.max(8, Math.min(70, radius + e.deltaY * 0.02));
      updateCameraPos();
    }, { passive: false });
  }

  // Raycaster & Hover HUD
  const raycaster = new THREE.Raycaster();
  const mouse = new THREE.Vector2();
  const tooltip = document.getElementById('view3dTooltip');
  const tipTitle = document.getElementById('view3dTooltipTitle');
  const tipVal = document.getElementById('view3dTooltipValue');
  const tipShare = document.getElementById('view3dTooltipShare');

  function onPointerMove(e) {
    const rect = renderer.domElement.getBoundingClientRect();
    mouse.x = ((e.clientX - rect.left) / rect.width) * 2 - 1;
    mouse.y = -((e.clientY - rect.top) / rect.height) * 2 + 1;

    raycaster.setFromCamera(mouse, camera);
    const intersects = raycaster.intersectObjects(AppState.threeJsDedicated.meshes);

    if (intersects.length > 0) {
      const hit = intersects[0].object;
      if (AppState.threeJsDedicated.activeHoverMesh !== hit) {
        if (AppState.threeJsDedicated.activeHoverMesh && AppState.threeJsDedicated.activeHoverMesh.material) {
          if (AppState.threeJsDedicated.activeHoverMesh.material.emissive) {
            AppState.threeJsDedicated.activeHoverMesh.material.emissive.setHex(0x000000);
          }
        }
        AppState.threeJsDedicated.activeHoverMesh = hit;

        if (hit.material && hit.material.emissive) {
          hit.material.emissive.setHex(0x38bdf8);
        }
      }

      if (hit.userData && tooltip) {
        tipTitle.textContent = hit.userData.label || 'Segment';
        tipVal.textContent = Number(hit.userData.val || 0).toLocaleString(undefined, { maximumFractionDigits: 1 });
        tipShare.textContent = `${hit.userData.percent}% of total • Rank #${hit.userData.rank} of ${hit.userData.total}`;

        const tipX = Math.min(rect.width - 180, Math.max(10, e.clientX - rect.left + 14));
        const tipY = Math.max(10, e.clientY - rect.top - 55);
        tooltip.style.left = `${tipX}px`;
        tooltip.style.top = `${tipY}px`;
        tooltip.classList.remove('hidden');
      }
      container.style.cursor = 'pointer';
    } else {
      if (AppState.threeJsDedicated.activeHoverMesh) {
        if (AppState.threeJsDedicated.activeHoverMesh.material && AppState.threeJsDedicated.activeHoverMesh.material.emissive) {
          AppState.threeJsDedicated.activeHoverMesh.material.emissive.setHex(0x000000);
        }
        AppState.threeJsDedicated.activeHoverMesh = null;
      }
      if (tooltip) tooltip.classList.add('hidden');
      container.style.cursor = 'default';
    }
  }

  function onPointerLeave() {
    if (AppState.threeJsDedicated.activeHoverMesh) {
      if (AppState.threeJsDedicated.activeHoverMesh.material && AppState.threeJsDedicated.activeHoverMesh.material.emissive) {
        AppState.threeJsDedicated.activeHoverMesh.material.emissive.setHex(0x000000);
      }
      AppState.threeJsDedicated.activeHoverMesh = null;
    }
    if (tooltip) tooltip.classList.add('hidden');
  }

  renderer.domElement.addEventListener('pointermove', onPointerMove);
  renderer.domElement.addEventListener('pointerleave', onPointerLeave);

  // Animation Loop
  let rotationAngle = 0;
  function animate() {
    AppState.threeJsDedicated.animationId = requestAnimationFrame(animate);

    if (AppState.threeJsDedicated.autoRotate) {
      const speed = AppState.threeJsDedicated.rotationSpeed || 1.0;
      if (controls) {
        controls.autoRotate = true;
        controls.autoRotateSpeed = 1.6 * speed;
      } else {
        rotationAngle += 0.005 * speed;
        const currentDist = camera.position.distanceTo(new THREE.Vector3(0, 3, 0)) || 28;
        camera.position.x = Math.sin(rotationAngle) * currentDist * 0.8;
        camera.position.z = Math.cos(rotationAngle) * currentDist * 0.8;
        camera.lookAt(0, 3, 0);
      }
    } else {
      if (controls) controls.autoRotate = false;
    }

    if (controls) controls.update();
    renderer.render(scene, camera);
  }

  animate();

  AppState.threeJsDedicated.renderer = renderer;
  AppState.threeJsDedicated.scene = scene;
  AppState.threeJsDedicated.camera = camera;
}

// Dedicated 3D Control Handlers
function toggleDedicated3DAutoRotate() {
  AppState.threeJsDedicated.autoRotate = !AppState.threeJsDedicated.autoRotate;
  const label = document.getElementById('view3dRotateLabel');
  const hudText = document.getElementById('view3dHudRotateText');
  const status = AppState.threeJsDedicated.autoRotate ? 'Active' : 'Paused';
  if (label) label.textContent = status;
  if (hudText) hudText.textContent = AppState.threeJsDedicated.autoRotate ? 'Pause' : 'Play';
  showToast(`3D Auto-Rotation: ${status}`);
}

function setDedicated3DRotationSpeed(speed) {
  AppState.threeJsDedicated.rotationSpeed = parseFloat(speed) || 1.0;
  showToast(`3D Rotation Speed: ${speed}x`);
}

function toggleDedicated3DFloorGrid() {
  AppState.threeJsDedicated.showGrid = !AppState.threeJsDedicated.showGrid;
  if (AppState.threeJsDedicated.gridHelper) {
    AppState.threeJsDedicated.gridHelper.visible = AppState.threeJsDedicated.showGrid;
  }
  const label = document.getElementById('view3dGridLabel');
  if (label) label.textContent = AppState.threeJsDedicated.showGrid ? 'Enabled' : 'Hidden';
  showToast(`3D Floor Grid: ${AppState.threeJsDedicated.showGrid ? 'Enabled' : 'Hidden'}`);
}

function toggleDedicated3DWireframe() {
  AppState.threeJsDedicated.wireframe = !AppState.threeJsDedicated.wireframe;
  const label = document.getElementById('view3dWireLabel');
  if (label) label.textContent = AppState.threeJsDedicated.wireframe ? 'Wireframe' : 'Solid';
  renderDedicated3DView();
  showToast(`3D Material: ${AppState.threeJsDedicated.wireframe ? 'Cyber Wireframe' : 'Solid Surfaces'}`);
}

function setDedicated3DHeightScale(scale) {
  AppState.threeJsDedicated.heightScale = parseFloat(scale) || 1.0;
  const valSpan = document.getElementById('view3dHeightVal');
  if (valSpan) valSpan.textContent = `${parseFloat(scale).toFixed(1)}x`;
  renderDedicated3DView();
}

function updateDedicated3DTheme(themeName) {
  AppState.threeJsDedicated.theme = themeName;
  renderDedicated3DView();
  showToast(`Applied 3D theme: ${themeName}`);
}

function setDedicated3DCamera(mode) {
  const camera = AppState.threeJsDedicated.camera;
  if (!camera) return;

  if (mode === 'isometric') {
    camera.position.set(16, 18, 24);
    camera.lookAt(0, 3, 0);
  } else if (mode === 'top') {
    camera.position.set(0, 34, 0.01);
    camera.lookAt(0, 0, 0);
  } else if (mode === 'front') {
    camera.position.set(0, 8, 30);
    camera.lookAt(0, 3, 0);
  } else if (mode === 'side') {
    camera.position.set(30, 8, 0);
    camera.lookAt(0, 3, 0);
  }

  if (AppState.threeJsDedicated.controls) {
    AppState.threeJsDedicated.controls.target.set(0, 3, 0);
    AppState.threeJsDedicated.controls.update();
  }
  showToast(`Switched to ${mode.toUpperCase()} 3D angle`);
}

function resetDedicated3DCamera() {
  setDedicated3DCamera('isometric');
}

function toggleDedicated3DFullscreen() {
  const container = document.getElementById('dedicatedThreeJsCanvasContainer');
  if (!container) return;
  AppState.threeJsDedicated.isFullscreen = !AppState.threeJsDedicated.isFullscreen;
  container.classList.toggle('fullscreen-3d-active', AppState.threeJsDedicated.isFullscreen);

  setTimeout(() => {
    resizeDedicated3D();
  }, 100);

  showToast(AppState.threeJsDedicated.isFullscreen ? 'Entered 3D fullscreen mode (Press Esc to exit)' : 'Exited 3D fullscreen mode');
}

function downloadDedicated3DChart() {
  if (!AppState.threeJsDedicated.renderer || !AppState.threeJsDedicated.scene || !AppState.threeJsDedicated.camera) {
    showToast('3D Scene not ready to export', true);
    return;
  }
  AppState.threeJsDedicated.renderer.render(AppState.threeJsDedicated.scene, AppState.threeJsDedicated.camera);
  const url = AppState.threeJsDedicated.renderer.domElement.toDataURL('image/png');
  const a = document.createElement('a');
  a.href = url;
  a.download = `3D_Visual_${AppState.threeJsDedicated.currentType}_${Date.now()}.png`;
  a.click();
  showToast('Downloaded High-Res 3D snapshot!');
}

// Escape key listener for dedicated 3D fullscreen
window.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && AppState.threeJsDedicated.isFullscreen) {
    toggleDedicated3DFullscreen();
  }
});

// Window resize listener
window.addEventListener('resize', () => {
  if (AppState.threeJsDedicated.renderer) resizeDedicated3D();
});

// -------------------------------------------------------------
// 🗂️ 4. Dynamic Pivot Matrix Crosstab
// -------------------------------------------------------------
function populatePivotDropdowns() {
  const rowSel = document.getElementById('pivotRowSelect');
  const colSel = document.getElementById('pivotColSelect');
  const valSel = document.getElementById('pivotValSelect');
  if (!rowSel || !colSel || !valSel) return;

  rowSel.innerHTML = '';
  colSel.innerHTML = '';
  valSel.innerHTML = '';

  AppState.columns.forEach((col, idx) => {
    rowSel.innerHTML += `<option value="${col}">${col}</option>`;
    colSel.innerHTML += `<option value="${col}">${col}</option>`;
  });

  const numCols = AppState.columns.filter(c => AppState.colMeta[c]?.type === 'numeric');
  (numCols.length > 0 ? numCols : AppState.columns).forEach(col => {
    valSel.innerHTML += `<option value="${col}">${col}</option>`;
  });

  if (AppState.columns.length > 1) {
    rowSel.selectedIndex = 0;
    colSel.selectedIndex = 1;
  }
}

function renderPivotMatrix() {
  const rowCol = document.getElementById('pivotRowSelect')?.value;
  const colCol = document.getElementById('pivotColSelect')?.value;
  const valCol = document.getElementById('pivotValSelect')?.value;
  const agg = document.getElementById('pivotAggSelect')?.value || 'sum';

  if (!rowCol || !colCol || !valCol) return;

  const matrix = {};
  const distinctColVals = new Set();
  const distinctRowVals = new Set();

  AppState.currentRows.forEach(r => {
    const rKey = groupKeyOf(r[rowCol]);
    const cKey = groupKeyOf(r[colCol]);
    distinctRowVals.add(rKey);
    distinctColVals.add(cKey);
    if (!matrix[rKey]) matrix[rKey] = {};
    if (!matrix[rKey][cKey]) matrix[rKey][cKey] = { rows: 0, nums: [] };
    matrix[rKey][cKey].rows++;
    const val = toNumberOrNull(r[valCol]);
    if (val !== null) matrix[rKey][cKey].nums.push(val);
  });

  const sortKeys = arr => arr.sort((a, b) => a === '(Blank)' ? 1 : b === '(Blank)' ? -1 : a.localeCompare(b, undefined, { numeric: true }));
  const colHeaders = sortKeys(Array.from(distinctColVals));
  const rowHeaders = sortKeys(Array.from(distinctRowVals));

  const thead = document.getElementById('pivotTableHead');
  const tbody = document.getElementById('pivotTableBody');
  if (!thead || !tbody) return;

  const cellValue = (cells) => {
    const nums = cells.flatMap(c => c.nums);
    const rows = cells.reduce((s, c) => s + c.rows, 0);
    return aggregateNumbers(nums, agg, rows);
  };

  let headHtml = `<tr><th class="py-2.5 px-3 bg-slate-100 font-bold text-slate-800">${escapeHtml(rowCol)} \\ ${escapeHtml(colCol)}</th>`;
  colHeaders.forEach(ch => { headHtml += `<th class="py-2.5 px-3 text-right bg-slate-50">${escapeHtml(ch)}</th>`; });
  headHtml += `<th class="py-2.5 px-3 text-right font-bold bg-indigo-50 text-indigo-700">Total</th></tr>`;
  thead.innerHTML = headHtml;

  let bodyHtml = '';
  rowHeaders.forEach(rh => {
    bodyHtml += `<tr class="hover:bg-slate-50/70"><td class="py-2.5 px-3 font-sans font-semibold text-slate-900 border-r border-slate-100">${escapeHtml(rh)}</td>`;
    colHeaders.forEach(ch => {
      const cell = matrix[rh]?.[ch];
      const val = cell ? cellValue([cell]) : null;
      bodyHtml += `<td class="py-2.5 px-3 text-right ${val ? 'text-slate-800 font-medium' : 'text-slate-300'}">${val === null || val === undefined ? '–' : formatNumber(val, 1)}</td>`;
    });
    const rowTotal = cellValue(Object.values(matrix[rh] || {}));
    bodyHtml += `<td class="py-2.5 px-3 text-right font-bold text-indigo-600 bg-indigo-50/40">${formatNumber(rowTotal, 1)}</td></tr>`;
  });

  // Column totals
  bodyHtml += `<tr class="border-t-2 border-slate-200 bg-slate-50 font-semibold"><td class="py-2.5 px-3 font-sans">Total</td>`;
  colHeaders.forEach(ch => {
    const v = cellValue(rowHeaders.map(rh => matrix[rh]?.[ch]).filter(Boolean));
    bodyHtml += `<td class="py-2.5 px-3 text-right">${formatNumber(v, 1)}</td>`;
  });
  const grand = cellValue(rowHeaders.flatMap(rh => Object.values(matrix[rh] || {})));
  bodyHtml += `<td class="py-2.5 px-3 text-right text-indigo-700 bg-indigo-50">${formatNumber(grand, 1)}</td></tr>`;

  tbody.innerHTML = bodyHtml;
}

// -------------------------------------------------------------
// ✏️ 5. Double Click Inline Cell Editing
// -------------------------------------------------------------
function makeCellEditable(td) {
  if (td.querySelector('input')) return;
  const currentVal = td.querySelector('.cell-empty') ? '' : td.innerText.trim();
  const rowIdx = parseInt(td.getAttribute('data-row-idx'));
  const colName = td.getAttribute('data-col');

  const input = document.createElement('input');
  input.type = 'text';
  input.value = currentVal;
  input.className = 'w-full px-2 py-1 text-xs border border-indigo-500 rounded bg-white shadow-xs focus:outline-none';

  td.innerHTML = '';
  td.appendChild(input);
  input.focus();

  const saveCell = () => {
    const newVal = input.value.trim();
    if (AppState.filteredRows[rowIdx]) {
      AppState.filteredRows[rowIdx][colName] = AppState.colMeta[colName]?.type === 'numeric' && !isNaN(Number(newVal))
        ? Number(newVal)
        : newVal;
    }
    td.innerHTML = newVal ? escapeHtml(newVal) : '<span class="cell-empty" title="Empty cell">—</span>';
    td.classList.add('bg-amber-50');
    refreshPostTransformation(`Updated ${colName} on row #${rowIdx + 1}`, { keepGrid: true });
  };

  input.addEventListener('blur', saveCell);
  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      input.blur();
    } else if (e.key === 'Escape') {
      input.removeEventListener('blur', saveCell);
      td.innerHTML = currentVal ? escapeHtml(currentVal) : '<span class="cell-empty" title="Empty cell">—</span>';
    }
  });
}

// -------------------------------------------------------------
// 📥 6. Full Analysis PDF Report Generator (jsPDF + html2canvas)
// -------------------------------------------------------------
// Render the current Chart Studio settings on an off-screen canvas, so the PDF works from any tab
function renderChartImageForReport() {
  const xCol = document.getElementById('studioXCol')?.value;
  const yCol = document.getElementById('studioYCol')?.value;
  const agg = document.getElementById('studioAgg')?.value || 'sum';
  let type = document.getElementById('studioChartType')?.value || 'bar';
  if (!xCol || !yCol) return null;
  if (type.startsWith('3d-') || type === 'radar') type = 'bar';

  const grouped = {};
  AppState.currentRows.forEach(r => {
    const k = groupKeyOf(r[xCol]);
    grouped[k] = grouped[k] || { count: 0, nums: [] };
    grouped[k].count++;
    const v = toNumberOrNull(r[yCol]);
    if (v !== null) grouped[k].nums.push(v);
  });
  const labels = Object.keys(grouped).slice(0, 25);
  const data = labels.map(k => aggregateNumbers(grouped[k].nums, agg, grouped[k].count));

  const canvas = document.createElement('canvas');
  canvas.width = 1400;
  canvas.height = 640;
  canvas.style.cssText = 'position:fixed;left:-10000px;top:0;width:1400px;height:640px';
  document.body.appendChild(canvas);
  try {
    const isArea = type === 'area';
    const scatter = type === 'scatter';
    const chart = new Chart(canvas, {
      type: isArea ? 'line' : type,
      data: scatter
        ? { datasets: [{ data: AppState.currentRows.map(r => ({ x: toNumberOrNull(r[xCol]), y: toNumberOrNull(r[yCol]) })).filter(p => p.x !== null && p.y !== null), backgroundColor: '#0d44e8' }] }
        : { labels, datasets: [{ data, fill: isArea, backgroundColor: ['pie', 'doughnut'].includes(type) ? ['#0d44e8', '#5a84fb', '#0f172a', '#94a3b8', '#10b981', '#f59e0b', '#8eadff', '#334155'] : '#0d44e8', borderColor: '#0d44e8', borderRadius: type === 'bar' ? 6 : 0 }] },
      options: {
        responsive: false, animation: false, devicePixelRatio: 1,
        plugins: { legend: { display: ['pie', 'doughnut'].includes(type), position: 'right' } },
        scales: ['pie', 'doughnut'].includes(type) ? {} : { y: { grid: { color: '#eef2f7' } }, x: { grid: { display: false } } }
      },
      plugins: [{ id: 'whiteBg', beforeDraw: c => { const ctx = c.ctx; ctx.save(); ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, c.width, c.height); ctx.restore(); } }]
    });
    const url = canvas.toDataURL('image/jpeg', 0.92);
    chart.destroy();
    return { url, title: `${agg.toUpperCase()} of ${yCol} by ${xCol}` };
  } catch (e) {
    console.warn('Report chart failed:', e);
    return null;
  } finally {
    canvas.remove();
  }
}

function stripMarkdown(text) {
  return String(text || '')
    .replace(/^\s*#{1,6}\s+/gm, '')
    .replace(/\*\*(.+?)\*\*/g, '$1')
    .replace(/(^|\s)\*(?!\s)(.+?)\*/g, '$1$2')
    .replace(/^\s*[-*]\s+/gm, '• ')
    .replace(/`/g, '');
}

async function exportAnalysisReportPDF() {
  document.getElementById('exportDropdownMenu')?.classList.add('hidden');
  if (!window.jspdf) {
    showToast('PDF library is still loading, try again in a moment', true);
    return;
  }

  try {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('p', 'mm', 'a4');
    const W = 210, M = 16;
    let y = 0;
    const ensure = h => { if (y + h > 282) { doc.addPage(); y = 18; } };
    const heading = text => { ensure(14); doc.setFont('helvetica', 'bold'); doc.setFontSize(12); doc.setTextColor(13, 68, 232); doc.text(text, M, y); y += 7; };
    const para = (text, size = 9.5) => {
      doc.setFont('helvetica', 'normal'); doc.setFontSize(size); doc.setTextColor(43, 47, 54);
      doc.splitTextToSize(text, W - 2 * M).forEach(line => { ensure(5); doc.text(line, M, y); y += 4.6; });
      y += 2;
    };

    // Header
    doc.setFillColor(13, 68, 232);
    doc.rect(0, 0, W, 4, 'F');
    doc.setTextColor(0, 0, 0);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(20);
    doc.text('Data Vista report', M, 20);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(9.5);
    doc.setTextColor(100, 116, 139);
    doc.text(`${AppState.fileName} · sheet "${AppState.activeSheet}" · ${AppState.currentRows.length} rows × ${AppState.columns.length} columns · ${new Date().toLocaleString()}`, M, 27);
    y = 38;

    // Column summary
    heading('Numeric columns');
    const numCols = AppState.columns.filter(c => AppState.colMeta[c]?.type === 'numeric');
    if (numCols.length) {
      doc.setFontSize(8.5);
      const cols = ['Column', 'Filled', 'Sum', 'Average', 'Min', 'Max'];
      const xs = [M, M + 52, M + 72, M + 102, M + 132, M + 157];
      doc.setFont('helvetica', 'bold'); doc.setTextColor(0, 0, 0);
      cols.forEach((c, i) => doc.text(c, xs[i], y));
      y += 2; doc.setDrawColor(13, 68, 232); doc.line(M, y, W - M, y); y += 4.5;
      doc.setFont('helvetica', 'normal'); doc.setTextColor(43, 47, 54);
      numCols.slice(0, 15).forEach(c => {
        ensure(6);
        const m = AppState.colMeta[c];
        [String(c).slice(0, 28), `${m.nonNull}/${AppState.currentRows.length}`, formatNumber(m.sum, 1), formatNumber(m.avg, 2), formatNumber(m.min, 2), formatNumber(m.max, 2)]
          .forEach((v, i) => doc.text(String(v), xs[i], y));
        y += 5;
      });
      y += 3;
    } else {
      para('This sheet has no numeric columns.');
    }

    // Chart
    const chart = renderChartImageForReport();
    if (chart) {
      heading(`Chart: ${chart.title}`);
      ensure(84);
      doc.addImage(chart.url, 'JPEG', M, y, W - 2 * M, (W - 2 * M) * 640 / 1400);
      y += (W - 2 * M) * 640 / 1400 + 6;
    }

    // AutoML
    const a = AppState.automl;
    if (a.bestModel && a.models?.length) {
      const info = mlMetricInfo(a.taskType, a.primaryMetric);
      heading(`AutoML: predicting ${a.targetColumn} (${a.taskType})`);
      para(`Evaluation: ${a.evaluation?.label}. Ranked by ${info.short}. Winner: ${a.bestModel.name}.`);
      a.models.forEach(m => para(`#${m.rank}  ${m.name}: ${info.short} ${mlFormatMetric(a.taskType, a.primaryMetric, m.metrics[a.primaryMetric])}`, 9));
      if (a.featureImportance?.length) para('Top drivers: ' + a.featureImportance.slice(0, 4).map(f => `${f.feature} (${f.importance}%)`).join(', '), 9);
    }

    // AI summary
    if (AppState.lastAiInsight) {
      heading('AI executive summary');
      para(stripMarkdown(AppState.lastAiInsight));
    }

    const pages = doc.getNumberOfPages();
    for (let i = 1; i <= pages; i++) {
      doc.setPage(i);
      doc.setFontSize(8); doc.setTextColor(148, 163, 184);
      doc.text(`Data Vista · page ${i} of ${pages}`, M, 290);
    }

    doc.save(`DataVista_Report_${AppState.activeSheet}_${Date.now()}.pdf`);
    showToast('PDF report downloaded');
  } catch (err) {
    console.error(err);
    showToast(`Could not build the PDF: ${err.message}`, true);
  }
}

// -------------------------------------------------------------
// 🔐 7. User Authentication & Admin Management
// -------------------------------------------------------------
let authMode = 'login';

function openAuthModal() {
  document.getElementById('authModal').classList.remove('hidden');
}

function closeAuthModal() {
  document.getElementById('authModal').classList.add('hidden');
}

function switchAuthMode(mode) {
  authMode = mode;
  const isLogin = mode === 'login';
  document.getElementById('authTabLogin').className = isLogin
    ? 'flex-1 pb-2.5 text-xs font-bold text-indigo-600 border-b-2 border-indigo-600 text-center'
    : 'flex-1 pb-2.5 text-xs font-bold text-slate-400 border-b-2 border-transparent hover:text-slate-700 text-center';
  document.getElementById('authTabRegister').className = !isLogin
    ? 'flex-1 pb-2.5 text-xs font-bold text-indigo-600 border-b-2 border-indigo-600 text-center'
    : 'flex-1 pb-2.5 text-xs font-bold text-slate-400 border-b-2 border-transparent hover:text-slate-700 text-center';

  document.getElementById('authNameGroup').classList.toggle('hidden', isLogin);
  document.getElementById('authRoleGroup').classList.toggle('hidden', isLogin);
  document.getElementById('authSubmitBtn').textContent = isLogin ? 'Sign In' : 'Create Account';
  document.getElementById('authModalTitle').textContent = isLogin ? 'Sign in to Data Vista' : 'Create your Account';
}

async function handleAuthSubmit(e) {
  e.preventDefault();
  const email = document.getElementById('authEmailInput').value;
  const password = document.getElementById('authPasswordInput').value;
  const name = document.getElementById('authNameInput').value;
  const role = document.getElementById('authRoleInput').value;

  const endpoint = authMode === 'login' ? '/auth/login' : '/auth/register';
  const payload = authMode === 'login' ? { email, password } : { email, password, name, role };

  try {
    const res = await fetch(`${AppState.apiBaseUrl}${endpoint}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });

    const data = await res.json();
    if (!res.ok || !data.success) {
      throw new Error(data.message || 'Authentication error');
    }

    localStorage.setItem('excelinsight_token', data.token);
    localStorage.setItem('excelinsight_user', JSON.stringify(data.user));
    AppState.auth.token = data.token;
    AppState.auth.user = data.user;

    checkAuthState();
    closeAuthModal();
    showToast(`Welcome, ${data.user.name}!`);

    // Save current active spreadsheet to history automatically
    saveUploadHistoryRecord();
  } catch (err) {
    showToast(err.message, true);
  }
}

function handleUserLogout() {
  localStorage.removeItem('excelinsight_token');
  localStorage.removeItem('excelinsight_user');
  AppState.auth.token = null;
  AppState.auth.user = null;
  checkAuthState();
  showToast('Signed out successfully');
}

function checkAuthState() {
  const loginBtn = document.getElementById('authLoginBtn');
  const userMenu = document.getElementById('authUserMenu');
  const userName = document.getElementById('authUserName');
  const roleBadge = document.getElementById('authUserRoleBadge');
  const adminTab = document.getElementById('tabBtn-admin');
  const historyNotice = document.getElementById('historyAuthNotice');

  if (AppState.auth.user && AppState.auth.token) {
    loginBtn.classList.add('hidden');
    userMenu.classList.remove('hidden');
    userName.textContent = AppState.auth.user.name;

    const isAdmin = AppState.auth.user.role === 'admin';
    roleBadge.classList.toggle('hidden', !isAdmin);
    adminTab.classList.toggle('hidden', !isAdmin);
    if (historyNotice) historyNotice.classList.add('hidden');
  } else {
    loginBtn.classList.remove('hidden');
    userMenu.classList.add('hidden');
    adminTab.classList.add('hidden');
    if (historyNotice) historyNotice.classList.remove('hidden');
  }
  lucide.createIcons();
}

// -------------------------------------------------------------
// 📂 8. MongoDB Upload History Service
// -------------------------------------------------------------
async function saveUploadHistoryRecord() {
  if (!AppState.auth.token) return;

  try {
    await fetch(`${AppState.apiBaseUrl}/uploads`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${AppState.auth.token}`
      },
      body: JSON.stringify({
        fileName: AppState.fileName,
        sheetsCount: Object.keys(AppState.sheets).length,
        sheetNames: Object.keys(AppState.sheets),
        totalRows: AppState.currentRows.length,
        totalColumns: AppState.columns.length,
        columns: AppState.columns.map(c => ({ name: c, type: AppState.colMeta[c]?.type || 'categorical' }))
      })
    });
  } catch (err) {
    console.warn('Auto-save history note:', err);
  }
}

async function fetchUploadHistory() {
  if (!AppState.auth.token) {
    showToast('Sign in to view cloud upload history', true);
    return;
  }

  const tbody = document.getElementById('historyTableBody');
  tbody.innerHTML = `<tr><td colspan="6" class="py-6 text-center text-slate-400 font-sans"><i data-lucide="loader-2" class="w-5 h-5 animate-spin mx-auto text-indigo-600 mb-1"></i> Loading MongoDB History...</td></tr>`;
  lucide.createIcons();

  try {
    const res = await fetch(`${AppState.apiBaseUrl}/uploads`, {
      headers: { 'Authorization': `Bearer ${AppState.auth.token}` }
    });
    const result = await res.json();

    if (!result.data || result.data.length === 0) {
      tbody.innerHTML = `<tr><td colspan="6" class="py-6 text-center text-slate-400 font-sans">No saved uploads in MongoDB yet.</td></tr>`;
      return;
    }

    let rowsHtml = '';
    result.data.forEach(item => {
      const dateStr = new Date(item.uploadedAt).toLocaleDateString() + ' ' + new Date(item.uploadedAt).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
      rowsHtml += `
        <tr class="hover:bg-slate-50 transition border-b border-slate-100">
          <td class="py-2.5 px-3 font-semibold text-slate-900">${item.fileName}</td>
          <td class="py-2.5 px-3 text-slate-600">${item.sheetsCount} sheets</td>
          <td class="py-2.5 px-3 text-slate-700 font-mono">${item.totalRows}</td>
          <td class="py-2.5 px-3 text-slate-700 font-mono">${item.totalColumns}</td>
          <td class="py-2.5 px-3 text-slate-500 font-sans text-xs">${dateStr}</td>
          <td class="py-2.5 px-3 text-right">
            <button onclick="deleteUploadRecord('${item._id}')" class="text-rose-500 hover:text-rose-700 text-xs font-semibold px-2 py-1 rounded hover:bg-rose-50 transition">
              Delete
            </button>
          </td>
        </tr>
      `;
    });
    tbody.innerHTML = rowsHtml;
  } catch (err) {
    tbody.innerHTML = `<tr><td colspan="6" class="py-6 text-center text-rose-500 font-sans">Failed to load history: ${err.message}</td></tr>`;
  }
}

async function deleteUploadRecord(id) {
  try {
    const res = await fetch(`${AppState.apiBaseUrl}/uploads/${id}`, {
      method: 'DELETE',
      headers: { 'Authorization': `Bearer ${AppState.auth.token}` }
    });
    if (res.ok) {
      showToast('Record removed from history');
      fetchUploadHistory();
    }
  } catch (err) {
    showToast(err.message, true);
  }
}

// -------------------------------------------------------------
// 🛡️ 9. Admin Dashboard Service
// -------------------------------------------------------------
async function fetchAdminData() {
  if (!AppState.auth.token) return;

  try {
    // 1. Fetch Stats
    const statsRes = await fetch(`${AppState.apiBaseUrl}/admin/stats`, {
      headers: { 'Authorization': `Bearer ${AppState.auth.token}` }
    });
    const statsData = await statsRes.json();
    if (statsData.success) {
      document.getElementById('adminTotalUsers').textContent = statsData.data.totalUsers.toLocaleString();
      document.getElementById('adminTotalUploads').textContent = statsData.data.totalUploads.toLocaleString();
      document.getElementById('adminTotalRows').textContent = statsData.data.totalRowsProcessed.toLocaleString();
    }

    // 2. Fetch Users
    const usersRes = await fetch(`${AppState.apiBaseUrl}/admin/users`, {
      headers: { 'Authorization': `Bearer ${AppState.auth.token}` }
    });
    const usersData = await usersRes.json();
    if (usersData.success) {
      const tbody = document.getElementById('adminUsersTableBody');
      let html = '';
      usersData.users.forEach(u => {
        const isAdm = u.role === 'admin';
        html += `
          <tr class="hover:bg-slate-50 transition border-b border-slate-100">
            <td class="py-2.5 px-3 font-semibold text-slate-900">${u.name}</td>
            <td class="py-2.5 px-3 text-slate-600 font-sans">${u.email}</td>
            <td class="py-2.5 px-3">
              <span class="px-2 py-0.5 rounded-full text-[10px] font-bold uppercase ${isAdm ? 'bg-amber-100 text-amber-800 border border-amber-200' : 'bg-slate-100 text-slate-700'}">${u.role}</span>
            </td>
            <td class="py-2.5 px-3 font-mono">${u.uploadsCount}</td>
            <td class="py-2.5 px-3 text-slate-500 font-sans text-xs">${new Date(u.createdAt).toLocaleDateString()}</td>
            <td class="py-2.5 px-3 text-right">
              <button onclick="toggleUserRole('${u.id}', '${isAdm ? 'user' : 'admin'}')" class="text-indigo-600 hover:text-indigo-800 text-xs font-semibold px-2 py-1 rounded hover:bg-indigo-50 transition">
                Make ${isAdm ? 'User' : 'Admin'}
              </button>
            </td>
          </tr>
        `;
      });
      tbody.innerHTML = html;
    }
  } catch (err) {
    showToast(err.message, true);
  }
}

async function toggleUserRole(userId, newRole) {
  try {
    const res = await fetch(`${AppState.apiBaseUrl}/admin/users/${userId}/role`, {
      method: 'PATCH',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${AppState.auth.token}`
      },
      body: JSON.stringify({ role: newRole })
    });
    if (res.ok) {
      showToast(`Role updated to ${newRole}`);
      fetchAdminData();
    }
  } catch (err) {
    showToast(err.message, true);
  }
}

// -------------------------------------------------------------
// 🧮 10. Interactive In-Browser Formula Engine (fx)
// -------------------------------------------------------------
const FORMULA_FUNCTIONS = {
  SUM: nums => aggregateNumbers(nums, 'sum') ?? 0,
  AVERAGE: nums => aggregateNumbers(nums, 'avg'),
  AVG: nums => aggregateNumbers(nums, 'avg'),
  MEDIAN: nums => aggregateNumbers(nums, 'median'),
  MIN: nums => aggregateNumbers(nums, 'min'),
  MAX: nums => aggregateNumbers(nums, 'max'),
  COUNT: nums => nums.length,
  STDEV: nums => {
    if (nums.length < 2) return null;
    const m = aggregateNumbers(nums, 'avg');
    return Math.sqrt(nums.reduce((s, v) => s + (v - m) ** 2, 0) / (nums.length - 1));
  }
};

function resolveFormulaColumn(raw) {
  const name = raw.trim().replace(/^\[|\]$/g, '').replace(/^["']|["']$/g, '').trim();
  const exact = AppState.columns.find(c => c === name);
  return exact || AppState.columns.find(c => c.toLowerCase() === name.toLowerCase()) || null;
}

// Evaluate formulas like =SUM(Sales), =AVERAGE(Profit)*2, =SUM(Profit)/SUM(Sales)*100
function evaluateFormula(expr) {
  const rows = AppState.currentRows;
  const notes = [];
  const replaced = expr.replace(/([A-Za-z]+)\s*\(([^()]*)\)/g, (_, fnRaw, arg) => {
    const fn = fnRaw.toUpperCase();
    const col = resolveFormulaColumn(arg);
    if (!col) throw new Error(`Unknown column "${arg.trim()}". Columns: ${AppState.columns.join(', ')}`);
    if (fn === 'COUNTA') return `(${rows.filter(r => !isBlankValue(r[col])).length})`;
    if (fn === 'COUNTBLANK') return `(${rows.filter(r => isBlankValue(r[col])).length})`;
    const impl = FORMULA_FUNCTIONS[fn];
    if (!impl) throw new Error(`Unsupported function ${fn}. Try SUM, AVERAGE, MEDIAN, MIN, MAX, COUNT, COUNTA, COUNTBLANK or STDEV.`);
    const nums = numericValues(rows, col);
    if (fn === 'COUNT' && nums.length === 0 && rows.some(r => !isBlankValue(r[col]))) {
      notes.push(`COUNT only counts numbers; use COUNTA(${col}) to count text cells`);
    }
    const v = impl(nums);
    if (v === null) throw new Error(`${fn}(${col}) has no numeric values`);
    return `(${v})`;
  });

  if (/[A-Za-z_]/.test(replaced.replace(/e[+-]?\d/gi, ''))) {
    throw new Error('Column names must go inside a function, e.g. =SUM(Sales)');
  }
  if (!/^[0-9+\-*/().,\s^eE]+$/.test(replaced)) throw new Error('Only numbers, + - * / ^ and brackets are allowed');
  const result = Function(`'use strict'; return (${replaced.replace(/\^/g, '**')})`)();
  if (typeof result !== 'number' || !Number.isFinite(result)) throw new Error('The formula does not produce a finite number');
  return { result, notes };
}

function executeFormulaBar() {
  const input = document.getElementById('excelFormulaInput');
  const badge = document.getElementById('formulaResultBadge');
  if (!input) return;

  let expr = input.value.trim();
  if (!expr) return;
  if (expr.startsWith('=')) expr = expr.substring(1).trim();

  const show = (ok, text) => {
    badge.classList.remove('hidden', 'text-rose-700', 'bg-rose-50', 'border-rose-200', 'text-emerald-700', 'bg-emerald-50', 'border-emerald-200');
    badge.classList.add(...(ok ? ['text-emerald-700', 'bg-emerald-50', 'border-emerald-200'] : ['text-rose-700', 'bg-rose-50', 'border-rose-200']));
    badge.textContent = text;
    badge.title = text;
  };

  try {
    const { result, notes } = evaluateFormula(expr);
    const formatted = formatNumber(result, 4);
    show(true, `${expr} = ${formatted}`);
    showToast(notes.length ? `${formatted} · ${notes[0]}` : `Result: ${formatted}`);
  } catch (err) {
    show(false, 'Formula error');
    showToast(err.message, true);
  }
}

// -------------------------------------------------------------
// ⚖️ 11. Dual-Sheet & File Comparison (Excel Diff Engine)
// -------------------------------------------------------------
function populateDiffDropdowns() {
  const baseSel = document.getElementById('diffBaseSheetSelect');
  const targetSel = document.getElementById('diffTargetSheetSelect');
  if (!baseSel || !targetSel) return;

  baseSel.innerHTML = '';
  targetSel.innerHTML = '';
  const sheetNames = Object.keys(AppState.sheets);
  sheetNames.forEach(name => {
    baseSel.innerHTML += `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`;
    targetSel.innerHTML += `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`;
  });
  if (sheetNames.length > 1) targetSel.selectedIndex = 1;
  populateDiffKeyDropdown();
}

// A good key is filled and (almost) unique in both sheets: OrderID, EmpID, CustomerID...
// A few repeats are fine (duplicate rows); they are matched in order of appearance.
function detectDiffKeyColumn(baseRows, targetRows, cols) {
  const keyScore = (rows, c) => {
    const filledRows = rows.filter(r => cols.some(cc => !isBlankValue(r[cc]))).length || 1;
    const vals = rows.map(r => r[c]).filter(v => !isBlankValue(v)).map(v => String(v).trim());
    if (vals.length < filledRows * 0.9) return 0;
    return new Set(vals).size / vals.length;
  };
  const candidates = cols
    .map(c => ({ c, s: Math.min(keyScore(baseRows, c), keyScore(targetRows, c)) }))
    .filter(x => x.s >= 0.9)
    .sort((a, b) => b.s - a.s);
  const named = candidates.find(x => /(^id$|id$|_id|key|code|no$|number)/i.test(x.c));
  return (named || candidates[0] || {}).c || '';
}

function populateDiffKeyDropdown() {
  const keySel = document.getElementById('diffKeyColSelect');
  if (!keySel) return;
  const baseRows = AppState.sheets[document.getElementById('diffBaseSheetSelect')?.value] || [];
  const targetRows = AppState.sheets[document.getElementById('diffTargetSheetSelect')?.value] || [];
  const cols = Array.from(new Set([...(baseRows[0] ? Object.keys(baseRows[0]) : []), ...(targetRows[0] ? Object.keys(targetRows[0]) : [])]));
  const auto = detectDiffKeyColumn(baseRows, targetRows, cols);
  keySel.innerHTML = cols.map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('') + '<option value="__position__">(Row position)</option>';
  keySel.value = auto || '__position__';
}

function diffValuesEqual(a, b) {
  if (isBlankValue(a) && isBlankValue(b)) return true;
  const na = toNumberOrNull(a), nb = toNumberOrNull(b);
  if (na !== null && nb !== null) return na === nb;
  return String(a ?? '').trim() === String(b ?? '').trim();
}

function runSheetComparison() {
  const baseName = document.getElementById('diffBaseSheetSelect')?.value;
  const targetName = document.getElementById('diffTargetSheetSelect')?.value;
  if (!baseName || !targetName) return;
  const keyCol = document.getElementById('diffKeyColSelect')?.value || '__position__';
  const onlyChanges = document.getElementById('diffOnlyChangesToggle')?.checked;

  const nonEmpty = rows => rows.filter(r => Object.values(r).some(v => !isBlankValue(v)));
  const baseRows = nonEmpty(AppState.sheets[baseName] || []);
  const targetRows = nonEmpty(AppState.sheets[targetName] || []);

  const summary = document.getElementById('diffSummaryContainer');
  const thead = document.getElementById('diffHeaderRow');
  const tbody = document.getElementById('diffTableBody');
  summary.classList.remove('hidden');

  const allCols = Array.from(new Set([...(baseRows[0] ? Object.keys(baseRows[0]) : []), ...(targetRows[0] ? Object.keys(targetRows[0]) : [])]));
  thead.innerHTML = `<th class="py-2.5 px-3">Status</th><th class="py-2.5 px-3">${keyCol === '__position__' ? 'Row #' : escapeHtml(keyCol)}</th>` +
    allCols.filter(c => c !== keyCol).map(c => `<th class="py-2.5 px-3">${escapeHtml(c)}</th>`).join('');

  // Build matched pairs
  const pairs = [];
  if (keyCol === '__position__') {
    const max = Math.max(baseRows.length, targetRows.length);
    for (let i = 0; i < max; i++) pairs.push({ key: i + 1, b: baseRows[i], t: targetRows[i] });
  } else {
    // Key plus occurrence number, so a repeated key pairs 1st with 1st, 2nd with 2nd
    const withKeys = rows => {
      const seen = {};
      return rows.map(r => {
        if (isBlankValue(r[keyCol])) return { k: null, label: '(no key)', r };
        const label = String(r[keyCol]).trim();
        seen[label] = (seen[label] || 0) + 1;
        return { k: `${label}#${seen[label]}`, label, r };
      });
    };
    const baseK = withKeys(baseRows);
    const targetK = withKeys(targetRows);
    const tMap = new Map(targetK.filter(x => x.k).map(x => [x.k, x.r]));
    const used = new Set();
    baseK.forEach(({ k, label, r }) => {
      const match = k ? tMap.get(k) : undefined;
      if (match) used.add(k);
      pairs.push({ key: label, b: r, t: match });
    });
    targetK.forEach(({ k, label, r }) => { if (!k || !used.has(k)) pairs.push({ key: label, b: undefined, t: r }); });
  }

  let identical = 0, modified = 0, added = 0, removed = 0;
  const changeRows = [];
  const matchRows = [];
  const cols = allCols.filter(c => c !== keyCol);
  const cell = v => isBlankValue(v) ? '<span class="cell-empty">—</span>' : escapeHtml(v);

  pairs.forEach(({ key, b, t }) => {
    if (!b) {
      added++;
      changeRows.push(`<tr class="bg-emerald-50/60"><td class="py-2 px-3 text-emerald-700 font-semibold">+ Added</td><td class="py-2 px-3 font-semibold">${escapeHtml(key)}</td>${cols.map(c => `<td class="py-2 px-3 text-emerald-800">${cell(t[c])}</td>`).join('')}</tr>`);
    } else if (!t) {
      removed++;
      changeRows.push(`<tr class="bg-rose-50/60"><td class="py-2 px-3 text-rose-700 font-semibold">− Removed</td><td class="py-2 px-3 font-semibold">${escapeHtml(key)}</td>${cols.map(c => `<td class="py-2 px-3 text-rose-800 line-through">${cell(b[c])}</td>`).join('')}</tr>`);
    } else {
      let changed = false;
      const tds = cols.map(c => {
        if (!diffValuesEqual(b[c], t[c])) {
          changed = true;
          return `<td class="py-2 px-3 bg-amber-50 text-amber-900 font-semibold" title="Before: ${escapeHtml(b[c] ?? '')}">${cell(t[c])}<span class="block text-[10px] font-normal text-amber-700">was ${cell(b[c])}</span></td>`;
        }
        return `<td class="py-2 px-3">${cell(b[c])}</td>`;
      }).join('');
      if (changed) {
        modified++;
        changeRows.push(`<tr><td class="py-2 px-3 text-amber-700 font-semibold">Δ Modified</td><td class="py-2 px-3 font-semibold">${escapeHtml(key)}</td>${tds}</tr>`);
      } else {
        identical++;
        matchRows.push(`<tr class="text-slate-500"><td class="py-2 px-3 text-slate-400">Match</td><td class="py-2 px-3">${escapeHtml(key)}</td>${tds}</tr>`);
      }
    }
  });

  const html = changeRows.join('') + (onlyChanges ? '' : matchRows.join(''));
  tbody.innerHTML = html || `<tr><td colspan="${allCols.length + 2}" class="py-6 text-center text-slate-400">No differences found.</td></tr>`;

  document.getElementById('diffIdenticalCount').textContent = identical;
  document.getElementById('diffModifiedCount').textContent = modified;
  document.getElementById('diffVarianceCount').textContent = `+${added} / −${removed}`;
  showToast(`Compared on ${keyCol === '__position__' ? 'row position' : keyCol}: ${modified} modified, ${added} added, ${removed} removed`);
}

// -------------------------------------------------------------
// 🚨 12. Statistical Anomaly & Outlier Detector (Z-Score + IQR)
// -------------------------------------------------------------
function populateAnomalyDropdown() {
  const select = document.getElementById('anomalyColSelect');
  if (!select) return;

  select.innerHTML = '';
  const numCols = AppState.columns.filter(c => AppState.colMeta[c]?.type === 'numeric');
  (numCols.length > 0 ? numCols : AppState.columns).forEach(c => {
    select.innerHTML += `<option value="${c}">${c}</option>`;
  });

  if (numCols.length > 0) runAnomalyDetection();
}

// Isolation Forest (1-D) with a fixed seed so scores do not change between runs
function computeIsolationForestScores(values, numTrees = 100) {
  const n = values.length;
  if (n <= 2) return values.map(() => 0.5);
  const sampleSize = Math.min(256, n);
  const cOf = m => m <= 1 ? 0 : 2 * (Math.log(m - 1) + 0.5772156649) - (2 * (m - 1) / m);
  const c = cOf(sampleSize);
  const maxDepth = Math.ceil(Math.log2(sampleSize));
  const rnd = mlRng(1234);

  const pathLength = (v, sample, depth) => {
    if (depth >= maxDepth || sample.length <= 1) return depth + cOf(sample.length);
    const lo = Math.min(...sample), hi = Math.max(...sample);
    if (hi - lo <= 0) return depth + cOf(sample.length);
    const split = lo + rnd() * (hi - lo);
    const side = v < split ? sample.filter(s => s < split) : sample.filter(s => s >= split);
    return pathLength(v, side, depth + 1);
  };

  const totals = new Array(n).fill(0);
  for (let t = 0; t < numTrees; t++) {
    const sample = mlShuffle(values, 99 + t).slice(0, sampleSize);
    values.forEach((v, i) => { totals[i] += pathLength(v, sample, 0); });
  }
  return totals.map(sum => Number(Math.pow(2, -(sum / numTrees) / c).toFixed(3)));
}

function quantileSorted(sorted, q) {
  const pos = (sorted.length - 1) * q;
  const lo = Math.floor(pos), hi = Math.ceil(pos);
  return sorted[lo] + (sorted[hi] - sorted[lo]) * (pos - lo);
}

function computeOutlierBounds(nums) {
  const sorted = [...nums].sort((a, b) => a - b);
  const q1 = quantileSorted(sorted, 0.25);
  const q3 = quantileSorted(sorted, 0.75);
  const iqr = q3 - q1;
  return { q1, q3, iqr, lower: q1 - 1.5 * iqr, upper: q3 + 1.5 * iqr };
}

function runAnomalyDetection() {
  const col = document.getElementById('anomalyColSelect')?.value;
  if (!col) return;

  const method = document.getElementById('anomalyAlgorithmSelect')?.value || 'ensemble';
  const contamination = Number(document.getElementById('anomalyContaminationSelect')?.value || 0.05);

  // Blank cells are skipped, never treated as 0
  const vals = [];
  AppState.currentRows.forEach((r, idx) => { const v = toNumberOrNull(r[col]); if (v !== null) vals.push({ val: v, rowIdx: idx, row: r }); });
  const tbody = document.getElementById('anomalyTableBody');
  if (vals.length < 4) {
    document.getElementById('anomalyCount').textContent = '0';
    if (tbody) tbody.innerHTML = `<tr><td colspan="7" class="py-6 text-center text-slate-500">Need at least 4 numeric values in ${escapeHtml(col)}.</td></tr>`;
    return;
  }

  const raw = vals.map(v => v.val);
  const n = raw.length;
  const mean = raw.reduce((a, b) => a + b, 0) / n;
  const stdDev = Math.sqrt(raw.reduce((a, b) => a + (b - mean) ** 2, 0) / (n - 1)) || 1;
  const { lower: lowerBound, upper: upperBound } = computeOutlierBounds(raw);

  const iForestScores = computeIsolationForestScores(raw);
  const sortedScores = [...iForestScores].sort((a, b) => a - b);
  // Flag the top `contamination` share of scores, and only if they look anomalous (> 0.55)
  const scoreThreshold = Math.max(0.55, quantileSorted(sortedScores, 1 - contamination));

  document.getElementById('anomalyMean').textContent = formatNumber(mean, 1);
  document.getElementById('anomalyStdDev').textContent = formatNumber(stdDev, 1);
  document.getElementById('anomalyIqrRange').textContent = `[${formatNumber(lowerBound, 1)} to ${formatNumber(upperBound, 1)}]`;
  const iForestBadge = document.getElementById('anomalyIForestScore');
  if (iForestBadge) iForestBadge.textContent = `Peak score: ${Math.max(...iForestScores).toFixed(2)} · cut-off ${scoreThreshold.toFixed(2)}`;

  const flagged = [];
  vals.forEach((item, i) => {
    const zScore = (item.val - mean) / stdDev;
    const ifScore = iForestScores[i];
    const votes = {
      zscore: Math.abs(zScore) >= 2.5,
      iqr: item.val < lowerBound || item.val > upperBound,
      iforest: ifScore >= scoreThreshold
    };
    const voteCount = Object.values(votes).filter(Boolean).length;
    let shouldFlag = false;
    let devType = '';
    if (method === 'ensemble') {
      shouldFlag = voteCount >= 2;
      devType = `${voteCount}/3 methods agree`;
    } else {
      shouldFlag = votes[method];
      devType = method === 'zscore' ? `|Z| = ${Math.abs(zScore).toFixed(1)}` : method === 'iqr' ? (item.val > upperBound ? 'Above upper fence' : 'Below lower fence') : `Score ${ifScore.toFixed(2)}`;
    }
    if (shouldFlag) flagged.push({ ...item, ifScore, zScore: zScore.toFixed(2), type: devType });
  });

  document.getElementById('anomalyCount').textContent = flagged.length;
  if (flagged.length === 0) {
    tbody.innerHTML = `<tr><td colspan="7" class="py-6 text-center text-emerald-700 font-sans font-semibold">No outliers found in ${escapeHtml(col)} with this method.</td></tr>`;
    return;
  }

  const labelCol = AppState.columns.find(c => c !== col && AppState.colMeta[c]?.type !== 'numeric') || AppState.columns.find(c => c !== col) || col;
  tbody.innerHTML = flagged.map(f => {
    const isSevere = Math.abs(Number(f.zScore)) > 3 || f.ifScore > 0.7;
    const capTo = f.val > upperBound ? upperBound : lowerBound;
    return `
      <tr class="hover:bg-slate-50 transition border-b border-slate-100">
        <td class="py-2.5 px-3 font-mono text-slate-500">#${f.rowIdx + 1}</td>
        <td class="py-2.5 px-3 font-semibold text-slate-800">${isBlankValue(f.row[labelCol]) ? '—' : escapeHtml(f.row[labelCol])}</td>
        <td class="py-2.5 px-3 font-bold text-rose-600">${formatNumber(f.val)}</td>
        <td class="py-2.5 px-3 font-mono"><span class="px-2 py-0.5 rounded-md text-[10px] font-bold ${isSevere ? 'bg-rose-100 text-rose-800' : 'bg-amber-100 text-amber-800'}">${f.ifScore.toFixed(2)} ${isSevere ? '· severe' : '· elevated'}</span></td>
        <td class="py-2.5 px-3 font-mono text-indigo-600 font-bold">${f.zScore} σ</td>
        <td class="py-2.5 px-3"><span class="px-2 py-0.5 rounded-full text-[10px] font-semibold bg-slate-100 text-slate-700">${escapeHtml(f.type)}</span></td>
        <td class="py-2.5 px-3 text-right">
          <div class="flex items-center justify-end gap-1.5">
            <button onclick="isolateOutlier(${f.rowIdx})" class="text-xs text-indigo-600 hover:text-indigo-800 font-semibold px-2 py-1 rounded hover:bg-indigo-50">Inspect</button>
            <button onclick="clampSingleOutlier(${f.rowIdx}, ${escapeHtml(JSON.stringify(col))}, ${capTo})" title="Cap at the IQR fence (${formatNumber(capTo, 1)})" class="text-xs text-amber-600 hover:text-amber-800 font-semibold px-2 py-1 rounded hover:bg-amber-50">Cap</button>
          </div>
        </td>
      </tr>`;
  }).join('');
}

function clampSingleOutlier(rowIdx, col, targetVal) {
  if (AppState.currentRows[rowIdx]) {
    AppState.currentRows[rowIdx][col] = Math.round(targetVal);
    refreshPostTransformation(`Row #${rowIdx + 1}: ${col} capped at ${Math.round(targetVal)}`);
  }
}

function clampOutliers() {
  const col = document.getElementById('anomalyColSelect')?.value;
  if (!col) return;
  const nums = numericValues(AppState.currentRows, col);
  if (nums.length < 4) return;
  const { lower, upper } = computeOutlierBounds(nums);

  let clampedCount = 0;
  AppState.currentRows.forEach(r => {
    const v = toNumberOrNull(r[col]);
    if (v === null) return;
    if (v > upper) { r[col] = Math.round(upper); clampedCount++; }
    else if (v < lower) { r[col] = Math.round(lower); clampedCount++; }
  });
  refreshPostTransformation(clampedCount ? `Capped ${clampedCount} value(s) in ${col} to the IQR fences` : `Nothing to cap in ${col}`);
}

// -------------------------------------------------------------
// 🧼 Smart Data Cleaning & Profiling Pipeline (AIML Feature)
// -------------------------------------------------------------
function countCaseVariants(rows) {
  let variants = 0;
  AppState.columns.filter(c => AppState.colMeta[c]?.type !== 'numeric').forEach(col => {
    const forms = {};
    rows.forEach(r => {
      if (typeof r[col] !== 'string' || isBlankValue(r[col])) return;
      const k = r[col].trim().toLowerCase();
      forms[k] = forms[k] || new Set();
      forms[k].add(r[col]);
    });
    Object.values(forms).forEach(s => { if (s.size > 1) variants += s.size - 1; });
  });
  return variants;
}

function updateSmartHygieneMetrics() {
  const rows = AppState.currentRows || [];
  const set = (id, text) => { const el = document.getElementById(id); if (el) el.textContent = text; };
  if (rows.length === 0) return;

  let missing = 0;
  rows.forEach(r => AppState.columns.forEach(c => { if (isBlankValue(r[c])) missing++; }));
  const seen = new Set();
  let duplicates = 0;
  rows.forEach(r => { const k = rowSignature(r); if (seen.has(k)) duplicates++; else seen.add(k); });
  const emptyRows = rows.filter(r => AppState.columns.every(c => isBlankValue(r[c]))).length;
  const variants = countCaseVariants(rows);

  const totalCells = (rows.length * AppState.columns.length) || 1;
  const health = Math.max(0, Math.min(100, Math.round(100 - (missing / totalCells) * 100 - (duplicates / rows.length) * 50 - (variants / rows.length) * 20)));

  set('dataHealthScoreBadge', `Data Health Score: ${health}%`);
  set('hygieneMissingCount', `${missing} missing cell${missing === 1 ? '' : 's'}`);
  set('hygieneDupCount', `${duplicates} duplicate${duplicates === 1 ? '' : 's'}`);
  set('hygieneVariantCount', `${variants} case variant${variants === 1 ? '' : 's'} · ${emptyRows} empty row${emptyRows === 1 ? '' : 's'}`);
}

function runSmartDataCleaningPipeline() {
  let rows = AppState.currentRows;
  if (!rows || rows.length === 0) return;
  const before = rows.length;

  // 1. Drop rows with no data at all (imputing them would invent fake records)
  rows = rows.filter(r => AppState.columns.some(c => !isBlankValue(r[c])));
  const emptyDropped = before - rows.length;

  // 2. Trim spaces and merge case variants
  const { merged } = standardizeTextValues(rows);

  // 3. Deduplicate
  const seen = new Set();
  const deduped = [];
  rows.forEach(r => { const k = rowSignature(r); if (!seen.has(k)) { seen.add(k); deduped.push(r); } });
  const dupCount = rows.length - deduped.length;

  // 4. Impute remaining gaps (median for numbers, most common value for text); identifiers are left alone
  let imputedCount = 0;
  AppState.columns.forEach(c => {
    const meta = AppState.colMeta[c];
    const isId = meta?.type !== 'numeric' && meta?.distinctCount >= deduped.length * 0.9;
    if (isId) return;
    let fill;
    if (meta?.type === 'numeric') {
      const nums = numericValues(deduped, c).sort((a, b) => a - b);
      if (!nums.length) return;
      fill = aggregateNumbers(nums, 'median');
    } else {
      const freq = {};
      deduped.forEach(r => { if (!isBlankValue(r[c])) freq[r[c]] = (freq[r[c]] || 0) + 1; });
      fill = Object.keys(freq).sort((a, b) => freq[b] - freq[a])[0];
      if (fill === undefined) return;
    }
    deduped.forEach(r => { if (isBlankValue(r[c])) { r[c] = fill; imputedCount++; } });
  });

  // 5. Cap only extreme values (Tukey's outer fences, 3 × IQR), so ordinary spread is left alone
  let cappedCount = 0;
  AppState.columns.filter(c => AppState.colMeta[c]?.type === 'numeric').forEach(c => {
    const nums = numericValues(deduped, c);
    if (nums.length < 8) return;
    const b = computeOutlierBounds(nums);
    const lower = b.q1 - 3 * b.iqr;
    const upper = b.q3 + 3 * b.iqr;
    if (!(upper > lower)) return;
    deduped.forEach(r => {
      const v = toNumberOrNull(r[c]);
      if (v === null) return;
      if (v > upper) { r[c] = Math.round(upper * 100) / 100; cappedCount++; }
      else if (v < lower) { r[c] = Math.round(lower * 100) / 100; cappedCount++; }
    });
  });

  AppState.currentRows = deduped;
  AppState.filteredRows = [...deduped];
  refreshPostTransformation(`Cleaned: ${emptyDropped} empty row(s) dropped, ${merged} case variant(s) merged, ${dupCount} duplicate(s) removed, ${imputedCount} cell(s) filled, ${cappedCount} value(s) capped`);
  populateAutoMlSelectors();
}

function isolateOutlier(rowIdx) {
  switchTab('grid');
  AppState.grid.page = Math.floor(rowIdx / AppState.grid.pageSize) + 1;
  renderGridData();
  showToast(`Navigated to flagged record #${rowIdx + 1}`);
}

// -------------------------------------------------------------
// 🗣️ 13. AI NLQ Assistant Drawer (Chat with your Data)
// -------------------------------------------------------------
function toggleAiChatDrawer() {
  const drawer = document.getElementById('aiChatDrawer');
  const backdrop = document.getElementById('aiChatBackdrop');
  if (drawer) drawer.classList.toggle('hidden');
  if (backdrop) backdrop.classList.toggle('hidden');
}

function askPresetQuestion(text) {
  document.getElementById('aiChatInput').value = text.replace(/^[^A-Za-z"]+/, '').replace(/^"|"$/g, '');
  handleSendAiChatMessage(new Event('submit'));
}

function formatAiMarkdown(text) {
  if (!text) return '';
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
    .replace(/\*(.*?)\*/g, '<em>$1</em>')
    .replace(/`([^`]+)`/g, '<code class="px-1 py-0.5 rounded bg-slate-100 font-mono text-indigo-700 text-[11px]">$1</code>')
    .replace(/^\s*[\*\-]\s+(.*)$/gm, '<li class="ml-4 list-disc">$1</li>')
    .replace(/\n\n+/g, '<br/><br/>')
    .replace(/\n/g, '<br/>');
}

async function handleSendAiChatMessage(e) {
  if (e && e.preventDefault) e.preventDefault();
  const input = document.getElementById('aiChatInput');
  const question = input.value.trim();
  if (!question) return;

  const container = document.getElementById('aiChatMessages');

  // Append User Bubble
  container.innerHTML += `
    <div class="flex justify-end">
      <div class="p-3 rounded-2xl bg-indigo-600 text-white max-w-[85%] shadow-xs leading-relaxed text-xs">
        ${escapeHtml(question)}
      </div>
    </div>
  `;

  input.value = '';
  container.scrollTop = container.scrollHeight;

  // Append Thinking Indicator
  const thinkingId = 'thinking_' + Date.now();
  container.innerHTML += `
    <div id="${thinkingId}" class="flex items-center gap-2 text-slate-500 py-2 text-xs">
      <i data-lucide="loader-2" class="w-4 h-4 animate-spin text-indigo-600"></i>
      <span>Thinking about your data…</span>
    </div>
  `;
  lucide.createIcons();
  container.scrollTop = container.scrollHeight;

  try {
    const headers = { 'Content-Type': 'application/json' };
    if (AppState.auth.token) {
      headers['Authorization'] = `Bearer ${AppState.auth.token}`;
    }

    const res = await fetch(`${AppState.apiBaseUrl}/ai/nlq`, {
      method: 'POST',
      headers,
      body: JSON.stringify({
        question,
        fileName: AppState.fileName,
        columns: AppState.columns,
        columnStats: AppState.colMeta,
        sampleRows: (AppState.currentRows || []).slice(0, 15),
        totalRows: (AppState.currentRows || []).length,
        aggregates: buildDataAggregates(),
        lastInsight: AppState.lastAiInsight || ''
      })
    });

    const data = await res.json();
    document.getElementById(thinkingId)?.remove();

    if (!data.success && data.message) {
      throw new Error(data.message);
    }

    let chartActionHtml = '';
    if (data.suggestedChart && data.suggestedChart.x && data.suggestedChart.y) {
      chartActionHtml = `
        <div class="mt-2 pt-2 border-t border-slate-100 flex items-center justify-between">
          <span class="text-[10px] text-slate-500">Suggested: ${escapeHtml(data.suggestedChart.type || 'bar')} (${escapeHtml(data.suggestedChart.y)} by ${escapeHtml(data.suggestedChart.x)})</span>
          <button onclick='applyAiChart(${escapeHtml(JSON.stringify(data.suggestedChart.type || "bar"))}, ${escapeHtml(JSON.stringify(data.suggestedChart.x))}, ${escapeHtml(JSON.stringify(data.suggestedChart.y))}, ${escapeHtml(JSON.stringify(data.suggestedChart.agg || "sum"))})'
            class="px-2 py-1 rounded bg-indigo-50 text-indigo-600 text-[10px] font-bold hover:bg-indigo-100 transition">
            Plot in Studio
          </button>
        </div>
      `;
    }

    const answerText = extractAiAnswer(data.answer || 'No answer returned.');
    const formattedAnswer = renderMarkdownLite(answerText) + (data.truncated ? '<p class="md-p text-amber-700 text-[11px]">The answer hit the length limit. Ask a narrower question for the rest.</p>' : '');
    const sourceLabel = data.source === 'smart-data-engine' ? 'Offline data engine' : (data.source ? `Model: ${escapeHtml(data.source)}` : 'AI assistant');

    container.innerHTML += `
      <div class="p-3.5 rounded-2xl bg-white border border-slate-200 text-slate-800 shadow-xs space-y-2">
        <div class="flex items-center justify-between font-bold text-indigo-600 text-[11px]">
          <div class="flex items-center gap-1.5">
            <i data-lucide="sparkles" class="w-3.5 h-3.5"></i>
            <span>${sourceLabel}</span>
          </div>
          ${!AppState.auth.token ? '<span class="text-[9px] px-1.5 py-0.5 rounded bg-amber-50 text-amber-700 font-semibold border border-amber-200">Guest</span>' : ''}
        </div>
        <div class="ai-md leading-relaxed text-xs text-slate-700">${formattedAnswer}</div>
        ${chartActionHtml}
      </div>
    `;

    lucide.createIcons();
    container.scrollTop = container.scrollHeight;
  } catch (err) {
    document.getElementById(thinkingId)?.remove();
    container.innerHTML += `
      <div class="p-3 rounded-2xl bg-rose-50 border border-rose-200 text-rose-700 text-xs">
        Could not reach the AI service (${escapeHtml(err.message)}). The free server may be waking up; try again in a few seconds.
      </div>
    `;
    container.scrollTop = container.scrollHeight;
  }
}

function applyAiChart(type, x, y, agg) {
  switchTab('charts');
  const typeSel = document.getElementById('studioChartType');
  if (type && [...typeSel.options].some(o => o.value === type)) typeSel.value = type;
  if (x) document.getElementById('studioXCol').value = x;
  if (y) document.getElementById('studioYCol').value = y;
  if (agg) document.getElementById('studioAgg').value = agg;
  renderStudioChart();
  toggleAiChatDrawer();
  showToast(`Auto-plotted ${type || 'bar'} chart for ${y} by ${x}`);
}

// Mobile Responsive Window Resize & Orientation Change Listener
window.addEventListener('resize', () => {
  if (AppState.studioChartInstance) {
    AppState.studioChartInstance.resize();
  }
  if (AppState.categoryChartInstance) {
    AppState.categoryChartInstance.resize();
  }
  if (AppState.threeJs.renderer && AppState.threeJs.camera) {
    const container = document.getElementById('threeJsCanvasContainer');
    if (container && container.clientWidth) {
      const width = container.clientWidth;
      const height = container.clientHeight || 280;
      AppState.threeJs.camera.aspect = width / height;
      AppState.threeJs.camera.updateProjectionMatrix();
      AppState.threeJs.renderer.setSize(width, height);
    }
  }
  resizeAutoMlCharts();
});

// ==============================================================================
// 🤖 14. ENTERPRISE AUTOML & PREDICTIVE ANALYTICS STUDIO (AIML MINOR PROJECT)
// ==============================================================================

function resizeAutoMlCharts() {
  if (AppState.automl && AppState.automl.charts) {
    Object.values(AppState.automl.charts).forEach(ch => {
      if (ch && typeof ch.resize === 'function') ch.resize();
    });
  }
}

// ---------------------------------------------------------------------------
// Configuration wizard
// ---------------------------------------------------------------------------
const ML_METRICS = {
  regression: [
    { value: 'mae', label: 'MAE · Mean Absolute Error', short: 'MAE', higherIsBetter: false },
    { value: 'rmse', label: 'RMSE · Root Mean Squared Error', short: 'RMSE', higherIsBetter: false },
    { value: 'r2', label: 'R² · Variance Explained', short: 'R²', higherIsBetter: true },
    { value: 'mape', label: 'MAPE · Mean Absolute % Error', short: 'MAPE', higherIsBetter: false }
  ],
  classification: [
    { value: 'f1', label: 'F1 · Macro Average', short: 'F1', higherIsBetter: true },
    { value: 'accuracy', label: 'Accuracy', short: 'Accuracy', higherIsBetter: true },
    { value: 'precision', label: 'Precision · Macro', short: 'Precision', higherIsBetter: true },
    { value: 'recall', label: 'Recall · Macro', short: 'Recall', higherIsBetter: true }
  ],
  forecasting: [
    { value: 'mae', label: 'MAE · Mean Absolute Error', short: 'MAE', higherIsBetter: false },
    { value: 'rmse', label: 'RMSE · Root Mean Squared Error', short: 'RMSE', higherIsBetter: false },
    { value: 'mape', label: 'MAPE · Mean Absolute % Error', short: 'MAPE', higherIsBetter: false }
  ]
};

function mlMetricInfo(taskType, metric) {
  const list = ML_METRICS[taskType] || ML_METRICS.regression;
  return list.find(m => m.value === metric) || list[0];
}

function mlFormatMetric(taskType, metric, value) {
  if (value === null || value === undefined || !Number.isFinite(value)) return '–';
  if (taskType === 'classification') return `${value.toFixed(1)}%`;
  if (metric === 'r2') return value.toFixed(3);
  if (metric === 'mape') return `${value.toFixed(1)}%`;
  return value.toLocaleString(undefined, { maximumFractionDigits: 2 });
}

// Columns that should not be used as features by default: identifiers and dates.
function mlColumnRole(col) {
  const meta = AppState.colMeta[col] || {};
  const rows = AppState.currentRows || [];
  const filled = meta.nonNull || 0;
  const name = col.toLowerCase();
  if (/(date|time|period|month|year|timestamp|day)$/.test(name) || /^(date|time|period)/.test(name)) return 'date';
  if (meta.type !== 'numeric' && filled >= 5 && meta.distinctCount >= filled * 0.9) return 'id';
  if (/(^id$|_id$|id$|^code$|uuid|serial)/.test(name) && meta.distinctCount >= filled * 0.9) return 'id';
  if (meta.type !== 'numeric') {
    const sample = rows.slice(0, 30).map(r => r[col]).filter(v => !isBlankValue(v));
    const dateLike = sample.filter(v => /^\d{4}-\d{1,2}(-\d{1,2})?/.test(String(v).trim())).length;
    if (sample.length > 0 && dateLike / sample.length > 0.8) return 'date';
  }
  return 'feature';
}

function populateAutoMlSelectors() {
  const targetSelect = document.getElementById('mlTargetColSelect');
  const dateSelect = document.getElementById('mlDateColSelect');
  const featContainer = document.getElementById('mlFeaturesContainer');
  if (!targetSelect || !featContainer) return;

  const cols = AppState.columns || [];
  if (cols.length === 0) return;

  const previousTarget = targetSelect.value;
  targetSelect.innerHTML = '';
  if (dateSelect) dateSelect.innerHTML = '';

  const numCols = cols.filter(c => AppState.colMeta[c]?.type === 'numeric');
  const defaultTarget = (previousTarget && cols.includes(previousTarget)) ? previousTarget : (numCols.find(c => {
    const cl = c.toLowerCase();
    return cl.includes('sales') || cl.includes('revenue') || cl.includes('profit') || cl.includes('price');
  }) || numCols[0] || cols[0]);

  cols.forEach(c => {
    const opt = document.createElement('option');
    opt.value = c;
    opt.textContent = `${c} (${AppState.colMeta[c]?.type === 'numeric' ? 'Numeric' : 'Categorical'})`;
    if (c === defaultTarget) opt.selected = true;
    targetSelect.appendChild(opt);
  });

  if (dateSelect) {
    const dateCandidates = cols.filter(c => mlColumnRole(c) === 'date');
    cols.forEach(c => {
      const dOpt = document.createElement('option');
      dOpt.value = c;
      dOpt.textContent = c;
      if (c === dateCandidates[0]) dOpt.selected = true;
      dateSelect.appendChild(dOpt);
    });
  }

  handleMlTargetChange(true);
}

function renderMlFeaturePills(targetCol) {
  const featContainer = document.getElementById('mlFeaturesContainer');
  if (!featContainer) return;
  featContainer.innerHTML = '';

  const cols = (AppState.columns || []).filter(c => c !== targetCol);
  cols.forEach(c => {
    const isNum = AppState.colMeta[c]?.type === 'numeric';
    const role = mlColumnRole(c);
    const checked = role === 'feature';
    const pill = document.createElement('label');
    pill.className = 'ml-pill inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-xs font-medium bg-white border border-slate-200 text-slate-700 cursor-pointer hover:bg-slate-100 transition select-none';
    pill.title = role === 'id' ? 'Looks like an identifier (unique per row). Excluded by default to avoid memorising row IDs.'
      : role === 'date' ? 'Looks like a date. Use Time-Series mode to model dates.' : '';
    pill.innerHTML = `
      <input type="checkbox" name="mlFeatureCheckbox" value="${escapeHtml(c)}" ${checked ? 'checked' : ''} class="rounded w-3.5 h-3.5 accent-indigo-600">
      <span>${escapeHtml(c)}</span>
      <span class="text-[9px] font-mono ${isNum ? 'text-indigo-600' : 'text-amber-600'}">${isNum ? '#' : 'abc'}</span>
      ${role !== 'feature' ? `<span class="text-[9px] font-semibold uppercase tracking-wide text-slate-400">${role}</span>` : ''}
    `;
    featContainer.appendChild(pill);
  });
}

function selectAllMlFeatures(checkAll = true) {
  document.querySelectorAll('input[name="mlFeatureCheckbox"]').forEach(cb => {
    cb.checked = checkAll;
  });
}

function setMlMetricOptions(taskType) {
  const metricSelect = document.getElementById('mlPrimaryMetricSelect');
  if (!metricSelect) return;
  const previous = metricSelect.value;
  const list = ML_METRICS[taskType] || ML_METRICS.regression;
  metricSelect.innerHTML = list.map(m => `<option value="${m.value}">${m.label}</option>`).join('');
  metricSelect.value = list.some(m => m.value === previous) ? previous : list[0].value;
}

function handleMlTargetChange(reRenderPills = true) {
  const targetCol = document.getElementById('mlTargetColSelect')?.value;
  if (!targetCol) return;

  const isNum = AppState.colMeta[targetCol]?.type === 'numeric';
  const taskSelect = document.getElementById('mlTaskTypeSelect');
  if (taskSelect && taskSelect.value !== 'forecasting') {
    taskSelect.value = isNum ? 'regression' : 'classification';
  }
  handleMlTaskChange();
  if (reRenderPills) renderMlFeaturePills(targetCol);
}

function handleMlTaskChange() {
  const task = document.getElementById('mlTaskTypeSelect')?.value || 'regression';
  const forecastContainer = document.getElementById('mlForecastingOptionsContainer');
  const featuresBlock = document.getElementById('mlFeaturesBlock');
  if (forecastContainer) forecastContainer.classList.toggle('hidden', task !== 'forecasting');
  if (featuresBlock) featuresBlock.classList.toggle('hidden', task === 'forecasting');
  setMlMetricOptions(task);
}

function loadMlScenario(scenarioKey) {
  const scenarios = {
    sales_regression: { dataset: 'sales', target: 'Profit', task: 'regression' },
    churn_classification: { dataset: 'churn', target: 'Churned', task: 'classification' },
    demand_forecast: { dataset: 'forecast', target: 'Revenue', task: 'forecasting', date: 'Period' }
  };
  const s = scenarios[scenarioKey];
  if (!s) return;
  loadSampleDataset(s.dataset);
  const targetSelect = document.getElementById('mlTargetColSelect');
  if (targetSelect) targetSelect.value = s.target;
  const taskSelect = document.getElementById('mlTaskTypeSelect');
  if (taskSelect) taskSelect.value = s.task;
  handleMlTaskChange();
  renderMlFeaturePills(s.target);
  if (s.date) {
    const dateSelect = document.getElementById('mlDateColSelect');
    if (dateSelect) dateSelect.value = s.date;
  }
  switchTab('automl');
  runAutoMlTraining();
}

// ---------------------------------------------------------------------------
// Numerical helpers (seeded, so every run gives the same result)
// ---------------------------------------------------------------------------
const ML_SEED = 42;

function mlRng(seed) {
  let s = (seed >>> 0) || 1;
  return () => {
    s ^= s << 13; s >>>= 0;
    s ^= s >>> 17;
    s ^= s << 5; s >>>= 0;
    return s / 4294967296;
  };
}

function mlShuffle(arr, seed) {
  const rnd = mlRng(seed);
  const a = [...arr];
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(rnd() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

function mlMean(a) { return a.length ? a.reduce((s, v) => s + v, 0) / a.length : 0; }

// Solve (XᵀX + αI)β = Xᵀy with Gaussian elimination and partial pivoting.
function mlSolveLinear(A, b) {
  const n = A.length;
  const M = A.map((row, i) => [...row, b[i]]);
  for (let col = 0; col < n; col++) {
    let pivot = col;
    for (let r = col + 1; r < n; r++) if (Math.abs(M[r][col]) > Math.abs(M[pivot][col])) pivot = r;
    [M[col], M[pivot]] = [M[pivot], M[col]];
    const p = M[col][col];
    if (Math.abs(p) < 1e-12) continue;
    for (let c = col; c <= n; c++) M[col][c] /= p;
    for (let r = 0; r < n; r++) {
      if (r === col) continue;
      const f = M[r][col];
      if (f === 0) continue;
      for (let c = col; c <= n; c++) M[r][c] -= f * M[col][c];
    }
  }
  return M.map(row => (Number.isFinite(row[n]) ? row[n] : 0));
}

// ---------------------------------------------------------------------------
// Preprocessing: fitted on training rows only (median impute, scale, one-hot)
// ---------------------------------------------------------------------------
function fitPreprocessorClient(trainRows, featureCols) {
  const colStats = {};
  featureCols.forEach(c => {
    const isNumeric = AppState.colMeta[c]?.type === 'numeric';
    if (isNumeric) {
      const nums = numericValues(trainRows, c).sort((a, b) => a - b);
      const median = nums.length ? nums[Math.floor(nums.length / 2)] : 0;
      const mean = mlMean(nums);
      const std = Math.sqrt(mlMean(nums.map(v => (v - mean) ** 2))) || 1;
      colStats[c] = { type: 'numeric', median, mean, std, min: nums[0] ?? 0, max: nums[nums.length - 1] ?? 0 };
    } else {
      const freq = {};
      trainRows.forEach(r => {
        if (!isBlankValue(r[c])) {
          const k = String(r[c]).trim();
          freq[k] = (freq[k] || 0) + 1;
        }
      });
      const categories = Object.keys(freq).sort((a, b) => freq[b] - freq[a]).slice(0, 15);
      colStats[c] = { type: 'categorical', categories, mode: categories[0] || '' };
    }
  });

  const featureNames = [];
  const featureSource = [];
  featureCols.forEach(c => {
    if (colStats[c].type === 'numeric') {
      featureNames.push(c);
      featureSource.push(c);
    } else {
      colStats[c].categories.forEach(cat => {
        featureNames.push(`${c} = ${cat}`);
        featureSource.push(c);
      });
    }
  });

  const vectorizeRow = (row) => {
    const vec = [];
    featureCols.forEach(c => {
      const stat = colStats[c];
      if (stat.type === 'numeric') {
        let val = toNumberOrNull(row[c]);
        if (val === null) val = stat.median;
        vec.push((val - stat.mean) / stat.std);
      } else {
        const catVal = isBlankValue(row[c]) ? stat.mode : String(row[c]).trim();
        stat.categories.forEach(cat => vec.push(catVal === cat ? 1 : 0));
      }
    });
    return vec;
  };

  return { featureNames, featureSource, colStats, vectorizeRow };
}

// ---------------------------------------------------------------------------
// Decision trees (histogram splits) shared by Random Forest & Gradient Boosting
// ---------------------------------------------------------------------------
function mlBinFeatures(X, maxBins = 24) {
  const p = X[0] ? X[0].length : 0;
  const thresholds = [];
  for (let j = 0; j < p; j++) {
    const vals = Array.from(new Set(X.map(r => r[j]))).sort((a, b) => a - b);
    if (vals.length <= 1) { thresholds.push([]); continue; }
    const mids = [];
    for (let i = 0; i < vals.length - 1; i++) mids.push((vals[i] + vals[i + 1]) / 2);
    if (mids.length <= maxBins) thresholds.push(mids);
    else {
      const step = mids.length / maxBins;
      const picked = [];
      for (let k = 0; k < maxBins; k++) picked.push(mids[Math.floor(k * step)]);
      thresholds.push(Array.from(new Set(picked)));
    }
  }
  return thresholds;
}

// Precompute, for every row and feature, which threshold bucket it falls in.
function mlPrepareBins(X, maxBins = 16) {
  const thresholds = mlBinFeatures(X, maxBins);
  const bins = X.map(row => row.map((v, f) => {
    const th = thresholds[f];
    let lo = 0, hi = th.length;
    while (lo < hi) { const mid = (lo + hi) >> 1; if (v <= th[mid]) hi = mid; else lo = mid + 1; }
    return lo;
  }));
  return { thresholds, bins };
}

// Histogram-based CART. mode 'reg' (variance) or 'clf' (Gini); target holds y values or class indices.
function mlBuildTree(binData, indices, target, opts, depth, rnd) {
  const { thresholds, bins } = binData;
  const { maxDepth, minLeaf, maxFeatures, mode, K } = opts;
  const D = mode === 'reg' ? 3 : K + 1;
  const add = (s, off, i) => {
    s[off]++;
    if (mode === 'reg') { s[off + 1] += target[i]; s[off + 2] += target[i] * target[i]; }
    else s[off + 1 + target[i]]++;
  };
  const imp = (s) => {
    if (s[0] <= 0) return 0;
    if (mode === 'reg') return s[2] - (s[1] * s[1]) / s[0];
    let q = 0; for (let k = 1; k <= K; k++) q += s[k] * s[k];
    return s[0] - q / s[0];
  };
  const leaf = (s) => mode === 'reg' ? s[1] / (s[0] || 1) : Array.from({ length: K }, (_, k) => s[1 + k] / (s[0] || 1));

  const total = new Float64Array(D);
  for (const i of indices) add(total, 0, i);
  const node = { leaf: true, value: leaf(total) };
  const parentImp = imp(total);
  if (depth >= maxDepth || indices.length < minLeaf * 2 || parentImp <= 1e-12) return node;

  const p = thresholds.length;
  let features = [...Array(p).keys()];
  if (maxFeatures && maxFeatures < p) {
    for (let i = p - 1; i > 0; i--) { const j = Math.floor(rnd() * (i + 1)); [features[i], features[j]] = [features[j], features[i]]; }
    features = features.slice(0, maxFeatures);
  }

  let best = null;
  const left = new Float64Array(D);
  const right = new Float64Array(D);
  for (const f of features) {
    const T = thresholds[f].length;
    if (!T) continue;
    const hist = new Float64Array((T + 1) * D);
    for (const i of indices) add(hist, bins[i][f] * D, i);
    left.fill(0);
    for (let b = 0; b < T; b++) {
      for (let d = 0; d < D; d++) left[d] += hist[b * D + d];
      for (let d = 0; d < D; d++) right[d] = total[d] - left[d];
      if (left[0] < minLeaf || right[0] < minLeaf) continue;
      const gain = parentImp - imp(left) - imp(right);
      if (!best || gain > best.gain) best = { f, b, gain };
    }
  }
  if (!best || best.gain <= 1e-9) return node;

  const L = [], R = [];
  for (const i of indices) (bins[i][best.f] <= best.b ? L : R).push(i);
  return {
    leaf: false,
    f: best.f,
    t: thresholds[best.f][best.b],
    left: mlBuildTree(binData, L, target, opts, depth + 1, rnd),
    right: mlBuildTree(binData, R, target, opts, depth + 1, rnd)
  };
}

function mlTreePredict(node, x) {
  while (!node.leaf) node = x[node.f] <= node.t ? node.left : node.right;
  return node.value;
}

// ---------------------------------------------------------------------------
// Model zoo. Each factory returns { fit(X, y) -> { predict(x), proba?(x) } }
// ---------------------------------------------------------------------------
function mlLinearModel(alpha) {
  return {
    fit(X, y) {
      const p = X[0] ? X[0].length : 0;
      const d = p + 1;
      const XtX = Array.from({ length: d }, () => Array(d).fill(0));
      const Xty = Array(d).fill(0);
      X.forEach((row, n) => {
        const xb = [1, ...row];
        for (let i = 0; i < d; i++) {
          Xty[i] += xb[i] * y[n];
          for (let j = i; j < d; j++) XtX[i][j] += xb[i] * xb[j];
        }
      });
      for (let i = 0; i < d; i++) for (let j = 0; j < i; j++) XtX[i][j] = XtX[j][i];
      // Intercept is not penalised; a tiny ridge keeps OLS stable with one-hot columns
      for (let i = 1; i < d; i++) XtX[i][i] += alpha;
      const beta = mlSolveLinear(XtX, Xty);
      return {
        beta,
        predict: x => beta[0] + x.reduce((s, v, i) => s + beta[i + 1] * v, 0)
      };
    }
  };
}

function mlRandomForestRegressor(nTrees = 40, maxDepth = 6) {
  return {
    fit(X, y) {
      const binData = mlPrepareBins(X);
      const rnd = mlRng(ML_SEED);
      const n = X.length;
      const maxFeatures = Math.max(1, Math.round(binData.thresholds.length / 3));
      const opts = { maxDepth, minLeaf: 2, maxFeatures, mode: 'reg' };
      const trees = [];
      for (let t = 0; t < nTrees; t++) {
        const bag = Array.from({ length: n }, () => Math.floor(rnd() * n));
        trees.push(mlBuildTree(binData, bag, y, opts, 0, rnd));
      }
      return { predict: x => mlMean(trees.map(tr => mlTreePredict(tr, x))) };
    }
  };
}

function mlGradientBoostingRegressor(nStages = 80, lr = 0.1, maxDepth = 3) {
  return {
    fit(X, y) {
      const binData = mlPrepareBins(X);
      const rnd = mlRng(ML_SEED + 7);
      const init = mlMean(y);
      const pred = y.map(() => init);
      const all = [...Array(X.length).keys()];
      const opts = { maxDepth, minLeaf: 2, maxFeatures: null, mode: 'reg' };
      const trees = [];
      for (let s = 0; s < nStages; s++) {
        const resid = y.map((v, i) => v - pred[i]);
        const tree = mlBuildTree(binData, all, resid, opts, 0, rnd);
        trees.push(tree);
        for (let i = 0; i < X.length; i++) pred[i] += lr * mlTreePredict(tree, X[i]);
      }
      return { predict: x => init + lr * trees.reduce((s, tr) => s + mlTreePredict(tr, x), 0) };
    }
  };
}

function mlLogisticRegression(l2 = 0.01, iters = 400, lr = 0.3) {
  return {
    fit(X, y, classes) {
      if (X.length > 1000) iters = 150;
      const p = X[0] ? X[0].length : 0;
      const K = classes.length;
      const W = Array.from({ length: K }, () => Array(p + 1).fill(0));
      const yIdx = y.map(v => classes.indexOf(v));
      const n = X.length || 1;
      const softmax = (x) => {
        const z = W.map(w => w[0] + x.reduce((s, v, i) => s + w[i + 1] * v, 0));
        const m = Math.max(...z);
        const e = z.map(v => Math.exp(v - m));
        const s = e.reduce((a, b) => a + b, 0);
        return e.map(v => v / s);
      };
      for (let it = 0; it < iters; it++) {
        const grad = Array.from({ length: K }, () => Array(p + 1).fill(0));
        X.forEach((x, idx) => {
          const pr = softmax(x);
          for (let k = 0; k < K; k++) {
            const err = pr[k] - (yIdx[idx] === k ? 1 : 0);
            grad[k][0] += err;
            for (let j = 0; j < p; j++) grad[k][j + 1] += err * x[j];
          }
        });
        for (let k = 0; k < K; k++) {
          W[k][0] -= lr * grad[k][0] / n;
          for (let j = 1; j <= p; j++) W[k][j] -= lr * (grad[k][j] / n + l2 * W[k][j]);
        }
      }
      return {
        proba: x => softmax(x),
        predict: x => { const pr = softmax(x); return classes[pr.indexOf(Math.max(...pr))]; }
      };
    }
  };
}

function mlRandomForestClassifier(nTrees = 40, maxDepth = 6) {
  return {
    fit(X, y, classes) {
      const binData = mlPrepareBins(X);
      const rnd = mlRng(ML_SEED + 3);
      const n = X.length;
      const maxFeatures = Math.max(1, Math.round(Math.sqrt(binData.thresholds.length)));
      const yIdx = y.map(v => classes.indexOf(v));
      const opts = { maxDepth, minLeaf: 1, maxFeatures, mode: 'clf', K: classes.length };
      const trees = [];
      for (let t = 0; t < nTrees; t++) {
        const bag = Array.from({ length: n }, () => Math.floor(rnd() * n));
        trees.push(mlBuildTree(binData, bag, yIdx, opts, 0, rnd));
      }
      const proba = x => {
        const acc = Array(classes.length).fill(0);
        trees.forEach(tr => mlTreePredict(tr, x).forEach((v, k) => { acc[k] += v; }));
        return acc.map(v => v / trees.length);
      };
      return { proba, predict: x => { const pr = proba(x); return classes[pr.indexOf(Math.max(...pr))]; } };
    }
  };
}

function mlKnnClassifier(k = 5) {
  return {
    fit(X, y, classes) {
      const proba = x => {
        const dists = X.map((row, i) => ({ d: Math.sqrt(row.reduce((s, v, j) => s + (v - x[j]) ** 2, 0)), c: y[i] }))
          .sort((a, b) => a.d - b.d).slice(0, Math.min(k, X.length));
        const acc = Array(classes.length).fill(0);
        dists.forEach(({ d, c }) => { acc[classes.indexOf(c)] += 1 / (d + 1e-6); });
        const s = acc.reduce((a, b) => a + b, 0) || 1;
        return acc.map(v => v / s);
      };
      return { proba, predict: x => { const pr = proba(x); return classes[pr.indexOf(Math.max(...pr))]; } };
    }
  };
}

// ---------------------------------------------------------------------------
// Metrics (sklearn-compatible definitions)
// ---------------------------------------------------------------------------
function computeRegMetrics(actuals, preds) {
  const n = actuals.length || 1;
  let abs = 0, sq = 0, ape = 0, apeN = 0;
  const meanAct = mlMean(actuals);
  let ssTot = 0;
  actuals.forEach((a, i) => {
    const d = a - preds[i];
    abs += Math.abs(d);
    sq += d * d;
    if (Math.abs(a) > 1e-9) { ape += Math.abs(d / a); apeN++; }
    ssTot += (a - meanAct) ** 2;
  });
  return {
    mae: abs / n,
    rmse: Math.sqrt(sq / n),
    r2: ssTot > 0 ? 1 - sq / ssTot : 0, // can be negative: worse than predicting the mean
    mape: apeN ? (ape / apeN) * 100 : null
  };
}

function computeClfMetrics(actuals, preds, classes) {
  const n = actuals.length || 1;
  let correct = 0;
  actuals.forEach((a, i) => { if (a === preds[i]) correct++; });
  let precSum = 0, recSum = 0, f1Sum = 0;
  classes.forEach(c => {
    let tp = 0, fp = 0, fn = 0;
    actuals.forEach((a, i) => {
      if (preds[i] === c && a === c) tp++;
      else if (preds[i] === c) fp++;
      else if (a === c) fn++;
    });
    // zero_division = 0: a class the model never predicts scores 0, not 100
    const prec = tp + fp > 0 ? tp / (tp + fp) : 0;
    const rec = tp + fn > 0 ? tp / (tp + fn) : 0;
    precSum += prec;
    recSum += rec;
    f1Sum += prec + rec > 0 ? (2 * prec * rec) / (prec + rec) : 0;
  });
  const k = classes.length || 1;
  return {
    accuracy: (correct / n) * 100,
    precision: (precSum / k) * 100,
    recall: (recSum / k) * 100,
    f1: (f1Sum / k) * 100
  };
}

function buildConfusionMatrixClient(actuals, preds, classes) {
  const matrix = {};
  classes.forEach(a => { matrix[a] = {}; classes.forEach(p => { matrix[a][p] = 0; }); });
  actuals.forEach((a, i) => { if (matrix[a] && matrix[a][preds[i]] !== undefined) matrix[a][preds[i]]++; });
  return { classes, matrix };
}

// ---------------------------------------------------------------------------
// Evaluation plan: 5-fold cross-validation for small data, 80/20 holdout for large
// ---------------------------------------------------------------------------
function mlEvaluationPlan(n) {
  const order = mlShuffle([...Array(n).keys()], ML_SEED);
  if (n < 400) {
    const k = Math.min(5, n);
    const folds = Array.from({ length: k }, () => []);
    order.forEach((idx, i) => folds[i % k].push(idx));
    return { method: 'cv', k, label: `${k}-fold cross-validation`, folds };
  }
  const cut = Math.floor(n * 0.8);
  return { method: 'holdout', k: 1, label: '80/20 holdout', folds: [order.slice(cut)] };
}

function mlRankModels(models, taskType, metric) {
  const info = mlMetricInfo(taskType, metric);
  const score = m => {
    const v = m.metrics[metric];
    if (v === null || v === undefined || !Number.isFinite(v)) return info.higherIsBetter ? -Infinity : Infinity;
    return v;
  };
  models.sort((a, b) => info.higherIsBetter ? score(b) - score(a) : score(a) - score(b));
  models.forEach((m, i) => { m.rank = i + 1; });
  return models;
}

// Permutation importance on the final model, aggregated back to original columns
function mlPermutationImportance(model, allRows, featureCols, vectorizeRow, yAll, taskType, classes) {
  // Score on at most 300 rows to keep it quick
  const pick = allRows.length > 300 ? mlShuffle([...allRows.keys()], ML_SEED + 5).slice(0, 300) : [...allRows.keys()];
  const rows = pick.map(i => allRows[i]);
  const y = pick.map(i => yAll[i]);
  const evalScore = (Xs) => {
    const preds = Xs.map(x => model.predict(x));
    if (taskType === 'classification') return computeClfMetrics(y, preds, classes).accuracy;
    return -computeRegMetrics(y, preds).rmse;
  };
  const baseX = rows.map(vectorizeRow);
  const base = evalScore(baseX);
  const raw = featureCols.map((col, ci) => {
    let drop = 0;
    for (let rep = 0; rep < 3; rep++) {
      const shuffledVals = mlShuffle(rows.map(r => r[col]), ML_SEED + 11 * (ci + 1) + rep);
      const Xp = rows.map((r, i) => vectorizeRow({ ...r, [col]: shuffledVals[i] }));
      drop += base - evalScore(Xp);
    }
    return { feature: col, weight: Math.max(0, drop / 3) };
  });
  const total = raw.reduce((s, r) => s + r.weight, 0);
  return raw
    .map(r => ({ feature: r.feature, importance: total > 0 ? Math.round((r.weight / total) * 100) : 0 }))
    .sort((a, b) => b.importance - a.importance);
}

// ---------------------------------------------------------------------------
// Tabular pipeline (regression & classification)
// ---------------------------------------------------------------------------
function executeTabularPipeline(allRows, targetCol, featureCols, taskType, metric) {
  // Rows without a target value cannot be used for training or scoring
  let rows = allRows.filter(r => taskType === 'classification' ? !isBlankValue(r[targetCol]) : toNumberOrNull(r[targetCol]) !== null);
  const droppedRows = allRows.length - rows.length;
  if (rows.length > 3000) rows = mlShuffle(rows, ML_SEED).slice(0, 3000);
  if (rows.length < 8) throw new Error(`Need at least 8 rows with a value in "${targetCol}" (found ${rows.length}).`);

  const yOf = r => taskType === 'classification' ? String(r[targetCol]).trim() : toNumberOrNull(r[targetCol]);
  const yAll = rows.map(yOf);
  const classes = taskType === 'classification' ? Array.from(new Set(yAll)).sort() : null;
  if (classes && classes.length < 2) throw new Error(`"${targetCol}" has only one class, so there is nothing to classify.`);
  if (classes && classes.length > 20) throw new Error(`"${targetCol}" has ${classes.length} distinct values. Pick a column with fewer classes or use Regression.`);

  const candidates = taskType === 'classification'
    ? [
      { algorithm: 'dummy', name: 'Baseline (Majority Class)', isBaseline: true, make: () => ({ fit: (X, y, cls) => {
        const counts = cls.map(c => y.filter(v => v === c).length);
        const top = cls[counts.indexOf(Math.max(...counts))];
        const prior = counts.map(c => c / (y.length || 1));
        return { predict: () => top, proba: () => prior };
      } }) },
      { algorithm: 'logistic_regression', name: 'Logistic Regression (Softmax)', make: () => mlLogisticRegression() },
      { algorithm: 'random_forest_clf', name: 'Random Forest Classifier', make: () => mlRandomForestClassifier() },
      { algorithm: 'knn', name: 'k-Nearest Neighbours (k=5)', make: () => mlKnnClassifier(5) }
    ]
    : [
      { algorithm: 'dummy', name: 'Baseline (Mean)', isBaseline: true, make: () => ({ fit: (X, y) => { const m = mlMean(y); return { predict: () => m }; } }) },
      { algorithm: 'linear_regression', name: 'Linear Regression (OLS)', make: () => mlLinearModel(1e-6) },
      { algorithm: 'ridge_regression', name: 'Ridge Regression (L2)', make: () => mlLinearModel(1.0) },
      { algorithm: 'random_forest', name: 'Random Forest Regressor', make: () => mlRandomForestRegressor() },
      { algorithm: 'gradient_boosting', name: 'Gradient Boosted Trees', make: () => mlGradientBoostingRegressor() }
    ];

  const plan = mlEvaluationPlan(rows.length);
  const oof = candidates.map(() => new Array(rows.length).fill(null));

  plan.folds.forEach(testIdx => {
    const testSet = new Set(testIdx);
    const trainIdx = rows.map((_, i) => i).filter(i => !testSet.has(i));
    const trainRows = trainIdx.map(i => rows[i]);
    const prep = fitPreprocessorClient(trainRows, featureCols);
    const Xtr = trainRows.map(prep.vectorizeRow);
    const ytr = trainIdx.map(i => yAll[i]);
    candidates.forEach((cand, ci) => {
      const model = cand.make().fit(Xtr, ytr, classes);
      testIdx.forEach(i => { oof[ci][i] = model.predict(prep.vectorizeRow(rows[i])); });
    });
  });

  const evalIdx = rows.map((_, i) => i).filter(i => oof[0][i] !== null);
  const yEval = evalIdx.map(i => yAll[i]);

  // Refit every candidate on all rows for inference
  const finalPrep = fitPreprocessorClient(rows, featureCols);
  const Xall = rows.map(finalPrep.vectorizeRow);

  const models = candidates.map((cand, ci) => {
    const preds = evalIdx.map(i => oof[ci][i]);
    const metrics = taskType === 'classification' ? computeClfMetrics(yEval, preds, classes) : computeRegMetrics(yEval, preds);
    return {
      algorithm: cand.algorithm,
      name: cand.name,
      isBaseline: !!cand.isBaseline,
      metrics,
      evalPreds: preds,
      fitted: cand.make().fit(Xall, yAll, classes)
    };
  });

  mlRankModels(models, taskType, metric);
  const bestModel = models[0];
  const featureImportance = bestModel.isBaseline ? [] :
    mlPermutationImportance(bestModel.fitted, rows, featureCols, finalPrep.vectorizeRow, yAll, taskType, classes);

  const bestCi = candidates.findIndex(c => c.algorithm === bestModel.algorithm);
  const evalPredictions = evalIdx.map((i, k) => {
    const actual = yAll[i];
    const predicted = oof[bestCi][i];
    return taskType === 'classification'
      ? { row: i + 1, actual, predicted, residual: actual === predicted ? 'correct' : 'wrong' }
      : { row: i + 1, actual, predicted: Number(predicted.toFixed(2)), residual: Number((actual - predicted).toFixed(2)) };
  });

  const targetNums = taskType === 'classification' ? [] : yAll;
  return {
    taskType,
    models,
    bestModel,
    featureImportance,
    testPredictions: evalPredictions,
    confusionMatrix: taskType === 'classification' ? buildConfusionMatrixClient(yEval, bestModel.evalPreds, classes) : null,
    classes,
    colStats: finalPrep.colStats,
    vectorizeRow: finalPrep.vectorizeRow,
    featureCols,
    evaluation: { method: plan.method, label: plan.label, evaluatedRows: evalIdx.length, trainingRows: rows.length, droppedRows },
    targetRange: targetNums.length ? { min: Math.min(...targetNums), max: Math.max(...targetNums) } : null
  };
}

// ---------------------------------------------------------------------------
// Time-series pipeline
// ---------------------------------------------------------------------------
function mlParseDate(v) {
  if (v instanceof Date) return v;
  if (typeof v === 'number' && v > 20000 && v < 80000) return new Date(Math.round((v - 25569) * 86400000)); // Excel serial
  const s = String(v ?? '').trim();
  if (!s) return null;
  if (/^\d{4}$/.test(s)) return new Date(`${s}-01-01T00:00:00Z`);
  if (/^\d{4}-\d{1,2}$/.test(s)) return new Date(`${s.length === 6 ? s.replace('-', '-0') : s}-01T00:00:00Z`);
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function mlFormatPeriod(d, unit) {
  const iso = d.toISOString();
  if (unit === 'year') return iso.slice(0, 4);
  if (unit === 'month') return iso.slice(0, 7);
  return iso.slice(0, 10);
}

function executeForecastingPipeline(rows, dateCol, valCol, horizon, metric) {
  if (!dateCol || dateCol === valCol) throw new Error('Choose a date column that is different from the target.');
  const byDate = new Map();
  let skipped = 0;
  rows.forEach(r => {
    const d = mlParseDate(r[dateCol]);
    const v = toNumberOrNull(r[valCol]);
    if (!d || v === null) { skipped++; return; }
    const key = d.getTime();
    byDate.set(key, (byDate.get(key) || 0) + v); // several rows on one date are summed
  });
  const series = Array.from(byDate.entries()).sort((a, b) => a[0] - b[0]).map(([t, v]) => ({ t, value: v }));
  const n = series.length;
  if (n < 6) throw new Error(`Forecasting needs at least 6 dated values in "${valCol}" (found ${n}).`);

  // Infer the spacing of the series
  const gaps = series.slice(1).map((s, i) => (s.t - series[i].t) / 86400000).sort((a, b) => a - b);
  const medGap = gaps[Math.floor(gaps.length / 2)];
  const unit = medGap >= 360 ? 'year' : medGap >= 27 ? 'month' : 'day';
  const stepDays = unit === 'day' ? Math.max(1, Math.round(medGap)) : null;
  const nextDate = (t, h) => {
    const d = new Date(t);
    if (unit === 'year') d.setUTCFullYear(d.getUTCFullYear() + h);
    else if (unit === 'month') d.setUTCMonth(d.getUTCMonth() + h);
    else d.setUTCDate(d.getUTCDate() + h * stepDays);
    return d;
  };
  series.forEach(s => { s.date = mlFormatPeriod(new Date(s.t), unit); });

  const values = series.map(s => s.value);
  const testLen = Math.max(2, Math.round(n * 0.2));
  const train = values.slice(0, n - testLen);
  const test = values.slice(n - testLen);

  const holtFit = (vals) => {
    let best = null;
    for (const a of [0.2, 0.4, 0.6, 0.8]) {
      for (const b of [0.05, 0.15, 0.3]) {
        let level = vals[0], trend = vals.length > 1 ? vals[1] - vals[0] : 0, sse = 0;
        for (let t = 1; t < vals.length; t++) {
          const f = level + trend;
          sse += (vals[t] - f) ** 2;
          const prev = level;
          level = a * vals[t] + (1 - a) * (level + trend);
          trend = b * (level - prev) + (1 - b) * trend;
        }
        if (!best || sse < best.sse) best = { sse, level, trend, a, b };
      }
    }
    return h => best.level + h * best.trend;
  };
  const forecasters = [
    { algorithm: 'naive', name: 'Naive (Last Value)', isBaseline: true, fit: vals => { const last = vals[vals.length - 1]; return () => last; } },
    { algorithm: 'moving_average', name: 'Moving Average (3)', fit: vals => { const m = mlMean(vals.slice(-3)); return () => m; } },
    { algorithm: 'linear_trend', name: 'Linear Trend', fit: vals => {
      const ts = vals.map((_, i) => i), mt = mlMean(ts), mv = mlMean(vals);
      const slope = ts.reduce((s, t, i) => s + (t - mt) * (vals[i] - mv), 0) / (ts.reduce((s, t) => s + (t - mt) ** 2, 0) || 1);
      const icpt = mv - slope * mt;
      return h => icpt + slope * (vals.length - 1 + h);
    } },
    { algorithm: 'holt', name: "Holt's Exponential Smoothing", fit: holtFit }
  ];

  const nonNegative = values.every(v => v >= 0);
  const models = forecasters.map(fc => {
    const f = fc.fit(train);
    const preds = test.map((_, h) => f(h + 1));
    const m = computeRegMetrics(test, preds);
    return { algorithm: fc.algorithm, name: fc.name, isBaseline: !!fc.isBaseline, metrics: { mae: m.mae, rmse: m.rmse, mape: m.mape }, holdoutPreds: preds, fc };
  });
  mlRankModels(models, 'forecasting', metric);
  const best = models[0];

  // Refit the winner on the whole series and project forward
  const finalF = best.fc.fit(values);
  const sigma = best.metrics.rmse;
  const lastT = series[n - 1].t;
  const forecast = [];
  for (let h = 1; h <= horizon; h++) {
    let pred = finalF(h);
    if (nonNegative) pred = Math.max(0, pred);
    const margin = 1.96 * sigma * Math.sqrt(h);
    forecast.push({
      step: h,
      date: mlFormatPeriod(nextDate(lastT, h), unit),
      predicted: Number(pred.toFixed(2)),
      lowerBound: Number((nonNegative ? Math.max(0, pred - margin) : pred - margin).toFixed(2)),
      upperBound: Number((pred + margin).toFixed(2))
    });
  }

  const testDates = series.slice(n - testLen).map(s => s.date);
  const testPredictions = [
    ...test.map((a, i) => ({ type: 'holdout', date: testDates[i], actual: a, predicted: Number(best.holdoutPreds[i].toFixed(2)), lowerBound: '', upperBound: '' })),
    ...forecast.map(f => ({ type: 'forecast', date: f.date, actual: '', predicted: f.predicted, lowerBound: f.lowerBound, upperBound: f.upperBound }))
  ];

  return {
    taskType: 'forecasting',
    models,
    bestModel: best,
    historical: series.map(s => ({ date: s.date, value: s.value })),
    forecast,
    testPredictions,
    featureImportance: [],
    evaluation: { method: 'holdout', label: `last ${testLen} periods held out`, evaluatedRows: testLen, trainingRows: n, droppedRows: skipped, unit }
  };
}

// ---------------------------------------------------------------------------
// Master training entry point
// ---------------------------------------------------------------------------
async function runAutoMlTraining() {
  if (AppState.automl.isTraining) return;
  AppState.automl.isTraining = true;
  const btn = document.getElementById('mlRunTrainBtn');

  try {
    const targetCol = document.getElementById('mlTargetColSelect')?.value;
    if (!targetCol) { showToast('Please select a target variable to predict', true); return; }

    const taskType = document.getElementById('mlTaskTypeSelect')?.value || 'regression';
    const metric = document.getElementById('mlPrimaryMetricSelect')?.value || ML_METRICS[taskType][0].value;
    const featureCols = Array.from(document.querySelectorAll('input[name="mlFeatureCheckbox"]:checked')).map(b => b.value).filter(c => c !== targetCol);

    if (taskType !== 'forecasting' && featureCols.length === 0) {
      showToast('Please select at least one feature column for training', true);
      return;
    }
    if (taskType === 'regression' && AppState.colMeta[targetCol]?.type !== 'numeric') {
      showToast(`"${targetCol}" is not numeric. Use Classification for categorical targets.`, true);
      return;
    }

    const rows = AppState.currentRows || [];
    if (btn) {
      btn.disabled = true;
      btn.innerHTML = `<i data-lucide="loader" class="w-4 h-4 animate-spin"></i> <span>Training models…</span>`;
      lucide.createIcons();
    }
    await new Promise(r => setTimeout(r, 30)); // let the button repaint before the heavy work

    const start = performance.now();
    let result;
    if (taskType === 'forecasting') {
      const dateCol = document.getElementById('mlDateColSelect')?.value;
      const horizon = Number(document.getElementById('mlHorizonStepsSelect')?.value || 6);
      result = executeForecastingPipeline(rows, dateCol, targetCol, horizon, metric);
    } else {
      result = executeTabularPipeline(rows, targetCol, featureCols, taskType, metric);
    }
    const duration = Math.round(performance.now() - start);

    // Store the active pipeline BEFORE rendering, so every panel shows this run
    Object.assign(AppState.automl, {
      taskType,
      targetColumn: targetCol,
      primaryMetric: metric,
      models: result.models,
      bestModel: result.bestModel,
      vectorizeRow: result.vectorizeRow || null,
      colStats: result.colStats || {},
      featureCols: result.featureCols || [],
      testPredictions: result.testPredictions || [],
      featureImportance: result.featureImportance || [],
      evaluation: result.evaluation,
      classes: result.classes || null,
      targetRange: result.targetRange || null,
      forecast: result.forecast || null,
      durationMs: duration
    });

    renderAutoMlLeaderboard(result, taskType, metric, duration);
    renderAutoMlXaiChart(result.featureImportance || [], taskType);
    renderAutoMlModelComparisonChart(result.models, metric, taskType);
    renderAutoMlValidationDiagnostics(taskType, result);
    buildWhatIfSimulationInputs();

    const info = mlMetricInfo(taskType, metric);
    showToast(`AutoML complete: ${result.bestModel.name} (${info.short} ${mlFormatMetric(taskType, metric, result.bestModel.metrics[metric])})`);

    // Persist a summary for signed-in users (results are computed in the browser)
    if (AppState.apiBaseUrl && AppState.auth.token) {
      fetch(`${AppState.apiBaseUrl}/ml/train`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${AppState.auth.token}` },
        body: JSON.stringify({ clientResults: buildMlAuditPayload() })
      }).catch(() => {});
    }
  } catch (err) {
    console.error('AutoML Training Error:', err);
    showToast(`Training error: ${err.message}`, true);
  } finally {
    AppState.automl.isTraining = false;
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = `<i data-lucide="play" class="w-4 h-4 fill-white"></i> <span>Train AutoML Models</span>`;
      lucide.createIcons();
    }
  }
}

// ---------------------------------------------------------------------------
// Rendering
// ---------------------------------------------------------------------------
function renderAutoMlLeaderboard(result, taskType, metric) {
  const { models, bestModel, evaluation } = result;
  const info = mlMetricInfo(taskType, metric);
  const setText = (id, text) => { const el = document.getElementById(id); if (el) el.textContent = text; };

  setText('mlKpiWinner', bestModel?.name || '–');
  setText('mlKpiWinnerNote', bestModel?.isBaseline ? 'No model beat the baseline' : `Ranked by ${info.short}`);
  setText('mlKpiPrimaryScore', `${mlFormatMetric(taskType, metric, bestModel.metrics[metric])}`);
  setText('mlKpiPrimaryScoreLabel', `${info.short} · ${evaluation.label}`);

  const baseline = models.find(m => m.isBaseline);
  let lift = '–';
  if (baseline && bestModel !== baseline) {
    const b = baseline.metrics[metric], w = bestModel.metrics[metric];
    if (Number.isFinite(b) && Number.isFinite(w)) {
      if (metric === 'r2') lift = `${w - b >= 0 ? '+' : ''}${(w - b).toFixed(3)} R²`;
      else if (info.higherIsBetter) lift = `${w - b >= 0 ? '+' : ''}${(w - b).toFixed(1)} pts`;
      else lift = b > 0 ? `${Math.round(((b - w) / b) * 100)}% less error` : '–';
    }
  } else if (bestModel?.isBaseline) {
    lift = 'none';
  }
  setText('mlKpiBaselineLift', lift);
  setText('mlKpiLatency', `${result.durationMs ?? AppState.automl.durationMs ?? 0} ms`);
  setText('mlKpiTestRows', `${evaluation.evaluatedRows} ${taskType === 'forecasting' ? 'periods' : 'rows'}`);
  setText('mlKpiTestRowsNote', evaluation.method === 'cv' ? `every row scored once (${evaluation.label})` : evaluation.label);

  const subtitle = document.getElementById('mlLeaderboardSubtitle');
  if (subtitle) {
    let note = `Ranked by ${info.short} (${info.higherIsBetter ? 'higher' : 'lower'} is better) using ${evaluation.label}.`;
    if (evaluation.droppedRows) note += ` ${evaluation.droppedRows} row(s) without a usable value were skipped.`;
    if (evaluation.trainingRows < 30) note += ' Small dataset: treat scores as rough.';
    subtitle.textContent = note;
  }

  const thead = document.getElementById('mlLeaderboardHeadRow');
  const otherMetrics = (ML_METRICS[taskType] || []).filter(m => m.value !== metric);
  if (thead) {
    thead.innerHTML = `
      <th class="py-2.5 px-3">Rank</th>
      <th class="py-2.5 px-3">Model</th>
      <th class="py-2.5 px-3">${info.short}</th>
      ${otherMetrics.map(m => `<th class="py-2.5 px-3">${m.short}</th>`).join('')}
      <th class="py-2.5 px-3 text-right">Status</th>`;
  }

  const tbody = document.getElementById('mlLeaderboardTableBody');
  if (!tbody) return;
  tbody.innerHTML = models.map((m, idx) => {
    const isWinner = idx === 0;
    const rankBadge = isWinner
      ? `<span class="ml-rank-badge"><i data-lucide="crown" class="w-3 h-3"></i> #1</span>`
      : `<span class="font-semibold text-slate-500">#${idx + 1}</span>`;
    return `
      <tr class="${isWinner ? 'bg-indigo-50/60' : 'hover:bg-slate-50'} transition border-b border-slate-100">
        <td class="py-2.5 px-3 whitespace-nowrap">${rankBadge}</td>
        <td class="py-2.5 px-3">
          <div class="font-semibold text-slate-900 font-sans">${escapeHtml(m.name)}${m.isBaseline ? ' <span class="text-[10px] font-medium text-slate-400">baseline</span>' : ''}</div>
          <div class="text-[10px] text-slate-400">${m.algorithm}</div>
        </td>
        <td class="py-2.5 px-3 font-semibold ${isWinner ? 'text-indigo-700' : 'text-slate-900'}">${mlFormatMetric(taskType, metric, m.metrics[metric])}</td>
        ${otherMetrics.map(om => `<td class="py-2.5 px-3 text-slate-500">${mlFormatMetric(taskType, om.value, m.metrics[om.value])}</td>`).join('')}
        <td class="py-2.5 px-3 text-right">
          <span class="px-2.5 py-1 rounded-lg text-[10px] font-semibold ${isWinner ? 'bg-indigo-600 text-white' : 'bg-slate-100 text-slate-600'}">${isWinner ? 'Active model' : 'Benchmarked'}</span>
        </td>
      </tr>`;
  }).join('');
  lucide.createIcons();
}

function renderAutoMlXaiChart(featureImportance, taskType) {
  const canvas = document.getElementById('automlFeatureImportanceChart');
  const summary = document.getElementById('automlXaiSummary');
  if (!canvas) return;
  if (AppState.automl.charts.xai) { AppState.automl.charts.xai.destroy(); AppState.automl.charts.xai = null; }

  if (taskType === 'forecasting' || featureImportance.length === 0) {
    const ctx = canvas.getContext('2d');
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    if (summary) summary.innerHTML = taskType === 'forecasting'
      ? 'Feature importance does not apply to a single time series. The forecast uses only past values of the target.'
      : 'The winning model is the baseline, so no feature carries signal yet. Try different features or more rows.';
    return;
  }

  const top = featureImportance.slice(0, 8);
  AppState.automl.charts.xai = new Chart(canvas, {
    type: 'bar',
    data: {
      labels: top.map(f => f.feature),
      datasets: [{ label: 'Share of importance (%)', data: top.map(f => f.importance), backgroundColor: top.map((_, i) => i === 0 ? '#0d44e8' : '#8eadff'), borderRadius: 6 }]
    },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => ` ${ctx.parsed.x}% of total importance` } } },
      scales: { x: { grid: { color: '#f1f5f9' }, ticks: { callback: v => `${v}%` } }, y: { grid: { display: false } } }
    }
  });

  if (summary) {
    const [a, b] = featureImportance;
    summary.innerHTML = `<strong>${escapeHtml(a.feature)}</strong> matters most (${a.importance}% of total importance)` +
      (b ? `, followed by <strong>${escapeHtml(b.feature)}</strong> (${b.importance}%)` : '') +
      '. Measured by shuffling each column and checking how much the winning model gets worse.';
  }
}

function renderAutoMlModelComparisonChart(models, metric, taskType) {
  const canvas = document.getElementById('automlModelComparisonChart');
  if (!canvas) return;
  if (AppState.automl.charts.comparison) AppState.automl.charts.comparison.destroy();
  const info = mlMetricInfo(taskType, metric);
  const data = models.map(m => Number.isFinite(m.metrics[metric]) ? Number(m.metrics[metric].toFixed(metric === 'r2' ? 3 : 2)) : null);

  AppState.automl.charts.comparison = new Chart(canvas, {
    type: 'bar',
    data: {
      labels: models.map(m => m.name.split(' (')[0]),
      datasets: [{ label: info.short, data, backgroundColor: models.map((m, i) => i === 0 ? '#0d44e8' : (m.isBaseline ? '#cbd5e1' : '#8eadff')), borderRadius: 6 }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => ` ${info.short}: ${mlFormatMetric(taskType, metric, ctx.parsed.y)}` } } },
      scales: {
        y: { title: { display: true, text: `${info.short} (${info.higherIsBetter ? 'higher' : 'lower'} is better)`, font: { size: 10 } }, grid: { color: '#f1f5f9' } },
        x: { grid: { display: false }, ticks: { maxRotation: 25, minRotation: 0 } }
      }
    }
  });
}

function renderAutoMlValidationDiagnostics(taskType, payload) {
  const regContainer = document.getElementById('automlRegressionValidationContainer');
  const clfContainer = document.getElementById('automlClassificationValidationContainer');
  const foreContainer = document.getElementById('automlForecastingValidationContainer');
  const badge = document.getElementById('automlValidationModeBadge');
  [regContainer, clfContainer, foreContainer].forEach(el => el && el.classList.add('hidden'));
  if (badge) badge.textContent = payload.evaluation.label;

  if (taskType === 'forecasting') {
    if (foreContainer) foreContainer.classList.remove('hidden');
    const canvas = document.getElementById('automlForecastChart');
    if (!canvas) return;
    if (AppState.automl.charts.forecast) AppState.automl.charts.forecast.destroy();
    const hist = payload.historical, fut = payload.forecast;
    const holdout = payload.testPredictions.filter(p => p.type === 'holdout');
    const labels = [...hist.map(h => h.date), ...fut.map(f => f.date)];
    const pad = arr => [...hist.map(() => null), ...arr];
    const holdoutSeries = hist.map(h => { const m = holdout.find(p => p.date === h.date); return m ? m.predicted : null; });
    AppState.automl.charts.forecast = new Chart(canvas, {
      type: 'line',
      data: {
        labels,
        datasets: [
          { label: 'Actual', data: [...hist.map(h => h.value), ...fut.map(() => null)], borderColor: '#0f172a', backgroundColor: '#0f172a', borderWidth: 2, tension: 0.2, pointRadius: 2 },
          { label: 'Holdout prediction', data: [...holdoutSeries, ...fut.map(() => null)], borderColor: '#f59e0b', backgroundColor: '#f59e0b', borderDash: [3, 3], borderWidth: 2, pointRadius: 3 },
          { label: `Forecast (${payload.bestModel.name})`, data: pad(fut.map(f => f.predicted)), borderColor: '#0d44e8', backgroundColor: '#0d44e8', borderDash: [6, 4], borderWidth: 2.5, tension: 0.2 },
          { label: '95% range (upper)', data: pad(fut.map(f => f.upperBound)), borderColor: 'rgba(13,68,232,0.25)', backgroundColor: 'rgba(13,68,232,0.08)', borderWidth: 1, fill: '+1', pointRadius: 0 },
          { label: '95% range (lower)', data: pad(fut.map(f => f.lowerBound)), borderColor: 'rgba(13,68,232,0.25)', borderWidth: 1, pointRadius: 0 }
        ]
      },
      options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'top', labels: { boxWidth: 10 } } }, scales: { x: { grid: { display: false } }, y: { grid: { color: '#f1f5f9' } } } }
    });
    return;
  }

  if (taskType === 'classification') {
    if (clfContainer) clfContainer.classList.remove('hidden');
    const grid = document.getElementById('automlConfusionMatrixGrid');
    if (!grid || !payload.confusionMatrix) return;
    const { classes, matrix } = payload.confusionMatrix;
    grid.innerHTML = `
      <div class="text-xs font-semibold text-slate-700 mb-2">Confusion matrix · rows are actual, columns are predicted (${escapeHtml(payload.evaluation.label)})</div>
      <table class="text-xs border-collapse bg-white rounded-xl overflow-hidden">
        <thead><tr class="bg-slate-100 text-slate-700">
          <th class="p-2 border border-slate-200">Actual \\ Pred</th>
          ${classes.map(c => `<th class="p-2 border border-slate-200 text-indigo-700">${escapeHtml(c)}</th>`).join('')}
        </tr></thead>
        <tbody>
          ${classes.map(a => `<tr><td class="p-2 border border-slate-200 font-semibold bg-slate-50">${escapeHtml(a)}</td>${classes.map(p => {
            const count = matrix[a]?.[p] || 0;
            const cls = a === p ? 'bg-emerald-50 text-emerald-800' : (count > 0 ? 'bg-rose-50 text-rose-800' : 'text-slate-400');
            return `<td class="p-2 border border-slate-200 text-center font-semibold ${cls}">${count}</td>`;
          }).join('')}</tr>`).join('')}
        </tbody>
      </table>`;
    return;
  }

  if (regContainer) regContainer.classList.remove('hidden');
  const canvas = document.getElementById('automlValidationChart');
  if (!canvas) return;
  if (AppState.automl.charts.validation) AppState.automl.charts.validation.destroy();
  const preds = payload.testPredictions;
  const actuals = preds.map(p => p.actual);
  const lo = Math.min(...actuals, ...preds.map(p => p.predicted));
  const hi = Math.max(...actuals, ...preds.map(p => p.predicted));
  AppState.automl.charts.validation = new Chart(canvas, {
    type: 'scatter',
    data: {
      datasets: [
        { label: 'Each row (predicted while held out)', data: preds.map(p => ({ x: p.actual, y: p.predicted })), backgroundColor: '#0d44e8', pointRadius: 5, pointHoverRadius: 7 },
        { type: 'line', label: 'Perfect prediction (y = x)', data: [{ x: lo, y: lo }, { x: hi, y: hi }], borderColor: '#94a3b8', borderDash: [6, 6], borderWidth: 1.5, pointRadius: 0 }
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      scales: {
        x: { title: { display: true, text: `Actual ${AppState.automl.targetColumn}`, font: { size: 11 } }, grid: { color: '#f1f5f9' } },
        y: { title: { display: true, text: 'Predicted', font: { size: 11 } }, grid: { color: '#f1f5f9' } }
      }
    }
  });
}

// ---------------------------------------------------------------------------
// What-if playground
// ---------------------------------------------------------------------------
function buildWhatIfSimulationInputs() {
  const container = document.getElementById('mlInferenceInputsContainer');
  const runBtn = document.getElementById('mlRunInferenceBtn');
  const resVal = document.getElementById('mlInferenceResultValue');
  const confText = document.getElementById('mlInferenceConfidenceText');
  const automl = AppState.automl;
  if (!container) return;
  container.innerHTML = '';

  const setText = (id, text) => { const el = document.getElementById(id); if (el) el.textContent = text; };
  setText('mlActivePipelineName', automl.bestModel?.name || '–');
  if (resVal) resVal.textContent = '–';
  if (confText) confText.textContent = 'Set the inputs and run the model.';
  setText('mlActiveConfidencePct', '–');

  if (automl.taskType === 'forecasting') {
    if (runBtn) runBtn.classList.add('hidden');
    const next = automl.forecast?.[0];
    container.innerHTML = `<div class="sm:col-span-2 text-xs text-slate-600 bg-slate-50 border border-slate-200 rounded-xl p-3 leading-relaxed">
      Forecasts need no inputs. The chart above shows every future period; the card on the right shows the next one.
      Use <strong>Download Predictions</strong> for the full table.</div>`;
    if (next && resVal) resVal.textContent = next.predicted.toLocaleString(undefined, { maximumFractionDigits: 2 });
    if (next && confText) confText.textContent = `Next period (${next.date}) · 95% range ${next.lowerBound.toLocaleString()} – ${next.upperBound.toLocaleString()}`;
    setText('mlActiveConfidencePct', '95% range');
    return;
  }
  if (runBtn) runBtn.classList.remove('hidden');

  (automl.featureCols || []).forEach(col => {
    const stat = automl.colStats[col] || { type: 'numeric', median: 0 };
    const div = document.createElement('div');
    const id = `mlInput_${col}`;
    if (stat.type === 'numeric') {
      div.innerHTML = `
        <label class="text-[11px] font-semibold text-slate-700 block mb-1">${escapeHtml(col)}</label>
        <input type="number" step="any" id="${escapeHtml(id)}" value="${stat.median}"
          class="w-full text-xs bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 font-mono text-slate-800 focus:outline-none focus:border-indigo-500">
        <span class="text-[10px] text-slate-400">seen in data: ${stat.min.toLocaleString()} – ${stat.max.toLocaleString()}</span>`;
    } else {
      div.innerHTML = `
        <label class="text-[11px] font-semibold text-slate-700 block mb-1">${escapeHtml(col)}</label>
        <select id="${escapeHtml(id)}" class="w-full text-xs bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 font-medium text-slate-800 focus:outline-none focus:border-indigo-500">
          ${stat.categories.map(cat => `<option value="${escapeHtml(cat)}">${escapeHtml(cat)}</option>`).join('')}
        </select>`;
    }
    container.appendChild(div);
  });
}

function runSinglePrediction() {
  const automl = AppState.automl;
  if (!automl.bestModel || !automl.vectorizeRow) {
    showToast('Train the AutoML models first', true);
    return;
  }
  const row = {};
  const warnings = [];
  (automl.featureCols || []).forEach(col => {
    const el = document.getElementById(`mlInput_${col}`);
    if (!el) return;
    row[col] = el.value;
    const stat = automl.colStats[col];
    if (stat?.type === 'numeric') {
      const v = toNumberOrNull(el.value);
      if (v !== null && (v < stat.min || v > stat.max)) warnings.push(`${col} is outside the range seen in the data`);
    }
  });

  const model = automl.bestModel.fitted;
  const x = automl.vectorizeRow(row);
  const resVal = document.getElementById('mlInferenceResultValue');
  const confText = document.getElementById('mlInferenceConfidenceText');
  const confPct = document.getElementById('mlActiveConfidencePct');

  if (automl.taskType === 'classification') {
    const pred = model.predict(x);
    const proba = model.proba ? model.proba(x) : null;
    const p = proba ? proba[automl.classes.indexOf(pred)] : null;
    if (resVal) resVal.innerHTML = `<span class="px-3 py-1 bg-indigo-100 text-indigo-800 rounded-xl">${escapeHtml(pred)}</span>`;
    if (confText) confText.textContent = p !== null ? `The model gives this class a ${(p * 100).toFixed(0)}% probability.` : 'Predicted class.';
    if (confPct) confPct.textContent = p !== null ? `${(p * 100).toFixed(0)}% probability` : '–';
  } else {
    let pred = model.predict(x);
    const range = automl.targetRange;
    let note = '';
    if (range && range.min >= 0 && pred < 0) {
      pred = 0;
      note = ' Raw output was negative; shown as 0 because the target is never negative in your data.';
    }
    const rmse = automl.bestModel.metrics.rmse || 0;
    const lo = pred - 1.96 * rmse, hi = pred + 1.96 * rmse;
    if (resVal) resVal.textContent = pred.toLocaleString(undefined, { maximumFractionDigits: 2 });
    if (confText) confText.textContent = `Typical range ${(range && range.min >= 0 ? Math.max(0, lo) : lo).toLocaleString(undefined, { maximumFractionDigits: 0 })} – ${hi.toLocaleString(undefined, { maximumFractionDigits: 0 })} (±1.96 × cross-validated RMSE).${note}`;
    if (confPct) confPct.textContent = `±${(1.96 * rmse).toLocaleString(undefined, { maximumFractionDigits: 0 })}`;
  }
  if (warnings.length && confText) confText.textContent += ` Caution: ${warnings.join('; ')}, so this is an extrapolation.`;
}

// ---------------------------------------------------------------------------
// Exports
// ---------------------------------------------------------------------------
function mlDownload(content, type, filename) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function csvCell(v) {
  const s = String(v ?? '');
  return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

function exportHoldoutPredictionsCsv() {
  const automl = AppState.automl;
  const preds = automl.testPredictions;
  if (!preds || preds.length === 0) { showToast('Run AutoML training first', true); return; }
  const target = automl.targetColumn || 'Target';
  let lines;
  if (automl.taskType === 'forecasting') {
    lines = [['Type', 'Period', `Actual_${target}`, `Predicted_${target}`, 'Lower95', 'Upper95'].join(',')];
    preds.forEach(p => lines.push([p.type, p.date, p.actual, p.predicted, p.lowerBound, p.upperBound].map(csvCell).join(',')));
  } else if (automl.taskType === 'classification') {
    lines = [['Row', `Actual_${target}`, `Predicted_${target}`, 'Result'].join(',')];
    preds.forEach(p => lines.push([p.row, p.actual, p.predicted, p.residual].map(csvCell).join(',')));
  } else {
    lines = [['Row', `Actual_${target}`, `Predicted_${target}`, 'Residual'].join(',')];
    preds.forEach(p => lines.push([p.row, p.actual, p.predicted, p.residual].map(csvCell).join(',')));
  }
  mlDownload(lines.join('\n'), 'text/csv', `DataVista_AutoML_${automl.bestModel.algorithm}_${Date.now()}.csv`);
  showToast(`Exported ${preds.length} predictions from ${automl.bestModel.name}`);
}

function buildMlAuditPayload() {
  const a = AppState.automl;
  return {
    platform: 'Data Vista',
    timestamp: new Date().toISOString(),
    dataset: AppState.fileName,
    sheet: AppState.activeSheet,
    datasetRows: AppState.currentRows.length,
    taskType: a.taskType,
    targetColumn: a.targetColumn,
    featureColumns: a.featureCols,
    primaryMetric: a.primaryMetric,
    evaluation: a.evaluation,
    randomSeed: ML_SEED,
    winningModel: { name: a.bestModel?.name, algorithm: a.bestModel?.algorithm, metrics: a.bestModel?.metrics },
    leaderboard: (a.models || []).map(m => ({ rank: m.rank, name: m.name, algorithm: m.algorithm, isBaseline: m.isBaseline, metrics: m.metrics })),
    featureImportance: a.featureImportance
  };
}

function exportMlReport() {
  if (!AppState.automl.bestModel) { showToast('Train AutoML models first to export the audit', true); return; }
  mlDownload(JSON.stringify(buildMlAuditPayload(), null, 2), 'application/json', `DataVista_ML_Audit_${Date.now()}.json`);
  showToast('Model audit JSON downloaded');
}
