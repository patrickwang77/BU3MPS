import { initializeApp } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-app.js";
import { getFirestore, collection, getDocs, doc, getDoc, setDoc, writeBatch, query, where } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-firestore.js";

// Firebase Configuration
const firebaseConfig = {
    projectId: "bu3mps",
    appId: "1:742896593160:web:b645ace237a18b4764a4bd",
    storageBucket: "bu3mps.firebasestorage.app",
    apiKey: "AIzaSyA92tsbdo45FAOTndA53mXDHRFedPVHyaM",
    authDomain: "bu3mps.firebaseapp.com",
    messagingSenderId: "742896593160",
    measurementId: "G-6CNP9P3MVL"
};

// Initialize Firebase
const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

// State Management
let allRecords = [];
let versionsList = [];
let currentVersion = "";
let currentSheet = "總表";

// View Level states
let sumViewLevel = "weekly"; // "weekly", "monthly", "quarterly", "yearly"
let wbViewLevel = "weekly";  // "weekly", "monthly", "quarterly", "yearly"
let diffViewLevel = "weekly"; // "weekly", "monthly", "quarterly", "yearly"

// Summary Tab state
let currentSumGroupBy = "customer"; // "customer" or "projectCode"
let sumStartPeriod = "";
let sumEndPeriod = "";
let sumPeriodsList = [];

// Whiteboard Tab filters
let currentWbFilters = {
    customer: "ALL",
    projectCode: "ALL",
    se: "ALL",
    attribute: "ALL",
    search: ""
};
let isDetailMode = false;
let trendsChart = null;

// Difference Analysis Tab state
let diffBaseVersion = "";
let diffCompVersion = "";
let diffSheet = "總表";
let diffGroupByDim = "customer"; // "customer" or "projectCode"
let diffStartPeriod = "";
let diffEndPeriod = "";
let diffPeriodsList = [];
let diffTypeDisplayVal = "ALL"; // "ALL", "PO", "FCST"

// AI Comment Tab state
let aiSelectedVersions = [""]; // List of version selects
let aiApiKeyVal = "";
let aiModelVal = "gemini-2.5-flash";
let currentRawReportMd = "";

// Waterfall Tab state
let wfSelectedVersions = [""]; // List of version selects
let wfSheet = "總表";
let wfViewLevel = "weekly"; // "weekly", "monthly", "quarterly", "yearly"
let wfCustomer = "ALL";
let wfProjectCode = "ALL";
let wfStartPeriod = "";
let wfEndPeriod = "";
let wfPeriodsList = [];
let wfRecordsCache = new Map(); // Cache: version_sheet -> array of records

// Sorting states
let summarySortColumn = null;
let summarySortDirection = null;
let wbSortColumn = null;
let wbSortDirection = null;
let diffSortColumn = null;
let diffSortDirection = null;

// UI Elements (Shared / Sync)
const syncVersionSelects = document.querySelectorAll(".version-select-sync");
const syncSheetSelects = document.querySelectorAll(".sheet-select-sync");

// Summary Tab Elements
const sumGroupBy = document.getElementById("sumGroupBy");
const sumStartPeriodSelect = document.getElementById("sumStartPeriod");
const sumEndPeriodSelect = document.getElementById("sumEndPeriod");
const summaryTableTitle = document.getElementById("summaryTableTitle");
const summaryHeaderRow = document.getElementById("summaryHeaderRow");
const summaryTableBody = document.getElementById("summaryTableBody");
const summaryTotalsRow = document.getElementById("summaryTotalsRow");

// Whiteboard Tab Elements
const filterCustomer = document.getElementById("filterCustomer");
const filterProjectCode = document.getElementById("filterProjectCode");
const filterSe = document.getElementById("filterSe");
const filterAttribute = document.getElementById("filterAttribute");
const searchInput = document.getElementById("searchInput");
const detailModeToggle = document.getElementById("detailModeToggle");
const tableHeaderRow = document.getElementById("tableHeaderRow");
const tableBody = document.getElementById("tableBody");
const totalsRow = document.getElementById("totalsRow");

// Difference Tab Elements
const diffBaseVersionSelect = document.getElementById("diffBaseVersionSelect");
const diffCompVersionSelect = document.getElementById("diffCompVersionSelect");
const diffSheetSelect = document.getElementById("diffSheetSelect");
const diffGroupBy = document.getElementById("diffGroupBy");
const diffStartPeriodSelect = document.getElementById("diffStartPeriod");
const diffEndPeriodSelect = document.getElementById("diffEndPeriod");
const diffTypeDisplay = document.getElementById("diffTypeDisplay");
const diffTableTitle = document.getElementById("diffTableTitle");
const diffHeaderRow = document.getElementById("diffHeaderRow");
const diffTableBody = document.getElementById("diffTableBody");
const diffTotalsRow = document.getElementById("diffTotalsRow");

// AI Tab Elements
const aiApiKeyInput = document.getElementById("aiApiKey");
const saveApiKeyBtn = document.getElementById("saveApiKeyBtn");
const clearApiKeyBtn = document.getElementById("clearApiKeyBtn");
const aiVersionsContainer = document.getElementById("aiVersionsContainer");
const addAiVersionBtn = document.getElementById("addAiVersionBtn");
const aiPrompt = document.getElementById("aiPrompt");
const generateAiBtn = document.getElementById("generateAiBtn");
const aiOutput = document.getElementById("aiOutput");
const aiStatusLabel = document.getElementById("aiStatusLabel");
const aiProgressBarContainer = document.getElementById("aiProgressBarContainer");

// Waterfall Tab Elements
const wfVersionsContainer = document.getElementById("wfVersionsContainer");
const addWfVersionBtn = document.getElementById("addWfVersionBtn");
const wfSheetSelect = document.getElementById("wfSheetSelect");
const filterWfCustomer = document.getElementById("filterWfCustomer");
const filterWfProjectCode = document.getElementById("filterWfProjectCode");
const wfStartPeriodSelect = document.getElementById("wfStartPeriod");
const wfEndPeriodSelect = document.getElementById("wfEndPeriod");
const exportWfExcelBtn = document.getElementById("exportWfExcelBtn");
const wfHeaderRow = document.getElementById("wfHeaderRow");
const wfTableBody = document.getElementById("wfTableBody");
const wfTotalsRow = document.getElementById("wfTotalsRow");

// Excel Export Buttons
const exportSumExcelBtn = document.getElementById("exportSumExcelBtn");
const exportDiffExcelBtn = document.getElementById("exportDiffExcelBtn");

// Upload Tab Elements
const uploadZone = document.getElementById("uploadZone");
const fileInput = document.getElementById("fileInput");
const consoleBody = document.getElementById("consoleBody");
const progressBarContainer = document.getElementById("progressBarContainer");
const progressBar = document.getElementById("progressBar");

// Initialize Application
window.addEventListener("DOMContentLoaded", async () => {
    setupEventListeners();
    loadStoredApiKey();
    addConsoleLog("System initialized. Loading available plan versions...", "info");
    await loadVersions();
});

// Setup Events
function setupEventListeners() {
    // Navigation Tabs
    const tabBtns = document.querySelectorAll(".tab-btn");
    const panels = document.querySelectorAll(".view-panel");
    tabBtns.forEach(btn => {
        btn.addEventListener("click", () => {
            tabBtns.forEach(b => b.classList.remove("active"));
            panels.forEach(p => p.classList.remove("active"));
            btn.classList.add("active");
            document.getElementById(btn.dataset.target).classList.add("active");
            
            // Refresh views
            if (btn.dataset.target === "summaryPanel") {
                renderSummaryGrid();
            } else if (btn.dataset.target === "waterfallPanel") {
                renderWfVersionSelectors();
                initWaterfallPeriods();
                renderWaterfallGrid();
            } else if (btn.dataset.target === "whiteboardPanel") {
                renderWhiteboardGrid();
            } else if (btn.dataset.target === "diffPanel") {
                initDiffPeriods();
                renderDiffGrid();
            } else if (btn.dataset.target === "aiPanel") {
                renderAiVersionSelectors();
            }
        });
    });

    // Sync Version Dropdowns (for Summary and Whiteboard)
    syncVersionSelects.forEach(select => {
        select.addEventListener("change", (e) => {
            currentVersion = e.target.value;
            syncVersionSelects.forEach(sel => { sel.value = currentVersion; });
            loadDataAndRefresh();
        });
    });

    // Sync Sheet Dropdowns (for Summary and Whiteboard)
    syncSheetSelects.forEach(select => {
        select.addEventListener("change", (e) => {
            currentSheet = e.target.value;
            syncSheetSelects.forEach(sel => { sel.value = currentSheet; });
            loadDataAndRefresh();
        });
    });

    // Summary Tab events
    sumGroupBy.addEventListener("change", (e) => {
        currentSumGroupBy = e.target.value;
        summarySortColumn = null;
        summarySortDirection = null;
        renderSummaryGrid();
    });
    sumStartPeriodSelect.addEventListener("change", (e) => {
        sumStartPeriod = e.target.value;
        renderSummaryGrid();
    });
    sumEndPeriodSelect.addEventListener("change", (e) => {
        sumEndPeriod = e.target.value;
        renderSummaryGrid();
    });

    // Summary Tab - View Level Selection
    const sumLevelBtns = document.querySelectorAll("#sumViewLevelSelector .level-btn");
    sumLevelBtns.forEach(btn => {
        btn.addEventListener("click", () => {
            sumLevelBtns.forEach(b => b.classList.remove("active"));
            btn.classList.add("active");
            sumViewLevel = btn.dataset.level;
            summarySortColumn = null;
            summarySortDirection = null;
            initSummaryPeriods();
            renderSummaryGrid();
        });
    });

    // Whiteboard Tab - Dimensional Filters
    filterCustomer.addEventListener("change", () => {
        currentWbFilters.customer = filterCustomer.value;
        renderWhiteboardGrid();
    });
    filterProjectCode.addEventListener("change", () => {
        currentWbFilters.projectCode = filterProjectCode.value;
        renderWhiteboardGrid();
    });
    filterSe.addEventListener("change", () => {
        currentWbFilters.se = filterSe.value;
        renderWhiteboardGrid();
    });
    filterAttribute.addEventListener("change", () => {
        currentWbFilters.attribute = filterAttribute.value;
        renderWhiteboardGrid();
    });
    searchInput.addEventListener("input", () => {
        currentWbFilters.search = searchInput.value.toLowerCase().trim();
        renderWhiteboardGrid();
    });

    // Whiteboard Tab - View Level Selection
    const wbLevelBtns = document.querySelectorAll("#wbViewLevelSelector .level-btn");
    wbLevelBtns.forEach(btn => {
        btn.addEventListener("click", () => {
            wbLevelBtns.forEach(b => b.classList.remove("active"));
            btn.classList.add("active");
            wbViewLevel = btn.dataset.level;
            wbSortColumn = null;
            wbSortDirection = null;
            renderWhiteboardGrid();
        });
    });

    // Whiteboard Tab - Detail Mode Toggle
    detailModeToggle.addEventListener("change", (e) => {
        isDetailMode = e.target.checked;
        renderWhiteboardGrid();
    });

    // Difference Tab events
    diffBaseVersionSelect.addEventListener("change", (e) => {
        diffBaseVersion = e.target.value;
        initDiffPeriods();
        renderDiffGrid();
    });
    diffCompVersionSelect.addEventListener("change", (e) => {
        diffCompVersion = e.target.value;
        initDiffPeriods();
        renderDiffGrid();
    });
    diffSheetSelect.addEventListener("change", (e) => {
        diffSheet = e.target.value;
        initDiffPeriods();
        renderDiffGrid();
    });
    diffGroupBy.addEventListener("change", (e) => {
        diffGroupByDim = e.target.value;
        diffSortColumn = null;
        diffSortDirection = null;
        renderDiffGrid();
    });
    diffStartPeriodSelect.addEventListener("change", (e) => {
        diffStartPeriod = e.target.value;
        renderDiffGrid();
    });
    diffEndPeriodSelect.addEventListener("change", (e) => {
        diffEndPeriod = e.target.value;
        renderDiffGrid();
    });
    diffTypeDisplay.addEventListener("change", (e) => {
        diffTypeDisplayVal = e.target.value;
        renderDiffGrid();
    });

    // Difference Tab - View Level Selection
    const diffLevelBtns = document.querySelectorAll("#diffViewLevelSelector .level-btn");
    diffLevelBtns.forEach(btn => {
        btn.addEventListener("click", () => {
            diffLevelBtns.forEach(b => b.classList.remove("active"));
            btn.classList.add("active");
            diffViewLevel = btn.dataset.level;
            diffSortColumn = null;
            diffSortDirection = null;
            initDiffPeriods();
            renderDiffGrid();
        });
    });

    // AI Tab - API Key and Model storage
    saveApiKeyBtn.addEventListener("click", () => {
        const key = aiApiKeyInput.value.trim();
        const model = document.getElementById("aiModelSelect").value;
        if (key) {
            localStorage.setItem("mps_gemini_api_key", key);
            localStorage.setItem("mps_gemini_model", model);
            aiApiKeyVal = key;
            aiModelVal = model;
            alert("金鑰與模型設定已安全儲存至您的瀏覽器本機 (localStorage)！");
        } else {
            alert("請輸入有效的 API 金鑰！");
        }
    });

    clearApiKeyBtn.addEventListener("click", () => {
        localStorage.removeItem("mps_gemini_api_key");
        localStorage.removeItem("mps_gemini_model");
        aiApiKeyInput.value = "";
        aiApiKeyVal = "";
        aiModelVal = "gemini-2.5-flash";
        document.getElementById("aiModelSelect").value = "gemini-2.5-flash";
        alert("金鑰與模型設定已從瀏覽器中清除！");
    });

    // AI Tab - Version setup
    addAiVersionBtn.addEventListener("click", () => {
        aiSelectedVersions.push("");
        renderAiVersionSelectors();
    });

    generateAiBtn.addEventListener("click", generateAiAnalysis);

    document.getElementById("exportMdBtn").addEventListener("click", exportReportAsMarkdown);
    document.getElementById("exportHtmlBtn").addEventListener("click", exportReportAsHtml);

    // Drag and Drop Uploader
    uploadZone.addEventListener("dragover", (e) => {
        e.preventDefault();
        uploadZone.classList.add("dragover");
    });
    uploadZone.addEventListener("dragleave", () => {
        uploadZone.classList.remove("dragover");
    });
    uploadZone.addEventListener("drop", (e) => {
        e.preventDefault();
        uploadZone.classList.remove("dragover");
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            handleFileUpload(files[0]);
        }
    });
    fileInput.addEventListener("change", (e) => {
        const files = e.target.files;
        if (files.length > 0) {
            handleFileUpload(files[0]);
        }
    });
    
    // Waterfall Tab events
    addWfVersionBtn.addEventListener("click", () => {
        wfSelectedVersions.push("");
        renderWfVersionSelectors();
    });

    wfSheetSelect.addEventListener("change", (e) => {
        wfSheet = e.target.value;
        initWaterfallPeriods();
        renderWaterfallGrid();
    });

    // Waterfall Tab - View Level Selection
    const wfLevelBtns = document.querySelectorAll("#wfViewLevelSelector .level-btn");
    wfLevelBtns.forEach(btn => {
        btn.addEventListener("click", () => {
            wfLevelBtns.forEach(b => b.classList.remove("active"));
            btn.classList.add("active");
            wfViewLevel = btn.dataset.level;
            
            initWaterfallPeriods();
            renderWaterfallGrid();
        });
    });

    filterWfCustomer.addEventListener("change", (e) => {
        wfCustomer = e.target.value;
        renderWaterfallGrid();
    });

    filterWfProjectCode.addEventListener("change", (e) => {
        wfProjectCode = e.target.value;
        renderWaterfallGrid();
    });

    wfStartPeriodSelect.addEventListener("change", (e) => {
        wfStartPeriod = e.target.value;
        renderWaterfallGrid();
    });

    wfEndPeriodSelect.addEventListener("change", (e) => {
        wfEndPeriod = e.target.value;
        renderWaterfallGrid();
    });

    // Excel Export Buttons
    exportSumExcelBtn.addEventListener("click", () => {
        exportHtmlTableToExcel("summaryTable", `產銷加總表_${currentVersion}_${currentSheet}.xlsx`, "產銷加總表");
    });
    exportWfExcelBtn.addEventListener("click", () => {
        exportHtmlTableToExcel("waterfallTable", `Waterfall需求推移表_${wfSheet}.xlsx`, "Waterfall推移表");
    });
    exportDiffExcelBtn.addEventListener("click", () => {
        exportHtmlTableToExcel("diffTable", `版本差異分析_${diffBaseVersion}_vs_${diffCompVersion}.xlsx`, "版本差異分析");
    });

    // Theme Toggle Handler
    const themeToggleBtn = document.getElementById("themeToggleBtn");
    const themeIcon = document.getElementById("themeIcon");
    
    // SVG icons
    const sunSvg = `<path d="M8 11a3 3 0 1 1 0-6 3 3 0 0 1 0 6zm0 1a4 4 0 1 0 0-8 4 4 0 0 0 0 8zM8 0a.5.5 0 0 1 .5.5v2a.5.5 0 0 1-1 0v-2A.5.5 0 0 1 8 0zm0 13a.5.5 0 0 1 .5.5v2a.5.5 0 0 1-1 0v-2A.5.5 0 0 1 8 13zm8-5a.5.5 0 0 1-.5.5h-2a.5.5 0 0 1 0-1h2a.5.5 0 0 1 .5.5zM3 8a.5.5 0 0 1-.5.5h-2a.5.5 0 0 1 0-1h2a.5.5 0 0 1 .5.5zm10.657-5.657a.5.5 0 0 1 0 .707l-1.414 1.415a.5.5 0 1 1-.707-.708l1.414-1.414a.5.5 0 0 1 .707 0zm-9.193 9.193a.5.5 0 0 1 0 .707L3.05 13.657a.5.5 0 0 1-.707-.707l1.414-1.414a.5.5 0 0 1 .707 0zm9.193 2.121a.5.5 0 0 1-.707 0l-1.414-1.414a.5.5 0 0 1 .707-.707l1.414 1.414a.5.5 0 0 1 0 .707zM4.464 4.465a.5.5 0 0 1-.707 0L2.343 3.05a.5.5 0 1 1 .707-.707l1.414 1.414a.5.5 0 0 1 0 .708z"/>`;
    const moonSvg = `<path d="M6 .278a.768.768 0 0 1 .08.858 7.208 7.208 0 0 0-.878 3.46c0 4.021 3.278 7.277 7.318 7.277.527 0 1.04-.055 1.533-.16a.787.787 0 0 1 .81.316.733.733 0 0 1-.031.893A8.349 8.349 0 0 1 8.344 16C3.734 16 0 12.286 0 7.71 0 4.266 2.114 1.312 5.124.06A.752.752 0 0 1 6 .278z"/>`;
    
    // Set initial icon based on theme
    const updateThemeIcon = (theme) => {
        if (theme === "light") {
            if (themeIcon) themeIcon.innerHTML = sunSvg;
        } else {
            if (themeIcon) themeIcon.innerHTML = moonSvg;
        }
    };
    
    const initialTheme = document.documentElement.getAttribute("data-theme") || "dark";
    updateThemeIcon(initialTheme);
    
    if (themeToggleBtn) {
        themeToggleBtn.addEventListener("click", () => {
            const currentTheme = document.documentElement.getAttribute("data-theme") || "dark";
            const newTheme = currentTheme === "dark" ? "light" : "dark";
            
            document.documentElement.setAttribute("data-theme", newTheme);
            localStorage.setItem("theme", newTheme);
            updateThemeIcon(newTheme);
            addConsoleLog(`切換佈景主題為: ${newTheme === 'dark' ? '深色模式' : '淺色模式'}`, "info");
        });
    }
}

// Load Stored Gemini API key and model
function loadStoredApiKey() {
    const key = localStorage.getItem("mps_gemini_api_key");
    if (key) {
        aiApiKeyInput.value = key;
        aiApiKeyVal = key;
    }
    const model = localStorage.getItem("mps_gemini_model");
    if (model) {
        const select = document.getElementById("aiModelSelect");
        if (select) {
            select.value = model;
            aiModelVal = model;
        }
    }
}

// Load available plan versions
async function loadVersions() {
    try {
        let versionsSet = new Set();
        
        // 1. Try reading metadata versions list doc (1 read)
        const metaRef = doc(db, "metadata", "versions");
        const metaSnap = await getDoc(metaRef);
        
        if (metaSnap.exists()) {
            const metaData = metaSnap.data();
            if (metaData && Array.isArray(metaData.list) && metaData.list.length > 0) {
                metaData.list.forEach(v => versionsSet.add(v));
                addConsoleLog(`Metadata version list loaded (1 read).`, "success");
            }
        }
        
        // 2. Fallback if metadata doesn't exist yet or is empty
        if (versionsSet.size === 0) {
            addConsoleLog("No cached versions list found. Scanning collection...", "info");
            const q = query(collection(db, "mps_records"));
            const snapshot = await getDocs(q);
            
            snapshot.forEach(doc => {
                const data = doc.data();
                if (data.version) {
                    versionsSet.add(data.version);
                }
            });
            
            versionsList = Array.from(versionsSet).sort((a, b) => b.localeCompare(a));
            
            if (versionsList.length > 0) {
                // Save versions to metadata document (1 write)
                await setDoc(metaRef, { list: versionsList, updatedAt: new Date().toISOString() });
                addConsoleLog(`Saved scanned versions list to metadata collection.`, "success");
            }
        } else {
            versionsList = Array.from(versionsSet).sort((a, b) => b.localeCompare(a));
        }
        
        if (versionsList.length === 0) {
            addConsoleLog("No versions found in Firestore. Please upload a file first.", "warn");
            return;
        }
        
        addConsoleLog(`Loaded ${versionsList.length} version(s) successfully.`, "success");
        
        // Populate sync dropdowns (Summary, Whiteboard)
        syncVersionSelects.forEach(select => {
            select.innerHTML = "";
            versionsList.forEach(ver => {
                const opt = document.createElement("option");
                opt.value = ver;
                opt.textContent = ver;
                select.appendChild(opt);
            });
        });
        
        // Populate Difference Version selectors
        populateDiffVersionSelects();
        
        // Set active version to the latest
        currentVersion = versionsList[0];
        syncVersionSelects.forEach(sel => { sel.value = currentVersion; });
        
        // Default AI Selected versions: starts with the latest
        aiSelectedVersions = [currentVersion];
        
        // Default Waterfall Selected versions: starts with the latest
        wfSelectedVersions = [currentVersion];
        
        await loadDataAndRefresh();
        
    } catch (err) {
        addConsoleLog(`Failed to load versions: ${err.message}`, "error");
    }
}

function populateDiffVersionSelects() {
    diffBaseVersionSelect.innerHTML = "";
    diffCompVersionSelect.innerHTML = "";
    
    versionsList.forEach(ver => {
        const optBase = document.createElement("option");
        optBase.value = ver;
        optBase.textContent = ver;
        diffBaseVersionSelect.appendChild(optBase);
        
        const optComp = document.createElement("option");
        optComp.value = ver;
        optComp.textContent = ver;
        diffCompVersionSelect.appendChild(optComp);
    });
    
    diffBaseVersion = versionsList[0];
    diffCompVersion = versionsList[1] || versionsList[0];
    
    diffBaseVersionSelect.value = diffBaseVersion;
    diffCompVersionSelect.value = diffCompVersion;
}

// Load data and refresh active panels
async function loadDataAndRefresh() {
    if (!currentVersion) return;
    
    addConsoleLog(`Loading records for version [${currentVersion}], sheet [${currentSheet}]...`, "info");
    
    try {
        const q = query(
            collection(db, "mps_records"),
            where("version", "==", currentVersion),
            where("sheet", "==", currentSheet)
        );
        const snapshot = await getDocs(q);
        allRecords = [];
        snapshot.forEach(doc => {
            allRecords.push(doc.data());
        });
        
        addConsoleLog(`Retrieved ${allRecords.length} records successfully.`, "success");
        
        populateWhiteboardFilters();
        initSummaryPeriods();
        
        // Refresh active panel
        const activeTab = document.querySelector(".tab-btn.active").id;
        if (activeTab === "tabSummary") {
            renderSummaryGrid();
        } else if (activeTab === "tabWhiteboard") {
            renderWhiteboardGrid();
        } else if (activeTab === "tabDiff") {
            initDiffPeriods();
            renderDiffGrid();
        } else if (activeTab === "tabAi") {
            renderAiVersionSelectors();
        }
        
        updateGlobalSummaryStats();
        
    } catch (err) {
        addConsoleLog(`Failed to load data: ${err.message}`, "error");
    }
}

// Update Top Global Statistics Row
function updateGlobalSummaryStats() {
    let totalPO = 0;
    let totalFCST = 0;
    
    allRecords.forEach(rec => {
        if (rec.type === "PO") {
            totalPO += rec.value;
        } else if (rec.type === "FCST") {
            totalFCST += rec.value;
        }
    });
    
    document.getElementById("statVersion").textContent = currentVersion || "-";
    document.getElementById("statTotalQty").textContent = (totalPO + totalFCST).toLocaleString();
    document.getElementById("statPOQty").textContent = totalPO.toLocaleString();
    document.getElementById("statFCSTQty").textContent = totalFCST.toLocaleString();
}

// Chronological sorting of periods
function sortPeriodKeys(periodsSet, viewLevel) {
    const list = Array.from(periodsSet);
    if (viewLevel === "weekly") {
        return list.sort((a, b) => {
            const dateA = calculateMondayDate(a);
            const dateB = calculateMondayDate(b);
            return dateA.localeCompare(dateB);
        });
    } else {
        return list.sort((a, b) => a.localeCompare(b));
    }
}

// ----------------------------------------------------
// TAB 1: 產銷加總表邏輯 (Summary Tab)
// ----------------------------------------------------

function initSummaryPeriods() {
    const periodsSet = new Set();
    allRecords.forEach(rec => {
        periodsSet.add(getTimeColumnKey(rec, sumViewLevel));
    });
    
    sumPeriodsList = sortPeriodKeys(periodsSet, sumViewLevel);
    
    sumStartPeriodSelect.innerHTML = "";
    sumEndPeriodSelect.innerHTML = "";
    
    if (sumPeriodsList.length === 0) return;
    
    sumPeriodsList.forEach(period => {
        const optStart = document.createElement("option");
        optStart.value = period;
        optStart.textContent = period;
        sumStartPeriodSelect.appendChild(optStart);
        
        const optEnd = document.createElement("option");
        optEnd.value = period;
        optEnd.textContent = period;
        sumEndPeriodSelect.appendChild(optEnd);
    });
    
    sumStartPeriod = sumPeriodsList[0];
    sumEndPeriod = sumPeriodsList[sumPeriodsList.length - 1];
    
    sumStartPeriodSelect.value = sumStartPeriod;
    sumEndPeriodSelect.value = sumEndPeriod;
}

function renderSummaryGrid() {
    if (allRecords.length === 0) {
        summaryTableBody.innerHTML = '<tr><td colspan="3" style="text-align:center;padding:2rem;">無可顯示資料</td></tr>';
        return;
    }
    
    const startIdx = sumPeriodsList.indexOf(sumStartPeriod);
    const endIdx = sumPeriodsList.indexOf(sumEndPeriod);
    
    if (startIdx === -1 || endIdx === -1 || startIdx > endIdx) {
        summaryTableBody.innerHTML = '<tr><td colspan="3" style="text-align:center;padding:2rem;color:var(--accent-po);">起始區間不能大於結束區間！</td></tr>';
        return;
    }
    
    const selectedPeriods = sumPeriodsList.slice(startIdx, endIdx + 1);
    
    const rangeFilteredRecords = allRecords.filter(rec => {
        const pKey = getTimeColumnKey(rec, sumViewLevel);
        const idx = sumPeriodsList.indexOf(pKey);
        return idx >= startIdx && idx <= endIdx;
    });
    
    const summaryRowsMap = new Map();
    rangeFilteredRecords.forEach(rec => {
        const groupVal = currentSumGroupBy === "customer" 
            ? (rec.customer ? rec.customer.trim() : "Unknown")
            : (rec.projectCode ? rec.projectCode.trim() : "Unknown");
            
        if (!summaryRowsMap.has(groupVal)) {
            summaryRowsMap.set(groupVal, {
                groupKey: groupVal,
                cells: {}
            });
        }
        
        const pKey = getTimeColumnKey(rec, sumViewLevel);
        const rowData = summaryRowsMap.get(groupVal);
        rowData.cells[pKey] = (rowData.cells[pKey] || 0) + rec.value;
    });
    
    const summaryRows = Array.from(summaryRowsMap.values());
    if (summarySortColumn !== null && summarySortDirection !== null) {
        summaryRows.sort((a, b) => {
            let valA, valB;
            if (summarySortColumn === 0) {
                valA = a.groupKey;
                valB = b.groupKey;
                return summarySortDirection === 'asc'
                    ? valA.localeCompare(valB)
                    : valB.localeCompare(valA);
            } else if (summarySortColumn === 'total') {
                valA = selectedPeriods.reduce((sum, p) => sum + (a.cells[p] || 0), 0);
                valB = selectedPeriods.reduce((sum, p) => sum + (b.cells[p] || 0), 0);
            } else {
                valA = a.cells[summarySortColumn] || 0;
                valB = b.cells[summarySortColumn] || 0;
            }
            if (valA < valB) return summarySortDirection === 'asc' ? -1 : 1;
            if (valA > valB) return summarySortDirection === 'asc' ? 1 : -1;
            return 0;
        });
    } else {
        summaryRows.sort((a, b) => a.groupKey.localeCompare(b.groupKey));
    }
    
    const groupLabel = currentSumGroupBy === "customer" ? "客戶 (Customer)" : "專案 (Project Code)";
    summaryTableTitle.textContent = `產銷小計匯總表 - ${groupLabel}`;
    
    let headerHtml = `<th onclick="toggleSummarySort(0)" style="position: sticky; left: 0; z-index: 5; background-color: #161c2b; border-right: 2px solid var(--border-color); cursor: pointer;">${groupLabel} <span class="sort-icon ${summarySortColumn === 0 ? 'active' : ''}">${getSortIconStr(0, summarySortColumn, summarySortDirection)}</span></th>`;
    selectedPeriods.forEach(p => {
        headerHtml += `<th onclick="toggleSummarySort('${p}')" style="cursor: pointer;">${formatColumnHeaderLabel(p, sumViewLevel)} <span class="sort-icon ${summarySortColumn === p ? 'active' : ''}">${getSortIconStr(p, summarySortColumn, summarySortDirection)}</span></th>`;
    });
    headerHtml += `<th onclick="toggleSummarySort('total')" class="total-column" style="cursor: pointer;">加總小計 <span class="sort-icon ${summarySortColumn === 'total' ? 'active' : ''}">${getSortIconStr('total', summarySortColumn, summarySortDirection)}</span></th>`;
    summaryHeaderRow.innerHTML = headerHtml;
    
    const colTotals = {};
    selectedPeriods.forEach(p => { colTotals[p] = 0; });
    let totalSum = 0;
    
    let bodyHtml = "";
    summaryRows.forEach(row => {
        let rowHtml = `<tr class="data-row"><td style="position: sticky; left: 0; z-index: 4; background-color: #111624; font-weight: 500; border-right: 2px solid var(--border-color);">${row.groupKey}</td>`;
        let rowSum = 0;
        
        selectedPeriods.forEach(p => {
            const val = row.cells[p] || 0;
            rowSum += val;
            colTotals[p] += val;
            
            rowHtml += `<td class="data-cell">${val > 0 ? val.toLocaleString() : "-"}</td>`;
        });
        
        totalSum += rowSum;
        rowHtml += `<td class="total-column">${rowSum.toLocaleString()}</td></tr>`;
        bodyHtml += rowHtml;
    });
    
    summaryTableBody.innerHTML = bodyHtml;
    
    let totalsHtml = `<td style="position: sticky; left: 0; z-index: 4; background-color: #161d2b; font-weight: 600; border-right: 2px solid var(--border-color);">加總小計</td>`;
    selectedPeriods.forEach(p => {
        const val = colTotals[p];
        totalsHtml += `<td class="data-cell">${val > 0 ? val.toLocaleString() : "-"}</td>`;
    });
    totalsHtml += `<td class="total-column">${totalSum.toLocaleString()}</td>`;
    summaryTotalsRow.innerHTML = totalsHtml;
    
    renderSummaryChart(rangeFilteredRecords, selectedPeriods);
}

function renderSummaryChart(filteredRecs, timePeriods) {
    const ctx = document.getElementById("trendsChartCanvas");
    if (!ctx) return;
    
    if (trendsChart) {
        trendsChart.destroy();
    }
    
    const poValues = Array(timePeriods.length).fill(0);
    const fcstValues = Array(timePeriods.length).fill(0);
    
    filteredRecs.forEach(rec => {
        const colKey = getTimeColumnKey(rec, sumViewLevel);
        const idx = timePeriods.indexOf(colKey);
        if (idx !== -1) {
            if (rec.type === "PO") {
                poValues[idx] += rec.value;
            } else if (rec.type === "FCST") {
                fcstValues[idx] += rec.value;
            }
        }
    });
    
    trendsChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: timePeriods.map(l => l.replace("-", "/")),
            datasets: [
                {
                    label: 'PO (採購單)',
                    data: poValues,
                    backgroundColor: 'rgba(16, 185, 129, 0.65)',
                    borderColor: '#10b981',
                    borderWidth: 1,
                    borderRadius: 4
                },
                {
                    label: 'FCST (預估數)',
                    data: fcstValues,
                    backgroundColor: 'rgba(59, 130, 246, 0.65)',
                    borderColor: '#3b82f6',
                    borderWidth: 1,
                    borderRadius: 4
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    labels: {
                        color: '#94a3b8',
                        font: { family: 'Outfit', size: 11 }
                    }
                }
            },
            scales: {
                x: {
                    grid: { color: 'rgba(255, 255, 255, 0.03)' },
                    ticks: { color: '#94a3b8', font: { family: 'Outfit', size: 10 } }
                },
                y: {
                    grid: { color: 'rgba(255, 255, 255, 0.03)' },
                    ticks: { color: '#94a3b8', font: { family: 'Outfit', size: 10 } }
                }
            }
        }
    });
}

// ----------------------------------------------------
// TAB 2: 大白板明細邏輯 (Whiteboard Tab)
// ----------------------------------------------------

function populateWhiteboardFilters() {
    const customers = new Set();
    const projectCodes = new Set();
    const ses = new Set();
    const attributes = new Set();
    
    allRecords.forEach(rec => {
        if (rec.customer) customers.add(rec.customer);
        if (rec.projectCode) projectCodes.add(rec.projectCode);
        if (rec.se) ses.add(rec.se);
        if (rec.attribute) attributes.add(rec.attribute);
    });
    
    setupDropdown(filterCustomer, customers);
    setupDropdown(filterProjectCode, projectCodes);
    setupDropdown(filterSe, ses);
    setupDropdown(filterAttribute, attributes);
}

function renderWhiteboardGrid() {
    const wbFilteredRecords = allRecords.filter(rec => {
        if (currentWbFilters.customer !== "ALL" && rec.customer !== currentWbFilters.customer) return false;
        if (currentWbFilters.projectCode !== "ALL" && rec.projectCode !== currentWbFilters.projectCode) return false;
        if (currentWbFilters.se !== "ALL" && rec.se !== currentWbFilters.se) return false;
        if (currentWbFilters.attribute !== "ALL" && rec.attribute !== currentWbFilters.attribute) return false;
        
        if (currentWbFilters.search) {
            const matches = 
                (rec.pn && rec.pn.toLowerCase().includes(currentWbFilters.search)) ||
                (rec.pcba && rec.pcba.toLowerCase().includes(currentWbFilters.search)) ||
                (rec.projectCode && rec.projectCode.toLowerCase().includes(currentWbFilters.search)) ||
                (rec.customer && rec.customer.toLowerCase().includes(currentWbFilters.search));
            if (!matches) return false;
        }
        return true;
    });
    
    if (wbFilteredRecords.length === 0) {
        tableHeaderRow.innerHTML = '<th colspan="8" style="text-align: center;">無符合篩選條件之明細</th>';
        tableBody.innerHTML = '<tr><td colspan="8" style="text-align: center; padding: 2rem;">請嘗試調整篩選條件</td></tr>';
        totalsRow.innerHTML = "";
        return;
    }
    
    const timeColsSet = new Set();
    wbFilteredRecords.forEach(rec => {
        timeColsSet.add(getTimeColumnKey(rec, wbViewLevel));
    });
    const timeColumns = sortPeriodKeys(timeColsSet, wbViewLevel);
    
    const rowsMap = new Map();
    wbFilteredRecords.forEach(rec => {
        const rowKey = `${rec.attribute}||${rec.se}||${rec.projectCode}||${rec.pn}||${rec.pcba}||${rec.customer}`;
        if (!rowsMap.has(rowKey)) {
            rowsMap.set(rowKey, {
                attribute: rec.attribute,
                se: rec.se,
                projectCode: rec.projectCode,
                pn: rec.pn,
                pcba: rec.pcba,
                customer: rec.customer,
                cells: {}
            });
        }
        
        const colKey = getTimeColumnKey(rec, wbViewLevel);
        const rowData = rowsMap.get(rowKey);
        if (!rowData.cells[colKey]) {
            rowData.cells[colKey] = [];
        }
        rowData.cells[colKey].push(rec);
    });
    
    const rows = Array.from(rowsMap.values());
    if (wbSortColumn !== null && wbSortDirection !== null) {
        rows.sort((a, b) => {
            let valA, valB;
            if (['attribute', 'se', 'projectCode', 'pn', 'pcba', 'customer'].includes(wbSortColumn)) {
                valA = a[wbSortColumn] || "";
                valB = b[wbSortColumn] || "";
                return wbSortDirection === 'asc'
                    ? valA.localeCompare(valB)
                    : valB.localeCompare(valA);
            } else if (wbSortColumn === 'total') {
                valA = timeColumns.reduce((sum, col) => sum + (a.cells[col] || []).reduce((s, r) => s + r.value, 0), 0);
                valB = timeColumns.reduce((sum, col) => sum + (b.cells[col] || []).reduce((s, r) => s + r.value, 0), 0);
            } else {
                valA = (a.cells[wbSortColumn] || []).reduce((s, r) => s + r.value, 0);
                valB = (b.cells[wbSortColumn] || []).reduce((s, r) => s + r.value, 0);
            }
            if (valA < valB) return wbSortDirection === 'asc' ? -1 : 1;
            if (valA > valB) return wbSortDirection === 'asc' ? 1 : -1;
            return 0;
        });
    } else {
        rows.sort((a, b) => {
            let comp = a.customer.localeCompare(b.customer);
            if (comp !== 0) return comp;
            comp = a.projectCode.localeCompare(b.projectCode);
            if (comp !== 0) return comp;
            return a.pn.localeCompare(b.pn);
        });
    }
    
    let headerHtml = `
        <th class="sticky-c1" style="cursor: pointer;" onclick="toggleWbSort('attribute')">料號屬性 <span class="sort-icon ${wbSortColumn === 'attribute' ? 'active' : ''}">${getSortIconStr('attribute', wbSortColumn, wbSortDirection)}</span></th>
        <th class="sticky-c2" style="cursor: pointer;" onclick="toggleWbSort('se')">SE <span class="sort-icon ${wbSortColumn === 'se' ? 'active' : ''}">${getSortIconStr('se', wbSortColumn, wbSortDirection)}</span></th>
        <th class="sticky-c3" style="cursor: pointer;" onclick="toggleWbSort('projectCode')">Project Code <span class="sort-icon ${wbSortColumn === 'projectCode' ? 'active' : ''}">${getSortIconStr('projectCode', wbSortColumn, wbSortDirection)}</span></th>
        <th class="sticky-c4" style="cursor: pointer;" onclick="toggleWbSort('pn')">PN <span class="sort-icon ${wbSortColumn === 'pn' ? 'active' : ''}">${getSortIconStr('pn', wbSortColumn, wbSortDirection)}</span></th>
        <th class="sticky-c5" style="cursor: pointer;" onclick="toggleWbSort('pcba')">PCBA <span class="sort-icon ${wbSortColumn === 'pcba' ? 'active' : ''}">${getSortIconStr('pcba', wbSortColumn, wbSortDirection)}</span></th>
        <th class="sticky-c6" style="cursor: pointer;" onclick="toggleWbSort('customer')">Customer <span class="sort-icon ${wbSortColumn === 'customer' ? 'active' : ''}">${getSortIconStr('customer', wbSortColumn, wbSortDirection)}</span></th>
    `;
    timeColumns.forEach(col => {
        headerHtml += `<th style="cursor: pointer;" onclick="toggleWbSort('${col}')">${formatColumnHeaderLabel(col, wbViewLevel)} <span class="sort-icon ${wbSortColumn === col ? 'active' : ''}">${getSortIconStr(col, wbSortColumn, wbSortDirection)}</span></th>`;
    });
    headerHtml += `<th class="total-column" style="cursor: pointer;" onclick="toggleWbSort('total')">小計 <span class="sort-icon ${wbSortColumn === 'total' ? 'active' : ''}">${getSortIconStr('total', wbSortColumn, wbSortDirection)}</span></th>`;
    tableHeaderRow.innerHTML = headerHtml;
    
    const columnTotals = {};
    timeColumns.forEach(col => {
        columnTotals[col] = { total: 0, po: 0, fcst: 0 };
    });
    let grandTotal = 0;
    
    let bodyHtml = "";
    rows.forEach(row => {
        let rowHtml = `
            <tr class="data-row">
                <td class="sticky-c1" title="${row.attribute}">${row.attribute}</td>
                <td class="sticky-c2" title="${row.se}">${row.se}</td>
                <td class="sticky-c3" title="${row.projectCode}">${row.projectCode}</td>
                <td class="sticky-c4" title="${row.pn}">${row.pn}</td>
                <td class="sticky-c5" title="${row.pcba}">${row.pcba}</td>
                <td class="sticky-c6" title="${row.customer}">${row.customer}</td>
        `;
        
        let rowSum = 0;
        timeColumns.forEach(col => {
            const cellRecords = row.cells[col] || [];
            let cellTotal = 0;
            let cellPO = 0;
            let cellFCST = 0;
            
            cellRecords.forEach(rec => {
                cellTotal += rec.value;
                if (rec.type === "PO") {
                    cellPO += rec.value;
                } else if (rec.type === "FCST") {
                    cellFCST += rec.value;
                }
            });
            
            rowSum += cellTotal;
            columnTotals[col].total += cellTotal;
            columnTotals[col].po += cellPO;
            columnTotals[col].fcst += cellFCST;
            
            let cellContent = "-";
            if (cellTotal > 0) {
                if (isDetailMode) {
                    let poBlock = "";
                    let fcstBlock = "";
                    if (cellPO > 0) {
                        poBlock = `<span class="detail-block po"><span class="cell-badge po">PO</span>${cellPO.toLocaleString()}</span>`;
                    }
                    if (cellFCST > 0) {
                        fcstBlock = `<span class="detail-block fcst"><span class="cell-badge fcst">FCST</span>${cellFCST.toLocaleString()}</span>`;
                    }
                    cellContent = `
                        <div class="cell-detail-container">
                            ${poBlock}
                            ${fcstBlock}
                        </div>
                    `;
                } else {
                    cellContent = cellTotal.toLocaleString();
                }
            }
            
            const cellParams = JSON.stringify({
                rowKey: `${row.attribute}||${row.se}||${row.projectCode}||${row.pn}||${row.pcba}||${row.customer}`,
                col: col,
                level: wbViewLevel
            }).replace(/"/g, '&quot;');
            
            rowHtml += `
                <td class="data-cell" onclick="showCellDetails('${cellParams}')" style="cursor: pointer;">
                    ${cellContent}
                </td>
            `;
        });
        
        grandTotal += rowSum;
        rowHtml += `<td class="total-column">${rowSum.toLocaleString()}</td></tr>`;
        bodyHtml += rowHtml;
    });
    tableBody.innerHTML = bodyHtml;
    
    let totalsHtml = `
        <td class="sticky-c1">小計總數</td>
        <td class="sticky-c2"></td>
        <td class="sticky-c3"></td>
        <td class="sticky-c4"></td>
        <td class="sticky-c5"></td>
        <td class="sticky-c6"></td>
    `;
    timeColumns.forEach(col => {
        const colVal = columnTotals[col].total;
        let colCellContent = "-";
        if (colVal > 0) {
            if (isDetailMode) {
                let poBlock = "";
                let fcstBlock = "";
                if (columnTotals[col].po > 0) {
                    poBlock = `<span class="detail-block po"><span class="cell-badge po">PO</span>${columnTotals[col].po.toLocaleString()}</span>`;
                }
                if (columnTotals[col].fcst > 0) {
                    fcstBlock = `<span class="detail-block fcst"><span class="cell-badge fcst">FCST</span>${columnTotals[col].fcst.toLocaleString()}</span>`;
                }
                colCellContent = `
                    <div class="cell-detail-container">
                        ${poBlock}
                        ${fcstBlock}
                    </div>
                `;
            } else {
                colCellContent = colVal.toLocaleString();
            }
        }
        totalsHtml += `<td class="data-cell">${colCellContent}</td>`;
    });
    totalsHtml += `<td class="total-column">${grandTotal.toLocaleString()}</td>`;
    totalsRow.innerHTML = totalsHtml;
}

// ----------------------------------------------------
// TAB 3: 差異分析邏輯 (Difference Tab)
// ----------------------------------------------------

// Load dynamic period list for difference tab based on two versions
async function initDiffPeriods() {
    if (!diffBaseVersion || !diffCompVersion) return;
    
    try {
        // Query both versions' records
        const qBase = query(collection(db, "mps_records"), where("version", "==", diffBaseVersion), where("sheet", "==", diffSheet));
        const qComp = query(collection(db, "mps_records"), where("version", "==", diffCompVersion), where("sheet", "==", diffSheet));
        
        const [snapBase, snapComp] = await Promise.all([getDocs(qBase), getDocs(qComp)]);
        
        const baseRecords = [];
        const compRecords = [];
        
        snapBase.forEach(doc => baseRecords.push(doc.data()));
        snapComp.forEach(doc => compRecords.push(doc.data()));
        
        const periodsSet = new Set();
        baseRecords.forEach(r => periodsSet.add(getTimeColumnKey(r, diffViewLevel)));
        compRecords.forEach(r => periodsSet.add(getTimeColumnKey(r, diffViewLevel)));
        
        diffPeriodsList = sortPeriodKeys(periodsSet, diffViewLevel);
        
        diffStartPeriodSelect.innerHTML = "";
        diffEndPeriodSelect.innerHTML = "";
        
        if (diffPeriodsList.length === 0) return;
        
        diffPeriodsList.forEach(period => {
            const optStart = document.createElement("option");
            optStart.value = period;
            optStart.textContent = period;
            diffStartPeriodSelect.appendChild(optStart);
            
            const optEnd = document.createElement("option");
            optEnd.value = period;
            optEnd.textContent = period;
            diffEndPeriodSelect.appendChild(optEnd);
        });
        
        diffStartPeriod = diffPeriodsList[0];
        diffEndPeriod = diffPeriodsList[diffPeriodsList.length - 1];
        
        diffStartPeriodSelect.value = diffStartPeriod;
        diffEndPeriodSelect.value = diffEndPeriod;
        
    } catch (err) {
        console.error("Error loading diff periods:", err);
    }
}

// Render Difference table grid
async function renderDiffGrid() {
    if (!diffBaseVersion || !diffCompVersion) {
        diffTableBody.innerHTML = '<tr><td colspan="2" style="text-align: center; padding: 2rem;">請選取兩個版本進行比較！</td></tr>';
        return;
    }
    
    const startIdx = diffPeriodsList.indexOf(diffStartPeriod);
    const endIdx = diffPeriodsList.indexOf(diffEndPeriod);
    
    if (startIdx === -1 || endIdx === -1 || startIdx > endIdx) {
        diffTableBody.innerHTML = '<tr><td colspan="2" style="text-align: center; padding: 2rem; color:var(--accent-po);">起始區間不能大於結束區間！</td></tr>';
        return;
    }
    
    const selectedPeriods = diffPeriodsList.slice(startIdx, endIdx + 1);
    
    try {
        // Query records for both versions
        const qBase = query(collection(db, "mps_records"), where("version", "==", diffBaseVersion), where("sheet", "==", diffSheet));
        const qComp = query(collection(db, "mps_records"), where("version", "==", diffCompVersion), where("sheet", "==", diffSheet));
        
        const [snapBase, snapComp] = await Promise.all([getDocs(qBase), getDocs(qComp)]);
        
        let recordsBase = [];
        let recordsComp = [];
        
        snapBase.forEach(d => recordsBase.push(d.data()));
        snapComp.forEach(d => recordsComp.push(d.data()));
        
        // Filter by PO / FCST Display selection
        if (diffTypeDisplayVal !== "ALL") {
            recordsBase = recordsBase.filter(r => r.type === diffTypeDisplayVal);
            recordsComp = recordsComp.filter(r => r.type === diffTypeDisplayVal);
        }
        
        // Group by Dimension (Customer or Project Code)
        const dimGroupMap = new Map();
        
        const addToMap = (recs, isBase) => {
            recs.forEach(rec => {
                const groupVal = diffGroupByDim === "customer"
                    ? (rec.customer ? rec.customer.trim() : "Unknown")
                    : (rec.projectCode ? rec.projectCode.trim() : "Unknown");
                    
                const pKey = getTimeColumnKey(rec, diffViewLevel);
                const colIdx = diffPeriodsList.indexOf(pKey);
                if (colIdx < startIdx || colIdx > endIdx) return; // out of period range
                
                if (!dimGroupMap.has(groupVal)) {
                    dimGroupMap.set(groupVal, {
                        groupKey: groupVal,
                        baseCells: {},
                        compCells: {}
                    });
                }
                
                const groupData = dimGroupMap.get(groupVal);
                if (isBase) {
                    groupData.baseCells[pKey] = (groupData.baseCells[pKey] || 0) + rec.value;
                } else {
                    groupData.compCells[pKey] = (groupData.compCells[pKey] || 0) + rec.value;
                }
            });
        };
        
        addToMap(recordsBase, true);
        addToMap(recordsComp, false);
        
        const diffRows = Array.from(dimGroupMap.values());
        if (diffSortColumn !== null && diffSortDirection !== null) {
            diffRows.sort((a, b) => {
                let valA, valB;
                if (diffSortColumn === 0) {
                    valA = a.groupKey;
                    valB = b.groupKey;
                    return diffSortDirection === 'asc'
                        ? valA.localeCompare(valB)
                        : valB.localeCompare(valA);
                } else if (diffSortColumn === 'total') {
                    valA = selectedPeriods.reduce((sum, p) => sum + ((a.baseCells[p] || 0) - (a.compCells[p] || 0)), 0);
                    valB = selectedPeriods.reduce((sum, p) => sum + ((b.baseCells[p] || 0) - (b.compCells[p] || 0)), 0);
                } else {
                    valA = (a.baseCells[diffSortColumn] || 0) - (a.compCells[diffSortColumn] || 0);
                    valB = (b.baseCells[diffSortColumn] || 0) - (b.compCells[diffSortColumn] || 0);
                }
                if (valA < valB) return diffSortDirection === 'asc' ? -1 : 1;
                if (valA > valB) return diffSortDirection === 'asc' ? 1 : -1;
                return 0;
            });
        } else {
            diffRows.sort((a, b) => a.groupKey.localeCompare(b.groupKey));
        }
        
        // Render Header
        const groupLabel = diffGroupByDim === "customer" ? "客戶 (Customer)" : "專案 (Project Code)";
        diffTableTitle.textContent = `版本差異分析：${diffBaseVersion} - ${diffCompVersion} (工作表: ${diffSheet})`;
        
        let headerHtml = `<th onclick="toggleDiffSort(0)" style="position: sticky; left: 0; z-index: 5; background-color: #161c2b; border-right: 2px solid var(--border-color); cursor: pointer;">${groupLabel} <span class="sort-icon ${diffSortColumn === 0 ? 'active' : ''}">${getSortIconStr(0, diffSortColumn, diffSortDirection)}</span></th>`;
        selectedPeriods.forEach(p => {
            headerHtml += `<th onclick="toggleDiffSort('${p}')" style="cursor: pointer;">${formatColumnHeaderLabel(p, diffViewLevel)} <span class="sort-icon ${diffSortColumn === p ? 'active' : ''}">${getSortIconStr(p, diffSortColumn, diffSortDirection)}</span></th>`;
        });
        headerHtml += `<th onclick="toggleDiffSort('total')" class="total-column" style="cursor: pointer;">小計差異 <span class="sort-icon ${diffSortColumn === 'total' ? 'active' : ''}">${getSortIconStr('total', diffSortColumn, diffSortDirection)}</span></th>`;
        diffHeaderRow.innerHTML = headerHtml;
        
        // Pre-allocate totals
        const colDiffTotals = {};
        selectedPeriods.forEach(p => { colDiffTotals[p] = 0; });
        let grandDiffTotal = 0;
        
        // Render Body
        let bodyHtml = "";
        diffRows.forEach(row => {
            let rowHtml = `<tr class="data-row"><td style="position: sticky; left: 0; z-index: 4; background-color: #111624; font-weight: 500; border-right: 2px solid var(--border-color);">${row.groupKey}</td>`;
            let rowDiffSum = 0;
            
            selectedPeriods.forEach(p => {
                const valBase = row.baseCells[p] || 0;
                const valComp = row.compCells[p] || 0;
                const diffVal = valBase - valComp;
                
                rowDiffSum += diffVal;
                colDiffTotals[p] += diffVal;
                
                rowHtml += `<td class="data-cell">${formatDiffValueHTML(diffVal)}</td>`;
            });
            
            grandDiffTotal += rowDiffSum;
            rowHtml += `<td class="total-column">${formatDiffValueHTML(rowDiffSum)}</td></tr>`;
            bodyHtml += rowHtml;
        });
        
        diffTableBody.innerHTML = bodyHtml;
        
        // Render Totals Row
        let totalsHtml = `<td style="position: sticky; left: 0; z-index: 4; background-color: #161d2b; font-weight: 600; border-right: 2px solid var(--border-color);">加總小計差異</td>`;
        selectedPeriods.forEach(p => {
            const val = colDiffTotals[p];
            totalsHtml += `<td class="data-cell">${formatDiffValueHTML(val)}</td>`;
        });
        totalsHtml += `<td class="total-column">${formatDiffValueHTML(grandDiffTotal)}</td>`;
        diffTotalsRow.innerHTML = totalsHtml;
        
    } catch (err) {
        console.error("Error rendering diff grid:", err);
        diffTableBody.innerHTML = `<tr><td colspan="2" style="text-align: center; padding: 2rem; color:#ef4444;">載入與計算比較數據時出錯: ${err.message}</td></tr>`;
    }
}

// Format Difference Value for HTML (negative values inside red parenthesis)
function formatDiffValueHTML(val) {
    if (val < 0) {
        return `<span class="neg-val">(${Math.abs(val).toLocaleString()})</span>`;
    } else if (val > 0) {
        return `<span class="pos-val">+${val.toLocaleString()}</span>`;
    } else {
        return `<span style="color:var(--text-secondary);">-</span>`;
    }
}

// ----------------------------------------------------
// TAB 4: AI comment 邏輯 (AI Comment Tab)
// ----------------------------------------------------

// Render active version selector dropdown rows dynamically
function renderAiVersionSelectors() {
    aiVersionsContainer.innerHTML = "";
    
    aiSelectedVersions.forEach((selectedVal, index) => {
        const row = document.createElement("div");
        row.className = "ai-version-row";
        
        // Create Select dropdown
        const select = document.createElement("select");
        select.style.flex = "1";
        
        // Populate options
        versionsList.forEach(ver => {
            const opt = document.createElement("option");
            opt.value = ver;
            opt.textContent = ver;
            select.appendChild(opt);
        });
        
        // Set selected value
        if (selectedVal && versionsList.includes(selectedVal)) {
            select.value = selectedVal;
        } else {
            select.value = versionsList[0] || "";
            aiSelectedVersions[index] = select.value;
        }
        
        // On change, update state
        select.addEventListener("change", (e) => {
            aiSelectedVersions[index] = e.target.value;
        });
        
        // Create Remove Button
        const removeBtn = document.createElement("button");
        removeBtn.type = "button";
        removeBtn.className = "ai-remove-btn";
        removeBtn.textContent = "移除";
        
        if (aiSelectedVersions.length <= 1) {
            removeBtn.disabled = true;
            removeBtn.style.opacity = "0.5";
            removeBtn.style.cursor = "not-allowed";
        } else {
            removeBtn.addEventListener("click", () => {
                aiSelectedVersions.splice(index, 1);
                renderAiVersionSelectors();
            });
        }
        
        row.appendChild(select);
        row.appendChild(removeBtn);
        aiVersionsContainer.appendChild(row);
    });
}

// Generate AI comment using Gemini API
async function generateAiAnalysis() {
    if (!aiApiKeyVal) {
        alert("請先於上方設定您的 Gemini API 金鑰，並點擊儲存！");
        return;
    }
    
    // Get unique non-empty versions
    const selectedVers = Array.from(new Set(aiSelectedVersions.filter(v => v !== "")));
    if (selectedVers.length === 0) {
        alert("請至少選取一個產銷計畫版本！");
        return;
    }
    
    const promptText = aiPrompt.value.trim();
    if (!promptText) {
        alert("請輸入向 AI 顧問提問的指令或分析問題！");
        return;
    }
    
    // Show loading
    document.getElementById("aiExportActions").style.display = "none";
    currentRawReportMd = "";
    aiStatusLabel.textContent = "正在撈取各版本數據並由 AI 進行分析中...";
    aiProgressBarContainer.style.display = "block";
    aiOutput.innerHTML = `
        <div style="text-align: center; padding: 2rem;">
            <div class="loading-spinner"></div>
            <p style="color: var(--text-secondary);">產銷 AI 顧問正在研讀並比較您選取的 ${selectedVers.length} 個產銷計劃版本...</p>
        </div>
    `;
    
    try {
        // 1. Ingest datasets for all selected versions
        const dataSummaries = [];
        
        for (const ver of selectedVers) {
            // Query total records for version
            const q = query(collection(db, "mps_records"), where("version", "==", ver));
            const snap = await getDocs(q);
            const records = [];
            snap.forEach(d => records.push(d.data()));
            
            // Calculate aggregations
            let totalPO = 0;
            let totalFCST = 0;
            const customerTotals = {};
            const projectTotals = {};
            const monthlyTotals = {};
            
            records.forEach(r => {
                if (r.type === "PO") totalPO += r.value;
                if (r.type === "FCST") totalFCST += r.value;
                
                // Customer Rollup
                if (r.customer) {
                    customerTotals[r.customer] = (customerTotals[r.customer] || 0) + r.value;
                }
                // Project Rollup
                if (r.projectCode) {
                    projectTotals[r.projectCode] = (projectTotals[r.projectCode] || 0) + r.value;
                }
                // Monthly Rollup
                const yr = r.year;
                const mo = String(r.month).padStart(2, '0');
                const mKey = `${yr}-${mo}`;
                monthlyTotals[mKey] = (monthlyTotals[mKey] || 0) + r.value;
            });
            
            // Get Top 3 Customers
            const topCustomers = Object.entries(customerTotals)
                .sort((a, b) => b[1] - a[1])
                .slice(0, 3)
                .map(([name, val]) => `${name}: ${val.toLocaleString()}`);
                
            // Get Top 3 Projects
            const topProjects = Object.entries(projectTotals)
                .sort((a, b) => b[1] - a[1])
                .slice(0, 3)
                .map(([code, val]) => `${code}: ${val.toLocaleString()}`);
                
            // Sorted Months
            const sortedMonths = Object.entries(monthlyTotals)
                .sort((a, b) => a[0].localeCompare(b[0]))
                .map(([month, val]) => `  - ${month}: ${val.toLocaleString()}`);
                
            dataSummaries.push({
                version: ver,
                totalQty: totalPO + totalFCST,
                poQty: totalPO,
                fcstQty: totalFCST,
                topCustomers: topCustomers,
                topProjects: topProjects,
                monthlyBreakdown: sortedMonths
            });
        }
        
        // 2. Format Context prompt for Gemini
        let dataSummaryText = "";
        dataSummaries.forEach((sum, idx) => {
            dataSummaryText += `
版本 ${idx + 1}: [${sum.version}]
- 總加總需求量: ${sum.totalQty.toLocaleString()} (PO: ${sum.poQty.toLocaleString()}, FCST: ${sum.fcstQty.toLocaleString()})
- 前三大客戶與需求量: ${sum.topCustomers.join(", ") || "無"}
- 前三大專案代碼與需求量: ${sum.topProjects.join(", ") || "無"}
- 按月份加總歷史與預估排程:
${sum.monthlyBreakdown.join("\n") || "  (無月別數據)"}
`;
        });
        
        const systemPrompt = `你是一位專業的產銷計畫 (MPS) 顧問，協助製造業與供應鏈管理者分析產銷需求變化、客戶訂單規律、預估準確性與排程波動異常。
以下是使用者選定的多個產銷計劃版本的匯總數據：
${dataSummaryText}

使用者提出的分析需求或問題：
"${promptText}"

請依據上述提供的產銷計畫數據與使用者的指示，以繁體中文撰寫一份高階且結構清晰的產銷顧問建議與分析報告。
請包含以下內容結構：
1. 【總體產銷分析】：對比選取的多個版本，分析其總量、PO (已鎖定採購量) 與 FCST (預估需求量) 的消長和增減趨勢。
2. 【異常波動與客戶/專案警示】：分析大客戶或關鍵專案在各版本間是否有大幅度的訂單拉貨或刪單波動，並明確指出哪些專案和客戶有潛在風險。
3. 【採購與產線排程顧問建議】：針對 PO 的實際執行和 FCST 預估的波動，提供關於備料策略、產能分配以及安全庫存水位等具體且可落地的戰略建議。

請使用 Markdown 格式輸出，排版要精美、段落分明，多用列表和粗體字。`;

        // 3. Post to Google Gemini API
        const url = `https://generativelanguage.googleapis.com/v1beta/models/${aiModelVal}:generateContent?key=${aiApiKeyVal}`;
        const response = await fetch(url, {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                contents: [{
                    parts: [{ text: systemPrompt }]
                }]
            })
        });
        
        if (!response.ok) {
            const errRes = await response.json();
            throw new Error(errRes.error?.message || response.statusText);
        }
        
        const resData = await response.json();
        const responseText = resData.candidates?.[0]?.content?.parts?.[0]?.text;
        
        if (!responseText) {
            throw new Error("Gemini API 未返回有效的文字內容。");
        }
        
        // 4. Render Markdown response as HTML
        aiOutput.innerHTML = parseMarkdownToHTML(responseText);
        aiStatusLabel.textContent = "分析完成";
        
        currentRawReportMd = responseText;
        document.getElementById("aiExportActions").style.display = "flex";
        
    } catch (err) {
        addConsoleLog(`AI Analysis failed: ${err.message}`, "error");
        aiStatusLabel.textContent = "分析失敗";
        aiOutput.innerHTML = `
            <div style="color: #ef4444; padding: 1.5rem; border: 1px solid rgba(239, 68, 68, 0.3); border-radius: 0.5rem; background: rgba(239, 68, 68, 0.1);">
                <h4>AI 顧問呼叫失敗！</h4>
                <p style="margin-top: 0.5rem; font-size: 0.85rem;">錯誤訊息: ${err.message}</p>
                <p style="margin-top: 0.5rem; font-size: 0.85rem; color: var(--text-secondary);">請確認您的 Gemini API 金鑰是否有效，且網路連接正常。</p>
            </div>
        `;
    } finally {
        aiProgressBarContainer.style.display = "none";
    }
}

// Client-side simple Markdown parser
function parseMarkdownToHTML(md) {
    // 1. Escape only raw comparison brackets that do not form valid HTML tags
    let escaped = md.replace(/<(?![a-zA-Z\/!])/g, "&lt;");
    
    // 2. Parse tables
    let lines = escaped.split('\n');
    let inTable = false;
    let tableRows = [];
    let processedLines = [];
    
    for (let i = 0; i < lines.length; i++) {
        let line = lines[i].trim();
        if (line.startsWith('|') && line.endsWith('|')) {
            if (!inTable) {
                inTable = true;
                tableRows = [];
            }
            tableRows.push(line);
        } else {
            if (inTable) {
                processedLines.push(renderMarkdownTableHTML(tableRows));
                inTable = false;
            }
            processedLines.push(lines[i]);
        }
    }
    if (inTable) {
        processedLines.push(renderMarkdownTableHTML(tableRows));
    }
    
    let html = processedLines.join('\n');
    
    // 3. Parse headers
    html = html.replace(/^#### (.*$)/gim, '<h4>$1</h4>');
    html = html.replace(/^### (.*$)/gim, '<h3>$1</h3>');
    html = html.replace(/^## (.*$)/gim, '<h2>$1</h2>');
    html = html.replace(/^# (.*$)/gim, '<h1>$1</h1>');
    
    // Bold
    html = html.replace(/\*\*(.*?)\*\*/gim, '<strong>$1</strong>');
    
    // Blockquotes
    html = html.replace(/^>\s+(.*$)/gim, '<blockquote>$1</blockquote>');
    
    // Inline code
    html = html.replace(/`(.*?)`/gim, '<code>$1</code>');
    
    // Lists formatting helper
    html = html.replace(/^\s*[\-\*]\s+(.*$)/gim, '<li>$1</li>');
    
    // Wrap paragraph blocks
    const blocks = html.split(/\n\n+/);
    const parsedBlocks = blocks.map(block => {
        block = block.trim();
        if (!block) return "";
        if (block.startsWith("<h") || block.startsWith("<blockquote") || block.startsWith("<li") || block.startsWith("<table") || block.startsWith("<div class=\"report-table-wrapper\"")) {
            return block;
        }
        if (block.startsWith("<li>") || block.includes("<li>")) {
            return `<ul>${block}</ul>`;
        }
        return `<p>${block.replace(/\n/g, "<br>")}</p>`;
    });
    
    return parsedBlocks.join("\n");
}

function renderMarkdownTableHTML(rows) {
    if (rows.length < 2) return rows.join('\n');
    
    let headerLine = rows[0];
    let headerCells = headerLine.split('|').slice(1, -1).map(c => c.trim());
    
    let sepLine = rows[1];
    let isSep = sepLine.split('|').slice(1, -1).every(c => /^[:\-\s]+$/.test(c));
    
    let startIdx = 1;
    if (isSep) {
        startIdx = 2;
    }
    
    let tableHtml = `<div class="report-table-wrapper"><table><thead><tr>`;
    headerCells.forEach(cell => {
        tableHtml += `<th>${cell}</th>`;
    });
    tableHtml += `</tr></thead><tbody>`;
    
    for (let i = startIdx; i < rows.length; i++) {
        let rowCells = rows[i].split('|').slice(1, -1).map(c => c.trim());
        tableHtml += `<tr>`;
        rowCells.forEach((cell, cIdx) => {
            let isNumeric = !isNaN(Number(cell.replace(/,/g, '').replace(/[\+\-\(\)]/g, '')));
            let alignStyle = isNumeric ? 'style="text-align: right;"' : '';
            tableHtml += `<td ${alignStyle}>${cell}</td>`;
        });
        tableHtml += `</tr>`;
    }
    tableHtml += `</tbody></table></div>`;
    return tableHtml;
}

// ----------------------------------------------------
// SHARED UTILITIES
// ----------------------------------------------------

function getTimeColumnKey(rec, viewLevel) {
    if (viewLevel === "weekly") {
        return rec.week;
    }
    
    const d = new Date(rec.date);
    const yr = d.getFullYear();
    const mo = String(d.getMonth() + 1).padStart(2, '0');
    
    if (viewLevel === "monthly") {
        return `${yr}-${mo}`;
    } else if (viewLevel === "quarterly") {
        const qtr = Math.floor(d.getMonth() / 3) + 1;
        return `${yr}-Q${qtr}`;
    } else {
        return `${yr}`;
    }
}

function formatColumnHeaderLabel(colKey, viewLevel) {
    if (viewLevel === "weekly") {
        const matchingRec = allRecords.find(r => r.week === colKey);
        if (matchingRec && matchingRec.date) {
            const mon = new Date(matchingRec.date);
            const sun = new Date(mon);
            sun.setDate(mon.getDate() + 6);
            const range = `${mon.getMonth() + 1}/${mon.getDate()}-${sun.getMonth() + 1}/${sun.getDate()}`;
            return `<div style="text-align:center;">${colKey}<div style="font-size:0.65rem;font-weight:400;color:var(--text-secondary);margin-top:0.15rem;">${range}</div></div>`;
        }
        return colKey;
    }
    return colKey;
}

// ----------------------------------------------------
// EXCEL FILE PARSING AND UPLOADING LOGIC (Step 7)
// ----------------------------------------------------

async function handleFileUpload(file) {
    if (!file.name.endsWith(".xlsx")) {
        alert("請選擇副檔名為 .xlsx 的 Excel 檔案！");
        return;
    }
    
    let extractedVersion = file.name.replace(".xlsx", "");
    const match = file.name.match(/WK\d+\s+\d+/i);
    if (match) {
        extractedVersion = match[0];
    } else {
        const wkMatch = file.name.match(/WK\d+/i);
        if (wkMatch) {
            extractedVersion = wkMatch[0];
        }
    }
    
    addConsoleLog(`----------------------------------------`, "info");
    addConsoleLog(`File selected: ${file.name}`, "info");
    addConsoleLog(`Parsed version label: "${extractedVersion}"`, "info");
    
    progressBarContainer.style.display = "block";
    progressBar.style.width = "0%";
    
    try {
        addConsoleLog("Reading file data into workbook...", "info");
        
        const reader = new FileReader();
        reader.onload = async (e) => {
            const arrayBuffer = e.target.result;
            
            try {
                const workbook = new ExcelJS.Workbook();
                await workbook.xlsx.load(arrayBuffer);
                
                addConsoleLog(`Workbook sheets loaded: [${workbook.worksheets.map(s => s.name).join(", ")}]`, "info");
                
                const sheetsToParse = workbook.worksheets.filter(s => s.name === "總表" || s.name === "QSI TAA");
                if (sheetsToParse.length === 0) {
                    addConsoleLog("No valid sheets found! File must contain '總表' or 'QSI TAA' sheets.", "error");
                    progressBarContainer.style.display = "none";
                    return;
                }
                
                const uploadRecords = [];
                const PO_COLORS = ["FF92D050", "FFFFFF00", "FFFF66FF"];
                const FCST_COLORS = ["FF66FFFF", "FF00FFFF", "FFFF99FF"];
                
                sheetsToParse.forEach(ws => {
                    addConsoleLog(`Parsing sheet: "${ws.name}"`, "info");
                    
                    const validCols = [];
                    for (let col = 7; col <= 84; col++) {
                        const weekCell = ws.getCell(2, col);
                        const weekVal = weekCell.value;
                        if (weekVal && String(weekVal).startsWith("WK")) {
                            validCols.push({ colNum: col, weekStr: String(weekVal) });
                        }
                    }
                    
                    addConsoleLog(`  Found ${validCols.length} valid week columns (Col 7-84)`, "info");
                    
                    const sheetKey = ws.name === "總表" ? "general" : "qsi_taa";
                    
                    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
                        if (rowNumber < 3) return;
                        
                        const attribute = ws.getCell(rowNumber, 1).value;
                        const se = ws.getCell(rowNumber, 2).value;
                        const projectCode = ws.getCell(rowNumber, 3).value;
                        const pn = ws.getCell(rowNumber, 4).value;
                        const pcba = ws.getCell(rowNumber, 5).value;
                        const customer = ws.getCell(rowNumber, 6).value;
                        
                        if (!attribute && !se && !projectCode && !pn && !pcba && !customer) return;
                        
                        validCols.forEach(col => {
                            const cell = ws.getCell(rowNumber, col.colNum);
                            const val = cell.value;
                            
                            if (val !== null && val !== undefined && !isNaN(val) && Number(val) !== 0) {
                                let colorHex = "NONE";
                                const fill = cell.fill;
                                if (fill && fill.type === "solid" && fill.fgColor && fill.fgColor.argb) {
                                    colorHex = fill.fgColor.argb;
                                }
                                
                                let cellType = "UNKNOWN";
                                if (PO_COLORS.includes(colorHex)) {
                                    cellType = "PO";
                                } else if (FCST_COLORS.includes(colorHex)) {
                                    cellType = "FCST";
                                } else {
                                    const parts = col.weekStr.replace("WK", "").split("-");
                                    const ww = parseInt(parts[0], 10);
                                    const yy = parseInt(parts[1], 10);
                                    if (yy < 26 || (yy === 26 && ww <= 26)) {
                                        cellType = "PO";
                                    } else {
                                        cellType = "FCST";
                                    }
                                }
                                
                                const mon = calculateMondayDate(col.weekStr);
                                const docId = `${extractedVersion.replace(/\s+/g, '_')}_${sheetKey}_R${rowNumber}_C${col.colNum}`;
                                
                                uploadRecords.push({
                                    id: docId,
                                    version: extractedVersion,
                                    sheet: ws.name,
                                    rowNumber: rowNumber,
                                    attribute: attribute ? String(attribute).trim() : "",
                                    se: se ? String(se).trim() : "",
                                    projectCode: projectCode ? String(projectCode).trim() : "",
                                    pn: pn ? String(pn).trim() : "",
                                    pcba: pcba ? String(pcba).trim() : "",
                                    customer: customer ? String(customer).trim() : "",
                                    week: col.weekStr,
                                    date: mon,
                                    value: Number(val),
                                    type: cellType,
                                    year: parseInt(mon.split("-")[0], 10),
                                    month: parseInt(mon.split("-")[1], 10),
                                    quarter: Math.floor((parseInt(mon.split("-")[1], 10) - 1) / 3) + 1,
                                    uploadedAt: new Date().toISOString()
                                });
                            }
                        });
                    });
                });
                
                addConsoleLog(`Parsed finished. Prepared ${uploadRecords.length} records to write.`, "success");
                await uploadInBatches(uploadRecords);
                
            } catch (err) {
                addConsoleLog(`Failed parsing Excel: ${err.message}`, "error");
                progressBarContainer.style.display = "none";
            }
        };
        
        reader.readAsArrayBuffer(file);
        
    } catch (err) {
        addConsoleLog(`FileReader error: ${err.message}`, "error");
        progressBarContainer.style.display = "none";
    }
}

function calculateMondayDate(weekStr) {
    const parts = weekStr.replace("WK", "").split("-");
    const ww = parseInt(parts[0], 10);
    const yy = parseInt(parts[1], 10);
    const year = 2000 + yy;
    
    const jan1 = new Date(year, 0, 1);
    let dayIndex = jan1.getDay();
    dayIndex = dayIndex === 0 ? 6 : dayIndex - 1;
    
    const wk1_mon = new Date(jan1);
    wk1_mon.setDate(jan1.getDate() - dayIndex);
    
    const mon = new Date(wk1_mon);
    mon.setDate(wk1_mon.getDate() + (ww - 1) * 7);
    
    const yyyy = mon.getFullYear();
    const mm = String(mon.getMonth() + 1).padStart(2, '0');
    const dd = String(mon.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
}

async function uploadInBatches(records) {
    const batchSize = 100;
    const totalRecords = records.length;
    addConsoleLog(`Uploading to Firestore in batches...`, "info");
    
    let uploadedCount = 0;
    
    for (let i = 0; i < totalRecords; i += batchSize) {
        const batchRecords = records.slice(i, i + batchSize);
        const batch = writeBatch(db);
        
        batchRecords.forEach(rec => {
            const docRef = doc(db, "mps_records", rec.id);
            const { id, ...data } = rec;
            batch.set(docRef, data);
        });
        
        try {
            await batch.commit();
            uploadedCount += batchRecords.length;
            const percentage = Math.round((uploadedCount / totalRecords) * 100);
            progressBar.style.width = `${percentage}%`;
            addConsoleLog(`Batch uploaded: ${uploadedCount}/${totalRecords} records (${percentage}%)`, "success");
        } catch (err) {
            addConsoleLog(`Firestore batch commit failed: ${err.message}`, "error");
            progressBarContainer.style.display = "none";
            return;
        }
    }
    
    // Update metadata versions list
    if (records.length > 0) {
        const newVer = records[0].version;
        try {
            const metaRef = doc(db, "metadata", "versions");
            const metaSnap = await getDoc(metaRef);
            let currentList = [];
            if (metaSnap.exists()) {
                const metaData = metaSnap.data();
                if (metaData && Array.isArray(metaData.list)) {
                    currentList = metaData.list;
                }
            }
            if (!currentList.includes(newVer)) {
                currentList.push(newVer);
                currentList.sort((a, b) => b.localeCompare(a));
                await setDoc(metaRef, { list: currentList, updatedAt: new Date().toISOString() });
                addConsoleLog(`Updated metadata versions list with new version "${newVer}".`, "success");
            }
        } catch (metaErr) {
            addConsoleLog(`Failed to update metadata document: ${metaErr.message}`, "warn");
        }
    }
    
    addConsoleLog("All records written to Firestore successfully!", "success");
    progressBarContainer.style.display = "none";
    
    await loadVersions();
    alert("檔案上傳並解析完成！已自動載入至儀表板。");
}

// ----------------------------------------------------
// MISSING HELPER FUNCTIONS & EXPORTS
// ----------------------------------------------------

function addConsoleLog(msg, type = "info") {
    if (!consoleBody) return;
    const entry = document.createElement("div");
    entry.className = `log-entry ${type}`;
    entry.textContent = `[${type.toUpperCase()}] ${msg}`;
    consoleBody.appendChild(entry);
    consoleBody.scrollTop = consoleBody.scrollHeight;
}

function setupDropdown(selectEl, valuesSet) {
    if (!selectEl) return;
    selectEl.innerHTML = '<option value="ALL">ALL (全部)</option>';
    
    const sortedVals = Array.from(valuesSet).sort((a, b) => String(a).localeCompare(String(b)));
    sortedVals.forEach(val => {
        if (val) {
            const opt = document.createElement("option");
            opt.value = val;
            opt.textContent = val;
            selectEl.appendChild(opt);
        }
    });
}

function showCellDetails(paramsStr) {
    let params;
    try {
        params = typeof paramsStr === 'string' ? JSON.parse(paramsStr) : paramsStr;
    } catch(e) {
        console.error("Failed to parse cell params", e);
        return;
    }
    const [attribute, se, projectCode, pn, pcba, customer] = params.rowKey.split("||");
    const col = params.col;
    const level = params.level;
    
    const matchingRecords = allRecords.filter(rec => {
        const matchesRow = 
            (rec.attribute || "") === attribute &&
            (rec.se || "") === se &&
            (rec.projectCode || "") === projectCode &&
            (rec.pn || "") === pn &&
            (rec.pcba || "") === pcba &&
            (rec.customer || "") === customer;
            
        if (!matchesRow) return false;
        
        const recColKey = getTimeColumnKey(rec, level);
        return recColKey === col;
    });
    
    const overlay = document.createElement("div");
    overlay.className = "modal-overlay";
    overlay.id = "detailModalOverlay";
    
    overlay.addEventListener("click", (e) => {
        if (e.target === overlay) {
            overlay.remove();
        }
    });
    
    let poTotal = 0;
    let fcstTotal = 0;
    matchingRecords.forEach(r => {
        if (r.type === "PO") poTotal += r.value;
        else if (r.type === "FCST") fcstTotal += r.value;
    });
    
    const content = document.createElement("div");
    content.className = "modal-content";
    content.innerHTML = `
        <div class="modal-header">
            <h3>產銷詳細排程明細 (${col})</h3>
            <button class="modal-close-btn" onclick="document.getElementById('detailModalOverlay').remove()">&times;</button>
        </div>
        <div class="modal-body">
            <div class="modal-info-row">
                <div class="info-item">
                    <span class="info-item-label">客戶 (Customer)</span>
                    <span class="info-item-value">${customer || "-"}</span>
                </div>
                <div class="info-item">
                    <span class="info-item-label">專案代碼 (Project Code)</span>
                    <span class="info-item-value">${projectCode || "-"}</span>
                </div>
                <div class="info-item">
                    <span class="info-item-label">料號屬性 / SE</span>
                    <span class="info-item-value">${attribute || "-"} / ${se || "-"}</span>
                </div>
                <div class="info-item">
                    <span class="info-item-label">料號 (PN) / PCBA</span>
                    <span class="info-item-value" style="word-break: break-all;">${pn || "-"} / ${pcba || "-"}</span>
                </div>
            </div>
            
            <div style="display: flex; gap: 1rem; margin: 0.5rem 0;">
                <div style="flex:1; padding: 0.5rem; background: rgba(16, 185, 129, 0.1); border-left: 3px solid #10b981; border-radius: 4px;">
                    <div style="font-size: 0.7rem; color: var(--text-secondary);">PO 小計</div>
                    <div style="font-size: 1.1rem; font-weight: 700; color: #10b981;">${poTotal.toLocaleString()}</div>
                </div>
                <div style="flex:1; padding: 0.5rem; background: rgba(59, 130, 246, 0.1); border-left: 3px solid #3b82f6; border-radius: 4px;">
                    <div style="font-size: 0.7rem; color: var(--text-secondary);">FCST 小計</div>
                    <div style="font-size: 1.1rem; font-weight: 700; color: #3b82f6;">${fcstTotal.toLocaleString()}</div>
                </div>
            </div>
            
            <div class="modal-table-container">
                <table class="modal-table">
                    <thead>
                        <tr>
                            <th>類別 (Type)</th>
                            <th>排程週別 (Week)</th>
                            <th>對應日期 (Date)</th>
                            <th style="text-align: right;">數量 (Qty)</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${matchingRecords.map(r => `
                            <tr>
                                <td>
                                    <span class="cell-badge ${r.type.toLowerCase()}">${r.type}</span>
                                </td>
                                <td>${r.week}</td>
                                <td>${r.date}</td>
                                <td class="qty" style="text-align: right; color: ${r.type === 'PO' ? '#10b981' : '#3b82f6'};">${r.value.toLocaleString()}</td>
                            </tr>
                        `).join("")}
                    </tbody>
                </table>
            </div>
        </div>
    `;
    
    overlay.appendChild(content);
    document.body.appendChild(overlay);
}

// Bind to window for global access (ES6 module workaround)
window.showCellDetails = showCellDetails;

function exportReportAsMarkdown() {
    if (!currentRawReportMd) {
        alert("無可用報告內容匯出！");
        return;
    }
    const blob = new Blob([currentRawReportMd], { type: "text/markdown;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `產銷顧問報告_${new Date().toISOString().slice(0,10)}.md`;
    a.click();
    URL.revokeObjectURL(url);
}

function exportReportAsHtml() {
    if (!currentRawReportMd) {
        alert("無可用報告內容匯出！");
        return;
    }
    
    const reportHtmlBody = parseMarkdownToHTML(currentRawReportMd);
    
    const fullHtml = `
<!DOCTYPE html>
<html lang="zh-Hant">
<head>
    <meta charset="UTF-8">
    <title>AI 智慧產銷顧問報告</title>
    <style>
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            background-color: #0b0f19;
            color: #f8fafc;
            line-height: 1.8;
            padding: 2.5rem;
            max-width: 800px;
            margin: 0 auto;
        }
        h1, h2, h3 {
            color: #ffffff;
            font-weight: 600;
            margin-top: 2rem;
            margin-bottom: 1rem;
        }
        h1 {
            border-bottom: 2px solid #3b82f6;
            padding-bottom: 0.5rem;
            font-size: 1.8rem;
        }
        h2 {
            border-left: 4px solid #10b981;
            padding-left: 0.75rem;
            font-size: 1.4rem;
        }
        h3 {
            font-size: 1.1rem;
            color: #cbd5e1;
        }
        p {
            margin-bottom: 1.25rem;
            color: #cbd5e1;
        }
        ul, ol {
            margin-bottom: 1.5rem;
            padding-left: 1.5rem;
            color: #cbd5e1;
        }
        li {
            margin-bottom: 0.5rem;
        }
        strong {
            color: #ffffff;
        }
        blockquote {
            background: rgba(255, 255, 255, 0.03);
            border-left: 4px solid #64748b;
            padding: 0.75rem 1.25rem;
            margin: 1.5rem 0;
            border-radius: 4px;
            font-style: italic;
        }
        code {
            font-family: monospace;
            background: rgba(255, 255, 255, 0.08);
            padding: 0.2rem 0.4rem;
            border-radius: 4px;
            font-size: 0.9em;
        }
        h4 {
            font-size: 1.0rem;
            color: #f8fafc;
            margin-top: 1.5rem;
            margin-bottom: 0.75rem;
            font-weight: 600;
        }
        .report-table-wrapper {
            margin: 1.5rem 0;
            overflow-x: auto;
            border: 1px solid rgba(255, 255, 255, 0.08);
            border-radius: 6px;
            background: rgba(255, 255, 255, 0.02);
        }
        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 0.9rem;
        }
        th, td {
            padding: 0.75rem 1rem;
            border-bottom: 1px solid rgba(255, 255, 255, 0.05);
            border-right: 1px solid rgba(255, 255, 255, 0.05);
        }
        th:last-child, td:last-child {
            border-right: none;
        }
        th {
            background-color: rgba(255, 255, 255, 0.05);
            font-weight: 600;
            text-align: left;
        }
        tr:last-child td {
            border-bottom: none;
        }
        .footer {
            margin-top: 3rem;
            padding-top: 1rem;
            border-top: 1px solid rgba(255,255,255,0.05);
            font-size: 0.8rem;
            color: #64748b;
            text-align: center;
        }
    </style>
</head>
<body>
    ${reportHtmlBody}
    <div class="footer">
        報告產出時間: ${new Date().toLocaleString()} | BU3 MPS 產銷計劃系統
    </div>
</body>
</html>
    `;
    
    const blob = new Blob([fullHtml], { type: "text/html;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `產銷顧問報告_${new Date().toISOString().slice(0,10)}.html`;
    a.click();
    URL.revokeObjectURL(url);
}

// ----------------------------------------------------
// TAB 1.5: WATERFALL ANALYSIS LOGIC
// ----------------------------------------------------

function renderWfVersionSelectors() {
    wfVersionsContainer.innerHTML = "";
    
    wfSelectedVersions.forEach((selectedVal, index) => {
        const row = document.createElement("div");
        row.className = "ai-version-row";
        
        const select = document.createElement("select");
        select.style.flex = "1";
        
        versionsList.forEach(ver => {
            const opt = document.createElement("option");
            opt.value = ver;
            opt.textContent = ver;
            select.appendChild(opt);
        });
        
        if (selectedVal && versionsList.includes(selectedVal)) {
            select.value = selectedVal;
        } else {
            select.value = versionsList[0] || "";
            wfSelectedVersions[index] = select.value;
        }
        
        select.addEventListener("change", (e) => {
            wfSelectedVersions[index] = e.target.value;
            initWaterfallPeriods();
            renderWaterfallGrid();
        });
        
        const removeBtn = document.createElement("button");
        removeBtn.type = "button";
        removeBtn.className = "ai-remove-btn";
        removeBtn.textContent = "移除";
        
        if (wfSelectedVersions.length <= 1) {
            removeBtn.disabled = true;
            removeBtn.style.opacity = "0.5";
            removeBtn.style.cursor = "not-allowed";
        } else {
            removeBtn.addEventListener("click", () => {
                wfSelectedVersions.splice(index, 1);
                renderWfVersionSelectors();
                initWaterfallPeriods();
                renderWaterfallGrid();
            });
        }
        
        row.appendChild(select);
        row.appendChild(removeBtn);
        wfVersionsContainer.appendChild(row);
    });
}

async function initWaterfallPeriods() {
    const selectedVers = Array.from(new Set(wfSelectedVersions.filter(v => v !== "")));
    if (selectedVers.length === 0) return;
    
    try {
        const periodsSet = new Set();
        
        for (const ver of selectedVers) {
            const cacheKey = `${ver}_${wfSheet}`;
            let records = [];
            
            if (wfRecordsCache.has(cacheKey)) {
                records = wfRecordsCache.get(cacheKey);
            } else {
                const q = query(
                    collection(db, "mps_records"),
                    where("version", "==", ver),
                    where("sheet", "==", wfSheet)
                );
                const snap = await getDocs(q);
                snap.forEach(doc => {
                    records.push(doc.data());
                });
                wfRecordsCache.set(cacheKey, records);
            }
            
            records.forEach(rec => {
                periodsSet.add(getTimeColumnKey(rec, wfViewLevel));
            });
        }
        
        wfPeriodsList = sortPeriodKeys(periodsSet, wfViewLevel);
        
        wfStartPeriodSelect.innerHTML = "";
        wfEndPeriodSelect.innerHTML = "";
        
        if (wfPeriodsList.length === 0) return;
        
        wfPeriodsList.forEach(period => {
            const optStart = document.createElement("option");
            optStart.value = period;
            optStart.textContent = period;
            wfStartPeriodSelect.appendChild(optStart);
            
            const optEnd = document.createElement("option");
            optEnd.value = period;
            optEnd.textContent = period;
            wfEndPeriodSelect.appendChild(optEnd);
        });
        
        wfStartPeriod = wfPeriodsList[0];
        wfEndPeriod = wfPeriodsList[wfPeriodsList.length - 1];
        
        wfStartPeriodSelect.value = wfStartPeriod;
        wfEndPeriodSelect.value = wfEndPeriod;
        
        populateWaterfallFilters();
        
    } catch (err) {
        console.error("Error loading waterfall periods:", err);
    }
}

function populateWaterfallFilters() {
    const selectedVers = Array.from(new Set(wfSelectedVersions.filter(v => v !== "")));
    const customers = new Set();
    const projectCodes = new Set();
    
    selectedVers.forEach(ver => {
        const cacheKey = `${ver}_${wfSheet}`;
        const records = wfRecordsCache.get(cacheKey) || [];
        records.forEach(rec => {
            if (rec.customer) customers.add(rec.customer);
            if (rec.projectCode) projectCodes.add(rec.projectCode);
        });
    });
    
    const prevCust = filterWfCustomer.value || "ALL";
    const prevProj = filterWfProjectCode.value || "ALL";
    
    setupDropdown(filterWfCustomer, customers);
    setupDropdown(filterWfProjectCode, projectCodes);
    
    if (Array.from(filterWfCustomer.options).some(o => o.value === prevCust)) {
        filterWfCustomer.value = prevCust;
        wfCustomer = prevCust;
    } else {
        filterWfCustomer.value = "ALL";
        wfCustomer = "ALL";
    }
    
    if (Array.from(filterWfProjectCode.options).some(o => o.value === prevProj)) {
        filterWfProjectCode.value = prevProj;
        wfProjectCode = prevProj;
    } else {
        filterWfProjectCode.value = "ALL";
        wfProjectCode = "ALL";
    }
}

function renderWaterfallGrid() {
    const selectedVers = Array.from(new Set(wfSelectedVersions.filter(v => v !== "")));
    if (selectedVers.length === 0) {
        wfTableBody.innerHTML = '<tr><td colspan="2" style="text-align: center; padding: 2rem;">請選取至少一個版本進行分析！</td></tr>';
        return;
    }
    
    const startIdx = wfPeriodsList.indexOf(wfStartPeriod);
    const endIdx = wfPeriodsList.indexOf(wfEndPeriod);
    
    if (startIdx === -1 || endIdx === -1 || startIdx > endIdx) {
        wfTableBody.innerHTML = '<tr><td colspan="2" style="text-align: center; padding: 2rem; color:var(--accent-po);">起始區間不能大於結束區間！</td></tr>';
        return;
    }
    
    const selectedPeriods = wfPeriodsList.slice(startIdx, endIdx + 1);
    
    let headerHtml = `<th style="position: sticky; left: 0; z-index: 5; background-color: #161c2b; border-right: 2px solid var(--border-color);">產銷計劃版本</th>`;
    selectedPeriods.forEach(p => {
        headerHtml += `<th>${formatColumnHeaderLabel(p, wfViewLevel)}</th>`;
    });
    wfHeaderRow.innerHTML = headerHtml;
    
    let bodyHtml = "";
    
    selectedVers.forEach(ver => {
        const cacheKey = `${ver}_${wfSheet}`;
        const records = wfRecordsCache.get(cacheKey) || [];
        
        const filteredRecs = records.filter(rec => {
            if (wfCustomer !== "ALL" && rec.customer !== wfCustomer) return false;
            if (wfProjectCode !== "ALL" && rec.projectCode !== wfProjectCode) return false;
            return true;
        });
        
        const periodValues = {};
        filteredRecs.forEach(rec => {
            const pKey = getTimeColumnKey(rec, wfViewLevel);
            periodValues[pKey] = (periodValues[pKey] || 0) + rec.value;
        });
        
        let rowHtml = `<tr class="data-row"><td style="position: sticky; left: 0; z-index: 4; background-color: #111624; font-weight: 500; border-right: 2px solid var(--border-color);">${ver}</td>`;
        
        selectedPeriods.forEach(p => {
            const val = periodValues[p] || 0;
            rowHtml += `<td class="data-cell">${val > 0 ? val.toLocaleString() : "-"}</td>`;
        });
        
        rowHtml += `</tr>`;
        bodyHtml += rowHtml;
    });
    
    wfTableBody.innerHTML = bodyHtml;
    wfTotalsRow.innerHTML = "";
}

// ----------------------------------------------------
// EXCEL EXPORTER UTILITY
// ----------------------------------------------------

async function exportHtmlTableToExcel(tableId, filename, sheetName) {
    const table = document.getElementById(tableId);
    if (!table) {
        alert("找不到資料表！");
        return;
    }
    
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(sheetName);
    
    // 1. Process Headers
    const thead = table.querySelector("thead");
    if (thead) {
        const headerRows = thead.querySelectorAll("tr");
        headerRows.forEach(tr => {
            const rowData = [];
            const cells = tr.querySelectorAll("th");
            cells.forEach(cell => {
                rowData.push(cell.textContent.trim());
            });
            const excelRow = worksheet.addRow(rowData);
            excelRow.eachCell(cell => {
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FF1E293B' } // Slate-800
                };
                cell.font = {
                    name: 'Segoe UI',
                    color: { argb: 'FFFFFFFF' },
                    bold: true,
                    size: 11
                };
                cell.alignment = { vertical: 'middle', horizontal: 'center' };
                cell.border = {
                    top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
                    left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
                    bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
                    right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
                };
            });
        });
    }
    
    // 2. Process Body
    const tbody = table.querySelector("tbody");
    if (tbody) {
        const bodyRows = tbody.querySelectorAll("tr");
        bodyRows.forEach(tr => {
            const rowData = [];
            const cells = tr.querySelectorAll("td");
            cells.forEach(cell => {
                let text = cell.innerText.trim();
                let isNeg = false;
                
                if (text.startsWith("(") && text.endsWith(")")) {
                    isNeg = true;
                    text = text.slice(1, -1);
                }
                if (text.startsWith("+")) {
                    text = text.slice(1);
                }
                if (text === "-") {
                    rowData.push(0);
                    return;
                }
                
                const numVal = Number(text.replace(/,/g, ""));
                if (text !== "" && !isNaN(numVal)) {
                    rowData.push(isNeg ? -numVal : numVal);
                } else {
                    rowData.push(text);
                }
            });
            
            const excelRow = worksheet.addRow(rowData);
            excelRow.eachCell((cell, cIdx) => {
                cell.border = {
                    top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
                    left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
                    bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
                    right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
                };
                
                if (typeof cell.value === 'number') {
                    if (cell.value < 0) {
                        cell.font = { name: 'Segoe UI', size: 10, color: { argb: 'FFEF4444' } }; // Red for negatives
                        cell.numFmt = '#,##0;[Red](#,##0);"-"';
                    } else if (cell.value === 0) {
                        cell.font = { name: 'Segoe UI', size: 10, color: { argb: 'FF94A3B8' } };
                        cell.value = "-";
                    } else {
                        cell.font = { name: 'Segoe UI', size: 10 };
                        cell.numFmt = '#,##0;(#,##0);"-"';
                    }
                    cell.alignment = { horizontal: 'right' };
                } else {
                    cell.font = { name: 'Segoe UI', size: 10 };
                    cell.alignment = { horizontal: 'left' };
                }
            });
        });
    }
    
    // 3. Process Footer
    const tfoot = table.querySelector("tfoot");
    if (tfoot) {
        const footerRows = tfoot.querySelectorAll("tr");
        footerRows.forEach(tr => {
            const rowData = [];
            const cells = tr.querySelectorAll("td, th");
            cells.forEach(cell => {
                let text = cell.innerText.trim();
                let isNeg = false;
                
                if (text.startsWith("(") && text.endsWith(")")) {
                    isNeg = true;
                    text = text.slice(1, -1);
                }
                if (text === "-") {
                    rowData.push(0);
                    return;
                }
                
                const numVal = Number(text.replace(/,/g, ""));
                if (text !== "" && !isNaN(numVal)) {
                    rowData.push(isNeg ? -numVal : numVal);
                } else {
                    rowData.push(text);
                }
            });
            
            const excelRow = worksheet.addRow(rowData);
            excelRow.eachCell(cell => {
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFF1F5F9' } // Grey-100
                };
                cell.border = {
                    top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
                    left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
                    bottom: { style: 'double', color: { argb: 'FF1E293B' } }, // double bottom line
                    right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
                };
                
                if (typeof cell.value === 'number') {
                    if (cell.value < 0) {
                        cell.font = { name: 'Segoe UI', bold: true, size: 11, color: { argb: 'FFEF4444' } };
                        cell.numFmt = '#,##0;[Red](#,##0);"-"';
                    } else if (cell.value === 0) {
                        cell.font = { name: 'Segoe UI', bold: true, size: 11, color: { argb: 'FF94A3B8' } };
                        cell.value = "-";
                    } else {
                        cell.font = { name: 'Segoe UI', bold: true, size: 11 };
                        cell.numFmt = '#,##0;(#,##0);"-"';
                    }
                    cell.alignment = { horizontal: 'right' };
                } else {
                    cell.font = { name: 'Segoe UI', bold: true, size: 11 };
                    cell.alignment = { horizontal: 'left' };
                }
            });
        });
    }
    
    // Auto-fit columns
    worksheet.columns.forEach(col => {
        let maxLen = 12;
        col.eachCell(cell => {
            let val = "";
            if (cell.value !== null && cell.value !== undefined) {
                if (typeof cell.value === 'number') {
                    val = cell.value.toLocaleString();
                } else {
                    val = String(cell.value);
                }
            }
            if (val.length > maxLen) maxLen = val.length;
        });
        col.width = maxLen + 4;
    });
    
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
}

// ----------------------------------------------------
// COLUMN SORTING HELPERS
// ----------------------------------------------------

function getSortIconStr(colId, currentSortCol, currentSortDir) {
    if (colId === currentSortCol) {
        if (currentSortDir === 'asc') return '▲';
        if (currentSortDir === 'desc') return '▼';
    }
    return '↕';
}

function toggleSummarySort(colId) {
    if (summarySortColumn === colId) {
        if (summarySortDirection === 'asc') {
            summarySortDirection = 'desc';
        } else if (summarySortDirection === 'desc') {
            summarySortColumn = null;
            summarySortDirection = null;
        }
    } else {
        summarySortColumn = colId;
        summarySortDirection = 'asc';
    }
    renderSummaryGrid();
}

function toggleWbSort(colId) {
    if (wbSortColumn === colId) {
        if (wbSortDirection === 'asc') {
            wbSortDirection = 'desc';
        } else if (wbSortDirection === 'desc') {
            wbSortColumn = null;
            wbSortDirection = null;
        }
    } else {
        wbSortColumn = colId;
        wbSortDirection = 'asc';
    }
    renderWhiteboardGrid();
}

function toggleDiffSort(colId) {
    if (diffSortColumn === colId) {
        if (diffSortDirection === 'asc') {
            diffSortDirection = 'desc';
        } else if (diffSortDirection === 'desc') {
            diffSortColumn = null;
            diffSortDirection = null;
        }
    } else {
        diffSortColumn = colId;
        diffSortDirection = 'asc';
    }
    renderDiffGrid();
}

// Bind to window for global access (ES Module scope compatibility)
window.toggleSummarySort = toggleSummarySort;
window.toggleWbSort = toggleWbSort;
window.toggleDiffSort = toggleDiffSort;


