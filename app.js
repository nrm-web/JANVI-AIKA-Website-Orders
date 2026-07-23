// Live Dashboard JavaScript Controller

// Application State
const state = {
    apiUrl: '',
    rawSheets: null,
    orders: [],
    filteredOrders: [],
    monthFilteredOrders: [],
    currentPage: 1,
    pageSize: 15,
    dateRangeFrom: null, // format: "YYYY-MM-DD"
    dateRangeTo: null,   // format: "YYYY-MM-DD"
    pickerLeftMonth: (() => {
        const today = new Date();
        return new Date(today.getFullYear(), today.getMonth() - 1, 1);
    })(),
    hoveredDate: null,
    charts: {
        salesTrend: null,
        dailyTrend: null,
        finance: null,
        delivery: null
    }
};

// DOM Elements
const elements = {
    apiUrlInput: document.getElementById('api-url-input'),
    saveUrlBtn: document.getElementById('save-url-btn'),
    syncStatus: document.getElementById('sync-status'),
    
    // KPIs
    kpiOrders: document.getElementById('kpi-total-orders'),
    kpiRevenue: document.getElementById('kpi-total-revenue'),
    kpiReturnedCount: document.getElementById('kpi-returned-count'),
    kpiRefunded: document.getElementById('kpi-total-refunded'),
    kpiSuccessfulCount: document.getElementById('kpi-successful-count'),
    kpiProfit: document.getElementById('kpi-total-profit'),
    kpiReturnRate: document.getElementById('kpi-return-rate'),
    kpiDenialRate: document.getElementById('kpi-denial-rate'),
    kpiReturnProgress: document.getElementById('kpi-return-progress'),
    kpiDenialProgress: document.getElementById('kpi-denial-progress'),
    kpiReturnRateSub: document.getElementById('kpi-return-rate-sub'),
    kpiDenialRateSub: document.getElementById('kpi-denial-rate-sub'),
    kpiCanceledCount: document.getElementById('kpi-canceled-count'),
    kpiCanceledCountSub: document.getElementById('kpi-canceled-count-sub'),
    kpiTotalCanceled: document.getElementById('kpi-total-canceled'),
    kpiTotalCanceledSub: document.getElementById('kpi-total-canceled-sub'),
    
    // Pipeline
    pipeUnfulfilled: document.querySelector('#step-unfulfilled .step-count'),
    pipeReturned: document.querySelector('#step-returned .step-count'),
    pipePickup: document.querySelector('#step-pickup .step-count'),
    pipeTransit: document.querySelector('#step-transit .step-count'),
    pipeDelivered: document.querySelector('#step-delivered .step-count'),
    pipeCanceled: document.querySelector('#step-canceled .step-count'),
    
    // Filters & Table
    searchInput: document.getElementById('search-input'),
    topFilterMonth: document.getElementById('top-filter-month'),
    dateRangeBtn: document.getElementById('date-range-btn'),
    datePickerPopup: document.getElementById('date-picker-popup'),
    dpPrevMonth: document.getElementById('dp-prev-month'),
    dpNextMonth: document.getElementById('dp-next-month'),
    dpLeftMonthName: document.getElementById('dp-left-month-name'),
    dpRightMonthName: document.getElementById('dp-right-month-name'),
    dpLeftDays: document.getElementById('dp-left-days'),
    dpRightDays: document.getElementById('dp-right-days'),
    filterMonth: document.getElementById('filter-month'),
    chartMonthSelect: document.getElementById('chart-month-select'),
    filterPayment: document.getElementById('filter-payment'),
    filterCategory: document.getElementById('filter-category'),
    filterStatus: document.getElementById('filter-status'),
    pageSizeSelect: document.getElementById('rows-per-page'),
    clearFiltersBtn: document.getElementById('clear-filters-btn'),
    headerClearBtn: document.getElementById('header-clear-btn'),
    searchClearBtn: document.getElementById('search-clear-btn'),
    ordersTbody: document.getElementById('orders-tbody'),
    tableCount: document.getElementById('table-record-count'),
    prevPageBtn: document.getElementById('prev-page-btn'),
    nextPageBtn: document.getElementById('next-page-btn'),
    paginationInfo: document.getElementById('pagination-numbers'),
    
    // Modal
    modal: document.getElementById('order-modal'),
    modalTitle: document.getElementById('modal-order-title'),
    modalCloseBtn: document.getElementById('modal-close-btn')
};

// Initial Setup
window.addEventListener('DOMContentLoaded', () => {
    const savedTheme = localStorage.getItem('theme') || 'light';
    document.body.classList.add(savedTheme + '-mode');
    document.body.classList.remove(savedTheme === 'dark' ? 'light-mode' : 'dark-mode');
    
    loadConfiguration();
    setupEventListeners();
});

async function loadConfiguration() {
    // 0. Try local offline data first (bypasses ALL network requests and CORS limits)
    if (window.DASHBOARD_DATA && window.DASHBOARD_DATA.sheets) {
        state.rawSheets = window.DASHBOARD_DATA.sheets;
        parseData();
        updateStatus('synced', 'Local Sync Mode (Offline)');
        // Hide API input container to look clean
        const configContainer = document.querySelector('.url-config-container');
        if (configContainer) configContainer.style.display = 'none';
        return;
    }

    let hasApi = false;

    // 1. Try global window.APP_CONFIG loaded via script tag (bypasses local file CORS)
    if (window.APP_CONFIG && window.APP_CONFIG.google_web_app_url) {
        const scriptUrl = window.APP_CONFIG.google_web_app_url;
        if (scriptUrl && !scriptUrl.includes('YOUR_URL') && !scriptUrl.includes('YOUR_GOOGLE_WEB_APP_URL_HERE')) {
            state.apiUrl = scriptUrl;
            elements.apiUrlInput.value = scriptUrl;
            hasApi = true;
        }
    }

    // 2. Try local storage
    if (!hasApi) {
        const savedUrl = localStorage.getItem('google_web_app_url');
        if (savedUrl) {
            state.apiUrl = savedUrl;
            elements.apiUrlInput.value = savedUrl;
            hasApi = true;
        }
    }

    // 3. Try config.json on the server
    if (!hasApi) {
        try {
            const response = await fetch('config.json');
            if (response.ok) {
                const config = await response.json();
                const configUrl = config.google_web_app_url;
                if (configUrl && !configUrl.includes('YOUR_URL') && !configUrl.includes('YOUR_GOOGLE_WEB_APP_URL_HERE')) {
                    state.apiUrl = configUrl;
                    elements.apiUrlInput.value = configUrl;
                    hasApi = true;
                }
            }
        } catch (e) {
            // config.json fetch might fail due to local file protocol CORS, which is expected
        }
    }

    // If we have an API URL, load live data from Google Sheets Web App
    if (hasApi) {
        updateStatus('loading', 'Loading data from Google Sheet...');
        fetchData();
        return;
    }

    updateStatus('error', 'API URL not set. Please paste it above.');
}

// Setup listeners
function setupEventListeners() {
    const themeToggle = document.getElementById('theme-toggle');
    if (themeToggle) {
        themeToggle.addEventListener('click', () => {
            const isDark = document.body.classList.contains('dark-mode');
            if (isDark) {
                document.body.classList.remove('dark-mode');
                document.body.classList.add('light-mode');
                localStorage.setItem('theme', 'light');
            } else {
                document.body.classList.remove('light-mode');
                document.body.classList.add('dark-mode');
                localStorage.setItem('theme', 'dark');
            }
            updateChartsTheme();
        });
    }

    elements.saveUrlBtn.addEventListener('click', () => {
        const url = elements.apiUrlInput.value.trim();
        if (url) {
            state.apiUrl = url;
            localStorage.setItem('google_web_app_url', url);
            updateStatus('loading', 'Fetching sheet data...');
            fetchData();
        }
    });

    if (elements.searchInput) {
        elements.searchInput.addEventListener('input', (e) => {
            if (elements.searchClearBtn) {
                elements.searchClearBtn.style.display = e.target.value.length > 0 ? 'block' : 'none';
            }
            applyFilters();
        });
    }
    if (elements.searchClearBtn) {
        elements.searchClearBtn.addEventListener('click', () => {
            elements.searchInput.value = '';
            elements.searchClearBtn.style.display = 'none';
            applyFilters();
        });
    }
    elements.topFilterMonth.addEventListener('change', () => {
        state.dateRangeFrom = null;
        state.dateRangeTo = null;
        state.hoveredDate = null;
        if (elements.dateRangeBtn) elements.dateRangeBtn.textContent = 'Select Date Range';
        elements.filterMonth.value = elements.topFilterMonth.value;
        if (elements.chartMonthSelect) elements.chartMonthSelect.value = elements.topFilterMonth.value;
        applyFilters();
    });
    elements.filterMonth.addEventListener('change', () => {
        state.dateRangeFrom = null;
        state.dateRangeTo = null;
        state.hoveredDate = null;
        if (elements.dateRangeBtn) elements.dateRangeBtn.textContent = 'Select Date Range';
        elements.topFilterMonth.value = elements.filterMonth.value;
        if (elements.chartMonthSelect) elements.chartMonthSelect.value = elements.filterMonth.value;
        applyFilters();
    });
    if (elements.chartMonthSelect) {
        elements.chartMonthSelect.addEventListener('change', () => {
            const val = elements.chartMonthSelect.value;
            state.dateRangeFrom = null;
            state.dateRangeTo = null;
            state.hoveredDate = null;
            if (elements.dateRangeBtn) elements.dateRangeBtn.textContent = 'Select Date Range';
            elements.topFilterMonth.value = val;
            elements.filterMonth.value = val;
            applyFilters();
        });
    }
    
    initDatePicker();
    
    elements.filterPayment.addEventListener('change', applyFilters);
    elements.filterCategory.addEventListener('change', applyFilters);
    elements.filterStatus.addEventListener('change', applyFilters);
    if (elements.pageSizeSelect) {
        elements.pageSizeSelect.addEventListener('change', () => {
            const val = elements.pageSizeSelect.value;
            if (val === 'all') {
                state.pageSize = 999999;
            } else {
                state.pageSize = parseInt(val, 10);
            }
            state.currentPage = 1;
            renderTable();
        });
    }
    elements.clearFiltersBtn.addEventListener('click', clearFilters);
    elements.headerClearBtn.addEventListener('click', () => {
        elements.topFilterMonth.value = '';
        elements.filterMonth.value = '';
        if (elements.chartMonthSelect) elements.chartMonthSelect.value = '';
        state.dateRangeFrom = null;
        state.dateRangeTo = null;
        state.hoveredDate = null;
        if (elements.dateRangeBtn) elements.dateRangeBtn.textContent = 'Select Date Range';
        applyFilters();
    });
    
    // Pagination event listeners
    elements.prevPageBtn.addEventListener('click', () => {
        if (state.currentPage > 1) {
            state.currentPage--;
            renderTable();
        }
    });

    elements.nextPageBtn.addEventListener('click', () => {
        const totalPages = Math.ceil(state.filteredOrders.length / state.pageSize) || 1;
        if (state.currentPage < totalPages) {
            state.currentPage++;
            renderTable();
        }
    });
    
    elements.modalCloseBtn.addEventListener('click', () => {
        elements.modal.style.display = 'none';
    });
    
    // Close modal on click outside
    window.addEventListener('click', (e) => {
        if (e.target === elements.modal) {
            elements.modal.style.display = 'none';
        }
    });
}

// Update UI status dot
function updateStatus(type, text) {
    elements.syncStatus.className = `sync-indicator ${type}`;
    elements.syncStatus.querySelector('.status-text').textContent = text;
}

// Fetch spreadsheet JSON via Apps Script doGet using JSONP (bypasses browser file:// CORS redirect limits)
function fetchData() {
    if (!state.apiUrl) return;
    
    // Generate a unique callback function name
    const callbackName = 'jsonpCallback_' + Math.round(Math.random() * 1000000);
    
    // Define callback function globally
    window[callbackName] = function(res) {
        // Clean up script tag and callback function
        delete window[callbackName];
        const script = document.getElementById(callbackName);
        if (script) script.remove();
        
        if (res.status === 'success') {
            state.rawSheets = res.data.sheets;
            parseData();
            updateStatus('synced', 'Connected & Synced');
        } else {
            updateStatus('error', `Sync failed: ${res.message}`);
        }
    };
    
    // Create script element
    const script = document.createElement('script');
    script.id = callbackName;
    
    // Append callback name to url
    const separator = state.apiUrl.includes('?') ? '&' : '?';
    script.src = state.apiUrl + separator + 'prefix=' + callbackName;
    
    // Error handler if loading fails
    script.onerror = function() {
        updateStatus('error', 'Failed to reach endpoint. Verify URL / Deployments.');
        delete window[callbackName];
        script.remove();
    };
    
    // Append script tag to execute JSONP load
    document.body.appendChild(script);
}

// Classify an order into one of the 6 logistics pipeline stages
function getPipelineStage(o) {
    const status = (o.logisticsStatus || '').toUpperCase().trim();
    
    if (status.includes('CANCELED') || status.includes('CANCELLED')) {
        return 'canceled';
    } else if (status === 'DELIVERED' || status === 'SELF FULFILED') {
        return 'delivered';
    } else if (o.returned || status.includes('RTO')) {
        return 'returned';
    } else if (status.includes('TRANSIT') || status.includes('PICKED UP') || status.includes('DELIVERY') || status.includes('HUB') || status.includes('SHIPPED') || status.includes('UNDELIVERED')) {
        return 'transit';
    } else if (status.includes('PICKUP') || status.includes('READY TO SHIP')) {
        if (status.includes('EXCEPTION')) {
            return 'unfulfilled';
        } else {
            return 'pickup';
        }
    } else {
        return 'unfulfilled';
    }
}

// Convert sheet rows to structured javascript objects
function parseData() {
    const masterRows = state.rawSheets['Master Sheet'];
    if (!masterRows || masterRows.length < 2) {
        elements.ordersTbody.innerHTML = `<tr><td colspan="11" class="empty-state">No master data found in sheet.</td></tr>`;
        return;
    }
    
    const headers = masterRows[0];
    const dataRows = masterRows.slice(1);
    
    state.orders = [];
    
    dataRows.forEach(row => {
        // Skip total/summary row at the bottom
        if (!row[0] || row[0] === 'Total Orders' || row[0] === 'Total Revenue') return;
        
        const isReturned = String(row[8]).toLowerCase() === 'true' || row[8] === true;
        const isPrepaid = String(row[6]).trim().toLowerCase() === 'yes' || row[6] === 'Yes';
        
        // Infer financial status
        let financialStatus = 'pending';
        if (isReturned) {
            financialStatus = 'refunded';
        } else if (isPrepaid) {
            financialStatus = 'paid';
        }
        
        // Infer Shopify fulfillment status
        const logisticsStatus = String(row[13] || 'NEW ORDER').toUpperCase().trim();
        let fulfillmentStatus = 'fulfilled';
        if (logisticsStatus === 'NEW ORDER' || logisticsStatus === 'UNFULFILLED' || logisticsStatus === '') {
            fulfillmentStatus = 'unfulfilled';
        }
        
        state.orders.push({
            orderNo: row[0],
            customerName: row[1],
            itemsOrdered: row[2],
            totalPrice: parseFloat(row[4]) || 0, // Index 4 is Total Price
            dateOfOrder: row[3],                 // Index 3 is Date of Order
            paymentMethod: row[5],
            prepaid: row[6],
            cod: row[7],
            returned: isReturned,
            codDenies: row[9],
            shiprocketComments: row[10],
            city: row[11],
            pinCode: String(row[12]),
            financialStatus: financialStatus,
            fulfillmentStatus: fulfillmentStatus,
            feedbackSent: row[14],
            feedbackReceived: row[15],
            awbCode: row[16],
            trackingLink: row[17],
            logisticsStatus: logisticsStatus,
            sku: row[18] || '-',
            category: row[19] || '-'
        });
    });
    
    // Sort orders by date descending
    state.orders.sort((a, b) => new Date(b.dateOfOrder) - new Date(a.dateOfOrder));
    state.filteredOrders = [...state.orders];
    state.monthFilteredOrders = [...state.orders];
    
    // Populate filter select options
    populateFilterOptions();
    
    // Populate Dashboard and charts
    renderDashboard();
}

// Extract unique months and logistics statuses for filter dropdowns
function populateFilterOptions() {
    const months = new Set();
    const statuses = new Set();
    const categories = new Set();
    
    state.orders.forEach(o => {
        if (o.dateOfOrder) {
            months.add(getMonthYearStr(o.dateOfOrder));
        }
        if (o.logisticsStatus) {
            let statusText = o.logisticsStatus.toUpperCase().trim();
            if (statusText.includes('UNDELIVERED')) {
                statuses.add('UNDELIVERED');
            } else {
                statuses.add(statusText);
            }
        }
        if (o.category) {
            o.category.split(',').forEach(c => {
                const clean = c.trim().toUpperCase();
                if (clean && clean !== '-' && clean !== 'OTHER') {
                    categories.add(clean);
                }
            });
        }
    });
    
    // Populate Months
    elements.filterMonth.innerHTML = '<option value="">All Months</option>';
    elements.topFilterMonth.innerHTML = '<option value="">All Months</option>';
    if (elements.chartMonthSelect) {
        elements.chartMonthSelect.innerHTML = '<option value="">All Months</option>';
    }
    // Chronological order (descending): compile list of months and sort
    const chronMonths = [
        "Mar 2027", "Feb 2027", "Jan 2027", "Dec 2026", "Nov 2026", 
        "Oct 2026", "Sep 2026", "Aug 2026", "Jul 2026", "Jun 2026"
    ];
    chronMonths.forEach(m => {
        if (months.has(m)) {
            const opt = `<option value="${m}">${m}</option>`;
            elements.filterMonth.innerHTML += opt;
            elements.topFilterMonth.innerHTML += opt;
            if (elements.chartMonthSelect) {
                elements.chartMonthSelect.innerHTML += opt;
            }
        }
    });
    
    // No dynamic population of statuses dropdown, since we use the 6 consolidated pipeline options

    // Populate Categories
    elements.filterCategory.innerHTML = '<option value="">All Categories</option>';
    Array.from(categories).sort().forEach(cat => {
        elements.filterCategory.innerHTML += `<option value="${cat}">${cat}</option>`;
    });
}

// Convert "2026-07-02" to "Jul 2026"
function getMonthYearStr(dateStr) {
    if (!dateStr) return '';
    const parts = dateStr.split('-');
    if (parts.length < 2) return '';
    const year = parts[0];
    const monthIndex = parseInt(parts[1], 10) - 1;
    const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    return `${months[monthIndex]} ${year}`;
}

// Format numbers to Rupees
function formatCurrency(amount) {
    return new Intl.NumberFormat('en-IN', {
        style: 'currency',
        currency: 'INR',
        maximumFractionDigits: 2
    }).format(amount);
}

// Render KPIs and visual metrics
function renderDashboard() {
    // 1. Calculate general stats
    const canceledList = state.monthFilteredOrders.filter(o => {
        const s = o.logisticsStatus.toUpperCase().trim();
        return s.includes('CANCELED') || s.includes('CANCELLED');
    });
    const canceledCount = canceledList.length;
    const totalCanceledAmount = canceledList.reduce((sum, o) => sum + o.totalPrice, 0);

    const totalOrders = state.monthFilteredOrders.length;
    const totalRevenue = state.monthFilteredOrders.reduce((sum, o) => sum + o.totalPrice, 0);
    const totalRefunded = state.monthFilteredOrders.filter(o => {
        const s = o.logisticsStatus.toUpperCase().trim();
        return o.returned && !s.includes('CANCELED') && !s.includes('CANCELLED') && s !== 'DELIVERED' && s !== 'SELF FULFILED';
    }).reduce((sum, o) => sum + o.totalPrice, 0);
    const totalProfit = totalRevenue - totalRefunded - totalCanceledAmount;
    
    const returnCount = state.monthFilteredOrders.filter(o => {
        const s = o.logisticsStatus.toUpperCase().trim();
        return o.returned && !s.includes('CANCELED') && !s.includes('CANCELLED') && s !== 'DELIVERED' && s !== 'SELF FULFILED';
    }).length;
    const returnRate = totalOrders > 0 ? (returnCount / totalOrders) * 100 : 0;
    const successfulCount = state.monthFilteredOrders.filter(o => {
        const status = o.logisticsStatus.toUpperCase().trim();
        return status === 'DELIVERED' || status === 'SELF FULFILED';
    }).length;
    
    const codOrders = state.monthFilteredOrders.filter(o => o.paymentMethod === 'COD');
    const codDenials = codOrders.filter(o => o.codDenies === 'Yes').length;
    const denialRate = codOrders.length > 0 ? (codDenials / codOrders.length) * 100 : 0;
    
    // Write values
    elements.kpiOrders.textContent = totalOrders;
    elements.kpiRevenue.textContent = formatCurrency(totalRevenue);
    elements.kpiReturnedCount.textContent = returnCount;
    elements.kpiRefunded.textContent = formatCurrency(totalRefunded);
    elements.kpiSuccessfulCount.textContent = successfulCount;
    elements.kpiProfit.textContent = formatCurrency(totalProfit);
    elements.kpiReturnRate.textContent = `${returnRate.toFixed(1)}%`;
    elements.kpiDenialRate.textContent = `${denialRate.toFixed(1)}%`;
    
    elements.kpiReturnRateSub.textContent = `${returnCount} returned of ${totalOrders}`;
    elements.kpiDenialRateSub.textContent = `${codDenials} denials of ${codOrders.length} COD`;
    
    elements.kpiCanceledCount.textContent = canceledCount;
    elements.kpiCanceledCountSub.textContent = `${canceledCount} canceled of ${totalOrders}`;
    elements.kpiTotalCanceled.textContent = formatCurrency(totalCanceledAmount);
    elements.kpiTotalCanceledSub.textContent = `From ${canceledCount} canceled`;
    
    // Progress bar animations
    elements.kpiReturnProgress.style.width = `${returnRate}%`;
    elements.kpiDenialProgress.style.width = `${denialRate}%`;
    
    // 2. Render Logistics Pipeline counts
    const pipeCounts = {
        unfulfilled: 0,
        returned: 0,
        pickup: 0,
        transit: 0,
        delivered: 0,
        canceled: 0
    };
    
    state.monthFilteredOrders.forEach(o => {
        const stage = getPipelineStage(o);
        if (pipeCounts.hasOwnProperty(stage)) {
            pipeCounts[stage]++;
        }
    });
    
    elements.pipeUnfulfilled.textContent = pipeCounts.unfulfilled;
    elements.pipeReturned.textContent = pipeCounts.returned;
    elements.pipePickup.textContent = pipeCounts.pickup;
    elements.pipeTransit.textContent = pipeCounts.transit;
    elements.pipeDelivered.textContent = pipeCounts.delivered;
    elements.pipeCanceled.textContent = pipeCounts.canceled;
    
    // Animate pipeline states
    document.querySelectorAll('.pipeline-step').forEach(step => {
        const id = step.id.replace('step-', '');
        if (pipeCounts[id] > 0) {
            step.classList.add('active');
        } else {
            step.classList.remove('active');
        }
    });

    // 2.5 Render Mode of Payment Breakdown
    renderPaymentModeBreakdown();

    // 3. Render charts
    renderCharts();
    
    // 4. Render Table
    renderTable();
    
    // 5. Render Category Breakdown
    renderCategoryBreakdown();
}

// Render Mode of Payment breakdown (COD & Prepaid cards)
function renderPaymentModeBreakdown() {
    const totalMonthOrders = state.monthFilteredOrders.length;
    
    // COD Orders
    const codOrders = state.monthFilteredOrders.filter(o => o.paymentMethod === 'COD' || o.cod === 'Yes');
    const codTotal = codOrders.length;
    const codTotalVal = codOrders.reduce((sum, o) => sum + o.totalPrice, 0);
    const codShare = totalMonthOrders > 0 ? (codTotal / totalMonthOrders) * 100 : 0;
    
    const codCounts = { delivered: 0, transit: 0, pickup: 0, unfulfilled: 0, returned: 0, canceled: 0 };
    codOrders.forEach(o => {
        const stage = getPipelineStage(o);
        if (codCounts.hasOwnProperty(stage)) codCounts[stage]++;
    });
    
    // Prepaid Orders
    const prepaidOrders = state.monthFilteredOrders.filter(o => o.paymentMethod === 'Prepaid' || o.prepaid === 'Yes');
    const prepaidTotal = prepaidOrders.length;
    const prepaidTotalVal = prepaidOrders.reduce((sum, o) => sum + o.totalPrice, 0);
    const prepaidShare = totalMonthOrders > 0 ? (prepaidTotal / totalMonthOrders) * 100 : 0;
    
    const prepaidCounts = { delivered: 0, transit: 0, pickup: 0, unfulfilled: 0, returned: 0, canceled: 0 };
    prepaidOrders.forEach(o => {
        const stage = getPipelineStage(o);
        if (prepaidCounts.hasOwnProperty(stage)) prepaidCounts[stage]++;
    });
    
    // Calculate Pcts
    const getPct = (cnt, tot) => tot > 0 ? ((cnt / tot) * 100).toFixed(1) : '0.0';
    
    // Update DOM Elements
    const setTxt = (id, txt) => {
        const el = document.getElementById(id);
        if (el) el.textContent = txt;
    };
    const setWidth = (id, pct) => {
        const el = document.getElementById(id);
        if (el) el.style.width = `${pct}%`;
    };
    
    // COD Card DOM values
    setTxt('payment-cod-share', `${codShare.toFixed(1)}% of total`);
    setTxt('payment-cod-total', codTotal);
    setTxt('payment-cod-total-val', formatCurrency(codTotalVal));
    setTxt('payment-cod-delivered', codCounts.delivered);
    setTxt('payment-cod-delivered-pct', `${getPct(codCounts.delivered, codTotal)}%`);
    setTxt('payment-cod-transit', codCounts.transit);
    setTxt('payment-cod-transit-pct', `${getPct(codCounts.transit, codTotal)}%`);
    setTxt('payment-cod-pickup', codCounts.pickup);
    setTxt('payment-cod-pickup-pct', `${getPct(codCounts.pickup, codTotal)}%`);
    setTxt('payment-cod-unfulfilled', codCounts.unfulfilled);
    setTxt('payment-cod-unfulfilled-pct', `${getPct(codCounts.unfulfilled, codTotal)}%`);
    setTxt('payment-cod-returned', codCounts.returned);
    setTxt('payment-cod-returned-pct', `${getPct(codCounts.returned, codTotal)}%`);
    setTxt('payment-cod-canceled', codCounts.canceled);
    setTxt('payment-cod-canceled-pct', `${getPct(codCounts.canceled, codTotal)}%`);
    
    const codReturnRate = getPct(codCounts.returned, codTotal);
    setTxt('payment-cod-progress-text', `${codReturnRate}%`);
    setWidth('payment-cod-progress-bar', parseFloat(codReturnRate));
    
    // Prepaid Card DOM values
    setTxt('payment-prepaid-share', `${prepaidShare.toFixed(1)}% of total`);
    setTxt('payment-prepaid-total', prepaidTotal);
    setTxt('payment-prepaid-total-val', formatCurrency(prepaidTotalVal));
    setTxt('payment-prepaid-delivered', prepaidCounts.delivered);
    setTxt('payment-prepaid-delivered-pct', `${getPct(prepaidCounts.delivered, prepaidTotal)}%`);
    setTxt('payment-prepaid-transit', prepaidCounts.transit);
    setTxt('payment-prepaid-transit-pct', `${getPct(prepaidCounts.transit, prepaidTotal)}%`);
    setTxt('payment-prepaid-pickup', prepaidCounts.pickup);
    setTxt('payment-prepaid-pickup-pct', `${getPct(prepaidCounts.pickup, prepaidTotal)}%`);
    setTxt('payment-prepaid-unfulfilled', prepaidCounts.unfulfilled);
    setTxt('payment-prepaid-unfulfilled-pct', `${getPct(prepaidCounts.unfulfilled, prepaidTotal)}%`);
    setTxt('payment-prepaid-returned', prepaidCounts.returned);
    setTxt('payment-prepaid-returned-pct', `${getPct(prepaidCounts.returned, prepaidTotal)}%`);
    setTxt('payment-prepaid-canceled', prepaidCounts.canceled);
    setTxt('payment-prepaid-canceled-pct', `${getPct(prepaidCounts.canceled, prepaidTotal)}%`);
    
    const prepaidReturnRate = getPct(prepaidCounts.returned, prepaidTotal);
    setTxt('payment-prepaid-progress-text', `${prepaidReturnRate}%`);
    setWidth('payment-prepaid-progress-bar', parseFloat(prepaidReturnRate));
}

// Group monthFilteredOrders by Category and render dynamic cards
function renderCategoryBreakdown() {
    const gridEl = document.getElementById('category-breakdown-grid');
    if (!gridEl) return;
    
    const defaultCategories = [
        "TOPS",
        "CHUDIDHAR",
        "ANARKALI",
        "LEHENGA",
        "HALF SAREE LEHENGA",
        "LONG GOWN",
        "SHARARA",
        "CO-ORD SET"
    ];
    
    const categoryStats = {};
    defaultCategories.forEach(cat => {
        categoryStats[cat] = {
            total: 0,
            delivered: 0,
            transit: 0,
            pickup: 0,
            unfulfilled: 0,
            returned: 0,
            canceled: 0
        };
    });

    const allStats = {
        total: 0,
        delivered: 0,
        transit: 0,
        pickup: 0,
        unfulfilled: 0,
        returned: 0,
        canceled: 0
    };
    
    state.monthFilteredOrders.forEach(o => {
        const cats = o.category ? o.category.split(',').map(c => c.trim()) : ['-'];
        const stage = getPipelineStage(o);
        
        cats.forEach(cat => {
            const cleanCat = cat === '-' || cat === '' || cat.toLowerCase() === 'nan' ? 'UNASSIGNED' : cat.toUpperCase();
            if (!categoryStats[cleanCat]) {
                categoryStats[cleanCat] = {
                    total: 0,
                    delivered: 0,
                    transit: 0,
                    pickup: 0,
                    unfulfilled: 0,
                    returned: 0,
                    canceled: 0
                };
            }
            
            categoryStats[cleanCat].total += 1;
            allStats.total += 1;
            
            if (categoryStats[cleanCat].hasOwnProperty(stage)) {
                categoryStats[cleanCat][stage] += 1;
                allStats[stage] += 1;
            }
        });
    });
    
    const sortedCats = Object.keys(categoryStats).sort((a, b) => categoryStats[b].total - categoryStats[a].total);
    gridEl.innerHTML = '';
    
    if (sortedCats.length === 0) {
        gridEl.innerHTML = `<div class="empty-state" style="grid-column: 1/-1;">No category data found.</div>`;
        return;
    }
    
    // Build list of all card items to render (ALL + categories)
    const cardsToRender = [
        { title: 'ALL', stats: allStats, isAll: true },
        ...sortedCats.map(cat => ({ title: cat, stats: categoryStats[cat], isAll: false }))
    ];
    
    const totalCards = cardsToRender.length;
    
    cardsToRender.forEach((item, index) => {
        const card = document.createElement('div');
        card.className = item.isAll ? 'category-card all-card' : 'category-card';
        if (item.isAll) {
            card.style.borderLeftColor = 'var(--color-primary)';
        }
        
        // Dynamic grid-column span logic for 12-column grid layout
        let span = 4; // Default: 3-by-3 layout for 9 cards (12 / 4 = 3 per row)
        
        if (totalCards <= 9) {
            // 3 x 3 layout for 9 cards
            span = 4;
        } else if (totalCards === 10) {
            // 4 + 4 + 2 layout (last 2 cards justify full width: 12 / 6 = 2)
            if (index < 8) {
                span = 3;
            } else {
                span = 6;
            }
        } else if (totalCards === 11) {
            // 4 + 4 + 3 layout (last 3 cards justify full width: 12 / 4 = 3)
            if (index < 8) {
                span = 3;
            } else {
                span = 4;
            }
        } else {
            // Generic fallback math for any arbitrary N cards:
            const remainder = totalCards % 4;
            const fullRowsCount = totalCards - remainder;
            if (remainder > 0 && index >= fullRowsCount) {
                span = Math.floor(12 / remainder);
            } else {
                span = 3;
            }
        }
        
        card.style.gridColumn = `span ${span}`;
        
        card.innerHTML = `
            <h3>${item.title}</h3>
            <div class="metrics">
                <div class="metric-group">
                    <span class="metric-label">Total</span>
                    <span class="metric-val" style="color:#85200c;">${item.stats.total}</span>
                </div>
                <div class="metric-group">
                    <span class="metric-label">Delivered</span>
                    <span class="metric-val" style="color:#1e8449;">${item.stats.delivered}</span>
                </div>
                <div class="metric-group">
                    <span class="metric-label">In Transit</span>
                    <span class="metric-val" style="color:#2980b9;">${item.stats.transit}</span>
                </div>
                <div class="metric-group">
                    <span class="metric-label">Pickup</span>
                    <span class="metric-val" style="color:#8e44ad;">${item.stats.pickup}</span>
                </div>
                <div class="metric-group">
                    <span class="metric-label">Unfulfilled</span>
                    <span class="metric-val" style="color:#f39c12;">${item.stats.unfulfilled}</span>
                </div>
                <div class="metric-group">
                    <span class="metric-label">Returned</span>
                    <span class="metric-val" style="color:#d35400;">${item.stats.returned}</span>
                </div>
                <div class="metric-group">
                    <span class="metric-label">Canceled</span>
                    <span class="metric-val" style="color:#c0392b;">${item.stats.canceled}</span>
                </div>
            </div>
        `;
        gridEl.appendChild(card);
    });
}


// Donut chart outer-label plugin — radial lines with label text, matching sample style
const donutLabelsLinePlugin = {
    id: 'donutLabelsLine',
    afterDraw(chart) {
        if (chart.config.type !== 'doughnut') return;
        const { ctx } = chart;
        const cssStyle  = getComputedStyle(document.body);
        const lineColor = cssStyle.getPropertyValue('--text-muted').trim()    || 'rgba(128,128,128,0.55)';
        const textColor = cssStyle.getPropertyValue('--text-primary').trim()  || '#1e293b';
        const subColor  = cssStyle.getPropertyValue('--text-secondary').trim()|| '#64748b';

        chart.data.datasets.forEach((dataset, dsIdx) => {
            const meta = chart.getDatasetMeta(dsIdx);
            if (meta.hidden) return;
            const sum = dataset.data.reduce((a, b) => a + b, 0);
            if (!sum) return;

            const labels  = chart.data.labels || [];
            const RADEXT  = 18;   // px beyond ring for the angled segment
            const TICK    = 8;    // px for horizontal tick
            const GAP     = 4;    // px between tick end and text
            const MINGAP  = 15;   // min px between adjacent label Y positions

            // Build label positions from midAngle
            const items = [];
            meta.data.forEach((el, idx) => {
                const val = dataset.data[idx];
                if (!val) return;
                const { x, y, outerRadius, startAngle, endAngle } = el;
                const mid  = startAngle + (endAngle - startAngle) / 2;
                const cosM = Math.cos(mid);
                const sinM = Math.sin(mid);
                const pct  = (val * 100 / sum).toFixed(1) + '%';
                const name = labels[idx] || '';

                // Anchor: outer ring midpoint
                const anchorX = x + cosM * outerRadius;
                const anchorY = y + sinM * outerRadius;

                // Elbow: extend radially
                const elbX = x + cosM * (outerRadius + RADEXT);
                const elbY = y + sinM * (outerRadius + RADEXT);

                const isLeft = cosM < 0;
                const tikX = elbX + (isLeft ? -TICK : TICK);
                const tikY = elbY;

                items.push({ anchorX, anchorY, elbX, elbY, tikX, tikY: elbY, isLeft, pct, name, sortY: elbY });
            });

            // Separate left / right, then nudge overlapping Y values
            ['left', 'right'].forEach(side => {
                const group = items
                    .filter(it => (side === 'left') === it.isLeft)
                    .sort((a, b) => a.sortY - b.sortY);

                // Push apart any labels that are too close vertically
                for (let pass = 0; pass < 20; pass++) {
                    let moved = false;
                    for (let k = 1; k < group.length; k++) {
                        const prev = group[k - 1];
                        const curr = group[k];
                        if (curr.tikY - prev.tikY < MINGAP) {
                            const shift = (MINGAP - (curr.tikY - prev.tikY)) / 2;
                            prev.tikY -= shift;
                            curr.tikY += shift;
                            moved = true;
                        }
                    }
                    if (!moved) break;
                }

                // Clamp: keep labels out of the bottom legend zone (~50px)
                group.forEach(it => {
                    it.tikY = Math.min(it.tikY, chart.height - 52);
                    it.tikY = Math.max(it.tikY, 8);
                });

                group.forEach(it => {
                    ctx.save();

                    // Line: slice edge → elbow → horizontal tick
                    ctx.beginPath();
                    ctx.moveTo(it.anchorX, it.anchorY);
                    ctx.lineTo(it.elbX, it.tikY);
                    ctx.lineTo(it.tikX, it.tikY);
                    ctx.strokeStyle = lineColor;
                    ctx.lineWidth   = 1;
                    ctx.lineJoin    = 'round';
                    ctx.stroke();

                    // Small dot at elbow
                    ctx.beginPath();
                    ctx.arc(it.tikX, it.tikY, 1.5, 0, Math.PI * 2);
                    ctx.fillStyle = lineColor;
                    ctx.fill();

                    // Text alignment
                    const tx = it.tikX + (it.isLeft ? -GAP : GAP);
                    ctx.textBaseline = 'middle';
                    ctx.textAlign    = it.isLeft ? 'right' : 'left';

                    // Percentage only — bold, centered on the tick line
                    ctx.font      = 'bold 10.5px Inter, sans-serif';
                    ctx.fillStyle = textColor;
                    ctx.fillText(it.pct, tx, it.tikY);

                    ctx.restore();
                });
            });
        });
    }
};


// Build an HTML legend for a donut chart and inject it into the given element
function renderDonutLegend(containerId, labels, colors) {
    const el = document.getElementById(containerId);
    if (!el) return;
    el.innerHTML = labels.map((label, i) =>
        `<span class="chart-legend-item">
            <span class="chart-legend-dot" style="background:${colors[i]}"></span>
            <span class="chart-legend-label">${label}</span>
        </span>`
    ).join('');
}

// Generate sales trend and donut breakdown charts
function renderCharts() {
    // Destroy previous chart instances if they exist
    if (state.charts.salesTrend) state.charts.salesTrend.destroy();
    if (state.charts.dailyTrend) state.charts.dailyTrend.destroy();
    if (state.charts.finance) state.charts.finance.destroy();
    if (state.charts.delivery) state.charts.delivery.destroy();
    
    // Read theme colors dynamically from computed styles (CSS variables)
    const style = getComputedStyle(document.body);
    const textSecondary = style.getPropertyValue('--text-secondary').trim() || '#94a3b8';
    const borderChart = style.getPropertyValue('--border-chart').trim() || '#111625';
    const gridColor = style.getPropertyValue('--grid-color').trim() || 'rgba(255,255,255,0.03)';
    const brandColor = style.getPropertyValue('--color-brand').trim() || '#3b82f6';
    
    // --- Chart 1: Sales Trend (Combo Bar/Line) ---
    // Group orders chronologically
    const chronMonths = [
        "Jun 2026", "Jul 2026", "Aug 2026", "Sep 2026", "Oct 2026", 
        "Nov 2026", "Dec 2026", "Jan 2027", "Feb 2027", "Mar 2027"
    ];
    const monthlySales = Array(10).fill(0);
    const monthlyCounts = Array(10).fill(0);
    
    state.orders.forEach(o => {
        const m = getMonthYearStr(o.dateOfOrder);
        const idx = chronMonths.indexOf(m);
        if (idx !== -1) {
            monthlySales[idx] += o.totalPrice;
            monthlyCounts[idx] += 1;
        }
    });
    
    const ctxTrend = document.getElementById('salesTrendChart').getContext('2d');
    state.charts.salesTrend = new Chart(ctxTrend, {
        type: 'bar',
        data: {
            labels: chronMonths,
            datasets: [
                {
                    label: 'Revenue (Rs.)',
                    data: monthlySales,
                    backgroundColor: 'rgba(16, 185, 129, 0.45)',
                    borderColor: '#10b981',
                    borderWidth: 2,
                    borderRadius: 6,
                    yAxisID: 'yRevenue'
                },
                {
                    label: 'Order Count',
                    data: monthlyCounts,
                    type: 'line',
                    borderColor: brandColor,
                    borderWidth: 3,
                    pointBackgroundColor: brandColor,
                    pointRadius: 4,
                    fill: false,
                    tension: 0.3,
                    yAxisID: 'yCount'
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { labels: { color: textSecondary, font: { family: 'Inter' } } }
            },
            scales: {
                x: { grid: { color: gridColor }, ticks: { color: textSecondary, font: { family: 'Inter' } } },
                yRevenue: {
                    type: 'linear',
                    position: 'left',
                    grid: { color: gridColor },
                    ticks: { color: '#10b981', font: { family: 'Inter' }, callback: val => '₹' + val.toLocaleString() }
                },
                yCount: {
                    type: 'linear',
                    position: 'right',
                    grid: { drawOnChartArea: false },
                    ticks: { color: brandColor, font: { family: 'Inter' }, stepSize: 5 }
                }
            }
        }
    });
    
    // --- Chart 2: Shopify Financial Status Donut ---
    const financeCounts = { paid: 0, pending: 0, refunded: 0 };
    state.monthFilteredOrders.forEach(o => {
        const s = o.financialStatus.toLowerCase();
        if (financeCounts[s] !== undefined) {
            financeCounts[s]++;
        }
    });
    
    const donutPadding = window.innerWidth <= 768 ? { left: 15, right: 15, top: 15, bottom: 15 } : { left: 75, right: 75, top: 30, bottom: 15 };

    const ctxFinance = document.getElementById('financeChart').getContext('2d');
    state.charts.finance = new Chart(ctxFinance, {
        type: 'doughnut',
        data: {
            labels: ['Paid', 'Pending', 'Refunded'],
            datasets: [{
                data: [financeCounts.paid, financeCounts.pending, financeCounts.refunded],
                backgroundColor: ['#10b981', '#f97316', '#ef4444'],
                borderColor: borderChart,
                borderWidth: 2
            }]
        },
        plugins: [donutLabelsLinePlugin],
        options: {
            responsive: true,
            maintainAspectRatio: false,
            layout: { padding: donutPadding },
            plugins: { legend: { display: false } },
            cutout: '65%'
        }
    });
    renderDonutLegend('financeChartLegend', ['Paid','Pending','Refunded'], ['#10b981','#f97316','#ef4444']);

    // --- Chart 3: Shiprocket Logistics Status Donut ---
    const deliveryCounts = { delivered: 0, transit: 0, rto: 0, canceled: 0 };
    state.monthFilteredOrders.forEach(o => {
        const s = o.logisticsStatus.toLowerCase();
        if (s === 'delivered') {
            deliveryCounts.delivered++;
        } else if (['rto initiated', 'rto in transit', 'rto delivered', 'returned'].includes(s)) {
            deliveryCounts.rto++;
        } else if (s === 'canceled') {
            deliveryCounts.canceled++;
        } else if (s && s !== 'unfulfilled') {
            deliveryCounts.transit++;
        }
    });
    
    const ctxDelivery = document.getElementById('deliveryChart').getContext('2d');
    state.charts.delivery = new Chart(ctxDelivery, {
        type: 'doughnut',
        data: {
            labels: ['Delivered', 'In Transit', 'RTO', 'Canceled'],
            datasets: [{
                data: [deliveryCounts.delivered, deliveryCounts.transit, deliveryCounts.rto, deliveryCounts.canceled],
                backgroundColor: ['#10b981', '#3b82f6', '#f97316', '#ef4444'],
                borderColor: borderChart,
                borderWidth: 2
            }]
        },
        plugins: [donutLabelsLinePlugin],
        options: {
            responsive: true,
            maintainAspectRatio: false,
            layout: { padding: donutPadding },
            plugins: { legend: { display: false } },
            cutout: '65%'
        }
    });
    renderDonutLegend('deliveryChartLegend', ['Delivered','In Transit','RTO','Canceled'], ['#10b981','#3b82f6','#f97316','#ef4444']);

    // --- Chart 4: Daily Sales & Volume Trend (Selected Period) ---
    const dailyData = {};
    state.monthFilteredOrders.forEach(o => {
        if (o.dateOfOrder) {
            if (!dailyData[o.dateOfOrder]) {
                dailyData[o.dateOfOrder] = { revenue: 0, count: 0 };
            }
            dailyData[o.dateOfOrder].revenue += o.totalPrice;
            dailyData[o.dateOfOrder].count += 1;
        }
    });
    
    let sortedDates = Object.keys(dailyData).sort();
    if (sortedDates.length > 30) {
        sortedDates = sortedDates.slice(-30);
    }
    const isMobileView = window.innerWidth <= 768;
    const monthsShort = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

    const dailyRevenues = sortedDates.map(d => dailyData[d].revenue);
    const dailyCounts = sortedDates.map(d => dailyData[d].count);
    const dailyLabels = sortedDates.map(d => {
        const parts = d.split('-');
        if (parts.length < 3) return d;
        if (isMobileView) {
            const mIdx = parseInt(parts[1], 10) - 1;
            return `${parts[2]} ${monthsShort[mIdx] || parts[1]}`;
        }
        return `${parts[2]}-${parts[1]}-${parts[0]}`;
    });

    const ctxDaily = document.getElementById('dailyTrendChart');
    if (ctxDaily) {
        const ctxDaily2d = ctxDaily.getContext('2d');
        state.charts.dailyTrend = new Chart(ctxDaily2d, {
            type: 'bar',
            data: {
                labels: dailyLabels,
                datasets: [
                    {
                        label: 'Revenue (Rs.)',
                        data: dailyRevenues,
                        backgroundColor: 'rgba(16, 185, 129, 0.45)',
                        borderColor: '#10b981',
                        borderWidth: 2,
                        borderRadius: 6,
                        yAxisID: 'y'
                    },
                    {
                        label: 'Orders Count',
                        data: dailyCounts,
                        type: 'line',
                        borderColor: brandColor,
                        borderWidth: 3,
                        pointBackgroundColor: brandColor,
                        pointBorderColor: '#fff',
                        pointBorderWidth: 2,
                        pointRadius: 4,
                        pointHoverRadius: 6,
                        fill: false,
                        yAxisID: 'y1'
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        position: 'top',
                        labels: {
                            color: textSecondary,
                            font: { family: 'Inter, sans-serif', size: 11, weight: '500' }
                        }
                    },
                    tooltip: {
                        mode: 'index',
                        intersect: false,
                        titleFont: { family: 'Inter, sans-serif', weight: '600' },
                        bodyFont: { family: 'Inter, sans-serif' }
                    }
                },
                scales: {
                    x: {
                        grid: { display: false },
                        ticks: {
                            color: textSecondary,
                            font: { family: 'Inter, sans-serif', size: 10, weight: '500' },
                            autoSkip: true,
                            maxTicksLimit: isMobileView ? 6 : 30,
                            maxRotation: isMobileView ? 45 : 0,
                            minRotation: isMobileView ? 45 : 0
                        }
                    },
                    y: {
                        type: 'linear',
                        display: true,
                        position: 'left',
                        grid: { color: gridColor },
                        ticks: {
                            color: textSecondary,
                            font: { family: 'Inter, sans-serif', size: 10 },
                            callback: (value) => {
                                if (isMobileView && value >= 1000) {
                                    return '₹' + (value / 1000).toFixed(0) + 'k';
                                }
                                return '₹' + value.toLocaleString();
                            }
                        }
                    },
                    y1: {
                        type: 'linear',
                        display: true,
                        position: 'right',
                        grid: { drawOnChartArea: false },
                        ticks: {
                            color: textSecondary,
                            font: { family: 'Inter, sans-serif', size: 10 }
                        }
                    }
                }
            }
        });
    }
}

// Update existing charts color theme without full destroy/recreate
function updateChartsTheme() {
    if (!state.charts.salesTrend && !state.charts.dailyTrend && !state.charts.finance && !state.charts.delivery) return;

    const style = getComputedStyle(document.body);
    const textSecondary = style.getPropertyValue('--text-secondary').trim() || '#94a3b8';
    const borderChart = style.getPropertyValue('--border-chart').trim() || '#111625';
    const gridColor = style.getPropertyValue('--grid-color').trim() || 'rgba(255,255,255,0.03)';
    const brandColor = style.getPropertyValue('--color-brand').trim() || '#3b82f6';

    if (state.charts.salesTrend) {
        state.charts.salesTrend.data.datasets[1].borderColor = brandColor;
        state.charts.salesTrend.data.datasets[1].pointBackgroundColor = brandColor;
        state.charts.salesTrend.options.plugins.legend.labels.color = textSecondary;
        state.charts.salesTrend.options.scales.x.grid.color = gridColor;
        state.charts.salesTrend.options.scales.x.ticks.color = textSecondary;
        state.charts.salesTrend.options.scales.yRevenue.grid.color = gridColor;
        state.charts.salesTrend.options.scales.yCount.ticks.color = brandColor;
        state.charts.salesTrend.update();
    }

    if (state.charts.dailyTrend) {
        state.charts.dailyTrend.data.datasets[1].borderColor = brandColor;
        state.charts.dailyTrend.data.datasets[1].pointBackgroundColor = brandColor;
        state.charts.dailyTrend.options.plugins.legend.labels.color = textSecondary;
        state.charts.dailyTrend.options.scales.x.ticks.color = textSecondary;
        state.charts.dailyTrend.options.scales.y.grid.color = gridColor;
        state.charts.dailyTrend.options.scales.y.ticks.color = textSecondary;
        state.charts.dailyTrend.options.scales.y1.ticks.color = textSecondary;
        state.charts.dailyTrend.update();
    }

    if (state.charts.finance) {
        state.charts.finance.data.datasets[0].borderColor = borderChart;
        state.charts.finance.options.plugins.legend.labels.color = textSecondary;
        state.charts.finance.update();
    }

    if (state.charts.delivery) {
        state.charts.delivery.data.datasets[0].borderColor = borderChart;
        state.charts.delivery.options.plugins.legend.labels.color = textSecondary;
        state.charts.delivery.update();
    }
}

// Render Table matching current filter sets
function renderTable() {
    const totalRecords = state.filteredOrders.length;
    
    if (totalRecords === 0) {
        elements.ordersTbody.innerHTML = `<tr><td colspan="12" class="empty-state">No matching orders found.</td></tr>`;
        elements.tableCount.textContent = 'Showing 0 orders';
        const pagWrapper = document.getElementById('pagination-wrapper');
        if (pagWrapper) pagWrapper.style.display = 'none';
        return;
    }
    
    const totalPages = Math.ceil(totalRecords / state.pageSize) || 1;
    if (state.currentPage > totalPages) state.currentPage = totalPages;
    if (state.currentPage < 1) state.currentPage = 1;
    
    const startIndex = (state.currentPage - 1) * state.pageSize;
    const endIndex = Math.min(startIndex + state.pageSize, totalRecords);
    const pageOrders = state.filteredOrders.slice(startIndex, endIndex);
    
    elements.tableCount.textContent = `Showing ${startIndex + 1} - ${endIndex} of ${totalRecords} active orders`;
    elements.ordersTbody.innerHTML = '';
    
    // Update pagination controls
    const pagWrapper = document.getElementById('pagination-wrapper');
    if (pagWrapper) {
        pagWrapper.style.display = 'flex';
        elements.paginationInfo.innerHTML = '';
        
        elements.prevPageBtn.disabled = state.currentPage === 1;
        elements.nextPageBtn.disabled = state.currentPage === totalPages;
        
        // Dynamic Pagination Number Rendering
        const startPage = Math.max(1, state.currentPage - 1);
        const endPage = Math.min(totalPages, state.currentPage + 1);
        
        // Helper to append page button
        const addPageButton = (page, text, isActive, isDisabled) => {
            const btn = document.createElement('button');
            btn.className = 'page-num-btn';
            if (isActive) btn.classList.add('active');
            if (isDisabled) btn.classList.add('disabled');
            btn.textContent = text || page;
            
            if (!isDisabled && !isActive) {
                btn.addEventListener('click', () => {
                    state.currentPage = page;
                    renderTable();
                });
            }
            elements.paginationInfo.appendChild(btn);
        };
        
        // First page
        if (startPage > 1) {
            addPageButton(1, '1', state.currentPage === 1, false);
            if (startPage > 2) {
                addPageButton(null, '...', false, true);
            }
        }
        
        // Page numbers range
        for (let i = startPage; i <= endPage; i++) {
            addPageButton(i, i.toString(), state.currentPage === i, false);
        }
        
        // Last page
        if (endPage < totalPages) {
            if (endPage < totalPages - 1) {
                addPageButton(null, '...', false, true);
            }
            addPageButton(totalPages, totalPages.toString(), state.currentPage === totalPages, false);
        }
    }
    
    pageOrders.forEach(o => {
        const tr = document.createElement('tr');
        tr.addEventListener('click', () => showOrderDetail(o));
        
        // Casing format for badges
        const finClass = o.financialStatus.toLowerCase() === 'paid' ? 'success' : o.financialStatus.toLowerCase() === 'pending' ? 'warning' : 'danger';
        const fulClass = o.fulfillmentStatus.toLowerCase() === 'fulfilled' ? 'success' : 'danger';
        
        let logClass = 'info';
        if (o.logisticsStatus === 'DELIVERED') logClass = 'success';
        if (o.logisticsStatus === 'CANCELED') logClass = 'danger';
        if (o.logisticsStatus.includes('RTO')) logClass = 'warning';
        
        // Feedback status pill
        let feedbackHTML = '';
        if (o.feedbackReceived === 'Yes') {
            feedbackHTML = '<span class="status-pill success">Received</span>';
        } else if (o.feedbackSent === 'Yes') {
            feedbackHTML = '<span class="status-pill warning">Sent</span>';
        } else {
            feedbackHTML = '<span class="status-pill danger">Not Sent</span>';
        }
        
        // AWB code cell
        const awbDisplay = (o.awbCode && o.awbCode !== '-') 
            ? `<a href="${o.trackingLink}" target="_blank" class="awb-link" onclick="event.stopPropagation();">${o.awbCode} ↗</a>` 
            : '-';
            
        // Row Cells
        const orderNoDisplay = String(o.orderNo).startsWith('#') ? o.orderNo : `#${o.orderNo}`;
        tr.innerHTML = `
            <td><strong>${orderNoDisplay}</strong></td>
            <td>${formatDateDisplay(o.dateOfOrder)}</td>
            <td>${o.customerName}</td>
            <td class="items-col">${o.itemsOrdered}</td>
            <td class="sku-col">${o.sku}</td>
            <td class="category-col">${o.category}</td>
            <td>${o.city}</td>
            <td><span class="status-pill info">${o.paymentMethod}</span></td>
            <td class="price-col">${formatCurrency(o.totalPrice)}</td>
            <td><span class="status-pill ${finClass}">${o.financialStatus}</span></td>
            <td><span class="status-pill ${fulClass}">${o.fulfillmentStatus}</span></td>
            <td><span class="status-pill ${logClass}">${o.logisticsStatus}</span></td>
            <td>${awbDisplay}</td>
            <td class="comments-col">${o.shiprocketComments || '-'}</td>
            <td>${feedbackHTML}</td>
        `;
        elements.ordersTbody.appendChild(tr);
    });
}

// Apply inputs and selects to filter the table list
function applyFilters() {
    const q = elements.searchInput.value.toLowerCase().trim();
    const month = elements.filterMonth.value;
    const pay = elements.filterPayment.value;
    const cat = elements.filterCategory.value;
    const status = elements.filterStatus.value;
    

    // 1. Filter by month/date range for dashboard stats, KPIs, and donut charts
    state.monthFilteredOrders = state.orders.filter(o => {
        if (state.dateRangeFrom && state.dateRangeTo) {
            return o.dateOfOrder >= state.dateRangeFrom && o.dateOfOrder <= state.dateRangeTo;
        }
        return !month || getMonthYearStr(o.dateOfOrder) === month;
    });
    
    // 2. Filter by all inputs for master table rows
    state.filteredOrders = state.orders.filter(o => {
        // Search text (with safe property checks)
        const matchesQuery = !q || 
            (o.orderNo && String(o.orderNo).toLowerCase().includes(q)) ||
            (o.customerName && o.customerName.toLowerCase().includes(q)) ||
            (o.itemsOrdered && o.itemsOrdered.toLowerCase().includes(q)) ||
            (o.shiprocketComments && o.shiprocketComments.toLowerCase().includes(q)) ||
            (o.city && o.city.toLowerCase().includes(q)) ||
            (o.pinCode && o.pinCode.toLowerCase().includes(q));
            
        // Date / Month range matching
        let matchesDateRange = true;
        if (state.dateRangeFrom && state.dateRangeTo) {
            matchesDateRange = o.dateOfOrder >= state.dateRangeFrom && o.dateOfOrder <= state.dateRangeTo;
        } else if (month) {
            matchesDateRange = getMonthYearStr(o.dateOfOrder) === month;
        }
        
        // Payment dropdown
        const matchesPayment = !pay || o.paymentMethod.toLowerCase().includes(pay.toLowerCase());
        
        // Category dropdown
        let matchesCategory = true;
        if (cat) {
            matchesCategory = o.category && o.category.toUpperCase().split(',').map(x => x.trim()).includes(cat.toUpperCase());
        }
        
        // Logistics Status dropdown (using the 6 simplified pipeline stages)
        let matchesStatus = true;
        if (status) {
            matchesStatus = getPipelineStage(o) === status.toLowerCase();
        }
        
        return matchesQuery && matchesDateRange && matchesPayment && matchesStatus && matchesCategory;
    });
    
    state.currentPage = 1; // Reset to page 1 on filter
    renderDashboard();
}

// Clear all filters
function clearFilters() {
    elements.searchInput.value = '';
    elements.topFilterMonth.value = '';
    state.dateRangeFrom = null;
    state.dateRangeTo = null;
    state.hoveredDate = null;
    state.pickerLeftMonth = (() => {
        const today = new Date();
        return new Date(today.getFullYear(), today.getMonth() - 1, 1);
    })();
    if (elements.dateRangeBtn) elements.dateRangeBtn.textContent = 'Select Date Range';
    elements.filterMonth.value = '';
    if (elements.chartMonthSelect) elements.chartMonthSelect.value = '';
    elements.filterPayment.value = '';
    elements.filterCategory.value = '';
    elements.filterStatus.value = '';
    state.filteredOrders = [...state.orders];
    state.monthFilteredOrders = [...state.orders];
    state.currentPage = 1; // Reset to page 1
    renderDashboard();
}

// Convert "2026-07-02" to "02-07-2026"
function formatDateDisplay(dateStr) {
    if (!dateStr) return '';
    const parts = dateStr.split('-');
    if (parts.length < 3) return dateStr;
    const year = parts[0];
    const month = parts[1];
    const day = parts[2];
    return `${day}-${month}-${year}`;
}

// Open and populate single order detail modal
function showOrderDetail(o) {
    elements.modalTitle.textContent = `Order Details #${o.orderNo}`;
    
    document.getElementById('det-name').textContent = o.customerName || 'N/A';
    document.getElementById('det-date').textContent = formatDateDisplay(o.dateOfOrder) || 'N/A';
    document.getElementById('det-pay-method').textContent = o.paymentMethod || 'N/A';
    document.getElementById('det-prepaid').textContent = o.prepaid || 'No';
    document.getElementById('det-cod').textContent = o.cod || 'No';
    document.getElementById('det-price').textContent = formatCurrency(o.totalPrice);
    document.getElementById('det-sku').textContent = o.sku || '-';
    document.getElementById('det-category').textContent = o.category || '-';
    document.getElementById('det-city').textContent = o.city || 'N/A';
    document.getElementById('det-pin').textContent = o.pinCode || 'N/A';
    
    document.getElementById('det-financial').textContent = o.financialStatus || 'N/A';
    document.getElementById('det-fulfillment').textContent = o.fulfillmentStatus || 'N/A';
    document.getElementById('det-logistics').textContent = o.logisticsStatus || 'N/A';
    document.getElementById('det-returned').textContent = o.returned ? 'Yes (Returned)' : 'No';
    document.getElementById('det-cod-denial').textContent = o.codDenies || 'No';
    document.getElementById('det-feedback-sent').textContent = o.feedbackSent || 'No';
    document.getElementById('det-feedback-received').textContent = o.feedbackReceived || 'No';
    document.getElementById('det-awb').textContent = o.awbCode || '-';
    
    const trLinkEl = document.getElementById('det-tracking');
    if (o.trackingLink && o.trackingLink !== '-') {
        trLinkEl.innerHTML = `<a href="${o.trackingLink}" target="_blank" class="track-btn" style="color:var(--color-brand); font-weight:600; text-decoration:underline;">Track Shipment ↗</a>`;
    } else {
        trLinkEl.textContent = '-';
    }
    
    document.getElementById('det-items').textContent = o.itemsOrdered || 'N/A';
    document.getElementById('det-comments').textContent = o.shiprocketComments || '-';
    
    elements.modal.style.display = 'flex';
}

// Populate calendar pane day buttons
function fillDaysPane(container, dateObj) {
    container.innerHTML = '';
    const year = dateObj.getFullYear();
    const month = dateObj.getMonth();
    
    // First day of month (0 = Sunday, 6 = Saturday)
    const firstDayIndex = new Date(year, month, 1).getDay();
    // Total days in month
    const totalDays = new Date(year, month + 1, 0).getDate();
    
    // Add empty spacer slots
    for (let i = 0; i < firstDayIndex; i++) {
        const spacer = document.createElement('div');
        spacer.className = 'dp-day-btn empty';
        container.appendChild(spacer);
    }
    
    // Add day numbers
    for (let day = 1; day <= totalDays; day++) {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'dp-day-btn';
        btn.textContent = day;
        
        // Build iso string for validation & comparisons
        const dStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        btn.dataset.date = dStr;
        
        // Apply class names based on state
        updateDayBtnState(btn, dStr);
        
        // Listeners for selection
        btn.addEventListener('click', () => handleDayClick(dStr));
        btn.addEventListener('mouseenter', () => handleDayMouseEnter(dStr));
        
        container.appendChild(btn);
    }
}

// Update class state for individual day button based on current From/To range
function updateDayBtnState(btn, dStr) {
    btn.className = 'dp-day-btn'; // Reset
    
    const dTime = new Date(dStr).getTime();
    
    // Enforce operational range: June 2026 to March 2027
    const minTime = new Date('2026-06-01').getTime();
    const maxTime = new Date('2027-03-31').getTime();
    if (dTime < minTime || dTime > maxTime) {
        btn.classList.add('disabled');
        return;
    }
    
    const fromTime = state.dateRangeFrom ? new Date(state.dateRangeFrom).getTime() : null;
    const toTime = state.dateRangeTo ? new Date(state.dateRangeTo).getTime() : null;
    const hoverTime = state.hoveredDate ? new Date(state.hoveredDate).getTime() : null;
    
    if (state.dateRangeFrom === dStr) {
        btn.classList.add('selected-start');
    }
    if (state.dateRangeTo === dStr) {
        btn.classList.add('selected-end');
    }
    
    if (fromTime && toTime) {
        if (dTime > fromTime && dTime < toTime) {
            btn.classList.add('in-range');
        }
    } else if (fromTime && hoverTime && dTime > fromTime && dTime <= hoverTime) {
        btn.classList.add('hovered-range');
    }
}

// Helper to format, apply, and close active date selection
function applyActiveSelectionAndClose() {
    if (state.dateRangeFrom) {
        if (!state.dateRangeTo) {
            state.dateRangeTo = state.dateRangeFrom;
        }
        const fromFormatted = formatDateDisplay(state.dateRangeFrom);
        const toFormatted = formatDateDisplay(state.dateRangeTo);
        elements.dateRangeBtn.textContent = state.dateRangeFrom === state.dateRangeTo ? fromFormatted : `${fromFormatted} - ${toFormatted}`;
        
        // Clear Month selections
        elements.topFilterMonth.value = '';
        elements.filterMonth.value = '';
        if (elements.chartMonthSelect) elements.chartMonthSelect.value = '';
        
        applyFilters();
    }
    elements.datePickerPopup.style.display = 'none';
}

// Handle clicking on a calendar day
function handleDayClick(dStr) {
    const dTime = new Date(dStr).getTime();
    const minTime = new Date('2026-06-01').getTime();
    const maxTime = new Date('2027-03-31').getTime();
    if (dTime < minTime || dTime > maxTime) return; // ignore disabled
    
    if (!state.dateRangeFrom || (state.dateRangeFrom && state.dateRangeTo)) {
        // Start new range selection
        state.dateRangeFrom = dStr;
        state.dateRangeTo = null;
        state.hoveredDate = null;
        renderDatePickerCalendars();
    } else if (state.dateRangeFrom && !state.dateRangeTo) {
        const fromTime = new Date(state.dateRangeFrom).getTime();
        if (dTime < fromTime) {
            // Clicked earlier day: set as new From date
            state.dateRangeFrom = dStr;
            renderDatePickerCalendars();
        } else {
            // Set as To date and auto-apply selection!
            state.dateRangeTo = dStr;
            state.hoveredDate = null;
            applyActiveSelectionAndClose();
        }
    }
}

// Handle hovering over a calendar day during range selection
function handleDayMouseEnter(dStr) {
    if (state.dateRangeFrom && !state.dateRangeTo) {
        const dTime = new Date(dStr).getTime();
        const fromTime = new Date(state.dateRangeFrom).getTime();
        if (dTime >= fromTime) {
            state.hoveredDate = dStr;
        } else {
            state.hoveredDate = null;
        }
        // Update only styles dynamically for performance
        const leftPaneBtns = elements.dpLeftDays.querySelectorAll('.dp-day-btn:not(.empty):not(.disabled)');
        const rightPaneBtns = elements.dpRightDays.querySelectorAll('.dp-day-btn:not(.empty):not(.disabled)');
        
        [...leftPaneBtns, ...rightPaneBtns].forEach(btn => {
            const btnDate = btn.dataset.date;
            updateDayBtnState(btn, btnDate);
        });
    }
}

// Render custom dual-pane date calendars
function renderDatePickerCalendars() {
    const leftMonth = state.pickerLeftMonth;
    const rightMonth = new Date(leftMonth.getFullYear(), leftMonth.getMonth() + 1, 1);
    
    // Set headers
    const formatMonthYear = (d) => {
        const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
        return `${months[d.getMonth()]} ${d.getFullYear()}`;
    };
    elements.dpLeftMonthName.textContent = formatMonthYear(leftMonth);
    elements.dpRightMonthName.textContent = formatMonthYear(rightMonth);
    
    // Fill left pane days
    fillDaysPane(elements.dpLeftDays, leftMonth);
    // Fill right pane days
    fillDaysPane(elements.dpRightDays, rightMonth);
}

// Initialize Date picker controls
function initDatePicker() {
    // Toggle popup visibility
    elements.dateRangeBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        const isShown = elements.datePickerPopup.style.display === 'flex';
        elements.datePickerPopup.style.display = isShown ? 'none' : 'flex';
        if (!isShown) {
            if (state.dateRangeFrom) {
                const selDate = new Date(state.dateRangeFrom);
                state.pickerLeftMonth = new Date(selDate.getFullYear(), selDate.getMonth(), 1);
            } else {
                state.pickerLeftMonth = (() => {
                    const today = new Date();
                    return new Date(today.getFullYear(), today.getMonth() - 1, 1);
                })();
            }
            renderDatePickerCalendars();
        }
    });
    
    // Close when clicking outside & auto-apply active selection
    document.addEventListener('click', (e) => {
        if (elements.datePickerPopup && elements.datePickerPopup.style.display === 'flex') {
            const isClickInside = elements.datePickerPopup.contains(e.target) || 
                                  (e.target.closest && e.target.closest('.date-picker-relative-container')) ||
                                  (e.target.classList && e.target.classList.contains('dp-day-btn')) ||
                                  (e.target.closest && e.target.closest('.dp-day-btn'));
            if (!isClickInside) {
                applyActiveSelectionAndClose();
            }
        }
    });
    
    // Prev/Next month navigation
    elements.dpPrevMonth.addEventListener('click', (e) => {
        e.stopPropagation();
        state.pickerLeftMonth.setMonth(state.pickerLeftMonth.getMonth() - 1);
        renderDatePickerCalendars();
    });
    
    elements.dpNextMonth.addEventListener('click', (e) => {
        e.stopPropagation();
        state.pickerLeftMonth.setMonth(state.pickerLeftMonth.getMonth() + 1);
        renderDatePickerCalendars();
    });
}
