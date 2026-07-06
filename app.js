// Live Dashboard JavaScript Controller

// Application State
const state = {
    apiUrl: '',
    rawSheets: null,
    orders: [],
    filteredOrders: [],
    currentPage: 1,
    pageSize: 15,
    charts: {
        salesTrend: null,
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
    kpiRefunded: document.getElementById('kpi-total-refunded'),
    kpiProfit: document.getElementById('kpi-total-profit'),
    kpiReturnRate: document.getElementById('kpi-return-rate'),
    kpiDenialRate: document.getElementById('kpi-denial-rate'),
    kpiReturnProgress: document.getElementById('kpi-return-progress'),
    kpiDenialProgress: document.getElementById('kpi-denial-progress'),
    kpiReturnRateSub: document.getElementById('kpi-return-rate-sub'),
    kpiDenialRateSub: document.getElementById('kpi-denial-rate-sub'),
    
    // Pipeline
    pipeUnfulfilled: document.querySelector('#step-unfulfilled .step-count'),
    pipeReturned: document.querySelector('#step-returned .step-count'),
    pipePickup: document.querySelector('#step-pickup .step-count'),
    pipeTransit: document.querySelector('#step-transit .step-count'),
    pipeDelivered: document.querySelector('#step-delivered .step-count'),
    
    // Filters & Table
    searchInput: document.getElementById('search-input'),
    topFilterMonth: document.getElementById('top-filter-month'),
    filterMonth: document.getElementById('filter-month'),
    filterPayment: document.getElementById('filter-payment'),
    filterStatus: document.getElementById('filter-status'),
    clearFiltersBtn: document.getElementById('clear-filters-btn'),
    ordersTbody: document.getElementById('orders-tbody'),
    tableCount: document.getElementById('table-record-count'),
    prevPageBtn: document.getElementById('prev-page-btn'),
    nextPageBtn: document.getElementById('next-page-btn'),
    paginationInfo: document.getElementById('pagination-info'),
    
    // Modal
    modal: document.getElementById('order-modal'),
    modalTitle: document.getElementById('modal-order-title'),
    modalCloseBtn: document.getElementById('modal-close-btn')
};

// Initial Setup
window.addEventListener('DOMContentLoaded', () => {
    const savedTheme = localStorage.getItem('theme') || 'dark';
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

    // 1. Try global window.APP_CONFIG loaded via script tag (bypasses local file CORS)
    if (window.APP_CONFIG && window.APP_CONFIG.google_web_app_url) {
        const scriptUrl = window.APP_CONFIG.google_web_app_url;
        if (scriptUrl && !scriptUrl.includes('YOUR_URL') && !scriptUrl.includes('YOUR_GOOGLE_WEB_APP_URL_HERE')) {
            state.apiUrl = scriptUrl;
            elements.apiUrlInput.value = scriptUrl;
            updateStatus('loading', 'Loading data from Google Sheet...');
            fetchData();
            return;
        }
    }

    // 2. Try local storage
    const savedUrl = localStorage.getItem('google_web_app_url');
    if (savedUrl) {
        state.apiUrl = savedUrl;
        elements.apiUrlInput.value = savedUrl;
        updateStatus('loading', 'Loading data from Google Sheet...');
        fetchData();
        return;
    }

    // 3. Try config.json on the server
    try {
        const response = await fetch('config.json');
        if (response.ok) {
            const config = await response.json();
            const configUrl = config.google_web_app_url;
            if (configUrl && !configUrl.includes('YOUR_URL') && !configUrl.includes('YOUR_GOOGLE_WEB_APP_URL_HERE')) {
                state.apiUrl = configUrl;
                elements.apiUrlInput.value = configUrl;
                updateStatus('loading', 'Loading data from Google Sheet...');
                fetchData();
                return;
            }
        }
    } catch (e) {
        // config.json fetch might fail due to local file protocol CORS, which is expected
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

    elements.searchInput.addEventListener('input', applyFilters);
    elements.topFilterMonth.addEventListener('change', () => {
        elements.filterMonth.value = elements.topFilterMonth.value;
        applyFilters();
    });
    elements.filterMonth.addEventListener('change', () => {
        elements.topFilterMonth.value = elements.filterMonth.value;
        applyFilters();
    });
    elements.filterPayment.addEventListener('change', applyFilters);
    elements.filterStatus.addEventListener('change', applyFilters);
    elements.clearFiltersBtn.addEventListener('click', clearFilters);
    
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
            logisticsStatus: logisticsStatus
        });
    });
    
    // Sort orders by date descending
    state.orders.sort((a, b) => new Date(b.dateOfOrder) - new Date(a.dateOfOrder));
    state.filteredOrders = [...state.orders];
    
    // Populate filter select options
    populateFilterOptions();
    
    // Populate Dashboard and charts
    renderDashboard();
}

// Extract unique months and logistics statuses for filter dropdowns
function populateFilterOptions() {
    const months = new Set();
    const statuses = new Set();
    
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
    });
    
    // Populate Months
    elements.filterMonth.innerHTML = '<option value="">All Months</option>';
    elements.topFilterMonth.innerHTML = '<option value="">All Months</option>';
    // Chronological order: compile list of months and sort
    const chronMonths = [
        "Jun 2026", "Jul 2026", "Aug 2026", "Sep 2026", "Oct 2026", 
        "Nov 2026", "Dec 2026", "Jan 2027", "Feb 2027", "Mar 2027"
    ];
    chronMonths.forEach(m => {
        if (months.has(m)) {
            const opt = `<option value="${m}">${m}</option>`;
            elements.filterMonth.innerHTML += opt;
            elements.topFilterMonth.innerHTML += opt;
        }
    });
    
    // Populate Statuses
    elements.filterStatus.innerHTML = '<option value="">All Logistics Statuses</option>';
    Array.from(statuses).sort().forEach(s => {
        let label = s;
        if (s === 'UNDELIVERED') {
            label = 'UNDELIVERED (Upto 5 Attempts)';
        }
        elements.filterStatus.innerHTML += `<option value="${s}">${label}</option>`;
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
    const totalOrders = state.filteredOrders.length;
    const totalRevenue = state.filteredOrders.reduce((sum, o) => sum + o.totalPrice, 0);
    const totalRefunded = state.filteredOrders.filter(o => o.returned).reduce((sum, o) => sum + o.totalPrice, 0);
    const totalProfit = totalRevenue - totalRefunded;
    
    const returnCount = state.filteredOrders.filter(o => o.returned).length;
    const returnRate = totalOrders > 0 ? (returnCount / totalOrders) * 100 : 0;
    
    const codOrders = state.filteredOrders.filter(o => o.paymentMethod === 'COD');
    const codDenials = codOrders.filter(o => o.codDenies === 'Yes').length;
    const denialRate = codOrders.length > 0 ? (codDenials / codOrders.length) * 100 : 0;
    
    // Write values
    elements.kpiOrders.textContent = totalOrders;
    elements.kpiRevenue.textContent = formatCurrency(totalRevenue);
    elements.kpiRefunded.textContent = formatCurrency(totalRefunded);
    elements.kpiProfit.textContent = formatCurrency(totalProfit);
    elements.kpiReturnRate.textContent = `${returnRate.toFixed(1)}%`;
    elements.kpiDenialRate.textContent = `${denialRate.toFixed(1)}%`;
    
    elements.kpiReturnRateSub.textContent = `${returnCount} returned of ${totalOrders}`;
    elements.kpiDenialRateSub.textContent = `${codDenials} denials of ${codOrders.length} COD`;
    
    // Progress bar animations
    elements.kpiReturnProgress.style.width = `${returnRate}%`;
    elements.kpiDenialProgress.style.width = `${denialRate}%`;
    
    // 2. Render Logistics Pipeline counts
    const pipeCounts = {
        unfulfilled: 0,
        returned: 0,
        pickup: 0,
        transit: 0,
        delivered: 0
    };
    
    state.filteredOrders.forEach(o => {
        const status = o.logisticsStatus.toUpperCase().trim();
        
        if (o.returned || status.includes('RTO') || status.includes('CANCELED') || status.includes('CANCELLED')) {
            pipeCounts.returned++;
        } else if (status === 'DELIVERED' || status === 'SELF FULFILED') {
            pipeCounts.delivered++;
        } else if (status.includes('TRANSIT') || status.includes('PICKED UP') || status.includes('DELIVERY') || status.includes('HUB') || status.includes('SHIPPED') || status.includes('UNDELIVERED')) {
            pipeCounts.transit++;
        } else if (status.includes('PICKUP') || status.includes('READY TO SHIP')) {
            if (status.includes('EXCEPTION')) {
                pipeCounts.unfulfilled++;
            } else {
                pipeCounts.pickup++;
            }
        } else {
            pipeCounts.unfulfilled++;
        }
    });
    
    elements.pipeUnfulfilled.textContent = pipeCounts.unfulfilled;
    elements.pipeReturned.textContent = pipeCounts.returned;
    elements.pipePickup.textContent = pipeCounts.pickup;
    elements.pipeTransit.textContent = pipeCounts.transit;
    elements.pipeDelivered.textContent = pipeCounts.delivered;
    
    // Animate pipeline states
    document.querySelectorAll('.pipeline-step').forEach(step => {
        const id = step.id.replace('step-', '');
        if (pipeCounts[id] > 0) {
            step.classList.add('active');
        } else {
            step.classList.remove('active');
        }
    });

    // 3. Render charts
    renderCharts();
    
    // 4. Render Table
    renderTable();
}

// Custom Chart.js Plugin to draw leader lines and percentages outside doughnut charts
const donutLabelsLinePlugin = {
    id: 'donutLabelsLine',
    afterDraw(chart) {
        const { ctx } = chart;
        const style = getComputedStyle(document.body);
        const lineStroke = style.getPropertyValue('--text-muted').trim() || 'rgba(128, 128, 128, 0.4)';
        const labelColor = style.getPropertyValue('--text-secondary').trim() || '#475569';
        
        chart.data.datasets.forEach((dataset, i) => {
            const meta = chart.getDatasetMeta(i);
            if (!meta.hidden && chart.config.type === 'doughnut') {
                // Keep track of drawn Y positions to prevent overlaps
                const drawnYLeft = [];
                const drawnYRight = [];
                
                meta.data.forEach((element, index) => {
                    const { x, y, outerRadius, startAngle, endAngle } = element;
                    const value = dataset.data[index];
                    if (!value || value === 0) return; // Skip zero/undefined values
                    
                    // Centroid angle of the slice
                    const midAngle = startAngle + (endAngle - startAngle) / 2;
                    
                    // Coordinates of the slice edge
                    const edgeX = x + Math.cos(midAngle) * outerRadius;
                    const edgeY = y + Math.sin(midAngle) * outerRadius;
                    
                    // Coordinates of the label start (external point)
                    const lineLength = 15; // Length of the pointer line
                    const labelX = x + Math.cos(midAngle) * (outerRadius + lineLength);
                    let labelY = y + Math.sin(midAngle) * (outerRadius + lineLength);
                    
                    const isLeft = Math.cos(midAngle) < 0;
                    const drawnYArray = isLeft ? drawnYLeft : drawnYRight;
                    
                    // Collision detection: adjust Y if too close to another label on the same side
                    const minDistance = 12; // Minimum vertical distance between labels
                    let attempts = 0;
                    let hasCollision = true;
                    
                    while (hasCollision && attempts < 10) {
                        hasCollision = false;
                        for (let j = 0; j < drawnYArray.length; j++) {
                            if (Math.abs(drawnYArray[j] - labelY) < minDistance) {
                                hasCollision = true;
                                // Shift up if in top half, down if in bottom half
                                const isTopHalf = Math.sin(midAngle) < 0;
                                labelY += isTopHalf ? -minDistance : minDistance;
                                break;
                            }
                        }
                        attempts++;
                    }
                    
                    // Boundary check: prevent label from going outside the canvas height
                    if (labelY < 12) labelY = 12;
                    if (labelY > chart.height - 12) labelY = chart.height - 12;
                    
                    drawnYArray.push(labelY);
                    
                    // Draw the pointer line
                    ctx.save();
                    ctx.beginPath();
                    ctx.moveTo(edgeX, edgeY);
                    ctx.lineTo(labelX, labelY);
                    
                    // Draw horizontal tick line
                    const tickLength = 8;
                    const tickX = labelX + (isLeft ? -tickLength : tickLength);
                    ctx.lineTo(tickX, labelY);
                    
                    ctx.strokeStyle = lineStroke;
                    ctx.lineWidth = 1.5;
                    ctx.stroke();
                    
                    // Draw text label next to the tick
                    ctx.fillStyle = labelColor;
                    ctx.font = '600 10px Inter';
                    ctx.textBaseline = 'middle';
                    ctx.textAlign = isLeft ? 'right' : 'left';
                    
                    // Calculate percentage
                    let sum = dataset.data.reduce((a, b) => a + b, 0);
                    let percentage = sum > 0 ? (value * 100 / sum).toFixed(1) + "%" : "0%";
                    
                    // Label text: percentage (value)
                    const labelText = `${percentage} (${value})`;
                    
                    const textX = tickX + (isLeft ? -4 : 4);
                    ctx.fillText(labelText, textX, labelY);
                    ctx.restore();
                });
            }
        });
    }
};

// Generate sales trend and donut breakdown charts
function renderCharts() {
    // Destroy previous chart instances if they exist
    if (state.charts.salesTrend) state.charts.salesTrend.destroy();
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
    state.filteredOrders.forEach(o => {
        const s = o.financialStatus.toLowerCase();
        if (financeCounts[s] !== undefined) {
            financeCounts[s]++;
        }
    });
    
    const ctxFinance = document.getElementById('financeChart').getContext('2d');
    state.charts.finance = new Chart(ctxFinance, {
        type: 'doughnut',
        data: {
            labels: ['Paid', 'Pending', 'Refunded'],
            datasets: [{
                data: [financeCounts.paid, financeCounts.pending, financeCounts.refunded],
                backgroundColor: ['#10b981', '#f97316', '#ef4444'],
                borderColor: borderChart,
                borderWidth: 2,
                radius: '70%'
            }]
        },
        plugins: [donutLabelsLinePlugin],
        options: {
            responsive: true,
            maintainAspectRatio: false,
            layout: {
                padding: 20
            },
            plugins: {
                legend: { position: 'bottom', labels: { color: textSecondary, boxWidth: 12, font: { family: 'Inter' } } }
            },
            cutout: '65%'
        }
    });

    // --- Chart 3: Shiprocket Logistics Status Donut ---
    const deliveryCounts = { delivered: 0, transit: 0, rto: 0, canceled: 0 };
    state.filteredOrders.forEach(o => {
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
                borderWidth: 2,
                radius: '70%'
            }]
        },
        plugins: [donutLabelsLinePlugin],
        options: {
            responsive: true,
            maintainAspectRatio: false,
            layout: {
                padding: 20
            },
            plugins: {
                legend: { position: 'bottom', labels: { color: textSecondary, boxWidth: 12, font: { family: 'Inter' } } }
            },
            cutout: '65%'
        }
    });
}

// Update existing charts color theme without full destroy/recreate
function updateChartsTheme() {
    if (!state.charts.salesTrend && !state.charts.finance && !state.charts.delivery) return;

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
        elements.ordersTbody.innerHTML = `<tr><td colspan="11" class="empty-state">No matching orders found.</td></tr>`;
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
        elements.paginationInfo.textContent = `Page ${state.currentPage} of ${totalPages}`;
        elements.prevPageBtn.disabled = state.currentPage === 1;
        elements.nextPageBtn.disabled = state.currentPage === totalPages;
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
        
        // Row Cells
        tr.innerHTML = `
            <td><strong>#${o.orderNo}</strong></td>
            <td>${formatDateDisplay(o.dateOfOrder)}</td>
            <td>${o.customerName}</td>
            <td class="items-col">${o.itemsOrdered}</td>
            <td>${o.city}</td>
            <td><span class="status-pill info">${o.paymentMethod}</span></td>
            <td class="price-col">${formatCurrency(o.totalPrice)}</td>
            <td><span class="status-pill ${finClass}">${o.financialStatus}</span></td>
            <td><span class="status-pill ${fulClass}">${o.fulfillmentStatus}</span></td>
            <td><span class="status-pill ${logClass}">${o.logisticsStatus}</span></td>
            <td class="comments-col">${o.shiprocketComments || '-'}</td>
        `;
        elements.ordersTbody.appendChild(tr);
    });
}

// Apply inputs and selects to filter the table list
function applyFilters() {
    const q = elements.searchInput.value.toLowerCase().trim();
    const month = elements.filterMonth.value;
    const pay = elements.filterPayment.value;
    const status = elements.filterStatus.value;
    
    state.filteredOrders = state.orders.filter(o => {
        // Search text
        const matchesQuery = !q || 
            String(o.orderNo).toLowerCase().includes(q) ||
            o.customerName.toLowerCase().includes(q) ||
            (o.itemsOrdered && o.itemsOrdered.toLowerCase().includes(q)) ||
            (o.shiprocketComments && o.shiprocketComments.toLowerCase().includes(q)) ||
            o.city.toLowerCase().includes(q) ||
            o.pinCode.toLowerCase().includes(q);
            
        // Month dropdown
        const matchesMonth = !month || getMonthYearStr(o.dateOfOrder) === month;
        
        // Payment dropdown
        const matchesPayment = !pay || o.paymentMethod.toLowerCase().includes(pay.toLowerCase());
        
        // Logistics Status dropdown
        let matchesStatus = false;
        if (!status) {
            matchesStatus = true;
        } else if (status === 'UNDELIVERED') {
            matchesStatus = o.logisticsStatus.toUpperCase().includes('UNDELIVERED');
        } else {
            matchesStatus = o.logisticsStatus.toLowerCase().trim() === status.toLowerCase().trim();
        }
        
        return matchesQuery && matchesMonth && matchesPayment && matchesStatus;
    });
    
    state.currentPage = 1; // Reset to page 1 on filter
    renderDashboard();
}

// Clear all filters
function clearFilters() {
    elements.searchInput.value = '';
    elements.topFilterMonth.value = '';
    elements.filterMonth.value = '';
    elements.filterPayment.value = '';
    elements.filterStatus.value = '';
    state.filteredOrders = [...state.orders];
    state.currentPage = 1; // Reset to page 1
    renderDashboard();
}

// Convert "2026-07-02" to "02 Jul 2026"
function formatDateDisplay(dateStr) {
    if (!dateStr) return '';
    const parts = dateStr.split('-');
    if (parts.length < 3) return dateStr;
    const day = parts[2];
    const year = parts[0];
    const monthIndex = parseInt(parts[1], 10) - 1;
    const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    return `${day} ${months[monthIndex]} ${year}`;
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
    document.getElementById('det-city').textContent = o.city || 'N/A';
    document.getElementById('det-pin').textContent = o.pinCode || 'N/A';
    
    document.getElementById('det-financial').textContent = o.financialStatus || 'N/A';
    document.getElementById('det-fulfillment').textContent = o.fulfillmentStatus || 'N/A';
    document.getElementById('det-logistics').textContent = o.logisticsStatus || 'N/A';
    document.getElementById('det-returned').textContent = o.returned ? 'Yes (Returned)' : 'No';
    document.getElementById('det-cod-denial').textContent = o.codDenies || 'No';
    document.getElementById('det-feedback-sent').textContent = o.feedbackSent || 'No';
    document.getElementById('det-feedback-received').textContent = o.feedbackReceived || 'No';
    
    document.getElementById('det-items').textContent = o.itemsOrdered || 'N/A';
    
    elements.modal.style.display = 'flex';
}
