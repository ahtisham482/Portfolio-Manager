/* ===================================
   PORTFOLIO OPTIMIZER PRO - APP LOGIC
   =================================== */

// Global State
let workbookData = null;
let portfoliosData = null;
let productsMap = {};          // { productName: { good: {name, id}, bad: {name, id} } }
let skippedPortfolios = [];
let campaignSheets = [];
let productsNeedingPrice = [];
let processedResults = {
    movedToGood: [],
    movedToBad: [],
    noAction: [],
    unassigned: []  // Campaigns with spend but no portfolio
};
let modifiedWorkbook = null;
let unassignedCampaigns = [];  // Campaigns without portfolio assignment

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    setupFileUpload();
});

// ===================================
// FILE UPLOAD HANDLING
// ===================================

function setupFileUpload() {
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');

    // Click to upload
    dropZone.addEventListener('click', () => fileInput.click());

    // Drag and drop
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('dragover');
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('dragover');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
        const files = e.dataTransfer.files;
        if (files.length > 0) handleFile(files[0]);
    });

    // File input change
    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) handleFile(e.target.files[0]);
    });
}

function handleFile(file) {
    if (!file.name.match(/\.(xlsx|xls)$/i)) {
        alert('Please upload an Excel file (.xlsx or .xls)');
        return;
    }

    showLoading('Reading Excel file...');

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            workbookData = XLSX.read(e.target.result, { type: 'binary' });

            // Show file info
            document.getElementById('fileInfo').classList.remove('hidden');
            document.querySelector('.file-name').textContent = file.name;

            // Process the workbook
            processWorkbook();
            hideLoading();
        } catch (err) {
            hideLoading();
            alert('Error reading file: ' + err.message);
        }
    };
    reader.readAsBinaryString(file);
}

function removeFile() {
    workbookData = null;
    portfoliosData = null;
    productsMap = {};
    skippedPortfolios = [];
    campaignSheets = [];

    document.getElementById('fileInput').value = '';
    document.getElementById('fileInfo').classList.add('hidden');
    document.getElementById('step-acos').classList.add('hidden');
    document.getElementById('step-price').classList.add('hidden');
    document.getElementById('processSection').classList.add('hidden');
    document.getElementById('step-results').classList.add('hidden');
}

// ===================================
// WORKBOOK PROCESSING
// ===================================

function processWorkbook() {
    const sheetNames = workbookData.SheetNames;

    // Find Portfolios sheet
    const portfolioSheetName = sheetNames.find(name =>
        name.toLowerCase().includes('portfolio')
    );

    if (!portfolioSheetName) {
        alert('Could not find Portfolios worksheet');
        return;
    }

    // Parse portfolios
    const portfolioSheet = workbookData.Sheets[portfolioSheetName];
    portfoliosData = XLSX.utils.sheet_to_json(portfolioSheet);

    // Find campaign sheets
    campaignSheets = sheetNames.filter(name => {
        const lower = name.toLowerCase();
        return lower.includes('sponsored') && lower.includes('campaign');
    });

    // Extract products from portfolios
    extractProducts();

    // Show ACOS input section
    showAcosInputSection();
}

function extractProducts() {
    productsMap = {};
    skippedPortfolios = [];

    // Find column names dynamically
    if (portfoliosData.length === 0) return;

    // Find Portfolio Name column
    const sampleRow = portfoliosData[0];
    let portfolioNameCol = null;
    let portfolioIdCol = null;

    for (let key of Object.keys(sampleRow)) {
        const keyLower = key.toLowerCase();
        if (keyLower.includes('portfolio') && keyLower.includes('name') && !keyLower.includes('informational')) {
            portfolioNameCol = key;
        }
        if (keyLower.includes('portfolio') && keyLower.includes('id')) {
            portfolioIdCol = key;
        }
    }

    if (!portfolioNameCol || !portfolioIdCol) {
        // Try alternative detection
        for (let key of Object.keys(sampleRow)) {
            if (key.toLowerCase() === 'portfolio name') portfolioNameCol = key;
            if (key.toLowerCase() === 'portfolio id') portfolioIdCol = key;
        }
    }

    if (!portfolioNameCol || !portfolioIdCol) {
        alert('Could not find Portfolio Name or Portfolio ID columns');
        return;
    }

    // Process each portfolio
    portfoliosData.forEach(row => {
        const name = row[portfolioNameCol];
        const id = row[portfolioIdCol];

        if (!name || !id) return;

        const nameLower = name.toLowerCase();

        // Check if it has Good Performing/Performance or Bad/Poor Performing/Performance
        const hasGood = nameLower.includes('good performing') || nameLower.includes('good performance');
        const hasBad = nameLower.includes('bad performing') || nameLower.includes('bad performance') ||
            nameLower.includes('poor performing') || nameLower.includes('poor performance');

        if (!hasGood && !hasBad) {
            skippedPortfolios.push(name);
            return;
        }

        // Extract product name (before Good/Bad/Poor Performing/Performance)
        let productName = '';
        const goodPatterns = ['good performing', 'good performance'];
        const badPatterns = ['bad performing', 'bad performance', 'poor performing', 'poor performance'];

        if (hasGood) {
            for (const pattern of goodPatterns) {
                if (nameLower.includes(pattern)) {
                    const idx = nameLower.indexOf(pattern);
                    productName = name.substring(0, idx).trim();
                    break;
                }
            }
        } else if (hasBad) {
            for (const pattern of badPatterns) {
                if (nameLower.includes(pattern)) {
                    const idx = nameLower.indexOf(pattern);
                    productName = name.substring(0, idx).trim();
                    break;
                }
            }
        }

        // Remove trailing separators
        productName = productName.replace(/[-|_:]+$/, '').trim();

        if (!productName) return;

        // Initialize product entry
        if (!productsMap[productName]) {
            productsMap[productName] = { good: null, bad: null };
        }

        // Assign portfolio
        if (hasGood) {
            productsMap[productName].good = { name: name, id: id };
        } else if (hasBad) {
            productsMap[productName].bad = { name: name, id: id };
        }
    });
}

// ===================================
// UI: ACOS INPUT SECTION
// ===================================

function showAcosInputSection() {
    const section = document.getElementById('step-acos');
    const grid = document.getElementById('productsGrid');

    grid.innerHTML = '';

    const productNames = Object.keys(productsMap);

    if (productNames.length === 0) {
        grid.innerHTML = '<p style="color: var(--text-secondary);">No products with Good/Bad Performing portfolios found.</p>';
        section.classList.remove('hidden');
        return;
    }

    productNames.forEach(productName => {
        const card = document.createElement('div');
        card.className = 'product-card';
        card.innerHTML = `
            <div class="product-name">${productName}</div>
            <div class="input-group">
                <input type="number" 
                       id="acos-${sanitizeId(productName)}" 
                       placeholder="Enter ACOS" 
                       min="0" 
                       max="100" 
                       step="0.1"
                       onchange="checkAllInputsFilled()">
                <span class="input-suffix">%</span>
            </div>
        `;
        grid.appendChild(card);
    });

    // Show skipped portfolios
    if (skippedPortfolios.length > 0) {
        const skippedDiv = document.getElementById('skippedPortfolios');
        const skippedList = document.getElementById('skippedList');
        skippedList.innerHTML = '';

        skippedPortfolios.forEach(name => {
            const li = document.createElement('li');
            li.textContent = `"${name}" - Will be skipped`;
            skippedList.appendChild(li);
        });

        skippedDiv.classList.remove('hidden');
    }

    section.classList.remove('hidden');

    // Show process section (will be enabled when all inputs filled)
    document.getElementById('processSection').classList.remove('hidden');
    document.getElementById('processBtn').disabled = true;
}

function sanitizeId(str) {
    return str.replace(/[^a-zA-Z0-9]/g, '_');
}

function checkAllInputsFilled() {
    const productNames = Object.keys(productsMap);
    let allFilled = true;

    productNames.forEach(productName => {
        const input = document.getElementById(`acos-${sanitizeId(productName)}`);
        if (!input || !input.value || parseFloat(input.value) < 0) {
            allFilled = false;
        }
    });

    // Also check price inputs if visible
    const priceSection = document.getElementById('step-price');
    if (!priceSection.classList.contains('hidden')) {
        productsNeedingPrice.forEach(productName => {
            const input = document.getElementById(`price-${sanitizeId(productName)}`);
            if (!input || !input.value || parseFloat(input.value) <= 0) {
                allFilled = false;
            }
        });
    }

    document.getElementById('processBtn').disabled = !allFilled;
}

// ===================================
// CAMPAIGN PROCESSING
// ===================================

function processCampaigns() {
    showLoading('Analyzing campaigns...');

    // Collect ACOS values
    const breakEvenAcos = {};
    Object.keys(productsMap).forEach(productName => {
        const input = document.getElementById(`acos-${sanitizeId(productName)}`);
        breakEvenAcos[productName] = parseFloat(input.value);
    });

    // Collect price values if needed
    const productPrices = {};
    productsNeedingPrice.forEach(productName => {
        const input = document.getElementById(`price-${sanitizeId(productName)}`);
        if (input) {
            productPrices[productName] = parseFloat(input.value);
        }
    });

    // First pass: Check if we need to ask for prices
    if (productsNeedingPrice.length === 0) {
        const needsPriceCheck = checkProductsNeedingPrice(breakEvenAcos);

        if (needsPriceCheck.length > 0) {
            productsNeedingPrice = needsPriceCheck;
            hideLoading();
            showPriceInputSection();
            return;
        }
    }

    // Process all campaign sheets
    processedResults = {
        movedToGood: [],
        movedToBad: [],
        noAction: [],
        unassigned: []
    };

    const processedCampaignIds = new Set(); // Track to avoid duplicates

    campaignSheets.forEach(sheetName => {
        const sheet = workbookData.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(sheet);

        processCampaignSheet(sheetName, data, breakEvenAcos, productPrices, processedCampaignIds);
    });

    hideLoading();
    showResults();
}

function checkProductsNeedingPrice(breakEvenAcos) {
    const productsWithoutAcos = new Set();

    campaignSheets.forEach(sheetName => {
        const sheet = workbookData.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(sheet);

        if (data.length === 0) return;

        // Find column names
        const cols = findCampaignColumns(data[0]);
        if (!cols.entity || !cols.acos) return;

        data.forEach(row => {
            // Only process campaign rows
            if (String(row[cols.entity]).toLowerCase() !== 'campaign') return;

            const acos = parseFloat(row[cols.acos]);
            const spend = cols.spend ? parseFloat(row[cols.spend]) : 0;
            const portfolioName = row[cols.portfolioName] || '';

            // If no ACOS but has spend
            if ((isNaN(acos) || acos === 0) && spend > 0) {
                const productName = extractProductFromPortfolioName(portfolioName);
                if (productName && productsMap[productName]) {
                    productsWithoutAcos.add(productName);
                }
            }
        });
    });

    return Array.from(productsWithoutAcos);
}

function findCampaignColumns(sampleRow) {
    const cols = {
        entity: null,
        campaignId: null,
        portfolioId: null,
        portfolioName: null,
        acos: null,
        spend: null,
        operation: null  // Added for setting 'Update'
    };

    for (let key of Object.keys(sampleRow)) {
        const keyLower = key.toLowerCase();

        if (keyLower === 'entity') cols.entity = key;
        if (keyLower === 'operation') cols.operation = key;  // Detect Operation column
        if (keyLower.includes('campaign') && keyLower.includes('id')) cols.campaignId = key;
        if (keyLower.includes('portfolio') && keyLower.includes('id')) cols.portfolioId = key;
        if (keyLower.includes('portfolio') && keyLower.includes('name') && keyLower.includes('informational')) {
            cols.portfolioName = key;
        }
        if (keyLower === 'acos' || keyLower.includes('advertising cost of sales')) cols.acos = key;
        if (keyLower === 'spend' || keyLower.includes('total spend')) cols.spend = key;
    }

    return cols;
}

function extractProductFromPortfolioName(portfolioName) {
    if (!portfolioName) return null;

    const nameLower = portfolioName.toLowerCase();

    // Define patterns for good and bad portfolios
    const goodPatterns = ['good performing', 'good performance'];
    const badPatterns = ['bad performing', 'bad performance', 'poor performing', 'poor performance'];

    let productName = '';

    // Check good patterns first
    for (const pattern of goodPatterns) {
        if (nameLower.includes(pattern)) {
            const idx = nameLower.indexOf(pattern);
            productName = portfolioName.substring(0, idx).trim();
            break;
        }
    }

    // If not found, check bad patterns
    if (!productName) {
        for (const pattern of badPatterns) {
            if (nameLower.includes(pattern)) {
                const idx = nameLower.indexOf(pattern);
                productName = portfolioName.substring(0, idx).trim();
                break;
            }
        }
    }

    // Remove trailing separators
    productName = productName.replace(/[-|_:]+$/, '').trim();

    return productName || null;
}

function processCampaignSheet(sheetName, data, breakEvenAcos, productPrices, processedCampaignIds) {
    if (data.length === 0) return;

    const cols = findCampaignColumns(data[0]);

    if (!cols.entity) {
        console.log(`Skipping ${sheetName}: No Entity column found`);
        return;
    }

    data.forEach((row, index) => {
        // Only process campaign rows
        if (String(row[cols.entity]).toLowerCase() !== 'campaign') return;

        // Get Campaign ID - handle numeric values and empty strings properly
        let campaignId = row[cols.campaignId];
        if (campaignId === undefined || campaignId === null || campaignId === '') {
            // Try Campaign Name as fallback
            const campaignNameCol = Object.keys(row).find(k =>
                k.toLowerCase() === 'campaign name' ||
                (k.toLowerCase().includes('campaign') && k.toLowerCase().includes('name') && !k.toLowerCase().includes('informational'))
            );
            campaignId = (campaignNameCol && row[campaignNameCol]) ? row[campaignNameCol] : `Row-${index + 2}`;
        }
        campaignId = String(campaignId);  // Convert to string for consistency

        // Skip duplicates
        if (processedCampaignIds.has(campaignId)) return;
        processedCampaignIds.add(campaignId);

        const currentPortfolioId = row[cols.portfolioId];
        const portfolioName = row[cols.portfolioName] || '';
        let acos = parseFloat(row[cols.acos]);
        const spend = cols.spend ? parseFloat(row[cols.spend]) : 0;

        // Convert ACOS from decimal to percentage if needed (e.g., 0.19 -> 19)
        if (!isNaN(acos) && acos > 0 && acos < 1) {
            acos = acos * 100;
        }

        // Extract product name
        const productName = extractProductFromPortfolioName(portfolioName);

        // Check if campaign has no portfolio assigned but has spend
        const hasNoPortfolio = !currentPortfolioId || currentPortfolioId === '' || portfolioName === '';

        if (hasNoPortfolio && spend > 0) {
            // Get campaign name for display
            const campaignNameCol = Object.keys(row).find(k =>
                k.toLowerCase() === 'campaign name' ||
                (k.toLowerCase().includes('campaign') && k.toLowerCase().includes('name') && !k.toLowerCase().includes('informational'))
            );
            const campaignName = campaignNameCol ? row[campaignNameCol] : 'Unknown';

            processedResults.unassigned.push({
                campaignId,
                campaignName,
                sheetName,
                rowIndex: index,
                row: row,
                cols: cols,
                acos: acos,
                spend: spend,
                acosSpend: !isNaN(acos) && acos > 0 ? `${acos.toFixed(2)}%` : `$${spend.toFixed(2)} Spend`,
                assigned: false,
                assignedPortfolio: null
            });
            return;
        }

        if (!productName || !productsMap[productName]) {
            processedResults.noAction.push({
                campaignId,
                sheetName,
                rowIndex: index,
                portfolioName,
                acosSpend: !isNaN(acos) && acos > 0 ? `${acos.toFixed(2)}%` : (spend > 0 ? `$${spend.toFixed(2)} Spend` : 'N/A'),
                reason: 'Portfolio not found or does not have Good/Bad Performing'
            });
            return;
        }

        const breakEven = breakEvenAcos[productName];
        const productPrice = productPrices[productName] || 0;

        // Determine current portfolio type
        const nameLower = portfolioName.toLowerCase();
        const isCurrentlyGood = nameLower.includes('good performing') || nameLower.includes('good performance');
        const isCurrentlyBad = nameLower.includes('bad performing') || nameLower.includes('bad performance') ||
            nameLower.includes('poor performing') || nameLower.includes('poor performance');

        // Calculate performance
        let shouldBeGood = null;
        let acosSpendDisplay = '';

        if (!isNaN(acos) && acos > 0) {
            // Has ACOS - display in proper percentage format
            acosSpendDisplay = `${acos.toFixed(2)}%`;

            if (acos < breakEven) {
                shouldBeGood = true;
            } else if (acos > breakEven) {
                shouldBeGood = false;
            } else {
                // ACOS equals break-even
                processedResults.noAction.push({
                    campaignId,
                    sheetName,
                    rowIndex: index,
                    portfolioName,
                    acosSpend: acosSpendDisplay,
                    reason: `ACOS equals break-even (${breakEven}%)`
                });
                return;
            }
        } else if (spend > 0 && productPrice > 0) {
            // No ACOS but has spend, use price calculation
            const maxSpend = productPrice * (breakEven / 100);
            acosSpendDisplay = `$${spend.toFixed(2)} Spend`;

            if (spend < maxSpend) {
                shouldBeGood = true;
            } else if (spend > maxSpend) {
                shouldBeGood = false;
            } else {
                processedResults.noAction.push({
                    campaignId,
                    sheetName,
                    rowIndex: index,
                    portfolioName,
                    acosSpend: acosSpendDisplay,
                    reason: `Spend equals max allowed ($${maxSpend.toFixed(2)})`
                });
                return;
            }
        } else {
            // No ACOS and no spend
            processedResults.noAction.push({
                campaignId,
                sheetName,
                rowIndex: index,
                portfolioName,
                acosSpend: 'N/A',
                reason: 'No ACOS or Spend data'
            });
            return;
        }

        // Determine if move is needed
        const goodPortfolio = productsMap[productName].good;
        const badPortfolio = productsMap[productName].bad;

        if (shouldBeGood && !isCurrentlyGood) {
            // Move to Good
            if (goodPortfolio) {
                row[cols.portfolioId] = goodPortfolio.id;
                // Set Operation to Update for Amazon bulk upload
                if (cols.operation) row[cols.operation] = 'Update';
                processedResults.movedToGood.push({
                    campaignId,
                    sheetName,
                    rowIndex: index,
                    row: row,
                    fromPortfolio: portfolioName,
                    toPortfolio: goodPortfolio.name,
                    acosSpend: acosSpendDisplay
                });
            } else {
                processedResults.noAction.push({
                    campaignId,
                    sheetName,
                    rowIndex: index,
                    portfolioName,
                    acosSpend: acosSpendDisplay,
                    reason: 'No Good Performing portfolio found for product'
                });
            }
        } else if (!shouldBeGood && !isCurrentlyBad) {
            // Move to Bad
            if (badPortfolio) {
                row[cols.portfolioId] = badPortfolio.id;
                // Set Operation to Update for Amazon bulk upload
                if (cols.operation) row[cols.operation] = 'Update';
                processedResults.movedToBad.push({
                    campaignId,
                    sheetName,
                    rowIndex: index,
                    row: row,
                    fromPortfolio: portfolioName,
                    toPortfolio: badPortfolio.name,
                    acosSpend: acosSpendDisplay
                });
            } else {
                processedResults.noAction.push({
                    campaignId,
                    sheetName,
                    rowIndex: index,
                    portfolioName,
                    acosSpend: acosSpendDisplay,
                    reason: 'No Bad Performing portfolio found for product'
                });
            }
        } else {
            // Already in correct portfolio
            processedResults.noAction.push({
                campaignId,
                sheetName,
                rowIndex: index,
                portfolioName,
                acosSpend: acosSpendDisplay,
                reason: 'Already in correct portfolio'
            });
        }
    });
}

// ===================================
// UI: PRICE INPUT SECTION
// ===================================

function showPriceInputSection() {
    const section = document.getElementById('step-price');
    const grid = document.getElementById('priceGrid');

    grid.innerHTML = '';

    productsNeedingPrice.forEach(productName => {
        const card = document.createElement('div');
        card.className = 'product-card';
        card.innerHTML = `
            <div class="product-name">${productName}</div>
            <div class="input-group">
                <span class="input-prefix">$</span>
                <input type="number" 
                       id="price-${sanitizeId(productName)}" 
                       placeholder="Enter Price" 
                       min="0.01" 
                       step="0.01"
                       onchange="checkAllInputsFilled()">
            </div>
        `;
        grid.appendChild(card);
    });

    section.classList.remove('hidden');
    document.getElementById('processBtn').disabled = true;
}

// ===================================
// UI: RESULTS SECTION
// ===================================

function showResults() {
    // Update counts
    document.getElementById('movedGoodCount').textContent = processedResults.movedToGood.length;
    document.getElementById('movedBadCount').textContent = processedResults.movedToBad.length;
    document.getElementById('noActionCount').textContent = processedResults.noAction.length;
    document.getElementById('unassignedCount').textContent = processedResults.unassigned.length;

    // Populate Good table
    const goodBody = document.getElementById('goodTableBody');
    goodBody.innerHTML = '';
    processedResults.movedToGood.forEach(item => {
        goodBody.innerHTML += `
            <tr>
                <td>${item.campaignId}</td>
                <td>${item.fromPortfolio}</td>
                <td>${item.toPortfolio}</td>
                <td>${item.acosSpend}</td>
            </tr>
        `;
    });
    if (processedResults.movedToGood.length === 0) {
        goodBody.innerHTML = '<tr><td colspan="4" style="text-align:center; color: var(--text-muted);">No campaigns moved to Good Performing</td></tr>';
    }

    // Populate Bad table
    const badBody = document.getElementById('badTableBody');
    badBody.innerHTML = '';
    processedResults.movedToBad.forEach(item => {
        badBody.innerHTML += `
            <tr>
                <td>${item.campaignId}</td>
                <td>${item.fromPortfolio}</td>
                <td>${item.toPortfolio}</td>
                <td>${item.acosSpend}</td>
            </tr>
        `;
    });
    if (processedResults.movedToBad.length === 0) {
        badBody.innerHTML = '<tr><td colspan="4" style="text-align:center; color: var(--text-muted);">No campaigns moved to Bad Performing</td></tr>';
    }

    // Populate No Action table
    const noActionBody = document.getElementById('noActionTableBody');
    noActionBody.innerHTML = '';
    processedResults.noAction.forEach(item => {
        noActionBody.innerHTML += `
            <tr>
                <td>${item.campaignId}</td>
                <td>${item.portfolioName}</td>
                <td>${item.acosSpend}</td>
                <td>${item.reason}</td>
            </tr>
        `;
    });
    if (processedResults.noAction.length === 0) {
        noActionBody.innerHTML = '<tr><td colspan="4" style="text-align:center; color: var(--text-muted);">All campaigns were moved</td></tr>';
    }

    // Populate Unassigned table
    const unassignedBody = document.getElementById('unassignedTableBody');
    unassignedBody.innerHTML = '';

    // Populate product dropdown
    const productSelect = document.getElementById('assignProductSelect');
    productSelect.innerHTML = '<option value="">-- Select Product --</option>';
    Object.keys(productsMap).forEach(productName => {
        productSelect.innerHTML += `<option value="${productName}">${productName}</option>`;
    });

    processedResults.unassigned.forEach((item, idx) => {
        unassignedBody.innerHTML += `
            <tr>
                <td><input type="checkbox" class="unassigned-checkbox" data-index="${idx}" checked></td>
                <td>${item.campaignId}</td>
                <td>${item.campaignName}</td>
                <td>${item.acosSpend}</td>
                <td id="assign-preview-${idx}">-</td>
            </tr>
        `;
    });
    if (processedResults.unassigned.length === 0) {
        unassignedBody.innerHTML = '<tr><td colspan="5" style="text-align:center; color: var(--text-muted);">No unassigned campaigns found</td></tr>';
    }

    // Show results section
    document.getElementById('step-results').classList.remove('hidden');

    // Scroll to results
    document.getElementById('step-results').scrollIntoView({ behavior: 'smooth' });
}

function showTab(tabName) {
    // Update tab buttons
    document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
    event.target.classList.add('active');

    // Update tab contents
    document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));

    if (tabName === 'good') document.getElementById('tabGood').classList.add('active');
    if (tabName === 'bad') document.getElementById('tabBad').classList.add('active');
    if (tabName === 'noaction') document.getElementById('tabNoaction').classList.add('active');
    if (tabName === 'unassigned') document.getElementById('tabUnassigned').classList.add('active');
}

// Toggle select all unassigned checkboxes
function toggleSelectAllUnassigned() {
    const selectAll = document.getElementById('selectAllUnassigned').checked;
    document.querySelectorAll('.unassigned-checkbox').forEach(cb => {
        cb.checked = selectAll;
    });
}

// Assign unassigned campaigns to selected product
function assignUnassignedCampaigns() {
    const productName = document.getElementById('assignProductSelect').value;

    if (!productName) {
        alert('Please select a product to assign campaigns to.');
        return;
    }

    const product = productsMap[productName];
    if (!product) {
        alert('Product not found.');
        return;
    }

    // Get break-even ACOS for this product
    const acosInput = document.getElementById(`acos-${sanitizeId(productName)}`);
    const breakEven = acosInput ? parseFloat(acosInput.value) : 25;

    // Get price for this product if needed
    const priceInput = document.getElementById(`price-${sanitizeId(productName)}`);
    const productPrice = priceInput ? parseFloat(priceInput.value) : 0;

    let assignedCount = 0;

    // Process each checked unassigned campaign
    document.querySelectorAll('.unassigned-checkbox:checked').forEach(cb => {
        const idx = parseInt(cb.dataset.index);
        const item = processedResults.unassigned[idx];

        if (!item) return;

        // Determine if should go to Good or Bad based on ACOS/Spend
        let shouldBeGood = true;

        if (!isNaN(item.acos) && item.acos > 0) {
            shouldBeGood = item.acos < breakEven;
        } else if (item.spend > 0 && productPrice > 0) {
            const maxSpend = productPrice * (breakEven / 100);
            shouldBeGood = item.spend < maxSpend;
        }

        // Assign portfolio
        const targetPortfolio = shouldBeGood ? product.good : product.bad;

        if (targetPortfolio) {
            item.row[item.cols.portfolioId] = targetPortfolio.id;
            // Set Operation to Update for Amazon bulk upload
            if (item.cols.operation) item.row[item.cols.operation] = 'Update';
            item.assigned = true;
            item.assignedPortfolio = targetPortfolio.name;

            // Update preview in table
            const previewCell = document.getElementById(`assign-preview-${idx}`);
            if (previewCell) {
                previewCell.textContent = targetPortfolio.name;
                previewCell.style.color = shouldBeGood ? 'var(--success)' : 'var(--danger)';
            }

            assignedCount++;
        }
    });

    alert(`${assignedCount} campaigns assigned to ${productName} portfolios. Click Download to get the file.`);
}

// ===================================
// FILE DOWNLOAD
// ===================================

function downloadFile() {
    showLoading('Generating Excel file...');

    try {
        // Create new workbook with separate sheets for SP and SB
        const newWorkbook = XLSX.utils.book_new();

        // Separate campaigns by sheet type
        const spCampaigns = [];  // Sponsored Products
        const sbCampaigns = [];  // Sponsored Brands

        // Helper function to categorize campaigns
        function addCampaign(item) {
            if (!item || !item.row) return;

            const sheetName = item.sheetName ? item.sheetName.toLowerCase() : '';

            if (sheetName.includes('sponsored products')) {
                spCampaigns.push(item.row);
            } else if (sheetName.includes('sponsored brands')) {
                sbCampaigns.push(item.row);
            } else {
                // Default to SP if unknown
                spCampaigns.push(item.row);
            }
        }

        // Add moved to good campaigns
        processedResults.movedToGood.forEach(item => addCampaign(item));

        // Add moved to bad campaigns
        processedResults.movedToBad.forEach(item => addCampaign(item));

        // Add newly assigned campaigns (unassigned that user assigned)
        if (processedResults.unassigned) {
            processedResults.unassigned.forEach(item => {
                if (item.assigned) addCampaign(item);
            });
        }

        if (spCampaigns.length === 0 && sbCampaigns.length === 0) {
            hideLoading();
            alert('No campaigns were changed. Nothing to download.');
            return;
        }

        // Create Sponsored Products worksheet
        if (spCampaigns.length > 0) {
            const spSheet = XLSX.utils.json_to_sheet(spCampaigns);
            XLSX.utils.book_append_sheet(newWorkbook, spSheet, 'Sponsored Products Campaigns');
        }

        // Create Sponsored Brands worksheet
        if (sbCampaigns.length > 0) {
            const sbSheet = XLSX.utils.json_to_sheet(sbCampaigns);
            XLSX.utils.book_append_sheet(newWorkbook, sbSheet, 'Sponsored Brands Campaigns');
        }

        // Generate file
        const today = new Date();
        const dateStr = today.toISOString().slice(0, 10);
        const filename = `Portfolio_Optimized_${dateStr}.xlsx`;

        XLSX.writeFile(newWorkbook, filename);

        hideLoading();
    } catch (err) {
        hideLoading();
        alert('Error generating file: ' + err.message);
        console.error(err);
    }
}

// ===================================
// UTILITY FUNCTIONS
// ===================================

function showLoading(message = 'Processing...') {
    const overlay = document.getElementById('loadingOverlay');
    overlay.querySelector('.loading-text').textContent = message;
    overlay.classList.remove('hidden');
}

function hideLoading() {
    document.getElementById('loadingOverlay').classList.add('hidden');
}
