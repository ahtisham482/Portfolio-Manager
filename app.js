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

// Debug mode state
let debugDecisions = [];       // Per-campaign decision logs
let debugPortfolioLog = [];    // Portfolio extraction log
let debugFileStats = {};       // File upload stats

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

    // Reset debug panel
    document.getElementById('debugPanel').style.display = 'none';
    debugDecisions = [];
    debugPortfolioLog = [];
    debugFileStats = {};
    document.getElementById('debugUploadDetails').innerHTML = '';
    document.getElementById('debugPortfolioDetails').innerHTML = '';
    document.getElementById('debugAcosDetails').innerHTML = '<p style="font-size:13px; color:#6b7280; font-style:italic;">Click "Process Campaigns" to see config.</p>';
    document.getElementById('debugDecisionsDetails').innerHTML = '<p style="font-size:13px; color:#6b7280; font-style:italic;">Click "Process Campaigns" to see decisions.</p>';
    document.getElementById('debugResultsDetails').innerHTML = '<p style="font-size:13px; color:#6b7280; font-style:italic;">Click "Process Campaigns" to see results.</p>';
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

    // Show debug panel with Chapter 1 & 2
    document.getElementById('debugPanel').style.display = 'block';
    renderDebugChapter1();
    renderDebugChapter2();
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

        // Log for debug
        debugPortfolioLog.push({
            portfolioName: name,
            portfolioId: id,
            productName: productName,
            type: hasGood ? 'Good Performing' : 'Bad Performing'
        });
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

    // Also check price inputs if visible (only for non-skipped products)
    const priceSection = document.getElementById('step-price');
    if (!priceSection.classList.contains('hidden')) {
        productsNeedingPrice.forEach(productName => {
            const skipCheckbox = document.getElementById(`skip-${sanitizeId(productName)}`);
            const isSkipped = skipCheckbox && skipCheckbox.checked;

            // Only require price for non-skipped products
            if (!isSkipped) {
                const input = document.getElementById(`price-${sanitizeId(productName)}`);
                if (!input || !input.value || parseFloat(input.value) <= 0) {
                    allFilled = false;
                }
            }
        });
    }

    document.getElementById('processBtn').disabled = !allFilled;
}

// Apply ACOS value to all product inputs
function applyAcosToAll() {
    const applyAllInput = document.getElementById('applyAllAcosInput');
    const value = parseFloat(applyAllInput.value);

    if (isNaN(value) || value < 0 || value > 100) {
        alert('Please enter a valid ACOS value between 0 and 100');
        return;
    }

    const productNames = Object.keys(productsMap);
    productNames.forEach(productName => {
        const input = document.getElementById(`acos-${sanitizeId(productName)}`);
        if (input) {
            input.value = value;
        }
    });

    // Check if all inputs are now filled
    checkAllInputsFilled();
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

    // Collect price values if needed (excluding skipped products)
    const productPrices = {};
    const skippedProducts = new Set();
    productsNeedingPrice.forEach(productName => {
        const skipCheckbox = document.getElementById(`skip-${sanitizeId(productName)}`);
        const isSkipped = skipCheckbox && skipCheckbox.checked;

        if (isSkipped) {
            skippedProducts.add(productName);
        } else {
            const input = document.getElementById(`price-${sanitizeId(productName)}`);
            if (input && input.value) {
                productPrices[productName] = parseFloat(input.value);
            }
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

    debugDecisions = [];  // Reset decisions log

    const processedCampaignIds = new Set(); // Track to avoid duplicates

    campaignSheets.forEach(sheetName => {
        const sheet = workbookData.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(sheet);

        processCampaignSheet(sheetName, data, breakEvenAcos, productPrices, processedCampaignIds);
    });

    hideLoading();
    showResults();

    // Render debug Chapters 3, 4, 5
    renderDebugChapter3(breakEvenAcos, productPrices, skippedProducts);
    renderDebugChapter4();
    renderDebugChapter5();
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
        operation: null,  // Added for setting 'Update'
        sales: null,      // For auto-price calculation
        units: null,      // For auto-price calculation
        campaignName: null // For extracting product name
    };

    for (let key of Object.keys(sampleRow)) {
        const keyLower = key.toLowerCase();

        if (keyLower === 'entity') cols.entity = key;
        if (keyLower === 'operation') cols.operation = key;
        if (keyLower.includes('campaign') && keyLower.includes('id')) cols.campaignId = key;
        if (keyLower.includes('portfolio') && keyLower.includes('id')) cols.portfolioId = key;
        if (keyLower.includes('portfolio') && keyLower.includes('name') && keyLower.includes('informational')) {
            cols.portfolioName = key;
        }
        if (keyLower === 'acos' || keyLower.includes('advertising cost of sales')) cols.acos = key;
        if (keyLower === 'spend' || keyLower.includes('total spend')) cols.spend = key;
        // Detect Sales column (exact match or contains 'sales' but not other metrics)
        if (keyLower === 'sales' || (keyLower.includes('sales') && !keyLower.includes('cost'))) {
            cols.sales = key;
        }
        // Detect Units column
        if (keyLower === 'units' || keyLower.includes('units sold')) {
            cols.units = key;
        }
        // Detect Campaign Name column (not informational)
        if ((keyLower === 'campaign name' || (keyLower.includes('campaign') && keyLower.includes('name'))) && !keyLower.includes('informational')) {
            cols.campaignName = key;
        }
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

            debugDecisions.push({
                campaignId, sheetName, portfolioName: '(none)',
                acos: !isNaN(acos) && acos > 0 ? acos.toFixed(2) + '%' : 'N/A',
                spend: spend > 0 ? '$' + spend.toFixed(2) : 'N/A',
                breakEven: 'N/A', currentType: 'N/A', targetPortfolio: 'N/A',
                decision: 'UNASSIGNED',
                reason: 'Campaign has spend but no portfolio assigned'
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

            debugDecisions.push({
                campaignId, sheetName, portfolioName,
                acos: !isNaN(acos) && acos > 0 ? acos.toFixed(2) + '%' : 'N/A',
                spend: spend > 0 ? '$' + spend.toFixed(2) : 'N/A',
                breakEven: 'N/A', currentType: 'N/A', targetPortfolio: 'N/A',
                decision: 'NO ACTION',
                reason: 'Portfolio not recognized (no Good/Bad Performing match)'
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

                debugDecisions.push({
                    campaignId, sheetName, portfolioName,
                    acos: acosSpendDisplay, spend: 'N/A',
                    breakEven: breakEven + '%',
                    currentType: isCurrentlyGood ? 'Good' : 'Bad',
                    targetPortfolio: portfolioName,
                    decision: 'NO ACTION',
                    reason: `ACOS (${acos.toFixed(2)}%) = Break-Even (${breakEven}%)`
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

            debugDecisions.push({
                campaignId, sheetName, portfolioName,
                acos: 'N/A', spend: 'N/A',
                breakEven: breakEven + '%',
                currentType: isCurrentlyGood ? 'Good' : 'Bad',
                targetPortfolio: portfolioName,
                decision: 'NO ACTION',
                reason: 'No ACOS or Spend data available'
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

                debugDecisions.push({
                    campaignId, sheetName, portfolioName,
                    acos: acosSpendDisplay, spend: spend > 0 ? '$' + spend.toFixed(2) : 'N/A',
                    breakEven: breakEven + '%',
                    currentType: isCurrentlyBad ? 'Bad' : 'Other',
                    targetPortfolio: goodPortfolio.name,
                    decision: 'MOVE TO GOOD',
                    reason: `ACOS (${acosSpendDisplay}) < Break-Even (${breakEven}%)`
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

                debugDecisions.push({
                    campaignId, sheetName, portfolioName,
                    acos: acosSpendDisplay, spend: spend > 0 ? '$' + spend.toFixed(2) : 'N/A',
                    breakEven: breakEven + '%',
                    currentType: isCurrentlyGood ? 'Good' : 'Other',
                    targetPortfolio: badPortfolio.name,
                    decision: 'MOVE TO BAD',
                    reason: `ACOS (${acosSpendDisplay}) > Break-Even (${breakEven}%)`
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

                debugDecisions.push({
                    campaignId, sheetName, portfolioName,
                    acos: acosSpendDisplay, spend: spend > 0 ? '$' + spend.toFixed(2) : 'N/A',
                    breakEven: breakEven + '%',
                    currentType: isCurrentlyGood ? 'Good' : 'Other',
                    targetPortfolio: 'N/A',
                    decision: 'NO ACTION',
                    reason: 'No Bad Performing portfolio found for this product'
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

            debugDecisions.push({
                campaignId, sheetName, portfolioName,
                acos: acosSpendDisplay, spend: spend > 0 ? '$' + spend.toFixed(2) : 'N/A',
                breakEven: breakEven + '%',
                currentType: isCurrentlyGood ? 'Good' : 'Bad',
                targetPortfolio: portfolioName,
                decision: 'NO ACTION',
                reason: 'Already in correct portfolio'
            });
        }
    });
}

// ===================================
// UI: PRICE INPUT SECTION
// ===================================

// Calculate product prices from campaign Sales/Units data
function calculateProductPrices() {
    const calculatedPrices = {};

    // Build a map of product names (from portfolio) to their variants in campaign names
    // We need to find campaigns for each product that needs a price

    campaignSheets.forEach(sheetName => {
        const sheet = workbookData.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(sheet);

        if (data.length === 0) return;

        const cols = findCampaignColumns(data[0]);
        if (!cols.entity || !cols.sales || !cols.units || !cols.campaignName) return;

        data.forEach(row => {
            // Only process campaign rows
            if (String(row[cols.entity]).toLowerCase() !== 'campaign') return;

            const sales = parseFloat(row[cols.sales]);
            const units = parseFloat(row[cols.units]);
            const campaignName = row[cols.campaignName] || '';

            // Skip if no valid sales/units data
            if (isNaN(sales) || isNaN(units) || sales <= 0 || units <= 0) return;

            // Extract product name from campaign name (text before first |)
            const pipeIndex = campaignName.indexOf('|');
            let campaignProductName = pipeIndex > 0 ? campaignName.substring(0, pipeIndex).trim() : campaignName.trim();

            // Remove trailing dash and extra spaces
            campaignProductName = campaignProductName.replace(/[-]+$/, '').trim();

            // Try to match this to a product that needs price
            // We'll check if any product needing price is contained in or matches the campaign product name
            productsNeedingPrice.forEach(productName => {
                // Skip if we already calculated price for this product
                if (calculatedPrices[productName]) return;

                // Check for match (case-insensitive)
                const productLower = productName.toLowerCase();
                const campaignProductLower = campaignProductName.toLowerCase();

                // Match if campaign product contains the product name or vice versa
                // Or if they share the same starting characters (like "TU" matching "TU - Silver")
                if (campaignProductLower.includes(productLower) ||
                    productLower.includes(campaignProductLower) ||
                    campaignProductLower.startsWith(productLower.split(' ')[0])) {

                    // Calculate price
                    const price = sales / units;
                    calculatedPrices[productName] = Math.round(price * 100) / 100; // Round to 2 decimal places
                }
            });
        });
    });

    return calculatedPrices;
}

function showPriceInputSection() {
    const section = document.getElementById('step-price');
    const grid = document.getElementById('priceGrid');

    grid.innerHTML = '';

    // Calculate prices from bulk file data (for auto-fill where possible)
    const calculatedPrices = calculateProductPrices();

    // Show ALL products - no auto-skipping
    // User will decide which to skip using checkboxes
    productsNeedingPrice.forEach(productName => {
        const calculatedPrice = calculatedPrices[productName];
        const hasCalculatedPrice = calculatedPrice && calculatedPrice > 0;

        const card = document.createElement('div');
        card.className = 'product-card';
        card.innerHTML = `
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                <div class="product-name" style="margin-bottom: 0;">${productName}</div>
                <label style="display: flex; align-items: center; gap: 6px; cursor: pointer; font-size: 12px; color: var(--text-secondary);">
                    <input type="checkbox" 
                           id="skip-${sanitizeId(productName)}" 
                           class="skip-price-checkbox"
                           onchange="handleSkipToggle('${sanitizeId(productName)}')"
                           style="cursor: pointer;">
                    Skip
                </label>
            </div>
            <div class="input-group" id="price-group-${sanitizeId(productName)}">
                <span class="input-prefix">$</span>
                <input type="number" 
                       id="price-${sanitizeId(productName)}" 
                       placeholder="Enter Price" 
                       min="0.01" 
                       step="0.01"
                       value="${hasCalculatedPrice ? calculatedPrice : ''}"
                       onchange="checkAllInputsFilled()">
            </div>
            ${hasCalculatedPrice ? '<div style="font-size: 11px; color: var(--success); margin-top: 4px;">✓ Auto-calculated from campaign data</div>' : '<div style="font-size: 11px; color: var(--warning); margin-top: 4px;">⚠️ No matching campaign data found - enter price manually or skip</div>'}
        `;
        grid.appendChild(card);
    });

    section.classList.remove('hidden');

    // Check if all inputs are filled (considering skip checkboxes)
    checkAllInputsFilled();
}

// Handle skip checkbox toggle - disable/enable price input
function handleSkipToggle(sanitizedProductName) {
    const checkbox = document.getElementById(`skip-${sanitizedProductName}`);
    const priceInput = document.getElementById(`price-${sanitizedProductName}`);
    const priceGroup = document.getElementById(`price-group-${sanitizedProductName}`);

    if (checkbox && priceInput && priceGroup) {
        if (checkbox.checked) {
            priceInput.disabled = true;
            priceGroup.style.opacity = '0.5';
        } else {
            priceInput.disabled = false;
            priceGroup.style.opacity = '1';
        }
    }

    checkAllInputsFilled();
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

// ===================================
// DEBUG MODE FUNCTIONS
// ===================================

function toggleDebugChapter(name) {
    const chapterMap = { upload: 'Upload', portfolio: 'Portfolio', acos: 'Acos', decisions: 'Decisions', results: 'Results' };
    const el = document.getElementById('chapter' + chapterMap[name]);
    const toggle = document.getElementById(name + 'Toggle');
    if (!el) return;
    if (el.style.display === 'block') {
        el.style.display = 'none';
        toggle.textContent = '\u25BC';
    } else {
        el.style.display = 'block';
        toggle.textContent = '\u25B2';
    }
}

function esc(str) {
    if (typeof str !== 'string') return '';
    return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
}

function renderDebugChapter1() {
    const sheetNames = workbookData ? workbookData.SheetNames : [];
    const portfolioSheetName = sheetNames.find(n => n.toLowerCase().includes('portfolio')) || 'N/A';
    const totalPortfolios = portfoliosData ? portfoliosData.length : 0;
    const totalProducts = Object.keys(productsMap).length;
    const totalSkipped = skippedPortfolios.length;

    // Count total campaign rows across all campaign sheets
    let totalCampaignRows = 0;
    campaignSheets.forEach(name => {
        const sheet = workbookData.Sheets[name];
        const data = XLSX.utils.sheet_to_json(sheet);
        totalCampaignRows += data.filter(r => {
            const entityCol = Object.keys(r).find(k => k.toLowerCase() === 'entity');
            return entityCol && String(r[entityCol]).toLowerCase() === 'campaign';
        }).length;
    });

    document.getElementById('debugUploadCount').textContent = `(${sheetNames.length} sheets)`;

    document.getElementById('debugUploadDetails').innerHTML = `
        <div class="debug-summary-grid">
            <div class="debug-stat"><div class="stat-label">Total Sheets</div><div class="stat-value">${sheetNames.length}</div></div>
            <div class="debug-stat"><div class="stat-label">Portfolio Sheet</div><div class="stat-value" style="font-size:13px;">${esc(portfolioSheetName)}</div></div>
            <div class="debug-stat"><div class="stat-label">Total Portfolios</div><div class="stat-value">${totalPortfolios}</div></div>
            <div class="debug-stat"><div class="stat-label">Campaign Sheets</div><div class="stat-value">${campaignSheets.length}</div></div>
            <div class="debug-stat"><div class="stat-label">Campaign Rows</div><div class="stat-value">${totalCampaignRows}</div></div>
            <div class="debug-stat"><div class="stat-label">Products Found</div><div class="stat-value" style="color:#4ade80;">${totalProducts}</div></div>
            <div class="debug-stat"><div class="stat-label">Skipped Portfolios</div><div class="stat-value" style="color:#f87171;">${totalSkipped}</div></div>
        </div>
        <div style="margin-top:12px;">
            <div style="font-size:11px; color:#6b7280; text-transform:uppercase; margin-bottom:4px;">All Sheet Names</div>
            <div style="font-size:12px; color:#94a3b8;">${sheetNames.map(n => esc(n)).join(' &bull; ')}</div>
        </div>`;
}

function renderDebugChapter2() {
    const entries = debugPortfolioLog;
    const productNames = Object.keys(productsMap);

    document.getElementById('debugPortfolioCount').textContent = `(${productNames.length} products, ${entries.length} portfolios)`;

    let html = '';

    // Products with their Good/Bad portfolios
    html += '<table class="debug-table"><thead><tr><th>Product</th><th>Good Portfolio</th><th>Bad Portfolio</th></tr></thead><tbody>';
    productNames.forEach(name => {
        const p = productsMap[name];
        html += `<tr>`;
        html += `<td style="font-weight:600;">${esc(name)}</td>`;
        html += `<td>${p.good ? esc(p.good.name) : '<span style="color:#f87171;">Missing</span>'}</td>`;
        html += `<td>${p.bad ? esc(p.bad.name) : '<span style="color:#f87171;">Missing</span>'}</td>`;
        html += `</tr>`;
    });
    html += '</tbody></table>';

    // Skipped portfolios
    if (skippedPortfolios.length > 0) {
        html += '<div style="margin-top:14px; padding:10px; background:rgba(250,204,21,0.1); border-radius:6px; border:1px solid rgba(250,204,21,0.3);">';
        html += '<div style="font-size:12px; font-weight:600; color:#fbbf24; margin-bottom:6px;">⚠️ Skipped Portfolios (no Good/Bad pattern)</div>';
        html += '<div style="font-size:12px; color:#94a3b8;">' + skippedPortfolios.map(n => esc(n)).join(', ') + '</div>';
        html += '</div>';
    }

    document.getElementById('debugPortfolioDetails').innerHTML = html;
}

function renderDebugChapter3(breakEvenAcos, productPrices, skippedProducts) {
    const productNames = Object.keys(productsMap);
    document.getElementById('debugAcosCount').textContent = `(${productNames.length} products)`;

    let html = '<table class="debug-table"><thead><tr><th>Product</th><th>Break-Even ACOS</th><th>Price</th><th>Price Source</th></tr></thead><tbody>';
    productNames.forEach(name => {
        const acos = breakEvenAcos[name];
        const price = productPrices[name];
        const isSkipped = skippedProducts && skippedProducts.has(name);
        const needsPrice = productsNeedingPrice.includes(name);

        let priceDisplay = 'N/A';
        let sourceDisplay = 'Not needed';
        if (needsPrice) {
            if (isSkipped) {
                priceDisplay = 'Skipped';
                sourceDisplay = '<span style="color:#fbbf24;">User skipped</span>';
            } else if (price > 0) {
                priceDisplay = '$' + price.toFixed(2);
                // Check if auto-calculated
                const priceInput = document.getElementById(`price-${sanitizeId(name)}`);
                const autoTag = priceInput && priceInput.parentElement && priceInput.parentElement.nextElementSibling;
                sourceDisplay = (autoTag && autoTag.textContent.includes('Auto')) ? '<span style="color:#4ade80;">Auto-calculated</span>' : 'Manual entry';
            } else {
                priceDisplay = 'Not entered';
                sourceDisplay = '<span style="color:#f87171;">Missing</span>';
            }
        }

        html += `<tr><td>${esc(name)}</td><td>${acos}%</td><td>${priceDisplay}</td><td>${sourceDisplay}</td></tr>`;
    });
    html += '</tbody></table>';

    document.getElementById('debugAcosDetails').innerHTML = html;
}

function renderDebugChapter4() {
    if (debugDecisions.length === 0) return;

    const movedGood = debugDecisions.filter(d => d.decision === 'MOVE TO GOOD').length;
    const movedBad = debugDecisions.filter(d => d.decision === 'MOVE TO BAD').length;
    const noAction = debugDecisions.filter(d => d.decision === 'NO ACTION').length;
    const unassigned = debugDecisions.filter(d => d.decision === 'UNASSIGNED').length;

    document.getElementById('debugDecisionsCount').textContent = `(${debugDecisions.length} campaigns)`;

    const badgeStyle = (decision) => {
        switch (decision) {
            case 'MOVE TO GOOD': return 'background:rgba(34,197,94,0.2); color:#4ade80;';
            case 'MOVE TO BAD': return 'background:rgba(239,68,68,0.2); color:#f87171;';
            case 'NO ACTION': return 'background:rgba(148,163,184,0.2); color:#94a3b8;';
            case 'UNASSIGNED': return 'background:rgba(250,204,21,0.2); color:#fbbf24;';
            default: return '';
        }
    };

    let html = `<div class="debug-summary-grid" style="margin-bottom:12px;">`;
    html += `<div class="debug-stat"><div class="stat-label">Moved to Good</div><div class="stat-value" style="color:#4ade80;">${movedGood}</div></div>`;
    html += `<div class="debug-stat"><div class="stat-label">Moved to Bad</div><div class="stat-value" style="color:#f87171;">${movedBad}</div></div>`;
    html += `<div class="debug-stat"><div class="stat-label">No Action</div><div class="stat-value">${noAction}</div></div>`;
    html += `<div class="debug-stat"><div class="stat-label">Unassigned</div><div class="stat-value" style="color:#fbbf24;">${unassigned}</div></div>`;
    html += `</div>`;

    html += '<table class="debug-table"><thead><tr>';
    html += '<th>#</th><th>Campaign ID</th><th>Current Portfolio</th><th>ACOS</th><th>Break-Even</th><th>Target</th><th>Decision</th><th>Reason</th>';
    html += '</tr></thead><tbody>';

    debugDecisions.forEach((d, i) => {
        html += `<tr>`;
        html += `<td>${i + 1}</td>`;
        html += `<td style="font-family:monospace; font-size:11px;">${esc(d.campaignId)}</td>`;
        html += `<td>${esc(d.portfolioName)}</td>`;
        html += `<td>${esc(d.acos)}</td>`;
        html += `<td>${esc(d.breakEven)}</td>`;
        html += `<td style="font-size:11px;">${esc(d.targetPortfolio)}</td>`;
        html += `<td><span class="debug-decision-badge" style="${badgeStyle(d.decision)}">${d.decision}</span></td>`;
        html += `<td style="font-size:11px; color:#6b7280;">${esc(d.reason)}</td>`;
        html += `</tr>`;
    });

    html += '</tbody></table>';
    document.getElementById('debugDecisionsDetails').innerHTML = html;
}

function renderDebugChapter5() {
    if (debugDecisions.length === 0) return;

    const total = debugDecisions.length;
    const movedGood = debugDecisions.filter(d => d.decision === 'MOVE TO GOOD').length;
    const movedBad = debugDecisions.filter(d => d.decision === 'MOVE TO BAD').length;
    const noAction = debugDecisions.filter(d => d.decision === 'NO ACTION').length;
    const unassigned = debugDecisions.filter(d => d.decision === 'UNASSIGNED').length;
    const totalMoved = movedGood + movedBad;
    const productsWithAcos = Object.keys(productsMap).length;

    // Group no-action reasons
    const noActionReasons = {};
    debugDecisions.filter(d => d.decision === 'NO ACTION').forEach(d => {
        const key = d.reason;
        noActionReasons[key] = (noActionReasons[key] || 0) + 1;
    });

    document.getElementById('debugResultsCount').textContent = `(${totalMoved} moved)`;

    let html = `<div class="debug-summary-grid">`;
    html += `<div class="debug-stat"><div class="stat-label">Total Campaigns</div><div class="stat-value">${total}</div></div>`;
    html += `<div class="debug-stat"><div class="stat-label">Moved to Good</div><div class="stat-value" style="color:#4ade80;">${movedGood}</div></div>`;
    html += `<div class="debug-stat"><div class="stat-label">Moved to Bad</div><div class="stat-value" style="color:#f87171;">${movedBad}</div></div>`;
    html += `<div class="debug-stat"><div class="stat-label">No Action</div><div class="stat-value">${noAction}</div></div>`;
    html += `<div class="debug-stat"><div class="stat-label">Unassigned</div><div class="stat-value" style="color:#fbbf24;">${unassigned}</div></div>`;
    html += `<div class="debug-stat"><div class="stat-label">Products Configured</div><div class="stat-value">${productsWithAcos}</div></div>`;
    html += `</div>`;

    // No-action breakdown
    if (Object.keys(noActionReasons).length > 0) {
        html += '<div style="margin-top:14px; padding:10px; background:rgba(148,163,184,0.1); border-radius:6px; border:1px solid rgba(148,163,184,0.2);">';
        html += '<div style="font-size:12px; font-weight:600; color:#94a3b8; margin-bottom:8px;">No Action Breakdown</div>';
        for (const [reason, count] of Object.entries(noActionReasons)) {
            html += `<div style="font-size:12px; color:#6b7280; margin-bottom:3px;">&bull; ${esc(reason)}: <strong>${count}</strong></div>`;
        }
        html += '</div>';
    }

    // Download preview
    html += `<div style="margin-top:12px; padding:10px; background:rgba(59,130,246,0.1); border-radius:6px; font-size:13px; color:#60a5fa;">`;
    html += `<strong>Download Preview:</strong> The exported file will contain <strong>${totalMoved}</strong> campaigns with updated portfolio assignments.`;
    html += `</div>`;

    document.getElementById('debugResultsDetails').innerHTML = html;
}
