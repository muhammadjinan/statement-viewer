// PDF.js configuration
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';

// Application State
let allTransactions = new Map();

// DOM Elements
const fileUpload = document.getElementById('fileUpload');
const clearDataBtn = document.getElementById('clearData');
const totalInterestCrEl = document.getElementById('totalInterestCr');
const totalInterestDrEl = document.getElementById('totalInterestDr');
const tableBody = document.getElementById('tableBody');

// Event Listeners
fileUpload.addEventListener('change', handleFileUpload);
clearDataBtn.addEventListener('click', resetApp);

async function handleFileUpload(event) {
    const files = event.target.files;
    if (!files.length) return;

    for (const file of files) {
        const ext = file.name.split('.').pop().toLowerCase();
        
        try {
            if (ext === 'xls' || ext === 'xlsx' || ext === 'csv') {
                await processExcel(file);
            } else if (ext === 'txt') {
                await processText(file);
            } else if (ext === 'pdf') {
                await processPDF(file);
            }
        } catch (error) {
            console.error(`Error processing ${file.name}:`, error);
            alert(`Failed to parse ${file.name}. Ensure it's a valid statement.`);
        }
    }
    
    updateDashboard();
    fileUpload.value = ''; 
}

function processExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            extractTransactions(rows);
            resolve();
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

function processText(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const text = e.target.result;
            const rows = text.split('\n').map(row => row.split(/[\t,]/));
            extractTransactions(rows);
            resolve();
        };
        reader.onerror = reject;
        reader.readAsText(file);
    });
}

async function processPDF(file) {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    let rows = [];

    for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        
        let itemsByY = {};
        textContent.items.forEach(item => {
            const y = Math.round(item.transform[5]);
            if (!itemsByY[y]) itemsByY[y] = [];
            itemsByY[y].push(item);
        });

        const yCoords = Object.keys(itemsByY).map(Number).sort((a, b) => b - a);
        
        yCoords.forEach(y => {
            const rowItems = itemsByY[y].sort((a, b) => a.transform[4] - b.transform[4]);
            rows.push(rowItems.map(item => item.str));
        });
    }
    extractTransactions(rows);
}

function extractTransactions(rows) {
    let isTransactionSection = false;

    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length < 2) continue;

        const cleanRow = row.map(cell => cell ? String(cell).trim() : '');

        if (cleanRow[0].includes('Date') && cleanRow[1].includes('Narration')) {
            isTransactionSection = true;
            continue;
        }

        if (cleanRow.join('').includes('STATEMENT SUMMARY') || cleanRow.join('').includes('Opening Balance')) {
            isTransactionSection = false;
        }

        if (isTransactionSection) {
            // Trim whitespace/hidden characters from the date string before testing
            const dateStr = cleanRow[0].trim();
            const datePattern = /^\d{2}\/\d{2}\/\d{2,4}$/;
            
            if (!datePattern.test(dateStr)) continue;

            const narration = cleanRow[1] || '';
            const withdrawal = parseFloat(cleanRow[4]) || 0;
            const deposit = parseFloat(cleanRow[5]) || 0;
            const balance = cleanRow[6] || '';

            // Expanded regex to catch variations of interest payments
            const isInterest = /interest|int\.?\s*pd|int\.?\s*rec/i.test(narration);

            const uniqueKey = `${dateStr}_${narration.substring(0, 20)}_${withdrawal}_${deposit}`;
            
            if (!allTransactions.has(uniqueKey)) {
                allTransactions.set(uniqueKey, {
                    date: dateStr, narration, withdrawal, deposit, balance, isInterest
                });
            }
        }
    }
}

function updateDashboard() {
    let totalInterestCr = 0;
    let totalInterestDr = 0;
    tableBody.innerHTML = '';

    if (allTransactions.size === 0) {
        tableBody.innerHTML = '<tr><td colspan="5" class="empty-state">No transactions found in the uploaded files.</td></tr>';
        totalInterestCrEl.innerText = '₹ 0.00';
        totalInterestDrEl.innerText = '₹ 0.00';
        return;
    }

    const sortedTxns = Array.from(allTransactions.values()).sort((a, b) => {
        const [d1, m1, y1] = a.date.split('/');
        const [d2, m2, y2] = b.date.split('/');
        // Normalize year for sorting (assuming 20xx)
        const year1 = y1.length === 2 ? `20${y1}` : y1;
        const year2 = y2.length === 2 ? `20${y2}` : y2;
        return new Date(`${year1}-${m1}-${d1}`) - new Date(`${year2}-${m2}-${d2}`);
    });

    sortedTxns.forEach(txn => {
        // Only sum the interest transactions for the dashboard cards
        if (txn.isInterest) {
            totalInterestCr += txn.deposit;
            totalInterestDr += txn.withdrawal;
        }

        const tr = document.createElement('tr');
        
        // Highlight interest rows in a subtle blue for easier scanning
        if (txn.isInterest) {
            tr.style.backgroundColor = '#f0f8ff';
        }

        tr.innerHTML = `
            <td>${txn.date}</td>
            <td>${txn.narration}</td>
            <td style="color: var(--danger);">${txn.withdrawal > 0 ? txn.withdrawal.toFixed(2) : '-'}</td>
            <td style="color: var(--success);">${txn.deposit > 0 ? txn.deposit.toFixed(2) : '-'}</td>
            <td>${txn.balance}</td>
        `;
        tableBody.appendChild(tr);
    });

    totalInterestCrEl.innerText = `₹ ${totalInterestCr.toFixed(2)}`;
    totalInterestDrEl.innerText = `₹ ${totalInterestDr.toFixed(2)}`;
}

function resetApp() {
    allTransactions.clear();
    updateDashboard();
}
