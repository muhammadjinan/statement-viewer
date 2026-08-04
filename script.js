// PDF.js configuration
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';

// Application State
let allTransactions = new Map(); // Using Map for deduplication

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
    // Reset input so the same files can be selected again if needed
    fileUpload.value = ''; 
}

// Process XLS / XLSX (Based on the HDFC format in your screenshots)
function processExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // Convert to array of arrays
            const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            extractTransactions(rows);
            resolve();
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// Process Text files (Assuming tab or comma separated)
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

// Process PDF Files (Extracts text lines and attempts to format as table rows)
async function processPDF(file) {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    let rows = [];

    for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        
        // Group text items by their vertical position (Y coordinate) to form rows
        let itemsByY = {};
        textContent.items.forEach(item => {
            const y = Math.round(item.transform[5]);
            if (!itemsByY[y]) itemsByY[y] = [];
            itemsByY[y].push(item);
        });

        // Sort Y coordinates descending (top to bottom)
        const yCoords = Object.keys(itemsByY).map(Number).sort((a, b) => b - a);
        
        yCoords.forEach(y => {
            // Sort items in a row by X coordinate (left to right)
            const rowItems = itemsByY[y].sort((a, b) => a.transform[4] - b.transform[4]);
            rows.push(rowItems.map(item => item.str));
        });
    }
    extractTransactions(rows);
}

// Core Logic: Extracting, Filtering, and Deduplicating
function extractTransactions(rows) {
    let isTransactionSection = false;

    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length < 2) continue;

        // Clean string values
        const cleanRow = row.map(cell => cell ? String(cell).trim() : '');

        // Detect table header based on your screenshots (1000737339.jpg)
        if (cleanRow[0].includes('Date') && cleanRow[1].includes('Narration')) {
            isTransactionSection = true;
            continue;
        }

        // Stop processing if we hit the statement summary (1000737375.jpg)
        if (cleanRow.join('').includes('STATEMENT SUMMARY') || cleanRow.join('').includes('Opening Balance')) {
            isTransactionSection = false;
        }

        if (isTransactionSection) {
            // Check for valid date format (e.g., DD/MM/YY) to skip masking rows like "********"
            const datePattern = /^\d{2}\/\d{2}\/\d{2,4}$/;
            if (!datePattern.test(cleanRow[0])) continue;

            const date = cleanRow[0];
            const narration = cleanRow[1] || '';
            const withdrawal = parseFloat(cleanRow[4]) || 0;
            const deposit = parseFloat(cleanRow[5]) || 0;
            const balance = cleanRow[6] || '';

            // Filter for Interest (e.g., 'INT.PD', 'INTEREST', 'CAPITALIZED')
            const isInterest = /interest|int\.?pd|int\.?rec/i.test(narration);

            if (isInterest) {
                // Deduplication Key: Date + First 20 chars of Narration + Withdrawal + Deposit
                // This ensures if overlapping date ranges are uploaded, duplicates are ignored.
                const uniqueKey = `${date}_${narration.substring(0, 20)}_${withdrawal}_${deposit}`;
                
                if (!allTransactions.has(uniqueKey)) {
                    allTransactions.set(uniqueKey, {
                        date, narration, withdrawal, deposit, balance
                    });
                }
            }
        }
    }
}

// Update DOM with calculations
function updateDashboard() {
    let totalCr = 0;
    let totalDr = 0;
    tableBody.innerHTML = '';

    if (allTransactions.size === 0) {
        tableBody.innerHTML = '<tr><td colspan="5" class="empty-state">No interest transactions found in the uploaded files.</td></tr>';
        totalInterestCrEl.innerText = '₹ 0.00';
        totalInterestDrEl.innerText = '₹ 0.00';
        return;
    }

    // Sort by Date (assuming DD/MM/YY)
    const sortedTxns = Array.from(allTransactions.values()).sort((a, b) => {
        const [d1, m1, y1] = a.date.split('/');
        const [d2, m2, y2] = b.date.split('/');
        return new Date(`20${y1}-${m1}-${d1}`) - new Date(`20${y2}-${m2}-${d2}`);
    });

    sortedTxns.forEach(txn => {
        totalCr += txn.deposit;
        totalDr += txn.withdrawal;

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${txn.date}</td>
            <td>${txn.narration}</td>
            <td style="color: var(--danger);">${txn.withdrawal > 0 ? txn.withdrawal.toFixed(2) : '-'}</td>
            <td style="color: var(--success);">${txn.deposit > 0 ? txn.deposit.toFixed(2) : '-'}</td>
            <td>${txn.balance}</td>
        `;
        tableBody.appendChild(tr);
    });

    totalInterestCrEl.innerText = `₹ ${totalCr.toFixed(2)}`;
    totalInterestDrEl.innerText = `₹ ${totalDr.toFixed(2)}`;
}

function resetApp() {
    allTransactions.clear();
    updateDashboard();
}

