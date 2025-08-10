let excelInput, templateInput, generateBtn, logElem;
let excelData, templateArrayBuffer;
let invoices = [];

function updateButtonGeneratePage() {
    generateBtn.disabled = !(excelInput.files.length && templateInput.files.length);
    if (!generateBtn.disabled) logElem.textContent = "Ready to generate invoices.";
}

function updateButtonManualPage() {
    //todo more validation here
    generateBtn.disabled = !(templateInput.files.length);
    if (!generateBtn.disabled) logElem.textContent = "Ready to generate invoices.";
}

function log(msg) {
    logElem.textContent += "\n" + msg;
    logElem.scrollTop = logElem.scrollHeight;
}

function parseExcel(arrayBuffer) {
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });

    let i = 1; // first row just header
    while (i < rows.length) {
        if (!rows[i] || rows[i].length === 0 || !rows[i][0]) { i++; continue; }
        const [name, address, bank, vatrate, duedateRaw, issuedateRaw, additional] = rows[i]; //order is important - should be like in excel file
        i++;

        const workDates = [], amounts = [], details = [];
        var issuedate;
        var duedate;

        while (i < rows.length && rows[i] && rows[i][0]) {
            const [workDateRaw, amountRaw, detail] = rows[i];
            let workDate;
            try {
                if (workDateRaw instanceof Date) workDate = workDateRaw;
                else if (typeof workDateRaw === 'string') workDate = new Date(workDateRaw);
                else if (typeof workDateRaw === 'number') {
                    const dateParts = XLSX.SSF.parse_date_code(workDateRaw);
                    workDate = new Date(Date.UTC(dateParts.y, dateParts.m - 1, dateParts.d));
                } else throw new Error();
            } catch {
                log(`⚠️ Invalid date "${workDateRaw}" at row ${i + 1}, skipping.`);
                i++; continue;
            }
            const amount = Number(amountRaw);
            if (isNaN(amount)) {
                log(`⚠️ Invalid amount "${amountRaw}" at row ${i + 1}, skipping.`);
                i++; continue;
            }

            issuedate = parseExcelDate(issuedateRaw);
            duedate = parseExcelDate(duedateRaw);

            workDates.push(workDate);
            amounts.push(amount);
            details.push(detail || '');
            i++;
        }
        //update global variable
        invoices.push({ name, address, bank, vatrate, issuedate, additional, duedate, workDates, amounts, details });
        i++;
    }
}

function formatDate(date) {
    return date.toLocaleDateString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' });
}

function generateInvoiceID(name, date) {
    let id = name.replace(/\s/g, '').toUpperCase().substring(0, 5);
    const ddmmyyyy = date.toLocaleDateString('en-GB').replace(/\//g, '');
    return id + ddmmyyyy;
}

function parseExcelDate(dateRaw) {
    if (dateRaw instanceof Date) return dateRaw;
    if (typeof dateRaw === 'string') {
        if (dateRaw.toLowerCase() === 'paid') {
            return dateRaw;
        }
        return new Date(dateRaw);
    }
    if (typeof dateRaw === 'number') {
        const dateParts = XLSX.SSF.parse_date_code(dateRaw);
        return new Date(Date.UTC(dateParts.y, dateParts.m - 1, dateParts.d));
    }
    return null;
}

function addMonthsClamped(date, months) {
    const year = date.getFullYear();
    const month = date.getMonth() + months;
    const targetMonth = (month % 12 + 12) % 12;
    const targetYear = year + Math.floor(month / 12);
    const lastDay = new Date(targetYear, targetMonth, 0).getDate();
    const day = Math.min(date.getDate(), lastDay);
    return new Date(targetYear, targetMonth, day);
}

async function generateInvoices() {
    logElem.textContent = "⏳ Starting invoice generation...";

    if (invoices.length === 0) {// if invoices emoty, just in case try read excel file again
        parseExcel(excelData); 
        if (invoices.length === 0) { log("❌ No invoices found."); return; }
    }
    log(`✅ Found ${invoices.length} invoices.`);

     for (const invoice of invoices) {
        const docZip = new PizZip(templateArrayBuffer);
        const doc = new window.docxtemplater(docZip, { paragraphLoop: true, linebreaks: true });

        const total = invoice.amounts.reduce((a, b) => a + b, 0);
        const vatRate = (!invoice.vatrate) ? 0.20 : invoice.vatrate / 100; // e.g., 20 → 0.20
        const vatRateStr = (vatRate * 100).toFixed(0) + '%';
        const vatAmount = total - (total / (1 + vatRate));
        const workDatesStr = invoice.workDates.map(d => formatDate(d));
        const minDate = new Date(Math.min(...invoice.workDates));

        const today = new Date();
        const issueDateFormat = invoice.issuedate == null ? today : invoice.issuedate; //minDate,
        const issueDateStr = formatDate(issueDateFormat);

        if (!invoice.duedate) { // Case: empty → add 3 months to issue date
            const baseDate = issueDateFormat;
            formattedDueDate = formatDate(addMonthsClamped(baseDate, 3));
        }
        else if (typeof invoice.duedate === 'string') {// Case: Paid
            formattedDueDate = 'Paid';
        }
        else {// Case: assume it's a date
            formattedDueDate = formatDate(new Date(invoice.duedate));
        }

        const data = {
            invoiceid: generateInvoiceID(invoice.name, minDate),
            mmYYYY: minDate.toLocaleDateString('en-GB', { month: 'long', year: 'numeric' }),
            address: (invoice.address || '').replace(/,/g, '\n'),
            additionaltext: invoice.additional ? invoice.additional.replace(/,/g, '\n') : "",
            duedate: formattedDueDate,
            bank: (invoice.bank || '').replace(/,/g, '\n'),
            name: invoice.name,
            issuedate: issueDateStr, 
            total: total.toFixed(2),
            subtotal: (total - vatAmount).toFixed(2),
            vat: vatAmount.toFixed(2),
            tax: (0).toFixed(2),
            vatrate: vatRateStr,
            items: invoice.amounts.map((amt, idx) => ({
                workdate: workDatesStr[idx],
                details: invoice.details[idx],
                amount: amt.toFixed(2),
            })),
        };

        doc.setData(data);
        try { doc.render(); }
        catch (error) {
            log(`❌ Template error: ${error.message}`);
            continue;
        }

        const outBlob = doc.getZip().generate({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
        saveAs(outBlob, `Invoice_${data.invoiceid}.docx`);
        log(`📄 Generated invoice for ${invoice.name} (${data.invoiceid})`);
    }
    log("🎉 All invoices generated.");
}

function displayInvoiceSummary(invoices) {
    const container = document.getElementById('invoicelog');
    container.textContent = ''; // clear previous content

    invoices.forEach((inv, idx) => {
        const total = inv.amounts.reduce((a, b) => a + b, 0);
        const vatRate = inv.vatrate ? inv.vatrate.toString() + '%' : 'Not set (default 20% will be used)';
        const issuedate = inv.issuedate ? formatDate(inv.issuedate) : 'Not set (today\'s date will be used)';
        const duedate = inv.duedate ? (typeof inv.duedate === 'string' ? inv.duedate : formatDate(inv.duedate)) : 'Not set (issue date + 3 month will be used)';

        // Create list of amounts with dates:
        const itemLines = inv.amounts.map((amt, i) => {
            const dateStr = formatDate(inv.workDates[i]);
            return `${i + 1}. ${dateStr} (£${amt.toFixed(2)})`;
        }).join('\n');

        const summary = [
            `Invoice ${idx + 1}: ${inv.name}`,
            `Items:\n${itemLines}`,  // Detailed amounts and dates here
            `Total Amount: £${total.toFixed(2)}`,
            `VAT Rate: ${vatRate}`,
            `Issue Date: ${issuedate}`,
            `Due Date: ${duedate}`,
            '------------------------'
        ].join('\n');

        container.textContent += summary + '\n';
    });
}

function initGenerateInvoicePage() {
    excelInput = document.getElementById('excelFile');
    templateInput = document.getElementById('templateFile');
    generateBtn = document.getElementById('generateBtn');
    logElem = document.getElementById('log');

    invoices = [];  // global variable for invoices

    excelInput.addEventListener('change', updateButtonGeneratePage);
    templateInput.addEventListener('change', updateButtonGeneratePage);


    templateInput.addEventListener('change', e => {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = evt => { templateArrayBuffer = evt.target.result; };
        reader.readAsArrayBuffer(file);
    });

    generateBtn.addEventListener('click', () => {
        if (!excelData || !templateArrayBuffer) {
            alert('Please select both Excel and template files.');
            return;
        }
        generateInvoices();
    });

    excelInput.addEventListener('change', e => {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = evt => {
            excelData = evt.target.result;
            parseExcel(excelData);
            displayInvoiceSummary(invoices);
            updateButtonGeneratePage();
        };
        reader.readAsArrayBuffer(file);
    });
}

function initManualInputPage() {
    generateBtn = document.getElementById('generateBtn');
    templateInput = document.getElementById('templateFile');
    logElem = document.getElementById('log');

    invoices = [];  // global variable for invoices

  //  templateInput.addEventListener('change', updateButtonManualPage);


    templateInput.addEventListener('change', e => {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = evt => { templateArrayBuffer = evt.target.result; };
        reader.readAsArrayBuffer(file);
        logElem.textContent = "Ready to generate invoices.";
        updateButtonManualPage()
    });

    document.getElementById('addItem').addEventListener('click', () => {
        const tbody = document.querySelector('#itemTable tbody');
        const row = document.createElement('tr');
        row.innerHTML = `
        <td><input type="date" name="workdate" class="form-control form-control-compact"></td>
        <td><input type="text" name="details" class="form-control form-control-compact"></td>
        <td><input type="number" name="amount" step="1.00" class="form-control form-control-compact"></td>
        <td><button type="button" class="removeItem btn btn-sm">❌</button></td>
  `;
        tbody.appendChild(row);
    });

    document.getElementById('itemTable').addEventListener('click', e => {
        if (e.target.classList.contains('removeItem')) {
            e.target.closest('tr').remove();
        }
    });

    document.getElementById('manualInvoiceForm').addEventListener('submit', e => {
        e.preventDefault();
        const form = e.target;
        const formData = new FormData(form);

        const items = Array.from(document.querySelectorAll('#itemTable tbody tr')).map(tr => {
            const dateVal = tr.querySelector('input[name="workdate"]').value;
            const dateObj = new Date(dateVal);
           // const formatted = dateObj.toLocaleDateString('en-GB'); // UK style DD/MM/YYYY

            return {
                workdate: dateObj,
                details: tr.querySelector('input[name="details"]').value,
                amount: parseFloat(tr.querySelector('input[name="amount"]').value)
            };
        });

        const dueDateVal = formData.get('duedate');
        const dueDateObj = dueDateVal ? new Date(dueDateVal) : null;

        const issueDateVal = formData.get('issuedate');
        const issueDateObj = issueDateVal ? new Date(issueDateVal) : null;

        const invoice = {
            name: formData.get('name'),
            address: formData.get('address'),
            bank: formData.get('bank'),
            vatrate: parseFloat(formData.get('vatrate')) || 20,
            issuedate: issueDateObj,
            duedate: dueDateObj,
            additional: '', // No additional text in manual input
            workDates: items.map(i => i.workdate),
            amounts: items.map(i => i.amount),
            details: items.map(i => i.details)
        };

        invoices.push(invoice);

        generateInvoices();
    });
}


let mainContent = document.getElementById("mainContent");

// Function to load a page into mainContent
function loadPage(page) {
    fetch(page)
        .then(res => {
            if (!res.ok) throw new Error(`Failed to load ${page}`);
            return res.text();
        })
        .then(html => {
            mainContent.innerHTML = html;
            if (page === 'generate.html') {
                initGenerateInvoicePage();
            } else if (page === 'manual.html') {
                initManualInputPage();
            } 
        })
        .catch(err => {
            mainContent.innerHTML = `<p style="color:red;">Error: ${err.message}</p>`;
        });
}

document.addEventListener("click", e => {
    const link = e.target.closest("[data-page]");
    if (!link) return; // click wasn't on a data-page element
    e.preventDefault();
    const page = link.getAttribute("data-page");
    if (page) {
        loadPage(page);
    }
});

loadPage("home.html");
