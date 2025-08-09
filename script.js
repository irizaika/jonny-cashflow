const excelInput = document.getElementById('excelFile');
const templateInput = document.getElementById('templateFile');
const generateBtn = document.getElementById('generateBtn');
const logElem = document.getElementById('log');

let excelData, templateArrayBuffer;

function updateButton() {
    generateBtn.disabled = !(excelInput.files.length && templateInput.files.length);
    if (!generateBtn.disabled) logElem.textContent = "Ready to generate invoices.";
}
excelInput.addEventListener('change', updateButton);
templateInput.addEventListener('change', updateButton);

function log(msg) {
    logElem.textContent += "\n" + msg;
    logElem.scrollTop = logElem.scrollHeight;
}

function parseExcel(arrayBuffer) {
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });

    const invoices = [];
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
        invoices.push({ name, address, bank, vatrate, issuedate, additional, duedate, workDates, amounts, details });
        i++;
    }
    return invoices;
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
    const invoices = parseExcel(excelData);
    if (invoices.length === 0) { log("❌ No invoices found."); return; }
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

        if (!invoice.duedate) {
            // Case: empty → add 3 months to issue date
            const baseDate = issueDateFormat; // just get it previously, value exist!
          //  formattedDueDate = formatDate(new Date(baseDate.getFullYear(), baseDate.getMonth() + 3, baseDate.getDate()));
            formattedDueDate = formatDate(addMonthsClamped(baseDate, 3));
        }
        else if (typeof invoice.duedate === 'string') {
            // Case: Paid
            formattedDueDate = 'Paid';
        }
        else {
            // Case: assume it's a date
            formattedDueDate = formatDate(new Date(invoice.duedate));
        }
        //const formattedDueDate = (!invoice.duedate || (typeof invoice.duedate === 'string' && invoice.duedate.toLowerCase() === 'paid'))
        //    ? formatDate(new Date(minDate.getFullYear(), minDate.getMonth() + 3, minDate.getDate()))
        //    : invoice.duedate;

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

excelInput.addEventListener('change', e => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = evt => { excelData = evt.target.result; };
    reader.readAsArrayBuffer(file);
});

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