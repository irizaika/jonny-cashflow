# Jonny CashFlow

Jonny CashFlow is your friendly invoice generator designed to simplify billing and speed up your cash flow.  
Load your invoice data from Excel, pick your custom Word template, and generate professional invoices in seconds — all in one easy-to-use web app.

## Features

- Simple drag & drop Excel and DOCX template inputs  
- Automated invoice generation with dynamic date formatting  
- Calculates totals and formats your data professionally  
- Clean, user-friendly interface for hassle-free operation  

## Getting Started

1. Upload your Excel invoice data file.  
2. Upload your DOCX invoice template.  
3. Click **Generate Invoices** and watch the magic happen!  
4. Download your personalized invoices and send them out.  
5. **Important:** Ensure dates in your Excel file use the `dd/mm/yyyy` format to avoid parsing errors.
6. Placeholders in the DOCX template **can be removed but not added**, and please note that they are **case-sensitive**.

Keep your payments flowing with **Jonny CashFlow** — because your business deserves smooth sailing.

---

### Tech Stack

- JavaScript (vanilla)  
- XLSX library for Excel parsing  
- Docxtemplater and PizZip for Word document templating  
- FileSaver for saving invoices on the client-side

---

### Excel Data Format and Placeholder Mapping

The invoice data is loaded from an Excel file formatted like this:

- **Invoice header row:**  
  - **Name** — name of the client  
  - **Address** — comma-separated for multiple lines (e.g. `123 Main St,City,ZIP`), comma will be replaced with new line
  - **Bank** — (Optional) payment details, also comma-separated if multiple lines for selfbilling invoice, comma will be replaced with new line
  - **Vat rate** — (Optional) number or left empty, if left empty default 20% will be used, add 0 if should not be calculated
  - **Due date** — (Optional) date string, “Paid”, or left empty (if left empty set to +3 month from issue date) 
  - **Issue date** — (Optional) date string, or left empty (if left empty set to today's date) 
  - **Additional text** — (Optional) optional notes, not in use

- **Invoice item rows:**  
  Each subsequent row contains:  
  - **Work date** — date of the service or product  
  - **Amount** — numeric value  
  - **Details** — description of the item or service  

- **Blank rows** separate different invoices.

---

#### Example snippet from Excel (order matters):

| Name        | Address                              | Bank Details                                 | Vat rate            | Due Date |Issue date  |Additional Text |
|-------------|--------------------------------------|----------------------------------------------|---------------------|----------|------------|----------------|
| Irina Z     | 123 Example St,Example City,EX 12345 | Nationwide,Acc no 4444444,Sort code 07 08 16 | 0                   | Paid     |            |                |
| 01/01/2025  | 100                                  | oak                                          |                     |          |            |                |
| 01/07/2025  | 200.5                                | dfs                                          |                     |          |            |                |
| 01/10/2025  | 50.75                                | H                                            |                     |          |            |                |
|             |                                      |                                              |                     |          |            |                |
| John Smith  | 123 Example St,Example City,EX 12345 | Nationwide,Acc no 442224444                  |                     |          | 20/08/2025 |                |
| 01/01/2025  | 140                                  | Web design                                   |                     |          |            |                |

---

### Placeholder Reference

| Placeholder             | Source from Excel                                         | Example                                     |
|-------------------------|-----------------------------------------------------------|---------------------------------------------|
| `{invoiceid}`           | Generated ID from contractor name and date                | JONNY11082024                               |
| `{issuedate}`           | Smallest work date in invoice items                       | 03/08/2025                                  |
| `{name}`                | Contractor name (header row)                              | Irina Z                                     |
| `{address}`             | Address (header row, commas replaced by line breaks)      | 123 Example St,Example City,EX 12345        |
| `{bank}`                | Bank details (header row, commas replaced by line breaks) | Nationwide,Acc no 171781,Sort code 27 02 06 |
| `{additionaltext}`      | Additional text (header row)                              | Not in use                                  |
| `{duedate}`             | Due date (header row)                                     | Paid                                        |
| `{mmYYYY}`              | Month and year from smallest work date in invoice items   | Paid                                        |
| `{#items}`...`{/items}` | Loop over invoice item rows                               | Work date, details, amount per row          |

---

### Invoice Item Placeholders (inside `{#items}` block)

- `{workdate}` — Date of the service  
- `{details}` — Description of the service/item  
- `{amount}` — Amount charged  

---

### Calculated Fields

- `{subtotal}`, `{total}` — Sum of all amounts  
- `{vat}` - 
- `{tax}` — Currently always 0.00

---

## Download Templates

To get started quickly, you can download sample templates here:

- [Invoice DOCX Template ](templates/invoice-template.docx)
- [Selfbilling Invoice DOCX Template](templates/selfbilling-invoice-template.docx)  
- [Invoice Data Excel Template](templates/invoice-data-template.xlsx)

---

## Version  
Version 1.0 — more features coming soon!

## License  
MIT License — free for personal and commercial use.

Happy invoicing!  
— Jonny CashFlow Team
