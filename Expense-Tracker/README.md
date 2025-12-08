
# Web-based Expense Tracker (Google Sheets + Apps Script)

A dark-themed **family expense tracker** built with **Google Apps Script + HTML/Bootstrap**, using **Google Sheets as the database**.

### 🔗 Web page link : [Click here](https://script.google.com/macros/s/AKfycbz0okfgRuwrc2pu-GDO-v3iM2_UXHBIvRflSG4Pt5QO6ck34YMFARSEIJ0mBgh6oBa4lw/exec)
You can:

- Select a **family member**
- Add **multiple expenses in one go**
- Save them to a Google Sheet
- Instantly see **transaction history + totals** for that member

---

## 📸 UI Preview

> Add your screenshot to the repo (e.g. `screenshot.png`) and uncomment this line:

<!-- ![Family Expense Tracker Screenshot](./screenshot.png) -->

---

## 🧩 Features

- 🔹 **Member-wise expense entry**
- 🔹 **Multiple rows per submission** (bulk add)
- 🔹 **Dropdowns from Google Sheet** (Members & Categories)
- 🔹 **Instant history panel** for the selected member
- 🔹 **Total transactions and total amount** summary
- 🔹 **Validation & error handling** (incomplete rows, missing member, etc.)
- 🔹 Modern **dark UI** using Bootstrap 5

---

## 🏗 Tech Stack

- **Frontend:**  
  - HTML, CSS, JavaScript  
  - Bootstrap 5  
  - Bootstrap Icons  

- **Backend:**  
  - Google Apps Script (`Code.gs`)  

- **Database:**  
  - Google Sheets (`DATA`, `DROPDOWNS`)

---

## 📊 Google Sheets Structure

Create a Google Sheet with **two tabs**:

### 1. `DROPDOWNS`

Used to populate the Member and Category dropdowns in the UI.

| Column | Header    | Description           |
|--------|-----------|-----------------------|
| A      | EMPLOYEE  | Family member names   |
| B      | CATEGORY  | Expense categories    |

Example:

| EMPLOYEE | CATEGORY               |
|---------|-------------------------|
| Father  | Housing & Utilities     |
| Mother  | Groceries & Household   |
| You     | Transportation          |
| Sibling | Education               |

> The script reads **A2:A** for members and **B2:B** for categories (ignores blanks and trims spaces).

---

### 2. `DATA`

All submitted expenses are stored here.

| Column | Header           | Description                          |
|--------|------------------|--------------------------------------|
| A      | Date Entered     | Timestamp when form was submitted    |
| B      | Employee         | Member name (from dropdown)          |
| C      | Transaction Date | Actual expense date                  |
| D      | Category         | Expense category                     |
| E      | Amount           | Amount spent                         |
| F      | Description      | Optional notes                       |

> **Do not** change the column positions. You can rename headers if you like, but keep the order A–F as above.

---

## ⚙️ How It Works

### UI (`Index.html`)

- Dark-themed Bootstrap form with:
  - **Member dropdown** (`#employeeSelect`)
  - Dynamic **transaction rows** (`date`, `category`, `amount`, `description`)
  - **Add row** and **Delete row** controls
  - **Submit Expenses** button

- Below the form:
  - **Transaction history section** that appears when a member is selected
  - Shows:
    - Date Entered
    - Transaction Date
    - Category
    - Amount
    - Description
  - Sorted by **latest transaction date**
  - Displays:
    - `Total Transactions`
    - `Total Amount (₹)`

- Uses `google.script.run` to talk to Apps Script backend:
  - `getDropdownData()` – loads members & categories
  - `submitExpenses(data)` – writes rows to `DATA`
  - `getEmployeeTransactions(name)` – loads history for a member

---

### Backend (`Code.gs`)

Key parts:

#### Config

```js
const SS = SpreadsheetApp.getActiveSpreadsheet();
const SH_DROPDOWNS = 'DROPDOWNS';
const SH_DATA = 'DATA';
````

#### Web App Entry

```js
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Expense Entry Form')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


#### Load Dropdown Data

```js
function getDropdownData() {
  // Reads DROPDOWNS!A2:A for members and B2:B for categories
  // Trims spaces and ignores empty cells
  ...
}
```

#### Submit Expenses

```js
function submitExpenses(data) {
  // Cleans member name
  // Loops over transactions, builds rows [DateEntered, Member, TxnDate, Category, Amount, Description]
  // Appends them to DATA sheet
  // Returns { success: true/false, message: '...' }
}
```

#### Get Member History

```js
function getEmployeeTransactions(employeeName) {
  // Reads DATA!A2:F
  // Normalizes name with trim + lowercase
  // Filters by member and returns an array of objects:
  // { dateEntered, employee, transactionDate, category, amount, description }
}
```

The front-end then renders this array as the history table and calculates totals.

---

## 🚀 Setup & Deployment

1. **Create Spreadsheet**

   * Create a new Google Sheet.
   * Add `DROPDOWNS` and `DATA` sheets.
   * Set up headers and sample data as shown above.

2. **Open Apps Script**

   * `Extensions` → `Apps Script`
   * Delete the default `Code.gs`.

3. **Add Files**

   * Create `Code.gs` and paste the backend code.
   * Create `Index.html` and paste the HTML/JS/CSS.

4. **Deploy as Web App**

   * Click **Deploy** → **New deployment**
   * Select type: **Web app**
   * Execute as: **Me**
   * Who has access: **Anyone with the link** (or your org)
   * Deploy and copy the URL.

5. **Use It**

   * Open the web app URL.
   * Select a member, add rows, submit.
   * Check `DATA` sheet and the **Transaction History** section.

---

## 💡 Possible Enhancements

* Date range filter in history (From / To)
* Filter by category
* Monthly summary per member
* Charts dashboard (Pie, Bar) using another HTML page / Looker Studio
* Export history to PDF

---
## Glance of Exepense Web Page

<img width="943" height="865" alt="image" src="https://github.com/user-attachments/assets/a334e0d5-2f6b-4ab4-bdd7-c2d710292c0d" />

---
## 👤 Author

Built by **Satanand**

> MIS Executive & automation enthusiast, turning Google Sheets into real-world web apps.

```
```
