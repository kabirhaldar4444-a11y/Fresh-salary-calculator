# ApexSalary — Professional Salary Calculator & Payroll Compliance

A premium, mathematically verified Enterprise Salary Calculator built directly from and fully compatible with **`Salary Calculater.xlsx`**.

---

## 🌟 Highlights & Key Features

1. **Mathematically Exact (1:1 Parity with `Salary Calculater.xlsx`)**:
   - **Basic Salary**: `=B4*B5/100` (configurable % of CTC, standard 50%).
   - **House Rent Allowance (HRA)**: `=B18*B6/100` (40% Non-Metro / 50% Metro).
   - **Employee Provident Fund (PF)**: `=IF(B9="Yes", IF(B12="Yes", B18*B10/100, MIN(B18, B11*12)*B10/100), 0)`.
   - **Employer Provident Fund (PF)**: Same statutory 12% calculation with support for the ₹15,000 monthly statutory wage ceiling or higher wage contribution.
   - **Gratuity Provision**: `=B18*B13/100` using the statutory 15/26/12 formula (`4.8077%` of Basic).
   - **Special / Flexi Allowance**: Exact balancing formula `=B24-SUM(B18:B19,B21:B23)` in Excel mode and standard corporate liability mode.
   - **Site / Project Completion Bonus**: Row 25 milestone bonus integration (default ₹5,00,000).

2. **Enterprise Payroll Compliance & Control Checks (B27:B31)**:
   - **Total CTC Check**: Real-time zero-deviation audit verifying calculated sum equals CTC (`Difference: ₹0.00`).
   - **Wage / Allowance Check (Code on Wages 50% Rule)**: Real-time gauge and auto-balancer ensuring wages stay above the 50% statutory threshold to prevent compliance penalties.
   - **PF Monthly Wage Check**: Compares monthly basic salary against the ₹15,000 statutory cap.
   - **Gratuity Provision & Accrual**: Annual accrual tracking and 5-year tenure payout projections.

3. **Real-World Take-Home Pay & Tax Calculator**:
   - Computes Gross Salary vs Deductions (Employee PF, Professional Tax ₹200/mo, and TDS).
   - Real-time Income Tax estimation under the **New Tax Regime (FY 2024–26)** with ₹75,000 standard deduction and Section 87A rebate (tax-free up to ₹7,75,000), as well as the Old Tax Regime.

4. **Rich Executive Aesthetics**:
   - Crisp Executive White Luxury theme with elegant indigo, emerald, and slate accents.
   - Interactive SVG Donut breakdown chart & stacked cashflow track.
   - Tabular figures (`JetBrains Mono`) for financial precision.

5. **Export & Print**:
   - **Export to Excel (`.xlsx`)**: Generates spreadsheet with formulas matching `Salary Calculater.xlsx`.
   - **Load Custom Excel**: Upload any `.xlsx` file to automatically populate fields.
   - **Print Official Annexure**: Formatted corporate CTC annexure ready for candidate signing and PDF saving.

---

## 🚀 How to Run Locally

1. Open a terminal in the project directory:
   ```bash
   cd e:\sallary
   node server.js
   ```

2. Open your web browser and visit:
   ```
   http://localhost:3000
   ```

---

## 📁 Project Structure

```
e:\sallary\
├── Salary Calculater.xlsx       # Original reference Excel spreadsheet
├── index.html                   # Semantic HTML5 layout with accessible tabs and modals
├── styles.css                   # Vanilla CSS design system with executive white theme & print styling
├── calculator.js                # Mathematical engine, Excel formula evaluator, SheetJS integration
├── server.js                    # Built-in lightweight Node.js static server
├── libs\
│   └── xlsx.full.min.js         # Offline SheetJS library for in-browser Excel import/export
└── README.md                    # Project documentation
```
