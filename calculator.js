/**
 * ApexSalary — Enterprise Payroll Engine & Mathematical Calculator
 * 1:1 Parity with 'Salary Calculater.xlsx' + Indian Payroll & Statutory Code on Wages Compliance
 */

(() => {
  'use strict';

  // --- STATE ---
  const state = {
    ctc: 1200000,
    isMonthlyInput: false,
    basicPercent: 50,
    hraPercent: 40,
    siteAllowance: 0,
    otherBenefits: 0,
    projectBonus: 500000,
    pfApplicable: true,
    pfRate: 12,
    pfWageCeiling: 15000,
    pfHigherWage: false,
    gratuityRate: 4.807692308, // 15/26/12 = 4.807692308%
    ptApplicable: true,
    taxRegime: 'new',
    mathMode: 'standard', // 'excel' or 'standard'
    currency: 'INR',      // 'INR' or 'USD'
    activeInspectorRow: 'B18',
    offerA: 1000000,
    offerB: 1500000
  };

  // --- DOM ELEMENTS ---
  const elements = {
    // Header & Toggles
    mathModeSelect: document.getElementById('mathModeSelect'),
    mathInfoBtn: document.getElementById('mathInfoBtn'),
    mathInfoModal: document.getElementById('mathInfoModal'),
    mathInfoCloseBtn: document.getElementById('mathInfoCloseBtn'),
    currInrBtn: document.getElementById('currInrBtn'),
    currUsdBtn: document.getElementById('currUsdBtn'),
    currSymbolPrefix: document.getElementById('currSymbolPrefix'),
    exportExcelBtn: document.getElementById('exportExcelBtn'),
    printAnnexureBtn: document.getElementById('printAnnexureBtn'),
    resetDefaultsBtn: document.getElementById('resetDefaultsBtn'),
    excelFileInput: document.getElementById('excelFileInput'),
    importExcelBtn: document.getElementById('importExcelBtn'),

    // Presets
    presetChips: document.querySelectorAll('.preset-chip'),

    // CTC Hero Inputs
    ctcInput: document.getElementById('ctcInput'),
    ctcSlider: document.getElementById('ctcSlider'),
    ctcAnnualToggle: document.getElementById('ctcAnnualToggle'),
    ctcMonthlyToggle: document.getElementById('ctcMonthlyToggle'),
    ctcInWords: document.getElementById('ctcInWords'),

    // Parameters
    basicPercentSlider: document.getElementById('basicPercentSlider'),
    basicPercentInput: document.getElementById('basicPercentInput'),
    hraPercentSlider: document.getElementById('hraPercentSlider'),
    hraPercentInput: document.getElementById('hraPercentInput'),
    hraMetroBtn: document.getElementById('hraMetroBtn'),
    hraNonMetroBtn: document.getElementById('hraNonMetroBtn'),
    siteAllowanceInput: document.getElementById('siteAllowanceInput'),
    otherBenefitsInput: document.getElementById('otherBenefitsInput'),
    projectBonusInput: document.getElementById('projectBonusInput'),

    // Statutory Controls
    statutoryAccordion: document.getElementById('statutoryAccordion'),
    statutoryToggleBtn: document.getElementById('statutoryToggleBtn'),
    pfApplicableCheck: document.getElementById('pfApplicableCheck'),
    pfRateInput: document.getElementById('pfRateInput'),
    pfCeilingInput: document.getElementById('pfCeilingInput'),
    pfHigherWageCheck: document.getElementById('pfHigherWageCheck'),
    gratuityRateInput: document.getElementById('gratuityRateInput'),
    ptApplicableCheck: document.getElementById('ptApplicableCheck'),
    taxRegimeSelect: document.getElementById('taxRegimeSelect'),

    // KPIs
    kpiNetMonthly: document.getElementById('kpiNetMonthly'),
    kpiNetAnnual: document.getElementById('kpiNetAnnual'),
    kpiGrossMonthly: document.getElementById('kpiGrossMonthly'),
    kpiGrossAnnual: document.getElementById('kpiGrossAnnual'),
    kpiCtcAnnual: document.getElementById('kpiCtcAnnual'),
    kpiCtcMonthly: document.getElementById('kpiCtcMonthly'),
    kpiRetiralsAnnual: document.getElementById('kpiRetiralsAnnual'),
    kpiRetiralsMonthly: document.getElementById('kpiRetiralsMonthly'),

    // Compliance Banner
    complianceBanner: document.getElementById('complianceBanner'),
    complianceIndicator: document.getElementById('complianceIndicator'),
    complianceTitle: document.getElementById('complianceTitle'),
    complianceDetail: document.getElementById('complianceDetail'),
    autoFixWageCodeBtn: document.getElementById('autoFixWageCodeBtn'),

    // Tabs
    tabButtons: document.querySelectorAll('.tab-btn'),
    tabPanels: document.querySelectorAll('.tab-panel'),

    // Tab 1: Breakdown Table & Inspector
    copySummaryBtn: document.getElementById('copySummaryBtn'),
    copyBtnText: document.getElementById('copyBtnText'),
    salaryTableBody: document.getElementById('salaryTableBody'),
    salaryTableFoot: document.getElementById('salaryTableFoot'),
    inspectorCellRef: document.getElementById('inspectorCellRef'),
    inspectorFormula: document.getElementById('inspectorFormula'),
    inspectorExplanation: document.getElementById('inspectorExplanation'),
    bonusDisplayAmount: document.getElementById('bonusDisplayAmount'),
    totalEarningPotential: document.getElementById('totalEarningPotential'),

    // Tab 2: Take-Home & Payslip
    eqGrossMonthly: document.getElementById('eqGrossMonthly'),
    eqDeductMonthly: document.getElementById('eqDeductMonthly'),
    eqNetMonthly: document.getElementById('eqNetMonthly'),
    activeRegimeBadge: document.getElementById('activeRegimeBadge'),
    thBasic: document.getElementById('thBasic'),
    thHra: document.getElementById('thHra'),
    thSpecial: document.getElementById('thSpecial'),
    thSite: document.getElementById('thSite'),
    thOther: document.getElementById('thOther'),
    thSiteRow: document.getElementById('thSiteRow'),
    thOtherRow: document.getElementById('thOtherRow'),
    thGrossTotal: document.getElementById('thGrossTotal'),
    thEmpPf: document.getElementById('thEmpPf'),
    thPt: document.getElementById('thPt'),
    thTax: document.getElementById('thTax'),
    thDeductTotal: document.getElementById('thDeductTotal'),
    taxNoteText: document.getElementById('taxNoteText'),

    // Tab 3: Statutory & Control Checks
    checkCtcCard: document.getElementById('checkCtcCard'),
    checkCtcStatus: document.getElementById('checkCtcStatus'),
    checkCtcVal: document.getElementById('checkCtcVal'),
    checkWageCard: document.getElementById('checkWageCard'),
    checkWageStatus: document.getElementById('checkWageStatus'),
    checkWageVal: document.getElementById('checkWageVal'),
    wageMeterBar: document.getElementById('wageMeterBar'),
    checkWageDesc: document.getElementById('checkWageDesc'),
    checkPfStatus: document.getElementById('checkPfStatus'),
    checkPfWageVal: document.getElementById('checkPfWageVal'),
    checkPfDesc: document.getElementById('checkPfDesc'),
    checkGratuityVal: document.getElementById('checkGratuityVal'),
    gratuity5yrPayout: document.getElementById('gratuity5yrPayout'),

    // Tab 4: Visuals
    donutSvg: document.getElementById('donutSvg'),
    donutCenterVal: document.getElementById('donutCenterVal'),
    chartLegend: document.getElementById('chartLegend'),
    flowCtcMonthly: document.getElementById('flowCtcMonthly'),
    flowNetMonthly: document.getElementById('flowNetMonthly'),
    segNet: document.getElementById('segNet'),
    segTax: document.getElementById('segTax'),
    segPf: document.getElementById('segPf'),
    segGratuity: document.getElementById('segGratuity'),

    // Tab 5: Comparison
    offerAInput: document.getElementById('offerAInput'),
    offerBInput: document.getElementById('offerBInput'),
    comparisonTableBody: document.getElementById('comparisonTableBody'),

    // Annexure Modal
    annexureModal: document.getElementById('annexureModal'),
    modalCloseBtn: document.getElementById('modalCloseBtn'),
    modalPrintBtn: document.getElementById('modalPrintBtn'),
    annexureDate: document.getElementById('annexureDate'),
    annexureRef: document.getElementById('annexureRef'),
    annexureCtcVal: document.getElementById('annexureCtcVal'),
    annexureTableBody: document.getElementById('annexureTableBody'),
    annexureTableFoot: document.getElementById('annexureTableFoot')
  };

  // --- NUMBER & CURRENCY FORMATTERS ---
  function formatCurrency(amount, includeSymbol = true, decimalPlaces = 2) {
    if (isNaN(amount) || amount === null) amount = 0;
    const isNegative = amount < 0;
    const absVal = Math.abs(amount);

    let formatted = '';
    if (state.currency === 'INR') {
      // Indian numbering system (Lakhs & Crores)
      formatted = absVal.toLocaleString('en-IN', {
        minimumFractionDigits: decimalPlaces,
        maximumFractionDigits: decimalPlaces
      });
      if (includeSymbol) formatted = '₹ ' + formatted;
    } else {
      // International USD system
      formatted = absVal.toLocaleString('en-US', {
        minimumFractionDigits: decimalPlaces,
        maximumFractionDigits: decimalPlaces
      });
      if (includeSymbol) formatted = '$ ' + formatted;
    }

    return isNegative ? `-${formatted}` : formatted;
  }

  function formatPercent(pct, decimalPlaces = 2) {
    if (isNaN(pct) || pct === null) pct = 0;
    return (pct * 100).toFixed(decimalPlaces) + '%';
  }

  function formatNumber(amount, decimalPlaces = 2) {
    if (isNaN(amount) || amount === null) amount = 0;
    const absVal = Math.abs(amount);
    const formatted = absVal.toLocaleString('en-IN', {
      minimumFractionDigits: decimalPlaces,
      maximumFractionDigits: decimalPlaces
    });
    return amount < 0 ? `-${formatted}` : formatted;
  }

  // Convert Indian Numbers to Words
  function numberToWordsINR(num) {
    if (num === 0) return 'Zero Rupees Only';
    if (num < 0) return 'Minus ' + numberToWordsINR(Math.abs(num));
    
    num = Math.round(num);
    const a = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten',
      'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
    const b = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];

    function inWords(n) {
      let str = '';
      if (n > 19) {
        str += b[Math.floor(n / 10)];
        if (n % 10) str += ' ' + a[n % 10];
      } else {
        str += a[n];
      }
      return str;
    }

    let crore = Math.floor(num / 10000000);
    num %= 10000000;
    let lakh = Math.floor(num / 100000);
    num %= 100000;
    let thousand = Math.floor(num / 1000);
    num %= 1000;
    let hundred = Math.floor(num / 100);
    let remaining = num % 100;

    let res = '';
    if (crore > 0) res += inWords(crore) + ' Crore ';
    if (lakh > 0) res += inWords(lakh) + ' Lakh ';
    if (thousand > 0) res += inWords(thousand) + ' Thousand ';
    if (hundred > 0) res += inWords(hundred) + ' Hundred ';
    if (remaining > 0) {
      if (res !== '') res += 'and ';
      res += inWords(remaining) + ' ';
    }
    return (res.trim() + ' Rupees Only');
  }

  // --- CORE MATHEMATICAL ENGINE ---
  function calculateSalaryModel(params) {
    const ctc = Number(params.ctc) || 0;
    const basicPercent = Number(params.basicPercent) || 50;
    const hraPercent = Number(params.hraPercent) || 40;
    const siteAllowance = Number(params.siteAllowance) || 0;
    const otherBenefits = Number(params.otherBenefits) || 0;
    const pfApplicable = !!params.pfApplicable;
    const pfRate = Number(params.pfRate) || 12;
    const pfWageCeiling = Number(params.pfWageCeiling) || 15000;
    const pfHigherWage = !!params.pfHigherWage;
    const gratuityRate = Number(params.gratuityRate) || 4.807692308;
    const ptApplicable = !!params.ptApplicable;
    const taxRegime = params.taxRegime || 'new';
    const mathMode = params.mathMode || 'standard';

    // 1. Basic Salary (Row 18)
    // Excel formula: =B4*B5/100
    const basicAnnual = ctc * (basicPercent / 100);
    const basicMonthly = basicAnnual / 12;

    // 2. House Rent Allowance (HRA) (Row 19)
    // Excel formula: =B18*B6/100
    const hraAnnual = basicAnnual * (hraPercent / 100);
    const hraMonthly = hraAnnual / 12;

    // 3. Provident Fund Calculations (Rows 21 & 22)
    // Statutory ceiling: 15,000 / month = 180,000 / year
    // Excel formula: =IF(B9="Yes",IF(B12="Yes",B18*B10/100,MIN(B18,B11*12)*B10/100),0)
    let pfEligibleAnnual = 0;
    if (pfApplicable) {
      if (pfHigherWage) {
        pfEligibleAnnual = basicAnnual;
      } else {
        pfEligibleAnnual = Math.min(basicAnnual, pfWageCeiling * 12);
      }
    }
    const empPfAnnual = pfApplicable ? pfEligibleAnnual * (pfRate / 100) : 0;
    const empPfMonthly = empPfAnnual / 12;

    const emprPfAnnual = pfApplicable ? pfEligibleAnnual * (pfRate / 100) : 0;
    const emprPfMonthly = emprPfAnnual / 12;

    // 4. Gratuity Provision (Row 23)
    // Excel formula: =B18*B13/100
    const gratuityAnnual = basicAnnual * (gratuityRate / 100);
    const gratuityMonthly = gratuityAnnual / 12;

    // 5. Special / Flexi Allowance (Row 20)
    // Balancing allowance to make calculated sum match CTC
    let specialAnnual = 0;
    if (mathMode === 'excel') {
      // Direct Excel Sheet2 Row 20 Formula:
      // =B24-SUM(B18:B19,B21:B23)
      // Note: Row 21 (Emp PF) + Row 22 (Empr PF) + Row 23 (Gratuity)
      specialAnnual = ctc - (basicAnnual + hraAnnual + siteAllowance + otherBenefits + empPfAnnual + emprPfAnnual + gratuityAnnual);
    } else {
      // Standard Corporate CTC Formula:
      // CTC = Gross (Basic + HRA + Special + Site + Other) + Employer PF + Gratuity
      // Employee PF is a deduction from Gross, not part of employer cost
      specialAnnual = ctc - (basicAnnual + hraAnnual + siteAllowance + otherBenefits + emprPfAnnual + gratuityAnnual);
    }
    const specialMonthly = specialAnnual / 12;

    // 6. Total CTC Check (Row 24 & Row 28)
    let calculatedCtcAnnual = 0;
    if (mathMode === 'excel') {
      calculatedCtcAnnual = basicAnnual + hraAnnual + specialAnnual + siteAllowance + otherBenefits + empPfAnnual + emprPfAnnual + gratuityAnnual;
    } else {
      calculatedCtcAnnual = basicAnnual + hraAnnual + specialAnnual + siteAllowance + otherBenefits + emprPfAnnual + gratuityAnnual;
    }
    const ctcVariance = calculatedCtcAnnual - ctc; // Must be 0

    // 7. Gross Salary
    // Standard Gross = Basic + HRA + Special Allowance + Site Allowance + Other Benefits
    let grossAnnual = basicAnnual + hraAnnual + specialAnnual + siteAllowance + otherBenefits;
    if (mathMode === 'excel') {
      // In Excel mode, employee PF was subtracted from Special Allowance, so Gross includes empPf
      grossAnnual += empPfAnnual;
    }
    const grossMonthly = grossAnnual / 12;

    // 8. Deductions: Professional Tax & Income Tax
    // Standard PT in India: ₹200 / month (approx ₹2,400 to ₹2,500 / year)
    const ptAnnual = ptApplicable ? (grossMonthly > 15000 ? 2400 : 0) : 0;
    const ptMonthly = ptAnnual / 12;

    // Income Tax (TDS) Estimation
    let taxAnnual = 0;
    if (taxRegime === 'new') {
      // FY 2024-25 / 2025-26 New Regime: Standard Deduction ₹75,000
      const standardDeduction = 75000;
      const taxableIncome = Math.max(0, grossAnnual - standardDeduction);

      if (taxableIncome > 700000) {
        // Compute Slabs
        let tax = 0;
        let rem = taxableIncome;

        // 0 - 3L: 0%
        rem = Math.max(0, taxableIncome - 300000);

        // 3L - 7L: 5% (on 4L = 20,000)
        let slab1 = Math.min(rem, 400000);
        tax += slab1 * 0.05;
        rem = Math.max(0, rem - 400000);

        // 7L - 10L: 10% (on 3L = 30,000)
        let slab2 = Math.min(rem, 300000);
        tax += slab2 * 0.10;
        rem = Math.max(0, rem - 300000);

        // 10L - 12L: 15% (on 2L = 30,000)
        let slab3 = Math.min(rem, 200000);
        tax += slab3 * 0.15;
        rem = Math.max(0, rem - 200000);

        // 12L - 15L: 20% (on 3L = 60,000)
        let slab4 = Math.min(rem, 300000);
        tax += slab4 * 0.20;
        rem = Math.max(0, rem - 300000);

        // Above 15L: 30%
        if (rem > 0) {
          tax += rem * 0.30;
        }

        // Add 4% Cess
        taxAnnual = tax * 1.04;
      } else {
        // Section 87A rebate makes tax zero for taxable income <= 7L
        taxAnnual = 0;
      }
    } else if (taxRegime === 'old') {
      // Old regime standard deduction ₹50,000 + 80C PF deduction
      const standardDeduction = 50000;
      const section80C = Math.min(150000, empPfAnnual);
      const taxableIncome = Math.max(0, grossAnnual - standardDeduction - section80C);

      if (taxableIncome > 500000) {
        let tax = 0;
        let rem = Math.max(0, taxableIncome - 250000);

        // 2.5L - 5L: 5% (12,500)
        let slab1 = Math.min(rem, 250000);
        tax += slab1 * 0.05;
        rem = Math.max(0, rem - 250000);

        // 5L - 10L: 20% (on 5L = 100,000)
        let slab2 = Math.min(rem, 500000);
        tax += slab2 * 0.20;
        rem = Math.max(0, rem - 500000);

        // Above 10L: 30%
        if (rem > 0) {
          tax += rem * 0.30;
        }
        taxAnnual = tax * 1.04;
      } else {
        taxAnnual = 0;
      }
    } else {
      taxAnnual = 0;
    }
    const taxMonthly = taxAnnual / 12;

    // Total Monthly Deductions
    const totalDeductAnnual = empPfAnnual + ptAnnual + taxAnnual;
    const totalDeductMonthly = totalDeductAnnual / 12;

    // Net Take-Home In Hand
    const netTakeHomeAnnual = grossAnnual - totalDeductAnnual;
    const netTakeHomeMonthly = netTakeHomeAnnual / 12;

    // Total Retirals (Employer Benefits)
    const retiralsAnnual = emprPfAnnual + gratuityAnnual;
    const retiralsMonthly = retiralsAnnual / 12;

    // Code on Wages 50% Compliance Check
    // Wages (Basic) must be >= 50% of (CTC or Gross)
    const basicRatio = ctc > 0 ? (basicAnnual / ctc) : 0;
    const isWageCodeCompliant = basicRatio >= 0.50;

    return {
      ctc,
      basicAnnual, basicMonthly,
      hraAnnual, hraMonthly,
      specialAnnual, specialMonthly,
      siteAllowance, siteMonthly: siteAllowance / 12,
      otherBenefits, otherMonthly: otherBenefits / 12,
      empPfAnnual, empPfMonthly,
      emprPfAnnual, emprPfMonthly,
      gratuityAnnual, gratuityMonthly,
      calculatedCtcAnnual,
      ctcVariance,
      grossAnnual, grossMonthly,
      ptAnnual, ptMonthly,
      taxAnnual, taxMonthly,
      totalDeductAnnual, totalDeductMonthly,
      netTakeHomeAnnual, netTakeHomeMonthly,
      retiralsAnnual, retiralsMonthly,
      basicRatio,
      isWageCodeCompliant,
      pfEligibleAnnual
    };
  }

  // --- RENDER TABLE & DATA ---
  function renderAll() {
    // 1. Calculate main salary model
    const m = calculateSalaryModel(state);

    // 2. Update CTC Inputs & In Words
    const displayCtc = state.isMonthlyInput ? (state.ctc / 12) : state.ctc;
    elements.ctcInput.value = Math.round(displayCtc);
    elements.ctcSlider.value = state.ctc;
    elements.ctcInWords.textContent = numberToWordsINR(state.ctc);

    // 3. Update Sliders & Percentage Inputs
    elements.basicPercentInput.value = state.basicPercent;
    elements.basicPercentSlider.value = state.basicPercent;
    elements.hraPercentInput.value = state.hraPercent;
    elements.hraPercentSlider.value = state.hraPercent;
    elements.siteAllowanceInput.value = state.siteAllowance;
    elements.otherBenefitsInput.value = state.otherBenefits;
    elements.projectBonusInput.value = state.projectBonus;

    // 4. Update KPIs
    elements.kpiNetMonthly.textContent = formatCurrency(m.netTakeHomeMonthly);
    elements.kpiNetAnnual.textContent = `Annual In-Hand: ${formatCurrency(m.netTakeHomeAnnual)}`;
    elements.kpiGrossMonthly.textContent = formatCurrency(m.grossMonthly);
    elements.kpiGrossAnnual.textContent = `Annual Gross: ${formatCurrency(m.grossAnnual)}`;
    elements.kpiCtcAnnual.textContent = formatCurrency(m.ctc);
    elements.kpiCtcMonthly.textContent = `Monthly CTC: ${formatCurrency(m.ctc / 12)}`;
    elements.kpiRetiralsAnnual.textContent = formatCurrency(m.retiralsAnnual);
    elements.kpiRetiralsMonthly.textContent = `PF + Gratuity: ${formatCurrency(m.retiralsMonthly)}/mo`;

    // 5. Update Compliance Banner
    if (m.isWageCodeCompliant) {
      elements.complianceBanner.className = 'compliance-banner';
      elements.complianceIndicator.className = 'compliance-indicator pass';
      elements.complianceTitle.textContent = `Code on Wages Compliance Check: PASSED (${(m.basicRatio * 100).toFixed(1)}% Basic)`;
      elements.complianceDetail.textContent = `Basic salary is ${(m.basicRatio * 100).toFixed(1)}% of CTC, safely exceeding the statutory 50% minimum wage rule.`;
      elements.autoFixWageCodeBtn.classList.add('hidden');
    } else {
      elements.complianceBanner.className = 'compliance-banner warning';
      elements.complianceIndicator.className = 'compliance-indicator warn';
      elements.complianceTitle.textContent = `Code on Wages Alert: Allowances exceed 50% threshold!`;
      elements.complianceDetail.textContent = `Basic salary is currently ${(m.basicRatio * 100).toFixed(1)}% of CTC. Under the Code on Wages, allowances over 50% are treated as wages for PF & Gratuity.`;
      elements.autoFixWageCodeBtn.classList.remove('hidden');
    }

    // 6. Render Tab 1: Excel Structure Table (Rows 17-25)
    renderExcelTable(m);

    // 7. Render Tab 2: Take-Home & Payslip View
    renderTakeHomeView(m);

    // 8. Render Tab 3: Statutory & Control Checks (B27:B31)
    renderStatutoryChecks(m);

    // 9. Render Tab 4: Visuals & Donut Chart
    renderVisuals(m);

    // 10. Render Tab 5: Offer Comparison
    renderOfferComparison();

    // 11. Render Annexure modal data
    renderAnnexureDocument(m);
  }

  // --- TAB 1: EXCEL SALARY STRUCTURE TABLE ---
  function renderExcelTable(m) {
    const specialFormula = state.mathMode === 'excel' 
      ? 'B24-SUM(...)' 
      : 'B24-SUM(B18:B19,B22:B23)';
    const specialFormulaFull = state.mathMode === 'excel'
      ? '=B24-SUM(B18:B19,B21:B23)'
      : '=B24-SUM(B18:B19,B22:B23)';

    const tableData = [
      {
        tag: 'A18',
        ref: 'B18',
        name: 'Basic Salary',
        annual: m.basicAnnual,
        monthly: m.basicMonthly,
        pct: m.ctc > 0 ? (m.basicAnnual / m.ctc) : 0,
        pill: 'B4*B5/100',
        basis: 'Basic as % of CTC',
        formula: '=B4*B5/100',
        explanation: 'Basic as % of CTC (Initial starting point; Indian payroll cornerstone).'
      },
      {
        tag: 'A19',
        ref: 'B19',
        name: 'House Rent Allowance',
        annual: m.hraAnnual,
        monthly: m.hraMonthly,
        pct: m.ctc > 0 ? (m.hraAnnual / m.ctc) : 0,
        pill: 'B18*B6/100',
        basis: 'HRA % of Basic',
        formula: '=B18*B6/100',
        explanation: 'HRA calculated as percentage of Basic Salary (40% Non-Metro, 50% Metro).'
      },
      {
        tag: 'A20',
        ref: 'B20',
        name: 'Special / Flexi Allowance',
        annual: m.specialAnnual,
        monthly: m.specialMonthly,
        pct: m.ctc > 0 ? (m.specialAnnual / m.ctc) : 0,
        pill: specialFormula,
        basis: 'Balancing allowance',
        formula: specialFormulaFull,
        explanation: 'Balancing allowance ensuring total calculated compensation matches Total CTC.'
      },
      {
        tag: 'A21',
        ref: 'B21',
        name: 'Employee PF Contribution',
        annual: m.empPfAnnual,
        monthly: m.empPfMonthly,
        pct: m.ctc > 0 ? (m.empPfAnnual / m.ctc) : 0,
        pill: 'IF(B9="Yes",...)',
        basis: 'Employee PF',
        formula: '=IF(B9="Yes",IF(B12="Yes",B18*B10/100,MIN(B18,B11*12)*B10/100),0)',
        explanation: '12% contribution by Employee, subject to statutory ₹15,000 monthly ceiling.'
      },
      {
        tag: 'A22',
        ref: 'B22',
        name: 'Employer PF Contribution',
        annual: m.emprPfAnnual,
        monthly: m.emprPfMonthly,
        pct: m.ctc > 0 ? (m.emprPfAnnual / m.ctc) : 0,
        pill: 'IF(B9="Yes",...)',
        basis: 'Employer PF',
        formula: '=IF(B9="Yes",IF(B12="Yes",B18*B10/100,MIN(B18,B11*12)*B10/100),0)',
        explanation: '12% statutory contribution by Employer (includes EPF & EPS pension provision).'
      },
      {
        tag: 'A23',
        ref: 'B23',
        name: 'Gratuity Provision',
        annual: m.gratuityAnnual,
        monthly: m.gratuityMonthly,
        pct: m.ctc > 0 ? (m.gratuityAnnual / m.ctc) : 0,
        pill: 'B18*B13/100',
        basis: '15/26/12 of eligible wage',
        formula: '=B18*B13/100',
        explanation: 'Annual provision based on statutory 15/26/12 method = 4.8077% of Basic.'
      }
    ];

    if (m.siteAllowance > 0) {
      tableData.splice(3, 0, {
        tag: 'A7',
        ref: 'B7',
        name: 'Site / Project Allowance',
        annual: m.siteAllowance,
        monthly: m.siteMonthly,
        pct: m.ctc > 0 ? (m.siteAllowance / m.ctc) : 0,
        pill: 'User Input',
        basis: 'Annual site allowance',
        formula: 'User Input (B7)',
        explanation: 'Annual site or field deployment allowance.'
      });
    }

    if (m.otherBenefits > 0) {
      tableData.splice(4, 0, {
        tag: 'A8',
        ref: 'B8',
        name: 'Other Benefits / Perks',
        annual: m.otherBenefits,
        monthly: m.otherMonthly,
        pct: m.ctc > 0 ? (m.otherBenefits / m.ctc) : 0,
        pill: 'User Input',
        basis: 'Annual perks amount',
        formula: 'User Input (B8)',
        explanation: 'Other company provided allowances and executive perks.'
      });
    }

    // Populate table body
    elements.salaryTableBody.innerHTML = '';
    
    // Rows A18 - A23
    tableData.forEach(row => {
      const tr = document.createElement('tr');
      if (row.ref === state.activeInspectorRow) tr.classList.add('selected');

      tr.innerHTML = `
        <td class="col-component">
          <div class="comp-cell">
            <span class="cell-tag">${row.tag}</span>
            <span class="comp-name">${row.name}</span>
          </div>
        </td>
        <td class="col-num">${formatNumber(row.annual)}</td>
        <td class="col-num">${formatNumber(row.monthly)}</td>
        <td class="col-pct">${formatPercent(row.pct)}</td>
        <td class="col-formula">
          <div class="formula-cell">
            <span class="formula-pill">${row.pill}</span>
            <span class="formula-basis-text">${row.basis}</span>
          </div>
        </td>
      `;

      tr.addEventListener('click', () => {
        state.activeInspectorRow = row.ref;
        elements.inspectorCellRef.textContent = `Cell ${row.tag} · ${row.name}`;
        elements.inspectorFormula.textContent = row.formula;
        elements.inspectorExplanation.textContent = row.explanation;
        document.querySelectorAll('#salaryTableBody tr').forEach(r => r.classList.remove('selected'));
        tr.classList.add('selected');
      });

      elements.salaryTableBody.appendChild(tr);
    });

    // Row A24: Total Annual CTC (Highlighted Total Row)
    const totalTr = document.createElement('tr');
    totalTr.className = 'total-highlight-row';
    if (state.activeInspectorRow === 'B24') totalTr.classList.add('selected');

    totalTr.innerHTML = `
      <td class="col-component">
        <div class="comp-cell">
          <span class="cell-tag">A24</span>
          <span class="comp-name">Total Annual CTC</span>
        </div>
      </td>
      <td class="col-num">${formatNumber(m.ctc)}</td>
      <td class="col-num">${formatNumber(m.ctc / 12)}</td>
      <td class="col-pct">100.00%</td>
      <td class="col-formula">
        <div class="formula-cell">
          <span class="formula-pill">=B4 / Sum(C18:C23)</span>
          <span class="formula-basis-text">Must equal Annual CTC</span>
        </div>
      </td>
    `;

    totalTr.addEventListener('click', () => {
      state.activeInspectorRow = 'B24';
      elements.inspectorCellRef.textContent = 'Cell A24 · Total Annual CTC';
      elements.inspectorFormula.textContent = '=B4';
      elements.inspectorExplanation.textContent = 'Total Annual Cost to Company. Must exactly equal entered CTC.';
      document.querySelectorAll('#salaryTableBody tr').forEach(r => r.classList.remove('selected'));
      totalTr.classList.add('selected');
    });

    elements.salaryTableBody.appendChild(totalTr);

    // Row A25: Site / Project Bonus (Italic user input row)
    const bonusTr = document.createElement('tr');
    bonusTr.className = 'bonus-row';
    if (state.activeInspectorRow === 'B25') bonusTr.classList.add('selected');

    const bonusValText = state.projectBonus 
      ? `${formatNumber(state.projectBonus)} or The Amount calculated on the basis of Project completion`
      : 'The Amount calculated on the basis of Project completion';

    bonusTr.innerHTML = `
      <td class="col-component">
        <div class="comp-cell">
          <span class="cell-tag">A25</span>
          <span class="comp-name">Site / Project Bonus</span>
        </div>
      </td>
      <td colspan="2" class="bonus-italic-text">${bonusValText}</td>
      <td class="col-pct">0.00%</td>
      <td class="col-formula">
        <div class="formula-cell">
          <span class="formula-basis-text">User input</span>
        </div>
      </td>
    `;

    bonusTr.addEventListener('click', () => {
      state.activeInspectorRow = 'B25';
      elements.inspectorCellRef.textContent = 'Cell A25 · Site / Project Bonus';
      elements.inspectorFormula.textContent = 'User input';
      elements.inspectorExplanation.textContent = 'Variable bonus based on milestone achievement or project completion.';
      document.querySelectorAll('#salaryTableBody tr').forEach(r => r.classList.remove('selected'));
      bonusTr.classList.add('selected');
    });

    elements.salaryTableBody.appendChild(bonusTr);
  }

  // --- TAB 2: TAKE-HOME (IN-HAND) VIEW ---
  function renderTakeHomeView(m) {
    elements.eqGrossMonthly.textContent = formatCurrency(m.grossMonthly);
    elements.eqDeductMonthly.textContent = formatCurrency(m.totalDeductMonthly);
    elements.eqNetMonthly.textContent = formatCurrency(m.netTakeHomeMonthly);

    elements.thBasic.textContent = formatCurrency(m.basicMonthly);
    elements.thHra.textContent = formatCurrency(m.hraMonthly);
    elements.thSpecial.textContent = formatCurrency(m.specialMonthly);

    if (m.siteAllowance > 0) {
      elements.thSiteRow.style.display = 'flex';
      elements.thSite.textContent = formatCurrency(m.siteMonthly);
    } else {
      elements.thSiteRow.style.display = 'none';
    }

    if (m.otherBenefits > 0) {
      elements.thOtherRow.style.display = 'flex';
      elements.thOther.textContent = formatCurrency(m.otherMonthly);
    } else {
      elements.thOtherRow.style.display = 'none';
    }

    elements.thGrossTotal.textContent = formatCurrency(m.grossMonthly);
    elements.thEmpPf.textContent = formatCurrency(m.empPfMonthly);
    elements.thPt.textContent = formatCurrency(m.ptMonthly);
    elements.thTax.textContent = formatCurrency(m.taxMonthly);
    elements.thDeductTotal.textContent = formatCurrency(m.totalDeductMonthly);

    // Regime Badge
    if (state.taxRegime === 'new') {
      elements.activeRegimeBadge.textContent = 'New Tax Regime (FY 2024-26)';
      elements.taxNoteText.textContent = 'Calculated with ₹75,000 standard deduction. Annual income up to ₹7,75,000 is 100% tax-free under Section 87A rebate.';
    } else if (state.taxRegime === 'old') {
      elements.activeRegimeBadge.textContent = 'Old Tax Regime (with 80C PF)';
      elements.taxNoteText.textContent = 'Includes ₹50,000 standard deduction and employee PF tax deduction under Section 80C (up to ₹1,50,000 cap).';
    } else {
      elements.activeRegimeBadge.textContent = 'Zero Tax Mode';
      elements.taxNoteText.textContent = 'Income tax TDS calculation is disabled. Showing gross take-home after statutory PF and PT deductions.';
    }
  }

  // --- TAB 3: STATUTORY & CONTROL CHECKS ---
  function renderStatutoryChecks(m) {
    // Row 28: Total CTC Check
    const diff = Math.abs(m.ctcVariance);
    elements.checkCtcVal.textContent = `${formatCurrency(diff)} Difference`;
    if (diff < 0.05) {
      elements.checkCtcStatus.className = 'check-status-badge pass';
      elements.checkCtcStatus.textContent = 'PASSED (0.00)';
    } else {
      elements.checkCtcStatus.className = 'check-status-badge warn';
      elements.checkCtcStatus.textContent = 'MISMATCH';
    }

    // Row 29: Wage / Allowance Check
    const basicPct = (m.basicRatio * 100).toFixed(1);
    const allowancePct = (100 - basicPct).toFixed(1);
    elements.checkWageVal.textContent = `${basicPct}% Basic / ${allowancePct}% Allowances`;
    elements.wageMeterBar.style.width = `${Math.min(100, Math.max(0, basicPct))}%`;

    if (m.isWageCodeCompliant) {
      elements.checkWageStatus.className = 'check-status-badge pass';
      elements.checkWageStatus.textContent = 'COMPLIANT';
      elements.wageMeterBar.className = 'meter-bar-fill';
      elements.checkWageDesc.textContent = `Complies with the 50% statutory wage rule under the Code on Wages. Basic salary is ${basicPct}%, preventing extra PF/Gratuity tax liabilities.`;
    } else {
      elements.checkWageStatus.className = 'check-status-badge warn';
      elements.checkWageStatus.textContent = 'NON-COMPLIANT';
      elements.wageMeterBar.className = 'meter-bar-fill warning';
      elements.checkWageDesc.textContent = `Allowances (${allowancePct}%) exceed the 50% limit. Under Indian Code on Wages, the surplus allowance (${(allowancePct - 50).toFixed(1)}%) must be deemed as wages.`;
    }

    // Row 30: PF Monthly Wage Check
    elements.checkPfWageVal.textContent = `${formatCurrency(m.basicMonthly)} / month`;
    if (!state.pfApplicable) {
      elements.checkPfStatus.className = 'check-status-badge info';
      elements.checkPfStatus.textContent = 'EXEMPT';
      elements.checkPfDesc.textContent = 'Provident Fund is marked as Not Applicable for this salary structure.';
    } else if (state.pfHigherWage) {
      elements.checkPfStatus.className = 'check-status-badge info';
      elements.checkPfStatus.textContent = 'HIGHER WAGE (FULL)';
      elements.checkPfDesc.textContent = `PF contribution is actively computed on the full monthly basic wage of ${formatCurrency(m.basicMonthly)} without statutory capping.`;
    } else {
      elements.checkPfStatus.className = 'check-status-badge info';
      elements.checkPfStatus.textContent = 'STATUTORY CAPPED';
      elements.checkPfDesc.textContent = `PF contribution is capped at the statutory ceiling of ₹15,000/month. Monthly PF contribution is ${formatCurrency(m.empPfMonthly)}.`;
    }

    // Row 31: Gratuity Provision Check
    elements.checkGratuityVal.textContent = `${formatCurrency(m.gratuityAnnual)} / year`;
    // Gratuity payout = 15/26 * Last drawn basic * tenure (e.g. 5 years)
    const gratuity5yr = (15 / 26) * m.basicMonthly * 5;
    elements.gratuity5yrPayout.textContent = formatCurrency(gratuity5yr);
  }

  // --- TAB 4: VISUAL ANALYTICS & CHARTS ---
  function renderVisuals(m) {
    elements.donutCenterVal.textContent = state.currency === 'INR' 
      ? (m.ctc >= 10000000 ? `${(m.ctc / 10000000).toFixed(2)} Cr` : `${(m.ctc / 100000).toFixed(1)}L`)
      : `$${(m.ctc / 1000).toFixed(0)}k`;

    // Segments: Basic, HRA, Special Allowance, Retirals (PF + Gratuity)
    const segments = [
      { name: 'Basic Salary', val: m.basicAnnual, color: 'var(--accent-indigo)' },
      { name: 'House Rent Allowance (HRA)', val: m.hraAnnual, color: 'var(--accent-cyan)' },
      { name: 'Special / Flexi Allowance', val: Math.max(0, m.specialAnnual), color: 'var(--accent-violet)' },
      { name: 'Provident Fund (Employer)', val: m.emprPfAnnual, color: '#3B82F6' },
      { name: 'Gratuity Provision', val: m.gratuityAnnual, color: '#EAB308' }
    ];

    if (m.siteAllowance > 0) {
      segments.push({ name: 'Site Allowance', val: m.siteAllowance, color: '#F97316' });
    }
    if (m.otherBenefits > 0) {
      segments.push({ name: 'Other Benefits', val: m.otherBenefits, color: '#EC4899' });
    }

    const total = segments.reduce((sum, s) => sum + s.val, 0);
    const circumference = 2 * Math.PI * 40; // ~251.32

    // Build SVG donut
    let strokeOffset = 0;
    let svgContent = `<circle cx="50" cy="50" r="40" class="donut-bg" stroke-width="14"></circle>`;
    let legendContent = '';

    segments.forEach(seg => {
      const pct = total > 0 ? (seg.val / total) : 0;
      const strokeLength = pct * circumference;
      const dashArray = `${strokeLength} ${circumference - strokeLength}`;

      svgContent += `
        <circle cx="50" cy="50" r="40"
          stroke="${seg.color}"
          stroke-width="14"
          stroke-dasharray="${dashArray}"
          stroke-dashoffset="${-strokeOffset}"
          class="donut-segment"
        ></circle>
      `;

      strokeOffset += strokeLength;

      legendContent += `
        <div class="legend-item">
          <div class="legend-item-left">
            <span class="legend-color-dot" style="background: ${seg.color};"></span>
            <span>${seg.name}</span>
          </div>
          <span class="legend-item-val">${formatCurrency(seg.val)} (${(pct * 100).toFixed(1)}%)</span>
        </div>
      `;
    });

    elements.donutSvg.innerHTML = svgContent;
    elements.chartLegend.innerHTML = legendContent;

    // Monthly Cashflow Flowchart
    elements.flowCtcMonthly.textContent = formatCurrency(m.ctc / 12);
    elements.flowNetMonthly.textContent = formatCurrency(m.netTakeHomeMonthly);

    const monthlyCtc = m.ctc / 12;
    if (monthlyCtc > 0) {
      const netPct = Math.max(0, (m.netTakeHomeMonthly / monthlyCtc) * 100);
      const taxPct = Math.max(0, ((m.taxMonthly + m.ptMonthly) / monthlyCtc) * 100);
      const pfPct = Math.max(0, ((m.empPfMonthly + m.emprPfMonthly) / monthlyCtc) * 100);
      const gratuityPct = Math.max(0, (m.gratuityMonthly / monthlyCtc) * 100);

      elements.segNet.style.width = `${netPct}%`;
      elements.segTax.style.width = `${taxPct}%`;
      elements.segPf.style.width = `${pfPct}%`;
      elements.segGratuity.style.width = `${gratuityPct}%`;
    }
  }

  // --- TAB 5: OFFER COMPARISON ---
  function renderOfferComparison() {
    const offerAVal = Number(elements.offerAInput.value) || 0;
    const offerBVal = Number(elements.offerBInput.value) || 0;

    const modA = calculateSalaryModel({ ...state, ctc: offerAVal });
    const modB = calculateSalaryModel({ ...state, ctc: offerBVal });

    const rows = [
      { metric: 'Annual CTC', a: modA.ctc, b: modB.ctc },
      { metric: 'Monthly In-Hand (Net)', a: modA.netTakeHomeMonthly, b: modB.netTakeHomeMonthly },
      { metric: 'Annual In-Hand (Net)', a: modA.netTakeHomeAnnual, b: modB.netTakeHomeAnnual },
      { metric: 'Gross Monthly', a: modA.grossMonthly, b: modB.grossMonthly },
      { metric: 'Total Annual PF (Employee + Employer)', a: modA.empPfAnnual + modA.emprPfAnnual, b: modB.empPfAnnual + modB.emprPfAnnual },
      { metric: 'Annual Gratuity Provision', a: modA.gratuityAnnual, b: modB.gratuityAnnual },
      { metric: 'Estimated Annual Tax (TDS)', a: modA.taxAnnual, b: modB.taxAnnual }
    ];

    elements.comparisonTableBody.innerHTML = '';
    rows.forEach(r => {
      const diff = r.b - r.a;
      const pctGrowth = r.a > 0 ? ((diff / r.a) * 100).toFixed(1) : '0';
      const isPositive = diff >= 0;

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td><strong>${r.metric}</strong></td>
        <td>${formatCurrency(r.a)}</td>
        <td>${formatCurrency(r.b)}</td>
        <td>
          <span class="variance-pill ${isPositive ? 'positive' : 'negative'}">
            ${isPositive ? '+' : ''}${formatCurrency(diff)} (${isPositive ? '+' : ''}${pctGrowth}%)
          </span>
        </td>
      `;
      elements.comparisonTableBody.appendChild(tr);
    });
  }

  // --- PRINTABLE ANNEXURE ---
  function renderAnnexureDocument(m) {
    const today = new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
    elements.annexureDate.textContent = today;
    elements.annexureCtcVal.textContent = formatCurrency(m.ctc);

    const items = [
      { name: 'Basic Salary', m: m.basicMonthly, a: m.basicAnnual, pct: m.ctc > 0 ? (m.basicAnnual / m.ctc) : 0 },
      { name: 'House Rent Allowance (HRA)', m: m.hraMonthly, a: m.hraAnnual, pct: m.ctc > 0 ? (m.hraAnnual / m.ctc) : 0 },
      { name: 'Special / Flexi Allowance', m: m.specialMonthly, a: m.specialAnnual, pct: m.ctc > 0 ? (m.specialAnnual / m.ctc) : 0 },
      { name: 'Site / Project Allowance', m: m.siteMonthly, a: m.siteAllowance, pct: m.ctc > 0 ? (m.siteAllowance / m.ctc) : 0 },
      { name: 'Other Benefits / Executive Perks', m: m.otherMonthly, a: m.otherBenefits, pct: m.ctc > 0 ? (m.otherBenefits / m.ctc) : 0 },
      { name: 'Employer PF Contribution (12%)', m: m.emprPfMonthly, a: m.emprPfAnnual, pct: m.ctc > 0 ? (m.emprPfAnnual / m.ctc) : 0 },
      { name: 'Gratuity Provision (15/26/12 method)', m: m.gratuityMonthly, a: m.gratuityAnnual, pct: m.ctc > 0 ? (m.gratuityAnnual / m.ctc) : 0 }
    ].filter(i => i.a > 0);

    elements.annexureTableBody.innerHTML = '';
    items.forEach(item => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${item.name}</td>
        <td>${formatCurrency(item.m)}</td>
        <td>${formatCurrency(item.a)}</td>
        <td>${formatPercent(item.pct)}</td>
      `;
      elements.annexureTableBody.appendChild(tr);
    });

    elements.annexureTableFoot.innerHTML = `
      <tr>
        <td><strong>Total Cost To Company (CTC)</strong></td>
        <td><strong>${formatCurrency(m.ctc / 12)}</strong></td>
        <td><strong>${formatCurrency(m.ctc)}</strong></td>
        <td><strong>100.00%</strong></td>
      </tr>
    `;
  }

  // --- COPY SALARY SUMMARY TO CLIPBOARD ---
  function copySalarySummary() {
    const m = calculateSalaryModel(state);
    const bonusText = state.projectBonus 
      ? `${formatNumber(state.projectBonus)} or The Amount calculated on the basis of Project completion`
      : 'The Amount calculated on the basis of Project completion';

    const specialPill = state.mathMode === 'excel' ? 'B24-SUM(...)' : 'B24-SUM(B18:B19,B22:B23)';

    const text = `
CALCULATED SALARY STRUCTURE
Live computation replicating Sheet 2 formulas cell-by-cell

COMPONENT                          ANNUAL (₹)       MONTHLY (₹)     % OF CTC    FORMULA / BASIS
---------------------------------------------------------------------------------------------------------
A18  Basic Salary                  ${formatNumber(m.basicAnnual).padStart(14)}   ${formatNumber(m.basicMonthly).padStart(13)}      ${(m.basicRatio * 100).toFixed(2)}%    [B4*B5/100] Basic as % of CTC
A19  House Rent Allowance          ${formatNumber(m.hraAnnual).padStart(14)}   ${formatNumber(m.hraMonthly).padStart(13)}      ${((m.hraAnnual / m.ctc) * 100).toFixed(2)}%    [B18*B6/100] HRA % of Basic
A20  Special / Flexi Allowance     ${formatNumber(m.specialAnnual).padStart(14)}   ${formatNumber(m.specialMonthly).padStart(13)}      ${((m.specialAnnual / m.ctc) * 100).toFixed(2)}%    [${specialPill}] Balancing allowance
A21  Employee PF Contribution      ${formatNumber(m.empPfAnnual).padStart(14)}   ${formatNumber(m.empPfMonthly).padStart(13)}      ${((m.empPfAnnual / m.ctc) * 100).toFixed(2)}%    [IF(B9="Yes",...)] Employee PF
A22  Employer PF Contribution      ${formatNumber(m.emprPfAnnual).padStart(14)}   ${formatNumber(m.emprPfMonthly).padStart(13)}      ${((m.emprPfAnnual / m.ctc) * 100).toFixed(2)}%    [IF(B9="Yes",...)] Employer PF
A23  Gratuity Provision            ${formatNumber(m.gratuityAnnual).padStart(14)}   ${formatNumber(m.gratuityMonthly).padStart(13)}      ${((m.gratuityAnnual / m.ctc) * 100).toFixed(2)}%    [B18*B13/100] 15/26/12 of eligible wage
---------------------------------------------------------------------------------------------------------
A24  Total Annual CTC              ${formatNumber(m.ctc).padStart(14)}   ${formatNumber(m.ctc / 12).padStart(13)}     100.00%    [=B4 / Sum(C18:C23)] Must equal Annual CTC
A25  Site / Project Bonus          ${bonusText}     0.00%    User input
---------------------------------------------------------------------------------------------------------
Est. Monthly In-Hand: ${formatCurrency(m.netTakeHomeMonthly)}
Gross Monthly:        ${formatCurrency(m.grossMonthly)}
`.trim();

    navigator.clipboard.writeText(text).then(() => {
      const btn = elements.copyBtnText || document.getElementById('copyBtnText');
      if (btn) {
        const orig = btn.textContent;
        btn.textContent = '✓ Copied!';
        setTimeout(() => { btn.textContent = orig; }, 2000);
      }
    }).catch(() => {
      alert('Salary summary copied to clipboard!');
    });
  }

  // --- EXCEL EXPORT (SHEETJS) ---
  function exportToExcel() {
    if (typeof XLSX === 'undefined') {
      alert('SheetJS library is not loaded. Cannot generate Excel file.');
      return;
    }

    const m = calculateSalaryModel(state);
    const wb = XLSX.utils.book_new();

    // Replicate exactly the Sheet2 layout from 'Salary Calculater.xlsx'
    const wsData = [
      ['ApexSalary CTC Breakdown Report'],
      ['Generated on', new Date().toISOString()],
      ['INPUTS'],
      ['Annual CTC', m.ctc, 'Enter annual CTC'],
      ['Basic % of CTC (initial)', state.basicPercent, 'Used as starting point; calculator adjusts for employer contributions'],
      ['HRA % of Basic', state.hraPercent, 'Common payroll structure; tax exemption differs by tax rules'],
      ['Site / Project Allowance', state.siteAllowance, 'Annual amount'],
      ['Other Benefits', state.otherBenefits, 'Annual amount'],
      ['PF Applicable?', state.pfApplicable ? 'Yes' : 'No', 'Yes/No'],
      ['PF Rate', state.pfRate, 'Employer statutory contribution rate'],
      ['PF Wage Ceiling', state.pfWageCeiling, 'Monthly statutory ceiling; change if organisation contributes on higher wages'],
      ['PF on Higher Wages?', state.pfHigherWage ? 'Yes' : 'No', 'Yes = PF on full PF wage; No = statutory ceiling'],
      ['Gratuity Rate', state.gratuityRate, '15/26/12 = 4.8077% of eligible monthly wage annually'],
      [''],
      [''],
      ['CALCULATED SALARY STRUCTURE'],
      ['Component', 'Annual', 'Monthly', '% of CTC', 'Formula / Basis'],
      ['Basic Salary', m.basicAnnual, m.basicMonthly, m.ctc > 0 ? (m.basicAnnual / m.ctc) : 0, '=B4*B5/100'],
      ['House Rent Allowance', m.hraAnnual, m.hraMonthly, m.ctc > 0 ? (m.hraAnnual / m.ctc) : 0, '=B18*B6/100'],
      ['Special / Flexi Allowance', m.specialAnnual, m.specialMonthly, m.ctc > 0 ? (m.specialAnnual / m.ctc) : 0, '=B24-SUM(B18:B19,B21:B23)'],
      ['Employee PF Contribution', m.empPfAnnual, m.empPfMonthly, m.ctc > 0 ? (m.empPfAnnual / m.ctc) : 0, 'Employee PF'],
      ['Employer PF Contribution', m.emprPfAnnual, m.emprPfMonthly, m.ctc > 0 ? (m.emprPfAnnual / m.ctc) : 0, 'Employer PF'],
      ['Gratuity Provision', m.gratuityAnnual, m.gratuityMonthly, m.ctc > 0 ? (m.gratuityAnnual / m.ctc) : 0, '15/26/12 of eligible wage'],
      ['Total Annual CTC', m.ctc, m.ctc / 12, 1, '=B4'],
      ['Site / Project Bonus', state.projectBonus, '', 0, 'User input'],
      [''],
      ['COMPLIANCE / CONTROL CHECKS'],
      ['Total CTC Check', Math.abs(m.ctcVariance), 'Should be 0'],
      ['Wage / Allowance Check', m.isWageCodeCompliant ? 0 : 1, 'Review if allowances exceed 50% threshold under Code on Wages'],
      ['PF Monthly Wage', m.basicMonthly, 'Compare with PF ceiling / higher-wage option'],
      ['Gratuity Annual Provision', m.gratuityAnnual, 'Based on 15/26/12 method']
    ];

    const ws = XLSX.utils.aoa_to_sheet(wsData);

    // Apply column widths
    ws['!cols'] = [
      { wch: 30 },
      { wch: 18 },
      { wch: 45 },
      { wch: 14 },
      { wch: 25 }
    ];

    XLSX.utils.book_append_sheet(wb, ws, 'Sheet2');
    XLSX.writeFile(wb, `Salary_Calculater_Output_${m.ctc}.xlsx`);
  }

  // --- IMPORT EXCEL FILE ---
  function handleExcelImport(file) {
    if (!file || typeof XLSX === 'undefined') return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        const sheetName = wb.SheetNames.includes('Sheet2') ? 'Sheet2' : wb.SheetNames[0];
        const sheet = wb.Sheets[sheetName];

        if (sheet['B4'] && sheet['B4'].v !== undefined) {
          state.ctc = Number(sheet['B4'].v);
        }
        if (sheet['B5'] && sheet['B5'].v !== undefined) {
          state.basicPercent = Number(sheet['B5'].v);
        }
        if (sheet['B6'] && sheet['B6'].v !== undefined) {
          state.hraPercent = Number(sheet['B6'].v);
        }
        if (sheet['B7'] && sheet['B7'].v !== undefined) {
          state.siteAllowance = Number(sheet['B7'].v) || 0;
        }
        if (sheet['B8'] && sheet['B8'].v !== undefined) {
          state.otherBenefits = Number(sheet['B8'].v) || 0;
        }
        if (sheet['B9'] && sheet['B9'].v !== undefined) {
          state.pfApplicable = String(sheet['B9'].v).toLowerCase().startsWith('y');
        }
        if (sheet['B10'] && sheet['B10'].v !== undefined) {
          state.pfRate = Number(sheet['B10'].v);
        }
        if (sheet['B11'] && sheet['B11'].v !== undefined) {
          state.pfWageCeiling = Number(sheet['B11'].v);
        }
        if (sheet['B12'] && sheet['B12'].v !== undefined) {
          state.pfHigherWage = String(sheet['B12'].v).toLowerCase().startsWith('y');
        }
        if (sheet['B13'] && sheet['B13'].v !== undefined) {
          state.gratuityRate = Number(sheet['B13'].v);
        }

        renderAll();
        alert(`Successfully loaded data from ${file.name}!`);
      } catch (err) {
        console.error(err);
        alert('Could not parse Excel file: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  // --- ATTACH EVENT LISTENERS ---
  function setupEventListeners() {
    // Math Mode
    elements.mathModeSelect.addEventListener('change', (e) => {
      state.mathMode = e.target.value;
      renderAll();
    });

    elements.mathInfoBtn.addEventListener('click', () => {
      elements.mathInfoModal.classList.remove('hidden');
    });

    elements.mathInfoCloseBtn.addEventListener('click', () => {
      elements.mathInfoModal.classList.add('hidden');
    });

    elements.mathInfoModal.addEventListener('click', (e) => {
      if (e.target === elements.mathInfoModal) elements.mathInfoModal.classList.add('hidden');
    });

    // Currency Switcher
    elements.currInrBtn.addEventListener('click', () => {
      state.currency = 'INR';
      elements.currInrBtn.classList.add('active');
      elements.currUsdBtn.classList.remove('active');
      elements.currSymbolPrefix.textContent = '₹';
      renderAll();
    });

    elements.currUsdBtn.addEventListener('click', () => {
      state.currency = 'USD';
      elements.currUsdBtn.classList.add('active');
      elements.currInrBtn.classList.remove('active');
      elements.currSymbolPrefix.textContent = '$';
      renderAll();
    });

    // Presets
    elements.presetChips.forEach(chip => {
      chip.addEventListener('click', () => {
        elements.presetChips.forEach(c => c.classList.remove('active'));
        chip.classList.add('active');
        state.ctc = Number(chip.getAttribute('data-ctc'));
        renderAll();
      });
    });

    // Reset Defaults (Matches Salary Calculater.xlsx initial values)
    elements.resetDefaultsBtn.addEventListener('click', () => {
      state.ctc = 50000;
      state.basicPercent = 50;
      state.hraPercent = 40;
      state.siteAllowance = 0;
      state.otherBenefits = 0;
      state.projectBonus = 500000;
      state.pfApplicable = true;
      state.pfRate = 12;
      state.pfWageCeiling = 15000;
      state.pfHigherWage = false;
      state.gratuityRate = 4.807692308;
      state.mathMode = 'excel';
      elements.mathModeSelect.value = 'excel';
      renderAll();
    });

    // Excel File Upload
    elements.importExcelBtn.addEventListener('click', () => {
      elements.excelFileInput.click();
    });

    elements.excelFileInput.addEventListener('change', (e) => {
      if (e.target.files && e.target.files[0]) {
        handleExcelImport(e.target.files[0]);
      }
    });

    // Copy Summary
    if (elements.copySummaryBtn) {
      elements.copySummaryBtn.addEventListener('click', copySalarySummary);
    }

    // Excel Export
    elements.exportExcelBtn.addEventListener('click', exportToExcel);

    // Print & Annexure Modal
    elements.printAnnexureBtn.addEventListener('click', () => {
      elements.annexureModal.classList.remove('hidden');
    });

    elements.modalCloseBtn.addEventListener('click', () => {
      elements.annexureModal.classList.add('hidden');
    });

    elements.annexureModal.addEventListener('click', (e) => {
      if (e.target === elements.annexureModal) elements.annexureModal.classList.add('hidden');
    });

    elements.modalPrintBtn.addEventListener('click', () => {
      window.print();
    });

    // CTC Annual vs Monthly input toggle
    elements.ctcAnnualToggle.addEventListener('click', () => {
      if (state.isMonthlyInput) {
        state.isMonthlyInput = false;
        elements.ctcAnnualToggle.classList.add('active');
        elements.ctcMonthlyToggle.classList.remove('active');
        renderAll();
      }
    });

    elements.ctcMonthlyToggle.addEventListener('click', () => {
      if (!state.isMonthlyInput) {
        state.isMonthlyInput = true;
        elements.ctcMonthlyToggle.classList.add('active');
        elements.ctcAnnualToggle.classList.remove('active');
        renderAll();
      }
    });

    // CTC Inputs
    elements.ctcInput.addEventListener('input', (e) => {
      const val = Number(e.target.value) || 0;
      state.ctc = state.isMonthlyInput ? (val * 12) : val;
      renderAll();
    });

    elements.ctcSlider.addEventListener('input', (e) => {
      state.ctc = Number(e.target.value);
      renderAll();
    });

    // Basic % of CTC
    elements.basicPercentSlider.addEventListener('input', (e) => {
      state.basicPercent = Number(e.target.value);
      elements.basicPercentInput.value = state.basicPercent;
      renderAll();
    });

    elements.basicPercentInput.addEventListener('input', (e) => {
      state.basicPercent = Number(e.target.value);
      elements.basicPercentSlider.value = state.basicPercent;
      renderAll();
    });

    // HRA % of Basic
    elements.hraPercentSlider.addEventListener('input', (e) => {
      state.hraPercent = Number(e.target.value);
      elements.hraPercentInput.value = state.hraPercent;
      elements.hraMetroBtn.classList.remove('active');
      elements.hraNonMetroBtn.classList.remove('active');
      renderAll();
    });

    elements.hraPercentInput.addEventListener('input', (e) => {
      state.hraPercent = Number(e.target.value);
      elements.hraPercentSlider.value = state.hraPercent;
      elements.hraMetroBtn.classList.remove('active');
      elements.hraNonMetroBtn.classList.remove('active');
      renderAll();
    });

    elements.hraMetroBtn.addEventListener('click', () => {
      state.hraPercent = 50;
      elements.hraMetroBtn.classList.add('active');
      elements.hraNonMetroBtn.classList.remove('active');
      renderAll();
    });

    elements.hraNonMetroBtn.addEventListener('click', () => {
      state.hraPercent = 40;
      elements.hraNonMetroBtn.classList.add('active');
      elements.hraMetroBtn.classList.remove('active');
      renderAll();
    });

    // Site Allowance & Other Benefits
    elements.siteAllowanceInput.addEventListener('input', (e) => {
      state.siteAllowance = Number(e.target.value) || 0;
      renderAll();
    });

    elements.otherBenefitsInput.addEventListener('input', (e) => {
      state.otherBenefits = Number(e.target.value) || 0;
      renderAll();
    });

    elements.projectBonusInput.addEventListener('input', (e) => {
      state.projectBonus = Number(e.target.value) || 0;
      renderAll();
    });

    // Statutory Accordion
    elements.statutoryToggleBtn.addEventListener('click', () => {
      elements.statutoryAccordion.classList.toggle('open');
    });

    elements.pfApplicableCheck.addEventListener('change', (e) => {
      state.pfApplicable = e.target.checked;
      renderAll();
    });

    elements.pfRateInput.addEventListener('input', (e) => {
      state.pfRate = Number(e.target.value) || 12;
      renderAll();
    });

    elements.pfCeilingInput.addEventListener('input', (e) => {
      state.pfWageCeiling = Number(e.target.value) || 15000;
      renderAll();
    });

    elements.pfHigherWageCheck.addEventListener('change', (e) => {
      state.pfHigherWage = e.target.checked;
      renderAll();
    });

    elements.gratuityRateInput.addEventListener('input', (e) => {
      state.gratuityRate = Number(e.target.value) || 4.8077;
      renderAll();
    });

    elements.ptApplicableCheck.addEventListener('change', (e) => {
      state.ptApplicable = e.target.checked;
      renderAll();
    });

    elements.taxRegimeSelect.addEventListener('change', (e) => {
      state.taxRegime = e.target.value;
      renderAll();
    });

    // Auto-fix Wage Code button
    elements.autoFixWageCodeBtn.addEventListener('click', () => {
      state.basicPercent = 50;
      renderAll();
    });

    // Tabs Navigation
    elements.tabButtons.forEach(btn => {
      btn.addEventListener('click', () => {
        const tabId = btn.getAttribute('data-tab');
        elements.tabButtons.forEach(b => {
          b.classList.remove('active');
          b.setAttribute('aria-selected', 'false');
        });
        elements.tabPanels.forEach(p => p.classList.remove('active'));

        btn.classList.add('active');
        btn.setAttribute('aria-selected', 'true');
        const targetPanel = document.getElementById(tabId);
        if (targetPanel) targetPanel.classList.add('active');
      });
    });

    // Offer Comparison inputs
    elements.offerAInput.addEventListener('input', renderOfferComparison);
    elements.offerBInput.addEventListener('input', renderOfferComparison);
  }

  // --- INITIALIZE APPLICATION ---
  function init() {
    setupEventListeners();
    renderAll();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

})();
