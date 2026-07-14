import { useState } from "react";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import Fields from "../models/Fields";


async function loadData(path) {
  try {
    const response = await fetch(path); 
    if (!response.ok) throw new Error('Network response was not ok');

    const data = await response.json();
    return data


  } catch (error) {
    console.error('Error reading JSON:', error);
  }
  
}

// ─── Sample data ──────────────────────────────────────────────────────────────
const SAMPLE_HEADER = {
  "data": {
    "ConsumerName": "Druvanshi Patil",
    "UnitNo": "105",
    "BuildingName": "ROYAL RESIDENCE 1",
    "InvoiceNumber": "RR16410",
    "AccountNumber": "RR1105T4",
    "BillingDate": null,
    "DueDate": null,
    "BillingPeriod": "01/06/2026 - 01/06/2026",
    "OperatingCompanyAddress": "LYNX District Cooling Services\nDubai Stadium point- Unit 312, Dubai, UAE",
    "ContactNumber": "Call: +971 45762953",
    "WhatsAppNumber": "Whatsapp: +971 585536949"
  },
  "labels": {
    "DueDate": "Due Date:\nتاريخ الاستحقاق",
    "Name": "Name:\nالاسم",
    "UnitNo": "Unit No.:\nرقم الوحدة",
    "BillingPeriod": "Billing Period:\nفترة الفوترة",
    "AccountNumber": "Account Number:\nرقم الحساب",
    "TaxInvoice": "TAX INVOICE\nفاتورة ضريبية",
    "BillingDate": "Billing Date:\nتاريخ الفاتورة",
    "InvoiceNumber": "Invoice Number:\nرقم الفاتورة",
    "TowerName": "Tower Name:\nاسم البرج"
  }
}
const SAMPLE_DETAIL = {
    "data": [
        {
            "SrNo": 1,
            "MeterID": 487,
            "MeterSerialNumber": "69294862",
            "CommercialDataTagID": 11,
            "Description": "Utility Consumption(Cooled energy) Rate(SerialNumber:69294862)\n(69294862: استهلاك المرافق (الطاقة المبردة) (الرقم التسلسلي",
            "PrevReading": 74309,
            "PrevReadingUOM": "kWh",
            "CurrReading": 53092,
            "CurrReadingUOM": "kWh",
            "Qty": 999978783,
            "QtyUOM": "kWh",
            "Rate": 0.1592,
            "RateUOM": "AED/kWh",
            "TaxableAmount": 159196622.25,
            "TaxableAmountUOM": "AED",
            "VAT": 7959831.11,
            "VATUOM": "AED",
            "VATPercetage": 5,
            "VATPercetageUOM": "%",
            "TotalAmount": 167156453.36,
            "TotalAmountUOM": "AED",
            "lblPreviousBalance": "Previous Balance (AED)",
            "lblAdjustmentAmount": "Adjustment Amount (AED)",
            "PreviousBalance": 768.17,
            "PreviousBalanceUOM": "AED",
            "lblTotalAmountPayableWithinDueDate": "Total Amount Payable Within Due Date (AED)",
            "TotalAmountPayableWithinDueDate": 189521886.07,
            "TotalAmountPayableWithinDueDateUOM": "AED",
            "lblAdvancePayment": "Wallet Balance / Advance Payment (AED)",
            "AdvancePayment": 0,
            "AdvancePaymentUOM": "AED",
            "IsAdvancePayment": true,
            "AdjustmentAmount": 0,
            "AdjustmentAmountUOM": "AED",
            "IsAdjustmentAmount": true,
            "SumTaxableAmount": 180496302.77,
            "SumTaxableAmountUOM": "AED",
            "SumVAT": 9024815.13,
            "SumVATUOM": "AED",
            "SumTotalAmount": 189521117.9,
            "SumTotalAmountUOM": "AED",
            "TRhCorrectionFactor": 1,
            "RTHMeterDataValueUOM": "TRh",
            "EstimatedFlag": null,
            "IsInitialBill": false
        },
        {
            "SrNo": 2,
            "MeterID": null,
            "MeterSerialNumber": null,
            "CommercialDataTagID": 64,
            "Description": "Fuel Surcharge Rate\nرسوم الوقود الإضافية",
            "PrevReading": null,
            "PrevReadingUOM": "kWh",
            "CurrReading": null,
            "CurrReadingUOM": "kWh",
            "Qty": 999978783,
            "QtyUOM": "kWh",
            "Rate": 0.0213,
            "RateUOM": "AED",
            "TaxableAmount": 21299548.08,
            "TaxableAmountUOM": "AED",
            "VAT": 1064977.4,
            "VATUOM": "AED",
            "VATPercetage": 5,
            "VATPercetageUOM": "%",
            "TotalAmount": 22364525.48,
            "TotalAmountUOM": "AED",
            "lblPreviousBalance": "Previous Balance (AED)",
            "lblAdjustmentAmount": "Adjustment Amount (AED)",
            "PreviousBalance": 768.17,
            "PreviousBalanceUOM": "AED",
            "lblTotalAmountPayableWithinDueDate": "Total Amount Payable Within Due Date (AED)",
            "TotalAmountPayableWithinDueDate": 189521886.07,
            "TotalAmountPayableWithinDueDateUOM": "AED",
            "lblAdvancePayment": "Wallet Balance / Advance Payment (AED)",
            "AdvancePayment": 0,
            "AdvancePaymentUOM": "AED",
            "IsAdvancePayment": true,
            "AdjustmentAmount": 0,
            "AdjustmentAmountUOM": "AED",
            "IsAdjustmentAmount": true,
            "SumTaxableAmount": 180496302.77,
            "SumTaxableAmountUOM": "AED",
            "SumVAT": 9024815.13,
            "SumVATUOM": "AED",
            "SumTotalAmount": 189521117.9,
            "SumTotalAmountUOM": "AED",
            "TRhCorrectionFactor": 1,
            "RTHMeterDataValueUOM": "TRh",
            "EstimatedFlag": null,
            "IsInitialBill": false
        },
        {
            "SrNo": 3,
            "MeterID": null,
            "MeterSerialNumber": null,
            "CommercialDataTagID": 53,
            "Description": "Contract Capacity \nرسوم السعة",
            "PrevReading": null,
            "PrevReadingUOM": "kWh",
            "CurrReading": null,
            "CurrReadingUOM": "kWh",
            "Qty": 90.7,
            "QtyUOM": "TR",
            "Rate": 1.01916,
            "RateUOM": "AED",
            "TaxableAmount": 92.44,
            "TaxableAmountUOM": "AED",
            "VAT": 4.62,
            "VATUOM": "AED",
            "VATPercetage": 5,
            "VATPercetageUOM": "%",
            "TotalAmount": 97.06,
            "TotalAmountUOM": "AED",
            "lblPreviousBalance": "Previous Balance (AED)",
            "lblAdjustmentAmount": "Adjustment Amount (AED)",
            "PreviousBalance": 768.17,
            "PreviousBalanceUOM": "AED",
            "lblTotalAmountPayableWithinDueDate": "Total Amount Payable Within Due Date (AED)",
            "TotalAmountPayableWithinDueDate": 189521886.07,
            "TotalAmountPayableWithinDueDateUOM": "AED",
            "lblAdvancePayment": "Wallet Balance / Advance Payment (AED)",
            "AdvancePayment": 0,
            "AdvancePaymentUOM": "AED",
            "IsAdvancePayment": true,
            "AdjustmentAmount": 0,
            "AdjustmentAmountUOM": "AED",
            "IsAdjustmentAmount": true,
            "SumTaxableAmount": 180496302.77,
            "SumTaxableAmountUOM": "AED",
            "SumVAT": 9024815.13,
            "SumVATUOM": "AED",
            "SumTotalAmount": 189521117.9,
            "SumTotalAmountUOM": "AED",
            "TRhCorrectionFactor": 1,
            "RTHMeterDataValueUOM": "TRh",
            "EstimatedFlag": null,
            "IsInitialBill": false
        },
        {
            "SrNo": 4,
            "MeterID": null,
            "MeterSerialNumber": null,
            "CommercialDataTagID": 13,
            "Description": "Billing Fee\nرسوم الفوترة",
            "PrevReading": null,
            "PrevReadingUOM": "kWh",
            "CurrReading": null,
            "CurrReadingUOM": "kWh",
            "Qty": 1,
            "QtyUOM": "kWh",
            "Rate": 40,
            "RateUOM": "AED",
            "TaxableAmount": 40,
            "TaxableAmountUOM": "AED",
            "VAT": 2,
            "VATUOM": "AED",
            "VATPercetage": 5,
            "VATPercetageUOM": "%",
            "TotalAmount": 42,
            "TotalAmountUOM": "AED",
            "lblPreviousBalance": "Previous Balance (AED)",
            "lblAdjustmentAmount": "Adjustment Amount (AED)",
            "PreviousBalance": 768.17,
            "PreviousBalanceUOM": "AED",
            "lblTotalAmountPayableWithinDueDate": "Total Amount Payable Within Due Date (AED)",
            "TotalAmountPayableWithinDueDate": 189521886.07,
            "TotalAmountPayableWithinDueDateUOM": "AED",
            "lblAdvancePayment": "Wallet Balance / Advance Payment (AED)",
            "AdvancePayment": 0,
            "AdvancePaymentUOM": "AED",
            "IsAdvancePayment": true,
            "AdjustmentAmount": 0,
            "AdjustmentAmountUOM": "AED",
            "IsAdjustmentAmount": true,
            "SumTaxableAmount": 180496302.77,
            "SumTaxableAmountUOM": "AED",
            "SumVAT": 9024815.13,
            "SumVATUOM": "AED",
            "SumTotalAmount": 189521117.9,
            "SumTotalAmountUOM": "AED",
            "TRhCorrectionFactor": 1,
            "RTHMeterDataValueUOM": "TRh",
            "EstimatedFlag": null,
            "IsInitialBill": false
        }
    ],
    "labels": {
        "VAT": "VAT (AED)\nضريبة القيمة المضافة",
        "Qty": "Qty (kWh)\nالكمية",
        "AdvancePayment": "Advance Payment\nدفعة مقدمة",
        "TaxableAmount": "Taxable Amount (AED)\nالمبلغ الخاضع للضريبة",
        "PreviousBalance": "Previous Balance\nالرصيد السابق",
        "LeaveDoorsWindowsOpen": "Leave doors and windows open:\nترك الأبواب والنوافذ مفتوحة",
        "Estimate": "E = Estimated value due to missing or unreadable meter data.\nE = قيمة تقديرية بسبب عدم توفر أو عدم وضوح بيانات العداد",
        "PleaseDoNot": "Please Do Not:\nيرجى عدم",
        "PleaseTryTo": "Please Try to:\nيرجى محاولة",
        "PreviousReading": "Previous Reading (kWh)\nالقراءة السابقة",
        "BlockVents": "Block vents:\nإغلاق فتحات التهوية",
        "TOTALUOM": "AED\nAED",
        "AdjustmentAmount": "Adjustment Amount\nمبلغ التعديل",
        "KeepVentsUnblocked": "Keep vents unblocked:\nإبقاء فتحات التهوية غير مغلقة",
        "KeepFiltersClean": "Keep filters clean:\nالحفاظ على نظافة الفلاتر",
        "TOTAL": "TOTAL\nالإجمالي",
        "WayToPayDesc": "Cash , Cheque , Bank Transfer, Payby Link\nنقدًا، شيك، تحويل بنكي، رابط PayBy",
        "TotalAmountPayableWithinDueDate": "Total Amount Payable Within Due Date\nإجمالي المبلغ المستحق السداد قبل تاريخ الاستحقاق",
        "CurrentReading": "Current Reading (kWh)\nالقراءة الحالية",
        "WayToPay": "Ways to pay your Bill:\nطرق سداد الفاتورة",
        "TotalAmount": "Total Amount (AED)\nإجمالي المبلغ",
        "SetThermostatBetween23CAnd25C": "Set thermostat between 23°C and 25°C:\nضبط منظم الحرارة بين 23°م و25°م",
        "UndertakeRegularMaintenance": "Undertake regular maintenance of the system:\nإجراء صيانة دورية للنظام",
        "LeaveHeatProducing": "Leave heat-producing appliances near thermostats:\nوضع الأجهزة المنتجة للحرارة بالقرب من منظم الحرارة",
        "Description": "Description\nالبيان",
        "KeepDoorsWindowsClosed": "Keep doors and windows closed:\nإبقاء الأبواب والنوافذ مغلقة",
        "Note": "Note: A Late Fee will be applied to unpaid accounts after the invoice Due Date\nملاحظة: سيتم تطبيق غرامة تأخير على الحسابات غير المسددة بعد تاريخ استحقاق الفاتورة",
        "Rate": "Rate (AED/kWh)\nسعر الوحدة",
        "SetThermostatTo20CORLower": "Set thermostat to 20°C or lower:\nضبط منظم الحرارة على 20°م أو أقل",
        "27_SecurityDeposit": "مبلغ التأمين",
        "25_PreviousBalance": "الرصيد السابق",
        "44_Adjustment Amount": "مبلغ التعديل",
        "38_TotalAmountPayableWithin DueDate": "إجمالي المبلغ المستحق السداد قبل تاريخ الاستحقاق",
        "11_UtilityConsumption": ":استهلاك المرافق (الطاقة المبردة) (الرقم التسلسلي",
        "32_LatePaymentFee": "غرامة التأخير",
        "31_ActivationFee": "رسوم التفعيل",
        "13_BillingFee": "رسوم الفوترة",
        "0_TOTAL": "الإجمالي",
        "37_Advance Payment": "دفعة مقدمة",
        "64_FuelSurcharge": "رسوم الوقود الإضافية",
        "14_LateFee": "غرامة التأخير",
        "58_ReconnectionCharges": "رسوم إعادة التوصيل",
        "53_CapacityCharges": "رسوم السعة"
    }
}

const SAMPLE_CONSUMPTION = [
  {
    "Year": 2026,
    "Month": 3,
    "MonthName": "Mar",
    "ConsumptionValue": 850.72,
    "MeterTagName": "Consumption",
    "UOM": "kWh"
  },
  {
    "Year": 2026,
    "Month": 4,
    "MonthName": "Apr",
    "ConsumptionValue": 823.28,
    "MeterTagName": "Consumption",
    "UOM": "kWh"
  },
  {
    "Year": 2026,
    "Month": 5,
    "MonthName": "May",
    "ConsumptionValue": 999,
    "MeterTagName": "Consumption",
    "UOM": "kWh"
  },
  {
    "Year": 2026,
    "Month": 6,
    "MonthName": "June",
    "ConsumptionValue": 0,
    "MeterTagName": "Consumption",
    "UOM": "kWh"
  }
]

const BANK_DETAILS = [
    ["Account Title", "LYNX TECHNICAL SERVICES L.L.C"],
    ["Account Number", "19101139550"],
    ["IBAN", "AE440330000019101139550"],
    ["CIF Number", "14492376.00"],
    ["Bank Name", "Mashreq Bank PSC"],
];


const enLine = (t) => (t ? String(t).split("\n")[0].trim() : "");
const arLine = (t) => {
    if (!t) return "";
    const p = String(t).split("\n");
    return p.length > 1 ? p[1].trim() : "";
}; const fmtMoney = (v) => (v === null || v === undefined) ? "" : Number(v).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtDate = (v) => { if (!v) return ""; const d = new Date(v); return isNaN(d) ? v : d.toLocaleDateString("en-GB", { day: "2-digit", month: "2-digit", year: "numeric" }); };


function drawConsumptionChart(doc, data, x, y, width, height) {
    if (!Array.isArray(data) || data.length === 0) return y;

    // Increased left and bottom padding for axis labels
    const padding = { top: 6, right: 8, bottom: 12, left: 15 };

    const chartW = width - padding.left - padding.right;
    const chartH = height - padding.top - padding.bottom;
    const plotX = x + padding.left;
    const plotY = y + padding.top;

    const values = data.map((d) => Number(d.ConsumptionValue) || 0);
    const maxV = Math.max(...values, 1);
    const minV = Math.min(0, ...values);
    const range = (maxV - minV) || 1;

    // =========================
    // Outer Border
    // =========================
    doc.setDrawColor(0, 0, 0);
    doc.setLineWidth(0.2);
    doc.rect(x, y, width, height);

    // =========================
    // Horizontal Gridlines + Y Labels
    // =========================
    const gridLines = 4;

    doc.setFont("helvetica", "normal");
    doc.setFontSize(6.5);
    doc.setTextColor(90, 90, 90);

    for (let i = 0; i <= gridLines; i++) {
        const gy = plotY + chartH - (chartH * i) / gridLines;
        const val = minV + (range * i) / gridLines;

        // Light horizontal gridline
        doc.setDrawColor(224, 224, 224);
        doc.setLineWidth(0.15);
        doc.line(plotX, gy, plotX + chartW, gy);

        // Y-axis values
        doc.text(
            Number(val).toLocaleString(undefined, {
                maximumFractionDigits: 0,
            }),
            x + 2,
            gy + 1.4
        );
    }

    // =========================
    // Bold Axes
    // =========================
    doc.setDrawColor(0, 0, 0);
    doc.setLineWidth(0.3);

    // Y-axis
    doc.line(plotX, plotY, plotX, plotY + chartH);

    // X-axis
    doc.line(plotX, plotY + chartH, plotX + chartW, plotY + chartH);

    // =========================
    // Data Points
    // =========================
    const n = data.length;

    const points = data.map((d, i) => {
        const px =
            n === 1
                ? plotX + chartW / 2
                : plotX + (chartW / (n - 1)) * i;

        const py =
            plotY +
            chartH -
            (((Number(d.ConsumptionValue) || 0) - minV) / range) * chartH;

        return { px, py, d };
    });

    // =========================
    // Line Graph
    // =========================
    doc.setDrawColor(10, 61, 98);
    doc.setLineWidth(0.5);

    for (let i = 0; i < points.length - 1; i++) {
        doc.line(
            points[i].px,
            points[i].py,
            points[i + 1].px,
            points[i + 1].py
        );
    }

    // =========================
    // Data Point Markers
    // =========================
    doc.setFillColor(10, 61, 98);

    points.forEach((p) => {
        doc.circle(p.px, p.py, 0.9, "F");
    });

    // =========================
    // Month Labels
    // =========================
    doc.setFont("helvetica", "normal");
    doc.setFontSize(7);
    doc.setTextColor(0, 0, 0);

    points.forEach((p) => {
        doc.text(
            String(p.d.MonthName || ""),
            p.px,
            plotY + chartH + 4,
            { align: "center" }
        );
    });

    // =========================
    // Axis Labels
    // =========================
    doc.setFont("helvetica", "bold");
    doc.setFontSize(8);

    // X-axis label
    doc.text(
        "Months",
        plotX + chartW / 2,
        y + height - 2,
        { align: "center" }
    );

    // Y-axis label (vertical)
    doc.text(
        "Consumption (TRh)",
        x + 8,
        plotY + chartH / 2,
        {
            angle: 90,
            align: "center",
        }
    );

    // Reset font
    doc.setFont("helvetica", "normal");
    doc.setTextColor(0, 0, 0);

    return y + height;
}
// ─── Styles ───────────────────────────────────────────────────────────────────
const S = {
    page: { maxWidth:900, boxSizing: "border-box", fontFamily: "Arial, Helvetica, sans-serif", fontSize: "12px", color: "#000",margin: "0 auto",
 },
    toolbar: { display: "flex", justifyContent: "flex-end", gap: "10px", marginBottom: "14px" },
    btn: { display: "flex", alignItems: "center", gap: "6px", padding: "7px 16px", border: "1px solid #444", borderRadius: "4px", background: "#fff", fontSize: "13px", fontFamily: "Arial, Helvetica, sans-serif", cursor: "pointer" },
    report: { border: "1px solid #000", padding: "16px" },
    topRow: { display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: "14px" },
    companyBlock: { fontSize: "11px", fontWeight: "bold", lineHeight: 1.5, whiteSpace: "pre-line" },
    titleBlock: { textAlign: "center", flex: 1 },
    titleEn: { fontSize: "18px", fontWeight: "bold" },
    titleAr: { fontSize: "13px", fontWeight: "bold" },
    logoBox: { width: "90px", textAlign: "right", fontWeight: "bold", fontSize: "16px", color: "#0a3d62" },
    infoTable: { width: "100%", borderCollapse: "collapse", marginBottom: "14px" },
    infoCellLabel: { border: "1px solid #000", padding: "6px 8px", width: "16%", fontWeight: "bold", fontSize: "11px", verticalAlign: "top" },
    infoCellLabelAr: { fontWeight: "normal", fontSize: "10px", display: "block" },
    infoCellValue: { border: "1px solid #000", padding: "6px 8px", width: "34%", fontSize: "12px", verticalAlign: "top" },
    table: { width: "100%", borderCollapse: "collapse", marginBottom: "0px" },
    th: { border: "1px solid #000", padding: "6px 8px", fontSize: "11px", textAlign: "left", verticalAlign: "bottom" },
    thAr: { display: "block", fontWeight: "normal", fontSize: "10px" },
    td: { border: "1px solid #000", padding: "6px 8px", fontSize: "12px", verticalAlign: "top" },
    tdRight: { border: "1px solid #000", padding: "6px 8px", fontSize: "12px", textAlign: "right" },
    descCell: { lineHeight: 1.4 },
    descAr: { fontSize: "10px", color: "#333", display: "block" },
    summaryTable: { width: "100%", borderCollapse: "collapse" },
    summaryLabelCell: { border: "1px solid #000", borderTop: "none", padding: "6px 8px", fontWeight: "bold", fontSize: "12px" },
    summaryLabelAr: { display: "block", fontWeight: "normal", fontSize: "10px" },
    summaryValueCell: { border: "1px solid #000", borderTop: "none", padding: "6px 8px", textAlign: "right", fontWeight: "bold", fontSize: "12px", width: "16%" },

    // footer row 1: ways to pay + note
    footerRow: { display: "flex", gap: "16px", marginTop: "14px", fontSize: "11px" },
    footerCol: { flex: 1 },
    noteCol: { flex: 1 },
    bold: { fontWeight: "bold" },

    // second page section
    page2: { marginTop: "28px", paddingTop: "16px" },
    twoColRow: { display: "flex", gap: "20px", marginTop: "16px", alignItems: "flex-start" },
    halfCol: { flex: 1, fontSize: "11px" },

    // bank details
    bankTable: { width: "100%", borderCollapse: "collapse", fontSize: "11px", marginTop: "4px" },
    bankTh: { padding: "5px 8px", fontWeight: "bold", textAlign: "left", background: "#fff" },
    bankTd: { padding: "5px 8px", textAlign: "left" },

    // contact block
    contactBlock: { padding: "8px", fontSize: "11px", marginTop: "4px" },
    contactTitle: { fontWeight: "bold", marginBottom: "4px" },
    contactLine: { marginBottom: "3px", color: "#0a3d62" },
    contactPlain: { marginBottom: "3px" },

    // tips block
    tipsBlock: { padding: "10px", fontSize: "11px", marginTop: "4px" },
    tipsTitle: { fontWeight: "bold", marginBottom: "6px" },
    tipItem: { display: "flex", gap: "6px", marginBottom: "4px", alignItems: "flex-start" },

    statusMsg: (ok) => ({ marginTop: "10px", fontSize: "12px", color: ok ? "#16a34a" : "#b91c1c", textAlign: "right" }),

    // consumption chart
    chartSection: { marginTop: "20px", border: "1px solid #000", padding: "12px" },
    chartTitleRow: { display: "flex", alignItems: "baseline", gap: "8px", marginBottom: "8px" },
    chartTitleEn: { fontSize: "13px", fontWeight: "bold", color: "#0a3d62" },
    chartTitleAr: { fontSize: "11px", color: "#0a3d62" },
    chartUnit: { fontSize: "10px", color: "#555" },
};
// ─── Consumption line chart (HTML/SVG) ────────────────────────────────────────
function ConsumptionChart({ data = [] }) {
    if (!data || data.length === 0) {
        return <div style={{ fontSize: "11px", color: "#666", padding: "6px 0" }}>No consumption data available.</div>;
    }

    const width = 700, height = 220;
    const padding = { top: 16, right: 24, bottom: 42, left: 64 };

    const chartW = width - padding.left - padding.right;
    const chartH = height - padding.top - padding.bottom;

    const values = data.map((d) => Number(d.ConsumptionValue) || 0);
    const maxV = Math.max(...values, 1);
    const minV = Math.min(0, ...values);
    const range = (maxV - minV) || 1;

    const n = data.length;

    const points = data.map((d, i) => {
        const x =
            n === 1
                ? padding.left + chartW / 2
                : padding.left + (chartW / (n - 1)) * i;

        const y =
            padding.top +
            chartH -
            (((Number(d.ConsumptionValue) || 0) - minV) / range) * chartH;

        return { x, y, d };
    });

    const polylinePoints = points.map((p) => `${p.x},${p.y}`).join(" ");
    const gridCount = 4;

    const gridLines = Array.from({ length: gridCount + 1 }, (_, i) => {
        const gy = padding.top + chartH - (chartH * i) / gridCount;
        const val = minV + (range * i) / gridCount;
        return { y: gy, val };
    });

    const unit = data[0]?.UOM || "";

    return (
        <svg
            viewBox={`0 0 ${width} ${height}`}
            style={{ width: "100%", height: "auto", display: "block", fontFamily: "Arial, Helvetica, sans-serif" }}
        >
            <rect x={0.5} y={0.5} width={width - 1} height={height - 1} fill="none" stroke="#000" strokeWidth="1" />

            {/* Grid + Y-axis values */}
            {gridLines.map((g, i) => (
                <g key={i}>
                    <line x1={padding.left} y1={g.y} x2={width - padding.right} y2={g.y} stroke="#e0e0e0" strokeWidth="1" />
                    <text x={padding.left - 6} y={g.y + 3} textAnchor="end" fontSize="9" fill="#555">
                        {Math.round(g.val).toLocaleString()}
                    </text>
                </g>
            ))}

            {/* Axes */}
            <line x1={padding.left} y1={padding.top} x2={padding.left} y2={height - padding.bottom} stroke="#000" strokeWidth="1" />
            <line x1={padding.left} y1={height - padding.bottom} x2={width - padding.right} y2={height - padding.bottom} stroke="#000" strokeWidth="1" />

            {/* Line */}
            <polyline points={polylinePoints} fill="none" stroke="#0a3d62" strokeWidth="2" />

            {/* Data points + X labels */}
            {points.map((p, i) => (
                <g key={i}>
                    <circle cx={p.x} cy={p.y} r="3" fill="#0a3d62">
                        <title>{`${p.d.MonthName} ${p.d.Year}: ${Number(p.d.ConsumptionValue).toLocaleString()} ${unit}`}</title>
                    </circle>
                    <text x={p.x} y={height - padding.bottom + 14} textAnchor="middle" fontSize="9" fill="#000">
                        {p.d.MonthName}
                    </text>
                </g>
            ))}

            {/* Axis Labels */}
            {/* X-axis label */}
            <text
                x={padding.left + chartW / 2}
                y={height - 4}
                textAnchor="middle"
                fontSize="10"
                fill="#333"
            >
                Months
            </text>

            {/* Y-axis label */}
            <text
                x={14}
                y={padding.top + chartH / 2}
                textAnchor="middle"
                fontSize="10"
                fill="#333"
                transform={`rotate(-90 14 ${padding.top + chartH / 2})`}
            >
                Consumption (TRh)
            </text>

            {/* Unit label */}
            <text x={4} y={padding.top - 4} fontSize="9" fill="#555">
                {unit ? `(${unit})` : ""}
            </text>
        </svg>
    );
}

// ─── Main Component ───────────────────────────────────────────────────────────
export default function TaxInvoiceReport({
    headerResponse = SAMPLE_HEADER,
    detailResponse = SAMPLE_DETAIL,
    // Consumption history comes from a separate API, e.g.:
    // [{ Year, Month, MonthName, ConsumptionValue, MeterTagName, UOM }, ...]
    consumptionResponse = SAMPLE_CONSUMPTION,
}) {
    const fields = new Fields(headerResponse, detailResponse);
    const [status, setStatus] = useState("");
    const [loadingPDF, setLoadingPDF] = useState(false);
    const [loadingExcel, setLoadingExcel] = useState(false);

    const consumption = Array.isArray(consumptionResponse) ? consumptionResponse : [];

    const details = fields.get("Details") || [];
    const lastRow = details[details.length - 1] || {};
    const sumTaxable = lastRow.SumTaxableAmount ?? details.reduce((s, r) => s + (r.TaxableAmount || 0), 0);
    const sumVAT = lastRow.SumVAT ?? details.reduce((s, r) => s + (r.VAT || 0), 0);
    const sumTotal = lastRow.SumTotalAmount ?? details.reduce((s, r) => s + (r.TotalAmount || 0), 0);
    const previousBalance = lastRow.PreviousBalance ?? 0;
    const adjustmentAmount = lastRow.AdjustmentAmount ?? 0;
    const payableWithinDueDate = lastRow.TotalAmountPayableWithinDueDate ?? sumTotal;

    const tryToItems = ["KeepFiltersClean", "SetThermostatBetween23CAnd25C", "KeepVentsUnblocked", "KeepDoorsWindowsClosed", "UndertakeRegularMaintenance"];
    const doNotItems = ["SetThermostatTo20CORLower", "BlockVents", "LeaveDoorsWindowsOpen", "LeaveHeatProducing"];

    // ── Excel Export ──────────────────────────────────────────────────────────────

    function exportExcel() {
        setLoadingExcel(true); setStatus("Generating Excel…");
        try {
            const wb = XLSX.utils.book_new();

            const infoRows = [
                [enLine(fields.getLabel("Name")), fields.get("ConsumerName"), enLine(fields.getLabel("AccountNumber")), fields.get("AccountNumber")],
                [enLine(fields.getLabel("UnitNo")), fields.get("UnitNo"), enLine(fields.getLabel("BillingDate")), fmtDate(fields.get("BillingDate"))],
                [enLine(fields.getLabel("TowerName")), fields.get("BuildingName"), enLine(fields.getLabel("DueDate")), fmtDate(fields.get("DueDate"))],
                [enLine(fields.getLabel("InvoiceNumber")), fields.get("InvoiceNumber"), enLine(fields.getLabel("BillingPeriod")), fields.get("BillingPeriod")],
            ];

            const detailHeader = ["Description", "Previous Reading", "Current Reading", "Qty", "Rate", "Taxable Amount", "VAT", "Total Amount"];
            const detailRows = details.map((r) => [
                enLine(r.Description), r.PrevReading ?? "", r.CurrReading ?? "",
                r.Qty ?? "", r.Rate ?? "", r.TaxableAmount ?? "", r.VAT ?? "", r.TotalAmount ?? "",
            ]);
            const totalsRow = ["", "", "", "", "TOTAL", sumTaxable, sumVAT, sumTotal];

            const summaryRows = [
                [enLine(fields.getLabel("PreviousBalance")), previousBalance],
                [enLine(fields.getLabel("AdjustmentAmount")), adjustmentAmount],
                [enLine(fields.getLabel("TotalAmountPayableWithinDueDate")), payableWithinDueDate],
            ];

            const sheetData = [
                [enLine(fields.getLabel("TaxInvoice")) || "TAX INVOICE"],
                [],
                ...infoRows,
                [],
                detailHeader,
                ...detailRows,
                totalsRow,
                [],
                ...summaryRows,
            ];

            const ws = XLSX.utils.aoa_to_sheet(sheetData);
            ws["!cols"] = [{ wch: 28 }, { wch: 20 }, { wch: 24 }, { wch: 20 }, { wch: 14 }, { wch: 14 }, { wch: 12 }, { wch: 14 }];
            XLSX.utils.book_append_sheet(wb, ws, "Tax Invoice");

            // Bank details on a second sheet for completeness
            const wsBank = XLSX.utils.aoa_to_sheet([["Bank Details"], [], ...BANK_DETAILS]);
            wsBank["!cols"] = [{ wch: 20 }, { wch: 30 }];
            XLSX.utils.book_append_sheet(wb, wsBank, "Bank Details");

            XLSX.writeFile(wb, `${fields.get("InvoiceNumber") || "tax_invoice"}.xlsx`);
            setStatus("\u2713 Excel downloaded.");
        } catch (err) { console.error(err); setStatus("Excel export failed."); }
        finally { setLoadingExcel(false); }
    }

    // ── PDF Export ────────────────────────────────────────────────────────────────
    async function exportPDF() {
        setLoadingPDF(true); setStatus("Loading Arabic font…");
        try {

            let amiriB64 = null;
            try {
                const fontResp = await fetch('/data/Amiri Regular.ttf');
                if (!fontResp.ok) throw new Error(`Font fetch failed: ${fontResp.status}`);
                const fontBuf = await fontResp.arrayBuffer();
                const bytes = new Uint8Array(fontBuf);
                let binary = "";
                for (let i = 0; i < bytes.byteLength; i++) binary += String.fromCharCode(bytes[i]);
                amiriB64 = btoa(binary);
            } catch (fontErr) {
                console.warn("Amiri font load failed, PDF will use latin-only font:", fontErr);
            }

            setStatus("Generating PDF…");
            const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });

            if (amiriB64) {
                doc.addFileToVFS("Amiri-Regular.ttf", amiriB64);
                doc.addFont("Amiri-Regular.ttf", "Amiri", "bold");
            }

            const setArabic = () => amiriB64 ? doc.setFont("Amiri", "bold") : doc.setFont("helvetica", "normal");
            const setLatin = (style = "normal") => doc.setFont("helvetica", style);



            const drawArSubline = (text, x, y, { size = 7.5, color = [10, 61, 98] } = {}) => {
                if (!amiriB64 || !text) return;
                doc.setFont("Amiri", "bold"); doc.setFontSize(size); doc.setTextColor(...color);
                doc.text(text, x, y, { align: "left", valign: 'top' });
                doc.setTextColor(0, 0, 0);
            };

            // ── Page 1 ─────────────────────────────────────────────────────────────
            setLatin("bold"); doc.setFontSize(10); doc.setTextColor(10, 61, 98);
            (fields.get("OperatingCompanyAddress") || "").split("\n").forEach((l, i) => doc.text(l, 14, 16 + i * 5));

            setLatin("bold"); doc.setFontSize(15);

            doc.text("TAX INVOICE", 105, 16, { align: "center" });

            setArabic(); doc.setFontSize(12);
            doc.text("\u0641\u0627\u062a\u0648\u0631\u0629 \u0636\u0631\u064a\u0628\u064a\u0629", 105, 23, { align: "center" });

            const logoUrl = "/output.png";

            function getBase64Image(url) {
                return new Promise((resolve) => {
                    const img = new Image();
                    img.crossOrigin = "anonymous";
                    img.src = url;

                    img.onload = () => {
                        const canvas = document.createElement("canvas");
                        canvas.width = img.width;
                        canvas.height = img.height;

                        const ctx = canvas.getContext("2d");
                        ctx.drawImage(img, 0, 0);

                        resolve(canvas.toDataURL("image/png"));
                    };
                });
            }

            const logoBase64 = await getBase64Image(logoUrl);

            doc.addImage(logoBase64, "PNG", 160, 8, 40, 12);


            const infoRows = [
                [fields.getLabel("Name"), fields.get("ConsumerName"), fields.getLabel("AccountNumber"), fields.get("AccountNumber")],
                [fields.getLabel("UnitNo"), fields.get("UnitNo"), fields.getLabel("BillingDate"), fmtDate(fields.get("BillingDate"))],
                [fields.getLabel("TowerName"), fields.get("BuildingName"), fields.getLabel("DueDate"), fmtDate(fields.get("DueDate"))],
                [fields.getLabel("InvoiceNumber"), fields.get("InvoiceNumber"), fields.getLabel("BillingPeriod"), fields.get("BillingPeriod")],
            ];


            const leftRows = infoRows.map(([l1, v1]) => [enLine(l1), v1]);
            const rightRows = infoRows.map(([, , l2, v2]) => [enLine(l2), v2]);


            const startY = 28;
            const gap = 101;

            const leftX = 30;
            const rightX = leftX + gap;


            autoTable(doc, {
                body: leftRows,
                startY: startY,
                theme: "grid",
                styles: {
                    fontSize: 8.5,
                    cellPadding: 3,
                    lineColor: [0, 0, 0],
                    lineWidth: 0.2,
                    textColor: [10, 61, 98],
                },
                columnStyles: {
                    0: { fontStyle: "bold", cellWidth: 28 },
                    1: { cellWidth: 42 },
                },
                didDrawCell(data) {
                    if (!amiriB64) return;
                    const isLabel = data.column.index === 0 || data.column.index === 2;
                    if (data.section === "body" && isLabel) {
                        const src = data.column.index === 0 ? infoRows[data.row.index][0] : infoRows[data.row.index][2];
                        const ar = arLine(src);
                        if (ar) drawArSubline(ar, data.cell.x + 3, data.cell.y + data.cell.height - 1.5, { size: 7.5 });
                    }
                },
            });

            autoTable(doc, {
                body: rightRows,
                startY: startY,
                margin: { left: rightX },
                theme: "grid",
                styles: {
                    fontSize: 8.5,
                    cellPadding: 3,
                    lineColor: [0, 0, 0],
                    lineWidth: 0.2,
                    textColor: [10, 61, 98],
                },
                columnStyles: {
                    0: { fontStyle: "bold", cellWidth: 28 },
                    1: { cellWidth: 42 },
                },
                didDrawCell(data) {
                    if (!amiriB64) return;
                    const isLabel = data.column.index === 0 || data.column.index === 2;
                    if (data.section === "body" && isLabel) {
                        const src = data.column.index === 0 ? infoRows[data.row.index][0] : infoRows[data.row.index][2];
                        const ar = arLine(src);
                        if (ar) drawArSubline(ar, data.cell.x + 3, data.cell.y + data.cell.height - 1.5, { size: 7.5 });
                    }
                },
            });
            const detailHeaders = [
                [enLine(fields.getLabel("Description")), arLine(fields.getLabel("Description"))],
                [enLine(fields.getLabel("PreviousReading")), arLine(fields.getLabel("PreviousReading"))],
                [enLine(fields.getLabel("CurrentReading")), arLine(fields.getLabel("CurrentReading"))],
                [enLine(fields.getLabel("Qty")), arLine(fields.getLabel("Qty"))],
                [enLine(fields.getLabel("Rate")), arLine(fields.getLabel("Rate"))],
                [enLine(fields.getLabel("TaxableAmount")), arLine(fields.getLabel("TaxableAmount"))],
                [enLine(fields.getLabel("VAT")), arLine(fields.getLabel("VAT"))],
                [enLine(fields.getLabel("TotalAmount")), arLine(fields.getLabel("TotalAmount"))],
            ];

            autoTable(doc, {
                head: [detailHeaders.map(([en]) => en)],
                body: details.map((r) => [
                    enLine(r.Description),
                    r.PrevReading ?? "-", r.CurrReading ?? "-",
                    fmtMoney(r.Qty), fmtMoney(r.Rate), fmtMoney(r.TaxableAmount), fmtMoney(r.VAT), fmtMoney(r.TotalAmount),
                ]),
                foot: [["TOTAL", "", "", "", "", fmtMoney(sumTaxable), fmtMoney(sumVAT), fmtMoney(sumTotal)]],
                startY: doc.lastAutoTable.finalY + 2,
                styles: { fontSize: 8, cellPadding: 1.5, lineColor: [0, 0, 0], lineWidth: 0.2, textColor: [0, 0, 0], overflow: 'linebreak' },
                headStyles: { fillColor: [255, 255, 255], textColor: [10, 61, 98], fontStyle: "bold", lineColor: [0, 0, 0], lineWidth: 0.2, halign: "left" },
                footStyles: { fillColor: [255, 255, 255], textColor: [10, 61, 98], fontStyle: "bold", lineColor: [0, 0, 0], lineWidth: 0.2, halign: "left" },
                columnStyles: {
                    0: { cellWidth: 52, textColor: [10, 61, 98], }, 1: { cellWidth: 16, halign: "right" }, 2: { cellWidth: 16, halign: "right" },
                    3: { cellWidth: 17, halign: "right" }, 4: { cellWidth: 20, halign: "right" },
                    5: { cellWidth: 25, halign: "right" }, 6: { cellWidth: 16, halign: "right" }, 7: { cellWidth: 25, halign: "right" },
                },
                didDrawCell(data) {
                    if (!amiriB64) return;
                    // Arabic sub-label in column headers — S.th is left-aligned in the HTML
                    // (S.thAr is just a block span under it), so keep these left-aligned too.
                    if (data.section === "head" && data.row.index === 0) {
                        const ar = detailHeaders[data.column.index]?.[1];
                        if (ar) {
                            drawArSubline(ar, data.cell.x + 2, data.cell.y + data.cell.height - 1, { size: 6.5 });
                        }
                    }
                    // Arabic description sub-line in body — S.descAr is a left-aligned block
                    // under the English description, not right-aligned.
                    if (data.section === "body" && data.column.index === 0) {
                        const ar = arLine(details[data.row.index]?.Description);
                        if (ar) drawArSubline(ar, data.cell.x + 2, data.cell.y + data.cell.height - 1.4, { size: 6.5, color: [51, 51, 51] });
                    }
                },
                didParseCell: function (data) {
                    if (data.section === "body" && data.column.index === 0) {
                        data.cell.styles.fontSize = 8;
                        data.cell.styles.minCellHeight = 10
                        data.cell.styles.cellPadding = {
                            top: 1,
                            right: 1,
                            bottom: 2,
                            left: 1.5,
                        };
                    }
                    if (data.section === "head" && (data.column.index === 1 || data.column.index === 2)) {
                        data.cell.styles.fontSize = 7;
                    }
                }

            });


            const summaryData = [
                [enLine(fields.getLabel("PreviousBalance")), arLine(fields.getLabel("PreviousBalance")), fmtMoney(previousBalance)],
                [enLine(fields.getLabel("AdjustmentAmount")), arLine(fields.getLabel("AdjustmentAmount")), fmtMoney(adjustmentAmount)],
                [enLine(fields.getLabel("TotalAmountPayableWithinDueDate")), arLine(fields.getLabel("TotalAmountPayableWithinDueDate")), fmtMoney(payableWithinDueDate)],
            ];

            autoTable(doc, {
                body: summaryData.map(([en, , val]) => [en, val]),
                startY: doc.lastAutoTable.finalY,
                theme: "grid",
                styles: { fontSize: 9, cellPadding: 3, lineColor: [0, 0, 0], lineWidth: 0.2, fontStyle: "bold", textColor: [10, 61, 98], },
                columnStyles: { 0: { cellWidth: 162, textColor: [10, 61, 98], }, 1: { cellWidth: 25, halign: "right", valign: 'bottom' } },
                didDrawCell(data) {
                    if (!amiriB64) return;
                    if (data.section === "body" && data.column.index === 0) {
                        const ar = summaryData[data.row.index]?.[1];
                        if (ar) drawArSubline(ar, data.cell.x + 2, data.cell.y + data.cell.height - 1.5, { size: 7.5 });
                    }
                },
            });

            // Ways to pay + Note
            let y = doc.lastAutoTable.finalY + 6;
            setLatin("bold"); doc.setFontSize(9); doc.setTextColor(10, 61, 98);
            doc.text(enLine(fields.getLabel("WayToPay")), 14, y);

            setArabic(); doc.setFontSize(8);
            doc.text(arLine(fields.getLabel("WayToPay")), 14, y + 5);
            setLatin("normal"); doc.setFontSize(8.5);
            doc.text(enLine(fields.getLabel("WayToPayDesc")), 48, y);
            setArabic(); doc.setFontSize(8);
            doc.text(arLine(fields.getLabel("WayToPayDesc")), 48, y + 5);
            setLatin("normal"); doc.setFontSize(8);
            doc.text(enLine(fields.getLabel("Note")), 110, y, { maxWidth: 88 });
            setArabic(); doc.setFontSize(7.5);
            doc.text(arLine(fields.getLabel("Note")), 110, y + 5, { maxWidth: 88 });
            setLatin("normal"); doc.setFontSize(7.5);
            doc.text(enLine(fields.getLabel("Estimate")), 110, y + 12, { maxWidth: 88 });
            setArabic(); doc.setFontSize(7.5);
            doc.text(arLine(fields.getLabel("Estimate")), 110, y + 17, { maxWidth: 88 });

            // ── Page 2 ─────────────────────────────────────────────────────────────
            y = y + 30;

            autoTable(doc, {
                head: [["Bank Details:", ""]],
                body: BANK_DETAILS,
                startY: y, theme: "grid", tableWidth: 90,
                styles: { fontSize: 9, cellPadding: 2.5, lineColor: [0, 0, 0], lineWidth: 0, textColor: [10, 61, 98], },
                headStyles: { fillColor: [255, 255, 255], textColor: [10, 61, 98], fontStyle: "bold", lineColor: [0, 0, 0], lineWidth: 0 },
                columnStyles: { 0: { fontStyle: "bold", cellWidth: 36 }, 1: { cellWidth: 54 } },
            });

            const contactY = y;
            setLatin("bold"); doc.setFontSize(9);
            doc.text("Contact Informations:", 127, contactY + 4);
            setLatin("normal"); doc.setFontSize(8.5);
            ["Enquiries@lynxllc.ae", "info@lynxllc.ae",
                fields.get("ContactNumber") || "", fields.get("WhatsAppNumber") || ""].forEach((line, i) => {
                    doc.text(line, 127, contactY + 10 + i * 6);
                });

            y = Math.max(doc.lastAutoTable.finalY, contactY + 44) + 10 + 32;

            const tryRows = tryToItems.map((k) => [enLine(fields.getLabel(k)), arLine(fields.getLabel(k))]);
            const doNotRows = doNotItems.map((k) => [enLine(fields.getLabel(k)), arLine(fields.getLabel(k))]);

            const makeTipTable = (headLabel, subLabel, rows, startY, marginLeft) => {
                autoTable(doc, {
                    head: [[headLabel]],
                    body: rows.map(([en]) => [`\u2022 ${en}`]),
                    startY, tableWidth: 90,
                    ...(marginLeft !== undefined ? { margin: { left: marginLeft } } : {}),
                    theme: "grid",
                    styles: { fontSize: 8.5, cellPadding: 2.5, lineColor: [0, 0, 0], lineWidth: 0, textColor: [10, 61, 98], },
                    headStyles: { fillColor: [255, 255, 255], textColor: [10, 61, 98], fontStyle: "bold", lineColor: [0, 0, 0], lineWidth: 0 },
                    didDrawCell(data) {
                        if (!amiriB64) return;
                        if (data.section === "body") {
                            const ar = rows[data.row.index]?.[1];
                            if (ar) drawArSubline(ar, data.cell.x + 2, data.cell.y + data.cell.height - 1, { size: 7 });
                        }
                        if (data.section === "head") {
                            const ar = subLabel;
                            if (ar) drawArSubline(ar, data.cell.x + 2, data.cell.y + data.cell.height - 1, { size: 7 });
                        }

                    },
                });
                return doc.lastAutoTable.finalY;
            };
            doc.addPage();
            y = 16;

            const tryFinalY = makeTipTable(`${enLine(fields.getLabel("PleaseTryTo"))}`, `${arLine(fields.getLabel("PleaseTryTo"))}`, tryRows, y, undefined);
            const doNotFinalY = makeTipTable(`${enLine(fields.getLabel("PleaseDoNot"))}`, `${arLine(fields.getLabel("PleaseDoNot"))}`, doNotRows, y, 107);

            // ── Consumption History Chart ─────────────────────────────────────────
            let chartY = Math.max(tryFinalY, doNotFinalY) + 12;
            const chartHeight = 65;
            const chartWidth = 182;
            if (chartY + 10 + chartHeight > 280) { doc.addPage(); chartY = 16; }

            setLatin("bold"); doc.setFontSize(10); doc.setTextColor(10, 61, 98);
            doc.text("Monthly Consumption Data", 14, chartY);
            drawConsumptionChart(doc, consumption, 14, chartY + 5, chartWidth, chartHeight);

            doc.save(`${fields.get("InvoiceNumber") || "tax_invoice"}.pdf`);
            setStatus("\u2713 PDF downloaded.");
        } catch (err) { console.error(err); setStatus("PDF export failed."); }
        finally { setLoadingPDF(false); }
    }

    return (
        <div style={S.page}>
            <div style={S.toolbar}>
                <button style={S.btn} onClick={exportExcel} disabled={loadingExcel}>⬇ {loadingExcel ? "Exporting…" : "Export Excel"}</button>
                <button style={S.btn} onClick={exportPDF} disabled={loadingPDF}>⬇ {loadingPDF ? "Exporting…" : "Export PDF"}</button>
            </div>

            <div style={S.report}>

                {/* ── Page 1: Header ── */}
                <div style={S.topRow}>
                    <div style={S.companyBlock}>{fields.get("OperatingCompanyAddress")}</div>
                    <div style={S.titleBlock}>
                        <div style={S.titleEn}>{enLine(fields.getLabel("TaxInvoice")) || "TAX INVOICE"}</div>
                        <div style={S.titleAr}>{arLine(fields.getLabel("TaxInvoice")) || "فاتورة ضريبية"}</div>
                    </div>
                    <div style={S.logoBox}>LYNX</div>
                </div>

                {/* Info grid */}
                {/* <table style={S.infoTable}>
                    <tbody>
                        {[
                            [fields.getLabel("Name"), fields.get("ConsumerName"), fields.getLabel("AccountNumber"), fields.get("AccountNumber")],
                            [fields.getLabel("UnitNo"), fields.get("UnitNo"), fields.getLabel("BillingDate"), fmtDate(fields.get("BillingDate"))],
                            [fields.getLabel("TowerName"), fields.get("BuildingName"), fields.getLabel("DueDate"), fmtDate(fields.get("DueDate"))],
                            [fields.getLabel("InvoiceNumber"), fields.get("InvoiceNumber"), fields.getLabel("BillingPeriod"), fields.get("BillingPeriod")],
                        ].map((row, i) => (
                            <tr key={i}>
                                <td style={S.infoCellLabel}>{enLine(row[0])}<span style={S.infoCellLabelAr}>{arLine(row[0])}</span></td>
                                <td style={S.infoCellValue}>{row[1]}</td>
                                <td style={S.infoCellLabel}>{enLine(row[2])}<span style={S.infoCellLabelAr}>{arLine(row[2])}</span></td>
                                <td style={S.infoCellValue}>{row[3]}</td>
                            </tr>
                        ))}
                    </tbody>
                </table> */}

                <div style={{ display: "flex", gap: "20px", marginTop: "10px" }}>

                    {/* LEFT TABLE */}
                    <table style={{ ...S.infoTable, width: "50%", marginRight: "30px" }}>
                        <tbody>
                            {[
                                [fields.getLabel("Name"), fields.get("ConsumerName")],
                                [fields.getLabel("UnitNo"), fields.get("UnitNo")],
                                [fields.getLabel("TowerName"), fields.get("BuildingName")],
                                [fields.getLabel("InvoiceNumber"), fields.get("InvoiceNumber")],
                            ].map((row, i) => (
                                <tr key={i}>
                                    <td style={S.infoCellLabel}>
                                        {enLine(row[0])}
                                        <span style={S.infoCellLabelAr}>{arLine(row[0])}</span>
                                    </td>
                                    <td style={S.infoCellValue}>{row[1]}</td>
                                </tr>
                            ))}
                        </tbody>
                    </table>

                    {/* RIGHT TABLE */}
                    <table style={{ ...S.infoTable, width: "50%", marginLeft: "40px" }}>
                        <tbody>
                            {[
                                [fields.getLabel("AccountNumber"), fields.get("AccountNumber")],
                                [fields.getLabel("BillingDate"), fmtDate(fields.get("BillingDate"))],
                                [fields.getLabel("DueDate"), fmtDate(fields.get("DueDate"))],
                                [fields.getLabel("BillingPeriod"), fields.get("BillingPeriod")],
                            ].map((row, i) => (
                                <tr key={i}>
                                    <td style={S.infoCellLabel}>
                                        {enLine(row[0])}
                                        <span style={S.infoCellLabelAr}>{arLine(row[0])}</span>
                                    </td>
                                    <td style={S.infoCellValue}>{row[1]}</td>
                                </tr>
                            ))}
                        </tbody>
                    </table>

                </div>

                {/* Detail table */}
                <table style={S.table}>
                    <thead>
                        <tr>
                            {["Description", "PreviousReading", "CurrentReading", "Qty", "Rate", "TaxableAmount", "VAT", "TotalAmount"].map((k) => (
                                <th key={k} style={S.th}>
                                    {enLine(fields.getLabel(k))}
                                    <span style={S.thAr}>{arLine(fields.getLabel(k))}</span>
                                </th>
                            ))}
                        </tr>
                    </thead>
                    <tbody>
                        {details.map((row) => (
                            <tr key={row.SrNo}>
                                <td style={{ ...S.td, ...S.descCell }}>
                                    {enLine(row.Description)}<span style={S.descAr}>{arLine(row.Description)}</span>
                                </td>
                                <td style={S.tdRight}>{row.PrevReading ?? ""}</td>
                                <td style={S.tdRight}>{row.CurrReading ?? ""}</td>
                                <td style={S.tdRight}>{fmtMoney(row.Qty)}</td>
                                <td style={S.tdRight}>{fmtMoney(row.Rate)}</td>
                                <td style={S.tdRight}>{fmtMoney(row.TaxableAmount)}</td>
                                <td style={S.tdRight}>{fmtMoney(row.VAT)}</td>
                                <td style={S.tdRight}>{fmtMoney(row.TotalAmount)}</td>
                            </tr>
                        ))}
                        <tr>
                            <td colSpan={5} style={{ ...S.td, fontWeight: "bold" }}>
                                {enLine(fields.getLabel("TOTAL"))}<span style={S.thAr}>{arLine(fields.getLabel("TOTAL"))}</span>
                            </td>
                            <td style={{ ...S.tdRight, fontWeight: "bold" }}>{fmtMoney(sumTaxable)}</td>
                            <td style={{ ...S.tdRight, fontWeight: "bold" }}>{fmtMoney(sumVAT)}</td>
                            <td style={{ ...S.tdRight, fontWeight: "bold" }}>{fmtMoney(sumTotal)}</td>
                        </tr>
                    </tbody>
                </table>

                {/* Summary */}
                <table style={S.summaryTable}>
                    <tbody>
                        {[
                            ["PreviousBalance", fmtMoney(previousBalance)],
                            ["AdjustmentAmount", fmtMoney(adjustmentAmount)],
                            ["TotalAmountPayableWithinDueDate", fmtMoney(payableWithinDueDate)],
                        ].map(([key, val]) => (
                            <tr key={key}>
                                <td style={S.summaryLabelCell} colSpan={7}>
                                    {enLine(fields.getLabel(key))}<span style={S.summaryLabelAr}>{arLine(fields.getLabel(key))}</span>
                                </td>
                                <td style={S.summaryValueCell}>{val}</td>
                            </tr>
                        ))}
                    </tbody>
                </table>

                {/* Footer row 1: Ways to pay + Note */}
                <div style={S.footerRow}>
                    <div style={S.footerCol}>
                        <div style={S.bold}>{enLine(fields.getLabel("WayToPay"))}</div>
                        <div style={{ fontSize: "11px", color: "#555" }}>{arLine(fields.getLabel("WayToPay"))}</div>
                        <div style={{ marginTop: "4px" }}>{enLine(fields.getLabel("WayToPayDesc"))}</div>
                    </div>
                    <div style={S.noteCol}>
                        <div>{enLine(fields.getLabel("Note"))}</div>
                        <div style={{ marginTop: "6px", fontSize: "11px", color: "#444" }}>{enLine(fields.getLabel("Estimate"))}</div>
                    </div>
                </div>

                {/* ── Page 2 section ── */}
                <div style={S.page2}>

                    {/* Bank Details (left) + Contact (right) */}
                    <div style={S.twoColRow}>

                        {/* Bank Details */}
                        <div style={S.halfCol}>
                            <table style={S.bankTable}>
                                <tbody>
                                    <tr>
                                        <th colSpan={2} style={S.bankTh}>Bank Details:</th>
                                    </tr>
                                    {BANK_DETAILS.map(([label, value]) => (
                                        <tr key={label}>
                                            <td style={{ ...S.bankTd, fontWeight: "bold", width: "40%" }}>{label}</td>
                                            <td style={S.bankTd}>{value}</td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>

                        {/* Contact Information */}
                        <div style={S.halfCol}>
                            <div style={S.contactBlock}>
                                <div style={S.contactTitle}>Contact Informations:</div>
                                <div style={S.contactLine}>Enquiries@lynxllc.ae</div>
                                <div style={S.contactLine}>info@lynxllc.ae</div>
                                <div style={S.contactPlain}>{fields.get("ContactNumber")}</div>
                                <div style={S.contactPlain}>{fields.get("WhatsAppNumber")}</div>
                            </div>
                        </div>
                    </div>

                    {/* Please Try To (left) + Please Do Not (right) */}
                    <div style={S.twoColRow}>

                        <div style={S.halfCol}>
                            <div style={S.tipsBlock}>
                                <div style={S.tipsTitle}>
                                    {enLine(fields.getLabel("PleaseTryTo")) || "Please Try to:"}
                                    <span style={{ fontWeight: "normal", marginLeft: "6px", fontSize: "10px" }}>{arLine(fields.getLabel("PleaseTryTo"))}</span>
                                </div>
                                {tryToItems.map((k) => (
                                    <div key={k} style={S.tipItem}>
                                        <span>•</span>
                                        <span>
                                            {enLine(fields.getLabel(k))}
                                            <span style={{ fontSize: "10px", color: "#555", display: "block" }}>{arLine(fields.getLabel(k))}</span>
                                        </span>
                                    </div>
                                ))}
                            </div>
                        </div>

                        <div style={S.halfCol}>
                            <div style={S.tipsBlock}>
                                <div style={S.tipsTitle}>
                                    {enLine(fields.getLabel("PleaseDoNot")) || "Please Do Not:"}
                                    <span style={{ fontWeight: "normal", marginLeft: "6px", fontSize: "10px" }}>{arLine(fields.getLabel("PleaseDoNot"))}</span>
                                </div>
                                {doNotItems.map((k) => (
                                    <div key={k} style={S.tipItem}>
                                        <span>•</span>
                                        <span>
                                            {enLine(fields.getLabel(k))}
                                            <span style={{ fontSize: "10px", color: "#555", display: "block" }}>{arLine(fields.getLabel(k))}</span>
                                        </span>
                                    </div>
                                ))}
                            </div>
                        </div>

                    </div>

                    {/* Consumption History Chart */}
                    <div style={S.chartSection}>
                        <div style={S.chartTitleRow}>
                            <span style={S.chartTitleEn}>Consumption Monthly data</span>

                        </div>
                        <ConsumptionChart data={consumption} />
                    </div>
                </div>

            </div>

            {status && <div style={S.statusMsg(status.startsWith("✓"))}>{status}</div>}
        </div>
    );
}