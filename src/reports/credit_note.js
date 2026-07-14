import { useState } from "react";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";

async function loadData(path) {
    try {
        const response = await fetch(path);
        if (!response.ok) throw new Error('Network response was not ok');
        const data = await response.json();
        return data;
    } catch (error) {
        console.error('Error reading JSON:', error);
        return null;
    }
}

const SAMPLE_DATA = await loadData('/data/creditnote.json');
const SAMPLE_HEADER = await loadData('/data/creditnote_header.json');
const getLineAmount = (row) => row?.["TotalAmount(AED)"] ?? row?.TotalAmount ?? 0;

const LABELS = {
    documentTitle: "TAX INVOICE - CREDIT NOTE",
    name: "Name",
    unitNo: "Unit No.",
    towerName: "Tower Name",
    accountNumber: "Account Number",
    billingDate: "Billing Date",
    invoiceNumber: "Invoice Number",
    billingPeriod: "Billing Period",
};

const BANK_DETAILS = [
    ["Account Title", "LYNX TECHNICAL SERVICES L.L.C"],
    ["Account Number", "19101139550"],
    ["IBAN", "AE440330000019101139550"],
    ["CIF Number", "14492376.00"],
    ["Bank Name", "Mashreq Bank PSC"],
];

const fmtMoney = (v) => (v === null || v === undefined || v === "") ? "" : Number(v).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtDate = (v) => { if (!v) return ""; const d = new Date(v); return isNaN(d) ? v : d.toLocaleDateString("en-GB", { day: "2-digit", month: "2-digit", year: "numeric" }); };

// ─── Styles ───────────────────────────────────────────────────────────────────
const S = {
    page: { maxWidth: 900, boxSizing: "border-box", fontFamily: "Arial, Helvetica, sans-serif", fontSize: "12px", color: "#000", margin: "0 auto" },
    toolbar: { display: "flex", justifyContent: "flex-end", gap: "10px", marginBottom: "14px" },
    btn: { display: "flex", alignItems: "center", gap: "6px", padding: "7px 16px", border: "1px solid #444", borderRadius: "4px", background: "#fff", fontSize: "13px", fontFamily: "Arial, Helvetica, sans-serif", cursor: "pointer" },
    report: { border: "1px solid #000", padding: "16px" },
    topRow: { display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: "14px" },
    companyBlock: { fontSize: "11px", fontWeight: "bold", lineHeight: 1.6, whiteSpace: "pre-line" },
    titleBlock: { textAlign: "center", flex: 1 },
    titleEn: { fontSize: "16px", fontWeight: "bold", letterSpacing: "0.5px" },
    logoBox: { width: "100px", textAlign: "right" },
    logoText: { fontWeight: "bold", fontSize: "20px", color: "#0a3d62", letterSpacing: "1px" },
    logoIcon: { fontSize: "10px", color: "#0a3d62" },
    infoTable: { width: "100%", borderCollapse: "collapse", marginBottom: "14px" },
    infoCellLabel: { border: "1px solid #000", padding: "6px 8px", width: "16%", fontWeight: "bold", fontSize: "11px", verticalAlign: "top" },
    infoCellValue: { border: "1px solid #000", padding: "6px 8px", width: "34%", fontSize: "12px", verticalAlign: "top" },
    table: { width: "100%", borderCollapse: "collapse", marginBottom: "0px" },
    th: { border: "1px solid #000", padding: "6px 8px", fontSize: "11px", textAlign: "left", verticalAlign: "bottom" },
    thRight: { border: "1px solid #000", padding: "6px 8px", fontSize: "11px", textAlign: "right", verticalAlign: "bottom" },
    td: { border: "1px solid #000", padding: "6px 8px", fontSize: "12px", verticalAlign: "top" },
    tdRight: { border: "1px solid #000", padding: "6px 8px", fontSize: "12px", textAlign: "right" },
    // footer row: ways to pay + note
    footerRow: { display: "flex", gap: "16px", marginTop: "14px", fontSize: "11px" },
    footerCol: { flex: 1.4 },
    noteCol: { flex: 1 },
    bold: { fontWeight: "bold" },
    // bank + contact section (same page, directly below footer)
    twoColRow: { display: "flex", gap: "20px", marginTop: "20px", alignItems: "flex-start" },
    halfCol: { flex: 1, fontSize: "11px" },
    bankTable: { width: "100%", borderCollapse: "collapse", fontSize: "11px", marginTop: "4px" },
    bankTd: { padding: "4px 0", textAlign: "left" },
    contactBlock: { padding: "0", fontSize: "11px", marginTop: "4px" },
    contactTitle: { fontWeight: "bold", marginBottom: "6px" },
    contactLine: { marginBottom: "3px", color: "#1155cc", textDecoration: "underline" },
    statusMsg: (ok) => ({ marginTop: "10px", fontSize: "12px", color: ok ? "#16a34a" : "#b91c1c", textAlign: "right" }),
};

const DEFAULT_HEADER = {
    OperatingCompanyAddress: "LYNX District Cooling Services\nTRN : 100066548700003\nFahad Bldg,Hor Al Anz, Dubai",
    ConsumerName: "",
    UnitNo: "",
    BuildingName: "",
    AccountNumber: "",
    CreditNoteBillingPeriod: "",
};



// ─── Main Component ───────────────────────────────────────────────────────────

export default function CreditNoteReport({
    data = SAMPLE_DATA,
    header = SAMPLE_HEADER['data'],
}) {
    const [status, setStatus] = useState("");
    const [loadingPDF, setLoadingPDF] = useState(false);
    const [loadingExcel, setLoadingExcel] = useState(false);
    const logoUrl = '/output.png';

    const h = { ...DEFAULT_HEADER, ...header };
    const details = Array.isArray(data) ? data : [];
    const sumTotal = details.reduce((s, r) => s + (Number(getLineAmount(r)) || 0), 0);

    // Invoice Number / Billing Date also ride along on each detail line
    // (InvoiceNUmber / BillDate) — use those if the header block doesn't have them.
    const firstLine = details[0] || {};
    const invoiceNumber = h.InvoiceNumber || firstLine.InvoiceNUmber || firstLine.InvoiceNumber || "";
    const billingDate = h.BillingDate || firstLine.BillDate || "";

    // ── Excel Export ──────────────────────────────────────────────────────────────
    function exportExcel() {
        setLoadingExcel(true); setStatus("Generating Excel…");
        try {
            const wb = XLSX.utils.book_new();
            const infoRows = [
                [LABELS.name, h.ConsumerName, LABELS.accountNumber, h.AccountNumber],
                [LABELS.unitNo, h.UnitNo, LABELS.billingDate, fmtDate(billingDate)],
                [LABELS.towerName, h.BuildingName, LABELS.invoiceNumber, invoiceNumber],
                ["", "", LABELS.billingPeriod, h.CreditNoteBillingPeriod],
            ];
            const detailHeader = ["Description", "Total Amount (AED)"];
            const detailRows = details.map((r) => [r.Description, fmtMoney(getLineAmount(r))]);
            const totalsRow = ["TOTAL", fmtMoney(sumTotal)];
            const sheetData = [
                [LABELS.documentTitle],
                [],
                ...infoRows,
                [],
                detailHeader,
                ...detailRows,
                totalsRow,
            ];
            const ws = XLSX.utils.aoa_to_sheet(sheetData);
            ws["!cols"] = [{ wch: 32 }, { wch: 24 }, { wch: 24 }, { wch: 20 }];
            XLSX.utils.book_append_sheet(wb, ws, "Credit Note");
            const wsBank = XLSX.utils.aoa_to_sheet([["Bank Details"], [], ...BANK_DETAILS]);
            wsBank["!cols"] = [{ wch: 20 }, { wch: 30 }];
            XLSX.utils.book_append_sheet(wb, wsBank, "Bank Details");
            XLSX.writeFile(wb, `${h.AccountNumber || "credit_note"}.xlsx`);
            setStatus("\u2713 Excel downloaded.");
        } catch (err) { console.error(err); setStatus("Excel export failed."); }
        finally { setLoadingExcel(false); }
    }

    // ── PDF Export ────────────────────────────────────────────────────────────────
    async function exportPDF() {
        setLoadingPDF(true); setStatus("Generating PDF…");
        try {
            const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
            const setLatin = (style = "normal") => doc.setFont("helvetica", style);

            // ── Header ───────────────────────────────────────────────────────────
            setLatin("bold"); doc.setFontSize(10); doc.setTextColor(0, 0, 0);
            h.OperatingCompanyAddress.split("\n").forEach((l, i) => doc.text(l, 14, 16 + i * 5));

            setLatin("bold"); doc.setFontSize(14);
            doc.text(LABELS.documentTitle, 105, 36, { align: "center" });

            let logoBase64 = null;
            try {
                logoBase64 = await new Promise((resolve, reject) => {
                    const img = new Image();
                    img.crossOrigin = "anonymous";
                    img.src = "/output.png";
                    img.onload = () => {
                        const canvas = document.createElement("canvas");
                        canvas.width = img.width; canvas.height = img.height;
                        canvas.getContext("2d").drawImage(img, 0, 0);
                        resolve(canvas.toDataURL("image/png"));
                    };
                    img.onerror = reject;
                });
            } catch (e) { console.warn("Logo load failed:", e); }
            if (logoBase64) doc.addImage(logoBase64, "PNG", 165, 8, 32, 14);

            // ── Info tables ──────────────────────────────────────────────────────
            const leftRows = [
                [LABELS.name, h.ConsumerName],
                [LABELS.unitNo, h.UnitNo],
                [LABELS.towerName, h.BuildingName],
            ];
            const rightRows = [
                [LABELS.accountNumber, h.AccountNumber],
                [LABELS.billingDate, fmtDate(billingDate)],
                [LABELS.invoiceNumber, invoiceNumber],
                [LABELS.billingPeriod, h.CreditNoteBillingPeriod],
            ];
            const startY = 40;
            autoTable(doc, {
                body: leftRows, startY, theme: "grid",
                styles: { fontSize: 8.5, cellPadding: 3, lineColor: [0, 0, 0], lineWidth: 0.2, textColor: [0, 0, 0] },
                columnStyles: { 0: { fontStyle: "bold", cellWidth: 30 }, 1: { cellWidth: 55 } },
                margin: { left: 14 },
            });
            autoTable(doc, {
                body: rightRows, startY, theme: "grid", margin: { left: 112 },
                styles: { fontSize: 8.5, cellPadding: 3, lineColor: [0, 0, 0], lineWidth: 0.2, textColor: [0, 0, 0] },
                columnStyles: { 0: { fontStyle: "bold", cellWidth: 34 }, 1: { cellWidth: 50 } },
            });

            // ── Detail table: Description / Total Amount ────────────────────────
            const tableTop = doc.lastAutoTable.finalY + 2;
            autoTable(doc, {
                head: [["Description", "Total Amount (AED)"]],
                body: details.map((r) => [r.Description, fmtMoney(getLineAmount(r))]),
                foot: [["Total", fmtMoney(sumTotal)]],
                startY: tableTop,
                styles: { fontSize: 9, cellPadding: 3, lineColor: [0, 0, 0], lineWidth: 0.2, textColor: [0, 0, 0] },
                headStyles: { fillColor: [255, 255, 255], textColor: [0, 0, 0], fontStyle: "bold", lineColor: [0, 0, 0], lineWidth: 0.2 },
                footStyles: { fillColor: [255, 255, 255], textColor: [0, 0, 0], fontStyle: "bold", lineColor: [0, 0, 0], lineWidth: 0.2 },
                columnStyles: { 0: { cellWidth: 140 }, 1: { cellWidth: 42, halign: "right" } },
            });

            // ── Ways to pay + Note ───────────────────────────────────────────────
            let y = doc.lastAutoTable.finalY + 8;
            setLatin("bold"); doc.setFontSize(9);
            doc.text("Ways to pay your Bill:", 14, y);
            setLatin("normal");
            doc.text("Cash , Cheque , Bank Transfer, Paybay Link", 55, y);

            setLatin("bold"); doc.setFontSize(9);
            doc.text("Note:", 120, y);
            setLatin("normal"); doc.setFontSize(8.5);
            doc.text("A Late Fee will be applied to unpaid\naccounts after the Invoice Due Date", 135, y, { maxWidth: 60 });

            // ── Bank Details + Contact Info (same page) ─────────────────────────
            y += 16;
            autoTable(doc, {
                head: [["Bank Details:", ""]],
                body: BANK_DETAILS,
                startY: y, theme: "plain", tableWidth: 95,
                styles: { fontSize: 9, cellPadding: 1.5, textColor: [0, 0, 0] },
                headStyles: { fontStyle: "bold", textColor: [0, 0, 0] },
                columnStyles: { 0: { fontStyle: "bold", cellWidth: 32 }, 1: { cellWidth: 55 } },
                margin: { left: 14 },
            });
            setLatin("bold"); doc.setFontSize(9); doc.setTextColor(0, 0, 0);
            doc.text("Contact Informations:", 127, y + 4);
            setLatin("normal"); doc.setFontSize(8.5); doc.setTextColor(17, 85, 204);
            ["Enquiries@lynxllc.ae", "info@lynxllc.ae"].forEach((line, i) => {
                doc.text(line, 127, y + 10 + i * 6);
            });

            doc.save(`${h.AccountNumber || "credit_note"}.pdf`);
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
                {/* Header */}
                <div style={S.topRow}>
                    <div style={S.companyBlock}>{h.OperatingCompanyAddress}</div>
                    <div style={S.titleBlock}>
                        <div style={S.titleEn}>{LABELS.documentTitle}</div>
                    </div>
                    <div style={S.logoBox}>
                        {logoUrl
              ? <img src={logoUrl} alt="Logo" style={{ maxWidth: "100%", maxHeight: "70px" }} />
              : <img src="https://via.placeholder.com/120x70?text=Logo" alt="Logo" style={{ maxWidth: "100%", maxHeight: "70px" }} />
            }
                    </div>
                </div>

                {/* Info tables */}
                <div style={{ display: "flex", gap: "20px" }}>
                    <table style={{ ...S.infoTable, width: "50%", marginRight: "30px" }}>
                        <tbody>
                            {[
                                [LABELS.name, h.ConsumerName],
                                [LABELS.unitNo, h.UnitNo],
                                [LABELS.towerName, h.BuildingName],
                            ].map((row, i) => (
                                <tr key={i}>
                                    <td style={S.infoCellLabel}>{row[0]}</td>
                                    <td style={S.infoCellValue}>{row[1]}</td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                    <table style={{ ...S.infoTable, width: "50%", marginLeft: "40px" }}>
                        <tbody>
                            {[
                                [LABELS.accountNumber, h.AccountNumber],
                                [LABELS.billingDate, fmtDate(billingDate)],
                                [LABELS.invoiceNumber, invoiceNumber],
                                [LABELS.billingPeriod, h.CreditNoteBillingPeriod],
                            ].map((row, i) => (
                                <tr key={i}>
                                    <td style={S.infoCellLabel}>{row[0]}</td>
                                    <td style={S.infoCellValue}>{row[1]}</td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>

                {/* Detail table: Description / Total Amount */}
                <table style={S.table}>
                    <thead>
                        <tr>
                            <th style={S.th}>Description</th>
                            <th style={S.thRight}>Total Amount (AED)</th>
                        </tr>
                    </thead>
                    <tbody>
                        {details.map((row, i) => (
                            <tr key={row.SrNo ?? i}>
                                <td style={S.td}>{row.Description}</td>
                                <td style={S.tdRight}>{fmtMoney(getLineAmount(row))}</td>
                            </tr>
                        ))}
                        <tr>
                            <td style={{ ...S.td, fontWeight: "bold" }}>Total</td>
                            <td style={{ ...S.tdRight, fontWeight: "bold" }}>{fmtMoney(sumTotal)}</td>
                        </tr>
                    </tbody>
                </table>

                {/* Ways to pay + Note */}
                <div style={S.footerRow}>
                    <div style={S.footerCol}>
                        <span style={S.bold}>Ways to pay your Bill:</span>{"  "}
                        Cash , Cheque , Bank Transfer, Paybay Link
                    </div>
                    <div style={S.noteCol}>
                        <span style={S.bold}>Note:</span>{"  "}
                        A Late Fee will be applied to unpaid accounts after the Invoice Due Date
                    </div>
                </div>

                {/* Bank Details + Contact Info (same page) */}
                <div style={S.twoColRow}>
                    <div style={S.halfCol}>
                        <div style={{ fontWeight: "bold", marginBottom: "4px" }}>Bank Details:</div>
                        <table style={S.bankTable}>
                            <tbody>
                                {BANK_DETAILS.map(([label, value]) => (
                                    <tr key={label}>
                                        <td style={{ ...S.bankTd, fontWeight: "bold", width: "40%" }}>{label}</td>
                                        <td style={S.bankTd}>{value}</td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                    <div style={S.halfCol}>
                        <div style={S.contactBlock}>
                            <div style={S.contactTitle}>Contact Informations:</div>
                            <div style={S.contactLine}>Enquiries@lynxllc.ae</div>
                            <div style={S.contactLine}>info@lynxllc.ae</div>
                        </div>
                    </div>
                </div>
            </div>
            {status && <div style={S.statusMsg(status.startsWith("✓"))}>{status}</div>}
        </div>
    );
}
