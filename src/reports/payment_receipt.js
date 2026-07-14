import { useState } from "react";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";

async function loadData() {
    try {
        const response = await fetch('/data/paymentreceipt.json');
        if (!response.ok) throw new Error('Network response was not ok');
        const data = await response.json();
        return data;
    } catch (error) {
        console.error('Error reading JSON:', error);
        return null;
    }
}

const SAMPLE_DATA = await loadData();

const COMPANY_ADDRESS = "LYNX District Cooling Services\nTRN : 100066548700003\nFahad Bldg,Hor Al Anz, Dubai";

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
    refRow: { display: "flex", gap: "50px", margin: "26px 0 20px 0", fontSize: "12px" },
    refLabel: { fontWeight: "bold" },
    infoTable: { width: "100%", borderCollapse: "collapse", marginBottom: "40px" },
    infoCellLabel: { border: "1px solid #000", padding: "7px 10px", width: "18%", fontWeight: "bold", fontSize: "11px", verticalAlign: "top" },
    infoCellValue: { border: "1px solid #000", padding: "7px 10px", width: "32%", fontSize: "12px", verticalAlign: "top" },
    signatureRow: { display: "flex", justifyContent: "flex-end", marginTop: "50px" },
    signatureLabel: { fontWeight: "bold", fontSize: "11px", borderTop: "1px solid #000", paddingTop: "4px" },
    statusMsg: (ok) => ({ marginTop: "10px", fontSize: "12px", color: ok ? "#16a34a" : "#b91c1c", textAlign: "right" }),
};

// ─── Main Component ───────────────────────────────────────────────────────────

export default function PaymentReceipt({ data = SAMPLE_DATA }) {
    const receipt = (Array.isArray(data) ? data[0] : data) || {};
    const [status, setStatus] = useState("");
    const [loadingPDF, setLoadingPDF] = useState(false);
    const [loadingExcel, setLoadingExcel] = useState(false);

    const referenceNo = receipt.ReferenceNumber || "";
    const date = fmtDate(receipt.CollectionDate);
    const consumerName = receipt.ConsumerName || "";
    const towerName = receipt.EstateName || "";
    const unitNo = receipt.UnitNumber ?? "";
    const accountNo = receipt.AccountNumber || "";
    const amount = fmtMoney(receipt.CollectionAmount);
    const paymentType = receipt.PaymentName || "";
    const logoUrl = '/output.png';

    // ── Excel Export ──────────────────────────────────────────────────────────────
    function exportExcel() {
        setLoadingExcel(true); setStatus("Generating Excel…");
        try {
            const wb = XLSX.utils.book_new();
            const sheetData = [
                ["Payment Receipt"],
                [],
                ["Reference No", referenceNo, "Date", date],
                [],
                ["Consumer Name", consumerName],
                ["Tower Name", towerName],
                ["Unit No", unitNo, "Account No", accountNo],
                ["Amount (AED)", amount, "Payment Type", paymentType],
            ];
            const ws = XLSX.utils.aoa_to_sheet(sheetData);
            ws["!cols"] = [{ wch: 20 }, { wch: 30 }, { wch: 18 }, { wch: 22 }];
            XLSX.utils.book_append_sheet(wb, ws, "Payment Receipt");
            XLSX.writeFile(wb, `${referenceNo || "payment_receipt"}.xlsx`);
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
            COMPANY_ADDRESS.split("\n").forEach((l, i) => doc.text(l, 14, 16 + i * 5));

            setLatin("bold"); doc.setFontSize(14);
            doc.text("Payment Receipt", 105, 22, { align: "center" });

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

            let y = 40;
            setLatin("bold"); doc.setFontSize(9); doc.setTextColor(0, 0, 0);
            doc.text("Reference No:", 14, y);
            setLatin("normal");
            doc.text(referenceNo, 45, y);
            setLatin("bold");
            doc.text("Date:", 150, y);
            setLatin("normal");
            doc.text(date, 165, y);

            // ── Info table ────────────────────────────────────────────────────────
            y += 10;
            autoTable(doc, {
                body: [
                    ["Consumer Name:", consumerName],
                    ["Tower Name:", towerName],
                ],
                startY: y,
                theme: "grid",
                styles: {
                    fontSize: 9,
                    cellPadding: 3.5,
                    lineColor: [0, 0, 0],
                    lineWidth: 0.2,
                    textColor: [0, 0, 0],
                },
                columnStyles: {
                    0: { fontStyle: "bold", cellWidth: 36 },
                    1: { cellWidth: 144 }, 
                },
            });

            const table1EndY = doc.lastAutoTable.finalY;

            autoTable(doc, {
                body: [
                    ["Unit No:", unitNo, "Account No:", accountNo],
                    ["Amount (AED):", amount, "Payment Type:", paymentType],
                ],
                startY: table1EndY,
                theme: "grid",
                styles: {
                    fontSize: 9,
                    cellPadding: 3.5,
                    lineColor: [0, 0, 0],
                    lineWidth: 0.2,
                    textColor: [0, 0, 0],
                },
                columnStyles: {
                    0: { fontStyle: "bold", cellWidth: 36 },
                    1: { cellWidth: 72 },
                    2: { fontStyle: "bold", cellWidth: 28 },
                    3: { cellWidth: 44 },
                },
            });

            const sigY = doc.lastAutoTable.finalY + 40;
            setLatin("bold"); doc.setFontSize(9);
            doc.line(150, sigY, 196, sigY);
            doc.text("Authorized Signature", 150, sigY + 5);

            doc.save(`${referenceNo || "payment_receipt"}.pdf`);
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
                    <div style={S.companyBlock}>{COMPANY_ADDRESS}</div>
                    <div style={S.titleBlock}>
                        <div style={S.titleEn}>Payment Receipt</div>
                    </div>
                    <div style={S.logoBox}>
                       {logoUrl
              ? <img src={logoUrl} alt="Logo" style={{ maxWidth: "100%", maxHeight: "70px" }} />
              : <img src="https://via.placeholder.com/120x70?text=Logo" alt="Logo" style={{ maxWidth: "100%", maxHeight: "70px" }} />
            }
                    </div>
                </div>

                {/* Reference No / Date */}
                <div style={S.refRow}>
                    <div><span style={S.refLabel}>Reference No: </span>{referenceNo}</div>
                    <div><span style={S.refLabel}>Date: </span>{date}</div>
                </div>

                {/* Info table */}
                <table style={S.infoTable}>
                    <tbody>
                        <tr>
                            <td style={S.infoCellLabel}>Consumer Name:</td>
                            <td style={{ ...S.infoCellValue, width: "82%" }} colSpan={3}>{consumerName}</td>
                        </tr>
                        <tr>
                            <td style={S.infoCellLabel}>Tower Name:</td>
                            <td style={{ ...S.infoCellValue, width: "82%" }} colSpan={3}>{towerName}</td>
                        </tr>
                        <tr>
                            <td style={S.infoCellLabel}>Unit No:</td>
                            <td style={S.infoCellValue}>{unitNo}</td>
                            <td style={S.infoCellLabel}>Account No:</td>
                            <td style={S.infoCellValue}>{accountNo}</td>
                        </tr>
                        <tr>
                            <td style={S.infoCellLabel}>Amount (AED):</td>
                            <td style={S.infoCellValue}>{amount}</td>
                            <td style={S.infoCellLabel}>Payment Type:</td>
                            <td style={S.infoCellValue}>{paymentType}</td>
                        </tr>
                    </tbody>
                </table>

                {/* Authorized Signature */}
                <div style={S.signatureRow}>
                    <div style={S.signatureLabel}>Authorized Signature</div>
                </div>
            </div>
            {status && <div style={S.statusMsg(status.startsWith("✓"))}>{status}</div>}
        </div>
    );
}