import React, { useState, useMemo, useCallback } from "react";
import * as XLSX from "xlsx-js-style";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import Select from "react-select";



async function loadData(path) {
    try {
                const response = await fetch(path);
        if (!response.ok) throw new Error('Network response was not ok');
        const data = await response.json()
        return data
    } catch (error) {
        console.error('Error reading JSON:', error);
        return [];
    }
}

const SAMPLE_DATA = await loadData('data/statement_of_account_detailed.json');
const SAMPLE_HEADER = await loadData('data/statement_of_account_header.json');

const COLUMNS = [
    { key: "CollectionDate", label: "Collection Date", type: "date" },
    { key: "BillingPeriod", label: "Billing Period", type: "text" },
    { key: "BillPostDate", label: "Bill Post Date", type: "date" },
        { key: "InvoiceNUmber", label: "Invoice Number", type: "text" },
  { key: "PaymentTypeName", label: "Transaction Type", type: "text" },
    { key: "PreviousBalance", label: "Previous Balance", type: "amount" },
    { key: "SecurityDeposit", label: "Security Deposit", type: "amount" },
    { key: "ActivationFee", label: "Activation Fee", type: "amount" },
    { key: "BillingFee", label: "Billing Fee", type: "amount" },
    { key: "CapacityCharges", label: "Capacity Charges", type: "amount" },
    { key: "ConsumptionCharges", label: "Consumption Charges", type: "amount" },
    { key: "VATOnBillingFee", label: "VAT On Billing Fee", type: "amount" },
    { key: "LateFee", label: "Late Fee", type: "amount" },
    { key: "VATAmount", label: "VAT Amount", type: "amount" },
        { key: "OtherCharges", label: "Other Charges", type: "amount" },
        { key: "AdvancePayment", label: "Advance Payment", type: "amount" },
    { key: "Adjustment", label: "Adjustment", type: "amount" },
    { key: "TotalAmountInclVAT", label: "Total Bill", type: "amount" },
        { key: "CollectionAmount", label: "Total Paid", type: "amount" },
    { key: "Balance", label: "Balance", type: "total" },



];

const firstRow = SAMPLE_DATA[0]; 

const columns = COLUMNS.map(column => {
  const uom = firstRow?.[`${column.key}UOM`];

  return {
    ...column,
    label: uom ? `${column.label}\n(${uom})` : column.label,
  };
});

const SUM_KEYS = [
    "PreviousBalance",
    "SecurityDeposit",
    "ActivationFee",
    "BillingFee",
    "CapacityCharges",
    "ConsumptionCharges",
    "VATOnBillingFee",
    "LateFee",
    "VATAmount",
    "TotalAmountInclVAT",
    "Balance",
    "CollectionAmount",
    "OtherCharges",
    "AdvancePayment",
    "Adjustment",
];

/* ------------------------------------------------------------------ */
/* Helpers                                                             */
/* ------------------------------------------------------------------ */
const fmt = (n) =>
    (Number(n) || 0).toLocaleString("en-US", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
    });

const fmtDateTime = (dateStr) => {
    if (!dateStr) return "—";
    try {
        const d = new Date(dateStr);
        const dd = String(d.getDate()).padStart(2, "0");
        const mm = String(d.getMonth() + 1).padStart(2, "0");
        const yyyy = d.getFullYear();
        return `${dd}/${mm}/${yyyy}`;
    } catch {
        return dateStr;
    }
};

const cellValue = (col, row) => {
    if (col.type === "text") return row[col.key] ?? "—";
    if (col.type === "amount" || col.type === "total") return fmt(row[col.key]);
    if (col.type === "date") return fmtDateTime(row[col.key]);
    return row[col.key] ?? "—";
};

function sumTotals(rows) {
    const totals = {};
    SUM_KEYS.forEach((k) => (totals[k] = 0));
    rows.forEach((r) => SUM_KEYS.forEach((k) => (totals[k] += Number(r[k]) || 0)));
    return totals;
}

export default function StatementOfAccountDetailedReport({ fetchData }) {
    const [data, setData] = useState(SAMPLE_DATA);
    const [header, setHeader] = useState(SAMPLE_HEADER);
    const [selectedAgreement, setSelectedAgreement] = useState(null);
    const [loading, setLoading] = useState(false);
        const logoUrl = '/output.png';


    const filteredData = useMemo(() => (data ? data : []), [data]);

    const generatedBy = data[0]?.GeneratedBy || "—";
    const generatedOn = data[0]?.GeneratedOn || fmtDateTime(new Date().toISOString());
    const uom =
        data[0]?.BalanceUOM ||
        data[0]?.TotalAmountInclVATUOM ||
        "AED";

    const totals = useMemo(() => sumTotals(filteredData), [filteredData]);

    const agreementOptions = useMemo(() => {
        const keys = new Map();
        (data || []).forEach((r) => {
            if (r.ServiceAgreementKey) keys.set(r.ServiceAgreementKey, r);
        });

        return Array.from(keys.values()).map((item) => {
            const key = item.ServiceAgreementKey;
            let label = header.ConsumerNumber ?  header.ConsumerNumber : item.StatementOfAccountID;

            return { value: key, label };
        });
    }, [data,header]);

    const loadImageAsDataUrl = useCallback((url) => {
        return new Promise((resolve, reject) => {
            if (!url) return resolve(null);
            const img = new Image();
            img.crossOrigin = "anonymous";
            img.onload = () => {
                try {
                    const canvas = document.createElement("canvas");
                    canvas.width = img.naturalWidth;
                    canvas.height = img.naturalHeight;
                    const ctx = canvas.getContext("2d");
                    ctx.drawImage(img, 0, 0);
                    resolve(canvas.toDataURL("image/png"));
                } catch (err) {
                    reject(err);
                }
            };
            img.onerror = (e) => reject(e);
            img.src = url;
        });
    }, []);

    const handleGenerate = useCallback(async () => {
        if (!fetchData) {
            setData(SAMPLE_DATA);
            setHeader(SAMPLE_HEADER);
            return;
        }
        setLoading(true);
        try {
            const param = selectedAgreement ? selectedAgreement.value : undefined;
            const result = await fetchData(param);
            if (result && result.header) setHeader(result.header);
            setData(Array.isArray(result) ? result : []);
        } catch (err) {
            console.error("Failed to load statement of account detailed report:", err);
        } finally {
            setLoading(false);
        }
    }, [fetchData, selectedAgreement]);

    const handleExportXlsx = useCallback(() => {
        const headerxl = columns.map((c) => c.label);
        const rows = filteredData.map((r) =>
            columns.map((c) => {
                if (c.type === "amount" || c.type === "total") return Number(r[c.key]) || 0;
                if (c.type === "date") return fmtDateTime(r[c.key]);
                return r[c.key] ?? "";
            })
        );

        const footRow = columns.map((c, i) => {
            if (i === 0) return "Totals";
            if (c.type === "text" || c.type === "date") return "";
            return totals[c.key];
        });

        const titleRow = [`Statement of Account Detailed Report`];
        const generatedRow = [
            `Generated On: ${generatedOn}`,
            ...Array(columns.length - 2).fill(""),
            `Generated By: ${generatedBy}`,
        ];

        const headerInfoRow = [
            `Estate: ${header?.EstateName || "—"}`,
            `Customer No.: ${header?.ConsumerNumber || "—"}`,
            `Customer Name: ${header?.ConsumerName || "—"}`,
            ...Array(Math.max(0, columns.length - 3)).fill(""),
        ];

        const aoa = [
            titleRow,
            generatedRow,
            headerInfoRow,
            [],
            headerxl,
            ...rows,
            footRow,
        ];

        const ws = XLSX.utils.aoa_to_sheet(aoa);
        ws["!cols"] = columns.map((c) => ({
            wch: c.type === "text" ? 22 : 16,
        }));
        ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: columns.length - 1 } }];

        
        const range = XLSX.utils.decode_range(ws["!ref"]);
        

           for (let R = range.s.r; R <= range.e.r; R++) {
                    for (let C = range.s.c; C <= range.e.c; C++) {
                        if (C===1) continue; 
                        const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
        
                        if (ws[cellRef]) {
                            ws[cellRef].s = {
                                alignment: {
                                    wrapText: true,
                                    overflowX:'auto',
                                    overflowY:'hidden',
                                    // vertical: "center",
                                    // horizontal: "center",
                                },
                            };
                        }
                    }
                }
        
                ws["!rows"] = aoa.map((_, index) => {
                    if (index === 4 || index === 2) {
                        return { hpt: 40 }; 
                    }
                    return { hpt: 20 };
                });
        
        

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Statement of Account");
        XLSX.writeFile(wb, `StatementOfAccountDetailed_${generatedOn.replace(/\//g, "-")}.xlsx`);
    }, [filteredData, generatedBy, generatedOn, totals, header]);

    const handleExportPdf = useCallback(async () => {
        const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4" });
        const pageWidth = doc.internal.pageSize.getWidth();

        // try to load and draw logo at top-right
        try {
            const dataUrl = await loadImageAsDataUrl(logoUrl);
            if (dataUrl) {
                const imgW = 120;
                const imgH = 70;
                doc.addImage(dataUrl, "PNG", pageWidth - 40 - imgW, 20, imgW, imgH);
            }
        } catch (err) {
            console.warn("Failed to load logo for PDF:", err);
        }

        doc.setFontSize(13);
        doc.setFont(undefined, "bold");
        doc.setTextColor(16, 36, 61);
        doc.text(`Statement of Account Detailed Report`, pageWidth / 2, 40, { align: "center" });

        doc.setFontSize(9);
        doc.setFont(undefined, "normal");
        doc.setTextColor(90, 105, 120);
       

        doc.text(`Estate: ${header?.EstateName || "—"}`, 40, 78);
        doc.text(`Customer No.: ${header?.ConsumerNumber || "—"}`, 40, 94);
        doc.text(`Customer Name: ${header?.ConsumerName || "—"}`, 40, 110);
         doc.text(`Generated On: ${generatedOn}`, 40, 126);
        // doc.text(`Generated By: ${generatedBy}`, 40, 142);

        const head = [columns.map((c) => c.label)];
        const body = filteredData.map((r) => columns.map((c) => cellValue(c, r)));

        body.push(
            columns.map((c, i) => {
                if (i === 0) return "Totals";
                if (c.type === "text" || c.type === "date") return "";
                return `${fmt(totals[c.key])} ${uom}`;
            })
        );

        const columnStyles = {};
        columns.forEach((c, i) => {
            columnStyles[i] = { halign: c.type === "text" || c.type === "date" ? "left" : "right" };
            if (c.type === "total") columnStyles[i].fontStyle = "bold";
        });

        autoTable(doc, {
            head,
            body,
            startY: 150,
            styles: { fontSize: 7, cellPadding: 4, lineColor: [221, 227, 234], lineWidth: 0.5 },
            headStyles: { fillColor: [234, 241, 251], textColor: [16, 36, 61], fontStyle: "bold", fontSize: 7 },
            columnStyles,
            alternateRowStyles: { fillColor: [249, 251, 253] },
            didParseCell: (cellData) => {
                if (cellData.row.index === body.length - 1 && cellData.section === "body") {
                    cellData.cell.styles.fillColor = [234, 241, 251];
                    cellData.cell.styles.fontStyle = "bold";
                }
            },
        });

        doc.save(`StatementOfAccountDetailed_${generatedOn.replace(/\//g, "-")}.pdf`);
    }, [filteredData, generatedBy, generatedOn, totals, uom, header, loadImageAsDataUrl, logoUrl]);

    return (
        <div style={styles.app}>
            {/* Toolbar */}
            <div style={styles.toolbar}>
                <div style={styles.controls}>
                    <div style={{ minWidth: 300 }}>
                        <Select
                            value={selectedAgreement}
                            onChange={setSelectedAgreement}
                            options={agreementOptions}
                            placeholder="Select service agreement..."
                            isClearable
                        />
                    </div>
                    <button style={styles.btnPrimary} onClick={handleGenerate} disabled={loading}>
                        {loading ? "Loading…" : "⟳ Generate Report"}
                    </button>
                </div>
            </div>

            {/* Export bar */}
            <div style={styles.exportBar}>
                <div style={styles.statusLine}>
                    Showing live data · {filteredData.length} record{filteredData.length === 1 ? "" : "s"}
                </div>
                <div style={styles.exportActions}>
                    <button style={styles.btnExport} onClick={handleExportXlsx}>
                        Export Excel
                    </button>
                    <button style={styles.btnExport} onClick={handleExportPdf}>
                        Export PDF
                    </button>
                </div>
            </div>

       <div style={styles.sheet}>
    <div style={styles.sheetInner}>
        <div style={styles.logoBox}>
            {logoUrl ? (
                <img
                    src={logoUrl}
                    alt="Logo"
                    style={{ maxWidth: "100%", maxHeight: "70px" }}
                />
            ) : (
                <img
                    src="https://via.placeholder.com/120x70?text=Logo"
                    alt="Logo"
                    style={{ maxWidth: "100%", maxHeight: "70px" }}
                />
            )}
        </div>

        <div style={styles.reportTitle}>
            Statement of Account Detailed Report
        </div>

        <div style={styles.reportHeader}>

             <div style={styles.reportHeaderItem}>
                <strong>Customer Name:</strong> {header?.ConsumerName || "—"}
            </div>
            <div style={styles.reportHeaderItem}>
                <strong>Tower Name:</strong> {header?.EstateName || "—"}
            </div>
            <div style={styles.reportHeaderItem}>
                <strong>Service Agreement No:</strong> {header?.ConsumerNumber || "—"}
            </div>
           
        </div>

        <div style={styles.reportMeta}>
            <span>
                Generated On: <b>{generatedOn}</b>
            </span>
            {/* <span>
                Generated By: <b>{generatedBy}</b>
            </span> */}
        </div>

                    {filteredData.length === 0 ? (
                        <div style={styles.empty}>
                            No statement of account data available.
                        </div>
                    ) : (
                        <div style={styles.tableWrapper}>
                            <table style={styles.table}>
                                <thead>
                                    <tr>
                                        {columns.map((c) => (
                                            <th
                                                key={c.key}
                                                style={{
                                                    ...styles.th,
                                                    textAlign: c.type === "text" || c.type === "date" ? "left" : "right",
                                                }}
                                            >
                                                {c.label}
                                            </th>
                                        ))}
                                    </tr>
                                </thead>

                                <tbody>
                                    {filteredData.map((r, i) => (
                                        <tr
                                            key={r.StatementOfAccountID || `${r.InvoiceDataID}-${i}`}
                                            style={i % 2 === 1 ? styles.rowAlt : undefined}
                                        >
                                            {columns.map((c) => (
                                                <td
                                                    key={c.key}
                                                    style={{
                                                        ...styles.td,
                                                        textAlign: c.type === "text" || c.type === "date" ? "left" : "right",
                                                        fontWeight: c.type === "total" ? 700 : 400,
                                                        color: c.type === "total" ? "#10243d" : "#1b2733",
                                                    }}
                                                >
                                                    {cellValue(c, r)}
                                                </td>
                                            ))}
                                        </tr>
                                    ))}
                                </tbody>

                                <tfoot>
                                    <tr>
                                        {columns.map((c, i) => (
                                            <td
                                                key={c.key}
                                                style={{
                                                    ...styles.tfootTd,
                                                    textAlign: c.type === "text" || c.type === "date" ? "left" : "right",
                                                }}
                                            >
                                                {i === 0
                                                    ? "Totals"
                                                    : c.type === "text" || c.type === "date"
                                                        ? ""
                                                        : `${fmt(totals[c.key])} ${uom}`}
                                            </td>
                                        ))}
                                    </tr>
                                </tfoot>
                            </table>
                        </div>
                    )}
                </div>
            </div>
        </div>
    );
}

/* ------------------------------------------------------------------ */
/* Inline styles — flat, matching the plain Excel output              */
/* ------------------------------------------------------------------ */
const styles = {
    app: {
        maxWidth: 900,
        margin: "0 auto",
        padding: "20px 20px 60px",
        fontFamily: "'Segoe UI', system-ui, -apple-system, sans-serif",
        color: "#1b2733",
        background: "#f2f5f8",
    },
    toolbar: {
        padding: "16px 20px",
        display: "flex",
        alignItems: "center",
        justifyContent: "space-between",
        gap: 16,
        flexWrap: "wrap",
    },
    toolbarTitle: {
        color: "#10243d",
        fontSize: 17,
        fontWeight: 600,
        margin: 0,
    },
    controls: { display: "flex", alignItems: "flex-end", gap: 14, flexWrap: "wrap" },
    field: { display: "flex", flexDirection: "column", gap: 4 },
    label: {
        fontSize: 10.5,
        color: "#10243d",
        textTransform: "uppercase",
        letterSpacing: "0.06em",
        fontWeight: 600,
    },
    input: {
        background: "#fff",
        color: "#1b2733",
        border: "1px solid #35516f",
        padding: "7px 10px",
        fontSize: 13,
        minWidth: 160,
        outline: "none",
    },
    btnPrimary: {
        border: "1px solid #35516f",
        padding: "8px 16px",
        fontSize: 13,
        fontWeight: 600,
        cursor: "pointer",
        background: "#fff",
        color: "#10243d",
    },
    exportBar: {
        display: "flex",
        justifyContent: "space-between",
        alignItems: "center",
        margin: "16px 0 0",
        padding: "10px 4px",
        flexWrap: "wrap",
        gap: 10,
    },
    statusLine: { fontSize: 12, color: "#5b6b7b" },
    exportActions: { display: "flex", gap: 10 },
    btnExport: {
        background: "#fff",
        border: "1px solid #dde3ea",
        padding: "8px 14px",
        fontSize: 12.5,
        fontWeight: 600,
        cursor: "pointer",
    },
    sheet: {
        background: "#fff",
        marginTop: 14,
        border: "1px solid #dde3ea",
    },
 sheetInner: {
        position: "relative",
        padding: "24px",
    },
    tableWrapper: {
        width: "100%",
        overflowX: "auto",
        overflowY: "hidden",
    },
        reportTitle: {
        textAlign: "center",
        fontSize: "17px",
        fontWeight: "bold",
        marginBottom: "50px",
    },

        reportHeader: {
            display: "flex",
                flexDirection: "column",
            justifyContent: "center",
            gap: 20,
            marginBottom: 12,
            flexWrap: "wrap",
        },

        reportHeaderItem: {
            fontSize: 13,
            color: "#1b2733",
        },

    reportMeta: {
        display: "flex",
                fontSize: 13,

        justifyContent: "space-between",
        marginBottom: "20px",
        paddingRight: "140px", // prevents text from overlapping the logo
    },

    table: {
        width: "max-content",
        minWidth: "100%",
        borderCollapse: "collapse",
        fontSize: 12.5,
    },
    th: {
        background: "#eaf1fb",
        color: "#10243d",
        fontWeight: 700,
        fontSize: 11.5,
        padding: "9px 10px",
        border: "1px solid #dde3ea",
        whiteSpace: "nowrap",
    },
    td: { padding: "8px 10px", border: "1px solid #dde3ea", whiteSpace: "nowrap" },
    rowAlt: { background: "#f9fbfd" },
    tfootTd: {
        padding: "10px 8px",
        border: "1px solid #dde3ea",
        borderTop: "2px solid #10243d",
        background: "#eaf1fb",
        fontWeight: 700,
        color: "#10243d",
    },
    empty: { textAlign: "center", padding: "40px 0", color: "#5b6b7b", fontSize: 13 },
        logoBox: {
        position: "absolute",
        top: "20px",
        right: "20px",
        width: "120px",
        height: "70px",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
    },
};