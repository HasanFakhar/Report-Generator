
import React, { useState, useMemo, useCallback } from "react";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import * as XLSX from "xlsx";
import Select from "react-select";


async function loadData() {
    try {
        const response = await fetch('/data/towerwise_data.json');
        if (!response.ok) throw new Error('Network response was not ok');
        console.log(response[1])
        const data = await response.json()
        return data


    } catch (error) {
        console.error('Error reading JSON:', error);
    }

}
const SAMPLE_DATA = await loadData();

const TOWERS = {
    "ROYAL RESIDENCE 1 key ": "ROYAL RESIDENCE 1",
    "ROYAL RESIDENCE 2 key ": "ROYAL RESIDENCE 2",
    "ROYAL RESIDENCE 3 key": "ROYAL RESIDENCE 3"
};

const COLUMNS = [
        { key: "BillGeneratedDate", label: "Bill Date", type: "date" },
    { key: "CollectionDate", label: "Collection Date", type: "date" },

    { key: "InvoiceNumber", label: "Invoice No.", type: "text" },
    { key: "AccountNo", label: "Account No.", type: "text" },
    { key: "UnitNumber", label: "Unit No.", type: "text" },
    { key: "ConsumerName", label: "Consumer Name", type: "text" },
    { key: "PaymentType", label: "Payment Type", type: "text" },
    { key: "DepositHeld", label: "Deposit Heldcolumns", type: "money" },
    { key: "AdminFee", label: "Admin Feecolumns", type: "money" },
    { key: "UtilityCharges", label: "Utility Chargescolumns", type: "money" },
    { key: "CapacityCharges", label: "Capacity Chargescolumns", type: "money" },
    { key: "BillingFee", label: "Billing Feecolumns", type: "money" },
    { key: "LateFee", label: "Late Feecolumns", type: "money" },
    { key: "OtherFee", label: "Other Feecolumns", type: "money" },
    { key: "Vat5Per", label: "VAT 5%columns", type: "money" },
    { key: "AdvanceAmount", label: "Advance Amountcolumns", type: "money" },
    { key: "AdjustmentAmount", label: "Adjustmentcolumns", type: "money" },
    { key: "Total", label: "Totalcolumns", type: "total" },
];
const firstRow = SAMPLE_DATA[0]; //replace sample data with resposne from api call

const columns = COLUMNS.map(column => {
  const uom = firstRow?.[`${column.key}UOM`];

  return {
    ...column,
    label: uom ? `${column.label}\n(${uom})` : column.label,
  };
});

const SUM_KEYS = [
    "DepositHeld",
    "AdminFee",
    "UtilityCharges",
    "CapacityCharges",
    "BillingFee",
    "LateFee",
    "OtherFee",
    "Vat5Per",
    "AdvanceAmount",
    "AdjustmentAmount",
    "Total",
];

/* ------------------------------------------------------------------ */
/* Helpers                                                             */
/* ------------------------------------------------------------------ */
const fmt = (n) =>
    (Number(n) || 0).toLocaleString("en-US", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
    });

const isoToApiDate = (iso) => {
    const [y, m, d] = iso.split("-");
    return `${d}/${m}/${y}`;
};

const fmtDate = (iso) => {
    const d = new Date(iso);
    const dd = String(d.getDate()).padStart(2, "0");
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const yyyy = d.getFullYear();
    return `${dd}/${mm}/${yyyy}`;
};

const cellValue = (col, row) => {
    if (col.type === "date") return fmtDate(row[col.key]);
    if (col.type === "text") return row[col.key];
    return fmt(row[col.key]);
};

function sumTotals(rows) {
    const totals = {};
    SUM_KEYS.forEach((k) => (totals[k] = 0));
    rows.forEach((r) => SUM_KEYS.forEach((k) => (totals[k] += Number(r[k]) || 0)));
    return totals;
}

/* ------------------------------------------------------------------ */
/* Component                                                           */
/* ------------------------------------------------------------------ */
export default function TowerwiseBillingReport({ fetchData }) {
    const [tower, setTower] = useState(Object.keys(TOWERS)[0]);
    const [fromDate, setFromDate] = useState("2026-06-01");
    const [toDate, setToDate] = useState("2026-07-13");
    const [data, setData] = useState(SAMPLE_DATA);
    const [loading, setLoading] = useState(false);

    const apiFrom = isoToApiDate(fromDate);
    const apiTo = isoToApiDate(toDate);
    const towerLabel = TOWERS[tower] || tower;
    const generatedBy = data[0]?.GeneratedBy || "—";
    const generatedOn = data[0]?.GeneratedOn || apiTo;
    const towerCount = useMemo(() => new Set(data.map((r) => r.EstateName)).size, [data]);

    const totals = useMemo(() => sumTotals(data), [data]);

    const handleGenerate = useCallback(async () => {
        if (!fetchData) {
            setData(tower === "ALL" ? SAMPLE_DATA : SAMPLE_DATA.filter((r) => r.EstateName === tower));
            return;
        }
        setLoading(true);
        try {
            const result = await fetchData(tower, apiFrom, apiTo);
            setData(Array.isArray(result) ? result : []);
        } catch (err) {
            console.error("Failed to load towerwise billing report:", err);
        } finally {
            setLoading(false);
        }
    }, [fetchData, tower, apiFrom, apiTo]);

    const handleExportPdf = useCallback(() => {
        const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4" });
        const pageWidth = doc.internal.pageSize.getWidth();

        doc.setFontSize(13);
        doc.setFont(undefined, "bold");
        doc.setTextColor(16, 36, 61);
        doc.text(`${towerLabel} Towerwise Billing Report ${apiFrom} to ${apiTo}`, pageWidth / 2, 40, { align: "center" });

        doc.setFontSize(9);
        doc.setFont(undefined, "normal");
        doc.setTextColor(90, 105, 120);
        doc.text(`Generated On: ${generatedOn}`, 40, 62);
        doc.text(`Generated By: ${generatedBy}`, pageWidth - 40, 62, { align: "right" });

        const head = [columns.map((c) => c.label)];
        const body = data.map((r) => columns.map((c) => cellValue(c, r)));

        body.push(
            columns.map((c, i) => {
                if (i === 0) return "Totals";
                if (c.type === "text" || c.type === "date") return "";
                return fmt(totals[c.key]);
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
            startY: 76,
            styles: { fontSize: 8, cellPadding: 5, lineColor: [221, 227, 234], lineWidth: 0.5 },
            headStyles: { fillColor: [234, 241, 251], textColor: [16, 36, 61], fontStyle: "bold", fontSize: 8 },
            columnStyles,
            alternateRowStyles: { fillColor: [249, 251, 253] },
            didParseCell: (cellData) => {
                if (cellData.row.index === body.length - 1 && cellData.section === "body") {
                    cellData.cell.styles.fillColor = [234, 241, 251];
                    cellData.cell.styles.fontStyle = "bold";
                }
            },
        });

        doc.save(`Towerwise_Billing_Report_${tower.replace(/\s+/g, "_")}_${apiFrom.replace(/\//g, "-")}_${apiTo.replace(/\//g, "-")}.pdf`);
    }, [data, tower, towerLabel, apiFrom, apiTo, generatedBy, generatedOn, totals]);

    const handleExportXlsx = useCallback(() => {
        const header = columns.map((c) => c.label);
        const rows = data.map((r) =>
            columns.map((c) => {
                if (c.type === "date") return fmtDate(r[c.key]);
                if (c.type === "text") return r[c.key];
                return Number(r[c.key]);
            })
        );

        const footRow = columns.map((c, i) => {
            if (i === 0) return "Totals";
            if (c.type === "text" || c.type === "date") return "";
            return totals[c.key];
        });

        const aoa = [
            [`${towerLabel} Towerwise Billing Report ${apiFrom} to ${apiTo}`],
            [
                `Generated On: ${generatedOn}`,
                ...Array(columns.length - 2).fill(""),
                `Generated By: ${generatedBy}`,
            ],
            [],
            header,
            ...rows,
            footRow,
        ];

        const ws = XLSX.utils.aoa_to_sheet(aoa);
        ws["!cols"] = columns.map((c) => ({ wch: c.type === "text" ? 18 : c.type === "date" ? 14 : 14 }));
        ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: columns.length - 1 } }];

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Towerwise Billing Report");
        XLSX.writeFile(wb, `Towerwise_Billing_Report_${tower.replace(/\s+/g, "_")}_${apiFrom.replace(/\//g, "-")}_${apiTo.replace(/\//g, "-")}.xlsx`);
    }, [data, tower, towerLabel, apiFrom, apiTo, generatedBy, generatedOn, totals]);

    return (
        <div style={styles.app}>
            {/* Toolbar */}
            <div style={styles.toolbar}>
                <div style={styles.controls}>
                    <div style={styles.field}>
                        <label style={styles.label}>Tower Name</label>
                        <Select
                            value={{
                                label: TOWERS[tower],
                                value: tower,
                            }}
                            onChange={(selected) => setTower(selected.value)}
                            options={Object.keys(TOWERS).map((t) => ({
                                label: TOWERS[t],
                                value: t,
                            }))}
                            isSearchable={true}
                            placeholder="Search Estate..."
                            styles={{
                                control: (base) => ({
                                    ...base,
                                    ...styles.input,
                                    minHeight: "36px",
                                    width: "220px",
                                }),
                                menu: (base) => ({
                                    ...base,
                                    zIndex: 9999,
                                }),
                            }}
                        />
                    </div>
                    <div style={styles.field}>
                        <label style={styles.label}>From Date</label>
                        <input
                            style={styles.input}
                            type="date"
                            value={fromDate}
                            onChange={(e) => setFromDate(e.target.value)}
                        />
                    </div>
                    <div style={styles.field}>
                        <label style={styles.label}>To Date</label>
                        <input style={styles.input} type="date" value={toDate} onChange={(e) => setToDate(e.target.value)} />
                    </div>
                    <button style={styles.btnPrimary} onClick={handleGenerate} disabled={loading}>
                        {loading ? "Loading…" : "⟳ Generate Report"}
                    </button>
                </div>
            </div>

            {/* Export bar */}
            <div style={styles.exportBar}>
                <div style={styles.statusLine}>
                    Showing live data · {data.length} line item{data.length === 1 ? "" : "s"} across {towerCount} tower
                    {towerCount === 1 ? "" : "s"}
                </div>
                <div style={styles.exportActions}>
                    <button style={{ ...styles.btnExport }} onClick={handleExportPdf}>
                         Export PDF
                    </button>
                    <button style={{ ...styles.btnExport }} onClick={handleExportXlsx}>
                         Export Excel
                    </button>
                </div>
            </div>

            {/* Report sheet / HTML preview */}
            <div style={styles.sheet}>
                <div style={styles.sheetInner}>
                    <div style={styles.reportTitle}>{`${towerLabel} Towerwise Billing Report ${apiFrom} to ${apiTo}`}</div>
                    <div style={styles.reportMeta}>
                        <span>
                            Generated On: <b>{generatedOn}</b>
                        </span>
                        <span>
                            Generated By: <b>{generatedBy}</b>
                        </span>
                    </div>

                    {data.length === 0 ? (
                        <div style={styles.empty}>No billing records for the selected date range.</div>
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
                                    {data.map((r, i) => (
                                        <tr key={r.InvoiceNumber || i} style={i % 2 === 1 ? styles.rowAlt : undefined}>
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
                                                {i === 0 ? "Totals" : c.type === "text" || c.type === "date" ? "" : fmt(totals[c.key])}
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
/* Inline styles — flat, matching the plain PDF/Excel output          */
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
    sheetInner: { padding: "34px 38px 28px" },
    tableWrapper: {
        width: "100%",
        overflowX: "auto",
        overflowY: "hidden",
    },
    reportTitle: {
        textAlign: "center",
        fontSize: 17,
        fontWeight: 700,
        color: "#10243d",
        marginBottom: 18,
        whiteSpace: "nowrap",
    },
    reportMeta: {
        display: "flex",
        justifyContent: "space-between",
        fontSize: 12,
        color: "#5b6b7b",
        marginBottom: 16,
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
};
