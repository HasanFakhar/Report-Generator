import React, { useState, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import Select from "react-select";

async function loadData() {
    try {
        const response = await fetch('/data/unitwise_data.json');
        if (!response.ok) throw new Error('Network response was not ok');
        const data = await response.json()
        return data
    } catch (error) {
        console.error('Error reading JSON:', error);
        return [];
    }
}

const SAMPLE_DATA = await loadData();

const TOWERS = {
    "ROYAL RESIDENCE 1 key ": "ROYAL RESIDENCE 1",
    "ROYAL RESIDENCE 2 key ": "ROYAL RESIDENCE 2",
    "ROYAL RESIDENCE 3 key": "ROYAL RESIDENCE 3"
};

const COLUMNS = [
    { key: "UnitNumber", label: "Unit No.", type: "text" },
    { key: "ConsumerName", label: "Consumer Name", type: "text" },
    { key: "AccountNo", label: "Account No.", type: "text" },
    { key: "InvoiceNumber", label: "Invoice No.", type: "text" },
    { key: "BillFromDate", label: "Bill From Date", type: "text" },
    { key: "BillToDate", label: "Bill To Date", type: "text" },
    { key: "BillDate", label: "Bill Date", type: "text" },
    { key: "SecurityDeposite", label: "Security Deposit", type: "amount" },
    { key: "RegistrationFee", label: "Registration Fee", type: "amount" },
    { key: "EneryConsumptionAmount", label: "Energy Consumption", type: "amount" },
    { key: "CapacityCharges", label: "Capacity Charges", type: "amount" },
    { key: "FuelSubCharges", label: "Fuel Sub Charges", type: "amount" },
    { key: "BillingFee", label: "Billing Fee", type: "amount" },
    { key: "LateFee", label: "Late Fee", type: "amount" },
    { key: "ReturnedChequeFee", label: "Returned Cheque Fee", type: "amount" },
    { key: "ReconnectionCharges", label: "Reconnection Charges", type: "amount" },
    { key: "InspectionCharges", label: "Inspection Charges", type: "amount" },
    { key: "VATOnUtility", label: "VAT On Utility", type: "amount" },
    { key: "VATOnOtherFee", label: "VAT On Other Fee", type: "amount" },
    { key: "TotalAmount", label: "Total Amount", type: "amount" },
    { key: "TotalVATAmount", label: "Total VAT Amount", type: "amount" },
    { key: "BilledAmount", label: "Billed Amount", type: "amount" },
    { key: "OpeningPreviousBalance", label: "Opening Balance", type: "amount" },
    { key: "AdvanceAmount", label: "Advance Amount", type: "amount" },
    { key: "TotalAdjustmentAmount", label: "Total Adjustment", type: "amount" },
    { key: "TotalBillingAmount", label: "Total Billing Amount", type: "total" },

];

const SUM_KEYS = [
    "SecurityDeposite",
    "RegistrationFee",
    "EneryConsumptionAmount",
    "CapacityCharges",
    "FuelSubCharges",
    "BillingFee",
    "LateFee",
    "ReturnedChequeFee",
    "ReconnectionCharges",
    "InspectionCharges",
    "VATOnUtility",
    "VATOnOtherFee",
    "TotalAmount",
    "TotalVATAmount",
    "BilledAmount",
    "OpeningPreviousBalance",
    "AdvanceAmount",
    "TotalAdjustmentAmount",
    "TotalBillingAmount"
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
    if (col.type === "text") return row[col.key] || "—";
    if (col.type === "amount" || col.type === "total") return fmt(row[col.key]);
    return row[col.key] || "—";
};

function sumTotals(rows) {
    const totals = {};
    SUM_KEYS.forEach((k) => (totals[k] = 0));
    rows.forEach((r) => SUM_KEYS.forEach((k) => (totals[k] += Number(r[k]) || 0)));
    return totals;
}

export default function UnitWiseBillingReport({ fetchData }) {
    const [estate, setEstate] = useState(Object.keys(TOWERS)[0]);
    const [data, setData] = useState(SAMPLE_DATA);
    const [loading, setLoading] = useState(false);

    const filteredData = useMemo(() => {
        if (!data) return [];
        return data.filter(row => row.EstateName === TOWERS[estate]);
    }, [data, estate]);

    const generatedBy = data[0]?.GeneratedBy || "—";
    const generatedOn = data[0]?.GeneratedOn || fmtDateTime(new Date().toISOString());
    const uom = data[0]?.BilledAmountUOM || "AED";

    const totals = useMemo(() => sumTotals(filteredData), [filteredData]);

    const handleGenerate = useCallback(async () => {
        if (!fetchData) {
            setData(SAMPLE_DATA);
            return;
        }
        setLoading(true);
        try {
            const result = await fetchData(estate);
            setData(Array.isArray(result) ? result : []);
        } catch (err) {
            console.error("Failed to load unit-wise billing report:", err);
        } finally {
            setLoading(false);
        }
    }, [fetchData, estate]);

    const handleExportXlsx = useCallback(() => {
        const header = COLUMNS.map((c) => c.label);
        const rows = filteredData.map((r) =>
            COLUMNS.map((c) => {
                if (c.type === "text") return r[c.key] || "";
                return Number(r[c.key]) || 0;
            })
        );

        const footRow = COLUMNS.map((c, i) => {
            if (i === 0) return "Totals";
            if (c.type === "text") return "";
            return totals[c.key];
        });

        const aoa = [
            [`${TOWERS[estate]} Unit-wise Billing Report`],
            [
                `Generated On: ${generatedOn}`,
                ...Array(COLUMNS.length - 2).fill(""),
                `Generated By: ${generatedBy}`,
            ],
            [],
            header,
            ...rows,
            footRow,
        ];

        const ws = XLSX.utils.aoa_to_sheet(aoa);
        ws["!cols"] = COLUMNS.map((c) => ({
            wch: c.type === "text" ? 25 : 14,
        }));
        ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: COLUMNS.length - 1 } }];

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Unit Billing");
        XLSX.writeFile(wb, `UnitWiseBilling_${TOWERS[estate].replace(/\s+/g, "_")}_${generatedOn.replace(/\//g, "-")}.xlsx`);
    }, [filteredData, estate, generatedBy, generatedOn, totals]);

    return (
        <div style={styles.app}>
            {/* Toolbar */}
            <div style={styles.toolbar}>
                <div style={styles.controls}>
                    <div style={styles.field}>
                        <label style={styles.label}>Estate Name</label>
                        <Select
                            value={{
                                label: TOWERS[estate],
                                value: estate,
                            }}
                            onChange={(selected) => setEstate(selected.value)}
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
                    <button style={styles.btnPrimary} onClick={handleGenerate} disabled={loading}>
                        {loading ? "Loading…" : "⟳ Generate Report"}
                    </button>
                </div>
            </div>

            {/* Export bar */}
            <div style={styles.exportBar}>
                <div style={styles.statusLine}>
                    Showing live data · {filteredData.length} unit{filteredData.length === 1 ? "" : "s"}
                </div>
                <div style={styles.exportActions}>
                    <button style={styles.btnExport} onClick={handleExportXlsx}>
                        Export Excel
                    </button>
                </div>
            </div>

            {/* Report sheet / HTML preview */}
            <div style={styles.sheet}>
                <div style={styles.sheetInner}>
                    <div style={styles.reportTitle}>
                        {`${TOWERS[estate]} Unit-wise Billing Report`}
                    </div>

                    <div style={styles.reportMeta}>
                        <span>
                            Generated On: <b>{generatedOn}</b>
                        </span>
                        <span>
                            Generated By: <b>{generatedBy}</b>
                        </span>
                    </div>

                    {filteredData.length === 0 ? (
                        <div style={styles.empty}>
                            No billing data for the selected estate.
                        </div>
                    ) : (
                        <div style={styles.tableWrapper}>
                            <table style={styles.table}>
                                <thead>
                                    <tr>
                                        {COLUMNS.map((c) => (
                                            <th
                                                key={c.key}
                                                style={{
                                                    ...styles.th,
                                                    textAlign: c.type === "text" ? "left" : "right",
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
                                            key={r.InvoiceNumber || i}
                                            style={i % 2 === 1 ? styles.rowAlt : undefined}
                                        >
                                            {COLUMNS.map((c) => (
                                                <td
                                                    key={c.key}
                                                    style={{
                                                        ...styles.td,
                                                        textAlign: c.type === "text" ? "left" : "right",
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
                                        {COLUMNS.map((c, i) => (
                                            <td
                                                key={c.key}
                                                style={{
                                                    ...styles.tfootTd,
                                                    textAlign: c.type === "text" ? "left" : "right",
                                                }}
                                            >
                                                {i === 0
                                                    ? "Totals"
                                                    : c.type === "text"
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
