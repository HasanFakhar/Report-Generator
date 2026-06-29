import { useState, useEffect } from "react";
import StatementOfAccount from "./statement_of_account";


// ─── Styles ───────────────────────────────────────────────────────────────────
const S = {
  root: {
    fontFamily: "Arial, Helvetica, sans-serif",
    fontSize: "13px",
    color: "#222",
    minHeight: "100vh",
    background: "#f4f4f4",
  },

  /* ── Header bar ── */
  header: {
    background: "#fff",
    borderBottom: "1px solid #ccc",
    padding: "0 24px",
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    height: "56px",
  },
  headerLeft: {
    display: "flex",
    alignItems: "center",
    gap: "12px",
  },
  headerTitle: {
    fontSize: "16px",
    fontWeight: "bold",
    margin: 0,
    whiteSpace: "nowrap",
  },
  divider: {
    width: "1px",
    height: "28px",
    background: "#ddd",
  },

  /* ── Dropdown area ── */
  dropdownGroup: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  dropdownLabel: {
    fontSize: "12px",
    color: "#555",
    whiteSpace: "nowrap",
    fontWeight: "bold",
  },
  select: {
    border: "1px solid #999",
    borderRadius: "3px",
    padding: "5px 8px",
    fontSize: "13px",
    fontFamily: "Arial, Helvetica, sans-serif",
    color: "#222",
    background: "#fff",
    cursor: "pointer",
    minWidth: "220px",
  },

  /* ── Meta chip strip ── */
  chipStrip: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    fontSize: "11px",
    color: "#555",
  },
  chip: (active) => ({
    padding: "2px 8px",
    border: "1px solid " + (active ? "#2a7" : "#bbb"),
    borderRadius: "3px",
    background: active ? "#e6f9f0" : "#f7f7f7",
    color: active ? "#1a6a40" : "#777",
    fontWeight: active ? "bold" : "normal",
    fontSize: "11px",
  }),

  /* ── Page body ── */
  body: {
    padding: "24px",
  },

  /* ── Info card shown after selection ── */
  infoCard: {
    background: "#fff",
    border: "1px solid #ddd",
    borderRadius: "4px",
    padding: "14px 18px",
    marginBottom: "18px",
    display: "grid",
    gridTemplateColumns: "repeat(auto-fill, minmax(200px, 1fr))",
    gap: "10px 24px",
  },
  infoItem: {
    display: "flex",
    flexDirection: "column",
    gap: "2px",
  },
  infoItemLabel: {
    fontSize: "10px",
    color: "#888",
    textTransform: "uppercase",
    letterSpacing: "0.04em",
  },
  infoItemValue: {
    fontSize: "13px",
    color: "#222",
    fontWeight: "bold",
  },

  /* ── Placeholder ── */
  placeholder: {
    textAlign: "center",
    padding: "60px 20px",
    color: "#aaa",
    fontSize: "14px",
    background: "#fff",
    border: "1px solid #ddd",
    borderRadius: "4px",
  },
};

async function loadData() {
  try {
    const response = await fetch("/dummylist.json");
    if (!response.ok) throw new Error("Network response was not ok");
    const data = await response.json();
    return data;
  } catch (error) {
    console.error("Error reading JSON:", error);
  }
}

// ─── Fetch SOA transactions for a given serviceAgreementKey ──────────────────
// Replace with your real API call e.g.:
//   const res = await fetch(`/api/statement-of-account?key=${key}`);
//   return res.json();
async function fetchTransactions(serviceAgreementKey) {
  return [
  {
    "StatementOfAccountID": 9732,
    "ServiceAgreementKey": serviceAgreementKey,
    "TransactionDate": "2026-06-25T15:56:41.13",
    "Description": "Auto Generated - Invoice Number : RR16394 - Bill Period : 21/06/2026 - 21/06/2026",
    "Debit_Amount": 744,
    "Debit_AmountUOM": "AED",
    "Credit_Amount": 0,
    "Credit_AmountUOM": "AED",
    "RunningBalance": 744,
    "RunningBalanceUOM": "AED",
    "InvoiceDataID": 16394,
    "CollectionKey": null,
    "CollectionName": "Initial Invoice",
    "PaymentName": null,
    "ReferenceNumber": null,
    "CreatedOn": "2026-06-25T15:56:41.13"
  },
  {
    "StatementOfAccountID": 9735,
    "ServiceAgreementKey": serviceAgreementKey,
    "TransactionDate": "2026-06-26T09:58:15.513",
    "Description": "Auto Generated - Invoice Number : RR16397 - Bill Period : 21/06/2026 - 16/07/2026",
    "Debit_Amount": 158.09,
    "Debit_AmountUOM": "AED",
    "Credit_Amount": 0,
    "Credit_AmountUOM": "AED",
    "RunningBalance": 902.09,
    "RunningBalanceUOM": "AED",
    "InvoiceDataID": 16397,
    "CollectionKey": null,
    "CollectionName": "Regular Invoice",
    "PaymentName": null,
    "ReferenceNumber": null,
    "CreatedOn": "2026-06-26T09:58:15.513"
  },
  {
    "StatementOfAccountID": 9738,
    "ServiceAgreementKey": serviceAgreementKey,
    "TransactionDate": "2026-06-25T16:57:38",
    "Description": "Collection for Invoice Number : RR16394",
    "Debit_Amount": 0,
    "Debit_AmountUOM": "AED",
    "Credit_Amount": 744,
    "Credit_AmountUOM": "AED",
    "RunningBalance": 158.09,
    "RunningBalanceUOM": "AED",
    "InvoiceDataID": 16394,
    "CollectionKey": "A7610C2B-0723-4FD5-9C3A-17798C451537",
    "CollectionName": "Invoice Collection",
    "PaymentName": "Debit Card",
    "ReferenceNumber": "e97efbde-2d39-465b-8548-58d64f0e42e5",
    "CreatedOn": "2026-06-26T10:14:53.983"
  },
  {
    "StatementOfAccountID": 9741,
    "ServiceAgreementKey": serviceAgreementKey,
    "TransactionDate": "2026-06-26T10:47:26.64",
    "Description": "Auto Generated - Invoice Number : RR16400 - Bill Period : 17/07/2026 - 16/08/2026",
    "Debit_Amount": 183.11,
    "Debit_AmountUOM": "AED",
    "Credit_Amount": 0,
    "Credit_AmountUOM": "AED",
    "RunningBalance": 341.2,
    "RunningBalanceUOM": "AED",
    "InvoiceDataID": 16400,
    "CollectionKey": null,
    "CollectionName": "Regular Invoice",
    "PaymentName": null,
    "ReferenceNumber": null,
    "CreatedOn": "2026-06-26T10:47:26.64"
  },
  {
    "StatementOfAccountID": 9744,
    "ServiceAgreementKey": serviceAgreementKey,
    "TransactionDate": "2026-06-26T11:04:11",
    "Description": "Collection for Invoice Number : RR16400",
    "Debit_Amount": 0,
    "Debit_AmountUOM": "AED",
    "Credit_Amount": 341.2,
    "Credit_AmountUOM": "AED",
    "RunningBalance": 0,
    "RunningBalanceUOM": "AED",
    "InvoiceDataID": 16400,
    "CollectionKey": "CC0BE192-03EC-4D26-8DAD-8321AC15CEC1",
    "CollectionName": "Invoice Collection",
    "PaymentName": "Debit Card",
    "ReferenceNumber": "c8c2ed76-ea73-4061-8425-c4dfbd53c0ad",
    "CreatedOn": "2026-06-27T08:44:53.567"
  }
]
}

// ─── Header Component ─────────────────────────────────────────────────────────
export default function StatementOfAccountPage() {
  const [serviceAgreements, setServiceAgreements] = useState([]);
  const [agreementsLoading, setAgreementsLoading] = useState(true);
  const [agreementsError, setAgreementsError] = useState("");

  const [selectedKey, setSelectedKey] = useState("");
  const [selectedAgreement, setSelectedAgreement] = useState(null);
  const [transactions, setTransactions] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  // Load service agreements on mount
  useEffect(() => {
    setAgreementsLoading(true);
    loadData()
      .then((data) => {
        if (data) setServiceAgreements(data);
        else setAgreementsError("Failed to load service agreements.");
      })
      .catch(() => setAgreementsError("Failed to load service agreements."))
      .finally(() => setAgreementsLoading(false));
  }, []);

  useEffect(() => {
    if (!selectedKey) {
      setSelectedAgreement(null);
      setTransactions([]);
      return;
    }

    const agreement = serviceAgreements.find(
      (sa) => sa.serviceAgreementKey === selectedKey
    );
    setSelectedAgreement(agreement || null);
    setTransactions([]);
    setError("");
    setLoading(true);

    fetchTransactions(selectedKey)
      .then((data) => setTransactions(data))
      .catch(() => setError("Failed to load transactions. Please try again."))
      .finally(() => setLoading(false));
  }, [selectedKey,serviceAgreements]);

  return (
    <div style={S.root}>

      {/* ── Header bar ───────────────────────────────────────────────────── */}
      <header style={S.header}>
        <div style={S.headerLeft}>
          <h1 style={S.headerTitle}>Statement of Account</h1>
          <div style={S.divider} />
          <div style={S.dropdownGroup}>
            <span style={S.dropdownLabel}>Service Agreement:</span>
            {agreementsLoading ? (
              <span style={{ fontSize: "12px", color: "#888" }}>Loading agreements…</span>
            ) : agreementsError ? (
              <span style={{ fontSize: "12px", color: "#b91c1c" }}>{agreementsError}</span>
            ) : (
              <select
                style={S.select}
                value={selectedKey}
                onChange={(e) => setSelectedKey(e.target.value)}
              >
                <option value="">— Select an agreement —</option>
                {serviceAgreements.map((sa) => (
                  <option key={sa.serviceAgreementKey} value={sa.serviceAgreementKey}>
                    {sa.serviceAgreementNumber}
                  </option>
                ))}
              </select>
            )}
          </div>
        </div>

       
        
      </header>

      {/* ── Page body ────────────────────────────────────────────────────── */}
      <div style={S.body}>

        {/* Agreement meta info card */}
        {/* {selectedAgreement && (
          <div style={S.infoCard}>
            {[
              ["Unit", selectedAgreement.unitName],
              ["Consumer", selectedAgreement.consumerName],
              ["Agreement No", selectedAgreement.serviceAgreementNumber],
              ["Agreement Key", selectedAgreement.serviceAgreementKey],
              ["Contract Capacity", `${selectedAgreement.contractCapacity} kWh`],
              ["Bill Start Day", selectedAgreement.billStartDay],
              ["Billing Days", selectedAgreement.noOfBillingDays],
              ["Start Date", formatDate(selectedAgreement.startTime)],
              ["End Date", formatDate(selectedAgreement.stopTime)],
              ["Last Modified By", selectedAgreement.lastModifiedBy],
            ].map(([label, value]) => (
              <div key={label} style={S.infoItem}>
                <span style={S.infoItemLabel}>{label}</span>
                <span style={S.infoItemValue}>{value ?? "—"}</span>
              </div>
            ))}
          </div>
        )} */}

        {/* States: placeholder / loading / error / report */}
        {!selectedKey && (
          <div style={S.placeholder}>
            Select a service agreement above to view the statement of account.
          </div>
        )}

        {selectedKey && loading && (
          <div style={S.placeholder}>Loading transactions…</div>
        )}

        {selectedKey && !loading && error && (
          <div style={{ ...S.placeholder, color: "#b91c1c" }}>{error}</div>
        )}

        {selectedKey && !loading && !error && transactions.length > 0 && (
          <StatementOfAccount
            transactions={transactions}
            unitName={selectedAgreement?.unitName ?? ""}
            consumerName={selectedAgreement?.consumerName ?? ""}
            serviceAgreementNo={selectedAgreement?.serviceAgreementNumber ?? ""}
          />
        )}
      </div>

    </div>
  );
}