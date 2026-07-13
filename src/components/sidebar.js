import React from 'react';

const NAV_ITEMS = [
  { id: 'tax-invoice', label: 'Tax Invoice' },
  { id: 'statement', label: 'Statement of Account' },
  { id: 'aging', label: 'Aging Report' },
  { id: 'noc', label: 'No Objection Certificate' },
  { id: 'collection', label: 'Collection Report' },
  { id: 'unitwise-billing', label: 'Unit-wise Billing Report' },
    { id: 'payment-summary', label: 'Unit-wise payment summary Report' },



];

export default function Sidebar({ activePage, onNavigate }) {
  return (
    <aside style={styles.sidebar}>
      <div style={styles.header}>Reports</div>
      <nav>
        {NAV_ITEMS.map((item) => (
          <button
            key={item.id}
            onClick={() => onNavigate(item.id)}
            style={{
              ...styles.navItem,
              ...(activePage === item.id ? styles.navItemActive : {}),
            }}
          >
            {item.label}
          </button>
        ))}
      </nav>
    </aside>
  );
}

const styles = {
  sidebar: {
    width: 220,
    minHeight: '100vh',
    borderRight: '1px solid #e2e2e2',
    padding: '16px 12px',
    boxSizing: 'border-box',
    fontFamily: 'sans-serif',
  },
  header: {
    fontSize: 13,
    fontWeight: 600,
    textTransform: 'uppercase',
    color: '#888',
    letterSpacing: 0.5,
    padding: '4px 8px 12px',
  },
  navItem: {
    display: 'block',
    width: '100%',
    textAlign: 'left',
    padding: '10px 12px',
    marginBottom: 4,
    border: 'none',
    borderRadius: 6,
    background: 'transparent',
    fontSize: 14,
    cursor: 'pointer',
    color: '#333',
  },
  navItemActive: {
    background: '#eef2ff',
    color: '#3b3bdb',
    fontWeight: 600,
  },
};