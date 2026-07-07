import React, { useState } from 'react';
import Sidebar from './components/sidebar';
import TaxInvoiceReport from './reports/tax_invoice';
import StatementOfAccount from './reportPage';
import AgingReport from './reports/aging_report';

export default function App() {
  const [activePage, setActivePage] = useState('tax-invoice');

  return (
    <div style={{ display: 'flex' }}>
      <Sidebar activePage={activePage} onNavigate={setActivePage} />
      <main style={{ flex: 1, padding: 24 }}>
        {activePage === 'tax-invoice' && <TaxInvoiceReport />}
        {activePage === 'statement' && <StatementOfAccount />}
        {activePage === 'aging' && <AgingReport />}

      </main>
    </div>
  );
}