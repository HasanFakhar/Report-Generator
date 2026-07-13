import React, { useState } from 'react';
import Sidebar from './components/sidebar';
import TaxInvoiceReport from './reports/tax_invoice';
import StatementOfAccount from './reportPage';
import AgingReport from './reports/aging_report';
import NOCReport from './reports/noc';
import CollectionReport from './reports/collection_report';
import UnitWiseBillingReport from './reports/unitwisebilling.js';
import PaymentSummaryReport from './reports/paymentSummaryReport.js';
import TowerWiseBillingReport from './reports/towerwisebilling.js';


export default function App() {
  const [activePage, setActivePage] = useState('tax-invoice');

  return (
    <div style={{ display: 'flex' }}>
      <Sidebar activePage={activePage} onNavigate={setActivePage} />
      <main style={{ flex: 1, padding: 24 }}>
        {activePage === 'tax-invoice' && <TaxInvoiceReport />}
        {activePage === 'statement' && <StatementOfAccount />}
        {activePage === 'aging' && <AgingReport />}
        {activePage === 'noc' && <NOCReport />}
        {activePage === 'collection' && <CollectionReport />}
        {activePage === 'unitwise-billing' && <UnitWiseBillingReport />}
        {activePage === 'payment-summary' && <PaymentSummaryReport />}
        {activePage === 'tower-wise-billing' && <TowerWiseBillingReport />}

      </main>
    </div>
  );
}