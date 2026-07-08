import React, { useState } from "react";
import jsPDF from "jspdf";


async function loadData() {
    try {
        const response = await fetch('/sample_noc.json');
        if (!response.ok) throw new Error('Network response was not ok');
        console.log(response[1])
        const data = await response.json()
        return data


    } catch (error) {
        console.error('Error reading JSON:', error);
    }

}
const SAMPLE_DATA = await loadData();


const parseDate = d => d || "";

const styles = {
    page: { fontFamily: "Arial,Helvetica,sans-serif", background: "#f3f3f3", minHeight: "100vh", padding: 30 },
    toolbar: { display: "flex", justifyContent: "flex-end", gap: 10, marginBottom: 20 },
    btn: { padding: "8px 16px", cursor: "pointer" },
    container: { width: "210mm", minHeight: "297mm", margin: "auto", background: "#fff", padding: "25mm", boxSizing: "border-box", boxShadow: "0 2px 10px rgba(0,0,0,.2)" },
    innerBox: { border: "2px solid #000", height: "220mm", padding: "20mm", display: "flex", flexDirection: "column", boxSizing: "border-box" },
    header: { display: "flex", justifyContent: "space-between" },
    logo: { width: 100 },
    title: { textAlign: "center", fontWeight: "bold", fontSize: 24, margin: "30px 0" },
    paragraph: { fontSize: 16, lineHeight: 2, textAlign: "justify", flex: 1 },
    signature: { marginTop: "auto", fontWeight: "bold", fontSize: 16 }
};

export default function NOCReport({ items = SAMPLE_DATA }) {
    const [selectedIndex, setSelectedIndex] = useState(0);
    const item = items[selectedIndex];
    const logoUrl = '/output.png'

    function exportPDF() {
        const doc = new jsPDF("p", "mm", "a4");
        const x = 20, y = 25, w = 170, h = 220;
        doc.rect(x, y, w, h);
        try { doc.addImage(logoUrl, "PNG", x + w - 35, y + 5, 25, 15); } catch (e) { }
        doc.setFontSize(11);
        doc.text("Date: " + parseDate(item.NOCGeneratedDate), x + 10, y + 12);
        doc.setFont("helvetica", "bold");
        doc.setFontSize(18);
        doc.text("No Objection Certificate", 105, y + 35, { align: "center" });
        doc.setFont("helvetica", "normal");
        doc.setFontSize(12);
        doc.setFontSize(12);

        let startX = 35;
        let currentY = y + 60;
        const lineHeight = 8;

        function write(text, bold = false) {
            doc.setFont("helvetica", bold ? "bold" : "normal");
            doc.text(text, startX, currentY);
            startX += doc.getTextWidth(text);
        }
        function drawMixedText(doc, segments, x, y, maxWidth, lineHeight = 7) {
            let cursorX = x;
            let cursorY = y;

            segments.forEach(({ text, bold }) => {
                doc.setFont("helvetica", bold ? "bold" : "normal");

                const words = text.split(" ");

                words.forEach((word, index) => {
                    const token = index === words.length - 1 ? word : word + " ";
                    const width = doc.getTextWidth(token);

                    if (cursorX + width > x + maxWidth) {
                        cursorX = x;
                        cursorY += lineHeight;
                    }

                    doc.text(token, cursorX, cursorY);
                    cursorX += width;
                });
            });

            return cursorY;
        } const endY = drawMixedText(
            doc,
            [
                { text: "This is to certify that  ", bold: false },
                { text: item.ConsumerName || "", bold: true },
                { text: ",  the ", bold: false },
                { text: item.ConsumerTypeName || "", bold: true },
                { text: "  of Unit No.  ", bold: false },
                { text: item.UnitName || "", bold: true },
                { text: " , Account No.  ", bold: false },
                { text: item.ConsumerNumber || "", bold: true },
                { text: "  in  ", bold: false },
                { text: item.EstateName || "", bold: true },
                { text: "  has cleared all outstanding utility charges up to  ", bold: false },
                { text: parseDate(item.ServiceAgreementStopDate), bold: true },
                { text: ".", bold: false },
            ],
            35,
            y + 60,
            140,
            7
        );

        doc.setFont("helvetica", "bold");
        doc.text("Sincerely,", 35, endY + 18); doc.save(`noc_${item.ConsumerNumber}.pdf`);
    }

    return (
        <div style={styles.page}>
            <div style={styles.toolbar}>
                <select value={selectedIndex} onChange={e => setSelectedIndex(Number(e.target.value))}>
                    {items.map((it, i) => <option key={i} value={i}>{it.ConsumerName} - {it.ConsumerNumber}</option>)}
                </select>
                <button style={styles.btn} onClick={exportPDF}>Export PDF</button>
            </div>
            <div style={styles.container}>
                <div style={styles.innerBox}>
                    <div style={styles.header}>
                        <div>Date: {parseDate(item.NOCGeneratedDate)}</div>
                        {logoUrl && <img src={logoUrl} alt="logo" style={styles.logo} />}
                    </div>
                    <div style={styles.title}>No Objection Certificate</div>
                    <div style={styles.paragraph}>
                        This is to certify that <b>{item.ConsumerName}</b>, the <b>{item.ConsumerTypeName}</b> of Unit No. <b>{item.UnitName}</b>, Account No. <b>{item.ConsumerNumber}</b> in <b>{item.EstateName}</b> has cleared all outstanding utility charges up to <b>{parseDate(item.ServiceAgreementStopDate)}</b>.
                    </div>
                    <div style={styles.signature}>Sincerely,</div>
                </div>
            </div>
        </div>);
}
