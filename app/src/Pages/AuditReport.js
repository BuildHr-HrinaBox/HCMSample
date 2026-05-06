import React, { useState, useEffect, useCallback } from 'react';
import { jsPDF } from 'jspdf';
import buildgrLogo from '../components/buildgr logo.png';
import './AuditReport.css';

const API_BASE = '/server/audit_function';

const AuditReport = ({ userRole, userEmail }) => {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [selectedMonth, setSelectedMonth] = useState('');
  const [months, setMonths] = useState([]);
  const [siteName, setSiteName] = useState('Delphi');

  const fetchData = useCallback(async () => {
    setLoading(true);
    try {
      const res = await fetch(`${API_BASE}/auditmaster?action=getAll`);
      const json = await res.json();
      if (json.status === 'success' && Array.isArray(json.data)) {
        setData(json.data);
        const uniqueMonths = [...new Set(json.data
          .filter(r => r.approvedForSite || r.createdTime)
          .map(r => {
            const d = r.createdTime ? new Date(r.createdTime) : new Date();
            return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
          })
        )].sort().reverse();
        if (uniqueMonths.length > 0 && !selectedMonth) {
          setSelectedMonth(uniqueMonths[0]);
        }
        setMonths(uniqueMonths.length > 0 ? uniqueMonths : [`${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, '0')}`]);
      } else {
        setData([]);
        const now = new Date();
        setMonths([`${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`]);
        setSelectedMonth(`${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`);
      }
    } catch {
      setData([]);
      const now = new Date();
      setMonths([`${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`]);
      setSelectedMonth(`${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchData();
  }, [fetchData]);

  useEffect(() => {
    if (months.length > 0 && !selectedMonth) setSelectedMonth(months[0]);
  }, [months, selectedMonth]);

  const formatMonthLabel = (ym) => {
    if (!ym) return '';
    const [y, m] = ym.split('-');
    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
    return `${monthNames[parseInt(m, 10) - 1]} ${y}`;
  };

  const filteredData = data.filter((r) => {
    const d = r.createdTime ? new Date(r.createdTime) : new Date();
    const ym = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
    return !selectedMonth || ym === selectedMonth;
  });

  const count = filteredData.length;

  const loadLogoAsBase64 = () => {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => {
        try {
          const canvas = document.createElement('canvas');
          canvas.width = img.width;
          canvas.height = img.height;
          const ctx = canvas.getContext('2d');
          ctx.drawImage(img, 0, 0);
          resolve(canvas.toDataURL('image/png'));
        } catch (e) {
          reject(e);
        }
      };
      img.onerror = () => reject(new Error('Logo load failed'));
      img.src = buildgrLogo;
    });
  };

  const generateAuditPDF = async () => {
    const doc = new jsPDF();
    const pageWidth = doc.internal.pageSize.getWidth();
    let y = 20;

    try {
      const logoData = await loadLogoAsBase64();
      doc.addImage(logoData, 'PNG', 14, 10, 30, 24);
    } catch {
      doc.setFontSize(16);
      doc.text('BuildHr', 14, 22);
    }

    doc.setFontSize(12);
    doc.text('BuildHr', 50, 20);
    doc.text('Teynapet, Chennai', 50, 28);

    doc.setFontSize(18);
    doc.text('Audit Report', pageWidth / 2, 45, { align: 'center' });

    doc.setFontSize(10);
    doc.text(`Period for Audit: ${formatMonthLabel(selectedMonth)}`, 14, 58);
    doc.text(`Date of Audit: ${new Date().toLocaleDateString()}`, 14, 65);

    y = 78;

    const cols = ['Act', 'Maximum Mark', 'Applicability', 'Percentage'];
    const colWidths = [50, 40, 45, 45];
    const startX = 14;

    doc.setFillColor(251, 191, 36);
    doc.rect(startX, y, colWidths.reduce((a, b) => a + b, 0), 8, 'F');
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(10);
    let x = startX;
    cols.forEach((col, i) => {
      doc.text(col, x + 2, y + 5.5);
      x += colWidths[i];
    });
    y += 8;

    doc.setFont('helvetica', 'normal');
    filteredData.forEach((row) => {
      if (y > 270) {
        doc.addPage();
        y = 20;
      }
      const maxM = row.maximumMarks || '-';
      const app = row.applicability || '-';
      const pct = (maxM !== '-' && app !== '-' && parseFloat(maxM) > 0)
        ? ((parseFloat(app) / parseFloat(maxM)) * 100).toFixed(2) + '%'
        : '-';
      const cells = [row.act || '-', maxM, app, pct];
      x = startX;
      cells.forEach((val, i) => {
        doc.setFontSize(9);
        doc.text(String(val).substring(0, 25), x + 2, y + 5);
        x += colWidths[i];
      });
      y += 7;
    });

    doc.save(`Audit_Report_${(selectedMonth || 'report').replace('-', '_')}.pdf`);
  };

  const generateCompliancePDF = async () => {
    const doc = new jsPDF();
    const pageWidth = doc.internal.pageSize.getWidth();
    let y = 20;

    try {
      const logoData = await loadLogoAsBase64();
      doc.addImage(logoData, 'PNG', 14, 10, 30, 24);
    } catch {
      doc.setFontSize(16);
      doc.text('BuildHr', 14, 22);
    }

    doc.setFontSize(12);
    doc.text('BuildHr', 50, 20);
    doc.text('Teynapet, Chennai', 50, 28);

    doc.setFontSize(18);
    doc.text('Compliance Assessment Report', pageWidth / 2, 45, { align: 'center' });

    doc.setFontSize(10);
    doc.text(`Period for Audit: ${formatMonthLabel(selectedMonth)}`, 14, 58);
    doc.text(`Date of Audit: ${new Date().toLocaleDateString()}`, 14, 65);

    y = 78;

    const cols = ['Act', 'Maximum Mark', 'Applicability', 'Percentage', 'Remarks'];
    const colWidths = [40, 35, 38, 38, 40];
    const startX = 14;

    doc.setFillColor(251, 191, 36);
    doc.rect(startX, y, colWidths.reduce((a, b) => a + b, 0), 8, 'F');
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(9);
    let x = startX;
    cols.forEach((col, i) => {
      doc.text(col, x + 2, y + 5.5);
      x += colWidths[i];
    });
    y += 8;

    doc.setFont('helvetica', 'normal');
    filteredData.forEach((row) => {
      if (y > 270) {
        doc.addPage();
        y = 20;
      }
      const maxM = row.maximumMarks || '-';
      const app = row.applicability || '-';
      const pct = (maxM !== '-' && app !== '-' && parseFloat(maxM) > 0)
        ? ((parseFloat(app) / parseFloat(maxM)) * 100).toFixed(2) + '%'
        : '-';
      const remarks = row.remarks || '-';
      const cells = [row.act || '-', maxM, app, pct, remarks];
      x = startX;
      cells.forEach((val, i) => {
        doc.setFontSize(8);
        doc.text(String(val).substring(0, 20), x + 2, y + 5);
        x += colWidths[i];
      });
      y += 7;
    });

    doc.save(`Compliance_Assessment_${(selectedMonth || 'report').replace('-', '_')}.pdf`);
  };

  const monthLabel = formatMonthLabel(selectedMonth);
  const monthSlug = (selectedMonth || '').replace('-', '_');

  return (
    <div className="audit-report-page">
      <div className="audit-report-header">
        <h1 className="audit-report-title">Audit Report</h1>
        <div className="audit-report-filter">
          <label htmlFor="audit-report-month">Filter by Month:</label>
          <select
            id="audit-report-month"
            className="audit-report-select"
            value={selectedMonth}
            onChange={(e) => setSelectedMonth(e.target.value)}
          >
            {months.map((m) => (
              <option key={m} value={m}>{formatMonthLabel(m)}</option>
            ))}
          </select>
        </div>
      </div>

      <section className="audit-report-section">
        <h2 className="audit-report-site">{siteName}:</h2>
        <div className="audit-report-cards">
          <div className="audit-report-card">
            <h3 className="audit-report-card-title">{monthSlug}_audit Report</h3>
            <p className="audit-report-card-count">{loading ? '...' : count} item(s)</p>
            <button type="button" className="audit-report-link" onClick={generateAuditPDF} disabled={loading}>
              {monthLabel}
            </button>
          </div>
          <div className="audit-report-card">
            <h3 className="audit-report-card-title">{monthSlug}_compliance Assessment</h3>
            <p className="audit-report-card-count">{loading ? '...' : count} item(s)</p>
            <button type="button" className="audit-report-link" onClick={generateCompliancePDF} disabled={loading}>
              {monthLabel}
            </button>
          </div>
        </div>
      </section>
    </div>
  );
};

export default AuditReport;
