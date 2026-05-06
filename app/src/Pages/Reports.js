import React, { useState, useEffect, useCallback } from 'react';
import './Reports.css';

const STATUTORY_API = '/server/statutoryreg_function/statutory';
const CHECKLIST_API = '/server/checklist_function/checklist';

const Reports = ({ userRole, userEmail }) => {
  const [loading, setLoading] = useState(true);
  const [message, setMessage] = useState('');
  const [stats, setStats] = useState({ completed: 0, pending: 0, reject: 0 });
  const [submittedForms, setSubmittedForms] = useState([]);

  const hasProof = (item) =>
    item.proofSubmissionFile &&
    item.proofSubmissionFile !== 'null' &&
    item.proofSubmissionFile !== null &&
    String(item.proofSubmissionFile).trim() !== '';

  const hasForm = (item) =>
    item.formFile &&
    item.formFile !== 'null' &&
    item.formFile !== null &&
    String(item.formFile).trim() !== '' &&
    !String(item.formFile).startsWith('STATUTORY_MASTER_FORM:');

  const getFileUrl = (id, docType) => {
    const rawId = String(id || '').startsWith('checklist_') ? id.replace('checklist_', '') : id;
    return `/server/${String(id).startsWith('checklist_') ? 'checklist_function/checklist' : 'statutoryreg_function/statutory'}/${rawId}/file/${docType}?format=template&disposition=attachment`;
  };

  const fetchData = useCallback(async () => {
    setLoading(true);
    setMessage('');
    try {
      const [statRes, checklistRes] = await Promise.all([
        fetch(STATUTORY_API, { cache: 'no-store' }),
        fetch(CHECKLIST_API, { cache: 'no-store' })
      ]);

      let allItems = [];

      if (statRes.ok) {
        const statJson = await statRes.json();
        if (statJson.status === 'success' && statJson.data?.statutoryData) {
          allItems = statJson.data.statutoryData.map(item => ({
            ...item,
            formFile: item.formFile || item.FormFile || null,
            formFileName: item.formFileName || item.FormFileName || null,
            formName: item.formName || item.FormName || '',
            act: item.act || item.Act || '',
            proofSubmissionFile: item.proofSubmissionFile || item.ProofSubmissionFile || null,
            proofSubmissionFileName: item.proofSubmissionFileName || item.ProofSubmissionFileName || null,
            description: item.description || item.Description || '',
            isFromChecklist: false
          }));
        }
      }

      if (checklistRes.ok) {
        const clJson = await checklistRes.json();
        if (clJson.status === 'success' && clJson.data?.checklistData) {
          const clItems = clJson.data.checklistData
            .filter(item => hasForm(item))
            .map(item => ({
              id: `checklist_${item.id}`,
              formName: item.formName || '',
              formFile: item.formFile,
              formFileName: item.formFileName || null,
              act: item.act || '',
              description: item.description || '',
              proofSubmissionFile: item.proofSubmissionFile || null,
              proofSubmissionFileName: item.proofSubmissionFileName || null,
              isFromChecklist: true
            }));
          allItems = [...allItems, ...clItems];
        }
      }

      const completed = allItems.filter(item => hasForm(item) && hasProof(item)).length;
      const pending = allItems.filter(item => hasForm(item) && !hasProof(item)).length;
      const reject = allItems.filter(item => !hasForm(item)).length;

      setStats({ completed, pending, reject });
      setSubmittedForms(allItems.filter(item => hasForm(item)));
    } catch (err) {
      setMessage('Failed to load reports: ' + (err.message || 'Unknown error'));
      setStats({ completed: 0, pending: 0, reject: 0 });
      setSubmittedForms([]);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchData();
  }, [fetchData]);

  return (
    <div className="reports-page">
      <header className="reports-header">
        <h1 className="reports-title">Reports</h1>
        <p className="reports-subtitle">Statutory proof submission & submitted forms</p>
      </header>

      {/* 3 Status cards - white background, colored outer border/glow */}
      <section className="reports-cards-row">
        <div className="reports-card reports-card-completed">
          <div className="reports-card-icon">
            <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14" />
              <polyline points="22 4 12 14.01 9 11.01" />
            </svg>
          </div>
          <div className="reports-card-label">Completed</div>
          <p className="reports-card-desc">View all completed reports</p>
          <div className="reports-card-value">{loading ? '—' : stats.completed}</div>
        </div>
        <div className="reports-card reports-card-pending">
          <div className="reports-card-icon">
            <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <circle cx="12" cy="12" r="10" />
              <polyline points="12 6 12 12 16 14" />
            </svg>
          </div>
          <div className="reports-card-label">Pending</div>
          <p className="reports-card-desc">View all pending reports</p>
          <div className="reports-card-value">{loading ? '—' : stats.pending}</div>
        </div>
        <div className="reports-card reports-card-reject">
          <div className="reports-card-icon">
            <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <line x1="18" y1="6" x2="6" y2="18" />
              <line x1="6" y1="6" x2="18" y2="18" />
            </svg>
          </div>
          <div className="reports-card-label">Reject</div>
          <p className="reports-card-desc">View all rejected reports</p>
          <div className="reports-card-value">{loading ? '—' : stats.reject}</div>
        </div>
      </section>

      {message && <p className="reports-message">{message}</p>}

      {/* Submitted Forms section */}
      <section className="reports-submitted-section">
        <h2 className="reports-section-title">Submitted Forms</h2>
        <p className="reports-section-desc">Forms that have been submitted (with Form file uploaded). Proof submission status below.</p>

        {loading ? (
          <p className="reports-loading">Loading...</p>
        ) : submittedForms.length === 0 ? (
          <p className="reports-empty">No submitted forms.</p>
        ) : (
          <div className="reports-table-wrap">
            <table className="reports-table">
              <thead>
                <tr>
                  <th>S.No</th>
                  <th>Form Name</th>
                  <th>Act</th>
                  <th>Description</th>
                  <th>Form File</th>
                  <th>Proof Submission</th>
                  <th>Status</th>
                </tr>
              </thead>
              <tbody>
                {submittedForms.map((row, i) => (
                  <tr key={row.id || i}>
                    <td>{i + 1}</td>
                    <td>{row.formName || '—'}</td>
                    <td>{row.act || '—'}</td>
                    <td>{row.description || '—'}</td>
                    <td>
                      {row.formFile ? (
                        <a href={getFileUrl(row.id, 'Form')} target="_blank" rel="noopener noreferrer" className="reports-file-link">
                          {row.formFileName || 'View File'}
                        </a>
                      ) : (
                        '—'
                      )}
                    </td>
                    <td>
                      {hasProof(row) ? (
                        <a href={getFileUrl(row.id, 'ProofSubmission')} target="_blank" rel="noopener noreferrer" className="reports-file-link">
                          {row.proofSubmissionFileName || 'View File'}
                        </a>
                      ) : (
                        <span className="reports-status-pending">Pending</span>
                      )}
                    </td>
                    <td>
                      <span className={`reports-badge ${hasProof(row) ? 'reports-badge-completed' : 'reports-badge-pending'}`}>
                        {hasProof(row) ? 'Completed' : 'Pending'}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </section>
    </div>
  );
};

export default Reports;
