import React, { useState, useEffect, useCallback } from 'react';
import './Compliance.css';

const API_BASE = '/server/compliance_function';

const Compliance = ({ userRole, userEmail }) => {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [message, setMessage] = useState('');
  const [modalOpen, setModalOpen] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [form, setForm] = useState({
    DateofAuditFrom: '',
    PeriodforAuditFrom: '',
    ClientName: '',
    ClientEmailAddress: '',
    Act: '',
    DateofAuditTo: '',
    PeriodforAuditTo: '',
    Address: '',
    ClientDetails: '',
    RecordCategory: '',
    MaximumMarks: '',
    Applicability: '',
    MarkObtained: ''
  });
  const [uploading, setUploading] = useState(false);

  const fetchData = useCallback(async () => {
    setLoading(true);
    setMessage('');
    try {
      const res = await fetch(`${API_BASE}/compliance`);
      const json = await res.json();
      if (json.status === 'success' && Array.isArray(json.data?.complianceRecords)) {
        setData(json.data.complianceRecords);
      } else {
        setData([]);
        setMessage(json.message || 'Failed to load compliance records');
      }
    } catch (err) {
      setData([]);
      setMessage('Failed to load compliance records: ' + (err.message || 'Unknown error'));
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchData();
  }, [fetchData]);

  const resetForm = () => {
    setForm({
      DateofAuditFrom: '',
      PeriodforAuditFrom: '',
      ClientName: '',
      ClientEmailAddress: '',
      Act: '',
      DateofAuditTo: '',
      PeriodforAuditTo: '',
      Address: '',
      ClientDetails: '',
      RecordCategory: '',
      MaximumMarks: '',
      Applicability: '',
      MarkObtained: ''
    });
    setEditingId(null);
    setModalOpen(false);
  };

  const openAdd = () => {
    resetForm();
    setModalOpen(true);
  };

  const openEdit = (row) => {
    setForm({
      DateofAuditFrom: row.DateofAuditFrom || '',
      PeriodforAuditFrom: row.PeriodforAuditFrom || '',
      ClientName: row.ClientName || '',
      ClientEmailAddress: row.ClientEmailAddress || '',
      Act: row.Act || '',
      DateofAuditTo: row.DateofAuditTo || '',
      PeriodforAuditTo: row.PeriodforAuditTo || '',
      Address: row.Address || '',
      ClientDetails: row.ClientDetails || '',
      RecordCategory: row.RecordCategory || '',
      MaximumMarks: row.MaximumMarks || '',
      Applicability: row.Applicability || '',
      MarkObtained: row.MarkObtained || ''
    });
    setEditingId(row.id);
    setModalOpen(true);
  };

  const handleChange = (e) => {
    const { name, value } = e.target;
    setForm(prev => ({ ...prev, [name]: value }));
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (!(form.ClientName || '').trim()) {
      setMessage('Client Name is required');
      return;
    }
    if (!(form.Act || '').trim()) {
      setMessage('Act is required');
      return;
    }
    setUploading(true);
    setMessage('');
    try {
      const payload = {
        ...form,
        DateofAuditFrom: form.DateofAuditFrom || null,
        PeriodforAuditFrom: form.PeriodforAuditFrom || null,
        ClientEmailAddress: form.ClientEmailAddress || null,
        DateofAuditTo: form.DateofAuditTo || null,
        PeriodforAuditTo: form.PeriodforAuditTo || null,
        Address: form.Address || null,
        ClientDetails: form.ClientDetails || null,
        RecordCategory: form.RecordCategory || null,
        MaximumMarks: form.MaximumMarks || null,
        Applicability: form.Applicability || null,
        MarkObtained: form.MarkObtained || null
      };

      const url = editingId ? `${API_BASE}/compliance/${editingId}` : `${API_BASE}/compliance`;
      const method = editingId ? 'PUT' : 'POST';
      const res = await fetch(url, {
        method,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      const json = await res.json();
      if (json.status === 'success') {
        setMessage(editingId ? 'Updated successfully' : 'Added successfully');
        resetForm();
        await fetchData();
      } else {
        setMessage(json.message || 'Operation failed');
      }
    } catch (err) {
      setMessage('Error: ' + (err.message || 'Unknown error'));
    } finally {
      setUploading(false);
    }
  };

  const handleDelete = async (id) => {
    if (!window.confirm('Delete this compliance record?')) return;
    setMessage('');
    try {
      const res = await fetch(`${API_BASE}/compliance/${id}`, { method: 'DELETE' });
      const json = await res.json();
      if (json.status === 'success') {
        setMessage('Deleted successfully');
        await fetchData();
      } else {
        setMessage(json.message || 'Delete failed');
      }
    } catch (err) {
      setMessage('Delete failed: ' + (err.message || 'Unknown error'));
    }
  };

  return (
    <div className="comp-page">
      <div className="comp-card-container">
        <div className="comp-header-row">
          <div className="comp-header-left">
            <div className="comp-title-section">
              <h2 className="comp-section-title">Compliance</h2>
              <p className="comp-section-subtitle">Manage audit compliance records, client details and marks</p>
            </div>
          </div>
          <div className="comp-actions-card">
            {message && <span className="comp-status-msg">{message}</span>}
            <button type="button" className="comp-action-btn comp-btn-add" onClick={openAdd} title="Add Record">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <line x1="12" y1="5" x2="12" y2="19" /><line x1="5" y1="12" x2="19" y2="12" />
              </svg>
            </button>
          </div>
        </div>

        <div className="comp-table-container">
          {loading && !data.length ? (
            <p className="comp-loading">Loading...</p>
          ) : (
            <table className="comp-table">
              <thead>
                <tr>
                  <th>S.No</th>
                  <th>Client Name</th>
                  <th>Client Email</th>
                  <th>Act</th>
                  <th>Audit From</th>
                  <th>Audit To</th>
                  <th>Period From</th>
                  <th>Period To</th>
                  <th>Address</th>
                  <th>Record Category</th>
                  <th>Max Marks</th>
                  <th>Marks Obtained</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>
                {data.length === 0 ? (
                  <tr>
                    <td colSpan={13} className="comp-empty">
                      No compliance records. Click &quot;+&quot; to add one.
                    </td>
                  </tr>
                ) : (
                  data.map((row, i) => (
                    <tr key={row.id || i}>
                      <td>{i + 1}</td>
                      <td>{row.ClientName}</td>
                      <td>{row.ClientEmailAddress || '—'}</td>
                      <td>{row.Act}</td>
                      <td>{row.DateofAuditFrom || '—'}</td>
                      <td>{row.DateofAuditTo || '—'}</td>
                      <td>{row.PeriodforAuditFrom || '—'}</td>
                      <td>{row.PeriodforAuditTo || '—'}</td>
                      <td className="comp-cell-address" title={row.Address}>{row.Address || '—'}</td>
                      <td>{row.RecordCategory || '—'}</td>
                      <td>{row.MaximumMarks ?? '—'}</td>
                      <td>{row.MarkObtained ?? '—'}</td>
                      <td>
                        <button type="button" className="comp-row-btn comp-row-btn-edit" onClick={() => openEdit(row)}>Edit</button>
                        <button type="button" className="comp-row-btn comp-row-btn-delete" onClick={() => handleDelete(row.id)}>Delete</button>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          )}
          <div className="comp-table-footer">
            Showing {data.length} compliance record(s)
          </div>
        </div>
      </div>

      {modalOpen && (
        <div className="comp-modal-overlay" onClick={resetForm}>
          <div className="comp-forms-wrap" onClick={e => e.stopPropagation()}>
            <div className="comp-card">
              <div className="comp-card-header">
                <h2>{editingId ? 'Edit Compliance' : 'Add Compliance'}</h2>
                <button type="button" className="comp-close-btn" onClick={resetForm} title="Close">&times;</button>
              </div>
              <form className="comp-form" onSubmit={handleSubmit}>
                <div className="comp-form-row">
                  <div className="comp-field">
                    <label htmlFor="comp-clientName">Client Name <span className="comp-req">*</span></label>
                    <input id="comp-clientName" name="ClientName" className="comp-input" value={form.ClientName} onChange={handleChange} required placeholder="Enter client name" />
                  </div>
                  <div className="comp-field">
                    <label htmlFor="comp-clientEmail">Client Email</label>
                    <input id="comp-clientEmail" name="ClientEmailAddress" type="email" className="comp-input" value={form.ClientEmailAddress} onChange={handleChange} placeholder="Enter client email" />
                  </div>
                </div>
                <div className="comp-form-row">
                  <div className="comp-field">
                    <label htmlFor="comp-act">Act <span className="comp-req">*</span></label>
                    <input id="comp-act" name="Act" className="comp-input" value={form.Act} onChange={handleChange} required placeholder="Enter act" />
                  </div>
                  <div className="comp-field">
                    <label htmlFor="comp-recordCategory">Record Category</label>
                    <input id="comp-recordCategory" name="RecordCategory" className="comp-input" value={form.RecordCategory} onChange={handleChange} placeholder="Enter record category" />
                  </div>
                </div>
                <div className="comp-form-row">
                  <div className="comp-field">
                    <label htmlFor="comp-dateFrom">Date of Audit From</label>
                    <input id="comp-dateFrom" name="DateofAuditFrom" type="date" className="comp-input" value={form.DateofAuditFrom} onChange={handleChange} />
                  </div>
                  <div className="comp-field">
                    <label htmlFor="comp-dateTo">Date of Audit To</label>
                    <input id="comp-dateTo" name="DateofAuditTo" type="date" className="comp-input" value={form.DateofAuditTo} onChange={handleChange} />
                  </div>
                </div>
                <div className="comp-form-row">
                  <div className="comp-field">
                    <label htmlFor="comp-periodFrom">Period for Audit From</label>
                    <input id="comp-periodFrom" name="PeriodforAuditFrom" className="comp-input" value={form.PeriodforAuditFrom} onChange={handleChange} placeholder="e.g. Jan 2024" />
                  </div>
                  <div className="comp-field">
                    <label htmlFor="comp-periodTo">Period for Audit To</label>
                    <input id="comp-periodTo" name="PeriodforAuditTo" className="comp-input" value={form.PeriodforAuditTo} onChange={handleChange} placeholder="e.g. Dec 2024" />
                  </div>
                </div>
                <div className="comp-row-full">
                  <label htmlFor="comp-address">Address</label>
                  <input id="comp-address" name="Address" className="comp-input" value={form.Address} onChange={handleChange} placeholder="Enter address" />
                </div>
                <div className="comp-row-full">
                  <label htmlFor="comp-clientDetails">Client Details</label>
                  <textarea id="comp-clientDetails" name="ClientDetails" className="comp-input comp-textarea" value={form.ClientDetails} onChange={handleChange} placeholder="Enter client details" rows={2} />
                </div>
                <div className="comp-form-row">
                  <div className="comp-field">
                    <label htmlFor="comp-maxMarks">Maximum Marks</label>
                    <input id="comp-maxMarks" name="MaximumMarks" type="number" className="comp-input" value={form.MaximumMarks} onChange={handleChange} placeholder="Max marks" />
                  </div>
                  <div className="comp-field">
                    <label htmlFor="comp-markObtained">Mark Obtained</label>
                    <input id="comp-markObtained" name="MarkObtained" type="number" className="comp-input" value={form.MarkObtained} onChange={handleChange} placeholder="Marks obtained" />
                  </div>
                </div>
                <div className="comp-row-full">
                  <label htmlFor="comp-applicability">Applicability</label>
                  <input id="comp-applicability" name="Applicability" className="comp-input" value={form.Applicability} onChange={handleChange} placeholder="Enter applicability" />
                </div>
                <div className="comp-actions">
                  <button type="submit" className="comp-btn comp-btn-primary" disabled={uploading}>
                    {uploading ? (editingId ? 'Updating...' : 'Submitting...') : (editingId ? 'Update' : 'Submit')}
                  </button>
                  <button type="button" className="comp-btn comp-btn-secondary" onClick={resetForm} disabled={uploading}>Cancel</button>
                </div>
              </form>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default Compliance;
