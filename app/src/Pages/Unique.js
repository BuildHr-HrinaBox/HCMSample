import React, { useState, useEffect, useCallback } from 'react';
import './Unique.css';

const API_BASE = '/server/unique_function';

const Unique = ({ userRole, userEmail }) => {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [message, setMessage] = useState('');
  const [modalOpen, setModalOpen] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [form, setForm] = useState({
    actUniqueId: '',
    actId: '',
    actCode: '',
    generatedCode: ''
  });
  const [uploading, setUploading] = useState(false);

  const fetchData = useCallback(async () => {
    setLoading(true);
    setMessage('');
    try {
      const res = await fetch(`${API_BASE}/unique-ids`);
      const json = await res.json();
      if (json.status === 'success' && Array.isArray(json.data?.uniqueIds)) {
        setData(json.data.uniqueIds);
      } else {
        setData([]);
        setMessage(json.message || 'Failed to load unique IDs');
      }
    } catch (err) {
      setData([]);
      setMessage('Failed to load unique IDs: ' + (err.message || 'Unknown error'));
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchData();
  }, [fetchData]);

  const resetForm = () => {
    setForm({
      actUniqueId: '',
      actId: '',
      actCode: '',
      generatedCode: ''
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
      actUniqueId: row.actUniqueId || '',
      actId: row.actId ?? '',
      actCode: row.actCode || '',
      generatedCode: row.generatedCode || ''
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
    if (!(form.actUniqueId || '').trim()) {
      setMessage('Act Unique ID is required');
      return;
    }
    setUploading(true);
    setMessage('');
    try {
      const payload = {
        actUniqueId: form.actUniqueId.trim(),
        actId: form.actId === '' ? null : form.actId,
        actCode: form.actCode?.trim() || null,
        generatedCode: form.generatedCode?.trim() || null
      };
      const url = editingId ? `${API_BASE}/unique-ids/${editingId}` : `${API_BASE}/unique-ids`;
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
    if (!window.confirm('Delete this unique ID?')) return;
    setMessage('');
    try {
      const res = await fetch(`${API_BASE}/unique-ids/${id}`, { method: 'DELETE' });
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
    <div className="uq-page">
      <div className="uq-card-container">
        <div className="uq-header-row">
          <div className="uq-header-left">
            <div className="uq-title-section">
              <h2 className="uq-section-title">Unique IDs</h2>
              <p className="uq-section-subtitle">Manage Act Unique IDs, Act Codes and Generated Codes</p>
            </div>
          </div>
          <div className="uq-actions-card">
            {message && <span className="uq-status-msg">{message}</span>}
            <button type="button" className="uq-action-btn uq-btn-add" onClick={openAdd} title="Add Unique ID">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <line x1="12" y1="5" x2="12" y2="19" /><line x1="5" y1="12" x2="19" y2="12" />
              </svg>
            </button>
          </div>
        </div>

        <div className="uq-table-container">
          {loading && !data.length ? (
            <p className="uq-loading">Loading...</p>
          ) : (
            <table className="uq-table">
              <thead>
                <tr>
                  <th>S.No</th>
                  <th>Act Unique ID</th>
                  <th>Act ID</th>
                  <th>Act Code</th>
                  <th>Generated Code</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>
                {data.length === 0 ? (
                  <tr>
                    <td colSpan={6} className="uq-empty">
                      No unique IDs. Click &quot;+&quot; to add one.
                    </td>
                  </tr>
                ) : (
                  data.map((row, i) => (
                    <tr key={row.id || i}>
                      <td>{i + 1}</td>
                      <td>{row.actUniqueId}</td>
                      <td>{row.actId ?? '—'}</td>
                      <td>{row.actCode || '—'}</td>
                      <td>{row.generatedCode || '—'}</td>
                      <td>
                        <button type="button" className="uq-row-btn uq-row-btn-edit" onClick={() => openEdit(row)}>Edit</button>
                        <button type="button" className="uq-row-btn uq-row-btn-delete" onClick={() => handleDelete(row.id)}>Delete</button>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          )}
          <div className="uq-table-footer">
            Showing {data.length} unique ID record(s)
          </div>
        </div>
      </div>

      {modalOpen && (
        <div className="uq-modal-overlay" onClick={resetForm}>
          <div className="uq-forms-wrap" onClick={e => e.stopPropagation()}>
            <div className="uq-card">
              <div className="uq-card-header">
                <h2>{editingId ? 'Edit Unique ID' : 'Add Unique ID'}</h2>
                <button type="button" className="uq-close-btn" onClick={resetForm} title="Close">&times;</button>
              </div>
              <form className="uq-form" onSubmit={handleSubmit}>
                <div className="uq-row-full">
                  <label htmlFor="uq-actUniqueId">Act Unique ID <span className="uq-req">*</span></label>
                  <input id="uq-actUniqueId" name="actUniqueId" className="uq-input" value={form.actUniqueId} onChange={handleChange} required placeholder="Enter Act Unique ID" />
                </div>
                <div className="uq-form-row">
                  <div className="uq-field">
                    <label htmlFor="uq-actId">Act ID</label>
                    <input id="uq-actId" name="actId" type="number" className="uq-input" value={form.actId} onChange={handleChange} placeholder="Act ID" />
                  </div>
                  <div className="uq-field">
                    <label htmlFor="uq-actCode">Act Code</label>
                    <input id="uq-actCode" name="actCode" className="uq-input" value={form.actCode} onChange={handleChange} placeholder="Act Code" />
                  </div>
                </div>
                <div className="uq-row-full">
                  <label htmlFor="uq-generatedCode">Generated Code</label>
                  <input id="uq-generatedCode" name="generatedCode" className="uq-input" value={form.generatedCode} onChange={handleChange} placeholder="Enter generated code" />
                </div>
                <div className="uq-actions">
                  <button type="submit" className="uq-btn uq-btn-primary" disabled={uploading}>
                    {uploading ? (editingId ? 'Updating...' : 'Submitting...') : (editingId ? 'Update' : 'Submit')}
                  </button>
                  <button type="button" className="uq-btn uq-btn-secondary" onClick={resetForm} disabled={uploading}>Cancel</button>
                </div>
              </form>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default Unique;
