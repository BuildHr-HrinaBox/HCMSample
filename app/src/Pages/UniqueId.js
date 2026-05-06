import React, { useState, useEffect, useCallback } from 'react';
import './UniqueId.css';

const API_BASE = '/server/uniqueid_function';

const UniqueId = ({ userRole, userEmail }) => {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [message, setMessage] = useState('');
  const [modalOpen, setModalOpen] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [form, setForm] = useState({ actUniqueId: '' });
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
    setForm({ actUniqueId: '' });
    setEditingId(null);
    setModalOpen(false);
  };

  const openAdd = () => {
    resetForm();
    setModalOpen(true);
  };

  const openEdit = (row) => {
    setForm({ actUniqueId: row.actUniqueId || '' });
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
      const payload = { actUniqueId: form.actUniqueId.trim() };
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
    <div className="uqi-page">
      <div className="uqi-card-container">
        <div className="uqi-header-row">
          <div className="uqi-header-left">
            <div className="uqi-title-section">
              <h2 className="uqi-section-title">Unique ID List</h2>
              <p className="uqi-section-subtitle">Manage unique IDs (Uniqueld table)</p>
            </div>
          </div>
          <div className="uqi-actions-card">
            {message && <span className="uqi-status-msg">{message}</span>}
            <button type="button" className="uqi-action-btn uqi-btn-add" onClick={openAdd} title="Add Unique ID">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <line x1="12" y1="5" x2="12" y2="19" /><line x1="5" y1="12" x2="19" y2="12" />
              </svg>
            </button>
          </div>
        </div>

        <div className="uqi-table-container">
          {loading && !data.length ? (
            <p className="uqi-loading">Loading...</p>
          ) : (
            <table className="uqi-table">
              <thead>
                <tr>
                  <th>S.No</th>
                  <th>Act Unique ID</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>
                {data.length === 0 ? (
                  <tr>
                    <td colSpan={3} className="uqi-empty">
                      No unique IDs. Click &quot;+&quot; to add one.
                    </td>
                  </tr>
                ) : (
                  data.map((row, i) => (
                    <tr key={row.id || i}>
                      <td>{i + 1}</td>
                      <td>{row.actUniqueId}</td>
                      <td>
                        <button type="button" className="uqi-row-btn uqi-row-btn-edit" onClick={() => openEdit(row)}>Edit</button>
                        <button type="button" className="uqi-row-btn uqi-row-btn-delete" onClick={() => handleDelete(row.id)}>Delete</button>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          )}
          <div className="uqi-table-footer">
            Showing {data.length} unique ID(s)
          </div>
        </div>
      </div>

      {modalOpen && (
        <div className="uqi-modal-overlay" onClick={resetForm}>
          <div className="uqi-forms-wrap" onClick={e => e.stopPropagation()}>
            <div className="uqi-card">
              <div className="uqi-card-header">
                <h2>{editingId ? 'Edit Unique ID' : 'Add Unique ID'}</h2>
                <button type="button" className="uqi-close-btn" onClick={resetForm} title="Close">&times;</button>
              </div>
              <form className="uqi-form" onSubmit={handleSubmit}>
                <div className="uqi-row-full">
                  <label htmlFor="uqi-actUniqueId">Act Unique ID <span className="uqi-req">*</span></label>
                  <input id="uqi-actUniqueId" name="actUniqueId" className="uqi-input" value={form.actUniqueId} onChange={handleChange} required placeholder="Enter Act Unique ID" />
                </div>
                <div className="uqi-actions">
                  <button type="submit" className="uqi-btn uqi-btn-primary" disabled={uploading}>
                    {uploading ? (editingId ? 'Updating...' : 'Submitting...') : (editingId ? 'Update' : 'Submit')}
                  </button>
                  <button type="button" className="uqi-btn uqi-btn-secondary" onClick={resetForm} disabled={uploading}>Cancel</button>
                </div>
              </form>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default UniqueId;
