import React, { useState, useEffect, useCallback } from 'react';
import './States.css';

const API_BASE = '/server/states_function';

const States = ({ userRole, userEmail }) => {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [message, setMessage] = useState('');
  const [modalOpen, setModalOpen] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [form, setForm] = useState({ stateName: '' });
  const [uploading, setUploading] = useState(false);

  const fetchData = useCallback(async () => {
    setLoading(true);
    setMessage('');
    try {
      const res = await fetch(`${API_BASE}/states`);
      const json = await res.json();
      if (json.status === 'success' && Array.isArray(json.data?.states)) {
        setData(json.data.states);
      } else {
        setData([]);
        setMessage(json.message || 'Failed to load states');
      }
    } catch (err) {
      setData([]);
      setMessage('Failed to load states: ' + (err.message || 'Unknown error'));
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchData();
  }, [fetchData]);

  const resetForm = () => {
    setForm({ stateName: '' });
    setEditingId(null);
    setModalOpen(false);
  };

  const openAdd = () => {
    resetForm();
    setModalOpen(true);
  };

  const openEdit = (row) => {
    setForm({ stateName: row.stateName || '' });
    setEditingId(row.id);
    setModalOpen(true);
  };

  const handleChange = (e) => {
    const { name, value } = e.target;
    setForm(prev => ({ ...prev, [name]: value }));
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (!(form.stateName || '').trim()) {
      setMessage('State name is required');
      return;
    }
    setUploading(true);
    setMessage('');
    try {
      const payload = { stateName: form.stateName.trim() };
      const url = editingId ? `${API_BASE}/states/${editingId}` : `${API_BASE}/states`;
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
    if (!window.confirm('Delete this state?')) return;
    setMessage('');
    try {
      const res = await fetch(`${API_BASE}/states/${id}`, { method: 'DELETE' });
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
    <div className="st-page">
      <div className="st-card-container">
        <div className="st-header-row">
          <div className="st-header-left">
            <div className="st-title-section">
              <h2 className="st-section-title">States</h2>
              <p className="st-section-subtitle">Manage states for compliance and statutory records</p>
            </div>
          </div>
          <div className="st-actions-card">
            {message && <span className="st-status-msg">{message}</span>}
            <button type="button" className="st-action-btn st-btn-add" onClick={openAdd} title="Add State">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <line x1="12" y1="5" x2="12" y2="19" /><line x1="5" y1="12" x2="19" y2="12" />
              </svg>
            </button>
          </div>
        </div>

        <div className="st-table-container">
          {loading && !data.length ? (
            <p className="st-loading">Loading...</p>
          ) : (
            <table className="st-table">
              <thead>
                <tr>
                  <th>S.No</th>
                  <th>State Name</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>
                {data.length === 0 ? (
                  <tr>
                    <td colSpan={3} className="st-empty">
                      No states. Click &quot;+&quot; to add one.
                    </td>
                  </tr>
                ) : (
                  data.map((row, i) => (
                    <tr key={row.id || i}>
                      <td>{i + 1}</td>
                      <td>{row.stateName}</td>
                      <td>
                        <button type="button" className="st-row-btn st-row-btn-edit" onClick={() => openEdit(row)}>Edit</button>
                        <button type="button" className="st-row-btn st-row-btn-delete" onClick={() => handleDelete(row.id)}>Delete</button>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          )}
          <div className="st-table-footer">
            Showing {data.length} state(s)
          </div>
        </div>
      </div>

      {modalOpen && (
        <div className="st-modal-overlay" onClick={resetForm}>
          <div className="st-forms-wrap" onClick={e => e.stopPropagation()}>
            <div className="st-card">
              <div className="st-card-header">
                <h2>{editingId ? 'Edit State' : 'Add State'}</h2>
                <button type="button" className="st-close-btn" onClick={resetForm} title="Close">&times;</button>
              </div>
              <form className="st-form" onSubmit={handleSubmit}>
                <div className="st-row-full">
                  <label htmlFor="st-stateName">State Name <span className="st-req">*</span></label>
                  <input id="st-stateName" name="stateName" className="st-input" value={form.stateName} onChange={handleChange} required placeholder="Enter state name" />
                </div>
                <div className="st-actions">
                  <button type="submit" className="st-btn st-btn-primary" disabled={uploading}>
                    {uploading ? (editingId ? 'Updating...' : 'Submitting...') : (editingId ? 'Update' : 'Submit')}
                  </button>
                  <button type="button" className="st-btn st-btn-secondary" onClick={resetForm} disabled={uploading}>Cancel</button>
                </div>
              </form>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default States;
