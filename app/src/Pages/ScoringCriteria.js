import React, { useState, useEffect, useCallback } from 'react';
import './ScoringCriteria.css';

const API_BASE = '/server/scoringcriteria_function';

const ScoringCriteria = ({ userRole, userEmail }) => {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [message, setMessage] = useState('');
  const [modalOpen, setModalOpen] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [form, setForm] = useState({
    levels: '',
    criteriaforscoring: ''
  });
  const [uploading, setUploading] = useState(false);

  const fetchData = useCallback(async () => {
    setLoading(true);
    setMessage('');
    try {
      const res = await fetch(`${API_BASE}/scoring-criteria`);
      const json = await res.json();
      if (json.status === 'success' && Array.isArray(json.data?.scoringCriteria)) {
        setData(json.data.scoringCriteria);
      } else {
        setData([]);
        setMessage(json.message || 'Failed to load scoring criteria');
      }
    } catch (err) {
      setData([]);
      setMessage('Failed to load scoring criteria: ' + (err.message || 'Unknown error'));
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchData();
  }, [fetchData]);

  const resetForm = () => {
    setForm({ levels: '', criteriaforscoring: '' });
    setEditingId(null);
    setModalOpen(false);
  };

  const openAdd = () => {
    resetForm();
    setModalOpen(true);
  };

  const openEdit = (row) => {
    setForm({
      levels: row.levels || '',
      criteriaforscoring: row.criteriaforscoring || ''
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
    if (!(form.criteriaforscoring || '').trim()) {
      setMessage('Criteria for scoring is required');
      return;
    }
    setUploading(true);
    setMessage('');
    try {
      const payload = {
        levels: form.levels?.trim() || null,
        criteriaforscoring: form.criteriaforscoring.trim()
      };
      const url = editingId ? `${API_BASE}/scoring-criteria/${editingId}` : `${API_BASE}/scoring-criteria`;
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
    if (!window.confirm('Delete this scoring criteria?')) return;
    setMessage('');
    try {
      const res = await fetch(`${API_BASE}/scoring-criteria/${id}`, { method: 'DELETE' });
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
    <div className="sc-page">
      <div className="sc-card-container">
        <div className="sc-header-row">
          <div className="sc-header-left">
            <div className="sc-title-section">
              <h2 className="sc-section-title">Scoring Criteria</h2>
              <p className="sc-section-subtitle">Manage levels and criteria for compliance scoring</p>
            </div>
          </div>
          <div className="sc-actions-card">
            {message && <span className="sc-status-msg">{message}</span>}
            <button type="button" className="sc-action-btn sc-btn-add" onClick={openAdd} title="Add Criteria">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <line x1="12" y1="5" x2="12" y2="19" /><line x1="5" y1="12" x2="19" y2="12" />
              </svg>
            </button>
          </div>
        </div>

        <div className="sc-table-container">
          {loading && !data.length ? (
            <p className="sc-loading">Loading...</p>
          ) : (
            <table className="sc-table">
              <thead>
                <tr>
                  <th>S.No</th>
                  <th>Levels</th>
                  <th>Criteria for Scoring</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>
                {data.length === 0 ? (
                  <tr>
                    <td colSpan={4} className="sc-empty">
                      No scoring criteria. Click &quot;+&quot; to add one.
                    </td>
                  </tr>
                ) : (
                  data.map((row, i) => (
                    <tr key={row.id || i}>
                      <td>{i + 1}</td>
                      <td>{row.levels || '—'}</td>
                      <td className="sc-cell-criteria" title={row.criteriaforscoring}>{row.criteriaforscoring || '—'}</td>
                      <td>
                        <button type="button" className="sc-row-btn sc-row-btn-edit" onClick={() => openEdit(row)}>Edit</button>
                        <button type="button" className="sc-row-btn sc-row-btn-delete" onClick={() => handleDelete(row.id)}>Delete</button>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          )}
          <div className="sc-table-footer">
            Showing {data.length} scoring criteria record(s)
          </div>
        </div>
      </div>

      {modalOpen && (
        <div className="sc-modal-overlay" onClick={resetForm}>
          <div className="sc-forms-wrap" onClick={e => e.stopPropagation()}>
            <div className="sc-card">
              <div className="sc-card-header">
                <h2>{editingId ? 'Edit Scoring Criteria' : 'Add Scoring Criteria'}</h2>
                <button type="button" className="sc-close-btn" onClick={resetForm} title="Close">&times;</button>
              </div>
              <form className="sc-form" onSubmit={handleSubmit}>
                <div className="sc-row-full">
                  <label htmlFor="sc-levels">Levels</label>
                  <input id="sc-levels" name="levels" className="sc-input" value={form.levels} onChange={handleChange} placeholder="e.g. Level 1, Level 2" />
                </div>
                <div className="sc-row-full">
                  <label htmlFor="sc-criteriaforscoring">Criteria for Scoring <span className="sc-req">*</span></label>
                  <textarea id="sc-criteriaforscoring" name="criteriaforscoring" className="sc-input sc-textarea" value={form.criteriaforscoring} onChange={handleChange} required placeholder="Enter criteria for scoring" rows={3} />
                </div>
                <div className="sc-actions">
                  <button type="submit" className="sc-btn sc-btn-primary" disabled={uploading}>
                    {uploading ? (editingId ? 'Updating...' : 'Submitting...') : (editingId ? 'Update' : 'Submit')}
                  </button>
                  <button type="button" className="sc-btn sc-btn-secondary" onClick={resetForm} disabled={uploading}>Cancel</button>
                </div>
              </form>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default ScoringCriteria;
