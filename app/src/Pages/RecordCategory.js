import React, { useState, useEffect, useCallback } from 'react';
import './RecordCategory.css';

const API_BASE = '/server/recordcategory_function';

const RecordCategory = ({ userRole, userEmail }) => {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [message, setMessage] = useState('');
  const [modalOpen, setModalOpen] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [form, setForm] = useState({ categoryName: '' });
  const [uploading, setUploading] = useState(false);

  const fetchData = useCallback(async () => {
    setLoading(true);
    setMessage('');
    try {
      const res = await fetch(`${API_BASE}/recordcategories`);
      const json = await res.json();
      if (json.status === 'success' && Array.isArray(json.data?.recordCategories)) {
        setData(json.data.recordCategories);
      } else {
        setData([]);
        setMessage(json.message || 'Failed to load record categories');
      }
    } catch (err) {
      setData([]);
      setMessage('Failed to load record categories: ' + (err.message || 'Unknown error'));
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchData();
  }, [fetchData]);

  const resetForm = () => {
    setForm({ categoryName: '' });
    setEditingId(null);
    setModalOpen(false);
  };

  const openAdd = () => {
    resetForm();
    setModalOpen(true);
  };

  const openEdit = (row) => {
    setForm({ categoryName: row.categoryName || '' });
    setEditingId(row.id);
    setModalOpen(true);
  };

  const handleChange = (e) => {
    const { name, value } = e.target;
    setForm(prev => ({ ...prev, [name]: value }));
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (!(form.categoryName || '').trim()) {
      setMessage('Category name is required');
      return;
    }
    setUploading(true);
    setMessage('');
    try {
      const payload = { categoryName: form.categoryName.trim(), description: '' };
      const url = editingId ? `${API_BASE}/recordcategories/${editingId}` : `${API_BASE}/recordcategories`;
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
    if (!window.confirm('Delete this record category?')) return;
    setMessage('');
    try {
      const res = await fetch(`${API_BASE}/recordcategories/${id}`, { method: 'DELETE' });
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
    <div className="rc-page">
      <div className="rc-card-container">
        <div className="rc-header-row">
          <div className="rc-header-left">
            <div className="rc-title-section">
              <h2 className="rc-section-title">Record Category</h2>
              <p className="rc-section-subtitle">Manage record categories for compliance classification</p>
            </div>
          </div>
          <div className="rc-actions-card">
            {message && <span className="rc-status-msg">{message}</span>}
            <button type="button" className="rc-action-btn rc-btn-add" onClick={openAdd} title="Add Category">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <line x1="12" y1="5" x2="12" y2="19" /><line x1="5" y1="12" x2="19" y2="12" />
              </svg>
            </button>
          </div>
        </div>

        <div className="rc-table-container">
          {loading && !data.length ? (
            <p className="rc-loading">Loading...</p>
          ) : (
            <table className="rc-table">
              <thead>
                <tr>
                  <th>S.No</th>
                  <th>Category Name</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>
                {data.length === 0 ? (
                  <tr>
                    <td colSpan={3} className="rc-empty">
                      No record categories. Click &quot;+&quot; to add one.
                    </td>
                  </tr>
                ) : (
                  data.map((row, i) => (
                    <tr key={row.id || i}>
                      <td>{i + 1}</td>
                      <td>{row.categoryName}</td>
                      <td>
                        <button type="button" className="rc-row-btn rc-row-btn-edit" onClick={() => openEdit(row)}>Edit</button>
                        <button type="button" className="rc-row-btn rc-row-btn-delete" onClick={() => handleDelete(row.id)}>Delete</button>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          )}
          <div className="rc-table-footer">
            Showing {data.length} record categor{data.length === 1 ? 'y' : 'ies'}
          </div>
        </div>
      </div>

      {modalOpen && (
        <div className="rc-modal-overlay" onClick={resetForm}>
          <div className="rc-forms-wrap" onClick={e => e.stopPropagation()}>
            <div className="rc-card">
              <div className="rc-card-header">
                <h2>{editingId ? 'Edit Record Category' : 'Add Record Category'}</h2>
                <button type="button" className="rc-close-btn" onClick={resetForm} title="Close">&times;</button>
              </div>
              <form className="rc-form" onSubmit={handleSubmit}>
                <div className="rc-row-full">
                  <label htmlFor="rc-categoryName">Category Name <span className="rc-req">*</span></label>
                  <input id="rc-categoryName" name="categoryName" className="rc-input" value={form.categoryName} onChange={handleChange} required placeholder="Enter category name" />
                </div>
                <div className="rc-actions">
                  <button type="submit" className="rc-btn rc-btn-primary" disabled={uploading}>
                    {uploading ? (editingId ? 'Updating...' : 'Submitting...') : (editingId ? 'Update' : 'Submit')}
                  </button>
                  <button type="button" className="rc-btn rc-btn-secondary" onClick={resetForm} disabled={uploading}>Cancel</button>
                </div>
              </form>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default RecordCategory;
