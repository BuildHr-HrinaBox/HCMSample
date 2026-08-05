import React, { useState, useEffect, useCallback } from 'react';
import './ClientDetails.css';

const API_BASE = '/server/clientdetails_function';

const ClientDetails = ({ userRole, userEmail }) => {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [message, setMessage] = useState('');
  const [modalOpen, setModalOpen] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [form, setForm] = useState({
    companyName: '',
    portalEmailID: '',
    addressLine1: '',
    addressLine2: '',
    city: '',
    state: '',
    pincode: '',
    clientCode: ''
  });
  const [logoFile, setLogoFile] = useState(null);
  const [logoPreview, setLogoPreview] = useState(null);
  const [uploading, setUploading] = useState(false);

  const fetchData = useCallback(async () => {
    setLoading(true);
    setMessage('');
    try {
      const res = await fetch(`${API_BASE}/clientdetails`);
      const json = await res.json();
      if (json.status === 'success' && Array.isArray(json.data?.clientDetails)) {
        setData(json.data.clientDetails);
      } else {
        setData([]);
        setMessage(json.message || 'Failed to load client details');
      }
    } catch (err) {
      setData([]);
      setMessage('Failed to load client details: ' + (err.message || 'Unknown error'));
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchData();
  }, [fetchData]);

  const resetForm = () => {
    setForm({
      companyName: '',
      portalEmailID: '',
      addressLine1: '',
      addressLine2: '',
      city: '',
      state: '',
      pincode: '',
      clientCode: ''
    });
    setLogoFile(null);
    setLogoPreview(null);
    setEditingId(null);
    setModalOpen(false);
  };

  const openAdd = () => {
    resetForm();
    setModalOpen(true);
  };

  const openEdit = (row) => {
    setForm({
      companyName: row.companyName || '',
      portalEmailID: row.portalEmailID || '',
      addressLine1: row.addressLine1 || '',
      addressLine2: row.addressLine2 || '',
      city: row.city || '',
      state: row.state || '',
      pincode: row.pincode || '',
      clientCode: row.clientCode || ''
    });
    setEditingId(row.id);
    setLogoFile(null);
    setLogoPreview(row.logo ? getLogoUrl(row.id) : null);
    setModalOpen(true);
  };

  const handleChange = (e) => {
    const { name, value } = e.target;
    setForm(prev => ({ ...prev, [name]: value }));
  };

  const handleLogoChange = (e) => {
    const file = e.target.files?.[0];
    setLogoFile(file || null);
    if (file && file.type.startsWith('image/')) {
      const reader = new FileReader();
      reader.onload = () => setLogoPreview(reader.result);
      reader.readAsDataURL(file);
    } else {
      setLogoPreview(null);
    }
  };

  const fileToBase64 = (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  };

  const getLogoUrl = (id) => {
    return `${API_BASE}/clientdetails/${id}/file/Logo`;
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    const required = ['companyName', 'portalEmailID', 'addressLine1', 'city', 'state', 'pincode', 'clientCode'];
    for (const key of required) {
      if (!(form[key] || '').trim()) {
        setMessage(`${key.replace(/([A-Z])/g, ' $1').trim()} is required`);
        return;
      }
    }
    setUploading(true);
    setMessage('');
    try {
      if (editingId) {
        let logoValue = data.find(r => r.id === editingId)?.logo || null;
        if (logoFile) {
          const fd = new FormData();
          fd.append('file', logoFile);
          const uploadRes = await fetch(`${API_BASE}/clientdetails/${editingId}/upload/Logo`, {
            method: 'POST',
            body: fd
          });
          const uploadJson = await uploadRes.json();
          if (uploadJson.status === 'success' && uploadJson.fileId) {
            logoValue = uploadJson.fileId;
          }
        }
        const payload = { ...form, logo: logoValue };
        const res = await fetch(`${API_BASE}/clientdetails/${editingId}`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });
        const json = await res.json();
        if (json.status === 'success') {
          setMessage('Updated successfully');
          resetForm();
          await fetchData();
        } else {
          setMessage(json.message || 'Update failed');
        }
      } else {
        let logoBase64 = null;
        if (logoFile) {
          logoBase64 = await fileToBase64(logoFile);
        }
        const payload = { ...form, logo: logoBase64 };
        const res = await fetch(`${API_BASE}/clientdetails`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });
        const json = await res.json();
        if (json.status === 'success') {
          setMessage('Added successfully');
          resetForm();
          await fetchData();
        } else {
          setMessage(json.message || 'Add failed');
        }
      }
    } catch (err) {
      setMessage('Error: ' + (err.message || 'Unknown error'));
    } finally {
      setUploading(false);
    }
  };

  const handleDelete = async (id) => {
    if (!window.confirm('Delete this client?')) return;
    setMessage('');
    try {
      const res = await fetch(`${API_BASE}/clientdetails/${id}`, { method: 'DELETE' });
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
    <div className="client-details-page">
      <div className="cd-card-container">
        <div className="cd-header-row">
          <div className="cd-header-left">
            <div className="cd-title-section">
              <h2 className="cd-section-title">Client Details</h2>
              <p className="cd-section-subtitle">Manage client companies, portal emails and addresses</p>
            </div>
          </div>
          <div className="cd-actions-card">
            {message && <span className="cd-status-msg">{message}</span>}
            <button type="button" className="cd-action-btn cd-btn-add" onClick={openAdd} title="Add Client">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <line x1="12" y1="5" x2="12" y2="19" /><line x1="5" y1="12" x2="19" y2="12" />
              </svg>
            </button>
          </div>
        </div>

        <div className="cd-table-container">
          {loading && !data.length ? (
            <p className="cd-loading">Loading...</p>
          ) : (
            <table className="cd-table">
              <thead>
                <tr>
                  <th>S.No</th>
                  <th>Company Name</th>
                  <th>Portal Email</th>
                  <th>Client Code</th>
                  <th>Address</th>
                  <th>City</th>
                  <th>State</th>
                  <th>Pincode</th>
                  <th>Logo</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>
                {data.length === 0 ? (
                  <tr>
                    <td colSpan={10} className="cd-empty">
                      No client data. Click &quot;+&quot; to add one.
                    </td>
                  </tr>
                ) : (
                  data.map((row, i) => (
                    <tr key={row.id || i}>
                      <td>{i + 1}</td>
                      <td>{row.companyName}</td>
                      <td>{row.portalEmailID}</td>
                      <td>{row.clientCode}</td>
                      <td className="cd-cell-address" title={`${row.addressLine1 || ''} ${row.addressLine2 || ''}`.trim()}>
                        {row.addressLine1 || '—'}
                      </td>
                      <td>{row.city || '—'}</td>
                      <td>{row.state || '—'}</td>
                      <td>{row.pincode || '—'}</td>
                      <td>
                        {row.logo ? (
                          <a href={getLogoUrl(row.id)} target="_blank" rel="noopener noreferrer" className="cd-file-link">View</a>
                        ) : '—'}
                      </td>
                      <td>
                        <button type="button" className="cd-row-btn cd-row-btn-edit" onClick={() => openEdit(row)}>Edit</button>
                        <button type="button" className="cd-row-btn cd-row-btn-delete" onClick={() => handleDelete(row.id)}>Delete</button>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          )}
          <div className="cd-table-footer">
            Showing {data.length} client record(s)
          </div>
        </div>
      </div>

      {modalOpen && (
        <div className="cd-modal-overlay" onClick={resetForm}>
          <div className="cd-forms-wrap" onClick={e => e.stopPropagation()}>
            <div className="cd-card">
              <div className="cd-card-header">
                <h2>{editingId ? 'Edit Client' : 'Add Client'}</h2>
                <button type="button" className="cd-close-btn" onClick={resetForm} title="Close">&times;</button>
              </div>
              <form className="cd-form" onSubmit={handleSubmit}>
                <div className="cd-form-row">
                  <div className="cd-field">
                    <label htmlFor="cd-companyName">Company Name <span className="cd-req">*</span></label>
                    <input id="cd-companyName" name="companyName" className="cd-input" value={form.companyName} onChange={handleChange} required placeholder="Enter company name" />
                  </div>
                  <div className="cd-field">
                    <label htmlFor="cd-portalEmailID">Portal Email ID <span className="cd-req">*</span></label>
                    <input id="cd-portalEmailID" name="portalEmailID" type="email" className="cd-input" value={form.portalEmailID} onChange={handleChange} required placeholder="Enter portal email" />
                  </div>
                </div>
                <div className="cd-form-row">
                  <div className="cd-field">
                    <label htmlFor="cd-clientCode">Client Code <span className="cd-req">*</span></label>
                    <input id="cd-clientCode" name="clientCode" className="cd-input" value={form.clientCode} onChange={handleChange} required placeholder="Enter client code" />
                  </div>
                  <div className="cd-field">
                    <label htmlFor="cd-pincode">Pincode <span className="cd-req">*</span></label>
                    <input id="cd-pincode" name="pincode" className="cd-input" value={form.pincode} onChange={handleChange} required placeholder="Enter pincode" />
                  </div>
                </div>
                <div className="cd-row-full">
                  <label htmlFor="cd-addressLine1">Address Line 1 <span className="cd-req">*</span></label>
                  <input id="cd-addressLine1" name="addressLine1" className="cd-input" value={form.addressLine1} onChange={handleChange} required placeholder="Enter address line 1" />
                </div>
                <div className="cd-row-full">
                  <label htmlFor="cd-addressLine2">Address Line 2</label>
                  <input id="cd-addressLine2" name="addressLine2" className="cd-input" value={form.addressLine2} onChange={handleChange} placeholder="Enter address line 2" />
                </div>
                <div className="cd-form-row">
                  <div className="cd-field">
                    <label htmlFor="cd-city">City <span className="cd-req">*</span></label>
                    <input id="cd-city" name="city" className="cd-input" value={form.city} onChange={handleChange} required placeholder="Enter city" />
                  </div>
                  <div className="cd-field">
                    <label htmlFor="cd-state">State <span className="cd-req">*</span></label>
                    <input id="cd-state" name="state" className="cd-input" value={form.state} onChange={handleChange} required placeholder="Enter state" />
                  </div>
                </div>
                <div className="cd-row-full">
                  <label htmlFor="cd-logo">Logo</label>
                  {logoPreview && (
                    <div className="cd-logo-preview">
                      <img src={logoPreview} alt="Logo preview" />
                    </div>
                  )}
                  {editingId && data.find(r => r.id === editingId)?.logo && !logoFile && (
                    <div className="cd-logo-current">
                      <a href={getLogoUrl(editingId)} target="_blank" rel="noopener noreferrer">Current Logo</a>
                    </div>
                  )}
                  <input id="cd-logo" name="logo" type="file" accept="image/*" onChange={handleLogoChange} className="cd-file-input" />
                  <small className="cd-file-hint">Supported: PNG, JPG, GIF, WebP. Max 5MB.</small>
                </div>
                <div className="cd-actions">
                  <button type="submit" className="cd-btn cd-btn-primary" disabled={uploading}>
                    {uploading ? (editingId ? 'Updating...' : 'Submitting...') : (editingId ? 'Update' : 'Submit')}
                  </button>
                  <button type="button" className="cd-btn cd-btn-secondary" onClick={resetForm} disabled={uploading}>Cancel</button>
                </div>
              </form>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default ClientDetails;
