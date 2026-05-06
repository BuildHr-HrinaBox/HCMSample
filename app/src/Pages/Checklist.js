import React, { useState, useEffect, useCallback } from 'react';
import './Checklist.css';

const API_BASE = '/server/checklist_function';
const CHECKLISTBULK_API = '/server/checklistbulk_function/checklistbulk';

// Map checklistbulk row to table row shape (checklistbulk has no formFile/proof/compliance/approval/marks)
const mapChecklistBulkToRow = (item, index) => ({
  id: item.id || item.ROWID || `bulk-${index}`,
  formName: item.formName || item.FormName || '—',
  act: item.act || item.Act || '—',
  consultgovtdep: item.concernedGovtDepartment || item.ConcernedGovtDepartment || '—',
  dueDate: item.dueDate || item.DueDate || '—',
  description: item.description || item.Description || '—',
  sector: item.sector || item.Sector || '—',
  state: item.state || item.State || '—',
  formFile: null,
  formFileName: null,
  proofSubmissionFile: null,
  proofSubmissionFileName: null,
  complianceStatus: '',
  approvalStatus: '',
  marks: ''
});

const Checklist = ({ userRole, userEmail }) => {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [message, setMessage] = useState('');
  const [modalOpen, setModalOpen] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [form, setForm] = useState({
    formName: '',
    act: '',
    consultgovtdep: '',
    dueDate: '',
    description: '',
    sector: '',
    state: '',
    complianceStatus: '',
    approvalStatus: '',
    marks: ''
  });
  const [formFile, setFormFile] = useState(null);
  const [proofFile, setProofFile] = useState(null);
  const [uploading, setUploading] = useState(false);

  const fetchData = useCallback(async () => {
    setLoading(true);
    setMessage('');
    try {
      const res = await fetch(`${CHECKLISTBULK_API}?action=getAll`);
      const json = await res.json();
      if (json.status === 'success' && Array.isArray(json.data)) {
        setData(json.data.map((item, i) => mapChecklistBulkToRow(item, i)));
      } else {
        setData([]);
        setMessage(json.message || 'No checklist data from Checklist Master.');
      }
    } catch (err) {
      setData([]);
      setMessage('Failed to load checklist data: ' + (err.message || 'Unknown error'));
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchData();
  }, [fetchData]);

  const resetForm = () => {
    setForm({
      formName: '',
      act: '',
      consultgovtdep: '',
      dueDate: '',
      description: '',
      sector: '',
      state: '',
      complianceStatus: '',
      approvalStatus: '',
      marks: ''
    });
    setFormFile(null);
    setProofFile(null);
    setEditingId(null);
    setModalOpen(false);
  };

  const openAdd = () => {
    resetForm();
    setModalOpen(true);
  };

  const openEdit = (row) => {
    setForm({
      formName: row.formName || '',
      act: row.act || '',
      consultgovtdep: row.consultgovtdep || '',
      dueDate: row.dueDate || '',
      description: row.description || '',
      sector: row.sector || '',
      state: row.state || '',
      complianceStatus: row.complianceStatus || '',
      approvalStatus: row.approvalStatus || '',
      marks: row.marks || ''
    });
    setEditingId(row.id);
    setFormFile(null);
    setProofFile(null);
    setModalOpen(true);
  };

  const handleChange = (e) => {
    const { name, value } = e.target;
    setForm(prev => ({ ...prev, [name]: value }));
  };

  const uploadFile = async (file, docType) => {
    if (!file) return { fileId: null, fileName: null };
    const fd = new FormData();
    fd.append('file', file);
    const res = await fetch(`${API_BASE}/checklist/upload/${docType}`, {
      method: 'POST',
      body: fd
    });
    const json = await res.json();
    if (json.status === 'success') {
      return { fileId: json.fileId, fileName: json.fileName || file.name };
    }
    throw new Error(json.message || 'Upload failed');
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (!(form.formName || '').trim()) {
      setMessage('Form Name is required');
      return;
    }
    setUploading(true);
    setMessage('');
    try {
      let formFileId = null;
      let formFileName = null;
      let proofFileId = null;
      let proofFileName = null;

      if (editingId) {
        const row = data.find(r => r.id === editingId);
        formFileId = row?.formFile || null;
        formFileName = row?.formFileName || null;
        proofFileId = row?.proofSubmissionFile || null;
        proofFileName = row?.proofSubmissionFileName || null;
      }

      if (formFile) {
        const u = await uploadFile(formFile, 'Form');
        formFileId = u.fileId;
        formFileName = u.fileName;
      }
      if (proofFile) {
        const u = await uploadFile(proofFile, 'ProofSubmission');
        proofFileId = u.fileId;
        proofFileName = u.fileName;
      }

      const payload = {
        ...form,
        formFile: formFileId,
        formFileName,
        proofSubmissionFile: proofFileId,
        proofSubmissionFileName: proofFileName
      };

      const url = editingId ? `${API_BASE}/checklist/${editingId}` : `${API_BASE}/checklist`;
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
    if (!window.confirm('Delete this checklist item?')) return;
    setMessage('');
    try {
      const res = await fetch(`${API_BASE}/checklist/${id}`, { method: 'DELETE' });
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

  const getFileUrl = (id, docType, disposition = 'attachment') => {
    return `${API_BASE}/checklist/${id}/file/${docType}?disposition=${disposition}`;
  };

  return (
    <div className="cl-page">
      <div className="cl-card-container">
        <div className="cl-header-row">
          <div className="cl-header-left">
            <div className="cl-title-section">
              <h2 className="cl-section-title">Checklist</h2>
              <p className="cl-section-subtitle">Manage compliance checklist items with forms and proof submissions</p>
            </div>
          </div>
          <div className="cl-actions-card">
            {message && <span className="cl-status-msg">{message}</span>}
            <button type="button" className="cl-action-btn cl-btn-add" onClick={openAdd} title="Add Item">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <line x1="12" y1="5" x2="12" y2="19" />
                <line x1="5" y1="12" x2="19" y2="12" />
              </svg>
            </button>
          </div>
        </div>

        <div className="cl-table-container">
          {loading && !data.length ? (
            <p className="cl-loading">Loading...</p>
          ) : (
            <table className="cl-table">
              <thead>
                <tr>
                  <th>S.No</th>
                  <th>Form Name</th>
                  <th>Act</th>
                  <th>Consult Govt Dept</th>
                  <th>Due Date</th>
                  <th>Description</th>
                  <th>Form File</th>
                  <th>Proof File</th>
                  <th>Compliance</th>
                  <th>Approval</th>
                  <th>Marks</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>
                {data.length === 0 ? (
                  <tr>
                    <td colSpan={12} className="cl-empty">
                      No checklist data. Click &quot;Add Item&quot; to create one.
                    </td>
                  </tr>
                ) : (
                  data.map((row, i) => (
                    <tr key={row.id || i}>
                      <td>{i + 1}</td>
                      <td>{row.formName}</td>
                      <td>{row.act}</td>
                      <td>{row.consultgovtdep}</td>
                      <td>{row.dueDate}</td>
                      <td className="cl-cell-desc" title={row.description}>{row.description || '—'}</td>
                      <td>
                        {row.formFile ? (
                          <a href={getFileUrl(row.id, 'Form')} target="_blank" rel="noopener noreferrer" className="cl-file-link">
                            {row.formFileName || 'Form'}
                          </a>
                        ) : '—'}
                      </td>
                      <td>
                        {row.proofSubmissionFile ? (
                          <a href={getFileUrl(row.id, 'ProofSubmission')} target="_blank" rel="noopener noreferrer" className="cl-file-link">
                            {row.proofSubmissionFileName || 'Proof'}
                          </a>
                        ) : '—'}
                      </td>
                      <td>{row.complianceStatus || '—'}</td>
                      <td>{row.approvalStatus || '—'}</td>
                      <td>{row.marks ?? '—'}</td>
                      <td>
                        <button type="button" className="cl-row-btn cl-row-btn-edit" onClick={() => openEdit(row)}>Edit</button>
                        <button type="button" className="cl-row-btn cl-row-btn-delete" onClick={() => handleDelete(row.id)}>Delete</button>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          )}
          <div className="cl-table-footer">
            Showing {data.length} checklist record(s)
          </div>
        </div>
      </div>

      {modalOpen && (
        <div className="cl-item-modal-overlay" onClick={resetForm}>
          <div className="cl-item-forms-wrap" onClick={e => e.stopPropagation()}>
            <div className="cl-item-card">
              <div className="cl-item-card-header">
                <h2>{editingId ? 'Edit Checklist Item' : 'Add Checklist Item'}</h2>
                <button type="button" className="cl-item-close-btn" onClick={resetForm} title="Close">&times;</button>
              </div>
              <form className="cl-item-form" onSubmit={handleSubmit}>
                <div className="cl-item-form-row">
                  <div className="cl-item-field">
                    <label htmlFor="cl-formName">Form Name <span className="cl-item-req">*</span></label>
                    <input id="cl-formName" name="formName" className="cl-item-input" value={form.formName} onChange={handleChange} required placeholder="Enter form name" />
                  </div>
                  <div className="cl-item-field">
                    <label htmlFor="cl-act">Act</label>
                    <input id="cl-act" name="act" className="cl-item-input" value={form.act} onChange={handleChange} placeholder="Enter act" />
                  </div>
                </div>
                <div className="cl-item-form-row">
                  <div className="cl-item-field">
                    <label htmlFor="cl-consultgovtdep">Consult Govt Dept</label>
                    <input id="cl-consultgovtdep" name="consultgovtdep" className="cl-item-input" value={form.consultgovtdep} onChange={handleChange} placeholder="Concerned government department" />
                  </div>
                  <div className="cl-item-field">
                    <label htmlFor="cl-dueDate">Due Date</label>
                    <input id="cl-dueDate" name="dueDate" type="date" className="cl-item-input" value={form.dueDate} onChange={handleChange} />
                  </div>
                </div>
                <div className="cl-item-row">
                  <label htmlFor="cl-description">Description</label>
                  <textarea id="cl-description" name="description" className="cl-item-input" value={form.description} onChange={handleChange} placeholder="Enter description" rows={3} style={{ height: 'auto', minHeight: '80px', resize: 'vertical', padding: '12px 16px' }} />
                </div>
                <div className="cl-item-form-row">
                  <div className="cl-item-field">
                    <label htmlFor="cl-sector">Sector</label>
                    <input id="cl-sector" name="sector" className="cl-item-input" value={form.sector} onChange={handleChange} placeholder="Enter sector" />
                  </div>
                  <div className="cl-item-field">
                    <label htmlFor="cl-state">State</label>
                    <input id="cl-state" name="state" className="cl-item-input" value={form.state} onChange={handleChange} placeholder="Enter state" />
                  </div>
                </div>
                <div className="cl-item-form-row">
                  <div className="cl-item-field">
                    <label htmlFor="cl-formFile">Form File</label>
                    {editingId && data.find(r => r.id === editingId)?.formFile && (
                      <div className="cl-item-file-info">
                        <span>Current File: {data.find(r => r.id === editingId)?.formFileName || 'View File'}</span>
                        <a href={getFileUrl(editingId, 'Form')} target="_blank" rel="noopener noreferrer" className="cl-item-file-view">View</a>
                      </div>
                    )}
                    <input id="cl-formFile" name="formFile" type="file" accept=".pdf,.doc,.docx,.xls,.xlsx,.png,.jpg,.jpeg" onChange={e => setFormFile(e.target.files?.[0] || null)} className="cl-item-file-input" />
                    {formFile && <div className="cl-item-file-selected">New File Selected: {formFile.name}</div>}
                    <small className="cl-item-file-hint">Maximum file size: 5MB. Supported: PDF, DOC, DOCX, XLS, XLSX, PNG, JPG</small>
                  </div>
                  <div className="cl-item-field">
                    <label htmlFor="cl-proofFile">Proof Submission File</label>
                    {editingId && data.find(r => r.id === editingId)?.proofSubmissionFile && (
                      <div className="cl-item-file-info">
                        <span>Current File: {data.find(r => r.id === editingId)?.proofSubmissionFileName || 'View File'}</span>
                        <a href={getFileUrl(editingId, 'ProofSubmission')} target="_blank" rel="noopener noreferrer" className="cl-item-file-view">View</a>
                      </div>
                    )}
                    <input id="cl-proofFile" name="proofFile" type="file" accept=".pdf,.doc,.docx,.xls,.xlsx,.png,.jpg,.jpeg" onChange={e => setProofFile(e.target.files?.[0] || null)} className="cl-item-file-input" />
                    {proofFile && <div className="cl-item-file-selected">New File Selected: {proofFile.name}</div>}
                    <small className="cl-item-file-hint">Maximum file size: 5MB. Supported: PDF, DOC, DOCX, XLS, XLSX, PNG, JPG</small>
                  </div>
                </div>
                <div className="cl-item-form-row">
                  <div className="cl-item-field">
                    <label htmlFor="cl-complianceStatus">Compliance Status</label>
                    <input id="cl-complianceStatus" name="complianceStatus" className="cl-item-input" value={form.complianceStatus} onChange={handleChange} placeholder="e.g. Compliant" />
                  </div>
                  <div className="cl-item-field">
                    <label htmlFor="cl-approvalStatus">Approval Status</label>
                    <input id="cl-approvalStatus" name="approvalStatus" className="cl-item-input" value={form.approvalStatus} onChange={handleChange} placeholder="e.g. Approved" />
                  </div>
                </div>
                <div className="cl-item-form-row">
                  <div className="cl-item-field">
                    <label htmlFor="cl-marks">Marks</label>
                    <input id="cl-marks" name="marks" className="cl-item-input" value={form.marks} onChange={handleChange} placeholder="Enter marks" type="text" />
                  </div>
                  <div className="cl-item-field" />
                </div>
                <div className="cl-item-actions">
                  <button type="submit" className="cl-item-btn cl-item-btn-primary" disabled={uploading}>
                    {uploading ? (editingId ? 'Updating...' : 'Submitting...') : (editingId ? 'Update' : 'Submit')}
                  </button>
                  <button type="button" className="cl-item-btn cl-item-btn-secondary" onClick={resetForm} disabled={uploading}>Reset</button>
                </div>
              </form>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default Checklist;
