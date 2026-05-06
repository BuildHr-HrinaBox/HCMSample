import React, { useState, useEffect, useCallback } from 'react';
import './Certificate.css';

const API_BASE = '/server/certificate_function';

const initialForm = {
  gstNo: '',
  gstDate: '',
  companyPANNo: '',
  companyPANDate: '',
  incorporationNo: '',
  incorporationDate: '',
  moaNo: '',
  moaDate: '',
  aoaNo: '',
  aoaDate: '',
  epfNo: '',
  epfDate: '',
  esiNo: '',
  esiDate: '',
  factoryLicenseNo: '',
  factoryLicenseDate: '',
  lafNo: '',
  lrfDate: '',
  ptNo: '',
  ptDate: ''
};

const toApiPayload = (form) => ({
  GSTNo: form.gstNo?.trim() || null,
  GSTDate: form.gstDate || null,
  CompanyPANNo: form.companyPANNo?.trim() || null,
  CompanyPANDate: form.companyPANDate || null,
  IncorporationNo: form.incorporationNo?.trim() || null,
  IncorporationDate: form.incorporationDate || null,
  MOANo: form.moaNo?.trim() || null,
  MOADate: form.moaDate || null,
  AOANo: form.aoaNo?.trim() || null,
  AOADate: form.aoaDate || null,
  EPFNo: form.epfNo?.trim() || null,
  EPFDate: form.epfDate || null,
  ESINo: form.esiNo?.trim() || null,
  ESIDate: form.esiDate || null,
  FactoryLicenseNo: form.factoryLicenseNo?.trim() || null,
  FactoryLicenseDate: form.factoryLicenseDate || null,
  LAFNo: form.lafNo?.trim() || null,
  LRFDate: form.lrfDate || null,
  PTNo: form.ptNo?.trim() || null,
  PTDate: form.ptDate || null
});

const certToForm = (cert) => ({
  gstNo: cert.GSTNo || '',
  gstDate: cert.GSTDate ? formatDateForInput(cert.GSTDate) : '',
  companyPANNo: cert.CompanyPANNo || '',
  companyPANDate: cert.CompanyPANDate ? formatDateForInput(cert.CompanyPANDate) : '',
  incorporationNo: cert.IncorporationNo || '',
  incorporationDate: cert.IncorporationDate ? formatDateForInput(cert.IncorporationDate) : '',
  moaNo: cert.MOANo || '',
  moaDate: cert.MOADate ? formatDateForInput(cert.MOADate) : '',
  aoaNo: cert.AOANo || '',
  aoaDate: cert.AOADate ? formatDateForInput(cert.AOADate) : '',
  epfNo: cert.EPFNo || '',
  epfDate: cert.EPFDate ? formatDateForInput(cert.EPFDate) : '',
  esiNo: cert.ESINo || '',
  esiDate: cert.ESIDate ? formatDateForInput(cert.ESIDate) : '',
  factoryLicenseNo: cert.FactoryLicenseNo || '',
  factoryLicenseDate: cert.FactoryLicenseDate ? formatDateForInput(cert.FactoryLicenseDate) : '',
  lafNo: cert.LAFNo || '',
  lrfDate: cert.LRFDate ? formatDateForInput(cert.LRFDate) : '',
  ptNo: cert.PTNo || '',
  ptDate: cert.PTDate ? formatDateForInput(cert.PTDate) : ''
});

function formatDateForInput(val) {
  if (!val) return '';
  const d = new Date(val);
  if (isNaN(d.getTime())) return '';
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

const Certificate = ({ userRole, userEmail }) => {
  const [certificates, setCertificates] = useState([]);
  const [loading, setLoading] = useState(true);
  const [message, setMessage] = useState('');
  const [showForm, setShowForm] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [form, setForm] = useState(initialForm);

  const fetchCertificates = useCallback(async () => {
    setLoading(true);
    setMessage('');
    try {
      const res = await fetch(`${API_BASE}/certificate`);
      const data = await res.json();
      if (data.status === 'success' && data.data && Array.isArray(data.data.certificateDetails)) {
        setCertificates(data.data.certificateDetails);
      } else {
        setCertificates([]);
      }
    } catch (err) {
      setMessage('Failed to load certificates');
      setCertificates([]);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchCertificates();
  }, [fetchCertificates]);

  const openAdd = () => {
    setEditingId(null);
    setForm(initialForm);
    setShowForm(true);
  };

  const openEdit = (cert) => {
    setEditingId(cert.id);
    setForm(certToForm(cert));
    setShowForm(true);
  };

  const closeForm = () => {
    setShowForm(false);
    setEditingId(null);
    setForm(initialForm);
  };

  const handleChange = (e) => {
    const { name, value } = e.target;
    setForm((prev) => ({ ...prev, [name]: value }));
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    setMessage('');
    try {
      const payload = toApiPayload(form);
      if (editingId) {
        const res = await fetch(`${API_BASE}/certificate/${editingId}`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });
        const data = await res.json();
        if (data.status === 'success') {
          setMessage('Certificate updated successfully');
          closeForm();
          fetchCertificates();
        } else {
          setMessage(data.message || 'Update failed');
        }
      } else {
        const res = await fetch(`${API_BASE}/certificate`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });
        const data = await res.json();
        if (data.status === 'success') {
          setMessage('Certificate added successfully');
          closeForm();
          fetchCertificates();
        } else {
          setMessage(data.message || 'Add failed');
        }
      }
    } catch (err) {
      setMessage(err.message || 'Request failed');
    }
  };

  const handleDelete = async (id) => {
    if (!window.confirm('Delete this certificate?')) return;
    setMessage('');
    try {
      const res = await fetch(`${API_BASE}/certificate/${id}`, { method: 'DELETE' });
      const data = await res.json();
      if (data.status === 'success') {
        setMessage('Certificate deleted');
        fetchCertificates();
      } else {
        setMessage(data.message || 'Delete failed');
      }
    } catch (err) {
      setMessage(err.message || 'Delete failed');
    }
  };

  const certLabel = (cert) => {
    if (cert.GSTNo) return `GST: ${cert.GSTNo}`;
    if (cert.CompanyPANNo) return `PAN: ${cert.CompanyPANNo}`;
    if (cert.IncorporationNo) return `Incorporation: ${cert.IncorporationNo}`;
    return `Certificate #${cert.id}`;
  };

  return (
    <div className={`certificate-management-page${showForm ? ' certificate-management-page-form-open' : ''}`}>
      {!showForm && (
        <header className="certificate-management-header">
          <div className="certificate-management-header-left">
            <div className="certificate-management-header-icon">
              <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
                <polyline points="14 2 14 8 20 8" />
                <path d="M16 13H8" /><path d="M16 17H8" /><path d="M10 9H8" />
              </svg>
            </div>
            <div>
              <h1 className="certificate-management-title">Certificate Management</h1>
              <p className="certificate-management-subtitle">Manage certificates and documents</p>
            </div>
          </div>
          <div className="certificate-management-actions">
            <button type="button" className="certificate-management-btn certificate-management-btn-icon" title="Filter" aria-label="Filter">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M22 3H2l8 9.46V19l4 2v-8.54L22 3z" /></svg>
            </button>
            <button type="button" className="certificate-management-btn certificate-management-btn-icon" title="Refresh" aria-label="Refresh" onClick={fetchCertificates} disabled={loading}>
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M23 4v6h-6M1 20v-6h6" /><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15" /></svg>
            </button>
            <button type="button" className="certificate-management-btn certificate-management-btn-add" onClick={openAdd}>
              <span className="certificate-management-btn-add-icon">+</span> Add
            </button>
          </div>
        </header>
      )}

      {showForm ? (
        <div className="certificate-management-form-center">
          {message && <div className="certificate-management-message">{message}</div>}
          <div className="certificate-management-card certificate-management-form-card">
            <div className="certificate-management-form-header">
              <h2 className="certificate-management-form-title">{editingId ? 'Edit Certificate' : 'Add Certificate'}</h2>
              <button type="button" className="certificate-management-form-close" onClick={closeForm} aria-label="Close">×</button>
            </div>
            <form onSubmit={handleSubmit} className="certificate-management-form">
              <div className="certificate-management-form-section">
                <strong>GST</strong>
                <div className="certificate-management-form-row">
                  <div className="certificate-management-form-col">
                    <label>GST No</label>
                    <input name="gstNo" value={form.gstNo} onChange={handleChange} placeholder="Enter GST number." />
                    <label>GST Date</label>
                    <input name="gstDate" type="date" value={form.gstDate} onChange={handleChange} />
                  </div>
                </div>
              </div>
              <div className="certificate-management-form-section">
                <strong>Company PAN</strong>
                <div className="certificate-management-form-row">
                  <div className="certificate-management-form-col">
                    <label>Company PAN No</label>
                    <input name="companyPANNo" value={form.companyPANNo} onChange={handleChange} placeholder="Enter company PAN number." />
                    <label>Company PAN Date</label>
                    <input name="companyPANDate" type="date" value={form.companyPANDate} onChange={handleChange} />
                  </div>
                </div>
              </div>
              <div className="certificate-management-form-section">
                <strong>Incorporation</strong>
                <div className="certificate-management-form-row">
                  <div className="certificate-management-form-col">
                    <label>Incorporation No</label>
                    <input name="incorporationNo" value={form.incorporationNo} onChange={handleChange} placeholder="Enter incorporation number." />
                    <label>Incorporation Date</label>
                    <input name="incorporationDate" type="date" value={form.incorporationDate} onChange={handleChange} />
                  </div>
                </div>
              </div>
              <div className="certificate-management-form-section">
                <strong>MOA / AOA</strong>
                <div className="certificate-management-form-row">
                  <div className="certificate-management-form-col">
                    <label>MOA No</label>
                    <input name="moaNo" value={form.moaNo} onChange={handleChange} placeholder="Enter MOA number." />
                    <label>MOA Date</label>
                    <input name="moaDate" type="date" value={form.moaDate} onChange={handleChange} />
                  </div>
                  <div className="certificate-management-form-col">
                    <label>AOA No</label>
                    <input name="aoaNo" value={form.aoaNo} onChange={handleChange} placeholder="Enter AOA number." />
                    <label>AOA Date</label>
                    <input name="aoaDate" type="date" value={form.aoaDate} onChange={handleChange} />
                  </div>
                </div>
              </div>
              <div className="certificate-management-form-section">
                <strong>EPF / ESI</strong>
                <div className="certificate-management-form-row">
                  <div className="certificate-management-form-col">
                    <label>EPF No</label>
                    <input name="epfNo" value={form.epfNo} onChange={handleChange} placeholder="Enter EPF number." />
                    <label>EPF Date</label>
                    <input name="epfDate" type="date" value={form.epfDate} onChange={handleChange} />
                  </div>
                  <div className="certificate-management-form-col">
                    <label>ESI No</label>
                    <input name="esiNo" value={form.esiNo} onChange={handleChange} placeholder="Enter ESI number." />
                    <label>ESI Date</label>
                    <input name="esiDate" type="date" value={form.esiDate} onChange={handleChange} />
                  </div>
                </div>
              </div>
              <div className="certificate-management-form-section">
                <strong>Factory License / LAF / PT</strong>
                <div className="certificate-management-form-row">
                  <div className="certificate-management-form-col">
                    <label>Factory License No</label>
                    <input name="factoryLicenseNo" value={form.factoryLicenseNo} onChange={handleChange} placeholder="Enter factory license number." />
                    <label>Factory License Date</label>
                    <input name="factoryLicenseDate" type="date" value={form.factoryLicenseDate} onChange={handleChange} />
                  </div>
                  <div className="certificate-management-form-col">
                    <label>LAF No</label>
                    <input name="lafNo" value={form.lafNo} onChange={handleChange} placeholder="Enter LAF number." />
                    <label>LRF Date</label>
                    <input name="lrfDate" type="date" value={form.lrfDate} onChange={handleChange} />
                  </div>
                  <div className="certificate-management-form-col">
                    <label>PT No</label>
                    <input name="ptNo" value={form.ptNo} onChange={handleChange} placeholder="Enter PT number." />
                    <label>PT Date</label>
                    <input name="ptDate" type="date" value={form.ptDate} onChange={handleChange} />
                  </div>
                </div>
              </div>
              <div className="certificate-management-form-footer">
                <button type="button" className="certificate-management-btn certificate-management-btn-secondary" onClick={closeForm}>Cancel</button>
                <button type="submit" className="certificate-management-btn certificate-management-btn-primary">{editingId ? 'Update' : 'Submit'}</button>
              </div>
            </form>
          </div>
        </div>
      ) : (
        <>
          {message && <div className="certificate-management-message">{message}</div>}
          <div className="certificate-management-card">
            {loading ? (
              <p className="certificate-management-loading">Loading...</p>
            ) : (
              <div className="certificate-management-grid">
                {certificates.length === 0 ? (
                  <p className="certificate-management-empty">No certificates. Click Add to create one.</p>
                ) : (
                  certificates.map((cert) => (
                    <div key={cert.id} className="certificate-management-item">
                      <div className="certificate-management-item-icon">
                        <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
                          <polyline points="14 2 14 8 20 8" />
                          <path d="M16 13H8" /><path d="M16 17H8" /><path d="M10 9H8" />
                        </svg>
                      </div>
                      <div className="certificate-management-item-name">{certLabel(cert)}</div>
                      <div className="certificate-management-item-actions">
                        <button type="button" className="certificate-management-item-btn edit" onClick={() => openEdit(cert)} title="Edit">
                          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7" /><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z" /></svg>
                        </button>
                        <button type="button" className="certificate-management-item-btn delete" onClick={() => handleDelete(cert.id)} title="Delete">
                          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polyline points="3 6 5 6 21 6" /><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3-3V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2" /><line x1="10" y1="11" x2="10" y2="17" /><line x1="14" y1="11" x2="14" y2="17" /></svg>
                        </button>
                      </div>
                    </div>
                  ))
                )}
              </div>
            )}
          </div>
        </>
      )}
    </div>
  );
};

export default Certificate;
