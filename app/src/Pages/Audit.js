import React, { useState, useEffect, useCallback } from 'react';
import { useSearchParams } from 'react-router-dom';
import './Audit.css';

const API_BASE = '/server/audit_function';

const Audit = ({ userRole, userEmail }) => {
  const [searchParams] = useSearchParams();
  const siteFromUrl = searchParams.get('site') || '';
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [message, setMessage] = useState('');
  const [selectedIds, setSelectedIds] = useState(new Set());
  const [sending, setSending] = useState(false);
  const [editModal, setEditModal] = useState(null);
  const [editForm, setEditForm] = useState({ maximumMarks: '', applicability: '' });
  const [saving, setSaving] = useState(false);
  const [editingRemarksId, setEditingRemarksId] = useState(null);
  const [editingRemarksValue, setEditingRemarksValue] = useState('');

  // When afrindinusha@gmail.com is logged in, show only CLRA audit records (Act or Sector = CLRA)
  const isCLRAOnlyUser = userEmail === 'afrindinusha@gmail.com';
  const isCLRARecord = (row) => {
    const act = String(row.act || '').toLowerCase().trim();
    const sector = String(row.sector || '').toLowerCase().trim();
    const isCLRA = (s) => s === 'clra' || s.includes('clra') || s.includes('contract labour') || s.includes('contract labor');
    return isCLRA(act) || isCLRA(sector);
  };

  // When afrindinusha.j@buildhr.co.in is logged in, show only Shops and Establishment audit records
  const isShopsAndEstablishmentOnlyUser = userEmail === 'afrindinusha.j@buildhr.co.in';
  // For these users: Applicability and Remarks are view-only (no edit), only view option
  const applicabilityRemarksViewOnlyEmails = ['afrindinusha.j@buildhr.co.in', 'afrindinusha@gmail.com', 'afrinatlin@gmail.com'];
  const isApplicabilityRemarksViewOnlyUser = applicabilityRemarksViewOnlyEmails.includes((userEmail || '').trim().toLowerCase());
  const isShopsAndEstablishmentRecord = (row) => {
    const act = String(row.act || '').toLowerCase().trim();
    const sector = String(row.sector || '').toLowerCase().trim();
    const isSE = (s) => s === 'shops and establishment' || s.includes('shops and establishment');
    return isSE(act) || isSE(sector);
  };

  // When afrinatlin@gmail.com is logged in, show only Factories Act audit records
  const isFactoriesActOnlyUser = userEmail === 'afrinatlin@gmail.com';
  const isFactoriesActRecord = (row) => {
    const act = String(row.act || '').toLowerCase().trim();
    const sector = String(row.sector || '').toLowerCase().trim();
    const isFA = (s) => s.includes('factories act') || s.includes('factory act') || s.includes('factories') || s.includes('factory');
    return isFA(act) || isFA(sector);
  };

  // When afrindinu14@gmail.com is on Audit — Delphi Kakinada, show only CLRA
  const isDelphiKakinadaCLRAOnly =
    (userEmail || '').trim().toLowerCase() === 'afrindinu14@gmail.com' &&
    siteFromUrl.trim().toLowerCase() === 'delphi kakinada';

  // When afrindinu14@gmail.com is on Audit — Delphi, show only Shops and Establishment
  const isDelphiShopsAndEstablishmentOnly =
    (userEmail || '').trim().toLowerCase() === 'afrindinu14@gmail.com' &&
    siteFromUrl.trim().toLowerCase() === 'delphi';

  // When afrindinu14@gmail.com is on Audit — Delphi Oragadam, show only Factories Act
  const isDelphiOragadamFactoriesActOnly =
    (userEmail || '').trim().toLowerCase() === 'afrindinu14@gmail.com' &&
    siteFromUrl.trim().toLowerCase() === 'delphi oragadam';

  // First apply act/sector filters (CLRA, Shops & Establishment, Factories Act, site-specific for afrindinu14)
  let filteredByAct = data;
  if (isCLRAOnlyUser) filteredByAct = data.filter(isCLRARecord);
  else if (isDelphiKakinadaCLRAOnly) filteredByAct = data.filter(isCLRARecord);
  else if (isDelphiShopsAndEstablishmentOnly) filteredByAct = data.filter(isShopsAndEstablishmentRecord);
  else if (isDelphiOragadamFactoriesActOnly) filteredByAct = data.filter(isFactoriesActRecord);
  else if (isShopsAndEstablishmentOnlyUser) filteredByAct = data.filter(isShopsAndEstablishmentRecord);
  else if (isFactoriesActOnlyUser) filteredByAct = data.filter(isFactoriesActRecord);

  // When ?site= is present (e.g. Audit — Delphi): show records for that site OR unassigned (empty approvedForSite) so user can assign them
  const displayData = siteFromUrl.trim()
    ? filteredByAct.filter((row) => {
        const site = String(row.approvedForSite || '').trim().toLowerCase();
        const selected = siteFromUrl.trim().toLowerCase();
        return site === selected || site === '';
      })
    : filteredByAct;

  const fetchData = useCallback(async () => {
    setLoading(true);
    setMessage('');
    try {
      const res = await fetch(`${API_BASE}/auditmaster?action=getAll`);
      const json = await res.json();
      if (json.status === 'success' && Array.isArray(json.data)) {
        setData(json.data);
      } else {
        setData([]);
        setMessage(json.message || 'Failed to load audit records');
      }
    } catch (err) {
      setData([]);
      setMessage('Failed to load audit records: ' + (err.message || 'Unknown error'));
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchData();
  }, [fetchData]);

  const toggleSelect = (id) => {
    setSelectedIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const toggleSelectAll = () => {
    if (selectedIds.size === displayData.length) {
      setSelectedIds(new Set());
    } else {
      setSelectedIds(new Set(displayData.map((r) => r.id)));
    }
  };

  const handleSendForApprovals = async () => {
    const ids = Array.from(selectedIds);
    if (!ids.length) {
      setMessage('Please select at least one record');
      return;
    }
    setSending(true);
    setMessage('');
    try {
      const res = await fetch(`${API_BASE}/auditmaster?action=sendForApproval`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          itemIds: ids,
          ...(siteFromUrl.trim() && { siteName: siteFromUrl.trim() })
        })
      });
      const json = await res.json();
      if (json.status === 'success') {
        setMessage(`Successfully sent ${ids.length} record(s) for approval`);
        setSelectedIds(new Set());
        await fetchData();
      } else {
        setMessage(json.message || 'Send for approval failed');
      }
    } catch (err) {
      setMessage('Send for approval failed: ' + (err.message || 'Unknown error'));
    } finally {
      setSending(false);
    }
  };

  const openEditModal = (row) => {
    setEditModal(row);
    setEditForm({
      maximumMarks: String(row.maximumMarks ?? '').trim(),
      applicability: String(row.applicability ?? '').trim()
    });
  };

  const closeEditModal = () => {
    setEditModal(null);
    setEditForm({ maximumMarks: '', applicability: '' });
  };

  const handleEditChange = (field, deltaOrValue) => {
    if (field !== 'maximumMarks' && field !== 'applicability') return;
    if (typeof deltaOrValue === 'number') {
      const num = parseInt(editForm[field], 10) || 0;
      setEditForm((prev) => ({ ...prev, [field]: String(Math.max(0, num + deltaOrValue)) }));
    } else {
      setEditForm((prev) => ({ ...prev, [field]: String(deltaOrValue ?? '') }));
    }
  };

  const handleSaveEdit = async () => {
    if (!editModal) return;
    setSaving(true);
    setMessage('');
    try {
      const res = await fetch(`${API_BASE}/auditmaster?action=update&id=${editModal.id}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          sector: editModal.sector,
          act: editModal.act,
          description: editModal.description,
          recordCategory: editModal.recordCategory,
          maximumMarks: editForm.maximumMarks || '',
          applicability: editForm.applicability || '',
          remarks: editModal.remarks
        })
      });
      const json = await res.json();
      if (json.status === 'success') {
        setMessage('Changes saved successfully');
        closeEditModal();
        await fetchData();
      } else {
        setMessage(json.message || 'Save failed');
      }
    } catch (err) {
      setMessage('Save failed: ' + (err.message || 'Unknown error'));
    } finally {
      setSaving(false);
    }
  };

  const openRemarksEdit = (row) => {
    setEditingRemarksId(row.id);
    setEditingRemarksValue(row.remarks || '');
  };

  const cancelRemarksEdit = () => {
    setEditingRemarksId(null);
    setEditingRemarksValue('');
  };

  const saveRemarksEdit = async () => {
    if (!editingRemarksId) return;
    const row = data.find((r) => r.id === editingRemarksId);
    if (!row) return;
    setSaving(true);
    setMessage('');
    try {
      const res = await fetch(`${API_BASE}/auditmaster?action=update&id=${editingRemarksId}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          sector: row.sector,
          act: row.act,
          description: row.description,
          recordCategory: row.recordCategory,
          maximumMarks: row.maximumMarks || '',
          applicability: row.applicability || '',
          remarks: editingRemarksValue
        })
      });
      const json = await res.json();
      if (json.status === 'success') {
        setMessage('Remarks saved successfully');
        cancelRemarksEdit();
        await fetchData();
      } else {
        setMessage(json.message || 'Save failed');
      }
    } catch (err) {
      setMessage('Save failed: ' + (err.message || 'Unknown error'));
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="audit-page">
      <div className="audit-card">
        <div className="audit-header-row">
          <div className="audit-header-left">
            <h1 className="audit-title">{siteFromUrl ? `Audit — ${siteFromUrl}` : 'Audit'}</h1>
            <p className="audit-subtitle">
              {siteFromUrl
                ? `View and manage audit records for ${siteFromUrl}${isDelphiKakinadaCLRAOnly ? ' (CLRA only)' : ''}${isDelphiShopsAndEstablishmentOnly ? ' (Shops and Establishment only)' : ''}${isDelphiOragadamFactoriesActOnly ? ' (Factories Act only)' : ''}`
                : isCLRAOnlyUser
                  ? 'View and manage audit records (CLRA only)'
                  : isShopsAndEstablishmentOnlyUser
                    ? 'View and manage audit records (Shops and Establishment only)'
                    : isFactoriesActOnlyUser
                      ? 'View and manage audit records (Factories Act only)'
                      : 'View and manage audit records'}
            </p>
          </div>
          <button
            type="button"
            className="audit-btn-send"
            onClick={handleSendForApprovals}
            disabled={sending || selectedIds.size === 0}
          >
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <line x1="22" y1="2" x2="11" y2="13" />
              <polygon points="22 2 15 22 11 13 2 9 22 2" />
            </svg>
            Send For Approvals
          </button>
        </div>

        {message && <p className="audit-message">{message}</p>}

        <div className="audit-table-wrap">
          {loading && !data.length ? (
            <p className="audit-loading">Loading...</p>
          ) : (
            <table className="audit-table">
              <thead>
                <tr>
                  <th className="audit-col-check">
                    <input
                      type="checkbox"
                      checked={displayData.length > 0 && selectedIds.size === displayData.length}
                      onChange={toggleSelectAll}
                      aria-label="Select all"
                    />
                  </th>
                  <th>S.No</th>
                  <th>Act</th>
                  <th>Description</th>
                  <th>Record Category</th>
                  <th>Maximum Marks</th>
                  <th>Applicability</th>
                  <th>Remarks</th>
                </tr>
              </thead>
              <tbody>
                {displayData.length === 0 ? (
                  <tr>
                    <td colSpan={8} className="audit-empty">
                      {siteFromUrl
                        ? isDelphiKakinadaCLRAOnly
                          ? `No CLRA audit records for ${siteFromUrl}. Send for approval with this site selected to assign records.`
                          : isDelphiShopsAndEstablishmentOnly
                            ? `No Shops and Establishment audit records for ${siteFromUrl}. Send for approval with this site selected to assign records.`
                            : isDelphiOragadamFactoriesActOnly
                              ? `No Factories Act audit records for ${siteFromUrl}. Send for approval with this site selected to assign records.`
                              : `No audit records for ${siteFromUrl}. Send for approval with this site selected to assign records.`
                        : isCLRAOnlyUser
                          ? 'No CLRA audit records.'
                          : isShopsAndEstablishmentOnlyUser
                            ? 'No Shops and Establishment audit records.'
                            : isFactoriesActOnlyUser
                              ? 'No Factories Act audit records.'
                              : 'No audit records. Import data from Audit Master Import.'}
                    </td>
                  </tr>
                ) : (
                  displayData.map((row, i) => (
                    <tr key={row.id || i}>
                      <td className="audit-col-check">
                        <input
                          type="checkbox"
                          checked={selectedIds.has(row.id)}
                          onChange={() => toggleSelect(row.id)}
                          aria-label={`Select row ${i + 1}`}
                        />
                      </td>
                      <td>{i + 1}</td>
                      <td>{row.act || '—'}</td>
                      <td>{row.description || '—'}</td>
                      <td>{row.recordCategory || '—'}</td>
                      <td>
                        <button type="button" className="audit-edit-link" onClick={() => openEditModal(row)} title="Edit Maximum Marks & Applicability">
                          {row.maximumMarks != null && row.maximumMarks !== '' ? row.maximumMarks : '—'}
                        </button>
                      </td>
                      <td>
                        {isApplicabilityRemarksViewOnlyUser ? (
                          <span>{row.applicability != null && row.applicability !== '' ? row.applicability : '—'}</span>
                        ) : (
                          <button type="button" className="audit-edit-link" onClick={() => openEditModal(row)} title="Edit Maximum Marks & Applicability">
                            {row.applicability != null && row.applicability !== '' ? row.applicability : '—'}
                          </button>
                        )}
                      </td>
                      <td className="audit-remarks-cell">
                        {isApplicabilityRemarksViewOnlyUser ? (
                          <span>{row.remarks && String(row.remarks).trim() ? row.remarks : '—'}</span>
                        ) : editingRemarksId === row.id ? (
                          <div className="audit-remarks-edit" onClick={(e) => e.stopPropagation()}>
                            <textarea
                              className="audit-remarks-textarea"
                              value={editingRemarksValue}
                              onChange={(e) => setEditingRemarksValue(e.target.value)}
                              placeholder="Enter remarks..."
                              rows={3}
                            />
                            <div className="audit-remarks-actions">
                              <button type="button" className="audit-btn audit-btn-cancel" onClick={cancelRemarksEdit} disabled={saving}>Cancel</button>
                              <button type="button" className="audit-btn audit-btn-save" onClick={saveRemarksEdit} disabled={saving}>{saving ? 'Saving...' : 'Save'}</button>
                            </div>
                          </div>
                        ) : (
                          <button type="button" className="audit-remarks-trigger" onClick={() => openRemarksEdit(row)}>
                            {row.remarks && String(row.remarks).trim() ? row.remarks : 'Click to add remarks'}
                          </button>
                        )}
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          )}
        </div>
      </div>

      {editModal && (
        <div className="audit-modal-overlay" onClick={closeEditModal}>
          <div className="audit-modal" onClick={(e) => e.stopPropagation()}>
            <div className="audit-modal-header">
              <h2>{isApplicabilityRemarksViewOnlyUser ? 'Edit Maximum Marks' : 'Edit Maximum Marks &amp; Applicability'}</h2>
              <button type="button" className="audit-modal-close" onClick={closeEditModal} title="Close">&times;</button>
            </div>
            <div className="audit-modal-body">
              <div className="audit-modal-field">
                <label>Act</label>
                <p className="audit-modal-static">{editModal.act || '—'}</p>
              </div>
              <div className="audit-modal-field">
                <label>Description</label>
                <p className="audit-modal-static">{editModal.description || '—'}</p>
              </div>
              <div className="audit-modal-field">
                <label>Maximum Marks</label>
                <div className="audit-modal-stepper">
                  <button type="button" className="audit-stepper-btn" onClick={() => handleEditChange('maximumMarks', -1)} aria-label="Decrease">−</button>
                  <input type="text" className="audit-modal-input" value={editForm.maximumMarks} onChange={(e) => handleEditChange('maximumMarks', e.target.value)} inputMode="numeric" />
                  <button type="button" className="audit-stepper-btn" onClick={() => handleEditChange('maximumMarks', 1)} aria-label="Increase">+</button>
                </div>
              </div>
              <div className="audit-modal-field">
                <label>Applicability</label>
                {isApplicabilityRemarksViewOnlyUser ? (
                  <p className="audit-modal-static">{editModal.applicability != null && editModal.applicability !== '' ? editModal.applicability : '—'}</p>
                ) : (
                  <div className="audit-modal-stepper">
                    <button type="button" className="audit-stepper-btn" onClick={() => handleEditChange('applicability', -1)} aria-label="Decrease">−</button>
                    <input type="text" className="audit-modal-input" value={editForm.applicability} onChange={(e) => handleEditChange('applicability', e.target.value)} inputMode="numeric" />
                    <button type="button" className="audit-stepper-btn" onClick={() => handleEditChange('applicability', 1)} aria-label="Increase">+</button>
                  </div>
                )}
              </div>
            </div>
            <div className="audit-modal-footer">
              <button type="button" className="audit-btn audit-btn-cancel" onClick={closeEditModal} disabled={saving}>Cancel</button>
              <button type="button" className="audit-btn audit-btn-save" onClick={handleSaveEdit} disabled={saving}>
                {saving ? 'Saving...' : 'Save Changes'}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default Audit;
