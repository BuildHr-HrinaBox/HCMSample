import React, { useState, useEffect, useCallback } from 'react';
import * as XLSX from 'xlsx';
import './ToDoList.css';

const API_BASE = '/server/todolist_function';
const CHECKLIST_BULK_API = '/server/checklistbulk_function/checklistbulk';
const FORMMASTER_TEMPLATES_API = '/server/formmaster_function/templates';
const FORMMASTER_DOWNLOAD_BASE = '/server/formmaster_function/templates/download';
const SPECIAL_CHECKLIST_BULK_EMAIL = 'afrindinusha.j@buildhr.co.in';
const TARGET_SECTOR_REGEX = /shop[s]?\s*(and|&)?\s*establishment/i;

const ToDoList = ({ userEmail }) => {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [message, setMessage] = useState('');
  const [modalOpen, setModalOpen] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [form, setForm] = useState({
    act: '',
    description: '',
    dueDate: '',
    formName: ''
  });
  const [attachment, setAttachment] = useState(null);
  const [proofAttachment, setProofAttachment] = useState(null);
  const [uploading, setUploading] = useState(false);
  const [autofilling, setAutofilling] = useState(false);
  const [autofillRowId, setAutofillRowId] = useState(null);
  const [inlinePreviewUrl, setInlinePreviewUrl] = useState('');
  const [inlinePreviewHtml, setInlinePreviewHtml] = useState('');
  const [excelPreviewTitle, setExcelPreviewTitle] = useState('');
  const normalizedEmail = String(userEmail || '').trim().toLowerCase();
  const useChecklistBulkSource = normalizedEmail === SPECIAL_CHECKLIST_BULK_EMAIL;

  const mapChecklistBulkToTodo = (item, index) => ({
    id: item.id || item.ROWID || `bulk-${index}`,
    sector: item.sector || item.Sector || '',
    act: item.act || item.Act || '',
    description: item.description || item.Description || '',
    dueDate: item.dueDate || item.DueDate || '',
    formName: item.formName || item.FormName || '',
    formFile: null,
    formFileName: null,
    proofSubmissionFile: null,
    proofSubmissionFileName: null
  });

  const normalizeFormText = (value) =>
    String(value || '')
      .toLowerCase()
      .replace(/\.(xlsx|xls|csv)$/i, '')
      .replace(/[^a-z0-9]/g, '');

  // Canonicalize common form code variants:
  // Form I / Form l / Form 1 -> form1
  const canonicalFormCode = (value) => {
    const raw = String(value || '').toLowerCase();
    const withoutExt = raw.replace(/\.(xlsx|xls|csv)$/i, '');
    const m = withoutExt.match(/form\s*([a-z0-9]+)/i);
    if (!m) return normalizeFormText(withoutExt);
    let code = m[1].trim();
    if (code === 'i' || code === 'l') code = '1';
    if (code === 'ii') code = '2';
    if (code === 'iii') code = '3';
    if (code === 'iv') code = '4';
    if (code === 'v') code = '5';
    if (code === 'vi') code = '6';
    if (code === 'vii') code = '7';
    if (code === 'viii') code = '8';
    if (code === 'ix') code = '9';
    if (code === 'x') code = '10';
    return `form${code}`.replace(/[^a-z0-9]/g, '');
  };

  const findBestTemplate = (templates, formName) => {
    const exactKey = normalizeFormText(formName);
    const canonicalKey = canonicalFormCode(formName);
    if (!exactKey && !canonicalKey) return null;

    // 1) strict exact match first (safest)
    const strict = templates.find((t) => normalizeFormText(t?.name) === exactKey);
    if (strict) return strict;

    // 2) canonical match (Form I == Form 1 == Form l)
    const canonical = templates.find((t) => canonicalFormCode(t?.name) === canonicalKey);
    if (canonical) return canonical;

    // 3) fallback contains match
    return templates.find((t) => {
      const tKey = normalizeFormText(t?.name);
      if (!tKey) return false;
      return tKey.includes(exactKey) || exactKey.includes(tKey);
    }) || null;
  };

  const extractZohoEmployees = (result) => {
    const data = result?.data;
    if (!data) return [];
    let employees = [];
    if (data.response && data.response.result) {
      if (Array.isArray(data.response.result)) {
        data.response.result.forEach((resultItem) => {
          if (typeof resultItem === 'object' && resultItem !== null) {
            Object.keys(resultItem).forEach((idKey) => {
              const employeeArray = resultItem[idKey];
              if (Array.isArray(employeeArray)) employees.push(...employeeArray);
              else if (employeeArray && typeof employeeArray === 'object') employees.push(employeeArray);
            });
          }
        });
      } else if (typeof data.response.result === 'object') {
        const keys = Object.keys(data.response.result);
        if (keys.length > 0) {
          const first = data.response.result[keys[0]];
          if (Array.isArray(first)) employees = first;
          else if (first && typeof first === 'object') employees = [first];
        }
      }
    } else if (Array.isArray(data)) {
      employees = data;
    } else if (Array.isArray(data.records)) {
      employees = data.records;
    } else if (Array.isArray(data.result)) {
      employees = data.result;
    }
    return Array.isArray(employees) ? employees : [];
  };

  const fetchData = useCallback(async () => {
    setLoading(true);
    setMessage('');
    try {
      if (useChecklistBulkSource) {
        const res = await fetch(`${CHECKLIST_BULK_API}?action=getAll`);
        const json = await res.json();
        if (json.status === 'success' && Array.isArray(json.data)) {
          const mappedRows = json.data.map((item, idx) => mapChecklistBulkToTodo(item, idx));
          const filteredRows = mappedRows.filter((row) =>
            TARGET_SECTOR_REGEX.test(String(row.sector || ''))
          );
          let templates = [];
          try {
            const formmasterRes = await fetch(FORMMASTER_TEMPLATES_API);
            const formmasterJson = await formmasterRes.json();
            if (formmasterJson.status === 'success' && Array.isArray(formmasterJson.data)) {
              templates = formmasterJson.data;
            }
          } catch (_) {
            templates = [];
          }
          const rowsWithTemplate = filteredRows.map((row) => {
            const matched = findBestTemplate(templates, row.formName);
            if (!matched) return row;
            return {
              ...row,
              formFile: matched.id || null,
              formFileName: matched.name || null
            };
          });
          setData(rowsWithTemplate);
        } else {
          setData([]);
          setMessage(json.message || 'Could not load checklist bulk data.');
        }
      } else {
        const res = await fetch(`${API_BASE}/todos`);
        const json = await res.json();
        if (json.status === 'success' && Array.isArray(json.data?.todos)) {
          setData(json.data.todos);
        } else {
          setData([]);
          setMessage(json.message || 'Could not load to-dos.');
        }
      }
    } catch (err) {
      setData([]);
      setMessage('Failed to load: ' + (err.message || 'Unknown error'));
    } finally {
      setLoading(false);
    }
  }, [useChecklistBulkSource]);

  useEffect(() => {
    fetchData();
  }, [fetchData]);

  const resetForm = () => {
    setForm({ act: '', description: '', dueDate: '', formName: '' });
    setAttachment(null);
    setProofAttachment(null);
    setEditingId(null);
    setModalOpen(false);
    setInlinePreviewUrl('');
    setInlinePreviewHtml('');
    setExcelPreviewTitle('');
  };

  const openAdd = () => {
    resetForm();
    setModalOpen(true);
  };

  const openEdit = (row) => {
    setForm({
      act: row.act || '',
      description: row.description || '',
      dueDate: row.dueDate || '',
      formName: row.formName || ''
    });
    setEditingId(row.id);
    setAttachment(null);
    setProofAttachment(null);
    setModalOpen(true);
  };

  const handleChange = (e) => {
    const { name, value } = e.target;
    setForm((prev) => ({ ...prev, [name]: value }));
  };

  const uploadFile = async (file, docType = 'FormFile') => {
    const fd = new FormData();
    fd.append('file', file);
    const res = await fetch(`${API_BASE}/todos/upload/${docType}`, { method: 'POST', body: fd });
    const json = await res.json();
    if (json.status === 'success') {
      return { fileId: json.fileId, fileName: json.fileName || file.name };
    }
    throw new Error(json.message || 'Upload failed');
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (!(form.act || '').trim()) {
      setMessage('Act is required');
      return;
    }
    setUploading(true);
    setMessage('');
    try {
      let formFileId = null;
      let formFileName = null;
      let proofSubmissionFileId = null;
      let proofSubmissionFileName = null;

      if (editingId) {
        const row = data.find((r) => r.id === editingId);
        formFileId = row?.formFile || null;
        formFileName = row?.formFileName || null;
        proofSubmissionFileId = row?.proofSubmissionFile || null;
        proofSubmissionFileName = row?.proofSubmissionFileName || null;
      }

      if (attachment) {
        const u = await uploadFile(attachment, 'FormFile');
        formFileId = u.fileId;
        formFileName = u.fileName;
      }
      if (proofAttachment) {
        const p = await uploadFile(proofAttachment, 'ProofSubmissionFile');
        proofSubmissionFileId = p.fileId;
        proofSubmissionFileName = p.fileName;
      }

      // Autofill flow: if proof file is empty, mirror selected FormFile into ProofSubmissionFile on Save.
      if (inlinePreviewUrl && formFileId && !proofSubmissionFileId) {
        proofSubmissionFileId = formFileId;
        proofSubmissionFileName = formFileName;
      }

      const payload = {
        act: form.act,
        description: form.description,
        dueDate: form.dueDate,
        formName: form.formName,
        formFile: formFileId,
        formFileName,
        proofSubmissionFile: proofSubmissionFileId,
        proofSubmissionFileName
      };

      // In checklist-bulk view mode, update current table row locally (no todolist backend row).
      if (useChecklistBulkSource) {
        setData((prev) =>
          prev.map((r) =>
            r.id === editingId
              ? {
                ...r,
                act: form.act,
                description: form.description,
                dueDate: form.dueDate,
                formName: form.formName,
                formFile: formFileId,
                formFileName,
                proofSubmissionFile: proofSubmissionFileId,
                proofSubmissionFileName
              }
              : r
          )
        );
        setMessage('Saved successfully');
        resetForm();
        return;
      }

      const url = editingId ? `${API_BASE}/todos/${editingId}` : `${API_BASE}/todos`;
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
    if (!window.confirm('Delete this to-do?')) return;
    setMessage('');
    try {
      const res = await fetch(`${API_BASE}/todos/${id}`, { method: 'DELETE' });
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

  const handleRemoveAttachment = async (id, docType = 'FormFile') => {
    if (!window.confirm('Remove the attached file from this to-do?')) return;
    setMessage('');
    try {
      const res = await fetch(`${API_BASE}/todos/${id}/file/${docType}`, { method: 'DELETE' });
      const json = await res.json();
      if (json.status === 'success') {
        setMessage('Attachment removed');
        await fetchData();
      } else {
        setMessage(json.message || 'Could not remove file');
      }
    } catch (err) {
      setMessage(err.message || 'Could not remove file');
    }
  };

  const normalizeRowKey = (row) =>
    [
      String(row?.act || '').trim().toLowerCase(),
      String(row?.description || '').trim().toLowerCase(),
      String(row?.dueDate || '').trim().toLowerCase(),
      String(row?.formName || '').trim().toLowerCase()
    ].join('|');

  const handleAutofillAllActs = async () => {
    setAutofilling(true);
    setMessage('');
    try {
      if (useChecklistBulkSource) {
        await fetchData();
        setMessage('Autofill completed from Checklist Bulk.');
        return;
      }
      const srcResp = await fetch(`${CHECKLIST_BULK_API}?action=getAll`);
      const srcJson = await srcResp.json();
      if (!(srcJson.status === 'success' && Array.isArray(srcJson.data))) {
        setMessage(srcJson.message || 'Could not load checklist bulk data for autofill.');
        return;
      }

      const existingKeys = new Set((data || []).map((row) => normalizeRowKey(row)));
      const mapped = srcJson.data.map((item, idx) => mapChecklistBulkToTodo(item, idx));
      const seenNewKeys = new Set();
      const candidates = mapped.filter((row) => {
        if (!String(row?.act || '').trim()) return false;
        const key = normalizeRowKey(row);
        if (existingKeys.has(key) || seenNewKeys.has(key)) return false;
        seenNewKeys.add(key);
        return true;
      });

      let added = 0;
      for (const row of candidates) {
        const payload = {
          act: row.act,
          description: row.description || '',
          dueDate: row.dueDate || '',
          formName: row.formName || '',
          formFile: null,
          formFileName: null,
          proofSubmissionFile: null,
          proofSubmissionFileName: null
        };
        const res = await fetch(`${API_BASE}/todos`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });
        const json = await res.json();
        if (json.status === 'success') added += 1;
      }

      await fetchData();
      setMessage(added > 0 ? `Autofill completed. Added ${added} act(s).` : 'Autofill completed. No new acts to add.');
    } catch (err) {
      setMessage('Autofill failed: ' + (err.message || 'Unknown error'));
    } finally {
      setAutofilling(false);
    }
  };

  const handleAutofillFormFile = async (row) => {
    if (!row || !row.id) return;
    setAutofilling(true);
    setAutofillRowId(row.id);
    setMessage('');
    try {
      const formmasterRes = await fetch(FORMMASTER_TEMPLATES_API);
      const formmasterJson = await formmasterRes.json();
      if (!(formmasterJson.status === 'success' && Array.isArray(formmasterJson.data))) {
        setMessage('Could not load Formmaster templates.');
        return;
      }
      const matched = findBestTemplate(formmasterJson.data, row.formName);
      if (!matched) {
        setMessage(`No matching Excel found in Formmaster for "${row.formName || 'this form'}".`);
        return;
      }
      const matchedFileUrl = formmasterDownloadUrl(matched.id, matched.name, 'inline');

      setData((prev) =>
        prev.map((r) =>
          r.id === row.id
            ? { ...r, formFile: matched.id || null, formFileName: matched.name || null }
            : r
        )
      );

      // Open only excel preview popup (no form fields, no download).
      setForm({
        act: row.act || '',
        description: row.description || '',
        dueDate: row.dueDate || '',
        formName: row.formName || ''
      });
      setEditingId(row.id);
      setModalOpen(true);
      setExcelPreviewTitle(matched.name || row.formName || 'Form Preview');
      // Use Office web embed so headers/borders/format render similar to source file.
      setInlinePreviewUrl(officeEmbedUrl(matchedFileUrl));
      setInlinePreviewHtml('');

      if (!useChecklistBulkSource) {
        const payload = {
          act: row.act || '',
          description: row.description || '',
          dueDate: row.dueDate || '',
          formName: row.formName || '',
          formFile: matched.id || null,
          formFileName: matched.name || null,
          proofSubmissionFile: row.proofSubmissionFile || null,
          proofSubmissionFileName: row.proofSubmissionFileName || null
        };
        await fetch(`${API_BASE}/todos/${row.id}`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });

      }

      // Same Zoho People fetch flow reference as Statutory page.
      let employeeCount = 0;
      try {
        const peopleResp = await fetch('/server/peopledata_function?form=employee&limit=100');
        const peopleResult = await peopleResp.json();
        if (peopleResult?.success) {
          const employees = extractZohoEmployees(peopleResult);
          employeeCount = employees.length;
          if (employeeCount > 0) {
            // Fill template with Zoho People rows and render in same modal.
            const tmplResp = await fetch(matchedFileUrl);
            const tmplBuffer = await tmplResp.arrayBuffer();
            const wb = XLSX.read(tmplBuffer, { type: 'array' });
            const wsName = wb.SheetNames[0];
            const ws = wb.Sheets[wsName];
            const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
            const headerIdx = rows.findIndex((r) =>
              Array.isArray(r) && r.some((c) => /emp\s*id|workman|designation/i.test(String(c || '')))
            );
            const start = (headerIdx >= 0 ? headerIdx : 3) + 1;
            const maxCols = Math.max(...rows.map((r) => (Array.isArray(r) ? r.length : 0)), 12);
            const pick = (emp, keys) => keys.map((k) => emp?.[k]).find((v) => v !== undefined && v !== null && String(v).trim() !== '') || '';
            const pad = (ri) => {
              if (!rows[ri]) rows[ri] = [];
              while (rows[ri].length < maxCols) rows[ri].push('');
            };

            employees.forEach((emp, i) => {
              const ri = start + i;
              pad(ri);
              const fullName = pick(emp, ['DisplayName', 'Name']) || [emp?.FirstName, emp?.MiddleName, emp?.LastName].filter(Boolean).join(' ');
              rows[ri][0] = i + 1;
              rows[ri][1] = pick(emp, ['EmployeeID', 'Employee_ID', 'ZUID', 'Zoho_ID']);
              rows[ri][2] = fullName || '';
              rows[ri][3] = pick(emp, ['Designation', 'JobTitle', 'DesignationName']);
              rows[ri][5] = pick(emp, ['Dateofjoining', 'Date_of_joining', 'DateOfJoining']);
              rows[ri][7] = pick(emp, ['DateOfConfirmation', 'Date_of_confirmation']);
            });

            const outWs = XLSX.utils.aoa_to_sheet(rows);
            outWs['!merges'] = ws['!merges'];
            outWs['!cols'] = ws['!cols'];
            outWs['!rows'] = ws['!rows'];
            const html = XLSX.utils.sheet_to_html(outWs);
            setInlinePreviewHtml(html || '');
            setInlinePreviewUrl('');
          }
        }
      } catch (_) {
        employeeCount = 0;
      }

      setMessage(
        employeeCount > 0
          ? `FormFile autofilled for "${row.formName || row.act || row.id}". Successfully loaded ${employeeCount} employee records from Zoho People.`
          : `FormFile autofilled for "${row.formName || row.act || row.id}".`
      );
    } catch (err) {
      setMessage('FormFile autofill failed: ' + (err.message || 'Unknown error'));
    } finally {
      setAutofilling(false);
      setAutofillRowId(null);
    }
  };

  const fileUrl = (id, docType = 'FormFile', disposition = 'inline') =>
    `${API_BASE}/todos/${id}/file/${docType}?disposition=${disposition}`;
  const formmasterDownloadUrl = (fileId, fileName, disposition = 'attachment') =>
    `${FORMMASTER_DOWNLOAD_BASE}/${fileId}?fileName=${encodeURIComponent(fileName || 'template.xlsx')}&disposition=${encodeURIComponent(disposition)}`;
  const toAbsoluteUrl = (url) =>
    String(url || '').startsWith('http') ? String(url) : `${window.location.origin}${String(url || '')}`;
  const officeEmbedUrl = (url) =>
    `https://view.officeapps.live.com/op/embed.aspx?src=${encodeURIComponent(toAbsoluteUrl(url))}`;

  return (
    <div className="td-page">
      <div className="td-card-container">
        <div className="td-header-row">
          <div className="td-header-left">
            <div className="td-title-section">
              <h2 className="td-section-title">To-Do List</h2>
              <p className="td-section-subtitle">
                {useChecklistBulkSource
                  ? 'Logged in as afrindinusha.j@buildhr.co.in: showing Checklist Bulk records for Sector = Shops and Establishment.'
                  : 'Track Act, Description, Form Name, and FormFile attachments for each item.'}
              </p>
            </div>
          </div>
          <div className="td-actions-card">
            {message ? <span className="td-status-msg">{message}</span> : null}
            {!useChecklistBulkSource ? (
              <button type="button" className="td-action-btn td-btn-add" onClick={openAdd} title="Add to-do">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <line x1="12" y1="5" x2="12" y2="19" />
                  <line x1="5" y1="12" x2="19" y2="12" />
                </svg>
              </button>
            ) : null}
          </div>
        </div>

        <div className="td-table-container">
          {loading && !data.length ? (
            <p className="td-loading">Loading...</p>
          ) : (
            <table className="td-table">
              <thead>
                <tr>
                  <th>S.No</th>
                  <th>Act</th>
                  <th>Description</th>
                  <th>DueDate</th>
                  <th>Form Name</th>
                  <th>FormFile</th>
                  <th>Autofill</th>
                  <th>ProofSubmissionFile</th>
                  {!useChecklistBulkSource ? <th>Actions</th> : null}
                </tr>
              </thead>
              <tbody>
                {data.length === 0 ? (
                  <tr>
                    <td colSpan={useChecklistBulkSource ? 8 : 9} className="td-empty">
                      {useChecklistBulkSource
                        ? 'No checklist bulk data found.'
                        : 'No items yet. Use the + button to add a to-do.'}
                    </td>
                  </tr>
                ) : (
                  data.map((row, i) => (
                    <tr key={row.id ?? i}>
                      <td>{i + 1}</td>
                      <td className="td-cell-wrap">{row.act || '—'}</td>
                      <td className="td-cell-wrap" title={row.description}>
                        {row.description || '—'}
                      </td>
                      <td>{row.dueDate || '—'}</td>
                      <td className="td-cell-wrap">{row.formName || '—'}</td>
                      <td>
                        {row.formFile ? (
                          <span className="td-file-cell">
                            <a
                              href={
                                useChecklistBulkSource
                                  ? formmasterDownloadUrl(row.formFile, row.formFileName)
                                  : fileUrl(row.id, 'FormFile', 'attachment')
                              }
                              target="_blank"
                              rel="noopener noreferrer"
                              className="td-file-link"
                            >
                              {row.formFileName || 'Download'}
                            </a>
                            {!useChecklistBulkSource ? (
                              <button
                                type="button"
                                className="td-file-remove"
                                onClick={() => handleRemoveAttachment(row.id, 'FormFile')}
                                title="Remove FormFile"
                              >
                                ×
                              </button>
                            ) : null}
                          </span>
                        ) : (
                          '—'
                        )}
                      </td>
                      <td>
                        <button
                          type="button"
                          className="td-row-btn td-row-btn-autofill"
                          onClick={() => handleAutofillFormFile(row)}
                          disabled={autofilling}
                        >
                          {autofilling && autofillRowId === row.id ? 'Loading...' : 'Autofill'}
                        </button>
                      </td>
                      <td>
                        {row.proofSubmissionFile ? (
                          <span className="td-file-cell">
                            <a
                              href={
                                useChecklistBulkSource
                                  ? formmasterDownloadUrl(row.proofSubmissionFile, row.proofSubmissionFileName)
                                  : fileUrl(row.id, 'ProofSubmissionFile', 'attachment')
                              }
                              target="_blank"
                              rel="noopener noreferrer"
                              className="td-file-link"
                            >
                              {row.proofSubmissionFileName || 'Download'}
                            </a>
                            {!useChecklistBulkSource ? (
                              <button
                                type="button"
                                className="td-file-remove"
                                onClick={() => handleRemoveAttachment(row.id, 'ProofSubmissionFile')}
                                title="Remove ProofSubmissionFile"
                              >
                                ×
                              </button>
                            ) : null}
                          </span>
                        ) : (
                          '—'
                        )}
                      </td>
                      {!useChecklistBulkSource ? (
                        <td>
                          <button type="button" className="td-row-btn td-row-btn-edit" onClick={() => openEdit(row)}>
                            Edit
                          </button>
                          <button
                            type="button"
                            className="td-row-btn td-row-btn-delete"
                            onClick={() => handleDelete(row.id)}
                          >
                            Delete
                          </button>
                        </td>
                      ) : null}
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          )}
          <div className="td-table-footer">Showing {data.length} to-do(s)</div>
        </div>
      </div>

      {modalOpen ? (
        <div className="td-modal-overlay" onClick={resetForm}>
          <div
            className={`td-modal ${inlinePreviewUrl ? 'td-modal-wide' : ''}`}
            onClick={(e) => e.stopPropagation()}
          >
            <div className="td-modal-header">
              <h2>
                {excelPreviewTitle || (useChecklistBulkSource
                  ? 'To-Do Preview'
                  : editingId ? 'Edit to-do' : 'New to-do')}
              </h2>
              <button type="button" className="td-modal-close" onClick={resetForm} title="Close">
                ×
              </button>
            </div>
            <form className="td-form" onSubmit={handleSubmit}>
              {!inlinePreviewUrl ? (
                <>
                  <div className="td-form-row">
                    <div className="td-field">
                      <label htmlFor="td-act">
                        Act <span className="td-req">*</span>
                      </label>
                      <input
                        id="td-act"
                        name="act"
                        className="td-input"
                        value={form.act}
                        onChange={handleChange}
                        required
                        placeholder="Act / task title"
                      />
                    </div>
                    <div className="td-field">
                      <label htmlFor="td-dueDate">DueDate</label>
                      <input
                        id="td-dueDate"
                        name="dueDate"
                        type="date"
                        className="td-input"
                        value={form.dueDate}
                        onChange={handleChange}
                      />
                    </div>
                    <div className="td-field">
                      <label htmlFor="td-formName">Form Name</label>
                      <input
                        id="td-formName"
                        name="formName"
                        className="td-input"
                        value={form.formName}
                        onChange={handleChange}
                        placeholder="Form Name"
                      />
                    </div>
                  </div>
                  <div className="td-field td-field-full">
                    <label htmlFor="td-description">Description</label>
                    <textarea
                      id="td-description"
                      name="description"
                      className="td-input td-textarea"
                      value={form.description}
                      onChange={handleChange}
                      placeholder="Details"
                      rows={4}
                    />
                  </div>
                  <div className="td-field td-field-full">
                    <label htmlFor="td-file">FormFile</label>
                    {editingId && data.find((r) => r.id === editingId)?.formFile ? (
                      <div className="td-current-file">
                        <span>Current: {data.find((r) => r.id === editingId)?.formFileName || 'file'}</span>
                        <a
                          href={
                            useChecklistBulkSource
                              ? formmasterDownloadUrl(
                                data.find((r) => r.id === editingId)?.formFile,
                                data.find((r) => r.id === editingId)?.formFileName,
                                'inline'
                              )
                              : fileUrl(editingId, 'FormFile', 'inline')
                          }
                          target="_blank"
                          rel="noopener noreferrer"
                          className="td-file-link"
                        >
                          Open
                        </a>
                      </div>
                    ) : null}
                    <input
                      id="td-file"
                      type="file"
                      accept=".pdf,.doc,.docx,.xls,.xlsx,.png,.jpg,.jpeg,.txt"
                      onChange={(e) => setAttachment(e.target.files?.[0] || null)}
                      className="td-file-input"
                    />
                    {attachment ? <div className="td-file-selected">Selected: {attachment.name}</div> : null}
                  </div>
                  <div className="td-field td-field-full">
                    <label htmlFor="td-proofFile">ProofSubmissionFile</label>
                    {editingId && data.find((r) => r.id === editingId)?.proofSubmissionFile ? (
                      <div className="td-current-file">
                        <span>Current: {data.find((r) => r.id === editingId)?.proofSubmissionFileName || 'file'}</span>
                        <a
                          href={fileUrl(editingId, 'ProofSubmissionFile', 'inline')}
                          target="_blank"
                          rel="noopener noreferrer"
                          className="td-file-link"
                        >
                          Open
                        </a>
                      </div>
                    ) : null}
                    <input
                      id="td-proofFile"
                      type="file"
                      accept=".pdf,.doc,.docx,.xls,.xlsx,.png,.jpg,.jpeg,.txt"
                      onChange={(e) => setProofAttachment(e.target.files?.[0] || null)}
                      className="td-file-input"
                    />
                    {proofAttachment ? <div className="td-file-selected">Selected: {proofAttachment.name}</div> : null}
                  </div>
                </>
              ) : null}
              {(inlinePreviewUrl || inlinePreviewHtml) ? (
                <div className="td-field td-field-full">
                  <label>Excel Preview (same page)</label>
                  {inlinePreviewHtml ? (
                    <div
                      style={{
                        width: '100%',
                        height: '380px',
                        overflow: 'auto',
                        border: '1px solid #e2e8f0',
                        borderRadius: '10px',
                        background: '#fff'
                      }}
                      dangerouslySetInnerHTML={{ __html: inlinePreviewHtml }}
                    />
                  ) : (
                    <iframe
                      title="Excel preview"
                      src={inlinePreviewUrl}
                      style={{
                        width: '100%',
                        height: '380px',
                        border: '1px solid #e2e8f0',
                        borderRadius: '10px',
                        background: '#fff'
                      }}
                    />
                  )}
                </div>
              ) : null}
              <div className="td-form-actions">
                <button type="submit" className="td-btn td-btn-primary" disabled={uploading}>
                  {uploading ? (editingId ? 'Saving…' : 'Adding…') : editingId ? 'Save' : 'Add'}
                </button>
                <button type="button" className="td-btn td-btn-secondary" onClick={resetForm} disabled={uploading}>
                  {useChecklistBulkSource ? 'Close' : 'Cancel'}
                </button>
              </div>
            </form>
          </div>
        </div>
      ) : null}
    </div>
  );
};

export default ToDoList;
