'use client';

import { useState, FormEvent } from 'react';

const APP_PASSWORD = 'BOP2026';

export default function Home() {
  const [authed, setAuthed] = useState(false);
  const [pwInput, setPwInput] = useState('');
  const [pwError, setPwError] = useState('');

  const [excels, setExcels] = useState<FileList | null>(null);
  const [emailPdf, setEmailPdf] = useState<File | null>(null);
  const [receipts, setReceipts] = useState<FileList | null>(null);
  const [titleNumber, setTitleNumber] = useState('');
  const [dateOverride, setDateOverride] = useState('');
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState('');
  const [info, setInfo] = useState('');

  function handleLogin(e: FormEvent) {
    e.preventDefault();
    if (pwInput === APP_PASSWORD) {
      setAuthed(true);
      setPwError('');
    } else {
      setPwError('Incorrect password.');
    }
  }

  async function handleSubmit(e: FormEvent) {
    e.preventDefault();
    setError('');
    setInfo('');

    if (!excels || excels.length === 0) {
      setError('Please upload at least one Excel file.');
      return;
    }
    if (!emailPdf) {
      setError('Please upload the CNX upload-notification email PDF.');
      return;
    }
    if (!titleNumber.trim()) {
      setError('Please enter the CNX Title # (e.g., 2120).');
      return;
    }
    if (!/^\d+$/.test(titleNumber.trim())) {
      setError('CNX Title # should be digits only (e.g., 2120).');
      return;
    }

    const formData = new FormData();
    Array.from(excels).forEach((f) => formData.append('excel', f));
    formData.append('email', emailPdf);
    if (receipts) Array.from(receipts).forEach((f) => formData.append('receipt', f));
    formData.append('titleNumber', titleNumber.trim());
    if (dateOverride) formData.append('dateOverride', dateOverride);

    setSubmitting(true);
    setInfo('Generating invoices… this can take 10–30 seconds.');
    try {
      const res = await fetch('/api/generate', { method: 'POST', body: formData });
      if (!res.ok) {
        const msg = await res.text();
        throw new Error(msg || `Server error ${res.status}`);
      }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `cnx-title-invoices-CC-${titleNumber.trim()}.zip`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      setInfo('Download started.');
    } catch (err: any) {
      setError(err.message || 'Generation failed.');
      setInfo('');
    } finally {
      setSubmitting(false);
    }
  }

  if (!authed) {
    return (
      <div className="container">
        <h1>BOP CNX Title Invoice Generator</h1>
        <p className="subtitle">Internal use only — password required.</p>
        <form onSubmit={handleLogin}>
          <div className="field">
            <label>Password</label>
            <input
              type="password"
              value={pwInput}
              onChange={(e) => setPwInput(e.target.value)}
              autoFocus
            />
          </div>
          <button type="submit">Sign In</button>
          {pwError && <div className="status error">{pwError}</div>}
        </form>
      </div>
    );
  }

  return (
    <div className="container">
      <h1>BOP CNX Title Invoice Generator</h1>
      <p className="subtitle">
        Upload Excel work logs, the CNX upload-notification email PDF, and any
        receipt/exception order PDFs. Output: a ZIP of CNX-format invoices.
      </p>

      <form onSubmit={handleSubmit}>
        <div className="field">
          <label>CNX Title # (digits only)</label>
          <input
            type="text"
            inputMode="numeric"
            pattern="[0-9]*"
            placeholder="e.g., 2120"
            value={titleNumber}
            onChange={(e) => setTitleNumber(e.target.value.replace(/\D/g, ''))}
          />
          <div className="hint">
            The handwritten number on the email (e.g., the &quot;2120&quot; in &quot;CC 2120&quot;).
            App will prefix &quot;CC &quot; automatically.
          </div>
        </div>

        <div className="field">
          <label>Excel Files (one per invoice)</label>
          <input
            type="file"
            accept=".xlsx,.xls"
            multiple
            onChange={(e) => setExcels(e.target.files)}
          />
          <div className="hint">Multi-select supported.</div>
        </div>

        <div className="field">
          <label>CNX Upload Notification Email PDF</label>
          <input
            type="file"
            accept="application/pdf"
            onChange={(e) => setEmailPdf(e.target.files?.[0] ?? null)}
          />
          <div className="hint">
            In Outlook, open the email and use <strong>File → Save As → PDF</strong> (not Print to PDF — that produces an image-only PDF the app can&apos;t read).
          </div>
        </div>

        <div className="field">
          <label>Receipt / Exception Order PDFs (optional)</label>
          <input
            type="file"
            accept="application/pdf"
            multiple
            onChange={(e) => setReceipts(e.target.files)}
          />
          <div className="hint">Filenames must contain the invoice number.</div>
        </div>

        <div className="field">
          <label>Invoice Date Override (optional)</label>
          <input
            type="date"
            value={dateOverride}
            onChange={(e) => setDateOverride(e.target.value)}
          />
          <div className="hint">
            Filename date snaps to nearest 15th or end-of-month.
          </div>
        </div>

        <button type="submit" disabled={submitting}>
          {submitting ? 'Generating…' : 'Generate Invoices'}
        </button>

        {error && <div className="status error">{error}</div>}
        {info && !error && <div className="status info">{info}</div>}
      </form>
    </div>
  );
}
