"use client";

import { useState, useRef, useCallback } from "react";
import { useRouter } from "next/navigation";

const API_URL = process.env.NEXT_PUBLIC_API_URL || "http://localhost:8000";

interface ConnectedPerson {
  name: string;
  designation: string;
  remuneration: string;
  roles: string;
}

interface FormData {
  company_name: string;
  company_short_name: string;
  nature_of_business: string;
  address: string;
  fiscal_year_start: string;
  fiscal_year_end: string;
  intangibles: string;
  activity_description: string;
  transaction_description: string;
  connected_persons: ConnectedPerson[];
  operating_revenue: string;
  cost_of_sales: string;
  admin_expenses: string;
  other_expenses: string;
  staff_salary: string;
  partner_salaries: string;
}

const defaultPerson: ConnectedPerson = {
  name: "",
  designation: "",
  remuneration: "",
  roles: "",
};

const STEPS = [
  { label: "Company Info", icon: "🏢" },
  { label: "Connected Persons", icon: "👥" },
  { label: "Financials", icon: "💰" },
  { label: "Upload Excel", icon: "📊" },
  { label: "Generate", icon: "📄" },
];

export default function HomePage() {
  const router = useRouter();
  const [step, setStep] = useState(0);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [sessionId, setSessionId] = useState("");
  const [uploadSummary, setUploadSummary] = useState<any>(null);
  const [uploadedFileName, setUploadedFileName] = useState("");
  const [generatedReportId, setGeneratedReportId] = useState("");
  const fileInputRef = useRef<HTMLInputElement>(null);
  const annexure3InputRef = useRef<HTMLInputElement>(null);
  const [annexure3FileName, setAnnexure3FileName] = useState("");
  const [annexure3Loading, setAnnexure3Loading] = useState(false);

  const [formData, setFormData] = useState<FormData>({
    company_name: "",
    company_short_name: "",
    nature_of_business: "",
    address: "",
    fiscal_year_start: "",
    fiscal_year_end: "",
    intangibles: "NA",
    activity_description: "",
    transaction_description: "",
    connected_persons: [{ ...defaultPerson }],
    operating_revenue: "",
    cost_of_sales: "",
    admin_expenses: "",
    other_expenses: "",
    staff_salary: "",
    partner_salaries: "",
  });

  const updateField = (field: keyof FormData, value: string) => {
    setFormData((prev) => ({ ...prev, [field]: value }));
  };

  const updatePerson = (index: number, field: keyof ConnectedPerson, value: string) => {
    setFormData((prev) => {
      const persons = [...prev.connected_persons];
      persons[index] = { ...persons[index], [field]: value };
      return { ...prev, connected_persons: persons };
    });
  };

  const addPerson = () => {
    setFormData((prev) => ({
      ...prev,
      connected_persons: [...prev.connected_persons, { ...defaultPerson }],
    }));
  };

  const removePerson = (index: number) => {
    if (formData.connected_persons.length <= 1) return;
    setFormData((prev) => ({
      ...prev,
      connected_persons: prev.connected_persons.filter((_, i) => i !== index),
    }));
  };

  const handleAnnexure3Upload = async (file: File) => {
    setAnnexure3Loading(true);
    setError("");
    try {
      const fd = new window.FormData();
      fd.append("file", file);
      const res = await fetch(`${API_URL}/api/upload-annexure3`, {
        method: "POST",
        body: fd,
      });
      if (!res.ok) {
        const errData = await res.json();
        throw new Error(errData.detail || "Upload failed");
      }
      const data = await res.json();
      setAnnexure3FileName(data.filename);

      // Auto-fill financials
      const fin = data.financials;
      setFormData((prev) => ({
        ...prev,
        operating_revenue: fin.operating_revenue ? String(fin.operating_revenue) : prev.operating_revenue,
        cost_of_sales: fin.cost_of_sales ? String(fin.cost_of_sales) : prev.cost_of_sales,
        admin_expenses: fin.admin_expenses ? String(fin.admin_expenses) : prev.admin_expenses,
        other_expenses: fin.other_expenses ? String(fin.other_expenses) : prev.other_expenses,
        staff_salary: fin.staff_salary ? String(fin.staff_salary) : prev.staff_salary,
        partner_salaries: fin.partner_salaries ? String(fin.partner_salaries) : prev.partner_salaries,
        // Auto-fill connected persons if found
        connected_persons:
          data.connected_persons && data.connected_persons.length > 0
            ? data.connected_persons.map((p: any) => ({
                name: p.name || "",
                designation: p.designation || "",
                remuneration: p.remuneration || "",
                roles: p.roles || "",
              }))
            : prev.connected_persons,
      }));
    } catch (e: any) {
      setError(e.message || "Failed to parse Annexure 3 file");
    } finally {
      setAnnexure3Loading(false);
    }
  };

  const handleFileUpload = async (file: File) => {
    setLoading(true);
    setError("");
    try {
      const fd = new window.FormData();
      fd.append("file", file);
      const res = await fetch(`${API_URL}/api/upload-excel`, {
        method: "POST",
        body: fd,
      });
      if (!res.ok) {
        const errData = await res.json();
        throw new Error(errData.detail || "Upload failed");
      }
      const data = await res.json();
      setSessionId(data.session_id);
      setUploadSummary(data.summary);
      setUploadedFileName(data.filename);
    } catch (e: any) {
      setError(e.message || "Failed to upload file");
    } finally {
      setLoading(false);
    }
  };

  const handleGenerate = async () => {
    setLoading(true);
    setError("");
    try {
      const payload = {
        company_info: {
          company_name: formData.company_name,
          company_short_name: formData.company_short_name,
          nature_of_business: formData.nature_of_business,
          address: formData.address,
          fiscal_year_start: formData.fiscal_year_start,
          fiscal_year_end: formData.fiscal_year_end,
          intangibles: formData.intangibles,
          activity_description: formData.activity_description,
          transaction_description: formData.transaction_description,
        },
        connected_persons: formData.connected_persons.map((p) => ({
          name: p.name,
          designation: p.designation,
          remuneration: parseFloat(p.remuneration) || 0,
          roles: p.roles,
        })),
        related_parties: [],
        financials: {
          operating_revenue: parseFloat(formData.operating_revenue) || 0,
          cost_of_sales: parseFloat(formData.cost_of_sales) || 0,
          admin_expenses: parseFloat(formData.admin_expenses) || 0,
          other_expenses: parseFloat(formData.other_expenses) || 0,
          staff_salary: parseFloat(formData.staff_salary) || 0,
          partner_salaries: parseFloat(formData.partner_salaries) || 0,
        },
      };

      const fd = new window.FormData();
      fd.append("form_data", JSON.stringify(payload));
      fd.append("session_id", sessionId);

      const res = await fetch(`${API_URL}/api/generate`, {
        method: "POST",
        body: fd,
      });

      if (!res.ok) {
        const errData = await res.json();
        throw new Error(errData.detail || "Generation failed");
      }

      const data = await res.json();
      setGeneratedReportId(data.report_id);

      // Navigate to review page
      router.push(`/review?reportId=${data.report_id}`);
    } catch (e: any) {
      setError(e.message || "Failed to generate report");
    } finally {
      setLoading(false);
    }
  };

  const canProceed = () => {
    switch (step) {
      case 0:
        return formData.company_name && formData.nature_of_business && formData.address && formData.fiscal_year_start && formData.fiscal_year_end;
      case 1:
        return formData.connected_persons.every((p) => p.name && p.designation && p.remuneration);
      case 2:
        return formData.operating_revenue && formData.cost_of_sales;
      case 3:
        return !!sessionId;
      default:
        return true;
    }
  };

  const computedOP = () => {
    const or_ = parseFloat(formData.operating_revenue) || 0;
    const cos = parseFloat(formData.cost_of_sales) || 0;
    const admin = parseFloat(formData.admin_expenses) || 0;
    const other = parseFloat(formData.other_expenses) || 0;
    const staff = parseFloat(formData.staff_salary) || 0;
    const partner = parseFloat(formData.partner_salaries) || 0;
    const oc = cos + admin + other + staff + partner;
    const op = or_ - oc;
    const ratio = or_ > 0 ? ((op / or_) * 100).toFixed(2) : "0.00";
    return { oc: oc.toFixed(2), op: op.toFixed(2), ratio };
  };

  return (
    <>
      <header className="app-header">
        <h1>TP Report Generator</h1>
        <p>Annexure 1 — Transfer Pricing Report (TNMM)</p>
      </header>

      {/* Step Progress */}
      <div className="step-progress">
        {STEPS.map((s, i) => (
          <div key={i} className="step-item">
            <div
              className={`step-circle ${i === step ? "active" : i < step ? "completed" : ""}`}
            >
              {i < step ? "✓" : s.icon}
            </div>
            <span className={`step-label ${i === step ? "active" : ""}`}>{s.label}</span>
            {i < STEPS.length - 1 && (
              <div className={`step-connector ${i < step ? "completed" : ""}`} />
            )}
          </div>
        ))}
      </div>

      {error && <div className="alert alert-error fade-in">{error}</div>}

      {/* Step 0: Company Info */}
      {step === 0 && (
        <div className="glass-card slide-up">
          <h2 className="card-title">Company Information</h2>
          <p className="card-subtitle">Enter the basic details about the company being assessed</p>
          <div className="form-row">
            <div className="form-group">
              <label className="form-label">Company Full Name *</label>
              <input
                className="form-input"
                placeholder="e.g. Amazon Travel and Tourism LLC"
                value={formData.company_name}
                onChange={(e) => updateField("company_name", e.target.value)}
              />
            </div>
            <div className="form-group">
              <label className="form-label">Short Name / Alias</label>
              <input
                className="form-input"
                placeholder="e.g. Amazon LLC"
                value={formData.company_short_name}
                onChange={(e) => updateField("company_short_name", e.target.value)}
              />
            </div>
          </div>

          <div className="form-group">
            <label className="form-label">Nature of Business *</label>
            <textarea
              className="form-textarea"
              placeholder="Detailed description of business activities, services offered, operating model..."
              value={formData.nature_of_business}
              onChange={(e) => updateField("nature_of_business", e.target.value)}
            />
          </div>

          <div className="form-group">
            <label className="form-label">Address *</label>
            <input
              className="form-input"
              placeholder="e.g. UAE"
              value={formData.address}
              onChange={(e) => updateField("address", e.target.value)}
            />
          </div>

          <div className="form-row">
            <div className="form-group">
              <label className="form-label">Fiscal Year Start *</label>
              <input
                className="form-input"
                placeholder="e.g. 1st Feb 2024"
                value={formData.fiscal_year_start}
                onChange={(e) => updateField("fiscal_year_start", e.target.value)}
              />
            </div>
            <div className="form-group">
              <label className="form-label">Fiscal Year End *</label>
              <input
                className="form-input"
                placeholder="e.g. 31st Jan 2025"
                value={formData.fiscal_year_end}
                onChange={(e) => updateField("fiscal_year_end", e.target.value)}
              />
            </div>
          </div>

          <div className="form-group">
            <label className="form-label">Intangible Assets</label>
            <input
              className="form-input"
              placeholder="Describe intangibles or enter NA"
              value={formData.intangibles}
              onChange={(e) => updateField("intangibles", e.target.value)}
            />
          </div>

          <div className="form-row">
            <div className="form-group">
              <label className="form-label">Activity / Service Type *</label>
              <input
                className="form-input"
                placeholder="e.g. Event Management Services, Tour Operating Services"
                value={formData.activity_description}
                onChange={(e) => updateField("activity_description", e.target.value)}
              />
            </div>
            <div className="form-group">
              <label className="form-label">Controlled Transaction Description *</label>
              <input
                className="form-input"
                placeholder="e.g. Event planning, venue management, and catering services"
                value={formData.transaction_description}
                onChange={(e) => updateField("transaction_description", e.target.value)}
              />
            </div>
          </div>
        </div>
      )}

      {/* Step 1: Connected Persons */}
      {step === 1 && (
        <div className="glass-card slide-up">
          <h2 className="card-title">Connected Persons / Related Parties</h2>
          <p className="card-subtitle">Add details for each connected person (KMP, directors, partners)</p>

          {/* Annexure 3 Upload Zone */}
          <input
            ref={annexure3InputRef}
            type="file"
            accept=".xlsx,.xls"
            style={{ display: "none" }}
            onChange={(e) => {
              const file = e.target.files?.[0];
              if (file) handleAnnexure3Upload(file);
            }}
          />
          <div
            className={`upload-zone ${annexure3FileName ? "uploaded" : ""}`}
            style={{ marginBottom: "1.5rem", padding: "1.25rem" }}
            onClick={() => annexure3InputRef.current?.click()}
            onDragOver={(e) => {
              e.preventDefault();
              e.currentTarget.classList.add("drag-over");
            }}
            onDragLeave={(e) => e.currentTarget.classList.remove("drag-over")}
            onDrop={(e) => {
              e.preventDefault();
              e.currentTarget.classList.remove("drag-over");
              const file = e.dataTransfer.files[0];
              if (file) handleAnnexure3Upload(file);
            }}
          >
            <div className="upload-icon" style={{ fontSize: "1.5rem" }}>
              {annexure3Loading ? "..." : annexure3FileName ? "\u2705" : "\uD83D\uDCC2"}
            </div>
            <div className="upload-text" style={{ fontSize: "0.95rem" }}>
              {annexure3Loading
                ? "Parsing Annexure 3..."
                : annexure3FileName
                ? `Auto-filled from: ${annexure3FileName}`
                : "Upload Annexure 3 to auto-fill connected persons & financials"}
            </div>
            <div className="upload-hint">
              {annexure3FileName
                ? "Click to re-upload a different file"
                : "Or enter details manually below"}
            </div>
          </div>

          {formData.connected_persons.map((person, idx) => (
            <div key={idx} className="person-card fade-in">
              <div className="person-header">
                <span className="person-number">Person #{idx + 1}</span>
                {formData.connected_persons.length > 1 && (
                  <button className="remove-btn" onClick={() => removePerson(idx)}>
                    ✕ Remove
                  </button>
                )}
              </div>
              <div className="form-row">
                <div className="form-group">
                  <label className="form-label">Full Name *</label>
                  <input
                    className="form-input"
                    placeholder="e.g. Mr. John Doe"
                    value={person.name}
                    onChange={(e) => updatePerson(idx, "name", e.target.value)}
                  />
                </div>
                <div className="form-group">
                  <label className="form-label">Designation *</label>
                  <input
                    className="form-input"
                    placeholder="e.g. Managing Partner"
                    value={person.designation}
                    onChange={(e) => updatePerson(idx, "designation", e.target.value)}
                  />
                </div>
              </div>
              <div className="form-row">
                <div className="form-group">
                  <label className="form-label">Annual Remuneration (AED) *</label>
                  <input
                    className="form-input"
                    type="number"
                    placeholder="e.g. 180000"
                    value={person.remuneration}
                    onChange={(e) => updatePerson(idx, "remuneration", e.target.value)}
                  />
                </div>
                <div className="form-group">
                  <label className="form-label">Roles & Responsibilities</label>
                  <input
                    className="form-input"
                    placeholder="e.g. Key decision making, Sales vertical..."
                    value={person.roles}
                    onChange={(e) => updatePerson(idx, "roles", e.target.value)}
                  />
                </div>
              </div>
            </div>
          ))}
          <button className="btn-add" onClick={addPerson}>+ Add Connected Person</button>
        </div>
      )}

      {/* Step 2: Financials */}
      {step === 2 && (
        <div className="glass-card slide-up">
          <h2 className="card-title">Financial Information</h2>
          <p className="card-subtitle">Enter the tested party&apos;s financial data (all amounts in AED)</p>
          <div className="form-row">
            <div className="form-group">
              <label className="form-label">Operating Revenue *</label>
              <input
                className="form-input"
                type="number"
                step="0.01"
                placeholder="e.g. 8961920.04"
                value={formData.operating_revenue}
                onChange={(e) => updateField("operating_revenue", e.target.value)}
              />
            </div>
            <div className="form-group">
              <label className="form-label">Cost of Sales *</label>
              <input
                className="form-input"
                type="number"
                step="0.01"
                placeholder="e.g. 5773895.76"
                value={formData.cost_of_sales}
                onChange={(e) => updateField("cost_of_sales", e.target.value)}
              />
            </div>
          </div>
          <div className="form-row">
            <div className="form-group">
              <label className="form-label">Admin & General Expenses</label>
              <input
                className="form-input"
                type="number"
                step="0.01"
                placeholder="e.g. 723121.70"
                value={formData.admin_expenses}
                onChange={(e) => updateField("admin_expenses", e.target.value)}
              />
            </div>
            <div className="form-group">
              <label className="form-label">Other Expenses</label>
              <input
                className="form-input"
                type="number"
                step="0.01"
                placeholder="e.g. 957146.11"
                value={formData.other_expenses}
                onChange={(e) => updateField("other_expenses", e.target.value)}
              />
            </div>
          </div>
          <div className="form-row">
            <div className="form-group">
              <label className="form-label">Staff Salary & Benefits</label>
              <input
                className="form-input"
                type="number"
                step="0.01"
                placeholder="e.g. 624583.23"
                value={formData.staff_salary}
                onChange={(e) => updateField("staff_salary", e.target.value)}
              />
            </div>
            <div className="form-group">
              <label className="form-label">Partners Salaries</label>
              <input
                className="form-input"
                type="number"
                step="0.01"
                placeholder="e.g. 600000"
                value={formData.partner_salaries}
                onChange={(e) => updateField("partner_salaries", e.target.value)}
              />
            </div>
          </div>

          {/* Computed Summary */}
          {formData.operating_revenue && (
            <div className="summary-grid" style={{ marginTop: "1.5rem" }}>
              <div className="summary-item">
                <div className="summary-value">{Number(computedOP().oc).toLocaleString()}</div>
                <div className="summary-label">Total Operating Cost</div>
              </div>
              <div className="summary-item">
                <div className="summary-value">{Number(computedOP().op).toLocaleString()}</div>
                <div className="summary-label">Operating Profit</div>
              </div>
              <div className="summary-item">
                <div className="summary-value">{computedOP().ratio}%</div>
                <div className="summary-label">OP/OR Ratio</div>
              </div>
            </div>
          )}
        </div>
      )}

      {/* Step 3: Upload Excel */}
      {step === 3 && (
        <div className="glass-card slide-up">
          <h2 className="card-title">Upload Benchmarking Data</h2>
          <p className="card-subtitle">Upload Annexure 2 Excel file containing benchmarking workings</p>

          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,.xls"
            style={{ display: "none" }}
            onChange={(e) => {
              const file = e.target.files?.[0];
              if (file) handleFileUpload(file);
            }}
          />

          <div
            className={`upload-zone ${uploadedFileName ? "uploaded" : ""}`}
            onClick={() => fileInputRef.current?.click()}
            onDragOver={(e) => {
              e.preventDefault();
              e.currentTarget.classList.add("drag-over");
            }}
            onDragLeave={(e) => e.currentTarget.classList.remove("drag-over")}
            onDrop={(e) => {
              e.preventDefault();
              e.currentTarget.classList.remove("drag-over");
              const file = e.dataTransfer.files[0];
              if (file) handleFileUpload(file);
            }}
          >
            <div className="upload-icon">{uploadedFileName ? "✅" : "📁"}</div>
            <div className="upload-text">
              {uploadedFileName
                ? `Uploaded: ${uploadedFileName}`
                : "Click or drag & drop your Excel file here"}
            </div>
            <div className="upload-hint">
              {uploadedFileName
                ? "Click to replace with a different file"
                : "Annexure 2 - Excel sheet containing workings (TNMM).xlsx"}
            </div>
          </div>

          {uploadSummary && (
            <div className="upload-summary fade-in">
              <h3 style={{ fontSize: "1rem", fontWeight: 600, marginBottom: "0.5rem", color: "var(--success)" }}>
                ✓ Excel parsed successfully
              </h3>
              <div className="summary-grid">
                <div className="summary-item">
                  <div className="summary-value">{uploadSummary.total_comparables}</div>
                  <div className="summary-label">Comparable Companies</div>
                </div>
                <div className="summary-item">
                  <div className="summary-value">{uploadSummary.search_strategies_count}</div>
                  <div className="summary-label">Search Regions</div>
                </div>
                <div className="summary-item">
                  <div className="summary-value">{uploadSummary.total_rejections}</div>
                  <div className="summary-label">Manual Rejections</div>
                </div>
                {uploadSummary.quartiles && (
                  <div className="summary-item">
                    <div className="summary-value">{uploadSummary.quartiles.median}</div>
                    <div className="summary-label">Median OP/OR</div>
                  </div>
                )}
              </div>
              {uploadSummary.regions && (
                <p style={{ fontSize: "0.85rem", color: "var(--text-muted)", marginTop: "0.75rem" }}>
                  Regions: {uploadSummary.regions.join(" → ")}
                </p>
              )}
            </div>
          )}
        </div>
      )}

      {/* Step 4: Generate */}
      {step === 4 && (
        <div className="glass-card slide-up">
          <h2 className="card-title">Review & Generate Report</h2>
          <p className="card-subtitle">Verify your inputs and generate the Annexure 1 TP Report</p>

          <div style={{ display: "grid", gap: "1rem" }}>
            {/* Company Summary */}
            <div className="person-card">
              <div className="person-number">🏢 Company Information</div>
              <p style={{ color: "var(--text-secondary)", fontSize: "0.9rem", marginTop: "0.5rem" }}>
                <strong>{formData.company_name}</strong><br />
                {formData.address} | FY: {formData.fiscal_year_start} to {formData.fiscal_year_end}
              </p>
            </div>

            {/* Persons Summary */}
            <div className="person-card">
              <div className="person-number">👥 Connected Persons ({formData.connected_persons.length})</div>
              {formData.connected_persons.map((p, i) => (
                <p key={i} style={{ color: "var(--text-secondary)", fontSize: "0.85rem", marginTop: "0.25rem" }}>
                  {p.name} — {p.designation} — AED {Number(p.remuneration).toLocaleString()}
                </p>
              ))}
            </div>

            {/* Financials Summary */}
            <div className="person-card">
              <div className="person-number">💰 Financial Summary</div>
              <div className="summary-grid" style={{ marginTop: "0.5rem" }}>
                <div className="summary-item">
                  <div className="summary-value" style={{ fontSize: "1.1rem" }}>
                    {Number(formData.operating_revenue).toLocaleString()}
                  </div>
                  <div className="summary-label">Revenue (AED)</div>
                </div>
                <div className="summary-item">
                  <div className="summary-value" style={{ fontSize: "1.1rem" }}>
                    {computedOP().ratio}%
                  </div>
                  <div className="summary-label">OP/OR Ratio</div>
                </div>
              </div>
            </div>

            {/* Excel Summary */}
            <div className="person-card">
              <div className="person-number">📊 Benchmarking Data</div>
              <p style={{ color: "var(--text-secondary)", fontSize: "0.85rem", marginTop: "0.25rem" }}>
                {uploadedFileName} — {uploadSummary?.total_comparables} comparable companies
              </p>
            </div>
          </div>

          <div style={{ textAlign: "center", marginTop: "2rem" }}>
            <button
              className="btn btn-success"
              onClick={handleGenerate}
              disabled={loading}
              style={{ padding: "1rem 3rem", fontSize: "1.1rem" }}
            >
              {loading ? (
                <>
                  <span className="loading-spinner" /> Generating Report...
                </>
              ) : (
                "🚀 Generate Annexure 1 Report"
              )}
            </button>
          </div>
        </div>
      )}

      {/* Navigation Buttons */}
      <div className="btn-group">
        {step > 0 && (
          <button className="btn btn-secondary" onClick={() => setStep(step - 1)}>
            ← Back
          </button>
        )}
        <div style={{ flex: 1 }} />
        {step < STEPS.length - 1 && (
          <button
            className="btn btn-primary"
            onClick={() => setStep(step + 1)}
            disabled={!canProceed()}
          >
            Next →
          </button>
        )}
      </div>

      {/* Loading Overlay */}
      {loading && (
        <div className="loading-overlay">
          <div className="loading-card">
            <div className="loading-spinner" />
            <p style={{ color: "var(--text-primary)", fontWeight: 600, marginTop: "1rem" }}>
              {step === 3 ? "Parsing Excel file..." : "Generating your TP Report..."}
            </p>
            <p style={{ color: "var(--text-muted)", fontSize: "0.85rem", marginTop: "0.5rem" }}>
              This may take a moment
            </p>
          </div>
        </div>
      )}
    </>
  );
}
