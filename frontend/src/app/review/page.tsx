"use client";

import { useState, useEffect, useCallback, Suspense } from "react";
import { useSearchParams, useRouter } from "next/navigation";

const API_URL = process.env.NEXT_PUBLIC_API_URL || "http://localhost:8000";

interface Section {
    heading: string;
    style: string;
    paragraphs: { text: string; style: string }[];
}

interface TableData {
    index: number;
    rows: string[][];
}

function ReviewContent() {
    const searchParams = useSearchParams();
    const router = useRouter();
    const reportId = searchParams.get("reportId") || "";

    const [sections, setSections] = useState<Section[]>([]);
    const [tables, setTables] = useState<TableData[]>([]);
    const [loading, setLoading] = useState(true);
    const [saving, setSaving] = useState(false);
    const [error, setError] = useState("");
    const [activeSection, setActiveSection] = useState(0);
    const [editedParagraphs, setEditedParagraphs] = useState<Record<string, string>>({});
    const [saveMessage, setSaveMessage] = useState("");

    useEffect(() => {
        if (!reportId) return;
        fetchContent();
    }, [reportId]);

    const fetchContent = async () => {
        try {
            const res = await fetch(`${API_URL}/api/report-content/${reportId}`);
            if (!res.ok) throw new Error("Failed to load report content");
            const data = await res.json();
            setSections(data.sections);
            setTables(data.tables);
        } catch (e: any) {
            setError(e.message);
        } finally {
            setLoading(false);
        }
    };

    const handleParagraphEdit = (sectionIdx: number, paraIdx: number, newText: string) => {
        const key = `${sectionIdx}-${paraIdx}`;
        const original = sections[sectionIdx]?.paragraphs[paraIdx]?.text || "";
        if (newText !== original) {
            setEditedParagraphs((prev) => ({ ...prev, [key]: newText }));
        }
    };

    const handleSaveEdits = async () => {
        if (Object.keys(editedParagraphs).length === 0) {
            setSaveMessage("No changes to save");
            setTimeout(() => setSaveMessage(""), 2000);
            return;
        }

        setSaving(true);
        setError("");

        // Build replacement map: old text -> new text
        const updates: Record<string, string> = {};
        for (const [key, newText] of Object.entries(editedParagraphs)) {
            const [sIdx, pIdx] = key.split("-").map(Number);
            const original = sections[sIdx]?.paragraphs[pIdx]?.text || "";
            if (original && original !== newText) {
                updates[original] = newText;
            }
        }

        try {
            const fd = new FormData();
            fd.append("report_id", reportId);
            fd.append("updates", JSON.stringify(updates));

            const res = await fetch(`${API_URL}/api/update-report`, {
                method: "POST",
                body: fd,
            });

            if (!res.ok) throw new Error("Failed to save edits");

            setSaveMessage("Changes saved successfully!");
            setEditedParagraphs({});

            // Refresh content
            await fetchContent();
            setTimeout(() => setSaveMessage(""), 3000);
        } catch (e: any) {
            setError(e.message);
        } finally {
            setSaving(false);
        }
    };

    const handleDownload = () => {
        window.open(`${API_URL}/api/download/${reportId}`, "_blank");
    };

    if (loading) {
        return (
            <div style={{ textAlign: "center", padding: "4rem" }}>
                <div className="loading-spinner" style={{ width: 40, height: 40, borderWidth: 3, margin: "0 auto" }} />
                <p style={{ color: "var(--text-secondary)", marginTop: "1rem" }}>Loading report content...</p>
            </div>
        );
    }

    return (
        <>
            <header className="app-header">
                <h1>Review & Edit Report</h1>
                <p>Click on any paragraph to edit it. Changes are saved back to the document.</p>
            </header>

            {error && <div className="alert alert-error fade-in">{error}</div>}
            {saveMessage && <div className="alert alert-success fade-in">{saveMessage}</div>}

            {/* Action Bar */}
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: "1.5rem", flexWrap: "wrap", gap: "0.75rem" }}>
                <button className="btn btn-secondary" onClick={() => router.push("/")}>
                    ← Back to Form
                </button>
                <div style={{ display: "flex", gap: "0.75rem" }}>
                    <button
                        className="btn btn-primary"
                        onClick={handleSaveEdits}
                        disabled={saving || Object.keys(editedParagraphs).length === 0}
                    >
                        {saving ? <><span className="loading-spinner" /> Saving...</> : `💾 Save Edits (${Object.keys(editedParagraphs).length})`}
                    </button>
                    <button className="btn btn-success" onClick={handleDownload}>
                        📥 Download DOCX
                    </button>
                </div>
            </div>

            {/* Editor */}
            <div className="editor-container">
                {/* Sidebar */}
                <div className="editor-sidebar">
                    <div className="sidebar-title">Sections</div>
                    {sections.map((section, idx) => (
                        <button
                            key={idx}
                            className={`sidebar-item ${idx === activeSection ? "active" : ""}`}
                            onClick={() => {
                                setActiveSection(idx);
                                const el = document.getElementById(`section-${idx}`);
                                el?.scrollIntoView({ behavior: "smooth", block: "start" });
                            }}
                        >
                            {section.heading.length > 30 ? section.heading.slice(0, 30) + "..." : section.heading}
                        </button>
                    ))}
                </div>

                {/* Main Content */}
                <div className="editor-main">
                    {sections.map((section, sIdx) => (
                        <div key={sIdx} id={`section-${sIdx}`} className="editor-section">
                            <h3 className="editor-section-title">{section.heading}</h3>
                            {section.paragraphs.map((para, pIdx) => (
                                <div
                                    key={pIdx}
                                    className="editor-paragraph"
                                    contentEditable
                                    suppressContentEditableWarning
                                    onBlur={(e) => {
                                        const newText = e.currentTarget.innerText;
                                        handleParagraphEdit(sIdx, pIdx, newText);
                                    }}
                                >
                                    {editedParagraphs[`${sIdx}-${pIdx}`] || para.text}
                                </div>
                            ))}
                        </div>
                    ))}

                    {/* Tables Section */}
                    {tables.length > 0 && (
                        <div className="editor-section">
                            <h3 className="editor-section-title">Tables</h3>
                            {tables.map((table, tIdx) => (
                                <div key={tIdx} style={{ marginBottom: "1.5rem", overflowX: "auto" }}>
                                    <p style={{ fontSize: "0.8rem", color: "var(--text-muted)", marginBottom: "0.5rem" }}>
                                        Table {tIdx + 1}
                                    </p>
                                    <table className="editor-table">
                                        <tbody>
                                            {table.rows.slice(0, 20).map((row, rIdx) => (
                                                <tr key={rIdx}>
                                                    {row.map((cell, cIdx) => {
                                                        const Tag = rIdx === 0 ? "th" : "td";
                                                        return <Tag key={cIdx}>{cell}</Tag>;
                                                    })}
                                                </tr>
                                            ))}
                                        </tbody>
                                    </table>
                                    {table.rows.length > 20 && (
                                        <p style={{ fontSize: "0.75rem", color: "var(--text-muted)", marginTop: "0.5rem" }}>
                                            Showing first 20 of {table.rows.length} rows
                                        </p>
                                    )}
                                </div>
                            ))}
                        </div>
                    )}
                </div>
            </div>
        </>
    );
}

export default function ReviewPage() {
    return (
        <Suspense fallback={
            <div style={{ textAlign: "center", padding: "4rem" }}>
                <div className="loading-spinner" style={{ width: 40, height: 40, borderWidth: 3, margin: "0 auto" }} />
                <p style={{ color: "var(--text-secondary)", marginTop: "1rem" }}>Loading...</p>
            </div>
        }>
            <ReviewContent />
        </Suspense>
    );
}
