import React, { useState, useEffect, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import {
  Search, Upload, LogOut, LayoutDashboard, FileSpreadsheet, Settings as SettingsIcon,
  Users, CheckCircle2, XCircle, Eye, EyeOff, Trash2, Pencil, Download, Printer,
  ShieldCheck, ChevronRight, AlertTriangle, Loader2, GraduationCap, KeyRound,
  BadgeCheck, ArrowLeft, RefreshCw, Menu, X, ClipboardList
} from "lucide-react";

/* ----------------------------------------------------------------------
   CONSTANTS
---------------------------------------------------------------------- */

const NAVY = "#0B2545";
const NAVY_DEEP = "#081b34";
const BLUE = "#2E6CE0";
const BLUE_SOFT = "#EAF1FD";
const PAPER = "#FBFBFA";
const OK = "#1B7A4A";
const BAD = "#B4382C";
const INK = "#1A2433";

const DEFAULT_SETTINGS = {
  institution_name: "Greenfield Public School & College",
  logo_text: "GS",
  address: "12 Lake View Road, Hyderabad, Telangana",
  contact: "results@greenfield.edu.in  |  +91 40 4567 8901",
  examination_name: "Annual Examination",
  session: "2025-26",
  result_date: new Date().toISOString().slice(0, 10),
  memo_title: "Statement of Marks",
  footer_text: "This is a computer-generated memo. For queries, contact the examination office.",
  signature_text: "Controller of Examinations",
  passing_percentage: 35,
  default_max_marks: 100,
};

const DEFAULT_ADMIN_EMAIL = "mdsahil44414@gmail.com";
const DEFAULT_ADMIN_PASSWORD = "Sahil@123";

// Canonical field -> accepted header aliases (normalized: lowercase, alnum only)
const ALIASES = {
  roll_number: ["rollno", "rollnumber", "roll", "rollnum", "hallticketno", "htno", "admissionno"],
  registration_number: ["regno", "registrationno", "registrationnumber", "enrollmentno", "enrollmentnumber", "enrollno", "enrolno"],
  student_name: ["name", "studentname", "candidatename", "fullname"],
  course: ["course", "class", "branch", "program", "stream", "coursename"],
  semester: ["semester", "sem", "yearofstudy"],
  examination: ["examination", "exam", "examname"],
  session: ["session", "examsession", "academicyear", "examyear", "academicsession"],
  total_obtained: ["totalmarks", "totalobtained", "totalmarksobtained", "grandtotal", "total"],
  total_max: ["maxtotal", "totalmaxmarks", "maximumtotalmarks", "outof", "totaloutof"],
  percentage: ["percentage", "percent", "pct"],
  grade: ["grade", "overallgrade", "finalgrade"],
  result_status: ["result", "status", "resultstatus", "passfail"],
  division: ["division", "classobtained", "divisionobtained"],
};

const norm = (s) => String(s ?? "").toLowerCase().replace(/[^a-z0-9]/g, "").trim();

function uid(prefix = "") {
  return (
    prefix +
    Date.now().toString(36) +
    Math.random().toString(36).slice(2, 10)
  ).toUpperCase();
}

async function sha256Hex(text) {
  const enc = new TextEncoder().encode(text);
  const buf = await crypto.subtle.digest("SHA-256", enc);
  return Array.from(new Uint8Array(buf)).map((b) => b.toString(16).padStart(2, "0")).join("");
}

function gradeFromPercentage(pct) {
  if (pct >= 90) return "A+";
  if (pct >= 80) return "A";
  if (pct >= 70) return "B+";
  if (pct >= 60) return "B";
  if (pct >= 50) return "C";
  if (pct >= 35) return "D";
  return "F";
}

/* ----------------------------------------------------------------------
   STORAGE HELPERS  (all institution data is shared so every visitor —
   students and admins — sees the same published results & settings)
---------------------------------------------------------------------- */

async function storageGet(key, shared) {
  try {
    const r = await window.storage.get(key, shared);
    return r ? JSON.parse(r.value) : null;
  } catch {
    return null;
  }
}
async function storageSet(key, value, shared) {
  try {
    await window.storage.set(key, JSON.stringify(value), shared);
    return true;
  } catch {
    return false;
  }
}

/* ----------------------------------------------------------------------
   ROOT APP
---------------------------------------------------------------------- */

export default function App() {
  const [booting, setBooting] = useState(true);
  const [settings, setSettings] = useState(DEFAULT_SETTINGS);
  const [results, setResults] = useState([]);
  const [adminCreds, setAdminCreds] = useState(null);
  const [view, setView] = useState("student"); // student | verify | adminLogin | admin
  const [session, setSession] = useState(null); // { email }
  const [toast, setToast] = useState(null);

  useEffect(() => {
    (async () => {
      const [s, r, a] = await Promise.all([
        storageGet("settings", true),
        storageGet("results", true),
        storageGet("admin-creds", true),
      ]);
      setSettings(s || DEFAULT_SETTINGS);
      setResults(r || []);
      if (a) {
        setAdminCreds(a);
      } else {
        const hash = await sha256Hex(DEFAULT_ADMIN_EMAIL + ":" + DEFAULT_ADMIN_PASSWORD);
        const fresh = { email: DEFAULT_ADMIN_EMAIL, hash };
        await storageSet("admin-creds", fresh, true);
        setAdminCreds(fresh);
      }
      setBooting(false);
    })();
  }, []);

  function notify(text, kind = "ok") {
    setToast({ text, kind });
    window.clearTimeout(notify._t);
    notify._t = window.setTimeout(() => setToast(null), 3200);
  }

  async function persistResults(next) {
    const prev = results;
    setResults(next);
    const ok = await storageSet("results", next, true);
    if (!ok) {
      setResults(prev);
      notify("Save failed: could not write to storage. Your change was not saved — please try again.", "error");
    }
    return ok;
  }
  async function persistSettings(next) {
    const prev = settings;
    setSettings(next);
    const ok = await storageSet("settings", next, true);
    if (!ok) {
      setSettings(prev);
      notify("Save failed: could not write settings to storage. Please try again.", "error");
    }
    return ok;
  }

  if (booting) {
    return (
      <div style={{ background: PAPER }} className="min-h-screen flex items-center justify-center">
        <Loader2 className="animate-spin" size={28} color={NAVY} />
      </div>
    );
  }

  return (
    <div style={{ background: PAPER, color: INK, fontFamily: "Inter, system-ui, sans-serif" }} className="min-h-screen">
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Fraunces:opsz,wght@9..144,500..700&display=swap');
        .font-display { font-family: 'Fraunces', serif; }
        @media print {
          body * { visibility: hidden; }
          #printable-memo, #printable-memo * { visibility: visible; }
          #printable-memo { position: absolute; inset: 0; width: 100%; }
        }
      `}</style>

      {toast && (
        <div
          className="fixed top-4 right-4 z-50 px-4 py-3 rounded-xl shadow-lg text-sm font-medium text-white print:hidden"
          style={{ background: toast.kind === "error" ? BAD : toast.kind === "warn" ? "#B07A17" : OK }}
        >
          {toast.text}
        </div>
      )}

      {view === "student" && (
        <StudentPortal
          settings={settings}
          results={results}
          notify={notify}
          onGoVerify={() => setView("verify")}
          onGoAdmin={() => setView(session ? "admin" : "adminLogin")}
        />
      )}

      {view === "verify" && (
        <VerifyPage settings={settings} results={results} onBack={() => setView("student")} />
      )}

      {view === "adminLogin" && (
        <AdminLogin
          settings={settings}
          adminCreds={adminCreds}
          onBack={() => setView("student")}
          onLoggedIn={(email) => {
            setSession({ email });
            setView("admin");
            notify("Signed in as " + email);
          }}
        />
      )}

      {view === "admin" &&
        (session ? (
          <AdminDashboard
            settings={settings}
            results={results}
            adminCreds={adminCreds}
            session={session}
            notify={notify}
            onLogout={() => {
              setSession(null);
              setView("student");
              notify("Signed out");
            }}
            onSaveSettings={persistSettings}
            onSaveResults={persistResults}
            onSaveCreds={async (next) => {
              const prev = adminCreds;
              setAdminCreds(next);
              const ok = await storageSet("admin-creds", next, true);
              if (!ok) {
                setAdminCreds(prev);
                notify("Save failed: could not write your new password to storage. Please try again.", "error");
              }
              return ok;
            }}
            onViewStudentSite={() => setView("student")}
          />
        ) : (
          <AdminLogin
            settings={settings}
            adminCreds={adminCreds}
            onBack={() => setView("student")}
            onLoggedIn={(email) => {
              setSession({ email });
              notify("Signed in as " + email);
            }}
          />
        ))}
    </div>
  );
}

/* ----------------------------------------------------------------------
   SHARED UI BITS
---------------------------------------------------------------------- */

function Logo({ settings, size = 40 }) {
  return (
    <div
      className="flex items-center justify-center rounded-xl font-display font-semibold text-white shrink-0"
      style={{ width: size, height: size, background: `linear-gradient(155deg, ${NAVY}, ${BLUE})`, fontSize: size * 0.4 }}
    >
      {settings.logo_text || "S"}
    </div>
  );
}

function Field({ label, children, hint }) {
  return (
    <label className="block mb-4">
      <span className="block text-xs font-semibold tracking-wide uppercase mb-1.5" style={{ color: "#5B6474" }}>{label}</span>
      {children}
      {hint && <span className="block text-xs mt-1 text-gray-400">{hint}</span>}
    </label>
  );
}

const inputCls =
  "w-full rounded-lg border px-3.5 py-2.5 text-sm outline-none transition focus:ring-2";
const inputStyle = { borderColor: "#D8DEE9", background: "#fff" };

function PrimaryButton({ children, onClick, disabled, type = "button", full, icon: Icon }) {
  return (
    <button
      type={type}
      onClick={onClick}
      disabled={disabled}
      className={`inline-flex items-center justify-center gap-2 rounded-full px-5 py-2.5 text-sm font-semibold text-white transition disabled:opacity-50 disabled:cursor-not-allowed ${full ? "w-full" : ""}`}
      style={{ background: `linear-gradient(155deg, ${BLUE}, ${NAVY})` }}
    >
      {Icon && <Icon size={16} />}
      {children}
    </button>
  );
}

function GhostButton({ children, onClick, icon: Icon, danger }) {
  return (
    <button
      onClick={onClick}
      className="inline-flex items-center justify-center gap-2 rounded-full px-4 py-2 text-sm font-semibold border transition hover:bg-gray-50"
      style={{ borderColor: danger ? "#F0C4BE" : "#D8DEE9", color: danger ? BAD : INK }}
    >
      {Icon && <Icon size={15} />}
      {children}
    </button>
  );
}

function StatusPill({ status }) {
  const pass = status === "PASS";
  return (
    <span
      className="inline-flex items-center gap-1.5 rounded-full px-3 py-1 text-xs font-bold tracking-wide"
      style={{ background: pass ? "#E6F5EC" : "#FBEAE8", color: pass ? OK : BAD }}
    >
      {pass ? <CheckCircle2 size={13} /> : <XCircle size={13} />}
      {status}
    </span>
  );
}

/* ----------------------------------------------------------------------
   STUDENT PORTAL
---------------------------------------------------------------------- */

function StudentPortal({ settings, results, notify, onGoVerify, onGoAdmin }) {
  const [roll, setRoll] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [record, setRecord] = useState(null);
  const [attempts, setAttempts] = useState(0);
  const [cooldownUntil, setCooldownUntil] = useState(0);

  function search(e) {
    e.preventDefault();
    setError("");
    setRecord(null);
    const now = Date.now();
    if (now < cooldownUntil) {
      setError("Too many attempts. Please wait a few seconds and try again.");
      return;
    }
    const q = norm(roll);
    if (!q) {
      setError("Please enter your roll number.");
      return;
    }
    setLoading(true);
    window.setTimeout(() => {
      const match = results.find((r) => r.published && norm(r.roll_number) === q);
      if (!match) {
        const next = attempts + 1;
        setAttempts(next);
        if (next >= 5) {
          setCooldownUntil(Date.now() + 15000);
          setAttempts(0);
        }
        setError("Result Not Found — please double-check your roll number and try again.");
      } else {
        setRecord(match);
      }
      setLoading(false);
    }, 450);
  }

  return (
    <div className="min-h-screen flex flex-col">
      <header className="border-b bg-white print:hidden" style={{ borderColor: "#E7EAF0" }}>
        <div className="max-w-5xl mx-auto px-5 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <Logo settings={settings} />
            <div>
              <div className="font-display font-semibold leading-tight" style={{ color: NAVY }}>{settings.institution_name}</div>
              <div className="text-xs text-gray-500">{settings.examination_name} &middot; {settings.session}</div>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <button onClick={onGoVerify} className="text-xs font-semibold text-gray-500 hover:text-gray-800 hidden sm:inline-flex items-center gap-1">
              <ShieldCheck size={14} /> Verify a result
            </button>
            <button onClick={onGoAdmin} className="text-xs font-semibold text-gray-500 hover:text-gray-800 inline-flex items-center gap-1">
              <KeyRound size={14} /> Admin
            </button>
          </div>
        </div>
      </header>

      <main className="flex-1 max-w-3xl w-full mx-auto px-5 py-12 print:hidden">
        {!record && (
          <div className="text-center mb-8">
            <h1 className="font-display text-3xl sm:text-4xl font-semibold" style={{ color: NAVY }}>Examination Results Portal</h1>
            <p className="text-gray-500 mt-2">Enter your Roll Number to view your result</p>
          </div>
        )}

        {!record && (
          <form onSubmit={search} className="bg-white rounded-2xl shadow-sm border p-6 sm:p-8 max-w-md mx-auto" style={{ borderColor: "#E7EAF0" }}>
            <Field label="Roll Number">
              <input
                autoFocus
                value={roll}
                onChange={(e) => setRoll(e.target.value)}
                className={inputCls}
                style={inputStyle}
                placeholder="e.g. 22CS0451"
              />
            </Field>
            <PrimaryButton type="submit" full disabled={loading} icon={loading ? Loader2 : Search}>
              {loading ? "Searching…" : "Search Result"}
            </PrimaryButton>
            {error && (
              <div className="mt-4 flex items-start gap-2 text-sm rounded-lg px-3 py-2.5" style={{ background: "#FBEAE8", color: BAD }}>
                <AlertTriangle size={16} className="mt-0.5 shrink-0" /> {error}
              </div>
            )}
            <button type="button" onClick={onGoVerify} className="sm:hidden mt-4 text-xs font-semibold text-gray-500 inline-flex items-center gap-1">
              <ShieldCheck size={14} /> Verify a result instead
            </button>
          </form>
        )}

        {record && <ResultCard record={record} settings={settings} notify={notify} onBack={() => setRecord(null)} />}
      </main>

      {record && <PrintableMemo record={record} settings={settings} />}

      <footer className="text-center text-xs text-gray-400 py-6 print:hidden">
        {settings.footer_text}
      </footer>
    </div>
  );
}

function ResultRow({ label, value }) {
  if (value === undefined || value === null || value === "") return null;
  return (
    <div className="flex justify-between gap-4 py-1.5 text-sm border-b last:border-0" style={{ borderColor: "#EEF1F5" }}>
      <span className="text-gray-500">{label}</span>
      <span className="font-medium text-right">{value}</span>
    </div>
  );
}

function ResultCard({ record, settings, notify, onBack }) {
  function printMemo() {
    try {
      window.print();
    } catch (err) {
      notify?.("Couldn't open the print dialog: " + (err?.message || "your browser blocked it") + ". Try again or use your browser's Print option from the menu.", "error");
    }
  }
  return (
    <div>
      <button onClick={onBack} className="mb-4 text-sm text-gray-500 inline-flex items-center gap-1 hover:text-gray-800">
        <ArrowLeft size={15} /> Search another roll number
      </button>
      <div className="bg-white rounded-2xl shadow-sm border overflow-hidden" style={{ borderColor: "#E7EAF0" }}>
        <div className="px-6 sm:px-8 py-6" style={{ background: `linear-gradient(155deg, ${NAVY}, ${NAVY_DEEP})` }}>
          <div className="flex items-center justify-between gap-4 flex-wrap">
            <div>
              <div className="text-white font-display text-xl font-semibold">{record.student_name}</div>
              <div className="text-white/70 text-sm">Roll No. {record.roll_number}</div>
            </div>
            <StatusPill status={record.result_status} />
          </div>
        </div>

        <div className="px-6 sm:px-8 py-6 grid sm:grid-cols-2 gap-x-8">
          <div>
            <ResultRow label="Registration No." value={record.registration_number} />
            <ResultRow label="Course / Class" value={record.course} />
            <ResultRow label="Semester / Year" value={record.semester} />
            <ResultRow label="Examination" value={record.examination || settings.examination_name} />
            <ResultRow label="Session" value={record.session || settings.session} />
          </div>
          <div>
            <ResultRow label="Total Marks" value={`${record.total_obtained} / ${record.total_max}`} />
            <ResultRow label="Percentage" value={`${record.percentage}%`} />
            <ResultRow label="Grade" value={record.grade} />
            <ResultRow label="Division" value={record.division} />
            <ResultRow label="Verification ID" value={record.verification_id} />
          </div>
        </div>

        {record.subjects?.length > 0 && (
          <div className="px-6 sm:px-8 pb-6">
            <div className="text-xs font-semibold uppercase tracking-wide text-gray-500 mb-2">Subject-wise Marks</div>
            <div className="overflow-x-auto rounded-xl border" style={{ borderColor: "#EEF1F5" }}>
              <table className="w-full text-sm">
                <thead>
                  <tr style={{ background: BLUE_SOFT }}>
                    <th className="text-left px-3 py-2 font-semibold">Subject</th>
                    <th className="text-right px-3 py-2 font-semibold">Max</th>
                    <th className="text-right px-3 py-2 font-semibold">Obtained</th>
                    {record.subjects.some((s) => s.grade) && <th className="text-right px-3 py-2 font-semibold">Grade</th>}
                  </tr>
                </thead>
                <tbody>
                  {record.subjects.map((s, i) => (
                    <tr key={i} className="border-t" style={{ borderColor: "#EEF1F5" }}>
                      <td className="px-3 py-2">{s.name}</td>
                      <td className="px-3 py-2 text-right text-gray-500">{s.max}</td>
                      <td className="px-3 py-2 text-right font-medium">{s.obtained}</td>
                      {record.subjects.some((x) => x.grade) && <td className="px-3 py-2 text-right text-gray-500">{s.grade || "—"}</td>}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        <div className="px-6 sm:px-8 pb-7 flex flex-wrap gap-3">
          <PrimaryButton icon={Printer} onClick={printMemo}>Print Memo</PrimaryButton>
          <GhostButton icon={Download} onClick={printMemo}>Download PDF</GhostButton>
        </div>
        <div className="px-6 sm:px-8 pb-6 -mt-2 text-xs text-gray-400">
          "Download PDF" opens your browser's print dialog — choose "Save as PDF" as the destination.
        </div>
      </div>
    </div>
  );
}

function PrintableMemo({ record, settings }) {
  return (
    <div id="printable-memo" className="hidden print:block bg-white p-10">
      <div className="flex items-center gap-4 border-b-2 pb-4 mb-6" style={{ borderColor: NAVY }}>
        <div className="w-14 h-14 rounded-lg flex items-center justify-center text-white font-display font-semibold text-xl" style={{ background: NAVY }}>
          {settings.logo_text}
        </div>
        <div>
          <div className="font-display text-xl font-semibold" style={{ color: NAVY }}>{settings.institution_name}</div>
          <div className="text-xs text-gray-600">{settings.address}</div>
          <div className="text-xs text-gray-600">{settings.contact}</div>
        </div>
      </div>
      <div className="text-center mb-6">
        <div className="font-display text-lg font-semibold">{settings.memo_title}</div>
        <div className="text-sm text-gray-600">{record.examination || settings.examination_name} &middot; {record.session || settings.session}</div>
      </div>
      <table className="w-full text-sm mb-6">
        <tbody>
          <tr><td className="py-1 text-gray-500 w-1/3">Student Name</td><td className="py-1 font-medium">{record.student_name}</td></tr>
          <tr><td className="py-1 text-gray-500">Roll Number</td><td className="py-1 font-medium">{record.roll_number}</td></tr>
          {record.registration_number && <tr><td className="py-1 text-gray-500">Registration No.</td><td className="py-1 font-medium">{record.registration_number}</td></tr>}
          {record.course && <tr><td className="py-1 text-gray-500">Course / Class</td><td className="py-1 font-medium">{record.course}</td></tr>}
          {record.semester && <tr><td className="py-1 text-gray-500">Semester / Year</td><td className="py-1 font-medium">{record.semester}</td></tr>}
        </tbody>
      </table>
      {record.subjects?.length > 0 && (
        <table className="w-full text-sm mb-6 border-collapse">
          <thead>
            <tr>
              <th className="text-left border-b-2 pb-1" style={{ borderColor: NAVY }}>Subject</th>
              <th className="text-right border-b-2 pb-1" style={{ borderColor: NAVY }}>Max Marks</th>
              <th className="text-right border-b-2 pb-1" style={{ borderColor: NAVY }}>Marks Obtained</th>
            </tr>
          </thead>
          <tbody>
            {record.subjects.map((s, i) => (
              <tr key={i}><td className="py-1 border-b" style={{ borderColor: "#eee" }}>{s.name}</td><td className="py-1 text-right border-b" style={{ borderColor: "#eee" }}>{s.max}</td><td className="py-1 text-right border-b" style={{ borderColor: "#eee" }}>{s.obtained}</td></tr>
            ))}
          </tbody>
        </table>
      )}
      <table className="w-full text-sm mb-8">
        <tbody>
          <tr><td className="py-1 text-gray-500 w-1/3">Total Marks</td><td className="py-1 font-medium">{record.total_obtained} / {record.total_max}</td></tr>
          <tr><td className="py-1 text-gray-500">Percentage</td><td className="py-1 font-medium">{record.percentage}%</td></tr>
          <tr><td className="py-1 text-gray-500">Grade</td><td className="py-1 font-medium">{record.grade}</td></tr>
          <tr><td className="py-1 text-gray-500">Result</td><td className="py-1 font-medium">{record.result_status}</td></tr>
          <tr><td className="py-1 text-gray-500">Verification ID</td><td className="py-1 font-medium">{record.verification_id}</td></tr>
        </tbody>
      </table>
      <div className="flex justify-between items-end mt-16">
        <div className="text-xs text-gray-500">Date of Result: {settings.result_date}<br />{settings.footer_text}</div>
        <div className="text-center">
          <div className="border-t pt-1 text-xs" style={{ borderColor: "#999", width: 160 }}>{settings.signature_text}</div>
        </div>
      </div>
    </div>
  );
}

/* ----------------------------------------------------------------------
   VERIFY PAGE  (simulates a public /verify/{id} destination for QR codes)
---------------------------------------------------------------------- */

function VerifyPage({ settings, results, onBack }) {
  const [id, setId] = useState("");
  const [found, setFound] = useState(undefined); // undefined = not searched, null = not found

  function check(e) {
    e.preventDefault();
    const match = results.find((r) => r.published && norm(r.verification_id) === norm(id));
    setFound(match || null);
  }

  return (
    <div className="min-h-screen flex flex-col">
      <header className="border-b bg-white">
        <div className="max-w-3xl mx-auto px-5 py-4 flex items-center gap-3">
          <Logo settings={settings} size={34} />
          <div className="font-display font-semibold" style={{ color: NAVY }}>{settings.institution_name}</div>
        </div>
      </header>
      <main className="flex-1 max-w-md w-full mx-auto px-5 py-12">
        <button onClick={onBack} className="mb-6 text-sm text-gray-500 inline-flex items-center gap-1 hover:text-gray-800">
          <ArrowLeft size={15} /> Back to results portal
        </button>
        <h1 className="font-display text-2xl font-semibold mb-1" style={{ color: NAVY }}>Verify a Result</h1>
        <p className="text-gray-500 text-sm mb-6">Enter the Verification ID printed on the memo or found in its QR code.</p>
        <form onSubmit={check} className="bg-white rounded-2xl shadow-sm border p-6" style={{ borderColor: "#E7EAF0" }}>
          <Field label="Verification ID">
            <input value={id} onChange={(e) => setId(e.target.value)} className={inputCls} style={inputStyle} placeholder="e.g. VER8K2N4Q1X" />
          </Field>
          <PrimaryButton type="submit" full icon={ShieldCheck}>Verify</PrimaryButton>
        </form>

        {found === null && (
          <div className="mt-5 flex items-center gap-2 text-sm rounded-lg px-3 py-2.5" style={{ background: "#FBEAE8", color: BAD }}>
            <XCircle size={16} /> No published result matches this verification ID.
          </div>
        )}
        {found && (
          <div className="mt-5 bg-white rounded-2xl shadow-sm border p-6 text-sm" style={{ borderColor: "#E7EAF0" }}>
            <div className="flex items-center gap-2 mb-4 font-semibold" style={{ color: OK }}>
              <BadgeCheck size={18} /> This result is genuine
            </div>
            <ResultRow label="Student Name" value={found.student_name} />
            <ResultRow label="Roll Number" value={found.roll_number} />
            <ResultRow label="Course / Class" value={found.course} />
            <ResultRow label="Examination" value={found.examination || settings.examination_name} />
            <ResultRow label="Session" value={found.session || settings.session} />
            <ResultRow label="Result" value={found.result_status} />
          </div>
        )}
      </main>
    </div>
  );
}

/* ----------------------------------------------------------------------
   ADMIN LOGIN
---------------------------------------------------------------------- */

function AdminLogin({ settings, adminCreds, onBack, onLoggedIn }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [show, setShow] = useState(false);
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  async function submit(e) {
    e.preventDefault();
    setError("");
    if (!email || !password) {
      setError("Enter both email and password.");
      return;
    }
    setBusy(true);
    const hash = await sha256Hex(email.trim().toLowerCase() + ":" + password);
    setBusy(false);
    if (adminCreds && email.trim().toLowerCase() === adminCreds.email.toLowerCase() && hash === adminCreds.hash) {
      onLoggedIn(adminCreds.email);
    } else {
      setError("Incorrect email or password.");
    }
  }

  return (
    <div className="min-h-screen flex items-center justify-center px-5" style={{ background: `radial-gradient(circle at 20% 20%, ${NAVY} 0%, ${NAVY_DEEP} 60%)` }}>
      <div className="w-full max-w-sm">
        <button onClick={onBack} className="mb-6 text-sm text-white/70 inline-flex items-center gap-1 hover:text-white">
          <ArrowLeft size={15} /> Back to student portal
        </button>
        <div className="bg-white rounded-2xl shadow-xl p-7">
          <div className="flex items-center gap-3 mb-6">
            <Logo settings={settings} />
            <div>
              <div className="font-display font-semibold" style={{ color: NAVY }}>Admin Sign In</div>
              <div className="text-xs text-gray-500">{settings.institution_name}</div>
            </div>
          </div>
          <form onSubmit={submit}>
            <Field label="Email">
              <input value={email} onChange={(e) => setEmail(e.target.value)} type="email" className={inputCls} style={inputStyle} autoComplete="username" />
            </Field>
            <Field label="Password">
              <div className="relative">
                <input value={password} onChange={(e) => setPassword(e.target.value)} type={show ? "text" : "password"} className={inputCls} style={inputStyle} autoComplete="current-password" />
                <button type="button" onClick={() => setShow((s) => !s)} className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-400">
                  {show ? <EyeOff size={16} /> : <Eye size={16} />}
                </button>
              </div>
            </Field>
            {error && (
              <div className="mb-4 flex items-center gap-2 text-sm rounded-lg px-3 py-2.5" style={{ background: "#FBEAE8", color: BAD }}>
                <AlertTriangle size={15} /> {error}
              </div>
            )}
            <PrimaryButton type="submit" full disabled={busy} icon={busy ? Loader2 : KeyRound}>
              {busy ? "Signing in…" : "Sign In"}
            </PrimaryButton>
          </form>
        </div>
        <p className="text-center text-xs text-white/50 mt-5">
          Passwords are hashed before storage. This runs as a client-side prototype — pair it with a real backend before production use.
        </p>
      </div>
    </div>
  );
}

/* ----------------------------------------------------------------------
   ADMIN DASHBOARD
---------------------------------------------------------------------- */

const TABS = [
  { id: "overview", label: "Overview", icon: LayoutDashboard },
  { id: "upload", label: "Upload Results", icon: Upload },
  { id: "manage", label: "Manage Results", icon: ClipboardList },
  { id: "published", label: "Published Results", icon: CheckCircle2 },
  { id: "settings", label: "Settings", icon: SettingsIcon },
  { id: "profile", label: "Admin Profile", icon: KeyRound },
];

function AdminDashboard({ settings, results, adminCreds, session, notify, onLogout, onSaveSettings, onSaveResults, onSaveCreds, onViewStudentSite }) {
  const [tab, setTab] = useState("overview");
  const [navOpen, setNavOpen] = useState(false);

  const stats = useMemo(() => {
    const published = results.filter((r) => r.published).length;
    const lastUpload = results.reduce((max, r) => (r.created_at > max ? r.created_at : max), "");
    return {
      total: results.length,
      published,
      draft: results.length - published,
      lastUpload: lastUpload ? new Date(lastUpload).toLocaleString() : "—",
    };
  }, [results]);

  return (
    <div className="min-h-screen flex">
      {/* Sidebar */}
      <aside
        className={`fixed lg:static z-40 inset-y-0 left-0 w-64 shrink-0 text-white flex flex-col transition-transform ${navOpen ? "translate-x-0" : "-translate-x-full lg:translate-x-0"}`}
        style={{ background: `linear-gradient(180deg, ${NAVY}, ${NAVY_DEEP})` }}
      >
        <div className="flex items-center gap-3 px-5 py-5 border-b border-white/10">
          <Logo settings={settings} size={36} />
          <div>
            <div className="font-display font-semibold text-sm leading-tight">{settings.institution_name}</div>
            <div className="text-xs text-white/50">Admin Dashboard</div>
          </div>
        </div>
        <nav className="flex-1 px-3 py-4 space-y-1">
          {TABS.map((t) => (
            <button
              key={t.id}
              onClick={() => { setTab(t.id); setNavOpen(false); }}
              className="w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition"
              style={{ background: tab === t.id ? "rgba(255,255,255,0.12)" : "transparent", color: tab === t.id ? "#fff" : "rgba(255,255,255,0.65)" }}
            >
              <t.icon size={16} /> {t.label}
            </button>
          ))}
        </nav>
        <div className="px-3 py-4 border-t border-white/10 space-y-1">
          <button onClick={onViewStudentSite} className="w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium text-white/65 hover:text-white">
            <GraduationCap size={16} /> View Student Site
          </button>
          <button onClick={onLogout} className="w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium text-white/65 hover:text-white">
            <LogOut size={16} /> Logout
          </button>
        </div>
      </aside>
      {navOpen && <div className="fixed inset-0 bg-black/40 z-30 lg:hidden" onClick={() => setNavOpen(false)} />}

      {/* Main */}
      <div className="flex-1 min-w-0">
        <div className="lg:hidden flex items-center justify-between px-4 py-3 bg-white border-b" style={{ borderColor: "#E7EAF0" }}>
          <button onClick={() => setNavOpen(true)}><Menu size={20} /></button>
          <div className="font-display font-semibold text-sm" style={{ color: NAVY }}>{TABS.find((t) => t.id === tab)?.label}</div>
          <div style={{ width: 20 }} />
        </div>

        <div className="p-5 sm:p-8 max-w-6xl mx-auto">
          {tab === "overview" && <OverviewTab stats={stats} settings={settings} session={session} />}
          {tab === "upload" && <UploadTab settings={settings} results={results} onSaveResults={onSaveResults} notify={notify} setTab={setTab} />}
          {tab === "manage" && <ManageTab allResults={results} displayResults={results} onSaveResults={onSaveResults} notify={notify} settings={settings} />}
          {tab === "published" && <ManageTab allResults={results} displayResults={results.filter((r) => r.published)} onSaveResults={onSaveResults} notify={notify} settings={settings} readOnlyFilter />}
          {tab === "settings" && <SettingsTab settings={settings} onSaveSettings={onSaveSettings} notify={notify} />}
          {tab === "profile" && <ProfileTab adminCreds={adminCreds} onSaveCreds={onSaveCreds} notify={notify} />}
        </div>
      </div>
    </div>
  );
}

function StatCard({ label, value, icon: Icon, accent }) {
  return (
    <div className="bg-white rounded-2xl border p-5" style={{ borderColor: "#E7EAF0" }}>
      <div className="flex items-center justify-between mb-3">
        <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">{label}</div>
        <Icon size={16} color={accent || BLUE} />
      </div>
      <div className="font-display text-2xl font-semibold" style={{ color: NAVY }}>{value}</div>
    </div>
  );
}

function OverviewTab({ stats, settings, session }) {
  return (
    <div>
      <h2 className="font-display text-2xl font-semibold mb-1" style={{ color: NAVY }}>Overview</h2>
      <p className="text-sm text-gray-500 mb-6">Signed in as {session.email}</p>
      <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-4">
        <StatCard label="Total Results" value={stats.total} icon={Users} />
        <StatCard label="Published Results" value={stats.published} icon={CheckCircle2} accent={OK} />
        <StatCard label="Draft / Unpublished" value={stats.draft} icon={Eye} accent="#B07A17" />
        <StatCard label="Last Excel Upload" value={stats.lastUpload} icon={FileSpreadsheet} />
        <StatCard label="Current Examination" value={settings.examination_name} icon={GraduationCap} />
        <StatCard label="Current Session" value={settings.session} icon={ClipboardList} />
      </div>
    </div>
  );
}

/* ---------------- Upload tab ---------------- */

function UploadTab({ settings, results, onSaveResults, notify, setTab }) {
  const [rawRows, setRawRows] = useState(null); // array of raw objects from file
  const [headers, setHeaders] = useState([]);
  const [mapping, setMapping] = useState({}); // canonical field -> header name
  const [fileName, setFileName] = useState("");
  const [errors, setErrors] = useState([]);
  const [busy, setBusy] = useState(false);
  const inputRef = useRef(null);

  function reset() {
    setRawRows(null); setHeaders([]); setMapping({}); setFileName(""); setErrors([]);
    if (inputRef.current) inputRef.current.value = "";
  }

  async function handleFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setErrors([]);
    const okExt = /\.(xlsx|xls|csv)$/i.test(file.name);
    if (!okExt) { setErrors(["Unsupported file type. Please upload .xlsx, .xls, or .csv"]); return; }
    if (file.size > 8 * 1024 * 1024) { setErrors(["File is too large (max 8MB)."]); return; }

    setBusy(true);
    setFileName(file.name);
    try {
      let rows = [];
      if (/\.csv$/i.test(file.name)) {
        const text = await file.text();
        const parsed = Papa.parse(text, { header: true, skipEmptyLines: true });
        if (parsed.errors?.length) {
          const reasons = parsed.errors.slice(0, 5).map((e) => `Row ${e.row != null ? e.row + 2 : "?"}: ${e.message}`);
          setErrors(["This CSV has formatting problems and couldn't be fully read:", ...reasons]);
          setBusy(false);
          return;
        }
        rows = parsed.data;
      } else {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: "array" });
        if (!wb.SheetNames.length) {
          setErrors(["This workbook has no sheets to read."]);
          setBusy(false);
          return;
        }
        const sheet = wb.Sheets[wb.SheetNames[0]];
        rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      }
      rows = rows.filter((r) => Object.values(r).some((v) => String(v).trim() !== ""));
      if (rows.length === 0) {
        setErrors(["The file appears to be empty — no data rows were found below the header row."]);
        setBusy(false);
        return;
      }
      const hdrs = Object.keys(rows[0]).filter((h) => h && !h.startsWith("__EMPTY"));
      if (hdrs.length === 0) {
        setErrors(["Could not find any column headers in the first row of this file."]);
        setBusy(false);
        return;
      }
      const autoMap = autoDetectColumns(hdrs);
      setErrors([]);
      setHeaders(hdrs);
      setMapping(autoMap);
      setRawRows(rows);
    } catch (err) {
      setErrors(["Could not read this file: " + (err?.message || "the file appears to be corrupted or is not a valid Excel/CSV file") + "."]);
    }
    setBusy(false);
  }

  function autoDetectColumns(hdrs) {
    const map = {};
    for (const field of Object.keys(ALIASES)) {
      const hit = hdrs.find((h) => ALIASES[field].includes(norm(h)) || norm(h) === norm(field));
      if (hit) map[field] = hit;
    }
    return map;
  }

  const preview = useMemo(() => {
    if (!rawRows) return { rows: [], errs: [] };
    return buildPreview(rawRows, mapping, settings, results);
  }, [rawRows, mapping, settings, results]);

  const mappingIssues = useMemo(() => {
    const issues = [];
    if (!mapping.roll_number) issues.push('Missing Roll Number column — none of the headers matched common names like "Roll No / Roll Number / ROLLNO". Map it manually below.');
    if (!mapping.student_name) issues.push("Missing Student Name column — map it manually below.");
    return issues;
  }, [mapping]);

  async function confirmImport() {
    if (mappingIssues.length > 0) {
      notify("Can't import — " + mappingIssues[0], "error");
      return;
    }
    if (preview.rows.length === 0) {
      notify("Nothing valid to import: every row failed validation (see the issues listed above the preview).", "error");
      return;
    }
    const batchId = uid("BATCH-");
    const now = new Date().toISOString();
    const existingByRoll = new Map(results.map((r) => [norm(r.roll_number), r]));
    const toAdd = [];
    for (const row of preview.rows) {
      if (existingByRoll.has(norm(row.roll_number))) continue; // prevent duplicates
      toAdd.push({
        id: uid("RES-"),
        ...row,
        verification_id: uid("VER"),
        published: false,
        batch_id: batchId,
        created_at: now,
        updated_at: now,
      });
    }
    if (toAdd.length === 0) {
      notify("All rows were duplicates of existing roll numbers — nothing imported.", "error");
      return;
    }
    await onSaveResults([...results, ...toAdd]);
    notify(`Imported ${toAdd.length} record(s) as drafts. Publish them from Manage Results.`);
    reset();
    setTab("manage");
  }

  return (
    <div>
      <h2 className="font-display text-2xl font-semibold mb-1" style={{ color: NAVY }}>Upload Results</h2>
      <p className="text-sm text-gray-500 mb-6">Upload Excel → Validate → Preview → Confirm Import → Publish Results</p>

      {!rawRows && (
        <div className="bg-white rounded-2xl border-2 border-dashed p-10 text-center" style={{ borderColor: "#D8DEE9" }}>
          <FileSpreadsheet size={30} className="mx-auto mb-3" color={BLUE} />
          <div className="font-medium mb-1">Upload a result file</div>
          <div className="text-sm text-gray-500 mb-5">.xlsx, .xls, or .csv — headers detected automatically</div>
          <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} className="hidden" id="file-up" />
          <label htmlFor="file-up">
            <span className="inline-flex items-center gap-2 rounded-full px-5 py-2.5 text-sm font-semibold text-white cursor-pointer" style={{ background: `linear-gradient(155deg, ${BLUE}, ${NAVY})` }}>
              {busy ? <Loader2 size={16} className="animate-spin" /> : <Upload size={16} />} Choose File
            </span>
          </label>
        </div>
      )}

      {errors.length > 0 && (
        <div className="mt-4 rounded-lg px-4 py-3 text-sm" style={{ background: "#FBEAE8", color: BAD }}>
          {errors.map((e, i) => <div key={i} className="flex items-center gap-2"><AlertTriangle size={15} /> {e}</div>)}
        </div>
      )}

      {rawRows && (
        <div className="space-y-6">
          <div className="bg-white rounded-2xl border p-5" style={{ borderColor: "#E7EAF0" }}>
            <div className="flex items-center justify-between mb-4 flex-wrap gap-2">
              <div className="text-sm"><span className="font-semibold">{fileName}</span> · {rawRows.length} row(s) detected</div>
              <GhostButton icon={RefreshCw} onClick={reset}>Start Over</GhostButton>
            </div>
            <div className="text-xs font-semibold uppercase tracking-wide text-gray-500 mb-3">Column Mapping</div>
            <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-3">
              {Object.keys(ALIASES).map((field) => (
                <label key={field} className="text-xs">
                  <span className="block mb-1 text-gray-500 capitalize">
                    {field.replace(/_/g, " ")} {(field === "roll_number" || field === "student_name") && <span style={{ color: BAD }}>*</span>}
                  </span>
                  <select
                    value={mapping[field] || ""}
                    onChange={(e) => setMapping((m) => ({ ...m, [field]: e.target.value || undefined }))}
                    className="w-full rounded-lg border px-2.5 py-2 text-sm"
                    style={inputStyle}
                  >
                    <option value="">— not mapped —</option>
                    {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                  </select>
                </label>
              ))}
            </div>
            <div className="text-xs text-gray-400 mt-3">
              Unmapped columns (e.g. subject names like "Maths", "Science") are automatically treated as subject-wise marks. A column named "&lt;subject&gt; Max" sets that subject's maximum marks.
            </div>
          </div>

          {mappingIssues.length > 0 && (
            <div className="rounded-lg px-4 py-3 text-sm space-y-1" style={{ background: "#FBEAE8", color: BAD }}>
              {mappingIssues.map((e, i) => <div key={i} className="flex items-center gap-2"><AlertTriangle size={15} /> {e}</div>)}
            </div>
          )}

          {preview.errs.length > 0 && (
            <div className="rounded-lg px-4 py-3 text-sm space-y-1" style={{ background: "#FFF6E8", color: "#8A5A0D" }}>
              {preview.errs.slice(0, 8).map((e, i) => <div key={i} className="flex items-center gap-2"><AlertTriangle size={14} /> {e}</div>)}
              {preview.errs.length > 8 && <div>…and {preview.errs.length - 8} more issue(s)</div>}
            </div>
          )}

          <div className="bg-white rounded-2xl border p-5" style={{ borderColor: "#E7EAF0" }}>
            <div className="flex items-center justify-between mb-3">
              <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">Preview ({preview.rows.length} valid row(s))</div>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-sm min-w-[640px]">
                <thead>
                  <tr style={{ background: BLUE_SOFT }}>
                    <th className="text-left px-3 py-2 font-semibold">Roll No.</th>
                    <th className="text-left px-3 py-2 font-semibold">Name</th>
                    <th className="text-right px-3 py-2 font-semibold">Total</th>
                    <th className="text-right px-3 py-2 font-semibold">%</th>
                    <th className="text-right px-3 py-2 font-semibold">Result</th>
                  </tr>
                </thead>
                <tbody>
                  {preview.rows.slice(0, 10).map((r, i) => (
                    <tr key={i} className="border-t" style={{ borderColor: "#EEF1F5" }}>
                      <td className="px-3 py-2">{r.roll_number}</td>
                      <td className="px-3 py-2">{r.student_name}</td>
                      <td className="px-3 py-2 text-right">{r.total_obtained}/{r.total_max}</td>
                      <td className="px-3 py-2 text-right">{r.percentage}%</td>
                      <td className="px-3 py-2 text-right"><StatusPill status={r.result_status} /></td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {preview.rows.length > 10 && <div className="text-xs text-gray-400 mt-2">…and {preview.rows.length - 10} more row(s)</div>}
            </div>
          </div>

          <PrimaryButton
            icon={Upload}
            onClick={confirmImport}
            disabled={mappingIssues.length > 0 || preview.rows.length === 0}
          >
            Confirm Import ({preview.rows.length} rows)
          </PrimaryButton>
          {(mappingIssues.length > 0 || preview.rows.length === 0) && (
            <div className="text-xs mt-2" style={{ color: BAD }}>
              {mappingIssues.length > 0
                ? "Import is disabled until the required columns above are mapped."
                : "Import is disabled: no rows passed validation — see the issues listed above."}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function buildPreview(rawRows, mapping, settings, existingResults) {
  const errs = [];
  const seenRoll = new Set();
  const usedHeaders = new Set(Object.values(mapping).filter(Boolean));
  const rows = [];

  rawRows.forEach((raw, idx) => {
    const rowNum = idx + 2; // account for header row
    const get = (field) => (mapping[field] ? raw[mapping[field]] : undefined);
    const roll = String(get("roll_number") ?? "").trim();
    const name = String(get("student_name") ?? "").trim();

    if (!roll) { errs.push(`Row ${rowNum}: missing Roll Number — skipped.`); return; }
    if (!name) { errs.push(`Row ${rowNum}: missing Student Name — skipped.`); return; }
    if (seenRoll.has(norm(roll))) { errs.push(`Row ${rowNum}: duplicate Roll Number "${roll}" within this file — skipped.`); return; }
    seenRoll.add(norm(roll));

    // Build subjects from unmapped, non-empty columns
    const subjectMaxCols = {};
    Object.keys(raw).forEach((h) => {
      if (usedHeaders.has(h)) return;
      const n = norm(h);
      const maxMatch = n.match(/^(.*)max$/) || n.match(/^max(.*)$/);
      if (maxMatch) {
        const base = (maxMatch[1] || "").trim();
        if (base) subjectMaxCols[base] = h;
      }
    });
    const subjects = [];
    Object.keys(raw).forEach((h) => {
      if (usedHeaders.has(h)) return;
      const n = norm(h);
      if (Object.values(subjectMaxCols).includes(h)) return;
      if (n.match(/max$/) && subjectMaxCols[n.replace(/max$/, "")]) return;
      const val = raw[h];
      if (val === "" || val === undefined || val === null) return;
      const num = Number(val);
      if (Number.isNaN(num)) {
        errs.push(`Row ${rowNum}: "${h}" has a non-numeric value ("${val}") — not counted as marks.`);
        return;
      }
      const maxHeader = subjectMaxCols[n];
      const max = maxHeader ? Number(raw[maxHeader]) || settings.default_max_marks : settings.default_max_marks;
      subjects.push({ name: h, obtained: num, max });
    });

    const rawTotalObtained = get("total_obtained");
    const rawTotalMax = get("total_max");
    const rawPercentage = get("percentage");
    let totalObtained = Number(rawTotalObtained);
    let totalMax = Number(rawTotalMax);
    if (mapping.total_obtained && rawTotalObtained !== "" && rawTotalObtained != null && Number.isNaN(totalObtained)) {
      errs.push(`Row ${rowNum}: invalid Total Marks value "${rawTotalObtained}" — recalculated from subject marks instead.`);
    }
    if (mapping.total_max && rawTotalMax !== "" && rawTotalMax != null && Number.isNaN(totalMax)) {
      errs.push(`Row ${rowNum}: invalid Maximum Marks value "${rawTotalMax}" — recalculated from subject marks instead.`);
    }
    if (!totalObtained || Number.isNaN(totalObtained)) {
      totalObtained = subjects.reduce((s, x) => s + (Number(x.obtained) || 0), 0);
    }
    if (!totalMax || Number.isNaN(totalMax)) {
      totalMax = subjects.length ? subjects.reduce((s, x) => s + (Number(x.max) || 0), 0) : settings.default_max_marks;
    }
    let percentage = Number(rawPercentage);
    if (mapping.percentage && rawPercentage !== "" && rawPercentage != null && Number.isNaN(percentage)) {
      errs.push(`Row ${rowNum}: invalid Percentage value "${rawPercentage}" — recalculated from totals instead.`);
    }
    if (!percentage || Number.isNaN(percentage)) {
      percentage = totalMax ? Math.round((totalObtained / totalMax) * 10000) / 100 : 0;
    }
    let status = String(get("result_status") ?? "").toUpperCase().trim();
    if (status !== "PASS" && status !== "FAIL") {
      status = percentage >= (settings.passing_percentage || 35) ? "PASS" : "FAIL";
    }
    let grade = String(get("grade") ?? "").trim();
    if (!grade) grade = gradeFromPercentage(percentage);

    if (existingResults.some((r) => norm(r.roll_number) === norm(roll))) {
      errs.push(`Row ${rowNum}: Roll Number "${roll}" already exists in the database — will be skipped on import.`);
    }

    rows.push({
      roll_number: roll,
      student_name: name,
      registration_number: String(get("registration_number") ?? "").trim(),
      course: String(get("course") ?? "").trim(),
      semester: String(get("semester") ?? "").trim(),
      examination: String(get("examination") ?? "").trim(),
      session: String(get("session") ?? "").trim(),
      subjects,
      total_obtained: totalObtained,
      total_max: totalMax,
      percentage,
      grade,
      result_status: status,
      division: String(get("division") ?? "").trim(),
    });
  });

  return { rows, errs };
}

/* ---------------- Manage / Published tab ---------------- */

function ManageTab({ allResults, displayResults, onSaveResults, notify, settings, readOnlyFilter }) {
  const [q, setQ] = useState("");
  const [editing, setEditing] = useState(null);
  const [confirmDelete, setConfirmDelete] = useState(null);

  const filtered = useMemo(() => {
    const query = norm(q);
    if (!query) return displayResults;
    return displayResults.filter((r) => norm(r.roll_number).includes(query) || norm(r.student_name).includes(query));
  }, [displayResults, q]);

  return (
    <div>
      <h2 className="font-display text-2xl font-semibold mb-1" style={{ color: NAVY }}>{readOnlyFilter ? "Published Results" : "Manage Results"}</h2>
      <p className="text-sm text-gray-500 mb-5">{readOnlyFilter ? "Results currently visible to students." : "Search, edit, publish, or remove imported records."}</p>

      <div className="flex items-center gap-3 mb-4 flex-wrap">
        <div className="relative flex-1 min-w-[220px]">
          <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
          <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Search by roll number or name" className={inputCls + " pl-9"} style={inputStyle} />
        </div>
        <ExportButtons rows={filtered} />
      </div>

      <div className="bg-white rounded-2xl border overflow-hidden" style={{ borderColor: "#E7EAF0" }}>
        <div className="overflow-x-auto">
          <table className="w-full text-sm min-w-[760px]">
            <thead>
              <tr style={{ background: BLUE_SOFT }}>
                <th className="text-left px-4 py-2.5 font-semibold">Roll No.</th>
                <th className="text-left px-4 py-2.5 font-semibold">Name</th>
                <th className="text-left px-4 py-2.5 font-semibold">Course</th>
                <th className="text-right px-4 py-2.5 font-semibold">%</th>
                <th className="text-right px-4 py-2.5 font-semibold">Status</th>
                <th className="text-right px-4 py-2.5 font-semibold">Published</th>
                <th className="text-right px-4 py-2.5 font-semibold">Actions</th>
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 && (
                <tr><td colSpan={7} className="text-center py-10 text-gray-400 text-sm">No matching results.</td></tr>
              )}
              {filtered.map((r) => (
                <tr key={r.id} className="border-t" style={{ borderColor: "#EEF1F5" }}>
                  <td className="px-4 py-2.5">{r.roll_number}</td>
                  <td className="px-4 py-2.5">{r.student_name}</td>
                  <td className="px-4 py-2.5 text-gray-500">{r.course || "—"}</td>
                  <td className="px-4 py-2.5 text-right">{r.percentage}%</td>
                  <td className="px-4 py-2.5 text-right"><StatusPill status={r.result_status} /></td>
                  <td className="px-4 py-2.5 text-right">
                    <ResultsToggle result={r} onSaveResults={onSaveResults} allResults={allResults} />
                  </td>
                  <td className="px-4 py-2.5">
                    <div className="flex justify-end gap-2">
                      <button title="Edit" onClick={() => setEditing(r)} className="p-1.5 rounded-lg hover:bg-gray-100"><Pencil size={15} /></button>
                      <button title="Delete" onClick={() => setConfirmDelete(r)} className="p-1.5 rounded-lg hover:bg-gray-100"><Trash2 size={15} color={BAD} /></button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {editing && (
        <EditModal
          record={editing}
          onClose={() => setEditing(null)}
          onSave={async (patch) => {
            const next = allResults.map((r) => (r.id === editing.id ? { ...r, ...patch, updated_at: new Date().toISOString() } : r));
            await onSaveResults(next);
            notify("Result updated.");
            setEditing(null);
          }}
        />
      )}

      {confirmDelete && (
        <ConfirmDialog
          title="Delete this result?"
          body={`This will permanently remove the record for ${confirmDelete.student_name} (Roll ${confirmDelete.roll_number}).`}
          confirmLabel="Delete"
          danger
          onCancel={() => setConfirmDelete(null)}
          onConfirm={async () => {
            const next = allResults.filter((r) => r.id !== confirmDelete.id);
            await onSaveResults(next);
            notify("Result deleted.");
            setConfirmDelete(null);
          }}
        />
      )}
    </div>
  );
}

function ResultsToggle({ result, onSaveResults, allResults }) {
  const [busy, setBusy] = useState(false);
  async function toggle() {
    setBusy(true);
    const next = allResults.map((r) => (r.id === result.id ? { ...r, published: !r.published, updated_at: new Date().toISOString() } : r));
    await onSaveResults(next);
    setBusy(false);
  }
  return (
    <button
      onClick={toggle}
      disabled={busy}
      className="inline-flex items-center gap-1.5 rounded-full px-3 py-1 text-xs font-semibold"
      style={{ background: result.published ? "#E6F5EC" : "#F1F3F6", color: result.published ? OK : "#5B6474" }}
    >
      {result.published ? <CheckCircle2 size={13} /> : <Eye size={13} />}
      {result.published ? "Published" : "Draft"}
    </button>
  );
}

function EditModal({ record, onClose, onSave }) {
  const [form, setForm] = useState({ ...record });
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));

  function recompute() {
    const totalMax = Number(form.total_max) || 0;
    const totalObtained = Number(form.total_obtained) || 0;
    const percentage = totalMax ? Math.round((totalObtained / totalMax) * 10000) / 100 : 0;
    setForm((f) => ({ ...f, percentage, grade: gradeFromPercentage(percentage), result_status: percentage >= 35 ? "PASS" : "FAIL" }));
  }

  return (
    <div className="fixed inset-0 bg-black/40 z-50 flex items-center justify-center p-4">
      <div className="bg-white rounded-2xl shadow-xl w-full max-w-lg max-h-[85vh] overflow-y-auto p-6">
        <div className="flex items-center justify-between mb-4">
          <div className="font-display text-lg font-semibold" style={{ color: NAVY }}>Edit Result</div>
          <button onClick={onClose}><X size={18} /></button>
        </div>
        <div className="grid grid-cols-2 gap-x-4">
          <Field label="Roll Number"><input className={inputCls} style={inputStyle} value={form.roll_number} onChange={set("roll_number")} /></Field>
          <Field label="Student Name"><input className={inputCls} style={inputStyle} value={form.student_name} onChange={set("student_name")} /></Field>
          <Field label="Course"><input className={inputCls} style={inputStyle} value={form.course} onChange={set("course")} /></Field>
          <Field label="Semester"><input className={inputCls} style={inputStyle} value={form.semester} onChange={set("semester")} /></Field>
          <Field label="Total Obtained"><input type="number" className={inputCls} style={inputStyle} value={form.total_obtained} onChange={set("total_obtained")} /></Field>
          <Field label="Total Max"><input type="number" className={inputCls} style={inputStyle} value={form.total_max} onChange={set("total_max")} /></Field>
          <Field label="Percentage"><input type="number" className={inputCls} style={inputStyle} value={form.percentage} onChange={set("percentage")} /></Field>
          <Field label="Grade"><input className={inputCls} style={inputStyle} value={form.grade} onChange={set("grade")} /></Field>
          <Field label="Result Status">
            <select className={inputCls} style={inputStyle} value={form.result_status} onChange={set("result_status")}>
              <option value="PASS">PASS</option>
              <option value="FAIL">FAIL</option>
            </select>
          </Field>
          <Field label="Division"><input className={inputCls} style={inputStyle} value={form.division} onChange={set("division")} /></Field>
        </div>
        <button onClick={recompute} className="text-xs font-semibold text-blue-600 mb-4">Recalculate percentage / grade / status from totals</button>
        <div className="flex gap-3">
          <PrimaryButton onClick={() => onSave(form)}>Save Changes</PrimaryButton>
          <GhostButton onClick={onClose}>Cancel</GhostButton>
        </div>
      </div>
    </div>
  );
}

function ConfirmDialog({ title, body, confirmLabel, danger, onCancel, onConfirm }) {
  return (
    <div className="fixed inset-0 bg-black/40 z-50 flex items-center justify-center p-4">
      <div className="bg-white rounded-2xl shadow-xl w-full max-w-sm p-6">
        <div className="flex items-center gap-2 mb-2 font-display text-lg font-semibold" style={{ color: danger ? BAD : NAVY }}>
          <AlertTriangle size={18} /> {title}
        </div>
        <p className="text-sm text-gray-500 mb-5">{body}</p>
        <div className="flex gap-3">
          <button onClick={onConfirm} className="rounded-full px-5 py-2.5 text-sm font-semibold text-white" style={{ background: danger ? BAD : BLUE }}>{confirmLabel}</button>
          <GhostButton onClick={onCancel}>Cancel</GhostButton>
        </div>
      </div>
    </div>
  );
}

function ExportButtons({ rows }) {
  function exportCsv() {
    const flat = rows.map((r) => ({
      roll_number: r.roll_number, student_name: r.student_name, registration_number: r.registration_number,
      course: r.course, semester: r.semester, examination: r.examination, session: r.session,
      total_obtained: r.total_obtained, total_max: r.total_max, percentage: r.percentage, grade: r.grade,
      result_status: r.result_status, division: r.division, published: r.published, verification_id: r.verification_id,
    }));
    const csv = Papa.unparse(flat);
    downloadBlob(csv, "results-export.csv", "text/csv");
  }
  function exportJson() {
    downloadBlob(JSON.stringify(rows, null, 2), "results-export.json", "application/json");
  }
  return (
    <div className="flex gap-2">
      <GhostButton icon={Download} onClick={exportCsv}>CSV</GhostButton>
      <GhostButton icon={Download} onClick={exportJson}>JSON</GhostButton>
    </div>
  );
}

function downloadBlob(content, filename, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}

/* ---------------- Settings tab ---------------- */

function SettingsTab({ settings, onSaveSettings, notify }) {
  const [form, setForm] = useState({ ...settings });
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));

  async function save(e) {
    e.preventDefault();
    await onSaveSettings(form);
    notify("Settings saved.");
  }

  return (
    <div>
      <h2 className="font-display text-2xl font-semibold mb-1" style={{ color: NAVY }}>Settings</h2>
      <p className="text-sm text-gray-500 mb-6">These appear automatically on the student portal and printed memos.</p>
      <form onSubmit={save} className="bg-white rounded-2xl border p-6 max-w-2xl" style={{ borderColor: "#E7EAF0" }}>
        <div className="grid sm:grid-cols-2 gap-x-5">
          <Field label="Institution Name"><input className={inputCls} style={inputStyle} value={form.institution_name} onChange={set("institution_name")} /></Field>
          <Field label="Logo Initials" hint="Shown inside the logo mark"><input className={inputCls} style={inputStyle} value={form.logo_text} onChange={set("logo_text")} maxLength={3} /></Field>
          <Field label="Address"><input className={inputCls} style={inputStyle} value={form.address} onChange={set("address")} /></Field>
          <Field label="Contact Details"><input className={inputCls} style={inputStyle} value={form.contact} onChange={set("contact")} /></Field>
          <Field label="Examination Name"><input className={inputCls} style={inputStyle} value={form.examination_name} onChange={set("examination_name")} /></Field>
          <Field label="Academic Session"><input className={inputCls} style={inputStyle} value={form.session} onChange={set("session")} /></Field>
          <Field label="Result Publication Date"><input type="date" className={inputCls} style={inputStyle} value={form.result_date} onChange={set("result_date")} /></Field>
          <Field label="Memo Title"><input className={inputCls} style={inputStyle} value={form.memo_title} onChange={set("memo_title")} /></Field>
          <Field label="Signature Line"><input className={inputCls} style={inputStyle} value={form.signature_text} onChange={set("signature_text")} /></Field>
          <Field label="Passing Percentage"><input type="number" className={inputCls} style={inputStyle} value={form.passing_percentage} onChange={set("passing_percentage")} /></Field>
          <Field label="Default Max Marks per Subject"><input type="number" className={inputCls} style={inputStyle} value={form.default_max_marks} onChange={set("default_max_marks")} /></Field>
        </div>
        <Field label="Footer Text"><textarea className={inputCls} style={inputStyle} rows={2} value={form.footer_text} onChange={set("footer_text")} /></Field>
        <PrimaryButton type="submit">Save Settings</PrimaryButton>
      </form>
    </div>
  );
}

/* ---------------- Profile tab ---------------- */

function ProfileTab({ adminCreds, onSaveCreds, notify }) {
  const [current, setCurrent] = useState("");
  const [next, setNext] = useState("");
  const [confirm, setConfirm] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  async function changePassword(e) {
    e.preventDefault();
    setError("");
    if (next.length < 8) { setError("New password must be at least 8 characters."); return; }
    if (next !== confirm) { setError("New password and confirmation do not match."); return; }
    setBusy(true);
    const curHash = await sha256Hex(adminCreds.email.toLowerCase() + ":" + current);
    if (curHash !== adminCreds.hash) {
      setBusy(false);
      setError("Current password is incorrect.");
      return;
    }
    const newHash = await sha256Hex(adminCreds.email.toLowerCase() + ":" + next);
    await onSaveCreds({ ...adminCreds, hash: newHash });
    setBusy(false);
    setCurrent(""); setNext(""); setConfirm("");
    notify("Password updated.");
  }

  return (
    <div>
      <h2 className="font-display text-2xl font-semibold mb-1" style={{ color: NAVY }}>Admin Profile</h2>
      <p className="text-sm text-gray-500 mb-6">Signed in as {adminCreds.email}</p>
      <form onSubmit={changePassword} className="bg-white rounded-2xl border p-6 max-w-sm" style={{ borderColor: "#E7EAF0" }}>
        <div className="text-sm font-semibold mb-4">Change Password</div>
        <Field label="Current Password"><input type="password" className={inputCls} style={inputStyle} value={current} onChange={(e) => setCurrent(e.target.value)} /></Field>
        <Field label="New Password"><input type="password" className={inputCls} style={inputStyle} value={next} onChange={(e) => setNext(e.target.value)} /></Field>
        <Field label="Confirm New Password"><input type="password" className={inputCls} style={inputStyle} value={confirm} onChange={(e) => setConfirm(e.target.value)} /></Field>
        {error && (
          <div className="mb-4 flex items-center gap-2 text-sm rounded-lg px-3 py-2.5" style={{ background: "#FBEAE8", color: BAD }}>
            <AlertTriangle size={15} /> {error}
          </div>
        )}
        <PrimaryButton type="submit" disabled={busy} icon={busy ? Loader2 : ShieldCheck}>{busy ? "Updating…" : "Update Password"}</PrimaryButton>
      </form>
    </div>
  );
}
