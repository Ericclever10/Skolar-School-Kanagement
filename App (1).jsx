import React, { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer } from "recharts";
import {
  Users, GraduationCap, ClipboardList, Upload, MessageSquare,
  Award, BookOpen, Send, Plus, Trash2, Edit2, X, Search,
  Download, LayoutDashboard, ChevronRight, Printer,
  Save, CheckCircle2, AlertCircle, FileSpreadsheet, Lock,
  Mail, UserCircle, Sparkles, Star, LogOut, TrendingUp, Check,
  School, Copy, RefreshCw, KeyRound, ShieldCheck, UserPlus, Phone,
  NotebookPen, FileText, Library, CheckCheck, Image as ImageIcon,
  CalendarCheck, Wallet, MessageCircle,
  RotateCcw, ShieldAlert, Grid3x3, ArrowRightLeft
} from "lucide-react";

// ---------- helpers ----------
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
const ACCOUNTS_KEY = "skolar-accounts";
const SCHOOLS_KEY = "skolar-schools";
const SESSION_KEY = "session";
const schoolDataKey = (schoolId) => `skolar-data:${schoolId}`;

const emptyState = { learners: [], teachers: [], exams: [], marks: [], messages: [], schemes: [], markingSchemes: [], lessonNotes: [], examSubjectStatus: [], resources: [], reportComments: [], attendance: [], feeStructures: [], feePayments: [], gradingSystems: [], subjectGradingSystems: {}, defaultGradingSystemId: "standard", trash: [], disciplineRecords: [], timetableSlots: [] };

const CODE_CHARS = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"; // no 0/O/1/I to avoid confusion

// Browser storage adapter: replaces the hosted-artifact storage API so Skolar
// works as a normal website on GitHub Pages, Netlify, Vercel, etc.
if (!window.storage) {
  window.storage = {
    async get(key) {
      const value = localStorage.getItem(key);
      return value === null ? null : { value };
    },
    async set(key, value) {
      localStorage.setItem(key, value);
      return { value };
    },
    async delete(key) {
      localStorage.removeItem(key);
      return true;
    },
  };
}

// ---------- Subscription ----------
const TRIAL_DAYS = 30;
const MONTHLY_PRICE = 150;
const YEARLY_PRICE = 1500;
const PAY_PHONE = "0795864556";
const PAYMENT_CODES_KEY = "skolar-used-payment-codes";

function daysSince(isoDate) {
  if (!isoDate) return 0;
  return Math.floor((Date.now() - new Date(isoDate).getTime()) / (1000 * 60 * 60 * 24));
}
function trialDaysLeft(school) {
  if (!school) return 0;
  return Math.max(0, TRIAL_DAYS - daysSince(school.trialStartDate));
}
function subscriptionActive(school) {
  if (!school) return false;
  if (!school.plan) return true; // pre-existing schools created before subscriptions existed
  if (school.plan === "trial") return trialDaysLeft(school) > 0;
  if (!school.subscriptionExpiresAt) return false;
  return new Date(school.subscriptionExpiresAt).getTime() > Date.now();
}
function generateSchoolCode(existingSchools) {
  let code;
  do {
    code = Array.from({ length: 6 }, () => CODE_CHARS[Math.floor(Math.random() * CODE_CHARS.length)]).join("");
  } while (existingSchools.some((s) => s.code === code));
  return code;
}

// Adds a lightweight teacher profile to a school's data set directly in storage.
// Used when someone joins a school via code — independent of whichever school's data
// (if any) happens to be loaded in the current browser session.
async function addTeacherProfileToSchool(schoolId, account) {
  try {
    const res = await window.storage.get(schoolDataKey(schoolId), false);
    const current = res && res.value ? JSON.parse(res.value) : {};
    const base = { learners: [], teachers: [], exams: [], marks: [], messages: [], ...current };
    const teacherProfile = { id: uid(), name: account.name, subject: "", email: account.email, phone: "", classesAssigned: "", accountId: account.id };
    base.teachers = [...base.teachers, teacherProfile];
    await window.storage.set(schoolDataKey(schoolId), JSON.stringify(base), false);
  } catch (e) {
    console.error("Couldn't add teacher profile to school", e);
  }
}

const ALL_SUBJECTS = [
  "Mathematics", "English Language", "Kiswahili", "Kenyan Sign Language (KSL)", "Science",
  "Integrated Science", "Social Studies", "History", "Geography",
  "Physics", "Chemistry", "Biology", "Computer Science", "Literature", "French", "Spanish",
  "Art", "Creative Arts", "Music", "Physical Education", "Religious Education", "CRE", "IRE", "HRE",
  "Economics", "Business Studies", "Agriculture", "Pre-Technical and Pre-Career Studies",
  "Woodwork", "Metalwork", "Home Science", "Life Skills Education", "Community Service Learning",
  "Civic Education", "Home Economics",
];

const BUILTIN_GRADING_SYSTEMS = [
  {
    id: "standard", name: "Standard (A–E)", builtin: true,
    bands: [
      { label: "A", min: 80 }, { label: "A-", min: 75 }, { label: "B+", min: 70 },
      { label: "B", min: 65 }, { label: "B-", min: 60 }, { label: "C+", min: 55 },
      { label: "C", min: 50 }, { label: "C-", min: 45 }, { label: "D+", min: 40 },
      { label: "D", min: 35 }, { label: "D-", min: 30 }, { label: "E", min: 0 },
    ],
  },
  {
    id: "cbc", name: "CBC (Exceeding / Meeting / Below Expectation)", builtin: true,
    bands: [
      { label: "EE1", min: 90 }, { label: "EE2", min: 75 },
      { label: "ME1", min: 58 }, { label: "ME2", min: 41 },
      { label: "BE1", min: 21 }, { label: "BE2", min: 0 },
    ],
  },
];

function getGradingSystems(data) {
  return [...BUILTIN_GRADING_SYSTEMS, ...((data && data.gradingSystems) || [])];
}
function getSystemForSubject(data, subject) {
  const systems = getGradingSystems(data);
  const assignedId = (data.subjectGradingSystems || {})[subject] || data.defaultGradingSystemId || "standard";
  return systems.find((s) => s.id === assignedId) || systems[0];
}
function getDefaultSystem(data) {
  const systems = getGradingSystems(data);
  return systems.find((s) => s.id === (data.defaultGradingSystemId || "standard")) || systems[0];
}
function gradeForPercent(system, percent) {
  const sorted = [...system.bands].sort((a, b) => b.min - a.min);
  const band = sorted.find((b) => percent >= b.min);
  return band ? band.label : sorted[sorted.length - 1].label;
}
// getGrade(percent) is kept as a fallback for the school-wide default system (e.g. overall averages).
// Per-subject grades should use gradeForPercent(getSystemForSubject(data, subject), percent) instead.
function getGrade(percent) {
  return gradeForPercent(BUILTIN_GRADING_SYSTEMS[0], percent);
}
function toPercent(score, outOf) {
  const max = Number(outOf) || 100;
  if (max <= 0) return 0;
  const pct = (Number(score || 0) / max) * 100;
  return Math.min(100, Math.max(0, pct));
}
function clamp(val, min, max) {
  if (val === "" || val === null || val === undefined) return val;
  const n = Number(val);
  if (isNaN(n)) return val;
  return Math.min(max, Math.max(min, n));
}

// Soft-delete: move a record into trash instead of destroying it, so it can be restored later.
function trashItem(data, persist, type, item, extraRemoval) {
  const entry = { id: uid(), type, item, deletedAt: new Date().toISOString() };
  persist({ ...data, trash: [...(data.trash || []), entry], ...(extraRemoval || {}) });
}

function serifStyle() {
  return { fontFamily: 'Georgia, "Iowan Old Style", "Times New Roman", serif' };
}

// ---------- small UI atoms ----------
function Seal({ letter = "R", size = 40, logo = null }) {
  if (logo) {
    return (
      <img
        src={logo}
        alt="School logo"
        style={{ width: size, height: size }}
        className="rounded-full object-cover border-2 border-amber-400 shrink-0"
      />
    );
  }
  return (
    <div
      style={{ width: size, height: size, ...serifStyle() }}
      className="rounded-full bg-emerald-800 text-amber-200 flex items-center justify-center font-bold border-2 border-amber-400 shrink-0"
    >
      {letter}
    </div>
  );
}

function Btn({ children, onClick, variant = "primary", size = "md", type = "button", className = "", disabled }) {
  const base = "inline-flex items-center gap-1.5 rounded-md font-medium transition-colors disabled:opacity-40 disabled:cursor-not-allowed";
  const sizes = { sm: "px-2.5 py-1.5 text-xs", md: "px-3.5 py-2 text-sm" };
  const variants = {
    primary: "bg-emerald-800 text-white hover:bg-emerald-900",
    secondary: "bg-slate-100 text-slate-800 hover:bg-slate-200 border border-slate-300",
    danger: "bg-red-50 text-red-700 hover:bg-red-100 border border-red-200",
    ghost: "text-slate-600 hover:bg-slate-100",
    gold: "bg-amber-500 text-slate-900 hover:bg-amber-400",
  };
  return (
    <button type={type} disabled={disabled} onClick={onClick} className={`${base} ${sizes[size]} ${variants[variant]} ${className}`}>
      {children}
    </button>
  );
}

function Field({ label, children }) {
  return (
    <label className="block mb-3">
      <span className="block text-xs font-semibold text-slate-500 uppercase tracking-wide mb-1">{label}</span>
      {children}
    </label>
  );
}

const inputCls = "w-full rounded-md border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-emerald-700 focus:border-emerald-700";

function Modal({ title, onClose, children, wide }) {
  return (
    <div className="fixed inset-0 bg-slate-900/50 z-50 flex items-end sm:items-center justify-center p-0 sm:p-4">
      <div className={`bg-white rounded-t-2xl sm:rounded-xl w-full ${wide ? "sm:max-w-2xl" : "sm:max-w-md"} max-h-[90vh] overflow-y-auto shadow-2xl`}>
        <div className="flex items-center justify-between px-5 py-4 border-b border-slate-200 sticky top-0 bg-white z-10">
          <h3 className="font-bold text-slate-800" style={serifStyle()}>{title}</h3>
          <button onClick={onClose} className="p-1 rounded-full hover:bg-slate-100"><X size={18} /></button>
        </div>
        <div className="p-5">{children}</div>
      </div>
    </div>
  );
}

function EmptyState({ icon: Icon, title, hint }) {
  return (
    <div className="text-center py-14 px-4 text-slate-400">
      <Icon size={32} className="mx-auto mb-3 opacity-50" />
      <p className="font-semibold text-slate-500">{title}</p>
      {hint && <p className="text-sm mt-1 max-w-xs mx-auto">{hint}</p>}
    </div>
  );
}

function Toast({ msg, onDone }) {
  useEffect(() => {
    const t = setTimeout(onDone, 2400);
    return () => clearTimeout(t);
  }, [msg]);
  if (!msg) return null;
  return (
    <div className="fixed bottom-4 left-1/2 -translate-x-1/2 bg-slate-900 text-white text-sm px-4 py-2.5 rounded-full shadow-lg z-[60] flex items-center gap-2">
      <CheckCircle2 size={15} className="text-emerald-400" /> {msg}
    </div>
  );
}

// Upload and attach a real document (PDF, Word, or Excel) to a scheme, exam, or marking scheme.
// The file is stored as-is (base64) so it can be downloaded again later — nothing is auto-generated.
const MAX_ATTACHMENT_BYTES = 3.5 * 1024 * 1024; // ~3.5MB raw file (storage has a 5MB-per-key ceiling)

function FileAttachment({ label, attachment, onUpload, onRemove }) {
  const inputRef = useRef(null);
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");

  const handleFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setErr("");
    if (file.size > MAX_ATTACHMENT_BYTES) {
      setErr(`That file is too large (max ~${Math.round(MAX_ATTACHMENT_BYTES / 1024 / 1024)}MB).`);
      if (inputRef.current) inputRef.current.value = "";
      return;
    }
    setBusy(true);
    const reader = new FileReader();
    reader.onload = () => {
      onUpload({ name: file.name, type: file.type, size: file.size, dataUrl: reader.result });
      setBusy(false);
      if (inputRef.current) inputRef.current.value = "";
    };
    reader.onerror = () => { setBusy(false); setErr("Couldn't read that file — please try again."); };
    reader.readAsDataURL(file);
  };

  return (
    <div className="border border-slate-200 rounded-lg p-3">
      {label && <p className="text-xs font-semibold text-slate-500 uppercase tracking-wide mb-2">{label}</p>}
      {attachment ? (
        <div className="flex items-center justify-between gap-2 bg-stone-50 rounded-md px-3 py-2">
          <div className="flex items-center gap-2 min-w-0">
            <FileText size={15} className="text-emerald-700 shrink-0" />
            <span className="text-sm text-slate-700 truncate">{attachment.name}</span>
          </div>
          <div className="flex items-center gap-2 shrink-0">
            <a href={attachment.dataUrl} download={attachment.name} className="text-xs text-emerald-800 font-medium">Download</a>
            <button type="button" onClick={onRemove} className="text-slate-400 hover:text-red-600"><X size={14} /></button>
          </div>
        </div>
      ) : (
        <div>
          <input ref={inputRef} type="file" accept=".pdf,.doc,.docx,.xls,.xlsx,.csv" onChange={handleFile} className="hidden" disabled={busy} />
          <button
            type="button"
            onClick={() => inputRef.current && inputRef.current.click()}
            disabled={busy}
            className="w-full flex items-center justify-center gap-2 border-2 border-dashed border-slate-300 rounded-lg py-3 text-sm text-slate-600 hover:border-emerald-600 hover:text-emerald-700 disabled:opacity-50"
          >
            <Upload size={15} /> {busy ? "Uploading…" : "Tap to choose a file"}
          </button>
          {err && <p className="text-xs text-red-600 mt-1 flex items-center gap-1"><AlertCircle size={12} /> {err}</p>}
        </div>
      )}
    </div>
  );
}

// ---------- Auth ----------
const REMEMBER_KEY = "remembered-credentials";

// ---------- Subscription paywall ----------
function SubscriptionPanel({ onActivate }) {
  const [plan, setPlan] = useState("monthly");
  const [confirming, setConfirming] = useState(false);
  const [code, setCode] = useState("");
  const [error, setError] = useState("");
  const amount = plan === "monthly" ? MONTHLY_PRICE : YEARLY_PRICE;

  const payViaMpesa = () => {
    window.location.href = `tel:*334*1*1*${PAY_PHONE}*${amount}%23`;
  };

  const submitCode = () => {
    setError("");
    const result = onActivate(plan, code);
    if (result && !result.ok) setError(result.error);
    else setCode("");
  };

  return (
    <>
      <div className="flex mb-4 bg-slate-100 rounded-lg p-1">
        {["monthly", "yearly"].map((p) => (
          <button key={p} onClick={() => setPlan(p)} className={`flex-1 py-1.5 rounded-md text-sm font-semibold transition-colors ${plan === p ? "bg-white text-emerald-800 shadow-sm" : "text-slate-500"}`}>
            {p === "monthly" ? `KES ${MONTHLY_PRICE}/mo` : `KES ${YEARLY_PRICE}/yr`}
          </button>
        ))}
      </div>

      {!confirming ? (
        <>
          <div className="bg-stone-50 rounded-lg p-4 mb-4 text-sm text-slate-700 space-y-1.5">
            <p>Pay via <b>M-Pesa</b> to:</p>
            <p className="text-lg font-bold text-slate-900" style={serifStyle()}>{PAY_PHONE}</p>
            <p className="text-slate-500">Amount: <b>KES {amount}</b> ({plan === "monthly" ? "1 month" : "1 year"})</p>
          </div>
          <Btn className="w-full justify-center mb-2" onClick={payViaMpesa}><Phone size={14} /> Pay via M-Pesa</Btn>
          <p className="text-[11px] text-slate-400 text-center mb-3">Opens your phone's dialer with the M-Pesa Send Money code pre-filled — check the number and amount before you confirm the send.</p>
          <Btn variant="secondary" className="w-full justify-center" onClick={() => setConfirming(true)}>I've already paid — enter my code</Btn>
        </>
      ) : (
        <>
          <p className="text-sm text-slate-600 mb-3">
            Enter the <b>M-Pesa confirmation code</b> from the payment message (e.g. QGH7X8Y2ZP) for the <b>KES {amount}</b> payment to <b>{PAY_PHONE}</b>.
          </p>
          <Field label="M-Pesa confirmation code">
            <input className={`${inputCls} uppercase tracking-widest`} value={code} onChange={(e) => setCode(e.target.value)} placeholder="e.g. QGH7X8Y2ZP" />
          </Field>
          {error && <p className="text-xs text-red-600 mb-3 flex items-center gap-1"><AlertCircle size={13} /> {error}</p>}
          <p className="text-[11px] text-slate-400 mb-4">This app can't contact Safaricom to verify the code — it only checks that this exact code hasn't already been used to activate a subscription before. Only submit a code from a payment you actually made.</p>
          <div className="flex gap-2">
            <Btn variant="secondary" onClick={() => { setConfirming(false); setError(""); }} className="flex-1 justify-center">Back</Btn>
            <Btn onClick={submitCode} disabled={!code.trim()} className="flex-1 justify-center"><CheckCircle2 size={14} /> Verify & activate</Btn>
          </div>
        </>
      )}
    </>
  );
}

function SubscriptionGate({ school, isOwner, onActivate, onLogout }) {
  return (
    <div className="min-h-screen bg-slate-900 flex items-center justify-center p-4">
      <div className="w-full max-w-sm">
        <div className="flex flex-col items-center mb-5 text-center">
          <Seal letter="S" size={48} />
          <h1 className="text-white font-bold text-xl mt-3" style={serifStyle()}>Your free trial has ended</h1>
          <p className="text-slate-400 text-sm mt-1">{school.name}</p>
        </div>

        <div className="bg-white rounded-2xl shadow-2xl p-6">
          {!isOwner ? (
            <p className="text-sm text-slate-600 text-center">
              This school's 30-day free trial has ended. Ask your school's Owner/Admin to renew the subscription to continue using Skolar.
            </p>
          ) : (
            <SubscriptionPanel onActivate={onActivate} />
          )}
        </div>

        <button onClick={onLogout} className="w-full text-center text-slate-400 text-sm mt-4 hover:text-white">Log out</button>
      </div>
    </div>
  );
}

function AuthGate({ onLogin, accounts, saveAccounts, schools, saveSchools }) {
  const [mode, setMode] = useState("login"); // login | signup
  const [path, setPath] = useState("register"); // register | join  (signup only)
  const [form, setForm] = useState({ name: "Eric Mutua", email: "ericmutua064@gmail.com", phone: "0795864556", username: "", password: "", schoolName: "", code: "", role: "Teacher" });
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const [remember, setRemember] = useState(true);
  const [autoLoginTried, setAutoLoginTried] = useState(false);
  const [checkingSaved, setCheckingSaved] = useState(true);

  // Pre-fill (and try) remembered credentials from a previous login on this device.
  // A short safety timeout guarantees the form is always usable even if storage is slow or unavailable.
  useEffect(() => {
    let done = false;
    const finish = () => { if (!done) { done = true; setCheckingSaved(false); } };
    const safety = setTimeout(finish, 2500);
    (async () => {
      try {
        const res = await window.storage.get(REMEMBER_KEY, false);
        if (res && res.value) {
          const saved = JSON.parse(res.value);
          setForm((f) => ({ ...f, username: saved.username, password: saved.password }));
        }
      } catch (e) {
        // nothing remembered yet
      }
      clearTimeout(safety);
      finish();
    })();
  }, []);

  // Once accounts are loaded and we have remembered credentials pre-filled, log in automatically.
  const attemptingAutoLogin = !checkingSaved && !autoLoginTried && !!accounts && !!form.username && !!form.password;
  useEffect(() => {
    if (checkingSaved || autoLoginTried || !accounts || !form.username || !form.password) return;
    setAutoLoginTried(true);
    const account = accounts.find((a) => a.username === form.username && a.password === form.password);
    if (account) onLogin(account);
  }, [checkingSaved, accounts, form.username, form.password, autoLoginTried]);

  const set = (k) => (e) => setForm({ ...form, [k]: e.target.value });

  const rememberCredentials = (uname, pwd) => {
    if (remember) {
      window.storage.set(REMEMBER_KEY, JSON.stringify({ username: uname, password: pwd }), false).catch(() => {});
    } else {
      window.storage.delete(REMEMBER_KEY, false).catch(() => {});
    }
  };

  const submit = async () => {
    setError("");
    if (!accounts || !schools) return;
    setBusy(true);
    const uname = form.username.trim().toLowerCase();

    if (mode === "login") {
      const account = accounts.find((a) => a.username === uname && a.password === form.password);
      if (!account) { setError("Username or password is incorrect."); setBusy(false); return; }
      rememberCredentials(uname, form.password);
      onLogin(account);
      setBusy(false);
      return;
    }

    // --- signup ---
    if (!form.name.trim() || !uname || !form.password) { setError("Please fill in your name, username and password."); setBusy(false); return; }
    if (accounts.some((a) => a.username === uname)) { setError("That username is already taken."); setBusy(false); return; }

    if (path === "register") {
      if (!form.schoolName.trim()) { setError("Give your school a name."); setBusy(false); return; }
      const school = { id: uid(), name: form.schoolName.trim(), code: generateSchoolCode(schools), ownerId: "", createdAt: new Date().toISOString(), trialStartDate: new Date().toISOString(), plan: "trial", subscriptionExpiresAt: null };
      const account = { id: uid(), name: form.name.trim(), email: form.email.trim(), phone: form.phone.trim(), username: uname, password: form.password, role: "Super Admin", schoolId: school.id };
      school.ownerId = account.id;
      // These update local state immediately and try to persist in the background — signup
      // succeeds either way, so a storage hiccup never blocks you from using the app.
      await saveSchools([...schools, school]);
      await saveAccounts([...accounts, account]);
      rememberCredentials(uname, form.password);
      onLogin(account);
    } else {
      const code = form.code.trim().toUpperCase();
      const school = schools.find((s) => s.code === code);
      if (!school) { setError("That school code doesn't match any school. Double-check it with your admin."); setBusy(false); return; }
      const account = { id: uid(), name: form.name.trim(), email: form.email.trim(), phone: form.phone.trim(), username: uname, password: form.password, role: form.role, schoolId: school.id };
      await saveAccounts([...accounts, account]);
      addTeacherProfileToSchool(school.id, account);
      rememberCredentials(uname, form.password);
      onLogin(account);
    }
    setBusy(false);
  };

  return (
    <div className="min-h-screen bg-slate-900 flex items-center justify-center p-4">
      <div className="w-full max-w-sm">
        <div className="flex flex-col items-center mb-6 text-center">
          <Seal letter="S" size={52} />
          <h1 className="text-white font-bold text-2xl mt-3" style={serifStyle()}>Skolar</h1>
          <p className="text-slate-400 text-sm">School records &amp; reporting</p>
        </div>

        <div className="bg-white rounded-2xl shadow-2xl p-6">
          {attemptingAutoLogin && (
            <p className="text-xs text-emerald-700 bg-emerald-50 rounded-md px-2.5 py-1.5 mb-3">Signing you in with your saved login…</p>
          )}
          <div className="flex mb-5 bg-slate-100 rounded-lg p-1">
            {["login", "signup"].map((m) => (
              <button key={m} onClick={() => { setMode(m); setError(""); }} className={`flex-1 py-1.5 rounded-md text-sm font-semibold transition-colors ${mode === m ? "bg-white text-emerald-800 shadow-sm" : "text-slate-500"}`}>
                {m === "login" ? "Log in" : "Create account"}
              </button>
            ))}
          </div>

          {mode === "signup" && (
            <div className="grid grid-cols-2 gap-2 mb-4">
              <button
                onClick={() => { setPath("register"); setError(""); }}
                className={`text-left rounded-lg border p-2.5 text-xs transition-colors ${path === "register" ? "border-emerald-700 bg-emerald-50" : "border-slate-200"}`}
              >
                <ShieldCheck size={15} className={path === "register" ? "text-emerald-800 mb-1" : "text-slate-400 mb-1"} />
                <p className="font-semibold text-slate-800">Register my school</p>
                <p className="text-slate-500">Sets you up as Super Admin, full access</p>
              </button>
              <button
                onClick={() => { setPath("join"); setError(""); }}
                className={`text-left rounded-lg border p-2.5 text-xs transition-colors ${path === "join" ? "border-emerald-700 bg-emerald-50" : "border-slate-200"}`}
              >
                <UserPlus size={15} className={path === "join" ? "text-emerald-800 mb-1" : "text-slate-400 mb-1"} />
                <p className="font-semibold text-slate-800">Join with a code</p>
                <p className="text-slate-500">My school's admin gave me a code (works from your own account for now)</p>
              </button>
            </div>
          )}

          <div>
            {mode === "signup" && (
              <>
                <Field label="Full name">
                  <div className="relative"><UserCircle size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" /><input className={`${inputCls} pl-9`} value={form.name} onChange={set("name")} autoComplete="name" placeholder="e.g. Mrs. Adaeze Obi" /></div>
                </Field>
                <Field label="Email">
                  <div className="relative"><Mail size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" /><input type="email" className={`${inputCls} pl-9`} value={form.email} onChange={set("email")} autoComplete="email" placeholder="you@school.edu" /></div>
                </Field>
                <Field label="Phone">
                  <div className="relative"><Phone size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" /><input type="tel" className={`${inputCls} pl-9`} value={form.phone} onChange={set("phone")} autoComplete="tel" placeholder="e.g. 0700 000 000" /></div>
                </Field>
                {path === "register" ? (
                  <Field label="School name">
                    <div className="relative"><School size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" /><input className={`${inputCls} pl-9`} value={form.schoolName} onChange={set("schoolName")} placeholder="e.g. Greenfield Academy" /></div>
                  </Field>
                ) : (
                  <>
                    <Field label="School code">
                      <div className="relative"><KeyRound size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" /><input className={`${inputCls} pl-9 uppercase tracking-widest`} value={form.code} onChange={set("code")} placeholder="e.g. 7K2QXR" maxLength={6} /></div>
                    </Field>
                    <Field label="Your role at this school">
                      <div className="grid grid-cols-3 gap-1.5">
                        {["Teacher", "Admin", "Registrar"].map((r) => (
                          <button
                            key={r}
                            type="button"
                            onClick={() => setForm({ ...form, role: r })}
                            className={`py-1.5 rounded-md text-xs font-semibold border transition-colors ${form.role === r ? "bg-emerald-800 text-white border-emerald-800" : "bg-white text-slate-600 border-slate-300"}`}
                          >
                            {r}
                          </button>
                        ))}
                      </div>
                    </Field>
                  </>
                )}
              </>
            )}
            <Field label="Username">
              <input className={inputCls} value={form.username} onChange={set("username")} autoCapitalize="none" autoComplete="username" placeholder="e.g. aobi" />
            </Field>
            <Field label="Password">
              <div className="relative">
                <Lock size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                <input
                  type="password"
                  className={`${inputCls} pl-9`}
                  value={form.password}
                  onChange={set("password")}
                  placeholder="••••••••"
                  autoComplete={mode === "login" ? "current-password" : "new-password"}
                  onKeyDown={(e) => { if (e.key === "Enter") submit(); }}
                />
              </div>
            </Field>

            <label className="flex items-center gap-2 text-xs text-slate-600 mb-4 cursor-pointer select-none">
              <input type="checkbox" checked={remember} onChange={(e) => setRemember(e.target.checked)} className="accent-emerald-700" />
              Remember my login on this device
            </label>

            {error && <p className="text-xs text-red-600 mb-3 flex items-center gap-1"><AlertCircle size={13} /> {error}</p>}

            <Btn onClick={submit} className="w-full justify-center" disabled={busy || !accounts}>
              {busy ? "Please wait…" : mode === "login" ? "Log in" : path === "register" ? "Create school & account" : "Join school"}
            </Btn>
          </div>
        </div>
        <p className="text-slate-500 text-[11px] text-center mt-4 leading-relaxed">
          This login is a lightweight access gate for your school team, not bank-grade security — avoid reusing a sensitive password here. With "Remember my login" checked, your username and password are saved on this device so the app can sign you straight into your dashboard next time. Log out to clear it.
        </p>
      </div>
    </div>
  );
}

// ---------- app ----------
const TABS = [
  { id: "dashboard", label: "Dashboard", icon: LayoutDashboard },
  { id: "profile", label: "School Profile", icon: ImageIcon },
  { id: "learners", label: "Learners", icon: Users },
  { id: "teachers", label: "Teachers", icon: GraduationCap },
  { id: "attendance", label: "Attendance", icon: CalendarCheck },
  { id: "fees", label: "Fees & Finance", icon: Wallet },
  { id: "discipline", label: "Discipline", icon: ShieldAlert },
  { id: "timetable", label: "Timetable", icon: Grid3x3 },
  { id: "academic", label: "Academic Resources", icon: Library },
  { id: "exams", label: "Exams & Marks", icon: ClipboardList },
  { id: "schemes", label: "Schemes of Work", icon: NotebookPen },
  { id: "lessonnotes", label: "Lesson Notes", icon: FileText },
  { id: "classlist", label: "Class Lists", icon: BookOpen },
  { id: "merit", label: "Merit List", icon: Award },
  { id: "reports", label: "Report Cards", icon: FileSpreadsheet },
  { id: "messages", label: "Message Parents", icon: MessageSquare },
];

export default function App() {
  const [user, setUser] = useState(null);
  const [sessionChecked, setSessionChecked] = useState(false);
  const [accounts, setAccounts] = useState(null);
  const [schools, setSchools] = useState(null);
  const [usedCodes, setUsedCodes] = useState(null);
  const [data, setData] = useState(emptyState);
  const [loaded, setLoaded] = useState(false);
  const [tab, setTab] = useState("dashboard");
  const [toast, setToast] = useState("");

  const safeGet = async (key, shared) => {
    try {
      const timeout = new Promise((resolve) => setTimeout(() => resolve(null), 4000));
      const res = await Promise.race([window.storage.get(key, shared), timeout]);
      return res && res.value ? res.value : null;
    } catch (e) {
      return null;
    }
  };

  // Fetch session, accounts, schools, and used payment codes all at once (in parallel) so the login screen appears fast.
  useEffect(() => {
    (async () => {
      const [sessionRaw, accountsRaw, schoolsRaw, codesRaw] = await Promise.all([
        safeGet(SESSION_KEY, false),
        safeGet(ACCOUNTS_KEY, false),
        safeGet(SCHOOLS_KEY, false),
        safeGet(PAYMENT_CODES_KEY, false),
      ]);

      const accountList = accountsRaw ? JSON.parse(accountsRaw) : [];
      const schoolList = schoolsRaw ? JSON.parse(schoolsRaw) : [];
      setAccounts(accountList);
      setSchools(schoolList);
      setUsedCodes(codesRaw ? JSON.parse(codesRaw) : []);

      if (sessionRaw) {
        const { id } = JSON.parse(sessionRaw);
        const account = accountList.find((a) => a.id === id);
        if (account) setUser(account);
      }
      setSessionChecked(true);
    })();
  }, []);

  // Once we know which school the logged-in user belongs to, load that school's data.
  useEffect(() => {
    if (!user) return;
    setLoaded(false);
    (async () => {
      const raw = await safeGet(schoolDataKey(user.schoolId), false);
      setData(raw ? { ...emptyState, ...JSON.parse(raw) } : emptyState);
      setLoaded(true);
    })();
  }, [user]);

  const handleLogin = (account) => {
    setUser(account);
    window.storage.set(SESSION_KEY, JSON.stringify({ id: account.id }), false).catch((e) => console.error("Couldn't save session", e));
  };

  const saveAccounts = async (next) => {
    setAccounts(next);
    try {
      await window.storage.set(ACCOUNTS_KEY, JSON.stringify(next), false);
    } catch (e) {
      // Persistence failed, but the account still exists in memory for this session,
      // so signup/login keep working even if it won't be remembered next time.
      console.error("Couldn't persist accounts", e);
    }
  };

  const saveSchools = async (next) => {
    setSchools(next);
    try {
      await window.storage.set(SCHOOLS_KEY, JSON.stringify(next), false);
    } catch (e) {
      console.error("Couldn't persist schools", e);
    }
  };

  const handleLogout = async () => {
    setUser(null);
    setLoaded(false);
    try {
      await window.storage.delete(SESSION_KEY, false);
      await window.storage.delete(REMEMBER_KEY, false);
    } catch (e) {
      // ignore
    }
  };

  const persist = async (next) => {
    setData(next);
    try {
      await window.storage.set(schoolDataKey(user.schoolId), JSON.stringify(next), false);
    } catch (e) {
      console.error("Storage save failed", e);
    }
  };

  const saveUsedCodes = async (next) => {
    setUsedCodes(next);
    try {
      await window.storage.set(PAYMENT_CODES_KEY, JSON.stringify(next), false);
    } catch (e) {
      console.error("Couldn't persist used payment codes", e);
    }
  };

  const activateSubscription = (plan, code) => {
    if (!schools || !usedCodes) return { ok: false, error: "Still loading — try again in a moment." };
    const normalized = code.trim().toUpperCase();
    if (!normalized) return { ok: false, error: "Enter the M-Pesa confirmation code from your payment message." };
    if (usedCodes.includes(normalized)) return { ok: false, error: "This code has already been used to activate a subscription. Each payment code can only be used once." };

    const currentSchoolRecord = schools.find((s) => s.id === user.schoolId);
    const base = (currentSchoolRecord && currentSchoolRecord.subscriptionExpiresAt && new Date(currentSchoolRecord.subscriptionExpiresAt).getTime() > Date.now())
      ? new Date(currentSchoolRecord.subscriptionExpiresAt)
      : new Date();
    const expires = new Date(base);
    if (plan === "monthly") expires.setDate(expires.getDate() + 30);
    else expires.setFullYear(expires.getFullYear() + 1);
    const next = schools.map((s) => (s.id === user.schoolId ? { ...s, plan, subscriptionExpiresAt: expires.toISOString() } : s));
    saveSchools(next);
    saveUsedCodes([...usedCodes, normalized]);
    notify(`Subscription activated (${plan === "monthly" ? "Monthly" : "Yearly"}) — thank you!`);
    return { ok: true };
  };

  const notify = (m) => setToast(m);

  if (!sessionChecked) {
    return <div className="min-h-screen flex items-center justify-center bg-slate-900 text-slate-400 text-sm">Loading…</div>;
  }

  if (!user) return <AuthGate onLogin={handleLogin} accounts={accounts} saveAccounts={saveAccounts} schools={schools} saveSchools={saveSchools} />;

  if (!loaded) {
    return <div className="min-h-screen flex items-center justify-center bg-stone-50 text-slate-400 text-sm">Loading your school's records…</div>;
  }

  const currentSchool = (schools || []).find((s) => s.id === user.schoolId);
  const isOwnerUser = currentSchool && user.id === currentSchool.ownerId;

  if (currentSchool && !subscriptionActive(currentSchool)) {
    return (
      <SubscriptionGate
        school={currentSchool}
        isOwner={isOwnerUser}
        onActivate={activateSubscription}
        onLogout={handleLogout}
      />
    );
  }

  return (
    <div className="min-h-screen bg-stone-50 pb-6">
      <header className="bg-slate-900 text-white sticky top-0 z-40 shadow-md">
        <div className="max-w-5xl mx-auto px-4 py-3 flex items-center gap-3">
          <Seal letter="S" logo={currentSchool?.logoDataUrl} />
          <div className="flex-1 min-w-0">
            <h1 className="font-bold text-lg leading-tight flex items-center gap-1.5" style={serifStyle()}>
              Skolar
              {currentSchool && user.id === currentSchool.ownerId && <Star size={14} className="text-amber-400" fill="#fbbf24" />}
            </h1>
            <p className="text-[11px] text-slate-400 -mt-0.5 truncate">{currentSchool ? currentSchool.name : "…"} · {user.name} · {user.role}</p>
          </div>
          <button onClick={handleLogout} title="Log out" className="p-2 rounded-full hover:bg-white/10 text-slate-300">
            <LogOut size={16} />
          </button>
        </div>
        <nav className="max-w-5xl mx-auto px-2 flex overflow-x-auto no-scrollbar border-t border-slate-700/60">
          {TABS.map((t) => (
            <button
              key={t.id}
              onClick={() => setTab(t.id)}
              className={`flex items-center gap-1.5 px-3.5 py-2.5 text-xs font-medium whitespace-nowrap border-b-2 transition-colors ${
                tab === t.id ? "border-amber-400 text-amber-300" : "border-transparent text-slate-300 hover:text-white"
              }`}
            >
              <t.icon size={14} /> {t.label}
            </button>
          ))}
        </nav>
      </header>

      <main className="max-w-5xl mx-auto px-4 pt-5">
        {tab === "dashboard" && <Dashboard data={data} setTab={setTab} user={user} school={currentSchool} onLogout={handleLogout} persist={persist} notify={notify} onActivateSubscription={activateSubscription} />}
        {tab === "profile" && <SchoolProfileTab school={currentSchool} schools={schools} saveSchools={saveSchools} user={user} notify={notify} data={data} persist={persist} />}
        {tab === "learners" && <Learners data={data} persist={persist} notify={notify} />}
        {tab === "teachers" && <Teachers data={data} persist={persist} notify={notify} user={user} accounts={accounts} saveAccounts={saveAccounts} />}
        {tab === "attendance" && <Attendance data={data} persist={persist} notify={notify} />}
        {tab === "fees" && <Fees data={data} persist={persist} notify={notify} />}
        {tab === "discipline" && <Discipline data={data} persist={persist} notify={notify} user={user} />}
        {tab === "timetable" && <Timetable data={data} persist={persist} notify={notify} />}
        {tab === "academic" && <AcademicResources data={data} persist={persist} notify={notify} setTab={setTab} />}
        {tab === "exams" && <ExamsMarks data={data} persist={persist} notify={notify} user={user} />}
        {tab === "schemes" && <Schemes data={data} persist={persist} notify={notify} />}
        {tab === "lessonnotes" && <LessonNotes data={data} persist={persist} notify={notify} />}
        {tab === "classlist" && <ClassLists data={data} school={currentSchool} />}
        {tab === "merit" && <MeritList data={data} school={currentSchool} />}
        {tab === "reports" && <ReportCards data={data} school={currentSchool} persist={persist} notify={notify} user={user} />}
        {tab === "messages" && <Messages data={data} persist={persist} notify={notify} />}
      </main>

      <Toast msg={toast} onDone={() => setToast("")} />
    </div>
  );
}

// ---------- Dashboard ----------
function classAverages(data) {
  const map = {};
  data.learners.forEach((l) => {
    const marks = data.marks.filter((m) => m.learnerId === l.id);
    if (!marks.length || !l.class) return;
    const avg = marks.reduce((s, m) => s + Number(m.score || 0), 0) / marks.length;
    (map[l.class] = map[l.class] || []).push(avg);
  });
  return Object.entries(map).map(([cls, arr]) => ({ class: cls, average: Number((arr.reduce((a, b) => a + b, 0) / arr.length).toFixed(1)) }));
}

function examTrend(data) {
  return [...data.exams]
    .filter((e) => (e.date || "").length > 0)
    .sort((a, b) => new Date(a.date) - new Date(b.date))
    .map((e) => {
      const { rows } = computeExamRanking(data, e.id);
      if (!rows.length) return null;
      const avg = rows.reduce((s, r) => s + r.average, 0) / rows.length;
      return { name: e.name.length > 14 ? e.name.slice(0, 14) + "…" : e.name, date: e.date, average: Number(avg.toFixed(1)) };
    })
    .filter(Boolean);
}

function learnerTrend(data, learnerId) {
  return [...data.exams]
    .filter((e) => (e.date || "").length > 0)
    .sort((a, b) => new Date(a.date) - new Date(b.date))
    .map((e) => {
      const { rows } = computeExamRanking(data, e.id);
      const row = rows.find((r) => r.learner.id === learnerId);
      if (!row) return null;
      return { name: e.name.length > 12 ? e.name.slice(0, 12) + "…" : e.name, average: Number(row.average.toFixed(1)), rank: row.rank };
    })
    .filter(Boolean);
}

function topPerformers(data, limit = 5) {
  return data.learners
    .map((l) => {
      const marks = data.marks.filter((m) => m.learnerId === l.id);
      if (!marks.length) return null;
      const avg = marks.reduce((s, m) => s + Number(m.score || 0), 0) / marks.length;
      return { learner: l, avg };
    })
    .filter(Boolean)
    .sort((a, b) => b.avg - a.avg)
    .slice(0, limit);
}

// ---------- School Profile (logo) ----------
function SchoolProfileTab({ school, schools, saveSchools, user, notify, data, persist }) {
  const inputRef = useRef(null);
  const restoreRef = useRef(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [restoreError, setRestoreError] = useState("");
  const [pendingRestore, setPendingRestore] = useState(null);
  const canEdit = user.role === "Super Admin" || user.role === "Admin";
  const MAX_LOGO_BYTES = 1.5 * 1024 * 1024;

  if (!school) return <EmptyState icon={ImageIcon} title="No school found" />;

  const handleFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setError("");
    if (!file.type.startsWith("image/")) { setError("Please choose an image file (PNG, JPG, etc)."); return; }
    if (file.size > MAX_LOGO_BYTES) { setError(`That image is too large (max ~${Math.round(MAX_LOGO_BYTES / 1024 / 1024 * 10) / 10}MB). Try a smaller or more compressed image.`); return; }
    setBusy(true);
    const reader = new FileReader();
    reader.onload = () => {
      const next = schools.map((s) => (s.id === school.id ? { ...s, logoDataUrl: reader.result } : s));
      saveSchools(next);
      notify("School logo updated");
      setBusy(false);
      if (inputRef.current) inputRef.current.value = "";
    };
    reader.onerror = () => { setBusy(false); setError("Couldn't read that image — please try again."); };
    reader.readAsDataURL(file);
  };

  const removeLogo = () => {
    const next = schools.map((s) => (s.id === school.id ? { ...s, logoDataUrl: null } : s));
    saveSchools(next);
    notify("Logo removed");
  };

  const exportBackup = () => {
    const payload = { skolarBackup: true, exportedAt: new Date().toISOString(), school: school.name, data };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${school.name.replace(/[^a-z0-9]+/gi, "-")}-skolar-backup-${new Date().toISOString().slice(0, 10)}.json`;
    a.click();
    URL.revokeObjectURL(url);
    notify("Backup downloaded");
  };

  const handleRestoreFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setRestoreError("");
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = JSON.parse(reader.result);
        const payload = parsed.skolarBackup ? parsed.data : parsed; // accept raw data too
        if (!payload || typeof payload !== "object" || !("learners" in payload)) throw new Error("This doesn't look like a Skolar backup file.");
        setPendingRestore(payload);
      } catch (err) {
        setRestoreError(err.message || "Couldn't read that file.");
      }
      if (restoreRef.current) restoreRef.current.value = "";
    };
    reader.readAsText(file);
  };

  const confirmRestore = () => {
    persist({ ...emptyState, ...pendingRestore });
    notify("Backup restored");
    setPendingRestore(null);
  };

  return (
    <div>
      <div className="bg-white rounded-xl border border-slate-200 p-5 mb-4">
        <h2 className="font-bold text-slate-800 mb-1" style={serifStyle()}>{school.name}</h2>
        <p className="text-xs text-slate-500 mb-4">This logo appears in the app header and on printed Class Lists, Merit Lists, and Report Cards.</p>

        <div className="flex flex-col items-center gap-4 py-4">
          <Seal letter="S" size={96} logo={school.logoDataUrl} />
          {canEdit ? (
            <>
              <input ref={inputRef} type="file" accept="image/*" onChange={handleFile} className="hidden" disabled={busy} />
              <div className="flex gap-2">
                <button
                  type="button"
                  onClick={() => inputRef.current && inputRef.current.click()}
                  disabled={busy}
                  className="flex items-center gap-2 border-2 border-dashed border-slate-300 rounded-lg px-4 py-2.5 text-sm text-slate-600 hover:border-emerald-600 hover:text-emerald-700 disabled:opacity-50"
                >
                  <Upload size={15} /> {busy ? "Uploading…" : school.logoDataUrl ? "Replace logo" : "Upload logo"}
                </button>
                {school.logoDataUrl && (
                  <Btn variant="danger" size="sm" onClick={removeLogo}><Trash2 size={13} /> Remove</Btn>
                )}
              </div>
              {error && <p className="text-xs text-red-600 flex items-center gap-1"><AlertCircle size={13} /> {error}</p>}
            </>
          ) : (
            <p className="text-xs text-slate-400">Only an Owner or Admin can change the school logo.</p>
          )}
        </div>
      </div>

      {canEdit && (
        <div className="bg-white rounded-xl border border-slate-200 p-5">
          <h3 className="font-semibold text-slate-800 mb-1 flex items-center gap-1.5" style={serifStyle()}><Download size={16} className="text-emerald-700" /> Backup & Restore</h3>
          <p className="text-xs text-slate-500 mb-4">All of Skolar's data lives on this device/account — download a backup regularly so you never lose it.</p>

          <Btn variant="secondary" onClick={exportBackup} className="w-full justify-center mb-3"><Download size={14} /> Download backup (.json)</Btn>

          <input ref={restoreRef} type="file" accept=".json" onChange={handleRestoreFile} className="hidden" />
          <button
            type="button"
            onClick={() => restoreRef.current && restoreRef.current.click()}
            className="w-full flex items-center justify-center gap-2 border-2 border-dashed border-slate-300 rounded-lg py-2.5 text-sm text-slate-600 hover:border-emerald-600 hover:text-emerald-700"
          >
            <Upload size={15} /> Restore from a backup file
          </button>
          {restoreError && <p className="text-xs text-red-600 mt-2 flex items-center gap-1"><AlertCircle size={13} /> {restoreError}</p>}

          {pendingRestore && (
            <div className="mt-3 bg-red-50 border border-red-200 rounded-lg p-3 text-xs text-red-800">
              <p className="mb-2"><b>This will replace all current data</b> — learners, exams, marks, fees, attendance, everything — with what's in the backup file. This can't be undone.</p>
              <div className="flex gap-2">
                <Btn size="sm" variant="danger" onClick={confirmRestore}>Yes, restore and replace everything</Btn>
                <Btn size="sm" variant="secondary" onClick={() => setPendingRestore(null)}>Cancel</Btn>
              </div>
            </div>
          )}
        </div>
      )}

      {canEdit && <GradingSystemsSection data={data} persist={persist} notify={notify} />}
      {canEdit && <TrashSection data={data} persist={persist} notify={notify} />}
    </div>

  );
}

function blankGradingSystem() { return { id: "", name: "", bands: [{ label: "", min: "" }] }; }

function GradingSystemsSection({ data, persist, notify }) {
  const systems = getGradingSystems(data);
  const custom = data.gradingSystems || [];
  const overrides = data.subjectGradingSystems || {};
  const [sysModal, setSysModal] = useState(null);

  const setDefault = (id) => {
    persist({ ...data, defaultGradingSystemId: id });
    notify("Default grading system updated");
  };

  const setSubjectOverride = (subject, id) => {
    const next = { ...overrides };
    if (id === "__default__") delete next[subject];
    else next[subject] = id;
    persist({ ...data, subjectGradingSystems: next });
  };

  const saveSystem = (sys) => {
    const exists = custom.some((s) => s.id === sys.id);
    const cleanBands = sys.bands.filter((b) => b.label.trim() !== "" && b.min !== "").map((b) => ({ label: b.label.trim(), min: Number(b.min) }));
    if (!cleanBands.length) { notify("Add at least one band with a label and a minimum %"); return; }
    const { forSubject, ...rest } = sys;
    const toSave = { ...rest, bands: cleanBands, id: exists ? sys.id : uid() };
    const next = exists ? custom.map((s) => (s.id === sys.id ? toSave : s)) : [...custom, toSave];
    const nextOverrides = forSubject ? { ...overrides, [forSubject]: toSave.id } : overrides;
    persist({ ...data, gradingSystems: next, subjectGradingSystems: nextOverrides });
    notify(forSubject ? `Grading for ${forSubject} updated` : exists ? "Grading system updated" : "Grading system created");
    setSysModal(null);
  };

  // Quick per-subject edit: fork the subject's currently-applied bands into a new subject-specific
  // system, so tweaking one subject's grading never accidentally changes another subject's.
  const editSubjectGrading = (subject) => {
    const current = getSystemForSubject(data, subject);
    setSysModal({ id: "", name: `${subject} Grading`, forSubject: subject, bands: current.bands.map((b) => ({ label: b.label, min: String(b.min) })) });
  };

  const deleteSystem = (id) => {
    const next = custom.filter((s) => s.id !== id);
    const clearedOverrides = { ...overrides };
    Object.keys(clearedOverrides).forEach((k) => { if (clearedOverrides[k] === id) delete clearedOverrides[k]; });
    persist({
      ...data,
      gradingSystems: next,
      subjectGradingSystems: clearedOverrides,
      defaultGradingSystemId: data.defaultGradingSystemId === id ? "standard" : data.defaultGradingSystemId,
    });
    notify("Grading system deleted");
  };

  return (
    <div className="bg-white rounded-xl border border-slate-200 p-5 mt-4">
      <h3 className="font-semibold text-slate-800 mb-1 flex items-center gap-1.5" style={serifStyle()}><Award size={16} className="text-emerald-700" /> Grading Systems</h3>
      <p className="text-xs text-slate-500 mb-4">Choose a school-wide default, and optionally assign a different grading system per subject — e.g. CBC's Exceeding/Meeting/Below Expectation scale for some subjects, Standard A–E for others.</p>

      <Field label="School default (used for overall averages, and any subject without an override)">
        <select className={inputCls} value={data.defaultGradingSystemId || "standard"} onChange={(e) => setDefault(e.target.value)}>
          {systems.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
        </select>
      </Field>

      <div className="flex items-center justify-between mb-2 mt-4">
        <p className="text-xs font-semibold text-slate-500 uppercase tracking-wide">Custom grading systems</p>
        <Btn size="sm" variant="secondary" onClick={() => setSysModal(blankGradingSystem())}><Plus size={13} /> New system</Btn>
      </div>
      {custom.length === 0 ? (
        <p className="text-xs text-slate-400 mb-4">No custom systems yet — using the built-in Standard and CBC scales below.</p>
      ) : (
        <div className="space-y-2 mb-4">
          {custom.map((s) => (
            <div key={s.id} className="flex items-center justify-between gap-2 bg-stone-50 rounded-md px-3 py-2">
              <div className="min-w-0">
                <p className="text-sm text-slate-700 truncate">{s.name}</p>
                <p className="text-xs text-slate-400">{s.bands.length} band(s)</p>
              </div>
              <div className="flex gap-1 shrink-0">
                <button onClick={() => setSysModal(s)} className="p-1 text-slate-400 hover:text-slate-700"><Edit2 size={13} /></button>
                <button onClick={() => deleteSystem(s.id)} className="p-1 text-slate-400 hover:text-red-600"><Trash2 size={13} /></button>
              </div>
            </div>
          ))}
        </div>
      )}

      <div className="mb-2">
        <p className="text-xs font-semibold text-slate-500 uppercase tracking-wide mb-2">Per-subject grading</p>
        <div className="max-h-72 overflow-y-auto space-y-1.5">
          {ALL_SUBJECTS.map((subj) => (
            <div key={subj} className="flex items-center justify-between gap-2">
              <span className="text-xs text-slate-600 truncate flex-1">{subj}</span>
              <select className="text-xs rounded border border-slate-300 px-2 py-1 max-w-[45%]" value={overrides[subj] || "__default__"} onChange={(e) => setSubjectOverride(subj, e.target.value)}>
                <option value="__default__">Use default</option>
                {systems.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
              <button onClick={() => editSubjectGrading(subj)} className="p-1 text-slate-400 hover:text-emerald-700 shrink-0" title={`Edit grading for ${subj}`}><Edit2 size={13} /></button>
            </div>
          ))}
        </div>
      </div>

      {sysModal && (
        <Modal title={sysModal.forSubject ? `Edit grading for ${sysModal.forSubject}` : sysModal.id ? "Edit grading system" : "New grading system"} onClose={() => setSysModal(null)} wide>
          <GradingSystemFormBody system={sysModal} onSave={saveSystem} onClose={() => setSysModal(null)} />
        </Modal>
      )}
    </div>
  );
}

function GradingSystemFormBody({ system, onSave, onClose }) {
  const [f, setF] = useState(system);
  const setBand = (i, key, val) => {
    const bands = [...f.bands];
    bands[i] = { ...bands[i], [key]: val };
    setF({ ...f, bands });
  };
  const addBand = () => setF({ ...f, bands: [...f.bands, { label: "", min: "" }] });
  const removeBand = (i) => setF({ ...f, bands: f.bands.filter((_, idx) => idx !== i) });

  return (
    <>
      <Field label="System name"><input className={inputCls} value={f.name} onChange={(e) => setF({ ...f, name: e.target.value })} placeholder="e.g. Our School's CBC Scale" /></Field>
      <span className="block text-xs font-semibold text-slate-500 uppercase tracking-wide mb-1.5">Bands (label + minimum %)</span>
      <div className="space-y-1.5 mb-2">
        {f.bands.map((b, i) => (
          <div key={i} className="flex items-center gap-2">
            <input className={inputCls} value={b.label} onChange={(e) => setBand(i, "label", e.target.value)} placeholder="e.g. EE1" />
            <div className="flex items-center gap-1 shrink-0">
              <span className="text-xs text-slate-400">min</span>
              <input type="number" min="0" max="100" className="w-16 rounded border border-slate-300 px-2 py-2 text-center text-sm" value={b.min} onChange={(e) => setBand(i, "min", clamp(e.target.value, 0, 100))} />
              <span className="text-xs text-slate-400">%</span>
            </div>
            <button onClick={() => removeBand(i)} className="text-slate-400 hover:text-red-600 shrink-0"><X size={14} /></button>
          </div>
        ))}
      </div>
      <Btn size="sm" variant="secondary" onClick={addBand}><Plus size={13} /> Add band</Btn>
      <p className="text-[11px] text-slate-400 mt-3 mb-3">A learner's grade is the label of the highest band whose minimum % they meet or exceed — bands don't need to be entered in order.</p>
      <div className="flex justify-end gap-2 mt-2">
        <Btn variant="secondary" onClick={onClose}>Cancel</Btn>
        <Btn onClick={() => f.name.trim() && onSave(f)} disabled={!f.name.trim()}><Save size={14} /> Save</Btn>
      </div>
    </>
  );
}

function TrashSection({ data, persist, notify }) {
  const trash = data.trash || [];
  const TRASH_DAYS = 30;

  const label = (entry) => {
    if (entry.type === "learner") return entry.item?.name || "Unnamed learner";
    if (entry.type === "teacher") return entry.item?.name || "Unnamed teacher";
    if (entry.type === "exam") return entry.item?.name || "Unnamed exam";
    return "Item";
  };

  const daysAgo = (iso) => Math.floor((Date.now() - new Date(iso).getTime()) / (1000 * 60 * 60 * 24));

  const restore = (entry) => {
    let next = { ...data, trash: trash.filter((t) => t.id !== entry.id) };
    if (entry.type === "learner") next.learners = [...data.learners, entry.item];
    if (entry.type === "teacher") next.teachers = [...data.teachers, entry.item];
    if (entry.type === "exam") next.exams = [...data.exams, entry.item];
    persist(next);
    notify(`${label(entry)} restored`);
  };

  const purge = (entry) => {
    persist({ ...data, trash: trash.filter((t) => t.id !== entry.id) });
    notify("Permanently deleted");
  };

  if (!trash.length) return null;

  return (
    <div className="bg-white rounded-xl border border-slate-200 p-5 mt-4">
      <h3 className="font-semibold text-slate-800 mb-1 flex items-center gap-1.5" style={serifStyle()}><Trash2 size={16} className="text-slate-500" /> Recently Deleted</h3>
      <p className="text-xs text-slate-500 mb-4">Learners, teachers, and exams stay here for {TRASH_DAYS} days before you'll want to clear them out for good — restore anything deleted by mistake.</p>
      <div className="divide-y divide-slate-100">
        {trash.map((entry) => (
          <div key={entry.id} className="py-2.5 flex items-center justify-between gap-2">
            <div className="min-w-0">
              <p className="text-sm text-slate-800 truncate">{label(entry)} <span className="text-xs text-slate-400 capitalize">({entry.type})</span></p>
              <p className="text-xs text-slate-400">Deleted {daysAgo(entry.deletedAt)} day(s) ago</p>
            </div>
            <div className="flex gap-1 shrink-0">
              <Btn size="sm" variant="secondary" onClick={() => restore(entry)}><RotateCcw size={13} /> Restore</Btn>
              <Btn size="sm" variant="danger" onClick={() => purge(entry)}><Trash2 size={13} /></Btn>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function Dashboard({ data, setTab, user, school, onLogout, persist, notify, onActivateSubscription }) {
  const classes = new Set(data.learners.map((l) => l.class).filter(Boolean));
  const chartData = classAverages(data);
  const trendData = examTrend(data);
  const top = topPerformers(data);
  const hour = new Date().getHours();
  const greeting = hour < 12 ? "Good morning" : hour < 17 ? "Good afternoon" : "Good evening";
  const [copied, setCopied] = useState(false);
  const [payModal, setPayModal] = useState(false);

  const copyCode = () => {
    if (!school) return;
    navigator.clipboard.writeText(school.code).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 1800);
    });
  };

  const cards = [
    { label: "Learners", value: data.learners.length, icon: Users, tab: "learners", color: "bg-emerald-800" },
    { label: "Teachers", value: data.teachers.length, icon: GraduationCap, tab: "teachers", color: "bg-slate-700" },
    { label: "Classes", value: classes.size, icon: BookOpen, tab: "classlist", color: "bg-amber-600" },
    { label: "Exams recorded", value: data.exams.length, icon: ClipboardList, tab: "exams", color: "bg-emerald-800" },
  ];

  const isTrueOwner = !!(school && user.id === school.ownerId);

  return (
    <div>
      <div className="bg-gradient-to-r from-slate-900 to-emerald-900 rounded-xl p-5 mb-5 text-white relative overflow-hidden">
        <Sparkles size={80} className="absolute -right-3 -top-3 text-amber-400/20" />
        <p className="text-amber-300 text-xs font-semibold uppercase tracking-wide mb-1 flex items-center gap-2">
          {greeting}
          {isTrueOwner && (
            <span className="inline-flex items-center gap-1 bg-amber-400 text-slate-900 px-2 py-0.5 rounded-full normal-case tracking-normal">
              <Star size={11} fill="#1e293b" /> Owner
            </span>
          )}
        </p>
        <h2 className="text-xl font-bold flex items-center gap-2" style={serifStyle()}>
          Welcome back, {user.name.split(" ")[0]}
          {isTrueOwner && <Star size={18} className="text-amber-400" fill="#fbbf24" />}
        </h2>
        <p className="text-slate-300 text-sm mt-1">
          {data.learners.length ? `You're overseeing ${data.learners.length} learner${data.learners.length === 1 ? "" : "s"} across ${classes.size} class${classes.size === 1 ? "" : "es"}.` : "Add your first learner to get started."}
        </p>
      </div>

      {(user.role === "Super Admin" || user.role === "Admin") && school && school.plan === "trial" && (
        <div className={`rounded-xl border p-3 mb-4 text-sm flex flex-wrap items-center justify-between gap-2 ${trialDaysLeft(school) <= 5 ? "bg-red-50 border-red-200 text-red-800" : "bg-amber-50 border-amber-200 text-amber-800"}`}>
          <span className="flex items-center gap-2"><Sparkles size={15} className="shrink-0" /> {trialDaysLeft(school)} day{trialDaysLeft(school) === 1 ? "" : "s"} left in your free trial — then KES {MONTHLY_PRICE}/mo or KES {YEARLY_PRICE}/yr.</span>
          <Btn size="sm" variant="gold" onClick={() => setPayModal(true)}><Phone size={13} /> Pay / renew now</Btn>
        </div>
      )}
      {(user.role === "Super Admin" || user.role === "Admin") && school && school.plan !== "trial" && (
        <div className="rounded-xl border border-emerald-200 bg-emerald-50 p-3 mb-4 text-sm text-emerald-800 flex flex-wrap items-center justify-between gap-2">
          <span className="flex items-center gap-2"><CheckCircle2 size={15} className="shrink-0" /> Subscribed ({school.plan === "monthly" ? "Monthly" : "Yearly"}) — active until {new Date(school.subscriptionExpiresAt).toLocaleDateString()}.</span>
          <Btn size="sm" variant="secondary" onClick={() => setPayModal(true)}><Phone size={13} /> Renew / top up</Btn>
        </div>
      )}

      {(user.role === "Super Admin" || user.role === "Admin") && school && (
        <div className="bg-white rounded-xl border border-amber-200 p-4 mb-5 flex items-center justify-between gap-3">
          <div>
            <p className="text-xs font-semibold text-amber-600 uppercase tracking-wide mb-0.5">Your school's invite code</p>
            <p className="text-2xl font-bold tracking-widest text-slate-800" style={serifStyle()}>{school.code}</p>
            <p className="text-xs text-slate-500 mt-0.5">Share this with teachers so they can join {school.name} — they enter it when creating their account.</p>
          </div>
          <Btn size="sm" variant="gold" onClick={copyCode}><Copy size={13} /> {copied ? "Copied!" : "Copy"}</Btn>
        </div>
      )}

      <div className="grid grid-cols-2 gap-3 mb-5">
        {cards.map((c) => (
          <button key={c.label} onClick={() => setTab(c.tab)} className="text-left bg-white rounded-xl border border-slate-200 p-4 hover:shadow-md transition-shadow">
            <div className={`w-8 h-8 rounded-lg ${c.color} text-white flex items-center justify-center mb-2`}>
              <c.icon size={16} />
            </div>
            <div className="text-2xl font-bold text-slate-800" style={serifStyle()}>{c.value}</div>
            <div className="text-xs text-slate-500">{c.label}</div>
          </button>
        ))}
      </div>

      {chartData.length > 0 && (
        <div className="bg-white rounded-xl border border-slate-200 p-4 mb-5">
          <h2 className="font-bold text-slate-800 mb-3 flex items-center gap-2" style={serifStyle()}><TrendingUp size={16} className="text-emerald-700" /> Average score by class</h2>
          <div style={{ width: "100%", height: 220 }}>
            <ResponsiveContainer>
              <BarChart data={chartData} margin={{ top: 4, right: 8, left: -20, bottom: 4 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" vertical={false} />
                <XAxis dataKey="class" tick={{ fontSize: 11, fill: "#64748b" }} />
                <YAxis tick={{ fontSize: 11, fill: "#64748b" }} domain={[0, 100]} />
                <Tooltip contentStyle={{ fontSize: 12, borderRadius: 8 }} />
                <Bar dataKey="average" fill="#065f46" radius={[4, 4, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      )}

      {trendData.length > 1 && (
        <div className="bg-white rounded-xl border border-slate-200 p-4 mb-5">
          <h2 className="font-bold text-slate-800 mb-3 flex items-center gap-2" style={serifStyle()}><TrendingUp size={16} className="text-emerald-700" /> Performance trend across exams</h2>
          <div style={{ width: "100%", height: 220 }}>
            <ResponsiveContainer>
              <LineChart data={trendData} margin={{ top: 4, right: 8, left: -20, bottom: 4 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" vertical={false} />
                <XAxis dataKey="name" tick={{ fontSize: 10, fill: "#64748b" }} />
                <YAxis tick={{ fontSize: 11, fill: "#64748b" }} domain={[0, 100]} />
                <Tooltip contentStyle={{ fontSize: 12, borderRadius: 8 }} />
                <Line type="monotone" dataKey="average" stroke="#065f46" strokeWidth={2} dot={{ r: 3 }} name="School average %" />
              </LineChart>
            </ResponsiveContainer>
          </div>
          <p className="text-[11px] text-slate-400 mt-1">Overall average percentage across learners, exam by exam in date order — watch for the trend, not just one result.</p>
        </div>
      )}

      <div className="grid sm:grid-cols-2 gap-5 mb-5">
        <div className="bg-white rounded-xl border border-slate-200 p-4">
          <h2 className="font-bold text-slate-800 mb-3 flex items-center gap-2" style={serifStyle()}><Star size={16} className="text-amber-500" /> Top performers</h2>
          {top.length === 0 ? (
            <p className="text-xs text-slate-400">Record some marks to see top performers here.</p>
          ) : (
            <div className="space-y-2">
              {top.map((t, i) => (
                <div key={t.learner.id} className="flex items-center gap-2.5">
                  <span className={`w-6 h-6 rounded-full flex items-center justify-center text-xs font-bold ${i === 0 ? "bg-amber-400 text-white" : "bg-slate-100 text-slate-500"}`}>{i + 1}</span>
                  <span className="flex-1 text-sm text-slate-700 truncate">{t.learner.name}</span>
                  <span className="text-sm font-semibold text-emerald-800">{t.avg.toFixed(1)}</span>
                </div>
              ))}
            </div>
          )}
        </div>

        <div className="bg-white rounded-xl border border-slate-200 p-4">
          <h2 className="font-bold text-slate-800 mb-3 flex items-center gap-2" style={serifStyle()}><ChevronRight size={16} className="text-amber-500" /> Quick actions</h2>
          <div className="flex flex-wrap gap-2">
            <Btn size="sm" variant="secondary" onClick={() => setTab("learners")}><Plus size={14} /> Add learner</Btn>
            <Btn size="sm" variant="secondary" onClick={() => setTab("profile")}><ImageIcon size={14} /> Upload school logo</Btn>
            <Btn size="sm" variant="secondary" onClick={() => setTab("exams")}><ClipboardList size={14} /> Create exam</Btn>
            <Btn size="sm" variant="secondary" onClick={() => setTab("exams")}><Upload size={14} /> Upload marks</Btn>
            <Btn size="sm" variant="secondary" onClick={() => setTab("merit")}><Award size={14} /> View merit list</Btn>
            <Btn size="sm" variant="secondary" onClick={() => setTab("messages")}><Send size={14} /> Message parents</Btn>
            <Btn size="sm" variant="danger" onClick={onLogout}><LogOut size={14} /> Log out</Btn>
          </div>
        </div>
      </div>

      {data.learners.length === 0 && (
        <div className="bg-amber-50 border border-amber-200 rounded-xl p-4 text-sm text-amber-800">
          No learners yet. Start in <b>Learners</b> to add profiles, or go to <b>Exams &amp; Marks</b> to bulk-import from a class list spreadsheet.
        </div>
      )}

      {payModal && (
        <Modal title="Pay / renew subscription" onClose={() => setPayModal(false)}>
          <SubscriptionPanel onActivate={(plan, code) => { const r = onActivateSubscription(plan, code); if (r && r.ok) setPayModal(false); return r; }} />
        </Modal>
      )}
    </div>
  );
}

// ---------- Learners ----------
function blankLearner() {
  return { id: "", name: "", admissionNo: "", class: "", gender: "", dob: "", parentName: "", parentEmail: "", parentPhone: "", address: "", kcpeScore: "", kjseaScore: "" };
}

function Learners({ data, persist, notify }) {
  const [modal, setModal] = useState(null); // learner obj or null
  const [importModal, setImportModal] = useState(false);
  const [promoteModal, setPromoteModal] = useState(false);
  const [q, setQ] = useState("");
  const [profileId, setProfileId] = useState(null);

  const filtered = data.learners.filter((l) =>
    `${l.name} ${l.admissionNo} ${l.class}`.toLowerCase().includes(q.toLowerCase())
  );

  const save = (learner) => {
    const exists = data.learners.some((l) => l.id === learner.id);
    const next = exists
      ? { ...data, learners: data.learners.map((l) => (l.id === learner.id ? learner : l)) }
      : { ...data, learners: [...data.learners, { ...learner, id: uid() }] };
    persist(next);
    notify(exists ? "Learner updated" : "Learner added");
    setModal(null);
  };

  const remove = (id) => {
    const learner = data.learners.find((l) => l.id === id);
    trashItem(data, persist, "learner", learner, { learners: data.learners.filter((l) => l.id !== id) });
    notify("Learner moved to trash — restore it from School Profile within 30 days");
    setProfileId(null);
  };

  const profile = data.learners.find((l) => l.id === profileId);

  if (profile) {
    const learnerMarks = data.marks.filter((m) => m.learnerId === profile.id);
    const byExam = {};
    learnerMarks.forEach((m) => { (byExam[m.examId] = byExam[m.examId] || []).push(m); });
    return (
      <div>
        <button onClick={() => setProfileId(null)} className="text-sm text-emerald-800 font-medium mb-3 flex items-center gap-1"><ChevronRight size={14} className="rotate-180" /> Back to learners</button>
        <div className="bg-white rounded-xl border border-slate-200 p-5 mb-4">
          <div className="flex items-start justify-between">
            <div className="flex items-center gap-3">
              <Seal letter={profile.name.charAt(0).toUpperCase() || "?"} size={48} />
              <div>
                <h2 className="text-xl font-bold text-slate-800" style={serifStyle()}>{profile.name}</h2>
                <p className="text-sm text-slate-500">{profile.class || "No class assigned"} · Adm# {profile.admissionNo || "—"}</p>
              </div>
            </div>
            <div className="flex gap-2">
              <Btn size="sm" variant="secondary" onClick={() => setModal(profile)}><Edit2 size={13} /> Edit</Btn>
              <Btn size="sm" variant="danger" onClick={() => remove(profile.id)}><Trash2 size={13} /> Delete</Btn>
            </div>
          </div>
          <div className="grid grid-cols-2 gap-3 mt-5 text-sm">
            <div><span className="text-slate-400 block text-xs uppercase tracking-wide">Gender</span>{profile.gender || "—"}</div>
            <div><span className="text-slate-400 block text-xs uppercase tracking-wide">Date of birth</span>{profile.dob || "—"}</div>
            <div><span className="text-slate-400 block text-xs uppercase tracking-wide">Parent / Guardian</span>{profile.parentName || "—"}</div>
            <div><span className="text-slate-400 block text-xs uppercase tracking-wide">Parent phone</span>{profile.parentPhone || "—"}</div>
            <div><span className="text-slate-400 block text-xs uppercase tracking-wide">Parent email</span>{profile.parentEmail || "—"}</div>
            <div><span className="text-slate-400 block text-xs uppercase tracking-wide">Address</span>{profile.address || "—"}</div>
            <div><span className="text-slate-400 block text-xs uppercase tracking-wide">KCPE score (/500)</span>{profile.kcpeScore || "—"}</div>
            <div><span className="text-slate-400 block text-xs uppercase tracking-wide">KJSEA score (/900)</span>{profile.kjseaScore || "—"}</div>
          </div>
        </div>

        {(() => {
          const trend = learnerTrend(data, profile.id);
          if (trend.length < 2) return null;
          return (
            <div className="bg-white rounded-xl border border-slate-200 p-4 mb-4">
              <h3 className="font-semibold text-sm text-slate-700 mb-2 flex items-center gap-1.5"><TrendingUp size={14} className="text-emerald-700" /> Progress across exams</h3>
              <div style={{ width: "100%", height: 180 }}>
                <ResponsiveContainer>
                  <LineChart data={trend} margin={{ top: 4, right: 8, left: -20, bottom: 4 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" vertical={false} />
                    <XAxis dataKey="name" tick={{ fontSize: 10, fill: "#64748b" }} />
                    <YAxis tick={{ fontSize: 11, fill: "#64748b" }} domain={[0, 100]} />
                    <Tooltip contentStyle={{ fontSize: 12, borderRadius: 8 }} />
                    <Line type="monotone" dataKey="average" stroke="#065f46" strokeWidth={2} dot={{ r: 3 }} name="Average %" />
                  </LineChart>
                </ResponsiveContainer>
              </div>
            </div>
          );
        })()}

        <h3 className="font-bold text-slate-800 mb-2" style={serifStyle()}>Exam history</h3>
        {Object.keys(byExam).length === 0 && <EmptyState icon={ClipboardList} title="No marks recorded yet" />}
        {Object.entries(byExam).map(([examId, marks]) => {
          const exam = data.exams.find((e) => e.id === examId);
          const total = marks.reduce((s, m) => s + Number(m.score || 0), 0);
          return (
            <div key={examId} className="bg-white rounded-xl border border-slate-200 p-4 mb-3">
              <p className="font-semibold text-sm text-slate-700 mb-2">{exam ? exam.name : "Unknown exam"}</p>
              <table className="w-full text-sm">
                <tbody>
                  {marks.map((m) => (
                    <tr key={m.id} className="border-t border-slate-100">
                      <td className="py-1 text-slate-500">{m.subject}</td>
                      <td className="py-1 text-right font-medium text-slate-800">{m.score}</td>
                    </tr>
                  ))}
                  <tr className="border-t border-slate-300">
                    <td className="py-1 font-semibold text-slate-800">Total</td>
                    <td className="py-1 text-right font-bold text-emerald-800">{total}</td>
                  </tr>
                </tbody>
              </table>
            </div>
          );
        })}
        {modal && <LearnerForm learner={modal} onSave={save} onClose={() => setModal(null)} />}
      </div>
    );
  }

  return (
    <div>
      <div className="relative mb-2">
        <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
        <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Search by name, admission no, class…" className={`${inputCls} pl-9`} />
      </div>
      <div className="flex gap-2 mb-2">
        <Btn variant="secondary" onClick={() => setImportModal(true)} className="flex-1 justify-center"><Upload size={15} /> Import from file</Btn>
        <Btn onClick={() => setModal(blankLearner())} className="flex-1 justify-center"><Plus size={15} /> Add</Btn>
      </div>
      <div className="mb-4">
        <Btn variant="secondary" onClick={() => setPromoteModal(true)} className="w-full justify-center"><ArrowRightLeft size={15} /> Promote a class</Btn>
      </div>

      {filtered.length === 0 ? (
        <EmptyState icon={Users} title="No learners found" hint="Add a learner profile, import from a file, or adjust your search." />
      ) : (
        <div className="bg-white rounded-xl border border-slate-200 divide-y divide-slate-100">
          {filtered.map((l) => (
            <button key={l.id} onClick={() => setProfileId(l.id)} className="w-full text-left px-4 py-3 flex items-center gap-3 hover:bg-stone-50">
              <Seal letter={l.name.charAt(0).toUpperCase() || "?"} size={34} />
              <div className="flex-1 min-w-0">
                <p className="font-medium text-slate-800 truncate">{l.name}</p>
                <p className="text-xs text-slate-500">{l.class || "No class"} · Adm# {l.admissionNo || "—"}</p>
              </div>
              <ChevronRight size={16} className="text-slate-300" />
            </button>
          ))}
        </div>
      )}

      {modal && <LearnerForm learner={modal} onSave={save} onClose={() => setModal(null)} />}
      {importModal && <ImportLearnersModal data={data} persist={persist} notify={notify} onClose={() => setImportModal(false)} />}
      {promoteModal && <PromoteClassModal data={data} persist={persist} notify={notify} onClose={() => setPromoteModal(false)} />}
    </div>
  );
}

// ---------- Import learners from a file (Excel/CSV/Word) ----------
function findHeader(headers, patterns, exclude) {
  for (const p of patterns) {
    const h = headers.find((x) => p.test(x) && (!exclude || !exclude.test(x)));
    if (h) return h;
  }
  return null;
}

function mapRowsToLearners(rows) {
  if (!rows.length) return { learners: [], skipped: 0 };
  const headers = Object.keys(rows[0]);
  const nameH = findHeader(headers, [/^name$/i, /full ?name/i, /learner ?name/i, /student ?name/i], /parent|guardian/i)
    || findHeader(headers, [/name/i], /parent|guardian/i);
  const admH = findHeader(headers, [/admission/i, /adm\s?no/i]);
  const classH = findHeader(headers, [/^class$/i, /grade/i]);
  const genderH = findHeader(headers, [/gender/i, /sex/i]);
  const dobH = findHeader(headers, [/dob/i, /birth/i]);
  const parentNameH = findHeader(headers, [/parent.*name/i, /guardian.*name/i]);
  const parentEmailH = findHeader(headers, [/parent.*email/i, /guardian.*email/i, /^email$/i]);
  const parentPhoneH = findHeader(headers, [/parent.*phone/i, /guardian.*phone/i, /^phone$/i, /contact/i]);
  const addressH = findHeader(headers, [/address/i]);

  let skipped = 0;
  const learners = [];
  rows.forEach((row) => {
    const name = nameH ? String(row[nameH] || "").trim() : "";
    if (!name) { skipped++; return; }
    learners.push({
      name,
      admissionNo: admH ? String(row[admH] || "").trim() : "",
      class: classH ? String(row[classH] || "").trim() : "",
      gender: genderH ? String(row[genderH] || "").trim() : "",
      dob: dobH ? String(row[dobH] || "").trim() : "",
      parentName: parentNameH ? String(row[parentNameH] || "").trim() : "",
      parentEmail: parentEmailH ? String(row[parentEmailH] || "").trim() : "",
      parentPhone: parentPhoneH ? String(row[parentPhoneH] || "").trim() : "",
      address: addressH ? String(row[addressH] || "").trim() : "",
    });
  });
  return { learners, skipped, matchedHeaders: { nameH, admH, classH, genderH, dobH, parentNameH, parentEmailH, parentPhoneH, addressH } };
}

async function parseExcelOrCsv(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  if (!rows.length) throw new Error("That file appears to be empty.");
  return rows;
}

function ImportLearnersModal({ data, persist, notify, onClose }) {
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const [preview, setPreview] = useState(null); // { learners, skipped, duplicates }
  const fileRef = useRef(null);

  const handleFile = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setError("");
    setPreview(null);

    const ext = file.name.split(".").pop().toLowerCase();
    if (ext === "pdf" || ext === "docx" || ext === "doc") {
      setError(`${ext.toUpperCase()} files can't be read automatically here — please save/export the file as Excel (.xlsx) or CSV first, then upload that instead.`);
      if (fileRef.current) fileRef.current.value = "";
      return;
    }

    setBusy(true);
    try {
      let rows;
      if (["xlsx", "xls", "csv"].includes(ext)) rows = await parseExcelOrCsv(file);
      else throw new Error("Unsupported file type. Please upload .xlsx, .xls, or .csv.");

      const { learners, skipped } = mapRowsToLearners(rows);
      if (!learners.length) throw new Error("Couldn't find a usable 'Name' column in this file. Make sure one column is clearly labeled Name.");

      const withDupFlag = learners.map((l) => ({
        ...l,
        isDuplicate: data.learners.some((ex) =>
          (l.admissionNo && ex.admissionNo && ex.admissionNo.trim().toLowerCase() === l.admissionNo.trim().toLowerCase()) ||
          (ex.name.trim().toLowerCase() === l.name.trim().toLowerCase() && ex.class.trim().toLowerCase() === l.class.trim().toLowerCase())
        ),
      }));
      setPreview({ learners: withDupFlag, skipped });
    } catch (err) {
      setError(err.message || "Couldn't read that file.");
    }
    setBusy(false);
    if (fileRef.current) fileRef.current.value = "";
  };

  const confirmImport = (skipDuplicates) => {
    const toImport = preview.learners.filter((l) => !skipDuplicates || !l.isDuplicate);
    const newLearners = toImport.map((l) => ({ ...blankLearner(), ...l, id: uid() }));
    persist({ ...data, learners: [...data.learners, ...newLearners] });
    notify(`Imported ${newLearners.length} learner(s)`);
    onClose();
  };

  const dupCount = preview ? preview.learners.filter((l) => l.isDuplicate).length : 0;

  return (
    <Modal title="Import learners from a file" onClose={onClose} wide>
      <p className="text-xs text-slate-500 mb-3">
        Upload an Excel or CSV file with a table of learners. Include a column for <b>Name</b> at minimum — Admission No, Class, Gender, Date of Birth, Parent Name, Parent Email, Parent Phone, and Address will be picked up automatically if present, however they're labeled. Word and PDF files aren't supported for automatic reading — export those as Excel or CSV first.
      </p>
      <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv,.docx,.doc,.pdf" onChange={handleFile} className="hidden" disabled={busy} />
      <button
        type="button"
        onClick={() => fileRef.current && fileRef.current.click()}
        disabled={busy}
        className="w-full flex items-center justify-center gap-2 border-2 border-dashed border-slate-300 rounded-lg py-3 text-sm text-slate-600 hover:border-emerald-600 hover:text-emerald-700 disabled:opacity-50 mb-2"
      >
        <Upload size={15} /> {busy ? "Reading file…" : "Tap to choose a file"}
      </button>
      {error && <p className="text-xs text-red-600 mb-3 flex items-center gap-1"><AlertCircle size={13} /> {error}</p>}

      {preview && (
        <div className="mt-2">
          <div className="bg-stone-50 border border-slate-200 rounded-lg p-3 text-xs mb-3">
            <p className="text-slate-700 mb-1">
              Found <b>{preview.learners.length}</b> learner(s) with a name{preview.skipped > 0 ? `, skipped ${preview.skipped} row(s) with no name` : ""}.
              {dupCount > 0 && ` ${dupCount} look like duplicates of learners already in your list.`}
            </p>
          </div>
          <div className="max-h-56 overflow-y-auto border border-slate-200 rounded-lg divide-y divide-slate-100 mb-3">
            {preview.learners.slice(0, 50).map((l, i) => (
              <div key={i} className="px-3 py-2 text-xs flex items-center justify-between gap-2">
                <div className="min-w-0">
                  <p className="font-medium text-slate-800 truncate">{l.name}</p>
                  <p className="text-slate-500">{l.class || "No class"} · Adm# {l.admissionNo || "—"}</p>
                </div>
                {l.isDuplicate && <span className="text-[10px] bg-amber-100 text-amber-700 px-1.5 py-0.5 rounded-full shrink-0">Possible duplicate</span>}
              </div>
            ))}
            {preview.learners.length > 50 && <p className="px-3 py-2 text-xs text-slate-400">…and {preview.learners.length - 50} more</p>}
          </div>
          <div className="flex flex-wrap gap-2">
            <Btn size="sm" onClick={() => confirmImport(true)}><CheckCircle2 size={13} /> Import {preview.learners.length - dupCount} new learner(s)</Btn>
            {dupCount > 0 && <Btn size="sm" variant="secondary" onClick={() => confirmImport(false)}>Import all {preview.learners.length} anyway</Btn>}
            <Btn size="sm" variant="secondary" onClick={() => setPreview(null)}>Cancel</Btn>
          </div>
        </div>
      )}

      {!preview && (
        <div className="flex justify-end mt-2">
          <Btn variant="secondary" onClick={onClose}>Close</Btn>
        </div>
      )}
    </Modal>
  );
}

function PromoteClassModal({ data, persist, notify, onClose }) {
  const classes = [...new Set(data.learners.map((l) => l.class).filter(Boolean))];
  const [fromClass, setFromClass] = useState(classes[0] || "");
  const [toClass, setToClass] = useState("");
  const [checked, setChecked] = useState({}); // learnerId -> boolean (true = promote)
  const [step, setStep] = useState("select"); // select | confirm

  const fromLearners = data.learners.filter((l) => l.class === fromClass);
  const isChecked = (id) => checked[id] !== false; // default to checked (promote)
  const toggle = (id) => setChecked((c) => ({ ...c, [id]: !isChecked(id) }));
  const promoteCount = fromLearners.filter((l) => isChecked(l.id)).length;
  const repeatCount = fromLearners.length - promoteCount;

  const doPromote = () => {
    const next = data.learners.map((l) => {
      if (l.class !== fromClass) return l;
      return isChecked(l.id) ? { ...l, class: toClass.trim() } : l;
    });
    persist({ ...data, learners: next });
    notify(`Promoted ${promoteCount} learner(s) to ${toClass}${repeatCount ? `, ${repeatCount} left in ${fromClass}` : ""}`);
    onClose();
  };

  return (
    <Modal title="Promote a class" onClose={onClose} wide>
      {step === "select" ? (
        <>
          <p className="text-xs text-slate-500 mb-3">Move every learner from one class up to the next. Uncheck anyone who's repeating — they'll stay in their current class.</p>
          <div className="grid sm:grid-cols-2 gap-x-4">
            <Field label="From class">
              <select className={inputCls} value={fromClass} onChange={(e) => setFromClass(e.target.value)}>
                {classes.map((c) => <option key={c} value={c}>{c}</option>)}
              </select>
            </Field>
            <Field label="To class"><input className={inputCls} value={toClass} onChange={(e) => setToClass(e.target.value)} placeholder="e.g. Grade 7A" /></Field>
          </div>

          {fromLearners.length === 0 ? (
            <EmptyState icon={Users} title="No learners in this class" />
          ) : (
            <div className="max-h-64 overflow-y-auto border border-slate-200 rounded-lg divide-y divide-slate-100 mb-3">
              {fromLearners.map((l) => (
                <label key={l.id} className="flex items-center gap-2 px-3 py-2 text-sm cursor-pointer">
                  <input type="checkbox" checked={isChecked(l.id)} onChange={() => toggle(l.id)} className="accent-emerald-700" />
                  <span className={isChecked(l.id) ? "text-slate-800" : "text-slate-400 line-through"}>{l.name}</span>
                  {!isChecked(l.id) && <span className="text-[10px] text-amber-600 ml-auto">Repeating</span>}
                </label>
              ))}
            </div>
          )}
          <p className="text-xs text-slate-500 mb-4">{promoteCount} promoted to {toClass || "…"}, {repeatCount} staying in {fromClass}.</p>
          <div className="flex justify-end gap-2">
            <Btn variant="secondary" onClick={onClose}>Cancel</Btn>
            <Btn onClick={() => setStep("confirm")} disabled={!toClass.trim() || !fromLearners.length}>Review</Btn>
          </div>
        </>
      ) : (
        <>
          <div className="bg-amber-50 border border-amber-200 rounded-lg p-3 text-sm text-amber-800 mb-4">
            <p className="mb-1">You're about to move <b>{promoteCount} learner(s)</b> from <b>{fromClass}</b> to <b>{toClass}</b>.</p>
            {repeatCount > 0 && <p><b>{repeatCount} learner(s)</b> will stay in {fromClass}.</p>}
          </div>
          <div className="flex justify-end gap-2">
            <Btn variant="secondary" onClick={() => setStep("select")}>Back</Btn>
            <Btn onClick={doPromote}><ArrowRightLeft size={14} /> Confirm promotion</Btn>
          </div>
        </>
      )}
    </Modal>
  );
}

function LearnerForm({ learner, onSave, onClose }) {
  const [f, setF] = useState(learner);
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  return (
    <Modal title={f.id ? "Edit learner" : "Add learner"} onClose={onClose} wide>
      <div className="grid sm:grid-cols-2 gap-x-4">
        <Field label="Full name"><input className={inputCls} value={f.name} onChange={set("name")} placeholder="e.g. Amara Chukwu" /></Field>
        <Field label="Admission number"><input className={inputCls} value={f.admissionNo} onChange={set("admissionNo")} placeholder="e.g. 2026-014" /></Field>
        <Field label="Class"><input className={inputCls} value={f.class} onChange={set("class")} placeholder="e.g. Grade 6B" /></Field>
        <Field label="Gender">
          <select className={inputCls} value={f.gender} onChange={set("gender")}>
            <option value="">Select…</option>
            <option>Female</option><option>Male</option><option>Other</option>
          </select>
        </Field>
        <Field label="Date of birth"><input type="date" className={inputCls} value={f.dob} onChange={set("dob")} /></Field>
        <Field label="Parent / guardian name"><input className={inputCls} value={f.parentName} onChange={set("parentName")} /></Field>
        <Field label="Parent email"><input type="email" className={inputCls} value={f.parentEmail} onChange={set("parentEmail")} /></Field>
        <Field label="Parent phone"><input className={inputCls} value={f.parentPhone} onChange={set("parentPhone")} /></Field>
      </div>
      <span className="block text-xs font-semibold text-slate-500 uppercase tracking-wide mb-1.5">Entry score (declared once, kept unchanged)</span>
      <div className="grid sm:grid-cols-2 gap-x-4">
        <Field label="KCPE score (out of 500)"><input type="number" min="0" max="500" className={inputCls} value={f.kcpeScore} onChange={(e) => setF({ ...f, kcpeScore: clamp(e.target.value, 0, 500) })} placeholder="e.g. 320" /></Field>
        <Field label="KJSEA score (out of 900)"><input type="number" min="0" max="900" className={inputCls} value={f.kjseaScore} onChange={(e) => setF({ ...f, kjseaScore: clamp(e.target.value, 0, 900) })} placeholder="e.g. 580" /></Field>
      </div>
      <p className="text-[11px] text-slate-400 -mt-2 mb-4">Fill in whichever applies to this learner's entry exam. This is their official declared score — used as-is for Value Added, it isn't meant to change term to term.</p>
      <Field label="Address"><textarea className={inputCls} rows={2} value={f.address} onChange={set("address")} /></Field>
      <div className="flex justify-end gap-2 mt-2">
        <Btn variant="secondary" onClick={onClose}>Cancel</Btn>
        <Btn onClick={() => f.name.trim() && onSave(f)} disabled={!f.name.trim()}><Save size={14} /> Save</Btn>
      </div>
    </Modal>
  );
}

// ---------- Attendance ----------
function Attendance({ data, persist, notify }) {
  const classes = [...new Set(data.learners.map((l) => l.class).filter(Boolean))];
  const [cls, setCls] = useState(classes[0] || "");
  const [date, setDate] = useState(new Date().toISOString().slice(0, 10));
  const [view, setView] = useState("mark"); // mark | summary
  const records = data.attendance || [];

  const classLearners = data.learners.filter((l) => l.class === cls);

  const getStatus = (learnerId) => {
    const r = records.find((a) => a.date === date && a.class === cls && a.learnerId === learnerId);
    return r ? r.status : "present";
  };
  const setStatus = (learnerId, status) => {
    const exists = records.some((a) => a.date === date && a.class === cls && a.learnerId === learnerId);
    const next = exists
      ? records.map((a) => (a.date === date && a.class === cls && a.learnerId === learnerId ? { ...a, status } : a))
      : [...records, { id: uid(), date, class: cls, learnerId, status }];
    persist({ ...data, attendance: next });
  };
  const markAllPresent = () => {
    const next = [...records];
    classLearners.forEach((l) => {
      const idx = next.findIndex((a) => a.date === date && a.class === cls && a.learnerId === l.id);
      if (idx >= 0) next[idx] = { ...next[idx], status: "present" };
      else next.push({ id: uid(), date, class: cls, learnerId: l.id, status: "present" });
    });
    persist({ ...data, attendance: next });
    notify(`Marked ${classLearners.length} learner(s) present for ${date}`);
  };

  const summary = classLearners.map((l) => {
    const learnerRecords = records.filter((a) => a.class === cls && a.learnerId === l.id);
    const present = learnerRecords.filter((a) => a.status === "present").length;
    const total = learnerRecords.length;
    return { learner: l, present, total, rate: total ? Math.round((present / total) * 100) : null };
  });

  if (!classes.length) return <EmptyState icon={CalendarCheck} title="No classes yet" hint="Assign a class to each learner profile to start tracking attendance." />;

  return (
    <div>
      <div className="flex gap-2 mb-4">
        <select className={inputCls} value={cls} onChange={(e) => setCls(e.target.value)}>
          {classes.map((c) => <option key={c} value={c}>{c}</option>)}
        </select>
      </div>
      <div className="flex gap-2 mb-4">
        <button onClick={() => setView("mark")} className={`px-3 py-1.5 rounded-full text-xs font-medium border ${view === "mark" ? "bg-emerald-800 text-white border-emerald-800" : "bg-white text-slate-600 border-slate-300"}`}>Mark attendance</button>
        <button onClick={() => setView("summary")} className={`px-3 py-1.5 rounded-full text-xs font-medium border ${view === "summary" ? "bg-emerald-800 text-white border-emerald-800" : "bg-white text-slate-600 border-slate-300"}`}>Attendance rates</button>
      </div>

      {view === "mark" ? (
        <div className="bg-white rounded-xl border border-slate-200 p-4">
          <div className="flex items-center justify-between mb-3 gap-2">
            <input type="date" className={`${inputCls} max-w-[180px]`} value={date} onChange={(e) => setDate(e.target.value)} />
            <Btn size="sm" variant="secondary" onClick={markAllPresent}><CheckCircle2 size={13} /> Mark all present</Btn>
          </div>
          {classLearners.length === 0 ? (
            <EmptyState icon={Users} title="No learners in this class" />
          ) : (
            <div className="divide-y divide-slate-100">
              {classLearners.map((l) => {
                const status = getStatus(l.id);
                return (
                  <div key={l.id} className="flex items-center justify-between py-2.5 gap-2">
                    <span className="text-sm text-slate-700 truncate">{l.name}</span>
                    <div className="flex gap-1 shrink-0">
                      {["present", "absent", "late"].map((s) => (
                        <button
                          key={s}
                          onClick={() => setStatus(l.id, s)}
                          className={`px-2.5 py-1 rounded-full text-[11px] font-medium capitalize border ${
                            status === s
                              ? s === "present" ? "bg-emerald-700 text-white border-emerald-700" : s === "absent" ? "bg-red-600 text-white border-red-600" : "bg-amber-500 text-white border-amber-500"
                              : "bg-white text-slate-500 border-slate-300"
                          }`}
                        >
                          {s}
                        </button>
                      ))}
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      ) : (
        <div className="bg-white rounded-xl border border-slate-200 p-4">
          {summary.length === 0 ? <EmptyState icon={CalendarCheck} title="No learners in this class" /> : (
            <table className="w-full text-sm">
              <thead><tr className="text-left text-xs text-slate-400 uppercase tracking-wide border-b border-slate-200">
                <th className="py-1.5">Learner</th><th className="py-1.5 text-right">Days recorded</th><th className="py-1.5 text-right">Present</th><th className="py-1.5 text-right">Rate</th>
              </tr></thead>
              <tbody>
                {summary.map((s) => (
                  <tr key={s.learner.id} className="border-b border-slate-100">
                    <td className="py-1.5 text-slate-800">{s.learner.name}</td>
                    <td className="py-1.5 text-right text-slate-400">{s.total}</td>
                    <td className="py-1.5 text-right text-slate-500">{s.present}</td>
                    <td className="py-1.5 text-right font-semibold">
                      {s.rate === null ? <span className="text-slate-300">—</span> : (
                        <span className={s.rate >= 90 ? "text-emerald-700" : s.rate >= 75 ? "text-amber-600" : "text-red-600"}>{s.rate}%</span>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      )}
    </div>
  );
}

// ---------- Fees & Finance ----------
function blankFeeStructure() { return { id: "", class: "", term: "Term 1", year: String(new Date().getFullYear()), amount: "" }; }
function blankPayment() { return { id: "", learnerId: "", amount: "", date: new Date().toISOString().slice(0, 10), method: "M-Pesa", reference: "", term: "Term 1", year: String(new Date().getFullYear()) }; }

function Fees({ data, persist, notify }) {
  const classes = [...new Set(data.learners.map((l) => l.class).filter(Boolean))];
  const [structModal, setStructModal] = useState(null);
  const [payModal, setPayModal] = useState(null);
  const [statementId, setStatementId] = useState(null);
  const [q, setQ] = useState("");
  const structures = data.feeStructures || [];
  const payments = data.feePayments || [];

  const expectedFor = (learner, term, year) => {
    const s = structures.find((x) => x.class === learner.class && x.term === term && x.year === year);
    return s ? Number(s.amount) : 0;
  };
  const paidFor = (learnerId, term, year) => payments.filter((p) => p.learnerId === learnerId && p.term === term && p.year === year).reduce((s, p) => s + Number(p.amount || 0), 0);

  const [term, setTerm] = useState("Term 1");
  const [year, setYear] = useState(String(new Date().getFullYear()));

  const rows = data.learners
    .filter((l) => `${l.name} ${l.admissionNo}`.toLowerCase().includes(q.toLowerCase()))
    .map((l) => {
      const expected = expectedFor(l, term, year);
      const paid = paidFor(l.id, term, year);
      return { learner: l, expected, paid, balance: expected - paid };
    });

  const saveStructure = (s) => {
    const exists = structures.some((x) => x.id === s.id);
    const next = exists ? structures.map((x) => (x.id === s.id ? s : x)) : [...structures, { ...s, id: uid() }];
    persist({ ...data, feeStructures: next });
    notify(exists ? "Fee structure updated" : "Fee structure set");
    setStructModal(null);
  };

  const savePayment = (p) => {
    const next = [...payments, { ...p, id: uid(), amount: Number(p.amount) }];
    persist({ ...data, feePayments: next });
    notify("Payment recorded");
    setPayModal(null);
  };

  return (
    <div>
      <div className="grid grid-cols-2 gap-2 mb-4">
        <select className={inputCls} value={term} onChange={(e) => setTerm(e.target.value)}>
          <option>Term 1</option><option>Term 2</option><option>Term 3</option>
        </select>
        <input className={inputCls} value={year} onChange={(e) => setYear(e.target.value)} placeholder="Year" />
      </div>

      <div className="flex gap-2 mb-4">
        <Btn size="sm" variant="secondary" onClick={() => setStructModal(blankFeeStructure())}><Plus size={14} /> Set fee structure</Btn>
        <Btn size="sm" onClick={() => setPayModal(blankPayment())}><Plus size={14} /> Record payment</Btn>
      </div>

      {structures.filter((s) => s.term === term && s.year === year).length > 0 && (
        <div className="bg-white rounded-xl border border-slate-200 p-3 mb-4">
          <p className="text-xs font-semibold text-slate-500 uppercase tracking-wide mb-2">Fee structure — {term} {year}</p>
          <div className="flex flex-wrap gap-2">
            {structures.filter((s) => s.term === term && s.year === year).map((s) => (
              <span key={s.id} className="text-xs bg-stone-100 text-slate-700 px-2.5 py-1 rounded-full">{s.class}: KES {Number(s.amount).toLocaleString()}</span>
            ))}
          </div>
        </div>
      )}

      <div className="relative mb-3">
        <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
        <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Search learner…" className={`${inputCls} pl-9`} />
      </div>

      {rows.length === 0 ? (
        <EmptyState icon={Wallet} title="No learners found" />
      ) : (
        <div className="bg-white rounded-xl border border-slate-200 divide-y divide-slate-100">
          {rows.map((r) => (
            <button key={r.learner.id} onClick={() => setStatementId(r.learner.id)} className="w-full text-left px-4 py-3 flex items-center gap-3 hover:bg-stone-50">
              <Seal letter={r.learner.name.charAt(0).toUpperCase() || "?"} size={32} />
              <div className="flex-1 min-w-0">
                <p className="font-medium text-slate-800 truncate">{r.learner.name}</p>
                <p className="text-xs text-slate-500">{r.learner.class} · Paid KES {r.paid.toLocaleString()} of {r.expected.toLocaleString()}</p>
              </div>
              <span className={`text-xs font-bold shrink-0 ${r.balance > 0 ? "text-red-600" : "text-emerald-700"}`}>
                {r.balance > 0 ? `Owes ${r.balance.toLocaleString()}` : "Cleared"}
              </span>
              <ChevronRight size={16} className="text-slate-300 shrink-0" />
            </button>
          ))}
        </div>
      )}

      {structModal && (
        <Modal title="Set fee structure" onClose={() => setStructModal(null)}>
          <FeeStructureForm structure={structModal} classes={classes} onSave={saveStructure} onClose={() => setStructModal(null)} />
        </Modal>
      )}
      {payModal && (
        <Modal title="Record payment" onClose={() => setPayModal(null)}>
          <PaymentForm payment={payModal} learners={data.learners} onSave={savePayment} onClose={() => setPayModal(null)} />
        </Modal>
      )}
      {statementId && <FeeStatementModal data={data} learnerId={statementId} onClose={() => setStatementId(null)} />}
    </div>
  );
}

function FeeStatementModal({ data, learnerId, onClose }) {
  const learner = data.learners.find((l) => l.id === learnerId);
  const payments = (data.feePayments || []).filter((p) => p.learnerId === learnerId).sort((a, b) => new Date(a.date) - new Date(b.date));
  const structures = data.feeStructures || [];
  const totalExpected = structures.filter((s) => s.class === learner?.class).reduce((sum, s) => sum + Number(s.amount || 0), 0);
  const totalPaid = payments.reduce((s, p) => s + Number(p.amount || 0), 0);

  const downloadExcel = () => {
    const rows = payments.map((p) => ({ Date: p.date, Term: p.term, Year: p.year, Method: p.method, Reference: p.reference, "Amount (KES)": p.amount }));
    rows.push({ Date: "", Term: "", Year: "", Method: "", Reference: "TOTAL PAID", "Amount (KES)": totalPaid });
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Statement");
    XLSX.writeFile(wb, `${learner.name.replace(/[^a-z0-9]+/gi, "-")}-fee-statement.xlsx`);
  };

  if (!learner) return null;

  return (
    <Modal title="Fee statement" onClose={onClose} wide>
      <div className="flex justify-end gap-2 mb-3 print:hidden">
        <Btn size="sm" variant="secondary" onClick={downloadExcel}><Download size={14} /> Excel</Btn>
        <Btn size="sm" variant="secondary" onClick={() => window.print()}><Printer size={14} /> PDF</Btn>
      </div>
      <div className="flex items-center gap-3 mb-4 border-b-2 border-double border-slate-300 pb-3">
        <Seal letter={learner.name.charAt(0).toUpperCase() || "?"} size={40} />
        <div>
          <p className="font-bold text-slate-800" style={serifStyle()}>{learner.name}</p>
          <p className="text-xs text-slate-500">{learner.class} · Adm# {learner.admissionNo || "—"}</p>
        </div>
      </div>

      {payments.length === 0 ? (
        <EmptyState icon={Wallet} title="No payments recorded yet" />
      ) : (
        <table className="w-full text-sm mb-4">
          <thead><tr className="text-left text-xs text-slate-400 uppercase tracking-wide border-b border-slate-200">
            <th className="py-1.5">Date</th><th className="py-1.5">Term</th><th className="py-1.5">Method</th><th className="py-1.5">Reference</th><th className="py-1.5 text-right">Amount</th>
          </tr></thead>
          <tbody>
            {payments.map((p) => (
              <tr key={p.id} className="border-b border-slate-100">
                <td className="py-1.5 text-slate-600 whitespace-nowrap">{p.date}</td>
                <td className="py-1.5 text-slate-600 whitespace-nowrap">{p.term} {p.year}</td>
                <td className="py-1.5 text-slate-600">{p.method}</td>
                <td className="py-1.5 text-slate-500">{p.reference || "—"}</td>
                <td className="py-1.5 text-right font-medium text-slate-800">{Number(p.amount).toLocaleString()}</td>
              </tr>
            ))}
          </tbody>
        </table>
      )}

      <div className="grid grid-cols-3 gap-2 text-center">
        <div className="bg-stone-50 rounded-lg p-3"><p className="text-[10px] text-slate-400 uppercase tracking-wide">Total paid</p><p className="text-lg font-bold text-emerald-800">{totalPaid.toLocaleString()}</p></div>
        <div className="bg-stone-50 rounded-lg p-3"><p className="text-[10px] text-slate-400 uppercase tracking-wide">Total expected</p><p className="text-lg font-bold text-slate-700">{totalExpected.toLocaleString()}</p></div>
        <div className={`rounded-lg p-3 ${totalExpected - totalPaid > 0 ? "bg-red-50" : "bg-emerald-50"}`}>
          <p className="text-[10px] text-slate-400 uppercase tracking-wide">Balance</p>
          <p className={`text-lg font-bold ${totalExpected - totalPaid > 0 ? "text-red-600" : "text-emerald-700"}`}>{(totalExpected - totalPaid).toLocaleString()}</p>
        </div>
      </div>
    </Modal>
  );
}

// ---------- Discipline ----------
function blankDisciplineRecord() { return { id: "", learnerId: "", date: new Date().toISOString().slice(0, 10), type: "Merit", description: "" }; }
const DISCIPLINE_TYPES = ["Merit", "Demerit", "Incident"];

function Discipline({ data, persist, notify, user }) {
  const [modal, setModal] = useState(null);
  const [q, setQ] = useState("");
  const records = data.disciplineRecords || [];

  const save = (r) => {
    const next = [...records, { ...r, id: uid(), recordedBy: user?.name || "" }];
    persist({ ...data, disciplineRecords: next });
    notify("Record added");
    setModal(null);
  };
  const remove = (id) => { persist({ ...data, disciplineRecords: records.filter((r) => r.id !== id) }); notify("Record removed"); };

  const learnerName = (id) => data.learners.find((l) => l.id === id)?.name || "Unknown learner";
  const filtered = [...records].reverse().filter((r) => learnerName(r.learnerId).toLowerCase().includes(q.toLowerCase()));

  const typeBadge = (type) => {
    if (type === "Merit") return "bg-emerald-100 text-emerald-800";
    if (type === "Demerit") return "bg-amber-100 text-amber-800";
    return "bg-red-100 text-red-800";
  };

  return (
    <div>
      <div className="relative mb-2">
        <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
        <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Search by learner name…" className={`${inputCls} pl-9`} />
      </div>
      <div className="flex justify-end mb-4">
        <Btn onClick={() => setModal(blankDisciplineRecord())}><Plus size={15} /> Add record</Btn>
      </div>

      {filtered.length === 0 ? (
        <EmptyState icon={ShieldAlert} title="No discipline records yet" hint="Log merits, demerits, or incidents for any learner." />
      ) : (
        <div className="bg-white rounded-xl border border-slate-200 divide-y divide-slate-100">
          {filtered.map((r) => (
            <div key={r.id} className="px-4 py-3 flex items-start gap-3">
              <span className={`text-[10px] font-semibold uppercase tracking-wide px-2 py-0.5 rounded-full shrink-0 mt-0.5 ${typeBadge(r.type)}`}>{r.type}</span>
              <div className="flex-1 min-w-0">
                <p className="text-sm font-medium text-slate-800">{learnerName(r.learnerId)}</p>
                <p className="text-xs text-slate-500">{r.description}</p>
                <p className="text-[11px] text-slate-400 mt-0.5">{r.date}{r.recordedBy ? ` · ${r.recordedBy}` : ""}</p>
              </div>
              <button onClick={() => remove(r.id)} className="p-1 text-slate-400 hover:text-red-600 shrink-0"><Trash2 size={14} /></button>
            </div>
          ))}
        </div>
      )}

      {modal && (
        <Modal title="Add discipline record" onClose={() => setModal(null)}>
          <DisciplineFormBody record={modal} learners={data.learners} onSave={save} onClose={() => setModal(null)} />
        </Modal>
      )}
    </div>
  );
}

function DisciplineFormBody({ record, learners, onSave, onClose }) {
  const [f, setF] = useState(record);
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  return (
    <>
      <Field label="Learner">
        <select className={inputCls} value={f.learnerId} onChange={set("learnerId")}>
          <option value="">Select learner…</option>
          {learners.map((l) => <option key={l.id} value={l.id}>{l.name} — {l.class}</option>)}
        </select>
      </Field>
      <div className="grid grid-cols-2 gap-x-4">
        <Field label="Type">
          <select className={inputCls} value={f.type} onChange={set("type")}>
            {DISCIPLINE_TYPES.map((t) => <option key={t}>{t}</option>)}
          </select>
        </Field>
        <Field label="Date"><input type="date" className={inputCls} value={f.date} onChange={set("date")} /></Field>
      </div>
      <Field label="Description"><textarea className={inputCls} rows={3} value={f.description} onChange={set("description")} placeholder="e.g. Helped organize the library shelves" /></Field>
      <div className="flex justify-end gap-2 mt-2">
        <Btn variant="secondary" onClick={onClose}>Cancel</Btn>
        <Btn onClick={() => f.learnerId && f.description.trim() && onSave(f)} disabled={!f.learnerId || !f.description.trim()}><Save size={14} /> Save</Btn>
      </div>
    </>
  );
}

// ---------- Timetable ----------
const TIMETABLE_DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];
const TIMETABLE_PERIODS = [1, 2, 3, 4, 5, 6, 7, 8];

function Timetable({ data, persist, notify }) {
  const classes = [...new Set(data.learners.map((l) => l.class).filter(Boolean))];
  const [cls, setCls] = useState(classes[0] || "");
  const slots = data.timetableSlots || [];
  const [cellModal, setCellModal] = useState(null); // { day, period }

  const getSlot = (day, period) => slots.find((s) => s.class === cls && s.day === day && s.period === period);

  const saveSlot = (subject, teacher) => {
    const { day, period } = cellModal;
    const exists = getSlot(day, period);
    const next = exists
      ? slots.map((s) => (s.class === cls && s.day === day && s.period === period ? { ...s, subject, teacher } : s))
      : [...slots, { id: uid(), class: cls, day, period, subject, teacher }];
    persist({ ...data, timetableSlots: next });
    setCellModal(null);
  };

  const clearSlot = () => {
    const { day, period } = cellModal;
    persist({ ...data, timetableSlots: slots.filter((s) => !(s.class === cls && s.day === day && s.period === period)) });
    setCellModal(null);
  };

  if (!classes.length) return <EmptyState icon={Grid3x3} title="No classes yet" hint="Assign a class to each learner profile to build a timetable." />;

  return (
    <div>
      <select className={`${inputCls} mb-4 max-w-[220px]`} value={cls} onChange={(e) => setCls(e.target.value)}>
        {classes.map((c) => <option key={c} value={c}>{c}</option>)}
      </select>

      <div className="bg-white rounded-xl border border-slate-200 p-3 overflow-x-auto">
        <table className="text-xs min-w-[600px] w-full border-collapse">
          <thead>
            <tr>
              <th className="p-1.5 text-slate-400 font-medium">Period</th>
              {TIMETABLE_DAYS.map((d) => <th key={d} className="p-1.5 text-slate-500 font-semibold">{d}</th>)}
            </tr>
          </thead>
          <tbody>
            {TIMETABLE_PERIODS.map((p) => (
              <tr key={p}>
                <td className="p-1.5 text-center text-slate-400 font-medium border-t border-slate-100">{p}</td>
                {TIMETABLE_DAYS.map((d) => {
                  const slot = getSlot(d, p);
                  return (
                    <td key={d} className="p-1 border-t border-slate-100">
                      <button
                        onClick={() => setCellModal({ day: d, period: p })}
                        className={`w-full h-14 rounded-md text-left px-1.5 py-1 ${slot ? "bg-emerald-50 hover:bg-emerald-100" : "bg-stone-50 hover:bg-stone-100"}`}
                      >
                        {slot ? (
                          <>
                            <p className="font-medium text-emerald-800 truncate">{slot.subject}</p>
                            <p className="text-slate-500 truncate">{slot.teacher}</p>
                          </>
                        ) : (
                          <span className="text-slate-300">+</span>
                        )}
                      </button>
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {cellModal && (
        <Modal title={`${cellModal.day} — Period ${cellModal.period}`} onClose={() => setCellModal(null)}>
          <TimetableCellForm data={data} cls={cls} existing={getSlot(cellModal.day, cellModal.period)} onSave={saveSlot} onClear={clearSlot} onClose={() => setCellModal(null)} />
        </Modal>
      )}
    </div>
  );
}

function TimetableCellForm({ data, cls, existing, onSave, onClear, onClose }) {
  const [subject, setSubject] = useState(existing?.subject || "");
  const [teacher, setTeacher] = useState(existing?.teacher || "");

  const changeSubject = (s) => {
    setSubject(s);
    const t = getSubjectTeacher(data, cls, s);
    if (t) setTeacher(t.name);
  };

  return (
    <>
      <Field label="Subject">
        <select className={inputCls} value={subject} onChange={(e) => changeSubject(e.target.value)}>
          <option value="">Select a subject…</option>
          {ALL_SUBJECTS.map((s) => <option key={s} value={s}>{s}</option>)}
        </select>
      </Field>
      <Field label="Teacher"><input className={inputCls} value={teacher} onChange={(e) => setTeacher(e.target.value)} placeholder="Auto-filled if assigned, or type a name" /></Field>
      <div className="flex justify-between gap-2 mt-2">
        {existing ? <Btn variant="danger" size="sm" onClick={onClear}><Trash2 size={13} /> Clear slot</Btn> : <span />}
        <div className="flex gap-2">
          <Btn variant="secondary" onClick={onClose}>Cancel</Btn>
          <Btn onClick={() => subject && onSave(subject, teacher)} disabled={!subject}><Save size={14} /> Save</Btn>
        </div>
      </div>
    </>
  );
}

function FeeStructureForm({ structure, classes, onSave, onClose }) {
  const [f, setF] = useState(structure);
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  return (
    <>
      <Field label="Class">
        <input list="feeclassopts" className={inputCls} value={f.class} onChange={set("class")} placeholder="e.g. Grade 6B" />
        <datalist id="feeclassopts">{classes.map((c) => <option key={c} value={c} />)}</datalist>
      </Field>
      <div className="grid grid-cols-2 gap-x-4">
        <Field label="Term">
          <select className={inputCls} value={f.term} onChange={set("term")}>
            <option>Term 1</option><option>Term 2</option><option>Term 3</option>
          </select>
        </Field>
        <Field label="Year"><input className={inputCls} value={f.year} onChange={set("year")} /></Field>
      </div>
      <Field label="Amount (KES)"><input type="number" className={inputCls} value={f.amount} onChange={set("amount")} /></Field>
      <div className="flex justify-end gap-2 mt-2">
        <Btn variant="secondary" onClick={onClose}>Cancel</Btn>
        <Btn onClick={() => f.class.trim() && f.amount && onSave(f)} disabled={!f.class.trim() || !f.amount}><Save size={14} /> Save</Btn>
      </div>
    </>
  );
}

function PaymentForm({ payment, learners, onSave, onClose }) {
  const [f, setF] = useState(payment);
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  return (
    <>
      <Field label="Learner">
        <select className={inputCls} value={f.learnerId} onChange={set("learnerId")}>
          <option value="">Select learner…</option>
          {learners.map((l) => <option key={l.id} value={l.id}>{l.name} — {l.class}</option>)}
        </select>
      </Field>
      <div className="grid grid-cols-2 gap-x-4">
        <Field label="Amount (KES)"><input type="number" className={inputCls} value={f.amount} onChange={set("amount")} /></Field>
        <Field label="Date"><input type="date" className={inputCls} value={f.date} onChange={set("date")} /></Field>
      </div>
      <div className="grid grid-cols-2 gap-x-4">
        <Field label="Term">
          <select className={inputCls} value={f.term} onChange={set("term")}>
            <option>Term 1</option><option>Term 2</option><option>Term 3</option>
          </select>
        </Field>
        <Field label="Year"><input className={inputCls} value={f.year} onChange={set("year")} /></Field>
      </div>
      <div className="grid grid-cols-2 gap-x-4">
        <Field label="Method">
          <select className={inputCls} value={f.method} onChange={set("method")}>
            <option>M-Pesa</option><option>Cash</option><option>Bank</option><option>Cheque</option>
          </select>
        </Field>
        <Field label="Reference (optional)"><input className={inputCls} value={f.reference} onChange={set("reference")} placeholder="e.g. M-Pesa code" /></Field>
      </div>
      <div className="flex justify-end gap-2 mt-2">
        <Btn variant="secondary" onClick={onClose}>Cancel</Btn>
        <Btn onClick={() => f.learnerId && f.amount && onSave(f)} disabled={!f.learnerId || !f.amount}><Save size={14} /> Save</Btn>
      </div>
    </>
  );
}

// ---------- Teachers ----------
function blankTeacher() { return { id: "", name: "", subject: "", email: "", phone: "", classesAssigned: "", classTeacherOf: "" }; }
function blankTeacherAccount() { return { name: "", email: "", username: "", password: "", subject: "", classesAssigned: "", classTeacherOf: "" }; }

function Teachers({ data, persist, notify, user, accounts, saveAccounts }) {
  const [modal, setModal] = useState(null);
  const [accountModal, setAccountModal] = useState(null);
  const isOwner = user && (user.role === "Super Admin" || user.role === "Admin");

  const save = (t) => {
    const exists = data.teachers.some((x) => x.id === t.id);
    const next = exists
      ? { ...data, teachers: data.teachers.map((x) => (x.id === t.id ? t : x)) }
      : { ...data, teachers: [...data.teachers, { ...t, id: uid() }] };
    persist(next);
    notify(exists ? "Teacher updated" : "Teacher added");
    setModal(null);
  };
  const remove = (id) => {
    const teacher = data.teachers.find((t) => t.id === id);
    trashItem(data, persist, "teacher", teacher, { teachers: data.teachers.filter((t) => t.id !== id) });
    notify("Teacher moved to trash — restore it from School Profile within 30 days");
  };

  const registerAccount = (f, setError) => {
    const uname = f.username.trim().toLowerCase();
    if (!f.name.trim() || !uname || !f.password) { setError("Name, username and password are required."); return; }
    if (accounts.some((a) => a.username === uname)) { setError("That username is already taken."); return; }
    const account = { id: uid(), name: f.name.trim(), email: f.email.trim(), username: uname, password: f.password, role: "Teacher", schoolId: user.schoolId };
    saveAccounts([...accounts, account]);
    persist({ ...data, teachers: [...data.teachers, { id: uid(), name: f.name.trim(), subject: f.subject, email: f.email.trim(), phone: "", classesAssigned: f.classesAssigned, accountId: account.id }] });
    notify(`Teacher account created — share the username "${uname}" and password with them`);
    setAccountModal(null);
  };

  return (
    <div>
      <div className="flex justify-end gap-2 mb-4">
        <Btn variant="secondary" onClick={() => setModal(blankTeacher())}><Plus size={15} /> Add teacher profile</Btn>
        {isOwner && <Btn onClick={() => setAccountModal(blankTeacherAccount())}><UserPlus size={15} /> Register teacher account</Btn>}
      </div>
      {data.teachers.length === 0 ? (
        <EmptyState icon={GraduationCap} title="No teachers added yet" hint={isOwner ? "Register a teacher account to give them login access, or add a profile-only record." : undefined} />
      ) : (
        <div className="bg-white rounded-xl border border-slate-200 divide-y divide-slate-100">
          {data.teachers.map((t) => (
            <div key={t.id} className="px-4 py-3 flex items-center gap-3">
              <Seal letter={t.name.charAt(0).toUpperCase() || "?"} size={34} />
              <div className="flex-1 min-w-0">
                <p className="font-medium text-slate-800 flex items-center gap-1.5">{t.name} {t.accountId && <span className="text-[10px] bg-emerald-50 text-emerald-700 px-1.5 py-0.5 rounded-full">has login</span>}</p>
                <p className="text-xs text-slate-500">{t.subject || "No subject"} {t.classesAssigned ? `· ${t.classesAssigned}` : ""}</p>
              </div>
              <button onClick={() => setModal(t)} className="p-1.5 text-slate-400 hover:text-slate-700"><Edit2 size={15} /></button>
              <button onClick={() => remove(t.id)} className="p-1.5 text-slate-400 hover:text-red-600"><Trash2 size={15} /></button>
            </div>
          ))}
        </div>
      )}
      {modal && (
        <Modal title={modal.id ? "Edit teacher" : "Add teacher profile"} onClose={() => setModal(null)}>
          <TeacherFormBody t={modal} onSave={save} onClose={() => setModal(null)} />
        </Modal>
      )}
      {accountModal && (
        <Modal title="Register teacher account" onClose={() => setAccountModal(null)}>
          <TeacherAccountFormBody f={accountModal} onSave={registerAccount} onClose={() => setAccountModal(null)} />
        </Modal>
      )}
    </div>
  );
}

function TeacherAccountFormBody({ f: initial, onSave, onClose }) {
  const [f, setF] = useState(initial);
  const [error, setError] = useState("");
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  return (
    <>
      <p className="text-xs text-slate-500 mb-3">This creates a login the teacher can use right away — set a temporary password and share it with them directly.</p>
      <Field label="Full name"><input className={inputCls} value={f.name} onChange={set("name")} /></Field>
      <Field label="Email"><input type="email" className={inputCls} value={f.email} onChange={set("email")} /></Field>
      <div className="grid sm:grid-cols-2 gap-x-4">
        <Field label="Username"><input className={inputCls} value={f.username} onChange={set("username")} /></Field>
        <Field label="Temporary password"><input className={inputCls} value={f.password} onChange={set("password")} /></Field>
      </div>
      <Field label="Subject(s) taught">
        <select className={inputCls} value={f.subject} onChange={set("subject")}>
          <option value="">Select a subject…</option>
          {ALL_SUBJECTS.map((s) => <option key={s} value={s}>{s}</option>)}
        </select>
      </Field>
      <Field label="Classes assigned"><input className={inputCls} value={f.classesAssigned} onChange={set("classesAssigned")} placeholder="e.g. Grade 6A, 6B" /></Field>
      <Field label="Class teacher of (optional)"><input className={inputCls} value={f.classTeacherOf} onChange={set("classTeacherOf")} placeholder="e.g. Grade 6B" /></Field>
      {error && <p className="text-xs text-red-600 mb-3 flex items-center gap-1"><AlertCircle size={13} /> {error}</p>}
      <div className="flex justify-end gap-2 mt-2">
        <Btn variant="secondary" onClick={onClose}>Cancel</Btn>
        <Btn onClick={() => onSave(f, setError)}><UserPlus size={14} /> Register</Btn>
      </div>
    </>
  );
}

function TeacherFormBody({ t, onSave, onClose }) {
  const [f, setF] = useState(t);
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  return (
    <>
      <Field label="Full name"><input className={inputCls} value={f.name} onChange={set("name")} /></Field>
      <Field label="Subject(s) taught">
        <select className={inputCls} value={f.subject} onChange={set("subject")}>
          <option value="">Select a subject…</option>
          {ALL_SUBJECTS.map((s) => <option key={s} value={s}>{s}</option>)}
        </select>
      </Field>
      <Field label="Classes assigned"><input className={inputCls} value={f.classesAssigned} onChange={set("classesAssigned")} placeholder="e.g. Grade 6A, 6B" /></Field>
      <Field label="Class teacher of (optional)"><input className={inputCls} value={f.classTeacherOf} onChange={set("classTeacherOf")} placeholder="e.g. Grade 6B" /></Field>
      <Field label="Email"><input type="email" className={inputCls} value={f.email} onChange={set("email")} /></Field>
      <Field label="Phone"><input className={inputCls} value={f.phone} onChange={set("phone")} /></Field>
      <div className="flex justify-end gap-2 mt-2">
        <Btn variant="secondary" onClick={onClose}>Cancel</Btn>
        <Btn onClick={() => f.name.trim() && onSave(f)} disabled={!f.name.trim()}><Save size={14} /> Save</Btn>
      </div>
    </>
  );
}

// ---------- Exams & Marks ----------
// ---------- Academic Resources hub ----------
function blankResource() { return { id: "", title: "", category: "Exam Paper", attachment: null }; }
const RESOURCE_CATEGORIES = ["Exam Paper", "Marking Scheme", "Scheme of Work", "Lesson Note", "Other"];

function AcademicResources({ data, persist, notify, setTab }) {
  const resources = data.resources || [];
  const [modal, setModal] = useState(null);

  const save = (r) => {
    const exists = resources.some((x) => x.id === r.id);
    const next = exists
      ? { ...data, resources: resources.map((x) => (x.id === r.id ? r : x)) }
      : { ...data, resources: [...resources, { ...r, id: uid() }] };
    persist(next);
    notify(exists ? "Resource updated" : "Resource uploaded");
    setModal(null);
  };
  const remove = (id) => { persist({ ...data, resources: resources.filter((r) => r.id !== id) }); notify("Resource removed"); };

  return (
    <div>
      <div className="grid grid-cols-2 sm:grid-cols-3 gap-3 mb-6">
        <button onClick={() => setTab("exams")} className="text-left bg-white rounded-xl border border-slate-200 p-4 hover:shadow-md transition-shadow">
          <ClipboardList size={18} className="text-emerald-800 mb-2" />
          <p className="font-semibold text-sm text-slate-800">Exams</p>
          <p className="text-xs text-slate-500">Create & manage</p>
        </button>
        <button onClick={() => setTab("schemes")} className="text-left bg-white rounded-xl border border-slate-200 p-4 hover:shadow-md transition-shadow">
          <NotebookPen size={18} className="text-emerald-800 mb-2" />
          <p className="font-semibold text-sm text-slate-800">Schemes of Work</p>
          <p className="text-xs text-slate-500">Term plans, week by week</p>
        </button>
        <button onClick={() => setTab("lessonnotes")} className="text-left bg-white rounded-xl border border-slate-200 p-4 hover:shadow-md transition-shadow">
          <FileText size={18} className="text-emerald-800 mb-2" />
          <p className="font-semibold text-sm text-slate-800">Lesson Notes</p>
          <p className="text-xs text-slate-500">Any subject, any class</p>
        </button>
      </div>

      <div className="flex items-center justify-between mb-3">
        <h2 className="font-bold text-slate-800" style={serifStyle()}>Document library</h2>
        <Btn size="sm" onClick={() => setModal(blankResource())}><Upload size={14} /> Upload document</Btn>
      </div>
      <p className="text-xs text-slate-500 mb-3">Upload any PDF, Word, or Excel file — exam papers, marking schemes, schemes of work, lesson notes, or anything else — and keep them all in one place.</p>

      {resources.length === 0 ? (
        <EmptyState icon={Library} title="No documents uploaded yet" hint="Upload your first PDF, Word, or Excel file above." />
      ) : (
        <div className="bg-white rounded-xl border border-slate-200 divide-y divide-slate-100">
          {resources.map((r) => (
            <div key={r.id} className="px-4 py-3 flex items-center gap-3">
              <FileText size={16} className="text-emerald-700 shrink-0" />
              <div className="flex-1 min-w-0">
                <p className="font-medium text-slate-800 truncate">{r.title || r.attachment?.name || "Untitled"}</p>
                <p className="text-xs text-slate-500">{r.category}{r.attachment ? ` · ${r.attachment.name}` : ""}</p>
              </div>
              {r.attachment && <a href={r.attachment.dataUrl} download={r.attachment.name} className="text-xs text-emerald-800 font-medium shrink-0">Download</a>}
              <button onClick={() => setModal(r)} className="p-1.5 text-slate-400 hover:text-slate-700 shrink-0"><Edit2 size={14} /></button>
              <button onClick={() => remove(r.id)} className="p-1.5 text-slate-400 hover:text-red-600 shrink-0"><Trash2 size={14} /></button>
            </div>
          ))}
        </div>
      )}

      {modal && (
        <Modal title={modal.id ? "Edit resource" : "Upload document"} onClose={() => setModal(null)}>
          <ResourceFormBody resource={modal} onSave={save} onClose={() => setModal(null)} />
        </Modal>
      )}
    </div>
  );
}

function ResourceFormBody({ resource, onSave, onClose }) {
  const [f, setF] = useState(resource);
  return (
    <>
      <Field label="Title"><input className={inputCls} value={f.title} onChange={(e) => setF({ ...f, title: e.target.value })} placeholder="e.g. Mathematics Term 1 Exam Paper" /></Field>
      <Field label="Category">
        <select className={inputCls} value={f.category} onChange={(e) => setF({ ...f, category: e.target.value })}>
          {RESOURCE_CATEGORIES.map((c) => <option key={c}>{c}</option>)}
        </select>
      </Field>
      <Field label="Document (PDF, Word, or Excel)">
        <FileAttachment attachment={f.attachment} onUpload={(a) => setF({ ...f, attachment: a })} onRemove={() => setF({ ...f, attachment: null })} />
      </Field>
      <div className="flex justify-end gap-2 mt-2">
        <Btn variant="secondary" onClick={onClose}>Cancel</Btn>
        <Btn onClick={() => f.attachment && onSave(f)} disabled={!f.attachment}><Save size={14} /> Save</Btn>
      </div>
    </>
  );
}

function ExamsMarks({ data, persist, notify, user }) {
  const [examModal, setExamModal] = useState(null);
  const [activeExamId, setActiveExamId] = useState(null);
  const [fileError, setFileError] = useState("");
  const [preview, setPreview] = useState(null);
  const fileRef = useRef(null);

  const isAdmin = user.role === "Super Admin" || user.role === "Admin";
  const isTeacher = user.role === "Teacher";
  const myTeacherProfile = data.teachers.find((t) => t.accountId === user.id);
  const mySubjects = myTeacherProfile && myTeacherProfile.subject ? [myTeacherProfile.subject] : [];
  const myClasses = myTeacherProfile && myTeacherProfile.classesAssigned
    ? myTeacherProfile.classesAssigned.split(",").map((s) => s.trim()).filter(Boolean)
    : [];

  const visibleExams = isTeacher ? data.exams.filter((e) => myClasses.includes(e.class)) : data.exams;
  const classes = [...new Set(data.learners.map((l) => l.class).filter(Boolean))];
  const activeExam = visibleExams.find((e) => e.id === activeExamId) || null;
  const gridSubjects = activeExam ? (isTeacher ? activeExam.subjects.filter((s) => mySubjects.includes(s)) : activeExam.subjects) : [];

  const statuses = data.examSubjectStatus || [];
  const getStatus = (examId, subject) => (statuses.find((s) => s.examId === examId && s.subject === subject) || {}).status || "draft";
  const lockedForTeacher = activeExam ? gridSubjects.filter((s) => getStatus(activeExam.id, s) !== "draft") : [];
  const setStatus = (examId, subject, status) => {
    const exists = statuses.some((s) => s.examId === examId && s.subject === subject);
    const now = new Date().toISOString();
    const next = exists
      ? statuses.map((s) => (s.examId === examId && s.subject === subject
          ? { ...s, status, ...(status === "submitted" ? { submittedAt: now, submittedBy: user.name } : {}), ...(status === "published" ? { publishedAt: now } : {}) }
          : s))
      : [...statuses, { id: uid(), examId, subject, status, submittedAt: status === "submitted" ? now : null, submittedBy: status === "submitted" ? user.name : null, publishedAt: status === "published" ? now : null }];
    persist({ ...data, examSubjectStatus: next });
    notify(status === "submitted" ? `Submitted ${subject} for review` : status === "published" ? `Published ${subject} results` : `Reopened ${subject}`);
  };

  const saveExam = (exam) => {
    const exists = data.exams.some((e) => e.id === exam.id);
    const next = exists
      ? { ...data, exams: data.exams.map((e) => (e.id === exam.id ? exam : e)) }
      : { ...data, exams: [...data.exams, { ...exam, id: uid() }] };
    persist(next);
    setActiveExamId(exists ? exam.id : next.exams[next.exams.length - 1].id);
    notify(exists ? "Exam updated" : "Exam created");
    setExamModal(null);
  };

  const removeExam = (id) => {
    const exam = data.exams.find((e) => e.id === id);
    trashItem(data, persist, "exam", exam, { exams: data.exams.filter((e) => e.id !== id) });
    if (activeExamId === id) setActiveExamId(null);
    notify("Exam moved to trash — restore it from School Profile within 30 days");
  };

  const upsertMarks = (newMarks) => {
    let marks = [...data.marks];
    newMarks.forEach(({ learnerId, subject, score }) => {
      const idx = marks.findIndex((m) => m.examId === activeExamId && m.learnerId === learnerId && m.subject === subject);
      if (idx >= 0) marks[idx] = { ...marks[idx], score };
      else marks.push({ id: uid(), examId: activeExamId, learnerId, subject, score });
    });
    persist({ ...data, marks });
  };

  const handleFile = (e) => {
    const file = e.target.files[0];
    if (!file || !activeExam) return;
    setFileError("");
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        if (!rows.length) { setFileError("The file appears to be empty."); return; }
        const headers = Object.keys(rows[0]);
        const idHeader = headers.find((h) => /admission/i.test(h)) || headers.find((h) => /name/i.test(h));
        if (!idHeader) { setFileError("Couldn't find a column for Admission No. or Name to match learners."); return; }
        const allowedSubjects = isTeacher ? gridSubjects : activeExam.subjects;
        const subjectHeaders = allowedSubjects.length
          ? allowedSubjects.filter((s) => headers.includes(s))
          : (isTeacher ? [] : headers.filter((h) => h !== idHeader && !/name|class|admission/i.test(h)));

        const matched = [];
        const unmatched = [];
        rows.forEach((row) => {
          const key = String(row[idHeader]).trim().toLowerCase();
          const learner = data.learners.find(
            (l) => l.admissionNo.trim().toLowerCase() === key || l.name.trim().toLowerCase() === key
          );
          if (!learner) { unmatched.push(row[idHeader]); return; }
          subjectHeaders.forEach((subj) => {
            const val = row[subj];
            if (val !== "" && val !== undefined) matched.push({ learnerId: learner.id, subject: subj, score: Number(val) });
          });
        });
        setPreview({ matchedCount: new Set(matched.map((m) => m.learnerId)).size, matched, unmatched, subjectHeaders, rowCount: rows.length });
      } catch (err) {
        setFileError("Couldn't read that file. Please upload a .xlsx, .xls, or .csv file exported from Excel.");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const confirmImport = () => {
    if (!preview) return;
    const subjSet = new Set([...(activeExam.subjects || []), ...preview.subjectHeaders]);
    const nextExam = { ...activeExam, subjects: [...subjSet] };
    persist({
      ...data,
      exams: data.exams.map((e) => (e.id === activeExam.id ? nextExam : e)),
      marks: (() => {
        let marks = [...data.marks];
        preview.matched.forEach(({ learnerId, subject, score }) => {
          const idx = marks.findIndex((m) => m.examId === activeExamId && m.learnerId === learnerId && m.subject === subject);
          if (idx >= 0) marks[idx] = { ...marks[idx], score };
          else marks.push({ id: uid(), examId: activeExamId, learnerId, subject, score });
        });
        return marks;
      })(),
    });
    notify(`Imported marks for ${preview.matchedCount} learner(s)`);
    setPreview(null);
    if (fileRef.current) fileRef.current.value = "";
  };

  return (
    <div>
      <div className="flex items-center justify-between mb-3 gap-2">
        <h2 className="font-bold text-slate-800" style={serifStyle()}>{isTeacher ? "Your exams" : "Manage exams"}</h2>
        {isAdmin && (
          <div className="flex gap-2 shrink-0">
            <Btn size="sm" onClick={() => setExamModal({ id: "", name: "", class: "", subjects: [], date: "", attachment: null, subjectMaxMarks: {}, term: "Term 1", year: String(new Date().getFullYear()) })}><Plus size={14} /> Create exam</Btn>
          </div>
        )}
      </div>

      {isTeacher && !myTeacherProfile && (
        <div className="bg-amber-50 border border-amber-200 rounded-xl p-4 text-sm text-amber-800 mb-4">
          Your account isn't linked to a teacher profile yet, so no exams or subjects are showing. Ask your admin to register your teacher account from the Teachers tab, or link your existing profile.
        </div>
      )}

      {visibleExams.length === 0 ? (
        <EmptyState icon={ClipboardList} title="No exams yet" hint={isTeacher ? "Exams for your assigned class will appear here once your admin creates them." : "Create your first exam — give it a name, class, and subjects — then record or upload marks."} />
      ) : (
        <div className="grid sm:grid-cols-2 gap-3 mb-6">
          {visibleExams.map((e) => (
            <div key={e.id} onClick={() => setActiveExamId(e.id)} className={`text-left bg-white rounded-xl border p-4 cursor-pointer transition-shadow hover:shadow-md ${activeExamId === e.id ? "border-emerald-700 ring-1 ring-emerald-700" : "border-slate-200"}`}>
              <div className="flex items-start justify-between mb-1.5">
                <p className="font-semibold text-slate-800">{e.name}</p>
                {isAdmin && (
                  <div className="flex gap-1 shrink-0">
                    <button onClick={(ev) => { ev.stopPropagation(); setExamModal(e); }} className="p-1 text-slate-400 hover:text-slate-700"><Edit2 size={13} /></button>
                    <button onClick={(ev) => { ev.stopPropagation(); removeExam(e.id); }} className="p-1 text-slate-400 hover:text-red-600"><Trash2 size={13} /></button>
                  </div>
                )}
              </div>
              <p className="text-xs text-slate-500 mb-2">{e.class || "No class"} · {e.term ? `${e.term} ${e.year}` : (e.date || "No date")}</p>
              <div className="flex flex-wrap gap-1">
                {e.subjects.slice(0, 4).map((s) => <span key={s} className="text-[10px] bg-stone-100 text-slate-600 px-1.5 py-0.5 rounded">{s}</span>)}
                {e.subjects.length > 4 && <span className="text-[10px] text-slate-400">+{e.subjects.length - 4} more</span>}
                {e.subjects.length === 0 && <span className="text-[10px] text-amber-600">No subjects set</span>}
              </div>
              {e.attachment && (
                <a href={e.attachment.dataUrl} download={e.attachment.name} onClick={(ev) => ev.stopPropagation()} className="text-[11px] text-emerald-800 font-medium flex items-center gap-1 mt-2">
                  <FileText size={11} /> {e.attachment.name}
                </a>
              )}
              <Btn
                size="sm"
                variant={activeExamId === e.id ? "primary" : "secondary"}
                className="w-full justify-center mt-3"
                onClick={(ev) => { ev.stopPropagation(); setActiveExamId(e.id); }}
              >
                <CheckCheck size={13} /> Manage marks & publish
              </Btn>
            </div>
          ))}
        </div>
      )}

      {activeExam && (
        <>
          {activeExam.generated && <GeneratedExamView exam={activeExam} />}

          {(!isTeacher || gridSubjects.length > 0) && (
            <div className="bg-white rounded-xl border border-slate-200 p-4 mb-4">
              <h3 className="font-semibold text-sm text-slate-700 mb-2 flex items-center gap-1.5"><Upload size={14} /> Upload marks for "{activeExam.name}" (Excel / CSV)</h3>
              <p className="text-xs text-slate-500 mb-3">First column should be <b>Admission No.</b> or <b>Name</b>, with one column per subject{isTeacher ? " (only your subject will be imported)" : ""}. PDF/Word mark sheets aren't parsed automatically — convert to Excel/CSV first, or use manual entry below.</p>
              <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} className="hidden" />
              <button
                type="button"
                onClick={() => fileRef.current && fileRef.current.click()}
                className="w-full flex items-center justify-center gap-2 border-2 border-dashed border-slate-300 rounded-lg py-3 text-sm text-slate-600 hover:border-emerald-600 hover:text-emerald-700"
              >
                <Upload size={15} /> Tap to choose a file
              </button>
              {fileError && <p className="text-xs text-red-600 mt-2 flex items-center gap-1"><AlertCircle size={13} /> {fileError}</p>}

              {preview && (
                <div className="mt-3 bg-stone-50 border border-slate-200 rounded-lg p-3 text-xs">
                  <p className="text-slate-700 mb-1"><b>{preview.matchedCount}</b> of {preview.rowCount} rows matched to learner profiles. Subjects found: {preview.subjectHeaders.join(", ") || "none"}.</p>
                  {preview.unmatched.length > 0 && <p className="text-amber-700 mb-2">Unmatched: {preview.unmatched.slice(0, 8).join(", ")}{preview.unmatched.length > 8 ? "…" : ""}</p>}
                  <div className="flex gap-2">
                    <Btn size="sm" onClick={confirmImport} disabled={preview.matchedCount === 0}><CheckCircle2 size={13} /> Import {preview.matchedCount} learner(s)</Btn>
                    <Btn size="sm" variant="secondary" onClick={() => { setPreview(null); if (fileRef.current) fileRef.current.value = ""; }}>Cancel</Btn>
                  </div>
                </div>
              )}
            </div>
          )}

          <ManualMarksGrid data={data} exam={activeExam} subjects={gridSubjects} upsertMarks={upsertMarks} notify={notify} persist={persist} readOnlySubjects={isAdmin ? [] : lockedForTeacher} isAdmin={isAdmin} />

          <div className="bg-white rounded-xl border border-slate-200 p-4 mt-4">
            <h3 className="font-semibold text-sm text-slate-700 mb-3 flex items-center gap-1.5"><CheckCheck size={14} /> {isAdmin ? "Review & publish" : "Submit for review"}</h3>
            <div className="space-y-2">
              {(isTeacher ? gridSubjects : activeExam.subjects).map((s) => {
                const status = getStatus(activeExam.id, s);
                const badge = status === "published" ? "bg-emerald-100 text-emerald-800" : status === "submitted" ? "bg-amber-100 text-amber-800" : "bg-slate-100 text-slate-600";
                return (
                  <div key={s} className="flex items-center justify-between gap-2 py-1.5 border-b border-slate-100 last:border-0">
                    <div>
                      <p className="text-sm text-slate-800">{s}</p>
                      <span className={`text-[10px] font-semibold uppercase tracking-wide px-2 py-0.5 rounded-full ${badge}`}>{status}</span>
                    </div>
                    {isTeacher && status === "draft" && <Btn size="sm" onClick={() => setStatus(activeExam.id, s, "submitted")}><Send size={13} /> Submit</Btn>}
                    {isTeacher && status === "submitted" && <span className="text-xs text-slate-400">Awaiting admin review</span>}
                    {isTeacher && status === "published" && <span className="text-xs text-emerald-700 flex items-center gap-1"><CheckCheck size={13} /> Published</span>}
                    {isAdmin && status !== "published" && <Btn size="sm" onClick={() => setStatus(activeExam.id, s, "published")}><CheckCheck size={13} /> Publish</Btn>}
                    {isAdmin && status === "published" && <Btn size="sm" variant="secondary" onClick={() => setStatus(activeExam.id, s, "submitted")}><Lock size={13} /> Reopen</Btn>}
                  </div>
                );
              })}
            </div>
            {isAdmin && <p className="text-[11px] text-slate-400 mt-2">Publishing marks the subject as finalized. Use Merit List and Report Cards to analyze results before publishing.</p>}
          </div>
        </>
      )}

      {examModal && (
        <Modal title={examModal.id ? "Edit exam" : "Create exam"} onClose={() => setExamModal(null)} wide>
          <ExamFormBody exam={examModal} classes={classes} onSave={saveExam} onClose={() => setExamModal(null)} />
        </Modal>
      )}


    </div>
  );
}

// ---------- Display for an auto-generated exam paper + marking scheme ----------
function GeneratedExamView({ exam }) {
  const [view, setView] = useState("paper"); // paper | scheme
  const g = exam.generated;

  const copyText = () => {
    const partA = view === "paper"
      ? g.sectionA.questions.map((q) => `${q.number}. ${q.question}\n${(q.options || []).join("\n")} (${q.marks} marks)`).join("\n\n")
      : g.markingScheme.sectionA.map((m) => `${m.number}. Correct: ${m.correctOption}${m.explanation ? " — " + m.explanation : ""} (${m.marks} marks)`).join("\n\n");
    const partB = view === "paper"
      ? g.sectionB.questions.map((q) => `${q.number}. ${q.question} (${q.marks} marks)`).join("\n\n")
      : g.markingScheme.sectionB.map((m) => `${m.number}. ${m.expectedAnswer} (${m.marks} marks)`).join("\n\n");

    const text =
      `Section A — Structured, Multiple Choice (out of ${g.sectionA.outOf})\n\n${partA}\n\n---\n\n` +
      `Section B — Unstructured (out of ${g.sectionB.outOf})\n\n${partB}`;
    navigator.clipboard.writeText(`${exam.name} — ${g.subject}\n\n${text}`);
  };

  const printId = `gen-exam-${exam.id}`;

  return (
    <div className="bg-white rounded-xl border border-slate-200 p-4 mb-4">
      <div className="flex items-center justify-between mb-2 print:hidden">
        <h3 className="font-semibold text-sm text-slate-700 flex items-center gap-1.5"><Sparkles size={14} className="text-amber-500" /> AI-generated {view === "paper" ? "exam paper" : "marking scheme"}</h3>
        <div className="flex gap-2">
          <Btn size="sm" variant="secondary" onClick={copyText}><Copy size={13} /> Copy</Btn>
          <Btn size="sm" variant="secondary" onClick={() => window.print()}><Printer size={13} /> Download PDF</Btn>
        </div>
      </div>
      <div className="flex gap-2 mb-3 print:hidden">
        <button onClick={() => setView("paper")} className={`px-3 py-1.5 rounded-full text-xs font-medium border ${view === "paper" ? "bg-emerald-800 text-white border-emerald-800" : "bg-white text-slate-600 border-slate-300"}`}>Exam paper</button>
        <button onClick={() => setView("scheme")} className={`px-3 py-1.5 rounded-full text-xs font-medium border ${view === "scheme" ? "bg-emerald-800 text-white border-emerald-800" : "bg-white text-slate-600 border-slate-300"}`}>Marking scheme</button>
      </div>

      <div id={printId}>
        <p className="font-bold text-slate-800 mb-1" style={serifStyle()}>{exam.name} — {g.subject}</p>
        <p className="text-xs text-slate-400 mb-3">Topics: {g.topics} · {view === "paper" ? "Exam Paper" : "Marking Scheme"}</p>

        {/* Section A — Multiple Choice */}
        <div className="mb-4">
          <p className="text-sm font-semibold text-slate-700 border-b border-slate-200 pb-1 mb-2">Section A — Structured, Multiple Choice (out of {g.sectionA.outOf})</p>
          {g.sectionA.instructions && view === "paper" && <p className="text-xs text-slate-500 italic mb-2">{g.sectionA.instructions}</p>}
          <div className="space-y-3 max-h-72 overflow-y-auto border border-slate-100 rounded-lg p-3 print:max-h-none print:overflow-visible print:border-0 print:p-0">
            {view === "paper"
              ? g.sectionA.questions.map((q) => (
                  <div key={q.number} className="text-sm">
                    <p className="text-slate-800"><b>{q.number}.</b> {q.question}</p>
                    <div className="pl-4 text-slate-600">
                      {(q.options || []).map((opt) => <p key={opt}>{opt}</p>)}
                    </div>
                    <p className="text-xs text-slate-400">{q.marks} marks</p>
                  </div>
                ))
              : g.markingScheme.sectionA.map((m) => (
                  <div key={m.number} className="text-sm">
                    <p className="text-slate-800"><b>{m.number}.</b> Correct answer: <span className="font-bold text-emerald-700">{m.correctOption}</span></p>
                    {m.explanation && <p className="text-slate-500 text-xs pl-4">{m.explanation}</p>}
                    <p className="text-xs text-slate-400">{m.marks} marks</p>
                  </div>
                ))}
          </div>
        </div>

        {/* Section B — Unstructured */}
        <div className="mb-4">
          <p className="text-sm font-semibold text-slate-700 border-b border-slate-200 pb-1 mb-2">Section B — Unstructured (out of {g.sectionB.outOf})</p>
          {g.sectionB.instructions && view === "paper" && <p className="text-xs text-slate-500 italic mb-2">{g.sectionB.instructions}</p>}
          <div className="space-y-2 max-h-72 overflow-y-auto border border-slate-100 rounded-lg p-3 print:max-h-none print:overflow-visible print:border-0 print:p-0">
            {view === "paper"
              ? g.sectionB.questions.map((q) => (
                  <div key={q.number} className="text-sm">
                    <p className="text-slate-800"><b>{q.number}.</b> {q.question}</p>
                    <p className="text-xs text-slate-400">{q.marks} marks</p>
                  </div>
                ))
              : g.markingScheme.sectionB.map((m) => (
                  <div key={m.number} className="text-sm">
                    <p className="text-slate-800"><b>{m.number}.</b> {m.expectedAnswer}</p>
                    <p className="text-xs text-slate-400">{m.marks} marks</p>
                  </div>
                ))}
          </div>
        </div>
      </div>
      <p className="text-[11px] text-slate-400 mt-2 print:hidden">Review before handing out — AI-generated questions should be checked for accuracy and curriculum fit. "Download PDF" opens your browser's print dialog — choose "Save as PDF" as the destination.</p>
    </div>
  );
}

// ---------- Exam form ----------
function ExamFormBody({ exam, classes, onSave, onClose }) {
  const [f, setF] = useState({ ...exam, subjects: exam.subjects || [], subjectMaxMarks: exam.subjectMaxMarks || {}, term: exam.term || "Term 1", year: exam.year || String(new Date().getFullYear()) });
  const [customSubj, setCustomSubj] = useState("");

  const toggleSubject = (s) => {
    setF((prev) => {
      const has = prev.subjects.includes(s);
      const nextMax = { ...prev.subjectMaxMarks };
      if (has) delete nextMax[s]; else nextMax[s] = nextMax[s] || 100;
      return {
        ...prev,
        subjects: has ? prev.subjects.filter((x) => x !== s) : [...prev.subjects, s],
        subjectMaxMarks: nextMax,
      };
    });
  };
  const addCustom = () => {
    const v = customSubj.trim();
    if (v && !f.subjects.includes(v)) setF({ ...f, subjects: [...f.subjects, v], subjectMaxMarks: { ...f.subjectMaxMarks, [v]: 100 } });
    setCustomSubj("");
  };

  return (
    <>
      <Field label="Exam name"><input className={inputCls} value={f.name} onChange={(e) => setF({ ...f, name: e.target.value })} placeholder="e.g. Term 2 Midterm" /></Field>
      <div className="grid sm:grid-cols-2 gap-x-4">
        <Field label="Class">
          <input list="classopts" className={inputCls} value={f.class} onChange={(e) => setF({ ...f, class: e.target.value })} placeholder="e.g. Grade 6B" />
          <datalist id="classopts">{classes.map((c) => <option key={c} value={c} />)}</datalist>
        </Field>
        <Field label="Date"><input type="date" className={inputCls} value={f.date} onChange={(e) => setF({ ...f, date: e.target.value })} /></Field>
      </div>
      <div className="grid sm:grid-cols-2 gap-x-4">
        <Field label="Term">
          <select className={inputCls} value={f.term} onChange={(e) => setF({ ...f, term: e.target.value })}>
            <option>Term 1</option><option>Term 2</option><option>Term 3</option>
          </select>
        </Field>
        <Field label="Year"><input className={inputCls} value={f.year} onChange={(e) => setF({ ...f, year: e.target.value })} /></Field>
      </div>

      <span className="block text-xs font-semibold text-slate-500 uppercase tracking-wide mb-1.5">Subjects examined</span>
      <div className="grid grid-cols-2 sm:grid-cols-3 gap-1.5 max-h-48 overflow-y-auto border border-slate-200 rounded-lg p-2.5 mb-2">
        {ALL_SUBJECTS.map((s) => (
          <label key={s} className="flex items-center gap-1.5 text-xs text-slate-700 cursor-pointer">
            <input type="checkbox" checked={f.subjects.includes(s)} onChange={() => toggleSubject(s)} className="accent-emerald-700" />
            {s}
          </label>
        ))}
      </div>
      <div className="flex gap-2 mb-3">
        <input className={inputCls} value={customSubj} onChange={(e) => setCustomSubj(e.target.value)} placeholder="Add another subject…" onKeyDown={(e) => e.key === "Enter" && (e.preventDefault(), addCustom())} />
        <Btn variant="secondary" size="sm" onClick={addCustom}><Plus size={13} /></Btn>
      </div>
      {f.subjects.length > 0 && (
        <div className="flex flex-wrap gap-1.5 mb-4">
          {f.subjects.map((s) => (
            <span key={s} className="flex items-center gap-1 text-xs bg-emerald-50 text-emerald-800 px-2 py-1 rounded-full">
              {s} <button onClick={() => toggleSubject(s)}><X size={11} /></button>
            </span>
          ))}
        </div>
      )}
      <p className="text-[11px] text-slate-400 -mt-2 mb-4">The subject teacher sets "out of" (max marks) themselves when entering marks — defaults to 100 until they change it.</p>

      <Field label="Exam paper document (optional)">
        <FileAttachment attachment={f.attachment} onUpload={(a) => setF({ ...f, attachment: a })} onRemove={() => setF({ ...f, attachment: null })} />
      </Field>

      <div className="flex justify-end gap-2 mt-2">
        <Btn variant="secondary" onClick={onClose}>Cancel</Btn>
        <Btn onClick={() => f.name.trim() && onSave(f)} disabled={!f.name.trim()}><Save size={14} /> Save exam</Btn>
      </div>
    </>
  );
}

function ManualMarksGrid({ data, exam, subjects, upsertMarks, notify, persist, readOnlySubjects = [], isAdmin = false }) {
  const learners = data.learners.filter((l) => l.class === exam.class);
  const [grid, setGrid] = useState({});
  const maxMarks = exam.subjectMaxMarks || {};

  const getVal = (learnerId, subject) => {
    if (grid[learnerId]?.[subject] !== undefined) return grid[learnerId][subject];
    const existing = data.marks.find((m) => m.examId === exam.id && m.learnerId === learnerId && m.subject === subject);
    return existing ? existing.score : "";
  };
  const setVal = (learnerId, subject, val) => {
    const max = maxMarks[subject] || 100;
    setGrid((g) => ({ ...g, [learnerId]: { ...g[learnerId], [subject]: clamp(val, 0, max) } }));
  };

  const setOutOf = (subject, val) => {
    const nextMax = { ...maxMarks, [subject]: val === "" ? "" : Number(val) };
    const nextExam = { ...exam, subjectMaxMarks: nextMax };
    persist({ ...data, exams: data.exams.map((e) => (e.id === exam.id ? nextExam : e)) });
  };

  const saveAll = () => {
    const updates = [];
    Object.entries(grid).forEach(([learnerId, subjects]) => {
      Object.entries(subjects).forEach(([subject, score]) => {
        if (score !== "" && !readOnlySubjects.includes(subject)) updates.push({ learnerId, subject, score: Number(score) });
      });
    });
    if (updates.length) { upsertMarks(updates); notify(`Saved marks for ${new Set(updates.map(u => u.learnerId)).size} learner(s)`); setGrid({}); }
  };

  if (!subjects.length) return <div className="text-sm text-slate-500 bg-white rounded-xl border border-slate-200 p-4">No subjects available for you to enter marks for on this exam. Ask your admin to assign your teaching subject.</div>;
  if (!learners.length) return <div className="text-sm text-slate-500 bg-white rounded-xl border border-slate-200 p-4">No learners are in class "{exam.class}" yet.</div>;

  return (
    <div className="bg-white rounded-xl border border-slate-200 p-4 overflow-x-auto">
      <h3 className="font-semibold text-sm text-slate-700 mb-1">Manual mark entry — {exam.class}</h3>
      <p className="text-xs text-slate-500 mb-3">
        Set "out of" for your subject below — marks are converted to a percentage and grade automatically.
        {isAdmin && " As admin, you can edit any subject at any time, even after it's been submitted or published."}
        {!isAdmin && readOnlySubjects.length > 0 && " A subject locks once you submit it — ask your admin to reopen it if you spot a mistake."}
      </p>
      <div className="flex flex-wrap gap-3 mb-3">
        {subjects.map((s) => (
          <label key={s} className="flex items-center gap-1.5 text-xs text-slate-600 bg-stone-50 rounded-md px-2.5 py-1.5">
            <span className="font-medium text-slate-700">{s}</span> out of
            <input
              type="number"
              className="w-14 rounded border border-slate-300 px-1.5 py-1 text-center text-xs"
              value={maxMarks[s] ?? 100}
              onChange={(e) => setOutOf(s, e.target.value)}
            />
          </label>
        ))}
      </div>
      <table className="text-sm min-w-full">
        <thead>
          <tr className="text-left text-xs text-slate-400 uppercase tracking-wide">
            <th className="pb-2 pr-3">Learner</th>
            {subjects.map((s) => (
              <th key={s} className="pb-2 px-1 text-center">
                {s}
                {readOnlySubjects.includes(s) && <span className="block normal-case font-normal text-amber-600 flex items-center justify-center gap-0.5"><Lock size={9} /> Locked</span>}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {learners.map((l) => (
            <tr key={l.id} className="border-t border-slate-100">
              <td className="py-1.5 pr-3 whitespace-nowrap">{l.name}</td>
              {subjects.map((s) => {
                const val = getVal(l.id, s);
                const pct = val !== "" ? toPercent(val, maxMarks[s] || 100) : null;
                const locked = readOnlySubjects.includes(s);
                return (
                  <td key={s} className="py-1.5 px-1">
                    <div className="flex flex-col items-center">
                      <input
                        type="number"
                        min="0"
                        max={maxMarks[s] || 100}
                        className={`w-16 rounded border px-1.5 py-1 text-center text-sm ${locked ? "border-slate-200 bg-slate-50 text-slate-400" : "border-slate-300"}`}
                        value={val}
                        onChange={(e) => setVal(l.id, s, e.target.value)}
                        disabled={locked}
                      />
                      {pct !== null && (
                        <span className="text-[10px] text-slate-400 mt-0.5">{pct.toFixed(0)}% · {gradeForPercent(getSystemForSubject(data, s), pct)}</span>
                      )}
                    </div>
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
      <Btn size="sm" className="mt-3" onClick={saveAll}><Save size={13} /> Save marks</Btn>
    </div>
  );
}

// ---------- Schemes of Work ----------
function blankScheme() { return { id: "", subject: "", class: "", term: "Term 1", year: String(new Date().getFullYear()), weeks: [], attachment: null }; }
function blankWeek() { return { id: "", week: "", lesson: "", topic: "", subtopic: "", objectives: "", activities: "", materials: "", references: "", remarks: "" }; }

function Schemes({ data, persist, notify }) {
  const schemes = data.schemes || [];
  const [modal, setModal] = useState(null);
  const [activeId, setActiveId] = useState(null);
  const [weekModal, setWeekModal] = useState(null); // { week }

  const active = schemes.find((s) => s.id === activeId);

  const saveScheme = (s) => {
    const exists = schemes.some((x) => x.id === s.id);
    const next = exists
      ? { ...data, schemes: schemes.map((x) => (x.id === s.id ? s : x)) }
      : { ...data, schemes: [...schemes, { ...s, id: uid(), weeks: [] }] };
    persist(next);
    notify(exists ? "Scheme updated" : "Scheme created");
    if (!exists) setActiveId(next.schemes[next.schemes.length - 1].id);
    setModal(null);
  };
  const removeScheme = (id) => {
    persist({ ...data, schemes: schemes.filter((s) => s.id !== id) });
    notify("Scheme deleted");
    if (activeId === id) setActiveId(null);
  };

  const saveWeek = (week) => {
    const exists = active.weeks.some((w) => w.id === week.id);
    const weeks = exists ? active.weeks.map((w) => (w.id === week.id ? week : w)) : [...active.weeks, { ...week, id: uid() }];
    persist({ ...data, schemes: schemes.map((s) => (s.id === active.id ? { ...s, weeks } : s)) });
    notify("Week saved");
    setWeekModal(null);
  };
  const removeWeek = (weekId) => {
    persist({ ...data, schemes: schemes.map((s) => (s.id === active.id ? { ...s, weeks: s.weeks.filter((w) => w.id !== weekId) } : s)) });
    notify("Week removed");
  };

  if (active) {
    return (
      <div>
        <button onClick={() => setActiveId(null)} className="text-sm text-emerald-800 font-medium mb-3 flex items-center gap-1"><ChevronRight size={14} className="rotate-180" /> Back to schemes</button>
        <div className="bg-white rounded-xl border border-slate-200 p-4 mb-4 flex items-start justify-between">
          <div>
            <h2 className="font-bold text-slate-800" style={serifStyle()}>{active.subject} — {active.class}</h2>
            <p className="text-xs text-slate-500">{active.term} · {active.year} · {active.weeks.length} week(s) planned</p>
            {active.attachment && (
              <a href={active.attachment.dataUrl} download={active.attachment.name} className="text-xs text-emerald-800 font-medium flex items-center gap-1 mt-1">
                <FileText size={12} /> {active.attachment.name} · Download
              </a>
            )}
          </div>
          <div className="flex gap-1.5 shrink-0">
            <button onClick={() => setModal(active)} className="p-1.5 text-slate-400 hover:text-slate-700"><Edit2 size={15} /></button>
            <button onClick={() => removeScheme(active.id)} className="p-1.5 text-slate-400 hover:text-red-600"><Trash2 size={15} /></button>
          </div>
        </div>

        <div className="flex justify-end mb-3">
          <Btn size="sm" onClick={() => setWeekModal({ week: blankWeek() })}><Plus size={14} /> Add week</Btn>
        </div>

        {active.weeks.length === 0 ? (
          <EmptyState icon={NotebookPen} title="No weeks planned yet" hint="Add your first week to start building this scheme of work." />
        ) : (
          <div className="space-y-3">
            {active.weeks.map((w, i) => (
              <div key={w.id} className="bg-white rounded-xl border border-slate-200 p-4">
                <div className="flex items-start justify-between mb-1.5">
                  <p className="font-semibold text-sm text-slate-800">Week {w.week || i + 1}{w.lesson ? ` · Lesson ${w.lesson}` : ""}</p>
                  <div className="flex gap-1 shrink-0">
                    <button onClick={() => setWeekModal({ week: w })} className="p-1 text-slate-400 hover:text-slate-700"><Edit2 size={13} /></button>
                    <button onClick={() => removeWeek(w.id)} className="p-1 text-slate-400 hover:text-red-600"><Trash2 size={13} /></button>
                  </div>
                </div>
                <p className="text-sm text-slate-700 font-medium mb-1">{w.topic}{w.subtopic ? ` — ${w.subtopic}` : ""}</p>
                {w.objectives && <p className="text-xs text-slate-500 mb-1"><b>Objectives:</b> {w.objectives}</p>}
                {w.activities && <p className="text-xs text-slate-500 mb-1"><b>Activities:</b> {w.activities}</p>}
                {w.materials && <p className="text-xs text-slate-500 mb-1"><b>Materials:</b> {w.materials}</p>}
                {w.references && <p className="text-xs text-slate-500 mb-1"><b>References:</b> {w.references}</p>}
                {w.remarks && <p className="text-xs text-slate-500"><b>Remarks:</b> {w.remarks}</p>}
              </div>
            ))}
          </div>
        )}

        {weekModal && (
          <Modal title={active.weeks.some((w) => w.id === weekModal.week.id) ? "Edit week" : "Add week"} onClose={() => setWeekModal(null)} wide>
            <WeekFormBody week={weekModal.week} onSave={saveWeek} onClose={() => setWeekModal(null)} />
          </Modal>
        )}
        {modal && (
          <Modal title="Edit scheme" onClose={() => setModal(null)}>
            <SchemeFormBody scheme={modal} onSave={saveScheme} onClose={() => setModal(null)} />
          </Modal>
        )}
      </div>
    );
  }

  return (
    <div>
      <div className="flex justify-end mb-4">
        <Btn onClick={() => setModal(blankScheme())}><Plus size={15} /> Create scheme</Btn>
      </div>
      {schemes.length === 0 ? (
        <EmptyState icon={NotebookPen} title="No schemes of work yet" hint="Create one per subject and class to plan out a term, week by week." />
      ) : (
        <div className="grid sm:grid-cols-2 gap-3">
          {schemes.map((s) => (
            <button key={s.id} onClick={() => setActiveId(s.id)} className="text-left bg-white rounded-xl border border-slate-200 p-4 hover:shadow-md transition-shadow">
              <p className="font-semibold text-slate-800">{s.subject || "Untitled subject"}</p>
              <p className="text-xs text-slate-500">{s.class || "No class"} · {s.term} {s.year}</p>
              <p className="text-xs text-emerald-700 mt-1">{s.weeks.length} week(s)</p>
            </button>
          ))}
        </div>
      )}
      {modal && (
        <Modal title={modal.id ? "Edit scheme" : "Create scheme of work"} onClose={() => setModal(null)}>
          <SchemeFormBody scheme={modal} onSave={saveScheme} onClose={() => setModal(null)} />
        </Modal>
      )}
    </div>
  );
}

function SchemeFormBody({ scheme, onSave, onClose }) {
  const [f, setF] = useState(scheme);
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  return (
    <>
      <Field label="Subject">
        <select className={inputCls} value={f.subject} onChange={set("subject")}>
          <option value="">Select a subject…</option>
          {ALL_SUBJECTS.map((s) => <option key={s} value={s}>{s}</option>)}
        </select>
      </Field>
      <Field label="Class"><input className={inputCls} value={f.class} onChange={set("class")} placeholder="e.g. Grade 6B" /></Field>
      <div className="grid grid-cols-2 gap-x-4">
        <Field label="Term">
          <select className={inputCls} value={f.term} onChange={set("term")}>
            <option>Term 1</option><option>Term 2</option><option>Term 3</option>
          </select>
        </Field>
        <Field label="Year"><input className={inputCls} value={f.year} onChange={set("year")} /></Field>
      </div>
      <Field label="Scheme document (optional)">
        <FileAttachment attachment={f.attachment} onUpload={(a) => setF({ ...f, attachment: a })} onRemove={() => setF({ ...f, attachment: null })} />
      </Field>
      <div className="flex justify-end gap-2 mt-2">
        <Btn variant="secondary" onClick={onClose}>Cancel</Btn>
        <Btn onClick={() => f.subject && f.class.trim() && onSave(f)} disabled={!f.subject || !f.class.trim()}><Save size={14} /> Save</Btn>
      </div>
    </>
  );
}

function WeekFormBody({ week, onSave, onClose }) {
  const [f, setF] = useState(week);
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  return (
    <>
      <div className="grid grid-cols-2 gap-x-4">
        <Field label="Week"><input className={inputCls} value={f.week} onChange={set("week")} placeholder="e.g. 1" /></Field>
        <Field label="Lesson"><input className={inputCls} value={f.lesson} onChange={set("lesson")} placeholder="e.g. 1 of 3" /></Field>
      </div>
      <Field label="Topic"><input className={inputCls} value={f.topic} onChange={set("topic")} /></Field>
      <Field label="Sub-topic"><input className={inputCls} value={f.subtopic} onChange={set("subtopic")} /></Field>
      <Field label="Specific objectives"><textarea className={inputCls} rows={2} value={f.objectives} onChange={set("objectives")} placeholder="By the end of the lesson, the learner should be able to…" /></Field>
      <Field label="Teaching/learning activities"><textarea className={inputCls} rows={2} value={f.activities} onChange={set("activities")} /></Field>
      <Field label="Materials/resources"><input className={inputCls} value={f.materials} onChange={set("materials")} /></Field>
      <Field label="References"><input className={inputCls} value={f.references} onChange={set("references")} /></Field>
      <Field label="Remarks"><input className={inputCls} value={f.remarks} onChange={set("remarks")} /></Field>
      <div className="flex justify-end gap-2 mt-2">
        <Btn variant="secondary" onClick={onClose}>Cancel</Btn>
        <Btn onClick={() => f.topic.trim() && onSave(f)} disabled={!f.topic.trim()}><Save size={14} /> Save week</Btn>
      </div>
    </>
  );
}

// ---------- Lesson Notes ----------
function blankLessonNote() {
  return { id: "", subject: "", class: "", date: "", topic: "", subtopic: "", objectives: "", introduction: "", development: "", conclusion: "", assignment: "", references: "" };
}

function LessonNotes({ data, persist, notify }) {
  const notes = data.lessonNotes || [];
  const [modal, setModal] = useState(null);
  const [q, setQ] = useState("");

  const filtered = notes.filter((n) => `${n.subject} ${n.class} ${n.topic}`.toLowerCase().includes(q.toLowerCase()));

  const save = (n) => {
    const exists = notes.some((x) => x.id === n.id);
    const next = exists
      ? { ...data, lessonNotes: notes.map((x) => (x.id === n.id ? n : x)) }
      : { ...data, lessonNotes: [...notes, { ...n, id: uid() }] };
    persist(next);
    notify(exists ? "Lesson note updated" : "Lesson note created");
    setModal(null);
  };
  const remove = (id) => { persist({ ...data, lessonNotes: notes.filter((n) => n.id !== id) }); notify("Lesson note removed"); };

  return (
    <div>
      <div className="flex items-center gap-2 mb-4">
        <div className="relative flex-1">
          <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
          <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Search by subject, class, topic…" className={`${inputCls} pl-9`} />
        </div>
        <Btn onClick={() => setModal(blankLessonNote())}><Plus size={15} /> New note</Btn>
      </div>

      {filtered.length === 0 ? (
        <EmptyState icon={FileText} title="No lesson notes yet" hint="Create a lesson note for any subject and class." />
      ) : (
        <div className="space-y-3">
          {filtered.map((n) => (
            <div key={n.id} className="bg-white rounded-xl border border-slate-200 p-4">
              <div className="flex items-start justify-between mb-1">
                <div>
                  <p className="font-semibold text-slate-800">{n.topic || "Untitled topic"}</p>
                  <p className="text-xs text-slate-500">{n.subject} · {n.class} {n.date ? `· ${n.date}` : ""}</p>
                </div>
                <div className="flex gap-1 shrink-0">
                  <button onClick={() => setModal(n)} className="p-1 text-slate-400 hover:text-slate-700"><Edit2 size={14} /></button>
                  <button onClick={() => remove(n.id)} className="p-1 text-slate-400 hover:text-red-600"><Trash2 size={14} /></button>
                </div>
              </div>
              {n.objectives && <p className="text-xs text-slate-600 mt-1"><b>Objectives:</b> {n.objectives}</p>}
            </div>
          ))}
        </div>
      )}

      {modal && (
        <Modal title={modal.id ? "Edit lesson note" : "New lesson note"} onClose={() => setModal(null)} wide>
          <LessonNoteFormBody note={modal} onSave={save} onClose={() => setModal(null)} />
        </Modal>
      )}
    </div>
  );
}

function LessonNoteFormBody({ note, onSave, onClose }) {
  const [f, setF] = useState(note);
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  return (
    <>
      <div className="grid sm:grid-cols-2 gap-x-4">
        <Field label="Subject">
          <select className={inputCls} value={f.subject} onChange={set("subject")}>
            <option value="">Select a subject…</option>
            {ALL_SUBJECTS.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
        </Field>
        <Field label="Class"><input className={inputCls} value={f.class} onChange={set("class")} placeholder="e.g. Grade 6B" /></Field>
      </div>
      <div className="grid sm:grid-cols-2 gap-x-4">
        <Field label="Date"><input type="date" className={inputCls} value={f.date} onChange={set("date")} /></Field>
        <Field label="Topic"><input className={inputCls} value={f.topic} onChange={set("topic")} /></Field>
      </div>
      <Field label="Sub-topic"><input className={inputCls} value={f.subtopic} onChange={set("subtopic")} /></Field>
      <Field label="Objectives"><textarea className={inputCls} rows={2} value={f.objectives} onChange={set("objectives")} placeholder="By the end of the lesson, the learner should be able to…" /></Field>
      <Field label="Introduction"><textarea className={inputCls} rows={2} value={f.introduction} onChange={set("introduction")} /></Field>
      <Field label="Lesson development"><textarea className={inputCls} rows={4} value={f.development} onChange={set("development")} placeholder={"Step 1: …\nStep 2: …"} /></Field>
      <Field label="Conclusion"><textarea className={inputCls} rows={2} value={f.conclusion} onChange={set("conclusion")} /></Field>
      <Field label="Assignment"><textarea className={inputCls} rows={2} value={f.assignment} onChange={set("assignment")} /></Field>
      <Field label="References"><input className={inputCls} value={f.references} onChange={set("references")} /></Field>
      <div className="flex justify-end gap-2 mt-2">
        <Btn variant="secondary" onClick={onClose}>Cancel</Btn>
        <Btn onClick={() => f.subject && f.topic.trim() && onSave(f)} disabled={!f.subject || !f.topic.trim()}><Save size={14} /> Save note</Btn>
      </div>
    </>
  );
}

// ---------- Class Lists ----------
function ClassLists({ data, school }) {
  const classes = [...new Set(data.learners.map((l) => l.class).filter(Boolean))];
  const [cls, setCls] = useState(classes[0] || "");
  const [extraRows, setExtraRows] = useState(5);
  const list = data.learners.filter((l) => l.class === cls).sort((a, b) => a.name.localeCompare(b.name));

  if (!classes.length) return <EmptyState icon={BookOpen} title="No classes yet" hint="Assign a class to each learner profile to generate class lists." />;

  const blanks = Array.from({ length: Math.max(0, Number(extraRows) || 0) });

  const downloadExcel = () => {
    const rows = list.map((l, i) => ({
      "#": i + 1, Name: l.name, "Adm#": l.admissionNo || "", Gender: l.gender || "", DOB: l.dob || "",
      "Parent Name": l.parentName || "", "Parent Phone": l.parentPhone || "", "Parent Email": l.parentEmail || "",
      "KCPE/KJSEA": entryScoreLabel(l),
    }));
    blanks.forEach((_, i) => rows.push({ "#": list.length + i + 1, Name: "", "Adm#": "", Gender: "", DOB: "", "Parent Name": "", "Parent Phone": "", "Parent Email": "", "KCPE/KJSEA": "" }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, cls.slice(0, 31));
    XLSX.writeFile(wb, `${cls.replace(/[^a-z0-9]+/gi, "-")}-class-list.xlsx`);
  };

  return (
    <div>
      <div className="flex flex-wrap items-center justify-between gap-2 mb-4 print:hidden">
        <select className={`${inputCls} max-w-[220px]`} value={cls} onChange={(e) => setCls(e.target.value)}>
          {classes.map((c) => <option key={c} value={c}>{c}</option>)}
        </select>
        <div className="flex items-center gap-2">
          <label className="text-xs text-slate-500 flex items-center gap-1.5">Extra blank rows
            <input type="number" min="0" max="30" className="w-14 rounded border border-slate-300 px-1.5 py-1 text-center text-xs" value={extraRows} onChange={(e) => setExtraRows(clamp(e.target.value, 0, 30))} />
          </label>
        </div>
      </div>
      <div className="flex gap-2 mb-4 print:hidden">
        <Btn size="sm" variant="secondary" onClick={downloadExcel}><Download size={14} /> Download Excel</Btn>
        <Btn size="sm" variant="secondary" onClick={() => window.print()}><Printer size={14} /> Download PDF</Btn>
      </div>
      <div className="bg-white rounded-xl border border-slate-200 p-5 overflow-x-auto">
        <div className="flex items-center gap-3 mb-4 border-b-2 border-double border-slate-300 pb-3">
          <Seal letter="S" logo={school?.logoDataUrl} />
          <div>
            <h2 className="font-bold text-slate-800" style={serifStyle()}>{school ? `${school.name} — ` : ""}Class List — {cls}</h2>
            <p className="text-xs text-slate-500">{list.length} learner(s) enrolled</p>
          </div>
        </div>
        <table className="w-full text-sm min-w-[720px]">
          <thead><tr className="text-left text-xs text-slate-400 uppercase tracking-wide border-b border-slate-200">
            <th className="py-1.5 pr-2">#</th><th className="py-1.5 pr-2">Name</th><th className="py-1.5 pr-2">Adm#</th>
            <th className="py-1.5 pr-2">Gender</th><th className="py-1.5 pr-2">DOB</th>
            <th className="py-1.5 pr-2">Parent Name</th><th className="py-1.5 pr-2">Parent Phone</th>
            <th className="py-1.5 pr-2">Parent Email</th><th className="py-1.5">KCPE/KJSEA</th>
          </tr></thead>
          <tbody>
            {list.map((l, i) => (
              <tr key={l.id} className="border-b border-slate-100">
                <td className="py-1.5 pr-2 text-slate-400">{i + 1}</td>
                <td className="py-1.5 pr-2 font-medium text-slate-800 whitespace-nowrap">{l.name}</td>
                <td className="py-1.5 pr-2 text-slate-500">{l.admissionNo || "—"}</td>
                <td className="py-1.5 pr-2 text-slate-500">{l.gender || "—"}</td>
                <td className="py-1.5 pr-2 text-slate-500 whitespace-nowrap">{l.dob || "—"}</td>
                <td className="py-1.5 pr-2 text-slate-500 whitespace-nowrap">{l.parentName || "—"}</td>
                <td className="py-1.5 pr-2 text-slate-500 whitespace-nowrap">{l.parentPhone || "—"}</td>
                <td className="py-1.5 pr-2 text-slate-500 whitespace-nowrap">{l.parentEmail || "—"}</td>
                <td className="py-1.5 text-slate-500 whitespace-nowrap">{entryScoreLabel(l)}</td>
              </tr>
            ))}
            {blanks.map((_, i) => (
              <tr key={`blank-${i}`} className="border-b border-slate-100">
                <td className="py-3 pr-2 text-slate-300">{list.length + i + 1}</td>
                <td className="py-3 pr-2"></td><td className="py-3 pr-2"></td><td className="py-3 pr-2"></td>
                <td className="py-3 pr-2"></td><td className="py-3 pr-2"></td><td className="py-3 pr-2"></td>
                <td className="py-3 pr-2"></td><td className="py-3"></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ---------- Merit List ----------
function computeExamRanking(data, examId) {
  const exam = data.exams.find((e) => e.id === examId);
  if (!exam) return { exam: null, rows: [] };
  const maxMarks = exam.subjectMaxMarks || {};
  const classLearners = data.learners.filter((l) => l.class === exam.class);

  // Per-subject percentages across the class, for computing each learner's rank within that subject.
  const subjectPercents = {};
  exam.subjects.forEach((s) => { subjectPercents[s] = []; });
  classLearners.forEach((l) => {
    exam.subjects.forEach((s) => {
      const mark = data.marks.find((m) => m.examId === examId && m.learnerId === l.id && m.subject === s);
      if (mark) subjectPercents[s].push({ learnerId: l.id, pct: toPercent(mark.score, maxMarks[s] || 100) });
    });
  });
  const subjectRanks = {};
  Object.entries(subjectPercents).forEach(([s, arr]) => {
    const sorted = [...arr].sort((a, b) => b.pct - a.pct);
    subjectRanks[s] = {};
    sorted.forEach((entry, i) => { subjectRanks[s][entry.learnerId] = i + 1; });
  });

  const rows = classLearners.map((l) => {
    const marks = data.marks.filter((m) => m.examId === examId && m.learnerId === l.id);
    const total = marks.reduce((s, m) => s + Number(m.score || 0), 0);
    const subjectResults = exam.subjects.map((s) => {
      const m = marks.find((mk) => mk.subject === s);
      if (!m) return null;
      const pct = toPercent(m.score, maxMarks[s] || 100);
      return {
        subject: s, score: m.score, outOf: maxMarks[s] || 100, pct, grade: gradeForPercent(getSystemForSubject(data, s), pct),
        rank: (subjectRanks[s] || {})[l.id] || null, outOfRank: (subjectPercents[s] || []).length,
      };
    }).filter(Boolean);
    const percentages = subjectResults.map((r) => r.pct);
    const average = percentages.length ? percentages.reduce((a, b) => a + b, 0) / percentages.length : 0;
    const baselinePercent = (l.kjseaScore !== undefined && l.kjseaScore !== null && l.kjseaScore !== "")
      ? toPercent(l.kjseaScore, 900)
      : (l.kcpeScore !== undefined && l.kcpeScore !== null && l.kcpeScore !== "")
        ? toPercent(l.kcpeScore, 500)
        : null;
    const valueAdded = baselinePercent !== null ? Number((average - baselinePercent).toFixed(1)) : null;
    return { learner: l, total, average, grade: gradeForPercent(getDefaultSystem(data), average), subjectResults, valueAdded, baselinePercent, subjectCount: marks.length };
  }).filter((r) => r.subjectCount > 0);
  rows.sort((a, b) => b.average - a.average);
  rows.forEach((r, i) => { r.rank = i + 1; });
  return { exam, rows };
}

function getSubjectTeacher(data, cls, subject) {
  return data.teachers.find((t) => t.subject === subject && (t.classesAssigned || "").split(",").map((s) => s.trim()).includes(cls));
}
function getClassTeacher(data, cls) {
  return data.teachers.find((t) => (t.classTeacherOf || "").trim() === cls);
}
function attendanceRate(data, learnerId, cls) {
  const records = (data.attendance || []).filter((a) => a.class === cls && a.learnerId === learnerId);
  if (!records.length) return null;
  const present = records.filter((a) => a.status === "present").length;
  return { present, total: records.length, rate: Math.round((present / records.length) * 100) };
}

function entryScoreLabel(learner) {
  if (learner.kjseaScore !== undefined && learner.kjseaScore !== null && learner.kjseaScore !== "") return `${learner.kjseaScore} KJSEA`;
  if (learner.kcpeScore !== undefined && learner.kcpeScore !== null && learner.kcpeScore !== "") return `${learner.kcpeScore} KCPE`;
  return "—";
}

function MeritList({ data, school }) {
  const [examId, setExamId] = useState(data.exams[0]?.id || "");
  const { exam, rows } = useMemo(() => computeExamRanking(data, examId), [data, examId]);
  const availableSystems = getGradingSystems(data);
  const [rankSystemId, setRankSystemId] = useState(data.defaultGradingSystemId || "standard");
  const rankSystem = availableSystems.find((s) => s.id === rankSystemId) || availableSystems[0];
  const overallGrade = (r) => gradeForPercent(rankSystem, r.average);

  if (!data.exams.length) return <EmptyState icon={Award} title="No exams yet" hint="Create an exam and record marks to generate a merit list." />;

  const downloadExcel = () => {
    if (!exam || !rows.length) return;
    const sheetRows = rows.map((r) => {
      const row = { Rank: r.rank, Name: r.learner.name, "KCPE/KJSEA": entryScoreLabel(r.learner) };
      exam.subjects.forEach((s) => {
        const sr = r.subjectResults.find((x) => x.subject === s);
        row[s] = sr ? `${sr.score} ${sr.grade}` : "";
      });
      row["Total"] = r.total;
      row["Avg %"] = Number(r.average.toFixed(1));
      row["Grade"] = overallGrade(r);
      row["Value Added"] = r.valueAdded === null ? "" : r.valueAdded;
      return row;
    });
    const ws = XLSX.utils.json_to_sheet(sheetRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Merit List");
    XLSX.writeFile(wb, `${exam.name.replace(/[^a-z0-9]+/gi, "-")}-merit-list.xlsx`);
  };

  return (
    <div>
      <div className="flex flex-wrap items-center justify-between gap-2 mb-2 print:hidden">
        <select className={`${inputCls} max-w-[240px]`} value={examId} onChange={(e) => setExamId(e.target.value)}>
          {data.exams.map((e) => <option key={e.id} value={e.id}>{e.name} — {e.class}</option>)}
        </select>
        <div className="flex gap-2">
          <Btn size="sm" variant="secondary" onClick={downloadExcel} disabled={!rows.length}><Download size={14} /> Excel</Btn>
          <Btn size="sm" variant="secondary" onClick={() => window.print()}><Printer size={14} /> PDF</Btn>
        </div>
      </div>
      <div className="flex items-center gap-2 mb-4 print:hidden">
        <label className="text-xs text-slate-500 shrink-0">Rank overall grade using</label>
        <select className={`${inputCls} max-w-[280px]`} value={rankSystemId} onChange={(e) => setRankSystemId(e.target.value)}>
          {availableSystems.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
        </select>
      </div>
      {exam && (
        <div className="bg-white rounded-xl border border-slate-200 p-5">
          <div className="flex items-center gap-3 mb-4 border-b-2 border-double border-slate-300 pb-3">
            <Seal letter="M" logo={school?.logoDataUrl} />
            <div>
              <h2 className="font-bold text-slate-800" style={serifStyle()}>Merit List — {exam.name}</h2>
              <p className="text-xs text-slate-500">{exam.class} · ranked by average percentage · overall grade: {rankSystem.name}</p>
            </div>
          </div>
          {rows.length === 0 ? <EmptyState icon={Award} title="No marks recorded for this exam yet" /> : (
            <>
              <div className="overflow-x-auto">
                <table className="w-full text-sm min-w-[640px]">
                  <thead><tr className="text-left text-xs text-slate-400 uppercase tracking-wide border-b border-slate-200">
                    <th className="py-1.5 pr-2">Rank</th>
                    <th className="py-1.5 pr-2">Name</th>
                    <th className="py-1.5 px-2 text-center whitespace-nowrap">KCPE/KJSEA</th>
                    {exam.subjects.map((s) => <th key={s} className="py-1.5 px-2 text-center whitespace-nowrap">{s}</th>)}
                    <th className="py-1.5 px-2 text-right">Total</th>
                    <th className="py-1.5 px-2 text-right">Avg %</th>
                    <th className="py-1.5 px-2 text-right">Grade</th>
                    <th className="py-1.5 pl-2 text-right">Value added</th>
                  </tr></thead>
                  <tbody>
                    {rows.map((r) => (
                      <tr key={r.learner.id} className="border-b border-slate-100">
                        <td className="py-1.5 pr-2">
                          <span className={`inline-flex w-6 h-6 items-center justify-center rounded-full text-xs font-bold ${r.rank <= 3 ? "bg-amber-100 text-amber-700" : "text-slate-400"}`}>{r.rank}</span>
                        </td>
                        <td className="py-1.5 pr-2 font-medium text-slate-800 whitespace-nowrap">{r.learner.name}</td>
                        <td className="py-1.5 px-2 text-center text-slate-500 whitespace-nowrap">{entryScoreLabel(r.learner)}</td>
                        {exam.subjects.map((s) => {
                          const sr = r.subjectResults.find((x) => x.subject === s);
                          return (
                            <td key={s} className="py-1.5 px-2 text-center text-slate-600 whitespace-nowrap">
                              {sr ? `${sr.score} ${sr.grade}` : "—"}
                            </td>
                          );
                        })}
                        <td className="py-1.5 px-2 text-right font-semibold text-emerald-800">{r.total}</td>
                        <td className="py-1.5 px-2 text-right text-slate-500">{r.average.toFixed(1)}%</td>
                        <td className="py-1.5 px-2 text-right font-bold text-slate-700">{overallGrade(r)}</td>
                        <td className="py-1.5 pl-2 text-right font-semibold whitespace-nowrap">
                          {r.valueAdded === null ? <span className="text-slate-300">—</span> : (
                            <span className={r.valueAdded > 0 ? "text-emerald-700" : r.valueAdded < 0 ? "text-red-600" : "text-slate-500"}>
                              {r.valueAdded > 0 ? "+" : ""}{r.valueAdded}
                            </span>
                          )}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <p className="text-[11px] text-slate-400 mt-2">Value added = this exam's average % minus the learner's declared KCPE (out of 500) or KJSEA (out of 900) score, converted to a percentage — set per learner in their profile. "—" means no entry score is on file.</p>

              <div className="mt-5 pt-4 border-t border-slate-200">
                <p className="text-xs font-semibold text-slate-500 uppercase tracking-wide mb-2">Teachers</p>
                <div className="grid sm:grid-cols-2 gap-1.5 text-sm">
                  {(() => {
                    const ct = getClassTeacher(data, exam.class);
                    return <p className="text-slate-600"><span className="text-slate-400">Class teacher:</span> {ct ? ct.name : "—"}</p>;
                  })()}
                  {exam.subjects.map((s) => {
                    const t = getSubjectTeacher(data, exam.class, s);
                    return <p key={s} className="text-slate-600"><span className="text-slate-400">{s}:</span> {t ? t.name : "—"}</p>;
                  })}
                </div>
              </div>
              <GradingKey data={data} systems={[...new Map([rankSystem, ...exam.subjects.map((s) => getSystemForSubject(data, s))].map((sys) => [sys.id, sys])).values()]} />
            </>
          )}
        </div>
      )}
    </div>
  );
}

function gradeColor(pctOrGrade) {
  // Accepts a percentage (preferred, works for any grading system) or falls back to matching
  // legacy Standard-scale labels if a raw grade string is passed instead.
  if (typeof pctOrGrade === "number") {
    const pct = pctOrGrade;
    if (pct >= 80) return "bg-emerald-100 text-emerald-800";
    if (pct >= 65) return "bg-blue-100 text-blue-800";
    if (pct >= 50) return "bg-amber-100 text-amber-800";
    if (pct >= 35) return "bg-orange-100 text-orange-800";
    return "bg-red-100 text-red-800";
  }
  const g = pctOrGrade;
  if (["A", "A-", "EE1", "EE2"].includes(g)) return "bg-emerald-100 text-emerald-800";
  if (["B+", "B", "B-", "ME1"].includes(g)) return "bg-blue-100 text-blue-800";
  if (["C+", "C", "C-", "ME2"].includes(g)) return "bg-amber-100 text-amber-800";
  if (["D+", "D", "D-", "BE1"].includes(g)) return "bg-orange-100 text-orange-800";
  return "bg-red-100 text-red-800";
}

function GradingKey({ data, systems }) {
  const list = systems && systems.length ? systems : [getDefaultSystem(data)];
  return (
    <div className="border-t border-slate-100 pt-3 mt-4">
      <p className="text-xs font-semibold text-slate-500 uppercase tracking-wide mb-2">Grading key</p>
      {list.map((system) => {
        const bands = [...system.bands].sort((a, b) => a.min - b.min);
        return (
          <div key={system.id} className="mb-2 last:mb-0">
            {list.length > 1 && <p className="text-[10px] text-slate-400 mb-1">{system.name}</p>}
            <div className="flex flex-wrap gap-1.5">
              {bands.map((b, i) => {
                const nextMin = bands[i + 1]?.min;
                const rangeLabel = nextMin !== undefined ? `${b.min}-${nextMin - 1}` : `${b.min}+`;
                return (
                  <span key={b.label} className={`text-[11px] font-medium px-2 py-0.5 rounded-full ${gradeColor(b.min)}`}>
                    {b.label} ({rangeLabel}%)
                  </span>
                );
              })}
            </div>
          </div>
        );
      })}
    </div>
  );
}

// ---------- Report Cards ----------
function ReportCards({ data, school, persist, notify, user }) {
  const [examId, setExamId] = useState(data.exams[0]?.id || "");
  const exam = data.exams.find((e) => e.id === examId);
  const classLearners = exam ? data.learners.filter((l) => l.class === exam.class) : [];
  const [mode, setMode] = useState("one"); // one | all
  const [learnerId, setLearnerId] = useState("");
  const { rows } = useMemo(() => computeExamRanking(data, examId), [data, examId]);
  const row = rows.find((r) => r.learner.id === learnerId);

  if (!data.exams.length) return <EmptyState icon={FileSpreadsheet} title="No exams yet" hint="Create an exam and record marks to generate report cards." />;

  return (
    <div>
      <div className="flex gap-2 mb-3 print:hidden">
        <button onClick={() => setMode("one")} className={`flex-1 py-1.5 rounded-full text-xs font-medium border ${mode === "one" ? "bg-emerald-800 text-white border-emerald-800" : "bg-white text-slate-600 border-slate-300"}`}>One learner</button>
        <button onClick={() => setMode("all")} className={`flex-1 py-1.5 rounded-full text-xs font-medium border ${mode === "all" ? "bg-emerald-800 text-white border-emerald-800" : "bg-white text-slate-600 border-slate-300"}`}>All learners in class</button>
      </div>

      <div className={`grid ${mode === "one" ? "grid-cols-2" : "grid-cols-1"} gap-2 mb-4 print:hidden`}>
        <select className={inputCls} value={examId} onChange={(e) => { setExamId(e.target.value); setLearnerId(""); }}>
          {data.exams.map((e) => <option key={e.id} value={e.id}>{e.name} — {e.class}</option>)}
        </select>
        {mode === "one" && (
          <select className={inputCls} value={learnerId} onChange={(e) => setLearnerId(e.target.value)}>
            <option value="">Select learner…</option>
            {classLearners.map((l) => <option key={l.id} value={l.id}>{l.name}</option>)}
          </select>
        )}
      </div>

      {mode === "all" && rows.length > 0 && (
        <div className="flex justify-end mb-3 print:hidden">
          <Btn size="sm" variant="secondary" onClick={() => window.print()}><Printer size={14} /> Download PDF (all {rows.length})</Btn>
        </div>
      )}

      {mode === "one" ? (
        !row ? (
          <EmptyState icon={FileSpreadsheet} title="Select a learner" hint="Pick an exam and learner above to generate their report card." />
        ) : (
          <SingleReportCard data={data} school={school} exam={exam} row={row} rows={rows} persist={persist} notify={notify} user={user} showPrintButton />
        )
      ) : rows.length === 0 ? (
        <EmptyState icon={FileSpreadsheet} title="No marks recorded for this exam yet" />
      ) : (
        <div className="space-y-6">
          {rows.map((r) => (
            <div key={r.learner.id} className="print:break-after-page">
              <SingleReportCard data={data} school={school} exam={exam} row={r} rows={rows} persist={persist} notify={notify} user={user} />
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function SingleReportCard({ data, school, exam, row, rows, persist, notify, user, showPrintButton }) {
  const examId = exam.id;
  const learnerId = row.learner.id;
  const canEditComments = user && (user.role === "Super Admin" || user.role === "Admin" || user.role === "Teacher");
  const comments = data.reportComments || [];
  const existingComment = comments.find((c) => c.examId === examId && c.learnerId === learnerId);
  const [classComment, setClassComment] = useState(existingComment?.classTeacherComment || "");
  const [principalComment, setPrincipalComment] = useState(existingComment?.principalComment || "");
  const [editingComments, setEditingComments] = useState(false);
  const attendance = attendanceRate(data, learnerId, exam.class);

  const saveComments = () => {
    const exists = comments.some((c) => c.examId === examId && c.learnerId === learnerId);
    const next = exists
      ? comments.map((c) => (c.examId === examId && c.learnerId === learnerId ? { ...c, classTeacherComment: classComment, principalComment } : c))
      : [...comments, { id: uid(), examId, learnerId, classTeacherComment: classComment, principalComment }];
    persist({ ...data, reportComments: next });
    notify("Comments saved");
    setEditingComments(false);
  };

  return (
    <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-sm">
      {showPrintButton && (
        <div className="flex justify-end p-2 print:hidden bg-stone-50 border-b border-slate-100">
          <Btn size="sm" variant="secondary" onClick={() => window.print()}><Printer size={14} /> Download PDF</Btn>
        </div>
      )}

      {/* Letterhead */}
      <div className="bg-gradient-to-r from-slate-900 to-emerald-900 text-white p-5 relative overflow-hidden">
        <Sparkles size={90} className="absolute -right-4 -top-4 text-amber-400/20" />
        <div className="flex items-center gap-3 relative">
          <Seal letter="S" size={52} logo={school?.logoDataUrl} />
          <div>
            <p className="text-amber-300 text-[11px] font-semibold uppercase tracking-wide">{school?.name || "Skolar"}</p>
            <h2 className="font-bold text-xl" style={serifStyle()}>Report Card</h2>
            <p className="text-slate-300 text-xs">{exam.name} · {exam.class} · {exam.date || "No date"}</p>
          </div>
        </div>
      </div>

      {/* Learner strip */}
      <div className="flex items-center gap-3 p-4 border-b border-slate-100 bg-stone-50">
        <Seal letter={row.learner.name.charAt(0).toUpperCase() || "?"} size={44} />
        <div className="flex-1 min-w-0">
          <p className="font-bold text-slate-800 truncate" style={serifStyle()}>{row.learner.name}</p>
          <p className="text-xs text-slate-500">Adm# {row.learner.admissionNo || "—"} · {row.learner.class || exam.class}</p>
        </div>
        <div className={`px-3 py-1.5 rounded-full text-center ${gradeColor(row.average)}`}>
          <p className="text-lg font-bold leading-none" style={serifStyle()}>{row.grade}</p>
        </div>
      </div>

      <div className="p-5">
        {/* Learner details */}
        <div className="grid grid-cols-3 gap-2 mb-4 text-sm">
          <div><span className="text-slate-400 text-xs uppercase tracking-wide block">Gender</span>{row.learner.gender || "—"}</div>
          <div><span className="text-slate-400 text-xs uppercase tracking-wide block">Entry score</span>{entryScoreLabel(row.learner)}</div>
          <div><span className="text-slate-400 text-xs uppercase tracking-wide block">Attendance</span>{attendance ? `${attendance.rate}% (${attendance.present}/${attendance.total})` : "—"}</div>
        </div>

        {/* Subject table */}
        <div className="overflow-x-auto mb-4">
          <table className="w-full text-sm min-w-[520px]">
            <thead><tr className="text-left text-xs text-slate-400 uppercase tracking-wide border-b border-slate-200">
              <th className="py-1.5">Subject</th><th className="py-1.5 text-right">Score</th><th className="py-1.5 text-right">Out of</th>
              <th className="py-1.5 text-right">%</th><th className="py-1.5 text-right">Grade</th><th className="py-1.5 text-right">Rank</th>
            </tr></thead>
            <tbody>
              {row.subjectResults.map((sr) => (
                <tr key={sr.subject} className="border-b border-slate-100">
                  <td className="py-2 text-slate-700">{sr.subject}</td>
                  <td className="py-2 text-right font-medium text-slate-800">{sr.score}</td>
                  <td className="py-2 text-right text-slate-400">{sr.outOf}</td>
                  <td className="py-2 text-right text-slate-500">{sr.pct.toFixed(0)}%</td>
                  <td className="py-2 text-right">
                    <span className={`px-2 py-0.5 rounded-full text-xs font-bold ${gradeColor(sr.pct)}`}>{sr.grade}</span>
                  </td>
                  <td className="py-2 text-right text-slate-500">{sr.rank ? `${sr.rank}/${sr.outOfRank}` : "—"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {/* Summary cards */}
        <div className="grid grid-cols-2 sm:grid-cols-5 gap-2 mb-5">
          <div className="bg-stone-50 rounded-lg p-3 text-center"><p className="text-[10px] text-slate-400 uppercase tracking-wide">Total</p><p className="text-lg font-bold text-emerald-800">{row.total}</p></div>
          <div className="bg-stone-50 rounded-lg p-3 text-center"><p className="text-[10px] text-slate-400 uppercase tracking-wide">Average</p><p className="text-lg font-bold text-emerald-800">{row.average.toFixed(1)}%</p></div>
          <div className="bg-stone-50 rounded-lg p-3 text-center"><p className="text-[10px] text-slate-400 uppercase tracking-wide">Grade</p><p className="text-lg font-bold text-emerald-800">{row.grade}</p></div>
          <div className="bg-amber-50 rounded-lg p-3 text-center"><p className="text-[10px] text-slate-400 uppercase tracking-wide">Rank</p><p className="text-lg font-bold text-amber-700">{row.rank}/{rows.length}</p></div>
          <div className={`rounded-lg p-3 text-center ${row.valueAdded === null ? "bg-slate-50" : row.valueAdded >= 0 ? "bg-emerald-50" : "bg-red-50"}`}>
            <p className="text-[10px] text-slate-400 uppercase tracking-wide">Value added</p>
            <p className={`text-lg font-bold flex items-center justify-center gap-1 ${row.valueAdded === null ? "text-slate-300" : row.valueAdded >= 0 ? "text-emerald-700" : "text-red-600"}`}>
              {row.valueAdded === null ? "—" : <>{row.valueAdded >= 0 ? <TrendingUp size={14} /> : <TrendingUp size={14} className="rotate-180" />}{row.valueAdded > 0 ? "+" : ""}{row.valueAdded}</>}
            </p>
          </div>
        </div>

        {/* Teachers */}
        <div className="border-t border-slate-100 pt-4 mb-4">
          <p className="text-xs font-semibold text-slate-500 uppercase tracking-wide mb-2">Teachers</p>
          <div className="grid sm:grid-cols-2 gap-1.5 text-sm">
            {(() => {
              const ct = getClassTeacher(data, exam.class);
              return <p className="text-slate-600"><span className="text-slate-400">Class teacher:</span> {ct ? ct.name : "—"}</p>;
            })()}
            {exam.subjects.map((s) => {
              const t = getSubjectTeacher(data, exam.class, s);
              return <p key={s} className="text-slate-600"><span className="text-slate-400">{s}:</span> {t ? t.name : "—"}</p>;
            })}
          </div>
        </div>

        <GradingKey data={data} systems={[...new Map(exam.subjects.map((s) => getSystemForSubject(data, s)).map((sys) => [sys.id, sys])).values()]} />

        {/* Comments */}
        <div className="border-t border-slate-100 pt-4">
          <div className="flex items-center justify-between mb-2 print:hidden">
            <p className="text-xs font-semibold text-slate-500 uppercase tracking-wide">Comments</p>
            {canEditComments && !editingComments && (
              <button onClick={() => setEditingComments(true)} className="text-xs text-emerald-800 font-medium flex items-center gap-1"><Edit2 size={12} /> Edit</button>
            )}
          </div>
          {editingComments ? (
            <div className="space-y-3">
              <Field label="Class teacher's comment"><textarea className={inputCls} rows={2} value={classComment} onChange={(e) => setClassComment(e.target.value)} placeholder="e.g. A hardworking learner, keep it up." /></Field>
              <Field label="Principal's / admin's comment"><textarea className={inputCls} rows={2} value={principalComment} onChange={(e) => setPrincipalComment(e.target.value)} placeholder="e.g. Good overall performance this term." /></Field>
              <div className="flex justify-end gap-2">
                <Btn size="sm" variant="secondary" onClick={() => setEditingComments(false)}>Cancel</Btn>
                <Btn size="sm" onClick={saveComments}><Save size={13} /> Save comments</Btn>
              </div>
            </div>
          ) : (
            <div className="space-y-2 text-sm">
              <p className="text-slate-600"><span className="text-slate-400 text-xs uppercase tracking-wide block">Class teacher</span>{classComment || <span className="text-slate-300 italic">No comment yet</span>}</p>
              <p className="text-slate-600"><span className="text-slate-400 text-xs uppercase tracking-wide block">Principal / admin</span>{principalComment || <span className="text-slate-300 italic">No comment yet</span>}</p>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ---------- Messages to parents ----------
function Messages({ data, persist, notify }) {
  const classes = [...new Set(data.learners.map((l) => l.class).filter(Boolean))];
  const [channel, setChannel] = useState("email"); // email | sms
  const [scope, setScope] = useState("class");
  const [cls, setCls] = useState(classes[0] || "");
  const [learnerId, setLearnerId] = useState("");
  const [subject, setSubject] = useState("");
  const [text, setText] = useState("");

  const hasContact = (l) => (channel === "email" ? !!l.parentEmail : !!l.parentPhone);

  const recipients = useMemo(() => {
    if (scope === "all") return data.learners.filter(hasContact);
    if (scope === "class") return data.learners.filter((l) => l.class === cls && hasContact(l));
    if (scope === "individual") return data.learners.filter((l) => l.id === learnerId && hasContact(l));
    return [];
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [scope, cls, learnerId, data.learners, channel]);

  const withoutContact = useMemo(() => {
    if (scope === "all") return data.learners.filter((l) => !hasContact(l));
    if (scope === "class") return data.learners.filter((l) => l.class === cls && !hasContact(l));
    return [];
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [scope, cls, data.learners, channel]);

  const logAndSend = (learner) => {
    const body = text.replace(/\{name\}/g, learner.name);
    if (channel === "email") {
      const mailto = `mailto:${learner.parentEmail}?subject=${encodeURIComponent(subject || "Message from school")}&body=${encodeURIComponent(body)}`;
      window.open(mailto, "_blank");
      persist({ ...data, messages: [...data.messages, { id: uid(), date: new Date().toISOString(), learnerId: learner.id, recipient: learner.parentEmail, subject, text: body, channel: "email" }] });
    } else {
      const sms = `sms:${learner.parentPhone}?body=${encodeURIComponent(body)}`;
      window.open(sms, "_blank");
      persist({ ...data, messages: [...data.messages, { id: uid(), date: new Date().toISOString(), learnerId: learner.id, recipient: learner.parentPhone, subject: "", text: body, channel: "sms" }] });
    }
  };

  const sendAll = () => {
    if (!recipients.length || !text.trim()) return;
    recipients.forEach(logAndSend);
    notify(`Opened ${recipients.length} ${channel === "email" ? "email" : "SMS"} draft(s) for sending`);
  };

  const exportCsv = () => {
    const rows = recipients.map((l) => ({ Name: l.name, ParentEmail: l.parentEmail, ParentPhone: l.parentPhone, Message: text.replace(/\{name\}/g, l.name) }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Messages");
    XLSX.writeFile(wb, "parent-messages.csv");
  };

  return (
    <div>
      <div className="bg-white rounded-xl border border-slate-200 p-4 mb-4">
        <h3 className="font-semibold text-sm text-slate-700 mb-3">Send via</h3>
        <div className="flex gap-2 mb-3">
          <button onClick={() => setChannel("email")} className={`flex-1 py-1.5 rounded-full text-xs font-medium border flex items-center justify-center gap-1.5 ${channel === "email" ? "bg-emerald-800 text-white border-emerald-800" : "bg-white text-slate-600 border-slate-300"}`}><Mail size={13} /> Email</button>
          <button onClick={() => setChannel("sms")} className={`flex-1 py-1.5 rounded-full text-xs font-medium border flex items-center justify-center gap-1.5 ${channel === "sms" ? "bg-emerald-800 text-white border-emerald-800" : "bg-white text-slate-600 border-slate-300"}`}><MessageCircle size={13} /> SMS</button>
        </div>
      </div>

      <div className="bg-white rounded-xl border border-slate-200 p-4 mb-4">
        <h3 className="font-semibold text-sm text-slate-700 mb-3">Recipients</h3>
        <div className="flex gap-2 mb-3">
          {[["class", "By class"], ["individual", "One learner"], ["all", "Everyone"]].map(([v, l]) => (
            <button key={v} onClick={() => setScope(v)} className={`px-3 py-1.5 rounded-full text-xs font-medium border ${scope === v ? "bg-emerald-800 text-white border-emerald-800" : "bg-white text-slate-600 border-slate-300"}`}>{l}</button>
          ))}
        </div>
        {scope === "class" && (
          <select className={inputCls} value={cls} onChange={(e) => setCls(e.target.value)}>
            {classes.map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
        )}
        {scope === "individual" && (
          <select className={inputCls} value={learnerId} onChange={(e) => setLearnerId(e.target.value)}>
            <option value="">Select learner…</option>
            {data.learners.map((l) => <option key={l.id} value={l.id}>{l.name}</option>)}
          </select>
        )}
        <p className="text-xs text-slate-500 mt-2">{recipients.length} parent(s) with a {channel === "email" ? "parent email" : "parent phone"} on file will receive this.{withoutContact.length > 0 && ` ${withoutContact.length} learner(s) in scope have no parent ${channel === "email" ? "email" : "phone"} saved.`}</p>
      </div>

      <div className="bg-white rounded-xl border border-slate-200 p-4 mb-4">
        <h3 className="font-semibold text-sm text-slate-700 mb-3">Compose message</h3>
        {channel === "email" && <Field label="Subject"><input className={inputCls} value={subject} onChange={(e) => setSubject(e.target.value)} placeholder="e.g. Midterm results available" /></Field>}
        <Field label="Message (use {name} to personalize)"><textarea className={inputCls} rows={5} value={text} onChange={(e) => setText(e.target.value)} placeholder="Dear parent/guardian of {name}, …" /></Field>
        <div className="flex flex-wrap gap-2">
          <Btn onClick={sendAll} disabled={!recipients.length || !text.trim()}><Send size={14} /> Send to {recipients.length || 0}</Btn>
          <Btn variant="secondary" onClick={exportCsv} disabled={!recipients.length}><Download size={14} /> Export list as CSV</Btn>
        </div>
        <p className="text-[11px] text-slate-400 mt-2">
          {channel === "email"
            ? "\"Send\" opens a pre-filled email draft per parent in your device's mail app — nothing is sent automatically in the background."
            : "\"Send\" opens a pre-filled SMS draft per parent in your device's messaging app, one at a time — you'll need to tap send in each. No SMS gateway is connected, so nothing sends automatically."}
        </p>
      </div>

      <div className="bg-white rounded-xl border border-slate-200 p-4">
        <h3 className="font-semibold text-sm text-slate-700 mb-2">Message log</h3>
        {data.messages.length === 0 ? <p className="text-xs text-slate-400">No messages sent yet.</p> : (
          <div className="divide-y divide-slate-100 max-h-64 overflow-y-auto">
            {[...data.messages].reverse().map((m) => (
              <div key={m.id} className="py-2 text-xs">
                <p className="font-medium text-slate-700 flex items-center gap-1.5">
                  {m.channel === "sms" ? <MessageCircle size={11} className="text-slate-400" /> : <Mail size={11} className="text-slate-400" />}
                  {m.subject || (m.channel === "sms" ? "SMS" : "(no subject)")} → {m.recipient}
                </p>
                <p className="text-slate-400">{new Date(m.date).toLocaleString()}</p>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
