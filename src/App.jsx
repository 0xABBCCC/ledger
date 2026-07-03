import { useState, useEffect, useMemo, useCallback, useRef } from "react";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell, Legend,
} from "recharts";
import {
  Home, ShoppingCart, Car, Zap, HeartPulse, Film, ShoppingBag, MoreHorizontal,
  Wallet, PiggyBank, Plus, X, Pencil, Trash2, ChevronLeft, ChevronRight,
  BookOpen, PieChart as PieIcon, Loader2, Globe, Download, Upload,
} from "lucide-react";
import * as XLSX from "xlsx";

// window.storage only exists inside a live Claude.ai artifact preview.
// When this file runs anywhere else (local Vite/CRA/Next dev server,
// a static build, etc.) it's undefined, so every save silently fails.
// This polyfill makes the same code work in both places by falling
// back to localStorage when window.storage isn't present.
const storage = (typeof window !== "undefined" && window.storage)
  ? window.storage
  : {
      async get(key, shared = false) {
        const raw = window.localStorage.getItem(key);
        if (raw === null) throw new Error(`key not found: ${key}`);
        return { key, value: raw, shared };
      },
      async set(key, value, shared = false) {
        window.localStorage.setItem(key, value);
        return { key, value, shared };
      },
      async delete(key, shared = false) {
        window.localStorage.removeItem(key);
        return { key, deleted: true, shared };
      },
      async list(prefix = "", shared = false) {
        const keys = Object.keys(window.localStorage).filter((k) => k.startsWith(prefix));
        return { keys, prefix, shared };
      },
    };

const MONTH_NAMES = {
  en: ["January","February","March","April","May","June","July","August","September","October","November","December"],
  el: ["Ιανουάριος","Φεβρουάριος","Μάρτιος","Απρίλιος","Μάιος","Ιούνιος","Ιούλιος","Αύγουστος","Σεπτέμβριος","Οκτώβριος","Νοέμβριος","Δεκέμβριος"],
};
const MONTH_SHORT = {
  en: ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"],
  el: ["Ιαν","Φεβ","Μαρ","Απρ","Μαϊ","Ιουν","Ιουλ","Αυγ","Σεπ","Οκτ","Νοε","Δεκ"],
};
const WEEKDAYS = {
  en: ["S","M","T","W","T","F","S"],
  el: ["Κ","Δ","Τ","Τ","Π","Π","Σ"],
};

const LANG_KEY = "ledger_lang_v1";
const TX_KEY = "ledger_transactions_v1";

const CATEGORIES = [
  { id: "rent",      type: "expense", color: "#8B5E3C", Icon: Home },
  { id: "food",      type: "expense", color: "#4C7A5E", Icon: ShoppingCart },
  { id: "transport", type: "expense", color: "#3D6B8A", Icon: Car },
  { id: "utilities", type: "expense", color: "#B8863B", Icon: Zap },
  { id: "health",    type: "expense", color: "#A6432C", Icon: HeartPulse },
  { id: "fun",       type: "expense", color: "#7A5C9E", Icon: Film },
  { id: "shopping",  type: "expense", color: "#C46B4F", Icon: ShoppingBag },
  { id: "other_exp", type: "expense", color: "#847A68", Icon: MoreHorizontal },
  { id: "salary",    type: "income",  color: "#2F7D53", Icon: Wallet },
  { id: "savings",   type: "income",  color: "#1F6E5C", Icon: PiggyBank },
  { id: "other_inc", type: "income",  color: "#5C8A6A", Icon: MoreHorizontal },
];
const catById = (id) => CATEGORIES.find((c) => c.id === id);

const TRANSLATIONS = {
  en: {
    appTitle: "The Ledger",
    appSubtitle: "A running account of what comes in and what goes out",
    yearTotal: "Year total",
    loading: "Opening the ledger…",
    saveError: "Could not save. Your last change may not stick around.",
    tabLedger: "Ledger",
    tabInsights: "Insights",
    income: "Income",
    spent: "Spent",
    net: "Net",
    toggleSpent: "Spent",
    received: "Received",
    addEntry: "Add entry",
    saveChanges: "Save changes",
    cancelEdit: "Cancel edit",
    amount: "Amount",
    reason: "Reason",
    reasonPlaceholder: "e.g. weekly groceries",
    spendingByWeek: "Spending by week",
    whereItWent: "Where it went",
    yearAtGlance: "Year at a glance",
    noSpending: "No spending logged for this month yet.",
    weekLabel: "Wk",
    language: "Language",
    export: "Export",
    import: "Import",
    exportedSheetTransactions: "Transactions",
    exportedSheetMonthly: "Monthly Summary",
    exportedSheetCategory: "Category Breakdown",
    colDate: "Date",
    colType: "Type",
    colCategory: "Category",
    colAmount: "Amount",
    colNote: "Note",
    colCategoryId: "Category ID",
    colTypeCode: "Type Code",
    colMonth: "Month",
    confirmYes: "Yes, delete",
    confirmNo: "Cancel",
    cancel: "Cancel",
    save: "Save",
    importSuccess: "Imported {{count}} transactions.",
    importError: "Could not read that file. Try exporting from this app first, then importing that file back in.",
    cat_rent: "Rent & Housing",
    cat_food: "Food & Groceries",
    cat_transport: "Transport",
    cat_utilities: "Utilities",
    cat_health: "Health",
    cat_fun: "Entertainment",
    cat_shopping: "Shopping",
    cat_other_exp: "Other",
    cat_salary: "Salary & Income",
    cat_savings: "Savings & Interest",
    cat_other_inc: "Other Income",
  },
  el: {
    appTitle: "Το Λογιστικό Βιβλίο",
    appSubtitle: "Μια συνεχής καταγραφή εσόδων και εξόδων",
    yearTotal: "Σύνολο έτους",
    loading: "Άνοιγμα του βιβλίου…",
    saveError: "Δεν ήταν δυνατή η αποθήκευση. Η τελευταία αλλαγή ενδέχεται να μη διατηρηθεί.",
    tabLedger: "Βιβλίο",
    tabInsights: "Στατιστικά",
    income: "Έσοδα",
    spent: "Έξοδα",
    net: "Καθαρό",
    toggleSpent: "Πληρωμή",
    received: "Είσπραξη",
    addEntry: "Προσθήκη εγγραφής",
    saveChanges: "Αποθήκευση αλλαγών",
    cancelEdit: "Ακύρωση επεξεργασίας",
    amount: "Ποσό",
    reason: "Αιτιολογία",
    reasonPlaceholder: "π.χ. εβδομαδιαία ψώνια",
    spendingByWeek: "Έξοδα ανά εβδομάδα",
    whereItWent: "Πού πήγαν τα χρήματα",
    yearAtGlance: "Το έτος με μια ματιά",
    noSpending: "Δεν έχουν καταγραφεί έξοδα για αυτόν τον μήνα ακόμα.",
    weekLabel: "Εβδ",
    language: "Γλώσσα",
    export: "Εξαγωγή",
    import: "Εισαγωγή",
    exportedSheetTransactions: "Συναλλαγές",
    exportedSheetMonthly: "Μηνιαία Σύνοψη",
    exportedSheetCategory: "Ανάλυση Κατηγοριών",
    colDate: "Ημερομηνία",
    colType: "Τύπος",
    colCategory: "Κατηγορία",
    colAmount: "Ποσό",
    colNote: "Σημείωση",
    colCategoryId: "Κωδικός Κατηγορίας",
    colTypeCode: "Κωδικός Τύπου",
    colMonth: "Μήνας",
    confirmYes: "Ναι, διαγραφή",
    confirmNo: "Ακύρωση",
    cancel: "Ακύρωση",
    save: "Αποθήκευση",
    importSuccess: "Έγινε εισαγωγή {{count}} συναλλαγών.",
    importError: "Δεν ήταν δυνατή η ανάγνωση του αρχείου. Δοκιμάστε πρώτα να κάνετε εξαγωγή από αυτήν την εφαρμογή και έπειτα εισαγωγή του ίδιου αρχείου.",
    cat_rent: "Ενοίκιο & Στέγαση",
    cat_food: "Φαγητό & Ψώνια",
    cat_transport: "Μετακινήσεις",
    cat_utilities: "Λογαριασμοί",
    cat_health: "Υγεία",
    cat_fun: "Ψυχαγωγία",
    cat_shopping: "Αγορές",
    cat_other_exp: "Άλλο",
    cat_salary: "Μισθός & Εισόδημα",
    cat_savings: "Αποταμιεύσεις & Τόκοι",
    cat_other_inc: "Άλλο Εισόδημα",
  },
};

function uid() {
  return (crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random().toString(16).slice(2)}`);
}
function pad2(n) { return String(n).padStart(2, "0"); }
function dateStr(y, mIdx, d) { return `${y}-${pad2(mIdx + 1)}-${pad2(d)}`; }
function fmtMoney(n, locale) {
  const sign = n < 0 ? "-" : "";
  return `${sign}€${Math.abs(n).toLocaleString(locale, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}
function daysInMonth(y, mIdx) { return new Date(y, mIdx + 1, 0).getDate(); }
function firstWeekday(y, mIdx) { return new Date(y, mIdx, 1).getDay(); }

export default function Ledger() {
  const [loading, setLoading] = useState(true);
  const [saveError, setSaveError] = useState(null);
  const [transactions, setTransactions] = useState([]);
  const [year, setYear] = useState(new Date().getFullYear());
  const [monthIdx, setMonthIdx] = useState(new Date().getMonth());
  const [tab, setTab] = useState("ledger");
  const [openDay, setOpenDay] = useState(null);
  const [editingId, setEditingId] = useState(null);
  const [form, setForm] = useState({ type: "expense", categoryId: "food", amount: "", note: "" });

  const [lang, setLang] = useState("en");
  const [importMessage, setImportMessage] = useState(null);
  const fileInputRef = useRef(null);
  const firstLoadRef = useRef(true);

  // Serializes every write so rapid clicks can never race and clobber
  // each other in storage.
  const writeQueueRef = useRef(Promise.resolve());
  const queuedSet = useCallback((key, value, shared = false) => {
    const p = writeQueueRef.current.then(() => storage.set(key, value, shared).catch(() => null));
    writeQueueRef.current = p;
    return p;
  }, []);

  const tr = useCallback((key, vars) => {
    let s = (TRANSLATIONS[lang] && TRANSLATIONS[lang][key]) ?? TRANSLATIONS.en[key] ?? key;
    if (vars) for (const k in vars) s = s.replace(new RegExp(`{{${k}}}`, "g"), vars[k]);
    return s;
  }, [lang]);
  const locale = lang === "el" ? "el-GR" : "en-US";

  // Bootstrap: language + transactions. Only the very first load shows
  // the full-screen spinner.
  useEffect(() => {
    (async () => {
      try {
        const [langRes, txRes] = await Promise.all([
          storage.get(LANG_KEY, false).catch(() => null),
          storage.get(TX_KEY, false).catch(() => null),
        ]);

        if (langRes && langRes.value) {
          try { setLang(JSON.parse(langRes.value)); } catch (e) {}
        }
        if (txRes && txRes.value) {
          try { setTransactions(JSON.parse(txRes.value)); } catch (e) {}
        }
      } finally {
        setLoading(false);
        firstLoadRef.current = false;
      }
    })();
  }, []);

  const persist = useCallback(async (list) => {
    setTransactions(list);
    try {
      const res = await queuedSet(TX_KEY, JSON.stringify(list), false);
      if (!res) setSaveError(tr("saveError"));
      else setSaveError(null);
    } catch (e) {
      setSaveError(tr("saveError"));
    }
  }, [tr, queuedSet]);

  const changeLang = async (l) => {
    setLang(l);
    await queuedSet(LANG_KEY, JSON.stringify(l), false);
  };

  const resetForm = () => { setForm({ type: "expense", categoryId: "food", amount: "", note: "" }); setEditingId(null); };

  const openDayModal = (dstr) => { setOpenDay(dstr); resetForm(); };
  const closeDayModal = () => { setOpenDay(null); resetForm(); };

  const submitForm = () => {
    const amt = parseFloat(form.amount);
    if (!amt || amt <= 0 || !openDay) return;
    if (editingId) {
      const list = transactions.map((t) =>
        t.id === editingId ? { ...t, type: form.type, categoryId: form.categoryId, amount: amt, note: form.note.trim() } : t
      );
      persist(list);
    } else {
      const tx = { id: uid(), date: openDay, type: form.type, categoryId: form.categoryId, amount: amt, note: form.note.trim() };
      persist([...transactions, tx]);
    }
    resetForm();
  };

  const startEdit = (t) => {
    setEditingId(t.id);
    setForm({ type: t.type, categoryId: t.categoryId, amount: String(t.amount), note: t.note || "" });
  };

  const deleteTx = (id) => {
    persist(transactions.filter((t) => t.id !== id));
    if (editingId === id) resetForm();
  };

  const txByDate = useMemo(() => {
    const map = {};
    for (const t of transactions) {
      (map[t.date] = map[t.date] || []).push(t);
    }
    return map;
  }, [transactions]);

  const monthTx = useMemo(
    () => transactions.filter((t) => t.date.startsWith(`${year}-${pad2(monthIdx + 1)}`)),
    [transactions, year, monthIdx]
  );

  const monthTotals = useMemo(() => {
    let income = 0, expense = 0;
    for (const t of monthTx) (t.type === "income" ? (income += t.amount) : (expense += t.amount));
    return { income, expense, net: income - expense };
  }, [monthTx]);

  const yearTx = useMemo(() => transactions.filter((t) => t.date.startsWith(`${year}-`)), [transactions, year]);
  const yearTotals = useMemo(() => {
    let income = 0, expense = 0;
    for (const t of yearTx) (t.type === "income" ? (income += t.amount) : (expense += t.amount));
    return { income, expense, net: income - expense };
  }, [yearTx]);

  const perMonth = useMemo(() => {
    return MONTH_SHORT[lang].map((label, idx) => {
      let income = 0, expense = 0;
      for (const t of transactions) {
        if (t.date.startsWith(`${year}-${pad2(idx + 1)}`)) (t.type === "income" ? (income += t.amount) : (expense += t.amount));
      }
      return { label, income, expense, net: income - expense };
    });
  }, [transactions, year, lang]);

  const weekly = useMemo(() => {
    const buckets = [1, 2, 3, 4, 5].map((n) => ({
      label: `${tr("weekLabel")} ${n}`, start: (n - 1) * 7 + 1, end: n === 5 ? 31 : n * 7, expense: 0, income: 0,
    }));
    for (const t of monthTx) {
      const day = parseInt(t.date.slice(-2), 10);
      const b = buckets.find((b) => day >= b.start && day <= b.end);
      if (b) (t.type === "income" ? (b.income += t.amount) : (b.expense += t.amount));
    }
    return buckets.filter((b) => b.start <= daysInMonth(year, monthIdx));
  }, [monthTx, year, monthIdx, tr]);

  const categoryBreakdown = useMemo(() => {
    const totals = {};
    for (const t of monthTx) {
      if (t.type !== "expense") continue;
      totals[t.categoryId] = (totals[t.categoryId] || 0) + t.amount;
    }
    return Object.entries(totals)
      .map(([id, value]) => ({ id, name: tr(`cat_${id}`) || id, value, color: catById(id)?.color || "#999" }))
      .sort((a, b) => b.value - a.value);
  }, [monthTx, tr]);

  const grid = useMemo(() => {
    const total = daysInMonth(year, monthIdx);
    const lead = firstWeekday(year, monthIdx);
    const cells = [];
    for (let i = 0; i < lead; i++) cells.push(null);
    for (let d = 1; d <= total; d++) cells.push(d);
    return cells;
  }, [year, monthIdx]);

  const today = new Date();
  const isToday = (d) => d && today.getFullYear() === year && today.getMonth() === monthIdx && today.getDate() === d;

  // ---- Export to Excel ----
  const exportXlsx = () => {
    const wb = XLSX.utils.book_new();

    const txRows = [...transactions].sort((a, b) => a.date.localeCompare(b.date)).map((t) => ({
      [tr("colDate")]: t.date,
      [tr("colType")]: t.type === "income" ? tr("received") : tr("toggleSpent"),
      [tr("colCategory")]: tr(`cat_${t.categoryId}`),
      [tr("colAmount")]: t.amount,
      [tr("colNote")]: t.note || "",
      [tr("colCategoryId")]: t.categoryId,
      [tr("colTypeCode")]: t.type,
    }));
    const wsTx = XLSX.utils.json_to_sheet(txRows.length ? txRows : [{ [tr("colDate")]: "" }]);
    XLSX.utils.book_append_sheet(wb, wsTx, tr("exportedSheetTransactions").slice(0, 31));

    const monthlyRows = perMonth.map((m) => ({
      [tr("colMonth")]: m.label,
      [tr("income")]: m.income,
      [tr("spent")]: m.expense,
      [tr("net")]: m.net,
    }));
    const wsMonthly = XLSX.utils.json_to_sheet(monthlyRows);
    XLSX.utils.book_append_sheet(wb, wsMonthly, tr("exportedSheetMonthly").slice(0, 31));

    const catTotals = {};
    for (const t of transactions) {
      if (t.type !== "expense") continue;
      catTotals[t.categoryId] = (catTotals[t.categoryId] || 0) + t.amount;
    }
    const catRows = Object.entries(catTotals).sort((a, b) => b[1] - a[1]).map(([id, value]) => ({
      [tr("colCategory")]: tr(`cat_${id}`),
      [tr("colAmount")]: value,
    }));
    const wsCat = XLSX.utils.json_to_sheet(catRows.length ? catRows : [{ [tr("colCategory")]: "" }]);
    XLSX.utils.book_append_sheet(wb, wsCat, tr("exportedSheetCategory").slice(0, 31));

    XLSX.writeFile(wb, "ledger_export.xlsx");
  };

  // ---- Import from Excel ----
  const triggerImport = () => fileInputRef.current?.click();

  const handleImportFile = async (e) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const sheetName =
        wb.SheetNames.find((n) => n === TRANSLATIONS.en.exportedSheetTransactions || n === TRANSLATIONS.el.exportedSheetTransactions) ||
        wb.SheetNames[0];
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: "" });

      const catLookup = {};
      for (const c of CATEGORIES) {
        catLookup[TRANSLATIONS.en[`cat_${c.id}`].toLowerCase()] = c.id;
        catLookup[TRANSLATIONS.el[`cat_${c.id}`].toLowerCase()] = c.id;
        catLookup[c.id.toLowerCase()] = c.id;
      }
      const typeLookup = {
        [TRANSLATIONS.en.received.toLowerCase()]: "income",
        [TRANSLATIONS.el.received.toLowerCase()]: "income",
        [TRANSLATIONS.en.toggleSpent.toLowerCase()]: "expense",
        [TRANSLATIONS.el.toggleSpent.toLowerCase()]: "expense",
        income: "income",
        expense: "expense",
      };

      const findVal = (row, keys) => {
        for (const k of keys) if (row[k] !== undefined && row[k] !== "") return row[k];
        return undefined;
      };
      const dateKeys = [TRANSLATIONS.en.colDate, TRANSLATIONS.el.colDate, "Date"];
      const typeKeys = [TRANSLATIONS.en.colTypeCode, TRANSLATIONS.el.colTypeCode, TRANSLATIONS.en.colType, TRANSLATIONS.el.colType, "Type"];
      const catIdKeys = [TRANSLATIONS.en.colCategoryId, TRANSLATIONS.el.colCategoryId];
      const catKeys = [TRANSLATIONS.en.colCategory, TRANSLATIONS.el.colCategory, "Category"];
      const amtKeys = [TRANSLATIONS.en.colAmount, TRANSLATIONS.el.colAmount, "Amount"];
      const noteKeys = [TRANSLATIONS.en.colNote, TRANSLATIONS.el.colNote, "Note"];

      const imported = [];
      for (const row of rows) {
        const rawDate = findVal(row, dateKeys);
        const rawAmt = findVal(row, amtKeys);
        if (rawDate === undefined || rawDate === "" || rawAmt === undefined) continue;

        let dateStr2 = null;
        if (typeof rawDate === "number") {
          const d = XLSX.SSF.parse_date_code(rawDate);
          if (d) dateStr2 = `${d.y}-${pad2(d.m)}-${pad2(d.d)}`;
        } else {
          const s = String(rawDate).trim();
          const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
          if (m) dateStr2 = `${m[1]}-${m[2]}-${m[3]}`;
        }
        if (!dateStr2) continue;

        const rawType = String(findVal(row, typeKeys) || "").toLowerCase().trim();
        const type = typeLookup[rawType];
        if (!type) continue;

        let categoryId = findVal(row, catIdKeys);
        if (!categoryId || !catById(categoryId)) {
          const rawCat = String(findVal(row, catKeys) || "").toLowerCase().trim();
          categoryId = catLookup[rawCat];
        }
        if (!categoryId || !catById(categoryId)) categoryId = type === "income" ? "other_inc" : "other_exp";

        const amount = parseFloat(rawAmt);
        if (!amount || amount <= 0) continue;

        imported.push({
          id: uid(), date: dateStr2, type, categoryId, amount,
          note: String(findVal(row, noteKeys) || "").trim(),
        });
      }

      if (!imported.length) {
        setImportMessage({ type: "error", text: tr("importError") });
        return;
      }
      await persist([...transactions, ...imported]);
      setImportMessage({ type: "success", text: tr("importSuccess", { count: imported.length }) });
    } catch (err) {
      setImportMessage({ type: "error", text: tr("importError") });
    }
  };

  if (loading) {
    return (
      <div className="ledger-app ledger-loading">
        <Loader2 className="spin" size={28} />
        <span>{tr("loading")}</span>
        <Style />
      </div>
    );
  }

  const openDayLabel = openDay ? new Date(openDay + "T00:00:00").toLocaleDateString(locale, { weekday: "long", month: "long", day: "numeric" }) : "";
  const dayEntries = openDay ? (txByDate[openDay] || []) : [];

  return (
    <div className="ledger-app">
      <Style />
      <header className="cover">
        <div className="cover-title">
          <BookOpen size={22} />
          <div>
            <h1>{tr("appTitle")}</h1>
            <p>{tr("appSubtitle")}</p>
          </div>
        </div>
        <div className="year-switch">
          <button aria-label="Previous year" onClick={() => setYear((y) => y - 1)}><ChevronLeft size={18} /></button>
          <span className="year-num">{year}</span>
          <button aria-label="Next year" onClick={() => setYear((y) => y + 1)}><ChevronRight size={18} /></button>
        </div>
        <div className="year-total">
          <span className="year-total-label">{tr("yearTotal")}</span>
          <span className={`year-total-num ${yearTotals.net >= 0 ? "pos" : "neg"}`}>{fmtMoney(yearTotals.net, locale)}</span>
        </div>
      </header>

      <div className="controls-bar">
        <div className="control-group">
          <Globe size={14} />
          <select value={lang} onChange={(e) => changeLang(e.target.value)}>
            <option value="en">English</option>
            <option value="el">Ελληνικά</option>
          </select>
        </div>

        <div className="control-group">
          <button className="mini-btn" onClick={exportXlsx}><Download size={13} /> {tr("export")}</button>
          <button className="mini-btn" onClick={triggerImport}><Upload size={13} /> {tr("import")}</button>
          <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={handleImportFile} />
        </div>
      </div>

      {saveError && <div className="save-banner">{saveError}</div>}
      {importMessage && (
        <div className={`import-banner ${importMessage.type}`}>
          <span>{importMessage.text}</span>
          <button className="icon-btn tiny" onClick={() => setImportMessage(null)}><X size={12} /></button>
        </div>
      )}

      <nav className="tabbar">
        <button className={tab === "ledger" ? "active" : ""} onClick={() => setTab("ledger")}>
          <BookOpen size={16} /> {tr("tabLedger")}
        </button>
        <button className={tab === "insights" ? "active" : ""} onClick={() => setTab("insights")}>
          <PieIcon size={16} /> {tr("tabInsights")}
        </button>
      </nav>

      <div className="body">
        <div className="month-tabs">
          {MONTH_SHORT[lang].map((m, idx) => {
            let net = 0;
            for (const t of transactions) if (t.date.startsWith(`${year}-${pad2(idx + 1)}`)) net += t.type === "income" ? t.amount : -t.amount;
            return (
              <button key={m} className={`month-tab ${idx === monthIdx ? "active" : ""}`} onClick={() => setMonthIdx(idx)}>
                <span className="mt-name">{m}</span>
                <span className={`mt-net ${net >= 0 ? "pos" : "neg"}`}>{net === 0 ? "—" : fmtMoney(net, locale)}</span>
              </button>
            );
          })}
        </div>

        <main className="page">
          {tab === "ledger" ? (
            <>
              <div className="page-head">
                <h2>{MONTH_NAMES[lang][monthIdx]} {year}</h2>
                <div className="summary-chips">
                  <span className="chip income">{tr("income")} {fmtMoney(monthTotals.income, locale)}</span>
                  <span className="chip expense">{tr("spent")} {fmtMoney(monthTotals.expense, locale)}</span>
                  <span className={`chip net ${monthTotals.net >= 0 ? "pos" : "neg"}`}>{tr("net")} {fmtMoney(monthTotals.net, locale)}</span>
                </div>
              </div>

              <div className="weekday-row">
                {WEEKDAYS[lang].map((w, i) => <div key={i} className="weekday">{w}</div>)}
              </div>
              <div className="calendar">
                {grid.map((d, i) => {
                  if (!d) return <div key={i} className="cell empty" />;
                  const ds = dateStr(year, monthIdx, d);
                  const entries = txByDate[ds] || [];
                  const net = entries.reduce((s, t) => s + (t.type === "income" ? t.amount : -t.amount), 0);
                  return (
                    <button key={i} className={`cell ${isToday(d) ? "today" : ""}`} onClick={() => openDayModal(ds)}>
                      <span className="cell-day">{d}</span>
                      {entries.length > 0 && (
                        <>
                          <span className={`cell-amount ${net >= 0 ? "pos" : "neg"}`}>{fmtMoney(net, locale)}</span>
                          <span className="cell-dots">
                            {entries.slice(0, 4).map((t, idx2) => (
                              <span key={idx2} className="dot" style={{ background: catById(t.categoryId)?.color }} />
                            ))}
                          </span>
                        </>
                      )}
                    </button>
                  );
                })}
              </div>
            </>
          ) : (
            <Insights
              monthLabel={`${MONTH_NAMES[lang][monthIdx]} ${year}`}
              weekly={weekly}
              categoryBreakdown={categoryBreakdown}
              perMonth={perMonth}
              tr={tr}
              locale={locale}
            />
          )}
        </main>
      </div>

      {openDay && (
        <div className="modal-backdrop" onClick={closeDayModal}>
          <div className="modal" onClick={(e) => e.stopPropagation()}>
            <div className="modal-head">
              <h3>{openDayLabel}</h3>
              <button className="icon-btn" onClick={closeDayModal}><X size={18} /></button>
            </div>

            {dayEntries.length > 0 && (
              <ul className="entry-list">
                {dayEntries.map((t) => {
                  const cat = catById(t.categoryId);
                  const Icon = cat?.Icon || MoreHorizontal;
                  return (
                    <li key={t.id} className="entry">
                      <span className="entry-icon" style={{ background: cat?.color }}><Icon size={14} color="#fff" /></span>
                      <span className="entry-info">
                        <span className="entry-cat">{tr(`cat_${t.categoryId}`)}</span>
                        {t.note && <span className="entry-note">{t.note}</span>}
                      </span>
                      <span className={`entry-amt ${t.type === "income" ? "pos" : "neg"}`}>
                        {t.type === "income" ? "+" : "-"}{fmtMoney(t.amount, locale).replace("-", "")}
                      </span>
                      <span className="entry-actions">
                        <button className="icon-btn" onClick={() => startEdit(t)}><Pencil size={14} /></button>
                        <button className="icon-btn" onClick={() => deleteTx(t.id)}><Trash2 size={14} /></button>
                      </span>
                    </li>
                  );
                })}
              </ul>
            )}

            <div className="form">
              <div className="type-toggle">
                <button className={form.type === "expense" ? "active exp" : ""} onClick={() => setForm((f) => ({ ...f, type: "expense", categoryId: "food" }))}>{tr("toggleSpent")}</button>
                <button className={form.type === "income" ? "active inc" : ""} onClick={() => setForm((f) => ({ ...f, type: "income", categoryId: "salary" }))}>{tr("received")}</button>
              </div>

              <div className="cat-grid">
                {CATEGORIES.filter((c) => c.type === form.type).map((c) => {
                  const Icon = c.Icon;
                  return (
                    <button key={c.id} className={`cat-pill ${form.categoryId === c.id ? "active" : ""}`}
                      style={form.categoryId === c.id ? { background: c.color, borderColor: c.color } : {}}
                      onClick={() => setForm((f) => ({ ...f, categoryId: c.id }))}>
                      <Icon size={14} /> {tr(`cat_${c.id}`)}
                    </button>
                  );
                })}
              </div>

              <div className="form-row">
                <label>
                  {tr("amount")}
                  <div className="amount-input">
                    <span>€</span>
                    <input type="number" min="0" step="0.01" placeholder="0.00" value={form.amount}
                      onChange={(e) => setForm((f) => ({ ...f, amount: e.target.value }))} />
                  </div>
                </label>
                <label className="grow">
                  {tr("reason")}
                  <input type="text" placeholder={tr("reasonPlaceholder")} value={form.note}
                    onChange={(e) => setForm((f) => ({ ...f, note: e.target.value }))} />
                </label>
              </div>

              <div className="form-actions">
                {editingId && <button className="ghost" onClick={resetForm}>{tr("cancelEdit")}</button>}
                <button className="primary" onClick={submitForm}>
                  <Plus size={16} /> {editingId ? tr("saveChanges") : tr("addEntry")}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}

function Insights({ monthLabel, weekly, categoryBreakdown, perMonth, tr, locale }) {
  const expenseTotal = categoryBreakdown.reduce((s, c) => s + c.value, 0);
  return (
    <div className="insights">
      <section className="panel">
        <h3>{tr("spendingByWeek")} — {monthLabel}</h3>
        <ResponsiveContainer width="100%" height={220}>
          <BarChart data={weekly} margin={{ top: 8, right: 12, left: -12, bottom: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#DDD3B8" vertical={false} />
            <XAxis dataKey="label" tick={{ fill: "#6B6152", fontSize: 12, fontFamily: "Inter, sans-serif" }} axisLine={{ stroke: "#DDD3B8" }} tickLine={false} />
            <YAxis tick={{ fill: "#6B6152", fontSize: 11, fontFamily: "Inter, sans-serif" }} axisLine={false} tickLine={false} width={54} />
            <Tooltip formatter={(v) => fmtMoney(v, locale)} contentStyle={{ fontFamily: "Inter, sans-serif", borderRadius: 8, border: "1px solid #DDD3B8" }} />
            <Bar dataKey="expense" name={tr("spent")} fill="#A6432C" radius={[4, 4, 0, 0]} />
            <Bar dataKey="income" name={tr("received")} fill="#2F7D53" radius={[4, 4, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </section>

      <section className="panel">
        <h3>{tr("whereItWent")} — {monthLabel}</h3>
        {categoryBreakdown.length === 0 ? (
          <p className="empty-note">{tr("noSpending")}</p>
        ) : (
          <div className="pie-wrap">
            <ResponsiveContainer width="100%" height={220}>
              <PieChart>
                <Pie data={categoryBreakdown} dataKey="value" nameKey="name" innerRadius={55} outerRadius={85} paddingAngle={2}>
                  {categoryBreakdown.map((c) => <Cell key={c.id} fill={c.color} />)}
                </Pie>
                <Tooltip formatter={(v) => fmtMoney(v, locale)} contentStyle={{ fontFamily: "Inter, sans-serif", borderRadius: 8, border: "1px solid #DDD3B8" }} />
              </PieChart>
            </ResponsiveContainer>
            <ul className="legend">
              {categoryBreakdown.map((c) => (
                <li key={c.id}>
                  <span className="swatch" style={{ background: c.color }} />
                  <span className="leg-name">{c.name}</span>
                  <span className="leg-val">{fmtMoney(c.value, locale)}</span>
                  <span className="leg-pct">{expenseTotal ? Math.round((c.value / expenseTotal) * 100) : 0}%</span>
                </li>
              ))}
            </ul>
          </div>
        )}
      </section>

      <section className="panel wide">
        <h3>{tr("yearAtGlance")}</h3>
        <ResponsiveContainer width="100%" height={240}>
          <BarChart data={perMonth} margin={{ top: 8, right: 12, left: -12, bottom: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#DDD3B8" vertical={false} />
            <XAxis dataKey="label" tick={{ fill: "#6B6152", fontSize: 12, fontFamily: "Inter, sans-serif" }} axisLine={{ stroke: "#DDD3B8" }} tickLine={false} />
            <YAxis tick={{ fill: "#6B6152", fontSize: 11, fontFamily: "Inter, sans-serif" }} axisLine={false} tickLine={false} width={54} />
            <Tooltip formatter={(v) => fmtMoney(v, locale)} contentStyle={{ fontFamily: "Inter, sans-serif", borderRadius: 8, border: "1px solid #DDD3B8" }} />
            <Legend wrapperStyle={{ fontFamily: "Inter, sans-serif", fontSize: 12 }} />
            <Bar dataKey="income" name={tr("received")} fill="#2F7D53" radius={[4, 4, 0, 0]} />
            <Bar dataKey="expense" name={tr("spent")} fill="#A6432C" radius={[4, 4, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </section>
    </div>
  );
}

function Style() {
  return (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600;9..144,700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap');

      .ledger-app {
        --ink:#1F2A3C; --paper:#F6F1E1; --cover:#1B3B2F; --cover-light:#25493A;
        --gold:#B8863B; --income:#2F7D53; --expense:#A6432C; --line:#DDD3B8; --muted:#6B6152;
        font-family:'Inter',sans-serif; color:var(--ink); background:var(--paper);
        border-radius:14px; overflow:hidden; max-width:920px; margin:0 auto;
        box-shadow:0 1px 3px rgba(0,0,0,0.08); border:1px solid var(--line);
      }
      .ledger-loading { display:flex; flex-direction:column; align-items:center; justify-content:center; gap:10px; padding:60px 20px; color:var(--muted); }
      .spin { animation: spin 1s linear infinite; }
      @keyframes spin { to { transform: rotate(360deg); } }

      .cover { background:linear-gradient(160deg,var(--cover),var(--cover-light)); color:#F3EEDD; padding:22px 24px; display:flex; align-items:center; justify-content:space-between; gap:16px; flex-wrap:wrap; }
      .cover-title { display:flex; align-items:center; gap:12px; }
      .cover-title h1 { font-family:'Fraunces',serif; font-weight:600; font-size:1.5rem; margin:0; letter-spacing:0.2px; }
      .cover-title p { margin:2px 0 0; font-size:0.78rem; color:#C9C1A4; }
      .year-switch { display:flex; align-items:center; gap:10px; }
      .year-switch button { background:rgba(255,255,255,0.1); border:1px solid rgba(255,255,255,0.2); color:#F3EEDD; border-radius:8px; padding:5px; cursor:pointer; display:flex; }
      .year-switch button:hover { background:rgba(255,255,255,0.2); }
      .year-num { font-family:'IBM Plex Mono',monospace; font-size:1rem; min-width:44px; text-align:center; }
      .year-total { text-align:right; }
      .year-total-label { display:block; font-size:0.7rem; color:#C9C1A4; text-transform:uppercase; letter-spacing:0.08em; }
      .year-total-num { font-family:'IBM Plex Mono',monospace; font-size:1.3rem; font-weight:600; }
      .year-total-num.pos { color:#8FD4AC; } .year-total-num.neg { color:#E2917E; }

      .controls-bar { display:flex; align-items:center; justify-content:space-between; gap:10px; flex-wrap:wrap; padding:10px 20px; background:#FBF8EE; border-bottom:1px solid var(--line); }
      .control-group { display:flex; align-items:center; gap:6px; color:var(--muted); }
      .control-group select { border:1px solid var(--line); border-radius:7px; padding:5px 8px; font-family:'Inter',sans-serif; font-size:0.78rem; background:#fff; color:var(--ink); cursor:pointer; max-width:150px; }
      .icon-btn.small { padding:5px; }
      .icon-btn.tiny { padding:3px; background:transparent; }
      .icon-btn.tiny:hover { background:rgba(0,0,0,0.06); }
      .confirm-inline { display:flex; gap:4px; align-items:center; }
      .mini-danger { background:var(--expense); color:#fff; border:none; border-radius:6px; padding:5px 9px; font-size:0.72rem; font-weight:600; cursor:pointer; }
      .mini-ghost { background:transparent; border:1px solid var(--line); border-radius:6px; padding:5px 9px; font-size:0.72rem; cursor:pointer; color:var(--muted); }
      .mini-btn { display:flex; align-items:center; gap:5px; border:1px solid var(--line); background:#fff; border-radius:7px; padding:6px 11px; font-size:0.76rem; font-weight:600; color:var(--ink); cursor:pointer; }
      .mini-btn:hover { background:#F1ECDB; }

      .save-banner { background:#F6E3D6; color:#7A3B1E; font-size:0.8rem; padding:8px 20px; }
      .import-banner { display:flex; align-items:center; justify-content:space-between; gap:10px; font-size:0.8rem; padding:7px 20px; }
      .import-banner.success { background:#E4F2E9; color:#1F6E5C; }
      .import-banner.error { background:#F6E5DF; color:#7A3B1E; }

      .tabbar { display:flex; gap:4px; padding:10px 20px 0; background:var(--paper); }
      .tabbar button { display:flex; align-items:center; gap:6px; border:none; background:transparent; color:var(--muted); font-family:'Inter',sans-serif; font-weight:600; font-size:0.85rem; padding:8px 14px; border-radius:8px 8px 0 0; cursor:pointer; }
      .tabbar button.active { background:#fff; color:var(--cover); box-shadow:0 -1px 0 var(--line) inset; }

      .body { display:flex; background:#fff; }
      .month-tabs { display:flex; flex-direction:column; width:112px; flex-shrink:0; background:#FBF8EE; border-right:1px solid var(--line); padding:10px 0; }
      .month-tab { display:flex; flex-direction:column; align-items:flex-start; gap:2px; border:none; background:transparent; padding:9px 14px; cursor:pointer; border-left:3px solid transparent; }
      .month-tab .mt-name { font-family:'Fraunces',serif; font-size:0.85rem; color:var(--ink); }
      .month-tab .mt-net { font-family:'IBM Plex Mono',monospace; font-size:0.68rem; color:var(--muted); }
      .month-tab .mt-net.pos { color:var(--income); } .month-tab .mt-net.neg { color:var(--expense); }
      .month-tab.active { background:#fff; border-left-color:var(--gold); }
      .month-tab.active .mt-name { color:var(--cover); font-weight:600; }
      .month-tab:hover { background:#F2ECD8; }

      .page { flex:1; padding:20px 22px 26px; min-width:0; }
      .page-head { display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:10px; margin-bottom:14px; }
      .page-head h2 { font-family:'Fraunces',serif; font-size:1.25rem; margin:0; }
      .summary-chips { display:flex; gap:8px; flex-wrap:wrap; }
      .chip { font-family:'IBM Plex Mono',monospace; font-size:0.72rem; padding:4px 10px; border-radius:999px; background:#F1ECDB; color:var(--muted); }
      .chip.income { color:var(--income); } .chip.expense { color:var(--expense); }
      .chip.net.pos { background:#E4F2E9; color:var(--income); } .chip.net.neg { background:#F6E5DF; color:var(--expense); }

      .weekday-row { display:grid; grid-template-columns:repeat(7,1fr); margin-bottom:4px; }
      .weekday { text-align:center; font-size:0.68rem; color:var(--muted); text-transform:uppercase; letter-spacing:0.06em; padding-bottom:4px; }
      .calendar { display:grid; grid-template-columns:repeat(7,1fr); gap:5px; }
      .cell { aspect-ratio:1/1; border:1px solid var(--line); border-radius:8px; background:#fff; display:flex; flex-direction:column; align-items:flex-start; padding:6px; cursor:pointer; position:relative; min-height:56px; }
      .cell:hover { border-color:var(--gold); background:#FBF7EA; }
      .cell.empty { border:none; background:transparent; cursor:default; }
      .cell.today { border-color:var(--cover); border-width:2px; }
      .cell-day { font-family:'IBM Plex Mono',monospace; font-size:0.72rem; color:var(--muted); }
      .cell-amount { font-family:'IBM Plex Mono',monospace; font-size:0.66rem; font-weight:600; margin-top:auto; }
      .cell-amount.pos { color:var(--income); } .cell-amount.neg { color:var(--expense); }
      .cell-dots { display:flex; gap:2px; margin-top:2px; }
      .dot { width:5px; height:5px; border-radius:50%; }

      .insights { display:grid; grid-template-columns:1fr 1fr; gap:16px; }
      .panel { border:1px solid var(--line); border-radius:10px; padding:14px 16px; background:#fff; }
      .panel.wide { grid-column:1 / -1; }
      .panel h3 { font-family:'Fraunces',serif; font-size:1rem; margin:0 0 8px; }
      .empty-note { color:var(--muted); font-size:0.85rem; }
      .pie-wrap { display:flex; align-items:center; gap:10px; flex-wrap:wrap; }
      .legend { list-style:none; margin:0; padding:0; flex:1; min-width:160px; }
      .legend li { display:flex; align-items:center; gap:8px; font-size:0.78rem; padding:3px 0; }
      .swatch { width:9px; height:9px; border-radius:50%; flex-shrink:0; }
      .leg-name { flex:1; color:var(--ink); }
      .leg-val { font-family:'IBM Plex Mono',monospace; color:var(--muted); }
      .leg-pct { font-family:'IBM Plex Mono',monospace; color:var(--muted); width:32px; text-align:right; }

      .modal-backdrop { position:fixed; inset:0; background:rgba(27,31,23,0.5); display:flex; align-items:center; justify-content:center; padding:16px; z-index:50; }
      .modal { background:#fff; border-radius:14px; width:100%; max-width:440px; max-height:88vh; overflow-y:auto; padding:18px 20px 20px; }
      .modal.small { max-width:360px; }
      .modal-head { display:flex; align-items:center; justify-content:space-between; margin-bottom:10px; }
      .modal-head h3 { font-family:'Fraunces',serif; font-size:1.05rem; margin:0; }
      .icon-btn { border:none; background:#F1ECDB; border-radius:7px; padding:6px; cursor:pointer; color:var(--ink); display:flex; }
      .icon-btn:hover { background:var(--line); }

      .entry-list { list-style:none; margin:0 0 14px; padding:0; display:flex; flex-direction:column; gap:6px; }
      .entry { display:flex; align-items:center; gap:8px; border:1px solid var(--line); border-radius:9px; padding:7px 9px; }
      .entry-icon { width:26px; height:26px; border-radius:50%; display:flex; align-items:center; justify-content:center; flex-shrink:0; }
      .entry-info { display:flex; flex-direction:column; flex:1; min-width:0; }
      .entry-cat { font-size:0.8rem; font-weight:600; }
      .entry-note { font-size:0.72rem; color:var(--muted); overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
      .entry-amt { font-family:'IBM Plex Mono',monospace; font-size:0.82rem; font-weight:600; }
      .entry-amt.pos { color:var(--income); } .entry-amt.neg { color:var(--expense); }
      .entry-actions { display:flex; gap:4px; }

      .form { border-top:1px solid var(--line); padding-top:12px; }
      .type-toggle { display:flex; gap:6px; margin-bottom:10px; }
      .type-toggle button { flex:1; border:1px solid var(--line); background:#fff; padding:8px; border-radius:8px; cursor:pointer; font-weight:600; font-size:0.82rem; color:var(--muted); }
      .type-toggle button.active.exp { background:#F6E5DF; color:var(--expense); border-color:var(--expense); }
      .type-toggle button.active.inc { background:#E4F2E9; color:var(--income); border-color:var(--income); }

      .cat-grid { display:flex; flex-wrap:wrap; gap:6px; margin-bottom:12px; }
      .cat-pill { display:flex; align-items:center; gap:5px; border:1px solid var(--line); background:#fff; color:var(--ink); font-size:0.72rem; padding:6px 10px; border-radius:999px; cursor:pointer; }
      .cat-pill.active { color:#fff; }

      .form-row { display:flex; gap:10px; margin-bottom:12px; }
      .form-row label { display:flex; flex-direction:column; gap:4px; font-size:0.72rem; color:var(--muted); font-weight:600; }
      .form-row label.grow { flex:1; }
      .amount-input { display:flex; align-items:center; border:1px solid var(--line); border-radius:8px; padding:0 8px; width:110px; }
      .amount-input span { color:var(--muted); font-family:'IBM Plex Mono',monospace; }
      .amount-input input { border:none; outline:none; padding:8px 4px; width:100%; font-family:'IBM Plex Mono',monospace; font-size:0.9rem; }
      .form-row input[type=text] { border:1px solid var(--line); border-radius:8px; padding:8px 10px; font-family:'Inter',sans-serif; font-size:0.85rem; outline:none; width:100%; box-sizing:border-box; }
      .form-row input:focus, .amount-input:focus-within { border-color:var(--cover); }

      .form-actions { display:flex; justify-content:flex-end; gap:8px; }
      .form-actions .primary { display:flex; align-items:center; gap:6px; background:var(--cover); color:#fff; border:none; padding:9px 16px; border-radius:8px; font-weight:600; font-size:0.85rem; cursor:pointer; }
      .form-actions .primary:hover { background:var(--cover-light); }
      .form-actions .ghost { background:transparent; border:1px solid var(--line); color:var(--muted); padding:9px 14px; border-radius:8px; cursor:pointer; font-size:0.85rem; }

      button:focus-visible, .cell:focus-visible, select:focus-visible { outline:2px solid var(--gold); outline-offset:2px; }

      @media (max-width:640px) {
        .body { flex-direction:column; }
        .month-tabs { flex-direction:row; width:100%; overflow-x:auto; border-right:none; border-bottom:1px solid var(--line); }
        .month-tab { border-left:none; border-bottom:3px solid transparent; flex-shrink:0; }
        .month-tab.active { border-bottom-color:var(--gold); }
        .insights { grid-template-columns:1fr; }
        .cover { padding:18px; }
        .year-total { text-align:left; }
        .controls-bar { justify-content:flex-start; }
      }
    `}</style>
  );
}