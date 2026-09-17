/* ═══════════════════════════════════════════════════════════════════
   ADMIN.JS — جمعية إرث وحضارة بالقريات
   المحرك البرمجي للوحة إدارة الحضور المستقلة (Attendance Admin)
   Firebase Auth · Firestore Sync · CSV Export Engine
   ═══════════════════════════════════════════════════════════════════ */

import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import {
  getAuth,
  signInWithEmailAndPassword,
  signOut,
  onAuthStateChanged
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";
import {
  getFirestore,
  collection,
  query,
  where,
  orderBy,
  getDocs,
  doc,
  getDoc,
  updateDoc
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";

/* ── إعدادات Firebase العامة للمشروع ── */
export const FIREBASE_CONFIG = {
  apiKey: "AIzaSyDnWUCjJKkMqORT8SeLYHszAIP0bv8PCSg",
  authDomain: "arthwhdarh-782ec.firebaseapp.com",
  projectId: "arthwhdarh-782ec",
  storageBucket: "arthwhdarh-782ec.firebasestorage.app",
  messagingSenderId: "597405952213",
  appId: "1:597405952213:web:5ffeab6adc7e451a265a5c"
};

const app = initializeApp(FIREBASE_CONFIG);
export const auth = getAuth(app);
export const db = getFirestore(app);

/* ── حالة التطبيق العامة (Application State) ── */
const state = {
  currentAdmin: null,
  activeTab: "records", // 'records' | 'employees'
  allRecords: [],
  allEmployees: [],
  filteredRecords: [],
  activeDateFilter: "today", // 'today' | 'yesterday' | 'week' | 'month' | 'all' | 'custom_single' | 'custom_range'
  selectedDate: null,
  selectedRangeStart: null,
  selectedRangeEnd: null,
  searchTerm: "",
  statusFilter: "all", // 'all' | 'on-time' | 'late'
  sortOrder: "time-desc", // 'time-desc' | 'time-asc' | 'name-asc' | 'name-desc'
  selectedEmployee: null
};

/* ── عناصر واجهة المستخدم ── */
const els = {
  loginScreen: document.getElementById("loginScreen"),
  adminApp: document.getElementById("adminApp"),
  loginForm: document.getElementById("loginForm"),
  loginEmail: document.getElementById("loginEmail"),
  loginPassword: document.getElementById("loginPassword"),
  loginBtn: document.getElementById("loginBtn"),
  loginError: document.getElementById("loginError"),

  adminEmailBadge: document.getElementById("adminEmailBadge"),
  btnLogout: document.getElementById("btnLogout"),

  // المؤشرات الرقمية
  statTotalEmployees: document.getElementById("statTotalEmployees"),
  statPresentToday: document.getElementById("statPresentToday"),
  statLateToday: document.getElementById("statLateToday"),
  statAbsentToday: document.getElementById("statAbsentToday"),

  // التبويبات
  tabBtnRecords: document.getElementById("tabBtnRecords"),
  tabBtnEmployees: document.getElementById("tabBtnEmployees"),
  tabBadgeRecords: document.getElementById("tabBadgeRecords"),
  tabBadgeEmployees: document.getElementById("tabBadgeEmployees"),
  viewTabRecords: document.getElementById("viewTabRecords"),
  viewTabEmployees: document.getElementById("viewTabEmployees"),

  // أدوات التصفية
  datePills: document.querySelectorAll(".date-pill"),
  filterSingleDate: document.getElementById("filterSingleDate"),
  filterRangeStart: document.getElementById("filterRangeStart"),
  filterRangeEnd: document.getElementById("filterRangeEnd"),
  inputSearch: document.getElementById("inputSearch"),
  selectStatus: document.getElementById("selectStatus"),
  selectSort: document.getElementById("selectSort"),
  btnExportCSV: document.getElementById("btnExportCSV"),

  // الجداول والعرض
  recordsTableBody: document.getElementById("recordsTableBody"),
  recordsCountText: document.getElementById("recordsCountText"),
  employeesGrid: document.getElementById("employeesGrid"),

  // نافذة تفاصيل الموظف (Modal)
  modalEmpDetails: document.getElementById("modalEmpDetails"),
  modalCloseBtn: document.getElementById("modalCloseBtn"),
  modalEmpName: document.getElementById("modalEmpName"),
  modalEmpUid: document.getElementById("modalEmpUid"),
  modalEmpCreated: document.getElementById("modalEmpCreated"),
  modalEmpLastAtt: document.getElementById("modalEmpLastAtt"),
  modalEmpTotalAtt: document.getElementById("modalEmpTotalAtt"),
  modalEmpRecordsBody: document.getElementById("modalEmpRecordsBody"),
  btnExportSingleEmp: document.getElementById("btnExportSingleEmp")
};

/* ═══════════════ صيغ التواريخ المساعدة ═══════════════ */

function getTodayISO() {
  const d = new Date();
  return d.toISOString().split("T")[0];
}

function getYesterdayISO() {
  const d = new Date();
  d.setDate(d.getDate() - 1);
  return d.toISOString().split("T")[0];
}

function getStartOfWeekISO() {
  const d = new Date();
  const day = d.getDay(); // 0 = Sunday, 6 = Saturday
  const diff = (day + 1) % 7; // الأسبوع يبدأ السبت
  d.setDate(d.getDate() - diff);
  return d.toISOString().split("T")[0];
}

function getCurrentMonthPrefix() {
  const d = new Date();
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, "0");
  return `${year}-${month}`;
}

function formatFirebaseTimestamp(ts) {
  if (!ts) return "—";
  try {
    let dateObj;
    if (ts.toDate) {
      dateObj = ts.toDate();
    } else if (ts.seconds) {
      dateObj = new Date(ts.seconds * 1000);
    } else {
      dateObj = new Date(ts);
    }
    return new Intl.DateTimeFormat("ar-SA", {
      year: "numeric",
      month: "numeric",
      day: "numeric",
      hour: "2-digit",
      minute: "2-digit",
      hour12: true
    }).format(dateObj);
  } catch (e) {
    return "—";
  }
}

/* ═══════════════ المصادقة وتسجيل الدخول ═══════════════ */

async function handleLogin(e) {
  e.preventDefault();
  els.loginError.style.display = "none";

  const email = els.loginEmail.value.trim();
  const password = els.loginPassword.value;

  if (!email || !password) {
    showLoginError("يرجى إدخال البريد الإلكتروني وكلمة المرور.");
    return;
  }

  els.loginBtn.disabled = true;
  els.loginBtn.innerHTML = `<span class="spinner"></span> جاري التحقق...`;

  try {
    await signInWithEmailAndPassword(auth, email, password);
    // سيتم استقبال المستخدم عبر onAuthStateChanged
  } catch (err) {
    console.error("فشل تسجيل الدخول:", err);
    let msg = "البريد الإلكتروني أو كلمة المرور غير صحيحة.";
    if (err.code === "auth/user-not-found" || err.code === "auth/invalid-credential") {
      msg = "بيانات الدخول غير صحيحة. يرجى التأكد من البريد وكلمة المرور.";
    } else if (err.code === "auth/too-many-requests") {
      msg = "تم حظر المحاولات مؤقتاً بسبب تكرار المحاولة. يرجى الانتظار قليلاً.";
    } else if (err.code === "auth/network-request-failed") {
      msg = "تعذر الاتصال بالشبكة. يرجى التحقق من اتصالك بالإنترنت.";
    }
    showLoginError(msg);
  } finally {
    els.loginBtn.disabled = false;
    els.loginBtn.innerHTML = `<span>تسجيل الدخول إلى لوحة الحضور</span>`;
  }
}

function showLoginError(msg) {
  els.loginError.textContent = msg;
  els.loginError.style.display = "block";
}

async function handleLogout() {
  try {
    await signOut(auth);
    window.location.reload();
  } catch (err) {
    console.error("خطأ أثناء تسجيل الخروج:", err);
  }
}

/* ═══════════════ جلب البيانات من Cloud Firestore ═══════════════ */

async function fetchAllData() {
  setTableLoading(true);
  try {
    // 1. جلب جميع الموظفين
    const empSnap = await getDocs(collection(db, "employees"));
    state.allEmployees = empSnap.docs.map(doc => ({ id: doc.id, ...doc.data() }));

    // 2. جلب جميع سجلات الحضور
    let recSnap;
    try {
      const q = query(collection(db, "attendance_records"), orderBy("timestamp", "desc"));
      recSnap = await getDocs(q);
    } catch (e) {
      // احتياطاً في حال عدم وجود فهرس زمني
      recSnap = await getDocs(collection(db, "attendance_records"));
    }
    state.allRecords = recSnap.docs.map(doc => ({ id: doc.id, ...doc.data() }));

    // 3. تحديث بطاقات المؤشرات الرقمية
    updateKPIs();

    // 4. تطبيق الفلاتر وعرض النتائج
    applyFiltersAndRender();
    renderEmployeesGrid();

  } catch (err) {
    console.error("خطأ أثناء استرجاع البيانات من Firestore:", err);
    showTableError("تعذر تحميل البيانات من السحابة. تأكد من امتلاك صلاحية الإدارة للوصول.");
  } finally {
    setTableLoading(false);
  }
}

/* ═══════════════ حساب المؤشرات الرقمية (KPIs) ═══════════════ */

function updateKPIs() {
  const todayISO = getTodayISO();

  // إجمالي الموظفين المسجلين
  const totalEmployees = state.allEmployees.length;
  els.statTotalEmployees.textContent = totalEmployees;

  // الحاضرون اليوم (سجلات تاريخ اليوم)
  const todayRecords = state.allRecords.filter(r => r.date === todayISO);
  const presentUids = new Set(todayRecords.map(r => r.uid));
  const presentCount = presentUids.size;
  els.statPresentToday.textContent = presentCount;

  // المتأخرون اليوم
  const lateRecords = todayRecords.filter(r => r.isLate === true || r.status === "متأخر");
  const lateUids = new Set(lateRecords.map(r => r.uid));
  els.statLateToday.textContent = lateUids.size;

  // الغائبون / لم يحضروا اليوم
  const absentCount = Math.max(0, totalEmployees - presentCount);
  els.statAbsentToday.textContent = absentCount;

  // تحديث شارات التبويبات
  els.tabBadgeRecords.textContent = state.allRecords.length;
  els.tabBadgeEmployees.textContent = totalEmployees;
}

/* ═══════════════ محرك الفلترة والبحث (Filter & Search Engine) ═══════════════ */

function applyFiltersAndRender() {
  let records = [...state.allRecords];
  const todayISO = getTodayISO();
  const yesterdayISO = getYesterdayISO();
  const startOfWeekISO = getStartOfWeekISO();
  const monthPrefix = getCurrentMonthPrefix();

  // 1. تصفية التاريخ
  if (state.activeDateFilter === "today") {
    records = records.filter(r => r.date === todayISO);
  } else if (state.activeDateFilter === "yesterday") {
    records = records.filter(r => r.date === yesterdayISO);
  } else if (state.activeDateFilter === "week") {
    records = records.filter(r => r.date >= startOfWeekISO && r.date <= todayISO);
  } else if (state.activeDateFilter === "month") {
    records = records.filter(r => r.date && r.date.startsWith(monthPrefix));
  } else if (state.activeDateFilter === "custom_single" && state.selectedDate) {
    records = records.filter(r => r.date === state.selectedDate);
  } else if (state.activeDateFilter === "custom_range") {
    if (state.selectedRangeStart && state.selectedRangeEnd) {
      records = records.filter(r => r.date >= state.selectedRangeStart && r.date <= state.selectedRangeEnd);
    } else if (state.selectedRangeStart) {
      records = records.filter(r => r.date >= state.selectedRangeStart);
    } else if (state.selectedRangeEnd) {
      records = records.filter(r => r.date <= state.selectedRangeEnd);
    }
  }

  // 2. البحث بالاسم أو الـ UID
  if (state.searchTerm) {
    const term = state.searchTerm.toLowerCase();
    records = records.filter(r => {
      const nameMatch = r.employeeName && r.employeeName.toLowerCase().includes(term);
      const uidMatch = r.uid && r.uid.toLowerCase().includes(term);
      return nameMatch || uidMatch;
    });
  }

  // 3. التصفية حسب الحالة
  if (state.statusFilter === "on-time") {
    records = records.filter(r => r.status !== "متأخر" && !r.isLate);
  } else if (state.statusFilter === "late") {
    records = records.filter(r => r.status === "متأخر" || r.isLate);
  }

  // 4. الترتيب والفرز
  if (state.sortOrder === "time-desc") {
    records.sort((a, b) => (b.timestamp?.seconds || 0) - (a.timestamp?.seconds || 0));
  } else if (state.sortOrder === "time-asc") {
    records.sort((a, b) => (a.timestamp?.seconds || 0) - (b.timestamp?.seconds || 0));
  } else if (state.sortOrder === "name-asc") {
    records.sort((a, b) => (a.employeeName || "").localeCompare(b.employeeName || "", "ar"));
  } else if (state.sortOrder === "name-desc") {
    records.sort((a, b) => (b.employeeName || "").localeCompare(a.employeeName || "", "ar"));
  }

  state.filteredRecords = records;
  renderRecordsTable(records);
}

/* ═══════════════ رسم جدول السجلات (Render Table) ═══════════════ */

function renderRecordsTable(records) {
  els.recordsCountText.textContent = `(عرض ${records.length} سجل)`;

  if (!records.length) {
    els.recordsTableBody.innerHTML = `
      <tr>
        <td colspan="7">
          <div class="empty-state">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
              <rect x="3" y="4" width="18" height="18" rx="2" ry="2"></rect>
              <line x1="16" y1="2" x2="16" y2="6"></line>
              <line x1="8" y1="2" x2="8" y2="6"></line>
              <line x1="3" y1="10" x2="21" y2="10"></line>
            </svg>
            <h3 style="font-size: 1.1rem; font-weight: 700;">لا توجد سجلات حضور مطابقة</h3>
            <p style="font-size: 0.85rem;">جرّب تعديل خيارات التصفية أو اختيار نطاق تاريخ آخر.</p>
          </div>
        </td>
      </tr>
    `;
    return;
  }

  const rowsHtml = records.map(r => {
    const isLate = r.isLate || r.status === "متأخر";
    const statusClass = isLate ? "badge-status late" : "badge-status on-time";
    const statusText = isLate ? "متأخر" : "حاضر في الموعد";
    const credShort = r.deviceCredentialId ? `${r.deviceCredentialId.slice(0, 10)}...` : "Face ID";
    const uidShort = r.uid ? `${r.uid.slice(0, 8)}...` : "—";
    const formattedCreated = formatFirebaseTimestamp(r.createdAt || r.timestamp);

    return `
      <tr>
        <td>
          <div style="font-weight: 700; color: var(--c-text-primary);">${escapeHtml(r.employeeName || "—")}</div>
        </td>
        <td>
          <span class="mono-code" title="${escapeHtml(r.uid || "")}">
            ${escapeHtml(uidShort)}
            <button class="btn-copy" onclick="copyText('${escapeHtml(r.uid || "")}')" title="نسخ المعرف">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" width="12" height="12" stroke-width="2">
                <rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect>
                <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path>
              </svg>
            </button>
          </span>
        </td>
        <td>
          <div style="font-weight: 600;">${escapeHtml(r.dateFormatted || r.date || "—")}</div>
        </td>
        <td>
          <div style="font-family: var(--font-en); font-weight: 700; color: var(--c-gold-deep);">${escapeHtml(r.checkInTime || "—")}</div>
        </td>
        <td>
          <span class="${statusClass}">${statusText}</span>
        </td>
        <td>
          <span class="mono-code" title="${escapeHtml(r.deviceCredentialId || "")}">
            ${escapeHtml(credShort)}
          </span>
        </td>
        <td style="font-size: 0.78rem; color: var(--c-text-muted);">
          ${formattedCreated}
        </td>
      </tr>
    `;
  }).join("");

  els.recordsTableBody.innerHTML = rowsHtml;
}

function setTableLoading(isLoading) {
  if (isLoading) {
    els.recordsTableBody.innerHTML = `
      <tr>
        <td colspan="7" class="loading-row">
          <span class="spinner"></span>
          <div style="margin-top: 0.5rem; font-weight: 600; color: var(--c-text-muted);">جاري تحميل السجلات من Firebase...</div>
        </td>
      </tr>
    `;
  }
}

function showTableError(msg) {
  els.recordsTableBody.innerHTML = `
    <tr>
      <td colspan="7">
        <div class="empty-state" style="color: var(--c-error);">
          <h3 style="font-size: 1.1rem; font-weight: 700;">حدث خطأ في تحميل البيانات</h3>
          <p style="font-size: 0.85rem;">${escapeHtml(msg)}</p>
        </div>
      </td>
    </tr>
  `;
}

/* ═══════════════ رسم دليل الموظفين (Render Employees Grid) ═══════════════ */

function renderEmployeesGrid() {
  if (!state.allEmployees.length) {
    els.employeesGrid.innerHTML = `
      <div class="empty-state" style="grid-column: 1/-1;">
        <h3 style="font-size: 1.1rem; font-weight: 700;">لا يوجد موظفون مسجلون بعد</h3>
        <p style="font-size: 0.85rem;">عندما يسجل موظف Face ID لأول مرة، سيظهر حسابه وسجله هنا تلقائياً.</p>
      </div>
    `;
    return;
  }

  const cardsHtml = state.allEmployees.map(emp => {
    const empRecords = state.allRecords.filter(r => r.uid === emp.uid);
    const totalDays = empRecords.length;
    const initial = (emp.name || "م")[0];
    const createdDate = formatFirebaseTimestamp(emp.createdAt || emp.credentialCreatedAt);
    const lastAttDate = formatFirebaseTimestamp(emp.lastAttendanceAt);

    return `
      <div class="employee-card">
        <div class="emp-card-header">
          <div style="display: flex; align-items: center; gap: 0.85rem;">
            <div class="emp-avatar">${escapeHtml(initial)}</div>
            <div class="emp-main-info">
              <div class="emp-name">${escapeHtml(emp.name || "موظف")}</div>
              <div class="emp-uid-pill">UID: ${escapeHtml((emp.uid || "").slice(0, 10))}...</div>
            </div>
          </div>
          <span class="badge-status active">نشط</span>
        </div>

        <div class="emp-stats-list">
          <div class="emp-stat-item">
            <span class="emp-stat-label">تاريخ أول تسجيل جهاز:</span>
            <span class="emp-stat-val">${createdDate}</span>
          </div>
          <div class="emp-stat-item">
            <span class="emp-stat-label">آخر عملية حضور:</span>
            <span class="emp-stat-val">${lastAttDate}</span>
          </div>
          <div class="emp-stat-item">
            <span class="emp-stat-label">إجمالي مرات الحضور:</span>
            <span class="emp-stat-val" style="font-family: var(--font-en); color: var(--c-gold-deep);">${totalDays}</span>
          </div>
        </div>

        <button class="btn-view-emp" onclick="openEmployeeModal('${escapeHtml(emp.uid)}')">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" width="16" height="16" stroke-width="2">
            <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path>
            <circle cx="12" cy="12" r="3"></circle>
          </svg>
          <span>عرض سجل الحضور والتفاصيل</span>
        </button>
      </div>
    `;
  }).join("");

  els.employeesGrid.innerHTML = cardsHtml;
}

/* ═══════════════ نافذة تفاصيل الموظف المنبثقة (Employee Details Modal) ═══════════════ */

window.openEmployeeModal = function(uid) {
  const emp = state.allEmployees.find(e => e.uid === uid);
  if (!emp) return;

  state.selectedEmployee = emp;
  const empRecords = state.allRecords.filter(r => r.uid === uid);

  els.modalEmpName.textContent = emp.name || "—";
  els.modalEmpUid.textContent = emp.uid || "—";
  els.modalEmpCreated.textContent = formatFirebaseTimestamp(emp.createdAt || emp.credentialCreatedAt);
  els.modalEmpLastAtt.textContent = formatFirebaseTimestamp(emp.lastAttendanceAt);
  els.modalEmpTotalAtt.textContent = empRecords.length;

  if (!empRecords.length) {
    els.modalEmpRecordsBody.innerHTML = `
      <tr>
        <td colspan="4" style="text-align: center; padding: 2rem; color: var(--c-text-muted);">
          لا توجد سجلات حضور مسجلة لهذا الموظف حتى الآن.
        </td>
      </tr>
    `;
  } else {
    els.modalEmpRecordsBody.innerHTML = empRecords.map(r => {
      const isLate = r.isLate || r.status === "متأخر";
      const statusClass = isLate ? "badge-status late" : "badge-status on-time";
      const statusText = isLate ? "متأخر" : "حاضر في الموعد";
      return `
        <tr>
          <td><strong>${escapeHtml(r.dateFormatted || r.date || "—")}</strong></td>
          <td style="font-family: var(--font-en); font-weight: 700; color: var(--c-gold-deep);">${escapeHtml(r.checkInTime || "—")}</td>
          <td><span class="${statusClass}">${statusText}</span></td>
          <td style="font-size: 0.78rem; color: var(--c-text-muted);">${formatFirebaseTimestamp(r.createdAt || r.timestamp)}</td>
        </tr>
      `;
    }).join("");
  }

  els.modalEmpDetails.style.display = "flex";
  document.body.style.overflow = "hidden";
};

function closeEmployeeModal() {
  els.modalEmpDetails.style.display = "none";
  document.body.style.overflow = "";
  state.selectedEmployee = null;
}

/* ═══════════════ محرك تصدير ملفات Excel / CSV المعتمد (UTF-8 BOM) ═══════════════ */

function exportRecordsToCSV(recordsToExport, customFilename = null) {
  if (!recordsToExport || !recordsToExport.length) {
    alert("لا توجد سجلات لتصديرها.");
    return;
  }

  // ترويسات الأعمدة المنظمة
  const headers = [
    "اسم الموظف",
    "المعرف الفريد (UID)",
    "التاريخ",
    "وقت الحضور",
    "الحالة",
    "نوع العملية",
    "معرف مفتاح الجهاز",
    "وقت التوثيق"
  ];

  const rows = recordsToExport.map(r => {
    const isLate = r.isLate || r.status === "متأخر";
    const status = isLate ? "متأخر" : "حاضر في الموعد";
    const formattedTs = formatFirebaseTimestamp(r.createdAt || r.timestamp);

    return [
      escapeCsvField(r.employeeName || ""),
      escapeCsvField(r.uid || ""),
      escapeCsvField(r.date || ""),
      escapeCsvField(r.checkInTime || ""),
      escapeCsvField(status),
      escapeCsvField(r.type || "check-in"),
      escapeCsvField(r.deviceCredentialId || ""),
      escapeCsvField(formattedTs)
    ];
  });

  // إضافة BOM لضمان قراءة اللغة العربية بشكل سليم في Microsoft Excel
  const BOM = "\uFEFF";
  const csvContent = BOM + [headers.join(","), ...rows.map(e => e.join(","))].join("\r\n");

  // تحديد اسم الملف التلقائي
  let filename = customFilename;
  if (!filename) {
    const today = getTodayISO();
    if (state.activeDateFilter === "today") {
      filename = `attendance-${today}.csv`;
    } else if (state.activeDateFilter === "custom_range" && state.selectedRangeStart && state.selectedRangeEnd) {
      filename = `attendance-${state.selectedRangeStart}-to-${state.selectedRangeEnd}.csv`;
    } else {
      filename = `attendance-records-${today}.csv`;
    }
  }

  // إنشاء رابط التحميل وتنزيل الملف
  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.setAttribute("href", url);
  link.setAttribute("download", filename);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function escapeCsvField(field) {
  if (field === null || field === undefined) return '""';
  const str = String(field).replace(/"/g, '""');
  return `"${str}"`;
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

window.copyText = function(text) {
  if (!text) return;
  navigator.clipboard.writeText(text).then(() => {
    alert("تم نسخ المعرف إلى الحافظة: " + text);
  }).catch(() => {});
};

/* ═══════════════ ربط الأحداث وإعداد الواجهة ═══════════════ */

function initAdmin() {
  // مراقبة حالة المصادقة من Firebase
  onAuthStateChanged(auth, user => {
    if (user) {
      state.currentAdmin = user;
      els.loginScreen.style.display = "none";
      els.adminApp.style.display = "flex";
      els.adminEmailBadge.textContent = user.email || user.uid;
      fetchAllData();
    } else {
      state.currentAdmin = null;
      els.loginScreen.style.display = "flex";
      els.adminApp.style.display = "none";
    }
  });

  // تسجيل الدخول
  if (els.loginForm) {
    els.loginForm.addEventListener("submit", handleLogin);
  }

  // تسجيل الخروج
  if (els.btnLogout) {
    els.btnLogout.addEventListener("click", handleLogout);
  }

  // تبديل التبويبات
  if (els.tabBtnRecords) {
    els.tabBtnRecords.addEventListener("click", () => {
      state.activeTab = "records";
      els.tabBtnRecords.classList.add("active");
      els.tabBtnEmployees.classList.remove("active");
      els.viewTabRecords.style.display = "block";
      els.viewTabEmployees.style.display = "none";
    });
  }

  if (els.tabBtnEmployees) {
    els.tabBtnEmployees.addEventListener("click", () => {
      state.activeTab = "employees";
      els.tabBtnEmployees.classList.add("active");
      els.tabBtnRecords.classList.remove("active");
      els.viewTabEmployees.style.display = "block";
      els.viewTabRecords.style.display = "none";
      renderEmployeesGrid();
    });
  }

  // تصفية أزرار التواريخ السريعة
  els.datePills.forEach(pill => {
    pill.addEventListener("click", () => {
      els.datePills.forEach(p => p.classList.remove("active"));
      pill.classList.add("active");
      state.activeDateFilter = pill.getAttribute("data-filter");

      // تفريغ الحقول اليدوية
      if (els.filterSingleDate) els.filterSingleDate.value = "";
      if (els.filterRangeStart) els.filterRangeStart.value = "";
      if (els.filterRangeEnd) els.filterRangeEnd.value = "";
      state.selectedDate = null;
      state.selectedRangeStart = null;
      state.selectedRangeEnd = null;

      applyFiltersAndRender();
    });
  });

  // اختيار تاريخ محدد
  if (els.filterSingleDate) {
    els.filterSingleDate.addEventListener("change", (e) => {
      const val = e.target.value;
      if (val) {
        els.datePills.forEach(p => p.classList.remove("active"));
        state.activeDateFilter = "custom_single";
        state.selectedDate = val;
        applyFiltersAndRender();
      }
    });
  }

  // اختيار فترة زمنية من - إلى
  function handleRangeChange() {
    const start = els.filterRangeStart ? els.filterRangeStart.value : null;
    const end = els.filterRangeEnd ? els.filterRangeEnd.value : null;
    if (start || end) {
      els.datePills.forEach(p => p.classList.remove("active"));
      state.activeDateFilter = "custom_range";
      state.selectedRangeStart = start;
      state.selectedRangeEnd = end;
      applyFiltersAndRender();
    }
  }

  if (els.filterRangeStart) els.filterRangeStart.addEventListener("change", handleRangeChange);
  if (els.filterRangeEnd) els.filterRangeEnd.addEventListener("change", handleRangeChange);

  // البحث الفوري
  if (els.inputSearch) {
    els.inputSearch.addEventListener("input", (e) => {
      state.searchTerm = e.target.value.trim();
      applyFiltersAndRender();
    });
  }

  // تصفية الحالة
  if (els.selectStatus) {
    els.selectStatus.addEventListener("change", (e) => {
      state.statusFilter = e.target.value;
      applyFiltersAndRender();
    });
  }

  // الترتيب
  if (els.selectSort) {
    els.selectSort.addEventListener("change", (e) => {
      state.sortOrder = e.target.value;
      applyFiltersAndRender();
    });
  }

  // تصدير CSV للسجلات المفلترة
  if (els.btnExportCSV) {
    els.btnExportCSV.addEventListener("click", () => {
      exportRecordsToCSV(state.filteredRecords);
    });
  }

  // إغلاق النافذة المنبثقة للموظف
  if (els.modalCloseBtn) {
    els.modalCloseBtn.addEventListener("click", closeEmployeeModal);
  }
  if (els.modalEmpDetails) {
    els.modalEmpDetails.addEventListener("click", (e) => {
      if (e.target === els.modalEmpDetails) closeEmployeeModal();
    });
  }

  // تصدير سجلات موظف محدد فقط
  if (els.btnExportSingleEmp) {
    els.btnExportSingleEmp.addEventListener("click", () => {
      if (!state.selectedEmployee) return;
      const empRecords = state.allRecords.filter(r => r.uid === state.selectedEmployee.uid);
      const safeName = (state.selectedEmployee.name || "employee").replace(/\s+/g, "_");
      const filename = `attendance-${safeName}-${getTodayISO()}.csv`;
      exportRecordsToCSV(empRecords, filename);
    });
  }
}

document.addEventListener("DOMContentLoaded", initAdmin);
