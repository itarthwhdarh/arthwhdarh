/* ══════════════════════════════════════════════════════════
   سحابة إرث وحضارة — وحدة التحكم الرئيسية (cloud.js)
   جمعية إرث وحضارة بالقريات
   ──────────────────────────────────────────────────────────
   منصة تخزين سحابية كاملة بهوية الجمعية المستقلة
   - تحكم بالأقسام: الرئيسية، ملفاتي، المجلدات، الأخيرة، المفضلة، روابط المشاركة، سلة المحذوفات، الإعدادات
   - محرك رفع متعدد في الخلفية بدون تجميد الصفحة
   - نظام مشاركة متقدم محمي بكلمة مرور وصلاحيات حقيقية
   - ذاكرة كاش ذكية للمجلدات (In-Memory Folder Cache)
   - خلو تام وقطعي من الرموز التعبيرية (Emojis)
══════════════════════════════════════════════════════════ */

import {
  loginWithEmail,
  logoutUser,
  onAuthUserChanged,
  verifyAndFetchUserProfile,
  fetchStorageStats,
  listFolderItems,
  createFolderInSharePoint,
  renameSharePointItem,
  deleteSharePointItem,
  getFilePreview,
  searchFilesAndFolders,
  fetchRecycleBin,
  createShareLink,
  buildShareUrl,
  fetchUserShareLinks,
  subscribeUserShareLinks,
  revokeShareLink,
  getFavorites,
  toggleFavorite,
  isItemFavorite,
  folderCache
} from "./services.js";

import { UploadManager } from "./uploadManager.js";

/* ── حالة التطبيق العامة (Application State) ── */
const State = {
  currentUser: null,
  activeView: "home", // home, files, folders, recent, favorites, shares, recycle, settings
  currentFolderId: "root",
  currentFolderName: "سحابة إرث وحضارة",
  folderHistory: [{ id: "root", name: "سحابة إرث وحضارة" }],
  items: [],
  allRecentFiles: [],
  viewMode: localStorage.getItem("cloud_view_mode") || "grid",
  sortBy: "date_desc",
  searchQuery: "",
  searchDebounceTimer: null,
  activeTargetItem: null,
  storageData: null,
  openedFromHome: false,
  selectedShareId: null,
  sharesUnsubscribe: null,
  selectedItemsMap: new Map(),
  lastSelectedItem: null
};

/* ── محرك الرفع الخلفي المستقل ── */
let uploadManager = null;

/* ── عناصر واجهة المستخدم (DOM Elements) ── */
const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => document.querySelectorAll(selector);

const UI = {
  bootScreen: $("#bootScreen"),
  loginScreen: $("#loginScreen"),
  loginForm: $("#loginForm"),
  loginEmail: $("#loginEmail"),
  loginPassword: $("#loginPassword"),
  loginErrBox: $("#loginErrBox"),
  loginErrMsg: $("#loginErrMsg"),
  loginSubmitBtn: $("#loginSubmitBtn"),
  loginBtnText: $("#loginBtnText"),
  togglePassBtn: $("#togglePassBtn"),
  togglePassIcon: $("#togglePassIcon"),
  
  appShell: $("#appShell"),
  sidebar: $("#sidebar"),
  mobileMenuBtn: $("#mobileMenuBtn"),
  userAvatar: $("#userAvatar"),
  userName: $("#userName"),
  userRoleBadge: $("#userRoleBadge"),
  logoutBtn: $("#logoutBtn"),
  
  quotaPercent: $("#quotaPercent"),
  quotaBarFill: $("#quotaBarFill"),
  quotaText: $("#quotaText"),
  
  // التنقل
  navHome: $("#navHome"),
  navAllFiles: $("#navAllFiles"),
  navFolders: $("#navFolders"),
  navRecentFiles: $("#navRecentFiles"),
  navFavorites: $("#navFavorites"),
  navSharedLinks: $("#navSharedLinks"),
  navRecycleBin: $("#navRecycleBin"),
  navSettings: $("#navSettings"),
  
  // الواجهات
  subBar: $("#subBar"),
  viewHome: $("#viewHome"),
  viewFiles: $("#viewFiles"),
  viewShares: $("#viewShares"),
  viewFavorites: $("#viewFavorites"),
  viewRecycle: $("#viewRecycle"),
  viewSettings: $("#viewSettings"),
  
  // عناصر الرئيسية
  homeUploadBtn: $("#homeUploadBtn"),
  homeNewFolderBtn: $("#homeNewFolderBtn"),
  homeStatUsed: $("#homeStatUsed"),
  homeStatFolders: $("#homeStatFolders"),
  homeStatFiles: $("#homeStatFiles"),
  homeStatShares: $("#homeStatShares"),
  homeFoldersGrid: $("#homeFoldersGrid"),
  homeRecentFilesGrid: $("#homeRecentFilesGrid"),
  
  // عناصر ملفاتي والمستكشف
  searchInput: $("#searchInput"),
  searchClearBtn: $("#searchClearBtn"),
  refreshBtn: $("#refreshBtn"),
  newFolderBtn: $("#newFolderBtn"),
  uploadBtn: $("#uploadBtn"),
  emptyUploadBtn: $("#emptyUploadBtn"),
  
  breadcrumbNav: $("#breadcrumbNav"),
  goBackBtn: $("#goBackBtn"),
  folderActionBtns: $("#folderActionBtns"),
  folderUploadBtn: $("#folderUploadBtn"),
  folderNewFolderBtn: $("#folderNewFolderBtn"),
  folderRefreshBtn: $("#folderRefreshBtn"),
  sortSelect: $("#sortSelect"),
  viewGridBtn: $("#viewGridBtn"),
  viewListBtn: $("#viewListBtn"),
  
  explorerCanvas: $("#explorerCanvas"),
  dropZoneOverlay: $("#dropZoneOverlay"),
  loadingSkeleton: $("#loadingSkeleton"),
  
  foldersSection: $("#foldersSection"),
  foldersCount: $("#foldersCount"),
  foldersGrid: $("#foldersGrid"),
  
  filesSection: $("#filesSection"),
  filesCount: $("#filesCount"),
  filesGrid: $("#filesGrid"),
  filesListView: $("#filesListView"),
  filesTableBody: $("#filesTableBody"),
  emptyState: $("#emptyState"),
  errorState: $("#errorState"),
  errorStateDesc: $("#errorStateDesc"),
  retryBtn: $("#retryBtn"),
  
  // عناصر المشاركة ونافذة الإحصائيات المنبثقة
  sharesCount: $("#sharesCount"),
  sharesTableBody: $("#sharesTableBody"),
  emptySharesState: $("#emptySharesState"),
  shareStatsModal: $("#shareStatsModal"),
  modalStatItemName: $("#modalStatItemName"),
  modalStatOpenCount: $("#modalStatOpenCount"),
  modalStatViewCount: $("#modalStatViewCount"),
  modalStatDownloadCount: $("#modalStatDownloadCount"),
  modalStatLastVisited: $("#modalStatLastVisited"),
  modalStatLinkStatus: $("#modalStatLinkStatus"),
  modalStatAccessedItems: $("#modalStatAccessedItems"),
  
  // عناصر المفضلة
  favoritesCount: $("#favoritesCount"),
  favoritesGrid: $("#favoritesGrid"),
  emptyFavoritesState: $("#emptyFavoritesState"),
  
  // عناصر المحذوفات
  recycleCount: $("#recycleCount"),
  recycleGrid: $("#recycleGrid"),
  emptyRecycleState: $("#emptyRecycleState"),
  
  // عناصر الإعدادات
  settUserName: $("#settUserName"),
  settUserEmail: $("#settUserEmail"),
  settUserRole: $("#settUserRole"),
  settCanDelete: $("#settCanDelete"),
  
  // النوافذ
  newFolderModal: $("#newFolderModal"),
  newFolderForm: $("#newFolderForm"),
  folderNameInput: $("#folderNameInput"),
  
  shareModal: $("#shareModal"),
  shareForm: $("#shareForm"),
  shareItemIcon: $("#shareItemIcon"),
  shareItemName: $("#shareItemName"),
  shareItemMeta: $("#shareItemMeta"),
  enablePasswordToggle: $("#enablePasswordToggle"),
  passwordFieldWrap: $("#passwordFieldWrap"),
  sharePasswordInput: $("#sharePasswordInput"),
  toggleSharePassBtn: $("#toggleSharePassBtn"),
  sharePassEye: $("#sharePassEye"),
  enableExpiryToggle: $("#enableExpiryToggle"),
  expiryFieldWrap: $("#expiryFieldWrap"),
  shareExpiryInput: $("#shareExpiryInput"),
  shareResultBox: $("#shareResultBox"),
  shareGeneratedUrl: $("#shareGeneratedUrl"),
  copyShareUrlBtn: $("#copyShareUrlBtn"),
  downloadQrBtn: $("#downloadQrBtn"),
  createShareBtn: $("#createShareBtn"),
  
  previewModal: $("#previewModal"),
  previewFileName: $("#previewFileName"),
  previewViewer: $("#previewViewer"),
  previewDownloadBtn: $("#previewDownloadBtn"),
  
  renameModal: $("#renameModal"),
  renameForm: $("#renameForm"),
  renameInput: $("#renameInput"),
  renameModalTitle: $("#renameModalTitle"),
  
  deleteModal: $("#deleteModal"),
  deleteItemName: $("#deleteItemName"),
  confirmDeleteBtn: $("#confirmDeleteBtn"),
  
  revokeShareModal: $("#revokeShareModal"),
  revokeShareItemName: $("#revokeShareItemName"),
  confirmRevokeShareBtn: $("#confirmRevokeShareBtn"),

  selectionMarquee: $("#selectionMarquee"),
  customContextMenu: $("#customContextMenu"),

  toastContainer: $("#toastContainer")
};

/* ═══════════════ نظام الإشعارات العائمة (Toast Notifications) ═══════════════ */

export function showToast(message, type = "info", duration = 4000) {
  const toast = document.createElement("div");
  toast.className = `toast-item ${type}`;

  let icon = "fa-circle-info";
  if (type === "success") icon = "fa-circle-check";
  if (type === "error") icon = "fa-circle-exclamation";
  if (type === "warning") icon = "fa-triangle-exclamation";

  toast.innerHTML = `
    <i class="fa-solid ${icon} toast-icon"></i>
    <div class="toast-content">${escapeHtml(message)}</div>
    <i class="fa-solid fa-xmark toast-close"></i>
  `;

  toast.querySelector(".toast-close").addEventListener("click", () => {
    toast.style.opacity = "0";
    toast.style.transform = "translateY(10px)";
    setTimeout(() => toast.remove(), 200);
  });

  UI.toastContainer.appendChild(toast);

  setTimeout(() => {
    if (toast.parentElement) {
      toast.style.opacity = "0";
      toast.style.transform = "translateY(10px)";
      setTimeout(() => toast.remove(), 200);
    }
  }, duration);
}

/* ═══════════════ أدوات مساعدة ═══════════════ */

function escapeHtml(str) {
  if (!str) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function formatBytes(bytes) {
  if (!bytes || bytes === 0) return "0 بايت";
  const k = 1024;
  const sizes = ["بايت", "كيلوبايت", "ميجابايت", "جيجابايت", "تيرابايت"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + " " + sizes[i];
}

function formatDate(isoStr) {
  if (!isoStr) return "غير متوفر";
  try {
    const d = new Date(isoStr);
    return new Intl.DateTimeFormat("ar-SA", {
      year: "numeric",
      month: "short",
      day: "numeric",
      hour: "2-digit",
      minute: "2-digit"
    }).format(d);
  } catch (e) {
    return isoStr;
  }
}

function getFileTypeDetails(fileName = "", mimeType = "") {
  const ext = fileName.split(".").pop().toLowerCase();
  
  if (ext === "pdf" || mimeType.includes("pdf")) {
    return { class: "file-type-pdf", icon: "fa-solid fa-file-pdf", label: "PDF" };
  }
  if (["doc", "docx", "word"].includes(ext) || mimeType.includes("word") || mimeType.includes("officedocument.wordprocessingml")) {
    return { class: "file-type-word", icon: "fa-solid fa-file-word", label: "Word" };
  }
  if (["xls", "xlsx", "sheet", "csv"].includes(ext) || mimeType.includes("sheet") || mimeType.includes("excel")) {
    return { class: "file-type-excel", icon: "fa-solid fa-file-excel", label: "Excel" };
  }
  if (["ppt", "pptx"].includes(ext) || mimeType.includes("presentation") || mimeType.includes("powerpoint")) {
    return { class: "file-type-ppt", icon: "fa-solid fa-file-powerpoint", label: "PowerPoint" };
  }
  if (["jpg", "jpeg", "png", "gif", "webp", "svg", "bmp"].includes(ext) || mimeType.startsWith("image/")) {
    return { class: "file-type-image", icon: "fa-solid fa-file-image", label: "صورة" };
  }
  if (["mp4", "mov", "avi", "mkv", "webm"].includes(ext) || mimeType.startsWith("video/")) {
    return { class: "file-type-video", icon: "fa-solid fa-file-video", label: "فيديو" };
  }
  if (["mp3", "wav", "ogg", "m4a"].includes(ext) || mimeType.startsWith("audio/")) {
    return { class: "file-type-audio", icon: "fa-solid fa-file-audio", label: "صوت" };
  }
  if (["zip", "rar", "7z", "tar", "gz"].includes(ext) || mimeType.includes("zip") || mimeType.includes("compressed")) {
    return { class: "file-type-zip", icon: "fa-solid fa-file-zipper", label: "أرشيف" };
  }
  if (["txt", "md", "json", "html", "css", "js"].includes(ext)) {
    return { class: "file-type-code", icon: "fa-solid fa-file-lines", label: "نص" };
  }
  
  return { class: "file-type-other", icon: "fa-solid fa-file", label: ext.toUpperCase() || "ملف" };
}

/* ═══════════════ إدارة النوافذ المنبثقة (Modals) ═══════════════ */

function openModal(modalEl) {
  if (!modalEl) return;
  modalEl.classList.add("open");
}

function closeModal(modalEl) {
  if (!modalEl) return;
  modalEl.classList.remove("open");
}

document.addEventListener("click", (e) => {
  const closeBtn = e.target.closest("[data-close]");
  if (closeBtn) {
    const modalId = closeBtn.getAttribute("data-close");
    closeModal($(`#${modalId}`));
  } else if (e.target.classList.contains("modal-backdrop")) {
    closeModal(e.target);
  }
});

document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") {
    $$(".modal-backdrop.open").forEach(m => closeModal(m));
  }
});

/* ═══════════════ تدفق المصادقة (Auth Flow) ═══════════════ */

onAuthUserChanged(async (firebaseUser) => {
  if (!firebaseUser) {
    UI.appShell.classList.add("hidden");
    UI.loginScreen.classList.remove("hidden");
    UI.bootScreen.classList.add("hidden");
    State.currentUser = null;
    return;
  }

  try {
    const result = await verifyAndFetchUserProfile(firebaseUser);
    if (!result.authorized) {
      showLoginError(result.reason || "حسابك غير مصرح له بالوصول إلى سحابة الجمعية");
      await logoutUser();
      UI.appShell.classList.add("hidden");
      UI.loginScreen.classList.remove("hidden");
      UI.bootScreen.classList.add("hidden");
      return;
    }

    State.currentUser = result.user;
    updateUserInterface(result.user);

    UI.loginScreen.classList.add("hidden");
    UI.appShell.classList.remove("hidden");
    UI.bootScreen.classList.add("hidden");

    // تهيئة محرك الرفع الخلفي
    initUploadManager();

    // تحميل السعة والواجهة الافتراضية (الرئيسية)
    loadStorageQuota();
    switchView("home");

  } catch (err) {
    console.error("[Auth Verification Error]:", err);
    showLoginError("حدث خطأ أثناء التحقق من الصلاحيات. يرجى المحاولة لاحقاً.");
    UI.loginScreen.classList.remove("hidden");
    UI.bootScreen.classList.add("hidden");
  }
});

function updateUserInterface(user) {
  UI.userName.textContent = user.name;
  UI.userAvatar.textContent = (user.name || "م").charAt(0);
  
  const roleLabels = {
    tech_admin: "مسؤول تقني",
    executive: "مدير تنفيذي",
    hr: "الموارد البشرية",
    employee: "موظف"
  };
  UI.userRoleBadge.textContent = roleLabels[user.role] || user.jobTitle || "عضو الجمعية";

  // تعبئة بيانات الإعدادات
  UI.settUserName.textContent = user.name;
  UI.settUserEmail.textContent = user.email;
  UI.settUserRole.textContent = roleLabels[user.role] || "موظف";
  UI.settCanDelete.textContent = user.canDelete ? "مسموح (صلاحية إدارية)" : "غير مصرح (عرض وتنزيل فقط)";

  updateFolderActionButtonsVisibility();
}

export function isUserAdmin() {
  const u = State.currentUser;
  if (!u) return false;
  return Boolean(
    u.isTechAdmin ||
    u.isExec ||
    u.role === "tech_admin" ||
    u.role === "executive" ||
    u.role === "admin" ||
    u.isAdmin ||
    u.canManage ||
    u.canDelete
  );
}

export function updateFolderActionButtonsVisibility() {
  if (!UI.folderActionBtns) return;
  const isAdmin = isUserAdmin();
  const isInFolderView = State.activeView === "files";

  if (isAdmin && isInFolderView) {
    UI.folderActionBtns.style.display = "inline-flex";
  } else {
    UI.folderActionBtns.style.display = "none";
  }
}

function showLoginError(msg) {
  UI.loginErrMsg.textContent = msg;
  UI.loginErrBox.classList.add("visible");
}

function clearLoginError() {
  UI.loginErrMsg.textContent = "";
  UI.loginErrBox.classList.remove("visible");
}

UI.loginForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  clearLoginError();

  const email = UI.loginEmail.value.trim();
  const password = UI.loginPassword.value;

  if (!email || !password) {
    showLoginError("يرجى إدخال البريد الإلكتروني وكلمة المرور");
    return;
  }

  UI.loginSubmitBtn.disabled = true;
  UI.loginBtnText.textContent = "جاري التحقق والدخول...";

  try {
    await loginWithEmail(email, password);
  } catch (err) {
    console.error("[Login Error]:", err.code, err.message);
    let errorText = "بيانات الدخول غير صحيحة، يرجى التأكد وإعادة المحاولة";
    if (err.code === "auth/too-many-requests") {
      errorText = "تم حظر محاولات الدخول مؤقتاً لكثرة المحاولات الخاطئة. الرجاء الانتظار دقيقة.";
    }
    showLoginError(errorText);
  } finally {
    UI.loginSubmitBtn.disabled = false;
    UI.loginBtnText.textContent = "الدخول إلى سحابة الجمعية";
  }
});

UI.togglePassBtn.addEventListener("click", () => {
  const isPass = UI.loginPassword.type === "password";
  UI.loginPassword.type = isPass ? "text" : "password";
  UI.togglePassIcon.className = isPass ? "fa-regular fa-eye-slash" : "fa-regular fa-eye";
});

UI.logoutBtn.addEventListener("click", async () => {
  if (confirm("هل تريد بالفعل تسجيل الخروج من سحابة إرث وحضارة؟")) {
    await logoutUser();
    showToast("تم تسجيل الخروج بنجاح", "info");
  }
});

/* ═══════════════ تهيئة محرك الرفع في الخلفية ═══════════════ */

function initUploadManager() {
  if (uploadManager) return;
  uploadManager = new UploadManager({
    onItemComplete: async (uploadedItem, targetFolderId) => {
      // تحديث فوري مباشر لواجهة المستخدم دون الحاجة لإعادة تحميل الصفحة
      await refreshActiveView(true, true);
    },
    onQueueComplete: async () => {
      showToast("اكتملت جميع عمليات الرفع بنجاح إلى السحابة", "success");
      await refreshActiveView(true, true);
    }
  });

  // منطقة إفلات الملفات على كامل الشاشة
  const canvas = UI.explorerCanvas;
  
  ["dragenter", "dragover"].forEach(eventName => {
    canvas.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      UI.dropZoneOverlay.classList.add("active");
    });
  });

  ["dragleave", "drop"].forEach(eventName => {
    canvas.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      UI.dropZoneOverlay.classList.remove("active");
    });
  });

  canvas.addEventListener("drop", (e) => {
    const files = Array.from(e.dataTransfer.files || []);
    if (files.length > 0) {
      uploadManager.addFiles(files, State.currentFolderId, State.currentFolderName);
      showToast(`تمت إضافة ${files.length} ملف إلى طابور الرفع بالخلفية`, "info");
    }
  });

  // نافذة اختيار الملفات الخفية لزر الرفع
  const fileInput = document.createElement("input");
  fileInput.type = "file";
  fileInput.multiple = true;
  fileInput.style.display = "none";
  document.body.appendChild(fileInput);

  fileInput.addEventListener("change", (e) => {
    const files = Array.from(e.target.files || []);
    if (files.length > 0) {
      uploadManager.addFiles(files, State.currentFolderId, State.currentFolderName);
      showToast(`تمت إضافة ${files.length} ملف إلى طابور الرفع بالخلفية`, "info");
      fileInput.value = "";
    }
  });

  UI.uploadBtn?.addEventListener("click", () => fileInput.click());
  UI.homeUploadBtn?.addEventListener("click", () => fileInput.click());
  UI.emptyUploadBtn?.addEventListener("click", () => fileInput.click());
  UI.folderUploadBtn?.addEventListener("click", () => fileInput.click());
}

/* ═══════════════ نظام التنقل بين أقسام المنصة (Views Router) ═══════════════ */

export function switchView(viewName, pushHistory = true) {
  State.activeView = viewName;
  clearSelection();

  if (pushHistory && window.history && window.history.pushState) {
    try {
      window.history.pushState({ view: viewName }, "", `#${viewName}`);
    } catch (e) {
      // تجاهل في البيئات المقيدة
    }
  }

  // عند العودة للرئيسية يتم تصفير مؤشر الفتح من الرئيسية
  if (viewName === "home") {
    State.openedFromHome = false;
  }

  // تحديث تمييز عناصر الشريط الجانبي
  $$(".nav-item").forEach(item => {
    if (item.getAttribute("data-view") === viewName) {
      item.classList.add("active");
    } else {
      item.classList.remove("active");
    }
  });

  // إخفاء جميع الواجهات
  $$(".view-panel").forEach(p => p.style.display = "none");
  UI.emptyState.style.display = "none";
  UI.errorState.style.display = "none";

  // شريط المسار والترتيب يظهر في استعراض الملفات، المجلدات، والأخيرة
  if (viewName === "files" || viewName === "folders" || viewName === "recent") {
    UI.subBar.style.display = "flex";
  } else {
    UI.subBar.style.display = "none";
  }

  updateFolderActionButtonsVisibility();

  // تفعيل الواجهة المطلوبة
  switch (viewName) {
    case "home":
      UI.viewHome.style.display = "block";
      loadHomeView();
      break;

    case "files":
      UI.viewFiles.style.display = "block";
      loadFolder(State.currentFolderId);
      break;

    case "folders":
      UI.viewFiles.style.display = "block";
      UI.subBar.style.display = "flex";
      loadFoldersOnlyView();
      break;

    case "recent":
      UI.viewFiles.style.display = "block";
      UI.subBar.style.display = "flex";
      loadRecentFilesView();
      break;

    case "favorites":
      UI.viewFavorites.style.display = "block";
      loadFavoritesView();
      break;

    case "shares":
      UI.viewShares.style.display = "block";
      loadSharesView();
      break;

    case "recycle":
      UI.viewRecycle.style.display = "block";
      loadRecycleView();
      break;

    case "settings":
      UI.viewSettings.style.display = "block";
      break;
  }
}

// دعم زر الرجوع في المتصفح للعودة للرئيسية أو الأقسام السابقة
window.addEventListener("popstate", (e) => {
  const targetView = e.state?.view || "home";
  switchView(targetView, false);
});

// ربط أزرار الشريط الجانبي
[
  UI.navHome,
  UI.navAllFiles,
  UI.navFolders,
  UI.navRecentFiles,
  UI.navFavorites,
  UI.navSharedLinks,
  UI.navRecycleBin,
  UI.navSettings
].forEach(btn => {
  if (!btn) return;
  btn.addEventListener("click", (e) => {
    e.preventDefault();
    const target = btn.getAttribute("data-view");
    switchView(target);
    UI.sidebar.classList.remove("open");
  });
});

// ربط النقر على شعار وعنوان السحابة للرجوع للرئيسية دائماً
$(".sidebar-brand")?.addEventListener("click", () => {
  switchView("home");
  UI.sidebar.classList.remove("open");
});

// ربط أزرار الرجوع للرئيسية في ترويسات الأقسام المختلفة
$$(".back-to-home-btn").forEach(btn => {
  btn.addEventListener("click", (e) => {
    e.preventDefault();
    switchView("home");
  });
});

UI.homeNewFolderBtn?.addEventListener("click", () => {
  UI.folderNameInput.value = "";
  openModal(UI.newFolderModal);
  setTimeout(() => UI.folderNameInput.focus(), 100);
});

/* ═══════════════ تحميل سعة التخزين (Storage Quota) ═══════════════ */

async function loadStorageQuota() {
  try {
    const data = await fetchStorageStats();
    State.storageData = data;
    if (data && data.quota) {
      const used = data.quota.used || 0;
      const total = data.quota.total || 27487790694400; // 25 TB
      const pct = Math.min(100, Math.max(1, Math.round((used / total) * 100)));

      UI.quotaPercent.textContent = `${pct}%`;
      UI.quotaBarFill.style.width = `${pct}%`;
      UI.quotaText.textContent = `${formatBytes(used)} مستخدمة من ${formatBytes(total)}`;

      if (UI.homeStatUsed) {
        UI.homeStatUsed.textContent = formatBytes(used);
      }
    }
  } catch (err) {
    console.warn("[Storage Quota Fetch Error]:", err.message);
  }
}

/* ═══════════════ محرك التحديث الفوري الموحد (Unified In-Place Refresh Engine) ═══════════════ */

export async function refreshActiveView(forceRefresh = true, silent = false) {
  if (forceRefresh) {
    folderCache.clear();
  }
  loadStorageQuota();

  switch (State.activeView) {
    case "home":
      return await loadHomeView(forceRefresh);
    case "files":
      return await loadFolder(State.currentFolderId, false, forceRefresh, silent);
    case "folders":
      return await loadFoldersOnlyView(forceRefresh, silent);
    case "recent":
      return await loadRecentFilesView(forceRefresh, silent);
    case "favorites":
      return loadFavoritesView();
    case "shares":
      return await loadSharesView();
    case "recycle":
      return await loadRecycleView();
    default:
      return await loadFolder(State.currentFolderId, false, forceRefresh, silent);
  }
}

/* ═══════════════ 1. واجهة الرئيسية (Home Dashboard) ═══════════════ */

async function loadHomeView(forceRefresh = false) {
  try {
    if (forceRefresh) folderCache.invalidate("root");
    // جلب عناصر الجذر من السحابة أو الكاش
    const rootData = await listFolderItems("root", !forceRefresh);
    const items = rootData.items || [];
    const folders = items.filter(i => i.isFolder);
    const files = items.filter(i => !i.isFolder);

    if (UI.homeStatFolders) UI.homeStatFolders.textContent = folders.length;
    if (UI.homeStatFiles) UI.homeStatFiles.textContent = files.length;

    // استعلام عدد روابط المشاركة
    if (State.currentUser) {
      const shares = await fetchUserShareLinks(State.currentUser.uid);
      const activeShares = shares.filter(s => s.status === "active");
      if (UI.homeStatShares) UI.homeStatShares.textContent = activeShares.length;
    }

    // عرض المجلدات في الرئيسية
    const folderFrag = document.createDocumentFragment();
    folders.slice(0, 6).forEach(folder => {
      folderFrag.appendChild(createFolderCardElement(folder));
    });
    UI.homeFoldersGrid.innerHTML = "";
    UI.homeFoldersGrid.appendChild(folderFrag);

    // عرض أحدث الملفات
    const fileFrag = document.createDocumentFragment();
    const sortedFiles = [...files].sort((a, b) => new Date(b.lastModifiedDateTime || 0) - new Date(a.lastModifiedDateTime || 0));
    sortedFiles.slice(0, 6).forEach(file => {
      fileFrag.appendChild(createFileCardElement(file));
    });
    UI.homeRecentFilesGrid.innerHTML = "";
    UI.homeRecentFilesGrid.appendChild(fileFrag);

  } catch (err) {
    console.warn("[Load Home Error]:", err);
  }
}

/* ═══════════════ 2. واجهة ملفاتي (File Manager Explorer) ═══════════════ */

export async function loadFolder(folderId = "root", isBack = false, forceRefresh = false, silent = false) {
  State.currentFolderId = folderId;
  const cacheKey = folderId === "root" ? "root" : folderId;

  // فحص الكاش الفوري لتوفير استجابة فورية (0ms delay)
  const cachedRes = !forceRefresh ? folderCache.get(cacheKey) : null;

  if (cachedRes) {
    if (UI.errorState) UI.errorState.style.display = "none";
    UI.loadingSkeleton.style.display = "none";

    State.currentFolderName = cachedRes.currentFolder?.name || "سحابة إرث وحضارة";
    updateFolderHistory(folderId, isBack);
    updateBreadcrumbs();
    State.items = cachedRes.items || [];
    applySortingAndRender();
    updateFolderActionButtonsVisibility();

    // تحديث صامت في الخلفية لضمان تطابق البيانات
    listFolderItems(cacheKey, false).then(freshRes => {
      if (freshRes && State.currentFolderId === folderId) {
        State.currentFolderName = freshRes.currentFolder?.name || State.currentFolderName;
        State.items = freshRes.items || [];
        applySortingAndRender();
      }
    }).catch(err => {
      console.warn("[Background Revalidate Warning]:", err);
    });

    return;
  }

  if (!silent) {
    UI.loadingSkeleton.style.display = "block";
    UI.foldersSection.style.display = "none";
    UI.filesSection.style.display = "none";
    UI.emptyState.style.display = "none";
  }
  if (UI.errorState) UI.errorState.style.display = "none";

  try {
    if (forceRefresh) {
      folderCache.invalidate(cacheKey);
    }
    const res = await listFolderItems(cacheKey, !forceRefresh);
    State.currentFolderName = res.currentFolder?.name || "سحابة إرث وحضارة";

    updateFolderHistory(folderId, isBack);
    updateBreadcrumbs();

    State.items = res.items || [];
    applySortingAndRender();

  } catch (err) {
    console.error("[Load Folder Error]:", err);
    if (UI.errorState) {
      UI.foldersSection.style.display = "none";
      UI.filesSection.style.display = "none";
      UI.emptyState.style.display = "none";
      UI.errorState.style.display = "flex";
      if (UI.errorStateDesc) {
        UI.errorStateDesc.textContent = err.message || "تعذر الاتصال بالسحابة. يرجى المحاولة مرة أخرى.";
      }
    }
    showToast(err.message || "تعذر جلب ملفات المجلد", "error");
  } finally {
    if (!silent) {
      UI.loadingSkeleton.style.display = "none";
    }
    updateFolderActionButtonsVisibility();
  }
}

function updateFolderHistory(folderId, isBack) {
  if (folderId === "root") {
    State.folderHistory = [{ id: "root", name: "سحابة إرث وحضارة" }];
  } else if (!isBack) {
    const existsIndex = State.folderHistory.findIndex(h => h.id === folderId);
    if (existsIndex !== -1) {
      State.folderHistory = State.folderHistory.slice(0, existsIndex + 1);
    } else {
      State.folderHistory.push({ id: folderId, name: State.currentFolderName });
    }
  }
}

function updateBreadcrumbs() {
  UI.breadcrumbNav.innerHTML = "";
  
  // 1. عنصر "الرئيسية" يظهر دائماً في مقدمة مسار التنقل
  const homeBc = document.createElement("span");
  homeBc.className = "bc-item";
  homeBc.setAttribute("data-id", "home");
  homeBc.setAttribute("role", "button");
  homeBc.setAttribute("tabindex", "0");
  homeBc.setAttribute("title", "الرجوع إلى الصفحة الرئيسية للسحابة");
  homeBc.innerHTML = '<i class="fa-solid fa-house"></i> الرئيسية';
  homeBc.addEventListener("click", () => switchView("home"));
  homeBc.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      switchView("home");
    }
  });
  UI.breadcrumbNav.appendChild(homeBc);

  const sep1 = document.createElement("i");
  sep1.className = "fa-solid fa-chevron-left bc-sep";
  UI.breadcrumbNav.appendChild(sep1);

  // 2. إذا كان المستخدم في جذر ملفاتي (root)
  if (State.folderHistory.length <= 1) {
    const rootBc = document.createElement("span");
    rootBc.className = "bc-item current";
    rootBc.setAttribute("data-id", "root");
    rootBc.innerHTML = '<i class="fa-solid fa-folder-tree"></i> ملفاتي';
    UI.breadcrumbNav.appendChild(rootBc);

    // زر الرجوع في الجذر يرجع دائماً إلى الصفحة الرئيسية للسحابة
    UI.goBackBtn.style.display = "inline-flex";
    UI.goBackBtn.innerHTML = '<i class="fa-solid fa-arrow-right"></i> الرجوع للرئيسية';
    UI.goBackBtn.title = "الرجوع إلى الصفحة الرئيسية للسحابة";
    updateFolderActionButtonsVisibility();
    return;
  }

  // 3. إذا كان داخل مجلدات فرعية
  // عنصر "ملفاتي" كحلقة وسيطة قابلة للنقر
  const filesBc = document.createElement("span");
  filesBc.className = "bc-item";
  filesBc.setAttribute("data-id", "root");
  filesBc.setAttribute("role", "button");
  filesBc.setAttribute("tabindex", "0");
  filesBc.setAttribute("title", "عرض جميع الملفات (الجذر)");
  filesBc.innerHTML = '<i class="fa-solid fa-folder-tree"></i> ملفاتي';
  filesBc.addEventListener("click", () => {
    State.folderHistory = [{ id: "root", name: "سحابة إرث وحضارة" }];
    loadFolder("root", true);
  });
  filesBc.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      State.folderHistory = [{ id: "root", name: "سحابة إرث وحضارة" }];
      loadFolder("root", true);
    }
  });
  UI.breadcrumbNav.appendChild(filesBc);

  // المجلدات الفرعية في المسار
  for (let i = 1; i < State.folderHistory.length; i++) {
    const item = State.folderHistory[i];
    const isLast = i === State.folderHistory.length - 1;

    const sep = document.createElement("i");
    sep.className = "fa-solid fa-chevron-left bc-sep";
    UI.breadcrumbNav.appendChild(sep);

    const folderBc = document.createElement("span");
    folderBc.className = `bc-item ${isLast ? "current" : ""}`;
    folderBc.setAttribute("data-id", item.id);
    folderBc.innerHTML = `<i class="fa-solid fa-folder"></i> ${escapeHtml(item.name)}`;

    if (!isLast) {
      folderBc.setAttribute("role", "button");
      folderBc.setAttribute("tabindex", "0");
      folderBc.addEventListener("click", () => {
        State.folderHistory = State.folderHistory.slice(0, i + 1);
        loadFolder(item.id, true);
      });
      folderBc.addEventListener("keydown", (e) => {
        if (e.key === "Enter" || e.key === " ") {
          e.preventDefault();
          State.folderHistory = State.folderHistory.slice(0, i + 1);
          loadFolder(item.id, true);
        }
      });
    }

    UI.breadcrumbNav.appendChild(folderBc);
  }

  // زر الرجوع في المجلدات الفرعية
  UI.goBackBtn.style.display = "inline-flex";
  UI.goBackBtn.innerHTML = '<i class="fa-solid fa-arrow-right"></i> رجوع';
  UI.goBackBtn.title = "الرجوع للمجلد السابق";
  updateFolderActionButtonsVisibility();
}

UI.goBackBtn.addEventListener("click", () => {
  if (State.activeView !== "files") {
    switchView("home");
    return;
  }

  if (State.folderHistory.length > 1) {
    State.folderHistory.pop();
    const prev = State.folderHistory[State.folderHistory.length - 1];
    if (prev.id === "root" && State.openedFromHome) {
      State.openedFromHome = false;
      switchView("home");
    } else {
      loadFolder(prev.id, true);
    }
  } else {
    // الرجوع من جذر ملفاتي إلى الصفحة الرئيسية
    switchView("home");
  }
});

/* عرض المجلدات فقط */
async function loadFoldersOnlyView(forceRefresh = false, silent = false) {
  UI.breadcrumbNav.innerHTML = `
    <span class="bc-item" data-id="home" role="button" tabindex="0" title="الرجوع إلى الصفحة الرئيسية للسحابة">
      <i class="fa-solid fa-house"></i> الرئيسية
    </span>
    <i class="fa-solid fa-chevron-left bc-sep"></i>
    <span class="bc-item current">
      <i class="fa-solid fa-folder"></i> جميع مجلدات السحابة
    </span>
  `;
  UI.breadcrumbNav.querySelector('[data-id="home"]')?.addEventListener("click", () => switchView("home"));
  UI.goBackBtn.style.display = "inline-flex";
  UI.goBackBtn.innerHTML = '<i class="fa-solid fa-arrow-right"></i> الرجوع للرئيسية';
  UI.goBackBtn.title = "الرجوع إلى الصفحة الرئيسية للسحابة";
  updateFolderActionButtonsVisibility();
  if (!silent) UI.loadingSkeleton.style.display = "block";

  try {
    if (forceRefresh) folderCache.invalidate("root");
    const rootData = await listFolderItems("root", !forceRefresh);
    const folders = (rootData.items || []).filter(i => i.isFolder);

    UI.foldersSection.style.display = "block";
    UI.filesSection.style.display = "none";
    UI.foldersCount.textContent = folders.length;
    UI.foldersGrid.innerHTML = "";

    folders.forEach(f => UI.foldersGrid.appendChild(createFolderCardElement(f)));
  } catch (e) {
    showToast("تعذر جلب المجلدات", "error");
  } finally {
    if (!silent) UI.loadingSkeleton.style.display = "none";
  }
}

/* عرض العناصر الأخيرة */
async function loadRecentFilesView(forceRefresh = false, silent = false) {
  UI.breadcrumbNav.innerHTML = `
    <span class="bc-item" data-id="home" role="button" tabindex="0" title="الرجوع إلى الصفحة الرئيسية للسحابة">
      <i class="fa-solid fa-house"></i> الرئيسية
    </span>
    <i class="fa-solid fa-chevron-left bc-sep"></i>
    <span class="bc-item current">
      <i class="fa-solid fa-clock-rotate-left"></i> العناصر الأخيرة
    </span>
  `;
  UI.breadcrumbNav.querySelector('[data-id="home"]')?.addEventListener("click", () => switchView("home"));
  UI.goBackBtn.style.display = "inline-flex";
  UI.goBackBtn.innerHTML = '<i class="fa-solid fa-arrow-right"></i> الرجوع للرئيسية';
  UI.goBackBtn.title = "الرجوع إلى الصفحة الرئيسية للسحابة";
  updateFolderActionButtonsVisibility();
  if (!silent) UI.loadingSkeleton.style.display = "block";

  try {
    if (forceRefresh) folderCache.invalidate("root");
    const rootData = await listFolderItems("root", !forceRefresh);
    const files = (rootData.items || []).filter(i => !i.isFolder);
    files.sort((a, b) => new Date(b.lastModifiedDateTime || 0) - new Date(a.lastModifiedDateTime || 0));

    UI.foldersSection.style.display = "none";
    UI.filesSection.style.display = "block";
    UI.filesCount.textContent = files.length;

    renderFilesGrid(files);
  } catch (e) {
    showToast("تعذر جلب الملفات الأخيرة", "error");
  } finally {
    if (!silent) UI.loadingSkeleton.style.display = "none";
  }
}

/* ═══════════════ الترتيب والفرز والعرض ═══════════════ */

function applySortingAndRender() {
  let list = [...State.items];

  if (State.searchQuery) {
    const q = State.searchQuery.toLowerCase();
    list = list.filter(item => item.name.toLowerCase().includes(q));
  }

  list.sort((a, b) => {
    switch (State.sortBy) {
      case "name_asc":
        return a.name.localeCompare(b.name, "ar");
      case "name_desc":
        return b.name.localeCompare(a.name, "ar");
      case "date_desc":
        return new Date(b.lastModifiedDateTime || 0) - new Date(a.lastModifiedDateTime || 0);
      case "date_asc":
        return new Date(a.lastModifiedDateTime || 0) - new Date(b.lastModifiedDateTime || 0);
      case "size_desc":
        return (b.size || 0) - (a.size || 0);
      case "size_asc":
        return (a.size || 0) - (b.size || 0);
      default:
        return 0;
    }
  });

  renderItems(list);
}

function renderItems(items) {
  const folders = items.filter(i => i.isFolder);
  const files = items.filter(i => !i.isFolder);

  UI.foldersCount.textContent = folders.length;
  UI.filesCount.textContent = files.length;

  if (folders.length === 0 && files.length === 0) {
    UI.foldersSection.style.display = "none";
    UI.filesSection.style.display = "none";
    UI.emptyState.style.display = "flex";
    return;
  }

  UI.emptyState.style.display = "none";

  // 1. عرض المجلدات
  if (folders.length > 0) {
    UI.foldersSection.style.display = "block";
    const frag = document.createDocumentFragment();
    folders.forEach(folder => {
      frag.appendChild(createFolderCardElement(folder));
    });
    UI.foldersGrid.innerHTML = "";
    UI.foldersGrid.appendChild(frag);
  } else {
    UI.foldersSection.style.display = "none";
  }

  // 2. عرض الملفات
  if (files.length > 0) {
    UI.filesSection.style.display = "block";
    if (State.viewMode === "grid") {
      UI.filesGrid.style.display = "grid";
      UI.filesListView.style.display = "none";
      renderFilesGrid(files);
    } else {
      UI.filesGrid.style.display = "none";
      UI.filesListView.style.display = "block";
      renderFilesList(files);
    }
  } else {
    UI.filesSection.style.display = "none";
  }
}

/* ═══════════════ محرك القائمة المنسدلة الذكية للإجراءات (⋯) ═══════════════ */

let activeDropdownMenu = null;

function closeItemMenu() {
  if (activeDropdownMenu) {
    activeDropdownMenu.remove();
    activeDropdownMenu = null;
  }
}

document.addEventListener("click", (e) => {
  if (activeDropdownMenu && !activeDropdownMenu.contains(e.target)) {
    closeItemMenu();
  }
});

document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") {
    closeItemMenu();
  }
});

window.addEventListener("resize", closeItemMenu);
window.addEventListener("scroll", closeItemMenu, true);

function openItemMenu({ item, isFolder, triggerElement, isFav }) {
  closeItemMenu();

  const menu = document.createElement("div");
  menu.className = "cloud-dropdown-menu";

  if (isFolder) {
    menu.innerHTML = `
      <button type="button" class="cloud-dropdown-item" data-action="open">
        <i class="fa-solid fa-folder-open"></i>
        <span>فتح المجلد</span>
      </button>
      <button type="button" class="cloud-dropdown-item" data-action="share">
        <i class="fa-solid fa-share-nodes"></i>
        <span>مشاركة المجلد</span>
      </button>
      <button type="button" class="cloud-dropdown-item ${isFav ? 'active-fav' : ''}" data-action="fav">
        <i class="${isFav ? 'fa-solid' : 'fa-regular'} fa-star"></i>
        <span>${isFav ? 'إزالة من المفضلة' : 'إضافة للمفضلة'}</span>
      </button>
      <button type="button" class="cloud-dropdown-item" data-action="rename">
        <i class="fa-solid fa-pen"></i>
        <span>إعادة تسمية</span>
      </button>
      ${State.currentUser?.canDelete ? `
        <div class="cloud-dropdown-divider"></div>
        <button type="button" class="cloud-dropdown-item danger" data-action="delete">
          <i class="fa-solid fa-trash"></i>
          <span>حذف المجلد</span>
        </button>
      ` : ""}
    `;
  } else {
    menu.innerHTML = `
      <button type="button" class="cloud-dropdown-item" data-action="preview">
        <i class="fa-solid fa-eye"></i>
        <span>معاينة وفتح</span>
      </button>
      <a href="${item.downloadUrl || '#'}" download class="cloud-dropdown-item" data-action="download">
        <i class="fa-solid fa-download"></i>
        <span>تنزيل الملف</span>
      </a>
      <button type="button" class="cloud-dropdown-item" data-action="share">
        <i class="fa-solid fa-share-nodes"></i>
        <span>مشاركة الرابط</span>
      </button>
      <button type="button" class="cloud-dropdown-item ${isFav ? 'active-fav' : ''}" data-action="fav">
        <i class="${isFav ? 'fa-solid' : 'fa-regular'} fa-star"></i>
        <span>${isFav ? 'إزالة من المفضلة' : 'إضافة للمفضلة'}</span>
      </button>
      <button type="button" class="cloud-dropdown-item" data-action="rename">
        <i class="fa-solid fa-pen"></i>
        <span>إعادة تسمية</span>
      </button>
      ${State.currentUser?.canDelete ? `
        <div class="cloud-dropdown-divider"></div>
        <button type="button" class="cloud-dropdown-item danger" data-action="delete">
          <i class="fa-solid fa-trash"></i>
          <span>حذف الملف</span>
        </button>
      ` : ""}
    `;
  }

  menu.querySelector('[data-action="open"]')?.addEventListener("click", () => {
    closeItemMenu();
    switchView("files");
    loadFolder(item.id);
  });

  menu.querySelector('[data-action="preview"]')?.addEventListener("click", () => {
    closeItemMenu();
    openPreviewModal(item);
  });

  menu.querySelector('[data-action="download"]')?.addEventListener("click", () => {
    closeItemMenu();
  });

  menu.querySelector('[data-action="share"]')?.addEventListener("click", () => {
    closeItemMenu();
    openShareModal(item);
  });

  menu.querySelector('[data-action="fav"]')?.addEventListener("click", () => {
    closeItemMenu();
    handleToggleFavorite(item);
    if (State.activeView === "home") {
      loadHomeView();
    } else if (State.activeView === "favorites") {
      loadFavoritesView();
    } else {
      applySortingAndRender();
    }
  });

  menu.querySelector('[data-action="rename"]')?.addEventListener("click", () => {
    closeItemMenu();
    openRenameModal(item);
  });

  menu.querySelector('[data-action="delete"]')?.addEventListener("click", () => {
    closeItemMenu();
    openDeleteModal(item);
  });

  document.body.appendChild(menu);
  activeDropdownMenu = menu;

  const rect = triggerElement.getBoundingClientRect();
  const menuRect = menu.getBoundingClientRect();

  let left = rect.right - menuRect.width;
  if (left < 10) left = 10;
  if (left + menuRect.width > window.innerWidth - 10) {
    left = window.innerWidth - menuRect.width - 10;
  }

  let top = rect.bottom + 6;
  if (top + menuRect.height > window.innerHeight - 10) {
    top = Math.max(10, rect.top - menuRect.height - 6);
  }

  menu.style.left = `${left}px`;
  menu.style.top = `${top}px`;
}

/* بطاقة المجلد الموحدة */
function createFolderCardElement(folder) {
  const isFav = State.currentUser ? isItemFavorite(folder.id, State.currentUser.uid) : false;
  const card = document.createElement("div");
  card.className = "folder-card";
  card.setAttribute("role", "button");
  card.setAttribute("tabindex", "0");
  card.dataset.itemId = folder.id;
  card.dataset.itemType = "folder";
  card._itemObject = folder;

  card.innerHTML = `
    <div class="folder-info">
      <div class="folder-icon">
        <i class="fa-solid fa-folder"></i>
      </div>
      <div class="folder-text">
        <span class="folder-name" title="${escapeHtml(folder.name)}">${escapeHtml(folder.name)}</span>
        <span class="folder-meta">
          <span>${folder.childCount || 0} عنصر</span>
          ${isFav ? '<i class="fa-solid fa-star folder-meta-fav" title="في المفضلة"></i>' : ''}
        </span>
      </div>
    </div>
    <button type="button" class="item-menu-btn" data-action="menu" title="المزيد من الخيارات" aria-label="المزيد من الخيارات">
      <i class="fa-solid fa-ellipsis-vertical"></i>
    </button>
  `;

  card.addEventListener("click", (e) => {
    if (e.target.closest('[data-action="menu"]')) return;
    selectSingleItem(folder, true, card, e.ctrlKey || e.metaKey, e.shiftKey);
  });

  card.addEventListener("dblclick", (e) => {
    if (e.target.closest('[data-action="menu"]')) return;
    State.openedFromHome = (State.activeView === "home");
    switchView("files");
    loadFolder(folder.id);
  });

  card.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") {
      if (e.target.closest('[data-action="menu"]')) return;
      e.preventDefault();
      State.openedFromHome = (State.activeView === "home");
      switchView("files");
      loadFolder(folder.id);
    }
  });

  const menuBtn = card.querySelector('[data-action="menu"]');
  menuBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    openItemMenu({ item: folder, isFolder: true, triggerElement: menuBtn, isFav });
  });

  return card;
}

function formatModifiedBy(by) {
  if (!by || String(by).toLowerCase().includes("sharepoint") || String(by).toLowerCase().includes("app")) {
    return "مستخدم الجمعية";
  }
  return String(by);
}

function triggerDirectDownload(url, fileName) {
  if (!url || url === "#") {
    showToast("رابط التنزيل غير متوفر حالياً", "error");
    return;
  }

  let directUrl = url;
  if (directUrl.includes("sharepoint.com") && !directUrl.includes("download.aspx") && !directUrl.includes("download=1")) {
    directUrl += (directUrl.includes("?") ? "&" : "?") + "download=1";
  }

  showToast(`جاري بدء تنزيل "${fileName || 'الملف'}"...`, "info", 3000);

  const a = document.createElement("a");
  a.href = directUrl;
  if (fileName) {
    a.setAttribute("download", fileName);
  }
  a.style.display = "none";
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    try {
      a.remove();
    } catch (e) {}
  }, 1000);
}

/* بطاقة الملف الموحدة (Grid) */
function createFileCardElement(file) {
  const typeInfo = getFileTypeDetails(file.name, file.mimeType);
  const isFav = State.currentUser ? isItemFavorite(file.id, State.currentUser.uid) : false;
  const isImage = typeInfo.class === "file-type-image";

  const card = document.createElement("div");
  card.className = "file-card";
  card.setAttribute("role", "button");
  card.setAttribute("tabindex", "0");
  card.dataset.itemId = file.id;
  card.dataset.itemType = "file";
  card._itemObject = file;

  card.innerHTML = `
    <div class="file-card-header">
      <span class="file-type-badge ${typeInfo.class}">
        <i class="${typeInfo.icon}"></i> ${typeInfo.label}
      </span>
      <div style="display:flex;align-items:center;gap:0.25rem;">
        ${file.downloadUrl ? `
          <button type="button" class="item-menu-btn" data-action="download" title="تنزيل الملف" aria-label="تنزيل الملف">
            <i class="fa-solid fa-download"></i>
          </button>
        ` : ''}
        <button type="button" class="item-menu-btn" data-action="menu" title="المزيد من الخيارات" aria-label="المزيد من الخيارات">
          <i class="fa-solid fa-ellipsis-vertical"></i>
        </button>
      </div>
    </div>

    <div class="file-card-preview">
      ${isImage && file.downloadUrl ? `
        <img src="${file.downloadUrl}" alt="${escapeHtml(file.name)}" loading="lazy" decoding="async" onerror="this.style.display='none';this.nextElementSibling.style.display='flex';"/>
        <div class="file-thumb-icon ${typeInfo.class}" style="display:none;"><i class="${typeInfo.icon}"></i></div>
      ` : `
        <div class="file-thumb-icon ${typeInfo.class}"><i class="${typeInfo.icon}"></i></div>
      `}
    </div>

    <div class="file-card-body">
      <div class="file-name" title="${escapeHtml(file.name)}">${escapeHtml(file.name)}</div>
      <div class="file-sub-meta">
        <span>${formatBytes(file.size)}</span>
        <span class="sep">•</span>
        <span>${formatDate(file.lastModifiedDateTime)}</span>
        ${isFav ? '<span class="sep">•</span><i class="fa-solid fa-star folder-meta-fav" title="في المفضلة"></i>' : ''}
      </div>
    </div>
  `;

  card.addEventListener("click", (e) => {
    if (e.target.closest('[data-action="menu"]') || e.target.closest('[data-action="download"]')) return;
    selectSingleItem(file, false, card, e.ctrlKey || e.metaKey, e.shiftKey);
  });

  card.addEventListener("dblclick", (e) => {
    if (e.target.closest('[data-action="menu"]') || e.target.closest('[data-action="download"]')) return;
    openPreviewModal(file);
  });

  card.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") {
      if (e.target.closest('[data-action="menu"]') || e.target.closest('[data-action="download"]')) return;
      e.preventDefault();
      openPreviewModal(file);
    }
  });

  const cardDlBtn = card.querySelector('[data-action="download"]');
  if (cardDlBtn) {
    cardDlBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      triggerDirectDownload(file.downloadUrl, file.name);
    });
  }

  const menuBtn = card.querySelector('[data-action="menu"]');
  menuBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    openItemMenu({ item: file, isFolder: false, triggerElement: menuBtn, isFav });
  });

  return card;
}

function renderFilesGrid(files) {
  const frag = document.createDocumentFragment();
  files.forEach(file => {
    frag.appendChild(createFileCardElement(file));
  });
  UI.filesGrid.innerHTML = "";
  UI.filesGrid.appendChild(frag);
}

function renderFilesList(files) {
  const frag = document.createDocumentFragment();

  files.forEach(file => {
    const typeInfo = getFileTypeDetails(file.name, file.mimeType);
    const isFav = State.currentUser ? isItemFavorite(file.id, State.currentUser.uid) : false;
    const tr = document.createElement("tr");
    tr.style.cursor = "pointer";
    tr.dataset.itemId = file.id;
    tr.dataset.itemType = "file";
    tr._itemObject = file;

    tr.innerHTML = `
      <td style="width:40px;text-align:center;">
        <button type="button" class="star-btn ${isFav ? 'active' : ''}" data-action="fav" title="${isFav ? 'إزالة من المفضلة' : 'إضافة للمفضلة'}">
          <i class="${isFav ? 'fa-solid' : 'fa-regular'} fa-star"></i>
        </button>
      </td>
      <td>
        <div class="table-file-cell">
          <div class="table-file-icon ${typeInfo.class}">
            <i class="${typeInfo.icon}"></i>
          </div>
          <span class="table-file-name" title="${escapeHtml(file.name)}">${escapeHtml(file.name)}</span>
        </div>
      </td>
      <td><span class="file-type-badge ${typeInfo.class}">${typeInfo.label}</span></td>
      <td style="direction:ltr;text-align:right;">${formatBytes(file.size)}</td>
      <td>${formatDate(file.lastModifiedDateTime)}</td>
      <td>${escapeHtml(formatModifiedBy(file.lastModifiedBy))}</td>
      <td style="width:85px;text-align:center;">
        <div style="display:flex;align-items:center;justify-content:center;gap:0.25rem;">
          ${file.downloadUrl ? `
            <button type="button" class="item-menu-btn" data-action="download" title="تنزيل الملف" aria-label="تنزيل الملف">
              <i class="fa-solid fa-download"></i>
            </button>
          ` : ''}
          <button type="button" class="item-menu-btn" data-action="menu" title="المزيد من الخيارات" aria-label="المزيد من الخيارات">
            <i class="fa-solid fa-ellipsis-vertical"></i>
          </button>
        </div>
      </td>
    `;

    tr.addEventListener("click", (e) => {
      if (e.target.closest('[data-action="fav"]') || e.target.closest('[data-action="menu"]') || e.target.closest('[data-action="download"]')) return;
      selectSingleItem(file, false, tr, e.ctrlKey || e.metaKey, e.shiftKey);
    });

    tr.addEventListener("dblclick", (e) => {
      if (e.target.closest('[data-action="fav"]') || e.target.closest('[data-action="menu"]') || e.target.closest('[data-action="download"]')) return;
      openPreviewModal(file);
    });

    tr.querySelector('[data-action="fav"]')?.addEventListener("click", (e) => {
      e.stopPropagation();
      handleToggleFavorite(file, e.currentTarget);
    });

    const rowDlBtn = tr.querySelector('[data-action="download"]');
    if (rowDlBtn) {
      rowDlBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        triggerDirectDownload(file.downloadUrl, file.name);
      });
    }

    const menuBtn = tr.querySelector('[data-action="menu"]');
    menuBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      openItemMenu({ item: file, isFolder: false, triggerElement: menuBtn, isFav });
    });

    frag.appendChild(tr);
  });

  UI.filesTableBody.innerHTML = "";
  UI.filesTableBody.appendChild(frag);
}

/* ═══════════════ محرك التحديد المتقدم والقائمة السياقية (Windows Explorer Selection & Context Menu) ═══════════════ */

export function clearSelection() {
  State.selectedItemsMap.clear();
  State.lastSelectedItem = null;
  document.querySelectorAll(".folder-card.selected, .file-card.selected, .files-table tr.selected").forEach(el => {
    el.classList.remove("selected");
  });
}

function updateItemSelectionUI() {
  const selectables = document.querySelectorAll(".folder-card[data-item-id], .file-card[data-item-id], .files-table tr[data-item-id]");
  selectables.forEach(el => {
    const id = el.dataset.itemId;
    if (id && State.selectedItemsMap.has(id)) {
      el.classList.add("selected");
    } else {
      el.classList.remove("selected");
    }
  });
}

function getItemById(id) {
  if (State.items && State.items.length) {
    const found = State.items.find(i => i.id === id);
    if (found) return found;
  }
  if (State.allRecentFiles && State.allRecentFiles.length) {
    const found = State.allRecentFiles.find(i => i.id === id);
    if (found) return found;
  }
  return null;
}

function selectSingleItem(item, isFolder, element, isCtrl = false, isShift = false) {
  if (!item || !item.id) return;

  if (isCtrl) {
    if (State.selectedItemsMap.has(item.id)) {
      State.selectedItemsMap.delete(item.id);
    } else {
      State.selectedItemsMap.set(item.id, { item, isFolder, element });
    }
    State.lastSelectedItem = { item, isFolder, element };
  } else if (isShift && State.lastSelectedItem) {
    const allElements = Array.from(document.querySelectorAll(".folder-card[data-item-id], .file-card[data-item-id], .files-table tr[data-item-id]"));
    const lastIdx = allElements.findIndex(el => el.dataset.itemId === State.lastSelectedItem.item.id);
    const currIdx = allElements.findIndex(el => el.dataset.itemId === item.id);
    if (lastIdx !== -1 && currIdx !== -1) {
      const start = Math.min(lastIdx, currIdx);
      const end = Math.max(lastIdx, currIdx);
      State.selectedItemsMap.clear();
      for (let i = start; i <= end; i++) {
        const el = allElements[i];
        const id = el.dataset.itemId;
        const type = el.dataset.itemType;
        const itemObj = el._itemObject || getItemById(id);
        if (id && itemObj) {
          State.selectedItemsMap.set(id, { item: itemObj, isFolder: (type === "folder"), element: el });
        }
      }
    }
  } else {
    State.selectedItemsMap.clear();
    State.selectedItemsMap.set(item.id, { item, isFolder, element });
    State.lastSelectedItem = { item, isFolder, element };
  }

  updateItemSelectionUI();
}

function selectAllItemsInCurrentView() {
  clearSelection();
  const selectables = document.querySelectorAll(".folder-card[data-item-id], .file-card[data-item-id], .files-table tr[data-item-id]");
  selectables.forEach(el => {
    const id = el.dataset.itemId;
    const type = el.dataset.itemType;
    const itemObj = el._itemObject || getItemById(id);
    if (id && itemObj) {
      State.selectedItemsMap.set(id, { item: itemObj, isFolder: (type === "folder"), element: el });
    }
  });
  updateItemSelectionUI();
}

function downloadSelectedFiles(filesList) {
  if (!filesList || filesList.length === 0) return;
  filesList.forEach((file, index) => {
    if (file.downloadUrl) {
      setTimeout(() => {
        const a = document.createElement("a");
        a.href = file.downloadUrl;
        a.download = file.name;
        document.body.appendChild(a);
        a.click();
        a.remove();
      }, index * 350);
    }
  });
  showToast(`جاري بدء تنزيل ${filesList.length} ملف...`, "info");
}

async function addMultipleToFavorites(itemsList) {
  if (!State.currentUser || !itemsList || itemsList.length === 0) return;
  let count = 0;
  for (const itemObj of itemsList) {
    const isFav = isItemFavorite(itemObj.id, State.currentUser.uid);
    if (!isFav) {
      await toggleFavorite(itemObj, State.currentUser.uid);
      count++;
    }
  }
  showToast(`تمت إضافة ${count} عنصر إلى المفضلة`, "success");
  if (State.activeView === "home") loadHomeView();
  else if (State.activeView === "favorites") loadFavoritesView();
  else applySortingAndRender();
}

async function deleteMultipleItems(itemsList) {
  if (!State.currentUser?.canDelete) {
    showToast("غير مصرح لك بحذف العناصر من السحابة (صلاحية إدارية مطلوبة)", "warning");
    return;
  }
  if (!itemsList || itemsList.length === 0) return;

  const count = itemsList.length;
  if (!confirm(`هل أنت تأكد من نقل ${count} عناصر إلى سلة المحذوفات؟`)) return;

  let successCount = 0;
  for (const itemObj of itemsList) {
    try {
      await deleteSharePointItem(itemObj.id, State.currentFolderId);
      State.items = State.items.filter(i => i.id !== itemObj.id);
      successCount++;
    } catch (err) {
      console.error("Error deleting item:", itemObj.name, err);
    }
  }

  clearSelection();
  showToast(`تم حذف ${successCount} عنصر بنجاح`, "success");
  if (State.activeView === "home") loadHomeView();
  else applySortingAndRender();
}

/* ── تفريغ التحديد عند النقر على المساحة الفارغة ── */
document.addEventListener("mousedown", (e) => {
  if (e.button === 2) return;

  const target = e.target;
  const isInsideInteractive = target.closest(".folder-card, .file-card, .files-table tr, .cloud-dropdown-menu, .cloud-context-menu, .modal, .sidebar, .header, input, button, select, textarea, [role='button']");

  if (!isInsideInteractive) {
    if (!e.ctrlKey && !e.shiftKey) {
      clearSelection();
    }
  }
});

/* ── القائمة السياقية لكليك يمين (Custom Context Menu Engine) ── */
function closeCustomContextMenu() {
  const menu = UI.customContextMenu || document.getElementById("customContextMenu");
  if (menu) menu.style.display = "none";
}

function openCustomContextMenu(e, htmlContent) {
  closeCustomContextMenu();
  const menu = UI.customContextMenu || document.getElementById("customContextMenu");
  if (!menu) return;

  menu.innerHTML = htmlContent;
  menu.style.display = "block";

  const menuWidth = menu.offsetWidth || 230;
  const menuHeight = menu.offsetHeight || 260;
  const windowWidth = window.innerWidth;
  const windowHeight = window.innerHeight;

  let left = e.clientX;
  let top = e.clientY;

  if (left + menuWidth > windowWidth - 10) {
    left = Math.max(10, windowWidth - menuWidth - 10);
  }
  if (top + menuHeight > windowHeight - 10) {
    top = Math.max(10, windowHeight - menuHeight - 10);
  }

  menu.style.left = `${left}px`;
  menu.style.top = `${top}px`;
}

document.addEventListener("contextmenu", (e) => {
  if (e.target.closest("input, textarea")) return;
  e.preventDefault();

  closeItemMenu();

  const itemEl = e.target.closest(".folder-card[data-item-id], .file-card[data-item-id], .files-table tr[data-item-id]");

  if (itemEl) {
    const id = itemEl.dataset.itemId;
    const type = itemEl.dataset.itemType;
    const itemObj = itemEl._itemObject || getItemById(id);

    if (!State.selectedItemsMap.has(id)) {
      if (!e.ctrlKey) clearSelection();
      if (id && itemObj) {
        State.selectedItemsMap.set(id, { item: itemObj, isFolder: (type === "folder"), element: itemEl });
        updateItemSelectionUI();
      }
    }

    const selectedCount = State.selectedItemsMap.size;

    if (selectedCount === 1 && itemObj) {
      const isFolder = (type === "folder");
      const isFav = State.currentUser ? isItemFavorite(itemObj.id, State.currentUser.uid) : false;
      renderSingleItemContextMenu(e, itemObj, isFolder, isFav);
    } else if (selectedCount > 1) {
      renderMultiItemContextMenu(e);
    }
  } else {
    renderEmptySpaceContextMenu(e);
  }
});

document.addEventListener("click", (e) => {
  const menu = UI.customContextMenu || document.getElementById("customContextMenu");
  if (menu && !menu.contains(e.target)) {
    closeCustomContextMenu();
  }
});

window.addEventListener("resize", closeCustomContextMenu);
window.addEventListener("scroll", closeCustomContextMenu, true);

function renderSingleItemContextMenu(e, itemObj, isFolder, isFav) {
  let html = `
    <div class="context-menu-header">
      <i class="${isFolder ? 'fa-solid fa-folder' : 'fa-solid fa-file'}"></i>
      <span class="truncate" style="max-width: 200px; display: inline-block; text-overflow: ellipsis; overflow: hidden; white-space: nowrap;">${escapeHtml(itemObj.name)}</span>
    </div>
  `;

  if (isFolder) {
    html += `
      <button type="button" class="context-menu-item" data-action="open">
        <i class="fa-solid fa-folder-open"></i>
        <span>فتح المجلد</span>
      </button>
      <button type="button" class="context-menu-item" data-action="share">
        <i class="fa-solid fa-share-nodes"></i>
        <span>مشاركة المجلد</span>
      </button>
      <button type="button" class="context-menu-item" data-action="fav">
        <i class="${isFav ? 'fa-solid' : 'fa-regular'} fa-star"></i>
        <span>${isFav ? 'إزالة من المفضلة' : 'إضافة للمفضلة'}</span>
      </button>
      <button type="button" class="context-menu-item" data-action="rename">
        <i class="fa-solid fa-pen"></i>
        <span>إعادة تسمية</span>
      </button>
      ${State.currentUser?.canDelete ? `
        <div class="context-menu-divider"></div>
        <button type="button" class="context-menu-item danger" data-action="delete">
          <i class="fa-solid fa-trash"></i>
          <span>نقل إلى السلة</span>
        </button>
      ` : ""}
    `;
  } else {
    html += `
      <button type="button" class="context-menu-item" data-action="preview">
        <i class="fa-solid fa-eye"></i>
        <span>معاينة وفتح</span>
      </button>
      ${itemObj.downloadUrl ? `
        <a href="${itemObj.downloadUrl}" download class="context-menu-item" data-action="download">
          <i class="fa-solid fa-download"></i>
          <span>تنزيل الملف</span>
        </a>
      ` : ""}
      <button type="button" class="context-menu-item" data-action="share">
        <i class="fa-solid fa-share-nodes"></i>
        <span>مشاركة الرابط</span>
      </button>
      <button type="button" class="context-menu-item" data-action="fav">
        <i class="${isFav ? 'fa-solid' : 'fa-regular'} fa-star"></i>
        <span>${isFav ? 'إزالة من المفضلة' : 'إضافة للمفضلة'}</span>
      </button>
      <button type="button" class="context-menu-item" data-action="rename">
        <i class="fa-solid fa-pen"></i>
        <span>إعادة تسمية</span>
      </button>
      ${State.currentUser?.canDelete ? `
        <div class="context-menu-divider"></div>
        <button type="button" class="context-menu-item danger" data-action="delete">
          <i class="fa-solid fa-trash"></i>
          <span>نقل إلى السلة</span>
        </button>
      ` : ""}
    `;
  }

  openCustomContextMenu(e, html);

  const menu = UI.customContextMenu || document.getElementById("customContextMenu");

  menu.querySelector('[data-action="open"]')?.addEventListener("click", () => {
    closeCustomContextMenu();
    switchView("files");
    loadFolder(itemObj.id);
  });

  menu.querySelector('[data-action="preview"]')?.addEventListener("click", () => {
    closeCustomContextMenu();
    openPreviewModal(itemObj);
  });

  menu.querySelector('[data-action="download"]')?.addEventListener("click", () => {
    closeCustomContextMenu();
  });

  menu.querySelector('[data-action="share"]')?.addEventListener("click", () => {
    closeCustomContextMenu();
    openShareModal(itemObj);
  });

  menu.querySelector('[data-action="fav"]')?.addEventListener("click", () => {
    closeCustomContextMenu();
    handleToggleFavorite(itemObj);
    if (State.activeView === "home") loadHomeView();
    else if (State.activeView === "favorites") loadFavoritesView();
    else applySortingAndRender();
  });

  menu.querySelector('[data-action="rename"]')?.addEventListener("click", () => {
    closeCustomContextMenu();
    openRenameModal(itemObj);
  });

  menu.querySelector('[data-action="delete"]')?.addEventListener("click", () => {
    closeCustomContextMenu();
    openDeleteModal(itemObj);
  });
}

function renderMultiItemContextMenu(e) {
  const selectedEntries = Array.from(State.selectedItemsMap.values());
  const count = selectedEntries.length;
  const filesList = selectedEntries.filter(entry => !entry.isFolder).map(entry => entry.item);

  let html = `
    <div class="context-menu-header">
      <i class="fa-solid fa-layer-group"></i>
      <span>تحديد ${count} عناصر</span>
    </div>
    ${filesList.length > 0 ? `
      <button type="button" class="context-menu-item" data-action="download-selected">
        <i class="fa-solid fa-download"></i>
        <span>تنزيل الملفات المحددة (${filesList.length})</span>
      </button>
    ` : ""}
    <button type="button" class="context-menu-item" data-action="fav-selected">
      <i class="fa-solid fa-star"></i>
      <span>إضافة للمفضلة</span>
    </button>
    <button type="button" class="context-menu-item" data-action="clear-selected">
      <i class="fa-solid fa-xmark"></i>
      <span>إلغاء التحديد</span>
    </button>
    ${State.currentUser?.canDelete ? `
      <div class="context-menu-divider"></div>
      <button type="button" class="context-menu-item danger" data-action="delete-selected">
        <i class="fa-solid fa-trash"></i>
        <span>حذف العناصر المحددة</span>
      </button>
    ` : ""}
  `;

  openCustomContextMenu(e, html);

  const menu = UI.customContextMenu || document.getElementById("customContextMenu");

  menu.querySelector('[data-action="download-selected"]')?.addEventListener("click", () => {
    closeCustomContextMenu();
    downloadSelectedFiles(filesList);
  });

  menu.querySelector('[data-action="fav-selected"]')?.addEventListener("click", () => {
    closeCustomContextMenu();
    addMultipleToFavorites(selectedEntries.map(e => e.item));
  });

  menu.querySelector('[data-action="clear-selected"]')?.addEventListener("click", () => {
    closeCustomContextMenu();
    clearSelection();
  });

  menu.querySelector('[data-action="delete-selected"]')?.addEventListener("click", () => {
    closeCustomContextMenu();
    deleteMultipleItems(selectedEntries.map(e => e.item));
  });
}

function renderEmptySpaceContextMenu(e) {
  let html = `
    <button type="button" class="context-menu-item" data-action="refresh">
      <i class="fa-solid fa-rotate"></i>
      <span>تحديث</span>
    </button>
    <button type="button" class="context-menu-item" data-action="new-folder">
      <i class="fa-solid fa-folder-plus"></i>
      <span>مجلد جديد</span>
    </button>
    <div class="context-menu-divider"></div>
    <button type="button" class="context-menu-item" data-action="select-all">
      <i class="fa-solid fa-object-group"></i>
      <span>تحديد الكل</span>
    </button>
  `;

  openCustomContextMenu(e, html);

  const menu = UI.customContextMenu || document.getElementById("customContextMenu");

  menu.querySelector('[data-action="refresh"]')?.addEventListener("click", () => {
    closeCustomContextMenu();
    if (State.activeView === "home") loadHomeView(true);
    else if (State.activeView === "files") loadCurrentFolderItems(true);
    else if (State.activeView === "folders") loadFoldersView(true);
    else if (State.activeView === "recent") loadRecentFilesView(true);
    else if (State.activeView === "favorites") loadFavoritesView();
  });

  menu.querySelector('[data-action="new-folder"]')?.addEventListener("click", () => {
    closeCustomContextMenu();
    openModal(UI.newFolderModal);
    UI.folderNameInput.value = "";
    UI.folderNameInput.focus();
  });

  menu.querySelector('[data-action="select-all"]')?.addEventListener("click", () => {
    closeCustomContextMenu();
    selectAllItemsInCurrentView();
  });
}

/* ═══════════════ التبديل بين Grid و List ═══════════════ */

UI.viewGridBtn.addEventListener("click", () => {
  State.viewMode = "grid";
  localStorage.setItem("cloud_view_mode", "grid");
  UI.viewGridBtn.classList.add("active");
  UI.viewListBtn.classList.remove("active");
  applySortingAndRender();
});

UI.viewListBtn.addEventListener("click", () => {
  State.viewMode = "list";
  localStorage.setItem("cloud_view_mode", "list");
  UI.viewListBtn.classList.add("active");
  UI.viewGridBtn.classList.remove("active");
  applySortingAndRender();
});

UI.sortSelect.addEventListener("change", (e) => {
  State.sortBy = e.target.value;
  applySortingAndRender();
});

UI.refreshBtn?.addEventListener("click", async () => {
  const icon = UI.refreshBtn.querySelector("i");
  if (icon) icon.classList.add("fa-spin");
  showToast("جاري تحديث السحابة...", "info", 1500);
  try {
    await refreshActiveView(true);
  } finally {
    if (icon) icon.classList.remove("fa-spin");
  }
});

if (UI.retryBtn) {
  UI.retryBtn.addEventListener("click", async () => {
    showToast("جاري إعادة محاولة الاتصال...", "info", 1500);
    await refreshActiveView(true);
  });
}

/* ═══════════════ البحث الفوري والعميق (Debounced Search) ═══════════════ */

UI.searchInput.addEventListener("input", (e) => {
  const q = e.target.value.trim();
  State.searchQuery = q;

  if (q.length > 0) {
    UI.searchClearBtn.classList.add("visible");
  } else {
    UI.searchClearBtn.classList.remove("visible");
  }

  clearTimeout(State.searchDebounceTimer);
  State.searchDebounceTimer = setTimeout(() => {
    applySortingAndRender();
  }, 250);
});

UI.searchClearBtn.addEventListener("click", () => {
  UI.searchInput.value = "";
  State.searchQuery = "";
  UI.searchClearBtn.classList.remove("visible");
  applySortingAndRender();
});

UI.searchInput.addEventListener("keydown", async (e) => {
  if (e.key === "Enter") {
    const q = UI.searchInput.value.trim();
    if (!q) return;

    showToast(`جاري البحث العميق عن "${q}" في السحابة...`, "info", 2000);
    try {
      const results = await searchFilesAndFolders(q);
      if (results.length > 0) {
        showToast(`تم العثور على ${results.length} عنصر مطابق`, "success");
        renderItems(results);
      } else {
        showToast("لم يتم العثور على عناصر مطابقة للبحث", "warning");
      }
    } catch (err) {
      console.warn("[Search Error]:", err);
    }
  }
});

/* ═══════════════ إنشاء مجلد جديد ═══════════════ */

UI.newFolderBtn?.addEventListener("click", () => {
  UI.folderNameInput.value = "";
  openModal(UI.newFolderModal);
  setTimeout(() => UI.folderNameInput.focus(), 100);
});

UI.folderNewFolderBtn?.addEventListener("click", () => {
  UI.folderNameInput.value = "";
  openModal(UI.newFolderModal);
  setTimeout(() => UI.folderNameInput.focus(), 100);
});

UI.folderRefreshBtn?.addEventListener("click", async () => {
  const icon = UI.folderRefreshBtn.querySelector("i");
  if (icon) icon.classList.add("fa-spin");
  try {
    folderCache.invalidate(State.currentFolderId === "root" ? null : State.currentFolderId);
    loadStorageQuota();
    await loadFolder(State.currentFolderId, false, true, false);
    showToast("تم تحديث محتويات المجلد بنجاح", "info");
  } catch (e) {
    showToast("تعذر تحديث محتويات المجلد", "error");
  } finally {
    if (icon) icon.classList.remove("fa-spin");
  }
});

UI.newFolderForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  const name = UI.folderNameInput.value.trim();
  if (!name) return;

  const btn = $("#createFolderSubmitBtn");
  btn.disabled = true;
  btn.innerHTML = '<span>جاري إنشاء المجلد...</span>';

  try {
    const parentId = State.currentFolderId === "root" ? null : State.currentFolderId;
    await createFolderInSharePoint(name, parentId);
    showToast(`تم إنشاء المجلد "${name}" بنجاح`, "success");
    closeModal(UI.newFolderModal);
    await refreshActiveView(true);
  } catch (err) {
    console.error("[Create Folder Error]:", err);
    showToast(err.message || "تعذر إنشاء المجلد في السحابة", "error");
  } finally {
    btn.disabled = false;
    btn.innerHTML = '<span>إنشاء المجلد</span>';
  }
});

/* ═══════════════ 3. نظام مشاركة الروابط المشفرة (Share Modal) ═══════════════ */

function openShareModal(item) {
  State.activeTargetItem = item;
  UI.shareItemName.textContent = item.name;
  UI.shareItemMeta.textContent = item.isFolder ? `${item.childCount || 0} عنصر` : formatBytes(item.size);

  const typeInfo = getFileTypeDetails(item.name, item.mimeType);
  UI.shareItemIcon.className = item.isFolder ? "fa-solid fa-folder share-item-icon" : `${typeInfo.icon} share-item-icon`;

  // إعادة ضبط خيارات الوصول دائماً على "عرض وتنزيل" افتراضياً
  const radioDownload = document.querySelector('input[name="accessType"][value="download"]');
  if (radioDownload) {
    radioDownload.checked = true;
  }
  syncAccessRadioCards();

  // إعادة ضبط الحقول
  UI.enablePasswordToggle.checked = false;
  UI.passwordFieldWrap.style.display = "none";
  UI.sharePasswordInput.value = "";
  UI.enableExpiryToggle.checked = false;
  UI.expiryFieldWrap.style.display = "none";
  UI.shareExpiryInput.value = "";
  UI.shareResultBox.style.display = "none";

  // إظهار زر إنشاء رابط المشاركة عند فتح النافذة لأي عنصر
  if (UI.createShareBtn) {
    UI.createShareBtn.style.display = "inline-flex";
  }

  openModal(UI.shareModal);
}

function syncAccessRadioCards() {
  document.querySelectorAll(".access-radio-card").forEach(card => {
    const radio = card.querySelector('input[name="accessType"]');
    if (radio && radio.checked) {
      card.classList.add("active");
    } else {
      card.classList.remove("active");
    }
  });
}

// التحديث الفوري المتبادل بين خيارات نوع الوصول
document.querySelectorAll('.access-radio-card').forEach(card => {
  card.addEventListener("click", () => {
    const radio = card.querySelector('input[name="accessType"]');
    if (radio && !radio.checked) {
      radio.checked = true;
      radio.dispatchEvent(new Event("change", { bubbles: true }));
    }
    syncAccessRadioCards();
  });
});

document.querySelectorAll('input[name="accessType"]').forEach(radio => {
  radio.addEventListener("change", syncAccessRadioCards);
});

UI.enablePasswordToggle.addEventListener("change", (e) => {
  UI.passwordFieldWrap.style.display = e.target.checked ? "block" : "none";
  if (e.target.checked) {
    UI.sharePasswordInput.focus();
    setTimeout(() => {
      UI.passwordFieldWrap.scrollIntoView({ behavior: "smooth", block: "nearest" });
    }, 50);
  }
});

UI.enableExpiryToggle.addEventListener("change", (e) => {
  UI.expiryFieldWrap.style.display = e.target.checked ? "block" : "none";
  if (e.target.checked) {
    setTimeout(() => {
      UI.expiryFieldWrap.scrollIntoView({ behavior: "smooth", block: "nearest" });
    }, 50);
  }
});

UI.toggleSharePassBtn.addEventListener("click", () => {
  const isPass = UI.sharePasswordInput.type === "password";
  UI.sharePasswordInput.type = isPass ? "text" : "password";
  UI.sharePassEye.className = isPass ? "fa-regular fa-eye-slash" : "fa-regular fa-eye";
});

UI.shareForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  if (!State.activeTargetItem) return;

  const btn = $("#createShareBtn");
  btn.disabled = true;
  btn.innerHTML = '<i class="fa-solid fa-circle-notch fa-spin"></i> <span>جاري تشفير وإنشاء الرابط...</span>';

  try {
    const selectedAccess = document.querySelector('input[name="accessType"]:checked')?.value || "download";
    const hasPassword = UI.enablePasswordToggle.checked;
    const password = hasPassword ? UI.sharePasswordInput.value.trim() : "";
    const hasExpiry = UI.enableExpiryToggle.checked;
    const expiry = hasExpiry ? UI.shareExpiryInput.value : null;

    if (hasPassword && (!password || password.length < 4)) {
      showToast("يرجى إدخال كلمة مرور تتكون من 4 أحرف على الأقل", "warning");
      btn.disabled = false;
      btn.innerHTML = '<i class="fa-solid fa-link"></i> <span>إنشاء رابط المشاركة</span>';
      return;
    }

    const res = await createShareLink({
      item: State.activeTargetItem,
      accessType: selectedAccess,
      allowDownload: selectedAccess === "download",
      allowUpload: selectedAccess === "upload",
      password,
      expiresAt: expiry,
      currentUser: State.currentUser
    });

    UI.shareGeneratedUrl.value = res.shareUrl;
    UI.shareResultBox.style.display = "block";
    setTimeout(() => {
      UI.shareResultBox.scrollIntoView({ behavior: "smooth", block: "nearest" });
    }, 50);
    showToast("تم إنشاء رابط المشاركة بنجاح", "success");

    // إخفاء زر إنشاء رابط المشاركة لهذا العنصر لمنع التكرار من نفس النافذة
    if (UI.createShareBtn) {
      UI.createShareBtn.style.display = "none";
    }

    // تحديث فوري مباشر لإحصائيات وروابط المشاركة
    if (State.activeView === "shares") {
      loadSharesView();
    } else if (State.activeView === "home") {
      loadHomeView(false);
    }

  } catch (err) {
    console.error("[Create Share Error]:", err);
    showToast(err.message || "تعذر إنشاء رابط المشاركة", "error");
  } finally {
    btn.disabled = false;
    btn.innerHTML = '<i class="fa-solid fa-link"></i> <span>إنشاء رابط المشاركة</span>';
  }
});

UI.copyShareUrlBtn.addEventListener("click", async () => {
  const url = UI.shareGeneratedUrl.value;
  if (!url) return;
  try {
    await navigator.clipboard.writeText(url);
    showToast("تم نسخ رابط المشاركة إلى الحافظة بنجاح", "success");
  } catch (e) {
    UI.shareGeneratedUrl.select();
    document.execCommand("copy");
    showToast("تم نسخ الرابط", "success");
  }
});

UI.downloadQrBtn?.addEventListener("click", async () => {
  const url = UI.shareGeneratedUrl.value;
  if (!url) return;

  const itemName = State.activeTargetItem?.name || "مشاركة";
  await downloadQrCode(url, itemName);
});

async function downloadQrCode(url, itemName) {
  if (!url) return;

  const safeName = (itemName || "مشاركة")
    .replace(/[\\/:*?"<>|]/g, "_")
    .trim();
  const fileName = `qr-${safeName}.png`;

  try {
    if (typeof window.QRCode === "undefined") {
      await new Promise((resolve, reject) => {
        const s = document.createElement("script");
        s.src = "qrcode.min.js";
        s.onload = resolve;
        s.onerror = () => {
          const s2 = document.createElement("script");
          s2.src = "https://cdn.jsdelivr.net/npm/qrcode@1.5.4/build/qrcode.min.js";
          s2.onload = resolve;
          s2.onerror = reject;
          document.head.appendChild(s2);
        };
        document.head.appendChild(s);
      });
    }

    if (typeof window.QRCode !== "undefined" && window.QRCode.toCanvas) {
      const canvas = document.createElement("canvas");
      await window.QRCode.toCanvas(canvas, url, {
        width: 1000,
        margin: 2,
        color: {
          dark: "#141210ff",
          light: "#00000000"
        }
      });

      const dataUrl = canvas.toDataURL("image/png");
      const a = document.createElement("a");
      a.href = dataUrl;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      showToast("تم تحميل رمز QR بخلفية شفافة بنجاح", "success");
    } else {
      throw new Error("مكتبة QR Code غير متاحة");
    }
  } catch (err) {
    console.error("[QR Code Download Error]:", err);
    showToast("تعذر إنشاء رمز QR، يرجى المحاولة مرة أخرى", "error");
  }
}

/* ═══════════════ إلغاء رابط المشاركة عبر نافذة مخصصة ═══════════════ */
let pendingRevokeShare = null;

function openRevokeShareModal(share) {
  pendingRevokeShare = share;
  if (UI.revokeShareItemName) {
    UI.revokeShareItemName.textContent = `"${share.itemName || "هذا العنصر"}"`;
  }
  openModal(UI.revokeShareModal);
}

UI.confirmRevokeShareBtn?.addEventListener("click", async () => {
  if (!pendingRevokeShare) return;
  const shareId = pendingRevokeShare.shareId;

  UI.confirmRevokeShareBtn.disabled = true;
  UI.confirmRevokeShareBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> <span>جاري الإلغاء...</span>';

  try {
    await revokeShareLink(shareId);
    showToast("تم إلغاء رابط المشاركة بنجاح", "info");
    closeModal(UI.revokeShareModal);
    loadSharesView();
    if (State.activeView === "home") loadHomeView(false);
  } catch (err) {
    console.error("[Revoke Share Error]:", err);
    showToast("تعذر إلغاء رابط المشاركة", "error");
  } finally {
    UI.confirmRevokeShareBtn.disabled = false;
    UI.confirmRevokeShareBtn.innerHTML = '<i class="fa-solid fa-ban"></i> <span>إلغاء الرابط</span>';
    pendingRevokeShare = null;
  }
});

/* ═══════════════ إحصائيات روابط المشاركة والتحديث اللحظي ═══════════════ */

function formatFullDateTime(dateStr) {
  if (!dateStr) return "لم يزار بعد";
  try {
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return "لم يزار بعد";
    const day = d.getDate();
    const months = [
      "يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
      "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"
    ];
    const month = months[d.getMonth()];
    const year = d.getFullYear();
    let hours = d.getHours();
    const minutes = d.getMinutes().toString().padStart(2, "0");
    const ampm = hours >= 12 ? "م" : "ص";
    hours = hours % 12 || 12;
    return `${day} ${month} ${year}، ${hours}:${minutes} ${ampm}`;
  } catch (e) {
    return "لم يزار بعد";
  }
}

let activeStatsShareId = null;

function openShareStatsModal(share) {
  if (!share) return;
  activeStatsShareId = share.shareId;
  updateShareStatsModalContent(share);
  openModal(UI.shareStatsModal);
}

function updateShareStatsModalContent(share) {
  if (!share) return;
  if (UI.modalStatItemName) UI.modalStatItemName.textContent = share.itemName || "رابط مشاركة";
  if (UI.modalStatOpenCount) UI.modalStatOpenCount.textContent = `${share.openCount || 0} مرة`;
  if (UI.modalStatViewCount) UI.modalStatViewCount.textContent = `${share.viewCount || 0} مشاهدة`;
  if (UI.modalStatDownloadCount) UI.modalStatDownloadCount.textContent = `${share.downloadCount || 0} مرة`;
  if (UI.modalStatLastVisited) UI.modalStatLastVisited.textContent = formatFullDateTime(share.lastVisitedAt);

  if (UI.modalStatLinkStatus) {
    let statusBadge = `<span class="badge-tag badge-active"><i class="fa-solid fa-circle-check"></i> نشط</span>`;
    if (share.status === "expired") {
      statusBadge = `<span class="badge-tag badge-expired"><i class="fa-solid fa-calendar-xmark"></i> منتهي</span>`;
    } else if (share.status === "revoked") {
      statusBadge = `<span class="badge-tag badge-revoked"><i class="fa-solid fa-ban"></i> ملغى</span>`;
    }
    UI.modalStatLinkStatus.innerHTML = statusBadge;
  }

  if (UI.modalStatAccessedItems) {
    if (share.isFolder) {
      UI.modalStatAccessedItems.textContent = `${share.childCount || 0} عناصر (مجلد)`;
    } else {
      UI.modalStatAccessedItems.textContent = `ملف واحد (${formatBytes(share.size)})`;
    }
  }
}

function renderSharesTable(shares) {
  if (!UI.sharesTableBody) return;
  UI.sharesCount.textContent = shares.length;

  if (shares.length === 0) {
    UI.sharesTableBody.innerHTML = "";
    UI.emptySharesState.style.display = "flex";
    return;
  }

  UI.emptySharesState.style.display = "none";

  // تحديث النافذة المنبثقة إذا كانت مفتوحة لحفظ التزامن اللحظي
  if (activeStatsShareId) {
    const activeShare = shares.find(s => s.shareId === activeStatsShareId);
    if (activeShare) {
      updateShareStatsModalContent(activeShare);
    }
  }

  UI.sharesTableBody.innerHTML = "";
  shares.forEach(s => {
    const tr = document.createElement("tr");

    let statusBadge = `<span class="badge-tag badge-active"><i class="fa-solid fa-circle-check"></i> نشط</span>`;
    if (s.status === "expired") {
      statusBadge = `<span class="badge-tag badge-expired"><i class="fa-solid fa-calendar-xmark"></i> منتهي</span>`;
    } else if (s.status === "revoked") {
      statusBadge = `<span class="badge-tag badge-revoked"><i class="fa-solid fa-ban"></i> ملغى</span>`;
    }

    const passBadge = s.hasPassword
      ? `<span class="badge-tag badge-password"><i class="fa-solid fa-lock"></i> محمي</span>`
      : `<span style="color:var(--ink-400);font-size:0.8rem;">بدون كلمة مرور</span>`;

    let accessLabel = "عرض وتنزيل";
    if (s.accessType === "view") accessLabel = "عرض فقط";
    if (s.accessType === "upload") accessLabel = "رفع ملفات";

    const fullUrl = buildShareUrl(s.shareId);

    tr.innerHTML = `
      <td><strong>${escapeHtml(s.itemName)}</strong></td>
      <td>${s.isFolder ? '<i class="fa-solid fa-folder text-gold"></i> مجلد' : '<i class="fa-solid fa-file"></i> ملف'}</td>
      <td>${accessLabel}</td>
      <td>${passBadge}</td>
      <td>${s.expiresAt ? formatDate(s.expiresAt) : "دائم"}</td>
      <td>${s.viewCount || 0}</td>
      <td>${statusBadge}</td>
      <td>
        <div style="display:flex;gap:0.35rem;justify-content:center;align-items:center;">
          <button type="button" class="btn-secondary btn-sm-stats" data-action="stats" title="عرض إحصائيات هذا الرابط">
            <i class="fa-solid fa-chart-pie" style="color:var(--gold);"></i>
            <span>الإحصائيات</span>
          </button>
          <button type="button" class="action-icon-btn" data-action="copy" title="نسخ الرابط">
            <i class="fa-solid fa-copy"></i>
          </button>
          ${s.status === "active" ? `
            <button type="button" class="action-icon-btn danger" data-action="revoke" title="إلغاء الرابط">
              <i class="fa-solid fa-ban"></i>
            </button>
          ` : ""}
        </div>
      </td>
    `;

    tr.querySelector('[data-action="stats"]')?.addEventListener("click", (e) => {
      e.stopPropagation();
      openShareStatsModal(s);
    });

    tr.querySelector('[data-action="copy"]')?.addEventListener("click", (e) => {
      e.stopPropagation();
      navigator.clipboard.writeText(fullUrl);
      showToast("تم نسخ رابط المشاركة", "success");
    });

    tr.querySelector('[data-action="revoke"]')?.addEventListener("click", (e) => {
      e.stopPropagation();
      openRevokeShareModal(s);
    });

    UI.sharesTableBody.appendChild(tr);
  });
}

/* عرض روابط المشاركة المنشأة مع التحديث اللحظي الفوري */
async function loadSharesView() {
  if (!State.currentUser) return;
  UI.emptySharesState.style.display = "none";

  if (typeof State.sharesUnsubscribe === "function") {
    State.sharesUnsubscribe();
    State.sharesUnsubscribe = null;
  }

  State.sharesUnsubscribe = subscribeUserShareLinks(State.currentUser.uid, (shares) => {
    renderSharesTable(shares);
  });
}

/* ═══════════════ 4. نظام المفضلة (Favorites) ═══════════════ */

function handleToggleFavorite(item, btnElement) {
  if (!State.currentUser) return;
  const added = toggleFavorite(item, State.currentUser.uid);
  if (btnElement) {
    if (added) {
      btnElement.classList.add("active");
      btnElement.querySelector("i").className = "fa-solid fa-star";
    } else {
      btnElement.classList.remove("active");
      btnElement.querySelector("i").className = "fa-regular fa-star";
    }
  }
  if (added) {
    showToast(`تمت إضافة "${item.name}" إلى المفضلة`, "success", 2000);
  } else {
    showToast(`تمت إزالة "${item.name}" من المفضلة`, "info", 2000);
  }

  // تحديث فوري مباشر لواجهة المفضلة أو الرئيسية إذا كانت معروضة
  if (State.activeView === "favorites") {
    loadFavoritesView();
  } else if (State.activeView === "home") {
    loadHomeView(false);
  }
}

function loadFavoritesView() {
  if (!State.currentUser) return;
  const favs = getFavorites(State.currentUser.uid);
  UI.favoritesCount.textContent = favs.length;

  if (favs.length === 0) {
    UI.favoritesGrid.innerHTML = "";
    UI.emptyFavoritesState.style.display = "flex";
    return;
  }

  UI.emptyFavoritesState.style.display = "none";
  UI.favoritesGrid.innerHTML = "";

  favs.forEach(item => {
    if (item.isFolder) {
      UI.favoritesGrid.appendChild(createFolderCardElement(item));
    } else {
      UI.favoritesGrid.appendChild(createFileCardElement(item));
    }
  });
}

/* ═══════════════ 5. سلة المحذوفات (Recycle Bin) ═══════════════ */

async function loadRecycleView() {
  UI.recycleGrid.innerHTML = '<div style="grid-column:1/-1;text-align:center;padding:2rem;"><i class="fa-solid fa-spinner fa-spin"></i> جاري فحص سلة المحذوفات...</div>';
  UI.emptyRecycleState.style.display = "none";

  try {
    const items = await fetchRecycleBin();
    UI.recycleCount.textContent = items.length;

    if (items.length === 0) {
      UI.recycleGrid.innerHTML = "";
      UI.emptyRecycleState.style.display = "flex";
      return;
    }

    UI.recycleGrid.innerHTML = "";
    items.forEach(item => {
      const card = document.createElement("div");
      card.className = "file-card";
      card.innerHTML = `
        <div class="file-thumb">
          <div class="file-thumb-icon file-type-other"><i class="fa-solid fa-trash"></i></div>
          <span class="file-type-badge">محذوف</span>
        </div>
        <div class="file-body">
          <div class="file-name">${escapeHtml(item.name)}</div>
          <div class="file-sub-meta">
            <span>${formatBytes(item.size)}</span>
            <span>•</span>
            <span>${formatDate(item.lastModifiedDateTime)}</span>
          </div>
        </div>
      `;
      UI.recycleGrid.appendChild(card);
    });

  } catch (err) {
    UI.recycleGrid.innerHTML = "";
    UI.emptyRecycleState.style.display = "flex";
  }
}

/* ═══════════════ 6. معاينة الملفات (Preview Modal) ═══════════════ */

async function openPreviewModal(file) {
  if (!file) return;
  UI.previewFileName.textContent = file.name;

  if (UI.previewDownloadBtn) {
    if (file.downloadUrl) {
      UI.previewDownloadBtn.style.display = "inline-flex";
      UI.previewDownloadBtn.onclick = (e) => {
        e.preventDefault();
        triggerDirectDownload(file.downloadUrl, file.name);
      };
    } else {
      UI.previewDownloadBtn.style.display = "none";
    }
  }

  UI.previewViewer.innerHTML = `
    <div class="preview-loading-spinner" style="display:flex;flex-direction:column;align-items:center;justify-content:center;padding:3rem;color:var(--ink-300);">
      <i class="fa-solid fa-spinner fa-spin" style="font-size:2rem;margin-bottom:0.75rem;color:var(--gold);"></i>
      <span>جاري إعداد المعاينة...</span>
    </div>
  `;
  openModal(UI.previewModal);

  try {
    const previewData = await getFilePreview(file.id);
    renderPreviewContent(file, previewData.previewUrl || previewData.downloadUrl || file.downloadUrl);
  } catch (err) {
    renderPreviewContent(file, file.downloadUrl);
  }
}

function renderPreviewContent(file, url) {
  const ext = (file.name.split(".").pop() || "").toLowerCase();
  const mime = file.mimeType || "";

  const isImage = ["jpg", "jpeg", "png", "gif", "webp", "svg", "bmp"].includes(ext) || mime.startsWith("image/");
  const isPdf = ext === "pdf" || mime.includes("pdf");
  const isVideo = ["mp4", "webm", "mov"].includes(ext) || mime.startsWith("video/");
  const isAudio = ["mp3", "wav", "m4a"].includes(ext) || mime.startsWith("audio/");

  if (isImage && url) {
    UI.previewViewer.innerHTML = `<img src="${url}" alt="${escapeHtml(file.name)}" class="preview-media-img"/>`;
    return;
  }

  if (isVideo && url) {
    UI.previewViewer.innerHTML = `<video src="${url}" controls autoplay class="preview-media-video"></video>`;
    return;
  }

  if (isAudio && url) {
    UI.previewViewer.innerHTML = `<audio src="${url}" controls autoplay class="preview-media-audio"></audio>`;
    return;
  }

  if ((isPdf || ext === "pdf") && url) {
    UI.previewViewer.innerHTML = `<iframe src="${url}" class="preview-iframe-viewer" title="${escapeHtml(file.name)}"></iframe>`;
    return;
  }

  if (url && (url.includes("sharepoint.com") || url.includes("office.com") || url.includes("embed") || url.includes("preview"))) {
    UI.previewViewer.innerHTML = `<iframe src="${url}" class="preview-iframe-viewer" title="${escapeHtml(file.name)}"></iframe>`;
    return;
  }

  renderFallbackPreview(file);
}

function renderFallbackPreview(file) {
  const typeInfo = getFileTypeDetails(file.name, file.mimeType);
  UI.previewViewer.innerHTML = `
    <div class="preview-unsupported-box">
      <i class="${typeInfo.icon}"></i>
      <h3>${escapeHtml(file.name)}</h3>
      <p>${formatBytes(file.size)} • ${typeInfo.label}</p>
      ${file.downloadUrl ? `
        <button type="button" class="btn-primary direct-dl-preview-btn" style="margin-top:1rem;">
          <i class="fa-solid fa-download"></i> تنزيل الملف الآن
        </button>
      ` : ''}
    </div>
  `;

  UI.previewViewer.querySelector(".direct-dl-preview-btn")?.addEventListener("click", () => {
    triggerDirectDownload(file.downloadUrl, file.name);
  });
}

/* ═══════════════ 7. إعادة التسمية (Rename) ═══════════════ */

function openRenameModal(item) {
  State.activeTargetItem = item;
  UI.renameInput.value = item.name;
  UI.renameModalTitle.textContent = item.isFolder ? "إعادة تسمية المجلد" : "إعادة تسمية الملف";
  openModal(UI.renameModal);
  setTimeout(() => UI.renameInput.focus(), 100);
}

UI.renameForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  const newName = UI.renameInput.value.trim();
  if (!newName || !State.activeTargetItem) return;

  if (newName === State.activeTargetItem.name) {
    closeModal(UI.renameModal);
    return;
  }

  const btn = $("#renameSubmitBtn");
  btn.disabled = true;
  btn.innerHTML = '<span>جاري حفظ الاسم الجديد...</span>';

  try {
    const targetId = State.activeTargetItem.id;
    await renameSharePointItem(targetId, newName, State.currentFolderId);

    // تحديث فوري محلي في الذاكرة (Optimistic In-Place Update)
    State.activeTargetItem.name = newName;
    const localItem = State.items.find(i => i.id === targetId);
    if (localItem) localItem.name = newName;

    // تحديث الاسم في المفضلة إذا كان مضافاً
    if (State.currentUser) {
      const favs = getFavorites(State.currentUser.uid);
      const favItem = favs.find(f => f.id === targetId);
      if (favItem) {
        favItem.name = newName;
        saveFavorites(favs, State.currentUser.uid);
      }
    }

    showToast(`تمت إعادة التسمية إلى "${newName}" بنجاح`, "success");
    closeModal(UI.renameModal);

    // إعادة فرز وعرض فورية، ثم مزامنة من السحابة في الخلفية دون وميض
    applySortingAndRender();
    await refreshActiveView(true, true);
  } catch (err) {
    console.error("[Rename Error]:", err);
    showToast(err.message || "تعذر إعادة التسمية في السحابة", "error");
  } finally {
    btn.disabled = false;
    btn.innerHTML = '<span>حفظ التعديل</span>';
  }
});

/* ═══════════════ 8. حذف ملف أو مجلد (Delete) ═══════════════ */

function openDeleteModal(item) {
  if (!State.currentUser?.canDelete) {
    showToast("غير مصرح لك بحذف العناصر من السحابة (صلاحية إدارية مطلوبة)", "warning");
    return;
  }

  State.activeTargetItem = item;
  UI.deleteItemName.textContent = item.name;
  openModal(UI.deleteModal);
}

UI.confirmDeleteBtn.addEventListener("click", async () => {
  if (!State.activeTargetItem) return;

  UI.confirmDeleteBtn.disabled = true;
  UI.confirmDeleteBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> جاري حذف العنصر...';

  try {
    const deletedName = State.activeTargetItem.name;
    const deletedId = State.activeTargetItem.id;
    await deleteSharePointItem(deletedId, State.currentFolderId);

    // إزالة فورية من عناصر الصفحة (Optimistic In-Place Removal)
    State.items = State.items.filter(i => i.id !== deletedId);

    // إزالة من المفضلة إن كان مضافاً
    if (State.currentUser) {
      const favs = getFavorites(State.currentUser.uid);
      const remainingFavs = favs.filter(f => f.id !== deletedId);
      if (remainingFavs.length !== favs.length) {
        saveFavorites(remainingFavs, State.currentUser.uid);
      }
    }

    showToast(`تم حذف "${deletedName}" نهائياً`, "success");
    closeModal(UI.deleteModal);

    // إعادة عرض فورية، ثم مزامنة السحابة وتحديث المساحة
    applySortingAndRender();
    await refreshActiveView(true, true);
  } catch (err) {
    console.error("[Delete Error]:", err);
    showToast(err.message || "تعذر حذف العنصر من السحابة", "error");
  } finally {
    UI.confirmDeleteBtn.disabled = false;
    UI.confirmDeleteBtn.innerHTML = '<i class="fa-solid fa-trash"></i> حذف نهائي';
  }
});

/* ═══════════════ استجابة الجوال ═══════════════ */

UI.mobileMenuBtn.addEventListener("click", () => {
  UI.sidebar.classList.toggle("open");
});

document.addEventListener("click", (e) => {
  if (!UI.sidebar.contains(e.target) && !UI.mobileMenuBtn.contains(e.target)) {
    UI.sidebar.classList.remove("open");
  }
});
