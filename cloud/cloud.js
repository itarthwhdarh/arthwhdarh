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
  storageData: null
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
  
  // عناصر المشاركة
  sharesCount: $("#sharesCount"),
  sharesTableBody: $("#sharesTableBody"),
  emptySharesState: $("#emptySharesState"),
  
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
    onItemComplete: (uploadedItem, targetFolderId) => {
      // إذا كان المستخدم يتصفح نفس المجلد المستهدف، نحدث الواجهة فورياً
      if (State.activeView === "files" && State.currentFolderId === targetFolderId) {
        loadFolder(targetFolderId, true);
      } else if (State.activeView === "home") {
        loadHomeView();
      }
      loadStorageQuota();
    },
    onQueueComplete: () => {
      showToast("اكتملت جميع عمليات الرفع بنجاح إلى SharePoint", "success");
      loadStorageQuota();
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

  UI.uploadBtn.addEventListener("click", () => fileInput.click());
  UI.homeUploadBtn?.addEventListener("click", () => fileInput.click());
  UI.emptyUploadBtn?.addEventListener("click", () => fileInput.click());
}

/* ═══════════════ نظام التنقل بين أقسام المنصة (Views Router) ═══════════════ */

function switchView(viewName) {
  State.activeView = viewName;

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

  // شريط المسار والترتيب يظهر فقط في استعراض الملفات
  if (viewName === "files") {
    UI.subBar.style.display = "flex";
  } else {
    UI.subBar.style.display = "none";
  }

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

/* ═══════════════ 1. واجهة الرئيسية (Home Dashboard) ═══════════════ */

async function loadHomeView() {
  try {
    // جلب عناصر الجذر من الكاش أو SharePoint
    const rootData = await listFolderItems("root", true);
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
    UI.homeFoldersGrid.innerHTML = "";
    folders.slice(0, 6).forEach(folder => {
      UI.homeFoldersGrid.appendChild(createFolderCardElement(folder));
    });

    // عرض أحدث الملفات
    UI.homeRecentFilesGrid.innerHTML = "";
    const sortedFiles = [...files].sort((a, b) => new Date(b.lastModifiedDateTime || 0) - new Date(a.lastModifiedDateTime || 0));
    sortedFiles.slice(0, 6).forEach(file => {
      UI.homeRecentFilesGrid.appendChild(createFileCardElement(file));
    });

  } catch (err) {
    console.warn("[Load Home Error]:", err);
  }
}

/* ═══════════════ 2. واجهة ملفاتي (File Manager Explorer) ═══════════════ */

export async function loadFolder(folderId = "root", isBack = false) {
  State.currentFolderId = folderId;

  UI.loadingSkeleton.style.display = "block";
  UI.foldersSection.style.display = "none";
  UI.filesSection.style.display = "none";
  UI.emptyState.style.display = "none";
  if (UI.errorState) UI.errorState.style.display = "none";

  try {
    const res = await listFolderItems(folderId === "root" ? null : folderId, true);
    State.currentFolderName = res.currentFolder?.name || "سحابة إرث وحضارة";

    // تحديث سجل المسار
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
        UI.errorStateDesc.textContent = err.message || "تعذر الاتصال بـ SharePoint. يرجى المحاولة مرة أخرى.";
      }
    }
    showToast(err.message || "تعذر جلب ملفات المجلد", "error");
  } finally {
    UI.loadingSkeleton.style.display = "none";
  }
}

function updateBreadcrumbs() {
  UI.breadcrumbNav.innerHTML = "";
  
  State.folderHistory.forEach((item, index) => {
    const isLast = index === State.folderHistory.length - 1;
    
    const bcSpan = document.createElement("span");
    bcSpan.className = `bc-item ${isLast ? "current" : ""}`;
    bcSpan.setAttribute("data-id", item.id);
    
    const icon = item.id === "root" ? '<i class="fa-solid fa-cloud"></i> ' : '<i class="fa-solid fa-folder"></i> ';
    bcSpan.innerHTML = `${icon} ${escapeHtml(item.name)}`;

    if (!isLast) {
      bcSpan.addEventListener("click", () => {
        State.folderHistory = State.folderHistory.slice(0, index + 1);
        loadFolder(item.id, true);
      });
    }

    UI.breadcrumbNav.appendChild(bcSpan);

    if (!isLast) {
      const sep = document.createElement("i");
      sep.className = "fa-solid fa-chevron-left bc-sep";
      UI.breadcrumbNav.appendChild(sep);
    }
  });

  if (State.folderHistory.length > 1) {
    UI.goBackBtn.style.display = "inline-flex";
  } else {
    UI.goBackBtn.style.display = "none";
  }
}

UI.goBackBtn.addEventListener("click", () => {
  if (State.folderHistory.length > 1) {
    State.folderHistory.pop();
    const prev = State.folderHistory[State.folderHistory.length - 1];
    loadFolder(prev.id, true);
  }
});

/* عرض المجلدات فقط */
async function loadFoldersOnlyView() {
  UI.breadcrumbNav.innerHTML = '<span class="bc-item current"><i class="fa-solid fa-folder"></i> جميع مجلدات السحابة</span>';
  UI.goBackBtn.style.display = "none";
  UI.loadingSkeleton.style.display = "block";

  try {
    const rootData = await listFolderItems("root", true);
    const folders = (rootData.items || []).filter(i => i.isFolder);

    UI.foldersSection.style.display = "block";
    UI.filesSection.style.display = "none";
    UI.foldersCount.textContent = folders.length;
    UI.foldersGrid.innerHTML = "";

    folders.forEach(f => UI.foldersGrid.appendChild(createFolderCardElement(f)));
  } catch (e) {
    showToast("تعذر جلب المجلدات", "error");
  } finally {
    UI.loadingSkeleton.style.display = "none";
  }
}

/* عرض العناصر الأخيرة */
async function loadRecentFilesView() {
  UI.breadcrumbNav.innerHTML = '<span class="bc-item current"><i class="fa-solid fa-clock-rotate-left"></i> العناصر الأخيرة</span>';
  UI.goBackBtn.style.display = "none";
  UI.loadingSkeleton.style.display = "block";

  try {
    const rootData = await listFolderItems("root", true);
    const files = (rootData.items || []).filter(i => !i.isFolder);
    files.sort((a, b) => new Date(b.lastModifiedDateTime || 0) - new Date(a.lastModifiedDateTime || 0));

    UI.foldersSection.style.display = "none";
    UI.filesSection.style.display = "block";
    UI.filesCount.textContent = files.length;

    renderFilesGrid(files);
  } catch (e) {
    showToast("تعذر جلب الملفات الأخيرة", "error");
  } finally {
    UI.loadingSkeleton.style.display = "none";
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
    UI.foldersGrid.innerHTML = "";
    folders.forEach(folder => {
      UI.foldersGrid.appendChild(createFolderCardElement(folder));
    });
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

/* بطاقة المجلد الموحدة */
function createFolderCardElement(folder) {
  const isFav = State.currentUser ? isItemFavorite(folder.id, State.currentUser.uid) : false;
  const card = document.createElement("div");
  card.className = "folder-card";
  card.innerHTML = `
    <div class="folder-info">
      <div class="folder-icon">
        <i class="fa-solid fa-folder"></i>
      </div>
      <div class="folder-text">
        <span class="folder-name" title="${escapeHtml(folder.name)}">${escapeHtml(folder.name)}</span>
        <span class="folder-meta">${folder.childCount} عنصر</span>
      </div>
    </div>
    <div class="file-actions">
      <button class="star-btn ${isFav ? 'active' : ''}" data-action="fav" title="إضافة للمفضلة">
        <i class="${isFav ? 'fa-solid' : 'fa-regular'} fa-star"></i>
      </button>
      <button class="action-icon-btn" data-action="share" title="مشاركة المجلد">
        <i class="fa-solid fa-share-nodes"></i>
      </button>
      <button class="action-icon-btn" data-action="rename" title="إعادة التسمية">
        <i class="fa-solid fa-pen"></i>
      </button>
      ${State.currentUser?.canDelete ? `
        <button class="action-icon-btn danger" data-action="delete" title="حذف المجلد">
          <i class="fa-solid fa-trash"></i>
        </button>
      ` : ""}
    </div>
  `;

  card.querySelector(".folder-info").addEventListener("click", () => {
    switchView("files");
    loadFolder(folder.id);
  });

  card.querySelector('[data-action="fav"]')?.addEventListener("click", (e) => {
    e.stopPropagation();
    handleToggleFavorite(folder, e.currentTarget);
  });

  card.querySelector('[data-action="share"]')?.addEventListener("click", (e) => {
    e.stopPropagation();
    openShareModal(folder);
  });

  card.querySelector('[data-action="rename"]')?.addEventListener("click", (e) => {
    e.stopPropagation();
    openRenameModal(folder);
  });

  card.querySelector('[data-action="delete"]')?.addEventListener("click", (e) => {
    e.stopPropagation();
    openDeleteModal(folder);
  });

  return card;
}

/* بطاقة الملف الموحدة (Grid) */
function createFileCardElement(file) {
  const typeInfo = getFileTypeDetails(file.name, file.mimeType);
  const isFav = State.currentUser ? isItemFavorite(file.id, State.currentUser.uid) : false;
  const isImage = typeInfo.class === "file-type-image";

  const card = document.createElement("div");
  card.className = "file-card";
  card.innerHTML = `
    <div class="file-thumb">
      ${isImage ? `
        <img src="${file.downloadUrl || ''}" alt="${escapeHtml(file.name)}" loading="lazy" onerror="this.style.display='none';this.nextElementSibling.style.display='flex';"/>
        <div class="file-thumb-icon ${typeInfo.class}" style="display:none;"><i class="${typeInfo.icon}"></i></div>
      ` : `
        <div class="file-thumb-icon ${typeInfo.class}"><i class="${typeInfo.icon}"></i></div>
      `}
      <span class="file-type-badge">${typeInfo.label}</span>
      <button class="star-btn card-star ${isFav ? 'active' : ''}" data-action="fav" title="إضافة للمفضلة">
        <i class="${isFav ? 'fa-solid' : 'fa-regular'} fa-star"></i>
      </button>
    </div>

    <div class="file-body">
      <div class="file-name" title="${escapeHtml(file.name)}">${escapeHtml(file.name)}</div>
      <div class="file-sub-meta">
        <span>${formatBytes(file.size)}</span>
        <span>•</span>
        <span>${formatDate(file.lastModifiedDateTime)}</span>
      </div>
      <div class="file-actions" style="margin-top:0.6rem;justify-content:space-between;">
        <div style="display:flex;gap:0.35rem;">
          <button class="action-icon-btn" data-action="preview" title="معاينة">
            <i class="fa-solid fa-eye"></i>
          </button>
          <a href="${file.downloadUrl || '#'}" download class="action-icon-btn" title="تنزيل">
            <i class="fa-solid fa-download"></i>
          </a>
          <button class="action-icon-btn" data-action="share" title="مشاركة">
            <i class="fa-solid fa-share-nodes"></i>
          </button>
        </div>
        <div style="display:flex;gap:0.35rem;">
          <button class="action-icon-btn" data-action="rename" title="إعادة التسمية">
            <i class="fa-solid fa-pen"></i>
          </button>
          ${State.currentUser?.canDelete ? `
            <button class="action-icon-btn danger" data-action="delete" title="حذف">
              <i class="fa-solid fa-trash"></i>
            </button>
          ` : ""}
        </div>
      </div>
    </div>
  `;

  card.querySelector('[data-action="fav"]')?.addEventListener("click", (e) => {
    e.stopPropagation();
    handleToggleFavorite(file, e.currentTarget);
  });

  card.querySelector('[data-action="preview"]')?.addEventListener("click", (e) => {
    e.stopPropagation();
    openPreviewModal(file);
  });

  card.querySelector('[data-action="share"]')?.addEventListener("click", (e) => {
    e.stopPropagation();
    openShareModal(file);
  });

  card.querySelector('[data-action="rename"]')?.addEventListener("click", (e) => {
    e.stopPropagation();
    openRenameModal(file);
  });

  card.querySelector('[data-action="delete"]')?.addEventListener("click", (e) => {
    e.stopPropagation();
    openDeleteModal(file);
  });

  return card;
}

function renderFilesGrid(files) {
  UI.filesGrid.innerHTML = "";
  files.forEach(file => {
    UI.filesGrid.appendChild(createFileCardElement(file));
  });
}

function renderFilesList(files) {
  UI.filesTableBody.innerHTML = "";

  files.forEach(file => {
    const typeInfo = getFileTypeDetails(file.name, file.mimeType);
    const isFav = State.currentUser ? isItemFavorite(file.id, State.currentUser.uid) : false;
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>
        <button class="star-btn ${isFav ? 'active' : ''}" data-action="fav" title="المفضلة">
          <i class="${isFav ? 'fa-solid' : 'fa-regular'} fa-star"></i>
        </button>
      </td>
      <td>
        <div class="table-file-cell">
          <i class="${typeInfo.icon} ${typeInfo.class}"></i>
          <span class="file-name" title="${escapeHtml(file.name)}">${escapeHtml(file.name)}</span>
        </div>
      </td>
      <td><span class="file-type-badge">${typeInfo.label}</span></td>
      <td style="direction:ltr;text-align:right;">${formatBytes(file.size)}</td>
      <td>${formatDate(file.lastModifiedDateTime)}</td>
      <td>${escapeHtml(file.lastModifiedBy || "مستخدم الجمعية")}</td>
      <td>
        <div style="display:flex;gap:0.35rem;justify-content:center;">
          <button class="action-icon-btn" data-action="preview" title="معاينة">
            <i class="fa-solid fa-eye"></i>
          </button>
          <a href="${file.downloadUrl || '#'}" download class="action-icon-btn" title="تنزيل">
            <i class="fa-solid fa-download"></i>
          </a>
          <button class="action-icon-btn" data-action="share" title="مشاركة">
            <i class="fa-solid fa-share-nodes"></i>
          </button>
          <button class="action-icon-btn" data-action="rename" title="تعديل الاسم">
            <i class="fa-solid fa-pen"></i>
          </button>
          ${State.currentUser?.canDelete ? `
            <button class="action-icon-btn danger" data-action="delete" title="حذف">
              <i class="fa-solid fa-trash"></i>
            </button>
          ` : ""}
        </div>
      </td>
    `;

    tr.querySelector('[data-action="fav"]')?.addEventListener("click", (e) => {
      e.stopPropagation();
      handleToggleFavorite(file, e.currentTarget);
    });

    tr.querySelector('[data-action="preview"]')?.addEventListener("click", () => openPreviewModal(file));
    tr.querySelector('[data-action="share"]')?.addEventListener("click", () => openShareModal(file));
    tr.querySelector('[data-action="rename"]')?.addEventListener("click", () => openRenameModal(file));
    tr.querySelector('[data-action="delete"]')?.addEventListener("click", () => openDeleteModal(file));

    UI.filesTableBody.appendChild(tr);
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

UI.refreshBtn.addEventListener("click", () => {
  showToast("جاري تحديث السحابة من SharePoint...", "info", 1500);
  folderCache.clear();
  if (State.activeView === "home") {
    loadHomeView();
  } else {
    loadFolder(State.currentFolderId, true);
  }
  loadStorageQuota();
});

if (UI.retryBtn) {
  UI.retryBtn.addEventListener("click", () => {
    showToast("جاري إعادة محاولة الاتصال...", "info", 1500);
    folderCache.clear();
    loadFolder(State.currentFolderId, true);
    loadStorageQuota();
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

    showToast(`جاري البحث العميق عن "${q}" في SharePoint...`, "info", 2000);
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

UI.newFolderBtn.addEventListener("click", () => {
  UI.folderNameInput.value = "";
  openModal(UI.newFolderModal);
  setTimeout(() => UI.folderNameInput.focus(), 100);
});

UI.newFolderForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  const name = UI.folderNameInput.value.trim();
  if (!name) return;

  const btn = $("#createFolderSubmitBtn");
  btn.disabled = true;
  btn.innerHTML = '<span>جاري الإنشاء في SharePoint...</span>';

  try {
    const parentId = State.currentFolderId === "root" ? null : State.currentFolderId;
    await createFolderInSharePoint(name, parentId);
    showToast(`تم إنشاء المجلد "${name}" في SharePoint بنجاح`, "success");
    closeModal(UI.newFolderModal);
    loadFolder(State.currentFolderId, true);
    loadStorageQuota();
  } catch (err) {
    console.error("[Create Folder Error]:", err);
    showToast(err.message || "تعذر إنشاء المجلد في SharePoint", "error");
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

  // إظهار خيار رفع الملفات إذا كان مجلداً فقط
  const optUpload = $("#optAccessUpload");
  if (optUpload) {
    optUpload.style.display = item.isFolder ? "flex" : "none";
  }

  // إعادة ضبط الحقول
  UI.enablePasswordToggle.checked = false;
  UI.passwordFieldWrap.style.display = "none";
  UI.sharePasswordInput.value = "";
  UI.enableExpiryToggle.checked = false;
  UI.expiryFieldWrap.style.display = "none";
  UI.shareExpiryInput.value = "";
  UI.shareResultBox.style.display = "none";

  openModal(UI.shareModal);
}

UI.enablePasswordToggle.addEventListener("change", (e) => {
  UI.passwordFieldWrap.style.display = e.target.checked ? "block" : "none";
  if (e.target.checked) UI.sharePasswordInput.focus();
});

UI.enableExpiryToggle.addEventListener("change", (e) => {
  UI.expiryFieldWrap.style.display = e.target.checked ? "block" : "none";
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
    showToast("تم إنشاء رابط المشاركة بنجاح", "success");

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

/* عرض روابط المشاركة المنشأة */
async function loadSharesView() {
  if (!State.currentUser) return;
  UI.sharesTableBody.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:2rem;"><i class="fa-solid fa-spinner fa-spin"></i> جاري جلب روابط المشاركة...</td></tr>';
  UI.emptySharesState.style.display = "none";

  try {
    const shares = await fetchUserShareLinks(State.currentUser.uid);
    UI.sharesCount.textContent = shares.length;

    if (shares.length === 0) {
      UI.sharesTableBody.innerHTML = "";
      UI.emptySharesState.style.display = "flex";
      return;
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
          <div style="display:flex;gap:0.35rem;justify-content:center;">
            <button class="action-icon-btn" data-action="copy" title="نسخ الرابط">
              <i class="fa-solid fa-copy"></i>
            </button>
            ${s.status === "active" ? `
              <button class="action-icon-btn danger" data-action="revoke" title="إلغاء الرابط">
                <i class="fa-solid fa-ban"></i>
              </button>
            ` : ""}
          </div>
        </td>
      `;

      tr.querySelector('[data-action="copy"]')?.addEventListener("click", () => {
        navigator.clipboard.writeText(fullUrl);
        showToast("تم نسخ رابط المشاركة", "success");
      });

      tr.querySelector('[data-action="revoke"]')?.addEventListener("click", async () => {
        if (confirm(`هل تريد بالتأكيد إلغاء رابط المشاركة لـ "${s.itemName}"؟ لن يتمكن أحد من فتحه بعد الآن.`)) {
          await revokeShareLink(s.shareId);
          showToast("تم إلغاء رابط المشاركة بنجاح", "info");
          loadSharesView();
        }
      });

      UI.sharesTableBody.appendChild(tr);
    });

  } catch (err) {
    console.error("[Load Shares Error]:", err);
    showToast("تعذر تحميل روابط المشاركة", "error");
  }
}

/* ═══════════════ 4. نظام المفضلة (Favorites) ═══════════════ */

function handleToggleFavorite(item, btnElement) {
  if (!State.currentUser) return;
  const added = toggleFavorite(item, State.currentUser.uid);
  if (btnElement) {
    if (added) {
      btnElement.classList.add("active");
      btnElement.querySelector("i").className = "fa-solid fa-star";
      showToast(`تمت إضافة "${item.name}" إلى المفضلة`, "success", 2000);
    } else {
      btnElement.classList.remove("active");
      btnElement.querySelector("i").className = "fa-regular fa-star";
      showToast(`تمت إزالة "${item.name}" من المفضلة`, "info", 2000);
    }
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
  UI.recycleGrid.innerHTML = '<div style="grid-column:1/-1;text-align:center;padding:2rem;"><i class="fa-solid fa-spinner fa-spin"></i> جاري فحص سلة مهملات SharePoint...</div>';
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
  UI.previewFileName.textContent = file.name;
  UI.previewDownloadBtn.href = file.downloadUrl || "#";
  UI.previewViewer.innerHTML = '<div class="preview-spinner"><i class="fa-solid fa-spinner fa-spin"></i> جاري إعداد المعاينة...</div>';
  openModal(UI.previewModal);

  try {
    const previewData = await getFilePreview(file.id);
    renderPreviewContent(file, previewData.previewUrl || previewData.downloadUrl);
  } catch (err) {
    renderFallbackPreview(file);
  }
}

function renderPreviewContent(file, url) {
  const typeInfo = getFileTypeDetails(file.name, file.mimeType);

  if (typeInfo.class === "file-type-image") {
    UI.previewViewer.innerHTML = `<img src="${url}" alt="${escapeHtml(file.name)}" class="preview-image"/>`;
    return;
  }

  if (typeInfo.class === "file-type-pdf") {
    UI.previewViewer.innerHTML = `<iframe src="${url}" class="preview-frame" title="${escapeHtml(file.name)}"></iframe>`;
    return;
  }

  if (typeInfo.class === "file-type-video") {
    UI.previewViewer.innerHTML = `<video src="${url}" controls class="preview-media"></video>`;
    return;
  }

  if (typeInfo.class === "file-type-audio") {
    UI.previewViewer.innerHTML = `
      <div class="preview-fallback">
        <i class="fa-solid fa-music" style="font-size:3.5rem;color:var(--gold);"></i>
        <audio src="${url}" controls style="margin-top:1.5rem;width:80%;max-width:400px;"></audio>
      </div>
    `;
    return;
  }

  // مستندات أوفيس أو ملفات أخرى عبر iframe
  if (url && (url.includes("sharepoint.com") || url.includes("office.com"))) {
    UI.previewViewer.innerHTML = `<iframe src="${url}" class="preview-frame" title="${escapeHtml(file.name)}"></iframe>`;
    return;
  }

  renderFallbackPreview(file);
}

function renderFallbackPreview(file) {
  const typeInfo = getFileTypeDetails(file.name, file.mimeType);
  UI.previewViewer.innerHTML = `
    <div class="preview-fallback">
      <i class="${typeInfo.icon}" style="font-size:3.5rem;color:var(--gold);"></i>
      <h3 style="font-size:1.1rem;font-weight:700;">${escapeHtml(file.name)}</h3>
      <p style="color:var(--ink-400);font-size:0.85rem;">${formatBytes(file.size)} • ${typeInfo.label}</p>
      <a href="${file.downloadUrl || '#'}" target="_blank" download class="btn-primary" style="margin-top:0.5rem;">
        <i class="fa-solid fa-download"></i> تنزيل الملف الآن
      </a>
    </div>
  `;
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

  const btn = $("#renameSubmitBtn");
  btn.disabled = true;
  btn.innerHTML = '<span>جاري التعديل في SharePoint...</span>';

  try {
    await renameSharePointItem(State.activeTargetItem.id, newName, State.currentFolderId);
    showToast(`تمت إعادة التسمية إلى "${newName}" بنجاح`, "success");
    closeModal(UI.renameModal);
    loadFolder(State.currentFolderId, true);
  } catch (err) {
    console.error("[Rename Error]:", err);
    showToast(err.message || "تعذر إعادة التسمية في SharePoint", "error");
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
  UI.confirmDeleteBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> جاري الحذف من SharePoint...';

  try {
    await deleteSharePointItem(State.activeTargetItem.id, State.currentFolderId);
    showToast(`تم حذف "${State.activeTargetItem.name}" نهائياً من SharePoint`, "success");
    closeModal(UI.deleteModal);
    loadFolder(State.currentFolderId, true);
    loadStorageQuota();
  } catch (err) {
    console.error("[Delete Error]:", err);
    showToast(err.message || "تعذر حذف العنصر من SharePoint", "error");
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
