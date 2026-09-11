/* ══════════════════════════════════════════════════════════
   محرك صفحة المشاركة العامة — سحابة إرث وحضارة (share.js)
   جمعية إرث وحضارة بالقريات
   ──────────────────────────────────────────────────────────
   التحقق من صحة الروابط وصلاحيتها
   فك التشفير والتحقق من كلمات المرور عبر Salted SHA-256
   تنزيل الملفات مباشرة إلى جهاز المستخدم دون فتح SharePoint
   واجهة بسيطة وواضحة تركز على اسم الملف/المجلد ومحتوياته فقط
   خلو تام من الرموز التعبيرية (Emojis)
══════════════════════════════════════════════════════════ */

import {
  fetchShareByToken,
  verifySharePassword,
  recordShareOpen,
  recordShareView,
  recordShareDownload,
  fetchSharedItemDetails
} from "./services.js";

const $ = (s) => document.querySelector(s);

const UI = {
  loading: $("#shareLoading"),
  errorBox: $("#shareErrorBox"),
  errTitle: $("#shareErrTitle"),
  errMsg: $("#shareErrMsg"),
  lockBox: $("#shareLockBox"),
  unlockForm: $("#unlockForm"),
  unlockPassInput: $("#unlockPassInput"),
  toggleUnlockPassBtn: $("#toggleUnlockPassBtn"),
  unlockPassEye: $("#unlockPassEye"),
  unlockErrBox: $("#unlockErrBox"),
  unlockSubmitBtn: $("#unlockSubmitBtn"),
  
  contentBox: $("#shareContentBox"),
  singleFileView: $("#sharedSingleFileView"),
  sharedIcon: $("#sharedIcon"),
  sharedTitle: $("#sharedTitle"),
  sharedSize: $("#sharedSize"),
  sharedPreviewBtn: $("#sharedPreviewBtn"),
  sharedDownloadBtn: $("#sharedDownloadBtn"),
  
  sharedFolderSection: $("#sharedFolderSection"),
  sharedFolderTitle: $("#sharedFolderTitle"),
  sharedFolderCount: $("#sharedFolderCount"),
  sharedNavBar: $("#sharedNavBar"),
  shareGoBackBtn: $("#shareGoBackBtn"),
  shareBreadcrumbs: $("#shareBreadcrumbs"),
  shareViewListBtn: $("#shareViewListBtn"),
  shareViewGridBtn: $("#shareViewGridBtn"),
  shareFolderLoading: $("#shareFolderLoading"),
  sharedItemsContainer: $("#sharedItemsContainer"),
  
  previewModal: $("#sharePreviewModal"),
  previewModalFileName: $("#previewModalFileName"),
  closePreviewModalBtn: $("#closePreviewModalBtn"),
  closePreviewFooterBtn: $("#closePreviewFooterBtn"),
  previewLoadingState: $("#previewLoadingState"),
  previewContentArea: $("#previewContentArea"),
  previewDownloadBtn: $("#previewDownloadBtn"),
  
  toastContainer: $("#toastContainer")
};

let currentShareInfo = null;
let currentShareItem = null;
let folderNavStack = []; // [{ id, name, item, children }]
let currentViewMode = localStorage.getItem("share_view_mode") || "list"; // الافتراضي هو وضع القائمة

function showToast(message, type = "info", duration = 4000) {
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

  toast.querySelector(".toast-close").addEventListener("click", () => toast.remove());
  UI.toastContainer.appendChild(toast);
  setTimeout(() => toast.remove(), duration);
}

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
  const sizes = ["بايت", "كيلوبايت", "ميجابايت", "جيجابايت"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + " " + sizes[i];
}

function getFileTypeDetails(fileName = "", mimeType = "") {
  const ext = fileName.split(".").pop().toLowerCase();
  if (ext === "pdf" || mimeType.includes("pdf")) return { class: "file-type-pdf", icon: "fa-solid fa-file-pdf", label: "PDF" };
  if (["doc", "docx", "word"].includes(ext) || mimeType.includes("word")) return { class: "file-type-word", icon: "fa-solid fa-file-word", label: "Word" };
  if (["xls", "xlsx", "sheet"].includes(ext) || mimeType.includes("sheet")) return { class: "file-type-excel", icon: "fa-solid fa-file-excel", label: "Excel" };
  if (["ppt", "pptx"].includes(ext) || mimeType.includes("presentation")) return { class: "file-type-ppt", icon: "fa-solid fa-file-powerpoint", label: "PowerPoint" };
  if (["jpg", "jpeg", "png", "gif", "webp", "svg"].includes(ext) || mimeType.startsWith("image/")) return { class: "file-type-image", icon: "fa-solid fa-file-image", label: "صورة" };
  if (["mp4", "mov", "avi", "mkv", "webm"].includes(ext) || mimeType.startsWith("video/")) return { class: "file-type-video", icon: "fa-solid fa-file-video", label: "فيديو" };
  if (["mp3", "wav", "m4a"].includes(ext) || mimeType.startsWith("audio/")) return { class: "file-type-audio", icon: "fa-solid fa-file-audio", label: "صوت" };
  if (["zip", "rar", "7z", "tar"].includes(ext)) return { class: "file-type-zip", icon: "fa-solid fa-file-zipper", label: "أرشيف" };
  return { class: "file-type-other", icon: "fa-solid fa-file", label: ext.toUpperCase() || "ملف" };
}

function showError(title, msg) {
  UI.loading.style.display = "none";
  UI.lockBox.style.display = "none";
  UI.contentBox.style.display = "none";
  UI.errTitle.textContent = title;
  UI.errMsg.textContent = msg;
  UI.errorBox.style.display = "block";
}

/* ═══════════════ آلية التنزيل المباشر ═══════════════ */

/**
 * تنزيل الملف مباشرة إلى جهاز المستخدم دون فتح أو توجيه إلى SharePoint
 */
function triggerDirectDownload(url, fileName) {
  if (!url || url === "#") {
    showToast("رابط التنزيل غير متوفر حالياً", "error");
    return;
  }

  // تحسين الرابط لضمان تنزيل مباشر
  let directUrl = url;
  if (directUrl.includes("sharepoint.com") && !directUrl.includes("download.aspx") && !directUrl.includes("download=1")) {
    directUrl += (directUrl.includes("?") ? "&" : "?") + "download=1";
  }

  // تسجيل عملية التنزيل في إحصائيات السحابة
  if (currentShareInfo?.share?.shareId) {
    recordShareDownload(currentShareInfo.share.shareId);
  }

  showToast(`جاري بدء تنزيل "${fileName || 'الملف'}"...`, "info", 3000);

  // تشغيل التنزيل عبر رابط خفي بدون target="_blank"
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

// ربط زر التنزيل الرئيسي للملف
UI.sharedDownloadBtn.addEventListener("click", (e) => {
  e.preventDefault();
  if (currentShareItem && currentShareItem.downloadUrl) {
    triggerDirectDownload(currentShareItem.downloadUrl, currentShareItem.name);
  } else {
    showToast("رابط التنزيل غير متوفر حالياً", "error");
  }
});

/* ═══════════════ محرك المعاينة المباشرة للملفات ═══════════════ */

function isPreviewable(fileName = "", mimeType = "") {
  const ext = (fileName.split(".").pop() || "").toLowerCase();
  const previewExts = [
    "png", "jpg", "jpeg", "gif", "webp", "svg", "bmp", "ico",
    "pdf",
    "mp4", "webm", "mov", "mp3", "wav", "m4a",
    "txt", "csv", "doc", "docx", "xls", "xlsx", "ppt", "pptx"
  ];
  return previewExts.includes(ext) ||
         mimeType.startsWith("image/") ||
         mimeType.includes("pdf") ||
         mimeType.startsWith("video/") ||
         mimeType.startsWith("audio/");
}

let activePreviewItem = null;

async function openSharePreview(fileItem) {
  if (!fileItem) return;
  activePreviewItem = fileItem;

  // تسجيل مشاهدة الملف في إحصائيات السحابة فقط إذا لم تُسجّل الزيارة لهذه الجلسة
  if (currentShareInfo?.share?.shareId) {
    const sId = currentShareInfo.share.shareId;
    if (!sessionStorage.getItem(`viewed_share_${sId}`)) {
      sessionStorage.setItem(`viewed_share_${sId}`, "true");
      recordShareView(sId);
    }
  }

  const ext = (fileItem.name.split(".").pop() || "").toLowerCase();
  const mime = fileItem.mimeType || "";
  const isImage = ["jpg", "jpeg", "png", "gif", "webp", "svg", "bmp"].includes(ext) || mime.startsWith("image/");
  const isPdf = ext === "pdf" || mime.includes("pdf");
  const isVideo = ["mp4", "webm", "mov"].includes(ext) || mime.startsWith("video/");
  const isAudio = ["mp3", "wav", "m4a"].includes(ext) || mime.startsWith("audio/");

  UI.previewModalFileName.textContent = fileItem.name;
  UI.previewContentArea.innerHTML = "";
  UI.previewLoadingState.style.display = "flex";
  UI.previewModal.style.display = "flex";

  // إدارة زر التنزيل داخل نافذة المعاينة حسب الصلاحية
  const canDownload = Boolean(currentShareInfo?.share?.allowDownload && fileItem.downloadUrl);
  if (canDownload) {
    UI.previewDownloadBtn.style.display = "inline-flex";
  } else {
    UI.previewDownloadBtn.style.display = "none";
  }

  try {
    if (isImage) {
      const img = document.createElement("img");
      img.className = "preview-media-img";
      img.alt = fileItem.name;
      img.onload = () => {
        UI.previewLoadingState.style.display = "none";
      };
      img.onerror = () => {
        UI.previewLoadingState.style.display = "none";
        UI.previewContentArea.innerHTML = `
          <div class="preview-unsupported-box">
            <i class="fa-solid fa-triangle-exclamation"></i>
            <h3>تعذر تحميل الصورة</h3>
            <p>قد يكون الرابط انتهت صلاحيته أو تعذر الوصول للمحتوى.</p>
          </div>
        `;
      };
      img.src = fileItem.downloadUrl;
      UI.previewContentArea.appendChild(img);
      return;
    }

    if (isVideo) {
      UI.previewLoadingState.style.display = "none";
      const video = document.createElement("video");
      video.className = "preview-media-video";
      video.controls = true;
      video.autoplay = true;
      video.src = fileItem.downloadUrl;
      UI.previewContentArea.appendChild(video);
      return;
    }

    if (isAudio) {
      UI.previewLoadingState.style.display = "none";
      const audio = document.createElement("audio");
      audio.className = "preview-media-audio";
      audio.controls = true;
      audio.autoplay = true;
      audio.src = fileItem.downloadUrl;
      UI.previewContentArea.appendChild(audio);
      return;
    }

    // لملفات PDF أو مستندات الأوفيس
    let targetPreviewUrl = fileItem.previewUrl;
    if (!targetPreviewUrl && fileItem.id) {
      try {
        const details = await fetchSharedItemDetails(fileItem.id);
        targetPreviewUrl = details.item?.previewUrl;
      } catch (e) {
        console.warn("[Fetch Item Preview Error]:", e);
      }
    }

    if (targetPreviewUrl) {
      const iframe = document.createElement("iframe");
      iframe.className = "preview-iframe-viewer";
      iframe.title = fileItem.name;
      iframe.onload = () => {
        UI.previewLoadingState.style.display = "none";
      };
      iframe.src = targetPreviewUrl;
      UI.previewContentArea.appendChild(iframe);
    } else if (fileItem.downloadUrl && isPdf) {
      const iframe = document.createElement("iframe");
      iframe.className = "preview-iframe-viewer";
      iframe.title = fileItem.name;
      iframe.onload = () => {
        UI.previewLoadingState.style.display = "none";
      };
      iframe.src = fileItem.downloadUrl;
      UI.previewContentArea.appendChild(iframe);
    } else {
      UI.previewLoadingState.style.display = "none";
      UI.previewContentArea.innerHTML = `
        <div class="preview-unsupported-box">
          <i class="fa-solid fa-file-lines"></i>
          <h3>المعاينة غير متاحة لهذا الملف</h3>
          <p>${canDownload ? "يمكنك تنزيل الملف لعرضه على جهازك." : "هذا الملف للمعاينة فقط ولكن صيغته تتطلب تطبيقاً مخصصاً."}</p>
        </div>
      `;
    }

  } catch (err) {
    console.error("[Open Preview Error]:", err);
    UI.previewLoadingState.style.display = "none";
    UI.previewContentArea.innerHTML = `
      <div class="preview-unsupported-box">
        <i class="fa-solid fa-triangle-exclamation"></i>
        <h3>حدث خطأ أثناء تحميل المعاينة</h3>
        <p>${escapeHtml(err.message || "")}</p>
      </div>
    `;
  }
}

function closeSharePreview() {
  UI.previewModal.style.display = "none";
  UI.previewContentArea.innerHTML = "";
  activePreviewItem = null;
}

// أزرار ومستمعات إغلاق المعاينة
UI.closePreviewModalBtn.addEventListener("click", closeSharePreview);
UI.closePreviewFooterBtn.addEventListener("click", closeSharePreview);
UI.previewModal.addEventListener("click", (e) => {
  if (e.target === UI.previewModal) closeSharePreview();
});
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape" && UI.previewModal.style.display === "flex") {
    closeSharePreview();
  }
});

// زر التنزيل داخل نافذة المعاينة
UI.previewDownloadBtn.addEventListener("click", () => {
  if (activePreviewItem && activePreviewItem.downloadUrl) {
    triggerDirectDownload(activePreviewItem.downloadUrl, activePreviewItem.name);
  }
});

/* ═══════════════ استخراج معرّف المشاركة الذكي ═══════════════ */

/**
 * استخراج معرّف المشاركة بذكاء وموثوقية عالية من عدة مصادر محتملة:
 * 1. معاملات الاستعلام: ?s=... أو ?shareId=... أو ?id=...
 * 2. أجزاء الهاش: #s=... أو #id=... أو #SHARE_ID
 * 3. مقاطع المسار: /share/:id أو /cloud/share/:id
 */
function extractShareId() {
  // 1. معاملات الاستعلام
  try {
    const urlParams = new URLSearchParams(window.location.search);
    const param = urlParams.get("s") || urlParams.get("shareId") || urlParams.get("id");
    if (param && param.trim()) return param.trim();
  } catch (e) {}

  // 2. أجزاء الهاش (Hash Fragment)
  try {
    const hash = window.location.hash ? window.location.hash.replace(/^#/, "").trim() : "";
    if (hash) {
      if (hash.includes("=")) {
        const hashParams = new URLSearchParams(hash);
        const hashVal = hashParams.get("s") || hashParams.get("shareId") || hashParams.get("id");
        if (hashVal && hashVal.trim()) return hashVal.trim();
      } else if (hash.length >= 4) {
        return hash;
      }
    }
  } catch (e) {}

  // 3. مقاطع المسار البرمجي
  try {
    const path = window.location.pathname;
    const match = path.match(/(?:^|\/)(?:cloud\/)?share\/([a-zA-Z0-9_-]+)/);
    if (match && match[1] && !match[1].includes(".html")) {
      return match[1].trim();
    }
  } catch (e) {}

  return null;
}

/* ═══════════════ بدء التشغيل وفحص الرابط ═══════════════ */

async function initSharePage() {
  const shareId = extractShareId();

  if (!shareId) {
    showError("رابط مشاركة غير صالح", "لم يتم توفير معرّف المشاركة في عنوان الصفحة.");
    return;
  }

  // تنظيف شريط العنوان وعرض الرابط بصيغة موحدة نقية ?s=ID دون وسم الهاش
  try {
    const currentParam = new URLSearchParams(window.location.search).get("s");
    if (currentParam !== shareId || window.location.hash) {
      const cleanPath = window.location.pathname.replace(/\/+$/, "") || "/";
      window.history.replaceState(null, "", `${cleanPath}?s=${encodeURIComponent(shareId)}`);
    }
  } catch (e) {}

  try {
    const result = await fetchShareByToken(shareId);
    if (!result.ok) {
      showError(
        result.error === "revoked" ? "هذا الرابط لم يعد متاحاً" : "تعذر الوصول للرابط المشترك",
        result.message || "الرابط غير موجود أو تم إلغاؤه من قِبل المالك."
      );
      return;
    }

    currentShareInfo = result;
    const share = result.share;

    // تسجيل فتح الرابط في إحصائيات السحابة فقط عند زيارة جلسة جديدة (وليس Refresh)
    const sessionOpenKey = `opened_share_${share.shareId}`;
    if (!sessionStorage.getItem(sessionOpenKey)) {
      sessionStorage.setItem(sessionOpenKey, "true");
      recordShareOpen(share.shareId);
    }

    // فحص الحماية بكلمة مرور
    if (share.hasPassword) {
      const isSessionUnlocked = sessionStorage.getItem(`unlocked_${share.shareId}`);
      if (!isSessionUnlocked) {
        UI.loading.style.display = "none";
        UI.lockBox.style.display = "block";
        setTimeout(() => UI.unlockPassInput.focus(), 100);
        return;
      }
    }

    // فتح وعرض المحتوى مباشرة
    await renderUnlockedContent();

  } catch (err) {
    console.error("[Share Load Error]:", err);
    showError("حدث خطأ أثناء تحميل الرابط", err.message || "يرجى التأكد من اتصال الإنترنت وإعادة المحاولة.");
  }
}

/* ═══════════════ التحقق من كلمة المرور ═══════════════ */

UI.toggleUnlockPassBtn.addEventListener("click", () => {
  const isPass = UI.unlockPassInput.type === "password";
  UI.unlockPassInput.type = isPass ? "text" : "password";
  UI.unlockPassEye.className = isPass ? "fa-regular fa-eye-slash" : "fa-regular fa-eye";
});

UI.unlockForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  UI.unlockErrBox.classList.remove("visible");

  const inputPass = UI.unlockPassInput.value.trim();
  if (!inputPass) return;

  UI.unlockSubmitBtn.disabled = true;
  UI.unlockSubmitBtn.innerHTML = '<span>جاري التحقق...</span>';

  try {
    const isValid = await verifySharePassword(currentShareInfo, inputPass);
    if (!isValid) {
      UI.unlockErrBox.classList.add("visible");
      UI.unlockPassInput.focus();
      return;
    }

    // حفظ حالة الفك للجلسة الحالية
    sessionStorage.setItem(`unlocked_${currentShareInfo.share.shareId}`, "true");
    UI.lockBox.style.display = "none";
    await renderUnlockedContent();

  } catch (err) {
    console.error("[Password Check Error]:", err);
    UI.unlockErrBox.classList.add("visible");
  } finally {
    UI.unlockSubmitBtn.disabled = false;
    UI.unlockSubmitBtn.innerHTML = '<span>فتح المحتوى المشترك</span> <div class="btn-icon-wrapper"><i class="fa-solid fa-unlock"></i></div>';
  }
});

/* ═══════════════ عرض المحتوى بعد الفك ═══════════════ */

async function renderUnlockedContent() {
  UI.loading.style.display = "flex";
  const share = currentShareInfo.share;

  // تسجيل المشاهدة في Firestore فقط عند زيارة جلسة جديدة (وليس Refresh)
  const sessionViewKey = `viewed_share_${share.shareId}`;
  if (!sessionStorage.getItem(sessionViewKey)) {
    sessionStorage.setItem(sessionViewKey, "true");
    recordShareView(share.shareId);
  }

  try {
    const details = await fetchSharedItemDetails(share.itemId);
    const item = details.item;
    currentShareItem = item;

    UI.loading.style.display = "none";
    UI.contentBox.style.display = "block";

    if (item.isFolder) {
      // استعراض محتويات المجلد
      UI.singleFileView.style.display = "none";
      UI.sharedFolderSection.style.display = "block";

      // تهيئة سجل تنقل المجلدات
      folderNavStack = [{
        id: item.id,
        name: item.name,
        item: item,
        children: details.children || []
      }];

      setViewMode(currentViewMode, false);
      renderFolderCurrentView();
    } else {
      // استعراض ملف منفرد: الاسم والحجم والمعاينة وزر التنزيل
      UI.sharedFolderSection.style.display = "none";
      UI.singleFileView.style.display = "block";

      UI.sharedTitle.textContent = item.name;
      UI.sharedSize.textContent = formatBytes(item.size);

      const typeInfo = getFileTypeDetails(item.name, item.mimeType);
      UI.sharedIcon.innerHTML = `<i class="${typeInfo.icon}"></i>`;

      const canPrev = isPreviewable(item.name, item.mimeType);
      if (canPrev) {
        UI.sharedPreviewBtn.style.display = "inline-flex";
        UI.sharedPreviewBtn.onclick = (e) => {
          e.preventDefault();
          openSharePreview(item);
        };
      } else {
        UI.sharedPreviewBtn.style.display = "none";
      }

      if (share.allowDownload && item.downloadUrl) {
        UI.sharedDownloadBtn.style.display = "inline-flex";
      } else {
        UI.sharedDownloadBtn.style.display = "none";
      }
    }

  } catch (err) {
    console.error("[Render Content Error]:", err);
    showError("تعذر قراءة بيانات العنصر من السحابة", err.message || "الملف ربما تم نقله أو حذفه.");
  }
}

/* ═══════════════ إدارة وضع العرض (قائمة / مربعات) ═══════════════ */

function setViewMode(mode, rerender = true) {
  currentViewMode = mode;
  try {
    localStorage.setItem("share_view_mode", mode);
  } catch (e) {}

  if (mode === "grid") {
    UI.sharedItemsContainer?.classList.remove("mode-list");
    UI.sharedItemsContainer?.classList.add("mode-grid");
    UI.shareViewGridBtn?.classList.add("active");
    UI.shareViewListBtn?.classList.remove("active");
  } else {
    UI.sharedItemsContainer?.classList.remove("mode-grid");
    UI.sharedItemsContainer?.classList.add("mode-list");
    UI.shareViewListBtn?.classList.add("active");
    UI.shareViewGridBtn?.classList.remove("active");
  }

  if (rerender && folderNavStack.length > 0) {
    const cur = folderNavStack[folderNavStack.length - 1];
    renderSharedItems(cur.children, currentShareInfo?.share?.allowDownload);
  }
}

UI.shareViewListBtn?.addEventListener("click", () => setViewMode("list"));
UI.shareViewGridBtn?.addEventListener("click", () => setViewMode("grid"));

/* ═══════════════ محرك التنقل داخل المجلدات المشتركة ═══════════════ */

function renderFolderCurrentView() {
  if (folderNavStack.length === 0) return;
  const currentLevel = folderNavStack[folderNavStack.length - 1];

  UI.sharedFolderTitle.textContent = currentLevel.name;
  const count = (currentLevel.children || []).length;
  UI.sharedFolderCount.textContent = `${count} عنصر`;

  // تحديث مسار المجلد (Breadcrumbs)
  renderBreadcrumbs();

  // زر الرجوع
  if (folderNavStack.length > 1) {
    UI.shareGoBackBtn.style.display = "inline-flex";
  } else {
    UI.shareGoBackBtn.style.display = "none";
  }

  // عرض العناصر بالوضع المختار
  renderSharedItems(currentLevel.children, currentShareInfo?.share?.allowDownload);
}

function renderBreadcrumbs() {
  if (!UI.shareBreadcrumbs) return;
  UI.shareBreadcrumbs.innerHTML = "";

  folderNavStack.forEach((entry, idx) => {
    const isLast = idx === folderNavStack.length - 1;

    if (idx > 0) {
      const sep = document.createElement("i");
      sep.className = "fa-solid fa-chevron-left share-bc-sep";
      UI.shareBreadcrumbs.appendChild(sep);
    }

    const item = document.createElement("span");
    item.className = `share-bc-item ${isLast ? "current" : ""}`;
    item.setAttribute("title", entry.name);
    item.innerHTML = `<i class="fa-solid ${idx === 0 ? 'fa-folder-open' : 'fa-folder'}"></i> <span>${escapeHtml(entry.name)}</span>`;

    if (!isLast) {
      item.setAttribute("role", "button");
      item.setAttribute("tabindex", "0");
      item.addEventListener("click", () => {
        navigateToStackIndex(idx);
      });
      item.addEventListener("keydown", (e) => {
        if (e.key === "Enter" || e.key === " ") {
          e.preventDefault();
          navigateToStackIndex(idx);
        }
      });
    }

    UI.shareBreadcrumbs.appendChild(item);
  });
}

function navigateToStackIndex(targetIdx) {
  if (targetIdx >= 0 && targetIdx < folderNavStack.length - 1) {
    folderNavStack = folderNavStack.slice(0, targetIdx + 1);
    renderFolderCurrentView();
  }
}

function navigateBackOneLevel() {
  if (folderNavStack.length > 1) {
    folderNavStack.pop();
    renderFolderCurrentView();
  }
}

UI.shareGoBackBtn?.addEventListener("click", navigateBackOneLevel);

async function openSubFolder(folderItem) {
  if (!folderItem || !folderItem.id) return;
  if (UI.shareFolderLoading) UI.shareFolderLoading.style.display = "flex";
  if (UI.sharedItemsContainer) UI.sharedItemsContainer.style.display = "none";

  try {
    const details = await fetchSharedItemDetails(folderItem.id);
    folderNavStack.push({
      id: folderItem.id,
      name: folderItem.name,
      item: details.item || folderItem,
      children: details.children || []
    });
    renderFolderCurrentView();
  } catch (err) {
    console.error("[Open Subfolder Error]:", err);
    showToast(err.message || "تعذر فتح المجلد المشترك", "error");
  } finally {
    if (UI.shareFolderLoading) UI.shareFolderLoading.style.display = "none";
    if (UI.sharedItemsContainer) UI.sharedItemsContainer.style.display = "";
  }
}

/* ═══════════════ عرض الملفات والمجلدات ═══════════════ */

function renderSharedItems(items, allowDownload) {
  if (!UI.sharedItemsContainer) return;
  UI.sharedItemsContainer.innerHTML = "";

  if (!items || items.length === 0) {
    UI.sharedItemsContainer.innerHTML = `
      <div class="share-empty-folder">
        <i class="fa-solid fa-folder-open"></i>
        <p>هذا المجلد لا يحتوي على ملفات أو مجلدات حالياً.</p>
      </div>
    `;
    return;
  }

  // فرز المجلدات أولاً ثم الملفات أبجدياً
  const sorted = [...items].sort((a, b) => {
    if (a.isFolder && !b.isFolder) return -1;
    if (!a.isFolder && b.isFolder) return 1;
    return (a.name || "").localeCompare(b.name || "", "ar");
  });

  if (currentViewMode === "grid") {
    renderGridItems(sorted, allowDownload);
  } else {
    renderListItems(sorted, allowDownload);
  }
}

/* ── وضع القائمة المرتبة (List View - الافتراضي) ── */
function renderListItems(items, allowDownload) {
  const fragment = document.createDocumentFragment();

  items.forEach(child => {
    const isFolder = Boolean(child.isFolder);
    const typeInfo = getFileTypeDetails(child.name, child.mimeType);
    const canPrev = !isFolder && isPreviewable(child.name, child.mimeType);
    const canDl = !isFolder && Boolean(allowDownload && child.downloadUrl);

    const row = document.createElement("div");
    row.className = `shared-list-item ${isFolder ? "is-folder" : "is-file"}`;

    if (isFolder) {
      row.setAttribute("role", "button");
      row.setAttribute("tabindex", "0");
      row.setAttribute("title", `اضغط لفتح مجلد "${child.name}"`);

      row.innerHTML = `
        <div class="item-primary-col">
          <div class="item-icon-wrap folder-icon-wrap">
            <i class="fa-solid fa-folder"></i>
          </div>
          <div class="item-details">
            <span class="item-name" title="${escapeHtml(child.name)}">${escapeHtml(child.name)}</span>
          </div>
        </div>
        <div class="item-meta-bar">
          <div class="item-meta-info">
            <span class="type-pill folder-pill">مجلد</span>
            <span class="item-size-text">${child.childCount || 0} عنصر</span>
          </div>
          <div class="item-actions-col">
            <button type="button" class="btn-row-open" title="فتح المجلد">
              <span>فتح</span>
              <i class="fa-solid fa-chevron-left"></i>
            </button>
          </div>
        </div>
      `;

      row.addEventListener("click", () => openSubFolder(child));
      row.addEventListener("keydown", (e) => {
        if (e.key === "Enter" || e.key === " ") {
          e.preventDefault();
          openSubFolder(child);
        }
      });
    } else {
      row.innerHTML = `
        <div class="item-primary-col">
          <div class="item-icon-wrap ${typeInfo.class}">
            <i class="${typeInfo.icon}"></i>
          </div>
          <div class="item-details">
            <span class="item-name" title="${escapeHtml(child.name)}">${escapeHtml(child.name)}</span>
          </div>
        </div>
        <div class="item-meta-bar">
          <div class="item-meta-info">
            <span class="type-pill ${typeInfo.class}">${typeInfo.label}</span>
            <span class="item-size-text">${formatBytes(child.size)}</span>
          </div>
          <div class="item-actions-col">
            ${canPrev ? `
              <button type="button" class="btn-row-action btn-row-preview direct-prev-item-btn" title="معاينة الملف">
                <i class="fa-solid fa-eye"></i>
                <span>معاينة</span>
              </button>
            ` : ""}
            ${canDl ? `
              <button type="button" class="btn-row-action btn-row-download direct-dl-item-btn" title="تنزيل الملف">
                <i class="fa-solid fa-download"></i>
                <span>تنزيل</span>
              </button>
            ` : ""}
          </div>
        </div>
      `;

      if (canPrev) {
        row.querySelector(".direct-prev-item-btn")?.addEventListener("click", (e) => {
          e.preventDefault();
          e.stopPropagation();
          openSharePreview(child);
        });
      }

      if (canDl) {
        row.querySelector(".direct-dl-item-btn")?.addEventListener("click", (e) => {
          e.preventDefault();
          e.stopPropagation();
          triggerDirectDownload(child.downloadUrl, child.name);
        });
      }
    }

    fragment.appendChild(row);
  });

  UI.sharedItemsContainer.innerHTML = "";
  UI.sharedItemsContainer.appendChild(fragment);
}

/* ── وضع المربعات والشبكة (Grid View) ── */
function renderGridItems(items, allowDownload) {
  const fragment = document.createDocumentFragment();

  items.forEach(child => {
    const isFolder = Boolean(child.isFolder);
    const typeInfo = getFileTypeDetails(child.name, child.mimeType);
    const canPrev = !isFolder && isPreviewable(child.name, child.mimeType);
    const canDl = !isFolder && Boolean(allowDownload && child.downloadUrl);

    const card = document.createElement("div");
    card.className = `file-card ${isFolder ? "is-folder-card" : ""}`;

    if (isFolder) {
      card.setAttribute("role", "button");
      card.setAttribute("tabindex", "0");
      card.setAttribute("title", `اضغط لفتح مجلد "${child.name}"`);

      card.innerHTML = `
        <div class="file-thumb">
          <div class="file-thumb-icon" style="background:var(--gold-subtle);color:var(--gold);">
            <i class="fa-solid fa-folder"></i>
          </div>
          <span class="file-type-badge">مجلد</span>
        </div>
        <div class="file-body">
          <div class="file-name" title="${escapeHtml(child.name)}">${escapeHtml(child.name)}</div>
          <div class="file-sub-meta">
            <span>${child.childCount || 0} عنصر</span>
          </div>
          <div class="file-card-actions">
            <button type="button" class="btn-preview direct-open-folder-btn" style="width:100%;" title="فتح المجلد">
              <i class="fa-solid fa-folder-open"></i> فتح المجلد
            </button>
          </div>
        </div>
      `;

      card.addEventListener("click", () => openSubFolder(child));
      card.addEventListener("keydown", (e) => {
        if (e.key === "Enter" || e.key === " ") {
          e.preventDefault();
          openSubFolder(child);
        }
      });
    } else {
      card.innerHTML = `
        <div class="file-thumb">
          <div class="file-thumb-icon ${typeInfo.class}">
            <i class="${typeInfo.icon}"></i>
          </div>
          <span class="file-type-badge">${typeInfo.label}</span>
        </div>
        <div class="file-body">
          <div class="file-name" title="${escapeHtml(child.name)}">${escapeHtml(child.name)}</div>
          <div class="file-sub-meta">
            <span>${formatBytes(child.size)}</span>
          </div>
          ${(canPrev || canDl) ? `
            <div class="file-card-actions">
              ${canPrev ? `
                <button type="button" class="btn-preview direct-prev-item-btn" title="معاينة الملف">
                  <i class="fa-solid fa-eye"></i> معاينة
                </button>
              ` : ""}
              ${canDl ? `
                <button type="button" class="btn-download direct-dl-item-btn" title="تنزيل الملف">
                  <i class="fa-solid fa-download"></i> تنزيل
                </button>
              ` : ""}
            </div>
          ` : ""}
        </div>
      `;

      if (canPrev) {
        card.querySelector(".direct-prev-item-btn")?.addEventListener("click", (e) => {
          e.preventDefault();
          e.stopPropagation();
          openSharePreview(child);
        });
      }

      if (canDl) {
        card.querySelector(".direct-dl-item-btn")?.addEventListener("click", (e) => {
          e.preventDefault();
          e.stopPropagation();
          triggerDirectDownload(child.downloadUrl, child.name);
        });
      }
    }

    fragment.appendChild(card);
  });

  UI.sharedItemsContainer.innerHTML = "";
  UI.sharedItemsContainer.appendChild(fragment);
}

// منع تحديد النص الافتراضي للمتصفح عند النقر والنقر المزدوج على العناصر
document.addEventListener("mousedown", (e) => {
  if (e.detail > 1 && !e.target.closest("input, textarea, [contenteditable='true']")) {
    e.preventDefault();
  }
});

document.addEventListener("dblclick", (e) => {
  if (!e.target.closest("input, textarea, [contenteditable='true']")) {
    if (window.getSelection) {
      window.getSelection().removeAllRanges();
    }
  }
});

// تشغيل الفحص الأولي عند تحميل الصفحة
window.addEventListener("DOMContentLoaded", initSharePage);
