/* ══════════════════════════════════════════════════════════
   محرك صفحة المشاركة العامة — سحابة إرث وحضارة (share.js)
   جمعية إرث وحضارة بالقريات
   ──────────────────────────────────────────────────────────
   التحقق من صحة الروابط وصلاحيتها
   فك التشفير والتحقق من كلمات المرور عبر Salted SHA-256
   عرض وتنزيل الملفات ومعاينة محتويات المجلدات المشتركة
   خلو تام من الرموز التعبيرية (Emojis)
══════════════════════════════════════════════════════════ */

import {
  fetchShareByToken,
  verifySharePassword,
  recordShareView,
  fetchSharedItemDetails,
  uploadFileToSharePoint
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
  sharedIcon: $("#sharedIcon"),
  sharedTitle: $("#sharedTitle"),
  sharedSize: $("#sharedSize"),
  sharedDate: $("#sharedDate"),
  sharedSender: $("#sharedSender"),
  sharedPreviewStage: $("#sharedPreviewStage"),
  sharedDownloadBtn: $("#sharedDownloadBtn"),
  accessNoticeText: $("#accessNoticeText"),
  
  sharedFolderSection: $("#sharedFolderSection"),
  sharedFolderCount: $("#sharedFolderCount"),
  sharedFolderUploadZone: $("#sharedFolderUploadZone"),
  sharedFolderFileInput: $("#sharedFolderFileInput"),
  sharedFolderGrid: $("#sharedFolderGrid"),
  
  headerAccessBadge: $("#headerAccessBadge"),
  toastContainer: $("#toastContainer")
};

let currentShareInfo = null;

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

/* ═══════════════ استخراج معرّف المشاركة الذكي ═══════════════ */

/**
 * استخراج معرّف المشاركة بذكاء وموثوقية عالية من عدة مصادر محتملة:
 * 1. معاملات الاستعلام: ?s=... أو ?shareId=... أو ?id=...
 * 2. أجزاء الهاش: #s=... أو #id=... أو #SHARE_ID (حصانة ضد إسقاط الاستعلام عند تحويل 301)
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

    // شارة نوع الوصول في الرأس
    let accessBadgeHtml = `<span class="badge-tag badge-active"><i class="fa-solid fa-download"></i> عرض وتنزيل</span>`;
    if (share.accessType === "view") {
      accessBadgeHtml = `<span class="badge-tag badge-password"><i class="fa-solid fa-eye"></i> عرض فقط</span>`;
    } else if (share.accessType === "upload") {
      accessBadgeHtml = `<span class="badge-tag badge-active"><i class="fa-solid fa-cloud-arrow-up"></i> إيداع ملفات</span>`;
    }
    UI.headerAccessBadge.innerHTML = accessBadgeHtml;

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

  // تسجيل المشاهدة في Firestore
  recordShareView(share.shareId);

  try {
    const details = await fetchSharedItemDetails(share.itemId);
    const item = details.item;

    UI.loading.style.display = "none";
    UI.contentBox.style.display = "block";

    // تعبئة البيانات العامة
    UI.sharedTitle.textContent = item.name;
    UI.sharedSize.textContent = item.isFolder ? `${item.childCount || 0} عنصر` : formatBytes(item.size);
    UI.sharedDate.textContent = `آخر تعديل: ${formatDate(item.lastModifiedDateTime)}`;
    UI.sharedSender.textContent = `بواسطة: ${escapeHtml(share.creatorName || "جمعية إرث وحضارة")}`;

    const typeInfo = getFileTypeDetails(item.name, item.mimeType);
    UI.sharedIcon.innerHTML = `<i class="${typeInfo.icon}"></i>`;

    // إدارة زر التنزيل بحسب الصلاحيات
    if (share.allowDownload && item.downloadUrl) {
      UI.sharedDownloadBtn.href = item.downloadUrl;
      UI.sharedDownloadBtn.style.display = "inline-flex";
      UI.accessNoticeText.innerHTML = '<i class="fa-solid fa-circle-check text-success"></i> المصرح به: المعاينة والتنزيل المباشر من سحابة إرث وحضارة.';
    } else {
      UI.sharedDownloadBtn.style.display = "none";
      UI.accessNoticeText.innerHTML = '<i class="fa-solid fa-eye text-gold"></i> المصرح به: المعاينة فقط (التنزيل معطّل من قِبل المالك).';
    }

    if (item.isFolder) {
      // استعراض محتويات المجلد
      UI.sharedPreviewStage.style.display = "none";
      UI.sharedFolderSection.style.display = "block";
      UI.sharedFolderCount.textContent = (details.children || []).length;
      renderSharedFolderChildren(details.children || [], share.allowDownload);

      // تفعيل رفع الملفات إذا كانت الصلاحية تسمح
      if (share.allowUpload) {
        UI.sharedFolderUploadZone.style.display = "flex";
        setupSharedFolderUpload(item.id);
      }

    } else {
      // استعراض ملف منفرد
      UI.sharedFolderSection.style.display = "none";
      UI.sharedPreviewStage.style.display = "flex";
      renderFilePreview(item, item.previewUrl || item.downloadUrl);
    }

  } catch (err) {
    console.error("[Render Content Error]:", err);
    showError("تعذر قراءة بيانات العنصر من السحابة", err.message || "الملف ربما تم نقله أو حذفه.");
  }
}

function renderFilePreview(item, url) {
  const typeInfo = getFileTypeDetails(item.name, item.mimeType);

  if (typeInfo.class === "file-type-image") {
    UI.sharedPreviewStage.innerHTML = `<img src="${url}" alt="${escapeHtml(item.name)}"/>`;
    return;
  }

  if (typeInfo.class === "file-type-pdf") {
    UI.sharedPreviewStage.innerHTML = `<iframe src="${url}" title="${escapeHtml(item.name)}"></iframe>`;
    return;
  }

  if (typeInfo.class === "file-type-video") {
    UI.sharedPreviewStage.innerHTML = `<video src="${url}" controls style="max-width:100%;max-height:500px;"></video>`;
    return;
  }

  if (typeInfo.class === "file-type-audio") {
    UI.sharedPreviewStage.innerHTML = `
      <div style="text-align:center;padding:2rem;">
        <i class="fa-solid fa-music" style="font-size:3.5rem;color:var(--gold);"></i>
        <div style="margin-top:1rem;font-weight:700;">${escapeHtml(item.name)}</div>
        <audio src="${url}" controls style="margin-top:1.5rem;width:80%;max-width:400px;"></audio>
      </div>
    `;
    return;
  }

  if (url && (url.includes("sharepoint.com") || url.includes("office.com"))) {
    UI.sharedPreviewStage.innerHTML = `<iframe src="${url}" title="${escapeHtml(item.name)}"></iframe>`;
    return;
  }

  UI.sharedPreviewStage.innerHTML = `
    <div style="text-align:center;padding:3rem 1rem;">
      <i class="${typeInfo.icon}" style="font-size:4rem;color:var(--gold);"></i>
      <h3 style="font-size:1.15rem;font-weight:700;margin-top:1rem;">${escapeHtml(item.name)}</h3>
      <p style="color:var(--ink-400);font-size:0.85rem;margin-top:0.25rem;">${formatBytes(item.size)} • ${typeInfo.label}</p>
    </div>
  `;
}

function renderSharedFolderChildren(items, allowDownload) {
  UI.sharedFolderGrid.innerHTML = "";
  if (items.length === 0) {
    UI.sharedFolderGrid.innerHTML = '<div style="grid-column:1/-1;text-align:center;padding:2rem;color:var(--ink-400);">هذا المجلد لا يحتوي على ملفات حالياً.</div>';
    return;
  }

  items.forEach(child => {
    const typeInfo = getFileTypeDetails(child.name, child.mimeType);
    const card = document.createElement("div");
    card.className = "file-card";
    card.innerHTML = `
      <div class="file-thumb">
        <div class="file-thumb-icon ${typeInfo.class}"><i class="${child.isFolder ? 'fa-solid fa-folder' : typeInfo.icon}"></i></div>
        <span class="file-type-badge">${child.isFolder ? 'مجلد' : typeInfo.label}</span>
      </div>
      <div class="file-body">
        <div class="file-name" title="${escapeHtml(child.name)}">${escapeHtml(child.name)}</div>
        <div class="file-sub-meta">
          <span>${child.isFolder ? `${child.childCount} عنصر` : formatBytes(child.size)}</span>
        </div>
        ${!child.isFolder && allowDownload && child.downloadUrl ? `
          <div style="margin-top:0.6rem;">
            <a href="${child.downloadUrl}" download class="btn-secondary" style="width:100%;justify-content:center;padding:0.4rem;font-size:0.8rem;">
              <i class="fa-solid fa-download"></i> تنزيل
            </a>
          </div>
        ` : ""}
      </div>
    `;
    UI.sharedFolderGrid.appendChild(card);
  });
}

function setupSharedFolderUpload(folderId) {
  const zone = UI.sharedFolderUploadZone;
  const input = UI.sharedFolderFileInput;

  zone.addEventListener("click", () => input.click());

  input.addEventListener("change", async (e) => {
    const files = Array.from(e.target.files || []);
    if (files.length === 0) return;

    showToast(`جاري رفع ${files.length} ملف إلى المجلد المشترك...`, "info");
    for (const file of files) {
      try {
        await uploadFileToSharePoint(file, folderId);
        showToast(`تم رفع "${file.name}" بنجاح`, "success");
      } catch (err) {
        showToast(`تعذر رفع "${file.name}"`, "error");
      }
    }

    // إعادة تحميل المحتويات
    renderUnlockedContent();
  });
}

// تشغيل الفحص الأولي عند تحميل الصفحة
window.addEventListener("DOMContentLoaded", initSharePage);
