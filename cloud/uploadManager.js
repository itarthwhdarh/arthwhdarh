/* ══════════════════════════════════════════════════════════
   محرك الرفع الخلفي المتقدم — سحابة إرث وحضارة
   Advanced Non-Blocking Background Upload Engine (uploadManager.js)
   ──────────────────────────────────────────────────────────
   - رفع متعدد للملفات في الخلفية دون تجميد الصفحة
   - دعم الملفات الكبيرة عبر Microsoft Graph Upload Sessions (Chunks)
   - تتبع حقيقي ودقيق للبايتات المرفوعة (Real Progress)
   - درج عائم (Floating Upload Drawer) لعرض حالة كل ملف
   - إعادة محاولة للملفات الفاشلة (Retry)
   - خلو تام من الإيموجي (رموز SVG وأيقونات رسمية فقط)
══════════════════════════════════════════════════════════ */

import { uploadFileToSharePoint, createUploadSession, folderCache } from "./services.js";

export class UploadManager {
  constructor(options = {}) {
    this.queue = [];
    this.isProcessing = false;
    this.maxConcurrent = 2;
    this.activeWorkers = 0;
    this.onItemComplete = options.onItemComplete || (() => {});
    this.onQueueComplete = options.onQueueComplete || (() => {});
    this.container = null;
    this.isMinimized = false;
    this.initUI();
  }

  initUI() {
    let el = document.getElementById("floatingUploadDrawer");
    if (!el) {
      el = document.createElement("div");
      el.id = "floatingUploadDrawer";
      el.className = "floating-upload-drawer hidden";
      el.innerHTML = `
        <div class="fud-shell">
          <div class="fud-core">
            <div class="fud-header">
              <div class="fud-title-wrap">
                <i class="fa-solid fa-cloud-arrow-up fud-main-icon"></i>
                <span class="fud-title" id="fudTitleText">جاري الرفع...</span>
              </div>
              <div class="fud-actions">
                <button class="fud-ctrl-btn" id="fudToggleMinimize" title="تصغير/توسيع">
                  <i class="fa-solid fa-chevron-down" id="fudMinIcon"></i>
                </button>
                <button class="fud-ctrl-btn" id="fudCloseBtn" title="إغلاق" style="display:none;">
                  <i class="fa-solid fa-xmark"></i>
                </button>
              </div>
            </div>
            <div class="fud-body" id="fudBody">
              <div class="fud-queue-list" id="fudQueueList"></div>
            </div>
          </div>
        </div>
      `;
      document.body.appendChild(el);
    }
    this.container = el;

    const minBtn = document.getElementById("fudToggleMinimize");
    const closeBtn = document.getElementById("fudCloseBtn");
    const body = document.getElementById("fudBody");
    const minIcon = document.getElementById("fudMinIcon");

    if (minBtn) {
      minBtn.addEventListener("click", () => {
        this.isMinimized = !this.isMinimized;
        if (this.isMinimized) {
          body.style.display = "none";
          minIcon.className = "fa-solid fa-chevron-up";
        } else {
          body.style.display = "block";
          minIcon.className = "fa-solid fa-chevron-down";
        }
      });
    }

    if (closeBtn) {
      closeBtn.addEventListener("click", () => {
        this.clearCompleted();
        this.container.classList.add("hidden");
      });
    }
  }

  addFiles(files, targetFolderId = "root", targetFolderName = "المستودع الرئيسي") {
    const newItems = Array.from(files).map((file, idx) => ({
      id: `up_${Date.now()}_${idx}_${Math.random().toString(36).slice(2, 7)}`,
      file,
      name: file.name,
      size: file.size,
      folderId: targetFolderId,
      folderName: targetFolderName,
      status: "pending", // pending, uploading, success, error
      progress: 0,
      loadedBytes: 0,
      errorMsg: "",
      uploadType: file.size > 4 * 1024 * 1024 ? "chunked" : "direct"
    }));

    this.queue.push(...newItems);
    this.container.classList.remove("hidden");
    if (this.isMinimized) {
      // فتح الدرج إذا كان مصغراً ليرى المستخدم الملفات الجديدة
      this.isMinimized = false;
      document.getElementById("fudBody").style.display = "block";
      document.getElementById("fudMinIcon").className = "fa-solid fa-chevron-down";
    }

    this.updateUI();
    this.processQueue();
  }

  async processQueue() {
    if (this.isProcessing && this.activeWorkers >= this.maxConcurrent) return;
    this.isProcessing = true;

    while (this.activeWorkers < this.maxConcurrent) {
      const nextItem = this.queue.find(i => i.status === "pending");
      if (!nextItem) break;

      this.activeWorkers++;
      this.uploadItem(nextItem).finally(() => {
        this.activeWorkers--;
        this.processQueue();
      });
    }

    if (this.activeWorkers === 0 && !this.queue.some(i => i.status === "pending" || i.status === "uploading")) {
      this.isProcessing = false;
      this.updateUI();
      this.onQueueComplete();
    }
  }

  async uploadItem(item) {
    item.status = "uploading";
    item.progress = 5;
    this.updateItemUI(item);

    try {
      let uploadedResult = null;

      if (item.size > 4 * 1024 * 1024) {
        // رفع بنظام الـ Chunks عبر جلسات رفع Microsoft Graph
        uploadedResult = await this.uploadChunked(item);
      } else {
        // رفع مباشر بنظام Base64 المحسن
        uploadedResult = await uploadFileToSharePoint(
          item.file,
          item.folderId,
          (pct, msg) => {
            item.progress = Math.max(item.progress, pct);
            this.updateItemUI(item);
          }
        );
      }

      item.status = "success";
      item.progress = 100;
      this.updateItemUI(item);
      folderCache.invalidate(item.folderId);
      this.onItemComplete(uploadedResult, item.folderId);

    } catch (err) {
      console.error("[Upload Manager Error]:", item.name, err);
      item.status = "error";
      item.errorMsg = err.message || "فشل الرفع إلى SharePoint";
      this.updateItemUI(item);
    }

    this.updateUI();
  }

  /**
   * رفع الملفات الكبيرة بنظام الأجزاء (Chunked Upload Session)
   */
  async uploadChunked(item) {
    const file = item.file;
    const totalBytes = file.size;
    const chunkSize = 5 * 1024 * 1024; // 5 MB لكل جزء (مطابق لمعايير Microsoft Graph)

    // 1. إنشاء جلسة الرفع
    item.progress = 10;
    this.updateItemUI(item);

    const session = await createUploadSession(file.name, item.folderId, "rename");
    const uploadUrl = session.uploadUrl;

    let start = 0;
    let finalItem = null;

    while (start < totalBytes) {
      const end = Math.min(start + chunkSize, totalBytes);
      const chunk = file.slice(start, end);

      const chunkRes = await fetch(uploadUrl, {
        method: "PUT",
        headers: {
          "Content-Length": String(end - start),
          "Content-Range": `bytes ${start}-${end - 1}/${totalBytes}`
        },
        body: chunk
      });

      if (!chunkRes.ok && chunkRes.status !== 200 && chunkRes.status !== 201 && chunkRes.status !== 202) {
        const errTxt = await chunkRes.text();
        throw new Error(`تعذر رفع جزء الملف (${chunkRes.status}): ${errTxt}`);
      }

      start = end;
      item.loadedBytes = start;
      item.progress = Math.min(98, Math.round((start / totalBytes) * 95) + 5);
      this.updateItemUI(item);

      if (chunkRes.status === 200 || chunkRes.status === 201) {
        finalItem = await chunkRes.json();
      }
    }

    return finalItem;
  }

  retryItem(itemId) {
    const item = this.queue.find(i => i.id === itemId);
    if (!item) return;
    item.status = "pending";
    item.progress = 0;
    item.errorMsg = "";
    this.updateItemUI(item);
    this.processQueue();
  }

  clearCompleted() {
    this.queue = this.queue.filter(i => i.status === "uploading" || i.status === "pending");
    this.updateUI();
  }

  updateUI() {
    const titleText = document.getElementById("fudTitleText");
    const closeBtn = document.getElementById("fudCloseBtn");
    const queueList = document.getElementById("fudQueueList");

    if (!queueList) return;

    const total = this.queue.length;
    const uploadingCount = this.queue.filter(i => i.status === "uploading").length;
    const pendingCount = this.queue.filter(i => i.status === "pending").length;
    const successCount = this.queue.filter(i => i.status === "success").length;
    const errorCount = this.queue.filter(i => i.status === "error").length;

    if (uploadingCount > 0 || pendingCount > 0) {
      titleText.textContent = `جاري رفع ${successCount + 1} من ${total} ملف...`;
      closeBtn.style.display = "none";
    } else {
      if (errorCount > 0) {
        titleText.textContent = `اكتمل الرفع (${successCount} بنجاح، ${errorCount} تعذر)`;
      } else {
        titleText.textContent = `اكتمل رفع جميع الملفات (${total})`;
      }
      closeBtn.style.display = "inline-flex";
    }

    // بناء أو تحديث عناصر الطابور
    this.queue.forEach(item => {
      let card = document.getElementById(`fud_item_${item.id}`);
      if (!card) {
        card = document.createElement("div");
        card.id = `fud_item_${item.id}`;
        card.className = "fud-item";
        queueList.appendChild(card);
      }
      this.renderItemContent(card, item);
    });
  }

  updateItemUI(item) {
    const card = document.getElementById(`fud_item_${item.id}`);
    if (card) {
      this.renderItemContent(card, item);
    }
  }

  renderItemContent(card, item) {
    let statusIcon = '<i class="fa-solid fa-clock text-muted"></i>';
    let statusLabel = "في الانتظار";
    let statusColor = "var(--ink-400)";

    if (item.status === "uploading") {
      statusIcon = '<i class="fa-solid fa-circle-notch fa-spin text-primary"></i>';
      statusLabel = `${item.progress}%`;
      statusColor = "var(--gold)";
    } else if (item.status === "success") {
      statusIcon = '<i class="fa-solid fa-circle-check text-success"></i>';
      statusLabel = "اكتمل";
      statusColor = "var(--success)";
    } else if (item.status === "error") {
      statusIcon = '<i class="fa-solid fa-circle-exclamation text-danger"></i>';
      statusLabel = "تعذر الرفع";
      statusColor = "var(--danger)";
    }

    card.innerHTML = `
      <div class="fud-item-row">
        <div class="fud-item-meta">
          <span class="fud-item-name" title="${escapeHtml(item.name)}">${escapeHtml(item.name)}</span>
          <span class="fud-item-size">${formatBytes(item.size)} • ${escapeHtml(item.folderName)}</span>
        </div>
        <div class="fud-item-status" style="color:${statusColor};">
          ${statusIcon}
          <span class="fud-item-pct">${statusLabel}</span>
          ${item.status === "error" ? `
            <button class="fud-retry-btn" data-retry="${item.id}" title="إعادة المحاولة">
              <i class="fa-solid fa-rotate-right"></i>
            </button>
          ` : ""}
        </div>
      </div>
      <div class="fud-progress-track">
        <div class="fud-progress-fill ${item.status}" style="width:${item.progress}%;"></div>
      </div>
    `;

    const retryBtn = card.querySelector(`[data-retry="${item.id}"]`);
    if (retryBtn) {
      retryBtn.addEventListener("click", () => this.retryItem(item.id));
    }
  }
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
