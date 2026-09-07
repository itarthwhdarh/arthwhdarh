/* ══════════════════════════════════════════════════════════
   المستودع السحابي لملفات الجمعية — Microsoft SharePoint Backend
   جمعية إرث وحضارة بالقريات
   ──────────────────────────────────────────────────────────
   يوفر هذا الملف دوال Cloud Functions الآمنة للتعامل المباشر
   مع SharePoint عبر Microsoft Graph API:
   - جلب وتصفح الملفات والمجلدات
   - إنشاء المجلدات
   - رفع الملفات المفردة والمتعددة
   - إعادة التسمية
   - الحذف الآمن (محمي بالصلاحيات)
   - المعاينة والتنزيل
   - البحث والإحصائيات
══════════════════════════════════════════════════════════ */

const { onCall, HttpsError } = require("firebase-functions/v2/https");
const { getFirestore } = require("firebase-admin/firestore");

const COL_USERS = "portal_users";

/* خيارات CORS المعتمدة */
const allowedCorsOrigins = [
  "https://arthwhdarh.com",
  "https://www.arthwhdarh.com",
  "http://localhost:3000",
  "http://127.0.0.1:3000",
  "http://localhost:5000",
  "http://127.0.0.1:5000",
  "http://localhost:5001",
  "http://127.0.0.1:5001",
  "http://localhost:8080",
  "http://127.0.0.1:8080",
  /^http:\/\/localhost(:\d+)?$/,
  /^http:\/\/127\.0\.0\.1(:\d+)?$/
];

const corsOptions = {
  cors: allowedCorsOrigins,
  maxInstances: 10
};

/* ── إدارة رمز الوصول لـ Microsoft Graph مع كاش مؤقت ── */
let _cachedToken = null;
let _tokenExpiresAt = 0;

async function getGraphAccessToken() {
  const now = Date.now();
  if (_cachedToken && now < _tokenExpiresAt - 300000) { // يتبقى 5 دقائق على الأقل
    return _cachedToken;
  }

  const tenantId = process.env.MICROSOFT_TENANT_ID || "5380057d-dc58-45d5-8ae2-230b3ef6a2ef";
  const clientId = process.env.MICROSOFT_CLIENT_ID || "e92632b0-43b5-40a0-a2e7-b3130aca7c35";
  const clientSecret = process.env.MICROSOFT_CLIENT_SECRET;

  if (!clientSecret) {
    throw new HttpsError("failed-precondition", "مفتاح MICROSOFT_CLIENT_SECRET غير مضبوط في ملف .env الخادمي");
  }

  const tokenParams = new URLSearchParams({
    client_id: clientId,
    scope: "https://graph.microsoft.com/.default",
    client_secret: clientSecret,
    grant_type: "client_credentials"
  });

  const tokenRes = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: tokenParams.toString()
  });

  if (!tokenRes.ok) {
    const errText = await tokenRes.text();
    const safeMsg = errText.replace(new RegExp(clientSecret, "g"), "[SECRET_REDACTED]");
    console.error("[Graph OAuth Error]:", safeMsg);
    throw new HttpsError("internal", "فشل المصادقة مع Microsoft Graph API");
  }

  const tokenData = await tokenRes.json();
  _cachedToken = tokenData.access_token;
  _tokenExpiresAt = now + ((tokenData.expires_in || 3600) * 1000);

  return _cachedToken;
}

function getDriveId() {
  return process.env.MICROSOFT_DRIVE_ID || "b!vag-u0AS6keay7P1gHkP54gtpOO87ZdFmdHyFfCVxqyBUJAhWFk3TrtY3uYtcmis";
}

/* ── التحقق الأمني من هوية وصلاحيات المستخدم المتصل ── */
async function verifyAuthorizedUser(request, requireAdmin = false) {
  if (!request.auth || !request.auth.uid) {
    throw new HttpsError("unauthenticated", "يجب تسجيل الدخول أولاً للوصول إلى المستودع السحابي");
  }

  const db = getFirestore();
  const uid = request.auth.uid;
  const email = (request.auth.token?.email || "").toLowerCase().trim();

  // فحص الحسابات الإدارية الرئيسية المعرفة في النظام
  const isExplicitMaster = uid === "dDcYQljmYFR1gq6cD0lG3wBaHys2" ||
                          email === "2falilh2@gmail.com" ||
                          email === "it@arthwhdarh.com" ||
                          email === "info@arthwhdarh.com" ||
                          email.endsWith("@arthwhdarh.com");

  let userData = null;
  const userSnap = await db.collection(COL_USERS).doc(uid).get();

  if (userSnap.exists) {
    userData = userSnap.data();
  } else {
    // محاولة فحص مجموعة users كخيار بديل
    const fallbackSnap = await db.collection("users").doc(uid).get();
    if (fallbackSnap.exists) {
      userData = fallbackSnap.data();
    } else if (isExplicitMaster) {
      userData = {
        name: email.split("@")[0] || "مسؤول النظام",
        email: email,
        role: "tech_admin",
        isTechAdmin: true,
        status: "active"
      };
    }
  }

  if (!userData && !isExplicitMaster) {
    throw new HttpsError("permission-denied", "هذا الحساب غير مصرح له بالوصول إلى مستودع ملفات الجمعية");
  }

  if (userData && (userData.status === "inactive" || userData.status === "disabled")) {
    throw new HttpsError("permission-denied", "تم إيقاف هذا الحساب، يرجى التواصل مع إدارة الجمعية");
  }

  const role = userData?.role || (isExplicitMaster ? "tech_admin" : "employee");
  const isTechAdmin = Boolean(userData?.isTechAdmin || role === "tech_admin" || isExplicitMaster);
  const isExec = role === "executive" || isTechAdmin;
  const isHR = role === "hr" || isExec;

  if (requireAdmin && !isExec && !isTechAdmin) {
    throw new HttpsError("permission-denied", "غير مصرح لك بتنفيذ هذه العملية؛ تتطلب صلاحية إدارية");
  }

  return {
    uid,
    name: userData?.name || (isExplicitMaster ? "مسؤول النظام" : "مستخدم الجمعية"),
    email: userData?.email || email,
    role,
    isTechAdmin,
    isExec,
    isHR
  };
}

/* تنقية بيانات العنصر القادمة من SharePoint */
function sanitizeItem(item) {
  return {
    id: item.id,
    name: item.name,
    isFolder: Boolean(item.folder),
    childCount: item.folder?.childCount ?? 0,
    size: item.size || 0,
    mimeType: item.file?.mimeType || (item.folder ? "folder" : "application/octet-stream"),
    lastModifiedDateTime: item.lastModifiedDateTime,
    lastModifiedBy: item.lastModifiedBy?.user?.displayName || "مستخدم الجمعية",
    webUrl: item.webUrl || "",
    downloadUrl: item["@microsoft.graph.downloadUrl"] || item.webUrl || "",
    parentReference: item.parentReference ? {
      id: item.parentReference.id,
      path: item.parentReference.path
    } : null
  };
}

/* ═══════════════ دوال Cloud Functions ═══════════════ */

/**
 * 1. إحصائيات المستودع وسعة التخزين (Stats)
 */
exports.spCloudGetStats = onCall(corsOptions, async (request) => {
  const user = await verifyAuthorizedUser(request);
  const token = await getGraphAccessToken();
  const driveId = getDriveId();

  try {
    const driveRes = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}`, {
      headers: { Authorization: `Bearer ${token}` }
    });

    if (!driveRes.ok) {
      throw new Error(`Graph HTTP ${driveRes.status}`);
    }

    const driveData = await driveRes.json();
    return {
      name: driveData.name || "مستودع ملفات الجمعية",
      quota: driveData.quota || {
        total: 27487790694400,
        used: 0,
        remaining: 27487790694400,
        state: "normal"
      },
      userRole: user.role,
      userName: user.name
    };
  } catch (err) {
    console.error("[spCloudGetStats Error]:", err.message);
    throw new HttpsError("internal", "تعذر جلب إحصائيات المستودع السحابي");
  }
});

/**
 * 2. جلب محتويات المجلد (List Items)
 */
exports.spCloudListItems = onCall(corsOptions, async (request) => {
  const user = await verifyAuthorizedUser(request);
  const token = await getGraphAccessToken();
  const driveId = getDriveId();

  const { folderId } = request.data || {};

  try {
    // 1. جلب العناصر داخل المجلد
    const listUrl = folderId
      ? `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${encodeURIComponent(folderId)}/children?$top=500`
      : `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children?$top=500`;

    const listRes = await fetch(listUrl, {
      headers: { Authorization: `Bearer ${token}` }
    });

    if (!listRes.ok) {
      const errTxt = await listRes.text();
      console.error("[spCloudListItems Fetch Error]:", listRes.status, errTxt);
      throw new HttpsError("not-found", "المجلد المطلوب غير موجود أو تم نقله");
    }

    const listData = await listRes.json();
    const rawItems = listData.value || [];
    const items = rawItems.map(sanitizeItem);

    // 2. جلب بيانات المجلد الحالي إذا لم يكن الجذر (Root)
    let currentFolder = {
      id: "root",
      name: "المستودع الرئيسي",
      path: "/",
      isRoot: true
    };

    if (folderId) {
      const metaRes = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${encodeURIComponent(folderId)}`, {
        headers: { Authorization: `Bearer ${token}` }
      });
      if (metaRes.ok) {
        const meta = await metaRes.json();
        currentFolder = {
          id: meta.id,
          name: meta.name,
          path: meta.parentReference?.path ? `${meta.parentReference.path}/${meta.name}` : meta.name,
          parentId: meta.parentReference?.id || "root",
          isRoot: false
        };
      }
    }

    return {
      currentFolder,
      items,
      canDelete: user.isExec || user.isTechAdmin,
      canUpload: true,
      canCreateFolder: true
    };
  } catch (err) {
    if (err instanceof HttpsError) throw err;
    console.error("[spCloudListItems Internal Error]:", err.message);
    throw new HttpsError("internal", "تعذر جلب ملفات المجلد من SharePoint");
  }
});

/**
 * 3. إنشاء مجلد جديد (Create Folder)
 */
exports.spCloudCreateFolder = onCall(corsOptions, async (request) => {
  const user = await verifyAuthorizedUser(request);
  const token = await getGraphAccessToken();
  const driveId = getDriveId();

  const { parentFolderId, folderName } = request.data || {};

  if (!folderName || typeof folderName !== "string" || !folderName.trim()) {
    throw new HttpsError("invalid-argument", "يرجى تحديد اسم صالح للمجلد");
  }

  const cleanName = folderName.trim().replace(/[\\/:*?"<>|]/g, "_");

  try {
    const createUrl = parentFolderId && parentFolderId !== "root"
      ? `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${encodeURIComponent(parentFolderId)}/children`
      : `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children`;

    const createRes = await fetch(createUrl, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        name: cleanName,
        folder: {},
        "@microsoft.graph.conflictBehavior": "rename"
      })
    });

    if (!createRes.ok) {
      const errTxt = await createRes.text();
      console.error("[spCloudCreateFolder Error]:", createRes.status, errTxt);
      throw new HttpsError("internal", "تعذر إنشاء المجلد في SharePoint");
    }

    const created = await createRes.json();
    console.log(`[SharePoint Cloud] Folder created: "${created.name}" by user ${user.uid}`);
    return sanitizeItem(created);
  } catch (err) {
    if (err instanceof HttpsError) throw err;
    console.error("[spCloudCreateFolder Internal Error]:", err.message);
    throw new HttpsError("internal", "حدث خطأ أثناء إنشاء المجلد");
  }
});

/**
 * 4. رفع ملف إلى SharePoint (Upload File)
 */
exports.spCloudUploadFile = onCall(corsOptions, async (request) => {
  const user = await verifyAuthorizedUser(request);
  const token = await getGraphAccessToken();
  const driveId = getDriveId();

  const { parentFolderId, fileName, fileBase64, mimeType } = request.data || {};

  if (!fileName || !fileBase64) {
    throw new HttpsError("invalid-argument", "بيانات الملف غير مكتملة");
  }

  const cleanFileName = fileName.trim().replace(/[\\/:*?"<>|]/g, "_");
  const encodedName = encodeURIComponent(cleanFileName);
  const fileBuffer = Buffer.from(fileBase64, "base64");

  try {
    const isRoot = !parentFolderId || parentFolderId === "root";
    let uploadItem = null;

    // للملفات حتى 4 ميجابايت: استخدام PUT المباشر
    if (fileBuffer.length <= 4 * 1024 * 1024) {
      const uploadUrl = isRoot
        ? `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encodedName}:/content`
        : `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${encodeURIComponent(parentFolderId)}:/${encodedName}:/content`;

      const uploadRes = await fetch(uploadUrl, {
        method: "PUT",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": mimeType || "application/octet-stream"
        },
        body: fileBuffer
      });

      if (!uploadRes.ok) {
        const errTxt = await uploadRes.text();
        console.error("[spCloudUploadFile Direct Error]:", uploadRes.status, errTxt);
        throw new HttpsError("internal", "تعذر رفع الملف إلى SharePoint");
      }

      uploadItem = await uploadRes.json();
    } else {
      // للملفات الأكبر من 4 ميجابايت: استخدام Upload Session المجزأة
      const sessionUrl = isRoot
        ? `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encodedName}:/createUploadSession`
        : `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${encodeURIComponent(parentFolderId)}:/${encodedName}:/createUploadSession`;

      const sessRes = await fetch(sessionUrl, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          item: { "@microsoft.graph.conflictBehavior": "rename" }
        })
      });

      if (!sessRes.ok) {
        const errTxt = await sessRes.text();
        console.error("[spCloudUploadFile Session Create Error]:", sessRes.status, errTxt);
        throw new HttpsError("internal", "تعذر فتح جلسة رفع الملف الكبير");
      }

      const sessData = await sessRes.json();
      const uploadUrl = sessData.uploadUrl;

      const chunkSize = 3200 * 1024;
      const totalSize = fileBuffer.length;
      let offset = 0;

      while (offset < totalSize) {
        const chunkEnd = Math.min(offset + chunkSize, totalSize);
        const chunk = fileBuffer.subarray(offset, chunkEnd);

        const chunkRes = await fetch(uploadUrl, {
          method: "PUT",
          headers: {
            "Content-Length": String(chunk.length),
            "Content-Range": `bytes ${offset}-${chunkEnd - 1}/${totalSize}`
          },
          body: chunk
        });

        if (chunkRes.status === 200 || chunkRes.status === 201) {
          uploadItem = await chunkRes.json();
          break;
        } else if (chunkRes.status === 202) {
          offset = chunkEnd;
        } else {
          const errTxt = await chunkRes.text();
          console.error("[spCloudUploadFile Chunk Error]:", chunkRes.status, errTxt);
          throw new HttpsError("internal", "فشل رفع أحد أجزاء الملف الكبير");
        }
      }
    }

    console.log(`[SharePoint Cloud] File uploaded: "${cleanFileName}" (${fileBuffer.length} bytes) by user ${user.uid}`);
    return sanitizeItem(uploadItem);
  } catch (err) {
    if (err instanceof HttpsError) throw err;
    console.error("[spCloudUploadFile Internal Error]:", err.message);
    throw new HttpsError("internal", "فشل رفع الملف إلى المستودع السحابي");
  }
});

/**
 * 5. إعادة تسمية ملف أو مجلد (Rename Item)
 */
exports.spCloudRenameItem = onCall(corsOptions, async (request) => {
  const user = await verifyAuthorizedUser(request);
  const token = await getGraphAccessToken();
  const driveId = getDriveId();

  const { itemId, newName } = request.data || {};

  if (!itemId || !newName || typeof newName !== "string" || !newName.trim()) {
    throw new HttpsError("invalid-argument", "يرجى تحديد العنصر والاسم الجديد المطلوب");
  }

  const cleanName = newName.trim().replace(/[\\/:*?"<>|]/g, "_");

  try {
    const renameRes = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${encodeURIComponent(itemId)}`, {
      method: "PATCH",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ name: cleanName })
    });

    if (!renameRes.ok) {
      const errTxt = await renameRes.text();
      console.error("[spCloudRenameItem Error]:", renameRes.status, errTxt);
      throw new HttpsError("internal", "تعذر إعادة تسمية العنصر في SharePoint");
    }

    const updated = await renameRes.json();
    console.log(`[SharePoint Cloud] Renamed item ${itemId} to "${cleanName}" by ${user.uid}`);
    return sanitizeItem(updated);
  } catch (err) {
    if (err instanceof HttpsError) throw err;
    console.error("[spCloudRenameItem Internal Error]:", err.message);
    throw new HttpsError("internal", "فشل إعادة تسمية العنصر");
  }
});

/**
 * 6. حذف ملف أو مجلد (Delete Item)
 * حماية صارمة: تتطلب صلاحية إدارة تنفيذي أو تقني
 */
exports.spCloudDeleteItem = onCall(corsOptions, async (request) => {
  const user = await verifyAuthorizedUser(request, true);
  const token = await getGraphAccessToken();
  const driveId = getDriveId();

  const { itemId } = request.data || {};

  if (!itemId) {
    throw new HttpsError("invalid-argument", "معرف العنصر المطلوب حذفه غير متوفر");
  }

  try {
    const delRes = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${encodeURIComponent(itemId)}`, {
      method: "DELETE",
      headers: { Authorization: `Bearer ${token}` }
    });

    if (!delRes.ok && delRes.status !== 204 && delRes.status !== 404) {
      const errTxt = await delRes.text();
      console.error("[spCloudDeleteItem Error]:", delRes.status, errTxt);
      throw new HttpsError("internal", "تعذر حذف العنصر من SharePoint");
    }

    console.log(`[SharePoint Cloud] Item deleted: ${itemId} by admin ${user.uid}`);
    return { success: true, itemId };
  } catch (err) {
    if (err instanceof HttpsError) throw err;
    console.error("[spCloudDeleteItem Internal Error]:", err.message);
    throw new HttpsError("internal", "فشل حذف العنصر من SharePoint");
  }
});

/**
 * 7. معاينة الملف وتجهيز رابط العرض (Preview Item)
 */
exports.spCloudGetPreview = onCall(corsOptions, async (request) => {
  await verifyAuthorizedUser(request);
  const token = await getGraphAccessToken();
  const driveId = getDriveId();

  const { itemId } = request.data || {};

  if (!itemId) {
    throw new HttpsError("invalid-argument", "معرف الملف غير متوفر");
  }

  try {
    const itemRes = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${encodeURIComponent(itemId)}`, {
      headers: { Authorization: `Bearer ${token}` }
    });

    if (!itemRes.ok) {
      throw new HttpsError("not-found", "الملف غير موجود");
    }

    const item = await itemRes.json();

    let previewUrl = null;
    try {
      const prevRes = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${encodeURIComponent(itemId)}/preview`, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json"
        }
      });
      if (prevRes.ok) {
        const prevData = await prevRes.json();
        previewUrl = prevData.getUrl || null;
      }
    } catch (e) {
      console.warn("[spCloudGetPreview] Embedded preview not available:", e.message);
    }

    return {
      id: item.id,
      name: item.name,
      mimeType: item.file?.mimeType || "",
      size: item.size || 0,
      previewUrl: previewUrl || item.webUrl || "",
      downloadUrl: item["@microsoft.graph.downloadUrl"] || item.webUrl || "",
      webUrl: item.webUrl || ""
    };
  } catch (err) {
    if (err instanceof HttpsError) throw err;
    console.error("[spCloudGetPreview Internal Error]:", err.message);
    throw new HttpsError("internal", "تعذر إنشاء معاينة الملف");
  }
});

/**
 * 8. البحث في المستودع السحابي (Search)
 */
exports.spCloudSearch = onCall(corsOptions, async (request) => {
  await verifyAuthorizedUser(request);
  const token = await getGraphAccessToken();
  const driveId = getDriveId();

  const { query } = request.data || {};

  if (!query || typeof query !== "string" || !query.trim()) {
    return { results: [] };
  }

  const cleanQuery = query.trim().toLowerCase();

  try {
    const matched = [];
    const queue = [{ id: null, path: "/" }];
    let scannedCount = 0;
    const maxScan = 15;

    while (queue.length > 0 && scannedCount < maxScan) {
      const current = queue.shift();
      scannedCount++;

      const url = current.id
        ? `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${encodeURIComponent(current.id)}/children?$top=100`
        : `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children?$top=100`;

      const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
      if (!res.ok) continue;

      const data = await res.json();
      const items = data.value || [];

      for (const item of items) {
        const itemName = (item.name || "").toLowerCase();
        if (itemName.includes(cleanQuery)) {
          matched.push({
            ...sanitizeItem(item),
            folderPath: current.path
          });
        }

        if (item.folder && item.folder.childCount > 0 && queue.length < 20) {
          queue.push({
            id: item.id,
            path: current.path === "/" ? `/${item.name}` : `${current.path}/${item.name}`
          });
        }
      }
    }

    return { results: matched };
  } catch (err) {
    console.error("[spCloudSearch Error]:", err.message);
    throw new HttpsError("internal", "تعذر إجراء البحث في المستودع السحابي");
  }
});
