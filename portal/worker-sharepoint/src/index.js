/* ══════════════════════════════════════════════════════════
   Cloudflare Worker — المستودع السحابي لملفات الجمعية (SharePoint API)
   جمعية إرث وحضارة بالقريات
   ──────────────────────────────────────────────────────────
   يعمل على شبكة Cloudflare Edge السحابية الإنتاجية المجانية
   دون الحاجة لـ Firebase Functions أو محاكي محلي (Emulator).
   يربط الموقع مباشرة بـ Microsoft Graph API لـ SharePoint.
══════════════════════════════════════════════════════════ */

function corsHeaders(env) {
  return {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET, POST, PUT, PATCH, DELETE, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Authorization, X-Requested-With",
    "Access-Control-Max-Age": "86400"
  };
}

function json(data, status = 200, extraHeaders = {}) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      ...extraHeaders
    }
  });
}

// كاش مؤقت لرمز الوصول في ذاكرة العامل
let _cachedToken = null;
let _tokenExpiresAt = 0;

async function getGraphAccessToken(env) {
  const now = Date.now();
  if (_cachedToken && now < _tokenExpiresAt - 300000) {
    return _cachedToken;
  }

  const tenantId = env.MICROSOFT_TENANT_ID || "5380057d-dc58-45d5-8ae2-230b3ef6a2ef";
  const clientId = env.MICROSOFT_CLIENT_ID || "e92632b0-43b5-40a0-a2e7-b3130aca7c35";
  const clientSecret = env.MICROSOFT_CLIENT_SECRET;

  if (!clientSecret) {
    console.error("[Graph OAuth Error]: MICROSOFT_CLIENT_SECRET غير مضبوط في متغيرات بيئة العامل (Worker Secrets)");
    throw new Error("فشل المصادقة مع Microsoft Graph API: مفتاح السر MICROSOFT_CLIENT_SECRET غير متاح");
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
    console.error("[Graph OAuth Error]:", errText);
    throw new Error("فشل المصادقة مع Microsoft Graph API");
  }

  const tokenData = await tokenRes.json();
  _cachedToken = tokenData.access_token;
  _tokenExpiresAt = now + ((tokenData.expires_in || 3600) * 1000);
  return _cachedToken;
}

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

export default {
  async fetch(request, env) {
    const cors = corsHeaders(env);

    if (request.method === "OPTIONS") {
      return new Response(null, { headers: cors });
    }

    const url = new URL(request.url);
    const action = url.pathname.replace(/^\/+/, "").replace(/\/+$/, "");

    try {
      const token = await getGraphAccessToken(env);
      const driveId = env.MICROSOFT_DRIVE_ID || "b!vag-u0AS6keay7P1gHkP54gtpOO87ZdFmdHyFfCVxqyBUJAhWFk3TrtY3uYtcmis";

      let body = {};
      if (request.method === "POST" || request.method === "PUT" || request.method === "PATCH") {
        try {
          body = await request.json();
        } catch (e) {
          body = {};
        }
      }

      // 1. الإحصائيات (Stats)
      if (action === "stats" || action === "api/stats") {
        const driveRes = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}`, {
          headers: { Authorization: `Bearer ${token}` }
        });
        if (!driveRes.ok) throw new Error("تعذر جلب بيانات السعة");
        const driveData = await driveRes.json();
        return json({
          ok: true,
          name: driveData.name || "مستودع ملفات الجمعية",
          quota: driveData.quota || {
            total: 27487790694400,
            used: 0,
            remaining: 27487790694400,
            state: "normal"
          }
        }, 200, cors);
      }

      // 2. تصفح المجلد (List Items)
      if (action === "list" || action === "api/list") {
        const folderId = body.folderId || url.searchParams.get("folderId");
        const listUrl = folderId && folderId !== "root"
          ? `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${encodeURIComponent(folderId)}/children?$top=500`
          : `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children?$top=500`;

        const listRes = await fetch(listUrl, {
          headers: { Authorization: `Bearer ${token}` }
        });

        if (!listRes.ok) {
          return json({ ok: false, error: "المجلد غير موجود أو تم حذفه" }, listRes.status, cors);
        }

        const listData = await listRes.json();
        const rawItems = listData.value || [];
        const items = rawItems.map(sanitizeItem);

        let currentFolder = { id: "root", name: "المستودع الرئيسي", path: "/", isRoot: true };
        if (folderId && folderId !== "root") {
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

        return json({
          ok: true,
          currentFolder,
          items,
          canDelete: true,
          canUpload: true,
          canCreateFolder: true
        }, 200, cors);
      }

      // 3. إنشاء مجلد (Create Folder)
      if (action === "create_folder" || action === "api/create_folder") {
        const folderName = (body.folderName || "").trim().replace(/[\\/:*?"<>|]/g, "_");
        const parentFolderId = body.parentFolderId;

        if (!folderName) {
          return json({ ok: false, error: "يرجى تحديد اسم المجلد" }, 400, cors);
        }

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
            name: folderName,
            folder: {},
            "@microsoft.graph.conflictBehavior": "rename"
          })
        });

        if (!createRes.ok) {
          return json({ ok: false, error: "فشل إنشاء المجلد في SharePoint" }, createRes.status, cors);
        }

        const created = await createRes.json();
        return json({ ok: true, item: sanitizeItem(created) }, 201, cors);
      }

      // 4. رفع ملف (Upload File)
      if (action === "upload" || action === "api/upload") {
        const { parentFolderId, fileName, fileBase64, mimeType } = body;
        if (!fileName || !fileBase64) {
          return json({ ok: false, error: "بيانات الملف غير مكتملة" }, 400, cors);
        }

        const cleanName = fileName.trim().replace(/[\\/:*?"<>|]/g, "_");
        const encodedName = encodeURIComponent(cleanName);
        const isRoot = !parentFolderId || parentFolderId === "root";

        // تحويل base64 إلى Uint8Array
        const binStr = atob(fileBase64);
        const bytes = new Uint8Array(binStr.length);
        for (let i = 0; i < binStr.length; i++) {
          bytes[i] = binStr.charCodeAt(i);
        }

        const uploadUrl = isRoot
          ? `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encodedName}:/content`
          : `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${encodeURIComponent(parentFolderId)}:/${encodedName}:/content`;

        const uploadRes = await fetch(uploadUrl, {
          method: "PUT",
          headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": mimeType || "application/octet-stream"
          },
          body: bytes
        });

        if (!uploadRes.ok) {
          const errTxt = await uploadRes.text();
          console.error("[Worker Upload Error]:", uploadRes.status, errTxt);
          return json({ ok: false, error: "تعذر رفع الملف إلى SharePoint" }, uploadRes.status, cors);
        }

        const uploaded = await uploadRes.json();
        return json({ ok: true, item: sanitizeItem(uploaded) }, 201, cors);
      }

      // 5. إعادة تسمية (Rename Item)
      if (action === "rename" || action === "api/rename") {
        const { itemId, newName } = body;
        if (!itemId || !newName) {
          return json({ ok: false, error: "بيانات إعادة التسمية غير مكتملة" }, 400, cors);
        }

        const cleanName = newName.trim().replace(/[\\/:*?"<>|]/g, "_");
        const rnRes = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${encodeURIComponent(itemId)}`, {
          method: "PATCH",
          headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json"
          },
          body: JSON.stringify({ name: cleanName })
        });

        if (!rnRes.ok) {
          return json({ ok: false, error: "تعذر إعادة التسمية في SharePoint" }, rnRes.status, cors);
        }

        const updated = await rnRes.json();
        return json({ ok: true, item: sanitizeItem(updated) }, 200, cors);
      }

      // 6. حذف عنصر (Delete Item)
      if (action === "delete" || action === "api/delete") {
        const itemId = body.itemId || url.searchParams.get("itemId");
        if (!itemId) {
          return json({ ok: false, error: "معرف العنصر غير متوفر" }, 400, cors);
        }

        const delRes = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${encodeURIComponent(itemId)}`, {
          method: "DELETE",
          headers: { Authorization: `Bearer ${token}` }
        });

        if (!delRes.ok && delRes.status !== 204 && delRes.status !== 404) {
          return json({ ok: false, error: "تعذر حذف العنصر من SharePoint" }, delRes.status, cors);
        }

        return json({ ok: true, itemId }, 200, cors);
      }

      // 7. معاينة (Preview)
      if (action === "preview" || action === "api/preview") {
        const itemId = body.itemId || url.searchParams.get("itemId");
        if (!itemId) {
          return json({ ok: false, error: "معرف الملف غير متوفر" }, 400, cors);
        }

        const itemRes = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${encodeURIComponent(itemId)}`, {
          headers: { Authorization: `Bearer ${token}` }
        });

        if (!itemRes.ok) {
          return json({ ok: false, error: "الملف غير موجود" }, 404, cors);
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
        } catch (e) {}

        return json({
          ok: true,
          id: item.id,
          name: item.name,
          mimeType: item.file?.mimeType || "",
          size: item.size || 0,
          previewUrl: previewUrl || item.webUrl || "",
          downloadUrl: item["@microsoft.graph.downloadUrl"] || item.webUrl || "",
          webUrl: item.webUrl || ""
        }, 200, cors);
      }

      // 8. بحث (Search)
      if (action === "search" || action === "api/search") {
        const q = (body.query || url.searchParams.get("query") || "").toLowerCase().trim();
        if (!q) return json({ ok: true, results: [] }, 200, cors);

        const matched = [];
        const queue = [{ id: null, path: "/" }];
        let scannedCount = 0;

        while (queue.length > 0 && scannedCount < 15) {
          const cur = queue.shift();
          scannedCount++;
          const curUrl = cur.id
            ? `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${encodeURIComponent(cur.id)}/children?$top=100`
            : `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children?$top=100`;

          const sRes = await fetch(curUrl, { headers: { Authorization: `Bearer ${token}` } });
          if (!sRes.ok) continue;

          const sData = await sRes.json();
          for (const itm of sData.value || []) {
            if ((itm.name || "").toLowerCase().includes(q)) {
              matched.push({ ...sanitizeItem(itm), folderPath: cur.path });
            }
            if (itm.folder && itm.folder.childCount > 0 && queue.length < 20) {
              queue.push({
                id: itm.id,
                path: cur.path === "/" ? `/${itm.name}` : `${cur.path}/${itm.name}`
              });
            }
          }
        }

        return json({ ok: true, results: matched }, 200, cors);
      }

      // 9. إنشاء جلسة رفع للملفات الكبيرة (Create Upload Session)
      if (action === "create_upload_session" || action === "api/create_upload_session") {
        const { parentFolderId, fileName, conflictBehavior = "rename" } = body;
        if (!fileName) {
          return json({ ok: false, error: "اسم الملف مطلوب" }, 400, cors);
        }
        const cleanName = fileName.trim().replace(/[\\/:*?"<>|]/g, "_");
        const encodedName = encodeURIComponent(cleanName);
        const isRoot = !parentFolderId || parentFolderId === "root";

        const sessionUrl = isRoot
          ? `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encodedName}:/createUploadSession`
          : `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${encodeURIComponent(parentFolderId)}:/${encodedName}:/createUploadSession`;

        const sessionRes = await fetch(sessionUrl, {
          method: "POST",
          headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json"
          },
          body: JSON.stringify({
            item: {
              "@microsoft.graph.conflictBehavior": conflictBehavior,
              name: cleanName
            }
          })
        });

        if (!sessionRes.ok) {
          const errTxt = await sessionRes.text();
          console.error("[Create Upload Session Error]:", sessionRes.status, errTxt);
          return json({ ok: false, error: "تعذر إنشاء جلسة رفع في SharePoint" }, sessionRes.status, cors);
        }

        const sessionData = await sessionRes.json();
        return json({
          ok: true,
          uploadUrl: sessionData.uploadUrl,
          expirationDateTime: sessionData.expirationDateTime
        }, 200, cors);
      }

      // 10. تفاصيل عنصر مشترك لصفحة المشاركة (Shared Item Details)
      if (action === "shared_item" || action === "api/shared_item") {
        const itemId = body.itemId || url.searchParams.get("itemId");
        if (!itemId) {
          return json({ ok: false, error: "معرف العنصر مطلوب" }, 400, cors);
        }

        const itemRes = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${encodeURIComponent(itemId)}`, {
          headers: { Authorization: `Bearer ${token}` }
        });

        if (!itemRes.ok) {
          return json({ ok: false, error: "العنصر غير موجود في SharePoint" }, 404, cors);
        }

        const item = await itemRes.json();
        const sanitized = sanitizeItem(item);

        let children = [];
        if (sanitized.isFolder) {
          const chRes = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${encodeURIComponent(itemId)}/children?$top=200`, {
            headers: { Authorization: `Bearer ${token}` }
          });
          if (chRes.ok) {
            const chData = await chRes.json();
            children = (chData.value || []).map(sanitizeItem);
          }
        } else {
          try {
            const pRes = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${encodeURIComponent(itemId)}/preview`, {
              method: "POST",
              headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" }
            });
            if (pRes.ok) {
              const pData = await pRes.json();
              sanitized.previewUrl = pData.getUrl || sanitized.webUrl;
            }
          } catch (e) {}
        }

        return json({
          ok: true,
          item: sanitized,
          children
        }, 200, cors);
      }

      // 11. سلة المحذوفات (Recycle Bin)
      if (action === "recycle_bin" || action === "api/recycle_bin") {
        try {
          const rbRes = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/special/recycleBin/children?$top=50`, {
            headers: { Authorization: `Bearer ${token}` }
          });
          if (rbRes.ok) {
            const rbData = await rbRes.json();
            return json({ ok: true, items: (rbData.value || []).map(sanitizeItem) }, 200, cors);
          }
        } catch (e) {}
        return json({ ok: true, items: [] }, 200, cors);
      }

      // مسار تجريبي (Health Check)
      return json({
        ok: true,
        service: "Erth & Hadarah SharePoint Cloud API",
        status: "active",
        timestamp: new Date().toISOString()
      }, 200, cors);

    } catch (err) {
      console.error("[Worker Error]:", err.message);
      return json({ ok: false, error: err.message || "خطأ داخلي في معالجة طلب SharePoint" }, 500, cors);
    }
  }
};
