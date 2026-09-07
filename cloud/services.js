/* ══════════════════════════════════════════════════════════
   سحابة إرث وحضارة — طبقة الخدمات المتقدمة (Services Layer)
   جمعية إرث وحضارة بالقريات
   ──────────────────────────────────────────────────────────
   المصادقة والصلاحيات (Firebase Auth + Firestore)
   محرك التشفير الآمن لكلمات المرور (Web Crypto API)
   نظام المشاركة المحمي بكلمة مرور وصلاحيات دقيقة
   محرك SharePoint السحابي الإنتاجي (Cloudflare Edge Worker)
   نظام التخزين المؤقت الذكي (In-Memory Folder Cache)
══════════════════════════════════════════════════════════ */

import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import {
  getAuth,
  signInWithEmailAndPassword,
  signOut,
  onAuthStateChanged
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";
import {
  getFirestore,
  doc,
  getDoc,
  setDoc,
  updateDoc,
  deleteDoc,
  collection,
  query,
  where,
  getDocs,
  serverTimestamp
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";

/* إعدادات Firebase العامة للمشروع */
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

/* ── نقطة النهاية الإنتاجية لـ SharePoint (Cloudflare Edge) ── */
export const SHAREPOINT_ENDPOINT = "https://erth-sharepoint.2falilh2.workers.dev";

/* ═══════════════ نظام التخزين المؤقت الذكي (Folder Cache) ═══════════════ */
class MemoryFolderCache {
  constructor(ttlMs = 120000) {
    this.cache = new Map();
    this.ttl = ttlMs;
  }

  get(folderId) {
    const key = folderId || "root";
    const entry = this.cache.get(key);
    if (!entry) return null;
    if (Date.now() - entry.time > this.ttl) {
      this.cache.delete(key);
      return null;
    }
    return entry.data;
  }

  set(folderId, data) {
    const key = folderId || "root";
    this.cache.set(key, { data, time: Date.now() });
  }

  invalidate(folderId = null) {
    if (folderId) {
      this.cache.delete(folderId);
    } else {
      this.cache.clear();
    }
  }

  clear() {
    this.cache.clear();
  }
}

export const folderCache = new MemoryFolderCache();

/* ═══════════════ أدوات التشفير الآمن (Web Crypto API) ═══════════════ */

/**
 * توليد Salt عشوائي مشفر بنسبة 100%
 */
export function generateRandomSalt(length = 16) {
  const bytes = new Uint8Array(length);
  crypto.getRandomValues(bytes);
  return Array.from(bytes).map(b => b.toString(16).padStart(2, "0")).join("");
}

/**
 * تشفير كلمة المرور مع Salt باستخدام SHA-256
 * (لا يتم حفظ كلمة المرور كنص صريح إطلاقاً)
 */
export async function hashPasswordWithSalt(password, salt) {
  if (!password || !salt) return "";
  const encoder = new TextEncoder();
  const data = encoder.encode(`${salt}__erth_salt__${password}`);
  const hashBuffer = await crypto.subtle.digest("SHA-256", data);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map(b => b.toString(16).padStart(2, "0")).join("");
}

/**
 * توليد معرّف مشاركة عشوائي فريد وآمن
 */
export function generateShareId(length = 14) {
  const chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
  const bytes = new Uint8Array(length);
  crypto.getRandomValues(bytes);
  let id = "";
  for (let i = 0; i < length; i++) {
    id += chars[bytes[i] % chars.length];
  }
  return id;
}

/* ═══════════════ خدمات المصادقة والمستخدمين ═══════════════ */

export async function loginWithEmail(email, password) {
  const cred = await signInWithEmailAndPassword(auth, email.trim(), password);
  return cred.user;
}

export async function logoutUser() {
  folderCache.clear();
  await signOut(auth);
}

export function onAuthUserChanged(callback) {
  return onAuthStateChanged(auth, callback);
}

export async function verifyAndFetchUserProfile(firebaseUser) {
  const uid = typeof firebaseUser === "string" ? firebaseUser : firebaseUser?.uid;
  const email = (firebaseUser?.email || "").toLowerCase().trim();

  const isExplicitMaster = uid === "dDcYQljmYFR1gq6cD0lG3wBaHys2" ||
                          email === "2falilh2@gmail.com" ||
                          email === "it@arthwhdarh.com" ||
                          email === "info@arthwhdarh.com" ||
                          email.endsWith("@arthwhdarh.com");

  let data = null;
  const docRef = doc(db, "portal_users", uid);
  const snap = await getDoc(docRef);

  if (snap.exists()) {
    data = snap.data();
  } else {
    try {
      const userSnap = await getDoc(doc(db, "users", uid));
      if (userSnap.exists()) {
        data = userSnap.data();
      }
    } catch (e) {}
  }

  if (!data && !isExplicitMaster) {
    return {
      authorized: false,
      reason: "الحساب غير مسجل في قائمة مستخدمي الجمعية المصرح لهم"
    };
  }

  if (data && (data.status === "inactive" || data.status === "disabled")) {
    return {
      authorized: false,
      reason: "تم تعطيل هذا الحساب، يرجى مراجعة إدارة الجمعية"
    };
  }

  const role = data?.role || (isExplicitMaster ? "tech_admin" : "employee");
  const isTechAdmin = Boolean(data?.isTechAdmin || role === "tech_admin" || isExplicitMaster);
  const isExec = role === "executive" || isTechAdmin;
  const isHR = role === "hr" || isExec;

  return {
    authorized: true,
    user: {
      uid,
      name: data?.name || (isExplicitMaster ? "مسؤول النظام التقني" : "مستخدم الجمعية"),
      email: data?.email || email,
      jobTitle: data?.jobTitle || (isExplicitMaster ? "المسؤول التقني" : "موظف"),
      role,
      isTechAdmin,
      isExec,
      isHR,
      canDelete: isExec || isTechAdmin,
      canManage: isHR || isExec || isTechAdmin
    }
  };
}

/* ═══════════════ محرك الاتصال بـ SharePoint الإنتاجي ═══════════════ */

async function callSharePointApi(action, payload = {}) {
  const url = `${SHAREPOINT_ENDPOINT}/${action}`;
  let response;

  try {
    response = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });
  } catch (netErr) {
    console.error("[SharePoint API Network Error]:", netErr);
    throw new Error("تعذر الاتصال بخادم SharePoint السحابي. يرجى التأكد من اتصال الإنترنت وإعادة المحاولة.");
  }

  let data;
  try {
    data = await response.json();
  } catch (jsonErr) {
    throw new Error(`استجابة غير صالحة من خادم SharePoint (HTTP ${response.status})`);
  }

  if (!response.ok || !data.ok) {
    const msg = data?.error || `فشل استدعاء SharePoint (HTTP ${response.status})`;
    throw new Error(msg);
  }

  return data;
}

export async function fetchStorageStats() {
  const data = await callSharePointApi("stats");
  return {
    name: data.name,
    quota: data.quota
  };
}

export async function listFolderItems(folderId = null, useCache = true) {
  const key = folderId && folderId !== "root" ? folderId : "root";
  
  if (useCache) {
    const cached = folderCache.get(key);
    if (cached) return cached;
  }

  const data = await callSharePointApi("list", {
    folderId: key === "root" ? null : key
  });

  const result = {
    currentFolder: data.currentFolder,
    items: data.items || [],
    canDelete: data.canDelete !== false,
    canUpload: true,
    canCreateFolder: true
  };

  folderCache.set(key, result);
  return result;
}

export async function createFolderInSharePoint(folderName, parentFolderId = null) {
  const data = await callSharePointApi("create_folder", {
    folderName,
    parentFolderId: parentFolderId && parentFolderId !== "root" ? parentFolderId : null
  });
  folderCache.invalidate(parentFolderId || "root");
  return data.item;
}

export function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const b64 = reader.result.split(",")[1];
      resolve(b64);
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

export async function uploadFileToSharePoint(file, parentFolderId = null, onProgress = null) {
  if (onProgress) onProgress(15, "جاري قراءة وتجهيز الملف...");
  const base64Data = await fileToBase64(file);

  if (onProgress) onProgress(50, "جاري إرسال الملف إلى SharePoint...");
  const data = await callSharePointApi("upload", {
    fileName: file.name,
    fileBase64: base64Data,
    mimeType: file.type || "application/octet-stream",
    parentFolderId: parentFolderId && parentFolderId !== "root" ? parentFolderId : null
  });

  folderCache.invalidate(parentFolderId || "root");
  if (onProgress) onProgress(100, "اكتمل الرفع بنجاح");
  return data.item;
}

export async function createUploadSession(fileName, parentFolderId = null, conflictBehavior = "rename") {
  const data = await callSharePointApi("create_upload_session", {
    fileName,
    parentFolderId: parentFolderId && parentFolderId !== "root" ? parentFolderId : null,
    conflictBehavior
  });
  return {
    uploadUrl: data.uploadUrl,
    expirationDateTime: data.expirationDateTime
  };
}

export async function renameSharePointItem(itemId, newName, parentFolderId = null) {
  const data = await callSharePointApi("rename", { itemId, newName });
  folderCache.invalidate(parentFolderId || "root");
  return data.item;
}

export async function deleteSharePointItem(itemId, parentFolderId = null) {
  const data = await callSharePointApi("delete", { itemId });
  folderCache.invalidate(parentFolderId || "root");
  return data;
}

export async function getFilePreview(itemId) {
  const data = await callSharePointApi("preview", { itemId });
  return data;
}

export async function searchFilesAndFolders(query) {
  const data = await callSharePointApi("search", { query });
  return data.results || [];
}

export async function fetchRecycleBin() {
  const data = await callSharePointApi("recycle_bin");
  return data.items || [];
}

export async function fetchSharedItemDetails(itemId) {
  const data = await callSharePointApi("shared_item", { itemId });
  return {
    item: data.item,
    children: data.children || []
  };
}

/* ═══════════════ نظام الروابط والمشاركة المشفرة (Share Links Hub) ═══════════════ */

/**
 * توليد رابط مشاركة مرن ومقاوم للتحويلات واختلاف بيئات الاستضافة
 */
export function buildShareUrl(shareId) {
  if (!shareId) return "";
  const loc = window.location;
  const origin = loc.origin;
  const path = loc.pathname;

  let basePath = "";
  if (path.includes("/cloud")) {
    basePath = "/cloud";
  }

  const cleanId = encodeURIComponent(shareId.trim());
  return `${origin}${basePath}/share.html?s=${cleanId}#s=${cleanId}`;
}

/**
 * إنشاء رابط مشاركة جديد محمي بكلمة مرور اختيارية
 */
export async function createShareLink({
  item,
  accessType = "download", // "view" | "download" | "upload" | "edit"
  allowDownload = true,
  allowUpload = false,
  password = "",
  expiresAt = null,
  currentUser
}) {
  if (!item || !item.id) throw new Error("بيانات العنصر غير صالحة");

  const shareId = generateShareId(14);
  const hasPassword = Boolean(password && password.trim().length > 0);
  let passwordSalt = "";
  let passwordHash = "";

  if (hasPassword) {
    passwordSalt = generateRandomSalt(16);
    passwordHash = await hashPasswordWithSalt(password.trim(), passwordSalt);
  }

  const shareDoc = {
    shareId,
    itemId: item.id,
    itemName: item.name,
    isFolder: Boolean(item.isFolder),
    size: item.size || 0,
    mimeType: item.mimeType || "",
    accessType, // "view" | "download" | "upload" | "edit"
    allowDownload: accessType === "view" ? false : Boolean(allowDownload),
    allowUpload: Boolean(item.isFolder && (accessType === "upload" || allowUpload)),
    hasPassword,
    passwordSalt,
    passwordHash,
    expiresAt: expiresAt ? new Date(expiresAt).toISOString() : null,
    status: "active", // "active" | "revoked" | "expired"
    createdBy: currentUser?.uid || "system",
    creatorName: currentUser?.name || "عضو الجمعية",
    creatorEmail: currentUser?.email || "",
    createdAt: new Date().toISOString(),
    viewCount: 0,
    downloadCount: 0
  };

  await setDoc(doc(db, "cloud_shares", shareId), shareDoc);

  // حساب الرابط الكامل بواسطة المحرك الموحد
  const shareUrl = buildShareUrl(shareId);

  return {
    ok: true,
    shareId,
    shareUrl,
    shareDoc
  };
}

/**
 * جلب جميع روابط المشاركة التي أنشأها المستخدم الحالي
 */
export async function fetchUserShareLinks(uid) {
  if (!uid) return [];
  try {
    const q = query(
      collection(db, "cloud_shares"),
      where("createdBy", "==", uid)
    );
    const snap = await getDocs(q);
    const list = [];
    snap.forEach(d => {
      const data = d.data();
      // فحص انتهاء الصلاحية وتحديث الحالة محلياً
      if (data.expiresAt && new Date(data.expiresAt) < new Date() && data.status === "active") {
        data.status = "expired";
      }
      list.push(data);
    });
    list.sort((a, b) => new Date(b.createdAt || 0) - new Date(a.createdAt || 0));
    return list;
  } catch (err) {
    console.warn("[Fetch Share Links Error]:", err);
    return [];
  }
}

/**
 * إلغاء أو حذف رابط مشاركة
 */
export async function revokeShareLink(shareId) {
  if (!shareId) return;
  const ref = doc(db, "cloud_shares", shareId);
  await updateDoc(ref, {
    status: "revoked",
    revokedAt: new Date().toISOString()
  });
}

/**
 * استعلام الرابط المشترك للجمهور العام (صفحة share.html)
 */
export async function fetchShareByToken(shareId) {
  if (!shareId) throw new Error("معرّف المشاركة غير صالح");

  const snap = await getDoc(doc(db, "cloud_shares", shareId));
  if (!snap.exists()) {
    return { ok: false, error: "not_found", message: "رابط المشاركة غير موجود أو تم حذفه" };
  }

  const data = snap.data();

  if (data.status === "revoked") {
    return { ok: false, error: "revoked", message: "هذا الرابط لم يعد متاحاً (تم إلغاؤه من قبل المالك)" };
  }

  if (data.expiresAt && new Date(data.expiresAt) < new Date()) {
    return { ok: false, error: "expired", message: "انتهت صلاحية رابط المشاركة" };
  }

  const isProtected = Boolean(data.hasPassword);

  return {
    ok: true,
    share: {
      shareId: data.shareId,
      itemId: data.itemId,
      itemName: isProtected ? "محتوى محمي بكلمة مرور" : data.itemName,
      isFolder: data.isFolder,
      size: isProtected ? 0 : data.size,
      mimeType: isProtected ? "" : data.mimeType,
      accessType: data.accessType,
      allowDownload: data.allowDownload,
      allowUpload: data.allowUpload,
      hasPassword: isProtected,
      passwordSalt: data.passwordSalt,
      expiresAt: data.expiresAt,
      creatorName: data.creatorName,
      createdAt: data.createdAt,
      status: data.status
    },
    // إخفاء الـ hash للأمان ونتركه فقط للمقارنة البرمجية
    _passwordHash: data.passwordHash,
    _realItemName: data.itemName,
    _realSize: data.size,
    _realMimeType: data.mimeType
  };
}

/**
 * التحقق من صحة كلمة مرور الرابط المشترك
 */
export async function verifySharePassword(shareInfo, enteredPassword) {
  if (!shareInfo || !shareInfo.share?.hasPassword) return true;
  if (!enteredPassword) return false;
  const hash = await hashPasswordWithSalt(enteredPassword.trim(), shareInfo.share.passwordSalt);
  const isValid = hash === shareInfo._passwordHash;
  if (isValid) {
    if (shareInfo._realItemName) shareInfo.share.itemName = shareInfo._realItemName;
    if (shareInfo._realSize !== undefined) shareInfo.share.size = shareInfo._realSize;
    if (shareInfo._realMimeType) shareInfo.share.mimeType = shareInfo._realMimeType;
  }
  return isValid;
}

/**
 * تسجيل زيادة عدد المشاهدات
 */
export async function recordShareView(shareId) {
  try {
    const ref = doc(db, "cloud_shares", shareId);
    const snap = await getDoc(ref);
    if (snap.exists()) {
      const cur = snap.data().viewCount || 0;
      await updateDoc(ref, { viewCount: cur + 1 });
    }
  } catch (e) {}
}

/* ═══════════════ نظام العناصر المفضلة (Favorites) ═══════════════ */

export function getFavorites(uid) {
  const key = `cloud_favs_${uid || "default"}`;
  try {
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : [];
  } catch (e) {
    return [];
  }
}

export function toggleFavorite(item, uid) {
  if (!item || !item.id) return false;
  const key = `cloud_favs_${uid || "default"}`;
  let list = getFavorites(uid);
  const exists = list.some(i => i.id === item.id);
  if (exists) {
    list = list.filter(i => i.id !== item.id);
  } else {
    list.unshift({
      id: item.id,
      name: item.name,
      isFolder: item.isFolder,
      size: item.size || 0,
      mimeType: item.mimeType || "",
      lastModifiedDateTime: item.lastModifiedDateTime,
      starredAt: new Date().toISOString()
    });
  }
  localStorage.setItem(key, JSON.stringify(list));
  return !exists;
}

export function isItemFavorite(itemId, uid) {
  const list = getFavorites(uid);
  return list.some(i => i.id === itemId);
}
