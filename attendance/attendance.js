/* ═══════════════════════════════════════════════════════════════════
   ATTENDANCE.JS — جمعية إرث وحضارة بالقريات
   نظام حضور الموظفين المستقل (Face ID / WebAuthn & Passkeys + Firebase)
   المحرك البرمجي المتكامل · Mobile-First Architecture
   ═══════════════════════════════════════════════════════════════════ */

import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import {
  getAuth,
  signInAnonymously,
  onAuthStateChanged
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";
import {
  getFirestore,
  doc,
  getDoc,
  setDoc,
  updateDoc,
  collection,
  query,
  where,
  getDocs,
  limit,
  serverTimestamp
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

// تهيئة Firebase
const app = initializeApp(FIREBASE_CONFIG);
export const auth = getAuth(app);
export const db = getFirestore(app);

/* ── مفاتيح التخزين المؤقت المحلي (Local Cache Hints Only) ── */
const CACHE_KEY_UID = "erth_attendance_cached_uid";
const CACHE_KEY_NAME = "erth_attendance_cached_name";
const CACHE_KEY_CRED = "erth_attendance_cached_cred_id";

/* ── عناصر واجهة المستخدم ── */
const els = {
  // شاشات
  viewRegister: document.getElementById("viewRegister"),
  viewReturning: document.getElementById("viewReturning"),
  viewSuccess: document.getElementById("viewSuccess"),
  viewDuplicate: document.getElementById("viewDuplicate"),

  // حقول وأزرار
  inputFullName: document.getElementById("inputFullName"),
  btnRegisterFaceId: document.getElementById("btnRegisterFaceId"),
  btnCheckInFaceId: document.getElementById("btnCheckInFaceId"),
  btnDirectVerify: document.getElementById("btnDirectVerify"),
  btnSwitchAccount: document.getElementById("btnSwitchAccount"),
  btnReturnFromSuccess: document.getElementById("btnReturnFromSuccess"),
  btnReturnFromDuplicate: document.getElementById("btnReturnFromDuplicate"),

  // نصوص وعناصر ديناميكية
  returningEmployeeName: document.getElementById("returningEmployeeName"),
  statusBanner: document.getElementById("statusBanner"),
  statusText: document.getElementById("statusText"),
  webviewNotice: document.getElementById("webviewNotice"),

  // ساعة الحضور
  clockDigits: document.getElementById("clockDigits"),
  clockPeriod: document.getElementById("clockPeriod"),
  clockDay: document.getElementById("clockDay"),
  clockDate: document.getElementById("clockDate"),

  // شاشة النجاح
  successGreeting: document.getElementById("successGreeting"),
  successEmpName: document.getElementById("successEmpName"),
  successTime: document.getElementById("successTime"),
  successDate: document.getElementById("successDate"),
  successRecordId: document.getElementById("successRecordId"),
  successUid: document.getElementById("successUid"),
  successStatusBadge: document.getElementById("successStatusBadge"),

  // شاشة التكرار
  dupGreeting: document.getElementById("dupGreeting"),
  dupEmpName: document.getElementById("dupEmpName"),
  dupTime: document.getElementById("dupTime"),
  dupDate: document.getElementById("dupDate"),
  dupRecordId: document.getElementById("dupRecordId"),

  // إطار المسح
  faceidBox: document.getElementById("faceidBox")
};

/* ═══════════════ أدوات تحويل بيانات WebAuthn الثنائية ═══════════════ */

/**
 * تحويل ArrayBuffer إلى Base64URL
 */
function bufferToBase64Url(buffer) {
  const bytes = new Uint8Array(buffer);
  let binary = "";
  for (let i = 0; i < bytes.byteLength; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary)
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");
}

/**
 * تحويل Base64URL إلى ArrayBuffer
 */
function base64UrlToBuffer(base64url) {
  if (!base64url) return new ArrayBuffer(0);
  let base64 = base64url.replace(/-/g, "+").replace(/_/g, "/");
  while (base64.length % 4) {
    base64 += "=";
  }
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes.buffer;
}

/**
 * إنشاء مصفوفة بايتات عشوائية للتحدي الأمني (Cryptographic Challenge)
 */
function generateChallenge() {
  const challenge = new Uint8Array(32);
  window.crypto.getRandomValues(challenge);
  return challenge;
}

/* ═══════════════ التفاعل الصوتي واللمسي (Audio & Haptic) ═══════════════ */

/**
 * نغمة تأكيد بيومترية ناعمة وفاخرة (Apple Pay Chime Style) عبر Web Audio API
 */
function playSuccessChime() {
  try {
    const AudioContext = window.AudioContext || window.webkitAudioContext;
    if (!AudioContext) return;
    const ctx = new AudioContext();

    const now = ctx.currentTime;
    const osc1 = ctx.createOscillator();
    const osc2 = ctx.createOscillator();
    const gain = ctx.createGain();

    osc1.type = "sine";
    osc2.type = "sine";

    // نغمة D5 (587.33 Hz) تنتقل إلى A5 (880 Hz)
    osc1.frequency.setValueAtTime(587.33, now);
    osc1.frequency.exponentialRampToValueAtTime(880, now + 0.12);

    osc2.frequency.setValueAtTime(880, now);
    osc2.frequency.exponentialRampToValueAtTime(1174.66, now + 0.14);

    gain.gain.setValueAtTime(0.18, now);
    gain.gain.exponentialRampToValueAtTime(0.001, now + 0.55);

    osc1.connect(gain);
    osc2.connect(gain);
    gain.connect(ctx.destination);

    osc1.start(now);
    osc2.start(now);
    osc1.stop(now + 0.55);
    osc2.stop(now + 0.55);
  } catch (e) {
    // تجاهل أخطاء تشغيل الصوت إذا منعت سياسة المتصفح
  }
}

/**
 * اهتزاز لمسي هادئ (Haptic Vibration)
 */
function triggerHaptic() {
  if ("vibrate" in navigator) {
    try {
      navigator.vibrate([35, 50, 45]);
    } catch (e) {}
  }
}

/* ═══════════════ صيغ التواريخ والأوقات الدقيقة ═══════════════ */

function getNowDateISO() {
  const d = new Date();
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`; // YYYY-MM-DD
}

function getFormattedArabicDate() {
  const d = new Date();
  return new Intl.DateTimeFormat("ar-SA-u-ca-gregory", {
    weekday: "long",
    day: "numeric",
    month: "long",
    year: "numeric"
  }).format(d);
}

function getFormattedArabicTime() {
  const d = new Date();
  return new Intl.DateTimeFormat("ar-SA", {
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: true
  }).format(d);
}

function getTimeOfDayGreeting() {
  const hour = new Date().getHours();
  return hour < 12 ? "صباح الخير" : "مساء الخير";
}

/**
 * هل توقيت الحضور متأخر؟ (افتراضياً بعد 08:30:00 ص)
 */
function checkIfLate() {
  const now = new Date();
  const hours = now.getHours();
  const minutes = now.getMinutes();
  // إذا كان الوقت بعد 08:30 صباحاً
  if (hours > 8 || (hours === 8 && minutes > 30)) {
    return true;
  }
  return false;
}

/* ═══════════════ إدارة ساعة الحضور الحية ═══════════════ */

function updateLiveClock() {
  const now = new Date();
  let hours = now.getHours();
  const minutes = String(now.getMinutes()).padStart(2, "0");
  const seconds = String(now.getSeconds()).padStart(2, "0");
  const isPm = hours >= 12;

  const period = isPm ? "م" : "ص";
  hours = hours % 12 || 12;
  const hoursStr = String(hours).padStart(2, "0");

  if (els.clockDigits) {
    els.clockDigits.textContent = `${hoursStr}:${minutes}:${seconds}`;
  }
  if (els.clockPeriod) {
    els.clockPeriod.textContent = period;
  }
  if (els.clockDay) {
    els.clockDay.textContent = new Intl.DateTimeFormat("ar-SA", { weekday: "long" }).format(now);
  }
  if (els.clockDate) {
    els.clockDate.textContent = new Intl.DateTimeFormat("ar-SA-u-ca-gregory", {
      day: "numeric",
      month: "long",
      year: "numeric"
    }).format(now);
  }
}

setInterval(updateLiveClock, 1000);
updateLiveClock();

/* ═══════════════ إشعارات وحالات الواجهة ═══════════════ */

function showStatus(message, type = "error") {
  if (!els.statusBanner || !els.statusText) return;
  els.statusText.textContent = message;
  els.statusBanner.className = `status-banner ${type}`;
  els.statusBanner.style.display = "flex";

  // إخفاء التنبيه تلقائياً بعد 7 ثوانٍ إذا لم يكن خطأً ثابتاً
  if (type === "success") {
    setTimeout(() => {
      els.statusBanner.style.display = "none";
    }, 5000);
  }
}

function hideStatus() {
  if (els.statusBanner) {
    els.statusBanner.style.display = "none";
  }
}

function setScanningState(isScanning) {
  if (els.faceidBox) {
    if (isScanning) {
      els.faceidBox.parentElement.classList.add("scanning");
    } else {
      els.faceidBox.parentElement.classList.remove("scanning");
    }
  }
}

/* ═══════════════ فحص دعم المتصفح والبيئة ═══════════════ */

function checkBrowserSupport() {
  // فحص المتصفحات الداخلية في التطبيقات مثل Instagram و WhatsApp
  const ua = navigator.userAgent || navigator.vendor || window.opera;
  const isInApp = /FBAN|FBAV|Instagram|WhatsApp|Telegram|Line/i.test(ua);
  if (isInApp && els.webviewNotice) {
    els.webviewNotice.style.display = "flex";
  }

  // فحص WebAuthn
  if (!window.PublicKeyCredential) {
    showStatus(
      "المتصفح الحالي لا يدعم ميزة Passkeys البيومترية. يرجى فتح الرابط في متصفح Safari على iPhone أو متصفح Chrome.",
      "error"
    );
    if (els.btnRegisterFaceId) els.btnRegisterFaceId.disabled = true;
    if (els.btnCheckInFaceId) els.btnCheckInFaceId.disabled = true;
    return false;
  }

  // التحقق من أن الصفحة تعمل في بيئة آمنة (HTTPS أو Localhost)
  if (!window.isSecureContext) {
    showStatus(
      "تقنية WebAuthn تتطلب اتصالاً آمناً مشفراً (HTTPS) لتعمل بصمة الوجه Face ID بشكل صحيح.",
      "warning"
    );
  }

  return true;
}

/* ═══════════════ إدارة الهوية والمصادقة في Firebase ═══════════════ */

/**
 * الحصول على UID ثابت وحقيقي للموظف عبر Firebase Auth
 */
async function getOrCreateEmployeeUid() {
  // إذا كان المستخدم مصادقاً بالفعل
  if (auth.currentUser && auth.currentUser.uid) {
    return auth.currentUser.uid;
  }

  try {
    const cred = await signInAnonymously(auth);
    return cred.user.uid;
  } catch (err) {
    console.warn("تعذر إنشاء حساب Anonymous عبر Firebase Auth، جاري توليد UID آمن:", err);
    // كود أمان احتياطي عالي الإنتروبيا في حال عدم تفعيل خيار Anonymous في Console
    const fallbackId = "emp_" + crypto.randomUUID().replace(/-/g, "");
    return fallbackId;
  }
}

/**
 * فحص ما إذا كان الموظف قد سجل حضوره اليوم مسبقاً (منع التكرار - شرط 9)
 */
async function checkExistingAttendanceToday(uid) {
  try {
    const todayISO = getNowDateISO();
    const q = query(
      collection(db, "attendance_records"),
      where("uid", "==", uid),
      where("date", "==", todayISO),
      limit(1)
    );
    const snap = await getDocs(q);
    if (!snap.empty) {
      const docData = snap.docs[0].data();
      return docData;
    }
    return null;
  } catch (err) {
    console.error("خطأ أثناء فحص سجل الحضور اليومي:", err);
    return null;
  }
}

/**
 * استرجاع موظف من Firestore بناءً على معرف الـ Passkey (حتى لو مسح بيانات Safari)
 */
async function findEmployeeByCredentialId(credentialId) {
  try {
    const q = query(
      collection(db, "employees"),
      where("credentialId", "==", credentialId),
      limit(1)
    );
    const snap = await getDocs(q);
    if (!snap.empty) {
      return snap.docs[0].data();
    }
    return null;
  } catch (err) {
    console.error("خطأ أثناء البحث عن الموظف بالمفتاح:", err);
    return null;
  }
}

/* ═══════════════ 1. أول استخدام للموظف (تسجيل Face ID لأول مرة) ═══════════════ */

async function handleRegisterFirstTime() {
  hideStatus();

  const fullName = els.inputFullName ? els.inputFullName.value.trim() : "";
  if (!fullName || fullName.length < 2) {
    showStatus("يرجى إدخال اسمك الكامل قبل المتابعة.", "warning");
    els.inputFullName?.focus();
    return;
  }

  if (!checkBrowserSupport()) return;

  const btn = els.btnRegisterFaceId;
  const originalBtnHtml = btn.innerHTML;
  btn.disabled = true;
  btn.innerHTML = `<span class="spinner"></span> جاري تهيئة Face ID...`;
  setScanningState(true);

  try {
    // 1. توليد هوية موظف حقيقية UID عبر Firebase Auth
    const uid = await getOrCreateEmployeeUid();

    // 2. تجهيز خيارات WebAuthn Passkey المتوافقة مع معايير Apple iOS
    const challenge = generateChallenge();
    const userIdBuffer = new TextEncoder().encode(uid);

    const publicKeyOptions = {
      challenge: challenge,
      rp: {
        name: "جمعية إرث وحضارة بالقريات",
        id: window.location.hostname
      },
      user: {
        id: userIdBuffer,
        name: fullName,
        displayName: fullName
      },
      pubKeyCredParams: [
        { alg: -7, type: "public-key" },  // ES256 (Apple Secure Enclave Passkey)
        { alg: -257, type: "public-key" } // RS256
      ],
      authenticatorSelection: {
        authenticatorAttachment: "platform", // بصمة الوجه Face ID أو Touch ID
        residentKey: "required",             // تخزين المفتاح كـ Passkey دائم
        requireResidentKey: true,
        userVerification: "required"         // إجبار فحص بصمة الوجه
      },
      timeout: 60000,
      attestation: "none"
    };

    // 3. استدعاء شاشة Face ID الحقيقية على جهاز iPhone
    const credential = await navigator.credentials.create({
      publicKey: publicKeyOptions
    });

    if (!credential) {
      throw new Error("لم يتم إنشاء مفتاح المصادقة البيومترية.");
    }

    const credentialId = credential.id || bufferToBase64Url(credential.rawId);

    // 4. حفظ بيانات الموظف في مجموعة employees/{uid} في Firestore
    const employeeRef = doc(db, "employees", uid);
    const employeeData = {
      uid: uid,
      name: fullName,
      createdAt: serverTimestamp(),
      credentialId: credentialId,
      credentialCreatedAt: serverTimestamp(),
      lastAttendanceAt: serverTimestamp(),
      active: true,
      registeredDevice: navigator.userAgent
    };

    await setDoc(employeeRef, employeeData);

    // 5. التحقق من منع التكرار لليوم
    const existing = await checkExistingAttendanceToday(uid);
    if (existing) {
      // إذا كان مسجلاً مسبقاً اليوم
      renderDuplicateScreen(existing, fullName);
      return;
    }

    // 6. تسجيل أول حركة حضور في attendance_records/{recordId}
    const recordRef = doc(collection(db, "attendance_records"));
    const recordId = recordRef.id;
    const todayISO = getNowDateISO();
    const formattedDate = getFormattedArabicDate();
    const formattedTime = getFormattedArabicTime();
    const isLate = checkIfLate();

    const attendanceData = {
      recordId: recordId,
      uid: uid,
      employeeName: fullName,
      date: todayISO,
      dateFormatted: formattedDate,
      checkInTime: formattedTime,
      timestamp: serverTimestamp(),
      deviceCredentialId: credentialId,
      type: "check-in",
      isLate: isLate,
      status: isLate ? "متأخر" : "حاضر",
      createdAt: serverTimestamp()
    };

    await setDoc(recordRef, attendanceData);

    // 7. تحديث الذاكرة المؤقتة المحلية (Cache Hint فقط)
    localStorage.setItem(CACHE_KEY_UID, uid);
    localStorage.setItem(CACHE_KEY_NAME, fullName);
    localStorage.setItem(CACHE_KEY_CRED, credentialId);

    // 8. تشغيل الصوت والاهتزاز اللمسي وعرض شاشة النجاح
    playSuccessChime();
    triggerHaptic();
    renderSuccessScreen(attendanceData, fullName, true);

  } catch (err) {
    console.error("خطأ أثناء تسجيل Face ID:", err);
    handleWebAuthnError(err);
  } finally {
    btn.disabled = false;
    btn.innerHTML = originalBtnHtml;
    setScanningState(false);
  }
}

/* ═══════════════ 2. الزيارات القادمة (تسجيل الحضور عبر Face ID) ═══════════════ */

async function handleCheckInReturning(forcedCredId = null) {
  hideStatus();

  if (!checkBrowserSupport()) return;

  const btn = els.btnCheckInFaceId || els.btnDirectVerify;
  const originalBtnHtml = btn ? btn.innerHTML : "";
  if (btn) {
    btn.disabled = true;
    btn.innerHTML = `<span class="spinner"></span> جاري فحص Face ID...`;
  }
  setScanningState(true);

  try {
    const challenge = generateChallenge();
    const cachedCredId = forcedCredId || localStorage.getItem(CACHE_KEY_CRED);

    const getOptions = {
      challenge: challenge,
      rpId: window.location.hostname,
      userVerification: "required", // استدعاء Face ID الحقيقي
      timeout: 60000
    };

    // إذا كان المعرف مخزناً محلياً، نضعه كمرشح، وإلا نترك المتصفح يكتشف الـ Passkey تلقائياً
    if (cachedCredId) {
      getOptions.allowCredentials = [
        {
          id: base64UrlToBuffer(cachedCredId),
          type: "public-key"
        }
      ];
    }

    // استدعاء Face ID
    let assertion;
    try {
      assertion = await navigator.credentials.get({ publicKey: getOptions });
    } catch (e) {
      // إذا فشل بالمعرف المحدد (مثلاً تم حذفه محلياً)، نحاول بدون allowCredentials لاكتشاف الـ Passkey
      if (getOptions.allowCredentials) {
        delete getOptions.allowCredentials;
        assertion = await navigator.credentials.get({ publicKey: getOptions });
      } else {
        throw e;
      }
    }

    if (!assertion) {
      throw new Error("لم يتم التحقق من الهوية البيومترية.");
    }

    const verifiedCredentialId = assertion.id || bufferToBase64Url(assertion.rawId);

    // ضمان وجود جلسة مصادقة نشطة للمستخدم للالتزام بقواعد أمان Firestore
    await getOrCreateEmployeeUid();

    // التعرف على الموظف في Firestore المرتبط بهذا الـ Passkey
    let employee = await findEmployeeByCredentialId(verifiedCredentialId);

    // محاولة إضافية عبر userHandle إن توفر في الاستجابة
    if (!employee && assertion.response && assertion.response.userHandle) {
      try {
        const decodedUid = new TextDecoder().decode(assertion.response.userHandle);
        const empSnap = await getDoc(doc(db, "employees", decodedUid));
        if (empSnap.exists()) {
          employee = empSnap.data();
        }
      } catch (ex) {}
    }

    // محاولة إضافية عبر الـ UID المخزن محلياً
    if (!employee) {
      const cachedUid = localStorage.getItem(CACHE_KEY_UID);
      if (cachedUid) {
        const empSnap = await getDoc(doc(db, "employees", cachedUid));
        if (empSnap.exists()) {
          employee = empSnap.data();
        }
      }
    }

    if (!employee) {
      showStatus(
        "تم التحقق من بصمة الوجه بنجاح، لكن لم يتم العثور على موظف مسجل بهذا المفتاح في النظام. يرجى تسجيل حسابك كموظف جديد أولاً.",
        "warning"
      );
      switchView("register");
      return;
    }

    if (employee.active === false) {
      showStatus("حساب هذا الموظف معطل حالياً. يرجى مراجعة إدارة الجمعية.", "error");
      return;
    }

    // استعادة وتحديث الكاش المحلي
    localStorage.setItem(CACHE_KEY_UID, employee.uid);
    localStorage.setItem(CACHE_KEY_NAME, employee.name);
    localStorage.setItem(CACHE_KEY_CRED, verifiedCredentialId);

    // التحقق من منع التكرار لليوم (شرط 9)
    const existing = await checkExistingAttendanceToday(employee.uid);
    if (existing) {
      renderDuplicateScreen(existing, employee.name);
      return;
    }

    // تسجيل حركة حضور جديدة في attendance_records
    const recordRef = doc(collection(db, "attendance_records"));
    const recordId = recordRef.id;
    const todayISO = getNowDateISO();
    const formattedDate = getFormattedArabicDate();
    const formattedTime = getFormattedArabicTime();
    const isLate = checkIfLate();

    const attendanceData = {
      recordId: recordId,
      uid: employee.uid,
      employeeName: employee.name,
      date: todayISO,
      dateFormatted: formattedDate,
      checkInTime: formattedTime,
      timestamp: serverTimestamp(),
      deviceCredentialId: verifiedCredentialId,
      type: "check-in",
      isLate: isLate,
      status: isLate ? "متأخر" : "حاضر",
      createdAt: serverTimestamp()
    };

    await setDoc(recordRef, attendanceData);

    // تحديث تاريخ آخر حضور للموظف
    await updateDoc(doc(db, "employees", employee.uid), {
      lastAttendanceAt: serverTimestamp()
    });

    // تشغيل التأكيد اللمسي والصوتي
    playSuccessChime();
    triggerHaptic();
    renderSuccessScreen(attendanceData, employee.name, false);

  } catch (err) {
    console.error("خطأ أثناء تسجيل الحضور:", err);
    handleWebAuthnError(err);
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = originalBtnHtml;
    }
    setScanningState(false);
  }
}

/* ═══════════════ معالجة الأخطاء باللغة العربية الواضحة ═══════════════ */

function handleWebAuthnError(err) {
  const name = err.name || "";
  const msg = err.message || "";

  if (name === "NotAllowedError" || msg.includes("canceled") || msg.includes("cancelled")) {
    showStatus(
      "تم إلغاء عملية Face ID أو لم يتم تأكيد البصمة. يمكنك المحاولة مرة أخرى بالضغط على الزر.",
      "warning"
    );
  } else if (name === "InvalidStateError") {
    showStatus(
      "تم تسجيل هذا المفتاح مسبقاً على هذا الجهاز أو يوجد تداخل مع مفتاح سابق.",
      "warning"
    );
  } else if (name === "NotSupportedError") {
    showStatus(
      "جهازك أو متصفحك لا يدعم التحقق البيومتري (Passkeys). يرجى التأكد من تشغيل Face ID واستخدام متصفح Safari.",
      "error"
    );
  } else if (name === "SecurityError") {
    showStatus(
      "خطأ أمني في النطاق. يجب أن تعمل الصفحة تحت اسم النطاق المعتمد للجمعية أو HTTPS.",
      "error"
    );
  } else if (msg.includes("network") || msg.includes("Failed to get document")) {
    showStatus(
      "تعذر الاتصال بقاعدة البيانات. يرجى التحقق من اتصال الإنترنت ثم إعادة المحاولة.",
      "error"
    );
  } else {
    showStatus(
      "حدث تعذر أثناء التحقق من بصمة الوجه. يرجى المحاولة مرة أخرى أو التأكد من إعدادات Face ID في جهازك.",
      "warning"
    );
  }
}

/* ═══════════════ التبديل بين الشاشات والواجهات ═══════════════ */

function switchView(viewName) {
  hideStatus();
  if (els.viewRegister) els.viewRegister.style.display = "none";
  if (els.viewReturning) els.viewReturning.style.display = "none";
  if (els.viewSuccess) els.viewSuccess.style.display = "none";
  if (els.viewDuplicate) els.viewDuplicate.style.display = "none";

  if (viewName === "register" && els.viewRegister) {
    els.viewRegister.style.display = "block";
    els.inputFullName?.focus();
  } else if (viewName === "returning" && els.viewReturning) {
    els.viewReturning.style.display = "block";
  } else if (viewName === "success" && els.viewSuccess) {
    els.viewSuccess.style.display = "flex";
  } else if (viewName === "duplicate" && els.viewDuplicate) {
    els.viewDuplicate.style.display = "flex";
  }
}

function renderSuccessScreen(data, employeeName, isFirstTime = false) {
  const greeting = getTimeOfDayGreeting();
  if (els.successGreeting) {
    els.successGreeting.textContent = `${greeting}، ${employeeName} 👋`;
  }
  if (els.successEmpName) {
    els.successEmpName.textContent = employeeName;
  }
  if (els.successTime) {
    els.successTime.textContent = data.checkInTime;
  }
  if (els.successDate) {
    els.successDate.textContent = data.dateFormatted || getFormattedArabicDate();
  }
  if (els.successRecordId) {
    els.successRecordId.textContent = `#${data.recordId.slice(0, 8).toUpperCase()}`;
  }
  if (els.successUid) {
    els.successUid.textContent = data.uid ? `${data.uid.slice(0, 10)}...` : "—";
  }
  if (els.successStatusBadge) {
    const isLate = data.isLate || data.status === "متأخر";
    els.successStatusBadge.className = isLate ? "receipt-badge warning" : "receipt-badge";
    els.successStatusBadge.textContent = isLate ? "متأخر" : "حاضر في الموعد";
  }

  switchView("success");
}

function renderDuplicateScreen(data, employeeName) {
  const greeting = getTimeOfDayGreeting();
  if (els.dupGreeting) {
    els.dupGreeting.textContent = `أهلاً بك، ${employeeName} 👋`;
  }
  if (els.dupEmpName) {
    els.dupEmpName.textContent = employeeName;
  }
  if (els.dupTime) {
    els.dupTime.textContent = data.checkInTime;
  }
  if (els.dupDate) {
    els.dupDate.textContent = data.dateFormatted || data.date;
  }
  if (els.dupRecordId) {
    els.dupRecordId.textContent = `#${data.recordId.slice(0, 8).toUpperCase()}`;
  }

  playSuccessChime();
  switchView("duplicate");
}

/* ═══════════════ التهيئة عند تحميل الصفحة ═══════════════ */

function initPage() {
  checkBrowserSupport();

  const cachedUid = localStorage.getItem(CACHE_KEY_UID);
  const cachedName = localStorage.getItem(CACHE_KEY_NAME);

  if (cachedUid && cachedName) {
    // الموظف مسجل مسبقاً على هذا الجهاز
    if (els.returningEmployeeName) {
      els.returningEmployeeName.textContent = cachedName;
    }
    switchView("returning");
  } else {
    // أول استخدام أو مسح الـ Cache
    switchView("register");
  }

  // ربط أحداث الأزرار
  if (els.btnRegisterFaceId) {
    els.btnRegisterFaceId.addEventListener("click", handleRegisterFirstTime);
  }

  if (els.btnCheckInFaceId) {
    els.btnCheckInFaceId.addEventListener("click", () => handleCheckInReturning());
  }

  if (els.btnDirectVerify) {
    els.btnDirectVerify.addEventListener("click", () => handleCheckInReturning());
  }

  if (els.btnSwitchAccount) {
    els.btnSwitchAccount.addEventListener("click", () => {
      localStorage.removeItem(CACHE_KEY_UID);
      localStorage.removeItem(CACHE_KEY_NAME);
      localStorage.removeItem(CACHE_KEY_CRED);
      if (els.inputFullName) els.inputFullName.value = "";
      switchView("register");
    });
  }

  if (els.btnReturnFromSuccess) {
    els.btnReturnFromSuccess.addEventListener("click", () => {
      const name = localStorage.getItem(CACHE_KEY_NAME);
      if (name && els.returningEmployeeName) {
        els.returningEmployeeName.textContent = name;
        switchView("returning");
      } else {
        switchView("register");
      }
    });
  }

  if (els.btnReturnFromDuplicate) {
    els.btnReturnFromDuplicate.addEventListener("click", () => {
      const name = localStorage.getItem(CACHE_KEY_NAME);
      if (name && els.returningEmployeeName) {
        els.returningEmployeeName.textContent = name;
        switchView("returning");
      } else {
        switchView("register");
      }
    });
  }

  // دعم الضغط على Enter في حقل الاسم
  if (els.inputFullName) {
    els.inputFullName.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        handleRegisterFirstTime();
      }
    });
  }
}

// تشغيل التهيئة
document.addEventListener("DOMContentLoaded", initPage);
