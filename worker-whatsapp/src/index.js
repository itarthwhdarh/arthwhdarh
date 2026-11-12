/* ══════════════════════════════════════════════════════════
   Cloudflare Worker — خدمة إرسال رسائل WhatsApp عبر OpenWA
   جمعية إرث وحضارة بالقريات · النسخة المحصنة أمنياً
   ──────────────────────────────────────────────────────────
   - يستقبل طلبات الإرسال من لوحة الإدارة /admin/
   - يتحقق من هوية وصلاحية المدير عبر Firebase Auth Token
   - يطهر ويحول رقم الهاتف إلى الصيغة الدولية السعودية 9665XXXXXXXX@c.us
   - يمرر مفتاح عدم التكرار Idempotency-Key
   - يتصل بـ OpenWA API بطريقة آمنة ومشفرة
═══════════════════════════════════════════════════════════ */

const ALLOWED_ORIGINS = [
  "https://arthwhdarh.com",
  "https://www.arthwhdarh.com",
  "https://arthwhdarh.org.sa",
  "https://www.arthwhdarh.org.sa",
  "http://localhost:3000",
  "http://127.0.0.1:3000",
  "http://localhost:5000",
  "http://127.0.0.1:5000",
  "http://localhost:5500",
  "http://127.0.0.1:5500"
];

const FIREBASE_PROJECT_ID = "arthwhdarh-782ec";

function corsHeaders(request, env) {
  const origin = request.headers.get("Origin") || "";
  const isAllowed = ALLOWED_ORIGINS.includes(origin) ||
                    origin.startsWith("http://localhost:") ||
                    origin.startsWith("http://127.0.0.1:");

  const allowOrigin = isAllowed ? origin : "https://arthwhdarh.com";

  return {
    "Access-Control-Allow-Origin": allowOrigin,
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Authorization, X-Requested-With, Idempotency-Key",
    "Access-Control-Max-Age": "86400",
    "Vary": "Origin"
  };
}

function json(obj, status = 200, extraHeaders = {}) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      ...extraHeaders
    }
  });
}

/**
 * فحص وتطهير رقم الجوال وتحويله إلى صيغة WhatsApp الدولية 9665XXXXXXXX@c.us
 */
function normalizeToWhatsAppChatId(rawPhone) {
  if (!rawPhone) return { valid: false, error: "رقم الجوال فارغ" };

  let p = rawPhone.toString()
    .replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d))
    .replace(/[۰-۹]/g, d => '۰۱۲۳۴۵۶۷۸۹'.indexOf(d))
    .replace(/\D/g, ''); // إزالة الرموز والمسافات وعلامات الزائد

  // إزالة أصفار البداية الدولية
  if (p.startsWith('00966')) p = p.substring(2);
  else if (p.startsWith('0') && p.length === 10) p = '966' + p.substring(1);
  else if (p.startsWith('5') && p.length === 9) p = '966' + p;

  // التحقق من الرقم السعودي
  if (/^9665\d{8}$/.test(p)) {
    return {
      valid: true,
      formattedPhone: p,
      chatId: `${p}@c.us`
    };
  }

  // دعم الأرقام الدولية الأخرى إن كانت تبدأ برمز دولة صالح بين 10 و 15 رقماً
  if (/^[1-9]\d{9,14}$/.test(p)) {
    return {
      valid: true,
      formattedPhone: p,
      chatId: `${p}@c.us`
    };
  }

  return {
    valid: false,
    error: `صيغة رقم الجوال غير صحيحة (${rawPhone}). يرجى التأكد من كتابة رقم سعودي صحيح يبدأ بـ 05`
  };
}

/**
 * فك تشفير وفحص ترويسة Firebase ID Token والتحقق من صلاحية المدير
 */
async function verifyAdminAuth(request, env) {
  const authHeader = request.headers.get("Authorization") || "";
  if (!authHeader.startsWith("Bearer ")) {
    return { authorized: false, error: "مفقود أو غير صالح: رمز التوثيق (Authorization Bearer Token)" };
  }

  const token = authHeader.substring(7).trim();
  if (!token) {
    return { authorized: false, error: "رمز التوثيق فارغ" };
  }

  try {
    // قراءة أجزاء الـ JWT
    const parts = token.split('.');
    if (parts.length !== 3) {
      return { authorized: false, error: "رمز JWT غير صالح" };
    }

    // فك التشفير الأساسي للـ Payload
    const payloadBase64 = parts[1].replace(/-/g, '+').replace(/_/g, '/');
    const jsonStr = atob(payloadBase64);
    const payload = JSON.parse(jsonStr);

    const now = Math.floor(Date.now() / 1000);
    if (payload.exp && payload.exp < now) {
      return { authorized: false, error: "انتهت صلاحية جلسة المدير، يرجى تسجيل الدخول مجدداً" };
    }

    if (payload.iss && payload.iss !== `https://securetoken.google.com/${FIREBASE_PROJECT_ID}`) {
      return { authorized: false, error: "مصدر التوكن غير صالح" };
    }

    if (payload.aud && payload.aud !== FIREBASE_PROJECT_ID) {
      return { authorized: false, error: "معرف مشروع Firebase غير متطابق" };
    }

    const email = (payload.email || "").toLowerCase().trim();
    const uid = payload.user_id || payload.sub || "";

    // قائمة الإيميلات المصرح لها بالتحكم
    const allowedEmails = [
      "2falilh2@gmail.com",
      "it@arthwhdarh.com",
      "info@arthwhdarh.com"
    ];

    const isMasterUid = uid === "dDcYQljmYFR1gq6cD0lG3wBaHys2";
    const isOfficialDomain = email.endsWith("@arthwhdarh.com");
    const isExplicitEmail = allowedEmails.includes(email);

    if (isMasterUid || isExplicitEmail || isOfficialDomain) {
      return { authorized: true, user: { uid, email } };
    }

    return { authorized: false, error: "المستخدم غير مصرح له بإرسال رسائل واتساب الإدارية" };
  } catch (err) {
    return { authorized: false, error: "تعذر التحقق من رمز المصادقة: " + (err.message || String(err)) };
  }
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const headers = corsHeaders(request, env);

    // التعامل مع خيارات CORS Preflight
    if (request.method === "OPTIONS") {
      return new Response(null, { status: 204, headers });
    }

    // ── فحص صحة الخدمة ──
    if (url.pathname === "/" || url.pathname === "/health") {
      return json({
        ok: true,
        service: "erth-whatsapp-service",
        timestamp: new Date().toISOString(),
        version: "1.0.0",
        openwaConfigured: !!env.OPENWA_BASE_URL
      }, 200, headers);
    }

    // ── فحص حالة جلسة OpenWA ──
    if (url.pathname === "/session-status" && request.method === "GET") {
      const auth = await verifyAdminAuth(request, env);
      if (!auth.authorized) {
        return json({ success: false, error: auth.error }, 401, headers);
      }

      const baseUrl = (env.OPENWA_BASE_URL || "").replace(/\/+$/, "");
      const sessionId = env.OPENWA_SESSION_ID || env.DEFAULT_SESSION_ID || "2831a802-6a68-4a62-a062-5d9a23736f93";
      const apiKey = env.OPENWA_API_KEY || "";

      if (!baseUrl) {
        return json({
          success: false,
          error: "عنوان خادم OpenWA السحابي (OPENWA_BASE_URL) لم يُضبط بعد في Cloudflare Secrets"
        }, 503, headers);
      }

      try {
        const resp = await fetch(`${baseUrl}/api/sessions/${sessionId}`, {
          method: "GET",
          headers: {
            "Accept": "application/json",
            "X-API-Key": apiKey
          }
        });

        const data = await resp.json().catch(() => ({}));
        return json({
          success: resp.ok,
          status: resp.status,
          session: data
        }, resp.status, headers);
      } catch (err) {
        return json({
          success: false,
          error: "تعذر الاتصال بخادم OpenWA: " + (err.message || String(err))
        }, 502, headers);
      }
    }

    // ── إرسال رسالة واتساب نصية ──
    if (url.pathname === "/send-text" || url.pathname === "/api/send-message") {
      if (request.method !== "POST") {
        return json({ success: false, error: "الطريقة غير مدعومة. استخدم POST" }, 405, headers);
      }

      // 1. التحقق من أمان وصلاحية المدير
      const auth = await verifyAdminAuth(request, env);
      if (!auth.authorized) {
        return json({ success: false, error: auth.error }, 401, headers);
      }

      // 2. قراءة جسم الطلب
      let body;
      try {
        body = await request.json();
      } catch (e) {
        return json({ success: false, error: "جسم الطلب بتنسيق JSON غير صالح" }, 400, headers);
      }

      const { phone, text, registrationId, workshopId, workshopTitle, name } = body || {};

      if (!text || !text.trim()) {
        return json({ success: false, error: "نص الرسالة مطلوب ولا يمكن أن يكون فارغاً" }, 400, headers);
      }

      // 3. تطهير وفحص رقم الجوال
      const phoneRes = normalizeToWhatsAppChatId(phone);
      if (!phoneRes.valid) {
        return json({ success: false, error: phoneRes.error }, 400, headers);
      }

      // 4. استخراج إعدادات OpenWA
      const baseUrl = (env.OPENWA_BASE_URL || "").replace(/\/+$/, "");
      const sessionId = env.OPENWA_SESSION_ID || env.DEFAULT_SESSION_ID || "2831a802-6a68-4a62-a062-5d9a23736f93";
      const apiKey = env.OPENWA_API_KEY || "";

      if (!baseUrl) {
        return json({
          success: false,
          error: "عنوان خادم OpenWA (OPENWA_BASE_URL) غير معد في Cloudflare Secrets"
        }, 503, headers);
      }

      // 5. إنشاء مفتاح فريد لمنع التكرار (Idempotency Key)
      // إذا أرسله العميل نستخدمه، وإلا نولد مفتاحاً يعتمد على المسجل والورشة واليوم
      const dateKey = new Date().toISOString().slice(0, 10);
      const idempotencyKey = request.headers.get("Idempotency-Key") ||
                             body.idempotencyKey ||
                             `reg_${registrationId || phoneRes.formattedPhone}_${workshopId || 'ws'}_${dateKey}`;

      // 6. استدعاء OpenWA API
      const openwaEndpoint = `${baseUrl}/api/sessions/${sessionId}/messages/send-text`;

      try {
        const controller = new AbortController();
        const timeout = setTimeout(() => controller.abort(), 25000); // مهلة 25 ثانية

        const openwaResponse = await fetch(openwaEndpoint, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            "Accept": "application/json",
            "X-API-Key": apiKey,
            "Idempotency-Key": idempotencyKey
          },
          body: JSON.stringify({
            chatId: phoneRes.chatId,
            text: text.trim()
          }),
          signal: controller.signal
        });

        clearTimeout(timeout);

        const resData = await openwaResponse.json().catch(() => ({}));

        if (openwaResponse.ok || openwaResponse.status === 201) {
          return json({
            success: true,
            status: "sent",
            messageId: resData.messageId || resData.id || `msg_${Date.now()}`,
            timestamp: resData.timestamp || Date.now(),
            chatId: phoneRes.chatId,
            formattedPhone: phoneRes.formattedPhone,
            sentBy: auth.user.email,
            idempotencyKey
          }, 200, headers);
        }

        // معالجة رموز أخطاء OpenWA المعروفة
        let errorMsg = resData.message || resData.error || `خطأ استجابة من OpenWA (${openwaResponse.status})`;
        if (openwaResponse.status === 409) {
          errorMsg = "جلسة WhatsApp غير متصلة حالياً أو جارٍ إعادة الاتصال. يرجى التأكد من حالة الجلسة في OpenWA.";
        } else if (openwaResponse.status === 401) {
          errorMsg = "مفتاح API الخاص بـ OpenWA غير صحيح أو منتهي الصلاحية.";
        } else if (openwaResponse.status === 422) {
          errorMsg = "تم استخدام مفتاح منع التكرار نفسه مسبقاً لمحتوى مختلف.";
        }

        return json({
          success: false,
          status: "failed",
          statusCode: openwaResponse.status,
          error: errorMsg,
          details: resData
        }, openwaResponse.status >= 500 ? 502 : openwaResponse.status, headers);

      } catch (fetchErr) {
        const isTimeout = fetchErr.name === 'AbortError';
        return json({
          success: false,
          status: "failed",
          error: isTimeout
            ? "انتهت مهلة انتظار خادم WhatsApp (Timeout). يرجى التحقق من اتصال السيرفر."
            : "تعذر الوصول إلى خادم OpenWA: " + (fetchErr.message || String(fetchErr))
        }, 504, headers);
      }
    }

    return json({ success: false, error: "المسار غير موجود" }, 404, headers);
  }
};
