// Cloudflare Worker: كل دقيقة يسأل Google Drive "فيه ملف اتعدّل؟" ولو أيوه يشغّل التحديث على GitHub فورًا.
// الأسرار (Settings ← Variables and Secrets): GOOGLE_SERVICE_ACCOUNT_JSON, DRIVE_FOLDER_ID, GH_TOKEN
// اختياري: GH_REPO (الافتراضي Acesmaze/makhzan), WINDOW_SEC (الافتراضي 150)
const b64u = (buf) => btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
const enc = (o) => b64u(new TextEncoder().encode(JSON.stringify(o)));

async function driveToken(sa) {
  const now = Math.floor(Date.now() / 1000);
  const head = enc({ alg: "RS256", typ: "JWT" });
  const body = enc({ iss: sa.client_email, scope: "https://www.googleapis.com/auth/drive.readonly", aud: "https://oauth2.googleapis.com/token", iat: now, exp: now + 3300 });
  const pem = sa.private_key.replace(/-----[^-]+-----/g, "").replace(/\s+/g, "");
  const der = Uint8Array.from(atob(pem), (c) => c.charCodeAt(0));
  const key = await crypto.subtle.importKey("pkcs8", der, { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" }, false, ["sign"]);
  const sig = await crypto.subtle.sign("RSASSA-PKCS1-v1_5", key, new TextEncoder().encode(head + "." + body));
  const r = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer", assertion: head + "." + body + "." + b64u(sig) }),
  });
  if (!r.ok) throw new Error("google token " + r.status + " " + (await r.text()).slice(0, 200));
  return (await r.json()).access_token;
}

async function changedRecently(env, token) {
  const since = new Date(Date.now() - (+env.WINDOW_SEC || 150) * 1000).toISOString();
  const H = { Authorization: "Bearer " + token };
  const q = `'${env.DRIVE_FOLDER_ID}' in parents and modifiedTime > '${since}'`;     // فيه ملف اتعدّل أو اتضاف أو اتمسح (الممسوح بيغيّر وقت الفولدر)
  const r = await fetch("https://www.googleapis.com/drive/v3/files?" + new URLSearchParams({ q, fields: "files(id,name,modifiedTime)", pageSize: "10", supportsAllDrives: "true", includeItemsFromAllDrives: "true" }), { headers: H });
  if (!r.ok) throw new Error("drive list " + r.status + " " + (await r.text()).slice(0, 200));
  const files = (await r.json()).files || [];
  return { hit: files.length > 0, files, since };      // بس ملفات اتعدّلت أو اتضافت جوا الفولدر في آخر دقيقتين ونص
}

async function dispatch(env) {
  const r = await fetch(`https://api.github.com/repos/${env.GH_REPO || "Acesmaze/makhzan"}/actions/workflows/sync.yml/dispatches`, {
    method: "POST",
    headers: { Authorization: "Bearer " + env.GH_TOKEN, Accept: "application/vnd.github+json", "X-GitHub-Api-Version": "2022-11-28", "User-Agent": "makhzan-watcher", "Content-Type": "application/json" },
    body: JSON.stringify({ ref: "main", inputs: { force: "0" } }),
  });
  if (r.status !== 204) throw new Error("github dispatch " + r.status + " " + (await r.text()).slice(0, 200));
}

async function check(env) {
  const sa = JSON.parse(env.GOOGLE_SERVICE_ACCOUNT_JSON);
  const res = await changedRecently(env, await driveToken(sa));
  if (res.hit) { await dispatch(env); return "changed -> dispatched: " + res.files.map((x) => x.name).join(", "); }
  return "no change (since " + res.since + ")";
}

export default {
  async scheduled(event, env, ctx) { ctx.waitUntil(check(env).then((m) => console.log(m)).catch((e) => console.error("ERR", e.message))); },
  async fetch(req, env) {   // فتح رابط الـ Worker في المتصفح = فحص يدوي فوري (بيرجّع النتيجة)
    try { return new Response(await check(env), { headers: { "Content-Type": "text/plain; charset=utf-8" } }); }
    catch (e) { return new Response("ERR " + e.message, { status: 500 }); }
  },
};
