const crypto = require("crypto");

const ALLOWED_ORIGINS = [
  "https://bot-mail-outlook-addin-chi.vercel.app",
  "http://localhost:3000",
  "http://127.0.0.1:3000",
];

const MAX_EMAIL_CHARS = 8000;
const MAX_INTENTION_CHARS = 2000;
const MAX_TOKEN_CHARS = 8192;
const RATE_WINDOW_MS = 15 * 60 * 1000;
const RATE_MAX_REQUESTS = 20;
const METADATA_TTL_MS = 12 * 60 * 60 * 1000;

function productionAudience() {
  if (process.env.ADDIN_AUDIENCE) {
    return process.env.ADDIN_AUDIENCE.replace(/\/$/, "");
  }

  const host = process.env.VERCEL_PROJECT_PRODUCTION_URL;
  if (host) {
    return host.startsWith("http") ? host.replace(/\/$/, "") : `https://${host}`;
  }

  return "https://bot-mail-outlook-addin-chi.vercel.app";
}

const DEFAULT_AUDIENCE = productionAudience();
const MANIFEST_ID = "6c6d9c3d-6c86-4f7d-9b5a-111111111111";
const DEFAULT_METADATA_URL =
  "https://outlook.office.com/autodiscover/metadata/json/1";

const ALLOWED_METADATA_HOSTS = new Set([
  "outlook.office.com",
  "outlook.office365.com",
  "outlook.live.com",
  "webshell.suite.office.com",
]);

function getStore() {
  const g = globalThis;
  if (!g.__botMailSecurity) {
    g.__botMailSecurity = {
      rate: new Map(),
      metadata: new Map(),
    };
  }
  return g.__botMailSecurity;
}

function allowedOrigin(origin) {
  return typeof origin === "string" && ALLOWED_ORIGINS.includes(origin);
}

function applyCors(req, res) {
  if (typeof res.setHeader !== "function") {
    return;
  }

  const origin = req.headers?.origin;
  if (allowedOrigin(origin)) {
    res.setHeader("Access-Control-Allow-Origin", origin);
    res.setHeader("Vary", "Origin");
  }

  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  res.setHeader("Access-Control-Max-Age", "86400");
  res.setHeader("Cache-Control", "no-store");
  res.setHeader("X-Content-Type-Options", "nosniff");
}

function readJsonBody(req) {
  const body = req.body;

  if (body == null || body === "") {
    return {};
  }

  if (typeof body === "string") {
    return JSON.parse(body);
  }

  if (typeof Buffer !== "undefined" && Buffer.isBuffer(body)) {
    return JSON.parse(body.toString("utf8"));
  }

  if (typeof body === "object") {
    return body;
  }

  return {};
}

function clientIp(req) {
  const forwarded = req.headers?.["x-forwarded-for"];

  if (typeof forwarded === "string" && forwarded.trim()) {
    return forwarded.split(",")[0].trim().slice(0, 64);
  }

  if (typeof req.headers?.["x-real-ip"] === "string") {
    return req.headers["x-real-ip"].slice(0, 64);
  }

  return "unknown";
}

function rateLimit(ip) {
  const now = Date.now();
  const store = getStore().rate;

  for (const [key, entry] of store) {
    if (now - entry.start > RATE_WINDOW_MS) {
      store.delete(key);
    }
  }

  const current = store.get(ip);

  if (!current || now - current.start > RATE_WINDOW_MS) {
    store.set(ip, { start: now, count: 1 });
    return { ok: true, remaining: RATE_MAX_REQUESTS - 1 };
  }

  current.count += 1;

  if (current.count > RATE_MAX_REQUESTS) {
    return { ok: false, remaining: 0 };
  }

  return { ok: true, remaining: RATE_MAX_REQUESTS - current.count };
}

function clip(value, max) {
  const text = String(value || "").replace(/\0/g, "").trim();
  return text.length > max ? text.slice(0, max) : text;
}

function anonymizeEmail(text) {
  return String(text || "")
    .replace(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g, "[EMAIL]")
    .replace(/\b[A-Z]{2}\d{2}[A-Z0-9]{10,30}\b/g, "[IBAN]")
    .replace(/\b(?:\d[ -]*?){13,19}\b/g, "[CARTE]")
    .replace(/\b\d{14}\b/g, "[SIRET]")
    .replace(/\b(?:DEV|FAC|REF|DOS|INV|PO|CMD)[-_]?\d+\b/gi, "[REFERENCE]")
    .replace(/(\+?\d[\d\s.-]{8,}\d)/g, "[TELEPHONE]")
    .replace(
      /\b\d+(?:[ .]\d{3})*(?:[.,]\d{2})?\s?(?:€|euros|EUR|\$|USD)/gi,
      "[MONTANT]"
    )
    .replace(/\b(?:\d{1,3}\.){3}\d{1,3}\b/g, "[IP]");
}

function confidentialContext() {
  return `Mode confidentiel activé.

Le contenu complet du mail n'est pas transmis.

Contexte général :
Un interlocuteur a envoyé un email professionnel.
L'utilisateur veut répondre à ce message en gardant un ton clair, poli et adapté.

Consigne :
Rédige une réponse sans mentionner de nom, montant, référence, adresse, entreprise ou détail confidentiel qui ne serait pas explicitement donné dans l'intention de l'utilisateur.`;
}

function sanitizeReply(text) {
  return String(text || "")
    .replace(/```[\s\S]*?```/g, "")
    .replace(/<script[\s\S]*?>[\s\S]*?<\/script>/gi, "")
    .replace(/<style[\s\S]*?>[\s\S]*?<\/style>/gi, "")
    .replace(/<[^>]+>/g, "")
    .replace(/javascript:/gi, "")
    .replace(/vbscript:/gi, "")
    .replace(/data:/gi, "")
    .replace(/\0/g, "")
    .trim()
    .slice(0, 8000);
}

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function b64urlToBuffer(value) {
  const padded = value.replace(/-/g, "+").replace(/_/g, "/");
  const pad = padded.length % 4 === 0 ? "" : "=".repeat(4 - (padded.length % 4));
  return Buffer.from(padded + pad, "base64");
}

function toB64Url(value) {
  return String(value)
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/g, "");
}

function parseJsonBuffer(buffer) {
  return JSON.parse(buffer.toString("utf8"));
}

function isAllowedMetadataUrl(rawUrl) {
  let parsed;

  try {
    parsed = new URL(rawUrl);
  } catch {
    return false;
  }

  if (parsed.protocol !== "https:") {
    return false;
  }

  return ALLOWED_METADATA_HOSTS.has(parsed.hostname.toLowerCase());
}

function audienceMatches(payload, expectedAudience) {
  const values = Array.isArray(payload.aud) ? payload.aud : [payload.aud];
  const expectedHost = new URL(expectedAudience).host;

  return values.some((value) => {
    if (typeof value !== "string" || !value) {
      return false;
    }

    if (value === expectedAudience || value.startsWith(expectedAudience) || value === MANIFEST_ID) {
      return true;
    }

    try {
      return new URL(value).host === expectedHost;
    } catch {
      return value.includes(expectedHost);
    }
  });
}

function issuerAllowed(issuer) {
  if (typeof issuer !== "string") {
    return false;
  }

  const lower = issuer.toLowerCase();
  return (
    lower.includes("outlook.office.com") ||
    lower.includes("outlook.office365.com") ||
    lower.includes("outlook.live.com") ||
    lower.includes("microsoft.com")
  );
}

async function fetchMetadata(url) {
  const store = getStore().metadata;
  const cached = store.get(url);

  if (cached && Date.now() - cached.at < METADATA_TTL_MS) {
    return cached.data;
  }

  const response = await fetch(url, {
    method: "GET",
    signal: AbortSignal.timeout(5000),
    headers: { Accept: "application/json" },
  });

  if (!response.ok) {
    throw new Error("metadata");
  }

  const data = await response.json();
  store.set(url, { at: Date.now(), data });
  return data;
}

function findRsaKey(metadata, header) {
  const keys = Array.isArray(metadata?.keys) ? metadata.keys : [];
  const thumb = header.x5t || header.kid;

  const match = keys.find((key) => {
    const info = key.keyinfo || {};
    return (
      (thumb && (info.x5t === thumb || key.x5t === thumb || key.kid === thumb)) ||
      false
    );
  });

  return match || keys.find((key) => key.keyvalue?.type === "RSA") || keys[0];
}

function verifyRs256(parts, rsaKey) {
  const keyvalue = rsaKey.keyvalue || rsaKey;
  const x5c = keyvalue.x5c || rsaKey.x5c;
  let key;

  if (x5c) {
    const cert = Array.isArray(x5c) ? x5c[0] : x5c;
    const pem = `-----BEGIN CERTIFICATE-----\n${cert}\n-----END CERTIFICATE-----`;
    key = crypto.createPublicKey(pem);
  } else {
    const modulus = keyvalue.modulus || keyvalue.n;
    const exponent = keyvalue.exponent || keyvalue.e;

    if (!modulus || !exponent) {
      return false;
    }

    key = crypto.createPublicKey({
      key: {
        kty: "RSA",
        n: toB64Url(modulus),
        e: toB64Url(exponent),
      },
      format: "jwk",
    });
  }

  return crypto.verify(
    "RSA-SHA256",
    Buffer.from(`${parts[0]}.${parts[1]}`),
    key,
    b64urlToBuffer(parts[2])
  );
}

async function verifyOfficeIdentityToken(token) {
  if (typeof token !== "string" || token.length < 40 || token.length > MAX_TOKEN_CHARS) {
    throw new Error("token");
  }

  const parts = token.split(".");
  if (parts.length !== 3) {
    throw new Error("token");
  }

  const header = parseJsonBuffer(b64urlToBuffer(parts[0]));
  const payload = parseJsonBuffer(b64urlToBuffer(parts[1]));
  const now = Date.now();

  if (header.alg && header.alg !== "RS256") {
    throw new Error("alg");
  }

  if (!payload.exp || payload.exp * 1000 < now - 120000) {
    throw new Error("exp");
  }

  if (payload.nbf && payload.nbf * 1000 > now + 120000) {
    throw new Error("nbf");
  }

  const expectedAudience = productionAudience();
  if (!audienceMatches(payload, expectedAudience)) {
    throw new Error("aud");
  }

  if (!issuerAllowed(payload.iss) && !issuerAllowed(payload.appctxsender)) {
    throw new Error("iss");
  }

  const appctx =
    typeof payload.appctx === "string"
      ? JSON.parse(payload.appctx)
      : payload.appctx || {};
  const metadataUrl = appctx.amurl || DEFAULT_METADATA_URL;

  if (!isAllowedMetadataUrl(metadataUrl)) {
    throw new Error("amurl");
  }

  const metadata = await fetchMetadata(metadataUrl);
  const rsaKey = findRsaKey(metadata, header);

  if (!rsaKey || !verifyRs256(parts, rsaKey)) {
    throw new Error("sig");
  }

  return payload;
}

function officeUserId(payload) {
  try {
    const appctx =
      typeof payload?.appctx === "string"
        ? JSON.parse(payload.appctx)
        : payload?.appctx || {};
    return String(appctx.msexchuid || payload.oid || payload.sub || "").slice(0, 128);
  } catch {
    return "";
  }
}

function logEvent(event, details) {
  const safe = {};

  if (details && typeof details === "object") {
    for (const [key, value] of Object.entries(details)) {
      if (["email", "intention", "officeToken", "reply", "body"].includes(key)) {
        continue;
      }
      safe[key] = value;
    }
  }

  console.error(JSON.stringify({ event, ...safe }));
}

function requiresOfficeToken() {
  return process.env.REQUIRE_OFFICE_TOKEN === "1";
}

module.exports = {
  MAX_EMAIL_CHARS,
  MAX_INTENTION_CHARS,
  applyCors,
  readJsonBody,
  clientIp,
  rateLimit,
  clip,
  anonymizeEmail,
  confidentialContext,
  sanitizeReply,
  escapeHtml,
  verifyOfficeIdentityToken,
  officeUserId,
  logEvent,
  requiresOfficeToken,
  allowedOrigin,
};
