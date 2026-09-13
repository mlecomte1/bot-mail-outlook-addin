const { test } = require("node:test");
const assert = require("node:assert/strict");
const handler = require("../api/generate-reply");
const { anonymizeEmail, sanitizeReply, escapeHtml, clip } = require("../api/security");

function mockRes() {
  return {
    headers: {},
    statusCode: 200,
    body: null,
    setHeader(name, value) {
      this.headers[name] = value;
      return this;
    },
    getHeader(name) {
      return this.headers[name];
    },
    status(code) {
      this.statusCode = code;
      return this;
    },
    json(body) {
      this.body = body;
      return this;
    },
    send(body) {
      this.body = body;
      return this;
    },
  };
}

test("anonymise email, iban et montant", () => {
  const text = anonymizeEmail("Contact jean@client.fr IBAN FR7612345678901234567890121 1200 €");
  assert.match(text, /\[EMAIL\]/);
  assert.match(text, /\[IBAN\]/);
  assert.match(text, /\[MONTANT\]/);
});

test("sanitizeReply conserve les accents", () => {
  assert.equal(sanitizeReply("C'est noté, merci."), "C'est noté, merci.");
});

test("sanitizeReply enlève le HTML", () => {
  assert.equal(sanitizeReply("<b>Bonjour</b>"), "Bonjour");
});

test("escapeHtml encode les chevrons", () => {
  assert.equal(escapeHtml("<hi>"), "&lt;hi&gt;");
});

test("clip tronque et retire les nuls", () => {
  assert.equal(clip("abcd\0", 3), "abc");
});

test("GET est refusé sans fuite de clé", async () => {
  const res = mockRes();
  await handler({ method: "GET", headers: {} }, res);
  assert.equal(res.statusCode, 405);
  assert.equal(res.body?.error, "Méthode non autorisée.");
  assert.equal(res.body?.hasGroqKey, undefined);
});

test("CORS refuse une origine inconnue", async () => {
  const res = mockRes();
  await handler({ method: "OPTIONS", headers: { origin: "https://evil.example" } }, res);
  assert.equal(res.headers["Access-Control-Allow-Origin"], undefined);
});

test("CORS autorise l'origine du complément", async () => {
  const res = mockRes();
  await handler(
    {
      method: "OPTIONS",
      headers: { origin: "https://bot-mail-outlook-addin-chi.vercel.app" },
    },
    res
  );
  assert.equal(
    res.headers["Access-Control-Allow-Origin"],
    "https://bot-mail-outlook-addin-chi.vercel.app"
  );
});

test("production sans jeton Exchange continue (jetons Outlook désactivés par Microsoft)", async () => {
  process.env.VERCEL = "1";
  delete process.env.REQUIRE_OFFICE_TOKEN;
  delete process.env.GROQ_API_KEY;
  const res = mockRes();
  await handler(
    {
      method: "POST",
      headers: { origin: "https://bot-mail-outlook-addin-chi.vercel.app" },
      body: {
        email: "Bonjour",
        intention: "Réponds poliment",
        privacyMode: "normal",
      },
    },
    res
  );
  delete process.env.VERCEL;
  assert.equal(res.statusCode, 500);
  assert.equal(res.body?.error, "Service IA indisponible.");
});

test("REQUIRE_OFFICE_TOKEN sans jeton renvoie 401", async () => {
  process.env.VERCEL = "1";
  process.env.REQUIRE_OFFICE_TOKEN = "1";
  const res = mockRes();
  await handler(
    {
      method: "POST",
      headers: { origin: "https://bot-mail-outlook-addin-chi.vercel.app" },
      body: {
        email: "Bonjour",
        intention: "Réponds poliment",
        privacyMode: "normal",
      },
    },
    res
  );
  delete process.env.VERCEL;
  delete process.env.REQUIRE_OFFICE_TOKEN;
  assert.equal(res.statusCode, 401);
});
