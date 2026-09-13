let officeReady = false;
let recognition = null;
let listening = false;
let uiBound = false;
let privacyMode = "confidential";
let language = "auto";

function getApiUrl() {
  return new URL("/api/generate-reply", window.location.href).toString();
}

document.addEventListener("DOMContentLoaded", bindUi);

if (document.readyState !== "loading") {
  bindUi();
}

if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
  Office.onReady((info) => {
    bindUi();

    if (info.host === Office.HostType.Outlook) {
      officeReady = true;
      setStatus("Complément prêt.");
      readCurrentMail();
      return;
    }

    officeReady = false;
    setStatus("Ouvre ce panneau depuis Outlook pour lire et insérer un mail.");
  });
} else {
  setStatus("Ouvre ce panneau depuis Outlook pour lire et insérer un mail.");
}

function bindUi() {
  if (uiBound) {
    return;
  }

  const generateBtn = document.getElementById("generateBtn");
  if (!generateBtn) {
    return;
  }

  uiBound = true;
  document.getElementById("readMailBtn").addEventListener("click", readCurrentMail);
  document.getElementById("voiceBtn").addEventListener("click", toggleDictation);
  generateBtn.addEventListener("click", generateReply);
  document.getElementById("insertBtn").addEventListener("click", insertReplyIntoOutlook);
  document.getElementById("clearBtn").addEventListener("click", clearAll);
  bindPills("privacyPills", "data-privacy", (value) => {
    privacyMode = value;
    updatePrivacyHint();
  });
  bindPills("languagePills", "data-lang", (value) => {
    language = value;
  });
  updatePrivacyHint();
}

function bindPills(listId, attr, onChange) {
  const root = document.getElementById(listId);
  if (!root) {
    return;
  }
  root.addEventListener("click", (event) => {
    const button = event.target.closest(`button[${attr}]`);
    if (!button) {
      return;
    }
    root.querySelectorAll("button").forEach((item) => item.classList.remove("active"));
    button.classList.add("active");
    onChange(button.getAttribute(attr));
  });
}

function updatePrivacyHint() {
  const hint = document.getElementById("privacyHint");
  if (!hint) {
    return;
  }
  hint.textContent =
    privacyMode === "normal"
      ? "Version anonymisée du mail envoyée à l’IA."
      : "Le contenu du mail n’est pas envoyé à l’IA.";
}

function setStatus(message, isError) {
  const statusEl = document.getElementById("status");
  if (!statusEl) {
    return;
  }
  statusEl.textContent = message;
  statusEl.classList.toggle("error", Boolean(isError));
  statusEl.classList.toggle("ok", !isError && /prêt|générée|insérée|lu|dictée/i.test(message));
}

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function toSafeHtml(text) {
  return escapeHtml(text).replace(/\n/g, "<br>");
}

function officeAsync(executor) {
  return new Promise((resolve) => {
    try {
      executor((result) => {
        if (result && result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
          return;
        }
        resolve(null);
      });
    } catch (error) {
      console.error(error);
      resolve(null);
    }
  });
}

function getOfficeIdentityToken() {
  return new Promise((resolve) => {
    if (!officeReady || !Office.context?.mailbox?.getUserIdentityTokenAsync) {
      resolve("");
      return;
    }

    Office.context.mailbox.getUserIdentityTokenAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
        resolve(result.value);
        return;
      }

      resolve("");
    });
  });
}

function formatRecipients(list) {
  if (!Array.isArray(list)) {
    return "";
  }

  return list
    .map((entry) => entry.displayName || entry.emailAddress || "")
    .filter(Boolean)
    .join(", ");
}

function formatAddress(entry) {
  if (!entry || typeof entry !== "object") {
    return "";
  }
  return entry.displayName || entry.emailAddress || "";
}

async function getTextField(field) {
  if (!field) {
    return "";
  }

  if (typeof field.getAsync === "function") {
    const value = await officeAsync((callback) => field.getAsync(callback));
    return typeof value === "string" ? value : "";
  }

  return typeof field === "string" ? field : "";
}

async function getAddressField(field) {
  if (!field) {
    return "";
  }

  try {
    if (typeof field.getAsync === "function") {
      const value = await officeAsync((callback) => field.getAsync(callback));
      return formatAddress(value) || formatRecipients(value);
    }

    return formatAddress(field) || formatRecipients(field);
  } catch (error) {
    console.error(error);
    return "";
  }
}

async function getMailBody(item) {
  if (!item?.body?.getAsync) {
    return "";
  }

  if (Office.MailboxEnums?.BodyMode?.Full) {
    const fullBody = await officeAsync((callback) => {
      item.body.getAsync(Office.CoercionType.Text, {
        bodyMode: Office.MailboxEnums.BodyMode.Full,
      }, callback);
    });

    if (fullBody) {
      return fullBody;
    }
  }

  return (
    (await officeAsync((callback) => {
      item.body.getAsync(Office.CoercionType.Text, callback);
    })) || ""
  );
}

async function readCurrentMail() {
  if (!officeReady || !Office.context?.mailbox?.item) {
    setStatus("Outlook n'est pas encore prêt.", true);
    return;
  }

  try {
    const item = Office.context.mailbox.item;
    const lines = [];

    let composeType = "";
    if (item.getComposeTypeAsync) {
      const compose = await officeAsync((callback) => item.getComposeTypeAsync(callback));
      composeType = compose?.composeType || "";
    }

    const from = await getAddressField(item.from);
    const to = await getAddressField(item.to);
    const cc = await getAddressField(item.cc);
    const subject = await getTextField(item.subject);

    if (composeType) {
      lines.push(`Type : ${composeType}`);
    }
    if (from) {
      lines.push(`De : ${from}`);
    }
    if (to) {
      lines.push(`À : ${to}`);
    }
    if (cc) {
      lines.push(`Cc : ${cc}`);
    }
    lines.push(`Objet : ${subject}`);
    lines.push("");
    lines.push(await getMailBody(item));

    document.getElementById("emailText").value = lines.join("\n").trim();
    document.getElementById("intentionText").focus();
    setStatus(composeType ? "Mail lu (mode rédaction)." : "Mail lu.");
  } catch (error) {
    console.error(error);
    setStatus("Impossible de lire le mail ouvert.", true);
  }
}

function getSpeechRecognition() {
  return window.SpeechRecognition || window.webkitSpeechRecognition || null;
}

function toggleDictation() {
  const SpeechRecognition = getSpeechRecognition();
  const voiceBtn = document.getElementById("voiceBtn");
  const intentionEl = document.getElementById("intentionText");

  if (!SpeechRecognition) {
    intentionEl.focus();
    setStatus("Dictée indisponible. Appuie sur Windows + H.");
    return;
  }

  if (listening && recognition) {
    recognition.stop();
    return;
  }

  recognition = new SpeechRecognition();
  recognition.lang = language === "en" ? "en-US" : "fr-FR";
  recognition.continuous = false;
  recognition.interimResults = false;

  recognition.onstart = () => {
    listening = true;
    voiceBtn.classList.add("listening");
    voiceBtn.textContent = "Stop";
    setStatus("Écoute en cours…");
  };

  recognition.onerror = () => {
    listening = false;
    voiceBtn.classList.remove("listening");
    voiceBtn.textContent = "Dicter";
    intentionEl.focus();
    setStatus("Dictée interrompue. Tu peux aussi utiliser Windows + H.");
  };

  recognition.onend = () => {
    listening = false;
    voiceBtn.classList.remove("listening");
    voiceBtn.textContent = "Dicter";
  };

  recognition.onresult = (event) => {
    const transcript = Array.from(event.results)
      .map((result) => result[0]?.transcript || "")
      .join(" ")
      .trim();

    if (!transcript) {
      return;
    }

    const current = intentionEl.value.trim();
    intentionEl.value = current ? `${current} ${transcript}` : transcript;
    setStatus("Intention dictée.");
  };

  try {
    recognition.start();
  } catch (error) {
    console.error(error);
    intentionEl.focus();
    setStatus("Impossible de démarrer la dictée. Appuie sur Windows + H.", true);
  }
}

async function generateReply() {
  const email = document.getElementById("emailText").value.trim();
  const intention = document.getElementById("intentionText").value.trim();
  const generateBtn = document.getElementById("generateBtn");

  if (privacyMode !== "confidential" && !email) {
    setStatus("Lis d’abord le mail ouvert.", true);
    return;
  }

  if (!intention) {
    setStatus("Ajoute ton intention avant de générer.", true);
    document.getElementById("intentionText").focus();
    return;
  }

  setStatus(
    privacyMode === "confidential"
      ? "Génération en cours (mode confidentiel)…"
      : "Génération en cours…"
  );

  generateBtn.disabled = true;
  generateBtn.textContent = "Génération…";

  try {
    let officeToken = "";

    if (window.location.hostname !== "localhost" && window.location.hostname !== "127.0.0.1") {
      officeToken = await getOfficeIdentityToken();
    }

    const payload = {
      intention,
      privacyMode: privacyMode === "normal" ? "normal" : "confidential",
      officeToken,
      tone: "standard",
      language,
      model: "openai/gpt-oss-120b",
      noSignature: true,
      email: privacyMode === "confidential" ? "confidential" : email,
    };

    const response = await fetch(getApiUrl(), {
      method: "POST",
      headers: {
        "Content-Type": "text/plain;charset=UTF-8",
      },
      body: JSON.stringify(payload),
    });

    const raw = await response.text();
    let data = {};

    try {
      data = raw ? JSON.parse(raw) : {};
    } catch (error) {
      console.error(error);
      setStatus("Erreur lors de la génération : réponse invalide du serveur.", true);
      return;
    }

    if (!response.ok) {
      setStatus(data.error || "Erreur lors de la génération du texte.", true);
      return;
    }

    if (!data.reply) {
      setStatus("L’IA n’a renvoyé aucun texte.", true);
      return;
    }

    document.getElementById("replyText").value = data.reply;
    document.getElementById("replyText").focus();
    setStatus("Réponse générée.");
  } catch (error) {
    console.error(error);
    setStatus(error?.message || "Impossible de contacter l'API IA.", true);
  } finally {
    generateBtn.disabled = false;
    generateBtn.textContent = "Générer";
  }
}

function insertReplyIntoOutlook() {
  const reply = document.getElementById("replyText").value.trim();

  if (!reply) {
    setStatus("Aucune réponse à insérer.", true);
    return;
  }

  if (!officeReady || !Office.context?.mailbox?.item) {
    setStatus("Outlook n'est pas prêt.", true);
    return;
  }

  const item = Office.context.mailbox.item;
  const safeHtml = toSafeHtml(reply);

  if (item.body && item.body.setSelectedDataAsync) {
    item.body.setSelectedDataAsync(
      reply,
      {
        coercionType: Office.CoercionType.Text,
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          setStatus("Réponse insérée dans Outlook.");
          return;
        }

        tryOpenReplyForm(safeHtml);
      }
    );

    return;
  }

  tryOpenReplyForm(safeHtml);
}

function tryOpenReplyForm(safeHtml) {
  try {
    Office.context.mailbox.item.displayReplyForm({
      htmlBody: safeHtml,
    });
    setStatus("Réponse ouverte dans Outlook.");
  } catch (error) {
    console.error(error);
    setStatus("Impossible d'insérer la réponse dans Outlook.", true);
  }
}

function clearAll() {
  document.getElementById("emailText").value = "";
  document.getElementById("intentionText").value = "";
  document.getElementById("replyText").value = "";
  setStatus("Champs effacés.");
}
