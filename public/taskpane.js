const DRAFTS_KEY = "botmail.drafts.v1";
const MAX_DRAFTS = 15;

let officeReady = false;
let recognition = null;
let listening = false;
let uiBound = false;

function getApiUrl() {
  return new URL("/api/generate-reply", window.location.href).toString();
}

Office.onReady(() => {
  officeReady = true;
  bindUi();
  setStatus("Complément prêt.");
});

document.addEventListener("DOMContentLoaded", () => {
  bindUi();
});

function bindUi() {
  if (uiBound) {
    return;
  }
  uiBound = true;
  document.getElementById("readMailBtn").addEventListener("click", readCurrentMail);
  document.getElementById("voiceBtn").addEventListener("click", toggleDictation);
  document.getElementById("generateBtn").addEventListener("click", generateReply);
  document.getElementById("insertBtn").addEventListener("click", insertReplyIntoOutlook);
  document.getElementById("clearBtn").addEventListener("click", clearAll);
  renderDrafts();
}

function setStatus(message, isError) {
  const statusEl = document.getElementById("status");
  statusEl.textContent = message;
  statusEl.classList.toggle("error", Boolean(isError));
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
  return new Promise((resolve, reject) => {
    if (!officeReady || !Office.context?.mailbox?.getUserIdentityTokenAsync) {
      reject(new Error("Outlook n'est pas prêt."));
      return;
    }

    Office.context.mailbox.getUserIdentityTokenAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
        resolve(result.value);
        return;
      }

      reject(new Error("Impossible d'authentifier le complément Outlook."));
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
  if (!officeReady || !Office.context.mailbox.item) {
    setStatus("Outlook n'est pas encore prêt.", true);
    return;
  }

  const item = Office.context.mailbox.item;
  const subject = item.subject || "";
  const lines = [];

  let composeType = "";
  if (item.getComposeTypeAsync) {
    const compose = await officeAsync((callback) => item.getComposeTypeAsync(callback));
    composeType = compose?.composeType || "";
  }

  let from = "";
  try {
    from = item.from?.displayName || item.from?.emailAddress || "";
  } catch (error) {
    console.error(error);
  }

  const to = item.to?.getAsync
    ? formatRecipients(await officeAsync((callback) => item.to.getAsync(callback)))
    : "";
  const cc = item.cc?.getAsync
    ? formatRecipients(await officeAsync((callback) => item.cc.getAsync(callback)))
    : "";

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
  setStatus(
    composeType
      ? "Mail lu (mode rédaction, citation incluse si disponible)."
      : "Mail lu depuis Outlook."
  );
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
    setStatus("Dictée du navigateur indisponible. Appuie sur Windows + H.");
    return;
  }

  if (listening && recognition) {
    recognition.stop();
    return;
  }

  recognition = new SpeechRecognition();
  recognition.lang = document.getElementById("languageSelect").value === "en" ? "en-US" : "fr-FR";
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

function loadDrafts() {
  try {
    const parsed = JSON.parse(localStorage.getItem(DRAFTS_KEY) || "[]");
    return Array.isArray(parsed) ? parsed.slice(0, MAX_DRAFTS) : [];
  } catch {
    return [];
  }
}

function saveDraft(entry) {
  const drafts = loadDrafts().filter((item) => item.reply !== entry.reply);
  drafts.unshift(entry);
  localStorage.setItem(DRAFTS_KEY, JSON.stringify(drafts.slice(0, MAX_DRAFTS)));
  renderDrafts();
}

function renderDrafts() {
  const list = document.getElementById("draftList");
  const drafts = loadDrafts();
  list.innerHTML = "";

  drafts.forEach((draft) => {
    const item = document.createElement("li");
    const button = document.createElement("button");
    const when = draft.at ? new Date(draft.at).toLocaleString("fr-FR") : "";
    const preview = (draft.intention || "Brouillon").slice(0, 70);
    button.type = "button";
    button.textContent = `${when} — ${preview}`;
    button.addEventListener("click", () => {
      document.getElementById("intentionText").value = draft.intention || "";
      document.getElementById("replyText").value = draft.reply || "";
      if (draft.tone) {
        document.getElementById("toneSelect").value = draft.tone;
      }
      setStatus("Brouillon restauré.");
    });
    item.appendChild(button);
    list.appendChild(item);
  });
}

async function generateReply() {
  const email = document.getElementById("emailText").value.trim();
  const intention = document.getElementById("intentionText").value.trim();
  const privacyMode = document.getElementById("privacyMode").value;
  const generateBtn = document.getElementById("generateBtn");

  if (!email) {
    setStatus("Aucun mail à traiter.", true);
    return;
  }

  if (!intention) {
    setStatus("Ajoute ton intention avant de générer.", true);
    return;
  }

  setStatus(
    privacyMode === "confidential"
      ? "Génération IA en cours en mode confidentiel..."
      : "Génération IA en cours en mode normal..."
  );

  generateBtn.disabled = true;

  try {
    let officeToken = "";

    if (window.location.hostname !== "localhost" && window.location.hostname !== "127.0.0.1") {
      officeToken = await getOfficeIdentityToken();
    }

    const payload = {
      intention,
      privacyMode: privacyMode === "normal" ? "normal" : "confidential",
      officeToken,
      tone: document.getElementById("toneSelect").value,
      language: document.getElementById("languageSelect").value,
      model: document.getElementById("modelSelect").value,
      noSignature: document.getElementById("noSignature").checked,
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
      const message = data.error || "Erreur lors de la génération du texte.";
      setStatus(message, true);
      document.getElementById("replyText").value = message;
      return;
    }

    if (!data.reply) {
      setStatus("Erreur lors de la génération : aucun texte renvoyé.", true);
      return;
    }

    document.getElementById("replyText").value = data.reply;
    saveDraft({
      at: Date.now(),
      intention,
      reply: data.reply,
      tone: payload.tone,
    });
    setStatus(
      payload.privacyMode === "confidential"
        ? "Réponse générée en mode confidentiel."
        : "Réponse générée en mode normal."
    );
  } catch (error) {
    console.error(error);
    setStatus(error?.message || "Impossible de contacter l'API IA.", true);
  } finally {
    generateBtn.disabled = false;
  }
}

function insertReplyIntoOutlook() {
  const reply = document.getElementById("replyText").value.trim();

  if (!reply) {
    setStatus("Aucune réponse à insérer.", true);
    return;
  }

  if (!officeReady || !Office.context.mailbox.item) {
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
