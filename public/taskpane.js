const API_URL = "/api/generate-reply";

let officeReady = false;

Office.onReady(() => {
  officeReady = true;

  document.getElementById("readMailBtn").addEventListener("click", readCurrentMail);
  document.getElementById("voiceBtn").addEventListener("click", showVoiceHelp);
  document.getElementById("generateBtn").addEventListener("click", generateReply);
  document.getElementById("insertBtn").addEventListener("click", insertReplyIntoOutlook);
  document.getElementById("clearBtn").addEventListener("click", clearAll);

  setStatus("Complément prêt.");
});

function setStatus(message) {
  document.getElementById("status").textContent = message;
}

function showVoiceHelp() {
  const intentionEl = document.getElementById("intentionText");

  intentionEl.focus();

  setStatus("Le champ est prêt. Appuie sur Windows + H pour dicter ton intention.");
}

function anonymizeEmail(text) {
  return text
    .replace(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g, "[EMAIL]")
    .replace(/\b(?:DEV|FAC|REF|DOS|INV|PO)[-_]?\d+\b/gi, "[REFERENCE]")
    .replace(/(\+?\d[\d\s.-]{8,}\d)/g, "[TELEPHONE]")
    .replace(
      /\b\d{1,3}(?:[ .]\d{3})*(?:,\d{2})?\s?(€|euros|EUR|\$|USD)\b/g,
      "[MONTANT]"
    );
}

function buildEmailContext(email, mode) {
  if (mode === "confidential") {
    return `Mode confidentiel activé.

Le contenu complet du mail n'est pas transmis.

Contexte général :
Un interlocuteur a envoyé un email professionnel.
L'utilisateur veut répondre à ce message en gardant un ton clair, poli et adapté.

Consigne :
Rédige une réponse sans mentionner de nom, montant, référence, adresse, entreprise ou détail confidentiel qui ne serait pas explicitement donné dans l'intention de l'utilisateur.`;
  }

  return anonymizeEmail(email);
}

function readCurrentMail() {
  if (!officeReady || !Office.context.mailbox.item) {
    setStatus("Outlook n'est pas encore prêt.");
    return;
  }

  const item = Office.context.mailbox.item;
  const subject = item.subject || "";
  const from = item.from?.displayName || item.from?.emailAddress || "";

  item.body.getAsync(Office.CoercionType.Text, (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      setStatus("Impossible de lire le corps du mail.");
      return;
    }

    const body = result.value || "";

    const content = `De : ${from}
Objet : ${subject}

${body}`;

    document.getElementById("emailText").value = content;
    setStatus("Mail lu depuis Outlook.");
  });
}

async function generateReply() {
  const email = document.getElementById("emailText").value.trim();
  const intention = document.getElementById("intentionText").value.trim();
  const privacyMode = document.getElementById("privacyMode").value;

  if (!email) {
    setStatus("Aucun mail à traiter.");
    return;
  }

  if (!intention) {
    setStatus("Ajoute ton intention avant de générer.");
    return;
  }

  const cleanEmail = buildEmailContext(email, privacyMode);

  if (privacyMode === "confidential") {
    setStatus("Génération IA en cours en mode confidentiel...");
  } else {
    setStatus("Génération IA en cours en mode normal...");
  }

  try {
    const response = await fetch(API_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        email: cleanEmail,
        intention,
        privacyMode
      })
    });

    const data = await response.json();

    if (!response.ok) {
      console.error(data);
      setStatus(data.error || "Erreur pendant la génération.");
      return;
    }

    document.getElementById("replyText").value = data.reply;

    if (privacyMode === "confidential") {
      setStatus("Réponse générée en mode confidentiel.");
    } else {
      setStatus("Réponse générée en mode normal.");
    }
  } catch (error) {
    console.error(error);
    setStatus("Impossible de contacter l'API IA.");
  }
}

function insertReplyIntoOutlook() {
  const reply = document.getElementById("replyText").value.trim();

  if (!reply) {
    setStatus("Aucune réponse à insérer.");
    return;
  }

  if (!officeReady || !Office.context.mailbox.item) {
    setStatus("Outlook n'est pas prêt.");
    return;
  }

  const item = Office.context.mailbox.item;
  const htmlReply = reply.replace(/\n/g, "<br>");

  if (item.body && item.body.setSelectedDataAsync) {
    item.body.setSelectedDataAsync(
      htmlReply,
      {
        coercionType: Office.CoercionType.Html,
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          setStatus("Réponse insérée dans Outlook.");
          return;
        }

        tryOpenReplyForm(htmlReply);
      }
    );

    return;
  }

  tryOpenReplyForm(htmlReply);
}

function tryOpenReplyForm(htmlReply) {
  try {
    Office.context.mailbox.item.displayReplyForm({
      htmlBody: htmlReply,
    });

    setStatus("Réponse ouverte dans Outlook.");
  } catch (error) {
    console.error(error);
    setStatus("Impossible d'insérer la réponse dans Outlook.");
  }
}

function clearAll() {
  document.getElementById("emailText").value = "";
  document.getElementById("intentionText").value = "";
  document.getElementById("replyText").value = "";
  setStatus("Champs effacés.");
}