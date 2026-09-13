# Bot Mail IA

Complément Outlook pour rédiger une réponse à partir du mail ouvert et de ton intention.

Site : https://bot-mail-outlook-addin-chi.vercel.app  
Code : https://github.com/mlecomte1/bot-mail-outlook-addin

## Pour les utilisateurs

1. Télécharge le [manifeste](https://bot-mail-outlook-addin-chi.vercel.app/manifest.xml).
2. Outlook : **Get Add-ins → Mes compléments → Ajouter à partir d’un fichier**.
3. Ouvre un mail → **Bot Mail IA**.

Le Microsoft Store (AppSource) est l’objectif suivant, pour installer sans fichier XML.

## Confidentialité

- **Confidentiel (défaut)** : le serveur n’envoie pas le corps du mail à Groq.
- **Plus précis** : version anonymisée (emails, IBAN, téléphones, montants). Les noms peuvent rester.

Détail : [politique de confidentialité](https://bot-mail-outlook-addin-chi.vercel.app/privacy.html).

Limites anti-abus : 10 générations / 15 min et 40 / jour par IP.

## Local

```powershell
npm test
npm run dev
```

Serveur : `http://127.0.0.1:3000` (accueil) et `/taskpane.html` (panneau).

```
GROQ_API_KEY=
ADDIN_AUDIENCE=https://bot-mail-outlook-addin-chi.vercel.app
```

Dans la [console Groq](https://console.groq.com), mets aussi un plafond de dépense : l’API est publique.

## Stack

Outlook Add-in (Office.js) · Vercel · Groq (`openai/gpt-oss-120b`)

## Publication Store (plus tard)

1. Compte [Partner Center](https://partner.microsoft.com/).
2. Fiche AppSource (description, icônes, URL support + confidentialité déjà en ligne).
3. Soumission du manifeste hébergé en HTTPS.
4. Validation Microsoft (plusieurs jours).
