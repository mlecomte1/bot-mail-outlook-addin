# Bot Mail IA

Complément Outlook pour rédiger une réponse professionnelle à partir du mail ouvert et de ton intention.

## Utilisation

1. Ouvre un mail dans Outlook (lecture ou réponse).
2. Lance **Bot Mail IA**.
3. Clique **Lire le mail ouvert**.
4. Écris ou dicte ton intention.
5. Vérifie le mode de confidentialité (confidentiel par défaut).
6. **Générer**, relis, puis **Insérer dans Outlook**.

En mode rédaction, le complément demande le corps complet (citation comprise) quand Outlook le permet, plus les destinataires.

## Confidentialité

- **Confidentiel (défaut)** : le serveur n’envoie pas le corps du mail à Groq, seulement l’intention.
- **Normal** : version anonymisée (emails, IBAN, téléphones, montants, etc.). Les noms peuvent rester.

La génération en production exige un jeton d’identité Exchange. Ça ne marche pas depuis un navigateur anonyme.

Pour limiter la conservation chez Groq : [Data Controls](https://console.groq.com/settings/data-controls) (zero retention / politique de ton compte). Ce n’est pas activable uniquement depuis le code.

## Stack

- Outlook Add-in (Office.js, Mailbox 1.3+)
- Panneau `public/taskpane.*`
- API Vercel `api/generate-reply.js`
- Groq : `openai/gpt-oss-120b` (défaut), `openai/gpt-oss-20b`, `qwen/qwen3.8-27b`

## Local

```powershell
npm test
npm run dev
```

Serveur : `http://127.0.0.1:3000` (localhost uniquement). Pas de jeton Outlook en local.

Variables (`.env`, jamais commitées) :

```
GROQ_API_KEY=
ADDIN_AUDIENCE=https://bot-mail-outlook-addin-chi.vercel.app
```

## Déploiement

Le projet se déploie sur Vercel (`bot-mail-outlook-addin-chi.vercel.app`).

Pour déployer depuis Git :

1. Pousse ce dépôt vers GitHub / GitLab.
2. Dans Vercel : **Import Git Repository**, projet existant ou nouveau.
3. Variables d’environnement Production : `GROQ_API_KEY`, `ADDIN_AUDIENCE`.
4. Chaque push sur la branche de production redéploie.

Sinon : `npx vercel --prod` depuis ce dossier.

Après déploiement, réinstalle le `manifest.xml` (version actuelle **1.0.3.0**) si Outlook garde l’ancien cache.

## Hors périmètre

- Publication Microsoft Store
- SSO Entra ID (app Azure + `WebApplicationInfo` dans le manifeste)
- Gestion d’équipe / comptes multiples
