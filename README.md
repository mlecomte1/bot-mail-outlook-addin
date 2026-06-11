# Bot Mail IA - Outlook Add-in

Bot Mail IA est un complément Outlook qui aide à rédiger rapidement des réponses professionnelles à partir d’un mail reçu et d’une intention donnée par l’utilisateur.

L’utilisateur ouvre un mail dans Outlook, lance le complément, lit le mail ouvert, écrit ou dicte son intention, puis l’IA génère une réponse claire et polie. La réponse peut ensuite être insérée directement dans Outlook.

## Fonctionnalités

- Lecture du mail ouvert dans Outlook
- Génération d’une réponse email avec IA
- Insertion directe de la réponse dans Outlook
- Mode normal pour une réponse plus précise
- Mode confidentiel pour limiter les données envoyées à l’IA
- Aide à la dictée via Windows + H
- Interface intégrée directement dans Outlook

## Technologies utilisées

- Outlook Add-in
- Office.js
- JavaScript
- HTML / CSS
- Vercel
- Groq API
- Llama 3.3 70B Versatile

## Architecture

```text
Outlook
  ↓
Complément Bot Mail IA
  ↓
Lecture du mail ouvert avec Office.js
  ↓
Anonymisation ou mode confidentiel
  ↓
API Vercel
  ↓
Groq API
  ↓
Réponse générée
  ↓
Insertion dans Outlook