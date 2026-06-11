export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Méthode non autorisée." });
  }

  try {
    const { email, intention, privacyMode } = req.body;

    if (!email || !intention) {
      return res.status(400).json({
        error: "Email et intention obligatoires.",
      });
    }

    if (!process.env.GROQ_API_KEY) {
      return res.status(500).json({
        error: "Clé GROQ_API_KEY manquante.",
      });
    }

    const systemPrompt =
      privacyMode === "confidential"
        ? "Tu es un assistant de rédaction d'emails professionnels. Mode confidentiel activé : tu ne dois jamais inventer ni réintroduire de noms, montants, références, adresses, entreprises, emails, numéros de téléphone ou données sensibles. Tu rédiges une réponse claire, polie et générale uniquement à partir de l'intention utilisateur. Retourne uniquement le corps du mail."
        : "Tu es un assistant de rédaction d'emails professionnels. Tu rédiges des réponses claires, polies et naturelles. Ne rajoute jamais d'informations non données par l'utilisateur. Ne mentionne jamais l'anonymisation. Retourne uniquement le corps du mail.";

    const groqResponse = await fetch(
      "https://api.groq.com/openai/v1/chat/completions",
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${process.env.GROQ_API_KEY}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          model: "llama-3.3-70b-versatile",
          temperature: 0.4,
          messages: [
            {
              role: "system",
              content: systemPrompt,
            },
            {
              role: "user",
              content: `Mail reçu ou contexte :
${email}

Intention de réponse de l'utilisateur :
${intention}

Rédige une réponse email professionnelle, claire et polie.`,
            },
          ],
        }),
      }
    );

    const data = await groqResponse.json();

    if (!groqResponse.ok) {
      console.error("ERREUR GROQ :", data);

      return res.status(500).json({
        error: "Erreur pendant la génération IA.",
        details: data,
      });
    }

    const reply = data.choices?.[0]?.message?.content;

    return res.status(200).json({
      reply: reply || "Aucune réponse générée.",
    });
  } catch (error) {
    console.error("ERREUR API :", error);

    return res.status(500).json({
      error: "Erreur serveur pendant la génération.",
    });
  }
}