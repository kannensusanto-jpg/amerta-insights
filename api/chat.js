const Anthropic = require("@anthropic-ai/sdk");

const client = new Anthropic.default({ apiKey: process.env.ANTHROPIC_API_KEY });

module.exports = async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).end();

  const { messages, context } = req.body;

  try {
    const response = await client.messages.create({
      model: "claude-sonnet-4-6",
      max_tokens: 1024,
      system: `You are a sales analytics assistant for Amerta Niagatama, an Indonesian trading company.
Answer questions in the same language the user writes in (Bahasa Indonesia or English).
Currency is Indonesian Rupiah (Rp). Use Jt for juta (million) and M for miliar (billion) when formatting numbers.
Be concise and direct. Lead with the specific numbers, then the insight.

Here is the company's sales data:
${context}`,
      messages,
    });

    res.status(200).json({ content: response.content[0].text });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
};
