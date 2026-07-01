// Serverless proxy for the Anthropic API.
// The browser calls /api/claude; this function adds the secret key server-side
// so it is never exposed to the client. Runs on Vercel's Node runtime.
export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).json({ error: { message: "Method not allowed" } });
    return;
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    res.status(500).json({
      error: {
        message:
          "ANTHROPIC_API_KEY is not set. Add it in your Vercel project under Settings > Environment Variables, then redeploy.",
      },
    });
    return;
  }

  try {
    const body =
      typeof req.body === "string" ? req.body : JSON.stringify(req.body || {});

    const upstream = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "content-type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body,
    });

    const text = await upstream.text();
    res.status(upstream.status);
    res.setHeader("content-type", "application/json");
    res.send(text);
  } catch (e) {
    res.status(502).json({
      error: { message: "Proxy error: " + (e && e.message ? e.message : String(e)) },
    });
  }
}
