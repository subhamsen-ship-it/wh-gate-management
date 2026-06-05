const GAS_URL =
  "https://script.google.com/a/macros/apnamart.in/s/AKfycbxrgWtVB8uM0aTkxlz_S10nfPNPdShty5MDWnlA5gfZGvjMekWmi3DntHfwJLiJuaKjKw/exec";

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") return res.status(200).end();

  try {
    if (req.method === "GET") {
      const params = new URLSearchParams(req.query).toString();
      const url = params ? `${GAS_URL}?${params}` : GAS_URL;
      const upstream = await fetch(url, { redirect: "follow" });
      const text = await upstream.text();
      res.setHeader("Content-Type", "application/json");
      return res.status(upstream.status).send(text);
    }

    if (req.method === "POST") {
      const body = await new Promise((resolve) => {
        let data = "";
        req.on("data", (chunk) => (data += chunk));
        req.on("end", () => resolve(data));
      });
      const upstream = await fetch(GAS_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain" },
        body,
        redirect: "follow",
      });
      const text = await upstream.text();
      res.setHeader("Content-Type", "application/json");
      return res.status(upstream.status).send(text);
    }
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
