// Fetches a page server-side and returns its own preview image (og:image).
// Runs on Vercel, so there is no browser cross-origin limit here.
export default async function handler(req, res) {
  const url = (req.query && req.query.url) || "";
  if (!/^https?:\/\//i.test(url)) {
    res.status(200).json({ image: null });
    return;
  }
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), 6000);
  try {
    const r = await fetch(url, {
      signal: ctrl.signal,
      redirect: "follow",
      headers: {
        "user-agent":
          "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
        accept: "text/html,application/xhtml+xml",
      },
    });
    clearTimeout(timer);
    if (!r.ok) {
      res.status(200).json({ image: null });
      return;
    }
    // Only read the head region; that is where the meta tags live.
    const html = (await r.text()).slice(0, 200000);
    const grab = (re) => {
      const m = html.match(re);
      return m ? m[1] : null;
    };
    let img =
      grab(/<meta[^>]+property=["']og:image:secure_url["'][^>]+content=["']([^"']+)["']/i) ||
      grab(/<meta[^>]+property=["']og:image["'][^>]+content=["']([^"']+)["']/i) ||
      grab(/<meta[^>]+content=["']([^"']+)["'][^>]+property=["']og:image["']/i) ||
      grab(/<meta[^>]+name=["']twitter:image["'][^>]+content=["']([^"']+)["']/i) ||
      grab(/<meta[^>]+content=["']([^"']+)["'][^>]+name=["']twitter:image["']/i);
    if (img) {
      img = img.replace(/&amp;/g, "&").trim();
      if (img.startsWith("//")) img = "https:" + img;
      else if (img.startsWith("/")) {
        try {
          img = new URL(img, url).href;
        } catch {
          img = null;
        }
      }
    }
    res.setHeader("cache-control", "public, max-age=86400, s-maxage=86400");
    res.status(200).json({ image: img || null });
  } catch (e) {
    clearTimeout(timer);
    res.status(200).json({ image: null });
  }
}
