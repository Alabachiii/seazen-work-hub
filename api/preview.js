// Given an article URL and a keyword, returns a real image:
//   1. the article's own preview image (og:image), if it is a real photo
//   2. otherwise a relevance-ranked photo from Openverse (open image search)
//   3. otherwise null, and the app shows a clean colored tile
// Runs on Vercel's server, so there is no browser cross-origin limit here.

const BAD = /logo|sprite|icon|favicon|placeholder|default-|avatar|blank|1x1|spacer|pixel\.|\.svg(\?|$)/i;

async function fetchText(url, ms) {
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), ms);
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
    clearTimeout(t);
    if (!r.ok) return null;
    return (await r.text()).slice(0, 200000);
  } catch {
    clearTimeout(t);
    return null;
  }
}

async function articleImage(url) {
  const html = await fetchText(url, 6000);
  if (!html) return null;
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
  if (!img) return null;
  img = img.replace(/&amp;/g, "&").trim();
  if (img.startsWith("//")) img = "https:" + img;
  else if (img.startsWith("/")) {
    try { img = new URL(img, url).href; } catch { return null; }
  }
  if (!/^https?:\/\//i.test(img)) return null;
  if (BAD.test(img)) return null; // reject logos / icons / placeholders
  return img;
}

async function stockImage(q) {
  if (!q) return null;
  const query = encodeURIComponent(String(q).replace(/,/g, " ").trim());
  const api = `https://api.openverse.org/v1/images/?q=${query}&page_size=8&mature=false`;
  const txt = await fetchText(api, 5000);
  if (!txt) return null;
  try {
    const j = JSON.parse(txt);
    const list = (j && j.results) || [];
    for (const it of list) {
      const u = it.thumbnail || it.url;
      if (u && /^https?:\/\//i.test(u) && !BAD.test(u)) return u;
    }
  } catch {}
  return null;
}

export default async function handler(req, res) {
  const q = (req.query && req.query.url) || "";
  const kw = (req.query && req.query.q) || "";
  let image = null;
  let source = null;
  if (/^https?:\/\//i.test(q)) {
    image = await articleImage(q);
    if (image) source = "article";
  }
  if (!image) {
    image = await stockImage(kw);
    if (image) source = "stock";
  }
  res.setHeader("cache-control", "public, max-age=86400, s-maxage=86400");
  res.status(200).json({ image, source });
}
