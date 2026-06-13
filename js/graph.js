// ── graph.js ─────────────────────────────────────────────────
// Minimal Microsoft Graph client with retry/paging support.
// No build required — pure ES module.

const BASE = "https://graph.microsoft.com/v1.0";

function sleep(ms) {
  return new Promise(r => setTimeout(r, ms));
}

export function createGraphClient(accessToken) {
  async function request(url, attempt = 0) {
    const abs = url.startsWith("http") ? url : `${BASE}${url}`;

    const res = await fetch(abs, {
      headers: { Authorization: `Bearer ${accessToken}`, Accept: "application/json" }
    });

    // Throttling: respect Retry-After with exponential back-off fallback
    if ((res.status === 429 || res.status === 503) && attempt < 6) {
      const ra = parseInt(res.headers.get("Retry-After") || "0", 10);
      await sleep((ra > 0 ? ra : Math.min(2 ** attempt, 20)) * 1000);
      return request(url, attempt + 1);
    }

    if (!res.ok) {
      const body = await res.text().catch(() => "");
      throw new Error(`Graph ${res.status}: ${body.slice(0, 200)}`);
    }

    return res.json();
  }

  /** Fetch a single page */
  function getJson(url) {
    return request(url);
  }

  /** Automatically follow @odata.nextLink pages */
  async function getAllPages(url, selector) {
    const items = [];
    let next = url;
    while (next) {
      const payload = await request(next);
      const page = selector(payload);
      if (Array.isArray(page)) items.push(...page);
      next = payload["@odata.nextLink"] || null;
    }
    return items;
  }

  return { getJson, getAllPages };
}
