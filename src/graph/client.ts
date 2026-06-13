export interface GraphClient {
  getJson<T>(url: string): Promise<T>;
  getAllPages<T>(url: string, itemSelector: (payload: any) => T[]): Promise<T[]>;
}

const BASE_URL = "https://graph.microsoft.com/v1.0";

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function toAbsoluteUrl(url: string): string {
  return url.startsWith("http") ? url : `${BASE_URL}${url}`;
}

export function createGraphClient(accessToken: string): GraphClient {
  async function request<T>(url: string, attempt = 0): Promise<T> {
    const absoluteUrl = toAbsoluteUrl(url);
    const response = await fetch(absoluteUrl, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: "application/json"
      }
    });

    if (response.status === 429 || response.status === 503) {
      const retryAfterHeader = response.headers.get("Retry-After");
      const retryAfterSeconds = retryAfterHeader ? Number.parseInt(retryAfterHeader, 10) : NaN;
      const backoffSeconds = Number.isFinite(retryAfterSeconds) ? retryAfterSeconds : Math.min(2 ** attempt, 20);

      if (attempt < 6) {
        await sleep(backoffSeconds * 1000);
        return request<T>(url, attempt + 1);
      }
    }

    if (!response.ok) {
      const message = await response.text();
      throw new Error(`Graph request failed (${response.status}): ${message}`);
    }

    return (await response.json()) as T;
  }

  async function getJson<T>(url: string): Promise<T> {
    return request<T>(url);
  }

  async function getAllPages<T>(url: string, itemSelector: (payload: any) => T[]): Promise<T[]> {
    const items: T[] = [];
    let nextUrl: string | undefined = url;

    while (nextUrl) {
      const payload = await request<any>(nextUrl);
      items.push(...itemSelector(payload));
      nextUrl = payload["@odata.nextLink"];
    }

    return items;
  }

  return {
    getJson,
    getAllPages
  };
}
