type FetchJsonParams = {
  url: string;
  token: string;
  label: string;
  method?: string;
  headers?: Record<string, string>;
  body?: unknown;
};

function truncateErrorText(value: string): string {
  const normalized = value.replace(/\s+/g, " ").trim();
  if (!normalized) {
    return "";
  }
  return normalized.length > 280 ? `${normalized.slice(0, 277)}...` : normalized;
}

export async function fetchJson<T>({
  url,
  token,
  label,
  method = "GET",
  headers = {},
  body,
}: FetchJsonParams): Promise<T> {
  const normalizedHeaders: Record<string, string> = {
    Authorization: `Bearer ${token}`,
    Accept: "application/json",
    ...headers,
  };
  if (body !== undefined && !normalizedHeaders["Content-Type"]) {
    normalizedHeaders["Content-Type"] = "application/json";
  }

  const response = await fetch(url, {
    method,
    headers: normalizedHeaders,
    body: body === undefined ? undefined : JSON.stringify(body),
  });

  const rawText = await response.text();
  if (!response.ok) {
    const detail = truncateErrorText(rawText);
    throw new Error(`${label} failed with ${response.status}${detail ? `: ${detail}` : ""}`);
  }

  if (!rawText.trim()) {
    return undefined as T;
  }

  return JSON.parse(rawText) as T;
}
