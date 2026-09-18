const RETRYABLE_STATUSES = new Set([408, 409, 429, 500, 502, 503, 504]);

const wait = (milliseconds: number) => new Promise((resolve) => setTimeout(resolve, milliseconds));

function retryDelay(response: Response | undefined, attempt: number) {
  const retryAfter = response?.headers.get("retry-after");
  if (retryAfter) {
    const seconds = Number(retryAfter);
    if (Number.isFinite(seconds)) return Math.min(Math.max(seconds * 1000, 0), 5_000);
    const dateDelay = Date.parse(retryAfter) - Date.now();
    if (Number.isFinite(dateDelay)) return Math.min(Math.max(dateDelay, 0), 5_000);
  }
  return Math.min(350 * (2 ** attempt) + Math.floor(Math.random() * 150), 2_000);
}

export async function fetchOpenAiWithRetry(
  url: string,
  createRequest: () => RequestInit,
  attempts = 3,
) {
  let lastError: unknown;
  for (let attempt = 0; attempt < attempts; attempt++) {
    try {
      const response = await fetch(url, createRequest());
      const retryable = RETRYABLE_STATUSES.has(response.status);
      if (!retryable || attempt === attempts - 1) return response;
      await response.body?.cancel().catch(() => undefined);
      await wait(retryDelay(response, attempt));
    } catch (error) {
      lastError = error;
      if (attempt === attempts - 1) throw error;
      await wait(retryDelay(undefined, attempt));
    }
  }
  throw lastError || new Error("OpenAI request failed after retries.");
}

export function temporaryOpenAiError(status: number) {
  return RETRYABLE_STATUSES.has(status)
    ? "The drawing service is temporarily busy. Please wait a moment and try the PDF again."
    : "";
}
