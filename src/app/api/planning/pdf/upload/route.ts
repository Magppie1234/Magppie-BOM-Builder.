import { NextResponse } from "next/server";

export const runtime = "nodejs";
export const maxDuration = 60;

const authError = () => NextResponse.json(
  { error: "PDF extraction is not configured: the server's OpenAI API key is invalid or expired." },
  { status: 503 },
);

async function openAiJson(response: Response) {
  const payload = await response.json();
  if (!response.ok) {
    if (response.status === 401) return { errorResponse: authError() };
    return { errorResponse: NextResponse.json({ error: payload?.error?.message || "OpenAI file upload failed." }, { status: response.status }) };
  }
  return { payload };
}

export async function POST(request: Request) {
  try {
    const key = process.env.OPENAI_API_KEY;
    if (!key) return NextResponse.json({ error: "OPENAI_API_KEY is not configured on the server." }, { status: 503 });
    const url = new URL(request.url);
    const action = url.searchParams.get("action");
    const headers = { Authorization: `Bearer ${key}` };

    if (action === "init") {
      const body = await request.json();
      const bytes = Number(body?.bytes);
      const filename = String(body?.filename || "production-drawing.pdf").slice(0, 180);
      if (!Number.isFinite(bytes) || bytes < 1 || bytes > 50 * 1024 * 1024) {
        return NextResponse.json({ error: "PDF must be between 1 byte and 50 MB." }, { status: 400 });
      }
      const response = await fetch("https://api.openai.com/v1/uploads", {
        method: "POST",
        headers: { ...headers, "Content-Type": "application/json" },
        body: JSON.stringify({ bytes, filename, mime_type: "application/pdf", purpose: "user_data" }),
      });
      const result = await openAiJson(response);
      if (result.errorResponse) return result.errorResponse;
      return NextResponse.json({ uploadId: result.payload.id });
    }

    const uploadId = String(url.searchParams.get("uploadId") || "");
    if (!/^upload[-_][A-Za-z0-9_-]+$/.test(uploadId)) return NextResponse.json({ error: "Invalid upload id." }, { status: 400 });

    if (action === "part") {
      const chunk = await request.arrayBuffer();
      if (!chunk.byteLength || chunk.byteLength > 3.25 * 1024 * 1024) {
        return NextResponse.json({ error: "Invalid PDF upload part." }, { status: 400 });
      }
      const form = new FormData();
      form.append("data", new Blob([chunk], { type: "application/octet-stream" }), "part.bin");
      const response = await fetch(`https://api.openai.com/v1/uploads/${encodeURIComponent(uploadId)}/parts`, { method: "POST", headers, body: form });
      const result = await openAiJson(response);
      if (result.errorResponse) return result.errorResponse;
      return NextResponse.json({ partId: result.payload.id });
    }

    if (action === "complete") {
      const body = await request.json();
      const partIds = Array.isArray(body?.partIds) ? body.partIds.map(String) : [];
      if (!partIds.length || partIds.some((id: string) => !/^part[-_][A-Za-z0-9_-]+$/.test(id))) {
        return NextResponse.json({ error: "Invalid upload parts." }, { status: 400 });
      }
      const response = await fetch(`https://api.openai.com/v1/uploads/${encodeURIComponent(uploadId)}/complete`, {
        method: "POST",
        headers: { ...headers, "Content-Type": "application/json" },
        body: JSON.stringify({ part_ids: partIds }),
      });
      const result = await openAiJson(response);
      if (result.errorResponse) return result.errorResponse;
      const fileId = result.payload?.file?.id;
      if (!fileId) return NextResponse.json({ error: "OpenAI completed the upload without returning a file id." }, { status: 502 });
      return NextResponse.json({ fileId });
    }

    return NextResponse.json({ error: "Unknown upload action." }, { status: 400 });
  } catch (error: any) {
    console.error("Planning PDF multipart upload failed", error);
    return NextResponse.json({ error: error?.message || "Could not upload this PDF." }, { status: 500 });
  }
}
