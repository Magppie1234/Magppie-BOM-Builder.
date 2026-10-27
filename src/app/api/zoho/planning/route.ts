import { NextResponse } from "next/server";
import {
  toApiError,
  zohoInventoryGet,
  zohoInventoryMultipartPost,
  zohoInventoryPost,
  zohoInventoryPut,
} from "@/lib/zoho";

export const dynamic = "force-dynamic";
export const maxDuration = 60;

const PLANNING_MODULE = "cm_planning";
const MULTILINE_CAP = 36_000;

type PlanningItem = {
  kind: "carcass" | "shutter" | "hardware";
  group: string;
  name: string;
  qty: number;
  uom?: string;
  colour?: string;
  pack?: string;
  pcs?: number;
  dim?: string;
  boxDim?: string;
};

type PlanningPayload = {
  id: string;
  createdAt: string;
  customerName: string;
  lines: Array<{ code: string; qty: number; elevation?: string; carcassMat?: string; shutterMat?: string }>;
  items: PlanningItem[];
};

type ModuleRecord = Record<string, unknown> & { module_record_id?: string };
type ListResponse = { module_records?: ModuleRecord[]; page_context?: { has_more_page?: boolean } };
type WriteResponse = { module_record?: { module_record_id?: string } };

const compact = (value: Record<string, unknown>) =>
  Object.fromEntries(Object.entries(value).filter(([, item]) => item !== "" && item !== null && item !== undefined));

const clip = (value: unknown, length = 255) => String(value ?? "").slice(0, length);

async function listPlanningRecords(maxPages = 25) {
  const records: ModuleRecord[] = [];
  for (let page = 1; page <= maxPages; page += 1) {
    const response = await zohoInventoryGet<ListResponse>(`/${PLANNING_MODULE}`, { page, per_page: 200 });
    records.push(...(response.module_records ?? []));
    if (!response.page_context?.has_more_page) break;
  }
  return records;
}

function planningDetails(items: PlanningItem[]) {
  const labels = {
    carcass: "CARCASS PACKING LIST",
    shutter: "SHUTTER PACKING LIST",
    hardware: "HARDWARE PACK & OTHER ITEMS",
  } as const;
  const lines: string[] = [];
  (["carcass", "shutter", "hardware"] as const).forEach((kind) => {
    const matching = items.filter((item) => item.kind === kind);
    if (!matching.length) return;
    if (lines.length) lines.push("");
    lines.push(labels[kind]);
    const groups = new Map<string, PlanningItem[]>();
    matching.forEach((item) => groups.set(item.group, [...(groups.get(item.group) ?? []), item]));
    groups.forEach((groupItems, group) => {
      lines.push(group);
      groupItems.forEach((item) => {
        lines.push(
          `  ${(item.pack ?? "").padEnd(10)} | ${item.name} | ${item.colour || "-"} | ${item.pcs ? `${item.pcs} pcs` : "-"} | ${item.qty}${item.uom ? ` ${item.uom}` : ""} | ${item.dim || "-"} | ${item.boxDim ? `box ${item.boxDim}` : "-"}`,
        );
      });
    });
  });
  const text = lines.join("\n");
  if (text.length <= MULTILINE_CAP) return text;
  const note = "\n… truncated — too many packs for one field.";
  return text.slice(0, MULTILINE_CAP - note.length) + note;
}

function validPayload(value: unknown): value is PlanningPayload {
  if (!value || typeof value !== "object") return false;
  const payload = value as Partial<PlanningPayload>;
  return Boolean(payload.id && payload.customerName && Array.isArray(payload.lines) && Array.isArray(payload.items));
}

export async function POST(request: Request) {
  let form: FormData;
  try {
    form = await request.formData();
  } catch {
    return NextResponse.json({ ok: false, error: "Invalid upload request." }, { status: 400 });
  }

  const rawPayload = form.get("payload");
  const upload = form.get("file");
  const knownRecordId = clip(form.get("zohoId") || "");
  let payload: unknown;
  try {
    payload = JSON.parse(typeof rawPayload === "string" ? rawPayload : "");
  } catch {
    return NextResponse.json({ ok: false, error: "The planning data is invalid." }, { status: 400 });
  }

  if (!validPayload(payload)) {
    return NextResponse.json({ ok: false, error: "The planning data is incomplete." }, { status: 400 });
  }
  if (!(upload instanceof File) || upload.size === 0) {
    return NextResponse.json({ ok: false, error: "Select the completed Excel MRP/form sheet." }, { status: 400 });
  }
  if (!/\.(xlsx|xls)$/i.test(upload.name)) {
    return NextResponse.json({ ok: false, error: "Only .xlsx or .xls files can be attached." }, { status: 400 });
  }

  const count = (kind: PlanningItem["kind"]) => payload.items.filter((item) => item.kind === kind).length;
  const recordPayload = compact({
    record_name: clip(`${payload.id} — ${payload.customerName}`, 100),
    cf_plan_id: clip(payload.id),
    cf_customer: clip(payload.customerName),
    cf_planned_on: payload.createdAt.slice(0, 10),
    cf_cabinets: payload.lines.reduce((sum, line) => sum + (Number(line.qty) || 0), 0),
    cf_total_packs: payload.items.length,
    cf_dispatched_packs: 0,
    cf_pending_packs: payload.items.length,
    cf_carcass_packs: count("carcass"),
    cf_shutter_packs: count("shutter"),
    cf_hardware_lines: count("hardware"),
    cf_plan_details: planningDetails(payload.items),
    cf_status: "Planned",
  });

  let recordId = knownRecordId;
  try {
    if (!recordId) {
      const records = await listPlanningRecords();
      recordId = clip(
        records.find((record) => record.cf_plan_id === payload.id)?.module_record_id || "",
      );
    }

    const response = recordId
      ? await zohoInventoryPut<WriteResponse>(`/${PLANNING_MODULE}/${encodeURIComponent(recordId)}`, recordPayload)
      : await zohoInventoryPost<WriteResponse>(`/${PLANNING_MODULE}`, recordPayload);
    recordId = response.module_record?.module_record_id || recordId;
    if (!recordId) throw new Error("Zoho saved the planning record but did not return its record ID.");
  } catch (error) {
    return NextResponse.json({ ok: false, planningSaved: false, ...toApiError(error) }, { status: 502 });
  }

  try {
    const attachment = new FormData();
    attachment.append("attachment", upload, upload.name);
    await zohoInventoryMultipartPost(`/${PLANNING_MODULE}/${encodeURIComponent(recordId)}/attachment`, attachment);
  } catch (error) {
    return NextResponse.json(
      { ok: false, planningSaved: true, recordId, error: `Planning was saved, but the Excel attachment failed: ${toApiError(error).error}` },
      { status: 502 },
    );
  }

  return NextResponse.json({ ok: true, planningSaved: true, attachmentSaved: true, recordId });
}
