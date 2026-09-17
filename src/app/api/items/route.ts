import { NextRequest, NextResponse } from "next/server";
import { searchCatalogItems } from "@/lib/itemCatalog";
import { toApiError, zohoInventoryPost } from "@/lib/zoho";
import type { ItemDetail } from "@/lib/types";

export const dynamic = "force-dynamic";

export async function GET(request: NextRequest) {
  const { searchParams } = new URL(request.url);
  const search = searchParams.get("search") || "";
  const name = searchParams.get("name") || "";
  const limit = Math.min(1000, Math.max(1, Number(searchParams.get("limit")) || 200));
  const items = searchCatalogItems(search, name);
  return NextResponse.json({ items: items.slice(0, limit), total: items.length });
}

export async function POST(request: NextRequest) {
  try {
    const input = await request.json();
    const customFields = [
      ["cf_group", input.cf_group], ["cf_sub_group", input.cf_sub_group],
      ["cf_finish", input.cf_finish], ["cf_profile_code", input.cf_profile_code],
      ["cf_height", input.cf_height], ["cf_width", input.cf_width],
      ["cf_thickness", input.cf_thickness], ["cf_msl", input.cf_msl], ["cf_rol", input.cf_rol],
    ].filter(([, value]) => value !== "" && value !== null && value !== undefined)
      .map(([api_name, value]) => ({ api_name, value }));

    const payload = {
      name: input.name,
      sku: input.sku || undefined,
      unit: input.unit || "Pcs",
      item_type: input.item_type || "inventory",
      description: input.description || undefined,
      custom_fields: customFields,
    };
    const body = await zohoInventoryPost<{ item?: ItemDetail; message?: string }>("/items", payload);
    if (!body.item) return NextResponse.json({ error: body.message || "Zoho item creation failed" }, { status: 500 });
    const item: ItemDetail = {
      ...body.item,
      name: body.item.name || input.name,
      item_name: body.item.item_name || body.item.name || input.name,
      sku: body.item.sku || input.sku || "",
      unit: body.item.unit || input.unit || "Pcs",
      cf_group: input.cf_group || "",
      cf_sub_group: input.cf_sub_group || "",
      cf_finish: input.cf_finish || "",
      cf_profile_code: input.cf_profile_code || "",
      cf_height: input.cf_height || "",
      cf_width: input.cf_width || "",
      cf_thickness: input.cf_thickness || "",
      cf_msl: input.cf_msl ?? null,
      cf_rol: input.cf_rol ?? null,
      status: body.item.status || "Active",
      source: "Item module",
      sync_status: "Synced",
      last_modified_time: new Date().toISOString(),
    };
    return NextResponse.json({ item });
  } catch (error) {
    return NextResponse.json(toApiError(error), { status: 500 });
  }
}
