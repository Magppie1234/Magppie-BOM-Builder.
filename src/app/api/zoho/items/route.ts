import { NextRequest, NextResponse } from "next/server";
import { toApiError, zohoInventoryPost } from "@/lib/zoho";
import { searchCatalogItems } from "@/lib/itemCatalog";
import type { ItemDetail } from "@/lib/types";

export const dynamic = "force-dynamic";

export async function GET(request: NextRequest) {
  const { searchParams } = new URL(request.url);
  const search = searchParams.get("search") || "";
  const name = searchParams.get("name") || "";
  const items = searchCatalogItems(search, name);
  return NextResponse.json({ items: items.slice(0, 200), total: items.length, source: "item-catalog" });
}

export async function POST(request: NextRequest) {
  try {
    const payload = await request.json();
    const body = await zohoInventoryPost<{ item?: ItemDetail; code?: number; message?: string }>("/items", payload);

    if (!body.item) {
      return NextResponse.json({ error: body.message ?? "Creation failed" }, { status: 500 });
    }

    return NextResponse.json({ item: body.item });
  } catch (error) {
    return NextResponse.json(toApiError(error), { status: 500 });
  }
}
