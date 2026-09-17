import { NextRequest, NextResponse } from "next/server";
import { toApiError, zohoInventoryPut } from "@/lib/zoho";
import { getCatalogItem } from "@/lib/itemCatalog";
import type { ItemDetail } from "@/lib/types";

export const dynamic = "force-dynamic";

type ItemResponse = {
  item?: ItemDetail;
};

export async function GET(_request: NextRequest, context: { params: Promise<{ id: string }> }) {
  const { id } = await context.params;
  const item = getCatalogItem(id);
  if (!item) return NextResponse.json({ error: "Item not found in item catalogue" }, { status: 404 });
  return NextResponse.json({ item, source: "item-catalog" });
}

export async function PUT(request: NextRequest, context: { params: Promise<{ id: string }> }) {
  try {
    const { id } = await context.params;
    const payload = await request.json();
    const body = await zohoInventoryPut<{ item?: ItemDetail; code?: number; message?: string }>(
      `/items/${encodeURIComponent(id)}`,
      payload,
    );

    if (!body.item) {
      return NextResponse.json({ error: body.message ?? "Update failed" }, { status: 500 });
    }

    return NextResponse.json({ item: body.item });
  } catch (error) {
    return NextResponse.json(toApiError(error), { status: 500 });
  }
}
