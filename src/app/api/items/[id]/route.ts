import { NextRequest, NextResponse } from "next/server";
import { getCatalogItem } from "@/lib/itemCatalog";

export const dynamic = "force-dynamic";

export async function GET(_request: NextRequest, context: { params: Promise<{ id: string }> }) {
  const { id } = await context.params;
  const item = getCatalogItem(id);
  if (!item) return NextResponse.json({ error: "Item not found in item catalogue" }, { status: 404 });
  return NextResponse.json({ item, source: "item-catalog" });
}
