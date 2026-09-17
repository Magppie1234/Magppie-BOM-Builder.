import catalogData from "@/data/item_catalog.json";
import type { ItemDetail } from "@/lib/types";

export const ITEM_CATALOG_BY_ID = catalogData as unknown as Record<string, ItemDetail>;
export const ITEM_CATALOG = Object.values(ITEM_CATALOG_BY_ID);

function searchable(item: ItemDetail): string {
  return [
    item.name,
    item.item_name,
    item.sku,
    item.cf_group,
    item.cf_sub_group,
    item.cf_finish,
    item.cf_profile_code,
    item.unit,
  ].filter(Boolean).join(" ").toLowerCase();
}

export function searchCatalogItems(query: string, name = ""): ItemDetail[] {
  const terms = (query || name).toLowerCase().trim().split(/\s+/).filter(Boolean);
  if (!terms.length) return ITEM_CATALOG;
  return ITEM_CATALOG.filter((item) => {
    const haystack = searchable(item);
    return terms.every((term) => haystack.includes(term));
  });
}

export function getCatalogItem(itemId: string): ItemDetail | undefined {
  return ITEM_CATALOG.find((item) => item.item_id === itemId);
}
