import type { ItemDetail } from "@/lib/types";
import { finishSearchAliases } from "@/lib/finishReferences";

export const LOCAL_ITEM_CATALOG_KEY = "magppie-item-catalog-v1";

export function readLocalCatalogItems(): ItemDetail[] {
  if (typeof window === "undefined") return [];
  try {
    const parsed = JSON.parse(window.localStorage.getItem(LOCAL_ITEM_CATALOG_KEY) || "[]");
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

export function saveLocalCatalogItem(item: ItemDetail): void {
  const items = readLocalCatalogItems();
  const index = items.findIndex((entry) => entry.item_id === item.item_id || (!!item.sku && entry.sku === item.sku));
  if (index >= 0) items[index] = item;
  else items.unshift(item);
  window.localStorage.setItem(LOCAL_ITEM_CATALOG_KEY, JSON.stringify(items));
  window.dispatchEvent(new CustomEvent("magppie-item-catalog-changed"));
}

function itemText(item: ItemDetail): string {
  return [item.name, item.item_name, item.sku, item.cf_group, item.cf_sub_group, item.cf_finish, item.cf_profile_code, item.unit]
    .filter(Boolean).join(" ").toLowerCase();
}

export function mergeLocalCatalogItems(items: ItemDetail[], query = ""): ItemDetail[] {
  const termGroups = finishSearchAliases(query)
    .map((alias) => alias.toLowerCase().trim().split(/\s+/).filter(Boolean))
    .filter((terms) => terms.length > 0);
  const additions = readLocalCatalogItems().filter((item) => {
    const text = itemText(item);
    return termGroups.length === 0 || termGroups.some((terms) => terms.every((term) => text.includes(term)));
  });
  const merged = new Map<string, ItemDetail>();
  [...additions, ...items].forEach((item) => merged.set(item.item_id, item));
  return Array.from(merged.values());
}
