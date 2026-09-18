import finishReferenceData from "@/data/finish_reference_map.json";

const referenceMap = finishReferenceData as Record<string, string[]>;

export function normalizeFinishName(value: string): string {
  return String(value || "")
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Za-z0-9]+/g, " ")
    .trim()
    .replace(/\s+/g, " ")
    .toUpperCase();
}

const aliasesByPlanningName = new Map<string, string[]>();
for (const [planningName, references] of Object.entries(referenceMap)) {
  const shortName = planningName.replace(/^MAGPPIE\s+/, "");
  aliasesByPlanningName.set(normalizeFinishName(planningName), Array.from(new Set([planningName, ...references].filter(Boolean))));
  aliasesByPlanningName.set(normalizeFinishName(shortName), Array.from(new Set([shortName, ...references].filter(Boolean))));
}

/** Returns the planning name plus every catalogue/reference-colour alias from the supplied workbook. */
export function finishSearchAliases(finish: string): string[] {
  const normalized = normalizeFinishName(finish);
  const mapped = aliasesByPlanningName.get(normalized);
  return mapped ? Array.from(new Set([finish, ...mapped].filter(Boolean))) : [finish];
}
