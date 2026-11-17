export type StoneOptiPriority = "quantity" | "length";

export type StoneOptiPiece = {
  id: string;
  item: string;
  group: string;
  length: number;
  width: number;
  qty: number;
  thickness: number;
  finish: string;
  kind: "sheet" | "length";
  filler?: boolean;
};

export type StoneOptiGroupSetting = { stockLength: number; stockWidth: number };
export type StoneOptiSettings = {
  priority: StoneOptiPriority;
  grainLocked: boolean;
  overproduction: boolean;
  kerf: number;
  trim: number;
  groups: Record<string, StoneOptiGroupSetting>;
};

export type StoneOptiPlacement = {
  x: number; y: number; length: number; width: number; rotated: boolean;
  piece: StoneOptiPiece; extra: boolean;
};
export type StoneOptiLayout = {
  number: number; length: number; width: number;
  placements: StoneOptiPlacement[];
  offcuts: Array<{ x: number; y: number; length: number; width: number }>;
};
export type StoneOptiGroupResult = {
  key: string; label: string; kind: StoneOptiPiece["kind"];
  stockLength: number; stockWidth: number; layouts: StoneOptiLayout[];
  requiredPieces: number; extraPieces: number; wastePct: number; remainingWastePct: number;
  unplaced: Array<{ piece: StoneOptiPiece; reason: string }>;
};
export type StoneOptiResult = { groups: StoneOptiGroupResult[]; generatedAt: string };

const pieceShapeKey = (p: StoneOptiPiece) => `${p.length}|${p.width}|${p.item}`;
export const stoneOptiGroupKey = (p: Pick<StoneOptiPiece, "kind" | "group" | "finish" | "thickness">) =>
  JSON.stringify([p.kind, p.group, p.finish, p.thickness]);

export function stoneOptiGroups(pieces: StoneOptiPiece[]) {
  const groups = new Map<string, { key: string; label: string; kind: StoneOptiPiece["kind"]; pieces: StoneOptiPiece[] }>();
  for (const piece of pieces) {
    const key = stoneOptiGroupKey(piece);
    const found = groups.get(key);
    if (found) found.pieces.push(piece);
    else groups.set(key, { key, label: `${piece.group}${piece.finish ? ` · ${piece.finish}` : ""}${piece.thickness ? ` · ${piece.thickness}mm` : ""}`, kind: piece.kind, pieces: [piece] });
  }
  return [...groups.values()];
}

function validate(settings: StoneOptiSettings, key: string, kind: StoneOptiPiece["kind"]) {
  const stock = settings.groups[key];
  if (!stock || !Number.isFinite(stock.stockLength) || stock.stockLength <= 0 || stock.stockLength > 50000)
    throw new Error("Enter a valid stock length for every material.");
  if (kind === "sheet" && (!Number.isFinite(stock.stockWidth) || stock.stockWidth <= 0 || stock.stockWidth > 50000))
    throw new Error("Enter a valid stock width for every sheet material.");
  if (![settings.kerf, settings.trim].every(n => Number.isFinite(n) && n >= 0 && n <= 30))
    throw new Error("Kerf and trim must be between 0 and 30 mm.");
}

function quantities(pieces: StoneOptiPiece[]) {
  const result = new Map<string, number>();
  pieces.forEach(p => result.set(pieceShapeKey(p), (result.get(pieceShapeKey(p)) || 0) + p.qty));
  return result;
}

function sortPieces(pieces: StoneOptiPiece[], priority: StoneOptiPriority) {
  const qty = quantities(pieces);
  return [...pieces].sort((a, b) => priority === "quantity"
    ? (qty.get(pieceShapeKey(b)) || 0) - (qty.get(pieceShapeKey(a)) || 0) || b.length - a.length || b.width - a.width
    : b.length - a.length || (qty.get(pieceShapeKey(b)) || 0) - (qty.get(pieceShapeKey(a)) || 0) || b.width - a.width);
}

function addOffcuts(layout: StoneOptiLayout, rect: StoneOptiLayout["offcuts"][number], length: number, width: number, kerf: number) {
  const right = rect.length - length - kerf;
  const below = rect.width - width - kerf;
  if (right > 0) layout.offcuts.push({ x: rect.x + length + kerf, y: rect.y, length: right, width });
  if (below > 0) layout.offcuts.push({ x: rect.x, y: rect.y + width + kerf, length: rect.length, width: below });
}

function optimizeGroup(group: ReturnType<typeof stoneOptiGroups>[number], settings: StoneOptiSettings): StoneOptiGroupResult {
  validate(settings, group.key, group.kind);
  const configured = settings.groups[group.key];
  const stockLength = configured.stockLength;
  const stockWidth = group.kind === "length" ? 1 : configured.stockWidth;
  const trim = group.kind === "length" ? 0 : settings.trim;
  const kerf = settings.kerf;
  const sorted = sortPieces(group.pieces, settings.priority);
  const expanded: StoneOptiPiece[] = [];
  for (const piece of sorted) {
    const count = Math.round(piece.qty);
    if (!Number.isSafeInteger(count) || count < 1 || piece.length <= 0 || (piece.kind === "sheet" && piece.width <= 0)) continue;
    if (expanded.length + count > 10000) throw new Error(`${group.label} exceeds the 10,000-piece browser limit.`);
    for (let i = 0; i < count; i++) expanded.push(piece);
  }

  const layouts: StoneOptiLayout[] = [];
  const unplaced: StoneOptiGroupResult["unplaced"] = [];
  expanded.forEach((piece) => {
    const shapes = [{ length: piece.length, width: group.kind === "length" ? 1 : piece.width, rotated: false },
      ...(!settings.grainLocked && group.kind === "sheet" && piece.length !== piece.width
        ? [{ length: piece.width, width: piece.length, rotated: true }] : [])];
    if (!shapes.some(s => s.length <= stockLength - trim * 2 && s.width <= stockWidth - trim * 2)) {
      if (!unplaced.some(row => row.piece.id === piece.id)) unplaced.push({ piece, reason: "Part exceeds usable stock size or grain direction." });
      return;
    }
    let best: { layout: StoneOptiLayout; rectIndex: number; shape: typeof shapes[number]; score: number } | undefined;
    const inspect = (layout: StoneOptiLayout) => layout.offcuts.forEach((rect, rectIndex) => shapes.forEach(shape => {
      if (shape.length <= rect.length && shape.width <= rect.width) {
        const score = rect.length * rect.width - shape.length * shape.width;
        if (!best || score < best.score) best = { layout, rectIndex, shape, score };
      }
    }));
    layouts.forEach(inspect);
    if (!best) {
      const layout: StoneOptiLayout = { number: layouts.length + 1, length: stockLength, width: stockWidth, placements: [],
        offcuts: [{ x: trim, y: trim, length: stockLength - trim * 2, width: stockWidth - trim * 2 }] };
      layouts.push(layout); inspect(layout);
    }
    if (!best) return;
    const choice = best as { layout: StoneOptiLayout; rectIndex: number; shape: typeof shapes[number]; score: number };
    const rect = choice.layout.offcuts.splice(choice.rectIndex, 1)[0];
    choice.layout.placements.push({ x: rect.x, y: rect.y, length: choice.shape.length, width: choice.shape.width,
      rotated: choice.shape.rotated, piece, extra: false });
    addOffcuts(choice.layout, rect, choice.shape.length, choice.shape.width, kerf);
  });

  if (settings.overproduction && !unplaced.length) {
    const unique = new Map<string, StoneOptiPiece>();
    sorted.forEach(p => { if (!unique.has(pieceShapeKey(p))) unique.set(pieceShapeKey(p), p); });
    const candidates = [...unique.values()].sort((a, b) => Number(!!a.filler) - Number(!!b.filler) || b.length - a.length);
    layouts.forEach(layout => {
      const existingLengths = new Set(layout.placements.map(p => p.length));
      let added = 0;
      while (added < 1000) {
        let match: { piece: StoneOptiPiece; rectIndex: number; length: number; width: number; rotated: boolean } | undefined;
        for (const tier of [0, 1, 2]) {
          for (const piece of candidates) {
            if (tier < 2 && piece.filler) continue;
            if (tier === 2 && !piece.filler) continue;
            const shapes = [{ length: piece.length, width: group.kind === "length" ? 1 : piece.width, rotated: false },
              ...(!settings.grainLocked && group.kind === "sheet" && piece.length !== piece.width ? [{ length: piece.width, width: piece.length, rotated: true }] : [])]
              .filter(s => tier !== 0 || existingLengths.has(s.length));
            for (let rectIndex = 0; rectIndex < layout.offcuts.length && !match; rectIndex++) {
              const rect = layout.offcuts[rectIndex];
              const shape = shapes.find(s => s.length <= rect.length && s.width <= rect.width);
              if (shape) match = { piece, rectIndex, ...shape };
            }
            if (match) break;
          }
          if (match) break;
        }
        if (!match) break;
        const rect = layout.offcuts.splice(match.rectIndex, 1)[0];
        layout.placements.push({ x: rect.x, y: rect.y, length: match.length, width: match.width, rotated: match.rotated, piece: match.piece, extra: true });
        addOffcuts(layout, rect, match.length, match.width, kerf); added++;
      }
    });
  }

  const stockArea = layouts.length * stockLength * stockWidth;
  const requiredArea = layouts.flatMap(l => l.placements).filter(p => !p.extra).reduce((n, p) => n + p.length * p.width, 0);
  const usedArea = layouts.flatMap(l => l.placements).reduce((n, p) => n + p.length * p.width, 0);
  return { key: group.key, label: group.label, kind: group.kind, stockLength, stockWidth, layouts,
    requiredPieces: layouts.flatMap(l => l.placements).filter(p => !p.extra).length,
    extraPieces: layouts.flatMap(l => l.placements).filter(p => p.extra).length,
    wastePct: stockArea ? Math.max(0, 100 * (1 - requiredArea / stockArea)) : 0,
    remainingWastePct: stockArea ? Math.max(0, 100 * (1 - usedArea / stockArea)) : 0, unplaced };
}

export function optimizeStoneProject(pieces: StoneOptiPiece[], settings: StoneOptiSettings): StoneOptiResult {
  if (!pieces.length) throw new Error("The current project has no cuttable stone or profile parts.");
  return { groups: stoneOptiGroups(pieces).map(group => optimizeGroup(group, settings)), generatedAt: new Date().toISOString() };
}

const htmlEscape = (value: unknown) => String(value).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]!);
export function stoneOptiHtml(result: StoneOptiResult, projectName: string, settings: StoneOptiSettings) {
  const rows = result.groups.map(g => `<tr><td>${htmlEscape(g.label)}</td><td>${g.layouts.length}</td><td>${g.requiredPieces}</td><td>${g.extraPieces}</td><td>${g.wastePct.toFixed(2)}%</td><td>${g.remainingWastePct.toFixed(2)}%</td><td>${g.unplaced.length}</td></tr>`).join("");
  const charts = result.groups.flatMap(g => g.layouts.map(layout => {
    const placements = layout.placements.map((p, i) => `<rect x="${p.x}" y="${p.y}" width="${p.length}" height="${p.width}" fill="${p.extra ? "#f7d77b" : "#d9e9df"}" stroke="#405c4c"/><text x="${p.x + p.length / 2}" y="${p.y + p.width / 2}" text-anchor="middle" font-size="${Math.max(8, Math.min(26, Math.min(p.length, p.width) / 6))}">${i + 1} · ${p.length}${g.kind === "sheet" ? `×${p.width}` : ""}${p.extra ? " STOCK" : ""}</text>`).join("");
    const parts = layout.placements.map((p, i) => `<tr><td>${i + 1}</td><td>${htmlEscape(p.piece.item)}</td><td>${p.piece.length}</td><td>${g.kind === "sheet" ? p.piece.width : "—"}</td><td>${p.extra ? "Extra stock" : "Required"}</td></tr>`).join("");
    return `<section><h2>${htmlEscape(g.label)} · ${g.kind === "sheet" ? "Sheet" : "Length"} ${layout.number}</h2><p>Stock: ${g.stockLength}${g.kind === "sheet" ? ` × ${g.stockWidth}` : ""} mm · Grain ${settings.grainLocked ? "locked" : "rotation allowed"} · Priority ${settings.priority}</p><svg viewBox="0 0 ${layout.length} ${Math.max(layout.width, 80)}">${placements}</svg><table><thead><tr><th>#</th><th>Part</th><th>Length</th><th>Width</th><th>Purpose</th></tr></thead><tbody>${parts}</tbody></table></section>`;
  })).join("");
  return `<!doctype html><html><head><meta charset="utf-8"><title>Stone OPTI</title><style>body{font:14px Arial;margin:24px}table{border-collapse:collapse;width:100%;margin:12px 0}th,td{border:1px solid #bbb;padding:6px;text-align:left}th{background:#e8eee9}section{break-after:page}svg{width:100%;max-height:62vh;border:1px solid #555}text{dominant-baseline:middle}@media print{button{display:none}}</style></head><body><button onclick="window.print()">Print / Save as PDF</button><h1>${htmlEscape(projectName)} · OPTI cutting plan</h1><p>Single retail project. Yellow pieces are optional extra stock cut only from existing offcuts.</p><table><thead><tr><th>Material</th><th>Stock units</th><th>Required</th><th>Extra stock</th><th>Waste before stock</th><th>Remaining waste</th><th>Unplaced</th></tr></thead><tbody>${rows}</tbody></table>${charts}</body></html>`;
}
