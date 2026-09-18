"use client";

import { FormEvent, useEffect, useMemo, useState } from "react";
import type { ItemDetail } from "@/lib/types";
import { mergeLocalCatalogItems, saveLocalCatalogItem } from "@/lib/itemCatalogClient";

type FormState = {
  name: string; sku: string; cf_group: string; cf_sub_group: string; cf_finish: string;
  cf_profile_code: string; cf_height: string; cf_width: string; cf_thickness: string;
  cf_msl: string; cf_rol: string; unit: string; item_type: string; description: string;
};

const EMPTY_FORM: FormState = {
  name: "", sku: "", cf_group: "", cf_sub_group: "", cf_finish: "", cf_profile_code: "",
  cf_height: "", cf_width: "", cf_thickness: "", cf_msl: "", cf_rol: "",
  unit: "Pcs", item_type: "inventory", description: "",
};

const text = (value: unknown) => value === null || value === undefined ? "" : String(value);

export function ItemCatalog() {
  const [items, setItems] = useState<ItemDetail[]>([]);
  const [loading, setLoading] = useState(true);
  const [query, setQuery] = useState("");
  const [group, setGroup] = useState("");
  const [subgroup, setSubgroup] = useState("");
  const [status, setStatus] = useState("");
  const [showCreate, setShowCreate] = useState(false);
  const [form, setForm] = useState<FormState>(EMPTY_FORM);
  const [saving, setSaving] = useState(false);
  const [message, setMessage] = useState("");

  async function loadItems() {
    setLoading(true);
    try {
      const response = await fetch("/api/items?limit=20000", { cache: "no-store" });
      const body = await response.json();
      if (!response.ok) throw new Error(body.error || "Unable to load items");
      setItems(mergeLocalCatalogItems(body.items || []));
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "Unable to load items");
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => { void loadItems(); }, []);

  const groups = useMemo(() => Array.from(new Set(items.map((i) => text(i.cf_group)).filter(Boolean))).sort(), [items]);
  const subgroups = useMemo(() => Array.from(new Set(items.filter((i) => !group || i.cf_group === group).map((i) => text(i.cf_sub_group)).filter(Boolean))).sort(), [items, group]);
  const filtered = useMemo(() => {
    const terms = query.toLowerCase().trim().split(/\s+/).filter(Boolean);
    return items.filter((item) => {
      const haystack = [item.name, item.sku, item.cf_group, item.cf_sub_group, item.cf_finish, item.cf_profile_code, item.unit, item.item_id]
        .filter(Boolean).join(" ").toLowerCase();
      return terms.every((term) => haystack.includes(term))
        && (!group || item.cf_group === group)
        && (!subgroup || item.cf_sub_group === subgroup)
        && (!status || text(item.status).toLowerCase() === status.toLowerCase());
    });
  }, [items, query, group, subgroup, status]);

  function update(field: keyof FormState, value: string) {
    setForm((current) => ({ ...current, [field]: value }));
  }

  async function createItem(event: FormEvent) {
    event.preventDefault();
    if (!form.name.trim()) return;
    setSaving(true);
    setMessage("");
    try {
      const response = await fetch("/api/items", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          ...form,
          cf_msl: form.cf_msl === "" ? null : Number(form.cf_msl),
          cf_rol: form.cf_rol === "" ? null : Number(form.cf_rol),
        }),
      });
      const body = await response.json();
      if (!response.ok || !body.item) throw new Error(body.error || "Item creation failed");
      saveLocalCatalogItem(body.item);
      setItems((current) => mergeLocalCatalogItems(current));
      setForm(EMPTY_FORM);
      setShowCreate(false);
      setMessage(`Created ${body.item.name || body.item.item_name} in Zoho and saved it in this browser.`);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "Item creation failed");
    } finally {
      setSaving(false);
    }
  }

  return (
    <main className="item-page">
      <header className="item-header">
        <div>
          <p className="item-kicker">Magppie item master</p>
          <h1>Items</h1>
          <p>{items.length.toLocaleString()} available items. Planning searches use this catalogue instead of Zoho item searches.</p>
        </div>
        <nav className="item-nav">
          <a href="/planning">Planning</a><a href="/builder">Builder</a><a href="/">BOM Dashboard</a>
          <button className="item-primary" onClick={() => setShowCreate((value) => !value)}>{showCreate ? "Close" : "Create item"}</button>
        </nav>
      </header>

      {message && <div className="item-message">{message}</div>}

      {showCreate && (
        <form className="item-create" onSubmit={createItem}>
          <div className="item-section-title"><h2>Create item</h2><p>Creates the item in Zoho, then adds it to this browser’s catalogue.</p></div>
          <div className="item-form-grid">
            {([
              ["name", "Item Name", true], ["sku", "SKU", false], ["cf_group", "Group", false],
              ["cf_sub_group", "Subgroup", false], ["cf_finish", "Finish", false], ["cf_profile_code", "Profile", false],
              ["cf_height", "Height", false], ["cf_width", "Width", false], ["cf_thickness", "Thickness", false],
              ["cf_msl", "MSL", false], ["cf_rol", "ROL", false], ["unit", "Unit/UOM", false],
            ] as Array<[keyof FormState, string, boolean]>).map(([field, label, required]) => (
              <label key={field}>{label}<input value={form[field]} required={required} onChange={(e) => update(field, e.target.value)} /></label>
            ))}
            <label>Item Type<select value={form.item_type} onChange={(e) => update("item_type", e.target.value)}><option value="inventory">Inventory</option><option value="sales">Sales</option><option value="purchases">Purchases</option><option value="service">Service</option></select></label>
            <label className="item-form-wide">Description<textarea value={form.description} onChange={(e) => update("description", e.target.value)} /></label>
          </div>
          <div className="item-form-actions"><button type="submit" className="item-primary" disabled={saving}>{saving ? "Creating…" : "Create in Zoho"}</button></div>
        </form>
      )}

      <section className="item-controls">
        <input className="item-search" placeholder="Search name, SKU, profile, finish, group or Zoho ID" value={query} onChange={(e) => setQuery(e.target.value)} />
        <select value={group} onChange={(e) => { setGroup(e.target.value); setSubgroup(""); }}><option value="">All groups</option>{groups.map((value) => <option key={value}>{value}</option>)}</select>
        <select value={subgroup} onChange={(e) => setSubgroup(e.target.value)}><option value="">All subgroups</option>{subgroups.map((value) => <option key={value}>{value}</option>)}</select>
        <select value={status} onChange={(e) => setStatus(e.target.value)}><option value="">All statuses</option><option>Active</option><option>Inactive</option></select>
        <strong>{filtered.length.toLocaleString()} results</strong>
      </section>

      <section className="item-table-wrap">
        {loading ? <div className="item-empty">Loading item catalogue…</div> : (
          <table className="item-table">
            <thead><tr><th>Item Name</th><th>SKU</th><th>Group</th><th>Subgroup</th><th>Finish</th><th>Profile</th><th>H</th><th>W</th><th>T</th><th>MSL</th><th>ROL</th><th>UOM</th><th>Status</th><th>Zoho Item ID</th></tr></thead>
            <tbody>{filtered.map((item) => <tr key={item.item_id}>
              <td className="item-name">{item.name || item.item_name}</td><td>{item.sku || "—"}</td><td>{item.cf_group || "—"}</td><td>{item.cf_sub_group || "—"}</td>
              <td>{item.cf_finish || "—"}</td><td>{item.cf_profile_code || "—"}</td><td>{text(item.cf_height) || "—"}</td><td>{text(item.cf_width) || "—"}</td><td>{text(item.cf_thickness) || "—"}</td>
              <td>{text(item.cf_msl) || "—"}</td><td>{text(item.cf_rol) || "—"}</td><td>{item.unit || "—"}</td><td>{item.status || "—"}</td><td className="item-id">{item.item_id}</td>
            </tr>)}</tbody>
          </table>
        )}
      </section>
    </main>
  );
}
