// Kho dữ liệu dạng "collection/doc" (cùng API mà UI đang dùng: collection().
// onSnapshot/add/doc(), doc().set/update/delete) chạy trên SharePoint list.
// Cache trong bộ nhớ: ghi xong cập nhật cache + báo listener ngay; dữ liệu
// người khác sửa được kéo về theo chu kỳ CONFIG.pollMs.
import { CONFIG } from "./config.js";

export class PermissionError extends Error {}

const newId = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
const clone = (o) => JSON.parse(JSON.stringify(o));

// ---------- adapters ----------
function spAdapter(getToken) {
  const site = `https://${CONFIG.siteHost}${CONFIG.sitePath}`;
  const items = (col) => `${site}/_api/web/lists/getbytitle('${encodeURIComponent(CONFIG.lists[col])}')/items`;
  const F = CONFIG.dataField;
  async function call(url, { method = "GET", body, headers = {} } = {}) {
    const token = await getToken();
    const res = await fetch(url, {
      method,
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=nometadata",
        ...headers,
      },
      body: body ? JSON.stringify(body) : undefined,
    });
    if (res.status === 401 || res.status === 403) throw new PermissionError(`SharePoint ${res.status}`);
    if (!res.ok) {
      let msg = `SharePoint ${res.status}`;
      try { msg = (await res.json())["odata.error"]?.message?.value || msg; } catch {}
      throw new Error(msg);
    }
    return res.status === 204 ? null : res.json().catch(() => null);
  }
  return {
    async load(col) {
      const out = [];
      let url = `${items(col)}?$select=Id,Title,${F}&$top=5000`;
      while (url) {
        const page = await call(url);
        for (const it of page.value) {
          try { out.push({ id: it.Title, itemId: it.Id, data: JSON.parse(it[F] || "{}") }); } catch {}
        }
        url = page["odata.nextLink"] || null;
      }
      return out;
    },
    async create(col, id, data) {
      const r = await call(items(col), { method: "POST", body: { Title: id, [F]: JSON.stringify(data) } });
      return r?.Id;
    },
    async save(col, itemId, data) {
      await call(`${items(col)}(${itemId})`, {
        method: "POST", headers: { "X-HTTP-Method": "MERGE", "IF-MATCH": "*" },
        body: { [F]: JSON.stringify(data) },
      });
    },
    async remove(col, itemId) {
      await call(`${items(col)}(${itemId})`, {
        method: "POST", headers: { "X-HTTP-Method": "DELETE", "IF-MATCH": "*" },
      });
    },
  };
}

function demoAdapter() {
  const KEY = "okr.demo.v1";
  const read = () => { try { return JSON.parse(localStorage.getItem(KEY)) || null; } catch { return null; } };
  const write = (db) => { try { localStorage.setItem(KEY, JSON.stringify(db)); } catch {} };
  let seeded = null;
  async function db() {
    let d = read();
    if (d) return d;
    if (!seeded) seeded = fetch("../data/okr-seed.json").then((r) => (r.ok ? r.json() : {})).catch(() => ({}));
    const s = await seeded;
    d = { objectives: s.objectives || {}, kpis: s.kpis || {}, checkins: s.checkins || {} };
    write(d);
    return d;
  }
  return {
    async load(col) { const d = await db(); return Object.entries(d[col] || {}).map(([id, data]) => ({ id, itemId: id, data })); },
    async create(col, id, data) { const d = await db(); d[col][id] = data; write(d); return id; },
    async save(col, itemId, data) { const d = await db(); d[col][itemId] = data; write(d); },
    async remove(col, itemId) { const d = await db(); delete d[col][itemId]; write(d); },
  };
}

// ---------- store ----------
export function createStore({ demo = false, getToken } = {}) {
  const api = demo ? demoAdapter() : spAdapter(getToken);
  const cache = {};      // col -> Map(id -> {itemId, data})
  const listeners = {};  // col -> [{fn, err, order, limit}]
  const loading = {};

  const snapshot = (col, l) => {
    let rows = [...(cache[col]?.entries() || [])].map(([id, v]) => ({ id, data: v.data }));
    if (l.order) {
      const [f, dir] = l.order;
      rows.sort((a, b) => String(a.data[f] ?? "").localeCompare(String(b.data[f] ?? "")) * (dir === "desc" ? -1 : 1));
    }
    if (l.limit) rows = rows.slice(0, l.limit);
    return { docs: rows.map((r) => ({ id: r.id, data: () => clone(r.data) })) };
  };
  const emit = (col) => (listeners[col] || []).forEach((l) => { try { l.fn(snapshot(col, l)); } catch (e) { console.error(e); } });

  async function refresh(col) {
    if (loading[col]) return loading[col];
    loading[col] = (async () => {
      try {
        const rows = await api.load(col);
        cache[col] = new Map(rows.map((r) => [r.id, { itemId: r.itemId, data: r.data }]));
        emit(col);
      } catch (e) {
        (listeners[col] || []).forEach((l) => l.err && l.err(e));
        throw e;
      } finally { loading[col] = null; }
    })();
    return loading[col].catch(() => {});
  }
  const ensure = async (col) => { if (!cache[col]) await refresh(col); return cache[col] || new Map(); };

  function docRef(col, id) {
    return {
      id,
      async set(data) {
        const m = await ensure(col);
        const cur = m.get(id);
        const body = clone(data);
        if (cur) await api.save(col, cur.itemId, body);
        const itemId = cur ? cur.itemId : await api.create(col, id, body);
        m.set(id, { itemId, data: body }); emit(col);
      },
      async update(patch) {
        const m = await ensure(col);
        const cur = m.get(id);
        if (!cur) throw new Error("Không tìm thấy dữ liệu để cập nhật — tải lại trang");
        const body = { ...cur.data, ...clone(patch) };
        await api.save(col, cur.itemId, body);
        m.set(id, { itemId: cur.itemId, data: body }); emit(col);
      },
      async delete() {
        const m = await ensure(col);
        const cur = m.get(id);
        if (cur) await api.remove(col, cur.itemId);
        m.delete(id); emit(col);
      },
    };
  }

  function query(col, order = null, limit = 0) {
    return {
      orderBy: (f, dir = "asc") => query(col, [f, dir], limit),
      limit: (n) => query(col, order, n),
      onSnapshot(fn, err) {
        const l = { fn, err, order, limit };
        (listeners[col] ||= []).push(l);
        if (cache[col]) fn(snapshot(col, l)); else refresh(col);
        return () => { listeners[col] = listeners[col].filter((x) => x !== l); };
      },
      doc: (id) => docRef(col, id || newId()),
      async add(data) { const ref = docRef(col, newId()); await ref.set(data); return ref; },
    };
  }

  // Kéo dữ liệu mới định kỳ (khi tab đang hiển thị) + khi quay lại tab.
  const pollAll = () => Object.keys(listeners).forEach((c) => refresh(c));
  setInterval(() => { if (document.visibilityState === "visible") pollAll(); }, CONFIG.pollMs);
  document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible") pollAll(); });

  return {
    collection: (col) => query(col),
    doc: (path) => { const [col, id] = path.split("/"); return docRef(col, id); },
    refresh: pollAll,
  };
}
