import { useState, useEffect, useRef } from "react";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import { supabase } from "./supabase";

/* ── Utils ──────────────────────────────────────────────────── */
const uid = () => Math.random().toString(36).slice(2, 11);
const n = (v) => parseFloat(v) || 0;
const fmt = (v) =>
  v != null && !isNaN(v) ? Math.round(v).toLocaleString("ru-RU") : "—";
const fmtDec = (v, d = 1) =>
  v != null && !isNaN(v) ? (+v).toFixed(d).replace(".", ",") : "—";

const compressPhoto = (file) =>
  new Promise((res) => {
    const fr = new FileReader();
    fr.onload = ({ target }) => {
      const img = new Image();
      img.onload = () => {
        const MAX = 260;
        const r = Math.min(MAX / img.width, MAX / img.height, 1);
        const [w, h] = [Math.round(img.width * r), Math.round(img.height * r)];
        const cv = document.createElement("canvas");
        cv.width = w; cv.height = h;
        cv.getContext("2d").drawImage(img, 0, 0, w, h);
        res(cv.toDataURL("image/jpeg", 0.75));
      };
      img.src = target.result;
    };
    fr.readAsDataURL(file);
  });

/* ── Calculations ───────────────────────────────────────────── */
const deliveryRub = (order) =>
  n(order.deliveryUsd) * n(order.usdRate) || n(order.delivery) || 0;

const itemQty = (item) => Math.max(1, n(item.qty) || 1);

const totalUnits = (order) =>
  (order.items || []).reduce((s, i) => s + itemQty(i), 0);

const itemCalc = (item, order, futureUnits = null) => {
  const qty = itemQty(item);
  const units = futureUnits ?? Math.max(1, totalUnits(order));
  const rate = n(item.rateOverride) || n(order.rate);
  const rubPU = n(item.yuanPrice) * rate;
  const delPU = deliveryRub(order) / units;
  const withDelPU = rubPU + delPU;
  const profitPU = n(item.salePrice) - withDelPU;
  return {
    rate, qty,
    rubPU, delPU, withDelPU, profitPU,
    rubles: rubPU * qty,
    delShare: delPU * qty,
    withDel: withDelPU * qty,
    profit: profitPU * qty,
  };
};

const orderCalc = (order) =>
  (order.items || []).reduce(
    (a, item) => {
      const c = itemCalc(item, order);
      return { cost: a.cost + c.withDel, rev: a.rev + n(item.salePrice) * c.qty, profit: a.profit + c.profit };
    },
    { cost: 0, rev: 0, profit: 0 }
  );

/* ── Excel export ───────────────────────────────────────────── */
const exportSalesTable = async (order) => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Продажи");

  ws.columns = [
    { header: "Цена", key: "price", width: 12 },
    { header: "Фото", key: "photo", width: 18 },
    { header: "Название", key: "name", width: 36 },
  ];
  ws.getRow(1).font = { bold: true };

  const PX = 60;
  for (const item of order.items) {
    const qty = itemQty(item);
    for (let q = 0; q < qty; q++) {
      const row = ws.addRow({ price: n(item.salePrice), photo: "", name: item.name || "Без названия" });
      row.height = PX;
      row.getCell("price").numFmt = '#,##0 "₽"';

      if (item.photo) {
        const base64 = item.photo.split(",")[1];
        const imgId = wb.addImage({ base64, extension: "jpeg" });
        const r = row.number - 1;
        ws.addImage(imgId, { tl: { col: 1, row: r }, br: { col: 2, row: r + 1 }, editAs: "oneCell" });
      }
    }
  }

  const buf = await wb.xlsx.writeBuffer();
  const a = Object.assign(document.createElement("a"), {
    href: URL.createObjectURL(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" })),
    download: `${order.name || "заказ"}.xlsx`,
  });
  a.click();
  URL.revokeObjectURL(a.href);
};

/* ── Photo ZIP export ───────────────────────────────────────── */
const exportPhotos = async (order) => {
  const items = order.items.filter(i => i.photo);
  if (!items.length) return alert("В заказе нет фотографий");
  const zip = new JSZip();
  items.forEach((item, idx) => {
    const base64 = item.photo.split(",")[1];
    const name = (item.name || `товар_${idx + 1}`).replace(/[/\\?%*:|"<>]/g, "_");
    zip.file(`${idx + 1}_${name}.jpg`, base64, { base64: true });
  });
  const blob = await zip.generateAsync({ type: "blob" });
  const a = Object.assign(document.createElement("a"), {
    href: URL.createObjectURL(blob),
    download: `${order.name || "заказ"}_фото.zip`,
  });
  a.click();
  URL.revokeObjectURL(a.href);
};

/* ── Storage ────────────────────────────────────────────────── */
const SK = "otrk_v1";
const AK = "otrk_active";

const loadActiveId = () => { try { return localStorage.getItem(AK); } catch { return null; } };
const saveActiveId = (id) => { try { localStorage.setItem(AK, id ?? ""); } catch {} };

const lsGetOrders = () => {
  try { const r = localStorage.getItem(SK); return r ? (JSON.parse(r).orders ?? []) : []; }
  catch { return []; }
};

const lsSave = (orders, activeId) => {
  try { localStorage.setItem(SK, JSON.stringify({ orders, activeId })); } catch {}
};

const loadOrders = async () => {
  if (supabase) {
    try {
      const { data, error } = await supabase.from("orders").select("*").order("created_at");
      if (error) throw error;
      if (data?.length) return data.map((r) => ({ id: r.id, ...r.data }));
    } catch (e) { console.warn("Supabase load failed, using localStorage:", e.message); }
  }
  return lsGetOrders();
};

const dbUpsert = (order) => {
  if (!supabase) return;
  const { id, ...data } = order;
  supabase.from("orders")
    .upsert({ id, data, updated_at: new Date().toISOString() })
    .then(({ error }) => { if (error) console.warn("dbUpsert error:", error.message); });
};

const dbDelete = (id) => {
  if (!supabase) return;
  supabase.from("orders").delete().eq("id", id)
    .then(({ error }) => { if (error) console.warn("dbDelete error:", error.message); });
};

/* ── Label colors ───────────────────────────────────────────── */
const LABEL_PALETTE = [
  { bg: "#EFF6FF", border: "#BFDBFE", text: "#1D4ED8" },
  { bg: "#F0FDF4", border: "#BBF7D0", text: "#15803D" },
  { bg: "#FFF7ED", border: "#FED7AA", text: "#C2410C" },
  { bg: "#FAF5FF", border: "#E9D5FF", text: "#7C3AED" },
  { bg: "#FFF1F2", border: "#FECDD3", text: "#BE123C" },
];
const labelColor = (label) => {
  if (!label) return null;
  let h = 0;
  for (let i = 0; i < label.length; i++) h = (h * 31 + label.charCodeAt(i)) >>> 0;
  return LABEL_PALETTE[h % LABEL_PALETTE.length];
};

/* ── Design tokens ──────────────────────────────────────────── */
const T = {
  bg: "#F7F8FC",
  card: "#FFFFFF",
  border: "#E6EAF2",
  accent: "#E8320A",       // Taobao/Chinese commerce red
  accentDim: "#FFF2EF",
  text: "#1A1C2A",
  sub: "#8896AE",
  subDark: "#5B6880",
  green: "#0F6B3E",
  greenBg: "#EDFAF3",
  greenBdr: "#A7F0C8",
  red: "#C01A1A",
  redBg: "#FEF2F2",
  redBdr: "#FDC5C5",
  gold: "#B45309",
  goldBg: "#FFFBEB",
  shadow: "0 2px 10px rgba(26,28,42,0.07), 0 0 0 1px rgba(26,28,42,0.03)",
};

/* ── Base styles ────────────────────────────────────────────── */
const $ = {
  input: {
    width: "100%", padding: "10px 12px", border: `1px solid ${T.border}`,
    borderRadius: 10, fontSize: 15, outline: "none", boxSizing: "border-box",
    background: T.bg, color: T.text, WebkitAppearance: "none", fontFamily: "inherit",
  },
  btnRed: {
    padding: "11px 18px", background: T.accent, color: "#fff", border: "none",
    borderRadius: 10, fontWeight: 700, cursor: "pointer", fontSize: 14, fontFamily: "inherit",
  },
  btnOutline: {
    padding: "9px 14px", background: T.bg, color: T.text,
    border: `1px solid ${T.border}`, borderRadius: 10, cursor: "pointer",
    fontSize: 14, fontFamily: "inherit",
  },
  ico: {
    background: "none", border: "none", cursor: "pointer",
    padding: "6px 8px", borderRadius: 8, fontSize: 16, fontFamily: "inherit",
  },
};

/* ═══════════════════════════════════════════════════════════════
   Root App
   ═══════════════════════════════════════════════════════════════ */
const getHashId = () => {
  const h = window.location.hash;
  return h.startsWith("#order-") ? h.slice(7) : null;
};

export default function App() {
  const [orders, setOrders] = useState([]);
  const [activeId, setActiveId] = useState(null);
  const [modal, setModal] = useState(null);
  const [hashId, setHashId] = useState(getHashId);
  const [mobile, setMobile] = useState(window.innerWidth < 720);
  const [ready, setReady] = useState(false);

  useEffect(() => {
    const h = () => setMobile(window.innerWidth < 720);
    window.addEventListener("resize", h);
    return () => window.removeEventListener("resize", h);
  }, []);

  useEffect(() => {
    const onHash = () => setHashId(getHashId());
    window.addEventListener("hashchange", onHash);
    return () => window.removeEventListener("hashchange", onHash);
  }, []);

  useEffect(() => {
    loadOrders().then((ords) => {
      const aid = loadActiveId();
      if (ords.length) {
        setOrders(ords);
        setActiveId(ords.find((o) => o.id === aid) ? aid : ords[0].id);
      }
      setReady(true);
    });
  }, []);

  useEffect(() => {
    if (!supabase) return;
    const ch = supabase.channel("orders-sync")
      .on("postgres_changes", { event: "*", schema: "public", table: "orders" }, (payload) => {
        if (payload.eventType === "DELETE") {
          setOrders((prev) => prev.filter((o) => o.id !== payload.old.id));
        } else {
          const updated = { id: payload.new.id, ...payload.new.data };
          setOrders((prev) => {
            const exists = prev.some((o) => o.id === updated.id);
            return exists ? prev.map((o) => o.id === updated.id ? updated : o) : [...prev, updated];
          });
        }
      })
      .subscribe();
    return () => supabase.removeChannel(ch);
  }, []);

  const active = orders.find((o) => o.id === activeId) ?? null;

  const openOrder = (id) => {
    setActiveId(id); saveActiveId(id);
    if (mobile) window.location.hash = `order-${id}`;
  };

  const goBack = () => {
    if (mobile) window.history.back();
  };

  /* Order ops */
  const createOrder = (d) => {
    const o = { id: uid(), items: [], ...d };
    const next = [...orders, o];
    setOrders(next); lsSave(next, o.id); dbUpsert(o);
    openOrder(o.id); setModal(null);
  };
  const updateOrder = (id, d) => {
    const next = orders.map((o) => o.id === id ? { ...o, ...d } : o);
    setOrders(next); lsSave(next, activeId); dbUpsert(next.find((o) => o.id === id));
    setModal(null);
  };
  const deleteOrder = (id) => {
    const next = orders.filter((o) => o.id !== id);
    const nextId = next[0]?.id ?? null;
    setOrders(next); setActiveId(nextId); saveActiveId(nextId);
    lsSave(next, nextId); dbDelete(id);
    if (mobile) window.location.hash = "";
    setModal(null);
  };

  /* Item ops */
  const createItem = (d) => {
    const next = orders.map((o) => o.id !== activeId ? o : { ...o, items: [...o.items, { id: uid(), ...d }] });
    setOrders(next); lsSave(next, activeId); dbUpsert(next.find((o) => o.id === activeId));
    setModal(null);
  };
  const updateItem = (iid, d) => {
    const next = orders.map((o) => o.id !== activeId ? o : { ...o, items: o.items.map((i) => i.id === iid ? { ...i, ...d } : i) });
    setOrders(next); lsSave(next, activeId); dbUpsert(next.find((o) => o.id === activeId));
    setModal(null);
  };
  const deleteItem = (iid) => {
    const next = orders.map((o) => o.id !== activeId ? o : { ...o, items: o.items.filter((i) => i.id !== iid) });
    setOrders(next); lsSave(next, activeId); dbUpsert(next.find((o) => o.id === activeId));
  };

  if (!ready) return (
    <div style={{ height: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: T.bg, color: T.sub, fontFamily: "system-ui" }}>
      Загрузка…
    </div>
  );

  return (
    <div style={{ height: "100vh", display: "flex", background: T.bg, fontFamily: "-apple-system,'Helvetica Neue',system-ui,sans-serif", fontSize: 14, color: T.text, overflow: "hidden" }}>

      {/* ── SIDEBAR ─────────────────────────────────────────── */}
      {(!mobile || !hashId) && (
        <aside style={{ width: mobile ? "100%" : 264, flexShrink: 0, background: T.card, borderRight: `1px solid ${T.border}`, display: "flex", flexDirection: "column", overflow: "hidden" }}>
          {/* Brand header */}
          <header style={{ padding: "18px 16px 14px", borderBottom: `1px solid ${T.border}` }}>
            <div style={{ display: "flex", alignItems: "baseline", gap: 8, marginBottom: 14 }}>
              <span style={{ fontSize: 20, fontWeight: 900, color: T.accent, letterSpacing: "-0.5px" }}>ТРЕКЕР</span>
              <span style={{ fontSize: 11, color: T.sub, fontWeight: 500, letterSpacing: "0.5px" }}>ЗАКАЗОВ</span>
            </div>
            <button onClick={() => setModal({ type: "order" })} style={{ ...$.btnRed, width: "100%", borderRadius: 12, letterSpacing: "0.2px" }}>
              + Новый заказ
            </button>
          </header>

          {/* Order list */}
          <div style={{ flex: 1, overflowY: "auto", WebkitOverflowScrolling: "touch" }}>
            {orders.length === 0 ? (
              <div style={{ padding: "40px 24px", textAlign: "center", color: T.sub }}>
                <div style={{ fontSize: 36, marginBottom: 10, opacity: 0.5 }}>📦</div>
                <div style={{ fontWeight: 700, color: T.subDark, marginBottom: 4 }}>Нет заказов</div>
                <div style={{ fontSize: 12, lineHeight: 1.5 }}>Создай первый заказ,<br/>чтобы начать учёт</div>
              </div>
            ) : orders.map((o, i) => {
              const { rev, profit } = orderCalc(o);
              const act = o.id === activeId;
              const margin = o.items.length > 0 && rev > 0 ? (profit / rev) * 100 : null;
              return (
                <div key={o.id}
                  onClick={() => openOrder(o.id)}
                  style={{ padding: "12px 16px", cursor: "pointer", borderBottom: `1px solid ${T.border}`, borderLeft: `3px solid ${act ? T.accent : "transparent"}`, background: act ? T.accentDim : "transparent", transition: "background 0.1s" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 4 }}>
                    <div style={{ fontWeight: 700, fontSize: 13, color: act ? T.accent : T.text }}>
                      {o.name || `Заказ ${i + 1}`}
                    </div>
                    <div style={{ fontSize: 10, color: T.sub, background: T.bg, borderRadius: 6, padding: "2px 6px", fontWeight: 600 }}>
                      {o.items.length} шт
                    </div>
                  </div>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    <span style={{ fontSize: 11, color: T.sub }}>{fmt(rev)} ₽</span>
                    <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                      {margin !== null && (
                        <span style={{ fontSize: 10, color: profit >= 0 ? T.green : T.red, background: profit >= 0 ? T.greenBg : T.redBg, padding: "1px 5px", borderRadius: 5, fontWeight: 600 }}>
                          {fmtDec(margin)}%
                        </span>
                      )}
                      <b style={{ fontSize: 12, color: profit >= 0 ? T.green : T.red }}>{profit >= 0 ? "+" : ""}{fmt(profit)} ₽</b>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>

          {/* Bottom total */}
          {orders.length > 1 && (() => {
            const totals = orders.reduce((a, o) => { const c = orderCalc(o); return { rev: a.rev + c.rev, profit: a.profit + c.profit }; }, { rev: 0, profit: 0 });
            return (
              <div style={{ padding: "10px 16px", borderTop: `1px solid ${T.border}`, background: T.bg }}>
                <div style={{ fontSize: 10, color: T.sub, textTransform: "uppercase", letterSpacing: "0.5px", marginBottom: 4, fontWeight: 600 }}>Итого по всем заказам</div>
                <div style={{ display: "flex", justifyContent: "space-between" }}>
                  <span style={{ fontSize: 12, color: T.sub }}>Выручка: <b style={{ color: T.text }}>{fmt(totals.rev)} ₽</b></span>
                  <b style={{ fontSize: 12, color: totals.profit >= 0 ? T.green : T.red }}>{totals.profit >= 0 ? "+" : ""}{fmt(totals.profit)} ₽</b>
                </div>
              </div>
            );
          })()}
        </aside>
      )}

      {/* ── MAIN ────────────────────────────────────────────── */}
      {(!mobile || hashId) && (
        <main style={{ flex: 1, overflow: "hidden", display: "flex", flexDirection: "column" }}>
          {active
            ? <Detail order={active} mobile={mobile}
                onBack={goBack}
                onEdit={() => setModal({ type: "order", data: active })}
                onAddItem={() => setModal({ type: "item" })}
                onEditItem={(item) => setModal({ type: "item", data: item })}
                onDelItem={deleteItem}
                onExport={() => exportSalesTable(active)}
                onExportPhotos={() => exportPhotos(active)} />
            : <EmptyMain mobile={mobile} onBack={goBack} />
          }
        </main>
      )}

      {/* ── MODALS ──────────────────────────────────────────── */}
      {modal && (
        <Sheet mobile={mobile} onClose={() => setModal(null)}>
          {modal.type === "order" && (
            <OrderForm data={modal.data} num={orders.length + 1}
              onSave={modal.data ? (d) => updateOrder(modal.data.id, d) : createOrder}
              onDel={modal.data ? () => deleteOrder(modal.data.id) : null}
              onClose={() => setModal(null)} />
          )}
          {modal.type === "item" && active && (
            <ItemForm data={modal.data} order={active}
              onSave={modal.data ? (d) => updateItem(modal.data.id, d) : createItem}
              onClose={() => setModal(null)} />
          )}
        </Sheet>
      )}
    </div>
  );
}

/* ── Detail view ────────────────────────────────────────────── */
function Detail({ order, mobile, onBack, onEdit, onAddItem, onEditItem, onDelItem, onExport, onExportPhotos }) {
  const { cost, rev, profit } = orderCalc(order);
  const delRub = deliveryRub(order);
  const perItem = order.items.length > 0 && delRub > 0 ? delRub / order.items.length : 0;
  const margin = rev > 0 ? (profit / rev) * 100 : null;
  const hasDelivery = delRub > 0;

  const deliveryLabel = (() => {
    if (!hasDelivery) return "не указана";
    if (n(order.deliveryUsd) > 0 && n(order.usdRate) > 0)
      return `$${order.deliveryUsd} × ${order.usdRate} = ${fmt(delRub)} ₽`;
    return `${fmt(delRub)} ₽`;
  })();

  return <>
    {/* Header */}
    <div style={{ background: T.card, borderBottom: `1px solid ${T.border}`, padding: "11px 14px 10px" }}>
      <div style={{ display: "flex", alignItems: "center", gap: 4, marginBottom: 8 }}>
        {mobile && (
          <button onClick={onBack} style={{ ...$.ico, fontSize: 26, color: T.accent, marginLeft: -4 }}>‹</button>
        )}
        <h2 style={{ margin: 0, flex: 1, fontSize: 17, fontWeight: 800, letterSpacing: "-0.3px" }}>{order.name}</h2>
        <button onClick={onEdit} style={{ ...$.ico, fontSize: 15 }}>✏️</button>
      </div>
      {/* Params chips */}
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
        {[
          { l: "Курс ¥", v: order.rate ? `${order.rate} ₽/¥` : "—" },
          { l: "Доставка", v: deliveryLabel, warn: !hasDelivery },
          ...(perItem > 0 ? [{ l: "На шт", v: `${fmt(perItem)} ₽` }] : []),
          { l: "Позиций", v: totalUnits(order) !== order.items.length ? `${order.items.length} (${totalUnits(order)} шт)` : String(order.items.length) },
        ].map(({ l, v, warn }) => (
          <div key={l} style={{ fontSize: 11, background: warn ? T.goldBg : T.bg, border: `1px solid ${warn ? T.gold : T.border}`, borderRadius: 8, padding: "3px 8px", display: "flex", gap: 4 }}>
            <span style={{ color: warn ? T.gold : T.sub }}>{l}:</span>
            <b style={{ color: warn ? T.gold : T.text }}>{v}</b>
          </div>
        ))}
      </div>
    </div>

    {/* Items grid */}
    <div style={{ flex: 1, overflowY: "auto", padding: 14, WebkitOverflowScrolling: "touch" }}>
      {order.items.length === 0 ? (
        <div style={{ paddingTop: 64, textAlign: "center", color: T.sub }}>
          <div style={{ fontSize: 44, marginBottom: 12, opacity: 0.4 }}>🏷️</div>
          <div style={{ fontWeight: 700, color: T.subDark, marginBottom: 4 }}>Нет товаров</div>
          <div style={{ fontSize: 12 }}>Нажми «+ Добавить товар» ниже</div>
        </div>
      ) : (
        <div style={{ display: "grid", gap: 12, gridTemplateColumns: mobile ? "1fr" : "repeat(auto-fill,minmax(300px,1fr))" }}>
          {order.items.map((item) => (
            <ItemCard key={item.id} item={item} order={order}
              onEdit={() => onEditItem(item)}
              onDel={() => onDelItem(item.id)} />
          ))}
        </div>
      )}
    </div>

    {/* Per-label breakdown */}
    {(() => {
      const labeled = order.items.filter(i => i.label);
      if (!labeled.length) return null;
      const map = {};
      for (const item of order.items) {
        const key = item.label || "—";
        if (!map[key]) map[key] = 0;
        map[key] += itemCalc(item, order).profit;
      }
      return (
        <div style={{ background: T.bg, borderTop: `1px solid ${T.border}`, padding: "8px 14px", display: "flex", flexWrap: "wrap", gap: "4px 14px" }}>
          {Object.entries(map).map(([name, lp]) => {
            const lc = name === "—" ? null : labelColor(name);
            return (
              <span key={name} style={{ fontSize: 12, display: "flex", alignItems: "center", gap: 5 }}>
                {lc && <span style={{ width: 8, height: 8, borderRadius: "50%", background: lc.text, display: "inline-block" }} />}
                <span style={{ color: T.sub }}>{name}:</span>
                <b style={{ color: lp >= 0 ? T.green : T.red }}>{lp >= 0 ? "+" : ""}{fmt(lp)} ₽</b>
              </span>
            );
          })}
        </div>
      );
    })()}

    {/* Footer */}
    <div style={{ background: T.card, borderTop: `1px solid ${T.border}`, padding: "10px 14px" }}>
      <div style={{ display: "flex", alignItems: "center", flexWrap: "wrap", gap: 10 }}>
        <button onClick={onAddItem} style={$.btnRed}>+ Добавить товар</button>
        {order.items.length > 0 && (
          <button onClick={onExport} style={$.btnOutline}>📊 Таблица для продажи</button>
        )}
        {order.items.some(i => i.photo) && (
          <button onClick={onExportPhotos} style={$.btnOutline}>🖼 Фото для отправки</button>
        )}
        <div style={{ flex: 1 }} />
        <div style={{ display: "flex", gap: 14, fontSize: 12, flexWrap: "wrap", alignItems: "center" }}>
          <span style={{ color: T.sub }}>Закупка+дост.: <b style={{ color: T.text }}>{fmt(cost)} ₽</b></span>
          <span style={{ color: T.sub }}>Выручка: <b style={{ color: T.text }}>{fmt(rev)} ₽</b></span>
          <span style={{ display: "flex", alignItems: "center", gap: 6 }}>
            {margin !== null && (
              <span style={{ fontSize: 11, background: profit >= 0 ? T.greenBg : T.redBg, color: profit >= 0 ? T.green : T.red, borderRadius: 6, padding: "2px 6px", fontWeight: 700 }}>
                {fmtDec(margin)}%
              </span>
            )}
            <b style={{ color: profit >= 0 ? T.green : T.red, fontSize: 13 }}>{profit >= 0 ? "+" : ""}{fmt(profit)} ₽</b>
          </span>
        </div>
      </div>
    </div>
  </>;
}

/* ── ItemCard ────────────────────────────────────────────────── */
function ItemCard({ item, order, onEdit, onDel }) {
  const [delConf, setDc] = useState(false);
  const { rate, qty, rubPU, delPU, withDelPU, profitPU, profit } = itemCalc(item, order);
  const pos = profit >= 0;
  const sp = n(item.salePrice);
  const margin = sp > 0 ? (profitPU / sp) * 100 : null;

  const hasTrack = !!item.trackNum;
  const borderColor = hasTrack
    ? (item.arrived ? T.greenBdr : T.redBdr)
    : (pos ? T.greenBdr : T.redBdr);

  return (
    <div style={{ background: T.card, borderRadius: 16, overflow: "hidden", border: `1px solid ${borderColor}`, boxShadow: T.shadow }}>
      {/* Top row: photo + name + actions */}
      <div style={{ display: "flex", gap: 10, padding: "12px 12px 10px", alignItems: "flex-start" }}>
        {item.photo
          ? <img src={item.photo} alt="" style={{ width: 66, height: 66, objectFit: "cover", borderRadius: 10, flexShrink: 0 }} />
          : <div style={{ width: 66, height: 66, borderRadius: 10, background: T.bg, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 24, flexShrink: 0, color: T.sub }}>🏷️</div>
        }
        <div style={{ flex: 1, minWidth: 0 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 3 }}>
            <div style={{ fontWeight: 700, fontSize: 14, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", flex: 1 }}>
              {item.name || "Без названия"}
            </div>
            {item.label && (() => { const lc = labelColor(item.label); return (
              <span style={{ fontSize: 10, background: lc.bg, border: `1px solid ${lc.border}`, borderRadius: 6, padding: "1px 7px", fontWeight: 700, color: lc.text, flexShrink: 0 }}>
                {item.label}
              </span>
            ); })()}
            {qty > 1 && (
              <span style={{ fontSize: 11, background: T.bg, border: `1px solid ${T.border}`, borderRadius: 6, padding: "1px 6px", fontWeight: 700, color: T.subDark, flexShrink: 0 }}>
                ×{qty}
              </span>
            )}
          </div>
          <div style={{ fontSize: 11, color: T.sub, lineHeight: 1.5 }}>
            ¥{item.yuanPrice} × {rate} ₽/¥ = <b style={{ color: T.text }}>{fmt(rubPU)} ₽</b>
          </div>
          {item.rateOverride && n(item.rateOverride) !== n(order.rate) && (
            <div style={{ fontSize: 10, color: T.gold, background: T.goldBg, borderRadius: 5, padding: "1px 5px", display: "inline-block", marginTop: 2 }}>
              свой курс: {item.rateOverride}
            </div>
          )}
        </div>
        <div style={{ display: "flex", gap: 2, flexShrink: 0 }}>
          <button onClick={onEdit} style={$.ico}>✏️</button>
          {delConf
            ? <>
                <button onClick={onDel} style={{ ...$.ico, color: T.red, fontSize: 12, fontWeight: 800, background: T.redBg, borderRadius: 8, padding: "4px 8px" }}>Удалить</button>
                <button onClick={() => setDc(false)} style={{ ...$.ico, fontSize: 13, opacity: 0.6 }}>✕</button>
              </>
            : <button onClick={() => setDc(true)} style={{ ...$.ico, opacity: 0.5 }}>🗑️</button>
          }
        </div>
      </div>

      {/* Track number */}
      {hasTrack && (
        <div style={{ margin: "0 12px 10px", padding: "6px 10px", background: item.arrived ? T.greenBg : T.redBg, border: `1px solid ${item.arrived ? T.greenBdr : T.redBdr}`, borderRadius: 8, display: "flex", alignItems: "center", gap: 8 }}>
          <span style={{ fontSize: 13 }}>{item.arrived ? "✅" : "🚚"}</span>
          <div style={{ flex: 1, minWidth: 0 }}>
            <div style={{ fontSize: 10, color: T.sub, fontWeight: 600 }}>{item.arrived ? "ПРИБЫЛ НА СКЛАД" : "В ПУТИ"}</div>
            <div style={{ fontSize: 12, fontWeight: 600, color: T.text, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{item.trackNum}</div>
          </div>
        </div>
      )}

      {/* 2×2 numbers grid (per unit) */}
      <div style={{ borderTop: `1px solid ${T.border}`, display: "grid", gridTemplateColumns: "1fr 1fr" }}>
        {[
          [qty > 1 ? "Закупка/шт" : "Закупка ¥→₽", rubPU],
          ["Доставка/шт", delPU],
          ["С доставкой/шт", withDelPU],
          ["Цена продажи", sp],
        ].map(([label, val], i) => (
          <div key={label} style={{
            padding: "9px 12px",
            borderBottom: i < 2 ? `1px solid ${T.border}` : undefined,
            borderRight: i % 2 === 0 ? `1px solid ${T.border}` : undefined,
          }}>
            <div style={{ fontSize: 10, color: T.sub, marginBottom: 3, fontWeight: 500 }}>{label}</div>
            <div style={{ fontWeight: 700, fontSize: 13 }}>{fmt(val)} ₽</div>
          </div>
        ))}
      </div>

      {/* Profit banner */}
      <div style={{ padding: "10px 14px", background: pos ? T.greenBg : T.redBg, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div>
          <span style={{ fontSize: 12, color: T.sub, fontWeight: 600 }}>ПРИБЫЛЬ</span>
          {qty > 1 && (
            <div style={{ fontSize: 10, color: T.sub }}>
              {profitPU >= 0 ? "+" : ""}{fmt(profitPU)} ₽/шт × {qty}
            </div>
          )}
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          {margin !== null && (
            <span style={{ fontSize: 11, fontWeight: 700, color: pos ? T.green : T.red, opacity: 0.7 }}>
              {fmtDec(margin)}%
            </span>
          )}
          <b style={{ fontSize: 16, color: pos ? T.green : T.red, letterSpacing: "-0.3px" }}>
            {pos ? "+" : ""}{fmt(profit)} ₽
          </b>
        </div>
      </div>
    </div>
  );
}

/* ── OrderForm ──────────────────────────────────────────────── */
function OrderForm({ data, num, onSave, onDel, onClose }) {
  const isEdit = !!data;
  const [name, setName] = useState(data?.name ?? `Заказ ${num}`);
  const [rate, setRate] = useState(data?.rate ?? "12");
  const [deliveryUsd, setDelUsd] = useState(data?.deliveryUsd ?? "");
  const [usdRate, setUsdRate] = useState(data?.usdRate ?? "90");
  const [delConf, setDc] = useState(false);

  const delRub = n(deliveryUsd) * n(usdRate);

  return (
    <div>
      <FormHead title={isEdit ? "Редактировать заказ" : "Новый заказ"} onClose={onClose} />
      <Field label="Название заказа">
        <input value={name} onChange={(e) => setName(e.target.value)} style={$.input} placeholder="Заказ 1" autoFocus />
      </Field>
      <Field label="Курс юаня (₽ за ¥)">
        <input type="number" value={rate} onChange={(e) => setRate(e.target.value)} style={$.input} placeholder="12" step="0.01" min="0" inputMode="decimal" />
      </Field>

      <div style={{ fontSize: 12, color: T.sub, fontWeight: 700, marginBottom: 8, marginTop: 4 }}>
        Доставка из Китая <span style={{ fontWeight: 400, color: T.sub }}>(можно заполнить позже)</span>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 4 }}>
        <Field label="Стоимость ($)" style={{ marginBottom: 0 }}>
          <input type="number" value={deliveryUsd} onChange={(e) => setDelUsd(e.target.value)} style={$.input} placeholder="0" min="0" step="0.01" inputMode="decimal" />
        </Field>
        <Field label="Курс доллара (₽/$)" style={{ marginBottom: 0 }}>
          <input type="number" value={usdRate} onChange={(e) => setUsdRate(e.target.value)} style={$.input} placeholder="90" step="0.1" min="0" inputMode="decimal" />
        </Field>
      </div>

      {n(deliveryUsd) > 0 && (
        <div style={{ fontSize: 12, background: T.bg, border: `1px solid ${T.border}`, borderRadius: 8, padding: "8px 12px", marginBottom: 12, color: T.subDark }}>
          📦 ${deliveryUsd} × {usdRate} ₽/$ = <b style={{ color: T.text }}>{fmt(delRub)} ₽</b> — поровну на все товары
        </div>
      )}
      {!n(deliveryUsd) && (
        <div style={{ fontSize: 12, color: T.sub, padding: "6px 0 12px" }}>
          📦 Оставь пустым — добавишь доставку после отправки из Китая
        </div>
      )}

      <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
        <button onClick={() => onSave({ name, rate, deliveryUsd, usdRate })} style={{ ...$.btnRed, flex: 1 }}>
          {isEdit ? "Сохранить" : "Создать заказ"}
        </button>
        {isEdit && onDel && (
          delConf
            ? <>
                <button onClick={onDel} style={{ ...$.btnOutline, color: T.red, borderColor: T.redBdr, fontWeight: 700 }}>Точно?</button>
                <button onClick={() => setDc(false)} style={$.btnOutline}>✕</button>
              </>
            : <button onClick={() => setDc(true)} style={{ ...$.btnOutline, color: T.red }}>🗑️</button>
        )}
      </div>
    </div>
  );
}

/* ── ItemForm ───────────────────────────────────────────────── */
function ItemForm({ data, order, onSave, onClose }) {
  const isEdit = !!data?.id;
  const [name, setName] = useState(data?.name ?? "");
  const [yuanPrice, setYp] = useState(data?.yuanPrice ?? "");
  const [salePrice, setSp] = useState(data?.salePrice ?? "");
  const [rateOvr, setRo] = useState(data?.rateOverride ?? "");
  const [qty, setQty] = useState(String(data?.qty ?? "1"));
  const [label, setLabel] = useState(data?.label ?? "");
  const [trackNum, setTrack] = useState(data?.trackNum ?? "");
  const [arrived, setArrived] = useState(data?.arrived ?? false);
  const [photo, setPhoto] = useState(data?.photo ?? null);
  const fRef = useRef();

  const existingLabels = [...new Set(order.items.filter(i => i.label && i.id !== data?.id).map(i => i.label))];

  const qtyNum = Math.max(1, n(qty) || 1);
  const rate = n(rateOvr) || n(order.rate);
  const existingUnits = totalUnits(order);
  const previewUnits = Math.max(1, isEdit
    ? existingUnits - itemQty(data) + qtyNum
    : existingUnits + qtyNum);
  const rubPU = n(yuanPrice) * rate;
  const delPU = deliveryRub(order) / previewUnits;
  const withDelPU = rubPU + delPU;
  const profitPU = n(salePrice) - withDelPU;
  const totalProfit = profitPU * qtyNum;
  const hasCalc = n(yuanPrice) > 0;

  const pickFile = async (e) => {
    const f = e.target.files?.[0];
    if (f) setPhoto(await compressPhoto(f));
  };

  useEffect(() => {
    const onPaste = async (e) => {
      const file = [...(e.clipboardData?.items || [])].find(i => i.type.startsWith("image/"))?.getAsFile();
      if (file) setPhoto(await compressPhoto(file));
    };
    document.addEventListener("paste", onPaste);
    return () => document.removeEventListener("paste", onPaste);
  }, []);

  return (
    <div>
      <FormHead title={isEdit ? "Редактировать товар" : "Добавить товар"} onClose={onClose} />

      {/* Photo + name row */}
      <div style={{ display: "flex", gap: 12, marginBottom: 14, alignItems: "flex-start" }}>
        <input ref={fRef} type="file" accept="image/*" onChange={pickFile} style={{ display: "none" }} />
        {photo
          ? <div style={{ position: "relative", flexShrink: 0 }}>
              <img src={photo} alt="" style={{ width: 78, height: 78, objectFit: "cover", borderRadius: 12 }} />
              <button onClick={() => setPhoto(null)} style={{ position: "absolute", top: -7, right: -7, width: 22, height: 22, background: T.red, color: "#fff", border: "none", borderRadius: "50%", fontSize: 15, cursor: "pointer", fontWeight: 900, display: "flex", alignItems: "center", justifyContent: "center", lineHeight: 1 }}>×</button>
            </div>
          : <div onClick={() => fRef.current?.click()}
              style={{ width: 78, height: 78, flexShrink: 0, border: `2px dashed ${T.border}`, borderRadius: 12, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", cursor: "pointer", background: T.bg, gap: 3 }}>
              <span style={{ fontSize: 22, opacity: 0.5 }}>📷</span>
              <span style={{ fontSize: 10, color: T.sub, fontWeight: 600 }}>ФОТО</span>
            </div>
        }
        <Field label="Название товара" style={{ flex: 1, marginBottom: 0 }}>
          <input value={name} onChange={(e) => setName(e.target.value)} style={$.input} placeholder="Кроссовки, сумка…" />
        </Field>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 80px", gap: 10 }}>
        <Field label="Цена в юанях (¥)">
          <input type="number" value={yuanPrice} onChange={(e) => setYp(e.target.value)} style={$.input} placeholder="0" min="0" step="0.01" inputMode="decimal" />
        </Field>
        <Field label="Цена продажи (₽)">
          <input type="number" value={salePrice} onChange={(e) => setSp(e.target.value)} style={$.input} placeholder="0" min="0" inputMode="numeric" />
        </Field>
        <Field label="Кол-во">
          <input type="number" value={qty} onChange={(e) => setQty(e.target.value)} style={$.input} placeholder="1" min="1" step="1" inputMode="numeric" />
        </Field>
      </div>

      <Field label={`Курс ₽/¥ (по умолч. для заказа: ${order.rate || "—"})`}>
        <input type="number" value={rateOvr} onChange={(e) => setRo(e.target.value)} style={$.input} placeholder={order.rate || "12"} step="0.01" min="0" inputMode="decimal" />
      </Field>

      {/* Label / buyer */}
      <Field label="Покупатель (необязательно)">
        {existingLabels.length > 0 && (
          <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 8 }}>
            {existingLabels.map((l) => {
              const lc = labelColor(l);
              const sel = label === l;
              return (
                <button key={l} onClick={() => setLabel(sel ? "" : l)}
                  style={{ padding: "4px 12px", borderRadius: 20, fontSize: 12, fontWeight: 700, cursor: "pointer", fontFamily: "inherit", border: `1.5px solid ${sel ? lc.text : T.border}`, background: sel ? lc.bg : T.bg, color: sel ? lc.text : T.sub }}>
                  {l}
                </button>
              );
            })}
          </div>
        )}
        <input value={label} onChange={(e) => setLabel(e.target.value)} style={$.input} placeholder="Макс, Витя, Кирилл…" />
      </Field>

      {/* Track number */}
      <Field label="Трек-номер (необязательно)">
        <input value={trackNum} onChange={(e) => setTrack(e.target.value)} style={$.input} placeholder="RL123456789CN" />
      </Field>
      {trackNum && (
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginTop: -6, marginBottom: 12 }}>
          {[
            { label: "🚚 В пути", val: false },
            { label: "✅ На складе", val: true },
          ].map(({ label, val }) => (
            <button key={String(val)} onClick={() => setArrived(val)}
              style={{ padding: "9px", borderRadius: 10, fontSize: 13, fontWeight: 700, cursor: "pointer", border: `2px solid ${arrived === val ? (val ? T.green : T.red) : T.border}`, background: arrived === val ? (val ? T.greenBg : T.redBg) : T.bg, color: arrived === val ? (val ? T.green : T.red) : T.sub, fontFamily: "inherit" }}>
              {label}
            </button>
          ))}
        </div>
      )}

      {/* Live preview */}
      {hasCalc && (
        <div style={{ padding: "12px 14px", background: T.bg, borderRadius: 12, marginTop: 2, marginBottom: 4 }}>
          <div style={{ fontSize: 10, color: T.sub, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.6px", marginBottom: 10 }}>Предпросмотр расчёта</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
            {[
              ["¥ → ₽/шт", rubPU],
              ["Доставка/шт", delPU],
              ["С доставкой/шт", withDelPU],
            ].map(([l, v]) => (
              <div key={l}>
                <div style={{ fontSize: 10, color: T.sub, marginBottom: 2 }}>{l}</div>
                <div style={{ fontWeight: 700, fontSize: 13 }}>{fmt(v)} ₽</div>
              </div>
            ))}
            <div>
              <div style={{ fontSize: 10, color: T.sub, marginBottom: 2 }}>
                Прибыль{qtyNum > 1 ? ` ×${qtyNum}` : ""}
              </div>
              <div style={{ fontWeight: 800, fontSize: 14, color: totalProfit >= 0 ? T.green : T.red }}>
                {totalProfit >= 0 ? "+" : ""}{fmt(totalProfit)} ₽
              </div>
            </div>
          </div>
        </div>
      )}

      <button
        onClick={() => onSave({ name, yuanPrice, salePrice, rateOverride: rateOvr, qty: qtyNum, label, trackNum, arrived, photo })}
        style={{ ...$.btnRed, width: "100%", marginTop: 14, padding: "13px", borderRadius: 12, fontSize: 15, opacity: !yuanPrice ? 0.45 : 1, cursor: !yuanPrice ? "not-allowed" : "pointer" }}
        disabled={!yuanPrice}
      >
        {isEdit ? "Сохранить изменения" : "Добавить товар"}
      </button>
    </div>
  );
}

/* ── Shared micro components ────────────────────────────────── */
function Field({ label, children, style }) {
  return (
    <div style={{ marginBottom: 12, ...style }}>
      <div style={{ fontSize: 12, color: T.sub, fontWeight: 600, marginBottom: 5 }}>{label}</div>
      {children}
    </div>
  );
}

function FormHead({ title, onClose }) {
  return (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18 }}>
      <h3 style={{ margin: 0, fontSize: 17, fontWeight: 800 }}>{title}</h3>
      <button onClick={onClose} style={{ ...$.ico, fontSize: 20, opacity: 0.5 }}>✕</button>
    </div>
  );
}

function EmptyMain({ mobile, onBack }) {
  return (
    <div style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center", flexDirection: "column", color: T.sub, gap: 12, position: "relative" }}>
      {mobile && <button onClick={onBack} style={{ ...$.btnOutline, position: "absolute", top: 16, left: 14, fontSize: 13 }}>← Назад</button>}
      <div style={{ fontSize: 52, opacity: 0.3 }}>📦</div>
      <div style={{ fontWeight: 700, color: T.subDark, fontSize: 16 }}>Выбери заказ</div>
      <div style={{ fontSize: 13 }}>или создай новый в списке слева</div>
    </div>
  );
}

function Sheet({ children, mobile, onClose }) {
  return (
    <div
      style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.48)", zIndex: 200, display: "flex", alignItems: mobile ? "flex-end" : "center", justifyContent: "center" }}
      onClick={onClose}
    >
      <div
        style={{ background: T.card, borderRadius: mobile ? "22px 22px 0 0" : 16, padding: "20px 20px 28px", width: mobile ? "100%" : 440, maxWidth: "100%", maxHeight: mobile ? "93vh" : "88vh", overflowY: "auto", WebkitOverflowScrolling: "touch", boxSizing: "border-box" }}
        onClick={(e) => e.stopPropagation()}
      >
        {children}
      </div>
    </div>
  );
}
