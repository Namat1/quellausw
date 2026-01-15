# quell.py
# Streamlit: Excel Upload -> Standalone interaktive HTML (alles läuft in der HTML)
# Excel: Blatt "Direkt", Spalten A–L
# A CSB | B SAP | C Marktname | D Straße | E PLZ | F Ort | G–L Mo–Sa (Tournummern)
#
# Feiertagsregel:
# - Grundprinzip: Wenn ein Tag als Feiertag markiert ist -> keine Belieferung an diesem Tag.
#   Betroffene Kunden/Lieferungen werden PRINCIPIell vorher beliefert (rückwärts verschoben).
# - Ausnahme: Ist der Feiertag Montag -> wird auf Dienstag geschoben (vorwärts).
# - Zusätzlich: Mindestabstand je Markt (minGapDays) wird eingehalten, sonst Konflikt.

import json
from typing import Any, Dict, List

import pandas as pd
import streamlit as st


# ----------------------------
# Streamlit setup
# ----------------------------
st.set_page_config(page_title="Excel → Interaktive HTML", layout="centered")
st.title("Excel hochladen → Interaktive HTML erzeugen (Standalone)")

uploaded = st.file_uploader("Excel-Datei auswählen", type=["xlsx", "xlsm", "xls"])


# ----------------------------
# Helpers
# ----------------------------
def norm_str(x: Any) -> str:
    if pd.isna(x):
        return ""
    return str(x).strip()


def norm_tour(x: Any) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip()
    # häufig: 1201.0 -> 1201
    if s.endswith(".0"):
        s = s[:-2]
    return s


def build_data(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Liest ALLE Zeilen aus dem Excel-Blatt ein.
    Leere Zeilen werden übersprungen.
    """
    markets: List[Dict[str, Any]] = []

    if df.shape[1] < 12:
        raise ValueError("Excel-Blatt hat weniger als 12 Spalten (A–L).")

    for _, r in df.iterrows():
        csb = norm_str(r.iloc[0])
        sap = norm_str(r.iloc[1])
        name = norm_str(r.iloc[2])
        street = norm_str(r.iloc[3])
        zipc = norm_str(r.iloc[4])
        city = norm_str(r.iloc[5])

        if not (csb or sap or name):
            continue

        pattern = {
            "mo": norm_tour(r.iloc[6]),
            "di": norm_tour(r.iloc[7]),
            "mi": norm_tour(r.iloc[8]),
            "do": norm_tour(r.iloc[9]),
            "fr": norm_tour(r.iloc[10]),
            "sa": norm_tour(r.iloc[11]),
        }

        markets.append(
            {
                "csb": csb,
                "sap": sap,
                "name": name,
                "street": street,
                "zip": zipc,
                "city": city,
                "pattern": pattern,
            }
        )

    return {
        "meta": {
            # feste Logik: Woche beginnt Sonntag (So–Sa)
            "weekStartsSunday": True,
            "minGapDays": 3,
        },
        "markets": markets,
    }


HTML_TEMPLATE = r"""<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Belieferungsschema – Interaktiv</title>
<style>
  body { font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial; margin: 16px; background:#f4f5f7; }
  .wrap { max-width: 1400px; margin: 0 auto; }
  .card { background:#fff; border:1px solid #ddd; border-radius:14px; padding:14px; box-shadow: 0 2px 10px rgba(0,0,0,.04); }
  .row { display:flex; gap:12px; flex-wrap:wrap; align-items:center; }
  .grow { flex:1; }
  input, select, button { padding:10px 12px; border-radius:10px; border:1px solid #ccc; background:#fff; }
  button { cursor:pointer; }
  .muted { color:#666; font-size: 13px; }
  .h2 { font-size:18px; margin: 8px 0; }
  .pill { display:inline-block; padding:4px 10px; border-radius:999px; border:1px solid #ddd; background:#fafafa; font-size:12px; }
  .tag { display:inline-flex; align-items:center; gap:6px; padding:6px 10px; border-radius:999px; border:1px solid #ddd; background:#fff; font-size:12px; }
  .tag input { margin:0; }
  .hr { height:1px; background:#eee; margin:10px 0; }
  .bad { color:#b00; }
  .ok { color:#0a6; }
  .small { font-size:12px; }

  /* Feiertage Buttons */
  .grid7 { display:grid; grid-template-columns: repeat(7, 1fr); gap:10px; }
  .daybtn { padding:10px; border-radius:12px; border:1px solid #ccc; background:#fff; text-align:center; user-select:none; cursor:pointer; }
  .holiday { border-color:#d33; background: #ffecec; }

  /* Matrix */
  .matrixWrap { overflow:auto; max-height: 72vh; border:1px solid #e3e3e3; border-radius:12px; background:#fff; }
  table.matrix { border-collapse: separate; border-spacing:0; width: 100%; font-size: 13px; }
  table.matrix th, table.matrix td { padding:8px 10px; border-bottom:1px solid #eee; border-right:1px solid #f0f0f0; white-space:nowrap; vertical-align:top; }
  table.matrix th { position: sticky; top: 0; background: #fafafa; z-index: 3; }
  table.matrix td.market { position: sticky; left: 0; background:#fff; z-index: 2; border-right:1px solid #e6e6e6; min-width: 280px; }
  table.matrix th.marketH { position: sticky; left:0; z-index: 4; background:#fafafa; border-right:1px solid #e6e6e6; min-width: 280px; }
  .tourCell { display:flex; gap:6px; align-items:center; justify-content:space-between; }
  .tourNum { font-weight:800; }
  .empty { color:#bbb; }
  .holidayCell { background:#ffecec; }
  .movedIn { background:#eafff0; }
  .badge { font-size:11px; padding:2px 8px; border-radius:999px; border:1px solid #ddd; background:#fff; }

  /* rechte Seite */
  .split { display:grid; grid-template-columns: 1.25fr .75fr; gap:12px; }
  @media (max-width: 900px) { .split { grid-template-columns: 1fr; } }

  .box { border:1px solid #e3e3e3; border-radius:12px; padding:10px; background:#fff; }
</style>
</head>
<body>
<div class="wrap">

  <div class="card">
    <div class="row">
      <div class="grow">
        <div class="h2">Belieferungsschema – Übersicht (Alles auf einen Blick)</div>
        <div class="muted">Datum wählen → KW wird angezeigt → Feiertage anklicken → Matrix aktualisiert sich sofort.</div>
      </div>

      <div>
        <label class="muted">Datum in KW</label><br/>
        <input id="datePick" type="date"/>
      </div>

      <div>
        <label class="muted">Suche</label><br/>
        <input id="q" placeholder="Markt / Ort / CSB / SAP / Tour…"/>
      </div>

      <div>
        <label class="muted">Ansicht</label><br/>
        <select id="view">
          <option value="matrix" selected>Matrix (Übersicht)</option>
          <option value="conflicts">Konflikte</option>
        </select>
      </div>

      <div>
        <label class="muted">Feiertage</label><br/>
        <button id="clearH">Feiertage löschen</button>
      </div>
    </div>

    <div style="margin-top:12px" class="row">
      <span id="kwLabel" class="pill"></span>
      <span class="pill">Mindestabstand: <b id="gapLabel"></b> Tage</span>
      <span class="pill">Woche: <b id="rangeLabel"></b></span>

      <span class="tag">
        <input type="checkbox" id="modeTourTogether"/>
        <label for="modeTourTogether">Touren zusammenhalten</label>
      </span>
    </div>

    <div class="muted small" style="margin-top:8px">
      Regeln: Feiertag = keine Lieferung. Normal: vorher liefern (rückwärts). Ausnahme: Feiertag Montag → auf Dienstag schieben.
      Farben: <span class="pill">Feiertag = rot</span> <span class="pill">verschoben = grün</span>
    </div>
  </div>

  <div style="height:12px"></div>

  <div class="card">
    <div class="h2">Feiertage in dieser Woche</div>
    <div id="weekDays" class="grid7"></div>
    <div class="muted" style="margin-top:8px">
      Klick auf Tag = Feiertag an/aus (nur für die aktuell gewählte KW).
    </div>
  </div>

  <div style="height:12px"></div>

  <div class="split">
    <div class="card">
      <div class="h2" id="leftTitle">Matrix</div>
      <div id="left"></div>
    </div>

    <div class="card">
      <div class="h2">Zusammenfassung</div>
      <div id="summary" class="muted"></div>
      <div class="hr"></div>
      <div class="box">
        <div class="muted small">
          <b>Hinweis:</b><br/>
          Diese Version verschiebt innerhalb der KW. (Vorwoche ist NICHT erlaubt.)<br/>
          Wenn Mindestabstand je Markt nicht einhaltbar ist → Konflikt.
        </div>
      </div>
    </div>
  </div>

</div>

<script>
// --------- embedded data ----------
const DATA = __DATA__;

// --------- config ----------
const weekStartsSunday = !!(DATA.meta && DATA.meta.weekStartsSunday);
const minGapDays = Number((DATA.meta && DATA.meta.minGapDays) || 3);

// --------- state ----------
const state = {
  date: null,
  holidays: new Set(),   // ISO date strings der aktuellen Woche
  view: "matrix",
  q: "",
  tourTogether: false,
};

const el = (id) => document.getElementById(id);

function pad(n){ return String(n).padStart(2,'0'); }
function iso(d){ return d.getFullYear()+"-"+pad(d.getMonth()+1)+"-"+pad(d.getDate()); }
function parseISO(s){
  const [y,m,da] = s.split("-").map(Number);
  return new Date(y, m-1, da);
}
function addDays(d, n){
  const x = new Date(d);
  x.setDate(x.getDate() + n);
  return new Date(x.getFullYear(), x.getMonth(), x.getDate());
}
function weekdayName(d){
  return ["So","Mo","Di","Mi","Do","Fr","Sa"][d.getDay()];
}
function weekRange(d){
  const dow = d.getDay();
  let start;
  if (weekStartsSunday){
    start = addDays(d, -dow);
  } else {
    const diff = (dow === 0) ? 6 : (dow - 1);
    start = addDays(d, -diff);
  }
  const end = addDays(start, 6);
  return {start, end};
}
function isoWeekNumber(date){
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(),0,1));
  const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return {year: d.getUTCFullYear(), week: weekNo};
}
function daterange(start, end){
  const out = [];
  let cur = new Date(start);
  while (cur <= end){
    out.push(cur);
    cur = addDays(cur, 1);
  }
  return out;
}
function getPatternForDow(market, dowJS){
  // Muster aus Excel: Mo–Sa; Sonntag = kein Plan
  const map = {1:"mo",2:"di",3:"mi",4:"do",5:"fr",6:"sa"};
  const k = map[dowJS];
  return k ? (market.pattern[k] || "") : "";
}
function diffDays(aISO, bISO){
  const a = parseISO(aISO);
  const b = parseISO(bISO);
  return Math.round((a - b) / 86400000);
}

function buildWeekDaysUI(){
  const wr = weekRange(state.date);
  const days = daterange(wr.start, wr.end);
  const cont = el("weekDays");
  cont.innerHTML = "";

  for (let i=0;i<days.length;i++){
    const day = days[i];
    const key = iso(day);
    const btn = document.createElement("div");
    btn.className = "daybtn" + (state.holidays.has(key) ? " holiday" : "");
    btn.innerHTML = `
      <div><b>${weekdayName(day)}</b></div>
      <div class="small">${pad(day.getDate())}.${pad(day.getMonth()+1)}</div>
    `;
    btn.onclick = () => {
      if (state.holidays.has(key)) state.holidays.delete(key);
      else state.holidays.add(key);
      render();
    };
    cont.appendChild(btn);
  }
}

function marketIdInit(){
  (DATA.markets || []).forEach((m, idx) => { m._id = idx; });
}

function planForWeek(){
  const wr = weekRange(state.date);
  const days = daterange(wr.start, wr.end);

  // deliveries: Map(dateISO -> Array of items)
  // item: { market, tour, originalDate }
  const deliveries = new Map();
  days.forEach(d => deliveries.set(iso(d), []));

  // Rohplan aus Muster
  for (const m of (DATA.markets || [])){
    for (const d of days){
      const tour = getPatternForDow(m, d.getDay());
      if (tour){
        deliveries.get(iso(d)).push({market:m, tour:tour, originalDate: iso(d)});
      }
    }
  }

  const moved = [];      // {from,to, market, tour}
  const conflicts = [];  // {type, msg, market, tour, from}

  function getExistingDatesForMarket(mid){
    const out = [];
    for (const [dk, arr] of deliveries.entries()){
      for (const it of arr){
        if (it.market._id === mid) out.push(dk);
      }
    }
    out.sort();
    return out;
  }

  function canPlace(market, targetISO){
    const existing = getExistingDatesForMarket(market._id);
    for (const ed of existing){
      const gap = Math.abs(diffDays(targetISO, ed));
      if (gap < minGapDays) return false;
    }
    return true;
  }

  // Zieltag finden nach deiner Regel:
  // - Montag-Feiertag: vorwärts (Di, Mi, ...)
  // - sonst: rückwärts (vorher liefern)
  function findTargetDateISO(baseDate, direction, marketOrItems){
    let target = addDays(baseDate, direction);

    while (target >= wr.start && target <= wr.end){
      const tISO = iso(target);

      // nicht auf Feiertag
      if (state.holidays.has(tISO)) { target = addDays(target, direction); continue; }

      // Gap-Check
      if (Array.isArray(marketOrItems)){
        let ok = true;
        for (const it of marketOrItems){
          if (!canPlace(it.market, tISO)) { ok = false; break; }
        }
        if (ok) return tISO;
      } else {
        if (canPlace(marketOrItems.market, tISO)) return tISO;
      }

      target = addDays(target, direction);
    }

    return null;
  }

  // Feiertage verschieben
  for (const d of days){
    const dayISO = iso(d);
    if (!state.holidays.has(dayISO)) continue;

    const items = deliveries.get(dayISO) || [];
    if (!items.length) continue;

    // Feiertag: keine Lieferung am Tag selbst
    deliveries.set(dayISO, []);

    const isMondayHoliday = (d.getDay() === 1);  // Mo
    const dir = isMondayHoliday ? +1 : -1;       // Mo -> vorwärts, sonst rückwärts

    if (state.tourTogether){
      // Touren gruppieren
      const groups = new Map();
      for (const it of items){
        if (!groups.has(it.tour)) groups.set(it.tour, []);
        groups.get(it.tour).push(it);
      }

      for (const [tour, gitems] of groups.entries()){
        const targetISO = findTargetDateISO(d, dir, gitems);

        if (!targetISO){
          for (const it of gitems){
            conflicts.push({
              type: "GAP_OR_RANGE",
              msg: `Kann ${it.market.name} (${it.market.city}) von ${dayISO} nicht verschieben (Tour ${tour}). Regel: ${isMondayHoliday ? "Mo → Di" : "vorher"}. Mindestabstand: ${minGapDays} Tage.`,
              market: it.market, tour: it.tour, from: dayISO
            });
          }
          continue;
        }

        for (const it of gitems){
          deliveries.get(targetISO).push(it);
          moved.push({from: dayISO, to: targetISO, market: it.market, tour: it.tour});
        }
      }
    } else {
      // itemweise
      for (const it of items){
        const targetISO = findTargetDateISO(d, dir, it);

        if (!targetISO){
          conflicts.push({
            type: "GAP_OR_RANGE",
            msg: `Kann ${it.market.name} (${it.market.city}) von ${dayISO} nicht verschieben. Regel: ${isMondayHoliday ? "Mo → Di" : "vorher"}. Mindestabstand: ${minGapDays} Tage.`,
            market: it.market, tour: it.tour, from: dayISO
          });
          continue;
        }

        deliveries.get(targetISO).push(it);
        moved.push({from: dayISO, to: targetISO, market: it.market, tour: it.tour});
      }
    }
  }

  return {wr, days, deliveries, moved, conflicts};
}

function renderSummary(plan){
  let totalStops = 0;
  for (const [, arr] of plan.deliveries.entries()) totalStops += arr.length;

  el("summary").innerHTML = `
    <div>Märkte gesamt: <b>${(DATA.markets||[]).length}</b></div>
    <div>Stops diese KW: <b>${totalStops}</b></div>
    <div>Feiertage markiert: <b>${state.holidays.size}</b></div>
    <div>Verschoben: <b>${plan.moved.length}</b></div>
    <div>Konflikte: <b class="${plan.conflicts.length ? "bad":"ok"}">${plan.conflicts.length}</b></div>
  `;
}

function renderConflicts(plan, q){
  const root = el("left");
  root.innerHTML = "";

  if (!plan.conflicts.length){
    root.innerHTML = `<div class="muted">Keine Konflikte 🎉</div>`;
    return;
  }

  for (const c of plan.conflicts){
    const hay = (c.market.name+" "+c.market.city+" "+c.market.csb+" "+c.market.sap+" "+c.tour).toLowerCase();
    if (q && !hay.includes(q)) continue;

    const div = document.createElement("div");
    div.className = "box";
    div.innerHTML = `
      <div><b class="bad">Konflikt</b> – Tour <b>${c.tour}</b></div>
      <div class="muted small">${c.from}</div>
      <div class="muted">${c.msg}</div>
    `;
    root.appendChild(div);
  }
}

function renderMatrix(plan, q){
  const root = el("left");
  root.innerHTML = "";

  // movedIn[dateISO][marketId] = fromDateISO
  const movedIn = new Map();
  for (const mv of plan.moved){
    if (!movedIn.has(mv.to)) movedIn.set(mv.to, new Map());
    movedIn.get(mv.to).set(mv.market._id, mv.from);
  }

  const wrap = document.createElement("div");
  wrap.className = "matrixWrap";

  const t = document.createElement("table");
  t.className = "matrix";

  const days = plan.days;

  // Header
  const thead = document.createElement("thead");
  const hr = document.createElement("tr");

  const th0 = document.createElement("th");
  th0.className = "marketH";
  th0.textContent = "Markt";
  hr.appendChild(th0);

  for (const d of days){
    const dk = iso(d);
    const th = document.createElement("th");
    th.textContent = `${weekdayName(d)} ${d.toLocaleDateString('de-DE')}${state.holidays.has(dk) ? " (FT)" : ""}`;
    hr.appendChild(th);
  }

  thead.appendChild(hr);
  t.appendChild(thead);

  // Precompute day->market->tour
  const dayMarketTour = new Map();
  for (const d of days){
    const dk = iso(d);
    const map = new Map();
    const arr = plan.deliveries.get(dk) || [];
    for (const it of arr){
      map.set(it.market._id, it.tour);
    }
    dayMarketTour.set(dk, map);
  }

  // Markets filter
  const markets = (DATA.markets || []).filter(m => {
    if (!q) return true;
    const hay = (m.name+" "+m.city+" "+m.csb+" "+m.sap).toLowerCase();
    return hay.includes(q);
  });

  const tbody = document.createElement("tbody");

  for (const m of markets){
    const tr = document.createElement("tr");

    const tdM = document.createElement("td");
    tdM.className = "market";
    tdM.innerHTML = `<div><b>${m.name}</b></div><div class="muted small">${m.city} · CSB ${m.csb} · SAP ${m.sap}</div>`;
    tr.appendChild(tdM);

    for (const d of days){
      const dk = iso(d);
      const td = document.createElement("td");

      if (state.holidays.has(dk)) td.classList.add("holidayCell");

      const tour = (dayMarketTour.get(dk) || new Map()).get(m._id) || "";
      const movedFrom = movedIn.get(dk)?.get(m._id);

      if (movedFrom) td.classList.add("movedIn");

      if (!tour){
        td.innerHTML = `<span class="empty">–</span>`;
      } else {
        td.innerHTML = `
          <div class="tourCell">
            <span class="tourNum">${tour}</span>
            ${movedFrom ? `<span class="badge">${movedFrom.slice(0,10)} →</span>` : ``}
          </div>
        `;
      }

      tr.appendChild(td);
    }

    tbody.appendChild(tr);
  }

  t.appendChild(tbody);
  wrap.appendChild(t);
  root.appendChild(wrap);
}

function render(){
  const wn = isoWeekNumber(state.date);
  const wr = weekRange(state.date);

  el("kwLabel").textContent = `KW ${wn.week} / ${wn.year}`;
  el("gapLabel").textContent = String(minGapDays);
  el("rangeLabel").textContent = `${wr.start.toLocaleDateString('de-DE')} – ${wr.end.toLocaleDateString('de-DE')}`;

  buildWeekDaysUI();

  const plan = planForWeek();
  renderSummary(plan);

  const q = state.q.trim().toLowerCase();

  el("leftTitle").textContent = (state.view === "conflicts") ? "Konflikte" : "Matrix (Übersicht)";

  if (state.view === "conflicts"){
    renderConflicts(plan, q);
  } else {
    renderMatrix(plan, q);
  }
}

function init(){
  marketIdInit();

  const today = new Date();
  state.date = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  el("datePick").value = iso(state.date);

  el("datePick").addEventListener("change", (e) => {
    state.date = parseISO(e.target.value);
    state.holidays.clear();
    render();
  });

  el("view").addEventListener("change", (e) => {
    state.view = e.target.value;
    render();
  });

  el("q").addEventListener("input", (e) => {
    state.q = e.target.value;
    render();
  });

  el("clearH").addEventListener("click", () => {
    state.holidays.clear();
    render();
  });

  el("modeTourTogether").addEventListener("change", (e) => {
    state.tourTogether = !!e.target.checked;
    render();
  });

  render();
}

init();
</script>
</body>
</html>
"""


def render_html(data: Dict[str, Any]) -> str:
    payload_json = json.dumps(data, ensure_ascii=False)
    # WICHTIG: kein f-string -> keine {} Probleme
    return HTML_TEMPLATE.replace("__DATA__", payload_json)


# ----------------------------
# Main
# ----------------------------
if uploaded:
    try:
        df = pd.read_excel(uploaded, sheet_name="Direkt 1 - 99", header=None)
    except Exception as e:
        st.error(f"Excel konnte nicht gelesen werden: {e}")
        st.stop()

    try:
        data = build_data(df)
    except Exception as e:
        st.error(f"Fehler beim Verarbeiten: {e}")
        st.stop()

    html = render_html(data)

    st.success(f"{len(data['markets'])} Märkte geladen. HTML bereit.")
    st.download_button(
        "Interaktive HTML herunterladen",
        data=html.encode("utf-8"),
        file_name="belieferung_interaktiv.html",
        mime="text/html",
    )
else:
    st.info("Bitte Excel hochladen.")
