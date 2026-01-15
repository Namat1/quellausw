# quell.py
# Streamlit: Excel Upload -> Standalone interaktive HTML (alles läuft in der HTML)
# Excel: Blatt "Direkt", Zeilen 1–99, Spalten A–L
# A CSB | B SAP | C Marktname | D Straße | E PLZ | F Ort | G–L Mo–Sa (Tournummern)

import json
from datetime import date
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
    s = str(x).strip()
    return s


def norm_tour(x: Any) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip()
    # Häufig: 1201.0 -> 1201
    if s.endswith(".0"):
        s = s[:-2]
    return s


def build_data(df: pd.DataFrame, *, max_rows: int = 99) -> Dict[str, Any]:
    # Zeilen 1–99 laut User (wir interpretieren: erste 99 Datenzeilen im Blatt)
    df = df.iloc[0:max_rows].copy()

    markets: List[Dict[str, Any]] = []
    for _, r in df.iterrows():
        # Spalten A–L => iloc 0..11
        csb = norm_str(r.iloc[0])
        sap = norm_str(r.iloc[1])
        name = norm_str(r.iloc[2])
        street = norm_str(r.iloc[3])
        zipc = norm_str(r.iloc[4])
        city = norm_str(r.iloc[5])

        # Leere Zeilen überspringen
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
            # Wenn du es fest willst: True (So–Sa)
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
  .wrap { max-width: 1200px; margin: 0 auto; }
  .card { background:#fff; border:1px solid #ddd; border-radius:14px; padding:14px; box-shadow: 0 2px 10px rgba(0,0,0,.04); }
  .row { display:flex; gap:12px; flex-wrap:wrap; align-items:center; }
  .grow { flex:1; }
  input, select, button { padding:10px 12px; border-radius:10px; border:1px solid #ccc; background:#fff; }
  button { cursor:pointer; }
  .grid7 { display:grid; grid-template-columns: repeat(7, 1fr); gap:10px; }
  .daybtn { padding:10px; border-radius:12px; border:1px solid #ccc; background:#fff; text-align:center; user-select:none; cursor:pointer; }
  .holiday { border-color:#d33; background: #ffecec; }
  .muted { color:#666; font-size: 13px; }
  .h2 { font-size:18px; margin: 8px 0; }
  .pill { display:inline-block; padding:4px 10px; border-radius:999px; border:1px solid #ddd; background:#fafafa; font-size:12px; }
  .split { display:grid; grid-template-columns: 1.1fr .9fr; gap:12px; }
  @media (max-width: 900px) { .split { grid-template-columns: 1fr; } }
  .list { display:flex; flex-direction:column; gap:10px; }
  .box { border:1px solid #e3e3e3; border-radius:12px; padding:10px; }
  .boxhead { display:flex; justify-content:space-between; align-items:center; gap:10px; }
  .bad { color:#b00; }
  .ok { color:#0a6; }
  .small { font-size:12px; }
  .tag { display:inline-flex; align-items:center; gap:6px; padding:6px 10px; border-radius:999px; border:1px solid #ddd; background:#fff; font-size:12px; }
  .tag input { margin:0; }
  .hr { height:1px; background:#eee; margin:10px 0; }
</style>
</head>
<body>
<div class="wrap">
  <div class="card">
    <div class="row">
      <div class="grow">
        <div class="h2">Belieferungsschema (Standalone HTML)</div>
        <div class="muted">KW wählen, Feiertage anklicken – Plan wird direkt neu berechnet (ohne Streamlit).</div>
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
          <option value="tour">Touren</option>
          <option value="market">Märkte</option>
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
        <label for="modeTourTogether">Touren zusammenhalten (weniger Splits)</label>
      </span>
    </div>

    <div class="muted small" style="margin-top:8px">
      Modus-Hinweis: „Touren zusammenhalten“ verschiebt am Feiertag möglichst komplette Tourgruppen rückwärts.
      Wenn Abstände &ge; Mindestabstand nicht einhaltbar sind, wird ein Konflikt erzeugt.
    </div>
  </div>

  <div style="height:12px"></div>

  <div class="card">
    <div class="h2">Feiertage in dieser Woche</div>
    <div id="weekDays" class="grid7"></div>
    <div class="muted" style="margin-top:8px">
      Klick auf Tag = Feiertag an/aus (nur für die aktuell gewählte Woche).
    </div>
  </div>

  <div style="height:12px"></div>

  <div class="split">
    <div class="card">
      <div class="h2" id="leftTitle">Ergebnis</div>
      <div id="left" class="list"></div>
    </div>
    <div class="card">
      <div class="h2">Zusammenfassung</div>
      <div id="summary" class="muted"></div>
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
  view: "tour",
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
  // normalisieren auf 00:00
  return new Date(x.getFullYear(), x.getMonth(), x.getDate());
}
function weekdayName(d){
  return ["So","Mo","Di","Mi","Do","Fr","Sa"][d.getDay()];
}
function weekRange(d){
  // JS getDay(): So=0..Sa=6
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

// ISO week number (Anzeige)
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
  // stabile IDs vergeben
  (DATA.markets || []).forEach((m, idx) => { m._id = idx; });
}

function planForWeek(){
  const wr = weekRange(state.date);
  const days = daterange(wr.start, wr.end);

  // deliveries: Map(dateISO -> Array of items)
  // item: { market, tour, originalDate }
  const deliveries = new Map();
  days.forEach(d => deliveries.set(iso(d), []));

  // Rohplan
  for (const m of (DATA.markets || [])){
    for (const d of days){
      const tour = getPatternForDow(m, d.getDay());
      if (tour){
        deliveries.get(iso(d)).push({market:m, tour:tour, originalDate: iso(d)});
      }
    }
  }

  // moved: {from,to, market, tour}
  const moved = [];
  // conflicts: {type, msg, market, tour, from}
  const conflicts = [];

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

  // Verschiebe-Strategie:
  // - Wenn "tourTogether": für jeden Feiertag gruppieren wir pro Tour und versuchen pro Tour gemeinsam zu schieben.
  // - Sonst: itemweise.
  for (const d of days){
    const dayISO = iso(d);
    if (!state.holidays.has(dayISO)) continue;

    const items = deliveries.get(dayISO) || [];
    if (!items.length) continue;

    // Feiertag: erst entfernen
    deliveries.set(dayISO, []);

    if (state.tourTogether){
      // gruppiere nach Tour
      const groups = new Map();
      for (const it of items){
        if (!groups.has(it.tour)) groups.set(it.tour, []);
        groups.get(it.tour).push(it);
      }

      for (const [tour, gitems] of groups.entries()){
        let target = addDays(d, -1);
        let targetISO = null;

        while (target >= wr.start){
          const tISO = iso(target);
          if (state.holidays.has(tISO)) { target = addDays(target, -1); continue; }

          // Tour zusammenhalten: alle Items müssen platzierbar sein
          let ok = true;
          for (const it of gitems){
            if (!canPlace(it.market, tISO)){
              ok = false; break;
            }
          }
          if (ok){
            targetISO = tISO; break;
          }
          target = addDays(target, -1);
        }

        if (!targetISO){
          for (const it of gitems){
            conflicts.push({
              type: "GAP_OR_RANGE",
              msg: `Kann ${it.market.name} (${it.market.city}) von ${dayISO} nicht nach vorne ziehen (Tour ${tour}) ohne Abstand < ${minGapDays} Tage oder außerhalb der KW.`,
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
        let target = addDays(d, -1);
        let targetISO = null;

        while (target >= wr.start){
          const tISO = iso(target);
          if (state.holidays.has(tISO)) { target = addDays(target, -1); continue; }

          if (canPlace(it.market, tISO)){
            targetISO = tISO; break;
          }
          target = addDays(target, -1);
        }

        if (!targetISO){
          conflicts.push({
            type: "GAP_OR_RANGE",
            msg: `Kann ${it.market.name} (${it.market.city}) von ${dayISO} nicht nach vorne ziehen ohne Abstand < ${minGapDays} Tage oder außerhalb der KW.`,
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
    <div class="hr"></div>
    <div class="small muted">
      Logik: Feiertage werden rückwärts verschoben („davor liefern“). Mindestabstand je Markt: ${minGapDays} Tage.
      Wenn nicht möglich innerhalb dieser KW ⇒ Konflikt.
    </div>
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
      <div class="boxhead">
        <div><b class="bad">Konflikt</b> – Tour ${c.tour}</div>
        <div class="pill">${c.from}</div>
      </div>
      <div class="muted">${c.msg}</div>
    `;
    root.appendChild(div);
  }
}

function renderMarkets(q){
  const root = el("left");
  root.innerHTML = "";

  const ms = (DATA.markets || []).filter(m => {
    if (!q) return true;
    const hay = (m.name+" "+m.city+" "+m.csb+" "+m.sap).toLowerCase();
    return hay.includes(q);
  });

  for (const m of ms){
    const p = m.pattern || {};
    const div = document.createElement("div");
    div.className = "box";
    div.innerHTML = `
      <div class="boxhead">
        <div><b>${m.name}</b> <span class="muted">(${m.city})</span></div>
        <div class="pill">CSB ${m.csb} · SAP ${m.sap}</div>
      </div>
      <div class="muted">${m.street}, ${m.zip} ${m.city}</div>
      <div style="margin-top:8px" class="row">
        <span class="pill">Mo: <b>${p.mo || "-"}</b></span>
        <span class="pill">Di: <b>${p.di || "-"}</b></span>
        <span class="pill">Mi: <b>${p.mi || "-"}</b></span>
        <span class="pill">Do: <b>${p.do || "-"}</b></span>
        <span class="pill">Fr: <b>${p.fr || "-"}</b></span>
        <span class="pill">Sa: <b>${p.sa || "-"}</b></span>
      </div>
    `;
    root.appendChild(div);
  }
}

function renderTours(plan, q){
  const root = el("left");
  root.innerHTML = "";

  for (const d of plan.days){
    const dk = iso(d);
    const arr = plan.deliveries.get(dk) || [];

    const filtered = arr.filter(it => {
      if (!q) return true;
      const hay = (it.market.name+" "+it.market.city+" "+it.market.csb+" "+it.market.sap+" "+it.tour).toLowerCase();
      return hay.includes(q);
    });

    // group by tour
    const byTour = new Map();
    for (const it of filtered){
      if (!byTour.has(it.tour)) byTour.set(it.tour, []);
      byTour.get(it.tour).push(it);
    }

    const dayBox = document.createElement("div");
    dayBox.className = "box";
    dayBox.innerHTML = `
      <div class="boxhead">
        <div><b>${weekdayName(d)} ${d.toLocaleDateString('de-DE')}</b>${state.holidays.has(dk) ? ' <span class="bad">(Feiertag)</span>' : ''}</div>
        <div class="pill">Stops: ${filtered.length}</div>
      </div>
      <div class="muted small">Touren: ${byTour.size}</div>
      <div class="list" style="margin-top:10px" id="inner_${dk.replaceAll("-","_")}"></div>
    `;
    root.appendChild(dayBox);

    const inner = dayBox.querySelector("#inner_"+dk.replaceAll("-","_"));
    for (const [tour, items] of [...byTour.entries()].sort((a,b)=>String(a[0]).localeCompare(String(b[0])))){
      const tdiv = document.createElement("div");
      tdiv.className = "box";
      tdiv.innerHTML = `
        <div class="boxhead">
          <div><b>Tour ${tour}</b></div>
          <div class="pill">${items.length} Märkte</div>
        </div>
        <div class="muted small">${items.map(x => x.market.name + " ("+x.market.city+")").join(" · ")}</div>
      `;
      inner.appendChild(tdiv);
    }
  }
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

  el("leftTitle").textContent =
    state.view === "tour" ? "Tourenansicht" :
    state.view === "market" ? "Marktansicht" : "Konflikte";

  if (state.view === "conflicts"){
    renderConflicts(plan, q);
  } else if (state.view === "market"){
    renderMarkets(q);
  } else {
    renderTours(plan, q);
  }
}

function init(){
  marketIdInit();

  const today = new Date();
  state.date = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  el("datePick").value = iso(state.date);

  el("datePick").addEventListener("change", (e) => {
    state.date = parseISO(e.target.value);
    state.holidays.clear(); // Feiertage pro KW neu setzen
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
    # WICHTIG: kein f-string! Nur Replace -> keine {} Probleme
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

    data = build_data(df, max_rows=99)
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
