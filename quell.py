import streamlit as st
import pandas as pd
import json
from io import BytesIO

st.set_page_config(page_title="Excel → Interaktive HTML", layout="centered")
st.title("Excel hochladen → Interaktive HTML erzeugen")

uploaded = st.file_uploader("Excel-Datei auswählen", type=["xlsx", "xlsm", "xls"])

def norm_tour(x):
    if pd.isna(x):
        return ""
    s = str(x).strip()
    # Tournummern kommen manchmal als float "1201.0"
    if s.endswith(".0"):
        s = s[:-2]
    return s

def build_data(df: pd.DataFrame):
    # Erwartet: Zeilen 1–99 (Excel) => in pandas i.d.R. 0-based,
    # du kannst das bei Bedarf anpassen (z.B. df.iloc[0:99])
    df = df.iloc[0:99].copy()

    markets = []
    for _, r in df.iterrows():
        csb = str(r.iloc[0]).strip() if not pd.isna(r.iloc[0]) else ""
        sap = str(r.iloc[1]).strip() if not pd.isna(r.iloc[1]) else ""
        name = str(r.iloc[2]).strip() if not pd.isna(r.iloc[2]) else ""
        street = str(r.iloc[3]).strip() if not pd.isna(r.iloc[3]) else ""
        zipc = str(r.iloc[4]).strip() if not pd.isna(r.iloc[4]) else ""
        city = str(r.iloc[5]).strip() if not pd.isna(r.iloc[5]) else ""

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

        markets.append({
            "csb": csb, "sap": sap, "name": name,
            "street": street, "zip": zipc, "city": city,
            "pattern": pattern
        })

    return {
        "meta": {
            "weekStartsSunday": True,   # kannst du fest einstellen
            "minGapDays": 3
        },
        "markets": markets
    }

def render_html(data: dict) -> str:
    payload = json.dumps(data, ensure_ascii=False)

    # Minimal-Standalone-HTML (App in der HTML)
    return f"""<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Belieferungsschema – Interaktiv</title>
<style>
  body {{ font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial; margin: 16px; background:#f4f5f7; }}
  .wrap {{ max-width: 1200px; margin: 0 auto; }}
  .card {{ background:#fff; border:1px solid #ddd; border-radius:14px; padding:14px; box-shadow: 0 2px 10px rgba(0,0,0,.04); }}
  .row {{ display:flex; gap:12px; flex-wrap:wrap; align-items:center; }}
  .grow {{ flex:1; }}
  input, select, button {{ padding:10px 12px; border-radius:10px; border:1px solid #ccc; background:#fff; }}
  button {{ cursor:pointer; }}
  .grid7 {{ display:grid; grid-template-columns: repeat(7, 1fr); gap:10px; }}
  .daybtn {{ padding:10px; border-radius:12px; border:1px solid #ccc; background:#fff; text-align:center; user-select:none; cursor:pointer; }}
  .holiday {{ border-color:#d33; background: #ffecec; }}
  .muted {{ color:#666; font-size: 13px; }}
  .h2 {{ font-size:18px; margin: 8px 0; }}
  .pill {{ display:inline-block; padding:4px 10px; border-radius:999px; border:1px solid #ddd; background:#fafafa; font-size:12px; }}
  .split {{ display:grid; grid-template-columns: 1.1fr .9fr; gap:12px; }}
  @media (max-width: 900px) {{ .split {{ grid-template-columns: 1fr; }} }}
  .list {{ display:flex; flex-direction:column; gap:10px; }}
  .tour {{ border:1px solid #e3e3e3; border-radius:12px; padding:10px; }}
  .tourhead {{ display:flex; justify-content:space-between; align-items:center; gap:10px; }}
  .bad {{ color:#b00; }}
  .ok {{ color:#0a6; }}
  .small {{ font-size:12px; }}
</style>
</head>
<body>
<div class="wrap">
  <div class="card">
    <div class="row">
      <div class="grow">
        <div class="h2">Belieferungsschema (Standalone HTML)</div>
        <div class="muted">KW wählen, Feiertage anklicken – Plan wird direkt neu berechnet.</div>
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
        <label class="muted">Aktion</label><br/>
        <button id="clearH">Feiertage löschen</button>
      </div>
    </div>
    <div style="margin-top:12px" class="row">
      <span id="kwLabel" class="pill"></span>
      <span class="pill">Mindestabstand: <b id="gapLabel"></b> Tage</span>
      <span class="pill">Woche: <b id="rangeLabel"></b></span>
    </div>
  </div>

  <div style="height:12px"></div>

  <div class="card">
    <div class="h2">Feiertage in dieser Woche</div>
    <div id="weekDays" class="grid7"></div>
    <div class="muted" style="margin-top:8px">
      Tipp: Klick auf Tag = Feiertag an/aus. (Nur diese KW)
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
const DATA = {payload};
const weekStartsSunday = !!DATA.meta.weekStartsSunday;
const minGapDays = Number(DATA.meta.minGapDays || 3);

const state = {{
  date: null,
  holidays: new Set(), // ISO date strings innerhalb der aktuellen Woche
  view: "tour",
  q: ""
}};

const el = (id) => document.getElementById(id);

function pad(n) {{ return String(n).padStart(2,'0'); }}
function iso(d) {{
  return d.getFullYear() + "-" + pad(d.getMonth()+1) + "-" + pad(d.getDate());
}}
function parseISO(s) {{
  const [y,m,da] = s.split("-").map(Number);
  return new Date(y, m-1, da);
}}
function addDays(d, n) {{
  const x = new Date(d);
  x.setDate(x.getDate() + n);
  return x;
}}
function weekdayName(d) {{
  return ["So","Mo","Di","Mi","Do","Fr","Sa"][d.getDay()];
}}
function weekRange(d) {{
  // JS getDay(): So=0..Sa=6
  const dow = d.getDay();
  let start;
  if (weekStartsSunday) {{
    start = addDays(d, -dow);
  }} else {{
    // Montag als Start: Mo=1.. So=0
    const diff = (dow === 0) ? 6 : (dow - 1);
    start = addDays(d, -diff);
  }}
  const end = addDays(start, 6);
  return {{start, end}};
}}

// ISO week number (simple, ok für Anzeige)
function isoWeekNumber(date) {{
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(),0,1));
  const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return {{year: d.getUTCFullYear(), week: weekNo}};
}}

function buildWeekDaysUI() {{
  const wr = weekRange(state.date);
  const cont = el("weekDays");
  cont.innerHTML = "";
  for (let i=0;i<7;i++) {{
    const day = addDays(wr.start, i);
    const key = iso(day);
    const btn = document.createElement("div");
    btn.className = "daybtn" + (state.holidays.has(key) ? " holiday" : "");
    btn.innerHTML = `<div><b>${{weekdayName(day)}}</b></div><div class="small">${{pad(day.getDate())}}.${{pad(day.getMonth()+1)}}</div>`;
    btn.onclick = () => {{
      if (state.holidays.has(key)) state.holidays.delete(key);
      else state.holidays.add(key);
      render();
    }};
    cont.appendChild(btn);
  }}
}}

function getPatternForDow(market, dowJS) {{
  // wir arbeiten nur Mo–Sa aus Excel; Sonntag = kein Plan
  // dowJS: So=0..Sa=6
  const map = {{
    1: "mo",
    2: "di",
    3: "mi",
    4: "do",
    5: "fr",
    6: "sa"
  }};
  const k = map[dowJS];
  return k ? (market.pattern[k] || "") : "";
}}

function planForWeek() {{
  const wr = weekRange(state.date);
  const days = [];
  for (let i=0;i<7;i++) days.push(addDays(wr.start, i));

  // 1) Rohplan aus Muster: deliveries[dateISO] = Array of {market, tour}
  const deliveries = new Map();
  for (const d of days) {{
    deliveries.set(iso(d), []);
  }}

  for (const m of DATA.markets) {{
    for (const d of days) {{
      const key = iso(d);
      const tour = getPatternForDow(m, d.getDay());
      if (tour) {{
        deliveries.get(key).push({{ market: m, tour, originalDate: key }});
      }}
    }}
  }}

  // 2) Feiertage verschieben "nach vorne" (rückwärts) + minGapDays prüfen
  // Wir machen eine einfache, nachvollziehbare Heuristik:
  // - Jede Lieferung an Feiertag wird auf vorherigen Nicht-Feiertag in derselben Woche geschoben.
  // - Wenn minGapDays verletzt wird, wird weiter rückwärts geschoben (bis Wochenstart).
  // - Wenn nicht möglich: Konflikt.
  const moved = [];      // {from,to, market, tour}
  const conflicts = [];  // {type, msg, market, tour, from}

  // Index bestehender Lieferdaten pro Markt zur Gap-Prüfung
  function marketDeliveriesDates(marketId, deliveriesMap) {{
    const out = [];
    for (const [dk, arr] of deliveriesMap.entries()) {{
      for (const it of arr) {{
        if (it.market._id === marketId) out.push(dk);
      }}
    }}
    out.sort();
    return out;
  }}

  // gib markets IDs
  DATA.markets.forEach((m, idx) => m._id = idx);

  // helper: date diff in days (absolute)
  function diffDays(aISO, bISO) {{
    const a = parseISO(aISO);
    const b = parseISO(bISO);
    return Math.round((a - b) / 86400000);
  }}

  // Rebuild deliveries by processing days in chronological order,
  // but we will "extract" holiday items and reinsert
  for (const day of days) {{
    const dayKey = iso(day);
    if (!state.holidays.has(dayKey)) continue;

    const items = deliveries.get(dayKey);
    if (!items.length) continue;

    // remove all items from holiday day
    deliveries.set(dayKey, []);

    for (const it of items) {{
      // target search backward
      let target = addDays(day, -1);
      let targetKey = null;

      while (target >= wr.start) {{
        const k = iso(target);
        if (state.holidays.has(k)) {{ target = addDays(target, -1); continue; }}

        // Gap check: compare against all other planned dates for this market (including previously moved)
        // We test whether placing it on k would create any pair with < minGapDays
        const existingDates = marketDeliveriesDates(it.market._id, deliveries);
        let ok = true;
        for (const ed of existingDates) {{
          const gap = Math.abs(diffDays(k, ed));
          if (gap < minGapDays) {{ ok = false; break; }}
        }}
        if (ok) {{ targetKey = k; break; }}
        target = addDays(target, -1);
      }}

      if (!targetKey) {{
        conflicts.push({{
          type: "GAP_OR_RANGE",
          msg: `Kann ${{it.market.name}} (${{it.market.city}}) von ${{dayKey}} nicht nach vorne ziehen ohne Abstand < ${{minGapDays}} Tage oder außerhalb der KW.`,
          market: it.market, tour: it.tour, from: dayKey
        }});
        // falls Konflikt: wir lassen es auf dem ursprünglichen Tag "stehen" aber markieren als nicht lieferbar
        // alternativ: komplett entfernen. Hier: entfernen und Konflikt.
        continue;
      }}

      deliveries.get(targetKey).push(it);
      moved.push({{from: dayKey, to: targetKey, market: it.market, tour: it.tour}});
    }}
  }}

  return {{wr, days, deliveries, moved, conflicts}};
}}

function render() {{
  // labels
  const wn = isoWeekNumber(state.date);
  const wr = weekRange(state.date);
  el("kwLabel").textContent = `KW ${wn.week} / ${wn.year}`;
  el("gapLabel").textContent = String(minGapDays);
  el("rangeLabel").textContent = `${wr.start.toLocaleDateString('de-DE')} – ${wr.end.toLocaleDateString('de-DE')}`;

  buildWeekDaysUI();

  const plan = planForWeek();
  const q = state.q.trim().toLowerCase();

  // summary
  let totalStops = 0;
  for (const [, arr] of plan.deliveries.entries()) totalStops += arr.length;

  el("summary").innerHTML = `
    <div>Stops diese KW: <b>${totalStops}</b></div>
    <div>Feiertage markiert: <b>${state.holidays.size}</b></div>
    <div>Verschoben: <b>${plan.moved.length}</b></div>
    <div>Konflikte: <b class="${plan.conflicts.length ? "bad":"ok"}">${plan.conflicts.length}</b></div>
    <hr/>
    <div class="small muted">Hinweis: Die Verschiebe-Logik ist bewusst simpel gehalten (rückwärts + Mindestabstand). Kann später um Lastverteilung / Toursplitting erweitert werden.</div>
  `;

  // views
  el("left").innerHTML = "";
  el("leftTitle").textContent =
    state.view === "tour" ? "Tourenansicht" :
    state.view === "market" ? "Marktansicht" : "Konflikte";

  if (state.view === "conflicts") {{
    if (!plan.conflicts.length) {{
      el("left").innerHTML = `<div class="muted">Keine Konflikte 🎉</div>`;
      return;
    }}
    for (const c of plan.conflicts) {{
      if (q && !(c.market.name.toLowerCase().includes(q) || c.market.city.toLowerCase().includes(q) || String(c.tour).toLowerCase().includes(q))) continue;
      const div = document.createElement("div");
      div.className = "tour";
      div.innerHTML = `
        <div class="tourhead">
          <div><b class="bad">Konflikt</b> – Tour ${c.tour}</div>
          <div class="pill">${c.from}</div>
        </div>
        <div class="muted">${c.msg}</div>
      `;
      el("left").appendChild(div);
    }}
    return;
  }}

  if (state.view === "market") {{
    // markets list filtered
    const ms = DATA.markets.filter(m => {{
      if (!q) return true;
      const hay = (m.name+" "+m.city+" "+m.csb+" "+m.sap).toLowerCase();
      return hay.includes(q);
    }});
    for (const m of ms) {{
      const div = document.createElement("div");
      div.className = "tour";
      const p = m.pattern;
      div.innerHTML = `
        <div class="tourhead">
          <div><b>${m.name}</b> <span class="muted">(${m.city})</span></div>
          <div class="pill">CSB ${m.csb} · SAP ${m.sap}</div>
        </div>
        <div class="muted">${m.street}, ${m.zip} ${m.city}</div>
        <div style="margin-top:8px" class="row">
          <span class="pill">Mo: <b>${p.mo||"-"}</b></span>
          <span class="pill">Di: <b>${p.di||"-"}</b></span>
          <span class="pill">Mi: <b>${p.mi||"-"}</b></span>
          <span class="pill">Do: <b>${p.do||"-"}</b></span>
          <span class="pill">Fr: <b>${p.fr||"-"}</b></span>
          <span class="pill">Sa: <b>${p.sa||"-"}</b></span>
        </div>
      `;
      el("left").appendChild(div);
    }}
    return;
  }}

  // tour view: group by day then tour
  for (const d of plan.days) {{
    const dk = iso(d);
    const arr = plan.deliveries.get(dk) || [];
    // filter per q
    const filtered = arr.filter(it => {{
      if (!q) return true;
      const hay = (it.market.name+" "+it.market.city+" "+it.market.csb+" "+it.market.sap+" "+it.tour).toLowerCase();
      return hay.includes(q);
    }});
    // group by tour
    const byTour = new Map();
    for (const it of filtered) {{
      if (!byTour.has(it.tour)) byTour.set(it.tour, []);
      byTour.get(it.tour).push(it);
    }}

    const dayBox = document.createElement("div");
    dayBox.className = "tour";
    dayBox.innerHTML = `
      <div class="tourhead">
        <div><b>${weekdayName(d)} ${d.toLocaleDateString('de-DE')}</b>${state.holidays.has(dk) ? ' <span class="bad">(Feiertag)</span>' : ''}</div>
        <div class="pill">Stops: ${filtered.length}</div>
      </div>
      <div class="muted small">Touren: ${byTour.size}</div>
      <div class="list" style="margin-top:10px" id="inner_${dk}"></div>
    `;
    el("left").appendChild(dayBox);

    const inner = dayBox.querySelector("#inner_"+dk.replaceAll("-","_"));
    // build inner tours
    for (const [tour, items] of [...byTour.entries()].sort((a,b)=>String(a[0]).localeCompare(String(b[0])))) {{
      const tdiv = document.createElement("div");
      tdiv.className = "tour";
      tdiv.innerHTML = `
        <div class="tourhead">
          <div><b>Tour ${tour}</b></div>
          <div class="pill">${items.length} Märkte</div>
        </div>
        <div class="muted small">${items.map(x => x.market.name + " ("+x.market.city+")").join(" · ")}</div>
      `;
      inner.appendChild(tdiv);
    }}
  }}
}}

function init() {{
  // default date = today
  const today = new Date();
  state.date = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  el("datePick").value = iso(state.date);

  el("datePick").addEventListener("change", (e) => {{
    state.date = parseISO(e.target.value);
    state.holidays.clear(); // Feiertage pro KW neu setzen
    render();
  }});
  el("view").addEventListener("change", (e) => {{
    state.view = e.target.value;
    render();
  }});
  el("q").addEventListener("input", (e) => {{
    state.q = e.target.value;
    render();
  }});
  el("clearH").addEventListener("click", () => {{
    state.holidays.clear();
    render();
  }});

  render();
}}
init();
</script>
</body>
</html>
"""

if uploaded:
    df = pd.read_excel(uploaded, sheet_name="Direkt", header=None)
    data = build_data(df)
    html = render_html(data)

    st.success(f"{len(data['markets'])} Märkte geladen. HTML bereit.")
    st.download_button(
        "Interaktive HTML herunterladen",
        data=html.encode("utf-8"),
        file_name="belieferung_interaktiv.html",
        mime="text/html"
    )
else:
    st.info("Bitte Excel hochladen.")
