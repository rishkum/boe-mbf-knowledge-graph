"""Streamlit explorer for the Bank of England MBF research knowledge graph.

Self-contained: reads kg.sqlite next to this file (built by build_demo_data.py). No API keys, no network calls.
Run locally:  streamlit run demo/streamlit_app.py
"""
import io
import json
import sqlite3
import zipfile
from collections import defaultdict
from datetime import date
from pathlib import Path

import altair as alt
import networkx as nx
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

HERE = Path(__file__).resolve().parent
DB = HERE / "kg.sqlite"
XLSX = HERE / "exports" / "BoE_MBF_KG.xlsx"
CSV_DIR = HERE / "exports" / "csv"

BLUE, ORANGE, GREY = "#2a78d6", "#eb6834", "#7d7b75"
SERIES_ORDER = ["Bank Overground", "Bank Insight", "Bank Underground", "Quarterly Bulletin", "Staff Working Paper"]
GENERIC_HUBS = {"financial_stability", "asset_prices", "monetary_transmission", "mbf"}  # vague as middle steps
FACETS = ["sector", "market", "vulnerability", "phenomenon", "policy"]

# the links graded in the insights note (Table 1): label, driver, outcome, direction (None = any)
PRESETS = [
    ("Price falls, volatility → bigger margin calls", "asset_prices", "margin", "increases"),
    ("Margin and collateral calls → forced sales", "margin", "fire_sales", "increases"),
    ("Forced sales → further price falls, volatility", "fire_sales", "asset_prices", None),
    ("Leverage → forced sales", "leverage", "fire_sales", "increases"),
    ("Dealer balance-sheet limits → market dysfunction", "dealer_constraints", "market_dysfunction", "increases"),
    ("Liquidity mismatch → forced sales", "liquidity_mismatch", "fire_sales", None),
    ("Central bank backstops → less dysfunction", "cb_backstop", "market_dysfunction", "decreases"),
    ("Central clearing → dealer capacity", "ccps", "dealer_constraints", None),
]

st.set_page_config(page_title="MBF Knowledge Graph", page_icon=":material/hub:", layout="wide")


# ====================================================================== data
@st.cache_data
def load():
    con = sqlite3.connect(f"file:{DB}?mode=ro", uri=True)
    t = {name: pd.read_sql_query(f"SELECT * FROM {name}", con) for (name,) in
         con.execute("SELECT name FROM sqlite_master WHERE type='table'")}
    con.close()
    t["publications"]["date"] = pd.to_datetime(t["publications"]["date"])
    t["claims"]["date"] = pd.to_datetime(t["claims"]["date"])
    return t


T = load()
PUBS, CLAIMS, CONCEPTS = T["publications"], T["claims"], T["concepts"]
PUB = PUBS.set_index("doc_id")
NAME = dict(zip(CONCEPTS.concept_id, CONCEPTS.name))
FACET = dict(zip(CONCEPTS.concept_id, CONCEPTS.facet))
SHOWN = set(CONCEPTS[CONCEPTS.shown == 1].concept_id)
KIDS = defaultdict(list)
for b, n in T["concept_hierarchy"].itertuples(index=False):
    KIDS[b].append(n)
PC = T["publication_concepts"]
MAIN = PC[PC.salience >= 0.6]  # central or analysed (not passing mentions)


def name(cid):
    return NAME.get(cid, cid)


def family(cid):
    """A concept plus all its narrower concepts."""
    out, stack = {cid}, [cid]
    while stack:
        for k in KIDS[stack.pop()]:
            if k not in out:
                out.add(k)
                stack.append(k)
    return out


def concept_picker(label, default=None, facets=None, key=None):
    ids = sorted((c for c in SHOWN if facets is None or FACET[c] in facets), key=lambda c: name(c).lower())
    idx = ids.index(default) if default in ids else 0
    return st.selectbox(label, ids, index=idx, key=key, format_func=lambda c: f"{name(c)}  ·  {FACET[c]}")


def docs_with(cid, min_salience=0.6, narrower=True):
    rows = PC[(PC.salience >= min_salience) & PC.concept_id.isin(family(cid) if narrower else {cid})]
    return set(rows.doc_id)


def with_pub(df, id_col="doc_id"):
    return df.assign(publication=df[id_col].map(PUB.title), series=df[id_col].map(PUB.series),
                     link=df[id_col].map(PUB.url))


LINK = st.column_config.LinkColumn("Link", display_text="Open")
DATE = st.column_config.DateColumn("Date", format="YYYY-MM-DD")


def claims_table(df, height="auto"):
    show = with_pub(df)[["date", "driver_raw", "direction", "outcome_raw", "evidence_type", "hedging", "quote",
                         "publication", "series", "link"]]
    st.dataframe(show.sort_values("date"), hide_index=True, width="stretch", height=height,
                 column_config={"date": DATE, "link": LINK, "driver_raw": "Driver (as written)",
                                "outcome_raw": "Outcome (as written)", "evidence_type": "Evidence",
                                "quote": st.column_config.TextColumn("Quote", width="large")})


def link_claims(driver, outcome, direction=None, narrower=True):
    d = family(driver) if narrower else {driver}
    o = family(outcome) if narrower else {outcome}
    df = CLAIMS[CLAIMS.driver.isin(d) & CLAIMS.outcome.isin(o)]
    return df[df.direction == direction] if direction else df


def grade(df):
    if df.empty:
        return {"Publications": 0, "Claims": 0, "Years": "–", "Measured in data": "–", "Stated firmly": "–"}
    return {"Publications": df.doc_id.nunique(), "Claims": len(df),
            "Years": f"{df.date.min().year}–{str(df.date.max().year)[2:]}",
            "Measured in data": f"{(df.evidence_type == 'empirical').mean():.0%}",
            "Stated firmly": f"{(df.hedging == 'firm').mean():.0%}"}


def metrics_row(d):
    for col, (k, v) in zip(st.columns(len(d)), d.items()):
        col.metric(k, v)


# ====================================================================== graph views
FACET_COLOUR = {"sector": "#2a78d6", "market": "#eb6834", "vulnerability": "#1baf7a", "phenomenon": "#eda100",
                "policy": "#e87ba4", "event": "#008300", "method": "#e34948", "emergent": "#6b5fc7"}
UP, DOWN, MIXED = "#d03b3b", "#2a78d6", "#a8a69e"
VIS_JS = HERE / "assets" / "vis-network.min.js"  # bundled so the map works offline
NODE_PUBS = MAIN.groupby("concept_id").doc_id.nunique()


@st.cache_data
def link_table():
    """Every causal link between concepts: publications, claims and how many say increases / decreases."""
    df = CLAIMS[(CLAIMS.driver != CLAIMS.outcome) & CLAIMS.driver.isin(SHOWN) & CLAIMS.outcome.isin(SHOWN)]
    g = df.groupby(["driver", "outcome"])
    return pd.DataFrame({"publications": g.doc_id.nunique(), "claims": g.size(),
                         "increases": g.direction.apply(lambda s: int((s == "increases").sum())),
                         "decreases": g.direction.apply(lambda s: int((s == "decreases").sum()))}).reset_index()


def direction_colour(up, down):
    if up + down == 0:
        return MIXED
    share = up / (up + down)
    return UP if share >= 0.75 else DOWN if share <= 0.25 else MIXED


def network(links, focus=None, height=560):
    """Interactive concept network (vis-network): drag, zoom, hover for details."""
    ids = sorted(set(links.driver) | set(links.outcome))
    nodes = [{"id": c, "label": name(c), "value": int(NODE_PUBS.get(c, 0)) + 1,
              "title": f"{name(c)}\n{FACET[c]} · {NODE_PUBS.get(c, 0)} publications",
              "color": {"background": FACET_COLOUR.get(FACET[c], GREY), "border": "#222" if c == focus else "#fff"},
              "borderWidth": 3 if c == focus else 1} for c in ids]
    edges = [{"from": r.driver, "to": r.outcome, "value": int(r.publications),
              "title": f"{name(r.driver)} → {name(r.outcome)}\n{r.publications} publications, {r.claims} claims"
                       f"\n{r.increases} say increases, {r.decreases} say decreases",
              "color": {"color": direction_colour(r.increases, r.decreases), "opacity": 0.8}}
             for r in links.itertuples()]
    data = json.dumps({"nodes": nodes, "edges": edges}).replace("</", "<\\/")
    components.html(f"""
<div id="net" style="height:{height}px;border:1px solid #e6e4dd;border-radius:8px;background:#fff"></div>
<script>{VIS_JS.read_text()}</script>
<script>
const data = {data};
for (const x of [...data.nodes, ...data.edges]) {{
  const el = document.createElement("div");
  el.style.cssText = "white-space:pre-line;font:13px sans-serif;padding:4px";
  el.textContent = x.title; x.title = el;
}}
const net = new vis.Network(document.getElementById("net"), data, {{
  nodes: {{shape: "dot", scaling: {{min: 8, max: 34, label: {{enabled: true, min: 12, max: 18}}}},
          font: {{face: "sans-serif", color: "#222", strokeWidth: 3, strokeColor: "#fff"}}}},
  edges: {{arrows: {{to: {{scaleFactor: 0.55}}}}, scaling: {{min: 1, max: 8}}, smooth: {{type: "continuous"}}}},
  physics: {{solver: "forceAtlas2Based", forceAtlas2Based: {{gravitationalConstant: -70, springLength: 130}},
            stabilization: {{iterations: 300}}}},
  interaction: {{hover: true, tooltipDelay: 80, zoomView: false, navigationButtons: true, keyboard: false}}
}});
net.once("stabilizationIterationsDone", () => net.fit({{animation: false}}));
</script>""", height=height + 12)


def legend():
    dots = "".join(f'<span style="margin-right:14px;white-space:nowrap"><span style="display:inline-block;width:10px;'
                   f'height:10px;border-radius:50%;background:{c};margin-right:5px"></span>{f}</span>'
                   for f, c in FACET_COLOUR.items() if f not in ("event", "method"))
    lines = "".join(f'<span style="margin-right:14px;white-space:nowrap"><span style="display:inline-block;width:18px;'
                    f'height:3px;background:{c};vertical-align:middle;margin-right:5px"></span>{t}</span>'
                    for t, c in (("mostly increases", UP), ("mostly decreases", DOWN), ("mixed", MIXED)))
    st.html(f'<div style="font-size:13px;color:#52514e;line-height:1.9">{dots}<br>{lines} · dot size = publications'
            f' · arrow width = publications making the claim</div>')


def neighbourhood(links, focus, depth):
    keep, frontier = {focus}, {focus}
    for _ in range(depth):
        nb = links[links.driver.isin(frontier) | links.outcome.isin(frontier)]
        new = set(nb.driver) | set(nb.outcome)
        frontier, keep = new - keep, keep | new
    return links[links.driver.isin(keep) & links.outcome.isin(keep)]


def dot_label(text):
    return text.replace('"', "'")


# ====================================================================== pages
def overview():
    st.title("What Bank of England research says about market-based finance")
    st.caption("A knowledge graph of Bank of England staff publications on market-based finance (MBF), 2015–2026")
    metrics_row({"Publications": len(PUBS), "Causal claims": f"{len(CLAIMS):,}",
                 "Concepts (fixed vocabulary)": int(CONCEPTS.seeded.sum()),
                 "Open questions": len(T["open_questions"]), "Stress episodes": len(T["events"])})
    st.markdown(
        "Each publication was read by a language model, which recorded its **causal claims** as "
        "*driver → direction → outcome* (for example, *margin calls increase forced sales*), with the type of "
        "evidence, how firmly it is stated, the date and a supporting quote checked against the source. "
        "Claims are linked through a shared vocabulary of concepts with synonyms, so the same idea is counted "
        "once however it is worded.")
    st.markdown("**Things to try**")
    c1, c2 = st.columns(2)
    c1.markdown("- **Grade the evidence** for a causal link before a briefing\n"
                "- **Audit foresight**: what was known before the 2022 LDI episode\n"
                "- **Trace a scenario** from a shock to an outcome, step by step")
    c2.markdown("- **Brief yourself** on a topic in one page\n"
                "- **Find research gaps** in the sector × vulnerability matrix\n"
                "- **See where findings differ** and under what conditions")

    corpus = T["corpus_by_year"].groupby("year")[["mbf_publications", "all_publications"]].sum().reset_index()
    corpus["share"] = corpus.mbf_publications / corpus.all_publications
    by_series = PUBS.groupby(["year", "series"]).size().reset_index(name="publications")
    left, right = st.columns(2)
    with left:
        st.markdown("**MBF publications by year and series**")
        st.altair_chart(alt.Chart(by_series).mark_bar().encode(
            x=alt.X("year:O", title=None), y=alt.Y("publications:Q", title="Publications"),
            color=alt.Color("series:N", sort=SERIES_ORDER, legend=alt.Legend(orient="bottom", title=None)),
            tooltip=["year", "series", "publications"]).properties(height=280), width="stretch")
    with right:
        st.markdown("**MBF share of all publications in the five series**")
        st.altair_chart(alt.Chart(corpus).mark_line(point=True, color=BLUE).encode(
            x=alt.X("year:O", title=None), y=alt.Y("share:Q", title="Share", axis=alt.Axis(format="%")),
            tooltip=["year", alt.Tooltip("share:Q", format=".0%"), "mbf_publications", "all_publications"])
            .properties(height=280), width="stretch")
        st.caption("2026 runs to 1 October. Which publications count as MBF was decided automatically.")

    with st.expander("Known limitations"):
        st.markdown(
            "- Direction, hedging, evidence type and concept matching were assigned by a language model and have "
            "not been reviewed by experts. Counts show where evidence sits, not its quality.\n"
            "- The *US Treasuries* concept also picks up generic statements about interest rates.\n"
            "- About 14% of quotes are the closest source passage rather than the model's exact quote.\n"
            "- Working papers were read from their abstract, introduction and conclusion. Staff publications set "
            "out their authors' views, not those of the Bank's policy committees.")


def causal_map():
    st.title("Causal map")
    st.caption("Concepts (dots) joined by the causal claims Bank publications make about them (arrows). "
               "Drag to rearrange, use the + and − buttons to zoom, hover for details.")
    links = link_table()
    ids = sorted(SHOWN, key=lambda c: name(c).lower())
    c1, c2, c3 = st.columns([3, 1, 1])
    with c1:
        focus = st.selectbox("Centre on a concept", [None] + ids, index=ids.index("fire_sales") + 1,
                             format_func=lambda c: "Whole graph" if c is None else f"{name(c)}  ·  {FACET[c]}")
    with c2:
        depth = st.radio("Steps out", [1, 2], horizontal=True, disabled=focus is None)
    with c3:
        min_pubs = st.number_input("Min publications per link", 1, 10, 2 if focus else 4)
    types = ["sector", "market", "vulnerability", "phenomenon", "policy", "emergent"]
    kinds = st.pills("Show concept types", types, default=types,
                     selection_mode="multi", format_func=lambda f: "proposed by model" if f == "emergent" else f)
    links = links[(links.publications >= min_pubs) & links.driver.map(FACET).isin(kinds)
                  & links.outcome.map(FACET).isin(kinds)]
    if focus:
        links = neighbourhood(links, focus, depth)
    links = links.nlargest(120, "publications")
    if links.empty:
        st.info("No links at this threshold. Lower the minimum publications or add concept types.")
        return
    metrics_row({"Concepts shown": len(set(links.driver) | set(links.outcome)), "Links shown": len(links),
                 "Claims behind them": int(links.claims.sum())})
    network(links, focus)
    legend()
    table = links.assign(link=links.driver.map(name) + " → " + links.outcome.map(name)) \
        .sort_values("publications", ascending=False)
    st.markdown("**Look at the evidence behind a link**")
    pick = st.selectbox("Link", table.index, format_func=lambda i: f"{table.link[i]}  ({table.publications[i]} "
                        f"publications)", label_visibility="collapsed")
    claims_table(link_claims(table.driver[pick], table.outcome[pick], narrower=False))


def spiral_dot(selected=None):
    """The collateral and leverage spiral from the insights note, with live publication counts."""
    lines = ['digraph { rankdir=LR; bgcolor="transparent"; nodesep=0.35; ranksep=0.5;',
             'node [shape=box, style="rounded,filled", fontname="Helvetica", fontsize=11, color="#c9c7bf", '
             'fillcolor="#f4f3ef"]; edge [fontname="Helvetica", fontsize=10, color="#9a988f"];']
    for c in {x for p in PRESETS for x in p[1:3]}:
        lines.append(f'"{c}" [label="{dot_label(name(c))}", fillcolor="{FACET_COLOUR.get(FACET[c], GREY)}22"];')
    for p in PRESETS:
        n = link_claims(p[1], p[2], p[3]).doc_id.nunique()
        on = selected is not None and p[0] == selected[0]
        style = f'color="{BLUE}", fontcolor="{BLUE}", penwidth=3' if on else f"penwidth={1 + n / 8:.1f}"
        dash = ', style="dashed"' if p[3] == "decreases" else ""
        lines.append(f'"{p[1]}" -> "{p[2]}" [label="{n}", {style}{dash}];')
    return "\n".join(lines + ["}"])


def chains_dot(found, g, s, t):
    edges = {}
    for _, p in found:
        for u, v in zip(p, p[1:]):
            edges[(u, v)] = len(g[u][v]["docs"])
    ends = family(s) | family(t)
    lines = ['digraph { rankdir=LR; bgcolor="transparent"; nodesep=0.3;',
             'node [shape=box, style="rounded,filled", fontname="Helvetica", fontsize=11, color="#c9c7bf"];',
             'edge [fontname="Helvetica", fontsize=10, color="#7d7b75"];']
    for c in {x for e in edges for x in e}:
        bold = ', penwidth=2, color="#222"' if c in ends else ""
        lines.append(f'"{c}" [label="{dot_label(name(c))}", fillcolor="{FACET_COLOUR.get(FACET[c], GREY)}22"{bold}];')
    for (u, v), n in edges.items():
        lines.append(f'"{u}" -> "{v}" [label="{n}", penwidth={1 + min(n, 8) * 0.5:.1f}];')
    return "\n".join(lines + ["}"])


def evidence():
    st.title("Grade the evidence for a causal link")
    st.caption("How many publications support a link, over what period, how it was measured and how firmly "
               "it is stated.")
    labels = ["Custom link"] + [p[0] for p in PRESETS]
    choice = st.selectbox("Start from a link in the spiral, or build your own", labels, index=2)
    preset = next((p for p in PRESETS if p[0] == choice), None)
    with st.expander("Where this link sits in the spiral", expanded=True):
        st.graphviz_chart(spiral_dot(preset), width="stretch")
        st.caption("Numbers are publications making each claim. The selected link is highlighted.")
    c1, c2, c3 = st.columns([2, 2, 1])
    with c1:
        drv = concept_picker("Driver", preset[1] if preset else "leverage", key=f"drv{choice}")
    with c2:
        out = concept_picker("Outcome", preset[2] if preset else "fire_sales", key=f"out{choice}")
    with c3:
        dirs = ["any", "increases", "decreases", "no_effect", "ambiguous"]
        direction = st.selectbox("Direction", dirs, index=dirs.index(preset[3]) if preset and preset[3] else 0,
                                 key=f"dir{choice}")
    narrower = st.toggle("Include narrower concepts", True,
                         help="e.g. procyclical margin counts as margin calls; basis trades count as leverage")
    df = link_claims(drv, out, None if direction == "any" else direction, narrower)
    st.subheader(f"{name(drv)} → {name(out)}")
    metrics_row(grade(df))
    if df.empty:
        st.info("No publication in the graph makes this claim. Try a broader concept or 'any' direction.")
        return
    per_year = df.drop_duplicates("doc_id").assign(year=lambda d: d.date.dt.year).groupby("year").size()
    left, right = st.columns([3, 2])
    with left:
        st.altair_chart(alt.Chart(per_year.reset_index(name="publications")).mark_bar(color=BLUE).encode(
            x=alt.X("year:O", title=None),
            y=alt.Y("publications:Q", title="Publications making the claim", axis=alt.Axis(tickMinStep=1, format="d")))
            .properties(height=240), width="stretch")
    with right:
        mix = df.groupby(["evidence_type"]).size().reset_index(name="claims")
        st.altair_chart(alt.Chart(mix).mark_bar(color=GREY).encode(
            y=alt.Y("evidence_type:N", title=None, sort="-x"),
            x=alt.X("claims:Q", title="Claims by evidence type", axis=alt.Axis(tickMinStep=1, format="d")))
            .properties(height=240), width="stretch")
    if direction == "any" and df.direction.nunique() > 1:
        st.caption("Directions: " + ", ".join(f"{k} {v}" for k, v in df.direction.value_counts().items()) +
                   ". Check the wording: 'lower prices' and 'higher volatility' can describe the same effect.")
    claims_table(df)


def foresight():
    st.title("Audit foresight: what was known before an event?")
    st.caption("Every claim is dated, so the graph can be cut at any point in time.")
    cut = st.date_input("Cut-off (claims published before this date count as 'before')", date(2022, 9, 23),
                        min_value=date(2015, 1, 1), max_value=date(2026, 10, 1))
    picks = st.multiselect("Causal links to check", [p[0] for p in PRESETS], default=[p[0] for p in PRESETS[:4]])
    st.markdown("**And publications analysing two topics together**")
    c1, c2 = st.columns(2)
    with c1:
        a = concept_picker("Topic", "ldi", key="fa")
    with c2:
        bs = st.multiselect("…together with any of", sorted(SHOWN, key=lambda c: name(c).lower()),
                            default=["leverage", "margin", "procyclical_margin", "repo"], format_func=name)
    pair_narrower = st.toggle("Include narrower concepts in the topic pair", False,
                              help="Off matches the insights note. On also counts e.g. basis trades as leverage.")
    rows = []
    for label, d, o, s in (p for p in PRESETS if p[0] in picks):
        for doc in link_claims(d, o, s).drop_duplicates("doc_id").doc_id:
            rows.append((label, doc))
    if bs:
        both = docs_with(a, narrower=pair_narrower) & set().union(
            *(docs_with(b, narrower=pair_narrower) for b in bs))
        rows += [(f"{name(a)} with {', '.join(name(b) for b in bs)}", doc) for doc in both]
    if not rows:
        st.info("Pick at least one link or topic pair.")
        return
    df = pd.DataFrame(rows, columns=["row", "doc_id"]).pipe(with_pub)
    df["date"] = df.doc_id.map(PUB.date)
    df["when"] = df.date.lt(pd.Timestamp(cut)).map({True: "before", False: "after"})
    counts = df.groupby(["row", "when"]).size().unstack(fill_value=0).reindex(columns=["before", "after"],
                                                                              fill_value=0)
    order = list(dict.fromkeys(df.row))
    base = alt.Chart(df).encode(x=alt.X("date:T", title=None),
                                y=alt.Y("row:N", sort=order, title=None, axis=alt.Axis(labelLimit=420)),
                                tooltip=["publication", alt.Tooltip("date:T"), "series"])
    rule = alt.Chart(pd.DataFrame({"date": [pd.Timestamp(cut)]})).mark_rule(strokeDash=[4, 4]).encode(x="date:T")
    st.altair_chart((base.transform_filter("datum.when == 'before'").mark_circle(size=90, color=BLUE)
                     + base.transform_filter("datum.when == 'after'").mark_point(size=90, color=BLUE, filled=False)
                     + rule).properties(height=60 * len(order) + 40), width="stretch")
    st.caption("Filled dots: before the cut-off. Open dots: after. Hover for the publication.")
    st.dataframe(counts.reindex(order), width="stretch")
    with st.expander("Publications before the cut-off"):
        pre = df[df.when == "before"].sort_values("date")
        st.dataframe(pre[["row", "date", "publication", "series", "link"]], hide_index=True,
                     width="stretch", column_config={"date": DATE, "link": LINK})


@st.cache_data
def claim_graph(min_docs):
    g = nx.DiGraph()
    for r in CLAIMS[CLAIMS.driver != CLAIMS.outcome].itertuples():
        if r.driver not in SHOWN or r.outcome not in SHOWN:
            continue
        e = g.get_edge_data(r.driver, r.outcome) or {"docs": set(), "claims": []}
        e["docs"].add(r.doc_id)
        e["claims"].append(r.claim_id)
        g.add_edge(r.driver, r.outcome, **e)
    g.remove_edges_from([(u, v) for u, v, d in g.edges(data=True) if len(d["docs"]) < min_docs])
    return g


def scenario():
    st.title("Trace a scenario")
    st.caption("Follow causal claims step by step from a shock to an outcome. Each step shows how many "
               "publications support it.")
    c1, c2, c3, c4 = st.columns([3, 3, 1, 1])
    with c1:
        s = concept_picker("From", "us_treasury", key="s")
    with c2:
        t = concept_picker("To", "market_dysfunction", key="t")
    with c3:
        hops = st.number_input("Max steps", 1, 4, 3)
    with c4:
        min_docs = st.number_input("Min pubs per step", 1, 5, 1)
    g = claim_graph(min_docs)
    claims = CLAIMS.set_index("claim_id")
    found = []
    for src in (c for c in family(s) if c in g):
        for tgt in (c for c in family(t) if c in g):
            if src == tgt:
                continue
            for p in nx.all_simple_paths(g, src, tgt, cutoff=hops):
                if any(c in GENERIC_HUBS for c in p[1:-1]):
                    continue
                found.append((min(len(g[u][v]["docs"]) for u, v in zip(p, p[1:])), p))
    found.sort(key=lambda x: (-x[0], len(x[1])))
    if not found:
        st.info("No chain of claims links these concepts. Try more steps or broader concepts.")
        return
    st.markdown(f"**{len(found)} chains found**, strongest first (a chain is only as strong as its weakest step).")
    st.graphviz_chart(chains_dot(found[:6], g, s, t), width="stretch")
    st.caption("The six strongest chains. Numbers on arrows are publications supporting each step; "
               "thicker arrows have more support.")
    for weakest, p in found[:10]:
        with st.expander(f"{' → '.join(name(c) for c in p)}   ·   weakest step: {weakest} publication"
                         f"{'s' if weakest > 1 else ''}", expanded=False):
            for u, v in zip(p, p[1:]):
                d = g[u][v]
                ex = claims.loc[d["claims"]].sort_values(["hedging", "date"]).iloc[0]
                st.markdown(f"**{name(u)} → {name(v)}** · {len(d['docs'])} publication"
                            f"{'s' if len(d['docs']) > 1 else ''}")
                st.markdown(f"> {ex.quote}  \n— *{PUB.title[ex.doc_id]}* ({ex.date:%Y}), "
                            f"[open]({PUB.url[ex.doc_id]})")
    st.caption("Concepts can be broader than they look: for example, US Treasuries also picks up generic "
               "interest-rate claims. Read each step with its source.")


def briefing():
    st.title("Brief yourself on a topic")
    cid = concept_picker("Topic", "private_credit", key="brief")
    fam = family(cid)
    docs = docs_with(cid)
    row = CONCEPTS.set_index("concept_id").loc[cid]
    st.subheader(name(cid))
    if row.definition:
        st.markdown(f"*{row.definition}*")
    narrower = [name(c) for c in fam - {cid}]
    learn = T["learning_path"]
    learn = [name(c) for c in learn[learn.concept == cid].learn_first]
    bits = [f"Facet: {row.facet}"] + ([f"Includes: {', '.join(narrower)}"] if narrower else []) + \
           ([f"Learn first: {', '.join(learn)}"] if learn else [])
    st.caption(" · ".join(bits))
    cl = CLAIMS[CLAIMS.driver.isin(fam) | CLAIMS.outcome.isin(fam)]
    oq = T["open_questions"]
    oq = oq[oq.concepts.fillna("").map(lambda s: bool(set(s.split("; ")) & fam))].sort_values("date",
                                                                                               ascending=False)
    pubs = PUBS[PUBS.doc_id.isin(docs)]
    metrics_row({"Publications (central or analysed)": len(pubs), "Causal claims": len(cl),
                 "Open questions": len(oq),
                 "Years": f"{pubs.year.min()}–{pubs.year.max()}" if len(pubs) else "–"})
    if pubs.empty:
        st.info("No publication analyses this concept.")
        return
    st.caption("By series: " + ", ".join(f"{v} {k}" for k, v in pubs.series.value_counts().items()))

    st.markdown("#### Concept map")
    near = neighbourhood(link_table(), cid, 1).nlargest(40, "publications")
    if len(near):
        network(near, focus=cid, height=430)
        legend()
    st.markdown("#### What the research claims")
    links = cl.assign(link=cl.driver.map(name) + " → " + cl.outcome.map(name)).groupby("link").agg(
        publications=("doc_id", "nunique"), claims=("claim_id", "size"),
        measured_in_data=("evidence_type", lambda s: (s == "empirical").mean()),
        stated_firmly=("hedging", lambda s: (s == "firm").mean())).sort_values("publications", ascending=False)
    st.dataframe(links, width="stretch", column_config={
        "measured_in_data": st.column_config.ProgressColumn("Measured in data", format="percent", max_value=1),
        "stated_firmly": st.column_config.ProgressColumn("Stated firmly", format="percent", max_value=1)})
    with st.expander(f"All {len(cl)} claims with quotes"):
        claims_table(cl)

    left, right = st.columns(2)
    with left:
        st.markdown("#### Open questions")
        for r in oq.head(8).itertuples():
            st.markdown(f"- {r.question} *({PUB.title.get(r.doc_id, '')}, {r.date[:4]})*")
        if oq.empty:
            st.caption("None recorded.")
    with right:
        st.markdown("#### Gaps")
        other = "sector" if row.facet == "vulnerability" else "vulnerability"
        gaps = [name(c) for c in sorted(SHOWN) if FACET[c] == other and c != "mbf" and not (docs & docs_with(c))]
        st.markdown(f"No publication analyses {name(cid).lower()} together with: "
                    + (", ".join(gaps) if gaps else "*(none: every pairing has at least one)*") + ".")

    st.markdown("#### Reading list")
    central = PUBS[PUBS.doc_id.isin(docs_with(cid, 1.0))].copy()
    central["order"] = central.series.map({s: i for i, s in enumerate(SERIES_ORDER)})
    central = central.sort_values(["order", "date"])
    st.caption("Publications where the topic is central, short explainers first.")
    st.dataframe(central[["date", "title", "series", "url"]].head(12), hide_index=True, width="stretch",
                 column_config={"date": DATE, "url": LINK})

    md = [f"# Briefing pack: {name(cid)}", "", f"{row.definition or ''}", "",
          f"Publications: {len(pubs)} ({pubs.year.min()}–{pubs.year.max()})", "", "## What the research claims"]
    md += [f"- {k}: {v.publications} publications, {v.claims} claims" for k, v in links.head(15).iterrows()]
    md += ["", "## Open questions"] + [f"- {r.question}" for r in oq.head(8).itertuples()]
    md += ["", "## Gaps", ", ".join(gaps) or "none", "", "## Reading list"]
    md += [f"- {r.title} ({r.date:%Y}) {r.url}" for r in central.head(12).itertuples()]
    st.download_button("Download this pack (Markdown)", "\n".join(md), f"briefing_{cid}.md", "text/markdown")


def disagreements():
    st.title("Know when a finding holds")
    st.caption("Pairs of claims on the same link that point in opposite directions, with the automated check's "
               "verdict on why they differ.")
    rel = T["claim_relations"]
    c1, c2 = st.columns([2, 1])
    with c1:
        any_c = st.toggle("Filter by topic", True)
        cid = concept_picker("Topic", "dealers", key="dis") if any_c else None
    with c2:
        options = sorted(rel.relation.unique())
        kinds = st.multiselect("Verdict", options,
                               default=[o for o in options if o != "same or supporting claim"])
    claims = CLAIMS.set_index("claim_id")
    df = rel[rel.relation.isin(kinds)].copy()
    if cid:
        fam = family(cid)
        later = claims.loc[df.later_claim_id]
        df = df[(later.driver.isin(fam) | later.outcome.isin(fam)).values]
    st.markdown(f"**{len(df)} pairs**")

    def side(ids):
        c = claims.loc[ids]
        return (c.driver_raw + " → " + c.direction + " → " + c.outcome_raw).values, \
            c.doc_id.map(PUB.title).values, c.date.dt.year.values

    if df.empty:
        return
    lc, lp, ly = side(df.later_claim_id)
    ec, ep, ey = side(df.earlier_claim_id)
    show = pd.DataFrame({"verdict": df.relation.values, "later claim": lc, "later publication": lp, "year": ly,
                         "earlier claim": ec, "earlier publication": ep, "year ": ey, "why they differ": df.reason})
    st.dataframe(show, hide_index=True, width="stretch",
                 column_config={"why they differ": st.column_config.TextColumn(width="large"),
                                "year": st.column_config.NumberColumn(format="%d"),
                                "year ": st.column_config.NumberColumn(format="%d")})


def gaps():
    st.title("Find research gaps")
    st.caption("Publications in which both concepts are central or analysed. Zeros are gaps within this collection, "
               "not proof that a topic has never been studied.")
    c1, c2 = st.columns(2)
    with c1:
        fa = st.selectbox("Rows", FACETS, index=0)
    with c2:
        fb = st.selectbox("Columns", FACETS, index=2)
    rows = sorted((c for c in SHOWN if FACET[c] == fa and c != "mbf"), key=name)
    cols = sorted((c for c in SHOWN if FACET[c] == fb and c != "mbf"), key=name)
    sets = {c: docs_with(c) for c in set(rows) | set(cols)}
    m = pd.DataFrame([(name(r), name(c), len(sets[r] & sets[c])) for r in rows for c in cols if r != c],
                     columns=["row", "col", "n"])
    heat = alt.Chart(m).encode(x=alt.X("col:N", title=None, axis=alt.Axis(labelAngle=-35, labelLimit=240)),
                               y=alt.Y("row:N", title=None, axis=alt.Axis(labelLimit=280)))
    st.altair_chart((heat.mark_rect().encode(
        color=alt.Color("n:Q", scale=alt.Scale(scheme="blues"), legend=alt.Legend(title="Publications")),
        tooltip=["row", "col", "n"])
        + heat.mark_text(fontSize=11).encode(
            text="n:Q", color=alt.condition("datum.n == 0", alt.value("#d03b3b"), alt.value("black"))))
        .properties(height=34 * len(rows) + 60), width="stretch")
    zeros = m[m.n == 0]
    st.markdown(f"**{len(zeros)} pairings with no publication**")
    st.dataframe(zeros[["row", "col"]], hide_index=True, width="stretch")


def agenda():
    st.title("Track the research agenda")
    split = st.slider("Compare publications before and from", 2017, 2026, 2023)
    early = PUBS[PUBS.year < split]
    late = PUBS[PUBS.year >= split]
    facets = st.multiselect("Facets", FACETS, default=["sector", "market", "vulnerability", "phenomenon"])
    rows = []
    for c in (c for c in SHOWN if FACET[c] in facets and c != "mbf"):
        d = docs_with(c)
        a, b = len(d & set(early.doc_id)), len(d & set(late.doc_id))
        if a + b >= 8:
            rows.append((name(c), a, b, a / max(len(early), 1), b / max(len(late), 1)))
    df = pd.DataFrame(rows, columns=["concept", "before_n", "after_n", "before", "after"])
    df["change"] = df.after - df.before
    df = df.sort_values("change")
    st.caption(f"Share of MBF publications in which each concept is central or analysed: "
               f"{early.year.min()}–{split - 1} ({len(early)} publications) vs {split}–2026 ({len(late)}). "
               "Concepts with at least 8 publications.")
    long = df.melt(id_vars=["concept", "change"], value_vars=["before", "after"], var_name="period",
                   value_name="share")
    y = alt.Y("concept:N", sort=list(df.concept), title=None, axis=alt.Axis(labelLimit=320))
    st.altair_chart((alt.Chart(df).mark_rule(color=GREY).encode(y=y, x="before:Q", x2="after:Q")
                     + alt.Chart(long).mark_circle(size=80).encode(
                         y=y, x=alt.X("share:Q", title="Share of MBF publications", axis=alt.Axis(format="%")),
                         color=alt.Color("period:N", scale=alt.Scale(domain=["before", "after"],
                                                                     range=["#9ec2ee", BLUE]),
                                         legend=alt.Legend(orient="bottom", title=None)),
                         tooltip=["concept", "period", alt.Tooltip("share:Q", format=".0%")]))
                    .properties(height=22 * len(df) + 40), width="stretch")

    st.markdown("#### Research response to stress episodes")
    ev = T["publication_events"].assign(year=lambda d: d.doc_id.map(PUB.year),
                                        event=lambda d: d.event_id.map(name))
    per = ev.groupby(["event", "year"]).size().reset_index(name="publications")
    st.altair_chart(alt.Chart(per).mark_line(point=True).encode(
        x=alt.X("year:O", title=None), y=alt.Y("publications:Q"),
        color=alt.Color("event:N", legend=alt.Legend(orient="bottom", title=None, columns=2)),
        tooltip=["event", "year", "publications"]).properties(height=300), width="stretch")


def publications():
    st.title("Publications")
    c1, c2, c3 = st.columns([3, 2, 2])
    with c1:
        text = st.text_input("Search titles and summaries")
    with c2:
        series = st.multiselect("Series", SERIES_ORDER)
    with c3:
        years = st.slider("Years", 2015, 2026, (2015, 2026))
    df = PUBS[PUBS.year.between(*years)]
    if series:
        df = df[df.series.isin(series)]
    if text:
        hit = df.title.str.contains(text, case=False) | df.summary.fillna("").str.contains(text, case=False)
        df = df[hit]
    df = df.assign(claims=df.doc_id.map(CLAIMS.doc_id.value_counts()).fillna(0).astype(int))
    st.caption(f"{len(df)} publications")
    st.dataframe(df.sort_values("date", ascending=False)[["date", "title", "series", "claims", "url"]],
                 hide_index=True, width="stretch", height=320, column_config={"date": DATE, "url": LINK})
    if df.empty:
        return
    pick = st.selectbox("Open a publication", df.sort_values("date", ascending=False).doc_id,
                        format_func=lambda d: f"{PUB.title[d]} ({PUB.date[d]:%Y})")
    p = PUB.loc[pick]
    st.subheader(p.title)
    st.caption(f"{p.series} · {p.date:%d %B %Y} · {p.authors or 'authors not recorded'} · [open]({p.url})")
    if p.summary:
        st.markdown(p.summary)
    conc = PC[(PC.doc_id == pick) & (PC.salience >= 0.6)].sort_values("salience", ascending=False)
    st.markdown("**Concepts:** " + ", ".join(
        f"{name(c)}{' (central)' if s >= 1 else ''}" for c, s in zip(conc.concept_id, conc.salience)))
    cl = CLAIMS[CLAIMS.doc_id == pick]
    if len(cl):
        st.markdown(f"**{len(cl)} causal claims**")
        claims_table(cl)
    q = T["open_questions"][T["open_questions"].doc_id == pick]
    if len(q):
        st.markdown("**Open questions**")
        for x in q.question:
            st.markdown(f"- {x}")


def downloads():
    st.title("Download the data")
    st.markdown("The whole graph as an Excel workbook (one sheet per table, with a README sheet explaining each "
                "one) or as CSV files. Full publication text is not included; each row links to the source.")
    if XLSX.exists():
        st.download_button("Excel workbook (.xlsx)", XLSX.read_bytes(), XLSX.name,
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
    if CSV_DIR.exists():
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for f in sorted(CSV_DIR.glob("*.csv")):
                z.write(f, f.name)
        st.download_button("CSV files (.zip)", buf.getvalue(), "BoE_MBF_KG_csv.zip", "application/zip")
    st.markdown("**Tables**")
    st.dataframe(pd.DataFrame([(k, len(v)) for k, v in T.items()], columns=["table", "rows"]), hide_index=True)


pages = [
    st.Page(overview, title="Overview", icon=":material/home:", default=True),
    st.Page(causal_map, title="Causal map", icon=":material/hub:", url_path="map"),
    st.Page(evidence, title="Grade the evidence", icon=":material/fact_check:", url_path="evidence"),
    st.Page(foresight, title="Audit foresight", icon=":material/history:", url_path="foresight"),
    st.Page(scenario, title="Trace a scenario", icon=":material/route:", url_path="scenario"),
    st.Page(briefing, title="Briefing pack", icon=":material/description:", url_path="briefing"),
    st.Page(disagreements, title="Where findings differ", icon=":material/compare_arrows:", url_path="differ"),
    st.Page(gaps, title="Research gaps", icon=":material/grid_on:", url_path="gaps"),
    st.Page(agenda, title="Research agenda", icon=":material/trending_up:", url_path="agenda"),
    st.Page(publications, title="Publications", icon=":material/library_books:", url_path="publications"),
    st.Page(downloads, title="Download data", icon=":material/download:", url_path="download"),
]
st.navigation(pages).run()
st.sidebar.caption("Bank of England staff publications on market-based finance, 2015–2026. "
                   "Claims extracted by a language model; see Overview → Known limitations.")
