from __future__ import annotations

import json
import base64
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title="AUiX Network Map", layout="wide", initial_sidebar_state="collapsed")

DATA_FILE = Path(__file__).with_name("Streamlit.xlsx")
AUIX_LOGO_FILE = Path(__file__).with_name("auix_logo.png")
DRADIS_LOGO_FILE = Path(__file__).with_name("dradis_logo.png")

CATEGORY_STYLE = {
    "Air University": {"color": "#e32119", "label": "AIR UNIVERSITY"},
    "Academia": {"color": "#0a55d5", "label": "ACADEMIA"},
    "Industry": {"color": "#ff9800", "label": "INDUSTRY"},
    "MIL & GOV": {"color": "#57a52c", "label": "MIL & GOV"},
}


def normalize_category(value: object) -> str:
    text = str(value).strip()
    aliases = {
        "Air Univerity": "Air University",
        "Air university": "Air University",
        "Mil & Gov": "MIL & GOV",
        "MIL&GOV": "MIL & GOV",
        "MIL & Gov": "MIL & GOV",
    }
    return aliases.get(text, text)


@st.cache_data(show_spinner=False)
def load_data(file_path: str, file_mtime_ns: int, file_size: int) -> pd.DataFrame:
    data = pd.read_excel(file_path)
    data.columns = [str(column).strip().lower() for column in data.columns]

    aliases = {
        "engagament": "engagement",
        "categories": "type",
        "category": "type",
        "expertise areas": "expertise",
        "engagement summary": "summary",
    }
    for source, target in aliases.items():
        if source in data.columns and target not in data.columns:
            data = data.rename(columns={source: target})

    required = {"name", "type", "engagement"}
    missing = required.difference(data.columns)
    if missing:
        raise ValueError(f"Missing required columns: {', '.join(sorted(missing))}")

    for optional in ("summary", "expertise"):
        if optional not in data.columns:
            data[optional] = "Not provided"

    data = data.dropna(subset=["name", "type"]).copy()
    data["name"] = data["name"].astype(str).str.strip()
    data["type"] = data["type"].map(normalize_category)
    data["engagement"] = pd.to_numeric(data["engagement"], errors="coerce").fillna(0)
    data["summary"] = data["summary"].fillna("Not provided").astype(str).str.strip()
    data["expertise"] = data["expertise"].fillna("Not provided").astype(str).str.strip()
    return data[data["type"].isin(CATEGORY_STYLE)].copy()


st.markdown(
    """
<style>
html, body, [data-testid="stAppViewContainer"], .stApp { background:#031630 !important; }
header[data-testid="stHeader"] { display:none; }
[data-testid="stToolbar"] { display:none; }
.block-container { padding:0 !important; max-width:none !important; }
iframe { display:block; width:100%; border:0; }
</style>
""",
    unsafe_allow_html=True,
)

try:
    stat = DATA_FILE.stat()
    df = load_data(str(DATA_FILE), stat.st_mtime_ns, stat.st_size)
except Exception as exc:
    st.error(f"The spreadsheet could not be loaded: {exc}")
    st.stop()

records = []
for row in df.to_dict(orient="records"):
    records.append(
        {
            "name": str(row["name"]),
            "type": str(row["type"]),
            "engagement": float(row["engagement"]),
            "summary": str(row["summary"]),
            "expertise": str(row["expertise"]),
        }
    )

payload = json.dumps(records, ensure_ascii=False).replace("</", "<\\/")
styles_payload = json.dumps(CATEGORY_STYLE, ensure_ascii=False).replace("</", "<\\/")

def image_data_uri(path: Path) -> str:
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:image/png;base64,{encoded}"

auix_logo_uri = image_data_uri(AUIX_LOGO_FILE)
dradis_logo_uri = image_data_uri(DRADIS_LOGO_FILE)

html = r'''<!doctype html>
<html>
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no" />
<style>
:root { --bg:#031630; --panel:#0a2448; --line:rgba(255,255,255,.74); --muted:#aac9f5; }
* { box-sizing:border-box; }
html, body { margin:0; width:100%; height:100%; min-height:100%; overflow:hidden; background:var(--bg); color:white; font-family:Arial, Helvetica, sans-serif; }
#app { width:100vw; height:100vh; height:100dvh; display:grid; grid-template-columns:minmax(0, 1fr) 340px; background:radial-gradient(circle at 42% 42%, #0b2a52 0%, #031630 62%, #020d1f 100%); }
#mapWrap { min-width:0; height:100vh; height:100dvh; position:relative; overflow:hidden; }
#brandBar { position:absolute; z-index:12; top:8px; left:50%; transform:translateX(-50%); display:flex; align-items:center; justify-content:center; gap:16px; height:58px; pointer-events:none; }
#brandBar img { display:block; object-fit:contain; max-height:54px; width:auto; }
#brandBar .auixLogo { max-width:92px; }
#brandBar .dradisLogo { max-width:68px; }
#viewNote { position:absolute; z-index:12; top:118px; left:50%; transform:translateX(-50%); color:#cbdcf6; font-size:12px; font-weight:700; letter-spacing:.02em; white-space:nowrap; pointer-events:none; }
#map { width:100%; height:100%; display:block; }
#panel { height:100vh; height:100dvh; overflow-y:auto; padding:88px 18px 18px; border-left:1px solid rgba(255,255,255,.18); background:rgba(2,13,31,.78); }
#panelTitle { position:absolute; top:22px; right:18px; width:304px; color:#d8e8ff; font-weight:800; font-size:14px; letter-spacing:.02em; }
.empty { background:rgba(10,36,72,.94); border:1px solid rgba(255,255,255,.3); border-radius:16px; padding:18px; color:#c8daf4; line-height:1.45; }
.card { position:relative; background:rgba(10,36,72,.96); border:1px solid rgba(255,255,255,.32); border-radius:16px; padding:17px 42px 17px 18px; margin-bottom:12px; box-shadow:0 8px 28px rgba(0,0,0,.3); }
.card h3 { margin:0 0 12px; font-size:22px; }
.close { position:absolute; right:12px; top:10px; width:26px; height:26px; border:0; border-radius:50%; background:white; color:#0a2448; font-weight:900; font-size:16px; cursor:pointer; }
.label { color:var(--muted); font-weight:800; font-size:13px; margin-top:9px; }
.value { font-size:15px; margin-top:2px; overflow-wrap:anywhere; }
#tooltip { position:absolute; z-index:20; min-width:190px; max-width:280px; pointer-events:none; opacity:0; transform:translate(-50%, calc(-100% - 16px)); background:#10284b; border:1px solid rgba(255,255,255,.38); border-radius:10px; padding:10px 12px; box-shadow:0 8px 24px rgba(0,0,0,.35); font-size:13px; line-height:1.35; transition:opacity .08s; }
#tooltip b { font-size:15px; }
.title { font-weight:900; fill:white; text-anchor:middle; }
.subtitle { fill:#d9e7ff; text-anchor:middle; font-size:12px; }
.edge { stroke:var(--line); stroke-linecap:round; }
.node { cursor:pointer; }
.node circle { stroke:white; stroke-width:2; transition:filter .12s, stroke-width .12s; }
.node:hover circle { filter:brightness(1.12); stroke-width:3; }
.node.selected circle { stroke:#ffe66d; stroke-width:5; }
.catText, .centerText { fill:white; font-weight:900; text-anchor:middle; dominant-baseline:middle; pointer-events:none; }
.orgText { fill:white; font-weight:800; font-size:13px; dominant-baseline:middle; pointer-events:none; paint-order:stroke; stroke:#031630; stroke-width:4px; stroke-linejoin:round; }
.mobileTabs { display:none; }
#mobileContent { display:none; }

@media (max-width: 700px) {
  html, body { height:auto; min-height:100%; overflow:visible; }
  #app { display:block; width:100%; height:auto; min-height:100vh; }
  #mapWrap { height:auto; min-height:100vh; overflow:visible; padding:0; }
  #map { display:none !important; }
  #panel { display:none !important; }
  #mobileContent { display:block; position:relative; width:100%; height:auto; min-height:100vh; overflow:visible; margin:0; padding:calc(14px + env(safe-area-inset-top)) 10px calc(100px + env(safe-area-inset-bottom)); touch-action:pan-y; background:radial-gradient(circle at 50% 24%, #0b2a52 0%, #031630 68%, #020d1f 100%); }
  #brandBar { display:none; }
  #viewNote { display:none; }
  .mobile-brand { display:flex; align-items:center; justify-content:center; gap:12px; margin:0 0 8px; }
  .mobile-brand img { height:42px; width:auto; object-fit:contain; }
  .mobile-brand .auixLogo { max-width:78px; }
  .mobile-brand .dradisLogo { max-width:58px; }
  .mobile-device-note { text-align:center; color:#cbdcf6; font-size:11px; font-weight:700; margin:5px 0 8px; }
  .mobile-header { text-align:center; margin:0 0 6px; }
  .mobile-header h1 { margin:0; font-size:27px; line-height:1.08; font-weight:900; }
  .mobile-header p { display:none; }
  .mobile-tabs { display:grid; grid-template-columns:1fr 1fr; gap:9px; margin:8px 0 0; }
  .mobile-tab { min-height:64px; border:2px solid white; border-radius:19px; color:white; font-weight:900; font-size:17px; line-height:1.05; padding:8px; cursor:pointer; }
  .mobile-tab.active { outline:4px solid rgba(255,230,109,.95); outline-offset:1px; }
  .mobile-help { display:none; }
  .mobile-visual { position:relative; width:100%; margin:-2px auto 0; }
  .mobile-network-svg { width:100%; height:auto; display:block; overflow:visible; }
  .mobile-link { stroke:rgba(255,255,255,.72); stroke-width:2; }
  .mobile-hub-label { fill:white; font-weight:900; text-anchor:middle; dominant-baseline:middle; font-size:21px; pointer-events:none; }
  .mobile-org-label { fill:white; font-weight:800; dominant-baseline:middle; font-size:12px; paint-order:stroke; stroke:#031630; stroke-width:4px; stroke-linejoin:round; pointer-events:none; }
  .mobile-org-node { cursor:pointer; }
  .mobile-org-node circle { stroke:white; stroke-width:3; }
  .mobile-org-node.selected circle { stroke:#ffe66d; stroke-width:6; }
  .mobile-detail { position:relative; background:rgba(10,36,72,.98); border:1px solid rgba(255,255,255,.38); border-radius:14px; padding:13px 40px 13px 14px; margin:8px 4px 4px; box-shadow:0 8px 24px rgba(0,0,0,.35); }
  .mobile-detail h3 { margin:0 0 8px; font-size:19px; }
  .mobile-detail-grid { display:grid; grid-template-columns:96px 1fr; gap:5px 8px; font-size:13px; line-height:1.3; }
  .mobile-detail-grid b { color:var(--muted); }
  #panelTitle { display:none; }
  #tooltip { display:none !important; }
}
</style>
</head>
<body>
<div id="app">
  <div id="mapWrap">
    <div id="brandBar"><img class="auixLogo" src="__AUIX_LOGO__" alt="AUiX logo"><img class="dradisLogo" src="__DRADIS_LOGO__" alt="DRADIS logo"></div>
    <div id="viewNote">Best viewed on desktop or tablet.</div>
    <svg id="map" role="img" aria-label="Interactive AUiX network map"></svg>
    <div id="tooltip"></div>
    <div id="mobileContent"></div>
  </div>
  <aside id="panel">
    <div id="panelTitle">Pinned organization details</div>
    <div id="cards"></div>
  </aside>
</div>
<script>
const DATA = __DATA__;
const STYLES = __STYLES__;
const categories = ["Air University", "Academia", "Industry", "MIL & GOV"];
const svg = document.getElementById('map');
const wrap = document.getElementById('mapWrap');
const tooltip = document.getElementById('tooltip');
const cards = document.getElementById('cards');
const panel = document.getElementById('panel');
let expanded = null;
let pinned = [];
let touchTipTimer = null;
let mobileSelected = null;
const mobileContent = document.getElementById('mobileContent');

function esc(v) { return String(v ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
function isMobile() { return window.matchMedia('(max-width: 700px)').matches; }
function make(tag, attrs={}) { const el=document.createElementNS('http://www.w3.org/2000/svg', tag); Object.entries(attrs).forEach(([k,v])=>el.setAttribute(k,v)); return el; }
function line(x1,y1,x2,y2,w=2) { return make('line',{x1,y1,x2,y2,'class':'edge','stroke-width':w}); }
function text(x,y,value,cls,size,anchor='middle') { const t=make('text',{x,y,'class':cls,'font-size':size,'text-anchor':anchor}); t.textContent=value; return t; }
function categoryData(cat) { return DATA.filter(d=>d.type===cat).sort((a,b)=>b.engagement-a.engagement || a.name.localeCompare(b.name)); }

function layout() {
  if (isMobile()) {
    svg.style.display='none';
    renderMobileApp();
    return;
  }
  svg.style.display='block';
  mobileContent.innerHTML='';
  const w = wrap.clientWidth, h = wrap.clientHeight;
  svg.setAttribute('viewBox', `0 0 ${w} ${h}`);
  while (svg.firstChild) svg.removeChild(svg.firstChild);
  const titleY = 92;
  svg.appendChild(text(w/2,titleY,'2025–2026 AUiX Network Map','title',34));
  svg.appendChild(text(w/2,titleY+24,'Click a category to open it. Hover for details. Click organizations to pin cards.','subtitle',12));
  drawDesktop(w,h);
  renderCards();
}

function drawDesktop(w,h) {
  const panelSafe = 0;
  const cx=w*0.50, cy=h*0.57;
  const xGap=Math.min(w*0.31, 390), yGap=Math.min(h*0.29, 250);
  const pos={
    'Air University':[cx-xGap,cy-yGap], 'Academia':[cx+xGap,cy-yGap],
    'Industry':[cx-xGap,cy+yGap], 'MIL & GOV':[cx+xGap,cy+yGap]
  };
  // Lower categories need extra room beneath them when their organization ring opens.
  // Move only the expanded lower hub upward; the normal four-hub view stays unchanged.
  if(expanded==='Industry' || expanded==='MIL & GOV') {
    pos[expanded][1] = Math.min(pos[expanded][1], h*0.68);
  }
  categories.forEach(cat=>{ const [x,y]=pos[cat]; svg.appendChild(line(cx,cy,x,y,3)); });
  drawHub(cx,cy,76,'#f4c542','AUiX','center');
  categories.forEach(cat=>{ const [x,y]=pos[cat]; drawHub(x,y,78,STYLES[cat].color,STYLES[cat].label,'category',cat); });
  if (expanded) drawOrganizationsDesktop(expanded,pos[expanded][0],pos[expanded][1],w,h);
}

function drawHub(x,y,r,color,label,type,cat=null) {
  const g=make('g',{'class':'node','data-type':type});
  g.appendChild(make('circle',{cx:x,cy:y,r,fill:color}));
  const centerSize = type==='center' ? 28 : 18;
  const longSize = type==='center' ? 28 : 17;
  if (label==='AIR UNIVERSITY' || label==='MIL & GOV') {
    const l1=label==='AIR UNIVERSITY'?'AIR':'MIL &';
    const l2=label==='AIR UNIVERSITY'?'UNIVERSITY':'GOV';
    g.appendChild(text(x,y-10,l1,'catText',label==='AIR UNIVERSITY' ? longSize : centerSize));
    g.appendChild(text(x,y+14,l2,'catText',label==='AIR UNIVERSITY' ? longSize : centerSize));
  } else {
    const size = label.length > 10 ? 17 : centerSize;
    g.appendChild(text(x,y,label,type==='center'?'centerText':'catText',size));
  }
  g.addEventListener('click',()=>{
    if(type==='center'){ expanded=null; layout(); return; }
    expanded = expanded===cat ? null : cat; layout();
  });
  g.addEventListener('mouseenter',e=>showTip(e, {name:cat||'AUiX', type:cat||'Center', engagement:'', summary:type==='category'?(expanded===cat?'Click to hide organizations':'Click to show organizations'):'Click to collapse map', expertise:''}));
  g.addEventListener('mouseleave',hideTip);
  svg.appendChild(g);
}

function drawOrganizationsDesktop(cat,hx,hy,w,h) {
  const items=categoryData(cat), n=items.length;
  if(!n) return;

  // Reserve room for labels on every side, then size the ellipse to the space that
  // actually exists around the selected category. This prevents the lower and left
  // Industry bubbles from being clipped by the map boundary.
  const sideLabelRoom = cat==='Industry' ? 165 : 135;
  const verticalRoom = 58;
  const maxRx = Math.max(120, Math.min(hx-sideLabelRoom, w-hx-sideLabelRoom));
  const maxRy = Math.max(105, Math.min(hy-verticalRoom, h-hy-verticalRoom));

  let desiredRx=Math.max(175,n*18);
  let desiredRy=Math.max(135,n*10);
  if(cat==='Industry') {
    desiredRx=Math.max(255,n*20);
    desiredRy=Math.max(145,n*10);
  }
  const rx=Math.min(desiredRx, w*0.25, maxRx);
  const ry=Math.min(desiredRy, h*0.23, maxRy);

  const max=Math.max(...items.map(d=>d.engagement),1);
  items.forEach((d,i)=>{
    const a=-Math.PI/2 + i*(Math.PI*2/n);
    const x=hx+rx*Math.cos(a), y=hy+ry*Math.sin(a);
    const r=17+17*Math.sqrt(d.engagement/max);
    svg.insertBefore(line(hx,hy,x,y,1.2+Math.min(d.engagement,60)/35), svg.querySelector('.node'));
    drawOrg(d,x,y,r,a,w,h);
  });
}

function setStreamlitHeight() {
  if (!isMobile()) return;
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      const contentHeight = Math.max(
        document.documentElement.scrollHeight,
        document.body.scrollHeight,
        mobileContent.scrollHeight
      ) + 24;
      window.parent.postMessage({
        isStreamlitMessage: true,
        type: 'streamlit:setFrameHeight',
        height: contentHeight
      }, '*');
    });
  });
}

function renderMobileApp() {
  mobileContent.innerHTML='';

  const header=document.createElement('div');
  const mobileBrand=document.createElement('div');
  mobileBrand.className='mobile-brand';
  mobileBrand.innerHTML='<img class="auixLogo" src="__AUIX_LOGO__" alt="AUiX logo"><img class="dradisLogo" src="__DRADIS_LOGO__" alt="DRADIS logo">';
  mobileContent.appendChild(mobileBrand);

  header.className='mobile-header';
  header.innerHTML='<h1>2025–2026 AUiX Network Map</h1><div class="mobile-device-note">Best viewed on desktop or tablet.</div>'; 
  mobileContent.appendChild(header);

  const tabs=document.createElement('div');
  tabs.className='mobile-tabs';
  categories.forEach(cat=>{
    const btn=document.createElement('button');
    btn.type='button';
    btn.className='mobile-tab'+(expanded===cat?' active':'');
    btn.style.background=STYLES[cat].color;
    btn.innerHTML=STYLES[cat].label==='AIR UNIVERSITY'?'AIR<br>UNIVERSITY':esc(STYLES[cat].label);
    btn.addEventListener('click',()=>{
      expanded=expanded===cat?null:cat;
      mobileSelected=null;
      renderMobileApp();
      window.scrollTo({top:0,behavior:'smooth'});
    });
    tabs.appendChild(btn);
  });
  mobileContent.appendChild(tabs);

  if(!expanded){
    setStreamlitHeight();
    return;
  }

  if(mobileSelected) renderMobileDetail(mobileSelected);

  const items=categoryData(expanded);
  const n=items.length;
  const visual=document.createElement('div');
  visual.className='mobile-visual';
  const NS='http://www.w3.org/2000/svg';
  const graph=document.createElementNS(NS,'svg');
  graph.setAttribute('class','mobile-network-svg');

  // Use two balanced columns around the category hub. This keeps every bubble
  // evenly spaced and gives each organization name its own dedicated row.
  const leftItems=[];
  const rightItems=[];
  items.forEach((item,index)=>{
    (index % 2 === 0 ? leftItems : rightItems).push(item);
  });
  const rowCount=Math.max(leftItems.length,rightItems.length,1);
  const topY=84;
  const rowGap=58;
  const bottomY=topY+(rowCount-1)*rowGap;
  const graphHeight=Math.max(450,bottomY+100);
  const cx=200;
  const cy=(topY+bottomY)/2;
  const hubR=56;
  const leftX=100;
  const rightX=300;
  const maxEng=Math.max(...items.map(d=>Number(d.engagement)||0),1);
  function orgRadius(d){
    const value = Math.max(0, Number(d.engagement)||0);
    return 14 + 9*Math.sqrt(value/maxEng);
  }

  graph.setAttribute('viewBox',`0 0 400 ${graphHeight}`);
  graph.setAttribute('role','img');
  graph.setAttribute('aria-label',`${STYLES[expanded].label} organizations`);

  const hub=document.createElementNS(NS,'circle');
  hub.setAttribute('cx',cx); hub.setAttribute('cy',cy); hub.setAttribute('r',hubR);
  hub.setAttribute('fill',STYLES[expanded].color); hub.setAttribute('stroke','white'); hub.setAttribute('stroke-width','3');
  graph.appendChild(hub);

  const hubText=document.createElementNS(NS,'text');
  hubText.setAttribute('x',cx); hubText.setAttribute('y',cy); hubText.setAttribute('class','mobile-hub-label');
  const hubLabel = STYLES[expanded].label;
  if(hubLabel==='AIR UNIVERSITY' || hubLabel==='MIL & GOV'){
    const first=document.createElementNS(NS,'tspan'); first.setAttribute('x',cx); first.setAttribute('dy','-10'); first.textContent=expanded==='Air University'?'AIR':'MIL &';
    const second=document.createElementNS(NS,'tspan'); second.setAttribute('x',cx); second.setAttribute('dy','23'); second.textContent=expanded==='Air University'?'UNIVERSITY':'GOV';
    hubText.appendChild(first); hubText.appendChild(second);
  } else {
    hubText.textContent=hubLabel;
  }
  graph.appendChild(hubText);

  function addMobileOrganization(d,side,rowIndex){
    const x=side==='left'?leftX:rightX;
    const y=topY+rowIndex*rowGap;
    const r=orgRadius(d);
    const line=document.createElementNS(NS,'line');
    line.setAttribute('x1',cx); line.setAttribute('y1',cy); line.setAttribute('x2',x); line.setAttribute('y2',y); line.setAttribute('class','mobile-link');
    graph.insertBefore(line,hub);

    const g=document.createElementNS(NS,'g');
    g.setAttribute('class','mobile-org-node'+(mobileSelected && mobileSelected.name===d.name?' selected':''));
    const c=document.createElementNS(NS,'circle');
    c.setAttribute('cx',x); c.setAttribute('cy',y); c.setAttribute('r',r); c.setAttribute('fill',STYLES[d.type].color);
    g.appendChild(c);

    const label=document.createElementNS(NS,'text');
    const onLeft=side==='left';
    const tx=onLeft?x-r-10:x+r+10;
    label.setAttribute('x',tx); label.setAttribute('y',y); label.setAttribute('text-anchor',onLeft?'end':'start'); label.setAttribute('class','mobile-org-label');
    const words=String(d.name).split(' ');
    const lines=[]; let current='';
    words.forEach(word=>{ const next=(current+' '+word).trim(); if(next.length>17 && current){ lines.push(current); current=word; } else current=next; });
    if(current) lines.push(current);
    const shown=lines.slice(0,2);
    shown.forEach((part,j)=>{ const sp=document.createElementNS(NS,'tspan'); sp.setAttribute('x',tx); sp.setAttribute('dy',j===0?(shown.length>1?'-7':'0'):'15'); sp.textContent=part; label.appendChild(sp); });
    g.appendChild(label);
    g.addEventListener('click',()=>{ mobileSelected=d; renderMobileApp(); window.scrollTo({top:0,behavior:'smooth'}); });
    graph.appendChild(g);
  }

  leftItems.forEach((d,i)=>addMobileOrganization(d,'left',i));
  rightItems.forEach((d,i)=>addMobileOrganization(d,'right',i));
  visual.appendChild(graph);
  mobileContent.appendChild(visual);
  setStreamlitHeight();
}
function renderMobileDetail(d){
  const detail=document.createElement('div');
  detail.className='mobile-detail';
  detail.innerHTML=`<button class="close" aria-label="Close details">×</button><h3>${esc(d.name)}</h3><div class="mobile-detail-grid"><b>Category</b><span>${esc(d.type)}</span><b>Engagements</b><span>${esc(Number(d.engagement).toLocaleString())}</span><b>Engagement Summary</b><span>${esc(d.summary)}</span><b>Expertise Areas</b><span>${esc(d.expertise)}</span></div>`;
  detail.querySelector('.close').addEventListener('click',()=>{mobileSelected=null; renderMobileApp();});
  mobileContent.appendChild(detail);
}

function drawOrg(d,x,y,r,a,w,h,mobile=false){
  const g=make('g',{'class':'node orgNode'+(pinned.some(p=>p.name===d.name)?' selected':''),'data-name':d.name});
  g.appendChild(make('circle',{cx:x,cy:y,r,fill:STYLES[d.type].color}));
  const cos=Math.cos(a), sin=Math.sin(a); let tx=x,ty=y,anchor='middle';
  const pad=r+8;
  if(cos>0.28){tx=x+pad;anchor='start';} else if(cos<-0.28){tx=x-pad;anchor='end';}
  else if(sin<0){ty=y-pad-4;} else {ty=y+pad+8;}
  // Keep names inside the visible map region.
  if(anchor==='start' && tx>w-145){anchor='end';tx=x-pad;}
  if(anchor==='end' && tx<145){anchor='start';tx=x+pad;}
  const t=text(tx,ty,d.name,'orgText',mobile?11:13,anchor);
  g.appendChild(t);
  g.addEventListener('mouseenter',e=>showTip(e,d));
  g.addEventListener('mousemove',moveTip);
  g.addEventListener('mouseleave',hideTip);
  g.addEventListener('touchstart',e=>{showTipAtNode(g,d); clearTimeout(touchTipTimer); touchTipTimer=setTimeout(hideTip,2300);},{passive:true});
  g.addEventListener('click',()=>{ pin(d); });
  svg.appendChild(g);
}

function tipHtml(d){ return `<b>${esc(d.name)}</b><br>Category: ${esc(d.type)}${d.engagement!==''?`<br>Engagements: ${esc(Number(d.engagement).toLocaleString())}`:''}<br>Engagement Summary: ${esc(d.summary)}${d.expertise?`<br>Expertise Areas: ${esc(d.expertise)}`:''}`; }
function showTip(e,d){ tooltip.innerHTML=tipHtml(d); tooltip.style.opacity='1'; moveTip(e); }
function moveTip(e){ const r=wrap.getBoundingClientRect(); tooltip.style.left=(e.clientX-r.left)+'px'; tooltip.style.top=(e.clientY-r.top)+'px'; }
function showTipAtNode(g,d){ const c=g.querySelector('circle'); if(!c)return; const p=svg.createSVGPoint(); p.x=parseFloat(c.getAttribute('cx')); p.y=parseFloat(c.getAttribute('cy')); const sp=p.matrixTransform(svg.getScreenCTM()); const r=wrap.getBoundingClientRect(); tooltip.innerHTML=tipHtml(d); tooltip.style.left=(sp.x-r.left)+'px'; tooltip.style.top=(sp.y-r.top)+'px'; tooltip.style.opacity='1'; }
function hideTip(){ tooltip.style.opacity='0'; }

function pin(d){ if(isMobile()){ mobileSelected=d; renderMobileApp(); window.scrollTo({top:0, behavior:'smooth'}); return; } if(!pinned.some(p=>p.name===d.name)){ pinned.push(d); if(pinned.length>4)pinned.shift(); } renderCards(); updateSelection(); }
function closePin(name){ pinned=pinned.filter(p=>p.name!==name); renderCards(); updateSelection(); }
function updateSelection(){ document.querySelectorAll('.orgNode').forEach(g=>g.classList.toggle('selected',pinned.some(p=>p.name===g.dataset.name))); }
function renderCards(){
  cards.innerHTML='';
  if(isMobile()){ panel.classList.remove('hasCards'); return; }
  if(!pinned.length){ cards.innerHTML='<div class="empty"><b>Organization details</b><br><br>Hover over a bubble for a quick preview. Click a bubble to pin its details here. Pinned cards stay open until you close them with ×.</div>'; panel.classList.remove('hasCards'); return; }
  panel.classList.add('hasCards');
  pinned.forEach(d=>{
    const card=document.createElement('div'); card.className='card';
    card.innerHTML=`<button class="close" aria-label="Close ${esc(d.name)}">×</button><h3>${esc(d.name)}</h3><div class="label">Category</div><div class="value">${esc(d.type)}</div><div class="label">Engagements</div><div class="value">${esc(Number(d.engagement).toLocaleString())}</div><div class="label">Engagement Summary</div><div class="value">${esc(d.summary)}</div><div class="label">Expertise Areas</div><div class="value">${esc(d.expertise)}</div>`;
    card.querySelector('.close').addEventListener('click',()=>closePin(d.name)); cards.appendChild(card);
  });
}

let resizeTimer; window.addEventListener('resize',()=>{clearTimeout(resizeTimer);resizeTimer=setTimeout(layout,120);});
layout();
</script>
</body>
</html>'''.replace('__DATA__', payload).replace('__STYLES__', styles_payload).replace('__AUIX_LOGO__', auix_logo_uri).replace('__DRADIS_LOGO__', dradis_logo_uri)

components.html(html, height=1050, scrolling=True)
