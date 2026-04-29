from __future__ import annotations

import csv
import json
from collections import OrderedDict
from datetime import date, datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import openpyxl

BASE_DIR = Path(__file__).resolve().parent
SOURCE = BASE_DIR / "펀드 기준가.xlsx"
MAPPING = BASE_DIR / "mapping.xlsx"
MAPPING_CSV = BASE_DIR / "mapping.csv"
HANA_EMP = BASE_DIR / "Hana_EMP.xlsx"
PLOTLY_JS = BASE_DIR / "plotly-2.35.2.min.js"
KST = ZoneInfo("Asia/Seoul")
OUTPUT = BASE_DIR / f"fund_return_chart({datetime.now(KST):%y%m%d}).html"
WEB_OUTPUT = BASE_DIR / "docs" / "index.html"


def as_date_text(value):
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    return str(value)[:10]


def as_number(value):
    if value is None or value == "":
        return None
    try:
        return round(float(value), 8)
    except (TypeError, ValueError):
        return None


def rows_from_mapping():
    if MAPPING_CSV.exists():
        with MAPPING_CSV.open("r", encoding="utf-8-sig", newline="") as f:
            return list(csv.reader(f))
    wb = openpyxl.load_workbook(MAPPING, read_only=True, data_only=True)
    ws = wb["Data"] if "Data" in wb.sheetnames else wb.active
    return list(ws.iter_rows(values_only=True))


def load_mapping():
    rows = rows_from_mapping()
    headers = [str(v).strip() if v is not None else "" for v in rows[0]]
    pos = {h: i for i, h in enumerate(headers)}
    mapping = OrderedDict()
    for row in rows[1:]:
        if not row:
            continue
        code = str(row[pos["Code"]]).strip() if row[pos["Code"]] is not None else ""
        name = str(row[pos["Fund"]]).strip() if row[pos["Fund"]] is not None else ""
        group = str(row[pos["Group"]]).strip() if row[pos["Group"]] is not None else "기타"
        manager = str(row[pos["Manager"]]).strip() if row[pos["Manager"]] is not None else "기타"
        if not code or not name or group == "BM":
            continue
        mapping[code] = {
            "code": code,
            "name": name,
            "type": group,
            "manager": manager,
            "investDate": as_date_text(row[pos["LaunchDate"]]) if "LaunchDate" in pos and pos["LaunchDate"] < len(row) else "",
            "returnMode": "simple" if code.upper().startswith("EMP") else "compound",
        }
    return mapping


def append_emp(mapping, funds, series, all_dates):
    if not HANA_EMP.exists():
        return
    wb = openpyxl.load_workbook(HANA_EMP, read_only=True, data_only=True)
    ws = wb["일별"] if "일별" in wb.sheetnames else wb.active
    rows = ws.iter_rows(values_only=True)
    headers = list(next(rows))
    cols = [(i, str(v).strip()) for i, v in enumerate(headers[:5]) if v and str(v).strip() in mapping]
    prev = {code: None for _, code in cols}
    for row in rows:
        d = as_date_text(row[0])
        if not d:
            continue
        for i, code in cols:
            cum = as_number(row[i] if i < len(row) else None)
            if cum is None:
                continue
            if code not in funds:
                funds[code] = len(funds)
                series[str(funds[code])] = []
            nav = 1000 * (1 + cum)
            daily = 0 if prev[code] is None else (cum - prev[code]) * 100
            series[str(funds[code])].append([d, round(daily, 6), round(nav, 6), None, 0, round(cum * 100, 6), round(nav, 6)])
            prev[code] = cum
            all_dates.append(d)


def build_payload():
    mapping = load_mapping()
    wb = openpyxl.load_workbook(SOURCE, read_only=True, data_only=True)
    ws = wb["Data"] if "Data" in wb.sheetnames else wb.active
    funds = OrderedDict()
    series = {}
    bm_levels = OrderedDict()
    all_dates = []

    for row in ws.iter_rows(min_row=4, values_only=True):
        d = as_date_text(row[1])
        if not d:
            continue
        if d not in bm_levels:
            bm_levels[d] = {
                "BM_KOSPI": as_number(row[11]),
                "BM_KOSDAQ": as_number(row[12]),
                "BM_SPX": as_number(row[13]),
                "BM_NASDAQ": as_number(row[14]),
            }
        code = str(row[3]).strip() if row[3] is not None else ""
        info = mapping.get(code)
        if not info:
            continue
        if code not in funds:
            funds[code] = len(funds)
            series[str(funds[code])] = []
        cum = as_number(row[10])
        level = round(cum * 10 + 1000, 6) if cum is not None else None
        series[str(funds[code])].append([d, as_number(row[6]), as_number(row[7]), as_number(row[8]), as_number(row[9]), cum, level])
        all_dates.append(d)

    append_emp(mapping, funds, series, all_dates)

    fund_rows = []
    sorted_series = {}
    for code, old_id in funds.items():
        info = mapping[code]
        rows = sorted(series[str(old_id)], key=lambda r: r[0])
        new_id = len(fund_rows)
        fund_rows.append({"id": new_id, **info, "count": len(rows)})
        sorted_series[str(new_id)] = rows

    bm_names = {"BM_KOSPI": "KOSPI", "BM_KOSDAQ": "KOSDAQ", "BM_SPX": "S&P500", "BM_NASDAQ": "NASDAQ"}
    bms = []
    for code, name in bm_names.items():
        rows = []
        prev = first = None
        for d, levels in sorted(bm_levels.items()):
            level = levels.get(code)
            if level is None:
                continue
            if first is None:
                first = level
            daily = 0 if prev is None else (level / prev - 1) * 100
            cum = (level / first - 1) * 100
            rows.append([d, round(daily, 6), level, prev, 0, round(cum, 6), level])
            prev = level
        bms.append({"id": code, "name": name, "code": code, "type": "BM", "manager": "BM", "investDate": "", "returnMode": "compound", "isBm": True, "count": len(rows)})
        sorted_series[code] = rows

    return {
        "sourceFile": SOURCE.name,
        "generatedAt": datetime.now(KST).strftime("%Y-%m-%d %H:%M"),
        "dateMin": min(all_dates),
        "dateMax": max(all_dates),
        "funds": fund_rows,
        "bms": bms,
        "series": sorted_series,
    }


def build_html(payload):
    plotly = PLOTLY_JS.read_text(encoding="utf-8") if PLOTLY_JS.exists() else ""
    data = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    template = r'''<!doctype html>
<html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>펀드 일별 수익률 대시보드</title>
<style>
:root{--ink:#17212b;--muted:#657385;--line:#d9e0e7;--panel:#fff;--wash:#f5f7fa;--accent:#0f766e;--good:#047857;--bad:#dc2626}*{box-sizing:border-box}body{margin:0;font-family:"Segoe UI","Malgun Gothic",Arial,sans-serif;color:var(--ink);background:var(--wash)}header{padding:26px 32px 18px;background:#fff;border-bottom:1px solid var(--line)}h1{margin:0 0 8px;font-size:26px;letter-spacing:0}.source{color:var(--muted);font-size:13px}main{padding:22px 32px 32px}.shell{display:grid;grid-template-columns:320px minmax(0,1fr);gap:18px;align-items:start}.main{display:grid;gap:18px;min-width:0}.panel{background:var(--panel);border:1px solid var(--line);border-radius:8px}.sidebar{position:sticky;top:14px;max-height:calc(100vh - 28px);overflow:hidden;padding:14px}.side-title{margin:0 0 8px;font-size:14px}.bm-grid{display:grid;grid-template-columns:1fr 1fr;gap:6px;padding-bottom:10px;border-bottom:1px solid var(--line)}.bm-grid label{display:grid;grid-template-columns:18px 1fr;gap:6px;align-items:center;padding:7px 8px;border:1px solid #d8e0e8;border-radius:6px;background:#fff;font-size:13px;font-weight:700}.filter-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:8px}.filter{position:relative}.filter button,.action-row button,.period button,.chart-actions button,.table-actions button{height:30px;border:1px solid #c8d1dc;border-radius:6px;background:#fff;color:var(--ink);font-weight:700;cursor:pointer}.filter>button{width:100%;display:flex;justify-content:space-between;align-items:center;padding:0 10px}.filter>button:after{content:"▾";font-size:11px;color:var(--muted)}.menu{display:none;position:absolute;left:0;right:0;top:34px;z-index:20;max-height:420px;overflow:auto;background:#fff;border:1px solid #c8d1dc;border-radius:6px;box-shadow:0 12px 28px rgba(15,23,42,.14);padding:6px}.filter.open .menu{display:grid;gap:2px}.menu label{display:grid;grid-template-columns:18px 1fr;gap:7px;align-items:center;padding:7px 8px;border-radius:5px;font-size:13px}.menu label:hover{background:#f0f5f8}.action-row{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin:8px 0}.fund-list{min-height:260px;max-height:calc(100vh - 230px);overflow:auto;border:1px solid #c8d1dc;border-radius:6px;background:#fff;padding:6px}.fund-option{display:grid;grid-template-columns:18px 1fr;gap:8px;padding:7px 8px;border-radius:5px;font-size:13px;font-weight:750}.fund-option:hover{background:#f0f5f8}.fund-meta{display:block;margin-top:2px;color:var(--muted);font-size:11px;font-weight:600}.period{padding:12px 14px;display:grid;grid-template-columns:minmax(220px,1fr) minmax(220px,1fr) auto;gap:10px;align-items:end}.date-card{display:grid;gap:6px}.label-row{display:flex;gap:6px;align-items:center;color:#314052;font-size:12px;font-weight:750}.date-card input{height:38px;border:1px solid #c8d1dc;border-radius:6px;padding:0 10px;font-size:14px;font-weight:700}.quick{display:flex;gap:6px}.quick button{height:26px;padding:0 8px;font-size:12px}.metrics{display:grid;grid-template-columns:repeat(6,minmax(140px,1fr));overflow:hidden}.metric{padding:15px 16px;border-right:1px solid var(--line);min-height:86px}.metric:last-child{border-right:0}.metric span{display:block;color:var(--muted);font-size:12px;font-weight:700;margin-bottom:9px}.metric strong{font-size:22px;white-space:nowrap}.chart-panel{padding:16px}.panel-title{display:flex;justify-content:space-between;gap:12px;align-items:baseline;margin-bottom:10px}.title-note{color:var(--muted);font-size:12px}.chart-actions{display:flex;gap:10px;align-items:center;flex-wrap:wrap}.chart-actions button,.table-actions button{height:26px;padding:0 9px;font-size:12px}h2{margin:0;font-size:17px}#chart{height:468px;background:#fbfcfe;border:1px solid #e3e8ee;border-radius:6px}.chart-panel.collapsed #chart,.chart-panel.collapsed .legend{display:none}.legend{display:flex;flex-wrap:wrap;justify-content:center;gap:8px 14px;margin-top:11px}.legend button{display:inline-flex;gap:6px;align-items:center;border:1px solid transparent;border-radius:6px;background:transparent;padding:3px 6px;font-size:12px;font-weight:700}.legend button.active{background:#e8f3f1;border-color:var(--accent)}.legend button.dim{opacity:.28}.swatch{width:20px;height:3px;border-radius:2px}.table-panel{overflow:hidden}.table-head{display:flex;justify-content:space-between;align-items:baseline;padding:14px 16px;border-bottom:1px solid var(--line)}.table-wrap{max-height:430px;overflow:auto}table{width:100%;border-collapse:collapse;background:#fff;font-size:13px}th,td{padding:9px 10px;border-bottom:1px solid #edf1f5;text-align:right;white-space:nowrap}th{position:sticky;top:0;background:#eef3f7;color:#314052;font-weight:800;cursor:pointer}th:first-child,td:first-child,th:nth-child(2),td:nth-child(2),th:nth-child(3),td:nth-child(3){text-align:left}.pos{color:var(--good)}.neg{color:var(--bad)}.empty{height:100%;min-height:320px;display:flex;align-items:center;justify-content:center;color:#5b6b82;font-size:16px}.hidden{display:none}@media(max-width:900px){.shell{grid-template-columns:1fr}.period{grid-template-columns:1fr}.metrics{grid-template-columns:1fr 1fr}.sidebar{position:relative;max-height:none}}
</style><script>%%PLOTLY%%</script></head>
<body><header><h1>펀드 일별 수익률 대시보드</h1><div class="source">원본: %%SOURCE%% · 데이터 %%ROWS%%건 · 펀드 %%FUNDS%%개 · 기간 %%MIN%% ~ %%MAX%% · 생성 %%GEN%%</div></header>
<main><div class="shell"><aside class="panel sidebar"><h3 class="side-title">BM 선택</h3><div id="bmBox" class="bm-grid"></div><h3 class="side-title" style="margin-top:12px">펀드 선택</h3><div class="filter-grid"><div class="filter" id="typeFilter"><div class="label-row">펀드 유형</div><button type="button" id="typeBtn">전체</button><div class="menu" id="typeMenu"></div></div><div class="filter" id="mgrFilter"><div class="label-row">운용사</div><button type="button" id="mgrBtn">전체</button><div class="menu" id="mgrMenu"></div></div></div><div class="action-row"><button id="selectBtn">일괄 선택</button><button id="clearBtn">일괄 해제</button></div><div id="fundBox" class="fund-list"></div></aside>
<div class="main"><section class="panel period"><div class="date-card"><div class="label-row"><span>시작일</span><span class="quick"><button id="ytd">연초</button><button id="mtd">월초</button></span></div><input type="date" id="start"></div><div class="date-card"><div class="label-row"><span>종료일</span><span class="quick"><button id="allPeriod">기간 전체</button></span></div><input type="date" id="end"></div></section><section class="panel metrics" id="metrics"></section><section class="panel chart-panel" id="chartPanel"><div class="panel-title"><h2>누적수익률 비교 추이</h2><div class="chart-actions"><button id="toggleChart">차트 접기</button><button id="colors">색상 변경</button><button id="labels">레이블 표시</button><span class="title-note" id="chartHint"></span></div></div><div id="chart"></div><div id="legend" class="legend"></div></section><section class="panel table-panel"><div class="table-head"><div><h2>성과지표 <span class="title-note">(펀드 기준가로 계산한 성과로, 실제 집행금액 기준 성과와 일부 차이가 발생할 수 있습니다)</span></h2><div class="title-note" id="perfHint"></div></div><div class="table-actions"><button id="expandTable">확대</button><button id="downloadCsv">CSV</button></div></div><div class="table-wrap"><table><thead><tr id="perfHead"></tr></thead><tbody id="perfBody"></tbody></table></div></section></div></div></main>
<script>
const DATA=%%DATA%%;const typeOrder=['주식형','혼합형','메자닌','멀티','롱숏','공모주','EMP','Pre-IPO','채권형(원화)','채권형(외화)','기타'];const fundPal=[['#0f766e','#f97316','#84cc16','#ec4899','#06b6d4','#f59e0b','#2563eb','#9333ea'],['#0891b2','#ea580c','#16a34a','#db2777','#ca8a04','#0d9488','#1d4ed8','#7c3aed']];const bmPal=[{BM_KOSPI:'#111827',BM_KOSDAQ:'#b91c1c',BM_SPX:'#2563eb',BM_NASDAQ:'#7c3aed'},{BM_KOSPI:'#0f766e',BM_KOSDAQ:'#dc2626',BM_SPX:'#1d4ed8',BM_NASDAQ:'#c026d3'}];let selected=new Set(),selectedBm=new Set(),types=new Set(),mgrs=new Set(),hi=new Set(),pal=0,showLabels=false,sortKey='periodReturn',sortDir='desc',models=[];const $=id=>document.getElementById(id);const cols=[['name','펀드명'],['type','유형'],['investDate','투자일'],['periodReturn','기간수익률'],['annualReturn','연환산 수익률'],['vol','연환산 변동성'],['sharpe','Sharpe'],['mdd','MDD'],['ytd','YTD'],['mtd','MTD'],['d1','1일'],['w1','1주'],['m1','1개월'],['m3','3개월'],['m6','6개월'],['y1','1년']];
function esc(s){return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}function pct(v){return v==null||isNaN(v)?'-':(v*100).toFixed(2)+'%'}function pp(v){return v==null||isNaN(v)?'-':v.toFixed(2)+'%'}function cls(v){return v>0?'pos':v<0?'neg':''}function rows(id,s=$('start').value,e=$('end').value){if(s>e)[s,e]=[e,s];return(DATA.series[String(id)]||[]).filter(r=>r[0]>=s&&r[0]<=e)}function stat(rs,f){let lr=rs.filter(r=>r[6]!=null);if(lr.length<2)return{};let simple=f.returnMode==='simple',first=lr[0][6],last=lr.at(-1)[6],pr=simple?(last-first)/1000:last/first-1,rets=lr.slice(1).map((r,i)=>simple?(r[6]-lr[i][6])/1000:r[6]/lr[i][6]-1),avg=rets.reduce((a,b)=>a+b,0)/rets.length,vol=Math.sqrt(rets.reduce((s,v)=>s+(v-avg)**2,0)/Math.max(1,rets.length-1))*Math.sqrt(252),ann=simple?pr*252/rets.length:Math.pow(1+pr,252/rets.length)-1,peak=1,mdd=0;for(const r of lr){let nav=simple?1+(r[6]-first)/1000:r[6]/first;peak=Math.max(peak,nav);mdd=Math.min(mdd,nav/peak-1)}return{periodReturn:pr,annualReturn:ann,vol,sharpe:vol?ann/vol:null,mdd,obs:lr.length}}
function shifted(d,t,n){let x=new Date(d+'T00:00:00');if(t==='d')x.setDate(x.getDate()-n);if(t==='m')x.setMonth(x.getMonth()-n);if(t==='y')x.setFullYear(x.getFullYear()-n);return x.toISOString().slice(0,10)}function trail(f){let e=$('end').value,d=new Date(e+'T00:00:00'),rs={ytd:`${d.getFullYear()}-01-01`,mtd:`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-01`,d1:shifted(e,'d',1),w1:shifted(e,'d',7),m1:shifted(e,'m',1),m3:shifted(e,'m',3),m6:shifted(e,'m',6),y1:shifted(e,'y',1)};return Object.fromEntries(Object.entries(rs).map(([k,s])=>[k,stat(rows(f.id,s<DATA.dateMin?DATA.dateMin:s,e),f).periodReturn]))}
function filteredFunds(){return DATA.funds.filter(f=>(!types.size||types.has(f.type))&&(!mgrs.size||mgrs.has(f.manager)))}function renderFilter(menu,set,values,btn){menu.innerHTML=`<label><input type="checkbox" data-all="1" ${set.size===0?'checked':''}>전체</label>`+values.map(v=>`<label><input type="checkbox" value="${esc(v)}" ${set.has(v)?'checked':''}>${esc(v)}</label>`).join('');btn.textContent=set.size?`${set.size}개 선택`:'전체';menu.querySelectorAll('input').forEach(i=>i.onchange=()=>{if(i.dataset.all){set.clear()}else{i.checked?set.add(i.value):set.delete(i.value)}renderFilters();renderFunds();render()})}
function renderFilters(){let allTypes=[...new Set(DATA.funds.map(f=>f.type))].sort((a,b)=>(typeOrder.indexOf(a)<0?99:typeOrder.indexOf(a))-(typeOrder.indexOf(b)<0?99:typeOrder.indexOf(b)));let baseForMgr=DATA.funds.filter(f=>!types.size||types.has(f.type));let allMgrs=[...new Set(baseForMgr.map(f=>f.manager))].sort((a,b)=>a.localeCompare(b,'ko-KR'));mgrs=new Set([...mgrs].filter(m=>allMgrs.includes(m)));renderFilter($('typeMenu'),types,allTypes,$('typeBtn'));renderFilter($('mgrMenu'),mgrs,allMgrs,$('mgrBtn'))}function renderFunds(){$('fundBox').innerHTML=filteredFunds().map(f=>`<label class="fund-option"><input type="checkbox" value="${f.id}" ${selected.has(String(f.id))?'checked':''}><span>${esc(f.name)}<span class="fund-meta">${esc(f.type)} · ${esc(f.manager)} · ${esc(f.code)}</span></span></label>`).join('');$('fundBox').querySelectorAll('input').forEach(i=>i.onchange=()=>{i.checked?selected.add(i.value):selected.delete(i.value);render()})}
function init(){ $('start').value=DATA.dateMin;$('end').value=DATA.dateMax;$('start').min=DATA.dateMin;$('end').max=DATA.dateMax;$('bmBox').innerHTML=DATA.bms.map(b=>`<label><input type="checkbox" value="${b.id}">${esc(b.name)}</label>`).join('');$('bmBox').querySelectorAll('input').forEach(i=>i.onchange=()=>{i.checked?selectedBm.add(i.value):selectedBm.delete(i.value);render()});['typeFilter','mgrFilter'].forEach(id=>$(id).querySelector('button').onclick=()=>$(id).classList.toggle('open'));document.addEventListener('click',e=>{if(!e.target.closest('.filter'))document.querySelectorAll('.filter').forEach(f=>f.classList.remove('open'))});$('selectBtn').onclick=()=>{filteredFunds().forEach(f=>selected.add(String(f.id)));renderFunds();render()};$('clearBtn').onclick=()=>{filteredFunds().forEach(f=>selected.delete(String(f.id)));renderFunds();render()};$('ytd').onclick=()=>{let y=$('end').value.slice(0,4);$('start').value=Math.max(DATA.dateMin,`${y}-01-01`);render()};$('mtd').onclick=()=>{$('start').value=$('end').value.slice(0,8)+'01';render()};$('allPeriod').onclick=()=>{$('start').value=DATA.dateMin;$('end').value=DATA.dateMax;render()};$('labels').onclick=()=>{showLabels=!showLabels;$('labels').textContent=showLabels?'레이블 숨김':'레이블 표시';render()};$('colors').onclick=()=>{pal=(pal+1)%fundPal.length;render()};$('toggleChart').onclick=()=>{$('chartPanel').classList.toggle('collapsed');$('toggleChart').textContent=$('chartPanel').classList.contains('collapsed')?'차트 열기':'차트 접기'};$('start').onchange=render;$('end').onchange=render;$('perfHead').innerHTML=cols.map(c=>`<th data-key="${c[0]}">${c[1]}</th>`).join('');$('perfHead').querySelectorAll('th').forEach(th=>th.onclick=()=>{sortDir=sortKey===th.dataset.key&&sortDir==='desc'?'asc':'desc';sortKey=th.dataset.key;renderPerf()});$('downloadCsv').onclick=downloadCsv;renderFilters();renderFunds();render()}
function selectedItems(){return[...DATA.funds.filter(f=>selected.has(String(f.id))),...DATA.bms.filter(b=>selectedBm.has(String(b.id)))]}function render(){let items=selectedItems();models=items.map((f,i)=>{let fundIdx=DATA.funds.findIndex(x=>String(x.id)===String(f.id));let c=f.isBm?bmPal[pal][f.id]:fundPal[pal][(fundIdx<0?i:fundIdx)%fundPal[pal].length];return{f,rows:rows(f.id),color:c,stats:stat(rows(f.id),f),tr:trail(f)}});renderMetrics();renderChart();renderPerf()}
function renderMetrics(){let v=models.filter(m=>m.stats.obs),best=v.slice().sort((a,b)=>(b.stats.periodReturn??-9)-(a.stats.periodReturn??-9))[0],avg=k=>{let a=v.map(m=>m.stats[k]).filter(x=>x!=null&&!isNaN(x));return a.length?a.reduce((s,x)=>s+x,0)/a.length:null},bestS=v.slice().sort((a,b)=>(b.stats.sharpe??-9)-(a.stats.sharpe??-9))[0],worst=v.slice().sort((a,b)=>(a.stats.mdd??9)-(b.stats.mdd??9))[0];let cards=[['선택 펀드',models.length+'개'],['최고 기간수익률',best?pct(best.stats.periodReturn):'-',best?.stats.periodReturn],['평균 연환산 수익률',pct(avg('annualReturn')),avg('annualReturn')],['평균 연환산 변동성',pct(avg('vol')),avg('vol')],['최고 Sharpe',bestS?.stats.sharpe==null?'-':bestS.stats.sharpe.toFixed(2),bestS?.stats.sharpe],['최대 MDD',worst?pct(worst.stats.mdd):'-',worst?.stats.mdd]];$('metrics').innerHTML=cards.map(c=>`<div class="metric"><span>${c[0]}</span><strong class="${cls(c[2])}">${c[1]}</strong></div>`).join('')}
function chartPts(m){let r=m.rows.filter(x=>x[6]!=null);if(r.length<2)return[];let base=r[0][6];return r.map(x=>[x[0],m.f.returnMode==='simple'?(x[6]-base)/1000*100:(x[6]/base-1)*100])}function renderChart(){let ss=models.map(m=>({...m,pts:chartPts(m)})).filter(m=>m.pts.length);$('chartHint').textContent=ss.length?`시작점 0 기준 · ${ss.length}개 펀드 비교`:'선택 기간에 누적수익률 데이터가 없습니다.';if(!ss.length){Plotly.purge('chart');$('chart').innerHTML='<div class="empty">1개 이상의 펀드를 선택하세요</div>';$('legend').innerHTML='';return}let traces=ss.map(m=>{let on=hi.has(String(m.f.id)),dim=hi.size&&!on;return{x:m.pts.map(p=>p[0]),y:m.pts.map(p=>p[1]),type:'scatter',mode:'lines',name:m.f.name,opacity:dim?.18:1,line:{color:m.color,width:on?(m.f.isBm?6:4.2):(m.f.isBm?4.2:2.4),dash:'solid'},hovertemplate:`<b>${esc(m.f.name)}</b><br>기준일=%{x}<br>수익률=%{y:.2f}%<extra></extra>`}});let anns=showLabels?ss.map(m=>{let p=m.pts.at(-1),on=hi.has(String(m.f.id));return{order:on?1:0,a:{x:p[0],y:p[1],xref:'x',yref:'y',text:`${m.f.name.slice(0,12)} ${p[1].toFixed(2)}%`,showarrow:true,arrowcolor:'rgba(0,0,0,0)',ax:24,ay:0,xanchor:'left',font:{size:on?13:11,color:on?'#111827':m.color},bgcolor:on?'rgba(255,255,255,.98)':'rgba(255,255,255,.78)',bordercolor:m.color,borderwidth:on?2:1}}}).sort((a,b)=>a.order-b.order).map(x=>x.a):[];Plotly.react('chart',traces,{margin:{l:58,r:28,t:18,b:42},paper_bgcolor:'#fbfcfe',plot_bgcolor:'#fbfcfe',hovermode:'closest',showlegend:false,annotations:anns,xaxis:{showgrid:false,zeroline:false},yaxis:{ticksuffix:'%',gridcolor:'#e5eaf0',zerolinecolor:'#94a3b8'}},{responsive:true,displayModeBar:false});$('legend').innerHTML=ss.map(m=>{let id=String(m.f.id),a=hi.has(id),d=hi.size&&!a;return`<button class="${a?'active ':''}${d?'dim':''}" data-id="${id}"><span class="swatch" style="background:${m.color};${m.f.isBm?'height:5px':''}"></span>${esc(m.f.name)}</button>`}).join('');$('legend').querySelectorAll('button').forEach(b=>b.onclick=()=>{hi.has(b.dataset.id)?hi.delete(b.dataset.id):hi.add(b.dataset.id);renderChart()})}
function perfModels(){let items=selectedItems();if(!items.length)items=DATA.funds;return items.map((f,i)=>{let fundIdx=DATA.funds.findIndex(x=>String(x.id)===String(f.id));let c=f.isBm?bmPal[pal][f.id]:fundPal[pal][(fundIdx<0?i:fundIdx)%fundPal[pal].length];return{f,rows:rows(f.id),color:c,stats:stat(rows(f.id),f),tr:trail(f)}})}function val(m,k){if(k==='name')return m.f.name;if(k==='type')return m.f.type;if(k==='investDate')return m.f.investDate;if(['periodReturn','annualReturn','vol','sharpe','mdd'].includes(k))return m.stats[k];return m.tr[k]}function renderPerf(){let ms=perfModels();$('perfHint').textContent=`${ms.length}개 펀드${selectedItems().length?'':' · 선택 없음: 전체 표시'}`;ms.sort((a,b)=>{let av=val(a,sortKey),bv=val(b,sortKey),dir=sortDir==='asc'?1:-1;if(typeof av==='string'||typeof bv==='string')return String(av??'').localeCompare(String(bv??''),'ko-KR')*dir;return ((av??(sortDir==='asc'?Infinity:-Infinity))-(bv??(sortDir==='asc'?Infinity:-Infinity)))*dir});$('perfBody').innerHTML=ms.map(m=>`<tr><td>${esc(m.f.name)}</td><td>${esc(m.f.type||'')}</td><td>${esc(m.f.investDate||'')}</td><td class="${cls(m.stats.periodReturn)}">${pct(m.stats.periodReturn)}</td><td class="${cls(m.stats.annualReturn)}">${pct(m.stats.annualReturn)}</td><td>${pct(m.stats.vol)}</td><td class="${cls(m.stats.sharpe)}">${m.stats.sharpe==null?'-':m.stats.sharpe.toFixed(2)}</td><td class="${cls(m.stats.mdd)}">${pct(m.stats.mdd)}</td><td>${pct(m.tr.ytd)}</td><td>${pct(m.tr.mtd)}</td><td>${pct(m.tr.d1)}</td><td>${pct(m.tr.w1)}</td><td>${pct(m.tr.m1)}</td><td>${pct(m.tr.m3)}</td><td>${pct(m.tr.m6)}</td><td>${pct(m.tr.y1)}</td></tr>`).join('')}
function downloadCsv(){let header=cols.map(c=>c[1]);let rows=[...document.querySelectorAll('#perfBody tr')].map(tr=>[...tr.children].map(td=>td.textContent));let csv=[header,...rows].map(r=>r.map(v=>`"${String(v).replaceAll('"','""')}"`).join(',')).join('\n');let a=document.createElement('a');a.href=URL.createObjectURL(new Blob(['\ufeff'+csv],{type:'text/csv'}));a.download='performance_metrics.csv';a.click()}init();
</script></body></html>'''
    return (template.replace("%%PLOTLY%%", plotly)
        .replace("%%DATA%%", data)
        .replace("%%SOURCE%%", payload["sourceFile"])
        .replace("%%ROWS%%", f"{sum(len(v) for v in payload['series'].values()):,}")
        .replace("%%FUNDS%%", f"{len(payload['funds']):,}")
        .replace("%%MIN%%", payload["dateMin"])
        .replace("%%MAX%%", payload["dateMax"])
        .replace("%%GEN%%", payload["generatedAt"]))


if __name__ == "__main__":
    html = build_html(build_payload())
    OUTPUT.write_text(html, encoding="utf-8")
    WEB_OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    WEB_OUTPUT.write_text(html, encoding="utf-8")
    print(OUTPUT)
    print(WEB_OUTPUT)
