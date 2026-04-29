from __future__ import annotations

import json
from collections import OrderedDict
from datetime import date, datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import openpyxl

BASE_DIR = Path(__file__).resolve().parent
SOURCE = BASE_DIR / "펀드 기준가.xlsx"
MAPPING = BASE_DIR / "mapping.xlsx"
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


def load_mapping():
    wb = openpyxl.load_workbook(MAPPING, read_only=True, data_only=True)
    ws = wb["Data"] if "Data" in wb.sheetnames else wb.active
    rows = ws.iter_rows(values_only=True)
    headers = [str(v).strip() if v is not None else "" for v in next(rows)]
    pos = {h: i for i, h in enumerate(headers)}
    mapping = {}
    for row in rows:
        code = str(row[pos["Code"]]).strip() if row[pos["Code"]] is not None else ""
        name = str(row[pos["Fund"]]).strip() if row[pos["Fund"]] is not None else ""
        if not code or not name:
            continue
        mapping[code] = {
            "code": code,
            "name": name,
            "type": str(row[pos["Group"]]).strip() if row[pos["Group"]] is not None else "기타",
            "manager": str(row[pos["Manager"]]).strip() if row[pos["Manager"]] is not None else "기타",
            "investDate": as_date_text(row[pos["LaunchDate"]]) if "LaunchDate" in pos else "",
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
    return f'''<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>펀드 누적수익률 대시보드</title>
<style>body{{margin:0;font-family:Segoe UI,Malgun Gothic,sans-serif;background:#f5f7fa;color:#17212b}}header{{padding:22px 28px;background:#fff;border-bottom:1px solid #d9e0e7}}main{{display:grid;grid-template-columns:320px 1fr;gap:18px;padding:18px 28px}}aside,section{{background:#fff;border:1px solid #d9e0e7;border-radius:8px;padding:14px}}.list{{max-height:72vh;overflow:auto;border:1px solid #c8d1dc;border-radius:6px;padding:6px}}label{{display:block;margin:6px 0;font-size:13px}}select,input,button{{height:32px;border:1px solid #c8d1dc;border-radius:6px;background:#fff}}.row{{display:flex;gap:8px;align-items:end;flex-wrap:wrap}}#chart{{height:520px}}table{{width:100%;border-collapse:collapse;font-size:13px}}th,td{{padding:8px;border-bottom:1px solid #edf1f5;white-space:nowrap;text-align:right}}th:first-child,td:first-child,th:nth-child(2),td:nth-child(2),th:nth-child(3),td:nth-child(3){{text-align:left}}th{{background:#eef3f7;position:sticky;top:0}}.legend button{{margin:4px;padding:4px 8px;height:auto}}.active{{outline:2px solid #0f766e;background:#e8f3f1}}.dim{{opacity:.25}}.pos{{color:#047857}}.neg{{color:#dc2626}}.note{{font-size:11px;color:#657385}}@media(max-width:900px){{main{{grid-template-columns:1fr}}}}</style>
<script>{plotly}</script></head><body><header><h1>펀드 누적수익률 대시보드</h1><div class="note">원본: {payload['sourceFile']} · 기간 {payload['dateMin']} ~ {payload['dateMax']} · 생성 {payload['generatedAt']}</div></header>
<main><aside><h3>BM 선택</h3><div id="bmBox"></div><h3>펀드 선택</h3><div class="row"><select id="typeSel" multiple size="6"></select><select id="mgrSel" multiple size="6"></select></div><div class="row"><button id="selectBtn">일괄 선택</button><button id="clearBtn">일괄 해제</button></div><div id="fundBox" class="list"></div></aside>
<div><section><div class="row"><label>시작일<br><input type="date" id="start"></label><button id="ytd">연초</button><button id="mtd">월초</button><label>종료일<br><input type="date" id="end"></label><button id="allPeriod">기간 전체</button><button id="labels">레이블 표시</button><button id="colors">색상 변경</button></div></section>
<section><h2>누적수익률 비교 추이</h2><div id="chart"></div><div id="legend" class="legend"></div></section>
<section><h2>성과지표 <span class="note">(펀드 기준가로 계산한 성과로, 실제 집행금액 기준 성과와 일부 차이가 발생할 수 있습니다)</span></h2><div style="overflow:auto"><table><thead><tr><th>펀드명</th><th>유형</th><th>투자일</th><th>기간수익률</th><th>연환산 수익률</th><th>연환산 변동성</th><th>Sharpe</th><th>MDD</th><th>YTD</th><th>MTD</th><th>1일</th><th>1주</th><th>1개월</th><th>3개월</th><th>6개월</th><th>1년</th></tr></thead><tbody id="perf"></tbody></table></div></section></div></main>
<script>const DATA={data};const typeOrder=['주식형','혼합형','메자닌','멀티','롱숏','공모주','EMP','Pre-IPO','채권형(원화)','채권형(외화)','기타'];const fundColors=[['#0f766e','#f97316','#84cc16','#ec4899','#06b6d4','#f59e0b'],['#0891b2','#ea580c','#16a34a','#db2777','#ca8a04','#0d9488']];const bmColors=[{{BM_KOSPI:'#111827',BM_KOSDAQ:'#b91c1c',BM_SPX:'#2563eb',BM_NASDAQ:'#7c3aed'}},{{BM_KOSPI:'#0f766e',BM_KOSDAQ:'#dc2626',BM_SPX:'#1d4ed8',BM_NASDAQ:'#c026d3'}}];let selected=new Set(),selectedBm=new Set(),hi=new Set(),showLabels=false,pal=0;const $=id=>document.getElementById(id);function esc(s){{return String(s).replace(/[&<>"']/g,c=>({{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}}[c]))}}function vals(sel){{return [...sel.selectedOptions].map(o=>o.value)}}function pct(v){{return v==null||isNaN(v)?'-':(v*100).toFixed(2)+'%'}}function pp(v){{return v==null||isNaN(v)?'-':v.toFixed(2)+'%'}}function num(v){{return v==null||isNaN(v)?'-':v.toFixed(2)}}function rows(id){{let s=$('start').value,e=$('end').value;if(s>e)[s,e]=[e,s];return (DATA.series[String(id)]||[]).filter(r=>r[0]>=s&&r[0]<=e)}}function stat(rs,f){{let lr=rs.filter(r=>r[6]!=null);if(lr.length<2)return{{}};let simple=f.returnMode==='simple',first=lr[0][6],last=lr.at(-1)[6],pr=simple?(last-first)/1000:last/first-1,rets=lr.slice(1).map((r,i)=>simple?(r[6]-lr[i][6])/1000:r[6]/lr[i][6]-1),avg=rets.reduce((a,b)=>a+b,0)/rets.length,vol=Math.sqrt(rets.reduce((s,v)=>s+(v-avg)**2,0)/Math.max(1,rets.length-1))*Math.sqrt(252),ann=simple?pr*252/rets.length:Math.pow(1+pr,252/rets.length)-1,peak=1,mdd=0;for(const r of lr){{let nav=simple?1+(r[6]-first)/1000:r[6]/first;peak=Math.max(peak,nav);mdd=Math.min(mdd,nav/peak-1)}}return{{periodReturn:pr,annualReturn:ann,vol,sharpe:vol?ann/vol:null,mdd}}}}function shifted(d,typ,n){{let x=new Date(d+'T00:00:00');if(typ==='d')x.setDate(x.getDate()-n);if(typ==='m')x.setMonth(x.getMonth()-n);if(typ==='y')x.setFullYear(x.getFullYear()-n);return x.toISOString().slice(0,10)}}function trail(f){{let e=$('end').value,d=new Date(e+'T00:00:00'),ranges={{ytd:`${{d.getFullYear()}}-01-01`,mtd:`${{d.getFullYear()}}-${{String(d.getMonth()+1).padStart(2,'0')}}-01`,d1:shifted(e,'d',1),w1:shifted(e,'d',7),m1:shifted(e,'m',1),m3:shifted(e,'m',3),m6:shifted(e,'m',6),y1:shifted(e,'y',1)}};return Object.fromEntries(Object.entries(ranges).map(([k,s])=>[k,stat((DATA.series[String(f.id)]||[]).filter(r=>r[0]>=s&&r[0]<=e),f).periodReturn]))}}function filters(){{let ts=vals($('typeSel')),ms=vals($('mgrSel'));return DATA.funds.filter(f=>(!ts.length||ts.includes(f.type))&&(!ms.length||ms.includes(f.manager)))}}function init(){{$('start').value=DATA.dateMin;$('end').value=DATA.dateMax;$('start').min=DATA.dateMin;$('end').max=DATA.dateMax;let types=[...new Set(DATA.funds.map(f=>f.type))].sort((a,b)=>(typeOrder.indexOf(a)<0?99:typeOrder.indexOf(a))-(typeOrder.indexOf(b)<0?99:typeOrder.indexOf(b)));$('typeSel').innerHTML=types.map(v=>`<option>${{esc(v)}}</option>`).join('');let mgrs=[...new Set(DATA.funds.map(f=>f.manager))].sort((a,b)=>a.localeCompare(b,'ko-KR'));$('mgrSel').innerHTML=mgrs.map(v=>`<option>${{esc(v)}}</option>`).join('');$('bmBox').innerHTML=DATA.bms.map(b=>`<label><input type="checkbox" value="${{b.id}}"> ${{esc(b.name)}}</label>`).join('');document.querySelectorAll('#bmBox input').forEach(i=>i.onchange=()=>{{i.checked?selectedBm.add(i.value):selectedBm.delete(i.value);render()}});['typeSel','mgrSel','start','end'].forEach(id=>$(id).onchange=()=>{{fundList();render()}});$('selectBtn').onclick=()=>{{filters().forEach(f=>selected.add(String(f.id)));fundList();render()}};$('clearBtn').onclick=()=>{{filters().forEach(f=>selected.delete(String(f.id)));fundList();render()}};$('allPeriod').onclick=()=>{{$('start').value=DATA.dateMin;$('end').value=DATA.dateMax;render()}};$('ytd').onclick=()=>{{let y=$('end').value.slice(0,4);$('start').value=Math.max(DATA.dateMin,`${{y}}-01-01`);render()}};$('mtd').onclick=()=>{{$('start').value=$('end').value.slice(0,8)+'01';render()}};$('labels').onclick=()=>{{showLabels=!showLabels;$('labels').textContent=showLabels?'레이블 숨김':'레이블 표시';render()}};$('colors').onclick=()=>{{pal=(pal+1)%fundColors.length;render()}};fundList();render()}}function fundList(){{$('fundBox').innerHTML=filters().map(f=>`<label><input type="checkbox" value="${{f.id}}" ${{selected.has(String(f.id))?'checked':''}}> ${{esc(f.name)}}<br><span class="note">${{esc([f.type,f.manager,f.code].join(' · '))}}</span></label>`).join('');document.querySelectorAll('#fundBox input').forEach(i=>i.onchange=()=>{{i.checked?selected.add(i.value):selected.delete(i.value);render()}})}}function selectedItems(){{let fs=DATA.funds.filter(f=>selected.has(String(f.id))),bs=DATA.bms.filter(b=>selectedBm.has(String(b.id)));return[...fs,...bs]}}function render(){{let items=selectedItems(),models=items.map((f,i)=>{{let c=f.isBm?bmColors[pal][f.id]:fundColors[pal][i%fundColors[pal].length];return{{f,rows:rows(f.id),color:c,stats:stat(rows(f.id),f),tr:trail(f)}}}});chart(models);perf(models.length?models:DATA.funds.map((f,i)=>{{return{{f,rows:rows(f.id),color:fundColors[pal][i%fundColors[pal].length],stats:stat(rows(f.id),f),tr:trail(f)}}}}))}}function chart(models){{let series=models.map(m=>{{let lr=m.rows.filter(r=>r[6]!=null);if(lr.length<2)return null;let base=lr[0][6],pts=lr.map(r=>[r[0],m.f.returnMode==='simple'?(r[6]-base)/1000*100:(r[6]/base-1)*100]);return{{...m,pts}}}}).filter(Boolean);if(!series.length){{Plotly.purge('chart');$('chart').innerHTML='<div style="padding:40px;text-align:center;color:#657385">1개 이상의 펀드를 선택하세요</div>';$('legend').innerHTML='';return}}let traces=series.map(m=>{{let on=hi.has(String(m.f.id)),dim=hi.size&&!on;return{{x:m.pts.map(p=>p[0]),y:m.pts.map(p=>p[1]),type:'scatter',mode:'lines',name:m.f.name,opacity:dim?.18:1,line:{{color:m.color,width:on?4.5:(m.f.isBm?4:2.3)}},hovertemplate:`<b>${{esc(m.f.name)}}</b><br>%{{x}}<br>%{{y:.2f}}%<extra></extra>`}}}});let anns=showLabels?series.map(m=>{{let p=m.pts.at(-1),on=hi.has(String(m.f.id));return{{order:on?1:0,a:{{x:p[0],y:p[1],xref:'x',yref:'y',text:`${{m.f.name.slice(0,12)}} ${{p[1].toFixed(2)}}%`,showarrow:true,ax:24,ay:0,arrowcolor:'rgba(0,0,0,0)',font:{{size:on?13:11,color:on?'#111827':m.color}},bgcolor:on?'rgba(255,255,255,.98)':'rgba(255,255,255,.7)',bordercolor:m.color,borderwidth:on?2:1}}}}}}).sort((a,b)=>a.order-b.order).map(x=>x.a):[];Plotly.react('chart',traces,{{margin:{{l:55,r:28,t:15,b:40}},showlegend:false,hovermode:'closest',annotations:anns,yaxis:{{ticksuffix:'%',gridcolor:'#e5eaf0'}},xaxis:{{showgrid:false}}}},{{responsive:true,displayModeBar:false}});$('legend').innerHTML=series.map(m=>`<button data-id="${{m.f.id}}" class="${{hi.has(String(m.f.id))?'active':hi.size?'dim':''}}"><span style="color:${{m.color}}">━</span> ${{esc(m.f.name)}}</button>`).join('');document.querySelectorAll('#legend button').forEach(b=>b.onclick=()=>{{let id=b.dataset.id;hi.has(id)?hi.delete(id):hi.add(id);render()}})}}function perf(models){{$('perf').innerHTML=models.map(m=>{{let s=m.stats,t=m.tr,cl=v=>v>0?'pos':v<0?'neg':'';return`<tr><td>${{esc(m.f.name)}}</td><td>${{esc(m.f.type||'')}}</td><td>${{esc(m.f.investDate||'')}}</td><td class="${{cl(s.periodReturn)}}">${{pct(s.periodReturn)}}</td><td class="${{cl(s.annualReturn)}}">${{pct(s.annualReturn)}}</td><td>${{pct(s.vol)}}</td><td>${{num(s.sharpe)}}</td><td class="${{cl(s.mdd)}}">${{pct(s.mdd)}}</td><td>${{pct(t.ytd)}}</td><td>${{pct(t.mtd)}}</td><td>${{pct(t.d1)}}</td><td>${{pct(t.w1)}}</td><td>${{pct(t.m1)}}</td><td>${{pct(t.m3)}}</td><td>${{pct(t.m6)}}</td><td>${{pct(t.y1)}}</td></tr>`}}).join('')}}init();</script></body></html>'''


if __name__ == "__main__":
    html = build_html(build_payload())
    OUTPUT.write_text(html, encoding="utf-8")
    WEB_OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    WEB_OUTPUT.write_text(html, encoding="utf-8")
    print(OUTPUT)
    print(WEB_OUTPUT)
