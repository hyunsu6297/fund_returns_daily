import json
import re
from pathlib import Path


def replace_data_payload(html):
    marker = "const DATA="
    if marker not in html:
        return html

    start = html.index(marker) + len(marker)
    end = html.index(";", start)
    data = json.loads(html[start:end])
    data["unmappedFunds"] = [
        item
        for item in data.get("unmappedFunds", [])
        if item.get("type") != "채권형(ETF)"
    ]
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    return html[:start] + payload + html[end:]


def normalize_ui(html):
    style = (
        ".control-title,.side-title{"
        "color:#0f766e!important;"
        "font-size:14px!important;"
        "font-weight:800!important;"
        "}"
    )
    if style not in html:
        html = html.replace("</style>", style + "</style>", 1)

    html = html.replace(">펀드 선택<", ">유형/운용사별 선택<")
    html = html.replace(">연초<", ">연초 이후<")
    html = html.replace(">월초<", ">월초 이후<")

    team_box = '<div id="teamButtons" class="team-actions"></div>'
    team_title = '<p class="side-title" style="margin-top:12px">팀별 선택</p>'
    html = html.replace(
        '<h3 class=\\"side-title\\" style=\\"margin-top:12px\\">팀별 선택</h3>',
        team_title,
    )
    html = re.sub(
        r'<h3 class="side-title" style="margin-top:\s*12px;?">팀별 선택</h3>',
        team_title,
        html,
    )
    if "팀별 선택" not in html and team_box in html:
        html = html.replace(team_box, team_title + team_box, 1)

    period_button = re.search(
        r'<button[^>]*id="allPeriod"[^>]*>기간 전체</button>', html
    )
    if period_button and html.find('id="allPeriod"') > html.find('id="mtd"'):
        button_html = period_button.group(0)
        html = html[: period_button.start()] + html[period_button.end() :]
        html = re.sub(
            r'(<button[^>]*id="mtd"[^>]*>월초 이후</button>)',
            r"\1" + button_html,
            html,
            count=1,
        )

    html = re.sub(r'<span class="quick">\s*</span>', "", html)
    html = re.sub(r'<span class="quick-buttons">\s*</span>', "", html)
    return html


def main():
    targets = [Path("docs/index.html"), *Path(".").glob("fund_return_chart(*.html")]
    for path in targets:
        if not path.exists():
            continue
        html = path.read_text(encoding="utf-8")
        html = replace_data_payload(html)
        html = normalize_ui(html)
        path.write_text(html, encoding="utf-8")


if __name__ == "__main__":
    main()
