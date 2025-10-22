import re
from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.enum.shapes import MSO_SHAPE_TYPE

# ---------------------------
# 🔹 Text / Chart Utilities
# ---------------------------

TOKEN_RE = re.compile(r"\{\{([^\{\}]+)\}\}")

def iter_shapes(slide):
    """모든 도형(그룹 내부 포함)을 순회"""
    for shp in slide.shapes:
        yield shp
        if shp.shape_type == MSO_SHAPE_TYPE.GROUP and hasattr(shp, "shapes"):
            for s in shp.shapes:
                yield s

# ---------------------------------
# ① 텍스트 토큰 치환
# ---------------------------------
def replace_text_tokens(prs: Presentation, mapping: dict):
    """
    PowerPoint 내 {{TOKEN}} 형태의 텍스트를 mapping 값으로 교체.
    예:
        mapping = {"SL1_VIS_YOY_P_1": "+12.3%", "SL1_SAL_WON_1": "4,070,000,000원"}
    """
    replaced = 0
    for slide in prs.slides:
        for shp in iter_shapes(slide):
            if not getattr(shp, "has_text_frame", False):
                continue
            for p in shp.text_frame.paragraphs:
                for r in p.runs:
                    before = r.text
                    r.text = TOKEN_RE.sub(
                        lambda m: str(mapping.get(m.group(1).strip(), m.group(0))), r.text
                    )
                    if r.text != before:
                        replaced += 1
    return replaced


# ---------------------------------
# ② 차트 찾기 및 데이터 교체
# ---------------------------------
def find_chart(prs: Presentation, name: str):
    """이름으로 차트 도형 찾기"""
    for slide in prs.slides:
        for shp in iter_shapes(slide):
            if getattr(shp, "name", "") == name and getattr(shp, "has_chart", False):
                return shp.chart
    return None


def replace_chart_data(chart, categories, series_dict):
    """
    차트 데이터 교체 (서식 유지)
    categories : X축 목록
    series_dict : {"시리즈명": [값 리스트], ...}
    """
    cd = ChartData()
    cd.categories = list(categories)
    for sname, svalues in series_dict.items():
        cd.add_series(sname, list(svalues))
    chart.replace_data(cd)


# ---------------------------------
# ③ 표(Table) 교체 (선택사항)
# ---------------------------------
def find_table(prs: Presentation, name: str):
    """이름으로 표 도형 찾기"""
    for slide in prs.slides:
        for shp in iter_shapes(slide):
            if getattr(shp, "name", "") == name and getattr(shp, "has_table", False):
                return shp.table
    return None


def fill_table(table, dataframe):
    """
    pandas DataFrame 데이터를 표에 삽입.
    컬럼 수와 행 수가 PPT 표 구조와 일치해야 함.
    """
    n_rows = len(dataframe)
    n_cols = len(dataframe.columns)
    for r in range(n_rows):
        for c in range(n_cols):
            val = str(dataframe.iat[r, c])
            table.cell(r, c).text = val


# ---------------------------
# 🔹 Formatter Utilities
# ---------------------------

def fmt_int_comma(x):
    """정수 천단위 콤마"""
    try:
        return f"{int(round(float(x))):,}"
    except Exception:
        return str(x)


def fmt_signed_percent_1(x):
    """+/- 표시, 소수점 1자리 %"""
    try:
        return f"{float(x):+0.1f}%"
    except Exception:
        return str(x)


def fmt_won_or_eok(x):
    """1억 이상이면 억원 단위로 변환"""
    try:
        v = float(x)
        if abs(v) >= 1e8:
            return f"{v/1e8:0.1f}억원"
        else:
            return f"{int(v):,}원"
    except Exception:
        return str(x)


FORMATTERS = {
    "int_comma": fmt_int_comma,
    "signed_percent_1": fmt_signed_percent_1,
    "won_or_eok": fmt_won_or_eok,
}


# ---------------------------
# 🔹 Helper (한 번에 실행용)
# ---------------------------
def apply_tokens_and_charts(prs_path, out_path, token_map, chart_map=None):
    """
    PPT에 토큰/차트 모두 반영하고 저장
    token_map: {TOKEN: 값}
    chart_map: {chart_name: (categories, series_dict)}
    """
    prs = Presentation(prs_path)
    replace_text_tokens(prs, token_map)
    if chart_map:
        for cname, (cats, sdict) in chart_map.items():
            ch = find_chart(prs, cname)
            if ch:
                replace_chart_data(ch, cats, sdict)
    prs.save(out_path)
    return out_path
