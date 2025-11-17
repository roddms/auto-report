#src/ppt_fillers.py
import os
import re
from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_MARKER_STYLE
import win32com.client as win32
import pandas as pd
import pythoncom
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN

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


def _set_cell(cell, text, align=PP_ALIGN.LEFT, bold=False, font_name="Pretendard", font_size_pt=10.5):
    tf = cell.text_frame
    tf.clear()                         # 기존 문단/런 비우기
    p = tf.paragraphs[0]               # clear() 후 항상 1개 생김
    run = p.add_run()                  # 최소 1개 run 보장
    run.text = "" if text is None else str(text)

    p.alignment = align
    font = run.font
    font.name = font_name
    font.size = Pt(font_size_pt)
    font.bold = bold
    font.color.rgb = RGBColor(0,0,0)

def fill_table_with_padding(table, rows):
    """
    헤더 없는 표에 데이터를 채움.
    table: python-pptx Table 객체
    rows:  [("주차장명", "1,234대"), ...]
    - 표 행 수보다 데이터가 많으면 잘라내고, 부족하면 빈칸으로 패딩.
    - 좌/우 정렬 및 서식 일괄 적용.
    """
    n_rows = len(table.rows)
    # 표는 2열 가정: [이름, 값]. 더 많은 열이면 필요에 맞게 수정.
    for i in range(n_rows):
        if i < len(rows):
            name, slots = rows[i]
            _set_cell(table.cell(i, 0), name, align=PP_ALIGN.LEFT)
            _set_cell(table.cell(i, 1), slots, align=PP_ALIGN.LEFT)
        else:
            _set_cell(table.cell(i, 0), "", align=PP_ALIGN.LEFT)
            _set_cell(table.cell(i, 1), "", align=PP_ALIGN.LEFT)


# ---------------------------
#  트리맵, 히트맵 이미지 삽입
# ---------------------------
def add_images_to_presentation(prs: Presentation, image_map: dict):
    """
    image_map: { "도형이름": "이미지경로", ... }
    - 모든 슬라이드를 반복하며, image_map에 있는 도형 이름과 일치하는 
      도형을 찾아 이미지를 삽입합니다.
    """
    import io
    
    # prs.slides는 인덱스 0부터 시작합니다.
    for slide_idx, slide in enumerate(prs.slides):
        idx_to_replace = []
        
        # 현재 슬라이드의 모든 도형을 확인
        for i, shp in enumerate(slide.shapes):
            name = shp.name
            # image_map에 현재 도형 이름이 있는지 확인
            if name in image_map:
                # 위치와 크기 정보를 가져옵니다.
                left = shp.left
                top = shp.top
                width = shp.width
                height = shp.height
                idx_to_replace.append((i, name, left, top, width, height))

        # 현재 슬라이드에서 대체할 도형이 있다면 처리
        if not idx_to_replace:
            continue

        print(f"INFO: 슬라이드 {slide_idx+1} ({len(idx_to_replace)}개 차트) 이미지 삽입 시작")
        
        # 뒤에서부터 삭제(인덱스 보존) 및 이미지 삽입
        for i, name, left, top, width, height in reversed(idx_to_replace):
            shp = slide.shapes[i]
            
            # 템플릿 도형(Placeholder) 제거
            slide.shapes._spTree.remove(shp._element)
            
            # 이미지 파일을 바이너리 스트림으로 읽어 삽입 (이전 수정사항 유지)
            img_path = image_map[name]
            
            try:
                with open(img_path, 'rb') as f:
                    image_stream = io.BytesIO(f.read())
                
                # 스트림을 add_picture 함수의 첫 번째 인수로 전달
                slide.shapes.add_picture(image_stream, left, top, width=width, height=height)
                
            except FileNotFoundError:
                print(f"⚠️ 경고: 이미지 파일 경로를 찾을 수 없습니다: {img_path}")
            except Exception as e:
                print(f"❌ 이미지 삽입 중 오류 발생 ({img_path}): {e}")


# 증감별 색 추가
def colorize_arrows(prs):
    """
    PowerPoint 전체 순회하며,
    문단 내에 ▲(상승) 있으면 문단 전체 빨강,
    ▼(하락) 있으면 문단 전체 파랑.
    """
    for slide in prs.slides:
        for shp in iter_shapes(slide):
            if not getattr(shp, "has_text_frame", False):
                continue
            for p in shp.text_frame.paragraphs:
                full_text = "".join(r.text for r in p.runs)
                if "▲" in full_text:
                    for r in p.runs:
                        r.font.color.rgb = RGBColor(231, 76, 60)     # 빨강
                elif "▼" in full_text:
                    for r in p.runs:
                        r.font.color.rgb = RGBColor(0, 112, 192)     # 파랑


def colorize_condition(prs):
    """
    텍스트 런이 '충분' 또는 '부족' 단독일 때만 해당 런에 색상 적용
    """
    for slide in prs.slides:
        for shp in iter_shapes(slide):
            if not getattr(shp, "has_text_frame", False):
                continue

            for p in shp.text_frame.paragraphs:
                for r in p.runs:
                    txt = r.text.strip()
                    if txt == "충분":
                        r.font.color.rgb = RGBColor(71, 187, 120)   # 초록
                    elif txt == "부족":
                        r.font.color.rgb = RGBColor(192, 57, 43)    # 빨강


# 막대 그래프 
def highlight_max_point_only(chart, series_idx=0, max_idx=0,
                             color=RGBColor(231, 76, 60)):
    """
    chart.series[series_idx]의 특정 point(=max_idx)만 강조색 적용.
    나머지 포인트는 건드리지 않아 기존 색 유지됨.
    """
    series = chart.series[series_idx]
    if max_idx < 0 or max_idx >= len(series.points):
        return
    pt = series.points[max_idx]
    fill = pt.format.fill
    fill.solid()
    fill.fore_color.rgb = color


def color_sl20_chart_1(chart, sdict):
    # 시리즈명 매핑
    name2idx = {chart.series[i].name: i for i in range(len(chart.series))}
    flags = sdict.get("_festival_flags")

    # 1) 막대(금액) 포인트별 색상: 축제=빨강, 비축제=파랑
    bar_name = "매출금액(백만원)"
    if flags and bar_name in name2idx:
        s = chart.series[name2idx[bar_name]]
        for i, pt in enumerate(s.points):
            fill = pt.format.fill
            fill.solid()
            # 축제 = 빨강, 비축제 = 파랑
            fill.fore_color.rgb = RGBColor(231, 76, 60) if flags[i] == 1 else RGBColor(31, 78, 121)

    # 2) 라인(건수) 기본색 지정
    line_name = "매출건수(건)"
    if line_name in name2idx:
        try:
            chart.series[name2idx[line_name]].format.line.color.rgb = RGBColor(96, 96, 96)
        except Exception:
            pass


def color_sl21_chart(chart, sdict, bar_series_name="방문인구(명)"):
    flags = sdict.get("_festival_flags")
    name2idx = {chart.series[i].name: i for i in range(len(chart.series))}
    if not flags or bar_series_name not in name2idx:
        return

    s = chart.series[name2idx[bar_series_name]]
    for i, pt in enumerate(s.points):
        try:
            fill = pt.format.fill
            fill.solid()
            # 축제 = 빨강, 비축제 = 파랑
            fill.fore_color.rgb = RGBColor(231, 76, 60) if flags[i] == 1 else RGBColor(31, 78, 121)
        except Exception:
            pass


def color_line_markers_by_flags(chart, series_name, flags,
                                color_festival=RGBColor(231, 76, 60),   # 빨강
                                color_default=RGBColor(96, 96, 96),     # 회색/기본
                                marker_size=6):
    """
    chart: python-pptx Chart
    series_name: 라인 시리즈 이름 (예: '매출건수(건)')
    flags: [0/1,...] (축제=1, 비축제=0)
    포인트별로 마커 색을 다르게 칠함 (선 색은 그대로 유지)
    """
    # 시리즈 찾기
    name2idx = {chart.series[i].name: i for i in range(len(chart.series))}
    if series_name not in name2idx or not flags:
        return

    s = chart.series[name2idx[series_name]]

    # 마커가 꺼져 있으면 켜기 (시리즈 공통 설정)
    if XL_MARKER_STYLE is not None:
        try:
            s.marker.style = XL_MARKER_STYLE.CIRCLE
            s.marker.size = marker_size
        except Exception:
            pass

    # 포인트별 마커 색상 적용
    for i, pt in enumerate(s.points):
        try:
            fill = pt.format.fill
            fill.solid()
            fill.fore_color.rgb = color_festival if flags[i] == 1 else color_default

        except Exception:
            # 길이 불일치 등은 조용히 스킵
            continue

# ---------------------------------
# 🔹 PowerPoint 트리맵 차트 자동 갱신
# ---------------------------------
import time
import os
import win32com.client as win32


# def update_treemap_chart(ppt_path, out_path, shape_name, rows, value_header="계열1"):
#     ppt_path = os.path.abspath(ppt_path)
#     out_path = os.path.abspath(out_path)
#     ppt_path = ppt_path.replace("/", "\\")
#     out_path = out_path.replace("/", "\\")

#     if not os.path.exists(ppt_path):
#         raise FileNotFoundError(f"PPT path not found: {ppt_path}")

#     pp = win32.gencache.EnsureDispatch("PowerPoint.Application")
#     pp.Visible = True

#     # 위치 인자(파일, ReadOnly, Untitled, WithWindow)
#     pres = pp.Presentations.Open(ppt_path, False, False, False)
#     try:
#         target = None
#         for s in pres.Slides:
#             for shp in s.Shapes:
#                 if shp.Name == shape_name and getattr(shp, "HasChart", False):
#                     target = shp
#                     break
#             if target:
#                 break

#         if not target:
#             raise RuntimeError(f"도형 '{shape_name}' 차트를 찾을 수 없습니다.")

#         ch = target.Chart
#         ch.ChartData.Activate()
#         wb = ch.ChartData.Workbook
#         ws = wb.Worksheets(1)

#         # 시트 초기화 + 헤더
#         ws.UsedRange.ClearContents()
#         ws.Cells(1, 1).Value = "계열"
#         ws.Cells(1, 2).Value = "상위"
#         ws.Cells(1, 3).Value = "하위"
#         ws.Cells(1, 4).Value = value_header

#         r = 2
#         for series, parent, child, val in rows:
#             ws.Cells(r, 1).Value = str(series)
#             ws.Cells(r, 2).Value = str(parent)
#             ws.Cells(r, 3).Value = str(child)
#             ws.Cells(r, 4).Value = float(val) if val is not None else 0.0
#             r += 1

#         wb.Close(SaveChanges=True)
#         ch.Refresh()

#         # 동일 파일로 저장
#         pres.SaveAs(out_path)
#         time.sleep(0.4)  # 파일 쓰기 안정화
#     finally:
#         pres.Close()
#         pp.Quit()

def update_treemaps_batch(ppt_path, out_path, updates):
    """
    updates: [(shape_name, rows, value_header), ...]
    rows: [(series, parent, child, value), ...]
    """
    ppt_path = os.path.abspath(ppt_path).replace("/", "\\")
    out_path = os.path.abspath(out_path).replace("/", "\\")
    if not os.path.exists(ppt_path):
        raise FileNotFoundError(f"PPT not found: {ppt_path}")

    pythoncom.CoInitialize()
    try:
        # 캐시(gen_py) 우회: 동적 바인딩
        pp = win32.DispatchEx("PowerPoint.Application")
    except Exception:
        pp = win32.Dispatch("PowerPoint.Application")
    pp.Visible = True

    pres = pp.Presentations.Open(ppt_path, False, False, False)
    try:
        for shape_name, rows, value_header in updates:
            target = None
            for s in pres.Slides:
                for shp in s.Shapes:
                    if shp.Name == shape_name and getattr(shp, "HasChart", False):
                        target = shp
                        break
                if target:
                    break
            if not target:
                print(f"⚠️ 도형 '{shape_name}' 차트를 찾을 수 없습니다.")
                continue

            ch = target.Chart
            ch.ChartData.Activate()
            wb = ch.ChartData.Workbook
            ws = wb.Worksheets(1)

            ws.UsedRange.ClearContents()
            ws.Cells(1, 1).Value = "계열"
            ws.Cells(1, 2).Value = "상위"
            ws.Cells(1, 3).Value = "하위"
            ws.Cells(1, 4).Value = value_header

            r = 2
            for series, parent, child, val in rows:
                ws.Cells(r, 1).Value = str(series)
                ws.Cells(r, 2).Value = str(parent)
                ws.Cells(r, 3).Value = str(child)
                ws.Cells(r, 4).Value = float(val) if val is not None else 0.0
                r += 1

            wb.Close(SaveChanges=True)
            ch.Refresh()
            time.sleep(0.1)

        pres.SaveAs(out_path)
        time.sleep(0.2)
    finally:
        try:
            for pres in list(pp.Presentations):
                try:
                    pres.Close()
                except:
                    pass
            pp.Quit()
        except:
            pass
        pythoncom.CoUninitialize()

# ---------------------------
# 🔹 Helper (한 번에 실행용)
# ---------------------------
def apply_tokens_and_charts(prs_path, out_path, token_map, chart_map=None, image_map=None):
    """
    PPT에 토큰/차트 모두 반영하고 저장
    token_map: {TOKEN: 값}
    chart_map: {chart_name: (categories, series_dict)}
    """
    prs = Presentation(prs_path)

    replace_text_tokens(prs, token_map)

    colorize_arrows(prs)

    colorize_condition(prs)


    if chart_map:
        for cname, (cats, sdict) in chart_map.items():
            ch = find_chart(prs, cname)
            if ch:
                data_sdict = {k: v for k, v in sdict.items() if not k.startswith("_")}
                replace_chart_data(ch, cats, data_sdict)

                if cname == "SL5_chart_2":
                    try:
                        # 첫 번째 시리즈의 값 리스트 추출
                        vals = next(iter(sdict.values()))
                        if vals:
                            max_idx = vals.index(max(vals))
                            highlight_max_point_only(ch, series_idx=0, max_idx=max_idx, color=RGBColor(231, 76, 60))
                    except Exception as e:
                        print(f"⚠️ SL5_chart_2 강조 처리 실패: {e}")

                if cname == "SL7_chart":
                    try:
                        # 첫 번째 시리즈의 값 리스트 추출
                        vals = next(iter(sdict.values()))
                        if vals:
                            max_idx = vals.index(max(vals))
                            highlight_max_point_only(ch, series_idx=0, max_idx=max_idx, color=RGBColor(91, 155, 213))
                    except Exception as e:
                        print(f"⚠️ SL7_chart 강조 처리 실패: {e}")

                pythoncom.PumpWaitingMessages()
                time.sleep(0.2)
    
                if cname == "SL20_chart":
                    color_sl20_chart_1(ch, sdict)
                    flags = sdict.get("_festival_flags")
                    color_line_markers_by_flags(ch, series_name="매출건수(건)", flags=flags)

                if cname == "SL21_chart":
                    color_sl21_chart(ch, sdict)

    if image_map:
        # out_path -> out_path 로 제자리 저장 (같은 경로 덮어씀)
        add_images_to_presentation(prs, image_map)

    prs.save(out_path)
    return out_path