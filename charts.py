def search_data_area_worksheet(ws):
    """
    차트에 입력한 데이터 영역을 가져오는 함수
 
    output : ='Sheet1'!$A$1:$D$6
    """
    used_range = ws.UsedRange
    # 1) 시작 셀과 끝 셀의 행/열 구하기
    first_row = used_range.Row
    first_col = used_range.Column
    last_row = first_row + used_range.Rows.Count - 1
    last_col = first_col + used_range.Columns.Count - 1
    # 2) 열 번호 → 알파벳 변환 함수
    def col_letter(n):
        letters = ""
        while n:
            n, remainder = divmod(n - 1, 26)
            letters = chr(65 + remainder) + letters
        return letters
    first_cell = f"${col_letter(first_col)}${first_row}"
    last_cell  = f"${col_letter(last_col)}${last_row}"
    # 3) 엑셀 수식 표기 형태로 문자열 생성
    sheet_name = ws.Name
    range_str = f"='{sheet_name}'!{first_cell}:{last_cell}"
    return range_str
 
def update_chart_data(chart, categories, series_data):
    """
    지정한 chart 의 데이터시트를 업데이트하는 함수
    """
    # 1) 차트 데이터 활성화(임베디드 Excel 열림)
    chart.ChartData.Activate()
    wb = chart.ChartData.Workbook
    ws = wb.Worksheets(1)  # 보통 첫 시트에 데이터가 있음
    # 2) 기존 데이터 삭제
    ws.Cells.Clear()
    # 3) 데이터 업데이트
    # categories 입력
    for n, cate in enumerate(categories):
        n += 2
        ws.Cells(n, 1).Value = cate
    # 헤더(시리즈명)
    ws.Cells(1, 1).Value = ""  # A1 비움
    col = 2
    for sname in series_data.keys():
        # 기존 데이터 제거
        ws.Columns(col).ClearContents()
        # 컬럼명 입력
        ws.Cells(1, col).Value = sname
        # 데이터 입력
        values = series_data[sname]
        for n, v in enumerate(values):
            n += 2 # 컬럼명 row 건너뛰기
            ws.Cells(n, col).Value = v
        col += 1
    data_range = search_data_area_worksheet(ws)
    chart.SetSourceData(data_range)
    wb.Close(SaveChanges=False)
 
##### 실행 #######
 
# 대상 차트 찾기
slide = presentation.Slides(4) # 슬라이드 페이지
shape = slide.Shapes("Chart 8") # 차트 이름
if not shape.HasChart:
    raise RuntimeError("해당 Shape는 차트가 아닙니다.")
 
 
chart = shape.Chart
 
# 1) 입력 데이터
categories = ["Q1", "Q2", "Q3", "Q4", "Q5"] # x축 
series_data = {
    "매출": [10, 8, 7, 2, 10],
    "이익": [30, 45, 42, 55, 30],
    "비용": [60, 30, 20, 10, 100]
} # y축
 
update_chart_data(chart, categories, series_data)