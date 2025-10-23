# src/run_build_report.py
import os
import yaml
from sqlalchemy import create_engine, text
from dotenv import load_dotenv, find_dotenv
from ppt_fillers import apply_tokens_and_charts
import time
import pandas as pd

# === 트리맵 이미지 생성 유틸 ===
def generate_treemap_image(data, out_path, color_hex, title=None, font_family="Malgun Gothic"):
    """
    data: list[tuple[str, float]] # [(업종명, 억원값)]
    out_path: 이미지 저장 경로
    
    디자인 목표: 
    1. 폰트: Noto Sans KR 적용
    2. 색상: 모든 타일 동일한
    3. 경계선: 굵은 흰색 테두리
    4. 텍스트: 흰색, 굵게
    """
    import matplotlib.pyplot as plt
    import squarify
    import os
    
    # 1. 데이터 준비 및 정렬
    data.sort(key=lambda item: item[1], reverse=True) 
    
    labels = [f"{k}\n{v:.1f}억" for k, v in data]
    sizes = [max(v, 0.0) for _, v in data]

    # 2. 색상 설정 (단일 색상 코드 사용)
    MONO_COLOR = color_hex # 전달받은 HEX 코드를 사용
    
    # 모든 타일에 동일한 색상을 적용하기 위해 리스트 생성
    colors = [MONO_COLOR] * len(sizes) 

    # 3. 폰트 설정
    try:
        # Noto Sans KR 폰트 적용
        plt.rcParams["font.family"] = font_family
    except Exception:
        # 폰트 설정을 실패하더라도 계속 진행
        pass

    # 4. 트리맵 그리기
    fig = plt.figure(figsize=(7, 4.5)) 
    fig.set_facecolor('none')

    plt.clf()
    
    squarify.plot(
        sizes=sizes, 
        label=labels, 
        color=colors,         # 단일 색상 적용
        # **경계선 설정: 굵고 흰색 테두리 적용**
        bar_kwargs={
            'linewidth': 2,       # 테두리 두께
            'edgecolor': 'white'  # 테두리 색상
        }, 
        # **텍스트 속성 설정: 흰색, 굵은 글씨**
        text_kwargs={
            'fontsize': 11, 
            'color': 'white',      
            'fontweight': 'bold'   
        } 
    )
    
    plt.axis("off") # 축 제거
    
    if title:
        plt.title(title, pad=10, fontsize=12, fontweight='bold')
        
    plt.tight_layout(pad=0.0)
    
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    plt.savefig(out_path, dpi=300, bbox_inches='tight', pad_inches=0) # 고해상도 저장
    plt.close()

# === 히트맵 이미지 생성 유틸 ===
def generate_heatmap_image(data_df, out_path, title=None, font_family="Malgun Gothic"):
    """
    data_df: pandas DataFrame (행/열 이름과 값 포함)
    out_path: 이미지 저장 경로
    """
    import matplotlib.pyplot as plt
    import seaborn as sns
    import os
    
    # 폰트 설정
    try:
        plt.rcParams["font.family"] = font_family
        plt.rcParams["axes.unicode_minus"] = False
    except Exception:
        pass

    # 1. Figure 생성 및 초기화
    fig, ax = plt.subplots(figsize=(8, 6))
    fig.set_facecolor('none')
    
    # 2. Seaborn 히트맵 생성
    # annot=True: 값 표시, fmt=".1f": 소수점 첫째 자리까지 표시
    # cmap="YlGnBu": 색상 맵 (원하는 색상 계열로 변경 가능)
    sns.heatmap(data_df, 
                annot=True, 
                fmt=".1f", 
                linewidths=.5, 
                linecolor='lightgray',
                cmap="Blues",
                cbar_kws={'shrink': .8}, # 컬러바 크기 조절
                ax=ax,
                alpha=0.7,
                annot_kws={"fontweight": "bold", "fontsize": 10})
    
    # 3. 축 레이블 회전 및 설정
    ax.tick_params(axis='x', rotation=0) # X축 레이블 유지
    ax.tick_params(axis='y', rotation=0) # Y축 레이블 유지

    ax.set_xlabel('')
    ax.set_ylabel('')
    
    # 4. 제목 설정
    if title:
        ax.set_title(title, pad=10, fontsize=12, fontweight='bold')
    
    # 5. 레이아웃 조정 및 저장
    plt.tight_layout()
    
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    
    # 여백을 최소화하여 저장
    plt.savefig(out_path, dpi=300, bbox_inches='tight', pad_inches=0.1, transparent=True)
    plt.close(fig) # Figure 객체 닫기



load_dotenv(find_dotenv(), override=True)

db_url = os.getenv("DB_URL")

engine = create_engine(
    db_url, 
    connect_args={"options": "-csearch_path=regionmonitor"}
)


with open("config/slides_text.yml", encoding="utf-8") as f:
    cfg = yaml.safe_load(f)

token_values = {}
chart_data = {}
image_map = {}

for s in cfg["slides"]:
    for token, meta in s.get("tokens", {}).items():
        with engine.connect() as conn:
            val = conn.execute(text(meta["sql"]), cfg["params"]).scalar()
            token_values[token] = val

    for chart_name, chart_conf in s.get("charts", {}).items():
        if "treemap_sql" in chart_conf:
            with engine.connect() as conn:
                rows = conn.execute(text(chart_conf["treemap_sql"]), cfg["params"]).fetchall()
                # rows: [(업종명, 억원값), ...]
                data = [(r[0], float(r[1]) if r[1] is not None else 0.0) for r in rows]

            outfile = chart_conf["outfile"]
            title = chart_conf.get("title", None)
            color_hex = chart_conf.get("color_hex", "#4682B4")
            generate_treemap_image(data, outfile, color_hex=color_hex, title=title)

            time.sleep(1.0)

            # PPT 도형 이름에 이 이미지를 꽂아라
            shape_name = chart_conf["shape"]
            image_map[shape_name] = outfile
            continue  # 다음 차트로

        if "heatmap_sql" in chart_conf:
            with engine.connect() as conn:
                # pandas를 사용하여 SQL 결과를 DataFrame으로 로드
                df = pd.read_sql(text(chart_conf["heatmap_sql"]), 
                                 conn, 
                                 params=cfg["params"])
                
                # 수정: 첫 번째 컬럼('industry')을 인덱스로 설정
                # 쿼리가 이미 피벗된 형태이므로, 인덱스만 설정
                data_df = df.set_index(df.columns[0]) 
            
            outfile = chart_conf["outfile"]
            title = chart_conf.get("title", None)
            color_hex = chart_conf.get("color_hex", "#4682B4") 
            
            # 히트맵 이미지 생성 (수정된 data_df 전달)
            generate_heatmap_image(data_df, outfile, title=title)
            
            time.sleep(1.0) # 파일 I/O 충돌 방지
            
            # PPT 도형 이름에 이 이미지를 꽂아라
            shape_name = chart_conf["shape"]
            image_map[shape_name] = outfile
            continue # 다음 차트로

        if "category_sql" in chart_conf:
            with engine.connect() as conn:
                categories = [r[0] for r in conn.execute(text(chart_conf["category_sql"]), cfg["params"]).fetchall()]
                series = {}
                for sname, ssql in chart_conf["series"].items():
                    series[sname] = [r[0] for r in conn.execute(text(ssql), cfg["params"]).fetchall()]
                chart_data[chart_name] = (categories, series)


print(f"DEBUG: 최종 Image Map: {image_map}")

apply_tokens_and_charts(
    prs_path="template/master.pptx",
    out_path="out/report_test_테스트최종2.pptx",
    token_map=token_values,
    chart_map=chart_data,
    image_map=image_map
)

print("✅ 테스트 슬라이드 생성 완료")
