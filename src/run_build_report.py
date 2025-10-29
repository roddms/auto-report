# src/run_build_report.py
import os
import time
import yaml
import pandas as pd
from sqlalchemy import create_engine, text
from dotenv import load_dotenv, find_dotenv

from ppt_fillers import apply_tokens_and_charts, update_treemap_chart

# ---------------------------------
# (옵션) 트리맵/히트맵 이미지 유틸
#   * PPT 내부 트리맵만 쓰려면 treemap 이미지 분기는 스킵해도 됨
# ---------------------------------
def generate_treemap_image(data, out_path, color_hex, title=None, font_family="Malgun Gothic"):
    import os, math
    import matplotlib.pyplot as plt
    import squarify

    try:
        plt.rcParams["font.family"] = font_family
    except Exception:
        pass

    cleaned = []
    for k, v in (data or []):
        try:
            x = float(v)
        except Exception:
            x = float("nan")
        if x is not None and not math.isnan(x) and x > 0:
            cleaned.append((k, x))

    if not cleaned:
        fig = plt.figure(figsize=(7, 4.5))
        fig.set_facecolor('none')
        ax = fig.add_subplot(111)
        ax.axis("off")
        if title:
            ax.set_title(title, pad=10, fontsize=12, fontweight='bold')
        ax.text(0.5, 0.5, "데이터 없음", ha="center", va="center", fontsize=13, color="#777")
        os.makedirs(os.path.dirname(out_path), exist_ok=True)
        fig.savefig(out_path, dpi=300, bbox_inches='tight', pad_inches=0)
        plt.close(fig)
        return

    cleaned.sort(key=lambda item: item[1], reverse=True)
    labels = [f"{k}\n{v:.1f}억" for k, v in cleaned]
    sizes  = [v for _, v in cleaned]
    colors = [color_hex] * len(sizes)

    fig = plt.figure(figsize=(7, 4.5))
    fig.set_facecolor('none')
    import matplotlib.pyplot as plt
    plt.clf()

    squarify.plot(
        sizes=sizes,
        label=labels,
        color=colors,
        bar_kwargs={"linewidth": 2, "edgecolor": "white"},
        text_kwargs={"fontsize": 11, "color": "white", "fontweight": "bold"},
        pad=True
    )
    plt.axis("off")
    if title:
        plt.title(title, pad=10, fontsize=12, fontweight='bold')
    plt.tight_layout(pad=0.0)

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    plt.savefig(out_path, dpi=300, bbox_inches='tight', pad_inches=0)
    plt.close()


def generate_heatmap_image(data_df, out_path, title=None, font_family="Malgun Gothic"):
    import matplotlib.pyplot as plt
    import seaborn as sns
    import os

    try:
        plt.rcParams["font.family"] = font_family
        plt.rcParams["axes.unicode_minus"] = False
    except Exception:
        pass

    fig, ax = plt.subplots(figsize=(8, 6))
    fig.set_facecolor('none')

    annot_data = data_df.applymap(lambda x: f'{x:.1f}%')
    annot_data[data_df.isna()] = ""

    sns.heatmap(
        data_df,
        annot=annot_data,
        fmt="",
        linewidths=.5,
        linecolor='lightgray',
        cmap="Blues",
        cbar_kws={'shrink': .8},
        ax=ax,
        alpha=0.8,
        annot_kws={"fontweight": 600, "fontsize": 10}
    )

    ax.tick_params(axis='x', rotation=0, colors='#404040', labelsize=9)
    ax.tick_params(axis='y', rotation=0, colors='#404040', labelsize=9)
    ax.set_xlabel('')
    ax.set_ylabel('')

    if title:
        ax.set_title(title, pad=10, fontsize=12, fontweight='bold')

    plt.tight_layout()
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    plt.savefig(out_path, dpi=300, bbox_inches='tight', pad_inches=0.1, transparent=True)
    plt.close(fig)


# ------------------------------
# 환경/DB 설정
# ------------------------------
load_dotenv(find_dotenv(), override=True)
db_url = os.getenv("DB_URL")

engine = create_engine(
    db_url,
    connect_args={"options": "-csearch_path=regionmonitor"}
)

with open("config/slides_tokens.yml", encoding="utf-8") as f:
    cfg = yaml.safe_load(f)

OUTPUT_PPT = "out/test_1537.pptx"
TEMPLATE_PPT = "template/test.pptx"

token_values = {}
chart_data = {}
image_map = {}

# ------------------------------
# 토큰 / 일반 차트 / (옵션) 이미지 차트 생성
# ------------------------------
for s in cfg["slides"]:
    # 텍스트 토큰
    for token, meta in s.get("tokens", {}).items():
        with engine.connect() as conn:
            val = conn.execute(text(meta["sql"]), cfg["params"]).scalar()
            token_values[token] = val

    # 차트
    for chart_name, chart_conf in s.get("charts", {}).items():
        # (옵션) 이미지 트리맵 생성 분기 — PPT 내부 트리맵만 쓸 거면 이 분기 자체를 제거해도 됩니다.
        if "treemap_sql" in chart_conf:
            with engine.connect() as conn:
                rows = conn.execute(text(chart_conf["treemap_sql"]), cfg["params"]).fetchall()
                data = [(r[0], float(r[1]) if r[1] is not None else 0.0) for r in rows]

            outfile = chart_conf["outfile"]
            title = chart_conf.get("title", None)
            color_hex = chart_conf.get("color_hex", "#4682B4")
            generate_treemap_image(data, outfile, color_hex=color_hex, title=title)
            time.sleep(0.5)

            shape_name = chart_conf["shape"]
            image_map[shape_name] = outfile
            continue

        # (옵션) 히트맵 이미지
        if "heatmap_sql" in chart_conf:
            with engine.connect() as conn:
                df = pd.read_sql(text(chart_conf["heatmap_sql"]), conn, params=cfg["params"])
                data_df = df.set_index(df.columns[0])

            outfile = chart_conf["outfile"]
            title = chart_conf.get("title", None)
            generate_heatmap_image(data_df, outfile, title=title)
            time.sleep(0.5)

            shape_name = chart_conf["shape"]
            image_map[shape_name] = outfile
            continue

        # 일반 카테고리/시리즈 차트
        if "category_sql" in chart_conf:
            with engine.connect() as conn:
                categories = [r[0] for r in conn.execute(text(chart_conf["category_sql"]), cfg["params"]).fetchall()]
                series = {}
                for sname, ssql in chart_conf["series"].items():
                    series[sname] = [r[0] for r in conn.execute(text(ssql), cfg["params"]).fetchall()]
                chart_data[chart_name] = (categories, series)

print(f"DEBUG: 최종 Image Map: {image_map}")

# ------------------------------
# PPT 저장 (토큰/일반차트/이미지 적용)
# ------------------------------
apply_tokens_and_charts(
    prs_path=TEMPLATE_PPT,
    out_path=OUTPUT_PPT,
    token_map=token_values,
    chart_map=chart_data,
    image_map=image_map
)

import os, time
ppt_abs = os.path.abspath(OUTPUT_PPT)
if not os.path.exists(ppt_abs):
    raise FileNotFoundError(f"PPT not found: {ppt_abs}")
time.sleep(0.5)  # 방금 저장한 파일 I/O 안정화

# ------------------------------
# PPT 내부 트리맵 차트 (Win32 COM) 갱신
#   - 이미지 트리맵 대신, 벡터 품질 유지
# ------------------------------
params = {
    "REGION_CD": cfg["params"]["REGION_CD"],
    "DATE_FROM": cfg["params"]["DATE_FROM"],
    "DATE_TO":   cfg["params"]["DATE_TO"],
}

sql_treemap_foreigner = """
WITH topk AS (
  SELECT i.svc_induty_sclas_cd_nm AS child, SUM(t.FRGNR_SALAMT) AS amt
  FROM regionmonitor.TB_NATION_SELNG t
  JOIN regionmonitor.tb_svc_induty_sclas i
    ON i.svc_induty_sclas_cd = t.SVC_INDUTY_SCLAS_CD
  WHERE t.REGION_CD = CAST(:REGION_CD AS VARCHAR)
    AND t.STDR_YMD BETWEEN :DATE_FROM AND :DATE_TO
  GROUP BY i.svc_induty_sclas_cd_nm
  ORDER BY amt DESC, i.svc_induty_sclas_cd_nm
  LIMIT 10
)
SELECT '업종별 매출금액(만원)' AS series, '외국인' AS parent, child, ROUND(amt/10000, 1) AS value FROM topk;
"""

sql_treemap_native = """
WITH topk AS (
  SELECT i.svc_induty_sclas_cd_nm AS child, SUM(t.NATIVE_SALAMT) AS amt
  FROM regionmonitor.TB_NATION_SELNG t
  JOIN regionmonitor.tb_svc_induty_sclas i
    ON i.svc_induty_sclas_cd = t.SVC_INDUTY_SCLAS_CD
  WHERE t.REGION_CD = CAST(:REGION_CD AS VARCHAR)
    AND t.STDR_YMD BETWEEN :DATE_FROM AND :DATE_TO
  GROUP BY i.svc_induty_sclas_cd_nm
  ORDER BY amt DESC, i.svc_induty_sclas_cd_nm
  LIMIT 10
)
SELECT '업종별 매출금액(만원)' AS series, '내국인' AS parent, child, ROUND(amt/10000, 1) AS value FROM topk;
"""

with engine.connect() as conn:
    rows_f = conn.execute(text(sql_treemap_foreigner), params).fetchall()
    rows_foreigner = [(r.series, r.parent, r.child, float(r.value or 0)) for r in rows_f]

    rows_n = conn.execute(text(sql_treemap_native), params).fetchall()
    rows_native = [(r.series, r.parent, r.child, float(r.value or 0)) for r in rows_n]

# 실제 PPT 파일(OUTPUT_PPT)을 열어 차트 시트에 4열 입력
# ⚠️ PowerPoint에서 도형 이름이 정확히 일치해야 함
update_treemap_chart(
    ppt_path=ppt_abs,   # <-- 절대경로
    out_path=ppt_abs,
    shape_name="SL19_chart_foreigner",
    rows=rows_foreigner,
    value_header="억원"
)
update_treemap_chart(
    ppt_path=ppt_abs,
    out_path=ppt_abs,
    shape_name="SL19_chart_native",
    rows=rows_native,
    value_header="억원"
)


print("✅ 보고서 생성 완료:", OUTPUT_PPT)
