# src/run_build_report.py
import os
import time
import yaml
import warnings
warnings.filterwarnings("ignore", category=UserWarning)

import pandas as pd
import matplotlib.pyplot as plt
import argparse

from collections import OrderedDict
from sqlalchemy import create_engine, text
from dotenv import load_dotenv, find_dotenv

from ppt_fillers import (
    apply_tokens_and_charts,
    update_treemaps_batch,
)

# ------------------------------
# 히트맵 이미지 유틸
# ------------------------------
def generate_heatmap_image(data_df, out_path, title=None, font_family="Malgun Gothic"):
    import seaborn as sns
    import numpy as np

    # 데이터 없음 안전 처리
    if (
        data_df is None
        or data_df.size == 0
        or data_df.dropna(how="all").empty
        or np.all(pd.isna(data_df.values))
    ):
        fig, ax = plt.subplots(figsize=(8, 6))
        fig.set_facecolor('none')
        ax.axis('off')
        msg = (title + "\n") if title else ""
        ax.text(0.5, 0.5, f"{msg}(데이터 없음)", ha='center', va='center',
                fontsize=12, fontweight='bold')
        os.makedirs(os.path.dirname(out_path), exist_ok=True)
        plt.tight_layout()
        plt.savefig(out_path, dpi=300, bbox_inches='tight', pad_inches=0.2, transparent=True)
        plt.close(fig)
        return

    try:
        plt.rcParams["font.family"] = font_family
        plt.rcParams["axes.unicode_minus"] = False
    except:
        pass

    fig, ax = plt.subplots(figsize=(8, 6))
    fig.set_facecolor('none')

    annot_data = data_df.map(lambda x: f'{x:.1f}%')
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
    ax.set_xlabel(''); ax.set_ylabel('')
    if title:
        ax.set_title(title, pad=10, fontsize=12, fontweight='bold')

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    plt.tight_layout()
    plt.savefig(out_path, dpi=300, bbox_inches='tight', pad_inches=0.1, transparent=True)
    plt.close(fig)

# ------------------------------
# 환경/DB 설정
# ------------------------------
load_dotenv(find_dotenv(), override=True)
db_url = os.getenv("DB_URL")
engine = create_engine(
    db_url,
    connect_args={"options": "-csearch_path=regionmonitor,public"}
)

with open("config/slides_tokens.yml", encoding="utf-8") as f:
    cfg = yaml.safe_load(f)

token_values = {}
chart_data = {}
image_map = {}  # 히트맵 이미지만 담김

# ------------------------------
# 인자 파싱 & 파라미터 반영
# ------------------------------
parser = argparse.ArgumentParser()
parser.add_argument("--REGION_CD", required=True)
parser.add_argument("--DATE_FROM", required=True)
parser.add_argument("--DATE_TO", required=True)
args = parser.parse_args()

cfg["params"]["REGION_CD"] = args.REGION_CD
cfg["params"]["DATE_FROM"] = args.DATE_FROM
cfg["params"]["DATE_TO"]   = args.DATE_TO

OUTPUT_PPT = f"out/report_{args.REGION_CD}.pptx"
TEMPLATE_PPT = "template/master_pretendard.pptx"

SQL_REGION_WKT = """
SELECT ST_AsText(
         ST_Transform(
           ST_Buffer(ST_Transform(r.popltn_relm, 5179), :BUFFER_M),
           4326
         )
       ) AS wkt
FROM regionmonitor.tb_intrst_region_relm r
WHERE r.region_cd = :REGION_CD;
"""

# BUFFER_M 기본값 보정 (yml에 없으면 500 사용)
BUFFER_M = int(cfg["params"].get("BUFFER_M", 500))

# REGION_WKT 계산해서 params에 주입
with engine.begin() as conn:
    region_wkt = conn.execute(
        text(SQL_REGION_WKT),
        {"REGION_CD": cfg["params"]["REGION_CD"], "BUFFER_M": BUFFER_M}
    ).scalar()

if not region_wkt:
    raise RuntimeError("REGION_WKT 생성 실패: REGION_CD/영역 데이터 확인 필요")

cfg["params"]["BUFFER_M"]   = BUFFER_M   # 혹시 yml에 없으면 넣어둠
cfg["params"]["REGION_WKT"] = region_wkt
print("✅ REGION_WKT 계산 및 파라미터 주입 완료")

# ------------------------------
# 토큰 / 차트 (커넥션 재사용)
# ------------------------------
with engine.begin() as conn:
    for s in cfg["slides"]:
        # 텍스트 토큰
        for token, meta in s.get("tokens", {}).items():
            val = conn.execute(text(meta["sql"]), cfg["params"]).scalar()
            token_values[token] = val

        # 차트
        for chart_name, chart_conf in s.get("charts", {}).items():

            # 히트맵 이미지
            if "heatmap_sql" in chart_conf:
                df = pd.read_sql(text(chart_conf["heatmap_sql"]), conn, params=cfg["params"])

                if df is None or df.empty or df.shape[1] < 2:
                    print(f"⚠️ heatmap '{chart_name}' 데이터 없음 → 스킵 (REGION_CD={cfg['params']['REGION_CD']})")
                    continue

                data_df = df.set_index(df.columns[0])
                outfile = chart_conf["outfile"]
                title = chart_conf.get("title", None)
                generate_heatmap_image(data_df, outfile, title=title)
                time.sleep(0.2)
                shape_name = chart_conf["shape"]
                image_map[shape_name] = outfile
                continue

            # 일반 카테고리/시리즈 차트
            if "category_sql" in chart_conf:
                cat_rows = conn.execute(text(chart_conf["category_sql"]), cfg["params"]).fetchall()
                categories = [r[0] for r in cat_rows]

                if not categories:
                    print(f"⚠️ chart '{chart_name}' 카테고리 결과가 비어 건너뜀 (REGION_CD={cfg['params']['REGION_CD']})")
                    continue

                series = OrderedDict()
                flags_from_cats = [1 if str(lbl).startswith("DAY") else 0 for lbl in categories]

                for sname, ssql in chart_conf["series"].items():
                    rows = conn.execute(text(ssql), cfg["params"]).fetchall()

                    if chart_name == "SL20_chart":
                        vals = [r[0] for r in rows]
                        if sname == "매출금액(백만원)":
                            series["매출금액(백만원)"] = vals
                            series["_festival_flags"] = flags_from_cats
                        elif sname == "매출건수(건)":
                            series["매출건수(건)"] = vals
                        else:
                            series[sname] = vals
                    elif chart_name == "SL21_chart" and rows and len(rows[0]) == 2:
                        vals  = [r[0] for r in rows]
                        flags = [int(r[1]) if r[1] is not None else 0 for r in rows]
                        if sname == "방문인구(명)":
                            series["방문인구(명)"] = vals
                            series["_festival_flags"] = flags
                        else:
                            series[sname] = vals
                    else:
                        series[sname] = [r[0] for r in rows]

                chart_data[chart_name] = (categories, series)

# ------------------------------
# PPT 저장 (토큰/차트/이미지)
# ------------------------------
apply_tokens_and_charts(
    prs_path=TEMPLATE_PPT,
    out_path=OUTPUT_PPT,
    token_map=token_values,
    chart_map=chart_data,
    image_map=image_map
)

ppt_abs = os.path.abspath(OUTPUT_PPT)
if not os.path.exists(ppt_abs):
    raise FileNotFoundError(f"PPT not found: {ppt_abs}")
time.sleep(0.3)

# ------------------------------
# 트리맵(내부 차트) 배치 갱신
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
SELECT '업종별 매출금액(백만원)' AS series, '외국인' AS parent, child, ROUND(amt/1000000, 1) AS value FROM topk;
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
SELECT '업종별 매출금액(백만원)' AS series, '내국인' AS parent, child, ROUND(amt/1000000, 1) AS value FROM topk;
"""

with engine.begin() as conn:
    rows_f = conn.execute(text(sql_treemap_foreigner), params).fetchall()
    rows_foreigner = [(r.series, r.parent, r.child, float(r.value or 0)) for r in rows_f]
    rows_n = conn.execute(text(sql_treemap_native), params).fetchall()
    rows_native = [(r.series, r.parent, r.child, float(r.value or 0)) for r in rows_n]

updates = [
    ("SL19_chart_foreigner", rows_foreigner, "백만원"),
    ("SL19_chart_native",    rows_native,    "백만원"),
]
update_treemaps_batch(
    ppt_path=ppt_abs,
    out_path=ppt_abs,
    updates=updates
)

print("✅ 보고서 생성 완료:", OUTPUT_PPT)