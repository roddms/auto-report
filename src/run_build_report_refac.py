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
from shapely.geometry import shape
from sqlalchemy import create_engine, text
from dotenv import load_dotenv, find_dotenv
from pptx import Presentation

from ppt_fillers import (
    apply_tokens_and_charts,
    update_treemaps_batch,
    find_table,
    fill_table_with_padding
)

# ------------------------------
# 히트맵 이미지 유틸
# ------------------------------
def generate_heatmap_image(data_df, out_path, title=None, font_family="Malgun Gothic"):
    import seaborn as sns
    import os
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
# Contextily 타일 캐시 (베이스맵)
# ------------------------------
BASEMAP_IMG_CACHE = {}

def get_basemap_img(bounds_3857, zoom=16):
    import contextily as ctx
    key = (tuple(round(x, 1) for x in bounds_3857), zoom)
    if key in BASEMAP_IMG_CACHE:
        return BASEMAP_IMG_CACHE[key]
    xmin, ymin, xmax, ymax = bounds_3857
    img, ext = ctx.bounds2img(xmin, ymin, xmax, ymax, zoom=zoom, source=ctx.providers.CartoDB.Positron)
    BASEMAP_IMG_CACHE[key] = (img, ext)
    return img, ext

# ------------------------------
# 그룹별 시설 지도 (REGION_WKT 사용)
# ------------------------------
def plot_facility_group_map(engine, region_cd, group_name, out_png,
                            region_wkt, buffer_m=500, title=None):
    import geopandas as gpd
    import json

    GROUP_MAP = {
        "그룹1":       ['G','H','J','M','N'],
        "그룹2(개별)": ['D'],
        "그룹3(개별)": ['E'],
        "그룹4(개별)": ['O'],
        "그룹5(개별)": ['P'],   # 주차장
        "그룹6":       ['B','F','I','L'],
        "그룹7(개별)": ['A'],
        "그룹8":       ['C','K'],
    }
    letters = GROUP_MAP[group_name]

    # REGION_WKT를 그대로 사용하여 영역 가져오기
    sql_region = """
    SELECT ST_AsGeoJSON(ST_GeomFromText(:REGION_WKT, 4326)) AS gj;
    """
    # 시설도 REGON_WKT로 공간조인
    sql_fac = """
    WITH reg AS (
      SELECT ST_GeomFromText(:REGION_WKT, 4326) AS geom
    )
    SELECT
      f.fclty_nm AS name,
      f.Y_CRDNT  AS x,      -- lon
      f.X_CRDNT  AS y,      -- lat
      SUBSTRING(TRIM(f.fclty_sclas_cd) FROM 1 FOR 1) AS lclas
    FROM regionmonitor.TB_MAIN_FCLTY_INFO f, reg
    WHERE ST_Within(ST_SetSRID(ST_MakePoint(f.Y_CRDNT,f.X_CRDNT),4326), reg.geom)
      AND SUBSTRING(TRIM(f.fclty_sclas_cd) FROM 1 FOR 1) = ANY(:LETTERS);
    """

    params = {"REGION_WKT": region_wkt, "LETTERS": letters}
    with engine.begin() as conn:
        row = conn.execute(text(sql_region), params).fetchone()
        region = shape(json.loads(row[0]))
        df = pd.read_sql(text(sql_fac), conn, params=params)

    gdf_region = gpd.GeoDataFrame(geometry=[region], crs="EPSG:4326")
    gdf = gpd.GeoDataFrame(df, geometry=gpd.points_from_xy(df["x"], df["y"]), crs="EPSG:4326")
    reg3857 = gdf_region.to_crs(3857)
    fac3857 = gdf.to_crs(3857)

    PALETTE = ["#1C75B3", "#FF7F0E", "#279D27", "#D62829", "#9467BD"]
    order_hint = ["G","H","J","M","N","D","E","O","P","B","F","I","L","A","C","K"]
    ordered_letters = [c for c in order_hint if c in letters]
    color_by_pos = {lc: PALETTE[i % len(PALETTE)] for i, lc in enumerate(ordered_letters)}
    label_map = {
        'G':'철도역','H':'버스정류장','J':'공항','M':'지하철','N':'터미널',
        'D':'상가업소','E':'의료기관','O':'공중화장실','P':'주차장',
        'B':'관공서','F':'공공기관','I':'교육기관','L':'은행점포',
        'A':'숙박시설','C':'문화여가시설','K':'상영관'
    }

    fig, ax = plt.subplots(figsize=(8,7))

    # bbox & basemap (캐시)
    xmin, ymin, xmax, ymax = reg3857.total_bounds
    pad = max((xmax-xmin), (ymax-ymin)) * 0.05
    bounds = (xmin - pad, ymin - pad, xmax + pad, ymax + pad)
    ax.set_xlim(bounds[0], bounds[2])
    ax.set_ylim(bounds[1], bounds[3])

    img, ext = get_basemap_img(bounds, zoom=16)
    ax.imshow(img, extent=ext, interpolation="bilinear", zorder=0)

    # points
    for lc, gsub in fac3857.groupby("lclas"):
        gsub.plot(
            ax=ax,
            markersize=26,
            color=color_by_pos.get(lc, "#666"),
            alpha=0.95,
            linewidth=0.2,
            edgecolor="#1a1a1a",
            label=label_map.get(lc, lc),
            zorder=10
        )

    # 경계선(버퍼선)
    reg3857.boundary.plot(ax=ax, color="#767171", linewidth=2, zorder=11)
    ax.set_axis_off()
    if title:
        ax.set_title(title, fontsize=12, fontweight="500", pad=6)

    os.makedirs(os.path.dirname(out_png), exist_ok=True)
    plt.savefig(out_png, dpi=200, bbox_inches="tight", pad_inches=0.1, transparent=True)
    plt.close(fig)
    print(f"✅ {region_cd} 시설 지도 생성 완료 → {out_png}  (시설:{len(fac3857)})")

# ------------------------------
# 환경/DB 설정
# ------------------------------
load_dotenv(find_dotenv(), override=True)
db_url = os.getenv("DB_URL")
engine = create_engine(
    db_url,
    connect_args={"options": "-csearch_path=regionmonitor,public"}
)

# ✅ REGION_WKT 미리 계산 (한 번만)
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

with open("config/slides_tokens.yml", encoding="utf-8") as f:
    cfg = yaml.safe_load(f)

token_values = {}
chart_data = {}
image_map = {}

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
BUFFER_M = int(cfg["params"].get("BUFFER_M", 500))

# REGION_WKT 계산
with engine.begin() as conn:
    region_wkt = conn.execute(
        text(SQL_REGION_WKT),
        {"REGION_CD": cfg["params"]["REGION_CD"], "BUFFER_M": BUFFER_M}
    ).scalar()

if not region_wkt:
    raise RuntimeError("⚠️ REGION_WKT 생성 실패 — REGION_CD 확인 필요")
print("✅ REGION_WKT 계산 완료")

# 이후 모든 쿼리에 공통 파라미터로 사용
PARAMS = {**cfg["params"], "REGION_WKT": region_wkt}

OUTPUT_PPT = f"out/report_{args.REGION_CD}.pptx"
TEMPLATE_PPT = "template/master_pretendard.pptx"

# ------------------------------
# 토큰 / 차트 (커넥션 재사용)
# ------------------------------
with engine.begin() as conn:
    for s in cfg["slides"]:
        # 텍스트 토큰
        for token, meta in s.get("tokens", {}).items():
            val = conn.execute(text(meta["sql"]), PARAMS).scalar()
            token_values[token] = val

        # 차트
        for chart_name, chart_conf in s.get("charts", {}).items():
            # 히트맵 이미지
            if "heatmap_sql" in chart_conf:
                df = pd.read_sql(text(chart_conf["heatmap_sql"]), conn, params=PARAMS)
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
                categories = [r[0] for r in conn.execute(text(chart_conf["category_sql"]), PARAMS).fetchall()]

                if not categories:
                    print(f"⚠️ chart '{chart_name}' 카테고리 결과가 비어 건너뜀 (REGION_CD={cfg['params']['REGION_CD']})")
                    continue

                series = OrderedDict()
                flags_from_cats = [1 if str(lbl).startswith("DAY") else 0 for lbl in categories]

                for sname, ssql in chart_conf["series"].items():
                    rows = conn.execute(text(ssql), PARAMS).fetchall()

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
# 지도 생성 (그룹 1~8 페이지) - REGION_WKT 사용
# ------------------------------
region_cd = cfg["params"]["REGION_CD"]

os.makedirs(f"out/img/{region_cd}", exist_ok=True)

plot_facility_group_map(engine, region_cd, "그룹1", out_png=f"out/img/{region_cd}/group1_map.png", region_wkt=region_wkt, buffer_m=BUFFER_M)
image_map["SL_G1_MAP"] = f"out/img/{region_cd}/group1_map.png"

plot_facility_group_map(engine, region_cd, "그룹2(개별)", out_png=f"out/img/{region_cd}/group2_map.png", region_wkt=region_wkt, buffer_m=BUFFER_M)
image_map["SL_G2_MAP"] = f"out/img/{region_cd}/group2_map.png"

plot_facility_group_map(engine, region_cd, "그룹3(개별)", out_png=f"out/img/{region_cd}/group3_map.png", region_wkt=region_wkt, buffer_m=BUFFER_M)
image_map["SL_G3_MAP"] = f"out/img/{region_cd}/group3_map.png"

plot_facility_group_map(engine, region_cd, "그룹4(개별)", out_png=f"out/img/{region_cd}/group4_map.png", region_wkt=region_wkt, buffer_m=BUFFER_M)
image_map["SL_G4_MAP"] = f"out/img/{region_cd}/group4_map.png"

plot_facility_group_map(engine, region_cd, "그룹5(개별)", out_png=f"out/img/{region_cd}/group5_map.png", region_wkt=region_wkt, buffer_m=BUFFER_M)
image_map["SL_G5_MAP"] = f"out/img/{region_cd}/group5_map.png"

plot_facility_group_map(engine, region_cd, "그룹6", out_png=f"out/img/{region_cd}/group6_map.png", region_wkt=region_wkt, buffer_m=BUFFER_M)
image_map["SL_G6_MAP"] = f"out/img/{region_cd}/group6_map.png"

plot_facility_group_map(engine, region_cd, "그룹7(개별)", out_png=f"out/img/{region_cd}/group7_map.png", region_wkt=region_wkt, buffer_m=BUFFER_M)
image_map["SL_G7_MAP"] = f"out/img/{region_cd}/group7_map.png"

plot_facility_group_map(engine, region_cd, "그룹8", out_png=f"out/img/{region_cd}/group8_map.png", region_wkt=region_wkt, buffer_m=BUFFER_M)
image_map["SL_G8_MAP"] = f"out/img/{region_cd}/group8_map.png"

print(f"DEBUG: 최종 Image Map: {image_map}")

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

# ------------------------------
# 주차장 표 채우기 (REGION_WKT 사용)
# ------------------------------
parking_list_sql = """
WITH reg AS (
  SELECT ST_GeomFromText(:REGION_WKT, 4326) AS geom
)
SELECT
  p.prkplce_nm AS name,
  COALESCE(p.prkcmprt_co,0)::int AS slots
FROM regionmonitor.TB_PRKPLCE_INFO p
JOIN reg
  ON ST_Within(ST_SetSRID(ST_MakePoint(p.Y_CRDNT, p.X_CRDNT), 4326), reg.geom)
WHERE p.X_CRDNT IS NOT NULL AND p.Y_CRDNT IS NOT NULL
ORDER BY slots DESC, name;
"""
with engine.begin() as conn:
    dfp = pd.read_sql(text(parking_list_sql), conn, params=PARAMS)

rows_for_table = [(row["name"], f"{int(row['slots']):,}대") for _, row in dfp.iterrows()]
prs = Presentation(OUTPUT_PPT)
tbl = find_table(prs, "SL23_table_parking")
if tbl is not None:
    fill_table_with_padding(tbl, rows_for_table)
    prs.save(OUTPUT_PPT)
else:
    print("⚠️ 도형 'SL23_table_parking' 표를 찾을 수 없습니다.")

print("✅ 보고서 생성 완료:", OUTPUT_PPT)
