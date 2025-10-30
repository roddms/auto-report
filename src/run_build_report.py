# src/run_build_report.py
import os
import math
import time
import yaml
import pandas as pd
import warnings
import geopandas as gpd
warnings.filterwarnings("ignore", category=UserWarning)

import matplotlib.pyplot as plt
import contextily as ctx

from collections import OrderedDict
from shapely.geometry import Point
from sqlalchemy import create_engine, text
from dotenv import load_dotenv, find_dotenv
from shapely.geometry import shape
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
    ax.set_xlabel('')
    ax.set_ylabel('')

    if title:
        ax.set_title(title, pad=10, fontsize=12, fontweight='bold')

    plt.tight_layout()
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    plt.savefig(out_path, dpi=300, bbox_inches='tight', pad_inches=0.1, transparent=True)
    plt.close(fig)


def add_naver_or_osm_basemap(ax, crs_epsg: int):
    import contextily as ctx
    try:
        naver_url = "https://map.pstatic.net/nrb/styles/basic/1477/{z}/{x}/{y}.png?mt=bg.ol.ts.ar.lko"
        src = ctx.TileClient(naver_url)
        ctx.add_basemap(ax, source=src, crs=f"EPSG:{crs_epsg}")
    except Exception:
        ctx.add_basemap(ax, crs=f"EPSG:{crs_epsg}", source=ctx.providers.OpenStreetMap.Mapnik)

def plot_facility_and_parking(engine, region_cd, out_png, buffer_m=500, title=None):
    import json, os
    import pandas as pd
    import geopandas as gpd
    import matplotlib.pyplot as plt
    from shapely.geometry import shape
    from sqlalchemy import text

    # SQL 정의
    sql_region = """
    WITH reg AS (
      SELECT ST_Transform(
               ST_Buffer(ST_Transform(r.popltn_relm, 5179), :BUFFER_M),
               4326
             ) AS geom
      FROM regionmonitor.tb_intrst_region_relm r
      WHERE r.region_cd = :REGION_CD
    )
    SELECT ST_AsGeoJSON(geom) AS gj FROM reg;
    """
    sql_facility = """
    WITH reg AS (
      SELECT ST_Transform(ST_Buffer(ST_Transform(r.popltn_relm, 5179), :BUFFER_M), 4326) AS geom
      FROM regionmonitor.tb_intrst_region_relm r
      WHERE r.region_cd = :REGION_CD
    )
    SELECT
      f.fclty_nm AS name,
      f.Y_CRDNT  AS x,
      f.X_CRDNT  AS y,
      f.fclty_sclas_cd AS type
    FROM regionmonitor.TB_MAIN_FCLTY_INFO f
    JOIN reg
      ON ST_Within(ST_SetSRID(ST_MakePoint(f.Y_CRDNT, f.X_CRDNT), 4326), reg.geom)
    WHERE f.X_CRDNT IS NOT NULL AND f.Y_CRDNT IS NOT NULL;
    """
    sql_parking = """
    WITH reg AS (
      SELECT ST_Transform(ST_Buffer(ST_Transform(r.popltn_relm, 5179), :BUFFER_M), 4326) AS geom
      FROM regionmonitor.tb_intrst_region_relm r
      WHERE r.region_cd = :REGION_CD
    )
    SELECT
      p.prkplce_nm AS name,
      p.Y_CRDNT    AS x,
      p.X_CRDNT    AS y, 
      p.prkcmprt_co AS slots
    FROM regionmonitor.TB_PRKPLCE_INFO p
    JOIN reg
      ON ST_Within(ST_SetSRID(ST_MakePoint(p.Y_CRDNT, p.X_CRDNT), 4326), reg.geom)
    WHERE p.X_CRDNT IS NOT NULL AND p.Y_CRDNT IS NOT NULL;
    """

    params = {"REGION_CD": region_cd, "BUFFER_M": buffer_m}
    with engine.connect() as conn:
        row = conn.execute(text(sql_region), params).fetchone()
        poly = shape(json.loads(row[0]))
        gdf_region = gpd.GeoDataFrame(geometry=[poly], crs="EPSG:4326")

        df_fac = pd.read_sql(text(sql_facility), conn, params=params)
        df_par = pd.read_sql(text(sql_parking), conn, params=params)

    if df_fac.empty and df_par.empty:
        print("⚠️ 시설/주차장 모두 없음")
        return

    gdf_fac = gpd.GeoDataFrame(df_fac, geometry=gpd.points_from_xy(df_fac["x"], df_fac["y"]), crs="EPSG:4326")
    gdf_par = gpd.GeoDataFrame(df_par, geometry=gpd.points_from_xy(df_par["x"], df_par["y"]), crs="EPSG:4326")

    # 투영
    reg3857 = gdf_region.to_crs(3857)
    fac3857 = gdf_fac.to_crs(3857)
    par3857 = gdf_par.to_crs(3857)

    # 색상 팔레트 (시설 코드별)
    unique_types = fac3857["type"].unique().tolist()
    cmap = plt.get_cmap("tab10")
    color_map = {t: cmap(i % 10) for i, t in enumerate(unique_types)}

    fig, ax = plt.subplots(figsize=(8, 7))

    # 베이스맵 마지막에 추가
    #reg3857.boundary.plot(ax=ax, color="#005BAC", linewidth=2, alpha=0.8, zorder=5)

    # 시설 (분류별 색상)
    for t, g in fac3857.groupby("type"):
        g.plot(ax=ax, markersize=25, color=color_map[t], alpha=0.8, label=f"시설:{t}", zorder=10)

    # 주차장
    #par3857.plot(ax=ax, color="black", markersize=30, marker="P", alpha=0.8, label="주차장", zorder=9)

    # 범례/제목/축
    ax.legend(loc="lower left", fontsize=8, frameon=True)
    ax.set_axis_off()
    if title:
        ax.set_title(title, fontsize=13, fontweight="bold", pad=6)

    from contextily import add_basemap, providers
    add_basemap(ax, source=providers.CartoDB.Positron, crs=3857)

    import os
    os.makedirs(os.path.dirname(out_png), exist_ok=True)
    plt.savefig(out_png, dpi=300, bbox_inches="tight", pad_inches=0.1, transparent=True)
    plt.close(fig)

    print(f"✅ 지도 이미지 생성 완료 → {out_png}")

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

OUTPUT_PPT = "out/test_1653.pptx"
TEMPLATE_PPT = "template/master.pptx"

token_values = {}
chart_data = {}
image_map = {}

# ------------------------------
# 토큰 / 일반 차트 / 이미지 차트 생성
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

                # ------- SL20 전용 임시 변수 -------
                amt_vals, amt_flags = None, None
                cnt_vals = None
                # -----------------------------------

                series = OrderedDict()

                flags_from_cats = [1 if str(lbl).startswith("DAY") else 0 for lbl in categories]

                for sname, ssql in chart_conf["series"].items():
                    rows = conn.execute(text(ssql), cfg["params"]).fetchall()

                    if chart_name == "SL20_chart":
                        vals = [r[0] for r in rows]  # 첫 컬럼만 사용
                        if sname == "매출금액(백만원)":
                            series["매출금액(백만원)"] = vals       # 막대(색칠 대상)
                            series["_festival_flags"] = flags_from_cats  # ← 여기서 확정
                        elif sname == "매출건수(건)":
                            series["매출건수(건)"] = vals           # 라인
                        else:
                            series[sname] = vals
                    else:
                        # 일반(단일 컬럼) 시리즈
                        series[sname] = [r[0] for r in rows]

                    if chart_name == "SL21_chart" and rows and len(rows[0]) == 2:
                        vals  = [r[0] for r in rows]
                        flags = [int(r[1]) if r[1] is not None else 0 for r in rows]

                        if sname == "방문인구(명)":
                            series["방문인구(명)"] = vals
                            series["_festival_flags"] = flags   # 메타(차트데이터로는 넣지 않음)
                        else:
                            series[sname] = vals
                    else:
                        series[sname] = [r[0] for r in rows]

                if chart_name == "SL20_chart":
                    # ⬇️ 콤보차트 타입 유지: 0번(막대)=금액, 1번(라인)=건수
                    if amt_vals is not None:
                        series["매출금액(백만원)"] = amt_vals      # 막대(포인트별 색칠 대상)
                    if cnt_vals is not None:
                        series["매출건수(건)"]   = cnt_vals        # 라인(연속선)
                    if amt_flags is not None:
                        series["_festival_flags"] = amt_flags      # 메타(차트 데이터 아님)

                chart_data[chart_name] = (categories, series)



# # 시설 지도
# facility_sql = """
# WITH reg AS (
#   SELECT ST_Transform(
#            ST_Buffer(ST_Transform(r.popltn_relm, 5179), :BUFFER_M),
#            4326
#          ) AS geom
#   FROM regionmonitor.tb_intrst_region_relm r
#   WHERE r.region_cd = :REGION_CD
# )
# SELECT
#   f.fclty_nm AS name,
#   /* x=경도(lon), y=위도(lat)로 맞춰서 반환 */
#   f.Y_CRDNT  AS x,
#   f.X_CRDNT  AS y,
#   f.fclty_sclas_cd AS type
# FROM regionmonitor.TB_MAIN_FCLTY_INFO f
# JOIN reg
#   ON ST_Within(
#        /* lon, lat 순서로 포인트 생성 */
#        ST_SetSRID(ST_MakePoint(f.Y_CRDNT, f.X_CRDNT), 4326),
#        reg.geom
#      )
# WHERE f.X_CRDNT IS NOT NULL
#   AND f.Y_CRDNT IS NOT NULL;

# """

# # 주차장 지도
# parking_sql = """
# WITH reg AS (
#   SELECT ST_Transform(
#            ST_Buffer(ST_Transform(r.popltn_relm, 5179), :BUFFER_M),
#            4326
#          ) AS geom
#   FROM regionmonitor.tb_intrst_region_relm r
#   WHERE r.region_cd = :REGION_CD
# )
# SELECT
#   p.prkplce_nm AS name,
#   /* x=경도(lon), y=위도(lat) */
#   p.Y_CRDNT    AS x,
#   p.X_CRDNT    AS y,
#   p.prkcmprt_co AS slots
# FROM regionmonitor.TB_PRKPLCE_INFO p
# JOIN reg
#   ON ST_Within(
#        ST_SetSRID(ST_MakePoint(p.Y_CRDNT, p.X_CRDNT), 4326),
#        reg.geom
#      )
# WHERE p.X_CRDNT IS NOT NULL
#   AND p.Y_CRDNT IS NOT NULL;
# """

region_cd = cfg["params"]["REGION_CD"]

# 3) 시설+주차장 지도
plot_facility_and_parking(
    engine=engine,
    region_cd=region_cd,
    buffer_m=500,
    out_png="out/img/facility_parking_map.png",
    title="500m내 주요 시설 및 주차장"
)
image_map["SL22_map_facility"] = "out/img/facility_parking_map.png"

# # 3-2) 주차장 지도
# plot_facility_and_parking(
#     engine=engine,
#     region_cd=region_cd,
#     buffer_m=500,
#     out_png="out/img/parking_map.png",
#     title="관심영역 500m 내 주차장"
# )
# image_map["SL23_map_parking"] = "out/img/parking_map.png"

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
    value_header="백만원"
)
update_treemap_chart(
    ppt_path=ppt_abs,
    out_path=ppt_abs,
    shape_name="SL19_chart_native",
    rows=rows_native,
    value_header="백만원"
)


print("✅ 보고서 생성 완료:", OUTPUT_PPT)
