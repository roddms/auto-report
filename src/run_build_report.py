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

def plot_points_within_region(engine, region_cd, out_png, buffer_m=500, point_sql=None, title=None):
    """
    - region_cd의 폴리곤을 GeoJSON으로 받아 shapely로 그림 (read_postgis 미사용)
    - point_sql 은 이미 ST_Within + 500m 버퍼로 필터된 쿼리(위의 facility_sql / parking_sql)
    """
    import json
    import pandas as pd
    import geopandas as gpd
    import matplotlib.pyplot as plt
    from shapely.geometry import shape
    from sqlalchemy import text
    import os

    # 0) 파라미터
    params = {"REGION_CD": region_cd, "BUFFER_M": buffer_m}

    # 1) 폴리곤(버퍼 적용) GeoJSON으로 가져오기
    sql_poly = """
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
    with engine.connect() as conn:
        row = conn.execute(text(sql_poly), params).fetchone()
    if not row or not row[0]:
        print(f"⚠️ region_cd={region_cd} 폴리곤 없음")
        return
    poly = shape(json.loads(row[0]))
    gdf_region = gpd.GeoDataFrame(geometry=[poly], crs="EPSG:4326")

    # 2) 포인트(이미 SQL에서 버퍼 내부로 필터된 결과만)
    if point_sql is None:
        print("⚠️ point_sql 이 필요합니다.")
        return
    df_points = pd.read_sql(text(point_sql), engine, params=params)
    if df_points.empty:
        print("⚠️ 포인트 없음")
        return
    gdf_points = gpd.GeoDataFrame(
        df_points, geometry=gpd.points_from_xy(df_points["x"], df_points["y"]), crs="EPSG:4326"
    )

    # 3) 투영(Basemap용): 3857
    reg_3857   = gdf_region.to_crs(3857)
    points_3857 = gdf_points.to_crs(3857)

    # 4) 시각화
    fig, ax = plt.subplots(figsize=(8, 7))
    add_naver_or_osm_basemap(ax, 3857)
    reg_3857.boundary.plot(ax=ax, color="#005BAC", linewidth=2, alpha=0.9)
    points_3857.plot(ax=ax, color="#E74C3C", markersize=18, alpha=0.85, edgecolor="k", linewidth=0.3)

    ax.set_axis_off()
    if title:
        ax.set_title(title, fontsize=13, fontweight="bold", pad=6)

    os.makedirs(os.path.dirname(out_png), exist_ok=True)
    plt.savefig(out_png, dpi=300, bbox_inches="tight", pad_inches=0.1, transparent=True)
    plt.close(fig)

    print(f"✅ 지도 이미지 생성 완료 → {out_png} (표시 건수: {len(points_3857)})")



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

OUTPUT_PPT = "out/test_전년대비수정.pptx"
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

                for sname, ssql in chart_conf["series"].items():
                    rows = conn.execute(text(ssql), cfg["params"]).fetchall()

                    if chart_name == "SL20_chart" and rows and len(rows[0]) == 2:
                        vals  = [r[0] for r in rows]
                        flags = [int(r[1]) if r[1] is not None else 0 for r in rows]

                        if sname == "매출금액(천만원)":
                            amt_vals, amt_flags = vals, flags
                        elif sname == "매출건수(건)":
                            cnt_vals = vals
                        else:
                            series[sname] = vals  # 예외 시 기본 처리
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
                        series["매출금액(천만원)"] = amt_vals      # 막대(포인트별 색칠 대상)
                    if cnt_vals is not None:
                        series["매출건수(건)"]   = cnt_vals        # 라인(연속선)
                    if amt_flags is not None:
                        series["_festival_flags"] = amt_flags      # 메타(차트 데이터 아님)

                chart_data[chart_name] = (categories, series)



# 시설 지도
facility_sql = """
WITH reg AS (
  SELECT ST_Transform(
           ST_Buffer(ST_Transform(r.popltn_relm, 5179), :BUFFER_M),
           4326
         ) AS geom
  FROM regionmonitor.tb_intrst_region_relm r
  WHERE r.region_cd = :REGION_CD
)
SELECT 
  f.fclty_nm AS name,
  f.X_CRDNT  AS x,
  f.Y_CRDNT  AS y
FROM regionmonitor.TB_MAIN_FCLTY_INFO f
JOIN reg 
  ON ST_Within(
       ST_SetSRID(ST_MakePoint(f.X_CRDNT, f.Y_CRDNT), 4326),
       reg.geom
     )
WHERE f.X_CRDNT IS NOT NULL 
  AND f.Y_CRDNT IS NOT NULL;
"""

# 주차장 지도
parking_sql = """
WITH reg AS (
  SELECT ST_Transform(
           ST_Buffer(ST_Transform(r.popltn_relm, 5179), :BUFFER_M),
           4326
         ) AS geom
  FROM regionmonitor.tb_intrst_region_relm r
  WHERE r.region_cd = :REGION_CD
)
SELECT 
  p.prkplce_nm AS name,
  p.X_CRDNT    AS x,
  p.Y_CRDNT    AS y,
  p.prkcmprt_co AS slots
FROM regionmonitor.TB_PRKPLCE_INFO p
JOIN reg 
  ON ST_Within(
       ST_SetSRID(ST_MakePoint(p.X_CRDNT, p.Y_CRDNT), 4326),
       reg.geom
     )
WHERE p.X_CRDNT IS NOT NULL 
  AND p.Y_CRDNT IS NOT NULL;
"""

region_cd = cfg["params"]["REGION_CD"]

# 3-1) 시설 지도
plot_points_within_region(
    engine=engine,
    region_cd=region_cd,
    buffer_m=500,
    point_sql=facility_sql,  # ← 단일 SQL (위 CTE 버전)
    out_png="out/img/facility_map.png",
    title="관심영역 500m 내 주요 시설"
)
image_map["SL22_map_facility"] = "out/img/facility_map.png"

# 3-2) 주차장 지도
plot_points_within_region(
    engine=engine,
    region_cd=region_cd,
    buffer_m=500,
    point_sql=parking_sql,   # ← 단일 SQL (위 CTE 버전)
    out_png="out/img/parking_map.png",
    title="관심영역 500m 내 주차장"
)
image_map["SL23_map_parking"] = "out/img/parking_map.png"

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
