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
    """
    관심영역(=tb_intrst_region_relm.popltn_relm) 주변 buffer_m 내
    - 시설(대분류→8개 그룹 매핑, 그룹별 색상)
    - 주차장(그룹5로 고정, P마커)
    를 한 장의 지도에 함께 표시한다.

    좌표 스키마 주의: x=경도=Y_CRDNT, y=위도=X_CRDNT
    버퍼 경계선(파란 원)은 표시하지 않음.
    """
    import json, os
    import pandas as pd
    import geopandas as gpd
    import matplotlib.pyplot as plt
    from shapely.geometry import shape
    from sqlalchemy import text
    from contextily import add_basemap, providers

    # 1) 버퍼 영역(4326) GeoJSON
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

    # 2) 시설: 대분류명 → 8개 그룹 매핑 포함
    sql_facility = """
    WITH reg AS (
        SELECT ST_Transform(ST_Buffer(ST_Transform(r.popltn_relm, 5179), :BUFFER_M), 4326) AS geom
        FROM regionmonitor.tb_intrst_region_relm r
        WHERE r.region_cd = :REGION_CD
        ),
        -- 코드(A~P) 기준으로 그룹/순서/대분류명 매핑
        code_map(code, lclas_nm, group_nm, group_ord) AS (VALUES
        ('A','숙박시설'     ,'그룹7(개별)',8),
        ('B','관공서'       ,'그룹6',2),
        ('C','문화여가시설' ,'그룹8',3),
        ('D','상가업소'     ,'그룹2(개별)',4),
        ('E','의료기관'     ,'그룹3(개별)',5),
        ('F','공공기관'     ,'그룹6',2),
        ('G','철도역'       ,'그룹1',1),
        ('H','버스정류장'   ,'그룹1',1),
        ('I','교육기관'     ,'그룹6',2),
        ('J','공항'         ,'그룹1',1),
        ('K','상영관'       ,'그룹8',3),
        ('L','은행점포'     ,'그룹6',2),
        ('M','지하철'       ,'그룹1',1),
        ('N','터미널'       ,'그룹1',1),
        ('O','공중화장실'   ,'그룹4(개별)',6),
        ('P','주차장'       ,'그룹5(개별)',7)   -- 주차장은 아래 WHERE에서 제외
        )
        SELECT
        f.fclty_nm                                  AS name,
        f.Y_CRDNT                                   AS x,       -- 경도(lon)
        f.X_CRDNT                                   AS y,       -- 위도(lat)
        cm.lclas_nm                                 AS lclas_nm,
        cm.group_nm                                 AS group_nm,
        cm.group_ord                                AS group_ord
        FROM regionmonitor.TB_MAIN_FCLTY_INFO f
        JOIN reg
        ON ST_Within(ST_SetSRID(ST_MakePoint(f.Y_CRDNT, f.X_CRDNT), 4326), reg.geom)
        LEFT JOIN code_map cm
        ON cm.code = SUBSTRING(TRIM(f.fclty_sclas_cd) FROM 1 FOR 1)
        WHERE f.X_CRDNT IS NOT NULL
        AND f.Y_CRDNT IS NOT NULL
        AND SUBSTRING(TRIM(f.fclty_sclas_cd) FROM 1 FOR 1) <> 'P';
    """

    # 3) 주차장: 그룹5(개별)로 고정
    sql_parking = """
    WITH reg AS (
      SELECT ST_Transform(ST_Buffer(ST_Transform(r.popltn_relm, 5179), :BUFFER_M), 4326) AS geom
      FROM regionmonitor.tb_intrst_region_relm r
      WHERE r.region_cd = :REGION_CD
    )
    SELECT
      p.prkplce_nm                                AS name,
      p.Y_CRDNT                                   AS x,      -- 경도(lon)
      p.X_CRDNT                                   AS y,      -- 위도(lat)
      '주차장'                                     AS lclas_nm,
      '그룹5(개별)'                                 AS group_nm,
      7                                            AS group_ord
    FROM regionmonitor.TB_PRKPLCE_INFO p
    JOIN reg
      ON ST_Within(ST_SetSRID(ST_MakePoint(p.Y_CRDNT, p.X_CRDNT), 4326), reg.geom)
    WHERE p.X_CRDNT IS NOT NULL AND p.Y_CRDNT IS NOT NULL;
    """

    params = {"REGION_CD": region_cd, "BUFFER_M": buffer_m}
    with engine.connect() as conn:
        gj = conn.execute(text(sql_region), params).scalar()
        if not gj:
            print("⚠️ 관심영역 폴리곤을 찾을 수 없습니다.")
            return
        poly = shape(json.loads(gj))
        df_fac = pd.read_sql(text(sql_facility), conn, params=params)
        df_par = pd.read_sql(text(sql_parking),  conn, params=params)

    gdf_region = gpd.GeoDataFrame(geometry=[poly], crs="EPSG:4326")
    gdf_fac = gpd.GeoDataFrame(df_fac, geometry=gpd.points_from_xy(df_fac["x"], df_fac["y"]), crs="EPSG:4326")
    gdf_par = gpd.GeoDataFrame(df_par, geometry=gpd.points_from_xy(df_par["x"], df_par["y"]), crs="EPSG:4326")

    # 투영(웹 타일용) : 3857
    reg3857 = gdf_region.to_crs(3857)
    fac3857 = gdf_fac.to_crs(3857) if not gdf_fac.empty else gdf_fac
    par3857 = gdf_par.to_crs(3857) if not gdf_par.empty else gdf_par

    # 고정 팔레트(8그룹)
    group_palette = {
        "그룹1": "#1F77B4",        # 철도/버스/공항/지하철/터미널
        "그룹2(개별)": "#FF7F0E",   # 상가업소
        "그룹3(개별)": "#2CA02C",   # 의료기관
        "그룹4(개별)": "#D62728",   # 공중화장실
        "그룹5(개별)": "#9467BD",   # 주차장
        "그룹6": "#8C564B",        # 관공서/공공기관/교육기관/은행점포
        "그룹7(개별)": "#E377C2",   # 숙박시설
        "그룹8": "#17BECF",        # 문화여가시설/상영관
        "기타": "#7F7F7F"
    }

    fig, ax = plt.subplots(figsize=(8, 7))

    # (버퍼 외곽선은 숨김)  ← 원 안 보이게
    # reg3857.boundary.plot(ax=ax, color="#005BAC", linewidth=2, alpha=0.8, zorder=5)

    # 시설: 그룹별 색상
    if not fac3857.empty:
        for gnm, g in fac3857.sort_values("group_ord").groupby("group_nm"):
            color = group_palette.get(gnm, "#7F7F7F")
            g.plot(ax=ax,
                   markersize=22, marker="o",
                   edgecolor="k", linewidth=0.2,
                   color=color, alpha=0.85,
                   label=f"{gnm}", zorder=10)

    # 주차장: 그룹5 색상, P 마커
    if not par3857.empty:
        color = group_palette["그룹5(개별)"]
        par3857.plot(ax=ax,
                     markersize=34, marker="P",
                     edgecolor="k", linewidth=0.3,
                     color=color, alpha=0.9,
                     label="그룹5(개별)", zorder=11)

    # 축 범위: 버퍼 전체 기준
    xmin, ymin, xmax, ymax = reg3857.total_bounds
    pad = 80
    ax.set_xlim(xmin - pad, xmax + pad)
    ax.set_ylim(ymin - pad, ymax + pad)

    # 베이스맵은 마지막에
    add_basemap(ax, source=providers.CartoDB.Positron, crs=3857)

    ax.set_axis_off()
    if title:
        ax.set_title(title, fontsize=13, fontweight="bold", pad=6)

    # 범례 정리
    # if not fac3857.empty or not par3857.empty:
    #     leg = ax.legend(loc="lower left", fontsize=8, frameon=True, ncol=2, markerscale=1.0)
    #     for lh in leg.legend_handles:
    #         lh.set_alpha(1.0)

    os.makedirs(os.path.dirname(out_png), exist_ok=True)
    plt.savefig(out_png, dpi=300, bbox_inches="tight", pad_inches=0.1, transparent=True)
    plt.close(fig)

    print(f"✅ 시설·주차장 지도 생성 완료 → {out_png}  "
          f"(시설:{len(fac3857)} / 주차장:{len(par3857)})")


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

OUTPUT_PPT = "out/test_.pptx"
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
