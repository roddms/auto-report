# src/run_build_report.py
import os
import yaml
from sqlalchemy import create_engine, text
from dotenv import load_dotenv, find_dotenv
from ppt_fillers import apply_tokens_and_charts

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
table_chart_map = {}

for s in cfg["slides"]:
    for token, meta in s.get("tokens", {}).items():
        with engine.connect() as conn:
            val = conn.execute(text(meta["sql"]), cfg["params"]).scalar()
            token_values[token] = val

    for chart_name, chart_conf in s.get("charts", {}).items():
        if "category_sql" in chart_conf:
            with engine.connect() as conn:
                categories = [r[0] for r in conn.execute(text(chart_conf["category_sql"]), cfg["params"]).fetchall()]
                series = {}
                for sname, ssql in chart_conf["series"].items():
                    series[sname] = [r[0] for r in conn.execute(text(ssql), cfg["params"]).fetchall()]
                chart_data[chart_name] = (categories, series)

        if "chart_data_sql" in chart_conf:
            with engine.connect() as conn:
                rows = conn.execute(text(chart_conf["chart_data_sql"]), cfg["params"]).fetchall()
                # SQLAlchemy Row -> tuple 로 맞추기
                rows = [tuple(r) for r in rows]
            table_chart_map[chart_name] = {
                "headers": chart_conf.get("headers", []),
                "rows": rows,
            }
            continue


table_chart_map = {
    "SL19_chart_1": {
        "headers": ["구분", "업종명", "소비금액"],
        "rows": [("외국인", "업종1", 22.0), ("외국인", "업종2", 12.0), ...]
    },
    "SL19_chart_2": {
        "headers": ["구분", "업종명", "소비금액"],
        "rows": [("내국인", "업종A", 25.3), ("내국인", "업종B", 19.4), ...]
    },
}


apply_tokens_and_charts(
    prs_path="template/master.pptx",
    out_path="out/report_test_차트18.pptx",
    token_map=token_values,
    chart_map=chart_data
)

print("✅ 테스트 슬라이드 생성 완료")
