# run_batch_serial.py
import subprocess
import time, sys

# =============================
# 실행할 지역 코드 및 기간 목록
# =============================
region_info = [
    ("단월동", "00000097", "20251022", "20251026"),
    ("대포동", "00000098", "20251022", "20251026"),
    ("호법면", "00000099", "20251022", "20251026"),
    ("대월면", "00000100", "20251022", "20251026"),
    ("모가면", "00000101", "20251022", "20251026"),
    ("설성면", "00000102", "20251022", "20251026"),
    ("이천농업테마공원", "00000103", "20251022", "20251026"),
    ("이천시 농업_민주화 공원 통합", "00000104", "20251022", "20251026"),
    ("민주화운동 기념공원", "00000105", "20251022", "20251026"),
    ("이천시", "00000096", "20251022", "20251026")
]

# =============================
# 실행 루프 (직렬)
# =============================
for name, region_cd, date_from, date_to in region_info:
    print(f"\n🚀 [{region_cd}] {name} 보고서 생성 시작 ({date_from} ~ {date_to})")

    cmd = [
        sys.executable,
        "src/run_build_report_refac_Icheon.py",
        f"--REGION_CD={region_cd}",
        f"--DATE_FROM={date_from}",
        f"--DATE_TO={date_to}"
    ]

    start = time.time()
    try:
        subprocess.run(cmd, check=True)
        elapsed = time.time() - start
        print(f"✅ [{region_cd}] 완료 ({elapsed:.1f}s)\n")
    except subprocess.CalledProcessError as e:
        print(f"❌ [{region_cd}] 실패 (exit {e.returncode})\n")
        continue

print("\n🎉 모든 지역 보고서 생성 완료 ✅")