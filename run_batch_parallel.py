# run_batch_parallel.py
from multiprocessing import Pool
import subprocess
import os

# 병렬 실행할 지역 목록
region_list = [
    "00000011",
    "00000012",
    "00000010"
]

# 공통 기간
DATE_FROM = "20250927"
DATE_TO   = "20251004"

MAX_PROCESSES = 3

def run_report(region_cd):
    print(f"\n🚀 {region_cd} 보고서 생성 시작...\n")
    out_name = f"out/report_{region_cd}.pptx"

    cmd = [
        "python", "src/run_build_report_refac.py",
        f"--REGION_CD={region_cd}",
        f"--DATE_FROM={DATE_FROM}",
        f"--DATE_TO={DATE_TO}"
    ]

    # 실행 (stdout/stderr 출력 그대로 전달)
    subprocess.run(cmd, check=True)
    print(f"✅ {region_cd} 완료 → {out_name}")

if __name__ == "__main__":
    os.makedirs("out", exist_ok=True)

    with Pool(processes=MAX_PROCESSES) as pool:
        pool.map(run_report, region_list)

    print("\n🎉 모든 보고서 생성 완료!")
