# 📊 지역 모니터링 보고서 자동 생성 시스템

PostgreSQL 데이터베이스에서 데이터를 조회하여 PowerPoint 보고서를 자동으로 생성하는 시스템입니다. 여러 지역 코드에 대해 배치로 보고서를 생성할 수 있습니다.

---

## 📁 프로젝트 구조

```
auto_report/
├── run_batch_serial.py          # 🎯 메인 실행 파일 (배치 처리)
├── src/
│   ├── run_build_report_refac.py # 보고서 생성 엔진
│   └── ppt_fillers.py            # PPT 템플릿 채우기 유틸리티
├── template/
│   └── master_pretendard.pptx   # PowerPoint 템플릿 파일
├── config/
│   └── slides_tokens.yml         # 슬라이드별 SQL 쿼리 설정
├── out/                          # 생성된 보고서 출력 디렉토리
│   ├── report_*.pptx
│   └── img/                      # 생성된 지도/이미지 파일
└── requirements.txt
```

---

## 🚀 빠른 시작

### 1단계: 환경 설정

```bash
# 패키지 설치
pip install -r requirements.txt

# 환경 변수 설정 (.env 파일 생성)
# DB_URL=postgresql://user:password@host:port/database
```

### 2단계: 실행

```bash
# 배치 실행 (여러 지역 코드 순차 처리)
python run_batch_serial.py
```

생성된 보고서는 `out/` 디렉토리에 저장됩니다.

---

## 📝 사용 방법

### 배치 실행 (권장)

`run_batch_serial.py`는 여러 지역 코드를 순차적으로 처리합니다:

```python
region_info = [ 
    ("제62회 수원화성문화제 : 연계지역포함", "00000011", "20250927", "20251004"),
    ("수원시 3대 축제", "00000027", "20250927", "20251012"),
    # ... 더 많은 지역 코드
]
```

각 항목은 `(지역명, 지역코드, 시작날짜, 종료날짜)` 형식입니다.

### 단일 보고서 생성

특정 지역 코드에 대해 단일 보고서만 생성하려면:

```bash
python src/run_build_report_refac.py \
    --REGION_CD=00000011 \
    --DATE_FROM=20250927 \
    --DATE_TO=20251004
```

---

## ⚙️ 설정 파일

### `config/slides_tokens.yml`

슬라이드별로 표시할 데이터를 SQL 쿼리로 정의합니다:

<!-- ```yaml
params:
  REGION_CD: "00000011"
  DATE_FROM: "20250927"
  DATE_TO: "20251004"
  BUFFER_M: 500

slides:
  - name: "1: 표지"
    tokens:
      SL1_p_1:
        sql: |-
          SELECT event_nm
          FROM regionmonitor.tb_analysis_report_tmp
          WHERE region_cd = CAST(:REGION_CD AS VARCHAR);
  
  - name: "3 : 행사 전체요약"
    tokens:
      SL3_p_1:
        sql: |-
          SELECT to_char(COALESCE(SUM(TOT_VIPOP), 0), 'FM999,999,999,999')
          FROM regionmonitor.tb_sexdstn_visit_popltn
          WHERE STDR_YMD BETWEEN :DATE_FROM AND :DATE_TO
            AND REGION_CD = CAST(:REGION_CD AS VARCHAR);
    
    charts:
      chart_name:
        category_sql: |-
          SELECT category FROM ...
        series:
          시리즈명: |-
            SELECT value FROM ...
``` -->

**주요 설정 항목:**
- `params`: 공통 파라미터 (REGION_CD, DATE_FROM, DATE_TO 등)
- `slides`: 슬라이드별 토큰 및 차트 설정
- `tokens`: 텍스트 토큰 (PPT의 `{{TOKEN}}` 치환)
- `charts`: 차트 데이터 (카테고리, 시리즈)

---

<!-- ## 🎨 PowerPoint 템플릿 설정

### 텍스트 토큰

템플릿 파일에서 변하는 부분만 `{{TOKEN_NAME}}` 형태로 표시:

```
총 방문인구 {{SL3_p_1}}명
전년대비 {{SL3_p_2}} 증가
```

### 차트 이름 설정

1. PowerPoint에서 차트 선택 → 우클릭 → **"이름 바꾸기"**
2. 의미있는 이름 설정 (예: `SL20_chart`, `SL21_chart`)
3. **선택 창(Selection Pane)**에서 확인 가능

### 이미지 도형 이름

지도, 히트맵 등 이미지를 삽입할 도형도 이름을 설정해야 합니다:
- `SL_G1_MAP` ~ `SL_G8_MAP`: 그룹별 시설 지도
- 기타 히트맵 이미지 도형 -->

---

## 🔧 주요 기능

### 1. 텍스트 토큰 치환
- 토큰을 데이터베이스 조회 결과로 자동 치환

### 2. 차트 데이터 교체
- 차트 데이터 자동 업데이트
- 축제일 표시 등의 특수 색상 처리 지원

### 3. 이미지 생성 및 삽입
- **지도 생성**: 그룹별 시설 지도 자동 생성 (GeoPandas + Contextily)
- **히트맵 생성**: Seaborn 기반 히트맵 이미지 생성
- 이미지 자동 삽입

### 4. 트리맵 차트 업데이트
- Windows COM을 통한 트리맵 차트 데이터 업데이트
- 외국인/내국인 업종별 매출 데이터 표시

### 5. 표(Table) 데이터 채우기
- 주차장 정보 등 표 형식 데이터 자동 채우기

### 6. 자동 색상 처리
- 증감 표시에 따른 자동 색상 적용
- 조건별 색상 처리 ("충분"=초록, "부족"=빨강)

---

<!-- ## 📊 데이터베이스 스키마

시스템은 다음 PostgreSQL 스키마를 사용합니다:

- `regionmonitor.tb_intrst_region_relm`: 관심 지역 영역 정보
- `regionmonitor.tb_sexdstn_visit_popltn`: 성별 방문 인구 데이터
- `regionmonitor.tb_sexdstn_selng`: 성별 매출 데이터
- `regionmonitor.TB_MAIN_FCLTY_INFO`: 주요 시설 정보
- `regionmonitor.TB_PRKPLCE_INFO`: 주차장 정보
- `regionmonitor.TB_NATION_SELNG`: 국적별 매출 데이터
- 기타 분석 관련 테이블

--- -->

## 🛠️ 기술 스택

- **Python 3.x**
- **python-pptx**: PowerPoint 파일 처리
- **SQLAlchemy**: 데이터베이스 연결
- **pandas**: 데이터 처리
- **matplotlib/seaborn**: 차트 및 히트맵 생성
- **GeoPandas**: 지리공간 데이터 처리
- **Contextily**: 베이스맵 타일
- **PyYAML**: 설정 파일 처리
- **win32com**: Windows COM (트리맵 차트 업데이트용)

---

<!-- ## 📤 출력 파일

- **보고서**: `out/report_{REGION_CD}.pptx`
- **지도 이미지**: `out/img/{REGION_CD}/group{1-8}_map.png`
- **히트맵 이미지**: `out/img/{SLIDE_NAME}_heatmap.png`

--- -->

## ⚠️ 주의사항

1. **Windows 환경 필요**: 트리맵 차트 업데이트를 위해 `win32com` 사용 (Windows 전용)
2. **데이터베이스 연결**: `.env` 파일에 `DB_URL` 설정 필수
3. **템플릿 파일**: `template/master_pretendard.pptx` 파일이 존재해야 함
4. **도형 이름**: PPT 템플릿에서 차트/이미지 도형 이름이 설정 파일과 일치해야 함

---

<!-- ## 🔍 문제 해결

### 보고서가 생성되지 않는 경우
- 데이터베이스 연결 확인 (`.env` 파일의 `DB_URL`)
- 지역 코드(`REGION_CD`)가 데이터베이스에 존재하는지 확인
- 날짜 범위(`DATE_FROM`, `DATE_TO`)가 올바른지 확인

### 차트가 업데이트되지 않는 경우
- PPT 템플릿에서 차트 도형 이름 확인 (선택 창에서 확인)
- 설정 파일(`slides_tokens.yml`)의 차트 이름과 일치하는지 확인

### 이미지가 삽입되지 않는 경우
- 이미지 도형 이름 확인
- 이미지 파일 경로 확인 (`out/img/` 디렉토리)

--- -->
