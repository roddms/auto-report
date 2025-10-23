# PPT 자동 업데이트 시스템 사용 가이드

## 📁 **프로젝트 구조**

```
auto_report/
├── template/
│   └── master.pptx              # 마스터 템플릿 ({{TOKEN}} 패턴 포함)
├── config/
│   └── slides.yml               # 슬라이드별 SQL 매핑 설정
├── secrets/
│   └── .env                     # DB 접속 정보
├── src/
│   ├── ppt_fillers.py           # 토큰/차트 교체 엔진
│   └── run_build_report.py      # 메인 실행 파일
├── out/
│   └── report_filled.pptx       # 결과물
└── requirements.txt
```

---

## 🚀 **사용 방법**

### **1단계: 환경 설정**
```bash
# 패키지 설치
pip install -r requirements.txt

# 환경 변수 설정
cp secrets/env_example.txt secrets/.env
# secrets/.env 파일에서 실제 DB 연결 정보 입력
```

### **2단계: 템플릿 생성**
```bash
# GPT 방식 템플릿 생성
python create_gpt_template.py
```

### **3단계: 설정 파일 수정**
`config/slides.yml`에서 실제 DB 테이블과 쿼리로 수정:

```yaml
params:
  EVENT_ID: 2025_Seokjangni
  DATE_FROM: '2025-10-01'
  DATE_TO: '2025-10-31'

slides:
  - name: overview
    tokens:
      TOT_VIS:
        sql: |
          SELECT SUM(visitors) FROM your_table WHERE event_id = :EVENT_ID
        fmt: int_comma
    charts:
      chart_daily_visitors:
        category_sql: |
          SELECT hour_label FROM your_table WHERE event_id = :EVENT_ID
        series:
          방문인구수: |
            SELECT visitors FROM your_table WHERE event_id = :EVENT_ID
```

### **4단계: 실행**
```bash
# 기본 실행
python src/run_build_report.py

# 파라미터와 함께 실행
python src/run_build_report.py --EVENT_ID 2025_Seokjangni --DATE_FROM 2025-10-01
```

---

## 🎨 **PPT 템플릿 설정 방법**

### **텍스트 토큰 설정**
1. **텍스트 박스에 `{{TOKEN_NAME}}` 패턴 입력**
   ```
   총 방문인구 {{TOT_VIS}}명
   전년대비 {{YOY_VIS_PCT}} 증가
   총 매출 {{TOT_SALES_WON}}
   ```

2. **고정 텍스트는 그대로 두기**
   - "총 방문인구", "전년대비", "총 매출" 등은 변경하지 않음
   - 변하는 숫자/라벨만 `{{}}`로 감싸기

### **차트 이름 설정**
1. **차트 선택** → **우클릭** → **"이름 바꾸기"**
2. **의미있는 이름 설정**:
   ```
   chart_daily_visitors
   chart_gender_spend_by_cat
   chart_top_products
   chart_trend_line
   ```

3. **선택 창(Selection Pane)에서 확인**:
   - PowerPoint에서 "홈" → "선택" → "선택 창" 클릭
   - 차트 이름이 설정된 것 확인

---

## 🔧 **포맷터 종류**

| 포맷터 | 설명 | 예시 |
|--------|------|------|
| `int_comma` | 천 단위 콤마 | 1,234,567 |
| `signed_percent_1` | 부호 있는 퍼센트 | +15.3% |
| `won_or_eok` | 원/억원 자동 변환 | 1,234,567원 / 1.2억원 |
| `date_kr` | 한국식 날짜 | 10월 17일 |
| `default` | 기본 문자열 | 그대로 출력 |

---

## 📊 **실행 흐름**

```
   ┌───────────────┐
   │ master.pptx   │  ← {{TOKEN}} 패턴과 차트 이름 포함
   └──────┬────────┘
          │
          ▼
   ┌───────────────┐
   │ slides.yml    │  ← 슬라이드별 SQL 설정
   └──────┬────────┘
          │
          ▼
   ┌───────────────┐
   │ run_build...  │  ← 메인 스크립트
   ├───────────────┤
   │ ① DB 연결     │
   │ ② SQL 실행    │
   │ ③ 포맷 변환   │
   │ ④ {{TOKEN}} 치환 │
   │ ⑤ 차트 데이터 교체 │
   └──────┬────────┘
          │
          ▼
   ┌───────────────┐
   │ report_filled │  ← 최종 PPT 자동 보고서
   └────────────────┘
```

---

## 💡 **핵심 포인트**

### **텍스트 처리**
- **전체 교체 X**: 텍스트 박스 전체를 바꾸지 않음
- **토큰 치환 O**: `{{TOKEN}}` 패턴만 값으로 교체
- **디자인 보존**: 폰트, 색상, 위치 등 모든 서식 유지

### **차트 처리**
- **이름 기반**: 선택 창에서 확인 가능한 이름으로 차트 식별
- **데이터만 교체**: 차트의 색상, 스타일, 축 설정 등 모두 보존
- **다중 시리즈**: 여러 데이터 시리즈를 하나의 차트에 표시 가능

### **설정 관리**
- **YAML 기반**: 직관적이고 읽기 쉬운 설정 파일
- **슬라이드별 독립**: 각 슬라이드마다 다른 테이블/쿼리 사용
- **파라미터 지원**: 동적 값으로 쿼리 실행

---

## 🎉 **결과**

- **클릭 한 번**으로 DBeaver DB → 실시간 쿼리 → PPT 자동 생성
- **디자인 완벽 보존**: 텍스트/차트 서식 모두 그대로 유지
- **유연한 확장**: YAML만 수정하면 새 슬라이드 자동 반영
- **한국식 표기**: 숫자, 날짜, 통화 등 한국식 포맷 자동 적용


## 🎯 **GPT 방식의 핵심 장점**

### ✅ **1. 텍스트 토큰 방식 (`{{TOKEN_NAME}}`)**
- **디자인 완벽 보존**: 고정 텍스트는 그대로, 변하는 부분만 `{{}}`로 감싸기
- **문장 내 삽입**: "총 방문인구 {{TOT_VIS}}명" 형태로 자연스러운 문장 구성
- **유연한 포맷팅**: `int_comma`, `won_or_eok` 등 한국식 표기 지원

### ✅ **2. 차트 이름 지정 방식**
- **선택 창(Selection Pane)**에서 차트 이름 확인 가능
- **서식 완벽 보존**: 색상, 폰트, 축 설정 등 디자인 그대로 유지
- **데이터만 교체**: `replace_data()`로 깔끔하게 처리

### ✅ **3. YAML 기반 설정**
- **슬라이드별 독립적 설정**: 각 슬라이드마다 다른 테이블/쿼리 사용
- **파라미터 지원**: `EVENT_ID`, `DATE_FROM` 등 동적 값 처리
- **포맷터 지원**: 한국식 숫자 표기 자동 변환

---