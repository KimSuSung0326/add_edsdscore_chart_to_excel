# EDSD Score Graph Generator

병원 병실별 EDSD 점수 데이터를 시계열 그래프로 시각화하여 Excel 파일에 자동으로 삽입하는 Python 도구입니다.

## 📋 기능

- **자동 데이터 수집**: 지정된 기간(기본 30일) 동안의 EDSD 점수 Excel 파일들을 자동으로 탐지
- **시계열 그래프 생성**: 병원별, 병동별, 병실별 EDSD 점수 변화를 시각화
- **Excel 통합**: 생성된 그래프를 해당 날짜의 Excel 파일에 자동 삽입
- **데이터 기간 조절**: 그래프에 표시할 데이터 기간을 자유롭게 설정 가능
- **누적 데이터 관리**: 이전 실행 데이터를 저장하여 연속성 있는 그래프 생성


## 📁 파일 구조

```
project/
├── main.py                    # 메인 실행 파일
├── excels/                    # Excel 파일들이 저장된 디렉토리
│   ├── 2024_01_01/
│   │   └── edsd_score_output.xlsx
│   ├── 2024_01_02/
│   │   └── edsd_score_output.xlsx
│   └── ...
├── data_store.pkl             # 누적 데이터 저장 파일 (자동 생성)
└── plotImg_YYYYMMDD/          # 임시 이미지 파일들 (자동 삭제)
```

## 🛠 설치 및 요구사항

### Python 버전
- Python 3.7 이상

### 필수 라이브러리
```bash
pip install openpyxl matplotlib
```
## 🚀 사용법

### 1. 기본 실행
```bash
python add_edsdscore_chart_to_excel.py
```

### 2. 설정 변경
코드 내의 설정값을 수정하여 동작을 조정할 수 있습니다:

```python
# main.py 하단 메인 실행 부분
excel_files = find_excel_files_in_date_range(base_dir, days_back=30)  # 날짜 파일에서 오늘 날짜 기준으로 탐색 할 일수(default =30)
image_files = create_plots_for_date(accumulated_data, current_date, days_back=30, save_dir=save_dir)  # 그래프에 표시할 일수 (default =30)
```

## ⚙️ 주요 설정값

| 설정 | 기본값 | 설명 |
|------|--------|------|
| `days_back` (탐색) | 30 | Excel 파일을 찾을 일수 범위 |
| `days_back` (그래프) | 30 | 그래프에 표시할 데이터 기간 |
| `prune_old_data` | 30 | 메모리에서 제거할 오래된 데이터 기준일 |

## 📊 출력 결과

### 그래프 특징
- **제목**: `{병원코드} - Room{병동번호} (~ YYYY-MM-DD)`
- **X축**: 날짜 (MM-DD 형식)
- **Y축**: EDSD SCORE
- **데이터 포인트**: 각 점 위에 실제 점수 표시
- **범례**: 병실별 색상 구분

### Excel 시트
- 각 병원별로 개별 시트 생성
- 병동별 그래프가 세로로 배열
- 기존 그래프는 자동으로 업데이트

## 📝 데이터 형식

### Excel 파일 요구사항
- 파일명: `edsd_score_output.xlsx`
- 위치: `excels/YYYY_MM_DD/` 폴더 내
- 데이터 시작: 4행부터
- 컬럼 구조:
  - B열: 병실명 (예: `101_01_yn`, `201_02_jj`)
  - C열: EDSD 점수

### 병실명 규칙
```
{병동}_{병실}_{병원코드}
```
- 예시: `101_01_yn` → 101병동 01호실, 영남병원
- 병원코드 생략시 기본값 `yn` 적용

## 🗂 파일 설명

### 주요 함수

| 함수명 | 기능 |
|--------|------|
| `find_excel_files_in_date_range()` | 지정 기간의 Excel 파일 탐색 |
| `read_edsd_scores()` | Excel에서 EDSD 점수 데이터 추출 |
| `create_plots_for_date()` | 시계열 그래프 생성 |
| `save_to_excel()` | Excel 파일에 그래프 삽입 |
| `prune_old_data()` | 오래된 데이터 정리 |

### 데이터 파일
- **`data_store.pkl`**: 누적 데이터 및 실행 상태 저장
  - 자동으로 생성/관리됨
  - 프로그램 재실행시 기존 데이터 유지

## 🎨 그래프 커스터마이징

### 색상 변경
```python
colors = ['r', 'b', 'g', 'orange', 'purple', 'pink']  # create_plots_for_date() 함수 내
```

### 그래프 크기 조정
```python
fig, ax = plt.subplots(figsize=(8, 5))  # 너비 8, 높이 5 인치
```

### 폰트 크기 변경
```python
ax.text(xi, yi + 0.15, f"{yi_orig}", fontsize=7, ...)  # 데이터 레이블 크기
```

## ⚠️ 주의사항

1. **파일 경로**: Excel 파일들이 정확한 폴더 구조(`excels/YYYY_MM_DD/`)에 위치해야 함
2. **파일 권한**: Excel 파일이 다른 프로그램에서 열려있으면 오류 발생 가능
3. **메모리 사용**: 대량의 데이터 처리시 메모리 사용량 증가
4. **날짜 형식**: 폴더명은 반드시 `YYYY_MM_DD` 형식 사용

## 📁 파일 구조

```
project/
├── main.py                    # 메인 실행 파일
├── excels/                    # Excel 파일들이 저장된 디렉토리
│   ├── 2024_01_01/
│   │   └── edsd_score_output.xlsx
│   ├── 2024_01_02/
│   │   └── edsd_score_output.xlsx
│   └── ...
├── data_store.pkl             # 누적 데이터 저장 파일 (자동 생성)
└── plotImg_YYYYMMDD/          # 임시 이미지 파일들 (자동 삭제)
```

## 🛠 설치 및 요구사항

### Python 버전
- Python 3.7 이상

### 필수 라이브러리
```bash
pip install openpyxl matplotlib
```

또는 requirements.txt 파일이 있다면:
```bash
pip install -r requirements.txt
```

## 🔄 실행 흐름

1. **파일 탐색**: 지정된 기간 내 Excel 파일들을 날짜순으로 탐색
2. **데이터 읽기**: 각 Excel 파일에서 EDSD 점수 데이터 추출
3. **누적 저장**: 기존 데이터와 병합하여 누적 데이터베이스 구성
4. **그래프 생성**: 지정된 기간의 데이터로 시계열 그래프 생성
5. **Excel 삽입**: 생성된 그래프를 해당 날짜의 Excel 파일에 삽입
6. **정리**: 임시 파일들 삭제 및 데이터 저장

