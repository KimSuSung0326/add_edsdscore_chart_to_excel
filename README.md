# 📊 EDSD Score Excel Processor

이 프로젝트는 병실별 **EDSD Score** 데이터를 Excel 파일에서 읽어와,  
최근 30일간의 데이터를 누적 저장하고 그래프로 시각화하여 Excel에 자동 삽입하는 도구입니다.  

---

## 🚀 주요 기능

- **엑셀 데이터 읽기**
  - `edsd_score_output.xlsx` 파일에서 병실별 점수 데이터를 읽어옴
  - 두 가지 엑셀 포맷 지원  
    - 기존 형식: `B열 = 병실`, `C열 = 점수`  
    - 새로운 형식: `C열 = 병실`, `D열 = 점수`

- **데이터 누적 관리**
  - 최근 30일간의 점수를 병실 단위로 누적
  - `data_store.pkl` 파일에 직렬화하여 저장
  - 오래된 데이터는 자동으로 삭제 (`days_back` 기준)

- **그래프 생성**
  - 병실별 점수를 날짜별 꺾은선 그래프로 생성
  - 0점 데이터는 시각적으로 보이도록 `+0.1` 처리
  - 각 점수 위에 **데이터 라벨** 표시

- **Excel 자동 업데이트**
  - 생성된 그래프를 해당 병원 시트에 삽입
  - 병원 코드(`yn`, `jj`, `h`, `gj`)에 따라 시트 자동 매핑
  - 기존 그래프 이미지는 삭제 후 갱신

---

## 📂 프로젝트 구조

project/
│── excels/
│ ├── 2025_09_01/edsd_score_output.xlsx
│ ├── 2025_09_02/edsd_score_output.xlsx
│ └── ...
│
│── data_store.pkl # 누적 데이터 저장 파일
│── main.py # 메인 실행 코드
│── README.md # 설명 문서

yaml
코드 복사

---

## ⚙️ 설치 및 실행 방법

### 1) 필수 라이브러리 설치
```bash
pip install openpyxl matplotlib
2) 코드 실행
bash
코드 복사
python main.py
🏥 병원 코드 매핑
코드	병원명
yn	영남(경산)
jj	전남제일(화순)
h	효사랑(영천)
gj	구미제일(구미)

📊 실행 결과
최근 30일간의 EDSD Score 데이터가 누적 저장됨

각 병실/병동 단위 그래프가 생성되어 Excel 시트에 삽입됨

Excel 파일은 기존 데이터를 유지하면서 그래프만 최신화

🧹 데이터 관리
data_store.pkl 파일에 누적 데이터 저장

필요 시 초기화를 위해 삭제 가능:

bash
코드 복사
rm data_store.pkl
