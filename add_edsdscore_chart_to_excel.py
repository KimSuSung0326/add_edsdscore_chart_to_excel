import os
import pickle
from datetime import datetime, timedelta
import openpyxl
from openpyxl.drawing.image import Image as ExcelImage
import re
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

# --- 엑셀 데이터 읽기 ---
def open_edsd_score_excel(excel_path):
    if os.path.exists(excel_path):
        return openpyxl.load_workbook(excel_path)
    return None

def read_edsd_scores(workbook):
    if workbook is None:
        return {}
    sheet = workbook.active
    scores = {}
    for row in sheet.iter_rows(min_row=4, min_col=2, max_col=3, values_only=True):
        room, score = row
        if room is None or score is None:
            continue
        match = re.match(r'(\d+_\d+)(?:_(\w+))?', room)
        if match:
            room_num = match.group(1).replace('_', '-')
            hospital_code = match.group(2) if match.group(2) else "yn"
            ward = room_num.split('-')[0]
            scores.setdefault(hospital_code, {}).setdefault(ward, {})[room_num] = score
    return scores

# --- 30일간의 엑셀 파일 찾기 ---
def find_excel_files_in_date_range(base_dir, days_back=30):
    excel_files = []
    today = datetime.now().date()
    for i in range(days_back, -1, -1):
        target_date = today - timedelta(days=i)
        folder_name = target_date.strftime('%Y_%m_%d')
        excel_path = os.path.join(base_dir, 'excels', folder_name, 'edsd_score_output.xlsx')
        if os.path.exists(excel_path):
            excel_files.append((target_date, excel_path))
            print(f"발견된 파일: {folder_name}/edsd_score_output.xlsx")
    return excel_files

# --- 누적 데이터 저장/불러오기 ---
def load_data_store(file_path="data_store.pkl"):
    if os.path.exists(file_path):
        with open(file_path, "rb") as f:
            data = pickle.load(f)
            if 'check_delete_date_flag' not in data:
                data['check_delete_date_flag'] = 0
            return data
    return {'check_delete_date_flag': 0}

def save_data_store(data_store, file_path="data_store.pkl"):
    with open(file_path, "wb") as f:
        pickle.dump(data_store, f)

def delete_data_store(file_path="data_store.pkl"):
    if os.path.exists(file_path):
        os.remove(file_path)
        print(f"data_store.pkl 파일이 삭제되었습니다.")

# --- 누적 데이터 오래된 것 삭제 ---
def prune_old_data(accumulated_data, days_back=30):
    today = datetime.now().date()
    cutoff = today - timedelta(days=days_back)
    for hospital_code, wards in accumulated_data.items():
        for ward, rooms in wards.items():
            for room_num, xy_data in rooms.items():
                accumulated_data[hospital_code][ward][room_num] = [
                    (d, s) for d, s in xy_data if d >= cutoff
                ]
    print(f"{days_back}일 이전 데이터 삭제 완료")
    return accumulated_data

# --- 그래프 생성 (오늘 기준 days_back만큼만) ---
def create_plots_for_date(accumulated_data, current_date, days_back=2, save_dir="plotImg"):
    os.makedirs(save_dir, exist_ok=True)
    image_files = {}
    colors = ['r', 'b', 'g', 'orange', 'purple', 'pink']

    today = current_date
    start_date = today - timedelta(days=days_back)

    for hospital_code, wards in accumulated_data.items():
        image_files.setdefault(hospital_code, {})
        for ward, rooms in wards.items():
            if not isinstance(rooms, dict):
                continue

            ward_shown = set()
            fig, ax = plt.subplots(figsize=(8, 5))
            ax.set_title(f"{hospital_code} - Room{ward} (~ {current_date.strftime('%Y-%m-%d')})")
            ax.set_xlabel("")
            ax.set_ylabel("EDSD SCORE")

            all_y_values = []

            for idx, (room_num, xy_data) in enumerate(rooms.items()):
                # 오늘 기준 days_back만큼만 데이터 필터링
                xy_data_filtered = [(d, s) for d, s in xy_data if start_date <= d <= today]
                if not xy_data_filtered:
                    continue

                x = [d for d, s in xy_data_filtered]
                y = [s + 0.1 if s == 0 else s for d, s in xy_data_filtered]
                all_y_values.extend(y)

                ax.plot(x, y, marker='o', label=room_num, color=colors[idx % len(colors)])

                for xi, yi, yi_orig in zip(x, y, [s for d, s in xy_data_filtered]):
                    key = (xi, yi_orig, ward)
                    if key not in ward_shown:
                        ax.text(xi, yi + 0.15, f"{yi_orig}", fontsize=7,
                                ha='center', va='bottom', rotation=0, color='black')
                        ward_shown.add(key)

            if all_y_values:
                ax.set_ylim(0, max(all_y_values) + 3)

            all_dates = sorted(set(d for room_data in rooms.values() for d, s in room_data if start_date <= d <= today))
            ax.set_xticks(all_dates)
            if len(all_dates) > 10:
                ax.xaxis.set_major_locator(mdates.DayLocator(interval=max(1, len(all_dates)//10)))
            fig.autofmt_xdate(rotation=45)

            ax.legend(loc='center left', bbox_to_anchor=(1, 0.5))

            img_path = os.path.join(save_dir, f"{hospital_code}_{ward}_{current_date.strftime('%Y%m%d')}.png")
            fig.savefig(img_path, bbox_inches='tight')
            plt.close(fig)

            image_files[hospital_code][ward] = img_path

    return image_files

# --- 엑셀에 그래프 삽입 ---
def save_to_excel(image_files, excel_path):
    """기존 엑셀 파일에 그래프 추가"""
    if os.path.exists(excel_path):
        wb = openpyxl.load_workbook(excel_path)
    else:
        print(f"엑셀 파일이 존재하지 않음: {excel_path}")
        return

    # hospital_code -> 시트 이름 매핑
    sheet_name_map = {
        'yn': "영남(경산)",
        'jj': "전남제일(화순)",
        'h': "효사랑(영천)",
        'gj': "구미제일(구미)"
    }

    for hospital_code, wards in image_files.items():
        # hospital_code에 맞는 시트 이름
        sheet_name = sheet_name_map.get(hospital_code, hospital_code)

        # 기존 시트가 있으면 그대로 사용, 없으면 새로 생성
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            # 기존 이미지들 삭제 (그래프 업데이트를 위해)
            ws._images.clear()
        else:
            ws = wb.create_sheet(title=sheet_name)

        # 이미지 삽입
        row_pos = 1  # 맨 위부터 삽입
        for ward, img_path in wards.items():
            img = ExcelImage(img_path)
            ws.add_image(img, f"A{row_pos}")
            row_pos += 24  # 이미지 간격

    wb.save(excel_path)
    print(f"엑셀 업데이트 완료: {excel_path}")


# --- 메인 실행 ---
if __name__ == '__main__':
    base_dir = os.path.dirname(os.path.abspath(__file__))
    excel_files = find_excel_files_in_date_range(base_dir, days_back=30)

    if not excel_files:
        print("30일 내에 존재하는 엑셀 파일이 없습니다.")
        exit()

    data_store = load_data_store()
    accumulated_data = {}

    print(f"\n각 날짜별 그래프 생성 시작... (총 {len(excel_files)}개 파일)")

    for i, (current_date, excel_path) in enumerate(excel_files):
        print(f"\n[{i+1}/{len(excel_files)}] 처리중: {current_date} - {os.path.basename(os.path.dirname(excel_path))}")

        wb = open_edsd_score_excel(excel_path)
        current_scores = read_edsd_scores(wb)
        if wb:
            wb.close()

        for hospital_code, wards in current_scores.items():
            accumulated_data.setdefault(hospital_code, {})
            for ward, rooms in wards.items():
                accumulated_data[hospital_code].setdefault(ward, {})
                for room_num, score in rooms.items():
                    accumulated_data[hospital_code][ward].setdefault(room_num, [])
                    accumulated_data[hospital_code][ward][room_num].append((current_date, score))

        accumulated_data = prune_old_data(accumulated_data, days_back=30)

        # 오늘 기준 days_back만큼만 그래프 생성
        save_dir = f"plotImg_{current_date.strftime('%Y%m%d')}"
        image_files = create_plots_for_date(accumulated_data, current_date, days_back=30, save_dir=save_dir)

        if image_files:
            save_to_excel(image_files, excel_path)

        import shutil
        if os.path.exists(save_dir):
            shutil.rmtree(save_dir)

    data_store['check_delete_date_flag'] += 1
    #print(f"\n실행 횟수: {data_store['check_delete_date_flag']}/30")

    for hospital_code, wards in accumulated_data.items():
        data_store.setdefault(hospital_code, {})
        for ward, rooms in wards.items():
            data_store[hospital_code].setdefault(ward, {})
            for room_num, xy_data in rooms.items():
                data_store[hospital_code][ward][room_num] = xy_data

    save_data_store(data_store)
    print("\n모든 날짜별 그래프 생성 및 저장 완료!")
