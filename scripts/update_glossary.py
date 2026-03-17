import gspread
import os
import json
import pandas as pd
import sys
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Dam bao output ho tro UTF-8
if sys.stdout.encoding != 'utf-8':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def update_glossary():
    print("[INFO] Dang tai du lieu tu Google Sheets...")
    
    # --- CAU HINH LOCAL DE TEST ---
    local_service_account_path = "service-account.json" 
    local_sheet_url = "https://docs.google.com/spreadsheets/d/1LzYUaFv0KfzIy0y_ju2IxA5U_Pyu7B6Noup8CCaeZM8/edit?usp=sharing"
    # -----------------------------

    service_account_str = os.environ.get("GOOGLE_SERVICE_ACCOUNT")
    sheet_url = local_sheet_url if local_sheet_url else os.environ.get("GOOGLE_SHEET_URL")
    
    # Kiem tra xac thuc
    has_local_auth = local_service_account_path and os.path.exists(local_service_account_path)
    if not (service_account_str or has_local_auth) or not sheet_url:
        print("[ERROR] Thieu thong tin xac thuc (service-account.json hoac GOOGLE_SERVICE_ACCOUNT) hoac URL cua Google Sheet.")
        return

    try:
        if has_local_auth:
            print(f"- Su dung file xac thuc local: {local_service_account_path}")
            gc = gspread.service_account(filename=local_service_account_path)
            with open(local_service_account_path, 'r') as f:
                info = json.load(f)
                email = info.get('client_email')
                print(f"- Email Service Account: {email}")
                print(f"!!! QUAN TRONG: Vao Google Sheet -> nut 'Share' -> Them email '{email}' voi quyen Viewer.")
        elif service_account_str:
            print("- Su dung thong tin xac thuc tu bien moi truong.")
            service_account_info = json.loads(service_account_str)
            gc = gspread.service_account_from_dict(service_account_info)
            if service_account_info:
                print(f"- Email Service Account: {service_account_info.get('client_email')}")
        else:
            # Truong hop nay ly thuyet khong xay ra do da check o tren
            print("[ERROR] Khong tim thay thong tin xac thuc.")
            return
            
        sh = gc.open_by_url(sheet_url)
        print(f"[OK] Da ket noi thanh cong voi: {sh.title}")
        
        output_md = 'glossary/glossary.md'
        os.makedirs('glossary', exist_ok=True)

        with open(output_md, 'w', encoding='utf-8') as f:
            f.write("# TAI LIEU THAM KHAO DICH THUAT\n\n")

            # Ham ho tro tim sheet khong phan biet hoa thuong
            def get_ws(name):
                try:
                    return sh.worksheet(name)
                except:
                    # Neu tim truc tiep khong thay, thu tim trong danh sach sheet hoa/thuong
                    for w in sh.worksheets():
                        if w.title.lower().strip() == name.lower().strip():
                            return w
                    return None

            # --- 1. XU LY SHEET XUNG HO ---
            print("- Dang xu ly sheet 'Xung ho'...")
            ws_xh = get_ws("Xưng hô")
            if ws_xh:
                try:
                    data_xh = ws_xh.get_all_values()
                    df_xh = pd.DataFrame(data_xh[1:], columns=data_xh[0])
                    df_xh.set_index(df_xh.columns[0], inplace=True)
                    
                    f.write("## 1. QUY TAC XUNG HO\n")
                    
                    # Gom nhom dai tu khi ke chuyen (Third-person)
                    narration_pronouns = []
                    for name in df_xh.index:
                        if name in df_xh.columns:
                            val = str(df_xh.loc[name, name]).strip()
                            if val and val != '-' and val.lower() != 'nan':
                                narration_pronouns.append(f"{name}: {val}")
                    
                    if narration_pronouns:
                        f.write("### Đại từ khi kể chuyện (Ngôi thứ 3):\n")
                        f.write(", ".join(narration_pronouns) + "\n\n")

                    f.write("### Cách gọi nhau trong đối thoại (Ngôi 1 gọi Ngôi 2):\n")
                    for speaker in df_xh.index:
                        for listener in df_xh.columns:
                            if speaker == listener: continue # Đã xử lý ở trên
                            call_name = str(df_xh.loc[speaker, listener]).strip()
                            if call_name and call_name != '-' and call_name.lower() != 'nan':
                                # Xử lý format "A - B" -> "xưng A gọi B"
                                lines = call_name.split('\n')
                                processed_lines = []
                                for line in lines:
                                    line = line.strip()
                                    if not line: continue
                                    if ' - ' in line:
                                        parts = line.split(' - ')
                                        processed_lines.append(f"xưng {parts[0].strip()} gọi {parts[1].strip()}")
                                    elif '-' in line:
                                        parts = line.split('-')
                                        processed_lines.append(f"xưng {parts[0].strip()} gọi {parts[1].strip()}")
                                    else:
                                        processed_lines.append(line)
                                
                                final_call = ", ".join(processed_lines)
                                f.write(f"- {speaker} gọi {listener} là: {final_call}\n")
                    f.write("\n")
                except Exception as e: print(f"Loi sheet Xung ho: {e}")

            # --- 2. XU LY SHEET NHAN VAT ---
            print("- Dang xu ly sheet 'Nhan vat'...")
            ws_nv = get_ws("Nhân vật")
            if ws_nv:
                try:
                    data_nv = ws_nv.get_all_values()
                    if len(data_nv) > 1:
                        # Sử dụng index-based access cho bố cục mới: A(0):Tên, B(1):Giới tính, D(3):Đại từ, I(8):Vai trò, J(9):Biệt danh
                        f.write("## 2. THONG TIN NHAN VAT\n")
                        for row in data_nv[1:]:
                            if len(row) < 1: continue
                            ten = str(row[0]).strip()
                            gioi_tinh = str(row[1]).strip() if len(row) > 1 else ""
                            dai_tu = str(row[3]).strip() if len(row) > 3 else ""
                            vai_tro = str(row[8]).strip() if len(row) > 8 else ""
                            biet_danh = str(row[9]).strip() if len(row) > 9 else ""
                            
                            if ten and ten.lower() != 'nan' and ten != '':
                                line = f"- {ten}"
                                details = []
                                if gioi_tinh and gioi_tinh.lower() != 'nan': details.append(gioi_tinh)
                                if dai_tu and dai_tu.lower() != 'nan': details.append(f"đại từ: {dai_tu}")
                                if vai_tro and vai_tro.lower() != 'nan': details.append(f"vai trò: {vai_tro}")
                                if biet_danh and biet_danh.lower() != 'nan': details.append(f"biệt danh: {biet_danh}")
                                
                                if details:
                                    line += " (" + ", ".join(details) + ")"
                                f.write(line + "\n")
                        f.write("\n")
                except Exception as e: print(f"Loi sheet Nhan vat: {e}")

            # --- 3. XU LY SHEET TU VUNG ---
            print("- Dang xu ly sheet 'Từ vựng'...")
            ws_tv = get_ws("Từ vựng")
            if ws_tv:
                try:
                    data_tv = ws_tv.get_all_values()
                    if len(data_tv) > 1:
                        f.write("## 3. THUAT NGU VA TEN RIENG\n")
                        for row in data_tv[1:]:
                            # Bố cục mới: A: Chap, B: Hàn, C: Việt, D: Anh
                            chap = str(row[0]).strip() if len(row) > 0 else ""
                            han = str(row[1]).strip() if len(row) > 1 else ""
                            viet = str(row[2]).strip() if len(row) > 2 else ""
                            anh = str(row[3]).strip() if len(row) > 3 else ""
                            
                            if han or viet or anh:
                                line = f"- {han}"
                                if anh: line += f" | {anh}"
                                if viet: line += f" -> {viet}"
                                if chap: line += f" (Chap {chap})"
                                f.write(line + "\n")
                        f.write("\n")
                except Exception as e: print(f"Loi sheet Tu vung: {e}")

        print(f"[OK] Da cap nhat thanh cong {output_md} tu Google Sheets!")

    except Exception as e:
        import traceback
        print(f"[ERROR] LOI khi ket noi Google Sheets: {e}")
        print("Chi tiet loi:")
        print(traceback.format_exc())

if __name__ == "__main__":
    update_glossary()
