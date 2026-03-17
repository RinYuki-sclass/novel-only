import gspread
import os
import json
import pandas as pd
import sys

# Dam bao output ho tro UTF-8
if sys.stdout.encoding != 'utf-8':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def update_glossary():
    print("[INFO] Dang tai du lieu tu Google Sheets...")
    
    # --- CAU HINH LOCAL DE TEST ---
    local_service_account_path = "service-account.json" 
    local_sheet_url = "https://docs.google.com/spreadsheets/d/1Rm6BLnW6yj019GMLHxxsQGDYCHQdz8z-cALrmdCfdro/edit?usp=sharing"
    # -----------------------------

    service_account_str = os.environ.get("GOOGLE_SERVICE_ACCOUNT")
    sheet_url = local_sheet_url if local_sheet_url else os.environ.get("GOOGLE_SHEET_URL")
    
    if not (service_account_str or local_service_account_path) or not sheet_url:
        print("[ERROR] Thieu thong tin xac thuc hoac URL cua Google Sheet.")
        return

    try:
        if local_service_account_path and os.path.exists(local_service_account_path):
            print(f"- Su dung file xac thuc local: {local_service_account_path}")
            gc = gspread.service_account(filename=local_service_account_path)
            with open(local_service_account_path, 'r') as f:
                info = json.load(f)
                email = info.get('client_email')
                print(f"- Email Service Account: {email}")
                print(f"!!! QUAN TRONG: Vao Google Sheet -> nut 'Share' -> Them email '{email}' voi quyen Viewer.")
        else:
            service_account_info = json.loads(service_account_str)
            gc = gspread.service_account_from_dict(service_account_info)
            if service_account_info:
                print(f"- Email Service Account: {service_account_info.get('client_email')}")
            
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
                            val = df_xh.loc[name, name]
                            if val and str(val).strip() and val != '-':
                                narration_pronouns.append(f"{name}: {val}")
                    
                    if narration_pronouns:
                        f.write("### Dai tu khi ke chuyen (Ngoi thu 3):\n")
                        f.write(", ".join(narration_pronouns) + "\n\n")

                    f.write("### Cach goi nhau trong doi thoai (Ngoi 1 goi Ngoi 2):\n")
                    for speaker in df_xh.index:
                        for listener in df_xh.columns:
                            if speaker == listener: continue # Da xu ly o tren
                            call_name = df_xh.loc[speaker, listener]
                            if call_name and str(call_name).strip() and call_name != '-':
                                f.write(f"- {speaker} goi {listener} la: \"{call_name}\"\n")
                    f.write("\n")
                except Exception as e: print(f"Loi sheet Xung ho: {e}")

            # --- 2. XU LY SHEET NHAN VAT ---
            print("- Dang xu ly sheet 'Nhan vat'...")
            ws_nv = get_ws("Nhân vật")
            if ws_nv:
                try:
                    data_nv = ws_nv.get_all_values()
                    if len(data_nv) > 1:
                        # Tao DataFrame tu gia tri thô, bo qua loi header trung
                        df_nv = pd.DataFrame(data_nv[1:], columns=data_nv[0])
                        f.write("## 2. THONG TIN NHAN VAT\n")
                        for _, row in df_nv.iterrows():
                            ten = str(row.get('Tên', '')).strip()
                            tuoi = str(row.get('Tuổi', '')).strip()
                            ghi_chu = str(row.get('Ghi chú', '')).strip()
                            sk_raw = str(row.get('Skill (raw)', '')).strip()
                            sk_eng = str(row.get('Skill (eng)', '')).strip()
                            sk_vn = str(row.get('Skill (vn)', '')).strip()
                            
                            if ten and ten != 'nan' and ten != '':
                                line = f"- {ten}"
                                if tuoi and tuoi != 'nan': line += f" ({tuoi} tuổi)"
                                if ghi_chu and ghi_chu != 'nan': line += f": {ghi_chu}"
                                
                                # Them thong tin skill neu co
                                skills = []
                                if sk_raw and sk_raw != 'nan': skills.append(sk_raw)
                                if sk_eng and sk_eng != 'nan': skills.append(sk_eng)
                                if sk_vn and sk_vn != 'nan': skills.append(f"-> {sk_vn}")
                                
                                if skills:
                                    line += " [Skill: " + " / ".join(skills) + "]"
                                
                                f.write(line + "\n")
                        f.write("\n")
                except Exception as e: print(f"Loi sheet Nhan vat: {e}")

            # --- 3. XU LY SHEET THUAT NGU ---
            print("- Dang xu ly sheet 'Thuat ngu'...")
            ws_tn = get_ws("Thuật ngữ chi tiết")
            if ws_tn:
                try:
                    data_tn = ws_tn.get_all_values()
                    if len(data_tn) > 1:
                        # Xu ly tuong tu cho Thuat ngu
                        df_tn = pd.DataFrame(data_tn[1:], columns=data_tn[0])
                        f.write("## 3. THUAT NGU VA TEN RIENG\n")
                        for _, row in df_tn.iterrows():
                            han = str(row.get('Tiếng hàn', '')).strip()
                            anh = str(row.get('Tiếng anh', '')).strip()
                            dich = str(row.get('Dịch', '')).strip()
                            if han or anh or dich:
                                f.write(f"- {han} | {anh} -> {dich}\n")
                        f.write("\n")
                except Exception as e: print(f"Loi sheet Thuat ngu: {e}")

        print(f"[OK] Da cap nhat thanh cong {output_md} tu Google Sheets!")

    except Exception as e:
        import traceback
        print(f"[ERROR] LOI khi ket noi Google Sheets: {e}")
        print("Chi tiet loi:")
        print(traceback.format_exc())

if __name__ == "__main__":
    update_glossary()
