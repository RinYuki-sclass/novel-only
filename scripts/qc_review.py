import os
import time
import sys
import json
from dotenv import load_dotenv
from google import genai
from google.genai import types

# Load API Key tu file .env
load_dotenv()
api_key = os.environ.get("GEMINI_API_KEY")

if not api_key:
    print("[ERROR] Thieu GEMINI_API_KEY. Vui long tao file .env voi noi dung: GEMINI_API_KEY=your_key")
    sys.exit(1)

client = genai.Client(api_key=api_key)

def generate_with_retry(model, contents, system_instruction="", retries=5):
    config = types.GenerateContentConfig(
        system_instruction=system_instruction,
        temperature=0.3,
    )
    for i in range(retries):
        try:
            response = client.models.generate_content(model=model, contents=contents, config=config)
            if response and response.text:
                return response.text
            return ""
        except Exception as e:
            if "429" in str(e):
                print(f"⚠️ Cham han muc (429). Dang cho 60s de reset token... (Lan {i+1}/{retries})")
                for remaining in range(60, 0, -1):
                    sys.stdout.write(f"\rTiep tuc sau {remaining}s...")
                    sys.stdout.flush()
                    time.sleep(1)
                print("\nDang thu lai...")
                continue
            print(f"❌ Loi API: {e}")
            time.sleep(5)
    return ""

def split_text(text, lines_per_chunk=50):
    lines = text.split('\n')
    chunks = []
    for i in range(0, len(lines), lines_per_chunk):
        chunks.append('\n'.join(lines[i:i + lines_per_chunk]))
    return chunks

def run_qc():
    print("🔍 DANG KHOI CHAY HE THONG QC REVIEW...")
    
    # Kiểm tra các file đầu vào trong thư mục qc
    vi_path = 'input/qc/vi_to_qc.txt'
    kor_path = 'input/qc/kor.txt'
    eng_path = 'input/qc/eng.txt'
    glossary_path = 'glossary/glossary.md'
    notes_path = 'glossary/personal_notes.md'
    
    if not os.path.exists(vi_path) or not os.path.exists(kor_path):
        print(f"[ERROR] Thieu file dau vao trong 'input/qc/'")
        print("Vui long copy ban dich can QC vao 'input/qc/vi_to_qc.txt' va ban goc vao 'input/qc/kor.txt'")
        return

    with open(vi_path, 'r', encoding='utf-8') as f: vi_text = f.read()
    with open(kor_path, 'r', encoding='utf-8') as f: kor_text = f.read()
    
    eng_text = ""
    if os.path.exists(eng_path):
        with open(eng_path, 'r', encoding='utf-8') as f: eng_text = f.read()
    
    glossary_text = ""
    if os.path.exists(glossary_path):
        with open(glossary_path, 'r', encoding='utf-8') as f: glossary_text = f.read()

    notes_text = ""
    if os.path.exists(notes_path):
        with open(notes_path, 'r', encoding='utf-8') as f: notes_text = f.read()
    
    vi_lines = vi_text.split('\n')
    kor_lines = kor_text.split('\n')
    eng_lines = eng_text.split('\n') if eng_text else []
    
    lines_per_chunk = 50
    num_chunks = (max(len(vi_lines), len(kor_lines)) + lines_per_chunk - 1) // lines_per_chunk
    
    full_report = [f"# BAO CAO QC REVIEW - {time.strftime('%d/%m/%Y %H:%M:%S')}\n"]
    new_terms_suggestions = []
    
    for i in range(num_chunks):
        start_idx = i * lines_per_chunk
        end_idx = start_idx + lines_per_chunk
        
        print(f"[{i+1}/{num_chunks}] Dang kiem tra tu dong {start_idx + 1} den {min(end_idx, len(vi_lines))}...")
        
        # Chuan bi ban dich Viet co kem so dong de AI doc
        vi_chunk_with_nums = ""
        for idx, line in enumerate(vi_lines[start_idx:end_idx], start=start_idx + 1):
            vi_chunk_with_nums += f"{idx}: {line}\n"
            
        kor_chunk = "\n".join(kor_lines[start_idx:end_idx])
        eng_reference = ""
        if eng_lines:
            eng_chunk = "\n".join(eng_lines[start_idx:end_idx])
            eng_reference = f"==== ENGLISH REFERENCE (REVISED VERSION) ====\n{eng_chunk}\n\n"

        prompt_qc = (
            "==== ROLE ====\n"
            "You are a professional Quality Control editor for Vietnamese novels. Audit the 'VIETNAMESE TRANSLATION' block.\n\n"
            "==== REFERENCE GLOSSARY (DO NOT AUDIT THIS) ====\n"
            f"{glossary_text}\n\n"
            "==== PERSONAL NOTES (HIGH PRIORITY) ====\n"
            f"{notes_text}\n\n"
            "==== TASK ====\n"
            "1. Compare the 'VIETNAMESE TRANSLATION' (with line numbers) with the 'KOREAN SOURCE' and 'ENGLISH REFERENCE'.\n"
            "2. IMPORTANT: The English version is the 'REVISED' master copy. Prioritize it over the old Korean source.\n"
            "3. REPORT BY LINE NUMBER: Use the exact line numbers provided in the 'VIETNAMESE TRANSLATION' block.\n"
            "4. NO PLOT 'CORRECTIONS': Accept the plot as provided (e.g., wings, skills).\n"
            "5. NO HALLUCINATIONS: Quote exact phrases. Search thoroughly before claiming missing content.\n"
            "6. DIALOGUE FORMAT: Do NOT suggest adding speaker tags (e.g., 'Name nói', 'Name bảo') if they are not in the source text. Follow the source's dialogue structure strictly.\n"
            "7. CHARACTER NAMES & SUFFIXES: Check if 'Yoohyun-ie', 'Yerim-ie', or other '-ie' suffixes exist in the 'ENGLISH REFERENCE'. If they do, they MUST be kept in the Vietnamese translation (e.g., 'Yoohyun-ie' should be 'Yoohyun-ie', not just 'Yoohyun' or 'Cậu Yoohyun'). Report as an error if missing.\n\n"
            "==== FORMAT (Only if errors exist) ====\n"
            "### Dòng [Số]:\n"
            "- [Lỗi]: \"[Trích dẫn cụm từ sai]\" - [Lý do sai]\n"
            "- [Gợi ý]: [Cách sửa đúng]\n\n"
            "==== KOREAN SOURCE (ORIGINAL) ====\n"
            f"{kor_chunk}\n\n"
            f"{eng_reference}"
            "==== VIETNAMESE TRANSLATION (WITH LINE NUMBERS) ====\n"
            f"{vi_chunk_with_nums}"
        )
        
        report_chunk = generate_with_retry(model='gemini-2.5-flash', contents=prompt_qc, system_instruction="Professional QC Editor. Report using line numbers. Focus on accuracy.")
        if report_chunk and report_chunk.strip():
            full_report.append(report_chunk)
            print(f"[{i+1}/{num_chunks}] Da phat hien loi o phan {i+1}.")
        else:
            print(f"[{i+1}/{num_chunks}] Phan {i+1} OK (Khong co loi).")
        
        # BUOC PHU: TRICH XUAT THUAT NGU MOI (Chua co trong Glossary)
        prompt_extract = (
            "==== TASK ====\n"
            "Identify character names, organizations, items, or special skills in the 'SOURCE TEXT' that are NOT present in the 'REFERENCE GLOSSARY'.\n\n"
            "==== REFERENCE GLOSSARY ====\n"
            f"{glossary_text}\n"
            f"{notes_text}\n\n"
            "==== SOURCE TEXT ====\n"
            f"--- KOREAN ---\n{kor_chunk}\n\n"
            "==== OUTPUT FORMAT ====\n"
            "- [Original Term]: [Suggested Vietnamese Translation]\n"
            "Return ONLY the list. If nothing new, return an empty string."
        )
        
        extract_chunk = generate_with_retry(model='gemini-2.5-flash', contents=prompt_extract, system_instruction="Glossary Extractor. Find new names/terms not in glossary.")
        if extract_chunk and extract_chunk.strip():
            new_terms_suggestions.append(extract_chunk.strip())

        time.sleep(2)

    # Lưu báo cáo
    os.makedirs('output', exist_ok=True)
    with open('output/qc_report.txt', 'w', encoding='utf-8') as f:
        # Loc bo cac gia tri None (neu co) truoc khi join
        clean_report = [str(r) for r in full_report if r is not None]
        f.write("\n".join(clean_report))
        
    # Ghi thêm phần thuật ngữ mới gợi ý ra file riêng
    with open('output/new_glossary_terms.txt', 'w', encoding='utf-8') as f:
        f.write(f"### THUẬT NGỮ/NHÂN VẬT MỚI GỢI Ý - {time.strftime('%d/%m/%Y %H:%M:%S')}\n")
        if new_terms_suggestions:
            f.write("\n")
            # Deduplicate basic
            unique_terms = list(set(new_terms_suggestions))
            f.write("\n".join(unique_terms))
            print(f"✨ Da luu thuật ngữ mới vào: output/new_glossary_terms.txt")
        else:
            f.write("\n(Không tìm thấy thuật ngữ mới nào trong lần chạy này.)")
            print(f"ℹ️ Không tìm thấy thuật ngữ mới nào.")
    
    print("\n✅ DA HOAN THANH QC! Vui long xem ket qua tai: output/qc_report.txt")

if __name__ == "__main__":
    run_qc()
