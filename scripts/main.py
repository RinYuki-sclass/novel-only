import os
import time
import sys
import argparse
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

def generate_with_retry(model, contents, system_instruction, retries=5):
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

def run_pipeline(mode='all'):
    print(f"🚀 DANG KHOI CHAY HE THONG {'DICH THUAT' if mode=='all' else 'RE-QC REFINE'}...")
    
    eng_path = 'input/trans/eng.txt'
    kor_path = 'input/trans/kor.txt'
    # Nếu chạy RE-QC, lấy input từ vi_final.txt của đợt trước hoặc vi_to_qc.txt
    draft_input_path = 'output/vi_final.txt' if mode == 'refine' else None
    
    glossary_path = 'glossary/glossary.md'
    notes_path = 'glossary/personal_notes.md'
    
    if not os.path.exists(eng_path) or not os.path.exists(kor_path):
        print("[ERROR] Thieu file dau vao trong 'input/trans/'")
        return

    with open(eng_path, 'r', encoding='utf-8') as f: eng_text = f.read()
    with open(kor_path, 'r', encoding='utf-8') as f: kor_text = f.read()
    with open(glossary_path, 'r', encoding='utf-8') as f: glossary_text = f.read() if os.path.exists(glossary_path) else ""
    with open(notes_path, 'r', encoding='utf-8') as f: notes_text = f.read() if os.path.exists(notes_path) else ""

    if mode == 'refine':
        if not os.path.exists(draft_input_path):
            print(f"[ERROR] Thieu file draft de re-QC: {draft_input_path}")
            return
        with open(draft_input_path, 'r', encoding='utf-8') as f: draft_text = f.read()
    
    # Chia theo doan van de dam bao khop 1:1
    eng_paragraphs = [p.strip() for p in eng_text.split('\n') if p.strip()]
    kor_paragraphs = [p.strip() for p in kor_text.split('\n') if p.strip()]
    draft_paragraphs = [p.strip() for p in draft_text.split('\n') if p.strip()] if mode == 'refine' else []
    
    # Check 1:1 match if refining
    if mode == 'refine':
        max_orig = max(len(eng_paragraphs), len(kor_paragraphs))
        if len(draft_paragraphs) != max_orig:
            print(f"⚠️ Canh bao: So doan van trong draft ({len(draft_paragraphs)}) khong khop voi ban goc ({max_orig}).")
            print("AI se co gang xu ly nhung ket qua co the bi lech.")
    
    # Gom 15 doan vao 1 chunk để dịch
    chunk_size = 15
    num_chunks = (max(len(eng_paragraphs), len(kor_paragraphs)) + chunk_size - 1) // chunk_size
    
    final_output = []
    print(f"📦 Da chia thanh {num_chunks} phan de xu ly.")

    for i in range(num_chunks):
        start = i * chunk_size
        end = start + chunk_size
        eng_chunk = "\n\n".join(eng_paragraphs[start:end])
        kor_chunk = "\n\n".join(kor_paragraphs[start:end])
        
        print(f"\n--- Dang xu ly Phan {i+1}/{num_chunks} ---")
        
        draft = ""
        if mode == 'all':
            # BUOC 1: DICH THO
            sys_draft = (
                "You are a professional novel translator. Translate English into natural Vietnamese. "
                "STRICT RULE for character names: ONLY include affectionate suffixes (e.g., -ie, -ah, -ya) IF they are already present in the source text. "
                "If a name appears without a suffix (e.g., 'Yoohyun'), do NOT add one. "
                "If it has a suffix (e.g., 'Yoohyun-ie'), keep it exactly as is.\n"
                "Output ONLY the translation. No explanations."
            )
            draft = generate_with_retry(model='gemini-2.5-flash', contents=eng_chunk, system_instruction=sys_draft)
        else:
            # Lấy draft từ file đã có
            draft = "\n\n".join(draft_paragraphs[start:end])

        # BUOC 2: REFINE & POLISH (QC)
        sys_refine = (
            "You are a strict novel editor. Refine the Vietnamese translation by comparing it with the English and Korean sources. "
            "STRICT RULES:\n"
            "1. Output ONLY the final Vietnamese text. No commentary.\n"
            "2. DIALOGUE STRUCTURE: Follow the source dialogue structure EXACTLY. Do NOT add speaker tags (e.g., 'Name said', 'Name nói') if they are not present in the source. Keep the dialogue format identical to the source.\n"
            "3. CHARACTER NAMES & SUFFIXES: "
            "STRICTLY keep affectionate suffixes like '-ie', '-ah', '-ya' IF they appear in the English source (e.g., 'Yoohyun-ie' -> 'Yoohyun-ie'). "
            "DO NOT remove them. DO NOT change them to 'Cậu' or 'Em'. If the English has the suffix, the Vietnamese MUST have it too.\n"
            "4. Keep 'ahjussi' as is. Keep suffixes like -ssi, -nim, -gun.\n"
            "5. Follow the Glossary strictly.\n"
            "6. NO creative rewriting, stay true to the source."
        )
        prompt_refine = (
            f"--- GLOSSARY ---\n{glossary_text}\n\n"
            f"--- PERSONAL NOTES ---\n{notes_text}\n\n"
            f"--- ENGLISH SOURCE (Check for suffixes like -ie, -ah here) ---\n{eng_chunk}\n\n"
            f"--- KOREAN SOURCE ---\n{kor_chunk}\n\n"
            f"--- CURRENT TRANSLATION ---\n{draft}"
        )
        refined = generate_with_retry(model='gemini-2.5-flash', contents=prompt_refine, system_instruction=sys_refine)
        
        # Cleanup any accidental AI chatter (loc bo cac dong nhan xet)
        lines = refined.strip().split('\n')
        clean_lines = [l for l in lines if not l.startswith(('*', 'Đây là', 'Bản dịch', 'Tuyệt vời', 'Đã sửa'))]
        final_output.append("\n".join(clean_lines))
        
        print(f"✅ Hoan tat phan {i+1}")
        time.sleep(1)

    # Luu ket qua
    os.makedirs('output', exist_ok=True)
    with open('output/vi_final.txt', 'w', encoding='utf-8') as f:
        f.write("\n\n".join(final_output))
    
    print(f"\n🎉 DA HOAN THANH{' RE-QC' if mode=='refine' else ''}!")

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('--mode', type=str, default='all', choices=['all', 'refine'], help='Set mode to all (Draft+Refine) or refine (QC only)')
    args = parser.parse_args()
    
    run_pipeline(mode=args.mode)
