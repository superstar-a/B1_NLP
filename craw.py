import asyncio
import re
import os
import sys
import json
from playwright.async_api import async_playwright
import pandas as pd

# Chỉnh sửa encoding console Windows tránh lỗi charmap print
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

# Danh sách các chủ đề thời sự cần cào
# Danh sách các chủ đề thời sự cần cào với các từ khóa con đa dạng
TOPICS = {
    "Kinh_Te": [
        "kinh tế giá vàng lạm phát",
        "bất động sản nhà đất",
        "lãi suất ngân hàng chứng khoán",
        "thị trường tài chính kinh doanh"
    ],
    "Xa_Hoi": [
        "giao thông hạ tầng kẹt xe",
        "đời sống đô thị xã hội",
        "an sinh xã hội y tế giáo dục",
        "thời sự xã hội tin tức"
    ],
    "Cong_Nghe": [
        "trí tuệ nhân tạo AI",
        "công nghệ bán dẫn vi mạch",
        "đổi mới sáng tạo công nghệ mới",
        "ứng dụng công nghệ phần mềm"
    ],
    "Moi_Truong": [
        "ô nhiễm không khí bụi mịn",
        "biến đổi khí hậu thời tiết",
        "môi trường rác thải nhựa",
        "năng lượng xanh bảo vệ môi trường"
    ]
}

TARGET_COUNT = 5000  # Số câu cần cào cho mỗi nhóm
DATA_DIR = os.path.abspath("./crawled_data")

def clean_sentence(text):
    """Làm sạch văn bản: xóa link, hashtag, icon thừa, tách câu ngắn."""
    text = re.sub(r'http\S+', '', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def save_json(group_name, sentences_list):
    """Lưu dữ liệu dạng JSON thô và tự động xuất ra Excel ngay lập tức."""
    os.makedirs(DATA_DIR, exist_ok=True)
    json_path = os.path.join(DATA_DIR, f"{group_name}.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(sentences_list, f, ensure_ascii=False, indent=2)
    # Tự động xuất ra file Excel để đĩa luôn có dữ liệu mới nhất
    try:
        export_to_excel()
    except Exception:
        pass

def load_json(group_name):
    """Tải dữ liệu đã cào trước đó nếu có."""
    json_path = os.path.join(DATA_DIR, f"{group_name}.json")
    if os.path.exists(json_path):
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                return set(json.load(f))
        except Exception:
            return set()
    return set()

def export_to_excel(excel_path="Threads_Thoi_Su_Output.xlsx"):
    """Tạo/cập nhật file Excel chuẩn từ các file JSON đã cào được."""
    os.makedirs(DATA_DIR, exist_ok=True)
    has_written = False
    
    try:
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            for group_name in TOPICS.keys():
                json_path = os.path.join(DATA_DIR, f"{group_name}.json")
                if os.path.exists(json_path):
                    with open(json_path, "r", encoding="utf-8") as f:
                        data_list = json.load(f)
                    if data_list:
                        df = pd.DataFrame({
                            "stt": range(1, len(data_list) + 1),
                            "nội dung câu văn": data_list
                        })
                        df.to_excel(writer, sheet_name=group_name, index=False)
                        has_written = True
            
            if not has_written:
                df_empty = pd.DataFrame({"stt": [], "nội dung câu văn": []})
                df_empty.to_excel(writer, sheet_name="General", index=False)
        print(f"-> Đã xuất/cập nhật file Excel thành công: {excel_path}")
    except PermissionError:
        print("\n" + "!"*70)
        print(f"[CẢNH BÁO] Không thể lưu file '{excel_path}' vì file đang được MỞ bằng Microsoft Excel!")
        print("-> VUI LÒNG ĐÓNG FILE EXCEL TRÊN MÁY BẠN LẠI để chương trình cập nhật dữ liệu.")
        print("!"*70 + "\n")

async def crawl_threads_topic(page, group_name, keywords_list, target_count=5000):
    sentences = load_json(group_name)
    print(f"\n[{group_name}] Đã có sẵn {len(sentences)} câu từ lần cào trước.")
    
    if len(sentences) >= target_count:
        print(f"[{group_name}] Đã đạt mục tiêu {target_count} câu, chuyển sang nhóm tiếp theo.")
        return list(sentences)[:target_count]

    for keyword in keywords_list:
        if len(sentences) >= target_count:
            break
            
        print(f"\n---> Đang tìm kiếm từ khóa: '{keyword}' (Đã có {len(sentences)}/{target_count} câu) <---")
        url = f"https://www.threads.net/search?q={keyword}&serp_type=default"
        
        try:
            await page.goto(url, wait_until="domcontentloaded", timeout=60000)
            await page.wait_for_timeout(2500)
        except Exception as e:
            print(f"Lỗi tải trang '{keyword}': {e}")
            continue

        no_new_count = 0  # Đếm số lần cuộn liên tiếp không thêm được câu mới
        max_scrolls_per_keyword = 500

        for scroll in range(max_scrolls_per_keyword):
            if len(sentences) >= target_count:
                break
                
            try:
                texts = await page.evaluate('''() => {
                    const nodes = document.querySelectorAll('div[data-pressable-container="true"] span, [dir="auto"]');
                    return Array.from(nodes).map(n => n.innerText || '');
                }''')
                
                added_this_round = 0
                for txt in texts:
                    txt_clean = clean_sentence(txt)
                    if len(txt_clean) >= 30 and txt_clean not in sentences:
                        sentences.add(txt_clean)
                        added_this_round += 1
                        if len(sentences) >= target_count:
                            break
                
                if added_this_round > 0:
                    no_new_count = 0
                    save_json(group_name, list(sentences))
                else:
                    no_new_count += 1

            except Exception as e:
                print(f"[{keyword}] Lỗi đọc DOM: {e}")

            # Nếu 15 lần cuộn liên tiếp không thu được thêm câu mới -> Từ khóa này đã cào hết bài!
            if no_new_count >= 15:
                print(f"-> [TỐI ƯU TỐC ĐỘ] Từ khóa '{keyword}' đã hết bài mới (15 lần cuộn không có dữ liệu mới). Chuyển từ khóa tiếp theo!")
                break

            await page.mouse.wheel(0, 3500)
            await page.wait_for_timeout(1000)
            print(f"[{group_name} | {keyword}] Tiến độ: {len(sentences)}/{target_count} câu (lần cuộn {scroll+1})...")

    final_list = list(sentences)[:target_count]
    save_json(group_name, final_list)
    return final_list

async def main():
    user_data_dir = os.path.abspath("./user_data")
    os.makedirs(user_data_dir, exist_ok=True)

    async with async_playwright() as p:
        context = await p.chromium.launch_persistent_context(
            user_data_dir=user_data_dir,
            headless=False,
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            viewport={"width": 1280, "height": 800}
        )
        
        page = context.pages[0] if context.pages else await context.new_page()

        print("="*60)
        print("Đang mở trình duyệt Threads...")
        await page.goto("https://www.threads.net/", wait_until="domcontentloaded")
        print("-> Vui lòng kiểm tra giao diện trình duyệt vừa mở.")
        print("-> Nếu CHƯA ĐĂNG NHẬP Instagram/Threads, hãy tiến hành đăng nhập trực tiếp trên trình duyệt đó.")
        print("="*60)
        
        input("\n>>> Sau khi đã đăng nhập thành công (hoặc nếu đã đăng nhập rồi), hãy quay lại đây và nhấn phím ENTER để bắt đầu cào... <<<\n")

        try:
            for group_name, keywords_list in TOPICS.items():
                print(f"\n================ BẮT ĐẦU NHÓM: {group_name} ================")
                await crawl_threads_topic(page, group_name, keywords_list, TARGET_COUNT)
                export_to_excel()
        except KeyboardInterrupt:
            print("\n\n[DỪNG THỦ CÔNG] Bạn đã bấm dừng (Ctrl+C). Đang xuất dữ liệu hiện có ra Excel...")
            export_to_excel()
        finally:
            try:
                await context.close()
            except Exception:
                pass
            print("=> Đã đóng trình duyệt an toàn.")

if __name__ == "__main__":
    asyncio.run(main())