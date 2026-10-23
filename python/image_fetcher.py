import json
import os
import urllib.parse
import requests

# ===== 配置区 ===== 
"""
# 暂时无法使用
# Pexels API（免费，注册即用：pexels.com/api）
PEXELS_API_KEY = os.environ.get("PEXELS_API_KEY", "")
PEXELS_URL = "https://api.pexels.com/v1/search"
"""

# Unsplash API（免费：unsplash.com/developers）
UNSPLASH_ACCESS_KEY = os.environ.get("UNSPLASH_ACCESS_KEY", "你的Unsplash KEY")

IMAGE_DIR = "images"
os.makedirs(IMAGE_DIR, exist_ok=True)

# ===== 图片获取函数 =====
def search_unsplash(query, output_path):
    """方案A: 从 Unsplash 搜索并下载图片"""
    url = "https://api.unsplash.com/search/photos"
    headers = {"Authorization": f"Client-ID {UNSPLASH_ACCESS_KEY}"}

    # 参数说明:
    # per_page=1: 只取最相关的一张
    # orientation=landscape: 只要横图（适合PPT）
    params = {"query": query, "per_page": 1, "orientation": "landscape"}

    try: 
        resp = requests.get(url, headers=headers, params=params, timeout=15)
        resp.raise_for_status() # 检查HTTP错误

        data = resp.json()
        if data.get("results") and len(data["results"]) > 0:
            # 获取图片的regular尺寸链接
            img_url = data["results"][0]["urls"]["regular"]

            # 下载图片
            img_data = requests.get(img_url, timeout=20).content
            with open(output_path, "wb") as f:
                f.write(img_data)
            return True
        else: 
            print(f"[Image] Unsplash 未找到关于 '{query}' 的图片")
            return False

    except Exception as e:
        print(f"[Image] 请求 Unsplash 失败: {e}")
        return False

def generate_image(prompt, output_path):
    """方案B: 使用 Pollinations 生成图片（无需API Key）"""
    encoded = urllib.parse.quote(prompt)
    url = f"https://image.pollinations.ai/prompt/{encoded}?width=1024&height=768&nologo=true"
    try:
        resp = requests.get(url, timeout=60)
        if resp.status.code == 200:
            with open(output_path, "wb") as f:
                f.write(resp.content)
            return True
        return False
    except Exception as e:
        print(f"[Image] AI 生成图片失败: {e}")
        return False

# ===== 主流程 =====
with open("optimized.json", "r", encoding="utf-8") as f:
    data = json.load(f)

for i, slide in enumerate(data.get("slides", [])):
    # 兼容旧的bullets结构和新的items结构
    layout = slide.get("layout", "")

    if (layout == "parallel_cards" or layout == "timeline") and "items" in slide:
        # 处理多图卡片
        for j, item in enumerate(slide["items"]):
            prompt = item.get("image_prompt", "")
            img_type = item.get("image_type", "search")
            if not prompt:
                continue

            output_path = os.path.join(IMAGE_DIR, f"slide_{i + 1}_card_{j + 1}.jpg")
            success = False

            if img_type == "generate":
                success = generate_image(prompt, output_path)
            else:
                success = search_unsplash(prompt, output_path)

            if success:
                item["image_path"] = output_path
                print(f"[Image] 第{i + 1}页 卡片{j + 1}配图已保存")
    else:    
        prompt = slide.get("image_prompt", "")
        img_type = slide.get("image_type", "search")
    
        if not prompt:
            continue

        output_path = os.path.join(IMAGE_DIR, f"slide_{i + 1}.jpg") # 使用join聚合路径
        success = False

        if img_type == "generate":
            print(f"[Image] 第{i + 1}页 正在 AI 生成图片...")
            success = generate_image(prompt, output_path)
        else:
            print(f"[Image] 第{i + 1}页 正在 Unsplash 搜索图片...")
            success = search_unsplash(prompt, output_path)
        
        if success:
            slide["image_path"] = output_path
            print(f"[Image] 第{i + 1}页 配图已保存: {output_path}")

# 保存更新后的JSON（包含 image_path 字段）
with open("optimized.json", "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=4)
