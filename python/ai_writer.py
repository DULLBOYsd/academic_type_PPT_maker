import json
import os
from openai import OpenAI

# 配置区
API_KEY = os.environ.get("DEEPSEEK_API_KEY", "你的API KEY");
BASE_URL = "https://api.deepseek.com";
MODEL = "deepseek-chat"

INPUT_JSON = "output.json"     # C++生成的文件
OUTPUT_JSON = "optimized.json" # AI优化后的文件

# 初始化客户端
client = OpenAI(api_key=API_KEY, base_url=BASE_URL)

# 读取C++输出的JSON
with open(INPUT_JSON, "r", encoding="utf-8") as f:
    data = json.load(f)

print(f"[Python] 读取到 {len(data['slides'])} 页原始内容")

# 构造提示词
prompt = f"""
你是一位资深的 PPT 设计师，请帮我优化以下PPT大纲。

要求：
# 任务一：优化文字内容（标题精炼、要点专业）

1.每个标题精炼、有吸引力，不超过 15 个字。
2.每页要点保留 3-5 条，每条不超过 25 个字。
3.使用专业的商务语言，去掉口语化表达。
4.可以适当提炼和归并，但不要编造原文没有的数据。
5.总结采用一句话概括，单独成句不与要点并列
6.严格输出 JSON 格式，不要添加任何解释文字

# 任务二：为每一页选择最合适的布局类型

可选的 layout 类型：
- "cover"：封面（只有第 1 页可用）
- "title_bullets"：标题 + 要点列表（最通用）
- "parallel_cards"：若干个并列要点，适合“三大战略”这类并列结构
- "left_right"：左右对比，适合“现状 vs 目标”这类对比内容
- "timeline"：时间线，适合带年份或步骤的内容
- "quote"：金句页，适合一句话结论
- "section"：章节过渡页，适合只有一个大标题的页

选择规则：
- 第 1 页必须用 "cover"
- 如果内容包含年份/步骤 → "timeline"
- 如果有若干个并列要点 → "parallel_cards"
- 如果是对比关系 → "left_right"（目前只支持两点对比（bullets里面只能填写两点）） 
- 如果是绘制图表 → "chart"（暂不支持，先别使用）
- 如果只有一句话 → "quote"（这句结论应该直接写在conclusion参数里面，且除了conclusion其他参数里面没有值）
- 如果是中间不同部分的过渡页 → "section"（描述应为第×部分 + 该部分内容。其中应该只有标题 + 副标题，且标题内容为第×部分，副标题为该部分内容的简介）
- 其他情况用 "title_bullets"

# 任务三：为每一页输出配图决策
- "image_prompt": 用于搜索或生成图片的英文描述（20词以内）
- "image_type": 图片获取方式，可选 "search"（搜索真实图片）或 "generate"（AI生成示意图）
  - "search": 需要真实照片、产品图、场景图时使用，通过图库API搜索
  - "generate": 需要示意图、概念图、流程图时使用，调用AI图像生成API
- "image_placement": 图片位置，可选 "right"（右侧）、"full"（全图）、"background"（背景）

原始内容：
{json.dumps(data, ensure_ascii=False, indent=2)}

# ===== 内容输出格式限定 ===== 

**极其重要：图片相关的字段（image_prompt等）以及高亮字段（highlights）都必须放在 slides 数组的每一页对象内部！绝对不能放在最外层！**

1. 如果需要嵌入图片的版式为 parallel_cards/timeline，按照以下格式输出 JSON：
{{
  "slides": [
    {{
      "layout": "parallel_cards",
      "title": "三大核心战略",
      "items": [
        {{
          "text": "技术突破：加大研发投入",
          "image_prompt": "futuristic technology lab research",
          "image_type": "search"
        }},
        {{
          "text": "市场拓展：进入华东华南",
          "image_prompt": "business expansion map china",
          "image_type": "search"
        }},
        {{
          "text": "团队建设：引进高端人才",
          "image_prompt": "diverse professional team meeting",
          "image_type": "search"
        }}
      ],
      "highlights": [
        {{"text": "技术突破", "color": "blue", "bold": true}},
        {{"text": "研发投入", "color": "red", "bold": true}},
        {{"text": "华东华南", "color": "green", "bold": true}}
      ],
      "conclusion": "坚持技术创新与人才驱动"
    }}
  ]
}}

2. 如果是其他普通版式（如 title_bullets / left_right），请按照以下结构输出：
{{
  "title": "整体标题",
  "slides": [
    {{
      "layout": "cover",
      "title": "整体标题",
      "subtitle": "副标题",
      "image_prompt": "business meeting presentation",
      "image_type": "search",
      "image_placement": "background"
    }},
    {{
      "layout": "title_bullets",
      "title": "市场分析",
      "bullets": ["营收达到1200万", "新增客户46家"],
      "image_prompt": "business growth chart",
      "image_type": "search",
      "image_placement": "right",
      "highlights": [
        {{"text": "1200万", "color": "red", "bold": true}},
        {{"text": "46家", "color": "blue", "bold": false}}
      ],
      "conclusion": "本季度业绩稳步增长"
    }}
  ]
}}

注意事项：
1. 标题页(cover)不需要conclusion和highlights，可以让它为空。
2. 除了cover/section，其他版式下都需要结论参数conclusion。
3. cover/section需要添加参数subtitle，其他版式暂不添加。
4. cover/title_bullets/section需要一张图片；parallel_cards/timeline需要若干张图片；其余的版式暂不需要图片。
5. **关于 highlights：** title_bullets/parallel_cards/left_right/timeline 这些版式，必须从 bullets 或 items 的文字中，提取 2-5 个核心词汇。**提取的 text 必须与原文字符100%完全一致（包含数字和标点）**，以便程序进行文本切割。
6. color 字段的可选值限定为："red", "blue", "green", "orange", "purple", "gray"。bold 字段限定为 true 或 false。
"""

# 调用Deepseek API
print("[Python] 正在调用 Deepseek API...")
response = client.chat.completions.create(
    model=MODEL,
    messages=[
        {"role": "system", "content": "你是一个专业的PPT大纲优化助手，只输出JSON。"},
        {"role": "user", "content": prompt},
    ],
    response_format={"type": "json_object"},
    temperature=0.7
)

# 解析AI返回结果并保存
optimized_text = response.choices[0].message.content;

try:
    optimized_data = json.loads(optimized_text)
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(optimized_data, f, ensure_ascii=False, indent=4)
    print(f"[Python] 已生成优化后的JSON: {OUTPUT_JSON}")
    print(f"[Python] 优化后共 {len(optimized_data['slides'])} 页")
except json.JSONDecodeError as e:
    print(f"[Python] AI返回的内容不是合法JSON: {e}")
    print("原始返回内容: ")
    print(optimized_text)

