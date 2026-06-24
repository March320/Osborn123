# 在仓库中生成《我的大学人生规划》PPT 的说明（熊怡嘉）

此文件包含：
1) 生成 PPTX 所需的说明与一键 Python 脚本（使用 python-pptx）。
2) 脚本会生成文件名：`WPS演示文稿202505050214熊怡嘉.pptx`，并放在本地运行目录中。

注意事项：
- 我在此尝试将生成脚本与说明提交到仓库（presentations/ 目录）。如果对仓库没有写入权限，提交将失败；失败时我会把内容直接在对话中发给你作为备份。
- 该脚本依赖 python-pptx 和 pillow。如要下载示例图片，脚本也使用 requests（可选）。

文件位置（仓库内）：presentations/README_生成PPT_熊怡嘉.md

----

使用说明（在本地运行���：

1. 在本地创建并进入工作目录。
2. 安装依赖（建议使用虚拟环境）：
   pip install python-pptx pillow requests
3. 将同目录下的脚本文件 `generate_presentation.py` 放入并运行：
   python generate_presentation.py
4. 运行后在当前目录将生成：`WPS演示文稿202505050214熊怡嘉.pptx`。

----

下面包含用于生成 PPT 的 Python 脚本（请保存为 generate_presentation.py）：

```python
# generate_presentation.py
# 生成《我的大学人生规划》演示文稿（16页示例）
# 依赖：python-pptx, pillow, requests (若使用在线示例图)

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import requests
from io import BytesIO

# 可替换的元数据
author_name = '熊怡嘉'
major = '财税大数据应用'
date_str = '2026-06-18'
output_filename = 'WPS演示文稿202505050214熊怡嘉.pptx'

# 示例图片（若无法下载，脚本会跳过图片插入）
image_urls = [
    'https://images.unsplash.com/photo-1503676260728-1c00da094a0b?w=1200',
    'https://images.unsplash.com/photo-1521737604893-d14cc237f11d?w=1200',
    'https://images.unsplash.com/photo-1504384308090-c894fdcc538d?w=1200'
]

prs = Presentation()
# 设置默认母版字体（只示例：后续可在 PowerPoint 中微调）
slide_width = prs.slide_width
slide_height = prs.slide_height

# 工具函数：添加标题+副标题布局
def add_title_slide(title, subtitle=''):
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    title_tf = slide.shapes.title.text_frame
    title_tf.text = title
    if subtitle:
        if slide.placeholders and len(slide.placeholders) > 1:
            try:
                slide.placeholders[1].text = subtitle
            except Exception:
                pass
    return slide

# 工具函数：创建带要点的普通幻灯片
def add_bullet_slide(title, bullets):
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = title
    body = slide.shapes.placeholders[1].text_frame
    body.clear()
    for i, b in enumerate(bullets):
        if i == 0:
            p = body.paragraphs[0]
            p.text = b
        else:
            p = body.add_paragraph()
            p.text = b
        p.level = 0
    return slide

# 插入图片（如果可下载）
def add_picture_slide(title, image_url=None):
    slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = title
    if image_url:
        try:
            r = requests.get(image_url, timeout=8)
            if r.status_code == 200:
                img = BytesIO(r.content)
                left = Inches(1)
                top = Inches(1.5)
                slide.shapes.add_picture(img, left, top, width=Inches(8))
        except Exception as e:
            print('无法下载图片:', e)
    return slide

# 1 封面
add_title_slide('我的大学人生规划', f'{author_name} • {major} • {date_str}')

# 2 目录
add_bullet_slide('目录', [
    '个人简介', '大学目标', '职业规划与能力评估', '实习与实践计划', '个人发展与技能提升', '未来展望', '致谢'
])

# 3 个人简介
add_bullet_slide('个人简介', [
    '兴趣爱好：数据分析、阅读、旅行',
    '自我评价：积极、严谨、善于学习和团队合作',
    '年级与背景：在校生，主修财税与大数据应用'
])

# 4 大学目标（总览）
add_bullet_slide('大学目标（总览）', [
    '短期目标：提高成绩，参加项目实践',
    '学业目标：GPA目标、证书（英语/岗位相关）',
    '技能目标：掌握 Python/SQL/财务软件',
    '长期目标：毕业后3-5年内稳定职业发展'
])

# 5 短期目标
add_bullet_slide('短期目标（大学期间）', [
    'GPA ≥ 3.2，完成核心课程',
    '参加校内外比赛与社团，拓展人脉',
    '每学期完成至少一个小项目'
])

# 6 学业目标
add_bullet_slide('学业目标', [
    '通过英语四/六级或同等级证书',
    '考取专业相关证书（税务/会计/数据分析）',
    '专业课程成绩稳定提升'
])

# 7 技能提升目标（雷达/图示在 PPT 中可手动制作）
add_bullet_slide('技能提升目标', [
    '掌握 Python、SQL、Excel 高级技能',
    '熟悉财税相关软件及数据可视化工具',
    '沟通与项目管理能力提升'
])

# 8 长期目标
add_bullet_slide('长期目标（毕业后3-5年）', [
    '进入财税或数据分析行业，担任数据分析师/税务顾问',
    '积累产品/项目经验，争取管理岗位',
    '实现个人生活目标（储蓄/旅行）'
])

# 9 职业规划与能力评估
add_bullet_slide('职业规划与能力评估', [
    '现有优势：专业背景、学习能力强',
    '需提升：实战经验、项目交付能力',
    '行动：实习、项目实践、证书路径'
])

# 10 实习与实践计划
add_bullet_slide('实习与实践计划', [
    '目标公司：会计师事务所、税务顾问公司、金融或互��网企业',
    '时间安排：寒暑假/学期实习相结合',
    '期望成果：完成真实数据分析或税务项目'
])

# 11 实习计划详细时间表
add_bullet_slide('实习时间表（示例）', [
    '大二暑期：岗位体验/短期实习',
    '大三学期：项目实习（负责数据分析）',
    '大四：毕业实习/求职准备'
])

# 12 个人发展与技能提升
add_bullet_slide('个人发展计划', [
    '课程学习 + 在线课程（Coursera/慕课）',
    '项目实践（数据清洗、可视化、报表）',
    '提升软技能：沟通、PPT、时间管理'
])

# 13 成果展示（示例）
add_picture_slide('项目与成果展示', image_url=image_urls[0])

# 14 未来展望
add_bullet_slide('未来展望', [
    '5 年愿景：成为行业内的资深分析师或税务顾问',
    '持续学习，保持职业敏感度和创新能力'
])

# 15 风险与应对
add_bullet_slide('风险与应对', [
    '风险：岗位竞争、技能滞后',
    '对策：持续学习，建立导师/人脉网络'
])

# 16 致谢
add_picture_slide('致谢', image_url=image_urls[1])

# 保存文件
prs.save(output_filename)
print('已生成：', output_filename)
```

----

如果你希望我���续尝试把实际 .pptx 文件直接上传到仓库（而非脚本），请先授权或确认我可写入该仓库；否则我会把此脚本和说明提交到仓库（当前操作即为提交此 README），并把脚本内容也在对话中提供，供你在本地运行生成 PPTX。

祝好，若需我立刻继续上传或切换为直接在对话中发送脚本，请回复“继续上传 PPTX”或“使用脚本（方案 C）”。