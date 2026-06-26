#!/usr/bin/env python3
"""
ファスティング指導士養成講座
カウンセリング講義スライド生成スクリプト
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import copy

# カラーパレット
COLOR_PRIMARY = RGBColor(0x2E, 0x7D, 0x32)      # 深緑
COLOR_ACCENT = RGBColor(0x81, 0xC7, 0x84)        # 薄緑
COLOR_WARM = RGBColor(0xF9, 0xA8, 0x25)          # アンバー
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_DARK = RGBColor(0x1A, 0x1A, 0x2E)          # 濃紺
COLOR_LIGHT_BG = RGBColor(0xF1, 0xF8, 0xE9)     # 薄緑背景
COLOR_GRAY = RGBColor(0x75, 0x75, 0x75)
COLOR_ORANGE = RGBColor(0xE6, 0x51, 0x00)        # オレンジ強調
COLOR_TEAL = RGBColor(0x00, 0x89, 0x8A)          # ティール

prs = Presentation()
prs.slide_width = Inches(13.33)
prs.slide_height = Inches(7.5)

BLANK_LAYOUT = prs.slide_layouts[6]  # blank


def add_rect(slide, left, top, width, height, fill_color=None, line_color=None, line_width=None):
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
        if line_width:
            shape.line.width = Pt(line_width)
    else:
        shape.line.fill.background()
    return shape


def add_text(slide, text, left, top, width, height,
             font_size=18, bold=False, color=COLOR_DARK,
             align=PP_ALIGN.LEFT, italic=False, wrap=True):
    txBox = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return txBox


def add_multiline_text(slide, lines, left, top, width, height,
                       font_size=16, bold=False, color=COLOR_DARK,
                       align=PP_ALIGN.LEFT, line_spacing=1.2):
    txBox = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = line
        run.font.size = Pt(font_size)
        run.font.bold = bold
        run.font.color.rgb = color
    return txBox


# ============================================================
# スライド 1: タイトルスライド
# ============================================================
slide = prs.slides.add_slide(BLANK_LAYOUT)

# 背景グラデーション風（二層）
add_rect(slide, 0, 0, 13.33, 7.5, fill_color=RGBColor(0x1B, 0x5E, 0x20))
add_rect(slide, 0, 4.5, 13.33, 3.0, fill_color=RGBColor(0x12, 0x43, 0x16))

# アクセントライン
add_rect(slide, 0.6, 1.0, 0.08, 5.0, fill_color=COLOR_WARM)

# タイトル
add_text(slide, "カウンセリングの本質", 1.0, 1.2, 11.5, 1.4,
         font_size=46, bold=True, color=COLOR_WHITE, align=PP_ALIGN.LEFT)
add_text(slide, "〜 答えを教えるのではなく、気づきを生む対話へ 〜", 1.0, 2.7, 11.5, 0.8,
         font_size=22, bold=False, color=COLOR_ACCENT, align=PP_ALIGN.LEFT)

# サブタイトル
add_rect(slide, 1.0, 3.7, 6.0, 0.06, fill_color=COLOR_WARM)
add_text(slide, "ファスティング指導士養成講座", 1.0, 3.9, 8.0, 0.6,
         font_size=18, color=RGBColor(0xC8, 0xE6, 0xC9), align=PP_ALIGN.LEFT)
add_text(slide, "カウンセリング講義", 1.0, 4.5, 8.0, 0.5,
         font_size=16, color=RGBColor(0xA5, 0xD6, 0xA7), align=PP_ALIGN.LEFT)

# 右下装飾
add_text(slide, "「悩みに気づかせる場が、最大の教育になる」", 2.0, 6.0, 10.0, 0.8,
         font_size=15, italic=True, color=COLOR_WARM, align=PP_ALIGN.CENTER)


# ============================================================
# スライド 2: 本日の講義内容（アジェンダ）
# ============================================================
slide = prs.slides.add_slide(BLANK_LAYOUT)
add_rect(slide, 0, 0, 13.33, 7.5, fill_color=COLOR_LIGHT_BG)
add_rect(slide, 0, 0, 13.33, 1.1, fill_color=COLOR_PRIMARY)
add_text(slide, "本日の講義内容", 0.4, 0.15, 12.0, 0.8,
         font_size=28, bold=True, color=COLOR_WHITE)

items = [
    ("01", "カウンセリングとは何か？", "よくある誤解を解く"),
    ("02", "カウンセリングの真の目的", "「気づき」を生む対話"),
    ("03", "悩みの構造を理解する", "なぜ一人では気づけないのか"),
    ("04", "なりたい自分を見つける", "目標設定のプロセス"),
    ("05", "環境と知識の重要性", "一緒に叶えるために"),
    ("06", "実践：カウンセリングの流れ", "問いかけの技術"),
]

for i, (num, title, sub) in enumerate(items):
    row = i // 2
    col = i % 2
    x = 0.5 + col * 6.4
    y = 1.4 + row * 1.85

    add_rect(slide, x, y, 6.0, 1.6, fill_color=COLOR_WHITE,
             line_color=COLOR_ACCENT, line_width=1.5)
    add_rect(slide, x, y, 0.7, 1.6, fill_color=COLOR_PRIMARY)
    add_text(slide, num, x + 0.05, y + 0.45, 0.6, 0.7,
             font_size=18, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)
    add_text(slide, title, x + 0.8, y + 0.15, 5.0, 0.6,
             font_size=16, bold=True, color=COLOR_DARK)
    add_text(slide, sub, x + 0.8, y + 0.85, 5.0, 0.5,
             font_size=12, color=COLOR_GRAY)


# ============================================================
# スライド 3: よくある誤解
# ============================================================
slide = prs.slides.add_slide(BLANK_LAYOUT)
add_rect(slide, 0, 0, 13.33, 7.5, fill_color=RGBColor(0xFF, 0xF8, 0xE1))
add_rect(slide, 0, 0, 13.33, 1.1, fill_color=COLOR_ORANGE)
add_text(slide, "01｜カウンセリングとは何か？ ── よくある誤解", 0.4, 0.15, 12.5, 0.8,
         font_size=24, bold=True, color=COLOR_WHITE)

# 左：誤解
add_rect(slide, 0.4, 1.3, 5.8, 5.6, fill_color=RGBColor(0xFF, 0xCC, 0xBC))
add_rect(slide, 0.4, 1.3, 5.8, 0.7, fill_color=RGBColor(0xBF, 0x36, 0x0C))
add_text(slide, "❌  よくある誤解", 0.6, 1.37, 5.4, 0.55,
         font_size=18, bold=True, color=COLOR_WHITE)

misconceptions = [
    "「全部答えてあげなきゃ」",
    "「全部伝えてあげなきゃ」",
    "「知識をたくさん教えること\nがカウンセリング」",
    "「専門家が正しい答えを\n提供する場」",
]
for j, m in enumerate(misconceptions):
    add_rect(slide, 0.6, 2.15 + j * 1.1, 5.2, 0.9,
             fill_color=COLOR_WHITE, line_color=RGBColor(0xBF, 0x36, 0x0C), line_width=1)
    add_text(slide, m, 0.85, 2.2 + j * 1.1, 4.8, 0.8,
             font_size=15, bold=True, color=RGBColor(0xBF, 0x36, 0x0C))

# 右：実態
add_rect(slide, 7.0, 1.3, 5.8, 5.6, fill_color=RGBColor(0xE8, 0xF5, 0xE9))
add_rect(slide, 7.0, 1.3, 5.8, 0.7, fill_color=COLOR_PRIMARY)
add_text(slide, "✅  カウンセリングの本質", 7.2, 1.37, 5.4, 0.55,
         font_size=18, bold=True, color=COLOR_WHITE)

truths = [
    "「気づきを生む場」",
    "「なりたい自分を見つける場」",
    "「悩みの本質に\n自分で気づいてもらう場」",
    "「最大の教育が\nできる場」",
]
for j, t in enumerate(truths):
    add_rect(slide, 7.2, 2.15 + j * 1.1, 5.4, 0.9,
             fill_color=COLOR_WHITE, line_color=COLOR_PRIMARY, line_width=1)
    add_text(slide, t, 7.45, 2.2 + j * 1.1, 5.0, 0.8,
             font_size=15, bold=True, color=COLOR_PRIMARY)

# 矢印代替
add_text(slide, "→", 6.2, 3.8, 0.8, 0.6,
         font_size=32, bold=True, color=COLOR_GRAY, align=PP_ALIGN.CENTER)


# ============================================================
# スライド 4: カウンセリングの真の目的
# ============================================================
slide = prs.slides.add_slide(BLANK_LAYOUT)
add_rect(slide, 0, 0, 13.33, 7.5, fill_color=COLOR_LIGHT_BG)
add_rect(slide, 0, 0, 13.33, 1.1, fill_color=COLOR_PRIMARY)
add_text(slide, "02｜カウンセリングの真の目的", 0.4, 0.15, 12.5, 0.8,
         font_size=28, bold=True, color=COLOR_WHITE)

# 中央メッセージ
add_rect(slide, 1.2, 1.3, 10.9, 1.5, fill_color=COLOR_PRIMARY)
add_text(slide, "カウンセリングとは、「答えを教える場」ではなく\n「自分に足りないものに気づかせる場」である",
         1.4, 1.35, 10.5, 1.4,
         font_size=20, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)

# 3つの目的
goals = [
    ("🔍", "気づきを生む", "クライアントが自分では見えていない\n悩みの本質や原因に気づく"),
    ("🌱", "なりたい自分を見つける", "理想の未来像を明確にし、\nそこへ向かう意欲を引き出す"),
    ("🤝", "行動への橋渡し", "気づいたことを一人で抱えるのではなく、\n必要な環境・知識につなげる"),
]

for i, (icon, title, desc) in enumerate(goals):
    x = 0.6 + i * 4.2
    add_rect(slide, x, 3.2, 3.9, 3.7, fill_color=COLOR_WHITE,
             line_color=COLOR_ACCENT, line_width=2)
    add_text(slide, icon, x + 0.1, 3.3, 3.7, 0.8,
             font_size=36, align=PP_ALIGN.CENTER)
    add_rect(slide, x, 4.15, 3.9, 0.6, fill_color=COLOR_ACCENT)
    add_text(slide, title, x + 0.1, 4.18, 3.7, 0.55,
             font_size=16, bold=True, color=RGBColor(0x1B, 0x5E, 0x20), align=PP_ALIGN.CENTER)
    add_text(slide, desc, x + 0.2, 4.9, 3.5, 1.8,
             font_size=13, color=COLOR_DARK, align=PP_ALIGN.CENTER)


# ============================================================
# スライド 5: なぜ一人では気づけないのか
# ============================================================
slide = prs.slides.add_slide(BLANK_LAYOUT)
add_rect(slide, 0, 0, 13.33, 7.5, fill_color=RGBColor(0xE3, 0xF2, 0xFD))
add_rect(slide, 0, 0, 13.33, 1.1, fill_color=COLOR_TEAL)
add_text(slide, "03｜悩みの構造 ── なぜ一人では気づけないのか？", 0.4, 0.15, 12.5, 0.8,
         font_size=26, bold=True, color=COLOR_WHITE)

# アイスバーグモデル
add_rect(slide, 0.5, 1.2, 5.5, 5.8, fill_color=COLOR_WHITE,
         line_color=RGBColor(0x90, 0xCA, 0xF9), line_width=1.5)
add_text(slide, "氷山モデル：悩みの構造", 0.7, 1.3, 5.1, 0.5,
         font_size=15, bold=True, color=COLOR_TEAL, align=PP_ALIGN.CENTER)

add_rect(slide, 0.8, 2.0, 4.9, 1.1, fill_color=RGBColor(0xB3, 0xE5, 0xFC))
add_text(slide, "水面上：表面的な悩み\n「体重が減らない」「続かない」", 1.0, 2.05, 4.5, 1.0,
         font_size=13, color=RGBColor(0x01, 0x57, 0x9B), align=PP_ALIGN.CENTER)

add_rect(slide, 0.8, 3.25, 4.9, 1.1, fill_color=RGBColor(0x81, 0xD4, 0xFA))
add_text(slide, "水面下：隠れた本音\n「自己否定感」「孤独感」", 1.0, 3.3, 4.5, 1.0,
         font_size=13, color=RGBColor(0x01, 0x57, 0x9B), align=PP_ALIGN.CENTER)

add_rect(slide, 0.8, 4.5, 4.9, 1.8, fill_color=RGBColor(0x29, 0xB6, 0xF6))
add_text(slide, "深部：根本原因\n「生活習慣の背景にある\nストレスや環境の問題」",
         1.0, 4.55, 4.5, 1.7,
         font_size=13, color=COLOR_WHITE, align=PP_ALIGN.CENTER)

# 右側説明
add_text(slide, "一人では気づけない理由", 6.3, 1.3, 6.6, 0.55,
         font_size=18, bold=True, color=COLOR_TEAL)
add_rect(slide, 6.3, 1.95, 6.7, 0.04, fill_color=COLOR_TEAL)

reasons = [
    ("👁️ 自分の中にいる限り\n客観視が難しい",
     "魚は水の中にいるから水が見えない。\n自分の当たり前が問題だと気づきにくい。"),
    ("💭 思い込みのフィルター",
     "「どうせ無理」「自分はダメ」という\n固定観念が本質を覆い隠す。"),
    ("🔄 孤独な試行錯誤の限界",
     "同じ視点でぐるぐると考えても\n新しい気づきは生まれにくい。"),
]

for i, (title, desc) in enumerate(reasons):
    y = 2.1 + i * 1.7
    add_rect(slide, 6.3, y, 6.7, 1.55, fill_color=COLOR_WHITE,
             line_color=RGBColor(0x29, 0xB6, 0xF6), line_width=1)
    add_text(slide, title, 6.5, y + 0.08, 3.0, 1.0,
             font_size=13, bold=True, color=COLOR_TEAL)
    add_text(slide, desc, 6.5, y + 0.7, 6.3, 0.8,
             font_size=12, color=COLOR_DARK)


# ============================================================
# スライド 6: 「気づき」が生まれる瞬間
# ============================================================
slide = prs.slides.add_slide(BLANK_LAYOUT)
add_rect(slide, 0, 0, 13.33, 7.5, fill_color=RGBColor(0xF3, 0xE5, 0xF5))
add_rect(slide, 0, 0, 13.33, 1.1, fill_color=RGBColor(0x6A, 0x1B, 0x9A))
add_text(slide, "「気づき」が生まれる瞬間", 0.4, 0.15, 12.5, 0.8,
         font_size=28, bold=True, color=COLOR_WHITE)

add_text(slide, "カウンセラーの問いかけが、クライアントの内側にある答えを引き出す",
         0.5, 1.2, 12.3, 0.7, font_size=18, bold=True,
         color=RGBColor(0x4A, 0x14, 0x8C), align=PP_ALIGN.CENTER)

# 対話例
add_rect(slide, 0.4, 2.0, 12.5, 4.8, fill_color=COLOR_WHITE,
         line_color=RGBColor(0xCE, 0x93, 0xD8), line_width=1.5)

dialogues = [
    ("カウンセラー", "「ファスティングで一番うまくいかないのはどんなときですか？」",
     COLOR_TEAL, False),
    ("クライアント", "「夜、仕事から帰ってくると、つい食べてしまうんです…」",
     COLOR_GRAY, False),
    ("カウンセラー", "「その時間、どんな気持ちになっていますか？」",
     COLOR_TEAL, False),
    ("クライアント", "「…疲れていて、ご褒美が欲しくなっている気がします」",
     COLOR_GRAY, False),
    ("カウンセラー", "「なるほど。疲れているときに何か別のご褒美があったら、変わると思いますか？」",
     COLOR_TEAL, False),
    ("クライアント", "「あ……そうか。食べること以外で自分を癒す方法、考えたことなかったです」",
     RGBColor(0x6A, 0x1B, 0x9A), True),
]

for i, (speaker, text, color, bold) in enumerate(dialogues):
    y = 2.1 + i * 0.73
    label_color = COLOR_TEAL if speaker == "カウンセラー" else COLOR_GRAY
    add_text(slide, f"【{speaker}】", 0.6, y, 2.2, 0.55,
             font_size=11, bold=True, color=label_color)
    add_text(slide, text, 2.8, y, 9.8, 0.6,
             font_size=13, bold=bold, color=color)

add_rect(slide, 0.4, 6.55, 12.5, 0.7, fill_color=RGBColor(0xF3, 0xE5, 0xF5))
add_text(slide, "💡 気づきは「教えられるもの」ではなく「問いかけによって自分の中から生まれるもの」",
         0.6, 6.58, 12.0, 0.65,
         font_size=14, bold=True, color=RGBColor(0x6A, 0x1B, 0x9A), align=PP_ALIGN.CENTER)


# ============================================================
# スライド 7: なりたい自分を見つける
# ============================================================
slide = prs.slides.add_slide(BLANK_LAYOUT)
add_rect(slide, 0, 0, 13.33, 7.5, fill_color=RGBColor(0xE8, 0xF5, 0xE9))
add_rect(slide, 0, 0, 13.33, 1.1, fill_color=COLOR_PRIMARY)
add_text(slide, "04｜なりたい自分を見つける ── 目標設定のプロセス", 0.4, 0.15, 12.5, 0.8,
         font_size=25, bold=True, color=COLOR_WHITE)

# ステップ図
steps = [
    ("STEP 1", "現在地を知る", "今、どんな状態にいるか？\n何に困っているか？\n何を感じているか？", RGBColor(0xE5, 0x73, 0x73)),
    ("STEP 2", "痛みを掘り下げる", "その悩みの奥にある\n本当の感情は何か？\nいつからそう感じている？", RGBColor(0xFF, 0xB7, 0x4D)),
    ("STEP 3", "理想を描く", "どんな自分になりたいか？\nそれが叶ったらどんな気持ち？\n大切にしたい価値観は？", RGBColor(0x4D, 0xB6, 0xAC)),
    ("STEP 4", "ギャップを言語化する", "現在地と理想の差は何か？\n何があれば埋まるか？\n何が邪魔しているか？", COLOR_PRIMARY),
]

for i, (step, title, desc, color) in enumerate(steps):
    x = 0.5 + i * 3.2
    add_rect(slide, x, 1.3, 3.0, 5.5, fill_color=COLOR_WHITE,
             line_color=color, line_width=2)
    add_rect(slide, x, 1.3, 3.0, 0.8, fill_color=color)
    add_text(slide, step, x + 0.05, 1.33, 2.9, 0.45,
             font_size=13, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)
    add_text(slide, title, x + 0.1, 2.2, 2.8, 0.7,
             font_size=16, bold=True, color=color, align=PP_ALIGN.CENTER)
    add_rect(slide, x + 0.2, 3.0, 2.6, 0.04, fill_color=color)
    add_text(slide, desc, x + 0.15, 3.15, 2.7, 3.5,
             font_size=13, color=COLOR_DARK, align=PP_ALIGN.CENTER)

    if i < 3:
        add_text(slide, "→", 3.35 + i * 3.2, 3.7, 0.5, 0.5,
                 font_size=24, bold=True, color=COLOR_GRAY, align=PP_ALIGN.CENTER)

add_rect(slide, 0.5, 6.95, 12.5, 0.45, fill_color=RGBColor(0xC8, 0xE6, 0xC9))
add_text(slide, "カウンセリングはクライアント自身が「なりたい自分」を言語化できるよう伴走するプロセス",
         0.7, 6.97, 12.0, 0.42,
         font_size=13, bold=True, color=COLOR_PRIMARY, align=PP_ALIGN.CENTER)


# ============================================================
# スライド 8: 一人ではできない ── 環境と知識の重要性
# ============================================================
slide = prs.slides.add_slide(BLANK_LAYOUT)
add_rect(slide, 0, 0, 13.33, 7.5, fill_color=RGBColor(0xFF, 0xF3, 0xE0))
add_rect(slide, 0, 0, 13.33, 1.1, fill_color=COLOR_WARM)
add_text(slide, "05｜一人ではできない ── 必要な環境と知識", 0.4, 0.15, 12.5, 0.8,
         font_size=26, bold=True, color=COLOR_WHITE)

# 中心メッセージ
add_rect(slide, 1.0, 1.25, 11.3, 1.2, fill_color=RGBColor(0xFF, 0xE0, 0xB2))
add_text(slide, "気づきは一歩目。でも、気づいただけでは変われない。\n「変わる」ためには、適切な環境と知識のサポートが必要。",
         1.2, 1.28, 11.0, 1.15,
         font_size=17, bold=True, color=RGBColor(0xE6, 0x51, 0x00), align=PP_ALIGN.CENTER)

# 三角形的な説明（三要素）
elements = [
    ("🌿", "環境づくり",
     ["仲間・コミュニティの存在", "記録できる仕組み", "定期的な振り返りの場", "家族や周囲の理解"]),
    ("📚", "知識のサポート",
     ["ファスティングの正しい方法", "体の仕組みと食の関係", "リバウンドを防ぐ考え方", "メンタルと食欲の関連"]),
    ("👥", "伴走するカウンセラー",
     ["定期的な対話と確認", "モチベーションの維持", "躓いたときの軌道修正", "次の目標の設定"]),
]

for i, (icon, title, items_list) in enumerate(elements):
    x = 0.5 + i * 4.2
    add_rect(slide, x, 2.65, 3.9, 4.5, fill_color=COLOR_WHITE,
             line_color=RGBColor(0xFF, 0xCC, 0x02), line_width=2)
    add_rect(slide, x, 2.65, 3.9, 1.0, fill_color=RGBColor(0xFF, 0xB3, 0x00))
    add_text(slide, f"{icon} {title}", x + 0.1, 2.72, 3.7, 0.55,
             font_size=17, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)
    for j, item in enumerate(items_list):
        add_text(slide, f"• {item}", x + 0.2, 3.75 + j * 0.78, 3.5, 0.65,
                 font_size=13, color=COLOR_DARK)


# ============================================================
# スライド 9: カウンセリングは最大の教育
# ============================================================
slide = prs.slides.add_slide(BLANK_LAYOUT)
add_rect(slide, 0, 0, 13.33, 7.5, fill_color=RGBColor(0xE8, 0xEA, 0xF6))
add_rect(slide, 0, 0, 13.33, 1.1, fill_color=RGBColor(0x28, 0x35, 0x93))
add_text(slide, "カウンセリングは「最大の教育の場」", 0.4, 0.15, 12.5, 0.8,
         font_size=28, bold=True, color=COLOR_WHITE)

# 左：教育との違い
add_rect(slide, 0.4, 1.3, 5.8, 5.6, fill_color=COLOR_WHITE,
         line_color=RGBColor(0x9F, 0xA8, 0xDA), line_width=1.5)
add_rect(slide, 0.4, 1.3, 5.8, 0.7, fill_color=RGBColor(0x5C, 0x6B, 0xC0))
add_text(slide, "一般的な「教える」教育", 0.6, 1.37, 5.4, 0.55,
         font_size=17, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)

teach_items = [
    "知識を一方向に伝える",
    "「正解」を外から与える",
    "受け手が受動的になりやすい",
    "「わかった」だけで終わりがち",
    "行動変容に繋がりにくい",
]
for j, item in enumerate(teach_items):
    add_text(slide, f"• {item}", 0.7, 2.2 + j * 0.85, 5.3, 0.7,
             font_size=15, color=RGBColor(0x28, 0x35, 0x93))

# 右：カウンセリング教育
add_rect(slide, 7.0, 1.3, 5.8, 5.6, fill_color=COLOR_WHITE,
         line_color=RGBColor(0xA5, 0xD6, 0xA7), line_width=1.5)
add_rect(slide, 7.0, 1.3, 5.8, 0.7, fill_color=COLOR_PRIMARY)
add_text(slide, "カウンセリングによる教育", 7.2, 1.37, 5.4, 0.55,
         font_size=17, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)

counsel_items = [
    "問いかけで自ら考えさせる",
    "自分の中から「答え」を引き出す",
    "当事者意識・主体性が生まれる",
    "「気づき」は深く記憶に残る",
    "行動変容に直結する",
]
for j, item in enumerate(counsel_items):
    add_text(slide, f"✓ {item}", 7.2, 2.2 + j * 0.85, 5.4, 0.7,
             font_size=15, bold=True, color=COLOR_PRIMARY)

# 中央
add_text(slide, "VS", 6.2, 3.7, 0.9, 0.7,
         font_size=24, bold=True, color=COLOR_GRAY, align=PP_ALIGN.CENTER)

add_rect(slide, 0.4, 6.95, 12.5, 0.45, fill_color=RGBColor(0x3F, 0x51, 0xB5))
add_text(slide, "「知識を教えてもらった」より「自分で気づいた」ことのほうが、人は動く",
         0.6, 6.97, 12.0, 0.42,
         font_size=14, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)


# ============================================================
# スライド 10: カウンセリングの実践フロー
# ============================================================
slide = prs.slides.add_slide(BLANK_LAYOUT)
add_rect(slide, 0, 0, 13.33, 7.5, fill_color=COLOR_LIGHT_BG)
add_rect(slide, 0, 0, 13.33, 1.1, fill_color=COLOR_PRIMARY)
add_text(slide, "06｜実践：カウンセリングの流れ", 0.4, 0.15, 12.5, 0.8,
         font_size=28, bold=True, color=COLOR_WHITE)

flow_steps = [
    ("🤝 ラポール形成", "信頼関係を築く\n「話しやすい場」をつくる", "0〜5分"),
    ("🔍 現状ヒアリング", "オープン質問で現状を聞く\n批判せず、ありのままを受容する", "5〜15分"),
    ("💡 悩みの深掘り", "「なぜ？」「どんな気持ち？」\nで本質に迫る", "15〜25分"),
    ("🌟 理想の明確化", "「どうなりたいか？」を引き出し\n言語化・視覚化する", "25〜35分"),
    ("🗺️ 行動計画の共創", "一緒に「次の一歩」を決める\n環境・知識のサポートを提案", "35〜45分"),
]

for i, (title, desc, time) in enumerate(flow_steps):
    x = 0.4 + i * 2.55
    # コネクター
    if i < 4:
        add_rect(slide, x + 2.35, 3.55, 0.25, 0.35, fill_color=COLOR_ACCENT)

    add_rect(slide, x, 1.3, 2.35, 5.6, fill_color=COLOR_WHITE,
             line_color=COLOR_ACCENT, line_width=1.5)
    add_rect(slide, x, 1.3, 2.35, 0.55, fill_color=COLOR_PRIMARY)
    add_text(slide, time, x + 0.05, 1.33, 2.25, 0.48,
             font_size=11, color=COLOR_ACCENT, align=PP_ALIGN.CENTER)

    add_rect(slide, x, 1.85, 2.35, 0.9,
             fill_color=RGBColor(0xE8, 0xF5, 0xE9))
    add_text(slide, title, x + 0.05, 1.88, 2.25, 0.85,
             font_size=12, bold=True, color=COLOR_PRIMARY, align=PP_ALIGN.CENTER)

    add_text(slide, desc, x + 0.1, 2.9, 2.15, 3.8,
             font_size=12, color=COLOR_DARK, align=PP_ALIGN.CENTER)

add_rect(slide, 0.4, 7.05, 12.5, 0.38, fill_color=RGBColor(0xC8, 0xE6, 0xC9))
add_text(slide, "⚡ ポイント：カウンセラーは「答えを持っている人」ではなく「正しい問いを持っている人」",
         0.6, 7.07, 12.0, 0.35,
         font_size=13, bold=True, color=COLOR_PRIMARY, align=PP_ALIGN.CENTER)


# ============================================================
# スライド 11: 問いかけの技術
# ============================================================
slide = prs.slides.add_slide(BLANK_LAYOUT)
add_rect(slide, 0, 0, 13.33, 7.5, fill_color=RGBColor(0xE0, 0xF7, 0xFA))
add_rect(slide, 0, 0, 13.33, 1.1, fill_color=COLOR_TEAL)
add_text(slide, "問いかけの技術 ── 「聴く力」と「問う力」", 0.4, 0.15, 12.5, 0.8,
         font_size=26, bold=True, color=COLOR_WHITE)

# 左：避けるべき問い
add_rect(slide, 0.4, 1.3, 5.8, 5.6, fill_color=COLOR_WHITE,
         line_color=RGBColor(0xEF, 0x9A, 0x9A), line_width=1.5)
add_rect(slide, 0.4, 1.3, 5.8, 0.65, fill_color=RGBColor(0xC6, 0x28, 0x28))
add_text(slide, "❌ 避けるべき問いかけ", 0.6, 1.35, 5.4, 0.55,
         font_size=16, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)

bad_questions = [
    ("「ちゃんと食事制限しましたか？」", "→ 詰問。相手が萎縮する"),
    ("「それはこうすればいいですよ」", "→ 答えを与えすぎ。主体性が失われる"),
    ("「なんでできないんですか？」", "→ 責める印象。関係が壊れる"),
    ("「〇〇すべきですよね」", "→ 押しつけ。反発を生みやすい"),
]
for j, (q, note) in enumerate(bad_questions):
    y = 2.1 + j * 1.18
    add_text(slide, q, 0.6, y, 5.3, 0.5,
             font_size=14, bold=True, color=RGBColor(0xC6, 0x28, 0x28))
    add_text(slide, note, 0.6, y + 0.5, 5.3, 0.5,
             font_size=12, color=COLOR_GRAY)

# 右：効果的な問い
add_rect(slide, 7.0, 1.3, 5.8, 5.6, fill_color=COLOR_WHITE,
         line_color=RGBColor(0x80, 0xCB, 0xC4), line_width=1.5)
add_rect(slide, 7.0, 1.3, 5.8, 0.65, fill_color=COLOR_TEAL)
add_text(slide, "✅ 効果的な問いかけ", 7.2, 1.35, 5.4, 0.55,
         font_size=16, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)

good_questions = [
    ("「今、どんなことを感じていますか？」", "→ 感情に寄り添うオープン質問"),
    ("「うまくいったとしたら、\n どんな違いがありますか？」", "→ 理想の未来を描かせる"),
    ("「それをやってみてどうでしたか？」", "→ 体験を通じた気づきを促す"),
    ("「一番大切にしたいことは何ですか？」", "→ 価値観に触れる深い問い"),
]
for j, (q, note) in enumerate(good_questions):
    y = 2.1 + j * 1.18
    add_text(slide, q, 7.2, y, 5.5, 0.6,
             font_size=14, bold=True, color=COLOR_TEAL)
    add_text(slide, note, 7.2, y + 0.62, 5.5, 0.45,
             font_size=12, color=COLOR_GRAY)

add_text(slide, "VS", 6.1, 3.7, 1.1, 0.7,
         font_size=24, bold=True, color=COLOR_GRAY, align=PP_ALIGN.CENTER)


# ============================================================
# スライド 12: 資格取得者へのメッセージ
# ============================================================
slide = prs.slides.add_slide(BLANK_LAYOUT)
add_rect(slide, 0, 0, 13.33, 7.5, fill_color=RGBColor(0x1B, 0x5E, 0x20))

# 装飾ライン
add_rect(slide, 0, 0, 0.12, 7.5, fill_color=COLOR_WARM)
add_rect(slide, 13.21, 0, 0.12, 7.5, fill_color=COLOR_WARM)

add_text(slide, "資格を取ったあなたへ", 0.5, 0.6, 12.3, 1.0,
         font_size=34, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)
add_rect(slide, 3.0, 1.7, 7.3, 0.06, fill_color=COLOR_WARM)

messages = [
    "「全部答えてあげなきゃ」と思わなくていい。",
    "「全部伝えてあげなきゃ」と思わなくていい。",
    "あなたの役割は、答えを渡すことではなく、",
    "クライアントが自分の答えに気づく手助けをすること。",
]
for i, msg in enumerate(messages):
    add_text(slide, msg, 0.5, 2.0 + i * 0.95, 12.3, 0.85,
             font_size=20, bold=(i >= 2), color=COLOR_WHITE, align=PP_ALIGN.CENTER)

add_rect(slide, 1.5, 5.9, 10.3, 1.2, fill_color=COLOR_WARM)
add_text(slide, "悩みに気づかせる場をつくれるカウンセラーが、\n最大の教育者になれる。",
         1.7, 5.92, 9.9, 1.15,
         font_size=20, bold=True, color=COLOR_DARK, align=PP_ALIGN.CENTER)


# ============================================================
# スライド 13: まとめ
# ============================================================
slide = prs.slides.add_slide(BLANK_LAYOUT)
add_rect(slide, 0, 0, 13.33, 7.5, fill_color=COLOR_LIGHT_BG)
add_rect(slide, 0, 0, 13.33, 1.1, fill_color=COLOR_PRIMARY)
add_text(slide, "まとめ ── カウンセリングの本質", 0.4, 0.15, 12.5, 0.8,
         font_size=28, bold=True, color=COLOR_WHITE)

summary_points = [
    ("🔑", "カウンセリングは「答えを教える場」ではない",
     "自分に足りないものに気づかせる、気づきの場である"),
    ("🌱", "「なりたい自分」を見つけるプロセスを支える",
     "クライアントが自分の言葉で理想を語れるよう伴走する"),
    ("👥", "一人では変われない、だから一緒に",
     "気づきの後には、環境・知識・伴走者が必要である"),
    ("📚", "カウンセリングは最大の教育の場",
     "「自分で気づいた」ことは深く記憶に残り、行動変容に直結する"),
    ("💬", "正しい問いを持つことがカウンセラーの技術",
     "答えではなく、問いかけによってクライアントの力を引き出す"),
]

for i, (icon, title, desc) in enumerate(summary_points):
    y = 1.3 + i * 1.18
    add_rect(slide, 0.4, y, 12.5, 1.05, fill_color=COLOR_WHITE,
             line_color=COLOR_ACCENT, line_width=1)
    add_text(slide, icon, 0.5, y + 0.15, 0.7, 0.7, font_size=24, align=PP_ALIGN.CENTER)
    add_text(slide, title, 1.3, y + 0.08, 11.3, 0.5,
             font_size=16, bold=True, color=COLOR_PRIMARY)
    add_text(slide, desc, 1.3, y + 0.58, 11.3, 0.42,
             font_size=13, color=COLOR_DARK)

add_rect(slide, 0.4, 7.15, 12.5, 0.28, fill_color=COLOR_PRIMARY)
add_text(slide, "悩みに気づかせ、なりたい自分を見つける場 ── それがカウンセリングの真髄",
         0.5, 7.16, 12.2, 0.26,
         font_size=12, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)


output_path = "/home/user/kiri-gohan/ファスティング指導士_カウンセリング講義スライド.pptx"
prs.save(output_path)
print(f"Saved: {output_path}")
print(f"Total slides: {len(prs.slides)}")
