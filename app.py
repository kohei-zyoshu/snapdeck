"""
パシャッと — スマホ・高齢者対応ユニバーサルデザイン版
ホワイトボード・手書きメモ → スライド自動変換
"""

import os, io, json, base64, re
from datetime import datetime

import streamlit as st
from PIL import Image, ImageChops, ImageEnhance, ImageFilter

# ─── ページ設定 ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="パシャッと",
    page_icon="📸",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ─── CSS（高齢者・スマホ ユニバーサルデザイン） ─────────────────────────────
st.markdown("""
<style>
/* ── 全体背景 ── */
.stApp { background: #F4F7FA; }
.main .block-container { padding: 1.4rem 1rem 5rem; max-width: 560px; }

/* ── ロゴ・タイトル ── */
.logo {
    text-align: center;
    font-size: 3rem;
    font-weight: 900;
    color: #0D1B2A;
    margin-bottom: 0.1rem;
    line-height: 1.2;
    letter-spacing: -0.01em;
}
.logo em { color: #0099BB; font-style: normal; }
.tagline {
    text-align: center;
    color: #4A6475;
    font-size: 1.15rem;
    margin-bottom: 0.6rem;
    line-height: 1.6;
    font-weight: 500;
}
.sub-tagline {
    text-align: center;
    color: #7A9AAD;
    font-size: 0.95rem;
    margin-bottom: 2rem;
    line-height: 1.5;
}

/* ── ステップラベル ── */
.step {
    display: flex;
    align-items: center;
    gap: 0.75rem;
    font-size: 1.45rem;
    font-weight: 800;
    color: #0D1B2A;
    margin: 2rem 0 0.6rem;
}
.snum {
    background: #0099BB;
    color: #fff;
    border-radius: 50%;
    width: 42px;
    height: 42px;
    line-height: 42px;
    text-align: center;
    font-size: 1.15rem;
    font-weight: 900;
    flex-shrink: 0;
}

/* ── ヒント・説明文 ── */
.hint {
    color: #3D5A6A;
    font-size: 1.1rem;
    margin: 0 0 1rem;
    line-height: 1.8;
    background: #EDF4F8;
    border-left: 4px solid #0099BB;
    border-radius: 0 8px 8px 0;
    padding: 0.75rem 1rem;
}

/* ── 案内バナー（注意・補足） ── */
.guide-box {
    background: #FFF8E7;
    border: 1.5px solid #F5C842;
    border-radius: 10px;
    padding: 0.85rem 1rem;
    font-size: 1.05rem;
    color: #5C4800;
    line-height: 1.7;
    margin-bottom: 1rem;
}

/* ── ボタン全般（大きく・見やすく） ── */
.stButton > button {
    font-size: 1.3rem !important;
    min-height: 72px !important;
    border-radius: 16px !important;
    font-weight: 800 !important;
    width: 100% !important;
    letter-spacing: 0.03em !important;
    transition: opacity 0.15s !important;
}
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #0099BB 0%, #007A99 100%) !important;
    color: white !important;
    border: none !important;
    box-shadow: 0 5px 16px rgba(0,153,187,0.35) !important;
}
.stButton > button[kind="primary"]:hover {
    opacity: 0.88 !important;
}
.stButton > button[kind="secondary"] {
    background: #EBF3F8 !important;
    color: #1C3A4A !important;
    border: 2px solid #B0CCD8 !important;
}
.stButton > button:disabled {
    opacity: 0.38 !important;
    cursor: not-allowed !important;
}

/* ── ダウンロードボタン ── */
.stDownloadButton > button {
    background: linear-gradient(135deg, #10B981 0%, #059669 100%) !important;
    color: white !important;
    font-size: 1.3rem !important;
    min-height: 72px !important;
    border-radius: 16px !important;
    font-weight: 800 !important;
    width: 100% !important;
    border: none !important;
    box-shadow: 0 5px 16px rgba(16,185,129,0.35) !important;
    letter-spacing: 0.03em !important;
}

/* ── 完了ボックス ── */
.done-box {
    background: #D1FAE5;
    border: 2.5px solid #10B981;
    border-radius: 16px;
    padding: 1.4rem 1.6rem;
    font-size: 1.2rem;
    color: #064E3B;
    font-weight: 700;
    margin: 1.2rem 0;
    line-height: 1.9;
}

/* ── タブ ── */
.stTabs [data-baseweb="tab"] {
    font-size: 1.15rem !important;
    font-weight: 700 !important;
    min-height: 56px !important;
}

/* ── ファイルアップローダー ── */
[data-testid="stFileUploaderDropzone"] {
    min-height: 150px;
    border-radius: 14px !important;
    border: 2.5px dashed #90B8C8 !important;
    background: #F0F8FB !important;
}
/* 英語の案内テキストを非表示にして日本語に置き換え */
[data-testid="stFileUploaderDropzoneInstructions"] span {
    display: none !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] small {
    display: none !important;
}
[data-testid="stFileUploaderDropzoneInstructions"]::before {
    content: "ここに画像をドラッグするか、下のボタンでファイルを選んでください";
    display: block;
    font-size: 1.05rem;
    color: #3D6070;
    margin-bottom: 0.6rem;
    text-align: center;
}
[data-testid="stFileUploaderDropzoneInstructions"]::after {
    content: "対応形式：JPG・PNG・WebP　最大20MBまで";
    display: block;
    font-size: 0.9rem;
    color: #7A9AAD;
    text-align: center;
    margin-top: 0.3rem;
}
/* 「Browse files」ボタンを「ファイルを選ぶ」に */
[data-testid="stFileUploaderDropzone"] button {
    font-size: 0 !important;
    min-height: 52px !important;
    border-radius: 10px !important;
    padding: 0 1.5rem !important;
}
[data-testid="stFileUploaderDropzone"] button::after {
    content: "📷　写真を撮る・選ぶ";
    font-size: 1.05rem;
    font-weight: 700;
}

/* ── トグル ── */
.stToggle label { font-size: 1.1rem !important; font-weight: 700 !important; }

/* ── expander ── */
.stExpander summary p { font-size: 1.1rem !important; font-weight: 600 !important; }

/* ── 区切り線 ── */
hr { margin: 1.6rem 0 !important; border-color: #C8D8E4 !important; border-width: 1.5px !important; }

/* ── 警告・エラー ── */
.stAlert p { font-size: 1.05rem !important; line-height: 1.7 !important; }

/* ── スピナー文字 ── */
.stSpinner p { font-size: 1.1rem !important; }
</style>
""", unsafe_allow_html=True)


# ─── バックエンド関数 ──────────────────────────────────────────────────────────

def get_api_key() -> str:
    """APIキーを Secrets または入力欄から取得（前後の空白・改行を除去）"""
    try:
        key = st.secrets["ANTHROPIC_API_KEY"]
        if key:
            return str(key).strip()
    except Exception:
        pass
    env_key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    if env_key:
        return env_key
    return st.session_state.get("api_key_input", "").strip()


def auto_trim(img: Image.Image, margin: int = 30) -> Image.Image:
    """白い余白を自動でカット（読み取り精度向上）"""
    if img.mode != "RGB":
        img = img.convert("RGB")
    bg   = Image.new("RGB", img.size, (255, 255, 255))
    diff = ImageChops.difference(img, bg)
    bbox = diff.getbbox()
    if not bbox:
        return img
    left   = max(0,          bbox[0] - margin)
    top    = max(0,          bbox[1] - margin)
    right  = min(img.width,  bbox[2] + margin)
    bottom = min(img.height, bbox[3] + margin)
    return img.crop((left, top, right, bottom))


def enhance_for_ocr(img: Image.Image) -> Image.Image:
    """OCR精度を上げるための画像前処理（コントラスト・シャープネス強化）"""
    # コントラスト強化
    img = ImageEnhance.Contrast(img).enhance(1.5)
    # シャープネス強化
    img = ImageEnhance.Sharpness(img).enhance(2.0)
    # 明るさを微調整（暗すぎる写真を補正）
    img = ImageEnhance.Brightness(img).enhance(1.1)
    return img


def pil_to_base64(img: Image.Image, max_px: int = 2048) -> tuple[str, str]:
    """PIL Image → base64（高解像度維持・OCR前処理済み）"""
    # OCR精度向上のための前処理
    img = enhance_for_ocr(img)
    # 大きすぎる画像は縮小（API上限対策）
    if img.width > max_px or img.height > max_px:
        img.thumbnail((max_px, max_px), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=92)
    return base64.standard_b64encode(buf.getvalue()).decode("utf-8"), "image/jpeg"


EXTRACTION_PROMPT = """
あなたはOCRと文書構造解析の専門家です。
添付の画像（ホワイトボード・手書きメモ・印刷物など）を注意深く分析し、
すべての文字・記号・数字を漏れなく読み取ってください。

【重要な読み取りルール】
- かすれた文字、薄い文字も最大限推測して読み取ること
- 矢印（→ ← ↑ ↓）、丸囲み、下線なども構造の手がかりとして活用すること
- 誤字と思われる文字もそのまま記録すること
- 日本語・英語・数字・記号を問わず、すべて抽出すること
- 画像の端や隅に小さく書かれた文字も見落とさないこと

以下のJSON形式のみで返答してください（説明文・前置き・コードブロック不要）:

{
  "title": "メモ全体のタイトルまたは主題（画像から推定）",
  "elements": [
    {
      "id": "el_001",
      "type": "heading",
      "content": "読み取ったテキスト",
      "level": 1,
      "confidence": 0.95
    }
  ],
  "structure": {
    "type": "list",
    "groups": [
      {
        "label": "グループ名",
        "items": ["el_001", "el_002"]
      }
    ]
  },
  "language": "ja",
  "notes": "読み取れなかった箇所や補足があれば記載"
}

typeは heading（大見出し）/ bullet（箇条書き）/ text（通常テキスト）/ table_row（表の行）のいずれか。
"""


@st.cache_data(show_spinner=False)
def analyze_with_claude(img_data: str, media_type: str, api_key: str) -> dict:
    """AIで画像を解析（同じ画像は再利用）"""
    import anthropic
    client = anthropic.Anthropic(api_key=api_key)
    message = client.messages.create(
        model="claude-opus-4-6",
        max_tokens=4096,
        messages=[{"role": "user", "content": [
            {"type": "image", "source": {
                "type": "base64", "media_type": media_type, "data": img_data
            }},
            {"type": "text", "text": EXTRACTION_PROMPT}
        ]}]
    )
    text = message.content[0].text.strip()
    # ```json ... ``` や ``` ... ``` のコードブロックを除去
    text = re.sub(r'^```(?:json)?\s*', '', text, flags=re.MULTILINE)
    text = re.sub(r'\s*```$',          '', text, flags=re.MULTILINE)
    text = text.strip()
    match = re.search(r'\{.*\}', text, re.DOTALL)
    raw = match.group() if match else text
    return json.loads(raw)


def get_demo_data() -> dict:
    return {
        "title": "新サービス企画 ブレインストーミング",
        "elements": [
            {"id": "el_001", "type": "heading", "content": "対象のお客様",                  "level": 1, "confidence": 0.97},
            {"id": "el_002", "type": "bullet",  "content": "20〜40代のビジネスパーソン",     "level": 2, "confidence": 0.95},
            {"id": "el_003", "type": "bullet",  "content": "会議が多い職種（営業・企画・PM）","level": 2, "confidence": 0.93},
            {"id": "el_004", "type": "heading", "content": "主な機能",                       "level": 1, "confidence": 0.98},
            {"id": "el_005", "type": "bullet",  "content": "写真1枚でスライド自動生成",      "level": 2, "confidence": 0.97},
            {"id": "el_006", "type": "bullet",  "content": "日本語・英語など多言語に対応",   "level": 2, "confidence": 0.94},
            {"id": "el_007", "type": "bullet",  "content": "パワーポイント形式でダウンロード","level": 2, "confidence": 0.96},
            {"id": "el_008", "type": "heading", "content": "料金プラン案",                   "level": 1, "confidence": 0.95},
            {"id": "el_009", "type": "bullet",  "content": "無料プラン（月10枚まで）",       "level": 2, "confidence": 0.93},
            {"id": "el_010", "type": "bullet",  "content": "有料プラン：月額1,980円",        "level": 2, "confidence": 0.91},
        ],
        "structure": {
            "type": "list",
            "groups": [
                {"label": "対象のお客様", "items": ["el_001", "el_002", "el_003"]},
                {"label": "主な機能",     "items": ["el_004", "el_005", "el_006", "el_007"]},
                {"label": "料金プラン案", "items": ["el_008", "el_009", "el_010"]},
            ]
        },
        "language": "ja",
        "notes": "これは動作確認用のサンプルデータです。実際の写真で試すと、その内容が読み取られます。"
    }


def generate_pptx(data: dict) -> bytes:
    """解析データからパワーポイントを生成"""
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN

    C_DARK   = RGBColor(0x0D, 0x1B, 0x2A)
    C_NAVY   = RGBColor(0x1B, 0x28, 0x38)
    C_ACCENT = RGBColor(0x00, 0xB4, 0xD8)
    C_WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
    C_LIGHT  = RGBColor(0xCB, 0xD5, 0xE1)
    C_MUTED  = RGBColor(0x94, 0xA3, 0xB8)

    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    def bg(slide, color):
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = color

    def rect(slide, x, y, w, h, fc, lc=None):
        s = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
        s.fill.solid(); s.fill.fore_color.rgb = fc
        if lc: s.line.color.rgb = lc
        else:  s.line.fill.background()

    def txt(slide, text, x, y, w, h, size=14, bold=False, color=None,
            align=PP_ALIGN.LEFT, wrap=True):
        color = color or C_WHITE
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame; tf.word_wrap = wrap
        p = tf.paragraphs[0]; p.alignment = align
        r = p.add_run(); r.text = text
        r.font.size = Pt(size); r.font.bold = bold
        r.font.color.rgb = color; r.font.name = "Arial"

    # ── 1枚のスライドに全内容を集約 ──
    s = prs.slides.add_slide(blank)
    bg(s, C_DARK)

    # 上部アクセントライン
    rect(s, 0, 0, 13.33, 0.07, C_ACCENT)

    # タイトル
    title = data.get("title", "ホワイトボードメモ")
    txt(s, title, 0.4, 0.12, 12.5, 0.75, size=26, bold=True)

    # 日時ラベル
    ts = datetime.now().strftime("%Y年%m月%d日 %H:%M")
    txt(s, f"変換日時：{ts}", 0.4, 0.82, 8.0, 0.35, size=10, color=C_MUTED)

    # コンテンツ背景
    rect(s, 0.3, 1.2, 12.73, 5.65, C_NAVY)

    # 全要素を左右2カラムに振り分けて配置
    elements = data.get("elements", [])
    elements_by_id = {el["id"]: el for el in elements}

    # グループを使わず全要素を順番に並べる
    all_items = []
    groups = data.get("structure", {}).get("groups", [])
    if groups:
        for group in groups:
            lbl = group.get("label", "")
            if lbl:
                all_items.append({"type": "heading", "content": lbl})
            for iid in group.get("items", []):
                el = elements_by_id.get(iid)
                if el:
                    all_items.append(el)
    else:
        all_items = elements

    # 2カラム配置
    mid = max(1, len(all_items) // 2)
    col_left  = all_items[:mid]
    col_right = all_items[mid:]

    def render_col(items, x_start, col_w):
        y = 1.3
        for el in items:
            if el.get("type") == "heading":
                if y > 1.3: y += 0.05
                txt(s, el.get("content", ""), x_start, y, col_w, 0.45,
                    size=15, bold=True, color=C_ACCENT)
                y += 0.5
            else:
                content = "• " + el.get("content", "")
                txt(s, content, x_start + 0.1, y, col_w - 0.1, 0.4,
                    size=12, color=C_LIGHT)
                y += 0.43
            if y > 6.65: break

    render_col(col_left,  0.45, 6.05)
    render_col(col_right, 6.7,  6.05)

    # フッター
    rect(s, 0, 7.15, 13.33, 0.35, C_NAVY)
    txt(s, "パシャッと  |  自動生成", 0.4, 7.15, 12.9, 0.35,
        size=10, color=C_MUTED, align=PP_ALIGN.RIGHT)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ─── 画面表示 ────────────────────────────────────────────
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

# ── ロゴ・説明 ──
st.markdown("""
<div class='logo'>パシャッ<em>と</em></div>
<div class='tagline'>写真を撮るだけで、スライドに変わります</div>
<div class='sub-tagline'>ホワイトボードや手書きメモを撮影 → AIが文字を読み取り → パワーポイントで保存</div>
""", unsafe_allow_html=True)

# ── APIキー入力（未設定時のみサイドバーに表示） ──
if not get_api_key():
    with st.sidebar:
        st.markdown("### 管理者設定")
        st.markdown("AIを利用するためのキーを入力してください。")
        key_in = st.text_input(
            "認証キー",
            type="password",
            placeholder="sk-ant-...",
            help="管理者から受け取ったキーを入力してください。",
        )
        if key_in:
            st.session_state["api_key_input"] = key_in
            st.success("設定できました")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ステップ１　写真を用意する
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
st.markdown(
    "<div class='step'><span class='snum'>１</span>写真を用意する</div>",
    unsafe_allow_html=True)
st.markdown(
    "<div class='hint'>"
    "ホワイトボードや手書きメモを、<strong>正面からまっすぐ</strong>撮影してください。<br>"
    "文字が大きくはっきり写っていると、より正確に読み取れます。"
    "</div>",
    unsafe_allow_html=True)

st.markdown(
    "<div class='guide-box'>📱 スマートフォンで撮影した写真や、カメラロールにある画像を選べます。</div>",
    unsafe_allow_html=True)
uploaded_file = st.file_uploader(
    "画像ファイルを選ぶ",
    type=["jpg", "jpeg", "png", "webp"],
    label_visibility="collapsed",
    help="JPG・PNG・WebP 形式の画像に対応しています。",
)

raw_input = uploaded_file

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 画像プレビュー ＆ 余白カット
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
pil_image  = None
img_data   = None
media_type = "image/jpeg"

if raw_input:
    pil_image = Image.open(raw_input)
    if pil_image.mode != "RGB":
        pil_image = pil_image.convert("RGB")

    st.markdown("---")

    do_trim = st.toggle(
        "✂️　余白を自動でカットして、文字を読み取りやすくする（おすすめ）",
        value=True,
        help="写真の周囲にある白い余白や余分な部分を自動で除去します。読み取り精度が上がります。",
    )

    display_img = auto_trim(pil_image) if do_trim else pil_image
    st.image(display_img, caption="この写真を変換します", use_container_width=True)

    img_data, media_type = pil_to_base64(display_img)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ステップ２　スライドに変換する
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
st.markdown("---")
st.markdown(
    "<div class='step'><span class='snum'>２</span>スライドに変換する</div>",
    unsafe_allow_html=True)

can_convert = bool(raw_input)

if not can_convert:
    st.markdown(
        "<div class='hint' style='border-color:#F5A623; background:#FFF8E7;'>"
        "↑ まず上の「ステップ１」で写真を撮るか選んでください。"
        "</div>",
        unsafe_allow_html=True)

convert_btn = st.button(
    "✨　この写真をスライドに変換する",
    type="primary",
    use_container_width=True,
    disabled=not can_convert,
)

st.markdown(
    "<p style='text-align:center; color:#7A9AAD; font-size:1.05rem; margin:1rem 0 0.5rem;'>"
    "— または —"
    "</p>",
    unsafe_allow_html=True)

demo_btn = st.button(
    "🎯　写真なしで動きを確認してみる（サンプル）",
    use_container_width=True,
)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 変換処理
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
if demo_btn:
    st.session_state["extracted"]  = get_demo_data()
    st.session_state["pptx_bytes"] = generate_pptx(get_demo_data())
    st.session_state["is_demo"]    = True
    st.rerun()

if convert_btn and img_data:
    st.session_state["is_demo"] = False
    try:
        api_key = get_api_key()
        if not api_key:
            st.warning("認証キーが設定されていません。左上のメニューから設定してください。")
            st.stop()

        status_text = st.empty()
        progress_bar = st.progress(0)

        status_text.markdown("**① 写真を準備しています…**")
        progress_bar.progress(15)

        status_text.markdown("**② AIに写真を送っています…**")
        progress_bar.progress(35)

        extracted = analyze_with_claude(img_data, media_type, api_key)

        status_text.markdown("**③ 文字の読み取りが完了しました ✅**")
        progress_bar.progress(75)

        status_text.markdown("**④ スライドを作成しています…**")
        progress_bar.progress(90)

        st.session_state["extracted"]  = extracted
        st.session_state["pptx_bytes"] = generate_pptx(extracted)

        progress_bar.progress(100)
        status_text.markdown("**✅ 完了しました！下にスクロールして保存してください。**")

    except Exception as e:
            err_msg = str(e)
            if "api_key" in err_msg.lower() or "authentication" in err_msg.lower() or "401" in err_msg:
                st.error("認証キーが正しくありません。左上のメニューから確認してください。")
            elif "JSONDecodeError" in type(e).__name__ or "json" in err_msg.lower():
                st.error("文字の読み取り結果を処理できませんでした。もう一度お試しください。")
            elif "timeout" in err_msg.lower() or "connection" in err_msg.lower():
                st.error("通信エラーが発生しました。インターネット接続をご確認のうえ、もう一度お試しください。")
            elif "overloaded" in err_msg.lower() or "529" in err_msg:
                st.error("AIサーバーが混み合っています。しばらく待ってから、もう一度お試しください。")
            else:
                st.error(f"エラーが発生しました。もう一度お試しください。\n（詳細：{err_msg}）")
            st.info("「写真なしで動きを確認してみる」ボタンを押すと、サンプルでお試しいただけます。")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ステップ３　スライドを保存する
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
if "extracted" in st.session_state:
    extracted  = st.session_state["extracted"]
    pptx_bytes = st.session_state.get("pptx_bytes")
    is_demo    = st.session_state.get("is_demo", False)

    st.markdown("---")
    st.markdown(
        "<div class='step'><span class='snum'>３</span>スライドを保存する</div>",
        unsafe_allow_html=True)

    el_count  = len(extracted.get("elements", []))
    grp_count = len(extracted.get("structure", {}).get("groups", []))
    demo_note = "（サンプルデータ）" if is_demo else ""

    st.markdown(f"""
<div class='done-box'>
完了しました{demo_note}<br>
読み取った項目：<strong>{el_count} 件</strong>　グループ：<strong>{grp_count} 件</strong>
</div>
""", unsafe_allow_html=True)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    if pptx_bytes:
        st.markdown(
            "<div class='hint' style='background:#EBF9F3; border-color:#10B981;'>"
            "⬇️　下のボタンを押すと、パワーポイントファイルが保存されます。"
            "</div>",
            unsafe_allow_html=True)
        st.download_button(
            label="💾　パワーポイントとして保存する",
            data=pptx_bytes,
            file_name=f"パシャッと_{ts}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )

    # 読み取り結果の確認（折りたたみ）
    with st.expander("📋　読み取った文字の内容を確認する"):
        elmap = {el["id"]: el for el in extracted.get("elements", [])}
        for group in extracted.get("structure", {}).get("groups", []):
            st.markdown(f"**{group.get('label', 'グループ')}**")
            for iid in group.get("items", []):
                el = elmap.get(iid)
                if el:
                    icon = "📌" if el.get("type") == "heading" else "・"
                    conf = el.get("confidence", 1.0)
                    st.markdown(
                        f"&nbsp;&nbsp;{icon} {el['content']} "
                        f"<span style='color:#9CA3AF; font-size:0.85rem;'>（確度 {conf*100:.0f}%）</span>",
                        unsafe_allow_html=True)
        st.markdown("---")
        json_str = json.dumps(
            {"バージョン": "1.0",
             "変換日時": datetime.now().isoformat(),
             "データ": extracted},
            ensure_ascii=False, indent=2
        )
        st.download_button(
            label="📄　テキストデータとして保存する",
            data=json_str,
            file_name=f"パシャッと_テキスト_{ts}.json",
            mime="application/json",
            use_container_width=True,
        )

# ── フッター ──
st.markdown("<br>", unsafe_allow_html=True)
st.markdown(
    "<div style='text-align:center; color:#90B0BE; font-size:0.9rem; line-height:2;'>"
    "AI文字認識による自動変換サービス<br>© 2026 パシャッと"
    "</div>",
    unsafe_allow_html=True)
