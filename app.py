"""
パシャッと — スマホ・高齢者対応ユニバーサルデザイン版
ホワイトボード・手書きメモ → スライド自動変換
"""

import os, io, json, base64, re
from datetime import datetime

import streamlit as st
from PIL import Image, ImageChops, ImageEnhance, ImageOps

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
/* ── Streamlit UI を非表示 ── */
header[data-testid="stHeader"] { display: none !important; }
#MainMenu { display: none !important; }
footer { display: none !important; }
[data-testid="stToolbar"] { display: none !important; }
[data-testid="stDecoration"] { display: none !important; }
a[href*="github.com"] { display: none !important; }
a[href*="streamlit.io"] { display: none !important; }
.viewerBadge_container__r5tak { display: none !important; }
.stActionButton { display: none !important; }

/* ── 全体背景 ── */
.stApp { background: #F4F7FA; }
.main .block-container { padding: 0 1rem 5rem; max-width: 560px; }

/* ── 固定ヘッダーバー ── */
.top-bar {
    position: sticky;
    top: 0;
    z-index: 999;
    background: #0D1B2A;
    padding: 0.55rem 1.2rem;
    margin: 0 -1rem 1.6rem;
    display: flex;
    align-items: center;
    gap: 0.6rem;
    box-shadow: 0 2px 10px rgba(0,0,0,0.18);
}
.top-bar-logo {
    font-size: 1.55rem;
    font-weight: 900;
    color: #ffffff;
    letter-spacing: -0.01em;
    line-height: 1;
}
.top-bar-logo em { color: #0099BB; font-style: normal; }
.top-bar-tagline {
    font-size: 0.88rem;
    color: #90B8C8;
    font-weight: 500;
    line-height: 1.3;
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
    width: 42px; height: 42px;
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
.stButton > button[kind="primary"]:hover { opacity: 0.88 !important; }
.stButton > button[kind="secondary"] {
    background: #EBF3F8 !important;
    color: #1C3A4A !important;
    border: 2px solid #B0CCD8 !important;
}
.stButton > button:disabled { opacity: 0.38 !important; cursor: not-allowed !important; }

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

/* ── ファイルアップローダー ── */
[data-testid="stFileUploaderDropzone"] {
    min-height: 150px;
    border-radius: 14px !important;
    border: 2.5px dashed #90B8C8 !important;
    background: #F0F8FB !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] span,
[data-testid="stFileUploaderDropzoneInstructions"] small { display: none !important; }
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

/* ── その他 ── */
.stToggle label { font-size: 1.1rem !important; font-weight: 700 !important; }
.stExpander summary p { font-size: 1.1rem !important; font-weight: 600 !important; }
hr { margin: 1.6rem 0 !important; border-color: #C8D8E4 !important; border-width: 1.5px !important; }
.stAlert p { font-size: 1.05rem !important; line-height: 1.7 !important; }
.stSpinner p { font-size: 1.1rem !important; }
</style>
""", unsafe_allow_html=True)


# ─── バックエンド関数 ──────────────────────────────────────────────────────────

def get_api_key() -> str:
    """APIキーを Secrets / 環境変数 / 入力欄から取得"""
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


def preprocess_image(img: Image.Image, do_trim: bool = True) -> Image.Image:
    """EXIF回転 → 余白カット"""
    try:
        img = ImageOps.exif_transpose(img)
    except Exception:
        pass
    if do_trim:
        if img.mode != "RGB":
            img = img.convert("RGB")
        bg   = Image.new("RGB", img.size, (255, 255, 255))
        diff = ImageChops.difference(img, bg)
        bbox = diff.getbbox()
        if bbox:
            margin = 30
            img = img.crop((
                max(0,          bbox[0] - margin),
                max(0,          bbox[1] - margin),
                min(img.width,  bbox[2] + margin),
                min(img.height, bbox[3] + margin),
            ))
    return img


def pil_to_base64(img: Image.Image, max_px: int = 2048) -> tuple[str, str]:
    """PIL Image → OCR用 base64（コントラスト強化・高解像度）"""
    img = ImageEnhance.Contrast(img).enhance(1.5)
    img = ImageEnhance.Sharpness(img).enhance(2.0)
    img = ImageEnhance.Brightness(img).enhance(1.1)
    if img.width > max_px or img.height > max_px:
        img.thumbnail((max_px, max_px), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=92)
    return base64.standard_b64encode(buf.getvalue()).decode("utf-8"), "image/jpeg"


SYSTEM_PROMPT = (
    "You are a JSON extraction engine. "
    "You must always respond with a single valid JSON object and absolutely nothing else. "
    "Do not include any explanation, markdown formatting, code fences, or commentary. "
    "Start your response with { and end with }."
)

EXTRACTION_PROMPT = """画像（ホワイトボード・手書きメモ）を解析し、以下のJSON形式だけを返してください。説明文・コードブロック・前置きは一切不要です。

{"title":"タイトル","elements":[{"id":"el_001","type":"heading","content":"テキスト","position":{"zone_row":0,"zone_col":0},"style":{"color":"black","bold":false,"underline":false,"circled":false,"size":"large"},"confidence":0.95}],"tables":[{"id":"tbl_001","position":{"zone_row":1,"zone_col":0},"headers":["列1","列2"],"rows":[["値1","値2"]]}],"structure":{"type":"list","groups":[{"label":"グループ名","items":["el_001"]}]},"language":"ja","notes":""}

ルール:
- すべての文字・数字・記号を漏れなく読む（かすれ・小さい文字も含む）
- type: heading / bullet / text / arrow のいずれか
- zone_row: 0=上 1=中 2=下、zone_col: 0=左 1=中 2=右
- color: black/red/blue/green/orange/purple/pink/gray/brown/yellow
- size: large/medium/small
- 表がなければ tables は []
- 必ず { で始まり } で終わる純粋なJSONのみ返すこと"""


def analyze_with_claude(img_data: str, media_type: str, api_key: str) -> dict:
    """AIで画像を解析（prefill方式でJSON確実取得）"""
    import anthropic
    client = anthropic.Anthropic(api_key=api_key)
    message = client.messages.create(
        model="claude-opus-4-6",
        max_tokens=4096,
        system=SYSTEM_PROMPT,
        messages=[
            {"role": "user", "content": [
                {"type": "image", "source": {
                    "type": "base64", "media_type": media_type, "data": img_data
                }},
                {"type": "text", "text": EXTRACTION_PROMPT},
            ]},
            # prefill: { を先置きすることでClaudeが必ずJSONを返す
            {"role": "assistant", "content": "{"},
        ]
    )
    # prefillの { を補って完全なJSONに
    text = "{" + message.content[0].text.strip()

    # 最後の } 以降を切り捨て
    end = text.rfind("}")
    if end == -1:
        raise ValueError(f"AIの応答にJSONが含まれていません。応答内容：{text[:300]}")
    text = text[:end + 1]

    # 制御文字を除去してパース
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)
    try:
        return json.loads(text)
    except json.JSONDecodeError as e:
        raise ValueError(f"JSONの解析に失敗しました: {e}\n応答の先頭: {text[:300]}")


def generate_pptx(data: dict) -> bytes:
    """解析データからパワーポイントを生成"""
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN

    # ─ カラーパレット ─
    C_DARK   = RGBColor(0x0D, 0x1B, 0x2A)
    C_NAVY   = RGBColor(0x1B, 0x28, 0x38)
    C_ACCENT = RGBColor(0x00, 0xB4, 0xD8)
    C_WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
    C_MUTED  = RGBColor(0x94, 0xA3, 0xB8)
    C_LIGHT  = RGBColor(0xCB, 0xD5, 0xE1)

    # ─ 手書き色 → RGBColor ─
    TEXT_COLORS = {
        "black":  RGBColor(0xF0, 0xF4, 0xF8),
        "white":  RGBColor(0xFF, 0xFF, 0xFF),
        "red":    RGBColor(0xFF, 0x6B, 0x6B),
        "blue":   RGBColor(0x64, 0xB5, 0xF6),
        "green":  RGBColor(0x81, 0xC7, 0x84),
        "yellow": RGBColor(0xFF, 0xF1, 0x76),
        "orange": RGBColor(0xFF, 0xB7, 0x4D),
        "purple": RGBColor(0xCE, 0x93, 0xD8),
        "pink":   RGBColor(0xF4, 0x8F, 0xB1),
        "gray":   RGBColor(0xB0, 0xBE, 0xC5),
        "brown":  RGBColor(0xBC, 0xAA, 0xA4),
    }

    def resolve_color(name: str) -> RGBColor:
        return TEXT_COLORS.get((name or "black").lower(), C_LIGHT)

    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)

    def bg(slide, color):
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = color

    def rect(slide, x, y, w, h, fc, lc=None):
        sh = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
        sh.fill.solid()
        sh.fill.fore_color.rgb = fc
        if lc:
            sh.line.color.rgb = lc
        else:
            sh.line.fill.background()

    def txt(slide, text, x, y, w, h, size=13, bold=False, color=None,
            align=PP_ALIGN.LEFT, wrap=True, underline=False, italic=False):
        color = color or C_LIGHT
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = wrap
        p = tf.paragraphs[0]
        p.alignment = align
        r = p.add_run()
        r.text           = text
        r.font.size      = Pt(size)
        r.font.bold      = bold
        r.font.underline = underline
        r.font.italic    = italic
        r.font.color.rgb = color
        r.font.name      = "Noto Sans JP"

    def add_table_shape(slide, tbl_data, x, y, max_w, max_h):
        headers = tbl_data.get("headers", [])
        rows    = tbl_data.get("rows", [])
        if not headers and not rows:
            return
        n_rows = len(rows) + (1 if headers else 0)
        n_cols = max(len(headers), max((len(r) for r in rows), default=1))
        if n_rows == 0 or n_cols == 0:
            return
        col_w = min(max_w / n_cols, 2.5)
        row_h = min(max_h / n_rows, 0.45)
        tbl = slide.shapes.add_table(
            n_rows, n_cols,
            Inches(x), Inches(y),
            Inches(col_w * n_cols), Inches(row_h * n_rows)
        ).table
        if headers:
            for ci, h in enumerate(headers[:n_cols]):
                cell = tbl.cell(0, ci)
                cell.text = str(h)
                cell.fill.solid()
                cell.fill.fore_color.rgb = C_ACCENT
                for para in cell.text_frame.paragraphs:
                    for run in para.runs:
                        run.font.bold      = True
                        run.font.size      = Pt(11)
                        run.font.color.rgb = C_DARK
        for ri, row in enumerate(rows):
            tr = ri + (1 if headers else 0)
            for ci, val in enumerate(row[:n_cols]):
                cell = tbl.cell(tr, ci)
                cell.text = str(val)
                bg_col = RGBColor(0x1E, 0x30, 0x40) if ri % 2 == 0 else C_NAVY
                cell.fill.solid()
                cell.fill.fore_color.rgb = bg_col
                for para in cell.text_frame.paragraphs:
                    for run in para.runs:
                        run.font.size      = Pt(10)
                        run.font.color.rgb = C_LIGHT

    # ── スライド作成 ──
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, C_DARK)
    rect(s, 0, 0, 13.33, 0.07, C_ACCENT)

    title = data.get("title", "ホワイトボードメモ")
    txt(s, title, 0.4, 0.10, 12.0, 0.72, size=24, bold=True, color=C_WHITE)
    ts_label = datetime.now().strftime("%Y年%m月%d日 %H:%M")
    txt(s, f"変換日時：{ts_label}", 0.4, 0.78, 6.0, 0.32, size=9, color=C_MUTED)
    rect(s, 0.3, 1.15, 12.73, 5.7, C_NAVY)

    # ─ 要素を zone 順にソート ─
    elements   = data.get("elements", [])
    sorted_els = sorted(elements, key=lambda e: (
        (e.get("position") or {}).get("zone_row", 1),
        (e.get("position") or {}).get("zone_col", 0),
    ))

    # ─ 左右カラムに振り分け ─
    col_left  = [e for e in sorted_els if (e.get("position") or {}).get("zone_col", 0) <= 1]
    col_right = [e for e in sorted_els if (e.get("position") or {}).get("zone_col", 2) == 2]
    if not col_right and len(col_left) > 3:
        mid       = len(col_left) // 2
        col_right = col_left[mid:]
        col_left  = col_left[:mid]

    # ─ 行高・フォントサイズを要素数から自動計算 ─
    Y_START = 1.25
    Y_MAX   = 6.65
    AVAIL_H = Y_MAX - Y_START
    max_col = max(len(col_left), len(col_right), 1)
    step_h  = max(0.28, min(0.52, AVAIL_H / max_col))
    base_pt = max(8, min(14, int(step_h * 26)))

    def render_column(items, x_start, col_w):
        y = Y_START
        for el in items:
            if y >= Y_MAX:
                break
            etype   = el.get("type", "text")
            content = el.get("content", "")
            style   = el.get("style") or {}
            color   = resolve_color(style.get("color", "black"))
            bold    = style.get("bold", False) or (etype == "heading")
            underl  = style.get("underline", False)
            circled = style.get("circled", False)
            size    = {"large": base_pt + 2, "medium": base_pt, "small": base_pt - 2}.get(
                          style.get("size", "medium"), base_pt)

            if etype == "heading":
                label = f"【{content}】" if circled else content
                txt(s, label, x_start, y, col_w, step_h,
                    size=min(size + 2, base_pt + 3), bold=True, color=color, underline=underl)
            elif etype == "arrow":
                txt(s, f"→ {content}", x_start + 0.1, y, col_w - 0.1, step_h,
                    size=max(size - 1, 8), color=RGBColor(0xFF, 0xB7, 0x4D), italic=True)
            else:
                prefix = "○ " if circled else "• "
                txt(s, f"{prefix}{content}", x_start + 0.1, y, col_w - 0.1, step_h,
                    size=size, bold=bold, color=color, underline=underl)
            y += step_h

    render_column(col_left,  0.45, 6.0)
    render_column(col_right, 6.75, 6.0)

    # ─ 表を配置 ─
    for tbl_data in data.get("tables", []):
        pos   = tbl_data.get("position") or {}
        tbl_y = 1.25 + pos.get("zone_row", 2) * 1.8
        tbl_x = 0.45 + pos.get("zone_col", 0) * 4.3
        add_table_shape(s, tbl_data, tbl_x, tbl_y, max_w=5.5, max_h=1.8)

    # フッター
    rect(s, 0, 7.15, 13.33, 0.35, C_NAVY)
    txt(s, "パシャッと  |  自動生成", 0.4, 7.17, 12.9, 0.30,
        size=9, color=C_MUTED, align=PP_ALIGN.RIGHT)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ─── 画面表示 ────────────────────────────────────────────
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

st.markdown("""
<div class='top-bar'>
  <div class='top-bar-logo'>パシャッ<em>と</em></div>
  <div class='top-bar-tagline'>メモの写真を<br>スライドに変換</div>
</div>
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
# ステップ１　写真を選ぶ
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
st.markdown(
    "<div class='step'><span class='snum'>１</span>写真を選ぶ</div>",
    unsafe_allow_html=True)
st.markdown(
    "<div class='hint'>ホワイトボードや手書きメモを、<strong>正面から</strong>撮った写真を選んでください。</div>",
    unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "画像ファイルを選ぶ",
    type=["jpg", "jpeg", "png", "webp"],
    label_visibility="collapsed",
    help="JPG・PNG・WebP 形式の画像に対応しています。",
)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 画像プレビュー ＆ 前処理
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
img_data   = None
media_type = "image/jpeg"

if uploaded_file:
    pil_image = Image.open(uploaded_file)
    if pil_image.mode != "RGB":
        pil_image = pil_image.convert("RGB")

    st.markdown("---")

    do_trim = st.toggle(
        "✂️　余白を自動でカット（おすすめ）",
        value=True,
        help="余白を除去して読み取り精度を上げます。",
    )

    display_img = preprocess_image(pil_image, do_trim=do_trim)
    st.image(display_img, caption="変換する写真", use_container_width=True)
    img_data, media_type = pil_to_base64(display_img)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ステップ２　変換する
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
st.markdown("---")
st.markdown(
    "<div class='step'><span class='snum'>２</span>変換する</div>",
    unsafe_allow_html=True)

if not uploaded_file:
    st.markdown(
        "<div class='hint' style='border-color:#F5A623; background:#FFF8E7;'>"
        "↑ まず写真を選んでください。"
        "</div>",
        unsafe_allow_html=True)

convert_btn = st.button(
    "✨　スライドに変換する",
    type="primary",
    use_container_width=True,
    disabled=not uploaded_file,
)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 変換処理
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
if convert_btn and img_data:
    try:
        api_key = get_api_key()
        if not api_key:
            st.warning("認証キーが設定されていません。左上のメニューから設定してください。")
            st.stop()

        status   = st.empty()
        progress = st.progress(0)

        status.markdown("**① 写真を準備しています…**")
        progress.progress(15)
        status.markdown("**② 写真を読み取っています…**")
        progress.progress(35)

        extracted = analyze_with_claude(img_data, media_type, api_key)

        status.markdown("**③ 文字の読み取りが完了しました ✅**")
        progress.progress(75)
        status.markdown("**④ スライドを作成しています…**")
        progress.progress(90)

        st.session_state["extracted"]  = extracted
        st.session_state["pptx_bytes"] = generate_pptx(extracted)

        progress.progress(100)
        status.markdown("**✅ 完了しました！下にスクロールして保存してください。**")

    except Exception as e:
        err_msg = str(e)
        if "api_key" in err_msg.lower() or "authentication" in err_msg.lower() or "401" in err_msg:
            st.error("認証キーが正しくありません。左上のメニューから確認してください。")
        elif "json" in err_msg.lower() or "JSON" in err_msg or "応答" in err_msg:
            st.error(f"文字の読み取り結果を処理できませんでした。もう一度お試しください。\n（詳細：{err_msg[:200]}）")
        elif "timeout" in err_msg.lower() or "connection" in err_msg.lower():
            st.error("通信エラーが発生しました。インターネット接続をご確認のうえ、もう一度お試しください。")
        elif "overloaded" in err_msg.lower() or "529" in err_msg:
            st.error("AIサーバーが混み合っています。しばらく待ってから、もう一度お試しください。")
        else:
            st.error(f"エラーが発生しました。もう一度お試しください。\n（詳細：{err_msg}）")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ステップ３　保存する
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
if "extracted" in st.session_state:
    extracted  = st.session_state["extracted"]
    pptx_bytes = st.session_state.get("pptx_bytes")

    st.markdown("---")
    st.markdown(
        "<div class='step'><span class='snum'>３</span>保存する</div>",
        unsafe_allow_html=True)

    el_count = len(extracted.get("elements", []))
    st.markdown(f"""
<div class='done-box'>
✅ 完了　読み取り：<strong>{el_count} 件</strong>
</div>
""", unsafe_allow_html=True)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    if pptx_bytes:
        st.download_button(
            label="💾　パワーポイントで保存する",
            data=pptx_bytes,
            file_name=f"パシャッと_{ts}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )

    with st.expander("📋　読み取り内容を確認する"):
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
            {"バージョン": "1.0", "変換日時": datetime.now().isoformat(), "データ": extracted},
            ensure_ascii=False, indent=2,
        )
        st.download_button(
            label="📄　テキストで保存する",
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
