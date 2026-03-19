"""
パシャッと — 手書きメモ・ホワイトボード → スライド自動変換
"""

import os, io, json, base64, re, copy
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

# ─── CSS（20代女性向け・グラデーションバイオレット） ────────────────────────
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

/* ── モバイルタップ最適化（効果音・ハイライト抑制） ── */
* { -webkit-tap-highlight-color: transparent; }
button, a { touch-action: manipulation; -webkit-appearance: none; }

/* ── 全体背景（クリーンホワイト） ── */
.stApp { background: #FFFFFF; }
.main .block-container { padding: 0 1rem 5rem; max-width: 560px; }

/* ── 固定ヘッダーバー（グラデーション） ── */
.top-bar {
    position: sticky;
    top: 0;
    z-index: 999;
    background: linear-gradient(135deg, #6D28D9 0%, #A855F7 55%, #EC4899 100%);
    padding: 0.6rem 1.2rem;
    margin: 0 -1rem 1.6rem;
    display: flex;
    align-items: center;
    gap: 0.6rem;
    box-shadow: 0 3px 16px rgba(109,40,217,0.30);
}
.top-bar-logo {
    font-size: 1.55rem;
    font-weight: 900;
    color: #ffffff;
    letter-spacing: -0.01em;
    line-height: 1;
}
.top-bar-logo em { color: #FDE8FF; font-style: normal; }
.top-bar-tagline {
    font-size: 0.88rem;
    color: rgba(255,255,255,0.80);
    font-weight: 500;
    line-height: 1.35;
}

/* ── ステップラベル ── */
.step {
    display: flex;
    align-items: center;
    gap: 0.75rem;
    font-size: 1.35rem;
    font-weight: 800;
    color: #1E1B4B;
    margin: 2rem 0 0.6rem;
}
.snum {
    background: linear-gradient(135deg, #7C3AED, #A855F7);
    color: #fff;
    border-radius: 50%;
    width: 40px; height: 40px;
    line-height: 40px;
    text-align: center;
    font-size: 1.1rem;
    font-weight: 900;
    flex-shrink: 0;
    box-shadow: 0 3px 10px rgba(124,58,237,0.35);
}

/* ── ヒント・説明文 ── */
.hint {
    color: #3B0764;
    font-size: 1.05rem;
    margin: 0 0 1rem;
    line-height: 1.8;
    background: #F5F3FF;
    border-left: 4px solid #7C3AED;
    border-radius: 0 8px 8px 0;
    padding: 0.75rem 1rem;
}

/* ── ボタン全般 ── */
.stButton > button {
    font-size: 1.2rem !important;
    min-height: 66px !important;
    border-radius: 18px !important;
    font-weight: 800 !important;
    width: 100% !important;
    letter-spacing: 0.02em !important;
    transition: opacity 0.12s, transform 0.1s !important;
    cursor: pointer !important;
}
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #7C3AED 0%, #A855F7 50%, #EC4899 100%) !important;
    color: white !important;
    border: none !important;
    box-shadow: 0 5px 20px rgba(124,58,237,0.40) !important;
}
.stButton > button[kind="primary"]:hover  { opacity: 0.88 !important; }
.stButton > button[kind="primary"]:active { transform: scale(0.97) !important; }
.stButton > button[kind="secondary"] {
    background: #F5F3FF !important;
    color: #4C1D95 !important;
    border: 2px solid #C4B5FD !important;
}
.stButton > button:disabled { opacity: 0.35 !important; cursor: not-allowed !important; }

/* ── ダウンロードボタン（メイン：グラデーション） ── */
.stDownloadButton > button {
    font-size: 1.2rem !important;
    min-height: 66px !important;
    border-radius: 18px !important;
    font-weight: 800 !important;
    width: 100% !important;
    border: none !important;
    letter-spacing: 0.02em !important;
    transition: opacity 0.12s, transform 0.1s !important;
    cursor: pointer !important;
}
.stDownloadButton > button:active { transform: scale(0.97) !important; }

/* ── 完了ボックス ── */
.done-box {
    background: #FAF5FF;
    border: 2px solid #A855F7;
    border-radius: 14px;
    padding: 1.2rem 1.4rem;
    font-size: 1.1rem;
    color: #4C1D95;
    font-weight: 700;
    margin: 1.2rem 0;
    line-height: 1.8;
}

/* ── ファイルアップローダー ── */
[data-testid="stFileUploaderDropzone"] {
    min-height: 150px;
    border-radius: 14px !important;
    border: 2.5px dashed #C4B5FD !important;
    background: #FAFAFF !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] span,
[data-testid="stFileUploaderDropzoneInstructions"] small { display: none !important; }
[data-testid="stFileUploaderDropzoneInstructions"]::before {
    content: "写真をドラッグするか、下のボタンで選んでください";
    display: block;
    font-size: 1.0rem;
    color: #5B21B6;
    margin-bottom: 0.5rem;
    text-align: center;
}
[data-testid="stFileUploaderDropzoneInstructions"]::after {
    content: "対応形式：JPG・PNG・WebP・PDF　最大20MBまで";
    display: block;
    font-size: 0.85rem;
    color: #8B5CF6;
    text-align: center;
    margin-top: 0.3rem;
}
[data-testid="stFileUploaderDropzone"] button {
    font-size: 0 !important;
    min-height: 50px !important;
    border-radius: 12px !important;
    padding: 0 1.5rem !important;
}
[data-testid="stFileUploaderDropzone"] button::after {
    content: "写真・ファイルを選ぶ";
    font-size: 1.0rem;
    font-weight: 700;
}

/* ── その他 ── */
.stToggle label { font-size: 1.05rem !important; font-weight: 700 !important; }
.stExpander summary p { font-size: 1.05rem !important; font-weight: 600 !important; }
hr { margin: 1.6rem 0 !important; border-color: #EDE9FE !important; border-width: 1.5px !important; }
.stAlert p { font-size: 1.0rem !important; line-height: 1.7 !important; }
.stSpinner p { font-size: 1.05rem !important; }
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


def get_pdf_info(file_bytes: bytes) -> int:
    """PDFのページ数を返す"""
    import fitz  # PyMuPDF
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    count = doc.page_count
    doc.close()
    return count


def pdf_page_to_pil(file_bytes: bytes, page_num: int = 0) -> Image.Image:
    """PDFの指定ページを PIL Image（RGB）に変換（2倍解像度）"""
    import fitz
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    page = doc[page_num]
    mat = fitz.Matrix(2.0, 2.0)        # 2×ズームで解像度を確保
    pix = page.get_pixmap(matrix=mat, colorspace=fitz.csRGB)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    doc.close()
    return img


def preprocess_image(img: Image.Image,
                     do_trim: bool = True) -> Image.Image:
    """EXIF回転 → OCR最適化補正（自動）→ 余白カット

    補正方針:
    1. 照明むらが大きい場合のみ Retinex 正規化（影・逆光を自動検出）
    2. autocontrast でヒストグラムを常に最大伸張
    3. シャープネス強化でテキスト輪郭を際立たせる
    """
    import numpy as np
    from PIL import ImageFilter

    try:
        img = ImageOps.exif_transpose(img)
    except Exception:
        pass
    if img.mode != "RGB":
        img = img.convert("RGB")

    # ── 照明むら自動検出 ──
    # ブラー画像の輝度標準偏差が大きいほど照明が不均一
    radius   = max(40, min(img.width, img.height) // 6)
    blur_pil = img.filter(ImageFilter.GaussianBlur(radius=radius))
    blur_l   = np.array(blur_pil.convert("L"), dtype=np.float32)
    if float(blur_l.std()) > 22:      # 照明むらが目立つ → Retinex で均一化
        arr  = np.array(img, dtype=np.float32)
        blur = np.array(blur_pil, dtype=np.float32)
        norm = arr / (blur / 160.0 + 0.5)
        img  = Image.fromarray(np.clip(norm, 0, 255).astype(np.uint8))

    # ── ヒストグラム自動伸張（常に適用） ──
    img = ImageOps.autocontrast(img, cutoff=1)

    # ── シャープネス強化 ──
    img = ImageEnhance.Sharpness(img).enhance(1.5)

    # ── 余白カット ──
    if do_trim:
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
    """PIL Image → OCR用 base64（preprocess 済み画像をそのままエンコード）"""
    if img.width > max_px or img.height > max_px:
        img.thumbnail((max_px, max_px), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=95)
    return base64.standard_b64encode(buf.getvalue()).decode("utf-8"), "image/jpeg"


@st.cache_data(show_spinner=False)
def cached_process_image(file_bytes: bytes,
                         do_trim: bool) -> tuple[str, str, bool, bytes]:
    """
    画像のバイト列を受け取り、前処理 → base64変換 → プレビュー用バイト列を返す。
    同じ入力なら Streamlit がキャッシュするため、リランのたびに再処理されない。
    戻り値: (img_data_b64, media_type, is_portrait, preview_jpeg_bytes)
    """
    pil = Image.open(io.BytesIO(file_bytes))
    if pil.mode != "RGB":
        pil = pil.convert("RGB")
    disp = preprocess_image(pil, do_trim=do_trim)
    is_portrait = disp.height > disp.width
    img_data, media_type = pil_to_base64(disp)
    # プレビュー用にJPEGバイト列として保持（PIL Imageはキャッシュ不可）
    preview_buf = io.BytesIO()
    disp.save(preview_buf, format="JPEG", quality=88)
    return img_data, media_type, is_portrait, preview_buf.getvalue()


# 色帯パレット（アイテムごとに割り当てる 12 色）
_BAND_COLORS = [
    (231, 76,  60),   # red
    (52,  152, 219),  # blue
    (39,  174, 96),   # green
    (155, 89,  182),  # purple
    (230, 126, 34),   # orange
    (32,  178, 170),  # teal
    (236, 72,  153),  # pink
    (99,  110, 250),  # indigo
    (243, 156, 18),   # yellow
    (76,  201, 240),  # cyan
    (67,  97,  238),  # royal blue
    (247, 37,  133),  # hot pink
]

def item_color_hex(global_idx: int) -> str:
    r, g, b = _BAND_COLORS[global_idx % len(_BAND_COLORS)]
    return f"#{r:02x}{g:02x}{b:02x}"


def build_annotated_image(preview_bytes: bytes, data: dict) -> bytes:
    """各アイテムの位置に色帯オーバーレイを描画した画像を返す。
    左端に不透明バー、全幅に薄い帯を重ねることで
    フォームの色インジケーターとの対応を視覚的に示す。
    """
    from PIL import ImageDraw

    img  = Image.open(io.BytesIO(preview_bytes)).convert("RGBA")
    W, H = img.size
    overlay = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    draw    = ImageDraw.Draw(overlay)

    band_h = max(8, H // 45)      # 帯の高さ
    side_w = max(20, W // 35)     # 左端インジケーター幅
    MIN_GAP = band_h + 2
    last_cy = -MIN_GAP

    all_items = [
        item
        for sec in (data.get("sections") or [])
        for item in (sec.get("items") or [])
    ]
    total = len(all_items)

    for idx, item in enumerate(all_items):
        r, g, b = _BAND_COLORS[idx % len(_BAND_COLORS)]

        yp = item.get("y_pct")
        if yp is None:
            yp = int(100 * (idx + 0.5) / max(total, 1))
        cy = max(band_h + 2, min(H - band_h - 2, int(H * yp / 100)))
        cy = max(cy, last_cy + MIN_GAP)
        last_cy = cy

        y0, y1 = cy - band_h // 2, cy + band_h // 2

        # 全幅：薄い透明帯（識別しやすいが文字を隠さない）
        draw.rectangle([0, y0, W, y1], fill=(r, g, b, 55))
        # 左端：不透明インジケーターバー
        draw.rectangle([0, y0, side_w, y1], fill=(r, g, b, 230))

    img = Image.alpha_composite(img, overlay).convert("RGB")
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=85)
    return buf.getvalue()


EXTRACTION_PROMPT = """あなたは優秀なOCRエンジンです。
この画像（ホワイトボード・手書きメモ・付箋など）の内容を読み取り、プレゼンテーション用に整理してください。

【解析手順】
STEP1: 画像全体を俯瞰して「1列構成か2列構成か」を判断する。
STEP2: 上から下・左から右の順にすべての文字を正確に読み取る。
STEP3: 意味的なまとまり（セクション）ごとにグループ化する。
  ・明確な区切り・見出しがあればセクションを分ける
  ・2列構成なら左側を column:1、右側を column:2 とする

【返却JSON形式】
{"title":"最も目立つタイトル・主題（スライドのタイトルとなる1行）","sections":[{"heading":"セクション見出し（なければ空文字）","column":1,"items":[{"text":"正確な文字列","type":"heading","shape":"none","color":"black","bold":true,"y_pct":10},{"text":"箇条書き","type":"bullet","shape":"none","color":"black","bold":false,"y_pct":25},{"text":"四角囲み","type":"text","shape":"rect","color":"red","bold":false,"y_pct":40},{"text":"丸囲み","type":"text","shape":"ellipse","color":"blue","bold":false,"y_pct":55},{"text":"矢印ラベル","type":"arrow","shape":"none","color":"black","bold":false,"y_pct":70}]}],"tables":[{"headers":["列1","列2"],"rows":[["値1","値2"]]}]}

【必須ルール】
- title はスライドヘッダー専用（sections の items には含めない）
- column: 左側・1列レイアウトは 1、右側コンテンツは 2
- type: heading（大きな文字・見出し）/ bullet（箇条書き）/ text（通常）/ arrow（矢印ラベル）
- shape: 四角囲み="rect"、丸・楕円囲み="ellipse"、囲みなし="none"
- color: black/red/blue/green/orange/purple/pink/gray/brown/yellow
- y_pct: その文字が画像の上端から何%の位置にあるか（0=最上部、100=最下部、整数）
- 画像内のすべての文字を漏れなく・正確に読む（推測・省略禁止）
- 表がなければ tables:[]
- JSONのみ返す（説明文・コードブロック不要）"""


def analyze_with_claude(img_data: str, media_type: str, api_key: str,
                        model: str = "claude-sonnet-4-6") -> dict:
    """AIで画像を解析（堅牢なJSON抽出）"""
    import anthropic
    client = anthropic.Anthropic(api_key=api_key)
    message = client.messages.create(
        model=model,
        max_tokens=8192,
        messages=[
            {"role": "user", "content": [
                {"type": "image", "source": {
                    "type": "base64", "media_type": media_type, "data": img_data
                }},
                {"type": "text", "text": EXTRACTION_PROMPT},
            ]},
        ]
    )
    raw = message.content[0].text.strip()

    # ① コードブロック除去
    text = re.sub(r"```(?:json)?\s*", "", raw)
    text = re.sub(r"```", "", text).strip()

    # ② 最初の { から最後の } を抽出
    start = text.find("{")
    end   = text.rfind("}")
    if start == -1 or end == -1 or end <= start:
        raise ValueError(f"応答にJSONが見つかりません（応答の先頭: {raw[:200]}）")
    text = text[start:end + 1]

    # ③ 制御文字を除去してパース
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)
    try:
        data = json.loads(text)
    except json.JSONDecodeError as e:
        raise ValueError(f"JSON解析エラー: {e}\n応答の先頭: {text[:200]}")

    # ── sections 形式への正規化 ──
    # null や欠落を安全に [] へ
    if not data.get("sections"):
        data["sections"] = []
    if not data.get("tables"):
        data["tables"] = []

    # 旧 items 形式が返ってきた場合は sections に変換
    if not data["sections"] and data.get("items"):
        data["sections"] = [{
            "heading": "",
            "column": 1,
            "items": [
                {"text": it.get("text", ""), "type": it.get("type", "text"),
                 "shape": it.get("shape", "none"), "color": it.get("color", "black"),
                 "bold": it.get("bold", False)}
                for it in (data["items"] or [])
            ],
        }]

    # sections 内の items が null のものを修復 + y_pct を 0〜100 に正規化
    for sec in data["sections"]:
        if sec.get("items") is None:
            sec["items"] = []
        for item in sec["items"]:
            yp = item.get("y_pct")
            if yp is not None:
                item["y_pct"] = max(0, min(100, int(yp)))

    # ── UI 表示用の elements リストを sections から構築 ──
    elements = []
    for sec in data.get("sections", []):
        for it in (sec.get("items") or []):
            idx = len(elements)
            elements.append({
                "id":      f"el_{idx:03d}",
                "type":    it.get("type", "text"),
                "content": it.get("text", ""),
                "shape":   it.get("shape", "none"),
                "style":   {"color": it.get("color", "black"),
                             "bold":  it.get("bold", False)},
                "confidence": 1.0,
            })
    data["elements"] = elements
    data.setdefault("structure", {"type": "list", "groups": [
        {"label": "内容", "items": [e["id"] for e in elements]}
    ]})
    return data


def generate_pptx(data: dict, is_portrait: bool = False) -> bytes:
    """解析データからパワーポイントを生成（縦横自動対応）"""
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

    # ─ 縦横でスライドサイズを切り替え ─
    if is_portrait:
        SLIDE_W, SLIDE_H = 7.5, 10.83   # A4縦
    else:
        SLIDE_W, SLIDE_H = 13.33, 7.5   # 16:9横

    prs = Presentation()
    prs.slide_width  = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)

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

    def shape_txt(slide, shape_type, text, x, y, w, h,
                  size=13, bold=False, color=None, underline=False):
        """テキスト付きオートシェイプ（rect / ellipse）を追加"""
        SHAPE_ID = {"rect": 1, "ellipse": 9, "rounded_rect": 5}
        sid = SHAPE_ID.get(shape_type, 1)
        color = color or C_LIGHT
        sh = slide.shapes.add_shape(sid,
                                    Inches(x), Inches(y), Inches(w), Inches(h))
        # 塗りつぶし：ネイビー + テキスト色のボーダー
        sh.fill.solid()
        sh.fill.fore_color.rgb = C_NAVY
        sh.line.color.rgb = color
        sh.line.width = Pt(1.5)
        tf = sh.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text           = text
        r.font.size      = Pt(size)
        r.font.bold      = bold
        r.font.underline = underline
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
    rect(s, 0, 0, SLIDE_W, 0.07, C_ACCENT)

    title    = data.get("title", "ホワイトボードメモ")
    ts_label = datetime.now().strftime("%Y年%m月%d日 %H:%M")
    CONTENT_X = 0.4
    TITLE_W   = SLIDE_W - 0.8

    txt(s, title, CONTENT_X, 0.10, TITLE_W, 0.65,
        size=20 if is_portrait else 24, bold=True, color=C_WHITE)
    txt(s, f"変換日時：{ts_label}", CONTENT_X, 0.73, TITLE_W, 0.30,
        size=9, color=C_MUTED)
    rect(s, 0.3, 1.08, SLIDE_W - 0.6, SLIDE_H - 1.55, C_NAVY)

    # ── コンテンツエリア定義 ──
    # タイトルバー下からフッター上までを「コンテンツエリア」とする
    CX = 0.35              # 左マージン（インチ）
    CY = 1.10              # コンテンツ開始Y（インチ）
    CW = SLIDE_W - 0.70    # コンテンツ幅
    CH = SLIDE_H - 1.55    # コンテンツ高さ（フッター分を引く）

    sections = data.get("sections") or []
    tables   = data.get("tables") or []

    # ── 座標を一切使わない sections ベースのクリーンレイアウト ──
    # AIには「何がどんな構造か」だけ聞き、配置はすべてコードが決める。

    col1_secs = [s for s in sections if s.get("column", 1) != 2]
    col2_secs = [s for s in sections if s.get("column", 1) == 2]
    has_two_cols = bool(col2_secs) and not is_portrait

    def count_rows(secs):
        """セクション群の総行数（見出し行 + アイテム行）"""
        n = 0
        for sec in secs:
            if sec.get("heading"):
                n += 1
            n += len(sec.get("items", []))
        return max(n, 1)

    if has_two_cols:
        col_w  = (CW - 0.30) / 2
        n_rows = max(count_rows(col1_secs), count_rows(col2_secs))
    else:
        col_w  = CW
        n_rows = count_rows(col1_secs + col2_secs)

    # ── 統一フォントサイズ（行数から逆算）──
    EL_H    = max(0.24, min(0.44, CH / n_rows))
    base_pt = max(9, min(16, int(EL_H * 72 * 0.50)))
    head_pt = min(base_pt + 3, 20)       # 見出し（heading type）
    sec_pt  = min(base_pt + 5, 22)       # セクション見出し

    def render_sections(secs, x_start, width):
        y = CY
        for sec in secs:
            heading = sec.get("heading", "").strip()
            items   = sec.get("items") or []

            # セクション見出し（アンダーライン付き）
            if heading:
                txt(s, heading, x_start, y, width, EL_H,
                    size=sec_pt, bold=True, color=C_WHITE, underline=True)
                y += EL_H

            for item in items:
                etype      = item.get("type", "text")
                content    = item.get("text", "")
                shape_type = item.get("shape", "none")
                color      = resolve_color(item.get("color", "black"))
                bold       = item.get("bold", False) or etype == "heading"

                if shape_type in ("rect", "ellipse"):
                    shape_txt(s, shape_type, content,
                              x_start + 0.05, y, width - 0.1, EL_H,
                              size=base_pt, bold=bold, color=color)
                elif etype == "heading":
                    txt(s, content, x_start, y, width, EL_H,
                        size=head_pt, bold=True, color=color)
                elif etype == "arrow":
                    txt(s, f"→ {content}", x_start + 0.1, y, width - 0.1, EL_H,
                        size=base_pt, color=RGBColor(0xFF, 0xB7, 0x4D), italic=True)
                elif etype == "bullet":
                    txt(s, f"・{content}", x_start + 0.1, y, width - 0.1, EL_H,
                        size=base_pt, bold=bold, color=color)
                else:
                    txt(s, content, x_start + 0.1, y, width - 0.1, EL_H,
                        size=base_pt, bold=bold, color=color)
                y += EL_H

    if has_two_cols:
        render_sections(col1_secs, CX,              col_w)
        render_sections(col2_secs, CX + col_w + 0.30, col_w)
    else:
        render_sections(col1_secs + col2_secs, CX, col_w)

    # ─ 表を配置（コンテンツエリア末尾）─
    tbl_y = CY + n_rows * EL_H + 0.1
    for tbl_data in tables:
        if tbl_y >= CY + CH:
            break
        tbl_h = min(1.8, CY + CH - tbl_y)
        add_table_shape(s, tbl_data, CX, tbl_y, max_w=CW, max_h=tbl_h)
        tbl_y += tbl_h + 0.1

    # ─ フッター ─
    FOOTER_Y = SLIDE_H - 0.35
    rect(s, 0, FOOTER_Y, SLIDE_W, 0.35, C_NAVY)
    txt(s, "パシャッと  |  自動生成", 0.4, FOOTER_Y + 0.02, SLIDE_W - 0.5, 0.30,
        size=9, color=C_MUTED, align=PP_ALIGN.RIGHT)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def generate_html(data: dict) -> bytes:
    """iPhone Safari でそのまま開けるHTMLを sections ベースで生成"""
    title    = data.get("title", "メモ")
    sections = data.get("sections") or []
    tables   = data.get("tables") or []
    CMAP = {
        "black": "#1E1B4B", "white": "#6B7280", "red": "#DC2626",
        "blue": "#2563EB", "green": "#16A34A", "yellow": "#B45309",
        "orange": "#EA580C", "purple": "#7C3AED", "pink": "#DB2777",
        "gray": "#6B7280", "brown": "#92400E",
    }

    def render_item(item):
        etype   = item.get("type", "text")
        content = item.get("text", "")
        shape   = item.get("shape", "none")
        color   = CMAP.get(item.get("color", "black"), "#1E1B4B")
        bold    = item.get("bold", False) or etype == "heading"
        fw      = "700" if bold else "400"
        if shape == "rect":
            return (f"<span class='badge rect' style='color:{color};"
                    f" border-color:{color}; font-weight:{fw};'>{content}</span>")
        if shape == "ellipse":
            return (f"<span class='badge ellipse' style='color:{color};"
                    f" border-color:{color}; font-weight:{fw};'>{content}</span>")
        if etype == "heading":
            return f"<h3 style='color:{color};'>{content}</h3>"
        if etype == "arrow":
            return f"<p class='arrow'>&rarr; {content}</p>"
        if etype == "bullet":
            return f"<li style='color:{color}; font-weight:{fw};'>{content}</li>"
        return f"<p style='color:{color}; font-weight:{fw};'>{content}</p>"

    def render_section(sec):
        heading = sec.get("heading", "").strip()
        items   = sec.get("items") or []
        out = ""
        if heading:
            out += f"<h2 class='sec-heading'>{heading}</h2>"
        # 箇条書きをまとめて <ul> に（インデックスを使わず状態フラグで管理）
        in_ul = False
        for item in items:
            if item.get("type") == "bullet":
                if not in_ul:
                    out += "<ul>"
                    in_ul = True
                out += render_item(item)
            else:
                if in_ul:
                    out += "</ul>"
                    in_ul = False
                out += render_item(item)
        if in_ul:
            out += "</ul>"
        return out

    body = ""
    col1 = [s for s in sections if s.get("column", 1) != 2]
    col2 = [s for s in sections if s.get("column", 1) == 2]
    if col2:
        body += "<div class='two-col'>"
        body += "<div class='col'>" + "".join(render_section(s) for s in col1) + "</div>"
        body += "<div class='col'>" + "".join(render_section(s) for s in col2) + "</div>"
        body += "</div>"
    else:
        body = "".join(render_section(s) for s in col1)

    # テーブル
    for tbl in tables:
        headers = tbl.get("headers", [])
        trows   = tbl.get("rows", [])
        if not headers and not trows:
            continue
        body += "<table>"
        if headers:
            body += "<tr>" + "".join(f"<th>{h}</th>" for h in headers) + "</tr>"
        for i, row in enumerate(trows):
            cls = "even" if i % 2 == 0 else ""
            body += f"<tr class='{cls}'>" + "".join(f"<td>{c}</td>" for c in row) + "</tr>"
        body += "</table>"

    ts = datetime.now().strftime("%Y年%m月%d日 %H:%M")
    html = f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{title} — パシャッと</title>
<style>
  * {{ box-sizing: border-box; }}
  body {{ font-family: -apple-system, 'Hiragino Sans', 'Yu Gothic', sans-serif;
          background: #F5F3FF; margin: 0; padding: 1rem; color: #1E1B4B;
          font-size: 16px; line-height: 1.7; }}
  .card {{ background: #fff; border-radius: 16px; padding: 1.5rem 1.6rem;
           box-shadow: 0 2px 16px rgba(109,40,217,0.12); max-width: 720px;
           margin: 0 auto; }}
  .header {{ background: linear-gradient(135deg,#6D28D9,#A855F7,#EC4899);
             border-radius: 12px; padding: 1rem 1.4rem; margin-bottom: 1.2rem; }}
  .header h1 {{ color:#fff; margin:0; font-size:1.3rem; }}
  .header p  {{ color:rgba(255,255,255,0.75); margin:0.25rem 0 0; font-size:0.8rem; }}
  hr {{ border:none; border-top:1px solid #EDE9FE; margin:1rem 0; }}
  h2.sec-heading {{ font-size:1.05rem; font-weight:800; color:#4C1D95;
                    border-left:4px solid #7C3AED; padding-left:0.6rem;
                    margin:1.2rem 0 0.5rem; }}
  h3 {{ font-size:1rem; margin:0.8rem 0 0.3rem; }}
  ul {{ padding-left:1.4rem; margin:0.3rem 0; }}
  li {{ margin:0.2rem 0; }}
  p  {{ margin:0.3rem 0; }}
  .arrow {{ color:#7C3AED; font-weight:600; }}
  .badge {{ display:inline-block; border:2px solid; border-radius:6px;
            padding:0.15rem 0.6rem; margin:0.25rem 0; font-weight:600; }}
  .badge.ellipse {{ border-radius:50px; padding:0.15rem 0.9rem; }}
  .two-col {{ display:flex; gap:1.2rem; }}
  .col {{ flex:1; min-width:0; }}
  table {{ border-collapse:collapse; width:100%; margin:1rem 0; font-size:0.9rem; }}
  th {{ background:#7C3AED; color:#fff; padding:0.4rem 0.6rem;
        text-align:left; border:1px solid #C4B5FD; }}
  td {{ padding:0.35rem 0.6rem; border:1px solid #EDE9FE; }}
  tr.even td {{ background:#F5F3FF; }}
  footer {{ text-align:center; color:#A78BFA; font-size:0.75rem; margin-top:1.5rem; }}
  @media (max-width:520px) {{ .two-col {{ flex-direction:column; }} }}
</style>
</head>
<body>
<div class="card">
  <div class="header">
    <h1>{title}</h1>
    <p>{ts} 変換 | パシャッと</p>
  </div>
  <hr>
  {body}
</div>
<footer>パシャッと — AI文字認識による自動変換</footer>
</body>
</html>"""
    return html.encode("utf-8")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ─── 画面表示 ────────────────────────────────────────────
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

st.markdown("""
<div class='top-bar'>
  <div class='top-bar-logo'>パシャッ<em>と</em></div>
  <div class='top-bar-tagline'>大切なメモを<br>きれいなスライドへ</div>
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
    "<div class='hint'>ホワイトボード・メモ帳の写真や、スキャンしたPDFを選ぶだけ。あとはAIがきれいに仕上げます。</div>",
    unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "画像ファイルを選ぶ",
    type=["jpg", "jpeg", "png", "webp", "pdf"],
    label_visibility="collapsed",
    help="JPG・PNG・WebP・PDF 形式に対応しています。",
)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 画像プレビュー ＆ 前処理
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# session_state から前回の処理済みデータを取得（ファイル未選択時のフォールバック）
img_data    = st.session_state.get("_img_data")
media_type  = st.session_state.get("_media_type", "image/jpeg")
is_portrait = st.session_state.get("_is_portrait", False)

if uploaded_file:
    st.markdown("---")

    is_pdf = uploaded_file.name.lower().endswith(".pdf")

    # ── PDF：ページ選択 ──
    pdf_page_num = 0
    if is_pdf:
        try:
            raw_bytes  = uploaded_file.getvalue()
            page_count = get_pdf_info(raw_bytes)
            if page_count > 1:
                st.markdown(
                    f"<div class='hint'>このPDFは <strong>{page_count} ページ</strong> あります。"
                    "変換するページを選んでください。</div>",
                    unsafe_allow_html=True)
                pdf_page_num = st.number_input(
                    "ページ番号",
                    min_value=1, max_value=page_count, value=1, step=1,
                    help=f"1〜{page_count} の範囲で選べます",
                ) - 1   # 0-indexed に変換
            else:
                st.caption("PDF（1ページ）を読み込みました")
        except Exception as pdf_info_err:
            st.error(f"PDFの読み込みに失敗しました。（{pdf_info_err}）")

    do_trim = st.toggle(
        "余白を自動カット（おすすめ）",
        value=True,
        help="不要な余白を取り除いて、読み取り精度をアップします。",
    )
    st.caption("影・逆光・コントラストはAIが自動で最適化します")

    try:
        file_bytes = uploaded_file.getvalue()

        # PDF の場合は先にページを画像に変換してから処理
        if is_pdf:
            pil_from_pdf = pdf_page_to_pil(file_bytes, page_num=pdf_page_num)
            buf_for_cache = io.BytesIO()
            pil_from_pdf.save(buf_for_cache, format="JPEG", quality=95)
            proc_bytes = buf_for_cache.getvalue()
        else:
            proc_bytes = file_bytes

        # キャッシュキー：ファイル名＋サイズ（＋PDFならページ番号も含める）
        file_id = f"{uploaded_file.name}_{len(file_bytes)}_p{pdf_page_num}"

        # キャッシュ済み処理（同一ファイル＋設定ならリランしても即返却）
        img_data, media_type, is_portrait, preview_bytes = cached_process_image(
            proc_bytes, do_trim
        )

        # ファイルが変わったら前の変換結果・編集内容をリセット
        if st.session_state.get("_file_id") != file_id:
            for k in ("extracted", "pptx_bytes", "html_bytes"):
                st.session_state.pop(k, None)
            for k in [k for k in st.session_state if k.startswith("edit_")]:
                del st.session_state[k]
            st.session_state["_file_id"] = file_id

        # session_state に保持（リラン後もボタン押下で参照できるよう）
        st.session_state["_img_data"]       = img_data
        st.session_state["_media_type"]     = media_type
        st.session_state["_is_portrait"]    = is_portrait
        st.session_state["_preview_bytes"]  = preview_bytes

        orient_label = "縦（ポートレート）" if is_portrait else "横（ランドスケープ）"
        cap_label    = "変換するページ" if is_pdf else "変換する写真"
        st.caption(f"スライド向き：{orient_label} で出力します")
        st.image(Image.open(io.BytesIO(preview_bytes)),
                 caption=cap_label, use_container_width=True)

    except Exception as img_err:
        st.error(f"ファイルの読み込みに失敗しました。別のファイルをお試しください。\n（{img_err}）")
        img_data = None

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ステップ２　変換する
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
st.markdown("---")
st.markdown(
    "<div class='step'><span class='snum'>２</span>変換する</div>",
    unsafe_allow_html=True)

if not uploaded_file:
    st.markdown(
        "<div class='hint' style='border-color:#C4B5FD; background:#F5F3FF;'>"
        "まず写真を選んでください。"
        "</div>",
        unsafe_allow_html=True)

# ── モデル選択（コスト vs 精度） ──
MODEL_OPTIONS = {
    "高精度（おすすめ）　約3円/回": "claude-sonnet-4-6",
    "省エネ　約0.3円/回": "claude-haiku-4-5-20251001",
}
selected_model_label = st.radio(
    "読み取りモード",
    list(MODEL_OPTIONS.keys()),
    index=0,
    horizontal=True,
    help="「省エネ」は速くて安い分、複雑な手書きや細かい文字の読み取り精度が下がる場合があります。",
)
selected_model = MODEL_OPTIONS[selected_model_label]

convert_btn = st.button(
    "スライドに変換する",
    type="primary",
    use_container_width=True,
    disabled=not uploaded_file,
)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 変換処理
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
if convert_btn and not img_data:
    st.error("写真を正しく読み込めませんでした。別の写真を選んでもう一度お試しください。")

if convert_btn and img_data:
    try:
        api_key = get_api_key()
        if not api_key:
            st.warning("認証キーが設定されていません。左上のメニューから設定してください。")
            st.stop()

        # ── ① 準備（瞬時に完了）──
        prog = st.progress(0, text="写真を準備しています…")
        prog.progress(30, text="写真を準備しています…")

        # ── ② AI読み取り（最も時間がかかる → スピナーで安心感を演出）──
        with st.spinner("AIが文字を読み取り中です。そのままお待ちください（通常10〜30秒）…"):
            extracted = analyze_with_claude(img_data, media_type, api_key,
                                            model=selected_model)

        # ── ③ スライド生成（数秒）──
        prog.progress(80, text="スライドを作成しています…")

        is_portrait = st.session_state.get("_is_portrait", False)
        st.session_state["extracted"]   = extracted
        st.session_state["pptx_bytes"]  = generate_pptx(extracted, is_portrait=is_portrait)
        st.session_state["html_bytes"]  = generate_html(extracted)
        # アノテーション画像を生成（番号丸印付き）
        _prev = st.session_state.get("_preview_bytes")
        if _prev:
            st.session_state["_annotated_bytes"] = build_annotated_image(_prev, extracted)
        # 新規変換時は編集キーをクリア（前回の編集内容を引き継がせない）
        for k in [k for k in st.session_state if k.startswith("edit_")]:
            del st.session_state[k]

        prog.progress(100, text="完了しました！")
        prog.empty()
        st.success("変換できました。下にスクロールして保存してください。")

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
    extracted   = st.session_state["extracted"]
    pptx_bytes  = st.session_state.get("pptx_bytes")
    html_bytes  = st.session_state.get("html_bytes")

    st.markdown("---")
    st.markdown(
        "<div class='step'><span class='snum'>３</span>保存する</div>",
        unsafe_allow_html=True)

    elements = extracted.get("elements", [])
    el_count = len(elements)
    st.markdown(
        f"<div class='done-box'>完了 — 読み取り：<strong>{el_count} 件</strong></div>",
        unsafe_allow_html=True)

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # テキスト編集エリア（元画像を見ながら修正）
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    st.markdown("---")
    st.markdown(
        "<div class='step'><span class='snum' style='font-size:0.95rem;'>✏</span>"
        "内容を修正する（任意）</div>",
        unsafe_allow_html=True)
    st.markdown(
        "<div class='hint'>元の写真を確認しながら、読み取り結果を直接修正できます。"
        "修正後は「作り直す」ボタンを押してください。</div>",
        unsafe_allow_html=True)

    # アノテーション画像（番号丸印付き）— エクスパンダーで表示
    ann_bytes = st.session_state.get("_annotated_bytes")
    plain_bytes = st.session_state.get("_preview_bytes")
    disp_bytes = ann_bytes or plain_bytes
    if disp_bytes:
        with st.expander("元の写真を確認する（色帯がフォームの■と対応）", expanded=True):
            st.image(Image.open(io.BytesIO(disp_bytes)), use_container_width=True)

    # 編集フォーム（全幅 — スマホ・PCどちらでも使いやすい）
    TYPE_LABEL  = {"heading": "見出し", "bullet": "箇条書き",
                   "text": "本文", "arrow": "矢印"}
    SHAPE_LABEL = {"rect": "【四角】", "ellipse": "（丸）"}

    # タイトル
    title_key = "edit_title"
    if title_key not in st.session_state:
        st.session_state[title_key] = extracted.get("title", "")
    st.markdown(
        "<p style='font-size:0.88rem; font-weight:700; color:#1E1B4B;"
        " margin:0.5rem 0 0.1rem;'>タイトル</p>",
        unsafe_allow_html=True)
    st.text_input("タイトル", key=title_key, label_visibility="collapsed")

    # 各セクション・アイテム（色インジケーター付き）
    edit_sections = extracted.get("sections") or []
    global_idx = 0
    for si, sec in enumerate(edit_sections):
        sec_heading = (sec.get("heading") or "").strip()
        if sec_heading:
            col_lbl = 1 if sec.get("column", 1) != 2 else 2
            st.markdown(
                f"<p style='margin:0.8rem 0 0.2rem; font-size:0.85rem;"
                f" color:#7C3AED; font-weight:700;'>── {sec_heading}"
                f"（列{col_lbl}）</p>",
                unsafe_allow_html=True)
        for ii, item in enumerate(sec.get("items") or []):
            etype   = item.get("type", "text")
            shape   = item.get("shape", "none")
            lbl     = SHAPE_LABEL.get(shape, "") + TYPE_LABEL.get(etype, "本文")
            chex    = item_color_hex(global_idx)
            ekey    = f"edit_{si}_{ii}"
            if ekey not in st.session_state:
                st.session_state[ekey] = item.get("text", "")
            # 色付きラベル（■ + テキスト種別）
            st.markdown(
                f"<p style='font-size:0.88rem; font-weight:700; color:#1E1B4B;"
                f" margin:0.6rem 0 0.1rem;'>"
                f"<span style='display:inline-block; width:13px; height:13px;"
                f" background:{chex}; border-radius:3px; margin-right:5px;"
                f" vertical-align:middle;'></span>{lbl}</p>",
                unsafe_allow_html=True)
            st.text_input(lbl, key=ekey, label_visibility="collapsed")
            global_idx += 1

    # 作り直しボタン
    if st.button("この内容でスライドを作り直す",
                 type="secondary", use_container_width=True):
        edited = copy.deepcopy(extracted)
        # タイトル更新
        edited["title"] = st.session_state.get("edit_title", edited.get("title", ""))
        # 各アイテムのテキスト更新
        for si, sec in enumerate(edited.get("sections") or []):
            for ii, item in enumerate(sec.get("items") or []):
                ekey = f"edit_{si}_{ii}"
                if ekey in st.session_state:
                    item["text"] = st.session_state[ekey]
        # elements リストを再構築
        new_elements = []
        for sec in edited.get("sections", []):
            for it in (sec.get("items") or []):
                idx = len(new_elements)
                new_elements.append({
                    "id": f"el_{idx:03d}", "type": it.get("type", "text"),
                    "content": it.get("text", ""), "shape": it.get("shape", "none"),
                    "style": {"color": it.get("color", "black"),
                              "bold": it.get("bold", False)},
                    "confidence": 1.0,
                })
        edited["elements"] = new_elements
        # 再生成
        ip = st.session_state.get("_is_portrait", False)
        st.session_state["extracted"]  = edited
        st.session_state["pptx_bytes"] = generate_pptx(edited, is_portrait=ip)
        st.session_state["html_bytes"] = generate_html(edited)
        # アノテーション画像も更新
        _prev = st.session_state.get("_preview_bytes")
        if _prev:
            st.session_state["_annotated_bytes"] = build_annotated_image(_prev, edited)
        st.success("スライドを更新しました。下のボタンから保存してください。")

    st.markdown("---")
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    # ── iPhone 向け HTML ダウンロード（メイン） ──
    if html_bytes:
        st.markdown(
            "<div style='font-size:0.9rem; color:#7C3AED; margin-bottom:0.4rem;"
            " font-weight:600;'>iPhone でそのまま開けます</div>",
            unsafe_allow_html=True)
        st.download_button(
            label="HTMLで保存する（iPhone対応）",
            data=html_bytes,
            file_name=f"パシャッと_{ts}.html",
            mime="text/html",
            use_container_width=True,
        )

    # ── PowerPoint ダウンロード（サブ） ──
    if pptx_bytes:
        st.markdown(
            "<div style='font-size:0.9rem; color:#6B7280; margin:0.8rem 0 0.4rem;"
            " font-weight:600;'>PC・PowerPoint 用</div>",
            unsafe_allow_html=True)
        st.download_button(
            label="PowerPoint で保存する",
            data=pptx_bytes,
            file_name=f"パシャッと_{ts}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )

    with st.expander("読み取り内容を確認する"):
        elmap = {el["id"]: el for el in extracted.get("elements", [])}
        for group in extracted.get("structure", {}).get("groups", []):
            st.markdown(f"**{group.get('label', 'グループ')}**")
            for iid in group.get("items", []):
                el = elmap.get(iid)
                if el:
                    prefix = "■ " if el.get("type") == "heading" else "・"
                    conf   = el.get("confidence", 1.0)
                    st.markdown(
                        f"&nbsp;&nbsp;{prefix} {el['content']} "
                        f"<span style='color:#9CA3AF; font-size:0.85rem;'>（確度 {conf*100:.0f}%）</span>",
                        unsafe_allow_html=True)
        st.markdown("---")
        json_str = json.dumps(
            {"バージョン": "1.0", "変換日時": datetime.now().isoformat(), "データ": extracted},
            ensure_ascii=False, indent=2,
        )
        st.download_button(
            label="テキスト（JSON）で保存する",
            data=json_str,
            file_name=f"パシャッと_テキスト_{ts}.json",
            mime="application/json",
            use_container_width=True,
        )

# ── フッター ──
st.markdown("<br>", unsafe_allow_html=True)
st.markdown(
    "<div style='text-align:center; color:#A78BFA; font-size:0.85rem; line-height:2;'>"
    "AI文字認識による自動変換サービス<br>&copy; 2026 パシャッと"
    "</div>",
    unsafe_allow_html=True)
