"""
パシャッと — スマホ・高齢者対応ユニバーサルデザイン版
ホワイトボード・手書きメモ → スライド自動変換
"""

import os, io, json, base64, re, math
from datetime import datetime

import numpy as np
import streamlit as st
from PIL import Image, ImageChops, ImageEnhance, ImageFilter, ImageOps

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
/* ── Streamlit ヘッダー・フッター・各種リンクを非表示 ── */
header[data-testid="stHeader"] { display: none !important; }
#MainMenu { display: none !important; }
footer { display: none !important; }
[data-testid="stToolbar"] { display: none !important; }
[data-testid="stDecoration"] { display: none !important; }
/* GitHub・Streamlit ロゴリンク */
a[href*="github.com"] { display: none !important; }
a[href*="streamlit.io"] { display: none !important; }
.viewerBadge_container__r5tak { display: none !important; }
.stActionButton { display: none !important; }

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


def fix_orientation(img: Image.Image) -> Image.Image:
    """EXIFの向き情報を使って自動回転（スマホ写真の天地補正）"""
    try:
        img = ImageOps.exif_transpose(img)
    except Exception:
        pass
    return img


def deskew_image(img: Image.Image) -> Image.Image:
    """投影プロファイル法で傾きを検出・補正する（numpy高速版）"""
    try:
        # 400px縮小版で角度検出（高速化）
        thumb = img.copy()
        thumb.thumbnail((400, 400), Image.LANCZOS)
        gray   = np.array(thumb.convert("L"), dtype=np.float32)
        binary = (gray < 128).astype(np.float32)
        h, w   = binary.shape
        ys     = np.arange(h, dtype=np.float32)
        xs     = np.arange(w, dtype=np.float32)

        best_angle = 0.0
        best_score = -1.0
        for angle in np.arange(-10, 10.5, 1.0):   # 1°刻みで高速探索
            rad = math.radians(angle)
            ny  = (ys[:, None] * math.cos(rad)
                   - xs[None, :] * math.sin(rad))
            ny_int = np.clip(ny.astype(np.int32) + w, 0, h + w - 1)
            proj   = np.bincount(ny_int[binary > 0].ravel(), minlength=h + w)
            score  = float(np.var(proj))
            if score > best_score:
                best_score = score
                best_angle = angle

        if abs(best_angle) > 0.5:
            img = img.rotate(-best_angle, expand=True,
                             fillcolor=(255, 255, 255), resample=Image.BICUBIC)
    except Exception:
        pass
    return img


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
    img = ImageEnhance.Contrast(img).enhance(1.5)
    img = ImageEnhance.Sharpness(img).enhance(2.0)
    img = ImageEnhance.Brightness(img).enhance(1.1)
    return img


def preprocess_image(img: Image.Image, do_trim: bool = True) -> Image.Image:
    """全前処理をまとめて実行: EXIF回転 → トリミング → デスキュー → 画質強化"""
    img = fix_orientation(img)
    if do_trim:
        img = auto_trim(img)
    img = deskew_image(img)
    return img


def pil_to_base64(img: Image.Image, max_px: int = 2048) -> tuple[str, str]:
    """PIL Image → base64（高解像度維持・OCR前処理済み）"""
    img = enhance_for_ocr(img)
    if img.width > max_px or img.height > max_px:
        img.thumbnail((max_px, max_px), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=92)
    return base64.standard_b64encode(buf.getvalue()).decode("utf-8"), "image/jpeg"


EXTRACTION_PROMPT = """
あなたはOCRと文書レイアウト解析の専門家です。
添付の画像（ホワイトボード・手書きメモ・印刷物など）を精密に分析し、
アナログ文書をデジタルで忠実に再現するためのすべての情報を抽出してください。

【読み取り対象】
1. すべての文字・数字・記号（かすれ・薄い・小さい文字も含む）
2. 文字の色（黒・赤・青・緑・オレンジ・紫・ピンク・茶色・グレー等）
3. 文字スタイル（太字・下線・丸囲み・二重線・取り消し線）
4. ページ上の位置（上段/中段/下段 × 左/中央/右 の9ゾーン）
5. 表・罫線の構造（ヘッダー行・データ行・列の区分け）
6. 矢印・接続線・囲み図形
7. 文字と文字の間隔（密集/通常/広め）
8. 画像全体の傾き（補正が必要な角度）

以下のJSON形式のみで返答してください（前置き・説明文・コードブロック不要）:

{
  "title": "文書のタイトルまたは主題（画像から推定）",
  "orientation_degrees": 0,
  "elements": [
    {
      "id": "el_001",
      "type": "heading",
      "content": "読み取ったテキスト",
      "position": { "zone_row": 0, "zone_col": 0 },
      "style": {
        "color": "black",
        "bold": false,
        "underline": false,
        "circled": false,
        "strikethrough": false,
        "size": "large"
      },
      "level": 1,
      "confidence": 0.95
    }
  ],
  "tables": [
    {
      "id": "tbl_001",
      "position": { "zone_row": 1, "zone_col": 1 },
      "headers": ["列名1", "列名2"],
      "rows": [["データ1", "データ2"], ["データ3", "データ4"]]
    }
  ],
  "structure": {
    "type": "list",
    "groups": [
      { "label": "グループ名", "items": ["el_001", "el_002"] }
    ]
  },
  "language": "ja",
  "notes": "読み取れなかった箇所や補足"
}

【フィールドの定義】
- orientation_degrees: 画像を正位置にするために必要な回転角度（0/90/180/270）
- zone_row: 0=上段, 1=中段, 2=下段
- zone_col: 0=左, 1=中央, 2=右
- type: heading（見出し）/ bullet（箇条書き）/ text（本文）/ table（表）/ arrow（矢印・接続）
- color: black / red / blue / green / orange / purple / pink / gray / brown / yellow
- size: large（大）/ medium（中）/ small（小）
- tablesがない場合は空配列 [] を返すこと
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
    """解析データからパワーポイントを生成（色・スタイル・表・ゾーン配置対応）"""
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from pptx.oxml.ns import qn
    from lxml import etree

    # ─ カラーパレット（スライド背景用） ─
    C_DARK   = RGBColor(0x0D, 0x1B, 0x2A)
    C_NAVY   = RGBColor(0x1B, 0x28, 0x38)
    C_ACCENT = RGBColor(0x00, 0xB4, 0xD8)
    C_WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
    C_MUTED  = RGBColor(0x94, 0xA3, 0xB8)
    C_LIGHT  = RGBColor(0xCB, 0xD5, 0xE1)

    # ─ 文字色マッピング（手書きの色 → RGBColor） ─
    TEXT_COLORS = {
        "black":  RGBColor(0xF0, 0xF4, 0xF8),  # 背景が暗いので白系
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

    def resolve_color(color_name: str) -> RGBColor:
        return TEXT_COLORS.get((color_name or "black").lower(), C_LIGHT)

    def pt_size(size_str: str) -> int:
        return {"large": 16, "medium": 13, "small": 11}.get(size_str or "medium", 13)

    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    def bg(slide, color):
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = color

    def rect(slide, x, y, w, h, fc, lc=None):
        sh = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
        sh.fill.solid(); sh.fill.fore_color.rgb = fc
        if lc: sh.line.color.rgb = lc
        else:  sh.line.fill.background()

    def txt(slide, text, x, y, w, h, size=13, bold=False, color=None,
            align=PP_ALIGN.LEFT, wrap=True, underline=False, italic=False):
        color = color or C_LIGHT
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame; tf.word_wrap = wrap
        p = tf.paragraphs[0]; p.alignment = align
        r = p.add_run(); r.text = text
        r.font.size      = Pt(size)
        r.font.bold      = bold
        r.font.underline = underline
        r.font.italic    = italic
        r.font.color.rgb = color
        r.font.name      = "Noto Sans JP"

    def add_table_shape(slide, tbl_data, x, y, max_w, max_h):
        """python-pptxで実際の表を作成"""
        headers = tbl_data.get("headers", [])
        rows    = tbl_data.get("rows", [])
        if not rows and not headers:
            return
        n_rows = len(rows) + (1 if headers else 0)
        n_cols = max(
            len(headers),
            max((len(r) for r in rows), default=1)
        )
        if n_rows == 0 or n_cols == 0:
            return
        col_w = min(max_w / n_cols, 2.5)
        row_h = min(max_h / n_rows, 0.45)
        tbl = slide.shapes.add_table(
            n_rows, n_cols,
            Inches(x), Inches(y),
            Inches(col_w * n_cols), Inches(row_h * n_rows)
        ).table
        # ヘッダー行
        if headers:
            for ci, h in enumerate(headers[:n_cols]):
                cell = tbl.cell(0, ci)
                cell.text = str(h)
                cell.fill.solid(); cell.fill.fore_color.rgb = C_ACCENT
                for para in cell.text_frame.paragraphs:
                    for run in para.runs:
                        run.font.bold = True
                        run.font.size = Pt(11)
                        run.font.color.rgb = C_DARK
        # データ行
        for ri, row in enumerate(rows):
            tr = ri + (1 if headers else 0)
            for ci, val in enumerate(row[:n_cols]):
                cell = tbl.cell(tr, ci)
                cell.text = str(val)
                bg_col = RGBColor(0x1E, 0x30, 0x40) if ri % 2 == 0 else C_NAVY
                cell.fill.solid(); cell.fill.fore_color.rgb = bg_col
                for para in cell.text_frame.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(10)
                        run.font.color.rgb = C_LIGHT

    # ━━━━━ スライド作成 ━━━━━
    s = prs.slides.add_slide(blank)
    bg(s, C_DARK)
    rect(s, 0, 0, 13.33, 0.07, C_ACCENT)

    # タイトル
    title = data.get("title", "ホワイトボードメモ")
    txt(s, title, 0.4, 0.10, 12.0, 0.72, size=24, bold=True, color=C_WHITE)

    # 日時
    ts = datetime.now().strftime("%Y年%m月%d日 %H:%M")
    txt(s, f"変換日時：{ts}", 0.4, 0.78, 6.0, 0.32, size=9, color=C_MUTED)

    # コンテンツ背景
    rect(s, 0.3, 1.15, 12.73, 5.7, C_NAVY)

    # ─ 要素を zone_row × zone_col でソートして読み取り順を保持 ─
    elements = data.get("elements", [])

    def sort_key(el):
        pos = el.get("position") or {}
        return (pos.get("zone_row", 1), pos.get("zone_col", 0))

    sorted_els = sorted(elements, key=sort_key)

    # ─ 表は別途処理 ─
    tables_data = data.get("tables", [])

    # ─ zone_col で左右カラムに分割 ─
    col_left  = [e for e in sorted_els if (e.get("position") or {}).get("zone_col", 0) <= 1]
    col_right = [e for e in sorted_els if (e.get("position") or {}).get("zone_col", 2) == 2]
    # 右カラムが少ない場合は左カラムを分割
    if len(col_right) == 0 and len(col_left) > 3:
        mid       = len(col_left) // 2
        col_right = col_left[mid:]
        col_left  = col_left[:mid]

    # ─ 利用可能な高さから1要素あたりの行高を自動計算 ─
    Y_START = 1.25
    Y_MAX   = 6.65
    AVAIL_H = Y_MAX - Y_START          # 5.4 インチ
    max_col = max(len(col_left), len(col_right), 1)
    # 1行あたりの高さ：最小0.28"、最大0.52"
    step_h  = max(0.28, min(0.52, AVAIL_H / max_col))
    # フォントサイズをステップ幅から逆算（比例縮小）
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
            # スタイル指定のサイズをベースptに対して相対調整
            size_map = {"large": base_pt + 2, "medium": base_pt, "small": base_pt - 2}
            size = size_map.get(style.get("size", "medium"), base_pt)
            row_h = step_h

            if etype == "heading":
                size = min(size + 2, base_pt + 3)
                label = f"【{content}】" if circled else content
                txt(s, label, x_start, y, col_w, row_h,
                    size=size, bold=True, color=color, underline=underl)
            elif etype == "arrow":
                txt(s, f"→ {content}", x_start + 0.1, y, col_w - 0.1, row_h,
                    size=max(size - 1, 8), color=RGBColor(0xFF, 0xB7, 0x4D), italic=True)
            else:
                prefix = "○ " if circled else "• "
                txt(s, f"{prefix}{content}", x_start + 0.1, y, col_w - 0.1, row_h,
                    size=size, bold=bold, color=color, underline=underl)

            y += row_h

    render_column(col_left,  0.45, 6.0)
    render_column(col_right, 6.75, 6.0)

    # ─ 表を下部またはゾーン位置に配置 ─
    tbl_y = 1.25
    tbl_x = 0.45
    for tbl_data in tables_data:
        pos    = tbl_data.get("position") or {}
        row_z  = pos.get("zone_row", 2)
        col_z  = pos.get("zone_col", 0)
        tbl_y  = 1.25 + row_z * 1.8
        tbl_x  = 0.45 + col_z * 4.3
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

# ── ロゴ・説明 ──
st.markdown("""
<div class='logo'>パシャッ<em>と</em></div>
<div class='tagline'>メモの写真を、スライドに変換します</div>
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
        "✂️　余白を自動でカット（おすすめ）",
        value=True,
        help="余白を除去して読み取り精度を上げます。傾き・天地も自動補正します。",
    )

    display_img = preprocess_image(pil_image, do_trim=do_trim)
    st.image(display_img, caption="変換する写真", use_container_width=True)

    img_data, media_type = pil_to_base64(display_img)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ステップ２　スライドに変換する
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
st.markdown("---")
st.markdown(
    "<div class='step'><span class='snum'>２</span>変換する</div>",
    unsafe_allow_html=True)

can_convert = bool(raw_input)

if not can_convert:
    st.markdown(
        "<div class='hint' style='border-color:#F5A623; background:#FFF8E7;'>"
        "↑ まず写真を選んでください。"
        "</div>",
        unsafe_allow_html=True)

convert_btn = st.button(
    "✨　スライドに変換する",
    type="primary",
    use_container_width=True,
    disabled=not can_convert,
)

st.markdown(
    "<p style='text-align:center; color:#7A9AAD; font-size:1.05rem; margin:0.8rem 0 0.4rem;'>"
    "— または —"
    "</p>",
    unsafe_allow_html=True)

demo_btn = st.button(
    "🎯　サンプルで試す",
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

        status_text.markdown("**② 写真を読み取っています…**")
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
        "<div class='step'><span class='snum'>３</span>保存する</div>",
        unsafe_allow_html=True)

    el_count = len(extracted.get("elements", []))
    demo_note = "（サンプル）" if is_demo else ""

    st.markdown(f"""
<div class='done-box'>
✅ 完了{demo_note}　読み取り：<strong>{el_count} 件</strong>
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
            {"バージョン": "1.0",
             "変換日時": datetime.now().isoformat(),
             "データ": extracted},
            ensure_ascii=False, indent=2
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
