"""
SnapDeck — スマホ対応・シンプル版
ホワイトボード・手書きメモ → スライド自動変換
"""

import os, io, json, base64, re, zipfile
from datetime import datetime

import streamlit as st
from PIL import Image, ImageChops

# ─── ページ設定 ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="パシャッと",
    page_icon="📸",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ─── CSS（スマホ・高齢者向け） ─────────────────────────────────────────────────
st.markdown("""
<style>
/* 背景 */
.stApp { background: #F2F5F8; }
.main .block-container { padding: 1.2rem 1rem 4rem; max-width: 540px; }

/* タイトル */
.logo { text-align:center; font-size:2.6rem; font-weight:900;
        color:#0D1B2A; margin-bottom:0.1rem; line-height:1.2; }
.logo em { color:#0099BB; font-style:normal; }
.tagline { text-align:center; color:#607D8B; font-size:1.1rem; margin-bottom:1.8rem; }

/* ステップラベル */
.step { display:flex; align-items:center; gap:0.6rem;
        font-size:1.3rem; font-weight:700; color:#0D1B2A;
        margin: 1.6rem 0 0.5rem; }
.snum { background:#0099BB; color:#fff; border-radius:50%;
        width:34px; height:34px; line-height:34px;
        text-align:center; font-size:1rem; font-weight:900; flex-shrink:0; }

/* ヒントテキスト */
.hint { color:#5A7080; font-size:1.05rem; margin:0 0 0.8rem; line-height:1.6; }

/* ボタン全般（大きく） */
.stButton > button {
    font-size: 1.25rem !important;
    min-height: 64px !important;
    border-radius: 14px !important;
    font-weight: 700 !important;
    width: 100% !important;
    letter-spacing: 0.02em !important;
}
.stButton > button[kind="primary"] {
    background: #0099BB !important;
    color: white !important;
    border: none !important;
    box-shadow: 0 4px 12px rgba(0,153,187,0.3) !important;
}
.stButton > button[kind="secondary"] {
    background: #E8F0F5 !important;
    color: #334155 !important;
    border: 1px solid #CBD5E0 !important;
}
.stButton > button:disabled {
    opacity: 0.45 !important;
}

/* ダウンロードボタン */
.stDownloadButton > button {
    background: #10B981 !important;
    color: white !important;
    font-size: 1.25rem !important;
    min-height: 64px !important;
    border-radius: 14px !important;
    font-weight: 700 !important;
    width: 100% !important;
    border: none !important;
    box-shadow: 0 4px 12px rgba(16,185,129,0.3) !important;
    letter-spacing: 0.02em !important;
}

/* 完了ボックス */
.done-box {
    background: #D1FAE5; border: 2px solid #10B981;
    border-radius: 14px; padding: 1.3rem 1.4rem;
    font-size: 1.15rem; color: #065F46; font-weight: 600;
    margin: 1rem 0; line-height: 1.7;
}

/* タブ */
.stTabs [data-baseweb="tab"] {
    font-size: 1.1rem !important;
    font-weight: 600 !important;
    min-height: 52px !important;
}

/* ファイルアップローダー */
[data-testid="stFileUploaderDropzone"] {
    min-height: 130px;
    border-radius: 12px !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] p {
    font-size: 1rem !important;
}

/* トグル */
.stToggle label { font-size: 1.05rem !important; font-weight: 600 !important; }

/* expander */
.stExpander summary p { font-size: 1.05rem !important; }

/* divider */
hr { margin: 1.2rem 0 !important; border-color: #D1D9E0 !important; }
</style>
""", unsafe_allow_html=True)


# ─── バックエンド関数 ──────────────────────────────────────────────────────────

def get_api_key() -> str:
    """APIキーをSecretsまたはサイドバーから取得"""
    try:
        return st.secrets["ANTHROPIC_API_KEY"]
    except Exception:
        pass
    env_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if env_key:
        return env_key
    return st.session_state.get("api_key_input", "")


def auto_trim(img: Image.Image, margin: int = 30) -> Image.Image:
    """白い余白を自動でカット（精度向上）"""
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


def pil_to_base64(img: Image.Image, max_px: int = 1600) -> tuple[str, str]:
    """PIL Image → base64（大きすぎる画像は縮小してAPIエラーを防ぐ）"""
    # スマホ写真は4000px超になることがあるため最大 max_px にリサイズ
    if img.width > max_px or img.height > max_px:
        img.thumbnail((max_px, max_px), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=88)
    return base64.standard_b64encode(buf.getvalue()).decode("utf-8"), "image/jpeg"


EXTRACTION_PROMPT = """
あなたは画像解析の専門家です。与えられた画像（ホワイトボードや手書きメモ）を分析し、
以下のJSON形式で情報を抽出してください。

出力するJSONの形式:
{
  "title": "メモ全体のタイトルまたは主題（推定）",
  "elements": [
    {
      "id": "el_001",
      "type": "text|heading|bullet|table_row",
      "content": "テキスト内容",
      "level": 1,
      "confidence": 0.95
    }
  ],
  "structure": {
    "type": "list|flow|table|freeform",
    "groups": [
      {
        "label": "グループ名（推定）",
        "items": ["el_001", "el_002"]
      }
    ]
  },
  "language": "ja|en|mixed",
  "notes": "補足コメント"
}

ルール:
- 読み取れるテキストはすべて抽出してください
- JSONのみを返してください（説明文は不要）
"""


@st.cache_data(show_spinner=False)
def analyze_with_claude(img_data: str, media_type: str, api_key: str) -> dict:
    """Claude Vision APIで画像を解析（結果をキャッシュ）"""
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
    text = re.sub(r'\s*```$', '', text, flags=re.MULTILINE)
    text = text.strip()
    # { ... } の部分だけ抽出（前後に余分な文字があっても安全に処理）
    match = re.search(r'\{.*\}', text, re.DOTALL)
    raw = match.group() if match else text
    return json.loads(raw)


def get_demo_data() -> dict:
    return {
        "title": "新サービス企画 ブレインストーミング",
        "elements": [
            {"id": "el_001", "type": "heading", "content": "ターゲットユーザー",             "level": 1, "confidence": 0.97},
            {"id": "el_002", "type": "bullet",  "content": "20〜40代ビジネスパーソン",        "level": 2, "confidence": 0.95},
            {"id": "el_003", "type": "bullet",  "content": "会議が多い職種（営業・企画・PM）", "level": 2, "confidence": 0.93},
            {"id": "el_004", "type": "heading", "content": "コア機能",                         "level": 1, "confidence": 0.98},
            {"id": "el_005", "type": "bullet",  "content": "写真1枚でスライド自動生成",        "level": 2, "confidence": 0.97},
            {"id": "el_006", "type": "bullet",  "content": "多言語OCR対応",                   "level": 2, "confidence": 0.94},
            {"id": "el_007", "type": "bullet",  "content": "PPTX / JSON 出力",                "level": 2, "confidence": 0.96},
            {"id": "el_008", "type": "heading", "content": "マネタイズ案",                    "level": 1, "confidence": 0.95},
            {"id": "el_009", "type": "bullet",  "content": "フリーミアム（月10枚無料）",      "level": 2, "confidence": 0.93},
            {"id": "el_010", "type": "bullet",  "content": "Pro: ¥1,980/月",                  "level": 2, "confidence": 0.91},
        ],
        "structure": {
            "type": "list",
            "groups": [
                {"label": "ターゲットユーザー", "items": ["el_001", "el_002", "el_003"]},
                {"label": "コア機能",           "items": ["el_004", "el_005", "el_006", "el_007"]},
                {"label": "マネタイズ案",        "items": ["el_008", "el_009", "el_010"]},
            ]
        },
        "language": "ja",
        "notes": "これはサンプルデータです。実際の画像を使うと本物の解析結果が得られます。"
    }


def generate_pptx(data: dict) -> bytes:
    """解析データからPowerPointを生成"""
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

    # タイトルスライド
    s = prs.slides.add_slide(blank)
    bg(s, C_DARK)
    rect(s, 0, 0, 0.08, 7.5, C_ACCENT)
    title = data.get("title", "ホワイトボードメモ")
    txt(s, title,                      0.3, 1.2, 12.7, 1.6, size=40, bold=True)
    txt(s, "パシャッと 自動変換レポート", 0.3, 2.9,  8.0, 0.6, size=18, color=C_ACCENT)
    ts = datetime.now().strftime("%Y年%m月%d日 %H:%M")
    txt(s, f"変換日時: {ts}",           0.3, 3.6,  8.0, 0.5, size=13, color=C_MUTED)
    if data.get("notes"):
        txt(s, f"📝 {data['notes']}",   0.3, 6.7, 12.7, 0.55, size=11, color=C_MUTED)

    # グループ別スライド
    elements_by_id = {el["id"]: el for el in data.get("elements", [])}
    groups = data.get("structure", {}).get("groups", [])
    if not groups:
        groups = [{"label": "内容", "items": [el["id"] for el in data.get("elements", [])]}]

    for group in groups:
        s = prs.slides.add_slide(blank)
        bg(s, C_DARK)
        rect(s, 0, 0, 13.33, 0.08, C_ACCENT)
        txt(s, group.get("label", "内容"), 0.4, 0.18, 12.5, 0.72, size=28, bold=True)
        rect(s, 0.3, 1.0, 12.73, 6.1, C_NAVY)
        y = 1.15
        for item_id in group.get("items", []):
            el = elements_by_id.get(item_id)
            if not el: continue
            if el.get("type") == "heading":
                if y > 1.15: y += 0.08
                txt(s, el["content"], 0.55, y, 12.2, 0.52, size=18, bold=True, color=C_ACCENT)
                y += 0.6
            else:
                txt(s, "  •  " + el["content"], 0.55, y, 12.1, 0.45, size=14, color=C_LIGHT)
                y += 0.48
            if y > 6.7: break
        rect(s, 0, 7.15, 13.33, 0.35, C_NAVY)
        txt(s, "パシャッと  |  自動生成", 0.4, 7.15, 12.9, 0.35,
            size=10, color=C_MUTED, align=PP_ALIGN.RIGHT)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ─── UI ────────────────────────────────────────────────────────────────────────

# タイトル
st.markdown("""
<div class='logo'>パシャッ<em>と</em></div>
<div class='tagline'>ホワイトボードの写真を、すぐにスライドへ</div>
""", unsafe_allow_html=True)

# APIキー — Secretsになければサイドバーに入力欄を表示
if not get_api_key():
    with st.sidebar:
        st.markdown("### 🔑 APIキー設定")
        key_in = st.text_input("Anthropic APIキー", type="password",
                               placeholder="sk-ant-api03-...")
        if key_in:
            st.session_state["api_key_input"] = key_in
            st.success("✅ 設定しました")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# STEP 1 ── 写真を撮る・選ぶ
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
st.markdown(
    "<div class='step'><span class='snum'>1</span>写真を撮る・選ぶ</div>",
    unsafe_allow_html=True)
st.markdown(
    "<p class='hint'>ホワイトボードや手書きメモを、できるだけ正面から撮影してください。</p>",
    unsafe_allow_html=True)

tab_cam, tab_file = st.tabs(["📷　カメラで撮影", "📁　ファイルを選択"])

with tab_cam:
    # facing_mode パラメータは Streamlit 1.42+ 以降でのみ対応。
    # Streamlit Cloud の実行バージョンに合わせて try/except でフォールバック。
    try:
        camera_photo = st.camera_input(
            "カメラを起動",
            facing_mode="environment",
            help="ブラウザのカメラ許可が必要です（初回のみポップアップが表示されます）",
            label_visibility="collapsed",
        )
    except TypeError:
        camera_photo = st.camera_input(
            "カメラを起動",
            help="ブラウザのカメラ許可が必要です（初回のみポップアップが表示されます）",
            label_visibility="collapsed",
        )

with tab_file:
    uploaded_file = st.file_uploader(
        "画像ファイルを選択",
        type=["jpg", "jpeg", "png", "webp"],
        label_visibility="collapsed",
        help="JPG・PNG・WebP に対応しています",
    )

# カメラ優先で統合
raw_input = camera_photo or uploaded_file

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 画像プレビュー ＆ 余白カット
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
pil_image = None
img_data  = None
media_type = "image/png"

if raw_input:
    pil_image = Image.open(raw_input)
    if pil_image.mode not in ("RGB",):
        pil_image = pil_image.convert("RGB")

    st.markdown("---")

    do_trim = st.toggle(
        "✂️  余白を自動でカットする（文字の精度が上がります）",
        value=True,
        help="写真の周囲にある白い余白や不要な部分を自動で除去します。",
    )

    display_img = auto_trim(pil_image) if do_trim else pil_image
    st.image(display_img, caption="変換する画像", use_container_width=True)

    # base64 に変換（変換ボタン押下時に使用）
    img_data, media_type = pil_to_base64(display_img)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# STEP 2 ── 変換する
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
st.markdown("---")
st.markdown(
    "<div class='step'><span class='snum'>2</span>スライドに変換する</div>",
    unsafe_allow_html=True)

can_convert = bool(raw_input)
convert_btn = st.button(
    "🚀　変換する",
    type="primary",
    use_container_width=True,
    disabled=not can_convert,
    help="写真を撮影またはファイルを選択してから押してください" if not can_convert else "",
)

if not can_convert:
    st.markdown(
        "<p style='text-align:center;color:#9CA3AF;font-size:0.95rem;margin-top:0.3rem;'>"
        "↑ まず写真を撮ってから押してください</p>",
        unsafe_allow_html=True)

st.markdown(
    "<p style='text-align:center;color:#9CA3AF;font-size:1rem;margin:0.8rem 0;'>または</p>",
    unsafe_allow_html=True)

demo_btn = st.button(
    "🎯　サンプルで試してみる（写真なしでOK）",
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
    with st.spinner("🤖 AIが文字を読み取っています… （10〜20秒かかります）"):
        try:
            api_key = get_api_key()
            if api_key:
                extracted = analyze_with_claude(img_data, media_type, api_key)
            else:
                st.warning("⚠️ APIキーが見つかりません。左上のメニューから設定してください。")
                extracted = get_demo_data()
            st.session_state["extracted"]  = extracted
            st.session_state["pptx_bytes"] = generate_pptx(extracted)
        except Exception as e:
            err_msg = str(e)
            if "api_key" in err_msg.lower() or "authentication" in err_msg.lower() or "401" in err_msg:
                st.error("❌ APIキーが正しくありません。左上のメニューから確認してください。")
            elif "model" in err_msg.lower() or "invalid" in err_msg.lower():
                st.error(f"❌ モデルエラー: {err_msg}")
            elif "JSONDecodeError" in type(e).__name__ or "json" in err_msg.lower():
                st.error("❌ AIの返答を解析できませんでした。もう一度お試しください。")
            elif "timeout" in err_msg.lower() or "connection" in err_msg.lower():
                st.error("❌ 通信エラーが発生しました。インターネット接続を確認してください。")
            elif "overloaded" in err_msg.lower() or "529" in err_msg:
                st.error("❌ AIサーバーが混み合っています。しばらく待ってから再度お試しください。")
            else:
                st.error(f"❌ エラーが発生しました: {err_msg}")
            st.info("💡 解決しない場合は「サンプルで試してみる」でアプリの動作確認ができます。")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# STEP 3 ── ダウンロード
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
if "extracted" in st.session_state:
    extracted  = st.session_state["extracted"]
    pptx_bytes = st.session_state.get("pptx_bytes")
    is_demo    = st.session_state.get("is_demo", False)

    st.markdown("---")
    st.markdown(
        "<div class='step'><span class='snum'>3</span>スライドをダウンロード</div>",
        unsafe_allow_html=True)

    el_count  = len(extracted.get("elements", []))
    grp_count = len(extracted.get("structure", {}).get("groups", []))
    demo_note = "（サンプルデータ）" if is_demo else ""

    st.markdown(f"""
<div class='done-box'>
✅ 変換完了{demo_note}<br>
検出した項目：<strong>{el_count} 件</strong>　グループ：<strong>{grp_count} 件</strong>
</div>
""", unsafe_allow_html=True)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    if pptx_bytes:
        st.download_button(
            label="⬇️　PowerPoint をダウンロード",
            data=pptx_bytes,
            file_name=f"snapdeck_{ts}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )

    # JSON／詳細は折りたたみで提供
    with st.expander("📋 テキストデータ・詳細を見る"):
        json_str = json.dumps(
            {"snapdeck_version": "1.0",
             "exported_at": datetime.now().isoformat(),
             "data": extracted},
            ensure_ascii=False, indent=2
        )
        st.download_button(
            label="⬇️　JSON をダウンロード",
            data=json_str,
            file_name=f"snapdeck_{ts}.json",
            mime="application/json",
            use_container_width=True,
        )
        st.markdown("---")
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
                        f"<span style='color:#9CA3AF;font-size:0.85rem;'>（確度 {conf*100:.0f}%）</span>",
                        unsafe_allow_html=True)

# フッター
st.markdown("<br>", unsafe_allow_html=True)
st.markdown(
    "<div style='text-align:center;color:#9CA3AF;font-size:0.85rem;'>"
    "Powered by Claude Vision API &nbsp;|&nbsp; © 2026 パシャッと</div>",
    unsafe_allow_html=True)
