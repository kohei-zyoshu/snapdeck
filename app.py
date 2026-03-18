"""
SnapDeck Web App — Streamlit Cloud対応版
ホワイトボード・手書きメモ → PowerPoint / JSON 自動変換
"""

import os
import io
import json
import base64
import re
import zipfile
from datetime import datetime

import streamlit as st

# ─── ページ設定 ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="SnapDeck",
    page_icon="📸",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── カスタムCSS ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* 全体背景 */
    .stApp { background-color: #0D1B2A; }
    .main .block-container { padding-top: 2rem; }

    /* サイドバー */
    section[data-testid="stSidebar"] { background-color: #1B2838; }

    /* タイトル */
    .snap-title {
        font-size: 2.8rem; font-weight: 900;
        color: #FFFFFF; line-height: 1.1; margin-bottom: 0;
    }
    .snap-accent { color: #00B4D8; }
    .snap-subtitle {
        font-size: 1.05rem; color: #94A3B8; margin-top: 0.3rem; margin-bottom: 1.5rem;
    }

    /* カード */
    .snap-card {
        background: #1B2838; border-radius: 10px;
        padding: 1.2rem 1.4rem; margin-bottom: 1rem;
        border: 1px solid #2D3F55;
    }
    .snap-card h4 { color: #00B4D8; margin: 0 0 0.4rem 0; font-size: 1rem; }
    .snap-card p  { color: #CBD5E1; margin: 0; font-size: 0.9rem; line-height: 1.5; }

    /* ステップバッジ */
    .step-badge {
        display: inline-block; background: #00B4D8;
        color: #0D1B2A; border-radius: 50%;
        width: 28px; height: 28px; line-height: 28px;
        text-align: center; font-weight: 900; font-size: 0.9rem;
        margin-right: 8px;
    }

    /* 結果ボックス */
    .result-box {
        background: #1E3448; border-left: 4px solid #00B4D8;
        border-radius: 6px; padding: 1rem 1.2rem; margin: 0.5rem 0;
    }
    .result-box .label { color: #94A3B8; font-size: 0.8rem; text-transform: uppercase; }
    .result-box .value { color: #FFFFFF; font-size: 1.05rem; font-weight: 600; }

    /* ダウンロードボタン上書き */
    .stDownloadButton button {
        background-color: #00B4D8 !important; color: #0D1B2A !important;
        font-weight: 700 !important; border: none !important;
        border-radius: 8px !important; width: 100%;
    }
    .stDownloadButton button:hover { background-color: #0094B3 !important; }

    /* ファイルアップローダー */
    .stFileUploader { border-radius: 10px; }

    /* 成功メッセージ */
    .stSuccess { background-color: #1E3448 !important; color: #10B981 !important; }

    /* コードブロック */
    .stCode { background-color: #1B2838 !important; }
</style>
""", unsafe_allow_html=True)

# ─── ユーティリティ ────────────────────────────────────────────────────────────

def get_api_key() -> str:
    """APIキーをSecretsまたはサイドバーから取得"""
    # Streamlit Secretsから取得（デプロイ時）
    try:
        return st.secrets["ANTHROPIC_API_KEY"]
    except Exception:
        pass
    # 環境変数から取得
    env_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if env_key:
        return env_key
    # サイドバー入力から取得
    return st.session_state.get("api_key_input", "")


def image_to_base64(uploaded_file) -> tuple[str, str]:
    """アップロードされた画像またはカメラ撮影をBase64に変換"""
    name = getattr(uploaded_file, "name", "camera.png")
    ext = name.rsplit(".", 1)[-1].lower()
    media_map = {"jpg": "image/jpeg", "jpeg": "image/jpeg",
                 "png": "image/png", "gif": "image/gif", "webp": "image/webp"}
    media_type = media_map.get(ext, "image/png")
    data = base64.standard_b64encode(uploaded_file.read()).decode("utf-8")
    return data, media_type


# ─── AI解析 ────────────────────────────────────────────────────────────────────

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
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {
                    "type": "base64", "media_type": media_type, "data": img_data
                }},
                {"type": "text", "text": EXTRACTION_PROMPT}
            ]
        }]
    )
    text = message.content[0].text.strip()
    match = re.search(r'\{.*\}', text, re.DOTALL)
    return json.loads(match.group() if match else text)


def get_demo_data() -> dict:
    return {
        "title": "新サービス企画 ブレインストーミング",
        "elements": [
            {"id": "el_001", "type": "heading", "content": "ターゲットユーザー", "level": 1, "confidence": 0.97},
            {"id": "el_002", "type": "bullet",  "content": "20〜40代ビジネスパーソン",        "level": 2, "confidence": 0.95},
            {"id": "el_003", "type": "bullet",  "content": "会議が多い職種（営業・企画・PM）", "level": 2, "confidence": 0.93},
            {"id": "el_004", "type": "heading", "content": "コア機能",                          "level": 1, "confidence": 0.98},
            {"id": "el_005", "type": "bullet",  "content": "写真1枚でスライド自動生成",        "level": 2, "confidence": 0.97},
            {"id": "el_006", "type": "bullet",  "content": "多言語OCR対応",                    "level": 2, "confidence": 0.94},
            {"id": "el_007", "type": "bullet",  "content": "PPTX / JSON / Figma 出力",          "level": 2, "confidence": 0.96},
            {"id": "el_008", "type": "heading", "content": "マネタイズ案",                     "level": 1, "confidence": 0.95},
            {"id": "el_009", "type": "bullet",  "content": "フリーミアム（月10枚無料）",       "level": 2, "confidence": 0.93},
            {"id": "el_010", "type": "bullet",  "content": "Pro: ¥1,980/月",                   "level": 2, "confidence": 0.91},
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
        "notes": "これはデモデータです。APIキーを設定すると実際の画像を解析できます。"
    }


# ─── PPTX生成 ──────────────────────────────────────────────────────────────────

def generate_pptx(data: dict) -> bytes:
    """解析データからPowerPointを生成してbytesで返す"""
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

    # ── タイトルスライド ──
    s = prs.slides.add_slide(blank)
    bg(s, C_DARK)
    rect(s, 0, 0, 0.08, 7.5, C_ACCENT)
    title = data.get("title", "ホワイトボードメモ")
    txt(s, title,                 0.3, 1.2, 12.7, 1.6, size=40, bold=True)
    txt(s, "SnapDeck 自動変換レポート", 0.3, 2.9,  8.0, 0.6, size=18, color=C_ACCENT)
    ts = datetime.now().strftime("%Y年%m月%d日 %H:%M")
    txt(s, f"変換日時: {ts}",     0.3, 3.6,  8.0, 0.5, size=13, color=C_MUTED)
    if data.get("notes"):
        txt(s, f"📝 {data['notes']}", 0.3, 6.7, 12.7, 0.55, size=11, color=C_MUTED)

    # ── グループ別スライド ──
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
            content  = el.get("content", "")
            el_type  = el.get("type", "text")

            if el_type == "heading":
                if y > 1.15: y += 0.08
                txt(s, content, 0.55, y, 12.2, 0.52, size=18, bold=True, color=C_ACCENT)
                y += 0.6
            else:
                txt(s, "  •  " + content, 0.55, y, 12.1, 0.45, size=14, color=C_LIGHT)
                y += 0.48
            if y > 6.7: break

        rect(s, 0, 7.15, 13.33, 0.35, C_NAVY)
        txt(s, "SnapDeck  |  自動生成", 0.4, 7.15, 12.9, 0.35,
            size=10, color=C_MUTED, align=PP_ALIGN.RIGHT)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ─── UI ────────────────────────────────────────────────────────────────────────

# サイドバー
with st.sidebar:
    st.markdown("## ⚙️ 設定")
    st.markdown("---")
    api_key_input = st.text_input(
        "🔑 Anthropic API キー",
        type="password",
        placeholder="sk-ant-api03-...",
        help="https://console.anthropic.com でAPIキーを取得できます"
    )
    if api_key_input:
        st.session_state["api_key_input"] = api_key_input
        st.success("✅ APIキーが設定されました")
    else:
        st.info("APIキー未設定のためデモモードで動作します")

    st.markdown("---")
    st.markdown("### 📤 出力設定")
    output_pptx = st.checkbox("PowerPoint (.pptx)", value=True)
    output_json = st.checkbox("JSON データ",        value=True)

    st.markdown("---")
    st.markdown("### 📖 使い方")
    st.markdown("""
<span class='step-badge'>1</span> 写真を選ぶ or カメラ撮影<br><br>
<span class='step-badge'>2</span> 「変換開始」をクリック<br><br>
<span class='step-badge'>3</span> ファイルをダウンロード
""", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown(
        "<div style='color:#94A3B8;font-size:0.78rem;'>Powered by Claude Vision API<br>"
        "© 2026 SnapDeck</div>",
        unsafe_allow_html=True
    )

# ─── メインエリア ──────────────────────────────────────────────────────────────
st.markdown("""
<div>
  <span class='snap-title'>Snap</span><span class='snap-title snap-accent'>Deck</span>
  <p class='snap-subtitle'>ホワイトボード・手書きメモを、瞬時にスライドへ</p>
</div>
""", unsafe_allow_html=True)

col_upload, col_preview = st.columns([1, 1], gap="large")

with col_upload:
    st.markdown("### 📸 画像を入力")
    tab_file, tab_cam = st.tabs(["📁 ファイルを選択", "📷 カメラで撮影"])

    with tab_file:
        uploaded_file = st.file_uploader(
            "ホワイトボードや手書きメモの写真を選択してください",
            type=["jpg", "jpeg", "png", "webp"],
            help="JPG / PNG / WebP 対応。最大20MB。"
        )

    with tab_cam:
        st.caption("📱 スマホ・PCのカメラを直接使用できます")
        camera_photo = st.camera_input(
            "ホワイトボードをカメラで撮影",
            help="ブラウザのカメラ許可が必要です（初回のみ確認あり）"
        )

    # ファイルとカメラどちらか優先（カメラが新しい場合は優先）
    uploaded = camera_photo or uploaded_file

    if uploaded:
        label = "📷 撮影した画像" if camera_photo and uploaded == camera_photo else uploaded_file.name if uploaded_file else "入力画像"
        st.image(uploaded, caption=label, use_container_width=True)

with col_preview:
    st.markdown("### 💡 使用例")
    st.markdown("""
<div class='snap-card'><h4>📋 対応コンテンツ</h4><p>
・ホワイトボードの板書<br>
・付箋・フリップチャート<br>
・手書きノート・メモ<br>
・ラフスケッチ・フローチャート
</p></div>

<div class='snap-card'><h4>✨ 変換できるもの</h4><p>
・テキスト（日本語・英語対応）<br>
・箇条書き・番号リスト<br>
・見出し・グループ構造<br>
・表・フロー（Phase 2〜）
</p></div>
""", unsafe_allow_html=True)

st.markdown("---")

# ─── 変換処理 ──────────────────────────────────────────────────────────────────

convert_btn = st.button(
    "🚀　変換開始",
    type="primary",
    use_container_width=True,
    disabled=(not uploaded and get_api_key() == "")
)

if st.button("🎯　デモモードで試す（画像不要）", use_container_width=True):
    st.session_state["force_demo"] = True
    st.session_state["extracted"] = get_demo_data()
    st.session_state["pptx_bytes"] = generate_pptx(get_demo_data()) if output_pptx else None
    st.session_state["json_str"]   = json.dumps(get_demo_data(), ensure_ascii=False, indent=2)
    st.rerun()

if convert_btn or st.session_state.get("force_demo"):
    st.session_state.pop("force_demo", None)
    api_key = get_api_key()

    with st.spinner("🤖 AIが解析中...（10〜20秒かかる場合があります）"):
        try:
            if uploaded and api_key:
                img_data, media_type = image_to_base64(uploaded)
                extracted = analyze_with_claude(img_data, media_type, api_key)
            elif uploaded and not api_key:
                st.warning("⚠️ APIキーが未設定のため、デモデータを表示します。サイドバーからAPIキーを入力してください。")
                extracted = get_demo_data()
            else:
                extracted = get_demo_data()

            st.session_state["extracted"]  = extracted
            st.session_state["pptx_bytes"] = generate_pptx(extracted) if output_pptx else None
            st.session_state["json_str"]   = json.dumps(
                {"snapdeck_version": "1.0", "exported_at": datetime.now().isoformat(), "data": extracted},
                ensure_ascii=False, indent=2
            ) if output_json else None

        except Exception as e:
            st.error(f"❌ エラーが発生しました: {e}")
            st.info("APIキーを確認するか、デモモードをお試しください。")

# ─── 結果表示 ──────────────────────────────────────────────────────────────────

if "extracted" in st.session_state:
    extracted = st.session_state["extracted"]
    st.success("✅ 変換完了！ファイルをダウンロードしてください。")

    # 統計
    el_count  = len(extracted.get("elements", []))
    grp_count = len(extracted.get("structure", {}).get("groups", []))
    lang      = extracted.get("language", "—")

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(f"""<div class='result-box'>
            <div class='label'>検出要素数</div>
            <div class='value'>{el_count} 個</div></div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""<div class='result-box'>
            <div class='label'>グループ数</div>
            <div class='value'>{grp_count} 件</div></div>""", unsafe_allow_html=True)
    with c3:
        st.markdown(f"""<div class='result-box'>
            <div class='label'>検出言語</div>
            <div class='value'>{"日本語" if lang=="ja" else "英語" if lang=="en" else lang}</div></div>""",
            unsafe_allow_html=True)

    st.markdown(f"**タイトル（推定）：** {extracted.get('title', '—')}")
    if extracted.get("notes"):
        st.caption(f"📝 {extracted['notes']}")

    # ダウンロード
    st.markdown("### 📥 ダウンロード")
    dl1, dl2, dl3 = st.columns(3)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    with dl1:
        if st.session_state.get("pptx_bytes"):
            st.download_button(
                label="📊 PowerPoint をダウンロード",
                data=st.session_state["pptx_bytes"],
                file_name=f"snapdeck_{ts}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )

    with dl2:
        if st.session_state.get("json_str"):
            st.download_button(
                label="📋 JSON をダウンロード",
                data=st.session_state["json_str"],
                file_name=f"snapdeck_{ts}.json",
                mime="application/json",
            )

    with dl3:
        # ZIPで両方まとめてダウンロード
        if st.session_state.get("pptx_bytes") and st.session_state.get("json_str"):
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.writestr(f"snapdeck_{ts}.pptx", st.session_state["pptx_bytes"])
                zf.writestr(f"snapdeck_{ts}.json",  st.session_state["json_str"].encode("utf-8"))
            st.download_button(
                label="📦 まとめてダウンロード (ZIP)",
                data=zip_buf.getvalue(),
                file_name=f"snapdeck_{ts}.zip",
                mime="application/zip",
            )

    # 抽出内容プレビュー
    with st.expander("🔍 抽出内容の詳細を確認する"):
        for group in extracted.get("structure", {}).get("groups", []):
            st.markdown(f"**{group.get('label', 'グループ')}**")
            elements_by_id = {el["id"]: el for el in extracted.get("elements", [])}
            for item_id in group.get("items", []):
                el = elements_by_id.get(item_id)
                if el:
                    conf = el.get("confidence", 1.0)
                    icon = "📌" if el.get("type") == "heading" else "•"
                    st.markdown(
                        f"&nbsp;&nbsp;{icon} {el['content']} "
                        f"<span style='color:#94A3B8;font-size:0.8rem;'>({conf*100:.0f}%)</span>",
                        unsafe_allow_html=True
                    )
