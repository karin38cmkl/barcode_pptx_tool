import io
import csv
import requests
import re
from datetime import datetime
import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.graphics.barcode import createBarcodeDrawing
from reportlab.graphics import renderPDF
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase import pdfmetrics

# ─── 日本語フォント ────────────────────────────────────────
pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
FONT = 'HeiseiKakuGo-W5'

# ─── 元PDFと同一の寸法 ────────────────────────────────────
# 実測値: 5列×13行, 左マージン22.6pt, 上マージン40.0pt
PAGE_W, PAGE_H = A4          # 595.28 × 841.89 pt
COLS           = 5
ROWS_PER_PAGE  = 13
MARGIN_LEFT    = 22.6        # pt
MARGIN_TOP     = 40.0        # pt（ページ上端からの距離）
CELL_W         = 115.4       # pt（列間隔）
CELL_H         = 60.0        # pt（行間隔）
# 元PDFから実測した値
CONTENT_W      = 87.9        # テキストブロック幅（セル内コンテンツ幅）
BC_WIDTH       = 60.5        # バーコードバー実幅（実測: 60.46pt）
BC_HEIGHT      = 18.0        # バーコードバー高さ（実測: 18.00pt）
TEXT_BLOCK_H   = 11.6        # 2行テキストブロック高さ（実測: 11.6pt）
GAP_TEXT_BC    = 2.56        # テキスト下端〜バーコード上端（実測: 2.56pt）
GAP_BC_NUM     = 1.12        # バーコード下端〜番号テキスト上端（実測: 1.12pt）
NUM_BLOCK_H    = 6.5         # 番号テキストブロック高さ（実測: 6.5pt）
FONT_TEXT_PT   = 5.5         # 商品コード・商品名フォントサイズ
FONT_NUM_PT    = 7.0         # バーコード番号フォントサイズ

# ─── ヘルパー関数 ──────────────────────────────────────────

def col_letter_to_index(letter: str) -> int:
    letter = letter.strip().upper()
    idx = 0
    for ch in letter:
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx - 1


def fit_text(c, text: str, font: str, size: float, max_w: float) -> str:
    while len(text) > 1 and c.stringWidth(text, font, size) > max_w:
        text = text[:-1]
    return text


def draw_cell(c, col: int, row_on_page: int, code: str, name: str, barcode_num: str):
    cell_x     = MARGIN_LEFT + col * CELL_W
    cell_y_top = PAGE_H - MARGIN_TOP - row_on_page * CELL_H
    center_x   = cell_x + CONTENT_W / 2
    bc_left    = cell_x + (CONTENT_W - BC_WIDTH) / 2   # バーコードをセル内中央に配置

    # 1行目：商品コード
    c.setFont(FONT, FONT_TEXT_PT)
    y_line1 = cell_y_top - TEXT_BLOCK_H * 0.45
    c.drawCentredString(center_x, y_line1, fit_text(c, code, FONT, FONT_TEXT_PT, CONTENT_W))

    # 2行目：商品名
    y_line2 = cell_y_top - TEXT_BLOCK_H
    c.drawCentredString(center_x, y_line2, fit_text(c, name, FONT, FONT_TEXT_PT, CONTENT_W))

    # バーコード（テキスト下端から GAP_TEXT_BC 空けて配置、幅は BC_WIDTH で中央寄せ）
    bc_bottom = cell_y_top - TEXT_BLOCK_H - GAP_TEXT_BC - BC_HEIGHT
    if barcode_num.strip():
        try:
            d = createBarcodeDrawing(
                'EAN13',
                value=barcode_num.strip()[:13].zfill(13),
                width=BC_WIDTH,
                height=BC_HEIGHT,
                barHeight=BC_HEIGHT,
                humanReadable=False,
                quiet=0,
            )
            renderPDF.draw(d, c, bc_left, bc_bottom)
        except Exception:
            try:
                d = createBarcodeDrawing(
                    'Code128',
                    value=barcode_num.strip(),
                    width=BC_WIDTH,
                    height=BC_HEIGHT,
                    barHeight=BC_HEIGHT,
                    humanReadable=False,
                    quiet=0,
                )
                renderPDF.draw(d, c, bc_left, bc_bottom)
            except Exception:
                pass

    # バーコード番号（バーコード下端から GAP_BC_NUM 空けて・大きめフォント）
    y_num = bc_bottom - GAP_BC_NUM - NUM_BLOCK_H * 0.8
    c.setFont(FONT, FONT_NUM_PT)
    c.drawCentredString(center_x, y_num, barcode_num)


def generate_pdf(data: list) -> bytes:
    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=A4)
    per_page = COLS * ROWS_PER_PAGE

    for i, (code, name, barcode_num) in enumerate(data):
        if i > 0 and i % per_page == 0:
            c.showPage()
        cell_idx = i % per_page
        draw_cell(c, cell_idx % COLS, cell_idx // COLS, code, name, barcode_num)

    c.save()
    return buf.getvalue()


def load_csv_rows(url=None, csv_bytes=None) -> list:
    if url:
        m = re.search(r'/spreadsheets/d/([a-zA-Z0-9_-]+)', url)
        if not m:
            st.error("Google SheetsのURLが正しくありません")
            return []
        export_url = f"https://docs.google.com/spreadsheets/d/{m.group(1)}/export?format=csv"
        resp = requests.get(export_url, allow_redirects=True, timeout=30)
        return list(csv.reader(io.StringIO(resp.content.decode('utf-8'))))
    elif csv_bytes:
        try:
            return list(csv.reader(io.StringIO(csv_bytes.decode('utf-8-sig'))))
        except Exception:
            return list(csv.reader(io.StringIO(csv_bytes.decode('shift-jis'))))
    return []


# ─── UI ───────────────────────────────────────────────────
st.set_page_config(page_title="バーコード発行", layout="wide")
st.title("バーコード発行")
st.caption("元PDFと同一レイアウト（A4・5列×13行）でバーコードPDFを生成します")

with st.expander("列設定", expanded=False):
    c1, c2, c3, c4 = st.columns(4)
    col_code    = c1.text_input("商品コード列",    "A")
    col_name    = c2.text_input("商品名列",        "B")
    col_barcode = c3.text_input("バーコード番号列", "C")
    header_rows = c4.number_input("ヘッダー行数",  min_value=0, max_value=10, value=1)

tab_csv, tab_url = st.tabs(["CSVアップロード", "Google Sheets URL"])
with tab_csv:
    csv_file = st.file_uploader("CSVファイル", type=["csv"])
with tab_url:
    sheets_url = st.text_input("URLを貼り付け",
                               placeholder="https://docs.google.com/spreadsheets/d/...")

st.divider()

if st.button("▶ バーコードPDFを生成", type="primary", use_container_width=True):
    use_url = sheets_url.strip() if sheets_url else None
    use_csv = csv_file.read() if csv_file else None

    if not use_url and not use_csv:
        st.error("CSVまたはGoogle SheetsのURLを指定してください")
        st.stop()

    with st.spinner("データ読み込み中..."):
        rows = load_csv_rows(url=use_url, csv_bytes=use_csv)

    if not rows:
        st.error("データを読み込めませんでした")
        st.stop()

    ci = col_letter_to_index(col_code)
    ni = col_letter_to_index(col_name)
    bi = col_letter_to_index(col_barcode)

    data = []
    for row in rows[int(header_rows):]:
        code    = row[ci].strip() if len(row) > ci else ''
        name    = row[ni].strip() if len(row) > ni else ''
        barcode = row[bi].strip() if len(row) > bi else ''
        if code or barcode:
            data.append((code, name, barcode))

    st.success(f"{len(data)}件 読み込み完了")

    with st.spinner(f"PDF生成中（{len(data)}件）..."):
        try:
            pdf_bytes = generate_pdf(data)
            pages = -(-len(data) // (COLS * ROWS_PER_PAGE))
            st.success(f"PDF生成完了: {len(data)}件 / {pages}ページ")
        except Exception as e:
            st.error(f"PDF生成エラー: {e}")
            st.stop()

    # 保管庫に保存
    if "archive" not in st.session_state:
        st.session_state.archive = []
    csv_fname = csv_file.name if csv_file else None
    st.session_state.archive.append({
        "type": "barcode_pdf",
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "filename": "barcodes.pdf",
        "data_source": f"Google Sheets: {use_url}" if use_url else f"CSV: {csv_fname}",
        "settings": {
            "商品コード列": col_code,
            "商品名列": col_name,
            "バーコード番号列": col_barcode,
            "ヘッダー行数": int(header_rows),
        },
        "item_count": len(data),
        "pages": pages,
        "pdf_bytes": pdf_bytes,
        "csv_bytes": use_csv,
        "csv_filename": csv_fname,
    })

    st.download_button(
        label="⬇ バーコードPDFをダウンロード",
        data=pdf_bytes,
        file_name="barcodes.pdf",
        mime="application/pdf",
        use_container_width=True,
        type="primary",
    )
