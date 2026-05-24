import io
import re
import csv
import zipfile
import tempfile
import os

import fitz
import requests
import streamlit as st
from PIL import Image
from lxml import etree
from pptx import Presentation
from pptx.util import Mm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

# ─── ページ設定 ───────────────────────────────────────────
st.set_page_config(page_title="バーコードPPTX生成ツール", layout="wide")
st.title("バーコードラベル PPTX生成ツール")

# ─── サイドバー設定 ────────────────────────────────────────
st.sidebar.header("設定")

line1_text = st.sidebar.text_input("1行目テキスト（固定）", "MimoRhea(みもれあ) 29×29")

st.sidebar.subheader("スプレッドシート列設定")
st.sidebar.caption("2行目テキスト形式: K列_H列(G列)_L列")
col_k = st.sidebar.text_input("商品コード列", "K")
col_h = st.sidebar.text_input("カテゴリ名列", "H")
col_g = st.sidebar.text_input("カテゴリコード列", "G")
col_l = st.sidebar.text_input("カラー名列", "L")
header_rows = st.sidebar.number_input("ヘッダー行数（スキップ行数）", min_value=0, max_value=5, value=1)

st.sidebar.subheader("ページ・デザイン設定")
page_w_mm = st.sidebar.number_input("ページ幅 (mm)", value=400, min_value=50, max_value=1000)
page_h_mm = st.sidebar.number_input("ページ高さ (mm)", value=200, min_value=50, max_value=1000)
white_box_pt = st.sidebar.number_input("白塗り高さ (pt)", value=193, min_value=10, max_value=500)
font_pt = st.sidebar.number_input("フォントサイズ (pt)", value=60, min_value=8, max_value=200)
scale = st.sidebar.slider("画像解像度 (PDF拡大率)", min_value=2, max_value=10, value=7,
                          help="7 = 700%相当。大きいほど高画質だが処理が遅い")

# ─── ヘルパー関数 ──────────────────────────────────────────

def col_letter_to_index(letter: str) -> int:
    """A→0, B→1 ... Z→25"""
    letter = letter.strip().upper()
    idx = 0
    for ch in letter:
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx - 1


def make_filename(code: str) -> str:
    """A01_01_29x29 → A-01"""
    parts = code.split('_')
    if len(parts) >= 2:
        return f"{parts[0][0]}-{parts[1]}"
    return code


def remove_shadow(shape):
    sp = shape._element
    spPr = sp.find(qn('p:spPr'))
    if spPr is None:
        spPr = sp.find('spPr')
    if spPr is not None:
        for el in spPr.findall(qn('a:effectLst')):
            spPr.remove(el)
        etree.SubElement(spPr, qn('a:effectLst'))


def add_textbox(slide, text, left, top, width, height, font_pt, bold=True):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = False
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = text
    r.font.bold = bold
    r.font.size = Pt(font_pt)
    r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    remove_shadow(txBox)
    return txBox


def extract_barcodes_from_pdf(pdf_bytes: bytes, scale: int):
    """PDFからバーコードラベル画像を切り出す。{filename: PIL.Image} を返す"""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    mat = fitz.Matrix(scale, scale)
    results = {}

    for page in doc:
        pix = page.get_pixmap(matrix=mat)
        page_img = Image.open(io.BytesIO(pix.tobytes('png')))
        blocks = page.get_text('blocks')

        product_blocks = []
        barcode_num_blocks = []
        for b in blocks:
            text = b[4].strip()
            first_line = text.split('\n')[0].strip()
            if re.match(r'^[A-Z]\d+_\d+_\d+x\d+', first_line):
                product_blocks.append(b)
            elif re.match(r'^\d{10,15}$', text.replace('\n', '')):
                barcode_num_blocks.append(b)

        for pb in product_blocks:
            first_line = pb[4].strip().split('\n')[0].strip()
            filename = make_filename(first_line)
            px0, py0, px2, py2 = pb[0], pb[1], pb[2], pb[3]
            pc = (px0 + px2) / 2

            matched_nb = None
            min_dist = float('inf')
            for nb in barcode_num_blocks:
                nc = (nb[0] + nb[2]) / 2
                if nb[1] > py0 and abs(nc - pc) < 30:
                    dist = nb[1] - py2
                    if dist < min_dist:
                        min_dist = dist
                        matched_nb = nb

            if matched_nb is None:
                continue

            cx0 = int(min(px0, matched_nb[0]) - 2)
            cy0 = int(py0 - 2)
            cx1 = int(max(px2, matched_nb[2]) + 2)
            cy1 = int(matched_nb[3] + 2)

            cropped = page_img.crop((
                cx0 * scale, cy0 * scale,
                cx1 * scale, cy1 * scale
            ))
            results[filename] = cropped

    doc.close()
    return results


def load_spreadsheet_data(url: str | None, csv_bytes: bytes | None,
                          ki: int, hi: int, gi: int, li: int,
                          header_rows: int) -> dict[str, str]:
    """スプシURL or CSVバイトからデータを読み込み {商品コード: 2行目テキスト} を返す"""
    if url:
        # Google Sheets URLからID抽出してCSVエクスポート
        match = re.search(r'/spreadsheets/d/([a-zA-Z0-9_-]+)', url)
        if not match:
            st.error("Google SheetsのURLが正しくありません")
            return {}
        sheet_id = match.group(1)
        export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
        resp = requests.get(export_url, allow_redirects=True, timeout=30)
        content = resp.content.decode('utf-8')
        reader = csv.reader(io.StringIO(content))
        rows = list(reader)
    elif csv_bytes:
        content = csv_bytes.decode('utf-8')
        reader = csv.reader(io.StringIO(content))
        rows = list(reader)
    else:
        return {}

    data = {}
    for row in rows[header_rows:]:
        k = row[ki].strip() if len(row) > ki else ''
        h = row[hi].strip() if len(row) > hi else ''
        g = row[gi].strip() if len(row) > gi else ''
        l = row[li].strip() if len(row) > li else ''
        if k and k != '-':
            data[k] = f"{k}_{h}({g})_{l}"
    return data


def build_pptx(barcode_images: dict, ss_data: dict,
               line1: str, page_w_mm: int, page_h_mm: int,
               white_box_pt: int, font_pt: int) -> bytes:
    """PPTXをバイトで返す"""
    img_aspect = None  # 最初の画像から取得

    prs = Presentation()
    prs.slide_width = Mm(page_w_mm)
    prs.slide_height = Mm(page_h_mm)
    blank = prs.slide_layouts[6]

    white_h = Pt(white_box_pt)
    half_h = white_h // 2

    for code in sorted(barcode_images.keys()):
        if code not in ss_data:
            continue

        line2 = ss_data[code]
        img = barcode_images[code]

        w_px, h_px = img.size
        img_h_mm = page_w_mm * h_px / w_px

        slide = prs.slides.add_slide(blank)

        # バーコード画像（フルwidth）
        buf = io.BytesIO()
        img.save(buf, format='JPEG', quality=95)
        buf.seek(0)
        slide.shapes.add_picture(buf, Mm(0), Mm(0), Mm(page_w_mm), Mm(img_h_mm))

        # 白塗り長方形
        rect = slide.shapes.add_shape(1, Mm(0), Mm(0), Mm(page_w_mm), white_h)
        rect.fill.solid()
        rect.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        rect.line.fill.background()
        remove_shadow(rect)

        # テキストボックス 1行目
        add_textbox(slide, line1, Mm(0), Mm(0), Mm(page_w_mm), half_h, font_pt)

        # テキストボックス 2行目
        add_textbox(slide, line2, Mm(0), half_h, Mm(page_w_mm), half_h, font_pt)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ─── メイン UI ─────────────────────────────────────────────

col1, col2 = st.columns(2)

with col1:
    st.subheader("① PDFをアップロード")
    pdf_file = st.file_uploader("バーコードPDF", type=["pdf"])

with col2:
    st.subheader("② スプレッドシートデータ")
    tab_url, tab_csv = st.tabs(["Google Sheets URL", "CSVアップロード"])
    with tab_url:
        sheets_url = st.text_input("URLを貼り付け", placeholder="https://docs.google.com/spreadsheets/d/...")
    with tab_csv:
        csv_file = st.file_uploader("CSVファイル", type=["csv"])

st.divider()

if st.button("▶ PPTXを生成", type="primary", use_container_width=True):
    if not pdf_file:
        st.error("PDFをアップロードしてください")
        st.stop()

    use_url = sheets_url.strip() if sheets_url else None
    use_csv = csv_file.read() if csv_file else None

    if not use_url and not use_csv:
        st.error("スプレッドシートのURLまたはCSVを指定してください")
        st.stop()

    ki = col_letter_to_index(col_k)
    hi = col_letter_to_index(col_h)
    gi = col_letter_to_index(col_g)
    li = col_letter_to_index(col_l)

    with st.spinner("PDFからバーコード画像を切り出し中..."):
        try:
            barcode_images = extract_barcodes_from_pdf(pdf_file.read(), scale)
            st.success(f"バーコード画像: {len(barcode_images)}件 切り出し完了")
        except Exception as e:
            st.error(f"PDF処理エラー: {e}")
            st.stop()

    with st.spinner("スプレッドシートデータを読み込み中..."):
        try:
            ss_data = load_spreadsheet_data(use_url, use_csv, ki, hi, gi, li, int(header_rows))
            st.success(f"スプレッドシート: {len(ss_data)}件 読み込み完了")
        except Exception as e:
            st.error(f"スプレッドシート読み込みエラー: {e}")
            st.stop()

    # 一致確認
    matched = set(barcode_images.keys()) & set(ss_data.keys())
    only_img = set(barcode_images.keys()) - set(ss_data.keys())
    only_ss = set(ss_data.keys()) - set(barcode_images.keys())

    if only_img:
        st.warning(f"スプシにデータなし（スキップ）: {sorted(only_img)}")
    if only_ss:
        st.warning(f"PDF画像なし（スキップ）: {sorted(only_ss)}")

    with st.spinner(f"PPTX生成中（{len(matched)}枚）..."):
        try:
            pptx_bytes = build_pptx(
                barcode_images, ss_data, line1_text,
                int(page_w_mm), int(page_h_mm),
                int(white_box_pt), int(font_pt)
            )
            st.success(f"PPTX生成完了: {len(matched)}スライド")
        except Exception as e:
            st.error(f"PPTX生成エラー: {e}")
            st.stop()

    st.download_button(
        label="⬇ PPTXをダウンロード",
        data=pptx_bytes,
        file_name="barcode_labels.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True,
        type="primary"
    )

    # プレビュー（最初の3枚）
    st.subheader("プレビュー（最初の3件）")
    preview_codes = sorted(barcode_images.keys())[:3]
    cols = st.columns(3)
    for i, code in enumerate(preview_codes):
        with cols[i]:
            st.image(barcode_images[code], caption=code, use_container_width=True)
            if code in ss_data:
                st.caption(ss_data[code])
