import streamlit as st

st.set_page_config(page_title="保管庫", layout="wide")
st.title("保管庫")
st.caption("このセッション中に生成したファイルを自動保存しています。ブラウザを閉じるかアプリを再起動すると消去されます。")

archive = st.session_state.get("archive", [])

if not archive:
    st.info("まだ保存されたファイルがありません。「バーコード発行」または「PPTXラベル生成」でファイルを生成すると自動で保存されます。")
    st.stop()

barcode_entries = [e for e in archive if e["type"] == "barcode_pdf"]
pptx_entries    = [e for e in archive if e["type"] == "pptx"]

# ─── ① バーコード発行 ─────────────────────────────────────
st.subheader(f"① バーコード発行　（{len(barcode_entries)}件）")

if not barcode_entries:
    st.write("まだバーコードPDFは生成されていません")
else:
    for i, entry in enumerate(reversed(barcode_entries)):
        label = f"{entry['timestamp']}　{entry['item_count']}件 / {entry['pages']}ページ　｜　{entry['data_source']}"
        with st.expander(label):
            c1, c2 = st.columns([3, 2])
            with c1:
                st.markdown("**参照データ**")
                st.write(entry["data_source"])
                st.markdown("**設定**")
                st.json(entry["settings"], expanded=False)
            with c2:
                st.download_button(
                    "⬇ バーコードPDFをダウンロード",
                    data=entry["pdf_bytes"],
                    file_name=entry["filename"],
                    mime="application/pdf",
                    use_container_width=True,
                    key=f"dl_bc_{i}",
                )
                if entry.get("csv_bytes"):
                    st.download_button(
                        f"⬇ 参照CSV（{entry['csv_filename']}）をダウンロード",
                        data=entry["csv_bytes"],
                        file_name=entry["csv_filename"],
                        mime="text/csv",
                        use_container_width=True,
                        key=f"dl_bc_csv_{i}",
                    )

st.divider()

# ─── ② PPTXラベル生成 ────────────────────────────────────
st.subheader(f"② PPTXラベル生成　（{len(pptx_entries)}件）")

if not pptx_entries:
    st.write("まだPPTXは生成されていません")
else:
    for i, entry in enumerate(reversed(pptx_entries)):
        label = f"{entry['timestamp']}　{entry['matched_count']}スライド　｜　{entry['data_source']}"
        with st.expander(label):
            c1, c2 = st.columns([3, 2])
            with c1:
                st.markdown("**参照データ**")
                st.write(entry["data_source"])
                st.markdown("**設定**")
                st.json(entry["settings"], expanded=False)
            with c2:
                st.download_button(
                    "⬇ PPTXをダウンロード",
                    data=entry["pptx_bytes"],
                    file_name=entry["filename"],
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                    key=f"dl_pptx_{i}",
                )
                if entry.get("csv_bytes"):
                    st.download_button(
                        f"⬇ 参照CSV（{entry['csv_filename']}）をダウンロード",
                        data=entry["csv_bytes"],
                        file_name=entry["csv_filename"],
                        mime="text/csv",
                        use_container_width=True,
                        key=f"dl_pptx_csv_{i}",
                    )
