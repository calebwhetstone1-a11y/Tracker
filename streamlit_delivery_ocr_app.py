import io
import os
import re
import tempfile
from typing import List, Dict, Tuple

import pandas as pd
import pytesseract
import streamlit as st
from openpyxl import load_workbook
from pdf2image import convert_from_bytes
from PIL import Image, ImageOps

st.set_page_config(page_title="Delivery OCR Tracker", layout="wide")

# -----------------------------
# Session state setup
# -----------------------------
if "processed" not in st.session_state:
    st.session_state.processed = False

if "raw_df" not in st.session_state:
    st.session_state.raw_df = pd.DataFrame()

if "summary_df" not in st.session_state:
    st.session_state.summary_df = pd.DataFrame()

if "ocr_workbook_bytes" not in st.session_state:
    st.session_state.ocr_workbook_bytes = None

if "updated_tracker_bytes" not in st.session_state:
    st.session_state.updated_tracker_bytes = None

if "unmatched_df" not in st.session_state:
    st.session_state.unmatched_df = pd.DataFrame()

# -----------------------------
# Helpers
# -----------------------------
def load_pages_from_upload(uploaded_file) -> List[Image.Image]:
    ext = os.path.splitext(uploaded_file.name)[1].lower()

    if ext == ".pdf":
        pdf_bytes = uploaded_file.read()
        pages = convert_from_bytes(pdf_bytes, dpi=300)
        uploaded_file.seek(0)
        return [page.convert("L") for page in pages]

    image = Image.open(uploaded_file).convert("L")
    uploaded_file.seek(0)
    return [image]


def crop_table_region(img: Image.Image) -> Image.Image:
    width, height = img.size
    top_crop = int(height * 0.25)
    bottom_crop = int(height * 0.90)
    cropped = img.crop((0, top_crop, width, bottom_crop))
    return ImageOps.autocontrast(cropped)


def ocr_page(img: Image.Image) -> str:
    return pytesseract.image_to_string(img, config="--psm 6")


def parse_items_from_text(text: str, source_file: str, page_number: int) -> List[Dict]:
    lines = text.split("\n")
    items = []
    capture = False

    for line in lines:
        line = " ".join(line.split())
        if not line:
            continue
        line_lower = line.lower()

        if "item" in line_lower and "descr" in line_lower:
            capture = True
            continue

        if "colli" in line_lower:
            capture = False

        if not capture:
            continue

        item_match = re.search(r'^(\d{4}-\d{4})\s+(.+)$', line)
        if not item_match:
            continue

        item_no = item_match.group(1)
        remainder = item_match.group(2).strip()

        qty_match = re.search(r'(.+?)\s+(\d+)(?:\s+\w+)?$', remainder)
        if not qty_match:
            continue

        description = qty_match.group(1).strip()
        quantity = int(qty_match.group(2))

        items.append({
            "ItemNo": item_no,
            "Description": description,
            "Quantity": quantity,
            "SourceFile": source_file,
            "PageNumber": page_number,
        })

    return items


def summarize_items(raw_df: pd.DataFrame) -> pd.DataFrame:
    if raw_df.empty:
        return pd.DataFrame(columns=["ItemNo", "Description", "Quantity"])

    summary = (
        raw_df.groupby("ItemNo")
        .agg(
            Quantity=("Quantity", "sum"),
            Description=("Description", lambda x: x.mode().iat[0] if not x.mode().empty else x.iloc[0]),
        )
        .reset_index()
    )
    return summary[["ItemNo", "Description", "Quantity"]]


def update_tracker_workbook(tracker_bytes: bytes, summary_df: pd.DataFrame, mode: str) -> Tuple[bytes, pd.DataFrame]:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(tracker_bytes)
        tmp_path = tmp.name

    wb = load_workbook(tmp_path)

    tracker_rows = {}
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        headers = {}
        for col in range(1, ws.max_column + 1):
            value = ws.cell(row=1, column=col).value
            if value is not None:
                headers[str(value).strip()] = col

        if "Part #" not in headers or "Qty Recei" not in headers:
            continue

        part_col = headers["Part #"]
        for row in range(2, ws.max_row + 1):
            part_value = ws.cell(row=row, column=part_col).value
            if part_value is None:
                continue
            tracker_rows[str(part_value).strip()] = {
                "sheet": ws,
                "row": row,
                "headers": headers,
            }

    not_found = []

    for _, row in summary_df.iterrows():
        item_no = str(row["ItemNo"]).strip()
        qty = row["Quantity"]
        desc = row["Description"]

        if item_no not in tracker_rows:
            not_found.append((item_no, desc, qty))
            continue

        entry = tracker_rows[item_no]
        ws = entry["sheet"]
        excel_row = entry["row"]
        qty_col = entry["headers"]["Qty Recei"]

        if mode == "add":
            current_val = ws.cell(row=excel_row, column=qty_col).value
            if current_val in [None, ""]:
                current_val = 0
            try:
                current_val = float(current_val)
            except Exception:
                current_val = 0
            ws.cell(row=excel_row, column=qty_col).value = current_val + qty
        else:
            ws.cell(row=excel_row, column=qty_col).value = qty

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)

    unmatched_df = pd.DataFrame(not_found, columns=["ItemNo", "Description", "Quantity"])
    return out.getvalue(), unmatched_df


def build_ocr_results_workbook(raw_df: pd.DataFrame, summary_df: pd.DataFrame) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        raw_df.to_excel(writer, sheet_name="Raw_OCR_Data", index=False)
        summary_df.to_excel(writer, sheet_name="Summarized_Totals", index=False)
    out.seek(0)
    return out.getvalue()


# -----------------------------
# UI
# -----------------------------
st.title("Delivery OCR Tracker")
st.write(
    "1. Upload delivery PDFs or images.\n"
    "2. Optionally upload your tracking workbook.\n"
    "3. Click Process Deliveries.\n"
    "4. Download the OCR results and updated tracker."
)

with st.sidebar:
    st.header("Settings")
    update_mode_label = st.radio(
        "Qty Received behavior",
        options=["Replace Qty Received", "Add to Qty Received"],
        index=0,
    )
    update_mode = "overwrite" if update_mode_label == "Replace Qty Received" else "add"
    show_page_previews = st.checkbox("Show cropped page previews", value=False)

col1, col2 = st.columns(2)

with col1:
    delivery_files = st.file_uploader(
        "Upload delivery images or PDFs",
        type=["pdf", "png", "jpg", "jpeg"],
        accept_multiple_files=True,
        key="delivery_files",
    )

with col2:
    tracker_file = st.file_uploader(
        "Upload tracking workbook (optional)",
        type=["xlsx"],
        accept_multiple_files=False,
        key="tracker_file",
    )

if st.button("Process Deliveries", type="primary", use_container_width=True):
    if not delivery_files:
        st.error("Upload at least one delivery image or PDF first.")
        st.stop()

    all_items: List[Dict] = []

    progress = st.progress(0)
    status = st.empty()

    for file_index, uploaded_file in enumerate(delivery_files, start=1):
        status.write(f"Processing {uploaded_file.name}...")
        pages = load_pages_from_upload(uploaded_file)

        for page_index, img in enumerate(pages, start=1):
            cropped_img = crop_table_region(img)

            if show_page_previews:
                with st.expander(f"Preview: {uploaded_file.name} | Page {page_index}"):
                    st.image(cropped_img, caption=f"{uploaded_file.name} - page {page_index}")

            text = ocr_page(cropped_img)
            items = parse_items_from_text(text, uploaded_file.name, page_index)
            all_items.extend(items)

        progress.progress(file_index / len(delivery_files))

    raw_df = pd.DataFrame(all_items)
    summary_df = summarize_items(raw_df)

    st.session_state.raw_df = raw_df
    st.session_state.summary_df = summary_df
    st.session_state.ocr_workbook_bytes = build_ocr_results_workbook(raw_df, summary_df)
    st.session_state.updated_tracker_bytes = None
    st.session_state.unmatched_df = pd.DataFrame()

    if tracker_file is not None and not summary_df.empty:
        updated_tracker_bytes, unmatched_df = update_tracker_workbook(
            tracker_file.getvalue(), summary_df, update_mode
        )
        st.session_state.updated_tracker_bytes = updated_tracker_bytes
        st.session_state.unmatched_df = unmatched_df

    st.session_state.processed = True

# -----------------------------
# Results area
# -----------------------------
if st.session_state.processed:
    raw_df = st.session_state.raw_df
    summary_df = st.session_state.summary_df

    st.subheader("OCR Results")
    c1, c2, c3 = st.columns(3)
    c1.metric("Raw OCR rows", 0 if raw_df.empty else len(raw_df))
    c2.metric("Unique items", 0 if summary_df.empty else len(summary_df))
    c3.metric("Tracker matches pending review", 0 if st.session_state.unmatched_df.empty else len(st.session_state.unmatched_df))

    if raw_df.empty:
        st.warning("No item rows were found.")
    else:
        tab1, tab2, tab3 = st.tabs(["Summarized Totals", "Raw OCR Data", "Downloads"])

        with tab1:
            st.dataframe(summary_df, use_container_width=True)

        with tab2:
            st.dataframe(raw_df, use_container_width=True)

        with tab3:
            @st.fragment
            def download_section():
                st.download_button(
                    label="Download OCR Results Workbook",
                    data=st.session_state.ocr_workbook_bytes,
                    file_name="OCR_Results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

                if st.session_state.updated_tracker_bytes is not None:
                    st.download_button(
                        label="Download Updated Tracking Workbook",
                        data=st.session_state.updated_tracker_bytes,
                        file_name="Updated_Tracking_File.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

                    if st.session_state.unmatched_df.empty:
                        st.success("All OCR items matched a tracker Part #.")
                    else:
                        st.warning("Some OCR items were not found in the tracker.")
                        st.dataframe(st.session_state.unmatched_df, use_container_width=True)

                        unmatched_bytes = io.BytesIO()
                        st.session_state.unmatched_df.to_excel(unmatched_bytes, index=False)
                        unmatched_bytes.seek(0)

                        st.download_button(
                            label="Download Unmatched Items",
                            data=unmatched_bytes,
                            file_name="Unmatched_OCR_Items.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                        )
                else:
                    st.info("No tracking workbook was uploaded for this run.")

            download_section()
