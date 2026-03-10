import io
import re
import os
import unicodedata
from datetime import datetime
from io import BytesIO
from typing import Optional

import pandas as pd
import streamlit as st
from openpyxl.styles import Font

st.set_page_config(page_title="Excel Tools", layout="centered")

tool = st.sidebar.radio(
    "Alət seç:",
    ["Excel Sheet Combiner", "Excel Cleaner"],
)

st.title(tool)

# ─────────────────────────────────────────────
# TOOL 1 – Excel Sheet Combiner
# ─────────────────────────────────────────────
if tool == "Excel Sheet Combiner":
    uploaded_files = st.file_uploader(
        "Upload Excel files", type="xlsx", accept_multiple_files=True
    )

    sheet_names_input = st.text_input(
        "Enter comma-separated sheet names to combine",
        "Tiktok,Facebook,News,YouTube,Linkedin,Twitter,Instagram",
    )

    if st.button("Combine Files"):
        if not uploaded_files:
            st.warning("Please upload at least one Excel file.")
        else:
            sheet_names = [s.strip().lower() for s in sheet_names_input.split(",")]
            combined_data = {sheet: [] for sheet in sheet_names}

            for uploaded_file in uploaded_files:
                try:
                    company_name = os.path.splitext(uploaded_file.name)[0]
                    xls = pd.ExcelFile(uploaded_file)
                    excel_sheets = {s.lower(): s for s in xls.sheet_names}

                    for sheet in sheet_names:
                        if sheet in excel_sheets:
                            df = pd.read_excel(xls, sheet_name=excel_sheets[sheet])
                            df["Company"] = company_name
                            combined_data[sheet].append(df)
                        else:
                            st.warning(
                                f"Sheet '{sheet}' not found in {uploaded_file.name}. Skipping."
                            )
                except Exception as e:
                    st.error(f"Error processing file {uploaded_file.name}: {e}")

            has_data = any(combined_data[s] for s in sheet_names)
            if not has_data:
                st.error("Heç bir sheet-də data tapılmadı.")
            else:
                output = BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    for sheet, data in combined_data.items():
                        if data:
                            combined_df = pd.concat(data, ignore_index=True)
                            combined_df.to_excel(
                                writer, sheet_name=sheet.capitalize(), index=False
                            )
                            st.success(f"Sheet '{sheet}' combined successfully.")
                        else:
                            st.warning(f"No data for sheet '{sheet}'.")

                output.seek(0)
                st.download_button(
                    label="Download Combined Excel File",
                    data=output,
                    file_name="combined_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

# ─────────────────────────────────────────────
# TOOL 2 – Excel Cleaner
# ─────────────────────────────────────────────
else:
    SENTIMENT_MAP = {
        "müsbət": "Positive", "musbet": "Positive",
        "müsbет": "Positive", "MÜSBƏT": "Positive",
        "pozitiv": "Positive", "POZİTİV": "Positive", "POZITIV": "Positive",
        "pozitif": "Positive", "POZİTİF": "Positive",
        "positive": "Positive", "POSITIVE": "Positive", "Positive": "Positive",
        "müspet": "Positive", "muspet": "Positive",
        "mənfi": "Negative", "menfi": "Negative",
        "MƏNFİ": "Negative", "MENFİ": "Negative", "MENFI": "Negative",
        "negativ": "Negative", "NEGATİV": "Negative", "NEGATIV": "Negative",
        "neqativ": "Negative", "NEQATİV": "Negative", "NEQATIV": "Negative",
        "negative": "Negative", "NEGATIVE": "Negative", "Negative": "Negative",
        "neytral": "Neutral", "NEYTRAL": "Neutral", "Neytral": "Neutral",
        "neutral": "Neutral", "NEUTRAL": "Neutral", "Neutral": "Neutral",
        "tərəfsiz": "Neutral", "TƏRƏFSİZ": "Neutral",
        "terefsiz": "Neutral", "TEREFSIZ": "Neutral",
        "1": "Positive", "0": "Neutral", "-1": "Negative",
    }

    def _fold(s: str) -> str:
        s = unicodedata.normalize("NFKD", s.lower())
        s = "".join(c for c in s if not unicodedata.combining(c))
        return s.replace("ı", "i").replace("ə", "e")

    def normalize_date_text(x) -> str:
        if pd.isna(x):
            return ""
        s = str(x).strip()
        s = re.sub(r"\s+\d{1,2}:\d{2}(:\d{2})?\s*(AM|PM)?$", "", s, flags=re.I)
        s = (
            s.replace("/", "-").replace(".", "-").replace("_", "-")
             .replace("–", "-").replace("—", "-")
        )
        s = re.sub(r"-{2,}", "-", s)
        return s.strip('"').strip("'").strip()

    def translate_sentiment(x) -> str:
        if pd.isna(x):
            return "Neutral"
        raw = str(x).strip()
        if not raw:
            return "Neutral"
        key = re.sub(r"\s+", " ", _fold(raw))
        return SENTIMENT_MAP.get(key, "Neutral")

    def best_col(df: pd.DataFrame, candidates) -> Optional[str]:
        cols_lower = {c.lower(): c for c in df.columns}
        for cand in candidates:
            for k, original in cols_lower.items():
                if cand in k:
                    return original
        return None

    def _guess_date_col(df: pd.DataFrame, exclude_cols=None) -> Optional[str]:
        """Try every column; return the one with the most parseable dates."""
        exclude = set(c.lower() for c in (exclude_cols or []))
        best, best_score = None, 0
        for col in df.columns:
            if col.lower() in exclude:
                continue
            try:
                vals = df[col].dropna().astype(str).str.strip()
                if vals.empty:
                    continue
                sample = vals.head(50)
                p1 = pd.to_datetime(sample, errors="coerce", dayfirst=True)
                p2 = pd.to_datetime(sample.map(normalize_date_text), errors="coerce", dayfirst=True)
                score = max(p1.notna().sum(), p2.notna().sum())
                if score > best_score:
                    best_score = score
                    best = col
            except Exception:
                continue
        if best and best_score >= max(2, len(df.head(50)) * 0.3):
            return best
        return None

    def process_sheet(df: pd.DataFrame) -> pd.DataFrame:
        df = df.reset_index(drop=True)
        url_col       = best_col(df, ["url", "link", "href", "source"])
        content_col   = best_col(df, ["content", "text", "metn", "mətn", "kontent",
                                       "message", "post", "caption", "description", "body"])
        date_col      = best_col(df, ["date", "tarix", "data", "datetime",
                                       "time", "timestamp", "created", "published",
                                       "posted", "vaxt", "zaman", "created_at",
                                       "publish", "gun", "gün"])
        if not date_col:
            date_col = _guess_date_col(df, exclude_cols=[
                url_col or "", content_col or ""])
        sentiment_col = best_col(df, ["sentiment", "hiss", "emosiya", "rating",
                                       "tone", "mood", "label", "class"])
        measures_col  = best_col(df, ["measures", "tədbir", "action"])

        # URL və Content tapılmasa → sheet-i skip et
        if not url_col and not content_col:
            raise ValueError(
                f"URL və Content sütunları tapılmadı. "
                f"Mövcud sütunlar: {list(df.columns)}"
            )

        # URL və ya Content tapılmasa, digərini istifadə et
        if not url_col:
            url_col = content_col
        if not content_col:
            content_col = url_col

        # Date tapılmasa boş burax
        if date_col:
            raw_vals = df[date_col].fillna("").astype(str).str.strip()
            # ISO format (YYYY-MM-DD, Excel datetime) → dayfirst=False
            parsed = pd.to_datetime(raw_vals, errors="coerce", dayfirst=False)
            # DD.MM.YYYY, DD/MM/YYYY → dayfirst=True
            alt = pd.to_datetime(raw_vals, errors="coerce", dayfirst=True)
            if alt.notna().sum() > parsed.notna().sum():
                parsed = alt
            # Normalize edib yenə cəhd et
            if parsed.isna().mean() > 0.4:
                normalized = raw_vals.map(normalize_date_text)
                p2 = pd.to_datetime(normalized, errors="coerce", dayfirst=False)
                p3 = pd.to_datetime(normalized, errors="coerce", dayfirst=True)
                best_norm = p2 if p2.notna().sum() >= p3.notna().sum() else p3
                if best_norm.notna().sum() > parsed.notna().sum():
                    parsed = best_norm
        else:
            parsed = pd.Series(pd.NaT, index=df.index)

        # Sentiment tapılmasa Neutral yaz
        if sentiment_col:
            sentiments = df[sentiment_col].map(translate_sentiment).values
        else:
            sentiments = ["Neutral"] * len(df)

        out = pd.DataFrame({
            "URL":       df[url_col].fillna("").astype(str).str.strip(),
            "Content":   df[content_col].fillna("").astype(str).str.strip(),
            "Date":      parsed.dt.strftime("%Y-%m-%d").where(parsed.notna(), ""),
            "Sentiment": sentiments,
            "_sort":     parsed.values,
        })
        if measures_col:
            out["Measures taken"] = df[measures_col].fillna("").astype(str).str.strip()
        out = (
            out.sort_values("_sort", ascending=True, na_position="last")
               .reset_index(drop=True)
               .drop(columns=["_sort"])
        )
        return out

    def process_excel(uploaded_bytes: bytes) -> tuple:
        all_sheets = pd.read_excel(
            io.BytesIO(uploaded_bytes), sheet_name=None, dtype=str
        )
        skipped = []
        processed = {}
        passthrough = {}
        for sheet_name, df in all_sheets.items():
            if sheet_name.lower() == "report":
                passthrough[sheet_name] = df
                continue
            try:
                processed[sheet_name] = process_sheet(df)
            except ValueError as e:
                skipped.append(f"\u26a0\ufe0f Sheet '{sheet_name}' skip olundu: {e}")

        if not processed and not passthrough:
            raise ValueError(
                "Heç bir sheet emal oluna bilmədi.\n" + "\n".join(skipped)
            )

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl", datetime_format="YYYY-MM-DD") as writer:
            # Report sheet-i oldu\u011fu kimi yaz
            for sheet_name, df in passthrough.items():
                df.to_excel(writer, index=False, sheet_name=sheet_name)

            for sheet_name, cleaned in processed.items():
                cleaned.to_excel(writer, index=False, sheet_name=sheet_name)

                ws = writer.sheets[sheet_name]
                header = {cell.value: cell.column for cell in ws[1]}
                url_col_idx  = header.get("URL")
                date_col_idx = header.get("Date")
                hyperlink_font = Font(color="0563C1", underline="single")

                for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                    if url_col_idx:
                        cell = row[url_col_idx - 1]
                        url_val = str(cell.value or "").strip()
                        if url_val.startswith("http"):
                            cell.hyperlink = url_val
                            cell.font = hyperlink_font
                    if date_col_idx:
                        dcell = row[date_col_idx - 1]
                        val = str(dcell.value or "").strip()
                        if val:
                            try:
                                dcell.value = datetime.strptime(val, "%Y-%m-%d")
                                dcell.number_format = "YYYY-MM-DD"
                            except ValueError:
                                pass
        return buf.getvalue(), skipped

    st.write(
        "Excel yüklə → bütün sheet-lər ayrı-ayrı eşlal olunacaq. "
        "Hər sheet-də çıxış: URL, Content, Date (YYYY-MM-DD), Sentiment."
    )

    uploaded = st.file_uploader("Excel faylını seç (.xlsx)", type=["xlsx"])

    col1, col2 = st.columns(2)
    with col1:
        run_btn = st.button("Təmizlə və hazırla", type="primary")
    with col2:
        st.caption("Çıxış: hər sheet ayrıca → URL · Content · Date · Sentiment")

    if run_btn:
        if not uploaded:
            st.error("Əvvəl Excel faylını yüklə.")
        else:
            try:
                result, skipped = process_excel(uploaded.getvalue())
                for msg in skipped:
                    st.warning(msg)
                st.success("Hazırdır ✅")
                st.download_button(
                    label="Cleaned Excel-i yüklə",
                    data=result,
                    file_name="cleaned_final.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error(str(e))
