"""
Katalog Fiyat Generator — Web Arayüzü
=====================================
PDF kataloğa Excel'den fiyat ekleyen Streamlit uygulaması.
Çalıştırma: streamlit run streamlit_app.py
"""
from __future__ import annotations

import io
import tempfile
from pathlib import Path

import pdfplumber
import streamlit as st

from extract_codes import build_template, extract_codes_with_pages
from inject_prices import (
    CODE_RE,
    apply_fuzzy_matches,
    build_dummy_prices,
    build_overlay,
    load_prices,
    merge,
)

st.set_page_config(
    page_title="Katalog Fiyat Generator",
    page_icon="🏷️",
    layout="centered",
)

st.title("🏷️ Katalog Fiyat Generator")
st.caption("PDF kataloğa Excel'den fiyat eklemek için bir araç.")


def _write_temp(data: bytes, suffix: str) -> Path:
    """Yüklenen byte'ları geçici bir dosyaya yazıp Path döner."""
    tf = tempfile.NamedTemporaryFile(suffix=suffix, delete=False)
    tf.write(data)
    tf.close()
    return Path(tf.name)


# -------- 1. PDF yükle --------
st.subheader("1. Katalog PDF'ini yükle")
pdf_file = st.file_uploader("PDF dosyası", type=["pdf"], key="pdf_uploader")

pdf_path: Path | None = None
if pdf_file is not None:
    pdf_bytes = pdf_file.getvalue()
    pdf_path = _write_temp(pdf_bytes, ".pdf")
    st.success(f"✓ {pdf_file.name} ({len(pdf_bytes) / 1024:.0f} KB)")

# -------- 2. Şablon indir --------
st.subheader("2. Fiyat şablonunu indir (opsiyonel)")
st.caption("PDF'teki tüm ürün kodlarını ön-doldurulmuş Excel olarak indirir.")
gen_template = st.button("🔍 Kodları çıkar ve şablon hazırla", disabled=pdf_path is None)

if gen_template and pdf_path is not None:
    with st.spinner("PDF taranıyor..."):
        codes = extract_codes_with_pages(pdf_path)
        out_xlsx = Path(tempfile.mktemp(suffix=".xlsx"))
        build_template(codes, out_xlsx)
        xlsx_bytes = out_xlsx.read_bytes()
    st.success(f"{len(codes)} adet ürün kodu bulundu.")
    st.download_button(
        "📥 fiyat-listesi.xlsx indir",
        data=xlsx_bytes,
        file_name="fiyat-listesi.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# -------- 3. Excel yükle --------
st.subheader("3. Doldurduğun fiyat listesini yükle")
xlsx_file = st.file_uploader("Excel dosyası", type=["xlsx"], key="xlsx_uploader")

xlsx_path: Path | None = None
if xlsx_file is not None:
    xlsx_path = _write_temp(xlsx_file.getvalue(), ".xlsx")
    st.success(f"✓ {xlsx_file.name}")

# -------- 4. Üret --------
st.subheader("4. Fiyatlı PDF üret")
col1, col2 = st.columns(2)
with col1:
    do_real = st.button(
        "🚀 Fiyatlı PDF üret",
        type="primary",
        disabled=pdf_path is None or xlsx_path is None,
        use_container_width=True,
    )
with col2:
    do_preview = st.button(
        "👁️ Dummy fiyatla önizle",
        disabled=pdf_path is None,
        use_container_width=True,
        help="Excel yüklemeden, sahte fiyatlarla nasıl görüneceğini görmek için.",
    )

if (do_real or do_preview) and pdf_path is not None:
    with st.spinner("Fiyatlı PDF hazırlanıyor..."):
        if do_real and xlsx_path is not None:
            prices = load_prices(xlsx_path)
            mode = "excel"
        else:
            prices = build_dummy_prices(pdf_path)
            mode = "dummy"

        if not prices:
            st.error("Excel'de fiyatlı satır bulunamadı.")
        else:
            # Fuzzy match (typo'lu kodlar için)
            pdf_codes: set[str] = set()
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text() or ""
                    for m in CODE_RE.finditer(text):
                        pdf_codes.add(m.group(0))
            fuzzy = apply_fuzzy_matches(prices, pdf_codes, max_distance=1)

            # Overlay + merge
            overlay_bytes, log, stats = build_overlay(pdf_path, prices, None)
            out_path = Path(tempfile.mktemp(suffix=".pdf"))
            merge(pdf_path, overlay_bytes, out_path)
            out_bytes = out_path.read_bytes()

        total_placed = sum(len(v) for v in log.values())
        st.success(
            f"✓ {total_placed} fiyat etiketi basıldı ({len(log)} sayfada)"
        )

        if fuzzy:
            with st.expander(f"Fuzzy eşleşmeler ({len(fuzzy)})", expanded=False):
                for pdf_c, ex_c, d in fuzzy:
                    st.text(f"PDF '{pdf_c}' ↔ Excel '{ex_c}' (mesafe={d})")

        base_name = Path(pdf_file.name).stem if pdf_file else "katalog"
        out_name = f"{base_name} - FIYATLI{'_PREVIEW' if mode == 'dummy' else ''}.pdf"
        st.download_button(
            "📥 Fiyatlı PDF indir",
            data=out_bytes,
            file_name=out_name,
            mime="application/pdf",
            type="primary",
        )

st.divider()
with st.expander("ℹ️ Nasıl kullanılır"):
    st.markdown(
        """
        1. **Katalog PDF'ini yükle** — orijinal, fiyatsız PDF.
        2. **Şablon hazırla** (ilk kez kullanıyorsan) — PDF'teki tüm kodları
           ön-doldurulmuş `fiyat-listesi.xlsx` olarak indir.
        3. Excel'de **Fiyat** sütununu doldur. Para Birimi default `TL`,
           istersen `USD`/`EUR` yazabilirsin. Boş bıraktığın kodlara fiyat
           basılmaz.
        4. Doldurduğun Excel'i yükle ve **Üret** butonuna bas.
        5. Çıkan PDF'i indir.

        **Otomatik özellikler:**
        - Fiyat ürün başlığının altında, sağa yaslı basılır.
        - Spec tablosu satırlarına dokunmaz.
        - Renkli arkaplan (turuncu banner vb.) varsa fiyatı kodun soluna alır.
        - PDF'te yazım hatalı kod varsa (örn. `MD.7AL.8030` yerine
          `MD.7AIL.8030`) otomatik eşler — Excel'de doğru kodu kullan.
        """
    )
