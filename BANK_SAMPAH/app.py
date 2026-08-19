import os
import shutil
from datetime import date, datetime
from pathlib import Path
from io import BytesIO

import pandas as pd
import streamlit as st

# PDF
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageBreak
)


# ============================================================
# CONFIGURATION
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
DATABASE_DIR = BASE_DIR / "database"
EXCEL_FILE = DATABASE_DIR / "bank_sampah.xlsx"
BACKUP_DIR = DATABASE_DIR / "backup"

DATABASE_DIR.mkdir(exist_ok=True)
BACKUP_DIR.mkdir(exist_ok=True)


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Bank Sampah",
    page_icon="♻️",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ============================================================
# CONSTANTS
# ============================================================

SHEET_NASABAH = "DATA NASABAH"
SHEET_MASTER_ITEM = "MASTER ITEM"
SHEET_HARGA = "DAFTAR HARGA PENJUALAN"
SHEET_SETORAN = "TRANSAKSI SETORAN"
SHEET_DETAIL_SETORAN = "DETAIL SETORAN"
SHEET_PENJUALAN = "TRANSAKSI PENJUALAN"
SHEET_DETAIL_PENJUALAN = "DETAIL PENJUALAN"
SHEET_STOK = "MUTASI STOK"
SHEET_KEUANGAN = "REKAPITULASI KEUANGAN"


# ============================================================
# REQUIRED COLUMNS
# ============================================================

REQUIRED_COLUMNS = {

    SHEET_NASABAH: [
        "NIK",
        "NAMA",
        "RT",
        "RW",
        "ALAMAT",
        "HP",
        "REKENING",
        "TANGGAL PENDAFTARAN"
    ],

    SHEET_MASTER_ITEM: [
        "ITEM_ID",
        "JENIS",
        "ITEM",
        "SATUAN"
    ],

    SHEET_HARGA: [
        "HARGA_ID",
        "TANGGAL MULAI BERLAKU",
        "ITEM_ID",
        "ITEM",
        "HARGA SETOR"
    ],

    SHEET_SETORAN: [
        "TRANSAKSI_ID",
        "NO",
        "TANGGAL TRANSAKSI",
        "NIK",
        "NAMA",
        "TOTAL KG",
        "TOTAL HARGA",
        "POTONGAN KAS",
        "SALDO NASABAH",
        "SUDAH DIBAYAR",     # <--- INI TAMBAHAN BARU
        "SISA SALDO",        # <--- INI TAMBAHAN BARU
        "STATUS PEMBAYARAN",
        "TANGGAL LUNAS"
    ],

    SHEET_DETAIL_SETORAN: [
        "DETAIL_ID",
        "TRANSAKSI_ID",
        "ITEM_ID",
        "ITEM",
        "KG",
        "HARGA",
        "SUBTOTAL",
        "STATUS PEMBAYARAN",  # <--- INI TAMBAHANNYA
        "TANGGAL LUNAS"       # <--- INI TAMBAHANNYA
    ],

    SHEET_PENJUALAN: [
        "PENJUALAN_ID",
        "TANGGAL",
        "PENGEPUL",
        "TOTAL KG",
        "TOTAL HARGA"
    ],

    SHEET_DETAIL_PENJUALAN: [
        "DETAIL_ID",
        "PENJUALAN_ID",
        "ITEM_ID",
        "ITEM",
        "KG",
        "HARGA",
        "SUBTOTAL"
    ],

    SHEET_STOK: [
        "MUTASI_ID",
        "TANGGAL",
        "ITEM_ID",
        "ITEM",
        "JENIS_MUTASI",
        "KG",
        "REFERENSI"
    ],

    SHEET_KEUANGAN: [
        "KEUANGAN_ID",
        "TANGGAL TRANSAKSI",
        "KETERANGAN",
        "DEBIT",
        "KREDIT",
        "REFERENSI"
    ]
}


# ============================================================
# DATABASE
# ============================================================

def check_database():
    return EXCEL_FILE.exists()


def ensure_columns(df, sheet_name):

    required = REQUIRED_COLUMNS.get(
        sheet_name,
        []
    )

    for col in required:

        if col not in df.columns:

            if col in [
                "STATUS PEMBAYARAN"
            ]:

                df[col] = "BELUM LUNAS"

            elif col in [
                "POTONGAN KAS",
                "SALDO NASABAH"
            ]:

                df[col] = 0

            else:

                df[col] = ""

    return df


@st.cache_data(ttl=2)
def load_sheet(sheet_name):

    if not EXCEL_FILE.exists():
        return pd.DataFrame()

    try:

        df = pd.read_excel(
            EXCEL_FILE,
            sheet_name=sheet_name
        )

        df = df.loc[
            :,
            ~df.columns.astype(str).str.startswith(
                "Unnamed"
            )
        ]

        df = ensure_columns(
            df,
            sheet_name
        )

        # ====================================================
        # TEXT
        # ====================================================

        text_columns = [

            "NIK",
            "NAMA",
            "RT",
            "RW",
            "ALAMAT",
            "HP",
            "REKENING",

            "ITEM_ID",
            "ITEM",
            "JENIS",
            "SATUAN",

            "TRANSAKSI_ID",
            "DETAIL_ID",

            "PENJUALAN_ID",

            "MUTASI_ID",
            "JENIS_MUTASI",
            "REFERENSI",

            "KEUANGAN_ID",
            "KETERANGAN",

            "STATUS PEMBAYARAN"
        ]

        for col in text_columns:

            if col in df.columns:

                df[col] = (
                    df[col]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                )

        # ====================================================
        # NUMERIC
        # ====================================================

        numeric_columns = [
            "KG",
            "HARGA",
            "SUBTOTAL",
            "TOTAL KG",
            "TOTAL HARGA",
            "POTONGAN KAS",
            "SALDO NASABAH",
            "SUDAH DIBAYAR",  # <--- INI TAMBAHAN BARU
            "SISA SALDO",     # <--- INI TAMBAHAN BARU
            "HARGA SETOR",
            "DEBIT",
            "KREDIT",
            "NO"
        ]

        for col in numeric_columns:

            if col in df.columns:

                df[col] = pd.to_numeric(
                    df[col],
                    errors="coerce"
                ).fillna(0)

        # ====================================================
        # DATE
        # ====================================================

        date_columns = [

            "TANGGAL",
            "TANGGAL TRANSAKSI",
            "TANGGAL PENDAFTARAN",
            "TANGGAL MULAI BERLAKU",
            "TANGGAL LUNAS"
        ]

        for col in date_columns:

            if col in df.columns:

                df[col] = pd.to_datetime(
                    df[col],
                    errors="coerce"
                )

        return df

    except Exception as e:

        st.error(
            f"Gagal membaca sheet {sheet_name}: {e}"
        )

        return pd.DataFrame()


def load_all_data():

    data = {}

    for sheet in REQUIRED_COLUMNS:

        data[sheet] = load_sheet(sheet)

    return data


def create_backup():

    if not EXCEL_FILE.exists():
        return

    timestamp = datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    backup_file = (
        BACKUP_DIR
        / f"bank_sampah_{timestamp}.xlsx"
    )

    try:

        shutil.copy2(
            EXCEL_FILE,
            backup_file
        )

    except Exception:
        pass


def save_sheet(sheet_name, df):

    try:

        create_backup()

        with pd.ExcelWriter(
            EXCEL_FILE,
            engine="openpyxl",
            mode="a",
            if_sheet_exists="replace"
        ) as writer:

            df.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False
            )

        load_sheet.clear()

        return True

    except Exception as e:

        st.error(
            f"Gagal menyimpan data: {e}"
        )

        return False


def save_multiple_sheets(sheet_data):

    try:

        create_backup()

        with pd.ExcelWriter(
            EXCEL_FILE,
            engine="openpyxl",
            mode="a",
            if_sheet_exists="replace"
        ) as writer:

            for sheet_name, df in sheet_data.items():

                df.to_excel(
                    writer,
                    sheet_name=sheet_name,
                    index=False
                )

        load_sheet.clear()

        return True

    except Exception as e:

        st.error(
            f"Gagal menyimpan Excel: {e}"
        )

        return False


# ============================================================
# HELPER
# ============================================================

def normalize_text(value):

    if pd.isna(value):
        return ""

    return str(value).strip()


def format_rupiah(value):

    try:

        value = float(value)

        return (
            "Rp "
            + f"{value:,.0f}".replace(
                ",",
                "."
            )
        )

    except Exception:

        return "Rp 0"


def generate_id(
    df,
    column,
    prefix,
    digits=6
):

    if df.empty or column not in df.columns:

        return f"{prefix}{1:0{digits}d}"

    values = (
        df[column]
        .dropna()
        .astype(str)
    )

    numbers = []

    for value in values:

        try:

            number = int(
                value.replace(
                    prefix,
                    ""
                )
            )

            numbers.append(
                number
            )

        except Exception:

            continue

    next_number = (
        max(numbers) + 1
        if numbers
        else 1
    )

    return (
        f"{prefix}"
        f"{next_number:0{digits}d}"
    )


def get_next_no(df):

    if df.empty:
        return 1

    if "NO" not in df.columns:
        return 1

    try:

        return int(
            pd.to_numeric(
                df["NO"],
                errors="coerce"
            ).max()
        ) + 1

    except Exception:

        return len(df) + 1


# ============================================================
# PRICE
# ============================================================

def get_price(
    item_id,
    tanggal_transaksi
):

    df_harga = load_sheet(
        SHEET_HARGA
    )

    if df_harga.empty:
        return None

    df_harga = df_harga.copy()

    df_harga[
        "TANGGAL MULAI BERLAKU"
    ] = pd.to_datetime(
        df_harga[
            "TANGGAL MULAI BERLAKU"
        ],
        errors="coerce"
    )

    tanggal_transaksi = pd.to_datetime(
        tanggal_transaksi
    )

    result = df_harga[
        (
            df_harga[
                "ITEM_ID"
            ].astype(str)
            == str(item_id)
        )
        &
        (
            df_harga[
                "TANGGAL MULAI BERLAKU"
            ]
            <= tanggal_transaksi
        )
    ].copy()

    if result.empty:
        return None

    result = result.sort_values(
        "TANGGAL MULAI BERLAKU",
        ascending=False
    )

    try:

        return float(
            result.iloc[0][
                "HARGA SETOR"
            ]
        )

    except Exception:

        return None


# ============================================================
# STOCK
# ============================================================

def get_stock():

    df = load_sheet(
        SHEET_STOK
    )

    df_item = load_sheet(
        SHEET_MASTER_ITEM
    )

    if df_item.empty:
        return pd.DataFrame()

    result = df_item[
        [
            "ITEM_ID",
            "JENIS",
            "ITEM",
            "SATUAN"
        ]
    ].copy()

    if df.empty:

        result["MASUK"] = 0
        result["KELUAR"] = 0
        result["STOK"] = 0

        return result

    df["KG"] = pd.to_numeric(
        df["KG"],
        errors="coerce"
    ).fillna(0)

    masuk = (
        df[
            df["JENIS_MUTASI"]
            == "MASUK"
        ]
        .groupby(
            "ITEM_ID"
        )["KG"]
        .sum()
    )

    keluar = (
        df[
            df["JENIS_MUTASI"]
            == "KELUAR"
        ]
        .groupby(
            "ITEM_ID"
        )["KG"]
        .sum()
    )

    result["MASUK"] = (
        result["ITEM_ID"]
        .map(masuk)
        .fillna(0)
    )

    result["KELUAR"] = (
        result["ITEM_ID"]
        .map(keluar)
        .fillna(0)
    )

    result["STOK"] = (
        result["MASUK"]
        - result["KELUAR"]
    )

    return result


# ============================================================
# PDF MUTASI
# ============================================================

def generate_mutasi_pdf(
    nasabah,
    transaksi,
    detail
):

    buffer = BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=1.2 * cm,
        leftMargin=1.2 * cm,
        topMargin=1.2 * cm,
        bottomMargin=1.2 * cm
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "TitleCustom",
        parent=styles["Title"],
        alignment=TA_CENTER,
        fontSize=16,
        leading=20,
        spaceAfter=10
    )

    normal_style = ParagraphStyle(
        "NormalCustom",
        parent=styles["Normal"],
        fontSize=9,
        leading=12
    )

    elements = []

    elements.append(
        Paragraph(
            "BANK SAMPAH",
            title_style
        )
    )

    elements.append(
        Paragraph(
            "MUTASI SALDO NASABAH",
            title_style
        )
    )

    elements.append(
        Spacer(1, 0.2 * cm)
    )

    # --------------------------------------------------------
    # IDENTITAS
    # --------------------------------------------------------
    
    # --- LOGIKA SENSOR NIK UNTUK PDF ---
    nik_asli = str(nasabah["NIK"])
    nik_sensor = nik_asli[:6] + "******" + nik_asli[-4:] if len(nik_asli) == 16 else "******"
    # -----------------------------------

    identity = [
        [
            Paragraph("<b>NIK</b>", normal_style),
            nik_sensor  # <--- Menggunakan NIK yang sudah disensor
        ],
        [
            Paragraph("<b>Nama</b>", normal_style),
            str(nasabah["NAMA"])
        ],
        [
            Paragraph("<b>Alamat</b>", normal_style),
            str(nasabah["ALAMAT"])
        ]
    ]

    identity_table = Table(
        identity,
        colWidths=[3 * cm, 14 * cm]
    )

    identity_table.setStyle(
        TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5)
        ])
    )

    elements.append(identity_table)
    elements.append(Spacer(1, 0.3 * cm))

    # --------------------------------------------------------
    # SUMMARY
    # --------------------------------------------------------

    # Handle kompatibilitas dengan data excel lama
    if "SUDAH DIBAYAR" not in transaksi.columns:
        transaksi["SUDAH DIBAYAR"] = 0
    if "SISA SALDO" not in transaksi.columns:
        transaksi["SISA SALDO"] = transaksi["SALDO NASABAH"]

    total_kg = pd.to_numeric(transaksi["TOTAL KG"], errors="coerce").fillna(0).sum()
    total_bruto = pd.to_numeric(transaksi["TOTAL HARGA"], errors="coerce").fillna(0).sum()
    total_kas = pd.to_numeric(transaksi["POTONGAN KAS"], errors="coerce").fillna(0).sum()
    total_saldo = pd.to_numeric(transaksi["SALDO NASABAH"], errors="coerce").fillna(0).sum()
    total_dibayar = pd.to_numeric(transaksi["SUDAH DIBAYAR"], errors="coerce").fillna(0).sum()
    total_sisa = pd.to_numeric(transaksi["SISA SALDO"], errors="coerce").fillna(0).sum()

    summary = [
        ["Total Sampah", f"{total_kg:,.2f} KG"],
        ["Total Bruto", format_rupiah(total_bruto)],
        ["Potongan Kas 10%", format_rupiah(total_kas)],
        ["Hak Nasabah (90%)", format_rupiah(total_saldo)],
        ["Sudah Ditarik/Dibayar", format_rupiah(total_dibayar)],
        ["Sisa Belum Dibayar", format_rupiah(total_sisa)]
    ]

    summary_table = Table(
        summary,
        colWidths=[7 * cm, 10 * cm]
    )

    summary_table.setStyle(
        TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
            ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
            ("ALIGN", (1, 0), (1, -1), "RIGHT"),
            ("BACKGROUND", (0, 0), (-1, -1), colors.whitesmoke)
        ])
    )

    elements.append(summary_table)
    elements.append(Spacer(1, 0.5 * cm))

    # --------------------------------------------------------
    # TRANSAKSI
    # --------------------------------------------------------

    transaction_data = [[
        "ID",
        "Tanggal",
        "KG",
        "Bruto",
        "Kas 10%",
        "Hak 90%",
        "Dibayar",
        "Sisa",
        "Status"
    ]]

    for _, row in transaksi.iterrows():

        tanggal = pd.to_datetime(row["TANGGAL TRANSAKSI"], errors="coerce")
        tanggal_text = tanggal.strftime("%d-%m-%Y") if not pd.isna(tanggal) else "-"

        dibayar_val = float(row.get("SUDAH DIBAYAR", 0))
        if pd.isna(dibayar_val): dibayar_val = 0
        
        sisa_val = float(row.get("SISA SALDO", row["SALDO NASABAH"]))
        if pd.isna(sisa_val): sisa_val = float(row["SALDO NASABAH"])

        transaction_data.append([
            str(row["TRANSAKSI_ID"]),
            tanggal_text,
            f"{float(row['TOTAL KG']):,.1f}",
            format_rupiah(row["TOTAL HARGA"]),
            format_rupiah(row["POTONGAN KAS"]),
            format_rupiah(row["SALDO NASABAH"]),
            format_rupiah(dibayar_val),
            format_rupiah(sisa_val),
            str(row["STATUS PEMBAYARAN"])
        ])

    # Sesuaikan lebar kolom agar muat di 1 kertas A4 (Total Max ~18.6 cm)
    transaction_table = Table(
        transaction_data,
        repeatRows=1,
        colWidths=[
            1.8 * cm, # ID
            1.8 * cm, # Tanggal
            1.0 * cm, # KG
            2.1 * cm, # Bruto
            2.0 * cm, # Kas 10%
            2.1 * cm, # Hak 90%
            2.1 * cm, # Dibayar
            2.1 * cm, # Sisa
            2.6 * cm  # Status
        ]
    )

    transaction_table.setStyle(
        TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 6),
            ("ALIGN", (2, 1), (-1, -1), "RIGHT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (-1, 0), "CENTER") # Judul kolom ditengah
        ])
    )

    elements.append(Paragraph("Riwayat Transaksi & Pembayaran", styles["Heading3"]))
    elements.append(transaction_table)

    # --------------------------------------------------------
    # DETAIL ITEM
    # --------------------------------------------------------

    if not detail.empty:

        elements.append(Spacer(1, 0.5 * cm))
        elements.append(Paragraph("Rekap Jenis Sampah Disetor", styles["Heading3"]))

        detail["KG"] = pd.to_numeric(detail["KG"], errors="coerce").fillna(0)
        detail["SUBTOTAL"] = pd.to_numeric(detail["SUBTOTAL"], errors="coerce").fillna(0)

        rekap = detail.groupby("ITEM", as_index=False).agg(
            KG=("KG", "sum"),
            NILAI=("SUBTOTAL", "sum")
        )

        detail_data = [["Item", "Total KG", "Nilai Bruto"]]

        for _, row in rekap.iterrows():
            detail_data.append([
                str(row["ITEM"]),
                f"{row['KG']:,.2f}",
                format_rupiah(row["NILAI"])
            ])

        detail_table = Table(
            detail_data,
            repeatRows=1,
            colWidths=[8 * cm, 4 * cm, 5 * cm]
        )

        detail_table.setStyle(
            TableStyle([
                ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
                ("FONTSIZE", (0, 0), (-1, -1), 8)
            ])
        )

        elements.append(detail_table)

    elements.append(Spacer(1, 1 * cm))
    elements.append(Paragraph("Dokumen ini merupakan rekapitulasi mutasi setoran dan pembayaran nasabah.", normal_style))
    elements.append(Spacer(1, 0.5 * cm))
    elements.append(Paragraph(f"Dicetak: {datetime.now().strftime('%d-%m-%Y %H:%M')}", normal_style))

    doc.build(elements)
    buffer.seek(0)

    return buffer

# ============================================================
# SIDEBAR
# ============================================================

def sidebar():

    col_logo, col_text = st.sidebar.columns([1, 3])
    with col_logo:
        st.image("assets/logo.png", width=500)  # Sesuaikan ukuran width jika kurang pas
    with col_text:
        st.markdown("### **BANK SAMPAH**<br>**CENDIKIA ARUTALA**", unsafe_allow_html=True)

    st.sidebar.caption("Sistem Informasi Pendataan Bank Sampah")
    st.sidebar.divider()

    # --- KODE ANTI-ERROR STREAMLIT ---
    # Jika ada pesanan pindah halaman dari tombol "Simpan", 
    # terapkan SEBELUM menu digambar.
    if "pindah_halaman" in st.session_state:
        st.session_state.menu_aktif = st.session_state.pindah_halaman
        del st.session_state.pindah_halaman # Hapus pesan setelah dipakai
    # ---------------------------------

    menu = st.sidebar.radio(
        "MENU",
        [
            "🏠 Dashboard",
            "👥 Nasabah",
            "🗂️ Master Item",
            "💰 Harga Sampah",
            "📥 Transaksi Setoran",
            "📤 Transaksi Penjualan",
            "📦 Stok",
            "💵 Keuangan"
        ],
        key="menu_aktif"
    )

    st.sidebar.divider()

    if EXCEL_FILE.exists():
        st.sidebar.success("Database Excel terhubung")
    else:
        st.sidebar.error("Database Excel tidak ditemukan")

    return menu

# ============================================================
# DASHBOARD
# ============================================================

def page_dashboard():

    st.title(
        "🏠 Dashboard Bank Sampah"
    )

    st.caption(
        "Ringkasan aktivitas Bank Sampah"
    )

    data = load_all_data()

    df_nasabah = data[
        SHEET_NASABAH
    ]

    df_setoran = data[
        SHEET_SETORAN
    ]

    df_penjualan = data[
        SHEET_PENJUALAN
    ]

    df_stok = get_stock()

    df_keuangan = data[
        SHEET_KEUANGAN
    ]

    # --------------------------------------------------------
    # KPI
    # --------------------------------------------------------

    total_nasabah = len(
        df_nasabah
    )

    total_setoran_kg = 0
    total_setoran_rp = 0
    total_saldo_nasabah = 0
    total_belum_lunas = 0

    if not df_setoran.empty:

        total_setoran_kg = pd.to_numeric(
            df_setoran[
                "TOTAL KG"
            ],
            errors="coerce"
        ).fillna(0).sum()

        total_setoran_rp = pd.to_numeric(
            df_setoran[
                "TOTAL HARGA"
            ],
            errors="coerce"
        ).fillna(0).sum()

        total_saldo_nasabah = pd.to_numeric(
            df_setoran[
                "SALDO NASABAH"
            ],
            errors="coerce"
        ).fillna(0).sum()

        total_belum_lunas = (
            df_setoran[
                "STATUS PEMBAYARAN"
            ]
            .astype(str)
            .eq("BELUM LUNAS")
            .sum()
        )

    total_penjualan_kg = 0
    total_penjualan_rp = 0

    if not df_penjualan.empty:

        total_penjualan_kg = pd.to_numeric(
            df_penjualan[
                "TOTAL KG"
            ],
            errors="coerce"
        ).fillna(0).sum()

        total_penjualan_rp = pd.to_numeric(
            df_penjualan[
                "TOTAL HARGA"
            ],
            errors="coerce"
        ).fillna(0).sum()

    total_stok = 0

    if not df_stok.empty:

        total_stok = df_stok[
            "STOK"
        ].sum()

    total_debit = 0
    total_kredit = 0

    if not df_keuangan.empty:

        total_debit = pd.to_numeric(
            df_keuangan[
                "DEBIT"
            ],
            errors="coerce"
        ).fillna(0).sum()

        total_kredit = pd.to_numeric(
            df_keuangan[
                "KREDIT"
            ],
            errors="coerce"
        ).fillna(0).sum()

    saldo = (
        total_debit
        - total_kredit
    )

    col1, col2, col3, col4 = st.columns(4)

    col1.metric(
        "Total Nasabah",
        f"{total_nasabah:,}".replace(
            ",",
            "."
        )
    )

    col2.metric(
        "Total Setoran",
        f"{total_setoran_kg:,.2f} KG"
    )

    col3.metric(
        "Saldo Belum Dibayar",
        format_rupiah(
            df_setoran[
                df_setoran[
                    "STATUS PEMBAYARAN"
                ]
                == "BELUM LUNAS"
            ]["SALDO NASABAH"].sum()
            if not df_setoran.empty
            else 0
        )
    )

    col4.metric(
        "Saldo Kas",
        format_rupiah(
            saldo
        )
    )

    st.divider()

    col1, col2 = st.columns(2)

    with col1:

        st.subheader(
            "Setoran per Bulan"
        )

        if not df_setoran.empty:

            df_header = df_setoran.copy()

            df_header[
                "TANGGAL TRANSAKSI"
            ] = pd.to_datetime(
                df_header[
                    "TANGGAL TRANSAKSI"
                ],
                errors="coerce"
            )

            df_header[
                "BULAN"
            ] = (
                df_header[
                    "TANGGAL TRANSAKSI"
                ]
                .dt.to_period("M")
                .astype(str)
            )

            chart = (
                df_header
                .groupby("BULAN")[
                    "TOTAL KG"
                ]
                .sum()
            )

            st.bar_chart(chart)

        else:

            st.info(
                "Belum ada transaksi setoran."
            )

    with col2:

        st.subheader(
            "Stok per Jenis Sampah"
        )

        if not df_stok.empty:

            chart_stock = (
                df_stok
                .groupby("JENIS")[
                    "STOK"
                ]
                .sum()
            )

            st.bar_chart(
                chart_stock
            )

        else:

            st.info(
                "Belum ada data stok."
            )

    st.subheader(
        "Status Pembayaran"
    )

    if not df_setoran.empty:

        belum = (
            df_setoran[
                df_setoran[
                    "STATUS PEMBAYARAN"
                ]
                == "BELUM LUNAS"
            ][
                "SALDO NASABAH"
            ]
            .sum()
        )

        lunas = (
            df_setoran[
                df_setoran[
                    "STATUS PEMBAYARAN"
                ]
                == "LUNAS"
            ][
                "SALDO NASABAH"
            ]
            .sum()
        )

        c1, c2 = st.columns(2)

        c1.metric(
            "Belum Lunas",
            format_rupiah(
                belum
            )
        )

        c2.metric(
            "Sudah Lunas",
            format_rupiah(
                lunas
            )
        )


# ============================================================
# NASABAH
# ============================================================

def page_nasabah():

    st.title(
        "👥 Data Nasabah"
    )

    df = load_sheet(
        SHEET_NASABAH
    )

    df_nasabah = df.copy()

    tab1, tab2, tab3, tab4 = st.tabs(
        [
            "📋 Data Nasabah",
            "➕ Tambah Nasabah",
            "✏️ Edit / Hapus",
            "📜 Detail Transaksi"
        ]
    )

    # ========================================================
    # TAB DATA
    # ========================================================

    with tab1:

        if df.empty:

            st.info(
                "Belum ada data nasabah."
            )

        else:

            search = st.text_input(
                "🔎 Cari NIK / Nama"
            )

            display = df.copy()

            if search:

                mask = (
                    display[
                        "NIK"
                    ]
                    .astype(str)
                    .str.contains(
                        search,
                        case=False,
                        na=False
                    )
                    |
                    display[
                        "NAMA"
                    ]
                    .astype(str)
                    .str.contains(
                        search,
                        case=False,
                        na=False
                    )
                )

                display = display[
                    mask
                ]

            st.dataframe(
                display,
                use_container_width=True,
                hide_index=True
            )

            st.caption(
                f"Jumlah nasabah: "
                f"{len(display)}"
            )

    # ========================================================
    # TAMBAH NASABAH
    # ========================================================

    with tab2:

        with st.form(
            "form_tambah_nasabah"
        ):

            col1, col2 = st.columns(2)

            nik = col1.text_input(
                "NIK",
                max_chars=16,
                help="Masukkan tepat 16 angka NIK"
            )
            
            nama = col2.text_input(
                "Nama"
            )

            rt = col1.text_input(
                "RT"
            )

            rw = col2.text_input(
                "RW"
            )

            alamat = st.text_area(
                "Alamat"
            )

            col1, col2, col3 = st.columns(3)

            hp = col1.text_input(
                "HP"
            )

            rekening = col2.text_input(
                "Rekening"
            )

            tanggal = col3.date_input(
                "Tanggal Pendaftaran",
                value=date.today()
            )

            submit = st.form_submit_button(
                "💾 Simpan Nasabah",
                use_container_width=True
            )

        if submit:
            
            # --- MULAI DARI SINI HARUS MENJOROK KE DALAM ---
            nik = nik.strip()
            nama = nama.strip()

            if not nik:
                st.error("NIK wajib diisi.")
                
            elif not nik.isdigit():
                st.error("NIK tidak valid! NIK hanya boleh berisi angka.")
                
            elif len(nik) != 16:
                st.error(f"NIK harus tepat 16 angka! (Saat ini: {len(nik)} angka)")

            elif not nama:
                st.error("Nama wajib diisi.")

            elif not df.empty and (
                df["NIK"]
                .astype(str)
                .str.strip()
                .eq(nik)
                .any()
            ):
                st.error("NIK sudah terdaftar.")
            
            else:
                new_data = pd.DataFrame(
                    [[
                        nik,
                        nama,
                        rt,
                        rw,
                        alamat,
                        hp,
                        rekening,
                        tanggal.strftime("%Y-%m-%d")
                    ]],
                    columns=REQUIRED_COLUMNS[SHEET_NASABAH]
                )

                df = pd.concat([df, new_data], ignore_index=True)

                if save_sheet(SHEET_NASABAH, df):
                    st.success("Nasabah berhasil ditambahkan.")
                    st.rerun()

    # ========================================================
    # EDIT / HAPUS NASABAH
    # ========================================================

    with tab3:

        if df.empty:

            st.info(
                "Belum ada data nasabah."
            )

        else:

            nik_list = (
                df["NIK"]
                .astype(str)
                .tolist()
            )

            selected_nik = st.selectbox(
                "Pilih NIK",
                nik_list,
                key="edit_nasabah_select"
            )

            index = df[
                df["NIK"].astype(str)
                == selected_nik
            ].index[0]

            row = df.loc[index]

            with st.form(
                "form_edit_nasabah"
            ):

                col1, col2 = st.columns(2)

                nama = col1.text_input(
                    "Nama",
                    value=normalize_text(
                        row["NAMA"]
                    )
                )

                rt = col2.text_input(
                    "RT",
                    value=normalize_text(
                        row["RT"]
                    )
                )

                rw = col1.text_input(
                    "RW",
                    value=normalize_text(
                        row["RW"]
                    )
                )

                alamat = col2.text_input(
                    "Alamat",
                    value=normalize_text(
                        row["ALAMAT"]
                    )
                )

                hp = col1.text_input(
                    "HP",
                    value=normalize_text(
                        row["HP"]
                    )
                )

                rekening = col2.text_input(
                    "Rekening",
                    value=normalize_text(
                        row["REKENING"]
                    )
                )

                tanggal_default = pd.to_datetime(
                    row[
                        "TANGGAL PENDAFTARAN"
                    ],
                    errors="coerce"
                )

                if pd.isna(
                    tanggal_default
                ):

                    tanggal_default = date.today()

                else:

                    tanggal_default = (
                        tanggal_default.date()
                    )

                tanggal = st.date_input(
                    "Tanggal Pendaftaran",
                    value=tanggal_default
                )

                col1, col2 = st.columns(2)

                update = col1.form_submit_button(
                    "💾 Update",
                    use_container_width=True
                )

                delete = col2.form_submit_button(
                    "🗑️ Hapus",
                    use_container_width=True
                )

            if update:

                df.at[
                    index,
                    "NAMA"
                ] = nama

                df.at[
                    index,
                    "RT"
                ] = rt

                df.at[
                    index,
                    "RW"
                ] = rw

                df.at[
                    index,
                    "ALAMAT"
                ] = alamat

                df.at[
                    index,
                    "HP"
                ] = hp

                df.at[
                    index,
                    "REKENING"
                ] = rekening

                df.at[
                    index,
                    "TANGGAL PENDAFTARAN"
                ] = tanggal.strftime(
                    "%Y-%m-%d"
                )

                if save_sheet(
                    SHEET_NASABAH,
                    df
                ):

                    st.success(
                        "Data nasabah berhasil diperbarui."
                    )

                    st.rerun()

            if delete:

                df = df.drop(
                    index
                ).reset_index(
                    drop=True
                )

                if save_sheet(
                    SHEET_NASABAH,
                    df
                ):

                    st.success(
                        "Nasabah berhasil dihapus."
                    )

                    st.rerun()

    # ========================================================
    # DETAIL TRANSAKSI
    # ========================================================

    with tab4:

        df_setoran = load_sheet(
            SHEET_SETORAN
        )

        df_detail = load_sheet(
            SHEET_DETAIL_SETORAN
        )

        if df_setoran.empty:

            st.info(
                "Belum ada transaksi setoran."
            )

        else:

            st.subheader(
                "📜 Riwayat Setoran Nasabah"
            )

            nasabah_options = (
                df_nasabah[
                    ["NIK", "NAMA"]
                ]
                .astype(str)
                .apply(
                    lambda x:
                    f"{x['NIK']} - {x['NAMA']}",
                    axis=1
                )
                .tolist()
            )

            selected_nasabah = st.selectbox(
                "Pilih Nasabah",
                nasabah_options,
                key="detail_nasabah"
            )

            selected_nik = (
                selected_nasabah
                .split(" - ")[0]
            )

            # ------------------------------------------------
            # FILTER TANGGAL
            # ------------------------------------------------

            col1, col2 = st.columns(2)

            tanggal_mulai = col1.date_input(
                "Tanggal Mulai",
                value=date(
                    date.today().year,
                    1,
                    1
                ),
                key="filter_tanggal_mulai"
            )

            tanggal_selesai = col2.date_input(
                "Tanggal Selesai",
                value=date.today(),
                key="filter_tanggal_selesai"
            )

            # ------------------------------------------------
            # FILTER STATUS
            # ------------------------------------------------

            status_filter = st.selectbox(
                "Status Pembayaran",
                [
                    "SEMUA",
                    "BELUM LUNAS",
                    "LUNAS"
                ],
                key="filter_status_nasabah"
            )

            # ------------------------------------------------
            # FILTER ITEM
            # ------------------------------------------------

            item_options = ["SEMUA"]

            if not df_detail.empty:

                item_options += sorted(
                    df_detail[
                        "ITEM"
                    ]
                    .dropna()
                    .astype(str)
                    .unique()
                    .tolist()
                )

            selected_item = st.selectbox(
                "Filter Item",
                item_options,
                key="filter_item_nasabah"
            )

            # ------------------------------------------------
            # TRANSAKSI NASABAH
            # ------------------------------------------------

            transaksi = df_setoran[
                df_setoran[
                    "NIK"
                ].astype(str)
                == str(selected_nik)
            ].copy()

            if transaksi.empty:

                st.info(
                    "Nasabah ini belum memiliki transaksi."
                )

            else:

                transaksi[
                    "TANGGAL TRANSAKSI"
                ] = pd.to_datetime(
                    transaksi[
                        "TANGGAL TRANSAKSI"
                    ],
                    errors="coerce"
                )

                transaksi = transaksi[
                    (
                        transaksi[
                            "TANGGAL TRANSAKSI"
                        ].dt.date
                        >= tanggal_mulai
                    )
                    &
                    (
                        transaksi[
                            "TANGGAL TRANSAKSI"
                        ].dt.date
                        <= tanggal_selesai
                    )
                ]

                if status_filter != "SEMUA":

                    transaksi = transaksi[
                        transaksi[
                            "STATUS PEMBAYARAN"
                        ]
                        == status_filter
                    ]

                if selected_item != "SEMUA":

                    if not df_detail.empty:

                        detail_filter = df_detail[
                            df_detail[
                                "ITEM"
                            ].astype(str)
                            == selected_item
                        ][
                            "TRANSAKSI_ID"
                        ].astype(str).unique()

                        transaksi = transaksi[
                            transaksi[
                                "TRANSAKSI_ID"
                            ]
                            .astype(str)
                            .isin(
                                detail_filter
                            )
                        ]

                # ------------------------------------------------
                # SUMMARY (UI DIPERBARUI)
                # ------------------------------------------------
                
                # Handle kompatibilitas dengan data excel lama (jika ada)
                if "SUDAH DIBAYAR" not in transaksi.columns:
                    transaksi["SUDAH DIBAYAR"] = 0
                if "SISA SALDO" not in transaksi.columns:
                    transaksi["SISA SALDO"] = transaksi["SALDO NASABAH"]

                total_transaksi = len(transaksi)
                total_kg = pd.to_numeric(transaksi["TOTAL KG"], errors="coerce").fillna(0).sum()
                total_bruto = pd.to_numeric(transaksi["TOTAL HARGA"], errors="coerce").fillna(0).sum()
                
                total_saldo = pd.to_numeric(transaksi["SALDO NASABAH"], errors="coerce").fillna(0).sum()
                total_dibayar = pd.to_numeric(transaksi["SUDAH DIBAYAR"], errors="coerce").fillna(0).sum()
                total_sisa = pd.to_numeric(transaksi["SISA SALDO"], errors="coerce").fillna(0).sum()

                st.markdown("##### 📊 Ringkasan Aktivitas Nasabah")
                col1, col2, col3 = st.columns(3)
                
                col1.metric("Jumlah Transaksi", total_transaksi)
                col2.metric("Total Berat Sampah", f"{total_kg:,.2f} KG")
                col3.metric("Total Uang Bruto", format_rupiah(total_bruto))
                
                st.markdown("##### 💰 Status Keuangan Nasabah")
                col4, col5, col6 = st.columns(3)
                
                col4.metric("Total Hak Nasabah (90%)", format_rupiah(total_saldo))
                col5.metric("Sudah Ditarik/Dibayar", format_rupiah(total_dibayar))
                
                # Highlight sisa saldo jika masih ada
                if total_sisa > 0:
                    col6.error(f"Sisa Belum Dibayar: \n{format_rupiah(total_sisa)}")
                else:
                    col6.success(f"Sisa Belum Dibayar: \n{format_rupiah(total_sisa)}")

                st.divider()

                # ------------------------------------------------
                # TABEL (UI DIPERBARUI)
                # ------------------------------------------------

                transaksi_display = transaksi.copy()

                transaksi_display["TANGGAL TRANSAKSI"] = transaksi_display["TANGGAL TRANSAKSI"].dt.strftime("%d-%m-%Y")
                transaksi_display["TOTAL HARGA"] = transaksi_display["TOTAL HARGA"].apply(format_rupiah)
                transaksi_display["POTONGAN KAS"] = transaksi_display["POTONGAN KAS"].apply(format_rupiah)
                transaksi_display["SALDO NASABAH"] = transaksi_display["SALDO NASABAH"].apply(format_rupiah)
                
                # Format rupiah untuk kolom baru
                transaksi_display["SUDAH DIBAYAR"] = pd.to_numeric(transaksi_display["SUDAH DIBAYAR"], errors="coerce").fillna(0).apply(format_rupiah)
                transaksi_display["SISA SALDO"] = pd.to_numeric(transaksi_display["SISA SALDO"], errors="coerce").fillna(0).apply(format_rupiah)

                st.dataframe(
                    transaksi_display[
                        [
                            "TRANSAKSI_ID",
                            "TANGGAL TRANSAKSI",
                            "TOTAL KG",
                            "TOTAL HARGA",
                            "SALDO NASABAH",
                            "SUDAH DIBAYAR",
                            "SISA SALDO",
                            "STATUS PEMBAYARAN"
                        ]
                    ],
                    use_container_width=True,
                    hide_index=True
                )

                # ------------------------------------------------
                # PELUNASAN NOMINAL BEBAS (BISA CICILAN)
                # ------------------------------------------------
                st.subheader("💵 Pembayaran Saldo Nasabah")

                # Kita gunakan data TRANSAKSI UTAMA
                belum_lunas = transaksi[transaksi["STATUS PEMBAYARAN"] == "BELUM LUNAS"].copy()

                if belum_lunas.empty:
                    st.success("Tidak ada transaksi yang belum lunas pada periode/filter ini.")
                else:
                    # Pengecekan aman untuk data Excel lama yang belum ada kolom SISA SALDO
                    if "SISA SALDO" not in belum_lunas.columns:
                        belum_lunas["SISA SALDO"] = belum_lunas["SALDO NASABAH"]

                    pilihan_lunas = belum_lunas.apply(
                        lambda r: f"{r['TRANSAKSI_ID']} | Tgl: {pd.to_datetime(r['TANGGAL TRANSAKSI']).strftime('%d-%m-%Y')} | Sisa: {format_rupiah(float(r['SISA SALDO']) if pd.notna(r.get('SISA SALDO')) else float(r['SALDO NASABAH']))}",
                        axis=1
                    ).tolist()

                    selected_lunas = st.selectbox(
                        "Pilih Transaksi yang akan dibayar",
                        pilihan_lunas,
                        key="pilih_trx_lunas_bebas"
                    )

                    selected_trx_id = selected_lunas.split(" | ")[0]
                    
                    baris_terpilih = belum_lunas[belum_lunas["TRANSAKSI_ID"] == selected_trx_id].iloc[0]
                    
                    sisa_saldo_awal = float(baris_terpilih.get("SISA SALDO", baris_terpilih["SALDO NASABAH"]))
                    if pd.isna(sisa_saldo_awal): 
                        sisa_saldo_awal = float(baris_terpilih["SALDO NASABAH"])

                    # Form Input Uang
                    nominal_bayar = st.number_input(
                        "Nominal Pembayaran (Rp)",
                        min_value=0.0,
                        max_value=float(sisa_saldo_awal),
                        step=1000.0,
                        value=float(sisa_saldo_awal), # Secara default otomatis terisi angka pelunasan penuh
                        key="nominal_bayar"
                    )

                    tanggal_lunas = st.date_input(
                        "Tanggal Pembayaran",
                        value=date.today(),
                        key="tanggal_bayar_bebas"
                    )

                    # --- PEMBAGIAN TOMBOL ---
                    col_btn1, col_btn2 = st.columns(2)

                    # TOMBOL 1: BAYAR SESUAI NOMINAL INPUT
                    with col_btn1:
                        btn_bayar = st.button("✅ Proses Pembayaran (Nota Terpilih)", type="primary", use_container_width=True)
                    
                    # TOMBOL 2: LUNASI SEMUA NOTA
                    with col_btn2:
                        btn_lunasi_semua = st.button("💰 Lunasi SEMUA Transaksi", type="secondary", use_container_width=True)

                    if btn_bayar:
                        if nominal_bayar <= 0:
                            st.error("Nominal pembayaran harus lebih dari 0!")
                        else:
                            df_setoran_all = load_sheet(SHEET_SETORAN)
                            
                            # Handle kompatibilitas dengan data excel lama
                            if "SUDAH DIBAYAR" not in df_setoran_all.columns:
                                df_setoran_all["SUDAH DIBAYAR"] = 0
                            if "SISA SALDO" not in df_setoran_all.columns:
                                df_setoran_all["SISA SALDO"] = df_setoran_all["SALDO NASABAH"]
                                
                            idx = df_setoran_all[df_setoran_all["TRANSAKSI_ID"].astype(str) == str(selected_trx_id)].index
                            
                            if len(idx) == 0:
                                st.error("Data transaksi tidak ditemukan.")
                            else:
                                idx = idx[0]
                                
                                # 1. Hitung matematika sisa saldonya
                                current_dibayar = float(df_setoran_all.at[idx, "SUDAH DIBAYAR"])
                                if pd.isna(current_dibayar): current_dibayar = 0
                                
                                new_dibayar = current_dibayar + nominal_bayar
                                new_sisa = float(df_setoran_all.at[idx, "SALDO NASABAH"]) - new_dibayar
                                
                                # 2. Update nilainya di Excel
                                df_setoran_all.at[idx, "SUDAH DIBAYAR"] = new_dibayar
                                df_setoran_all.at[idx, "SISA SALDO"] = new_sisa
                                
                                # 3. Cek apakah sudah lunas 100%
                                if new_sisa <= 0:
                                    df_setoran_all.at[idx, "STATUS PEMBAYARAN"] = "LUNAS"
                                    df_setoran_all.at[idx, "TANGGAL LUNAS"] = tanggal_lunas.strftime("%Y-%m-%d")
                                
                                # 4. CATAT PENGELUARAN KAS KEUANGAN OTOMATIS
                                df_keuangan = load_sheet(SHEET_KEUANGAN)
                                keuangan_id = generate_id(df_keuangan, "KEUANGAN_ID", "KEU", 6)
                                
                                # Mencatat pengeluaran kas sesuai nominal input (Satu Nota)
                                new_finance = pd.DataFrame(
                                    [[
                                        keuangan_id,
                                        tanggal_lunas.strftime("%Y-%m-%d"),
                                        f"Pembayaran Saldo Nasabah (Nota: {selected_trx_id})",
                                        0,              # DEBIT = 0
                                        nominal_bayar,  # KREDIT = Kas Berkurang (-Nominal Dibayar)
                                        selected_trx_id
                                    ]],
                                    columns=REQUIRED_COLUMNS[SHEET_KEUANGAN]
                                )
                                
                                df_keuangan = pd.concat([df_keuangan, new_finance], ignore_index=True)
                                
                                # 5. Simpan semuanya
                                if save_multiple_sheets({SHEET_SETORAN: df_setoran_all, SHEET_KEUANGAN: df_keuangan}):
                                    st.success(f"Pembayaran sebesar {format_rupiah(nominal_bayar)} berhasil dicatat.")
                                    if new_sisa <= 0:
                                        st.info("Status Nota telah menjadi LUNAS!")
                                    else:
                                        st.info(f"Sisa saldo nota ini yang belum dibayar: {format_rupiah(new_sisa)}")
                                    st.rerun()

                    # LOGIKA UNTUK TOMBOL LUNASI SEMUA
                    if btn_lunasi_semua:
                        df_setoran_all = load_sheet(SHEET_SETORAN)
                        df_keuangan = load_sheet(SHEET_KEUANGAN)
                        
                        # Pastikan kolom ada
                        if "SUDAH DIBAYAR" not in df_setoran_all.columns: df_setoran_all["SUDAH DIBAYAR"] = 0
                        if "SISA SALDO" not in df_setoran_all.columns: df_setoran_all["SISA SALDO"] = df_setoran_all["SALDO NASABAH"]
                        
                        total_dibayar_masal = 0
                        
                        # Looping semua transaksi nasabah ini yang berstatus "BELUM LUNAS"
                        for _, row in belum_lunas.iterrows():
                            trx_id = str(row['TRANSAKSI_ID'])
                            idx_list = df_setoran_all[df_setoran_all["TRANSAKSI_ID"].astype(str) == trx_id].index
                            
                            if len(idx_list) > 0:
                                idx = idx_list[0]
                                sisa_saldo = float(df_setoran_all.at[idx, "SISA SALDO"])
                                if pd.isna(sisa_saldo): sisa_saldo = float(df_setoran_all.at[idx, "SALDO NASABAH"])
                                
                                if sisa_saldo > 0:
                                    total_dibayar_masal += sisa_saldo
                                    current_dibayar = float(df_setoran_all.at[idx, "SUDAH DIBAYAR"]) if pd.notna(df_setoran_all.at[idx, "SUDAH DIBAYAR"]) else 0
                                    
                                    # Update Excel
                                    df_setoran_all.at[idx, "SUDAH DIBAYAR"] = current_dibayar + sisa_saldo
                                    df_setoran_all.at[idx, "SISA SALDO"] = 0
                                    df_setoran_all.at[idx, "STATUS PEMBAYARAN"] = "LUNAS"
                                    df_setoran_all.at[idx, "TANGGAL LUNAS"] = tanggal_lunas.strftime("%Y-%m-%d")
                                    
                                    # Catat per transaksi di keuangan
                                    keuangan_id = generate_id(df_keuangan, "KEUANGAN_ID", "KEU", 6)
                                    new_finance = pd.DataFrame([[
                                        keuangan_id,
                                        tanggal_lunas.strftime("%Y-%m-%d"),
                                        f"Pelunasan Masal Saldo Nasabah (Nota: {trx_id})",
                                        sisa_saldo,
                                        0,
                                        trx_id
                                    ]], columns=REQUIRED_COLUMNS[SHEET_KEUANGAN])
                                    df_keuangan = pd.concat([df_keuangan, new_finance], ignore_index=True)
                        
                        if total_dibayar_masal > 0:
                            if save_multiple_sheets({SHEET_SETORAN: df_setoran_all, SHEET_KEUANGAN: df_keuangan}):
                                st.success(f"Semua nota berhasil dilunasi dengan total pengeluaran {format_rupiah(total_dibayar_masal)}.")
                                st.rerun()
                        else:
                            st.warning("Tidak ada sisa saldo yang bisa dilunasi.")

                # ------------------------------------------------
                # PDF & WHATSAPP (UI BARU)
                # ------------------------------------------------

                st.subheader("📄 Cetak Mutasi & Bagikan")

                nasabah = df_nasabah[
                    df_nasabah["NIK"].astype(str) == str(selected_nik)
                ].iloc[0]

                pdf_buffer = generate_mutasi_pdf(
                    nasabah,
                    transaksi,
                    df_detail[
                        df_detail["TRANSAKSI_ID"].astype(str).isin(transaksi["TRANSAKSI_ID"].astype(str))
                    ].copy() if not df_detail.empty else pd.DataFrame()
                )

                # Bagi layout menjadi 2 tombol berdampingan
                col_pdf, col_wa = st.columns(2)

                with col_pdf:
                    st.download_button(
                        "📄 Download Mutasi PDF",
                        data=pdf_buffer,
                        file_name=f"mutasi_{selected_nik}_{date.today().strftime('%Y%m%d')}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )

                with col_wa:
                    import urllib.parse # Library bawaan python untuk mengubah teks jadi format link

                    # 1. Rapikan Format Nomor HP ke format internasional (62)
                    no_hp = str(nasabah["HP"]).strip()
                    no_hp = ''.join(filter(str.isdigit, no_hp)) # Buang spasi/tanda strip
                    
                    if no_hp.startswith("0"):
                        no_hp = "62" + no_hp[1:]
                    elif no_hp.startswith("8"):
                        no_hp = "62" + no_hp
                        
                    # 2. Hitung ulang rekap khusus untuk pesan WA
                    wa_saldo = pd.to_numeric(transaksi["SALDO NASABAH"], errors="coerce").fillna(0).sum()
                    wa_dibayar = pd.to_numeric(transaksi["SUDAH DIBAYAR"], errors="coerce").fillna(0).sum()
                    wa_sisa = pd.to_numeric(transaksi["SISA SALDO"], errors="coerce").fillna(0).sum()
                    
                    if no_hp and len(no_hp) >= 10:
                        # 3. Rangkai isi pesannya
                        pesan_wa = f"Halo Bapak/Ibu *{nasabah['NAMA']}*,\n\n"
                        pesan_wa += f"Berikut adalah informasi saldo Anda di Bank Sampah per tanggal {date.today().strftime('%d-%m-%Y')}:\n\n"
                        pesan_wa += f"💰 *Total Hak Anda:* {format_rupiah(wa_saldo)}\n"
                        pesan_wa += f"💸 *Sudah Ditarik:* {format_rupiah(wa_dibayar)}\n"
                        pesan_wa += f"💳 *SISA SALDO:* {format_rupiah(wa_sisa)}\n\n"
                        pesan_wa += "Terima kasih telah menabung di Bank Sampah! ♻️\n"
                        pesan_wa += "*(Silakan balas pesan ini jika Anda membutuhkan rincian mutasi lengkap dalam bentuk PDF)*"
                        
                        # Encode teks agar aman masuk ke URL
                        link_wa = f"https://wa.me/{no_hp}?text={urllib.parse.quote(pesan_wa)}"
                        
                        # Buat tombol Link WhatsApp (Fitur Streamlit)
                        st.link_button("💬 Kirim Info ke WhatsApp", link_wa, use_container_width=True)
                    else:
                        st.button("💬 Nomor WA Tidak Tersedia/Tidak Valid", disabled=True, use_container_width=True)

                # ------------------------------------------------
                # DETAIL ITEM
                # ------------------------------------------------

                st.subheader(
                    "📦 Detail Sampah yang Disetor"
                )

                if not df_detail.empty:

                    detail = df_detail[
                        df_detail[
                            "TRANSAKSI_ID"
                        ]
                        .astype(str)
                        .isin(
                            transaksi[
                                "TRANSAKSI_ID"
                            ]
                            .astype(str)
                        )
                    ].copy()

                    if selected_item != "SEMUA":

                        detail = detail[
                            detail[
                                "ITEM"
                            ].astype(str)
                            == selected_item
                        ]

                    if not detail.empty:

                        detail[
                            "KG"
                        ] = pd.to_numeric(
                            detail["KG"],
                            errors="coerce"
                        ).fillna(0)

                        detail[
                            "SUBTOTAL"
                        ] = pd.to_numeric(
                            detail["SUBTOTAL"],
                            errors="coerce"
                        ).fillna(0)

                        rekap_item = (
                            detail
                            .groupby(
                                "ITEM",
                                as_index=False
                            )
                            .agg(
                                TOTAL_KG=(
                                    "KG",
                                    "sum"
                                ),
                                TOTAL_NILAI=(
                                    "SUBTOTAL",
                                    "sum"
                                )
                            )
                        )

                        rekap_item[
                            "TOTAL NILAI"
                        ] = rekap_item[
                            "TOTAL_NILAI"
                        ].apply(
                            format_rupiah
                        )

                        st.dataframe(
                            rekap_item[
                                [
                                    "ITEM",
                                    "TOTAL_KG",
                                    "TOTAL NILAI"
                                ]
                            ],
                            use_container_width=True,
                            hide_index=True
                        )

                        st.subheader(
                            "Detail Per Transaksi"
                        )

                        detail_display = detail.merge(
                            transaksi[
                                [
                                    "TRANSAKSI_ID",
                                    "TANGGAL TRANSAKSI"
                                ]
                            ],
                            on="TRANSAKSI_ID",
                            how="left"
                        )

                        detail_display[
                            "TANGGAL TRANSAKSI"
                        ] = detail_display[
                            "TANGGAL TRANSAKSI"
                        ].dt.strftime(
                            "%d-%m-%Y"
                        )

                        detail_display[
                            "SUBTOTAL"
                        ] = detail_display[
                            "SUBTOTAL"
                        ].apply(
                            format_rupiah
                        )

                        detail_display[
                            "HARGA"
                        ] = detail_display[
                            "HARGA"
                        ].apply(
                            format_rupiah
                        )

                        st.dataframe(
                            detail_display[
                                [
                                    "TRANSAKSI_ID",
                                    "TANGGAL TRANSAKSI",
                                    "ITEM",
                                    "KG",
                                    "HARGA",
                                    "SUBTOTAL"
                                ]
                            ],
                            use_container_width=True,
                            hide_index=True
                        )


# ============================================================
# MASTER ITEM
# ============================================================

def page_master_item():

    st.title(
        "🗂️ Master Item Sampah"
    )

    df = load_sheet(
        SHEET_MASTER_ITEM
    )

    tab1, tab2, tab3 = st.tabs(
        [
            "📋 Data Item",
            "➕ Tambah Item",
            "✏️ Edit / Hapus"
        ]
    )

    # ========================================================
    # DATA ITEM
    # ========================================================

    with tab1:

        if df.empty:

            st.info(
                "Belum ada master item."
            )

        else:

            st.dataframe(
                df[
                    [
                        "ITEM_ID",
                        "JENIS",
                        "ITEM",
                        "SATUAN"
                    ]
                ],
                use_container_width=True,
                hide_index=True
            )

            st.caption(
                f"Total item: {len(df)}"
            )

    # ========================================================
    # TAMBAH ITEM
    # ========================================================

    with tab2:

        st.subheader(
            "➕ Tambah Item Sampah"
        )

        jenis = st.text_input(
            "Jenis Sampah",
            placeholder="Contoh: PLASTIK"
        )

        item = st.text_input(
            "Nama Item",
            placeholder="Contoh: Botol Plastik"
        )

        satuan = st.selectbox(
            "Satuan",
            [
                "KG",
                "PCS",
                "UNIT"
            ]
        )

        if st.button(
            "💾 Simpan Item",
            type="primary",
            use_container_width=True
        ):

            jenis = jenis.strip().upper()
            item = item.strip()

            if not jenis:

                st.error(
                    "Jenis sampah wajib diisi."
                )

            elif not item:

                st.error(
                    "Nama item wajib diisi."
                )

            else:

                duplicate = False

                if not df.empty:

                    duplicate = (
                        df[
                            "ITEM"
                        ]
                        .astype(str)
                        .str.strip()
                        .str.lower()
                        .eq(
                            item.lower()
                        )
                        .any()
                    )

                if duplicate:

                    st.error(
                        "Nama item sudah terdaftar."
                    )

                else:

                    item_id = generate_id(
                        df,
                        "ITEM_ID",
                        "ITM",
                        6
                    )

                    new_row = pd.DataFrame(
                        [[
                            item_id,
                            jenis,
                            item,
                            satuan
                        ]],
                        columns=REQUIRED_COLUMNS[
                            SHEET_MASTER_ITEM
                        ]
                    )

                    df = pd.concat(
                        [
                            df,
                            new_row
                        ],
                        ignore_index=True
                    )

                    if save_sheet(
                        SHEET_MASTER_ITEM,
                        df
                    ):

                        st.success(
                            f"Item {item_id} berhasil ditambahkan."
                        )

                        st.rerun()

    # ========================================================
    # EDIT / HAPUS
    # ========================================================

    with tab3:

        if df.empty:

            st.info(
                "Belum ada item."
            )

        else:

            pilihan = df.apply(
                lambda row:
                f"{row['ITEM_ID']} | "
                f"{row['ITEM']} | "
                f"{row['JENIS']}",
                axis=1
            ).tolist()

            selected = st.selectbox(
                "Pilih Item",
                pilihan,
                key="pilih_master_item"
            )

            selected_id = (
                selected
                .split(" | ")[0]
            )

            index_list = df[
                df[
                    "ITEM_ID"
                ].astype(str)
                == selected_id
            ].index.tolist()

            if not index_list:

                st.error(
                    "Item tidak ditemukan."
                )

                return

            index = index_list[0]

            row = df.loc[index]

            jenis_baru = st.text_input(
                "Jenis Sampah",
                value=str(
                    row["JENIS"]
                ),
                key="edit_jenis_item"
            )

            item_baru = st.text_input(
                "Nama Item",
                value=str(
                    row["ITEM"]
                ),
                key="edit_nama_item"
            )

            satuan_options = [
                "KG",
                "PCS",
                "UNIT"
            ]

            satuan_lama = str(
                row["SATUAN"]
            )

            if satuan_lama not in satuan_options:

                satuan_options.append(
                    satuan_lama
                )

            satuan_baru = st.selectbox(
                "Satuan",
                satuan_options,
                index=satuan_options.index(
                    satuan_lama
                ),
                key="edit_satuan_item"
            )

            col1, col2 = st.columns(2)

            update = col1.button(
                "💾 Update Item",
                type="primary",
                use_container_width=True
            )

            delete = col2.button(
                "🗑️ Hapus Item",
                use_container_width=True
            )

            if update:

                jenis_baru = (
                    jenis_baru
                    .strip()
                    .upper()
                )

                item_baru = (
                    item_baru
                    .strip()
                )

                if not jenis_baru:

                    st.error(
                        "Jenis sampah wajib diisi."
                    )

                elif not item_baru:

                    st.error(
                        "Nama item wajib diisi."
                    )

                else:

                    duplicate = (
                        df[
                            "ITEM"
                        ]
                        .astype(str)
                        .str.strip()
                        .str.lower()
                        .eq(
                            item_baru.lower()
                        )
                        &
                        (
                            df[
                                "ITEM_ID"
                            ].astype(str)
                            != selected_id
                        )
                    ).any()

                    if duplicate:

                        st.error(
                            "Nama item sudah digunakan "
                            "oleh item lain."
                        )

                    else:

                        old_item_name = str(
                            df.at[
                                index,
                                "ITEM"
                            ]
                        )

                        df.at[
                            index,
                            "JENIS"
                        ] = jenis_baru

                        df.at[
                            index,
                            "ITEM"
                        ] = item_baru

                        df.at[
                            index,
                            "SATUAN"
                        ] = satuan_baru

                        if save_sheet(
                            SHEET_MASTER_ITEM,
                            df
                        ):

                            # ------------------------------------------------
                            # UPDATE NAMA ITEM DI MASTER HARGA
                            # ------------------------------------------------

                            df_harga = load_sheet(
                                SHEET_HARGA
                            )

                            if not df_harga.empty:

                                mask = (
                                    df_harga[
                                        "ITEM_ID"
                                    ].astype(str)
                                    == selected_id
                                )

                                df_harga.loc[
                                    mask,
                                    "ITEM"
                                ] = item_baru

                                save_sheet(
                                    SHEET_HARGA,
                                    df_harga
                                )

                            # ------------------------------------------------
                            # UPDATE NAMA ITEM DI DETAIL SETORAN
                            # ------------------------------------------------

                            df_detail = load_sheet(
                                SHEET_DETAIL_SETORAN
                            )

                            if not df_detail.empty:

                                mask = (
                                    df_detail[
                                        "ITEM_ID"
                                    ].astype(str)
                                    == selected_id
                                )

                                df_detail.loc[
                                    mask,
                                    "ITEM"
                                ] = item_baru

                                save_sheet(
                                    SHEET_DETAIL_SETORAN,
                                    df_detail
                                )

                            # ------------------------------------------------
                            # UPDATE NAMA ITEM DI DETAIL PENJUALAN
                            # ------------------------------------------------

                            df_detail_jual = load_sheet(
                                SHEET_DETAIL_PENJUALAN
                            )

                            if not df_detail_jual.empty:

                                mask = (
                                    df_detail_jual[
                                        "ITEM_ID"
                                    ].astype(str)
                                    == selected_id
                                )

                                df_detail_jual.loc[
                                    mask,
                                    "ITEM"
                                ] = item_baru

                                save_sheet(
                                    SHEET_DETAIL_PENJUALAN,
                                    df_detail_jual
                                )

                            # ------------------------------------------------
                            # UPDATE NAMA ITEM DI STOK
                            # ------------------------------------------------

                            df_stok = load_sheet(
                                SHEET_STOK
                            )

                            if not df_stok.empty:

                                mask = (
                                    df_stok[
                                        "ITEM_ID"
                                    ].astype(str)
                                    == selected_id
                                )

                                df_stok.loc[
                                    mask,
                                    "ITEM"
                                ] = item_baru

                                save_sheet(
                                    SHEET_STOK,
                                    df_stok
                                )

                            st.success(
                                "Master item berhasil diperbarui."
                            )

                            st.rerun()

            if delete:

                # ---------------------------------------------
                # CEK APAKAH ITEM SUDAH DIGUNAKAN
                # ---------------------------------------------

                df_detail = load_sheet(
                    SHEET_DETAIL_SETORAN
                )

                df_detail_jual = load_sheet(
                    SHEET_DETAIL_PENJUALAN
                )

                df_stok = load_sheet(
                    SHEET_STOK
                )

                df_harga = load_sheet(
                    SHEET_HARGA
                )

                sedang_digunakan = False

                if not df_detail.empty:

                    sedang_digunakan = (
                        df_detail[
                            "ITEM_ID"
                        ].astype(str)
                        == selected_id
                    ).any()

                if not df_detail_jual.empty:

                    sedang_digunakan = (
                        sedang_digunakan
                        or
                        (
                            df_detail_jual[
                                "ITEM_ID"
                            ].astype(str)
                            == selected_id
                        ).any()
                    )

                if not df_stok.empty:

                    sedang_digunakan = (
                        sedang_digunakan
                        or
                        (
                            df_stok[
                                "ITEM_ID"
                            ].astype(str)
                            == selected_id
                        ).any()
                    )

                if not df_harga.empty:

                    sedang_digunakan = (
                        sedang_digunakan
                        or
                        (
                            df_harga[
                                "ITEM_ID"
                            ].astype(str)
                            == selected_id
                        ).any()
                    )

                if sedang_digunakan:

                    st.error(
                        "Item tidak dapat dihapus karena "
                        "sudah digunakan dalam transaksi, "
                        "stok, atau daftar harga."
                    )

                else:

                    df = df.drop(
                        index
                    ).reset_index(
                        drop=True
                    )

                    if save_sheet(
                        SHEET_MASTER_ITEM,
                        df
                    ):

                        st.success(
                            "Item berhasil dihapus."
                        )

                        st.rerun()


# ============================================================
# HARGA
# ============================================================

def page_harga():

    st.title(
        "💰 Daftar Harga Sampah"
    )

    df_harga = load_sheet(
        SHEET_HARGA
    )

    df_item = load_sheet(
        SHEET_MASTER_ITEM
    )

    tab1, tab2, tab3 = st.tabs(
        [
            "📋 Daftar Harga",
            "➕ Tambah Harga",
            "✏️ Edit Harga"
        ]
    )

    # ========================================================
    # TAB 1
    # ========================================================

    with tab1:

        if df_harga.empty:

            st.info(
                "Belum ada daftar harga."
            )

        else:

            display = df_harga.copy()

            display[
                "TANGGAL MULAI BERLAKU"
            ] = pd.to_datetime(
                display[
                    "TANGGAL MULAI BERLAKU"
                ],
                errors="coerce"
            ).dt.strftime(
                "%d-%m-%Y"
            )

            display[
                "HARGA SETOR"
            ] = display[
                "HARGA SETOR"
            ].apply(
                format_rupiah
            )

            st.dataframe(
                display.sort_values(
                    "TANGGAL MULAI BERLAKU",
                    ascending=False
                ),
                use_container_width=True,
                hide_index=True
            )

    # ========================================================
    # TAMBAH
    # ========================================================

    with tab2:

        if df_item.empty:

            st.error(
                "Master item belum tersedia."
            )

        else:

            tanggal_mulai = st.date_input(
                "Tanggal Mulai Berlaku",
                value=date.today(),
                key="tanggal_mulai_harga"
            )

            jenis = st.selectbox(
                "Jenis Sampah",
                sorted(
                    df_item[
                        "JENIS"
                    ]
                    .dropna()
                    .unique()
                    .tolist()
                ),
                key="jenis_tambah_harga"
            )

            items_jenis = df_item[
                df_item[
                    "JENIS"
                ] == jenis
            ]

            item_id = st.selectbox(
                "Item",
                items_jenis[
                    "ITEM_ID"
                ].tolist(),
                format_func=lambda x:
                items_jenis.loc[
                    items_jenis[
                        "ITEM_ID"
                    ] == x,
                    "ITEM"
                ].iloc[0],
                key="item_tambah_harga"
            )

            item_name = items_jenis.loc[
                items_jenis[
                    "ITEM_ID"
                ] == item_id,
                "ITEM"
            ].iloc[0]

            harga = st.number_input(
                f"Harga Setor {item_name} / KG",
                min_value=0,
                step=100,
                value=0,
                key="harga_tambah"
            )

            st.info(
                "Harga berlaku mulai tanggal yang dipilih."
            )

            if st.button(
                "💾 Simpan Harga",
                type="primary",
                use_container_width=True
            ):

                if harga <= 0:

                    st.error(
                        "Harga harus lebih dari 0."
                    )

                    st.stop()

                duplicate = False

                if not df_harga.empty:

                    tanggal_existing = pd.to_datetime(
                        df_harga[
                            "TANGGAL MULAI BERLAKU"
                        ],
                        errors="coerce"
                    )

                    duplicate = (
                        (
                            df_harga[
                                "ITEM_ID"
                            ].astype(str)
                            == str(item_id)
                        )
                        &
                        (
                            tanggal_existing.dt.date
                            == tanggal_mulai
                        )
                    ).any()

                if duplicate:

                    st.error(
                        "Sudah ada harga untuk item "
                        "tersebut pada tanggal tersebut."
                    )

                else:

                    harga_id = generate_id(
                        df_harga,
                        "HARGA_ID",
                        "H",
                        6
                    )

                    new_row = pd.DataFrame(
                        [[
                            harga_id,
                            tanggal_mulai.strftime(
                                "%Y-%m-%d"
                            ),
                            item_id,
                            item_name,
                            harga
                        ]],
                        columns=REQUIRED_COLUMNS[
                            SHEET_HARGA
                        ]
                    )

                    df_harga = pd.concat(
                        [
                            df_harga,
                            new_row
                        ],
                        ignore_index=True
                    )

                    if save_sheet(
                        SHEET_HARGA,
                        df_harga
                    ):

                        st.success(
                            "Harga berhasil ditambahkan."
                        )

                        st.rerun()

    # ========================================================
    # EDIT
    # ========================================================

    with tab3:

        if df_harga.empty:

            st.info(
                "Belum ada harga untuk diedit."
            )

        else:

            pilihan = df_harga.apply(
                lambda row:
                f"{row['HARGA_ID']} | "
                f"{row['ITEM']} | "
                f"{row['TANGGAL MULAI BERLAKU']}",
                axis=1
            ).tolist()

            selected = st.selectbox(
                "Pilih Data Harga",
                pilihan,
                key="pilih_harga_edit"
            )

            selected_id = (
                selected
                .split(" | ")[0]
            )

            index = df_harga[
                df_harga[
                    "HARGA_ID"
                ].astype(str)
                == selected_id
            ].index[0]

            row = df_harga.loc[
                index
            ]

            tanggal_existing = pd.to_datetime(
                row[
                    "TANGGAL MULAI BERLAKU"
                ],
                errors="coerce"
            )

            if pd.isna(
                tanggal_existing
            ):

                tanggal_existing = date.today()

            else:

                tanggal_existing = (
                    tanggal_existing.date()
                )

            tanggal_baru = st.date_input(
                "Tanggal Mulai Berlaku",
                value=tanggal_existing,
                key="tanggal_edit_harga"
            )

            st.text_input(
                "Item",
                value=str(
                    row["ITEM"]
                ),
                disabled=True
            )

            item_id = row[
                "ITEM_ID"
            ]

            harga_baru = st.number_input(
                "Harga Setor / KG",
                min_value=0,
                step=100,
                value=int(
                    float(
                        row[
                            "HARGA SETOR"
                        ]
                    )
                ),
                key="harga_edit"
            )

            col1, col2 = st.columns(2)

            update = col1.button(
                "💾 Update Harga",
                type="primary",
                use_container_width=True
            )

            delete = col2.button(
                "🗑️ Hapus Harga",
                use_container_width=True
            )

            if update:

                if harga_baru <= 0:

                    st.error(
                        "Harga harus lebih dari 0."
                    )

                    st.stop()

                tanggal_existing_all = pd.to_datetime(
                    df_harga[
                        "TANGGAL MULAI BERLAKU"
                    ],
                    errors="coerce"
                )

                duplicate = (
                    (
                        df_harga[
                            "ITEM_ID"
                        ].astype(str)
                        == str(item_id)
                    )
                    &
                    (
                        tanggal_existing_all.dt.date
                        == tanggal_baru
                    )
                    &
                    (
                        df_harga[
                            "HARGA_ID"
                        ].astype(str)
                        != str(selected_id)
                    )
                ).any()

                if duplicate:

                    st.error(
                        "Sudah ada harga lain untuk "
                        "item tersebut pada tanggal tersebut."
                    )

                else:

                    df_harga.at[
                        index,
                        "TANGGAL MULAI BERLAKU"
                    ] = tanggal_baru.strftime(
                        "%Y-%m-%d"
                    )

                    df_harga.at[
                        index,
                        "HARGA SETOR"
                    ] = harga_baru

                    if save_sheet(
                        SHEET_HARGA,
                        df_harga
                    ):

                        st.success(
                            "Harga berhasil diperbarui."
                        )

                        st.rerun()

            if delete:

                df_harga = df_harga.drop(
                    index
                ).reset_index(
                    drop=True
                )

                if save_sheet(
                    SHEET_HARGA,
                    df_harga
                ):

                    st.success(
                        "Harga berhasil dihapus."
                    )

                    st.rerun()


# ============================================================
# TRANSAKSI SETORAN
# ============================================================

def page_transaksi_setoran():

    st.title(
        "📥 Transaksi Setoran"
    )

    df_nasabah = load_sheet(
        SHEET_NASABAH
    )

    df_item = load_sheet(
        SHEET_MASTER_ITEM
    )

    df_setoran = load_sheet(
        SHEET_SETORAN
    )

    df_detail = load_sheet(
        SHEET_DETAIL_SETORAN
    )

    df_stok = load_sheet(
        SHEET_STOK
    )

    df_keuangan = load_sheet(
        SHEET_KEUANGAN
    )

    if df_nasabah.empty:

        st.warning(
            "Belum ada nasabah. "
            "Tambahkan nasabah terlebih dahulu."
        )

        return

    if df_item.empty:

        st.warning(
            "Master item belum tersedia."
        )

        return

    # --------------------------------------------------------
    # HEADER
    # --------------------------------------------------------

    st.subheader(
        "Data Transaksi"
    )

    col1, col2 = st.columns(2)

    tanggal = col1.date_input(
        "Tanggal Transaksi",
        value=date.today()
    )

    nasabah_options = (
        df_nasabah[
            "NIK"
        ].astype(str)
        .tolist()
    )

    selected_nik = col2.selectbox(
        "Nasabah",
        nasabah_options,
        format_func=lambda nik:
        f"{nik} - "
        +
        str(
            df_nasabah.loc[
                df_nasabah[
                    "NIK"
                ].astype(str)
                == str(nik),
                "NAMA"
            ].iloc[0]
        )
    )

    nasabah = df_nasabah[
        df_nasabah[
            "NIK"
        ].astype(str)
        == str(selected_nik)
    ].iloc[0]

    col1, col2, col3 = st.columns(3)

    col1.text_input(
        "NIK",
        value=str(
            nasabah["NIK"]
        ),
        disabled=True
    )

    col2.text_input(
        "Nama",
        value=str(
            nasabah["NAMA"]
        ),
        disabled=True
    )

    col3.text_input(
        "Periode Harga",
        value=tanggal.strftime(
            "%Y-%m"
        ),
        disabled=True
    )

    st.divider()

    st.subheader(
        "Detail Sampah"
    )

    detail_rows = []

    for jenis in (
        df_item[
            "JENIS"
        ]
        .dropna()
        .unique()
    ):

        st.markdown(
            f"### {jenis}"
        )

        items_jenis = df_item[
            df_item[
                "JENIS"
            ] == jenis
        ]

        for _, item in items_jenis.iterrows():

            item_id = item[
                "ITEM_ID"
            ]

            item_name = item[
                "ITEM"
            ]

            harga = get_price(
                item_id,
                tanggal
            )

            col1, col2, col3 = st.columns(
                [5, 2, 3]
            )

            col1.write(
                f"**{item_name}**"
            )

            if harga is None:

                col3.warning(
                    "Harga belum ada"
                )

            else:

                col3.write(
                    format_rupiah(
                        harga
                    )
                    + " / KG"
                )

            kg = col2.number_input(
                "KG",
                min_value=0.0,
                step=0.1,
                value=0.0,
                key=f"kg_{item_id}"
            )

            if kg > 0 and harga is not None:

                subtotal = (
                    kg * harga
                )

                detail_rows.append({
                    "ITEM_ID": item_id,
                    "ITEM": item_name,
                    "KG": kg,
                    "HARGA": harga,
                    "SUBTOTAL": subtotal
                })

    st.divider()

    if detail_rows:

        detail_input = pd.DataFrame(
            detail_rows
        )

        total_kg = detail_input[
            "KG"
        ].sum()

        total_harga = detail_input[
            "SUBTOTAL"
        ].sum()

        # ====================================================
        # POTONGAN KAS 10%
        # ====================================================

        potongan_kas = (
            total_harga
            * 0.10
        )

        saldo_nasabah = (
            total_harga
            - potongan_kas
        )

        col1, col2, col3 = st.columns(3)

        col1.metric(
            "Total Berat",
            f"{total_kg:,.2f} KG"
        )

        col2.metric(
            "Total Bruto",
            format_rupiah(
                total_harga
            )
        )

        col3.metric(
            "Saldo Nasabah",
            format_rupiah(
                saldo_nasabah
            )
        )

        st.info(
            f"Potongan kas 10%: "
            f"{format_rupiah(potongan_kas)}"
        )

        st.dataframe(
            detail_input,
            use_container_width=True,
            hide_index=True
        )

        st.warning(
            "Saldo nasabah belum dianggap dibayar. "
            "Status awal transaksi adalah BELUM LUNAS."
        )

        if st.button(
            "💾 SIMPAN TRANSAKSI",
            type="primary",
            use_container_width=True
        ):

            transaksi_id = generate_id(
                df_setoran,
                "TRANSAKSI_ID",
                "TRX",
                6
            )

            no = get_next_no(
                df_setoran
            )

            # =================================================
            # HEADER
            # =================================================

            new_header = pd.DataFrame(
                [[
                    transaksi_id,
                    no,
                    tanggal.strftime("%Y-%m-%d"),
                    str(nasabah["NIK"]),
                    str(nasabah["NAMA"]),
                    total_kg,
                    total_harga,
                    potongan_kas,
                    saldo_nasabah,
                    0,               # SUDAH DIBAYAR (Awalnya 0)
                    saldo_nasabah,   # SISA SALDO (Awalnya sama dengan total)
                    "BELUM LUNAS",
                    ""
                ]],
                columns=REQUIRED_COLUMNS[SHEET_SETORAN]
            )

            df_setoran = pd.concat(
                [
                    df_setoran,
                    new_header
                ],
                ignore_index=True
            )

            # =================================================
            # DETAIL
            # =================================================

            for _, row in detail_input.iterrows():

                detail_id = generate_id(
                    df_detail,
                    "DETAIL_ID",
                    "DT",
                    6
                )

                # --- KODE YANG BARU ---
                new_detail = pd.DataFrame(
                    [[
                        detail_id,
                        transaksi_id,
                        row["ITEM_ID"],
                        row["ITEM"],
                        row["KG"],
                        row["HARGA"],
                        row["SUBTOTAL"],
                        "BELUM LUNAS",  # <--- TAMBAHAN: Status awal per item
                        ""              # <--- TAMBAHAN: Tanggal Lunas kosong
                    ]],
                    columns=REQUIRED_COLUMNS[
                        SHEET_DETAIL_SETORAN
                    ]
                )

                df_detail = pd.concat(
                    [
                        df_detail,
                        new_detail
                    ],
                    ignore_index=True
                )

            # =================================================
            # MUTASI STOK MASUK
            # =================================================

            for _, row in detail_input.iterrows():

                mutasi_id = generate_id(
                    df_stok,
                    "MUTASI_ID",
                    "M",
                    6
                )

                new_stock = pd.DataFrame(
                    [[
                        mutasi_id,
                        tanggal.strftime(
                            "%Y-%m-%d"
                        ),
                        row["ITEM_ID"],
                        row["ITEM"],
                        "MASUK",
                        row["KG"],
                        transaksi_id
                    ]],
                    columns=REQUIRED_COLUMNS[
                        SHEET_STOK
                    ]
                )

                df_stok = pd.concat(
                    [
                        df_stok,
                        new_stock
                    ],
                    ignore_index=True
                )

            # =================================================
            # KEUANGAN
            # =================================================
            #
            # Potongan 10% langsung masuk kas.
            #
            # Saldo nasabah 90% merupakan kewajiban
            # yang masih harus dibayarkan.
            #
            # =================================================

            keuangan_id = generate_id(
                df_keuangan,
                "KEUANGAN_ID",
                "KEU",
                6
            )

            new_keuangan = pd.DataFrame(
                [[
                    keuangan_id,
                    tanggal.strftime("%Y-%m-%d"),
                    f"Kas 10% dari setoran {nasabah['NAMA']}",
                    potongan_kas,  # DEBIT = Pemasukan Kas (+10%)
                    0,             # KREDIT = 0
                    transaksi_id
                ]],
                columns=REQUIRED_COLUMNS[SHEET_KEUANGAN]
            )

            df_keuangan = pd.concat(
                [
                    df_keuangan,
                    new_keuangan
                ],
                ignore_index=True
            )

            # =================================================
            # SIMPAN
            # =================================================

            success = save_multiple_sheets({

                SHEET_SETORAN:
                    df_setoran,

                SHEET_DETAIL_SETORAN:
                    df_detail,

                SHEET_STOK:
                    df_stok,

                SHEET_KEUANGAN:
                    df_keuangan

            })

            if success:

                st.success(
                    f"Transaksi {transaksi_id} berhasil disimpan."
                )

                st.info(
                    f"Saldo nasabah: "
                    f"{format_rupiah(saldo_nasabah)} "
                    f"(BELUM LUNAS)"
                )

                # --- KODE BARU: PINDAH HALAMAN OTOMATIS ---
                
                # 1. Gunakan variabel 'titipan' agar tidak bentrok dengan widget
                st.session_state.pindah_halaman = "👥 Nasabah"
                
                # 2. Simpan nama nasabah ini agar otomatis terpilih di halaman detail
                st.session_state.detail_nasabah = f"{nasabah['NIK']} - {nasabah['NAMA']}"
                
                # 3. Muat ulang halaman
                st.rerun()
                # ------------------------------------------

    else:

        st.info(
            "Masukkan KG pada minimal satu jenis sampah."
        )
        
        

# ============================================================
# TRANSAKSI PENJUALAN
# ============================================================

def page_transaksi_penjualan():

    st.title(
        "📤 Transaksi Penjualan"
    )

    df_item = load_sheet(
        SHEET_MASTER_ITEM
    )

    df_penjualan = load_sheet(
        SHEET_PENJUALAN
    )

    df_detail = load_sheet(
        SHEET_DETAIL_PENJUALAN
    )

    df_stok = load_sheet(
        SHEET_STOK
    )

    df_keuangan = load_sheet(
        SHEET_KEUANGAN
    )

    if df_item.empty:

        st.warning(
            "Master item belum tersedia."
        )

        return

    tanggal = st.date_input(
        "Tanggal Penjualan",
        value=date.today()
    )

    pengepul = st.text_input(
        "Nama Pengepul"
    )

    st.divider()

    st.subheader(
        "Barang yang Dijual"
    )

    current_stock = get_stock()

    detail_rows = []

    for jenis in (
        df_item[
            "JENIS"
        ]
        .dropna()
        .unique()
    ):

        st.markdown(
            f"### {jenis}"
        )

        items_jenis = df_item[
            df_item[
                "JENIS"
            ] == jenis
        ]

        for _, item in items_jenis.iterrows():

            item_id = item[
                "ITEM_ID"
            ]

            item_name = item[
                "ITEM"
            ]

            stock_row = current_stock[
                current_stock[
                    "ITEM_ID"
                ]
                == item_id
            ]

            if stock_row.empty:

                stok = 0

            else:

                stok = float(
                    stock_row.iloc[0][
                        "STOK"
                    ]
                )

            col1, col2, col3 = st.columns(
                [5, 2, 2]
            )

            col1.write(
                f"**{item_name}**"
            )

            col2.write(
                f"Stok: {stok:.2f} KG"
            )

            kg = col3.number_input(
                "KG",
                min_value=0.0,
                max_value=max(
                    stok,
                    0
                ),
                step=0.1,
                value=0.0,
                key=f"jual_{item_id}"
            )

            harga = st.number_input(
                "Harga Jual / KG",
                min_value=0,
                step=100,
                value=0,
                key=f"harga_jual_{item_id}"
            )

            if kg > 0:

                subtotal = (
                    kg * harga
                )

                detail_rows.append({
                    "ITEM_ID": item_id,
                    "ITEM": item_name,
                    "KG": kg,
                    "HARGA": harga,
                    "SUBTOTAL": subtotal
                })

    st.divider()

    if detail_rows:

        detail_input = pd.DataFrame(
            detail_rows
        )

        total_kg = detail_input[
            "KG"
        ].sum()

        total_harga = detail_input[
            "SUBTOTAL"
        ].sum()

        col1, col2 = st.columns(2)

        col1.metric(
            "Total Berat",
            f"{total_kg:,.2f} KG"
        )

        col2.metric(
            "Total Penjualan",
            format_rupiah(
                total_harga
            )
        )

        st.dataframe(
            detail_input,
            use_container_width=True,
            hide_index=True
        )

        if st.button(
            "💾 SIMPAN PENJUALAN",
            type="primary",
            use_container_width=True
        ):

            if not pengepul.strip():

                st.error(
                    "Nama pengepul wajib diisi."
                )

                return

            penjualan_id = generate_id(
                df_penjualan,
                "PENJUALAN_ID",
                "JUAL",
                6
            )

            new_header = pd.DataFrame(
                [[
                    penjualan_id,
                    tanggal.strftime(
                        "%Y-%m-%d"
                    ),
                    pengepul.strip(),
                    total_kg,
                    total_harga
                ]],
                columns=REQUIRED_COLUMNS[
                    SHEET_PENJUALAN
                ]
            )

            df_penjualan = pd.concat(
                [
                    df_penjualan,
                    new_header
                ],
                ignore_index=True
            )

            for _, row in detail_input.iterrows():

                detail_id = generate_id(
                    df_detail,
                    "DETAIL_ID",
                    "DJ",
                    6
                )

                new_detail = pd.DataFrame(
                    [[
                        detail_id,
                        penjualan_id,
                        row["ITEM_ID"],
                        row["ITEM"],
                        row["KG"],
                        row["HARGA"],
                        row["SUBTOTAL"]
                    ]],
                    columns=REQUIRED_COLUMNS[
                        SHEET_DETAIL_PENJUALAN
                    ]
                )

                df_detail = pd.concat(
                    [
                        df_detail,
                        new_detail
                    ],
                    ignore_index=True
                )

                mutasi_id = generate_id(
                    df_stok,
                    "MUTASI_ID",
                    "M",
                    6
                )

                new_stock = pd.DataFrame(
                    [[
                        mutasi_id,
                        tanggal.strftime(
                            "%Y-%m-%d"
                        ),
                        row["ITEM_ID"],
                        row["ITEM"],
                        "KELUAR",
                        row["KG"],
                        penjualan_id
                    ]],
                    columns=REQUIRED_COLUMNS[
                        SHEET_STOK
                    ]
                )

                df_stok = pd.concat(
                    [
                        df_stok,
                        new_stock
                    ],
                    ignore_index=True
                )

            keuangan_id = generate_id(
                df_keuangan,
                "KEUANGAN_ID",
                "KEU",
                6
            )

            new_finance = pd.DataFrame(
                [[
                    keuangan_id,
                    tanggal.strftime("%Y-%m-%d"),
                    f"Penjualan Sampah ke {pengepul.strip()}",
                    total_harga,  # DEBIT = Pemasukan Kas
                    0,            # KREDIT = 0
                    penjualan_id
                ]],
                columns=REQUIRED_COLUMNS[SHEET_KEUANGAN]
            )

            df_keuangan = pd.concat(
                [
                    df_keuangan,
                    new_finance
                ],
                ignore_index=True
            )
# =================================================
            # SIMPAN TRANSAKSI PENJUALAN
            # =================================================

            success = save_multiple_sheets({

                SHEET_PENJUALAN:
                    df_penjualan,

                SHEET_DETAIL_PENJUALAN:
                    df_detail,

                SHEET_STOK:
                    df_stok,

                SHEET_KEUANGAN:
                    df_keuangan

            })

            if success:

                st.success(
                    f"Transaksi Penjualan {penjualan_id} berhasil disimpan."
                )

                st.info(
                    f"Pemasukan kas: "
                    f"{format_rupiah(total_harga)} "
                    f"dari {pengepul}"
                )

                # --- KODE BARU: PINDAH HALAMAN OTOMATIS ---
                
                # 1. Gunakan variabel 'titipan' untuk pindah ke Keuangan
                st.session_state.pindah_halaman = "💵 Keuangan"
                
                # 2. Muat ulang halaman agar menu sidebar berubah
                st.rerun()
                # ------------------------------------------

    else:

        st.info(
            "Masukkan KG pada minimal satu jenis sampah yang akan dijual."
        )
        

# ============================================================
# STOK
# ============================================================

def page_stok():

    st.title(
        "📦 Stok Bank Sampah"
    )

    df_stok = get_stock()

    if df_stok.empty:

        st.info(
            "Belum ada data stok."
        )

        return

    col1, col2, col3 = st.columns(3)

    col1.metric(
        "Total Stok Masuk",
        f"{df_stok['MASUK'].sum():,.2f} KG"
    )

    col2.metric(
        "Total Stok Keluar",
        f"{df_stok['KELUAR'].sum():,.2f} KG"
    )

    col3.metric(
        "Stok Saat Ini",
        f"{df_stok['STOK'].sum():,.2f} KG"
    )

    st.divider()

    jenis_filter = st.selectbox(
        "Filter Jenis",
        ["SEMUA"]
        +
        sorted(
            df_stok[
                "JENIS"
            ]
            .dropna()
            .unique()
            .tolist()
        )
    )

    display = df_stok.copy()

    if jenis_filter != "SEMUA":

        display = display[
            display[
                "JENIS"
            ]
            == jenis_filter
        ]

    st.dataframe(
        display,
        use_container_width=True,
        hide_index=True
    )

    st.subheader(
        "Riwayat Mutasi Stok"
    )

    df_mutasi = load_sheet(
        SHEET_STOK
    )

    if not df_mutasi.empty:

        st.dataframe(
            df_mutasi.sort_values(
                "TANGGAL",
                ascending=False
            ),
            use_container_width=True,
            hide_index=True
        )


# ============================================================
# KEUANGAN
# ============================================================

def page_keuangan():

    st.title(
        "💵 Rekapitulasi Keuangan"
    )

    df = load_sheet(
        SHEET_KEUANGAN
    )

    if df.empty:

        st.info(
            "Belum ada transaksi keuangan."
        )

        return

    df["DEBIT"] = pd.to_numeric(
        df["DEBIT"],
        errors="coerce"
    ).fillna(0)

    df["KREDIT"] = pd.to_numeric(
        df["KREDIT"],
        errors="coerce"
    ).fillna(0)

    total_debit = df[
        "DEBIT"
    ].sum()

    total_kredit = df[
        "KREDIT"
    ].sum()

    saldo = (
        total_debit
        - total_kredit
    )

    col1, col2, col3 = st.columns(3)

    col1.metric(
        "Debit / Pemasukan dari Pengepul",
        format_rupiah(
            total_debit
        )
    )

    col2.metric(
        "Kredit / Pemasukan dari Setoran",
        format_rupiah(
            total_kredit
        )
    )

    col3.metric(
        "Saldo Kas",
        format_rupiah(
            saldo
        )
    )

    st.divider()

    display = df.copy()

    display[
        "DEBIT"
    ] = display[
        "DEBIT"
    ].apply(
        format_rupiah
    )

    display[
        "KREDIT"
    ] = display[
        "KREDIT"
    ].apply(
        format_rupiah
    )

    st.dataframe(
        display,
        use_container_width=True,
        hide_index=True
    )

def render_floating_credits():
    # 1. CSS posisi melayang di pojok kanan bawah
    st.markdown(
        """
        <style>
        div[data-testid="stPopover"] {
            position: fixed;
            bottom: 20px;
            right: 20px;
            z-index: 999999;
        }
        div[data-testid="stPopover"] > button {
            border-radius: 50px !important;
            background-color: #2e7d32 !important;
            color: white !important;
            border: none !important;
            box-shadow: 0 4px 12px rgba(0,0,0,0.3) !important;
            padding: 10px 18px !important;
            font-weight: bold !important;
            transition: transform 0.2s ease-in-out;
        }
        div[data-testid="stPopover"] > button:hover {
            transform: scale(1.05);
            background-color: #1b5e20 !important;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    # 2. Isi Popover saat tombol diklik
    with st.popover("✨ Credits", help="Klik untuk melihat pengembang & info KKN"):
        
        # --- PEMBUAT / DEVELOPERS ---
        st.markdown("**💻 Development Team:**")
        st.markdown(
            """
            * **M.C. Raka Anugrah**
            * **M. Rafi Ardiansyah**
            * **Adreian Alexander A.**
            """
        )
        
        st.divider()
        
        # --- SPECIAL THANKS ---
        st.markdown("**🌟 Special Thanks to:**")
        st.markdown(
            """
            * **Tuhan Yang Maha Esa**
            * **Kelurahan Balas Klumprik**
            * **Dosen Pembimbing Lapangan (DPL)**
            * **Teman-teman KKN Kelompok 27 Surabaya UPNVJT**
            """
        )
        
        st.divider()
        
        # --- LINK SOSIAL MEDIA KKN ---
        st.markdown("**📲 Media Sosial KKN:**")
        
        col_ig, col_tt = st.columns(2)
        with col_ig:
            # Ganti dengan link Instagram KKN Kelompok 27 kamu
            st.link_button("📸 Instagram", "https://www.instagram.com/sabalas27/", use_container_width=True)
        
        with col_tt:
            # Ganti dengan link TikTok KKN Kelompok 27 kamu
            st.link_button("🎵 TikTok", "https://www.tiktok.com/@27sabalas", use_container_width=True)
        
        st.caption("Developed with ❤️ for Kelurahan Balas Klumprik")

# ============================================================
# MAIN
# ============================================================

def main():

    if not check_database():

        st.error(
            "Database Excel tidak ditemukan."
        )

        st.info(
            "Pastikan file berada di:"
        )

        st.code(
            str(EXCEL_FILE)
        )

        st.stop()

    render_floating_credits()

    menu = sidebar()

    if menu == "🏠 Dashboard":

        page_dashboard()

    elif menu == "👥 Nasabah":

        page_nasabah()

    elif menu == "🗂️ Master Item":

        page_master_item()

    elif menu == "💰 Harga Sampah":

        page_harga()

    elif menu == "📥 Transaksi Setoran":

        page_transaksi_setoran()

    elif menu == "📤 Transaksi Penjualan":

        page_transaksi_penjualan()

    elif menu == "📦 Stok":

        page_stok()

    elif menu == "💵 Keuangan":

        page_keuangan()


if __name__ == "__main__":
    main()