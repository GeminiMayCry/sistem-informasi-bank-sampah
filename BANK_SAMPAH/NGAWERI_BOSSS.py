import os
import shutil
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st


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
        "TOTAL HARGA"
    ],

    SHEET_DETAIL_SETORAN: [
        "DETAIL_ID",
        "TRANSAKSI_ID",
        "ITEM_ID",
        "ITEM",
        "KG",
        "HARGA",
        "SUBTOTAL"
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
# EXCEL DATABASE FUNCTIONS
# ============================================================

def check_database():
    """Memastikan file Excel tersedia."""
    return EXCEL_FILE.exists()


@st.cache_data(ttl=2)
def load_sheet(sheet_name):
    """Membaca satu sheet dari Excel dengan normalisasi tipe data."""

    if not EXCEL_FILE.exists():
        return pd.DataFrame()

    try:

        df = pd.read_excel(
            EXCEL_FILE,
            sheet_name=sheet_name
        )

        # =====================================================
        # HILANGKAN KOLOM UNNAMED
        # =====================================================

        df = df.loc[
            :,
            ~df.columns.astype(str).str.startswith("Unnamed")
        ]

        # =====================================================
        # NORMALISASI KOLOM TEKS
        # =====================================================

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

            "TRANSAKSI_ID",
            "DETAIL_ID",

            "PENJUALAN_ID",

            "MUTASI_ID",
            "JENIS_MUTASI",
            "REFERENSI",

            "KEUANGAN_ID",
            "KETERANGAN"
        ]

        for col in text_columns:

            if col in df.columns:

                df[col] = (
                    df[col]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                )

        # =====================================================
        # NORMALISASI KOLOM NUMERIK
        # =====================================================

        numeric_columns = [
            "KG",
            "HARGA",
            "SUBTOTAL",
            "TOTAL KG",
            "TOTAL HARGA",
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

        # =====================================================
        # NORMALISASI KOLOM TANGGAL
        # =====================================================

        date_columns = [
            "TANGGAL",
            "TANGGAL TRANSAKSI",
            "TANGGAL PENDAFTARAN",
            "TANGGAL MULAI BERLAKU"
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
    """Membaca seluruh sheet."""
    data = {}

    for sheet in REQUIRED_COLUMNS:
        data[sheet] = load_sheet(sheet)

    return data


def save_sheet(sheet_name, df):
    """
    Menyimpan satu sheet tanpa menghapus sheet lainnya.
    """

    try:
        # Backup terlebih dahulu
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

        # Bersihkan cache
        load_sheet.clear()

        return True

    except Exception as e:
        st.error(f"Gagal menyimpan data: {e}")
        return False


def save_multiple_sheets(sheet_data):
    """
    Menyimpan beberapa sheet sekaligus.
    """

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
        st.error(f"Gagal menyimpan Excel: {e}")
        return False


def create_backup():
    """Membuat backup Excel sebelum perubahan."""

    if not EXCEL_FILE.exists():
        return

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    backup_file = (
        BACKUP_DIR /
        f"bank_sampah_{timestamp}.xlsx"
    )

    try:
        shutil.copy2(
            EXCEL_FILE,
            backup_file
        )

    except Exception:
        pass


# ============================================================
# HELPER FUNCTIONS
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
            + f"{value:,.0f}"
            .replace(",", ".")
        )

    except Exception:
        return "Rp 0"


def generate_id(df, column, prefix, digits=6):
    """Membuat ID otomatis."""

    if df.empty or column not in df.columns:
        return f"{prefix}{1:0{digits}d}"

    values = df[column].dropna().astype(str)

    numbers = []

    for value in values:

        try:
            number = int(
                value.replace(prefix, "")
            )

            numbers.append(number)

        except Exception:
            continue

    next_number = (
        max(numbers) + 1
        if numbers
        else 1
    )

    return f"{prefix}{next_number:0{digits}d}"


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


def clean_date_column(df, column):
    if column in df.columns:
        df[column] = pd.to_datetime(
            df[column],
            errors="coerce"
        ).dt.strftime("%Y-%m-%d")

    return df


# ============================================================
# PRICE FUNCTION
# ============================================================

def get_price(item_id, tanggal_transaksi):
    """
    Mengambil harga terakhir yang berlaku
    sebelum atau sama dengan tanggal transaksi.

    Contoh:

    01-08 = 5.000
    10-08 = 5.500
    20-08 = 6.000

    Transaksi 15-08 -> 5.500
    """

    df_harga = load_sheet(SHEET_HARGA)

    if df_harga.empty:
        return None

    df_harga = df_harga.copy()

    df_harga["TANGGAL MULAI BERLAKU"] = pd.to_datetime(
        df_harga["TANGGAL MULAI BERLAKU"],
        errors="coerce"
    )

    tanggal_transaksi = pd.to_datetime(
        tanggal_transaksi
    )

    result = df_harga[
        (df_harga["ITEM_ID"].astype(str) == str(item_id))
        &
        (
            df_harga["TANGGAL MULAI BERLAKU"]
            <= tanggal_transaksi
        )
    ].copy()

    if result.empty:
        return None

    # Ambil harga paling baru yang sudah berlaku
    result = result.sort_values(
        "TANGGAL MULAI BERLAKU",
        ascending=False
    )

    try:
        return float(
            result.iloc[0]["HARGA SETOR"]
        )

    except Exception:
        return None


# ============================================================
# STOCK FUNCTION
# ============================================================

def get_stock():
    """
    Menghitung stok berdasarkan MUTASI STOK.
    """

    df = load_sheet(SHEET_STOK)
    df_item = load_sheet(SHEET_MASTER_ITEM)

    if df_item.empty:
        return pd.DataFrame()

    result = df_item[
        ["ITEM_ID", "JENIS", "ITEM", "SATUAN"]
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
        df[df["JENIS_MUTASI"] == "MASUK"]
        .groupby("ITEM_ID")["KG"]
        .sum()
    )

    keluar = (
        df[df["JENIS_MUTASI"] == "KELUAR"]
        .groupby("ITEM_ID")["KG"]
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
# SIDEBAR
# ============================================================

def sidebar():

    st.sidebar.title("♻️ BANK SAMPAH")

    st.sidebar.caption(
        "Sistem Informasi Bank Sampah"
    )

    st.sidebar.divider()

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
        ]
    )

    st.sidebar.divider()

    if EXCEL_FILE.exists():

        st.sidebar.success(
            "Database Excel terhubung"
        )

    else:

        st.sidebar.error(
            "Database Excel tidak ditemukan"
        )

    return menu


# ============================================================
# DASHBOARD
# ============================================================

def page_dashboard():

    st.title("🏠 Dashboard Bank Sampah")

    st.caption(
        "Ringkasan aktivitas Bank Sampah"
    )

    data = load_all_data()

    df_nasabah = data[SHEET_NASABAH]
    df_setoran = data[SHEET_SETORAN]
    df_penjualan = data[SHEET_PENJUALAN]
    df_stok = get_stock()
    df_keuangan = data[SHEET_KEUANGAN]

    # --------------------------------------------------------
    # KPI
    # --------------------------------------------------------

    total_nasabah = len(df_nasabah)

    total_setoran_kg = 0
    total_setoran_rp = 0

    if not df_setoran.empty:

        total_setoran_kg = pd.to_numeric(
            df_setoran["TOTAL KG"],
            errors="coerce"
        ).fillna(0).sum()

        total_setoran_rp = pd.to_numeric(
            df_setoran["TOTAL HARGA"],
            errors="coerce"
        ).fillna(0).sum()

    total_penjualan_kg = 0
    total_penjualan_rp = 0

    if not df_penjualan.empty:

        total_penjualan_kg = pd.to_numeric(
            df_penjualan["TOTAL KG"],
            errors="coerce"
        ).fillna(0).sum()

        total_penjualan_rp = pd.to_numeric(
            df_penjualan["TOTAL HARGA"],
            errors="coerce"
        ).fillna(0).sum()

    total_stok = 0

    if not df_stok.empty:
        total_stok = df_stok["STOK"].sum()

    total_debit = 0
    total_kredit = 0

    if not df_keuangan.empty:

        total_debit = pd.to_numeric(
            df_keuangan["DEBIT"],
            errors="coerce"
        ).fillna(0).sum()

        total_kredit = pd.to_numeric(
            df_keuangan["KREDIT"],
            errors="coerce"
        ).fillna(0).sum()

    saldo = total_kredit - total_debit

    col1, col2, col3, col4 = st.columns(4)

    col1.metric(
        "Total Nasabah",
        f"{total_nasabah:,}".replace(",", ".")
    )

    col2.metric(
        "Total Setoran",
        f"{total_setoran_kg:,.2f} KG"
    )

    col3.metric(
        "Total Stok",
        f"{total_stok:,.2f} KG"
    )

    col4.metric(
        "Saldo",
        format_rupiah(saldo)
    )

    st.divider()

    # --------------------------------------------------------
    # GRAFIK
    # --------------------------------------------------------

    col1, col2 = st.columns(2)

    with col1:

        st.subheader("Setoran per Bulan")

        df_detail = data[SHEET_DETAIL_SETORAN].copy()

        if not df_detail.empty:

            df_header = df_setoran.copy()

            if not df_header.empty:

                df_header["TANGGAL TRANSAKSI"] = pd.to_datetime(
                    df_header["TANGGAL TRANSAKSI"],
                    errors="coerce"
                )

                df_header["BULAN"] = (
                    df_header["TANGGAL TRANSAKSI"]
                    .dt.to_period("M")
                    .astype(str)
                )

                chart = (
                    df_header
                    .groupby("BULAN")["TOTAL KG"]
                    .sum()
                )

                st.bar_chart(chart)

            else:

                st.info(
                    "Belum ada transaksi setoran."
                )

        else:

            st.info(
                "Belum ada transaksi setoran."
            )

    with col2:

        st.subheader("Stok per Jenis Sampah")

        if not df_stok.empty:

            chart_stock = (
                df_stok
                .groupby("JENIS")["STOK"]
                .sum()
            )

            st.bar_chart(chart_stock)

        else:

            st.info(
                "Belum ada data stok."
            )

    # --------------------------------------------------------
    # STOK TERBANYAK
    # --------------------------------------------------------

    st.subheader("Stok Sampah")

    if not df_stok.empty:

        display = df_stok.copy()

        display["STOK"] = display["STOK"].round(2)

        st.dataframe(
            display[
                [
                    "JENIS",
                    "ITEM",
                    "STOK",
                    "SATUAN"
                ]
            ],
            use_container_width=True,
            hide_index=True
        )


# ============================================================
# NASABAH
# ============================================================

def page_nasabah():

    st.title("👥 Data Nasabah")

    df = load_sheet(SHEET_NASABAH)
    df_nasabah = df.copy()

    tab1, tab2, tab3, tab4 = st.tabs(
    [
        "📋 Data Nasabah",
        "➕ Tambah Nasabah",
        "✏️ Edit / Hapus",
        "📜 Detail Transaksi"
        ]
    )

    # --------------------------------------------------------
    # DATA
    # --------------------------------------------------------

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
                    display["NIK"]
                    .astype(str)
                    .str.contains(
                        search,
                        case=False,
                        na=False
                    )
                    |
                    display["NAMA"]
                    .astype(str)
                    .str.contains(
                        search,
                        case=False,
                        na=False
                    )
                )

                display = display[mask]

            st.dataframe(
                display,
                use_container_width=True,
                hide_index=True
            )

            st.caption(
                f"Jumlah nasabah: {len(display)}"
            )

    # --------------------------------------------------------
    # TAMBAH
    # --------------------------------------------------------

    with tab2:

        with st.form("form_tambah_nasabah"):

            col1, col2 = st.columns(2)

            nik = col1.text_input("NIK")
            nama = col2.text_input("Nama")

            rt = col1.text_input("RT")
            rw = col2.text_input("RW")

            alamat = st.text_area(
                "Alamat"
            )

            col1, col2, col3 = st.columns(3)

            hp = col1.text_input("HP")
            rekening = col2.text_input("Rekening")
            tanggal = col3.date_input(
                "Tanggal Pendaftaran",
                value=date.today()
            )

            submit = st.form_submit_button(
                "💾 Simpan Nasabah",
                use_container_width=True
            )

        if submit:

            nik = nik.strip()
            nama = nama.strip()

            if not nik:

                st.error(
                    "NIK wajib diisi."
                )

            elif not nama:

                st.error(
                    "Nama wajib diisi."
                )

            elif not df.empty and (
                df["NIK"]
                .astype(str)
                .str.strip()
                .eq(nik)
                .any()
            ):

                st.error(
                    "NIK sudah terdaftar."
                )

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

                df = pd.concat(
                    [df, new_data],
                    ignore_index=True
                )

                if save_sheet(
                    SHEET_NASABAH,
                    df
                ):

                    st.success(
                        "Nasabah berhasil ditambahkan."
                    )

                    st.rerun()

    # --------------------------------------------------------
    # EDIT / HAPUS
    # --------------------------------------------------------

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
                nik_list
            )

            index = df[
                df["NIK"].astype(str)
                == selected_nik
            ].index[0]

            row = df.loc[index]

            with st.form("form_edit_nasabah"):

                col1, col2 = st.columns(2)

                nama = col1.text_input(
                    "Nama",
                    value=normalize_text(row["NAMA"])
                )

                rt = col2.text_input(
                    "RT",
                    value=normalize_text(row["RT"])
                )

                rw = col1.text_input(
                    "RW",
                    value=normalize_text(row["RW"])
                )

                alamat = col2.text_input(
                    "Alamat",
                    value=normalize_text(row["ALAMAT"])
                )

                hp = col1.text_input(
                    "HP",
                    value=normalize_text(row["HP"])
                )

                rekening = col2.text_input(
                    "Rekening",
                    value=normalize_text(row["REKENING"])
                )

                tanggal_default = pd.to_datetime(
                    row["TANGGAL PENDAFTARAN"],
                    errors="coerce"
                )

                if pd.isna(tanggal_default):

                    tanggal_default = date.today()

                else:

                    tanggal_default = tanggal_default.date()

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

                df.at[index, "NAMA"] = nama
                df.at[index, "RT"] = rt
                df.at[index, "RW"] = rw
                df.at[index, "ALAMAT"] = alamat
                df.at[index, "HP"] = hp
                df.at[index, "REKENING"] = rekening
                df.at[index, "TANGGAL PENDAFTARAN"] = (
                    tanggal.strftime("%Y-%m-%d")
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

                df = df.drop(index).reset_index(
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
# TAB 4 - DETAIL TRANSAKSI NASABAH
# ========================================================

    with tab4:

        df_setoran = load_sheet(
            SHEET_SETORAN
        )

        df_detail = load_sheet(
            SHEET_DETAIL_SETORAN
        )

        df_item = load_sheet(
            SHEET_MASTER_ITEM
        )

        if df_setoran.empty:

            st.info(
                "Belum ada transaksi setoran."
            )

        else:

            st.subheader(
                "📜 Riwayat Setoran Nasabah"
            )

            # ------------------------------------------------
            # PILIH NASABAH
            # ------------------------------------------------

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
            # FILTER PERIODE
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
            # FILTER ITEM
            # ------------------------------------------------

            item_options = ["SEMUA"]

            if not df_detail.empty:

                item_options += sorted(
                    df_detail["ITEM"]
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
            # GABUNGKAN TRANSAKSI
            # ------------------------------------------------

            transaksi = df_setoran[
                df_setoran["NIK"]
                .astype(str)
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

                # Filter tanggal
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

                # ------------------------------------------------
                # FILTER ITEM
                # ------------------------------------------------

                if selected_item != "SEMUA":

                    if not df_detail.empty:

                        detail_filter = df_detail[
                            (
                                df_detail["ITEM"]
                                .astype(str)
                                == selected_item
                            )
                        ][
                            "TRANSAKSI_ID"
                        ].astype(str).unique()

                        transaksi = transaksi[
                            transaksi[
                                "TRANSAKSI_ID"
                            ]
                            .astype(str)
                            .isin(detail_filter)
                        ]

                # ------------------------------------------------
                # SUMMARY
                # ------------------------------------------------

                total_transaksi = len(
                    transaksi
                )

                total_kg = pd.to_numeric(
                    transaksi["TOTAL KG"],
                    errors="coerce"
                ).fillna(0).sum()

                total_uang = pd.to_numeric(
                    transaksi["TOTAL HARGA"],
                    errors="coerce"
                ).fillna(0).sum()

                col1, col2, col3 = st.columns(3)

                col1.metric(
                    "Jumlah Transaksi",
                    total_transaksi
                )

                col2.metric(
                    "Total Sampah",
                    f"{total_kg:,.2f} KG"
                )

                col3.metric(
                    "Total Setoran",
                    format_rupiah(total_uang)
                )

                st.divider()

                # ------------------------------------------------
                # TABEL TRANSAKSI
                # ------------------------------------------------

                transaksi_display = transaksi.copy()

                transaksi_display[
                    "TANGGAL TRANSAKSI"
                ] = transaksi_display[
                    "TANGGAL TRANSAKSI"
                ].dt.strftime(
                    "%d-%m-%Y"
                )

                transaksi_display[
                    "TOTAL HARGA"
                ] = transaksi_display[
                    "TOTAL HARGA"
                ].apply(format_rupiah)

                st.dataframe(
                    transaksi_display[
                        [
                            "TRANSAKSI_ID",
                            "NO",
                            "TANGGAL TRANSAKSI",
                            "NIK",
                            "NAMA",
                            "TOTAL KG",
                            "TOTAL HARGA"
                        ]
                    ],
                    use_container_width=True,
                    hide_index=True
                )

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
                            detail["ITEM"]
                            .astype(str)
                            == selected_item
                        ]

                    if not detail.empty:

                        detail["KG"] = pd.to_numeric(
                            detail["KG"],
                            errors="coerce"
                        ).fillna(0)

                        detail["SUBTOTAL"] = pd.to_numeric(
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
                                TOTAL_KG=("KG", "sum"),
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

                    else:

                        st.info(
                            "Tidak ada detail item "
                            "sesuai filter."
                        )

# ============================================================
# MASTER ITEM
# ============================================================

def page_master_item():

    st.title("🗂️ Master Item Sampah")

    df = load_sheet(SHEET_MASTER_ITEM)

    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True
    )

    st.caption(
        f"Total item: {len(df)}"
    )


# ============================================================
# HARGA
# ============================================================

def page_harga():

    st.title("💰 Daftar Harga Sampah")

    df_harga = load_sheet(SHEET_HARGA)
    df_item = load_sheet(SHEET_MASTER_ITEM)

    tab1, tab2, tab3 = st.tabs(
        [
            "📋 Daftar Harga",
            "➕ Tambah Harga",
            "✏️ Edit Harga"
        ]
    )

    # ========================================================
    # TAB 1 - DAFTAR HARGA
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
            ).dt.strftime("%d-%m-%Y")

            display["HARGA SETOR"] = display[
                "HARGA SETOR"
            ].apply(format_rupiah)

            st.dataframe(
                display.sort_values(
                    "TANGGAL MULAI BERLAKU",
                    ascending=False
                ),
                use_container_width=True,
                hide_index=True
            )

    # ========================================================
    # TAB 2 - TAMBAH HARGA
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
                    df_item["JENIS"]
                    .dropna()
                    .unique()
                    .tolist()
                ),
                key="jenis_tambah_harga"
            )

            items_jenis = df_item[
                df_item["JENIS"] == jenis
            ]

            item_id = st.selectbox(
                "Item",
                items_jenis["ITEM_ID"].tolist(),
                format_func=lambda x:
                    items_jenis.loc[
                        items_jenis["ITEM_ID"] == x,
                        "ITEM"
                    ].iloc[0],
                key="item_tambah_harga"
            )

            item_name = items_jenis.loc[
                items_jenis["ITEM_ID"] == item_id,
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
                "Harga ini akan digunakan untuk transaksi "
                "pada tanggal mulai berlaku sampai ada harga "
                "baru yang berlaku."
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

                # ------------------------------------------------
                # CEK DUPLIKAT TANGGAL + ITEM
                # ------------------------------------------------

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
                            df_harga["ITEM_ID"]
                            .astype(str)
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
    # TAB 3 - EDIT HARGA
    # ========================================================

    with tab3:

        if df_harga.empty:

            st.info(
                "Belum ada harga untuk diedit."
            )

        else:

            # -----------------------------------------------
            # PILIH HARGA
            # -----------------------------------------------

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

            selected_id = selected.split(" | ")[0]

            index_list = df_harga[
                df_harga["HARGA_ID"].astype(str)
                == selected_id
            ].index.tolist()

            if not index_list:

                st.error(
                    "Data harga tidak ditemukan."
                )

                return

            index = index_list[0]

            row = df_harga.loc[index]

            # -----------------------------------------------
            # FORM EDIT
            # -----------------------------------------------

            tanggal_existing = pd.to_datetime(
                row[
                    "TANGGAL MULAI BERLAKU"
                ],
                errors="coerce"
            )

            if pd.isna(tanggal_existing):

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

            item_id = row["ITEM_ID"]

            st.text_input(
                "Item",
                value=str(row["ITEM"]),
                disabled=True
            )

            harga_baru = st.number_input(
                "Harga Setor / KG",
                min_value=0,
                step=100,
                value=int(
                    float(row["HARGA SETOR"])
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

            # -----------------------------------------------
            # UPDATE
            # -----------------------------------------------

            if update:

                if harga_baru <= 0:

                    st.error(
                        "Harga harus lebih dari 0."
                    )

                    st.stop()

                # Cek bentrok dengan data lain
                tanggal_existing_all = pd.to_datetime(
                    df_harga[
                        "TANGGAL MULAI BERLAKU"
                    ],
                    errors="coerce"
                )

                duplicate = (
                    (
                        df_harga["ITEM_ID"]
                        .astype(str)
                        == str(item_id)
                    )
                    &
                    (
                        tanggal_existing_all.dt.date
                        == tanggal_baru
                    )
                    &
                    (
                        df_harga["HARGA_ID"]
                        .astype(str)
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

            # -----------------------------------------------
            # DELETE
            # -----------------------------------------------

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

    st.title("📥 Transaksi Setoran")

    df_nasabah = load_sheet(SHEET_NASABAH)
    df_item = load_sheet(SHEET_MASTER_ITEM)
    df_harga = load_sheet(SHEET_HARGA)
    df_setoran = load_sheet(SHEET_SETORAN)
    df_detail = load_sheet(SHEET_DETAIL_SETORAN)
    df_stok = load_sheet(SHEET_STOK)
    df_keuangan = load_sheet(SHEET_KEUANGAN)

    if df_nasabah.empty:

        st.warning(
            "Belum ada nasabah. Tambahkan nasabah terlebih dahulu."
        )

        return

    if df_item.empty:

        st.warning(
            "Master item belum tersedia."
        )

        return

    # --------------------------------------------------------
    # HEADER TRANSAKSI
    # --------------------------------------------------------

    st.subheader("Data Transaksi")

    col1, col2 = st.columns(2)

    tanggal = col1.date_input(
        "Tanggal Transaksi",
        value=date.today()
    )

    nasabah_options = (
        df_nasabah["NIK"]
        .astype(str)
        .tolist()
    )

    selected_nik = col2.selectbox(
        "Nasabah",
        nasabah_options,
        format_func=lambda nik:
            f"{nik} - "
            + str(
                df_nasabah.loc[
                    df_nasabah["NIK"].astype(str)
                    == str(nik),
                    "NAMA"
                ].iloc[0]
            )
    )

    nasabah = df_nasabah[
        df_nasabah["NIK"].astype(str)
        == str(selected_nik)
    ].iloc[0]

    col1, col2, col3 = st.columns(3)

    col1.text_input(
        "NIK",
        value=str(nasabah["NIK"]),
        disabled=True
    )

    col2.text_input(
        "Nama",
        value=str(nasabah["NAMA"]),
        disabled=True
    )

    col3.text_input(
        "Periode Harga",
        value=tanggal.strftime("%Y-%m"),
        disabled=True
    )

    st.divider()

    st.subheader("Detail Sampah")

    # --------------------------------------------------------
    # INPUT ITEM
    # --------------------------------------------------------

    detail_rows = []

    for jenis in df_item["JENIS"].dropna().unique():

        st.markdown(
            f"### {jenis}"
        )

        items_jenis = df_item[
            df_item["JENIS"] == jenis
        ]

        for _, item in items_jenis.iterrows():

            item_id = item["ITEM_ID"]
            item_name = item["ITEM"]

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

                kg = col2.number_input(
                    "KG",
                    min_value=0.0,
                    step=0.1,
                    value=0.0,
                    key=f"kg_{item_id}"
                )

            else:

                col3.write(
                    format_rupiah(harga)
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

                subtotal = kg * harga

                detail_rows.append({
                    "ITEM_ID": item_id,
                    "ITEM": item_name,
                    "KG": kg,
                    "HARGA": harga,
                    "SUBTOTAL": subtotal
                })

    # --------------------------------------------------------
    # RINGKASAN
    # --------------------------------------------------------

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
            "Total Setoran",
            format_rupiah(total_harga)
        )

        st.dataframe(
            detail_input,
            use_container_width=True,
            hide_index=True
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

            # --------------------------------------------
            # HEADER
            # --------------------------------------------

            new_header = pd.DataFrame(
                [[
                    transaksi_id,
                    no,
                    tanggal.strftime("%Y-%m-%d"),
                    str(nasabah["NIK"]),
                    str(nasabah["NAMA"]),
                    total_kg,
                    total_harga
                ]],
                columns=REQUIRED_COLUMNS[SHEET_SETORAN]
            )

            df_setoran = pd.concat(
                [df_setoran, new_header],
                ignore_index=True
            )

            # --------------------------------------------
            # DETAIL
            # --------------------------------------------

            new_details = []

            for _, row in detail_input.iterrows():

                detail_id = generate_id(
                    df_detail,
                    "DETAIL_ID",
                    "DT",
                    6
                )

                new_details.append([
                    detail_id,
                    transaksi_id,
                    row["ITEM_ID"],
                    row["ITEM"],
                    row["KG"],
                    row["HARGA"],
                    row["SUBTOTAL"]
                ])

                # Tambahkan ke dataframe agar
                # generate ID berikutnya berbeda
                df_detail = pd.concat(
                    [
                        df_detail,
                        pd.DataFrame(
                            [[
                                detail_id,
                                transaksi_id,
                                row["ITEM_ID"],
                                row["ITEM"],
                                row["KG"],
                                row["HARGA"],
                                row["SUBTOTAL"]
                            ]],
                            columns=REQUIRED_COLUMNS[
                                SHEET_DETAIL_SETORAN
                            ]
                        )
                    ],
                    ignore_index=True
                )

            # --------------------------------------------
            # MUTASI STOK
            # --------------------------------------------

            new_stock_rows = []

            for _, row in detail_input.iterrows():

                mutasi_id = generate_id(
                    df_stok,
                    "MUTASI_ID",
                    "M",
                    6
                )

                new_stock_rows.append([
                    mutasi_id,
                    tanggal.strftime("%Y-%m-%d"),
                    row["ITEM_ID"],
                    row["ITEM"],
                    "MASUK",
                    row["KG"],
                    transaksi_id
                ])

                df_stok = pd.concat(
                    [
                        df_stok,
                        pd.DataFrame(
                            [[
                                mutasi_id,
                                tanggal.strftime("%Y-%m-%d"),
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
                    ],
                    ignore_index=True
                )

            # --------------------------------------------
            # KEUANGAN
            # --------------------------------------------

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
                    f"Setoran {nasabah['NAMA']}",
                    0,
                    total_harga,
                    transaksi_id
                ]],
                columns=REQUIRED_COLUMNS[
                    SHEET_KEUANGAN
                ]
            )

            df_keuangan = pd.concat(
                [
                    df_keuangan,
                    new_keuangan
                ],
                ignore_index=True
            )

            # --------------------------------------------
            # SIMPAN SEMUA
            # --------------------------------------------

            success = save_multiple_sheets({
                SHEET_SETORAN: df_setoran,
                SHEET_DETAIL_SETORAN: df_detail,
                SHEET_STOK: df_stok,
                SHEET_KEUANGAN: df_keuangan
            })

            if success:

                st.success(
                    f"Transaksi {transaksi_id} berhasil disimpan."
                )

                st.rerun()

    else:

        st.info(
            "Masukkan KG pada minimal satu jenis sampah."
        )


# ============================================================
# TRANSAKSI PENJUALAN
# ============================================================

def page_transaksi_penjualan():

    st.title("📤 Transaksi Penjualan")

    df_item = load_sheet(SHEET_MASTER_ITEM)
    df_penjualan = load_sheet(SHEET_PENJUALAN)
    df_detail = load_sheet(SHEET_DETAIL_PENJUALAN)
    df_stok = load_sheet(SHEET_STOK)
    df_keuangan = load_sheet(SHEET_KEUANGAN)

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

    for jenis in df_item["JENIS"].dropna().unique():

        st.markdown(
            f"### {jenis}"
        )

        items_jenis = df_item[
            df_item["JENIS"] == jenis
        ]

        for _, item in items_jenis.iterrows():

            item_id = item["ITEM_ID"]
            item_name = item["ITEM"]

            stock_row = current_stock[
                current_stock["ITEM_ID"] == item_id
            ]

            if stock_row.empty:

                stok = 0

            else:

                stok = float(
                    stock_row.iloc[0]["STOK"]
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
                max_value=max(stok, 0),
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

                subtotal = kg * harga

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

        total_kg = detail_input["KG"].sum()
        total_harga = detail_input["SUBTOTAL"].sum()

        col1, col2 = st.columns(2)

        col1.metric(
            "Total Berat",
            f"{total_kg:,.2f} KG"
        )

        col2.metric(
            "Total Penjualan",
            format_rupiah(total_harga)
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

            # --------------------------------------------
            # HEADER
            # --------------------------------------------

            new_header = pd.DataFrame(
                [[
                    penjualan_id,
                    tanggal.strftime("%Y-%m-%d"),
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

            # --------------------------------------------
            # DETAIL
            # --------------------------------------------

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

                # ----------------------------------------
                # MUTASI STOK KELUAR
                # ----------------------------------------

                mutasi_id = generate_id(
                    df_stok,
                    "MUTASI_ID",
                    "M",
                    6
                )

                new_stock = pd.DataFrame(
                    [[
                        mutasi_id,
                        tanggal.strftime("%Y-%m-%d"),
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

            # --------------------------------------------
            # KEUANGAN
            # --------------------------------------------

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
                    f"Penjualan ke {pengepul.strip()}",
                    total_harga,
                    0,
                    penjualan_id
                ]],
                columns=REQUIRED_COLUMNS[
                    SHEET_KEUANGAN
                ]
            )

            df_keuangan = pd.concat(
                [
                    df_keuangan,
                    new_finance
                ],
                ignore_index=True
            )

            success = save_multiple_sheets({
                SHEET_PENJUALAN: df_penjualan,
                SHEET_DETAIL_PENJUALAN: df_detail,
                SHEET_STOK: df_stok,
                SHEET_KEUANGAN: df_keuangan
            })

            if success:

                st.success(
                    f"Penjualan {penjualan_id} berhasil disimpan."
                )

                st.rerun()

    else:

        st.info(
            "Masukkan KG pada minimal satu item."
        )


# ============================================================
# STOK
# ============================================================

def page_stok():

    st.title("📦 Stok Bank Sampah")

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
        + sorted(
            df_stok["JENIS"]
            .dropna()
            .unique()
            .tolist()
        )
    )

    display = df_stok.copy()

    if jenis_filter != "SEMUA":

        display = display[
            display["JENIS"] == jenis_filter
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

    st.title("💵 Rekapitulasi Keuangan")

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

    total_debit = df["DEBIT"].sum()
    total_kredit = df["KREDIT"].sum()

    saldo = total_kredit - total_debit

    col1, col2, col3 = st.columns(3)

    col1.metric(
        "Debit / Pengeluaran",
        format_rupiah(total_debit)
    )

    col2.metric(
        "Kredit / Pemasukan",
        format_rupiah(total_kredit)
    )

    col3.metric(
        "Saldo",
        format_rupiah(saldo)
    )

    st.divider()

    display = df.copy()

    display["DEBIT"] = display[
        "DEBIT"
    ].apply(format_rupiah)

    display["KREDIT"] = display[
        "KREDIT"
    ].apply(format_rupiah)

    st.dataframe(
        display,
        use_container_width=True,
        hide_index=True
    )


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
