import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

# ---- KONFIGURASI
DATA_FILE = "Data_Penjualan.xlsx"
NAMA      = "Afida Mar'atus Solihah"
NIM       = "231064005"
APP_TITLE = "Dashboard Analisis Data Penjualan"
SUMBER_DATA = "Dataset sekunder data transaksi supermarket (Data_Penjualan.xlsx). Isi dengan tautan/sumber asli yang dapat dipertanggungjawabkan (mis. Kaggle: Supermarket Sales)."
LATAR_BELAKANG = "Supermarket perlu memahami faktor yang berkaitan dengan besarnya nilai transaksi agar strategi promosi, penataan produk, dan pelayanan dapat dirancang berdasarkan data. Nilai transaksi dapat dipengaruhi oleh jumlah barang yang dibeli, karakteristik pelanggan, lini produk, cabang, metode pembayaran, dan waktu transaksi. Analisis eksplorasi dan regresi linear berganda digunakan untuk mengidentifikasi faktor-faktor tersebut."
TUJUAN = [
  "Menggambarkan karakteristik setiap variabel pada data transaksi (univariat).",
  "Mengidentifikasi hubungan dan perbedaan kelompok antar dua variabel (bivariat).",
  "Mengeksplorasi pola yang melibatkan tiga variabel atau lebih (multivariat).",
  "Membangun model regresi linear berganda untuk menjelaskan nilai transaksi (Total) dan mengevaluasi kemampuannya.",
  "Merumuskan kesimpulan dan rekomendasi bagi pengelola supermarket."
]

# ---- Konstanta variabel ------------------------------------------------------
cat_base = ["Branch", "Customer type", "Gender", "Product line", "Payment"]
cat_all  = [cat_base, "Bulan", "Hari"]
num_all  = ["Quantity", "Total", "Rating", "Hour"]
NONE = "(tidak ada)"
bulan_lv = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus",
              "September","Oktober","November","Desember"]
hari_lv  = ["Senin","Selasa","Rabu","Kamis","Jumat","Sabtu","Minggu"]

kamus = data.frame(
  Variabel = ["Branch","Customer type","Gender","Product line","Quantity","Total",
               "Date","Time","Payment","Rating","Hour (turunan)","Bulan (turunan)","Hari (turunan)"],
  Tipe = ["Kategorik","Kategorik","Kategorik","Kategorik","Numerik (diskret)","Numerik (kontinu)",
           "Tanggal","Waktu","Kategorik","Numerik (kontinu)","Numerik (diskret)","Kategorik","Kategorik"],
  Definisi = ["Cabang supermarket tempat transaksi (A, B, C).",
               "Tipe pelanggan: Member atau Normal.",
               "Jenis kelamin pelanggan: Female atau Male.",
               "Lini produk yang dibeli (6 kategori).",
               "Jumlah unit barang yang dibeli dalam satu transaksi (1-10).",
               "Nilai total transaksi termasuk pajak 5% (satuan mata uang pada data).",
               "Tanggal transaksi.",
               "Jam transaksi.",
               "Metode pembayaran: Cash, Credit card, Ewallet.",
               "Penilaian kepuasan pelanggan (skala 1-10, nilai teramati 4-10).",
               "Jam transaksi (0-23) yang diambil dari Time.",
               "Bulan transaksi yang diambil dari Date.",
               "Hari dalam seminggu yang diambil dari Date."],
  stringsAsFactors = FALSE)

# ---- Membaca dan menyiapkan data --------------------------------------------
import pandas as pd
import os

# Kolom kategorik
cat_base = [
    "Branch",
    "Customer type",
    "Gender",
    "Product line",
    "Payment"
]

# Nama bulan
bulan_lv = [
    "Januari", "Februari", "Maret", "April",
    "Mei", "Juni", "Juli", "Agustus",
    "September", "Oktober", "November", "Desember"
]

# Nama hari
hari_lv = [
    "Senin", "Selasa", "Rabu", "Kamis",
    "Jumat", "Sabtu", "Minggu"
]


def read_sales(path):

    # Membaca ekstensi file
    ext = os.path.splitext(path.name if hasattr(path, "name") else path)[1].lower()

    # Membaca dataset Excel atau CSV
    if ext in [".xlsx", ".xls"]:
        d = pd.read_excel(path)
    elif ext == ".csv":
        d = pd.read_csv(path)
    else:
        return None

    # Kolom yang wajib tersedia
    need = [
        "Branch", "Customer type", "Gender",
        "Product line", "Quantity", "Total",
        "Date", "Time", "Payment", "Rating"
    ]

    # Memeriksa kelengkapan kolom
    if not all(col in d.columns for col in need):
        return None

    # Mengubah Date menjadi datetime
    d["Date"] = pd.to_datetime(d["Date"], errors="coerce")

    # Mengambil jam dari kolom Time
    d["Hour"] = pd.to_numeric(
        d["Time"].astype(str).str.extract(r"(\d{1,2}):")[0],
        errors="coerce"
    )

    # Mengubah kolom numerik
    for v in ["Quantity", "Total", "Rating"]:
        d[v] = pd.to_numeric(d[v], errors="coerce")

    # Mengubah kolom kategorik
    for v in cat_base:
        d[v] = d[v].astype("category")

    # Membuat variabel Bulan
    d["Bulan"] = pd.Categorical(
        d["Date"].dt.month.map(
            lambda x: bulan_lv[x - 1] if pd.notna(x) else None
        ),
        categories=bulan_lv,
        ordered=True
    )

    # Membuat variabel Hari
    d["Hari"] = pd.Categorical(
        d["Date"].dt.dayofweek.map(
            lambda x: hari_lv[x] if pd.notna(x) else None
        ),
        categories=hari_lv,
        ordered=True
    )

    # Mengembalikan kolom yang diperlukan
    return d[
        cat_base + [
            "Bulan", "Hari", "Quantity",
            "Total", "Rating", "Hour", "Date"
        ]
    ]

# ---- Palet & tema grafik (pink formal) ---------------------------------------
pk <- list(d = "#880E4F", m = "#C2185B", s = "#AD1457", l = "#F48FB1", xl = "#F8BBD0",
           bg = "#FFF5F8", txt = "#3B2A33", grid = "#EFDDE4", mauve = "#6D4C5E", slate = "#4F6D8F")
def pal_cat(n):
    colors = [
        "#AD1457",
        "#F48FB1",
        "#6D4C5E",
        "#E8BBD0",
        "#D81B60",
        "#B39DAB",
        "#880E4F"
    ]

    return [colors[i % len(colors)] for i in range(n)]
def theme_pink(ax, pk, base=12):
    # Background putih
    ax.set_facecolor("white")
    ax.figure.set_facecolor("white")

    # Judul grafik
    ax.title.set_fontweight("bold")
    ax.title.set_color(pk["d"])
    ax.title.set_fontsize(base + 2)

    # Judul sumbu
    ax.xaxis.label.set_fontweight("bold")
    ax.xaxis.label.set_color(pk["mauve"])

    ax.yaxis.label.set_fontweight("bold")
    ax.yaxis.label.set_color(pk["mauve"])

    # Teks sumbu
    ax.tick_params(axis="both", colors=pk["txt"])

    # Grid
    ax.grid(
        True,
        which="major",
        color=pk["grid"],
        linewidth=0.7
    )

    ax.grid(False, which="minor")
    ax.minorticks_off()

    # Legend
    legend = ax.get_legend()

    if legend is not None:
        legend.get_title().set_fontweight("bold")
        legend.get_title().set_color(pk["mauve"])

    return ax
def rot(ax):
    # Memiringkan label sumbu X
    ax.tick_params(axis="x", labelrotation=20)

    for label in ax.get_xticklabels():
        label.set_horizontalalignment("right")

    # Mengatur margin grafik
    ax.figure.subplots_adjust(
        left=0.15,
        right=0.98,
        bottom=0.22,
        top=0.90
    )

    return ax
# ---- Fungsi bantu ------------------------------------------------------------
import numpy as np
import pandas as pd
import re


# Format angka dengan pemisah ribuan dan desimal
def fmt(x, k=2):
    if pd.isna(x):
        return "-"
    return f"{x:,.{k}f}"


# Format p-value
def fmt_p(p):
    if pd.isna(p):
        return "-"
    return "< 0.001" if p < 0.001 else f"{p:.3f}"


# Format p-value untuk kalimat
def pq(p):
    if pd.isna(p):
        return "p tidak tersedia"
    return "< 0.001" if p < 0.001 else f"= {p:.3f}"


# Tanda signifikansi
def stars(p):
    if pd.isna(p):
        return ""
    if p < 0.001:
        return "***"
    elif p < 0.01:
        return "**"
    elif p < 0.05:
        return "*"
    elif p < 0.1:
        return "."
    else:
        return ""


# Fungsi skewness
def skew_fun(x):
    x = np.asarray(x, dtype=float)
    x = x[~np.isnan(x)]

    n = len(x)

    if n < 3 or np.std(x, ddof=1) == 0:
        return np.nan

    z = (x - np.mean(x)) / np.std(x, ddof=1)

    return (
        n / ((n - 1) * (n - 2))
        * np.sum(z ** 3)
    )


# Fungsi kurtosis
def kurt_fun(x):
    x = np.asarray(x, dtype=float)
    x = x[~np.isnan(x)]

    n = len(x)

    if n < 4 or np.std(x, ddof=1) == 0:
        return np.nan

    z = (x - np.mean(x)) / np.std(x, ddof=1)

    return (
        n * (n + 1) / ((n - 1) * (n - 2) * (n - 3))
        * np.sum(z ** 4)
        - 3 * (n - 1) ** 2 / ((n - 2) * (n - 3))
    )


# Format teks dengan bold
def md(x):
    return str(x)


# Membuat bullet points
def bullets(lines):
    return "\n".join([f"- {line}" for line in lines])


# Interpretasi kekuatan korelasi
def strength(r):
    a = abs(r)

    if a < 0.2:
        return "sangat lemah"
    elif a < 0.4:
        return "lemah"
    elif a < 0.6:
        return "sedang"
    elif a < 0.8:
        return "kuat"
    else:
        return "sangat kuat"

# =============================================================================
# 2. UNIVARIAT
# =============================================================================
def uni_choices(v):
    if v == "Date":
        return ["Deret waktu harian"]

    elif v in cat_all:
        return [
            "Diagram batang",
            "Pareto",
            "Diagram lingkaran"
        ]

    else:
        return [
            "Histogram",
            "Histogram + Density",
            "Boxplot",
            "Density",
            "Stem-and-leaf",
            "QQ-plot"
        ]

def plot_uni(d, v, type):

    # Mengambil warna dari palet pink
    colors = pal_cat(20)

    # ==========================================
    # 1. DERET WAKTU HARIAN
    # ==========================================
    if v == "Date":

        dd = d.groupby("Date").size().reset_index(name="n")
        dd = dd.sort_values("Date")

        fig, ax = plt.subplots(figsize=(10, 5))

        ax.plot(
            dd["Date"], dd["n"],
            color=pk["l"],
            linewidth=1.5,
            label="Jumlah transaksi"
        )

        ax.scatter(
            dd["Date"], dd["n"],
            color=pk["m"],
            s=15
        )

        # Garis tren LOESS
        if len(dd) >= 3:
            x = np.arange(len(dd))
            trend = lowess(dd["n"], x, frac=0.3)

            ax.plot(
                dd["Date"],
                trend[:, 1],
                color=pk["d"],
                linewidth=2.5,
                label="Tren (LOESS)"
            )

        ax.set_title(
            "Jumlah transaksi per hari",
            fontweight="bold",
            color=pk["d"]
        )

        ax.set_xlabel("Tanggal")
        ax.set_ylabel("Jumlah transaksi")
        ax.legend()

        theme_pink(ax, pk)
        fig.autofmt_xdate()
        fig.tight_layout()

        return fig

    # ==========================================
    # 2. DATA KATEGORIK
    # ==========================================
    if v in cat_all:

        tab = (
            d[v]
            .value_counts()
            .sort_index()
            .reset_index()
        )

        tab.columns = ["Kategori", "n"]

        tab["pct"] = (
            100 * tab["n"] / tab["n"].sum()
        )

        # --------------------------------------
        # DIAGRAM LINGKARAN
        # --------------------------------------
        if type == "Diagram lingkaran":

            fig, ax = plt.subplots(figsize=(7, 6))

            ax.pie(
                tab["n"],
                labels=None,
                colors=pal_cat(len(tab)),
                startangle=90,
                autopct="%1.1f%%",
                pctdistance=0.65,
                textprops={
                    "color": "white",
                    "fontweight": "bold"
                },
                wedgeprops={
                    "edgecolor": "white",
                    "linewidth": 1
                }
            )

            ax.legend(
                tab["Kategori"],
                title=v,
                loc="center left",
                bbox_to_anchor=(1, 0.5)
            )

            ax.set_title(
                f"Proporsi {v}",
                fontweight="bold",
                color=pk["d"],
                fontsize=14
            )

            fig.patch.set_facecolor("white")
            fig.tight_layout()

            return fig

        # --------------------------------------
        # DIAGRAM PARETO
        # --------------------------------------
        if type == "Pareto":

            tab = tab.sort_values(
                "n",
                ascending=False
            ).reset_index(drop=True)

            tab["kum"] = (
                tab["n"].cumsum()
                / tab["n"].sum() * 100
            )

            fig, ax = plt.subplots(figsize=(10, 5))

            x = np.arange(len(tab))
            total = tab["n"].sum()

            ax.bar(
                x,
                tab["n"],
                color=pk["l"],
                width=0.7
            )

            ax.set_ylabel("Frekuensi")
            ax.set_xlabel(v)

            ax.set_xticks(x)
            ax.set_xticklabels(
                tab["Kategori"],
                rotation=20,
                ha="right"
            )

            # Sumbu kedua untuk persentase kumulatif
            ax2 = ax.twinx()

            ax2.plot(
                x,
                tab["kum"],
                color=pk["d"],
                marker="o",
                linewidth=1.5
            )

            ax2.set_ylabel("Kumulatif (%)")
            ax2.set_ylim(0, 110)

            for i, value in enumerate(tab["kum"]):
                ax2.annotate(
                    f"{value:.1f}%",
                    (i, value),
                    textcoords="offset points",
                    xytext=(0, 8),
                    ha="center",
                    color=pk["mauve"],
                    fontsize=9
                )

            ax.set_title(
                f"Pareto {v}",
                fontweight="bold",
                color=pk["d"]
            )

            theme_pink(ax, pk)
            fig.tight_layout()

            return fig

        # --------------------------------------
        # DIAGRAM BATANG
        # --------------------------------------
        fig, ax = plt.subplots(figsize=(9, 5))

        bars = ax.bar(
            tab["Kategori"].astype(str),
            tab["n"],
            color=pal_cat(len(tab)),
            width=0.7
        )

        ax.bar_label(
            bars,
            padding=3,
            color=pk["mauve"],
            fontweight="bold"
        )

        ax.set_title(
            f"Frekuensi {v}",
            fontweight="bold",
            color=pk["d"]
        )

        ax.set_xlabel(v)
        ax.set_ylabel("Frekuensi")

        ax.tick_params(
            axis="x",
            labelrotation=20
        )

        for label in ax.get_xticklabels():
            label.set_horizontalalignment("right")

        theme_pink(ax, pk)
        fig.tight_layout()

        return fig

    # ==========================================
    # 3. DATA NUMERIK
    # ==========================================

    x = pd.to_numeric(
        d[v],
        errors="coerce"
    ).dropna()

    bw = v in ["Quantity", "Hour"]

    # ------------------------------------------
    # HISTOGRAM
    # ------------------------------------------
    if type in ["Histogram", "Histogram + Density"]:

        fig, ax = plt.subplots(figsize=(9, 5))

        if bw:
            bins = np.arange(
                x.min() - 0.5,
                x.max() + 1.5,
                1
            )
        else:
            bins = 25

        density = type == "Histogram + Density"

        ax.hist(
            x,
            bins=bins,
            density=density,
            color=pk["xl"],
            edgecolor="white"
        )

        if density and len(x) > 1 and x.nunique() > 1:
            kde = stats.gaussian_kde(x)

            xx = np.linspace(
                x.min(),
                x.max(),
                300
            )

            ax.plot(
                xx,
                kde(xx),
                color=pk["d"],
                linewidth=2
            )

        ax.set_title(
            f"{'Histogram dan Density' if density else 'Distribusi'} {v}",
            fontweight="bold",
            color=pk["d"]
        )

        ax.set_xlabel(v)
        ax.set_ylabel(
            "Densitas" if density else "Frekuensi"
        )

        theme_pink(ax, pk)
        fig.tight_layout()

        return fig

    # ------------------------------------------
    # DENSITY
    # ------------------------------------------
    if type == "Density":

        fig, ax = plt.subplots(figsize=(9, 5))

        if len(x) > 1 and x.nunique() > 1:

            kde = stats.gaussian_kde(x)

            xx = np.linspace(
                x.min(),
                x.max(),
                300
            )

            yy = kde(xx)

            ax.fill_between(
                xx,
                yy,
                color=pk["l"],
                alpha=0.5
            )

            ax.plot(
                xx,
                yy,
                color=pk["d"],
                linewidth=1.5
            )

        ax.axvline(
            x.mean(),
            color=pk["mauve"],
            linestyle="--",
            label="Rata-rata"
        )

        ax.set_title(
            f"Density {v}",
            fontweight="bold",
            color=pk["d"]
        )

        ax.set_xlabel(v)
        ax.set_ylabel("Densitas")

        ax.legend()
        theme_pink(ax, pk)

        fig.tight_layout()

        return fig

    # ------------------------------------------
    # BOXPLOT
    # ------------------------------------------
    if type == "Boxplot":

        fig, ax = plt.subplots(figsize=(8, 4))

        ax.boxplot(
            x,
            vert=False,
            patch_artist=True,
            widths=0.45,
            boxprops={
                "facecolor": pk["xl"],
                "edgecolor": pk["d"]
            },
            medianprops={
                "color": pk["d"],
                "linewidth": 2
            },
            whiskerprops={
                "color": pk["d"]
            },
            capprops={
                "color": pk["d"]
            },
            flierprops={
                "markerfacecolor": pk["m"],
                "markeredgecolor": pk["m"]
            }
        )

        ax.scatter(
            x.mean(),
            1,
            marker="D",
            s=55,
            facecolor="white",
            edgecolor=pk["d"],
            label="Rata-rata"
        )

        ax.set_yticks([])
        ax.set_xlabel(v)

        ax.set_title(
            f"Boxplot {v}",
            fontweight="bold",
            color=pk["d"]
        )

        ax.legend()
        theme_pink(ax, pk)

        fig.tight_layout()

        return fig

    # ------------------------------------------
    # QQ-PLOT
    # ------------------------------------------
    if type == "QQ-plot":

        fig, ax = plt.subplots(figsize=(7, 6))

        stats.probplot(
            x,
            dist="norm",
            plot=ax
        )

        ax.get_lines()[0].set_color(pk["m"])
        ax.get_lines()[0].set_alpha(0.6)

        ax.get_lines()[1].set_color(pk["d"])

        ax.set_title(
            f"QQ-plot {v}",
            fontweight="bold",
            color=pk["d"]
        )

        ax.set_xlabel("Kuantil teoretis")
        ax.set_ylabel("Kuantil sampel")

        theme_pink(ax, pk)
        fig.tight_layout()

        return fig

    # ------------------------------------------
    # STEM-AND-LEAF
    # ------------------------------------------
    if type == "Stem-and-leaf":

        fig, ax = plt.subplots(figsize=(8, 5))

        stem_data = np.sort(x.to_numpy())

        stems = np.floor(stem_data).astype(int)
        leaves = np.round(
            (stem_data - stems) * 10
        ).astype(int)

        lines = ["STEM-AND-LEAF", ""]

        for stem in np.unique(stems):
            leaf_values = leaves[stems == stem]

            lines.append(
                f"{stem:>8} | {' '.join(map(str, leaf_values))}"
            )

        ax.text(
            0.05,
            0.95,
            "\n".join(lines),
            transform=ax.transAxes,
            verticalalignment="top",
            family="monospace",
            fontsize=10,
            color=pk["d"]
        )

        ax.axis("off")

        ax.set_title(
            f"Stem-and-leaf {v}",
            fontweight="bold",
            color=pk["d"]
        )

        fig.patch.set_facecolor("white")
        fig.tight_layout()

        return fig

    return None

def uni_table(d, v):

    # ==========================================
    # 1. VARIABEL TANGGAL
    # ==========================================
    if v == "Date":

        dates = pd.to_datetime(d["Date"]).dropna()

        unique_dates = dates.nunique()

        dd = d.groupby("Date").size()

        return pd.DataFrame({
            "Statistik": [
                "Tanggal awal",
                "Tanggal akhir",
                "Jumlah hari",
                "Jumlah hari dengan transaksi",
                "Rata-rata transaksi per hari"
            ],

            "Nilai": [
                str(dates.min().date()),
                str(dates.max().date()),
                str((dates.max() - dates.min()).days + 1),
                str(unique_dates),
                fmt(len(d) / unique_dates)
            ]
        })

    # ==========================================
    # 2. VARIABEL KATEGORIK
    # ==========================================
    if v in cat_all:

        t = d[v].value_counts().sort_index()

        out = pd.DataFrame({
            "Kategori": t.index.astype(str),
            "Frekuensi": t.values,
            "Persentase (%)": [
                fmt(100 * value / t.sum())
                for value in t.values
            ]
        })

        return out

    # ==========================================
    # 3. VARIABEL NUMERIK
    # ==========================================

    x = pd.to_numeric(
        d[v],
        errors="coerce"
    ).dropna()

    n = len(x)

    q1 = x.quantile(0.25)
    q3 = x.quantile(0.75)

    iqr = q3 - q1

    lo = q1 - 1.5 * iqr
    hi = q3 + 1.5 * iqr

    sd = x.std()
    mean = x.mean()

    variansi = x.var()

    cv = (
        100 * sd / mean
        if mean != 0 else np.nan
    )

    jumlah_out = ((x < lo) | (x > hi)).sum()

    statistik = [
        "N",
        "Rata-rata",
        "Median",
        "Simpangan baku",
        "Variansi",
        "Koef. variasi (%)",
        "Minimum",
        "Q1",
        "Q3",
        "Maksimum",
        "IQR",
        "Skewness",
        "Kurtosis (excess)",
        "Batas pencilan bawah",
        "Batas pencilan atas",
        "Jumlah pencilan"
    ]

    nilai = [
        str(n),
        fmt(mean),
        fmt(x.median()),
        fmt(sd),
        fmt(variansi),
        fmt(cv),
        fmt(x.min()),
        fmt(q1),
        fmt(q3),
        fmt(x.max()),
        fmt(iqr),
        fmt(skew_fun(x), 3),
        fmt(kurt_fun(x), 3),
        fmt(lo),
        fmt(hi),
        str(jumlah_out)
    ]

    return pd.DataFrame({
        "Statistik": statistik,
        "Nilai": nilai
    })

def interp_uni(d, v):

    # ==========================================
    # 1. INTERPRETASI VARIABEL TANGGAL
    # ==========================================
    if v == "Date":

        dd = (
            d.groupby("Date")
            .size()
            .reset_index(name="n")
        )

        top = dd.loc[dd["n"].idxmax()]

        tanggal_awal = pd.to_datetime(d["Date"]).min()
        tanggal_akhir = pd.to_datetime(d["Date"]).max()

        return [
            (
                f"Data mencakup **{tanggal_awal:%Y-%m-%d}** "
                f"sampai **{tanggal_akhir:%Y-%m-%d}** "
                f"({len(dd)} hari bertransaksi)."
            ),

            (
                f"Rata-rata **{fmt(dd['n'].mean())}** transaksi "
                f"per hari; puncak pada "
                f"**{top['Date']}** dengan "
                f"**{top['n']}** transaksi."
            ),

            (
                "Garis tren (loess) menunjukkan apakah volume "
                "transaksi harian meningkat, menurun, atau "
                "relatif stabil sepanjang periode."
            )
        ]

    # ==========================================
    # 2. INTERPRETASI VARIABEL KATEGORIK
    # ==========================================
    if v in cat_all:

        t = d[v].value_counts().sort_values(
            ascending=False
        )

        p = 100 * t / t.sum()

        sp = p.max() - p.min()

        return [
            (
                f"Kategori terbanyak: **{t.index[0]}** "
                f"({t.iloc[0]} transaksi; "
                f"{fmt(p.iloc[0], 1)}%). "
                f"Tersedikit: **{t.index[-1]}** "
                f"({t.iloc[-1]}; "
                f"{fmt(p.iloc[-1], 1)}%)."
            ),

            (
                f"Selisih proporsi tertinggi dan terendah "
                f"**{fmt(sp, 1)}** poin persentase, "
                f"sehingga sebaran kategori "
                f"**{'relatif seimbang' if sp < 5 else 'tidak seimbang'}**."
            ),

            (
                "Pareto menunjukkan apakah sedikit kategori "
                "mendominasi (garis kumulatif cepat mencapai 80%)."
            )
        ]

    # ==========================================
    # 3. INTERPRETASI VARIABEL NUMERIK
    # ==========================================

    x = pd.to_numeric(
        d[v],
        errors="coerce"
    ).dropna()

    m = x.mean()
    me = x.median()
    s = x.std()

    sk = skew_fun(x)
    ku = kurt_fun(x)

    q1 = x.quantile(0.25)
    q3 = x.quantile(0.75)

    iqr = q3 - q1

    out = (
        (x < q1 - 1.5 * iqr) |
        (x > q3 + 1.5 * iqr)
    ).sum()

    # Bentuk distribusi
    if pd.isna(sk):
        shape = "tidak dapat ditentukan"
    elif abs(sk) < 0.5:
        shape = "hampir simetris"
    elif sk > 0:
        shape = (
            "miring kuat ke kanan"
            if sk > 1
            else "miring sedang ke kanan"
        )
    else:
        shape = (
            "miring kuat ke kiri"
            if sk < -1
            else "miring sedang ke kiri"
        )

    # Perbandingan rata-rata dan median
    if s == 0 or pd.isna(s):
        perbandingan = "; rata-rata dan median sama"
    elif abs(m - me) / s > 0.1:
        arah = "kanan" if m > me else "kiri"

        perbandingan = (
            f"; rata-rata {'lebih besar' if m > me else 'lebih kecil'} "
            f"dari median sehingga sebaran condong ke {arah}"
        )
    else:
        perbandingan = (
            "; keduanya berdekatan (sebaran seimbang)"
        )

    # Interpretasi kurtosis
    if pd.isna(ku):
        kurt_desc = "tidak dapat ditentukan"
    elif ku < -1:
        kurt_desc = "datar (mendekati sebaran seragam)"
    elif ku > 1:
        kurt_desc = "runcing"
    else:
        kurt_desc = "mendekati normal"

    # Hasil interpretasi
    hasil = [
        (
            f"Rata-rata **{fmt(m)}** dan median **{fmt(me)}**"
            f"{perbandingan}."
        ),

        (
            f"Sebaran data: simpangan baku **{fmt(s)}**, "
            f"rentang **{fmt(x.min())}** sampai "
            f"**{fmt(x.max())}**, koefisien variasi "
            f"**{fmt(100 * s / m if m != 0 else np.nan)}%**."
        ),

        (
            f"Bentuk distribusi **{shape}** "
            f"(skewness {fmt(sk, 2)}); "
            f"kurtosis {fmt(ku, 2)} menunjukkan "
            f"puncak yang **{kurt_desc}**."
        )
    ]

    if out == 0:
        hasil.append(
            "Tidak ditemukan pencilan menurut aturan 1.5 x IQR."
        )
    else:
        hasil.append(
            f"Terdapat **{out} pencilan** menurut aturan "
            f"1.5 x IQR ({fmt(100 * out / len(x), 1)}% data). "
            "Pencilan dipertahankan karena merupakan transaksi "
            "yang valid."
        )

    return hasil

# =============================================================================
# 3. BIVARIAT
# =============================================================================
def bi_choices(x, y):

    xc = x in cat_all
    yc = y in cat_all

    if not xc and not yc:
        return [
            "Scatter plot + garis tren",
            "Scatter plot + kurva linear, kuadratik, nonparametrik",
            "Line plot rata-rata Y per X",
            "Correlation plot",
            "Heatmap kepadatan 2D"
        ]

    elif xc != yc:
        return [
            "Grouped boxplot",
            "Grouped violin plot",
            "Grouped kernel density plot",
            "Grouped ridgeline plot",
            "Grouped strip plot",
            "Grouped mean/SEM plot",
            "Bar chart statistik ringkasan",
            "Bar chart rata-rata (+/- SE)",
            "Cleveland dot chart"
        ]

    else:
        return [
            "Stacked bar chart",
            "Grouped bar chart",
            "Segmented bar chart (100%)",
            "Heatmap tabulasi silang"
        ]

# Correlation plot: lingkaran (segitiga bawah) + angka r (segitiga atas)
def plot_corr_circle(d, vars=None):

    if vars is None:
        vars = num_all

    m = d[vars].corr(method="pearson")

    fig, ax = plt.subplots(
        figsize=(9, 7)
    )

    n = len(vars)

    for i in range(n):
        for j in range(n):

            r = m.iloc[i, j]

            if i > j:

                color = (
                    pk["s"] if r > 0
                    else pk["slate"]
                )

                circle = plt.Circle(
                    (j, i),
                    radius=0.12 + abs(r) * 0.27,
                    color=color,
                    alpha=0.8
                )

                ax.add_patch(circle)

            elif i < j:

                ax.text(
                    j,
                    i,
                    f"{r:.2f}",
                    ha="center",
                    va="center",
                    color=pk["txt"],
                    fontweight="bold"
                )

            else:

                ax.text(
                    j,
                    i,
                    vars[i],
                    ha="center",
                    va="center",
                    color=pk["d"],
                    fontweight="bold"
                )

    ax.set_xticks(range(n))
    ax.set_yticks(range(n))

    ax.set_xticklabels(vars)
    ax.set_yticklabels(vars)

    ax.invert_yaxis()

    ax.set_xlim(-0.5, n - 0.5)
    ax.set_ylim(n - 0.5, -0.5)

    ax.set_aspect("equal")

    ax.set_title(
        "Correlation plot (Pearson)",
        fontweight="bold",
        color=pk["d"]
    )

    ax.set_xlabel("")
    ax.set_ylabel("")

    ax.grid(False)

    fig.tight_layout()

    return fig


import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

from scipy.stats import gaussian_kde
from statsmodels.nonparametric.smoothers_lowess import lowess
import statsmodels.api as sm
from sklearn.preprocessing import PolynomialFeatures
from sklearn.linear_model import LinearRegression


def plot_bi(d, x, y, type):

    # Menentukan tipe variabel
    xc = x in cat_all
    yc = y in cat_all

    # ---------- NUMERIK X NUMERIK ----------
    if not xc and not yc:

        # 1. Correlation plot
        if type == "Correlation plot":
            return plot_corr_circle(d, num_all)

        # 2. Scatter plot + kurva linear, kuadratik, nonparametrik
        if type == "Scatter plot + kurva linear, kuadratik, nonparametrik":

            fig, ax = plt.subplots(figsize=(9, 6))

            xv = d[x].to_numpy()
            yv = d[y].to_numpy()

            # Scatter plot
            ax.scatter(
                xv, yv,
                color=pk["l"],
                alpha=0.5,
                s=30,
                label="Data"
            )

            # Mengurutkan X agar garis kurva tidak berantakan
            order = np.argsort(xv)
            xs = xv[order]

            # Kurva linear
            linear = np.polyfit(xv, yv, 1)
            ax.plot(
                xs,
                np.polyval(linear, xs),
                color=pk["d"],
                linewidth=2,
                label="Linear"
            )

            # Kurva kuadratik
            quadratic = np.polyfit(xv, yv, 2)
            ax.plot(
                xs,
                np.polyval(quadratic, xs),
                color=pk["slate"],
                linewidth=2,
                label="Kuadratik"
            )

            # Kurva nonparametrik (LOESS)
            loess = lowess(yv, xv, frac=0.6, return_sorted=True)

            ax.plot(
                loess[:, 0],
                loess[:, 1],
                color="#2E7D32",
                linewidth=2,
                label="Nonparametrik (LOESS)"
            )

            ax.set_title(f"{y} terhadap {x}")
            ax.set_xlabel(x)
            ax.set_ylabel(y)

            ax.legend(title="Kurva")
            ax.grid(alpha=0.2)

            return fig

        # 3. Line plot rata-rata Y per X
        if type == "Line plot rata-rata Y per X":

            xv = d[x]
            yv = d[y]

            # Pengelompokan data
            if xv.nunique() <= 12:
                grp = xv
            else:
                grp = pd.cut(xv, bins=10)

            temp = pd.DataFrame({
                "gx": xv,
                "yv": yv,
                "grp": grp
            })

            agg = temp.groupby(
                "grp",
                observed=True
            ).agg(
                gx=("gx", "mean"),
                my=("yv", "mean"),
                sd=("yv", "std"),
                n=("yv", "count")
            ).reset_index(drop=True)

            # Standard error
            agg["se"] = agg["sd"] / np.sqrt(agg["n"])

            fig, ax = plt.subplots(figsize=(9, 6))

            ax.fill_between(
                agg["gx"],
                agg["my"] - agg["se"],
                agg["my"] + agg["se"],
                color=pk["xl"],
                alpha=0.8
            )

            ax.plot(
                agg["gx"],
                agg["my"],
                color=pk["d"],
                linewidth=2
            )

            ax.scatter(
                agg["gx"],
                agg["my"],
                color=pk["m"],
                s=50
            )

            ax.set_title(f"Rata-rata {y} menurut {x}")
            ax.set_xlabel(x)
            ax.set_ylabel(f"Rata-rata {y}")

            ax.grid(alpha=0.2)

            return fig

        # 4. Heatmap kepadatan 2D
        if type == "Heatmap kepadatan 2D":

            fig, ax = plt.subplots(figsize=(9, 6))

            sns.histplot(
                data=d,
                x=x,
                y=y,
                bins=20,
                cmap="pink",
                cbar=True,
                ax=ax
            )

            ax.set_title(f"Kepadatan {x} dan {y}")
            ax.set_xlabel(x)
            ax.set_ylabel(y)

            return fig

        # 5. Scatter plot + garis regresi linear
        fig, ax = plt.subplots(figsize=(9, 6))

        sns.regplot(
            data=d,
            x=x,
            y=y,
            scatter_kws={
                "color": pk["m"],
                "alpha": 0.5,
                "s": 30
            },
            line_kws={
                "color": pk["d"],
                "linewidth": 2
            },
            ci=None,
            ax=ax
        )

        ax.set_title(f"{y} terhadap {x}")
        ax.set_xlabel(x)
        ax.set_ylabel(y)

        ax.grid(alpha=0.2)

        return fig

  # ---------- kategorik x numerik ----------
  
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

from scipy.stats import gaussian_kde


def plot_bi(d, x, y, type):

    # Menentukan tipe variabel
    xc = x in cat_all
    yc = y in cat_all

    # ---------- KATEGORIK X NUMERIK ----------
    if xc != yc:

        # Menentukan variabel kategorik dan numerik
        cv = x if xc else y
        nv = y if xc else x

        # Mengambil data
        d2 = pd.DataFrame({
            "g": d[cv],
            "v": pd.to_numeric(d[nv], errors="coerce")
        }).dropna()

        # Mempertahankan urutan kategori
        if isinstance(d[cv].dtype, pd.CategoricalDtype):
            levels = list(d[cv].cat.categories)
        else:
            levels = list(d2["g"].drop_duplicates())

        d2["g"] = pd.Categorical(
            d2["g"],
            categories=levels,
            ordered=True
        )

        # Palet warna
        cols = pal_cat(len(levels))
        color_map = dict(zip(levels, cols))

        # Statistik ringkasan
        agg = d2.groupby(
            "g",
            observed=False
        )["v"].agg(
            m="mean",
            md="median",
            s="std",
            mn="min",
            mx="max",
            n="count"
        ).reset_index()

        agg["se"] = agg["s"] / np.sqrt(agg["n"])

        # Fungsi pengaturan grafik
        def finish(ax, title, subtitle=None):
            ax.set_title(title)
            if subtitle:
                ax.set_xlabel(cv)
                ax.set_ylabel(nv)
            ax.grid(alpha=0.2)

        # 1. Grouped violin plot
        if type == "Grouped violin plot":

            fig, ax = plt.subplots(figsize=(9, 6))

            sns.violinplot(
                data=d2,
                x="g",
                y="v",
                hue="g",
                palette=color_map,
                inner=None,
                alpha=0.75,
                legend=False,
                ax=ax
            )

            sns.boxplot(
                data=d2,
                x="g",
                y="v",
                width=0.12,
                color="white",
                showfliers=False,
                ax=ax
            )

            ax.set_title(
                f"Grouped violin plot: {nv} menurut {cv}"
            )
            ax.set_xlabel(cv)
            ax.set_ylabel(nv)
            ax.grid(axis="y", alpha=0.2)
            plt.xticks(rotation=35, ha="right")

            return fig

        # 2. Grouped kernel density plot
        if type == "Grouped kernel density plot":

            fig, ax = plt.subplots(figsize=(9, 6))

            for g in levels:
                values = d2.loc[d2["g"] == g, "v"].dropna()

                if len(values) < 2 or values.nunique() < 2:
                    continue

                sns.kdeplot(
                    x=values,
                    fill=True,
                    alpha=0.25,
                    color=color_map[g],
                    linewidth=1.5,
                    label=str(g),
                    ax=ax
                )

            ax.set_title(
                f"Grouped kernel density plot: {nv} per {cv}"
            )
            ax.set_xlabel(nv)
            ax.set_ylabel("Densitas")
            ax.legend(title=cv)
            ax.grid(alpha=0.2)

            return fig

        # 3. Grouped ridgeline plot
        if type == "Grouped ridgeline plot":

            fig, ax = plt.subplots(figsize=(9, 6))

            valid = d2["v"].dropna()
            xmin, xmax = valid.min(), valid.max()

            if xmin == xmax:
                xmin -= 0.5
                xmax += 0.5

            xs = np.linspace(xmin, xmax, 256)

            # Menghitung densitas setiap kategori
            densities = {}

            for g in levels:
                values = d2.loc[d2["g"] == g, "v"].dropna()

                if len(values) >= 2 and values.nunique() > 1:
                    kde = gaussian_kde(values)
                    densities[g] = kde(xs)
                else:
                    densities[g] = np.zeros(len(xs))

            max_density = max(
                [z.max() for z in densities.values()],
                default=0
            )

            # Membuat ridgeline
            for i, g in enumerate(reversed(levels)):

                density = densities[g]

                if max_density > 0:
                    height = density / max_density * 1.6
                else:
                    height = density

                base = i + 1

                ax.fill_between(
                    xs,
                    base,
                    base + height,
                    color=color_map[g],
                    alpha=0.8,
                    linewidth=1
                )

                ax.plot(
                    xs,
                    base + height,
                    color=pk["mauve"],
                    linewidth=0.8
                )

            ax.set_yticks(range(1, len(levels) + 1))
            ax.set_yticklabels(list(reversed(levels)))

            ax.set_title(
                f"Grouped ridgeline plot: {nv} per {cv}"
            )
            ax.set_xlabel(nv)
            ax.set_ylabel(cv)
            ax.grid(axis="x", alpha=0.2)

            return fig

        # 4. Grouped strip plot
        if type == "Grouped strip plot":

            fig, ax = plt.subplots(figsize=(9, 6))

            sns.stripplot(
                data=d2,
                x="g",
                y="v",
                hue="g",
                palette=color_map,
                jitter=0.22,
                alpha=0.45,
                size=4,
                legend=False,
                ax=ax
            )

            # Menambahkan titik rata-rata berbentuk berlian
            sns.pointplot(
                data=d2,
                x="g",
                y="v",
                color=pk["d"],
                estimator="mean",
                errorbar=None,
                markers="D",
                markersize=8,
                linestyle="none",
                ax=ax
            )

            ax.set_title(
                f"Grouped strip plot: {nv} menurut {cv}"
            )
            ax.set_xlabel(cv)
            ax.set_ylabel(nv)
            ax.text(
                0.5, 1.02,
                "Titik = observasi; berlian = rata-rata",
                transform=ax.transAxes,
                ha="center",
                fontsize=9
            )

            ax.grid(axis="y", alpha=0.2)
            plt.xticks(rotation=35, ha="right")

            return fig

        # 5. Grouped mean/SEM plot
        if type == "Grouped mean/SEM plot":

            fig, ax = plt.subplots(figsize=(9, 6))

            positions = np.arange(len(agg))

            for i, row in agg.iterrows():

                ax.errorbar(
                    i,
                    row["m"],
                    yerr=row["se"] if pd.notna(row["se"]) else 0,
                    fmt="o",
                    color=color_map[row["g"]],
                    markersize=8,
                    capsize=5,
                    linewidth=1.5
                )

            ax.set_xticks(positions)
            ax.set_xticklabels(
                agg["g"].astype(str),
                rotation=35,
                ha="right"
            )

            ax.set_title(
                f"Grouped mean/SEM plot: {nv} menurut {cv}"
            )
            ax.set_xlabel(cv)
            ax.set_ylabel(f"Rata-rata {nv}")

            ax.text(
                0.5, 1.02,
                "Titik = rata-rata; garis = +/- 1 standard error (SEM)",
                transform=ax.transAxes,
                ha="center",
                fontsize=9
            )

            ax.grid(axis="y", alpha=0.2)

            return fig

        # 6. Bar chart statistik ringkasan
        if type == "Bar chart statistik ringkasan":

            stats = [
                ("Rata-rata", "m"),
                ("Median", "md"),
                ("Simpangan baku", "s"),
                ("Minimum", "mn"),
                ("Maksimum", "mx")
            ]

            fig, axes = plt.subplots(
                1, 5,
                figsize=(17, 6)
            )

            for ax, (label, col) in zip(axes, stats):

                values = agg[col]

                ax.bar(
                    agg["g"].astype(str),
                    values,
                    color=cols,
                    width=0.7
                )

                for i, val in enumerate(values):
                    if pd.notna(val):
                        ax.text(
                            i,
                            val,
                            f"{val:.1f}",
                            ha="center",
                            va="bottom",
                            fontsize=8,
                            color=pk["txt"]
                        )

                ax.set_title(label, fontsize=10)
                ax.tick_params(axis="x", rotation=45)
                ax.grid(axis="y", alpha=0.2)

            fig.suptitle(
                f"Bar chart statistik ringkasan {nv} menurut {cv}"
            )
            fig.tight_layout()

            return fig

        # 7. Bar chart rata-rata (+/- SE)
        if type == "Bar chart rata-rata (+/- SE)":

            fig, ax = plt.subplots(figsize=(9, 6))

            positions = np.arange(len(agg))

            ax.bar(
                positions,
                agg["m"],
                color=cols,
                width=0.7
            )

            ax.errorbar(
                positions,
                agg["m"],
                yerr=agg["se"].fillna(0),
                fmt="none",
                color=pk["txt"],
                capsize=5
            )

            for i, val in enumerate(agg["m"]):
                if pd.notna(val):
                    ax.text(
                        i,
                        val,
                        f"{val:.2f}",
                        ha="center",
                        va="top",
                        color="white",
                        fontweight="bold"
                    )

            ax.set_xticks(positions)
            ax.set_xticklabels(
                agg["g"].astype(str),
                rotation=35,
                ha="right"
            )

            ax.set_title(
                f"Rata-rata {nv} menurut {cv}"
            )
            ax.set_xlabel(cv)
            ax.set_ylabel(f"Rata-rata {nv}")

            ax.text(
                0.5, 1.02,
                "Garis = +/- 1 standard error",
                transform=ax.transAxes,
                ha="center",
                fontsize=9
            )

            ax.grid(axis="y", alpha=0.2)

            return fig

        # 8. Cleveland dot chart
        if type == "Cleveland dot chart":

            a2 = agg.sort_values("m").reset_index(drop=True)

            lv = a2["g"].astype(str).tolist()

            cl = pd.DataFrame({
                "g": lv * 2,
                "stat": (
                    ["Rata-rata"] * len(a2)
                    + ["Median"] * len(a2)
                ),
                "val": list(a2["m"]) + list(a2["md"])
            })

            values = cl["val"].dropna()

            pad = 0.15 * max(
                values.max() - values.min(),
                abs(values.mean()) * 0.05,
                1e-9
            )

            lo = values.min() - pad
            hi = values.max() + pad

            fig, ax = plt.subplots(figsize=(9, 6))

            for i, row in a2.iterrows():

                ax.hlines(
                    y=i,
                    xmin=lo,
                    xmax=hi,
                    color=pk["grid"],
                    linestyle="dotted",
                    linewidth=1
                )

                ax.scatter(
                    row["m"],
                    i,
                    s=90,
                    color=pk["m"],
                    edgecolor=pk["d"],
                    label="Rata-rata" if i == 0 else None
                )

                ax.scatter(
                    row["md"],
                    i,
                    s=90,
                    color="white",
                    edgecolor=pk["d"],
                    label="Median" if i == 0 else None
                )

            ax.set_yticks(range(len(lv)))
            ax.set_yticklabels(lv)
            ax.set_xlim(lo, hi)

            ax.set_title(
                f"Cleveland dot chart: {nv} menurut {cv}"
            )

            ax.set_xlabel(nv)
            ax.set_ylabel(cv)

            ax.legend(
                loc="upper center",
                bbox_to_anchor=(0.5, -0.12),
                ncol=2
            )

            ax.grid(axis="x", alpha=0.2)

            return fig

        # 9. Grouped boxplot (default)
        fig, ax = plt.subplots(figsize=(9, 6))

        sns.boxplot(
            data=d2,
            x="g",
            y="v",
            hue="g",
            palette=color_map,
            showfliers=True,
            ax=ax
        )

        # Menambahkan rata-rata berbentuk berlian
        means = agg["m"]

        ax.scatter(
            np.arange(len(agg)),
            means,
            marker="D",
            s=55,
            color="white",
            edgecolor=pk["d"],
            zorder=5,
            label="Rata-rata"
        )

        ax.set_title(
            f"Grouped boxplot: {nv} menurut {cv}"
        )
        ax.set_xlabel(cv)
        ax.set_ylabel(nv)

        ax.text(
            0.5, 1.02,
            "Berlian = rata-rata",
            transform=ax.transAxes,
            ha="center",
            fontsize=9
        )

        ax.legend()
        ax.grid(axis="y", alpha=0.2)
        plt.xticks(rotation=35, ha="right")

        return fig

  # ---------- kategorik x kategorik ----------
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

from scipy import stats
from scipy.stats import chi2_contingency, kruskal, ttest_ind
from statsmodels.formula.api import ols
from statsmodels.stats.anova import anova_lm
from statsmodels.stats.multicomp import pairwise_tukeyhsd
from sklearn.linear_model import LinearRegression


def bi_chart(d, x, y, type, pk, rot=None):
    ny = d[y].nunique()

    df = (
        d.groupby([x, y], observed=False)
        .size()
        .reset_index(name="n")
    )

    df["pct"] = df["n"] / df.groupby(x)["n"].transform("sum") * 100

    palette = pal_cat(ny)

    if type == "Heatmap tabulasi silang":
        tab = pd.crosstab(d[x], d[y])
        mid = tab.values.max() / 2

        fig, ax = plt.subplots(figsize=(9, 6))

        sns.heatmap(
            tab,
            annot=True,
            fmt="d",
            linewidths=1,
            linecolor="white",
            cmap=sns.light_palette(pk["d"], as_cmap=True),
            cbar_kws={"label": "Jumlah"},
            ax=ax
        )

        for text in ax.texts:
            value = int(text.get_text())
            text.set_color("white" if value > mid else pk["txt"])
            text.set_fontweight("bold")

        ax.set_title(f"{x} dan {y}")
        ax.set_xlabel(x)
        ax.set_ylabel(y)

        if rot is not None:
            plt.xticks(rotation=rot)

        plt.tight_layout()
        return fig

    if type == "Stacked bar chart":
        pivot = df.pivot(index=x, columns=y, values="n").fillna(0)

        ax = pivot.plot(
            kind="bar",
            stacked=True,
            color=palette,
            edgecolor="white",
            figsize=(9, 6)
        )

        ax.set_title(f"Stacked bar chart: {y} dalam {x}")
        ax.set_xlabel(x)
        ax.set_ylabel("Jumlah transaksi")
        ax.legend(title=y)

        ax.text(
            0.5, 1.02,
            "Tinggi total batang = jumlah transaksi per kategori",
            transform=ax.transAxes,
            ha="center"
        )

        if rot is not None:
            plt.xticks(rotation=rot)

        plt.tight_layout()
        return ax.figure

    if type == "Segmented bar chart (100%)":
        pivot = df.pivot(index=x, columns=y, values="pct").fillna(0)

        ax = pivot.plot(
            kind="bar",
            stacked=True,
            color=palette,
            edgecolor="white",
            figsize=(9, 6)
        )

        ax.set_title(f"Segmented bar chart: {y} dalam {x}")
        ax.set_xlabel(x)
        ax.set_ylabel("Persentase")
        ax.set_ylim(0, 102)
        ax.set_yticks(np.arange(0, 101, 10))
        ax.set_yticklabels([f"{z}%" for z in range(0, 101, 10)])
        ax.legend(title=y)

        ax.text(
            0.5, 1.02,
            "Setiap batang = 100%; bandingkan komposisi antar kategori",
            transform=ax.transAxes,
            ha="center"
        )

        if rot is not None:
            plt.xticks(rotation=rot)

        plt.tight_layout()
        return ax.figure

    if type == "Grouped bar chart":
        fig, ax = plt.subplots(figsize=(9, 6))

        sns.barplot(
            data=df,
            x=x,
            y="n",
            hue=y,
            palette=palette,
            errorbar=None,
            ax=ax
        )

        for container in ax.containers:
            ax.bar_label(
                container,
                fmt="%.0f",
                padding=3,
                fontsize=9,
                color=pk["txt"]
            )

        ax.set_title(f"Grouped bar chart: {x} dan {y}")
        ax.set_xlabel(x)
        ax.set_ylabel("Jumlah transaksi")
        ax.legend(title=y)

        if rot is not None:
            plt.xticks(rotation=rot)

        plt.tight_layout()
        return fig

    return None


def bi_text(d, x, y):

    cat_all = globals().get("cat_all", [])

    xc = x in cat_all
    yc = y in cat_all

    if not xc and not yc:

        ct = stats.pearsonr(d[x], d[y])
        sp = stats.spearmanr(d[x], d[y])

        X = d[[x]]
        Y = d[y]

        model = LinearRegression()
        model.fit(X, Y)

        b = model.coef_[0]
        r = ct.statistic
        r2 = model.score(X, Y)

        tab = pd.DataFrame({
            "Ukuran": [
                "Korelasi Pearson (r)",
                "p-value Pearson",
                "Korelasi Spearman (rho)",
                "p-value Spearman",
                "Kemiringan garis (b1)",
                "R-kuadrat regresi sederhana"
            ],
            "Nilai": [
                fmt(r, 3),
                fmt_p(ct.pvalue),
                fmt(sp.statistic, 3),
                fmt_p(sp.pvalue),
                fmt(b, 3),
                fmt(r2, 3)
            ]
        })

        lines = [
            f"Hubungan linear **{strength(r)}** dan "
            f"**{'positif' if r > 0 else 'negatif'}** "
            f"(r = {fmt(r, 3)}; "
            f"p {'< 0.001' if ct.pvalue < 0.001 else '= ' + fmt_p(ct.pvalue)}).",

            (
                f"Hubungan **signifikan** pada taraf 5%: setiap kenaikan "
                f"1 satuan **{x}**, rata-rata **{y}** berubah sebesar "
                f"**{fmt(b, 3)}** satuan."
                if ct.pvalue < 0.05 else
                "Hubungan **tidak signifikan** pada taraf 5%, "
                "sehingga tidak ada bukti hubungan linear antar kedua variabel."
            ),

            "Korelasi menunjukkan keterkaitan, bukan sebab-akibat."
        ]

        return {"table": tab, "lines": lines}

    if xc != yc:

        cv = x if xc else y
        nv = y if xc else x

        g = (
            d.groupby(cv, observed=False)[nv]
            .agg(["mean", "median", "std", "count"])
            .reset_index()
        )

        g.columns = [
            cv,
            "Rata-rata",
            "Median",
            "Simpangan baku",
            "n"
        ]

        g["Rata-rata"] = g["Rata-rata"].apply(fmt)
        g["Median"] = g["Median"].apply(fmt)
        g["Simpangan baku"] = g["Simpangan baku"].apply(fmt)

        gm = d.groupby(cv, observed=False)[nv].mean()

        groups = [
            group[nv].dropna().values
            for _, group in d.groupby(cv, observed=False)
        ]

        groups = [group for group in groups if len(group) > 0]

        model = ols(f'Q("{nv}") ~ C(Q("{cv}"))', data=d).fit()
        a = anova_lm(model, typ=2)

        p = a["PR(>F)"].iloc[0]

        eta = (
            a["sum_sq"].iloc[0] /
            a["sum_sq"].sum()
        )

        kw = kruskal(*groups).pvalue

        if d[cv].nunique() == 2:

            tt = stats.ttest_ind(
                *groups,
                equal_var=False
            )

            p = tt.pvalue

            test = f"uji t Welch (t = {fmt(tt.statistic, 2)})"

        else:

            F = a["F"].iloc[0]

            test = f"ANOVA (F = {fmt(F, 2)})"

        highest = gm.idxmax()
        lowest = gm.idxmin()

        lines = [
            f"Rata-rata **{nv}** tertinggi pada "
            f"**{highest}** ({fmt(gm.max())}) dan terendah pada "
            f"**{lowest}** ({fmt(gm.min())}).",

            f"Hasil {test}: p {pq(p)}; "
            f"uji Kruskal-Wallis: p {pq(kw)}; "
            f"ukuran efek eta-kuadrat = {fmt(eta, 3)}.",

            (
                "Perbedaan rata-rata antar kelompok **signifikan** pada taraf 5%."
                if p < 0.05 else
                "Perbedaan rata-rata antar kelompok **tidak signifikan** "
                "pada taraf 5%; selisih yang tampak pada grafik dapat terjadi "
                "karena kebetulan."
            )
        ]

        return {"table": g, "lines": lines}

    tab = pd.crosstab(d[x], d[y])

    ch2, p_value, dof, expected = chi2_contingency(tab)

    n = tab.values.sum()

    V = np.sqrt(
        ch2 / (n * (min(tab.shape) - 1))
    )

    ix = np.unravel_index(
        np.argmax(tab.values),
        tab.shape
    )

    out = tab.reset_index()

    out.columns.name = None

    lines = [
        f"Kombinasi terbanyak: **{tab.index[ix[0]]}** x "
        f"**{tab.columns[ix[1]]}** "
        f"({int(tab.iloc[ix[0], ix[1]])} transaksi).",

        f"Uji Chi-square: X2 = {fmt(ch2, 2)}, "
        f"df = {dof}, p {pq(p_value)}; "
        f"Cramer's V = {fmt(V, 3)}.",

        (
            "Ada hubungan **signifikan** antar kedua variabel kategorik pada taraf 5%."
            if p_value < 0.05 else
            "**Tidak ada** hubungan signifikan antar kedua variabel kategorik pada taraf 5%."
        )
    ]

    return {"table": out, "lines": lines}
# =============================================================================
# 4. MULTIVARIAT
# =============================================================================

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

from scipy import stats
from scipy.stats import chi2
from scipy.spatial.distance import mahalanobis
from sklearn.decomposition import PCA
from statsmodels.nonparametric.smoothers_lowess import lowess


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def _palette(n):
    if "pal_cat" in globals():
        return pal_cat(n)
    return sns.color_palette("husl", n_colors=max(n, 1))


def _none(value):
    return value is None or value == NONE if "NONE" in globals() else value is None or value == "NONE"


def _format(value, digits=1):
    if "fmt" in globals():
        return fmt(value, digits)
    return f"{value:.{digits}f}".replace(".", ",")


def _theme(ax, size=11):
    p = globals().get("pk", {})
    bg = p.get("bg", "#FFF5F8")
    txt = p.get("txt", "#59434D")

    ax.set_facecolor(bg)
    ax.tick_params(colors=txt, labelsize=size)
    ax.xaxis.label.set_color(txt)
    ax.yaxis.label.set_color(txt)
    ax.title.set_color(txt)
    ax.title.set_fontweight("bold")

    for spine in ax.spines.values():
        spine.set_color("#E8D5DE")

    ax.grid(False)


def _get_pk():
    return globals().get("pk", {
        "bg": "#FFF5F8",
        "d": "#C45B83",
        "m": "#D98BA9",
        "s": "#B65C83",
        "slate": "#8B9BAA",
        "xl": "#F3C6D5",
        "txt": "#59434D",
        "l": "#E8B7CA"
    })


# ============================================================
# MULTIVARIATE SCATTER PLOT
# ============================================================

def plot_mv_scatter(d, x, y, col, size, facet, lines):

    pk = _get_pk()

    hasc = not _none(col)
    hass = not _none(size)

    fig, ax = plt.subplots(figsize=(9, 6))

    if hasc:
        if pd.api.types.is_numeric_dtype(d[col]):
            sc = ax.scatter(
                d[x], d[y],
                c=d[col],
                cmap="pink",
                alpha=0.6,
                s=40 if not hass else None
            )

            if hass:
                sc.set_sizes(
                    20 + 100 * (
                        (d[size] - d[size].min()) /
                        (d[size].max() - d[size].min() + 1e-10)
                    )
                )

            fig.colorbar(sc, ax=ax, label=col)

        else:
            categories = d[col].dropna().unique()
            palette = _palette(len(categories))

            for category, color in zip(categories, palette):
                subset = d[d[col] == category]

                ax.scatter(
                    subset[x],
                    subset[y],
                    color=color,
                    alpha=0.6,
                    s=40 if not hass else (
                        20 + 100 * (
                            (subset[size] - d[size].min()) /
                            (d[size].max() - d[size].min() + 1e-10)
                        )
                    ),
                    label=str(category)
                )

            ax.legend(title=col)

    else:
        if hass:
            sc = ax.scatter(
                d[x], d[y],
                c=pk["m"],
                s=20 + 100 * (
                    (d[size] - d[size].min()) /
                    (d[size].max() - d[size].min() + 1e-10)
                ),
                alpha=0.5
            )
        else:
            ax.scatter(
                d[x], d[y],
                color=pk["m"],
                alpha=0.6,
                s=30
            )

    if hass and not hasc:
        ax.set_title(f"Ukuran titik: {size}")

    if lines:
        if hasc and not pd.api.types.is_numeric_dtype(d[col]):
            for category, color in zip(
                d[col].dropna().unique(),
                _palette(d[col].nunique())
            ):
                subset = d[d[col] == category].dropna(subset=[x, y])

                if len(subset) >= 2:
                    sns.regplot(
                        data=subset,
                        x=x,
                        y=y,
                        scatter=False,
                        ci=None,
                        color=color,
                        line_kws={"linewidth": 1}
                    )
        else:
            sns.regplot(
                data=d,
                x=x,
                y=y,
                scatter=False,
                ci=None,
                color=pk["d"],
                line_kws={"linewidth": 1.5},
                ax=ax
            )

    if not _none(facet):
        plt.close(fig)

        g = sns.relplot(
            data=d,
            x=x,
            y=y,
            hue=col if hasc else None,
            size=size if hass else None,
            col=facet,
            kind="scatter",
            alpha=0.6,
            height=4,
            aspect=1.2
        )

        if lines:
            g.map_dataframe(
                sns.regplot,
                x=x,
                y=y,
                scatter=False,
                ci=None,
                color=pk["d"]
            )

        g.set_axis_labels(x, y)
        g.fig.suptitle(
            f"{y} terhadap {x}",
            y=1.04,
            fontweight="bold"
        )

        return g.fig

    ax.set_title(f"{y} terhadap {x}")
    ax.set_xlabel(x)
    ax.set_ylabel(y)

    ax.text(
        0.5, 1.02,
        f"Warna: {col} | Ukuran: {size} | Panel: {facet}",
        transform=ax.transAxes,
        ha="center",
        fontsize=9
    )

    _theme(ax)

    plt.tight_layout()
    return fig


# ============================================================
# MULTIVARIATE HEATMAP DATA
# ============================================================

def mv_heat_data(d, a, b, num, fun):

    if fun == "Rata-rata":
        v = d[num]
        f = "mean"

    elif fun == "Median":
        v = d[num]
        f = "median"

    elif fun == "Jumlah":
        v = d[num]
        f = "sum"

    elif fun == "Banyak transaksi":
        v = d[a]
        f = "count"

    else:
        raise ValueError("Fungsi agregasi tidak dikenal.")

    temp = d.copy()
    temp["_value"] = v

    res = (
        temp.groupby([a, b], observed=False)["_value"]
        .agg(f)
        .reset_index()
    )

    res.columns = ["A", "B", "val"]

    return res


# ============================================================
# MULTIVARIATE HEATMAP
# ============================================================

def plot_mv_heat(d, a, b, num, fun):

    pk = _get_pk()

    res = mv_heat_data(d, a, b, num, fun)

    mid = (res["val"].max() + res["val"].min()) / 2

    lab = (
        "Banyak transaksi"
        if fun == "Banyak transaksi"
        else f"{fun} {num}"
    )

    tab = res.pivot(
        index="B",
        columns="A",
        values="val"
    )

    fig, ax = plt.subplots(figsize=(9, 6))

    sns.heatmap(
        tab,
        annot=True,
        fmt=".0f" if fun == "Banyak transaksi" else ".1f",
        linewidths=1,
        linecolor="white",
        cmap=sns.light_palette(pk["d"], as_cmap=True),
        cbar_kws={"label": lab},
        ax=ax
    )

    for text in ax.texts:
        try:
            value = float(text.get_text())
            text.set_color(
                "white" if value > mid else pk["txt"]
            )
            text.set_fontweight("bold")
        except ValueError:
            pass

    ax.set_title(lab)
    ax.set_xlabel(a)
    ax.set_ylabel(b)

    fig.text(
        0.5, 0.94,
        f"menurut {a} dan {b}",
        ha="center",
        fontsize=9
    )

    _theme(ax)
    plt.tight_layout()

    return fig


# ============================================================
# CORRELATION MATRIX
# ============================================================

def plot_mv_cor(d, vars, method):

    pk = _get_pk()

    m = d[vars].corr(method=method)

    fig, ax = plt.subplots(figsize=(9, 7))

    sns.heatmap(
        m,
        annot=True,
        fmt=".2f",
        linewidths=1,
        linecolor="white",
        cmap=sns.diverging_palette(
            240, 340, as_cmap=True
        ),
        center=0,
        vmin=-1,
        vmax=1,
        cbar_kws={"label": "r"},
        ax=ax
    )

    for text in ax.texts:
        text.set_color(pk["txt"])
        text.set_fontweight("bold")
        text.set_fontsize(11)

    ax.set_title(
        f"Matriks korelasi {method.title()}",
        fontweight="bold"
    )

    ax.set_xlabel("")
    ax.set_ylabel("")

    _theme(ax)
    plt.tight_layout()

    return fig


# ============================================================
# TREND DATA
# ============================================================

def trend_data(d, metric, unit, grp):

    temp = d.copy()

    temp["Date"] = pd.to_datetime(temp["Date"])

    if unit == "Harian":
        temp["Periode"] = temp["Date"].dt.normalize()

    elif unit == "Mingguan":
        temp["Periode"] = (
            temp["Date"]
            .dt.to_period("W-SUN")
            .dt.start_time
        )

    elif unit == "Bulanan":
        temp["Periode"] = (
            temp["Date"]
            .dt.to_period("M")
            .dt.start_time
        )

    else:
        raise ValueError("Unit periode tidak dikenal.")

    if metric == "Jumlah transaksi":
        temp["_val"] = 1
        agg = "sum"

    elif metric == "Total penjualan (jumlah)":
        temp["_val"] = temp["Total"]
        agg = "sum"

    elif metric == "Rata-rata Total":
        temp["_val"] = temp["Total"]
        agg = "mean"

    elif metric == "Rata-rata Rating":
        temp["_val"] = temp["Rating"]
        agg = "mean"

    elif metric == "Total Quantity":
        temp["_val"] = temp["Quantity"]
        agg = "sum"

    else:
        raise ValueError("Metrik tidak dikenal.")

    group_cols = ["Periode"]

    if not _none(grp):
        group_cols.append(grp)

    td = (
        temp.groupby(group_cols, observed=False)["_val"]
        .agg(agg)
        .reset_index(name="val")
    )

    return td


# ============================================================
# TREND PLOT
# ============================================================

def plot_mv_trend(d, metric, unit, grp):

    pk = _get_pk()

    td = trend_data(d, metric, unit, grp)

    fig, ax = plt.subplots(figsize=(10, 6))

    if _none(grp):

        td = td.sort_values("Periode")

        ax.plot(
            td["Periode"],
            td["val"],
            color=pk["d"],
            linewidth=1.5
        )

        ax.scatter(
            td["Periode"],
            td["val"],
            color=pk["m"],
            s=25
        )

    else:

        categories = d[grp].dropna().unique()
        palette = _palette(len(categories))

        for category, color in zip(categories, palette):

            subset = td[td[grp] == category].sort_values("Periode")

            ax.plot(
                subset["Periode"],
                subset["val"],
                color=color,
                linewidth=1.5,
                label=str(category)
            )

            ax.scatter(
                subset["Periode"],
                subset["val"],
                color=color,
                s=20
            )

        ax.legend(title=grp)

    subtitle = (
        "Minggu pertama/terakhir dapat tidak lengkap"
        if unit == "Mingguan"
        else ""
    )

    ax.set_title(metric + " " + unit)
    ax.set_xlabel("Periode")
    ax.set_ylabel(metric)

    if subtitle:
        ax.text(
            0.5, 1.02,
            subtitle,
            transform=ax.transAxes,
            ha="center",
            fontsize=9
        )

    fig.autofmt_xdate()

    _theme(ax)
    plt.tight_layout()

    return fig


# ============================================================
# PRINCIPAL COMPONENT ANALYSIS
# ============================================================

def pca_fit(d, vars):

    model = PCA()

    X = d[vars].astype(float)

    X_scaled = (
        X - X.mean()
    ) / X.std(ddof=0).replace(0, 1)

    scores = model.fit_transform(X_scaled)

    return model, scores


# ============================================================
# PCA BIPLOT
# ============================================================

def plot_mv_pca(d, vars, colv):

    pk = _get_pk()

    pr, scores = pca_fit(d, vars)

    sc = pd.DataFrame(
        scores[:, :2],
        columns=["PC1", "PC2"],
        index=d.index
    )

    ve = pr.explained_variance_ratio_ * 100

    ld = pd.DataFrame(
        pr.components_[:2].T,
        columns=["PC1", "PC2"],
        index=vars
    )

    k = min(
        sc["PC1"].abs().max(),
        sc["PC2"].abs().max()
    ) * 0.85

    ld["PC1"] *= k
    ld["PC2"] *= k

    fig, ax = plt.subplots(figsize=(9, 7))

    if not _none(colv):

        categories = d[colv].dropna().unique()
        palette = _palette(len(categories))

        for category, color in zip(categories, palette):

            idx = d[colv] == category

            ax.scatter(
                sc.loc[idx, "PC1"],
                sc.loc[idx, "PC2"],
                color=color,
                alpha=0.6,
                s=30,
                label=str(category)
            )

        ax.legend(title=colv)

    else:

        ax.scatter(
            sc["PC1"],
            sc["PC2"],
            color=pk["m"],
            alpha=0.5,
            s=30
        )

    for variable, row in ld.iterrows():

        ax.annotate(
            "",
            xy=(row["PC1"], row["PC2"]),
            xytext=(0, 0),
            arrowprops={
                "arrowstyle": "->",
                "color": pk["d"],
                "lw": 1.3
            }
        )

        ax.text(
            row["PC1"],
            row["PC2"],
            variable,
            color=pk["d"],
            fontweight="bold",
            ha="center",
            va="bottom"
        )

    ax.axhline(0, color="gray", linewidth=0.5)
    ax.axvline(0, color="gray", linewidth=0.5)

    ax.set_title("PCA biplot")
    ax.set_xlabel(f"PC1 ({ve[0]:.1f}%)")
    ax.set_ylabel(f"PC2 ({ve[1]:.1f}%)")

    _theme(ax)
    plt.tight_layout()

    return fig


# ============================================================
# SCATTERPLOT MATRIX / PAIRS PLOT
# ============================================================

def plot_mv_pairs(d, vars, colv):

    pk = _get_pk()

    n = len(vars)

    if n < 2:
        raise ValueError("Pilih minimal dua variabel numerik.")

    grp = None if _none(colv) else d[colv]

    fig, axes = plt.subplots(
        n, n,
        figsize=(3 * n, 3 * n),
        squeeze=False
    )

    categories = (
        grp.dropna().unique()
        if grp is not None else [None]
    )

    palette = _palette(len(categories))

    for i, y in enumerate(vars):

        for j, x in enumerate(vars):

            ax = axes[i, j]

            if i == j:

                sns.histplot(
                    data=d,
                    x=x,
                    color=pk["xl"],
                    edgecolor="white",
                    ax=ax
                )

            elif i > j:

                if grp is None:

                    subset = d[[x, y]].dropna()

                    ax.scatter(
                        subset[x],
                        subset[y],
                        color=pk["m"],
                        alpha=0.55,
                        s=15
                    )

                    if len(subset) >= 3:
                        sm = lowess(
                            subset[y],
                            subset[x],
                            frac=0.6
                        )

                        ax.plot(
                            sm[:, 0],
                            sm[:, 1],
                            color=pk["d"],
                            linewidth=1.5
                        )

                else:

                    for category, color in zip(categories, palette):

                        subset = d[d[colv] == category][[x, y]].dropna()

                        ax.scatter(
                            subset[x],
                            subset[y],
                            color=color,
                            alpha=0.55,
                            s=15,
                            label=str(category)
                        )

                        if len(subset) >= 3:
                            sm = lowess(
                                subset[y],
                                subset[x],
                                frac=0.6
                            )

                            ax.plot(
                                sm[:, 0],
                                sm[:, 1],
                                color=pk["d"],
                                linewidth=1.2
                            )

            else:

                subset = d[[x, y]].dropna()

                if len(subset) > 1:

                    r = subset[x].corr(subset[y])

                    ax.text(
                        0.5, 0.5,
                        f"r = {r:.2f}",
                        transform=ax.transAxes,
                        ha="center",
                        va="center",
                        fontsize=12 + 12 * abs(r),
                        color=pk["d"],
                        fontweight="bold"
                    )

                ax.set_xticks([])
                ax.set_yticks([])

            if i == n - 1:
                ax.set_xlabel(x)

            else:
                ax.set_xlabel("")

            if j == 0:
                ax.set_ylabel(y)

            else:
                ax.set_ylabel("")

            _theme(ax, 8)

    fig.suptitle(
        "Scatterplot matrix (pairs plot)",
        fontsize=14,
        fontweight="bold",
        color=pk["d"]
    )

    if grp is not None:
        handles, labels = axes[0, 0].get_legend_handles_labels()

        if handles:
            fig.legend(
                handles,
                labels,
                title=colv,
                loc="lower center",
                ncol=min(len(labels), 5),
                frameon=False
            )

    plt.tight_layout(rect=[0, 0.04, 1, 0.96])

    return fig


# ============================================================
# MAHALANOBIS DISTANCE
# ============================================================

def mahal_fit(d, vars):

    X = d[vars].astype(float).dropna()

    if len(X) != len(d):
        raise ValueError(
            "Terdapat missing value pada variabel yang dipilih. "
            "Tangani missing value terlebih dahulu."
        )

    X = X.to_numpy()

    p = len(vars)

    if p < 2:
        raise ValueError(
            "Pilih minimal dua variabel numerik."
        )

    if np.linalg.matrix_rank(np.cov(X, rowvar=False)) < p:
        raise ValueError(
            "Matriks kovarians singular (ada variabel yang "
            "bergantung linear). Kurangi variabel yang dipilih."
        )

    mean = np.mean(X, axis=0)

    cov = np.cov(X, rowvar=False)

    inv_cov = np.linalg.inv(cov)

    diff = X - mean

    D2 = np.einsum(
        "ij,jk,ik->i",
        diff,
        inv_cov,
        diff
    )

    cut = chi2.ppf(0.975, df=p)

    df = pd.DataFrame({
        "idx": np.arange(1, len(X) + 1),
        "D2": D2,
        "pval": chi2.sf(D2, df=p),
        "out": D2 > cut
    })

    return {
        "df": df,
        "cut": cut,
        "p": p,
        "X": X
    }


# ============================================================
# MAHALANOBIS DISTANCE INDEX PLOT
# ============================================================

def plot_mv_mahal_idx(f):

    pk = _get_pk()

    df = f["df"].copy()

    df["Status"] = np.where(
        df["out"],
        "Pencilan",
        "Normal"
    )

    o = (
        df[df["out"]]
        .sort_values("D2", ascending=False)
        .head(5)
    )

    fig, ax = plt.subplots(figsize=(10, 6))

    colors = {
        "Normal": pk["l"],
        "Pencilan": pk["d"]
    }

    for status, color in colors.items():

        subset = df[df["Status"] == status]

        ax.scatter(
            subset["idx"],
            subset["D2"],
            color=color,
            alpha=0.7,
            s=30,
            label=status
        )

    ax.axhline(
        f["cut"],
        linestyle="--",
        color=pk["d"],
        linewidth=1.2
    )

    for _, row in o.iterrows():

        ax.text(
            row["idx"],
            row["D2"],
            str(row["idx"]),
            fontsize=8,
            ha="center",
            va="bottom"
        )

    percentage = 100 * df["out"].mean()

    ax.set_title("Jarak Mahalanobis (D2) per observasi")

    ax.set_xlabel("Indeks observasi")
    ax.set_ylabel("D2 (kuadrat jarak Mahalanobis)")

    ax.text(
        0.5, 1.03,
        (
            f"Garis putus-putus = batas Chi-Square 97.5% "
            f"(df = {f['p']}): {f['cut']:.2f} | "
            f"Pencilan: {df['out'].sum()} ({percentage:.1f}%)"
        ),
        transform=ax.transAxes,
        ha="center",
        fontsize=9
    )

    ax.legend(title="Status")

    _theme(ax)
    plt.tight_layout()

    return fig


# ============================================================
# MAHALANOBIS SCATTER PLOT
# ============================================================

def plot_mv_mahal_sc(f, vars):

    pk = _get_pk()

    df = pd.DataFrame({
        "x": f["X"][:, 0],
        "y": f["X"][:, 1],
        "Status": np.where(
            f["df"]["out"],
            "Pencilan",
            "Normal"
        )
    })

    fig, ax = plt.subplots(figsize=(9, 6))

    colors = {
        "Normal": pk["l"],
        "Pencilan": pk["d"]
    }

    sizes = {
        "Normal": 30,
        "Pencilan": 65
    }

    for status in ["Normal", "Pencilan"]:

        subset = df[df["Status"] == status]

        ax.scatter(
            subset["x"],
            subset["y"],
            color=colors[status],
            s=sizes[status],
            alpha=0.65,
            label=status
        )

    ax.set_title(
        "Pencilan multivariat pada ruang dua variabel"
    )

    ax.set_xlabel(vars[0])
    ax.set_ylabel(vars[1])

    ax.text(
        0.5, 1.02,
        f"Ditampilkan: {vars[0]} dan {vars[1]} "
        "(dua variabel pertama yang dipilih)",
        transform=ax.transAxes,
        ha="center",
        fontsize=9
    )

    ax.legend(title="Status")

    _theme(ax)
    plt.tight_layout()

    return fig


# ============================================================
# MAHALANOBIS TABLE
# ============================================================

def mahal_table(d, f, vars):

    df = f["df"].copy()

    o = (
        df.sort_values("D2", ascending=False)
        .head(10)
        .copy()
    )

    out = pd.DataFrame({
        "Baris": o["idx"].values,
        "D2": [
            _format(v, 2) for v in o["D2"]
        ],
        "p-value": [
            _format(v, 4) for v in o["pval"]
        ],
        "Status": np.where(
            o["out"],
            "Pencilan",
            "Normal"
        )
    })

    for v in vars:

        out[v] = [
            _format(d[v].iloc[int(i) - 1], 1)
            for i in o["idx"]
        ]

    return out


# ============================================================
# Q-Q PLOT CHI-SQUARE
# ============================================================

def plot_mv_qqchi(f):

    pk = _get_pk()

    df = f["df"].copy()

    n = len(df)

    o = np.argsort(df["D2"].values)

    s = pd.DataFrame({
        "D2": df["D2"].values[o],
        "idx": df["idx"].values[o],
        "out": df["out"].values[o]
    })

    s["q"] = chi2.ppf(
        (np.arange(1, n + 1) - 0.5) / n,
        df=f["p"]
    )

    s["Status"] = np.where(
        s["out"],
        "Pencilan",
        "Normal"
    )

    fig, ax = plt.subplots(figsize=(9, 6))

    colors = {
        "Normal": pk["m"],
        "Pencilan": pk["d"]
    }

    for status, color in colors.items():

        subset = s[s["Status"] == status]

        ax.scatter(
            subset["q"],
            subset["D2"],
            color=color,
            alpha=0.7,
            s=35,
            label=status
        )

    max_val = max(
        s["q"].max(),
        s["D2"].max()
    )

    ax.plot(
        [0, max_val],
        [0, max_val],
        color=pk["d"],
        linewidth=1.2
    )

    for _, row in s.tail(3).iterrows():

        ax.text(
            row["q"],
            row["D2"],
            str(row["idx"]),
            fontsize=8,
            va="bottom"
        )

    ax.set_title("Q-Q plot Chi-Square")

    ax.set_xlabel("Kuantil teoretis Chi-Square")
    ax.set_ylabel("D2 terurut (sampel)")

    ax.text(
        0.5, 1.03,
        (
            f"D2 terurut vs kuantil Chi-Square "
            f"(df = {f['p']}); titik di sekitar garis = "
            "mendekati normal multivariat"
        ),
        transform=ax.transAxes,
        ha="center",
        fontsize=9
    )

    ax.legend(title="Status")

    _theme(ax)
    plt.tight_layout()

    return fig

# =============================================================================
# 5. REGRESI LINEAR BERGANDA
# =============================================================================

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

import statsmodels.api as sm
import statsmodels.formula.api as smf
import statsmodels.stats.api as sms

from scipy import stats
from scipy.stats import chi2, shapiro
from statsmodels.stats.outliers_influence import variance_inflation_factor
from statsmodels.stats.diagnostic import acorr_ljungbox
from statsmodels.nonparametric.smoothers_lowess import lowess
from patsy import dmatrices, build_design_matrices
import re


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def _pk():
    return globals().get("pk", {
        "bg": "#FFF5F8",
        "d": "#C45B83",
        "m": "#D98BA9",
        "s": "#B65C83",
        "slate": "#8B9BAA",
        "xl": "#F3C6D5",
        "txt": "#59434D",
        "l": "#E8B7CA",
        "mauve": "#B58AA0"
    })


def _fmt(x, digits=3):
    if "fmt" in globals():
        return fmt(x, digits)
    return f"{x:.{digits}f}".replace(".", ",")


def _fmt_p(p):
    if "fmt_p" in globals():
        return fmt_p(p)
    return "< 0,001" if p < 0.001 else _fmt(p, 3)


def _pq(p):
    if "pq" in globals():
        return pq(p)
    return _fmt_p(p)


def _clean(x):
    return str(x).replace("`", "")


def _theme(ax, size=11):
    pk = _pk()
    ax.set_facecolor(pk["bg"])
    ax.tick_params(colors=pk["txt"], labelsize=size)
    ax.xaxis.label.set_color(pk["txt"])
    ax.yaxis.label.set_color(pk["txt"])
    ax.title.set_color(pk["txt"])
    ax.title.set_fontweight("bold")
    ax.grid(False)
    for spine in ax.spines.values():
        spine.set_color("#E8D5DE")


def _palette(n):
    if "pal_cat" in globals():
        return pal_cat(n)
    return sns.color_palette("husl", max(n, 1))


def _is_none(x):
    return x is None or x == "NONE"


# ============================================================
# BUILD REGRESSION MODEL
# ============================================================

def build_model(d, resp, preds, logy):

    if logy and (d[resp] <= 0).any():
        raise ValueError(
            "Transformasi log tidak dapat dilakukan karena "
            "variabel respons memiliki nilai nol atau negatif."
        )

    temp = d.copy()

    response = f"np.log(Q('{resp}'))" if logy else f"Q('{resp}')"

    predictors = " + ".join(
        [f"C(Q('{v}'))" if not pd.api.types.is_numeric_dtype(temp[v])
         else f"Q('{v}')"
         for v in preds]
    )

    formula = f"{response} ~ {predictors}"

    model = smf.ols(formula=formula, data=temp).fit()

    return model


# ============================================================
# CLEAN VARIABLE NAMES
# ============================================================

def clean(x):
    return _clean(x)


# ============================================================
# IDENTIFY TERM
# ============================================================

def term_of(m, cn):

    tl = [
        _clean(t)
        for t in m.model.exog_names
        if t != "Intercept"
    ]

    cn = _clean(cn)

    candidates = [
        t for t in tl
        if cn.startswith(t)
    ]

    if not candidates:
        return cn

    return max(candidates, key=len)


# ============================================================
# COEFFICIENT LABEL
# ============================================================

def coef_label(m, cn):

    tl = term_of(m, cn)
    cn2 = _clean(cn)

    if cn2 == tl:
        return tl

    return f"{tl} = {cn2[len(tl):]}"


# ============================================================
# MAKE REGRESSION EQUATION
# ============================================================

def make_equation(m, resp, logy):

    b = m.params

    lines = []

    lhs = f"ln({resp})" if logy else resp

    lines.append(
        f"{lhs}-hat = {_fmt(b.iloc[0], 4)}"
    )

    for name, value in b.iloc[1:].items():

        label = coef_label(m, name)

        sign = "+" if value >= 0 else "-"

        lines.append(
            f"   {sign} {_fmt(abs(value), 4)} x [{label}]"
        )

    return "\n".join(lines)


# ============================================================
# VIF TABLE / GENERALIZED VIF
# ============================================================

def vif_table(m):

    try:

        X = pd.DataFrame(
            m.model.exog,
            columns=m.model.exog_names
        )

        X = X.drop(columns=["Intercept"], errors="ignore")

        design_info = m.model.data.design_info

        term_slices = design_info.term_name_slices

        groups = {}

        for term, sl in term_slices.items():

            if term == "Intercept":
                continue

            start = sl.start - 1
            stop = sl.stop - 1

            groups[term] = list(range(start, stop))

        if len(groups) < 2:
            return None

        R = np.corrcoef(X.to_numpy(), rowvar=False)

        if R.ndim == 0:
            return None

        det_R = np.linalg.det(R)

        if abs(det_R) < 1e-12:
            return None

        results = []

        for term, indices in groups.items():

            i = indices
            j = [k for k in range(R.shape[0]) if k not in i]

            det_i = np.linalg.det(R[np.ix_(i, i)])

            det_j = (
                np.linalg.det(R[np.ix_(j, j)])
                if j else 1
            )

            gv = det_i * det_j / det_R

            results.append({
                "Variabel": _clean(term),
                "Df": len(i),
                "VIF setara": gv ** (1 / len(i))
            })

        return pd.DataFrame(results)

    except Exception:
        return None


# ============================================================
# BREUSCH-PAGAN TEST
# ============================================================

def bp_test(m):

    e = m.resid

    X = m.model.exog

    u = e ** 2 / np.mean(e ** 2)

    auxiliary = sm.OLS(u, X).fit()

    stat = np.sum(
        (auxiliary.fittedvalues - np.mean(u)) ** 2
    ) / 2

    df = X.shape[1] - 1

    p = chi2.sf(stat, df)

    return {
        "stat": stat,
        "df": df,
        "p": p
    }


# ============================================================
# REGRESSION DATA
# ============================================================

def reg_df(m):

    return pd.DataFrame({
        "fit": m.fittedvalues,
        "act": m.model.endog,
        "res": m.resid,
        "rs": m.get_influence().resid_studentized_internal,
        "cook": m.get_influence().cooks_distance[0],
        "idx": np.arange(1, len(m.resid) + 1)
    })


# ============================================================
# MODEL FIT TABLE
# ============================================================

def fit_table(m):

    e = m.resid

    f = m.fvalue
    df1 = m.df_model
    df2 = m.df_resid

    pF = m.f_pvalue

    data = {
        "Ukuran": [
            "Jumlah observasi (n)",
            "Derajat bebas residual",
            "R-kuadrat",
            "R-kuadrat terkoreksi (Adj. R2)",
            "Galat baku residual (sigma)",
            "RMSE",
            "MAE",
            "Statistik F",
            "p-value (uji F simultan)",
            "AIC",
            "BIC"
        ],
        "Nilai": [
            int(m.nobs),
            int(m.df_resid),
            _fmt(m.rsquared, 4),
            _fmt(m.rsquared_adj, 4),
            _fmt(np.sqrt(m.mse_resid), 3),
            _fmt(np.sqrt(np.mean(e ** 2)), 3),
            _fmt(np.mean(np.abs(e)), 3),
            f"{_fmt(f, 2)} (df1 = {int(df1)}, df2 = {int(df2)})",
            _fmt_p(pF),
            _fmt(m.aic, 2),
            _fmt(m.bic, 2)
        ]
    }

    return pd.DataFrame(data)


# ============================================================
# COEFFICIENT TABLE
# ============================================================

def coef_table(m):

    ci = m.conf_int()

    s = m.summary2().tables[1]

    rows = []

    for name in m.params.index:

        value = m.params[name]

        p = m.pvalues[name]

        label = (
            "Intersep"
            if name == "Intercept"
            else coef_label(m, name)
        )

        if p < 0.001:
            sig = "***"
        elif p < 0.01:
            sig = "**"
        elif p < 0.05:
            sig = "*"
        elif p < 0.1:
            sig = "."
        else:
            sig = ""

        rows.append({
            "Koefisien": label,
            "Estimasi": _fmt(value, 4),
            "Galat baku": _fmt(m.bse[name], 4),
            "t": _fmt(m.tvalues[name], 3),
            "p-value": _fmt_p(p),
            "CI 2.5%": _fmt(ci.loc[name, 0], 3),
            "CI 97.5%": _fmt(ci.loc[name, 1], 3),
            "Sig.": sig
        })

    return pd.DataFrame(rows)


# ============================================================
# ANOVA TABLE
# ============================================================

def anova_table(m):

    a = sm.stats.anova_lm(m, typ=1)

    rows = []

    for variable, row in a.iterrows():

        rows.append({
            "Variabel": _clean(variable),
            "Df": int(row["df"]),
            "Jumlah kuadrat": _fmt(row["sum_sq"], 1),
            "Kuadrat tengah": _fmt(
                row["sum_sq"] / row["df"], 1
            ),
            "F": (
                ""
                if pd.isna(row["F"])
                else _fmt(row["F"], 3)
            ),
            "p-value": (
                ""
                if pd.isna(row["PR(>F)"])
                else _fmt_p(row["PR(>F)"])
            )
        })

    return pd.DataFrame(rows)


# ============================================================
# ACTUAL VS PREDICTED PLOT
# ============================================================

def plot_reg_avp(m, logy, resp):

    pk = _pk()

    df = reg_df(m)

    r2 = m.rsquared

    sc = f"ln({resp})" if logy else resp

    fig, ax = plt.subplots(figsize=(8, 6))

    ax.scatter(
        df["fit"],
        df["act"],
        color=pk["m"],
        alpha=0.45,
        s=35
    )

    low = min(df["fit"].min(), df["act"].min())
    high = max(df["fit"].max(), df["act"].max())

    ax.plot(
        [low, high],
        [low, high],
        color=pk["d"],
        linestyle="--",
        linewidth=1.3
    )

    ax.set_title("Nilai aktual vs prediksi")

    ax.set_xlabel(f"Prediksi {sc}")
    ax.set_ylabel(f"Aktual {sc}")

    ax.text(
        0.5, 1.02,
        f"Garis putus-putus = prediksi sempurna | R2 = {r2:.3f}",
        transform=ax.transAxes,
        ha="center",
        fontsize=9
    )

    _theme(ax)
    plt.tight_layout()

    return fig


# ============================================================
# COEFFICIENT PLOT
# ============================================================

def plot_reg_coef(m):

    pk = _pk()

    ci = m.conf_int()

    rows = []

    for name in m.params.index:

        if name == "Intercept":
            continue

        p = m.pvalues[name]

        rows.append({
            "lab": coef_label(m, name),
            "est": m.params[name],
            "lo": ci.loc[name, 0],
            "hi": ci.loc[name, 1],
            "p": p,
            "sig": (
                "Signifikan (p < 0.05)"
                if p < 0.05
                else "Tidak signifikan"
            )
        })

    df = pd.DataFrame(rows)

    if df.empty:
        raise ValueError("Tidak ada koefisien prediktor untuk ditampilkan.")

    df["lab"] = pd.Categorical(
        df["lab"],
        categories=df["lab"].iloc[::-1],
        ordered=True
    )

    fig, ax = plt.subplots(figsize=(9, max(5, len(df) * 0.5)))

    colors = {
        "Signifikan (p < 0.05)": pk["d"],
        "Tidak signifikan": "#B39DAB"
    }

    ax.axvline(
        0,
        color=pk["mauve"],
        linestyle="--"
    )

    for _, row in df.iterrows():

        ax.errorbar(
            row["est"],
            row["lab"],
            xerr=[
                [row["est"] - row["lo"]],
                [row["hi"] - row["est"]]
            ],
            fmt="o",
            color=colors[row["sig"]],
            capsize=4,
            markersize=7
        )

    ax.set_title("Estimasi koefisien dan selang kepercayaan 95%")

    ax.set_xlabel("Estimasi koefisien")
    ax.set_ylabel("")

    ax.text(
        0.5, 1.02,
        "Selang yang memotong garis nol berarti tidak signifikan",
        transform=ax.transAxes,
        ha="center",
        fontsize=9
    )

    _theme(ax)
    plt.tight_layout()

    return fig


# ============================================================
# RESIDUAL VS FITTED
# ============================================================

def plot_reg_resid(m):

    pk = _pk()

    df = reg_df(m)

    fig, ax = plt.subplots(figsize=(8, 6))

    ax.scatter(
        df["fit"],
        df["res"],
        color=pk["m"],
        alpha=0.45,
        s=30
    )

    smoothed = lowess(
        df["res"],
        df["fit"],
        frac=0.6
    )

    ax.plot(
        smoothed[:, 0],
        smoothed[:, 1],
        color=pk["d"],
        linewidth=1.5
    )

    ax.axhline(
        0,
        color=pk["mauve"],
        linestyle="--"
    )

    ax.set_title("Residual vs nilai prediksi")

    ax.set_xlabel("Nilai prediksi")
    ax.set_ylabel("Residual")

    ax.text(
        0.5, 1.02,
        "Sebaran acak di sekitar nol = asumsi terpenuhi",
        transform=ax.transAxes,
        ha="center",
        fontsize=9
    )

    _theme(ax)
    plt.tight_layout()

    return fig


# ============================================================
# NORMAL Q-Q PLOT
# ============================================================

def plot_reg_qq(m):

    pk = _pk()

    df = reg_df(m)

    fig, ax = plt.subplots(figsize=(8, 6))

    stats.probplot(
        df["rs"],
        dist="norm",
        plot=ax
    )

    ax.get_lines()[0].set_color(pk["m"])
    ax.get_lines()[0].set_alpha(0.6)

    ax.get_lines()[1].set_color(pk["d"])

    ax.set_title(
        "Normal probability plot (QQ-plot residual terstandar)"
    )

    ax.set_xlabel("Kuantil teoretis")
    ax.set_ylabel("Kuantil sampel")

    ax.text(
        0.5, 1.02,
        "Titik mengikuti garis = residual mendekati normal",
        transform=ax.transAxes,
        ha="center",
        fontsize=9
    )

    _theme(ax)
    plt.tight_layout()

    return fig


# ============================================================
# SCALE LOCATION PLOT
# ============================================================

def plot_reg_scale(m):

    pk = _pk()

    df = reg_df(m)

    fig, ax = plt.subplots(figsize=(8, 6))

    ax.scatter(
        df["fit"],
        np.sqrt(np.abs(df["rs"])),
        color=pk["m"],
        alpha=0.45,
        s=30
    )

    smoothed = lowess(
        np.sqrt(np.abs(df["rs"])),
        df["fit"],
        frac=0.6
    )

    ax.plot(
        smoothed[:, 0],
        smoothed[:, 1],
        color=pk["d"],
        linewidth=1.5
    )

    ax.set_title("Scale-Location")

    ax.set_xlabel("Nilai prediksi")
    ax.set_ylabel("Akar |residual terstandar|")

    ax.text(
        0.5, 1.02,
        "Garis datar = ragam residual konstan (homoskedastis)",
        transform=ax.transAxes,
        ha="center",
        fontsize=9
    )

    _theme(ax)
    plt.tight_layout()

    return fig


# ============================================================
# COOK'S DISTANCE
# ============================================================

def plot_reg_cook(m):

    pk = _pk()

    df = reg_df(m)

    thr = 4 / len(df)

    top = df.nlargest(3, "cook")

    fig, ax = plt.subplots(figsize=(9, 6))

    ax.vlines(
        df["idx"],
        0,
        df["cook"],
        color=pk["l"]
    )

    ax.axhline(
        thr,
        color=pk["d"],
        linestyle="--"
    )

    ax.scatter(
        df["idx"],
        df["cook"],
        color=pk["m"],
        s=20
    )

    for _, row in top.iterrows():

        ax.text(
            row["idx"],
            row["cook"],
            str(int(row["idx"])),
            color=pk["d"],
            ha="center",
            va="bottom"
        )

    ax.set_title("Cook's distance")

    ax.set_xlabel("Indeks observasi")
    ax.set_ylabel("Cook's distance")

    ax.text(
        0.5, 1.02,
        f"Garis putus-putus = 4/n = {thr:.4f}",
        transform=ax.transAxes,
        ha="center",
        fontsize=9
    )

    _theme(ax)
    plt.tight_layout()

    return fig


# ============================================================
# REGRESSION ADDITIONAL VISUALIZATIONS
# ============================================================

def reg_ord_res(m, d):

    if "Date" in d.columns:
        order = np.argsort(
            pd.to_datetime(d["Date"]).values,
            kind="stable"
        )
    else:
        order = np.arange(len(d))

    return {
        "res": np.asarray(m.resid)[order],
        "date": (
            pd.to_datetime(d["Date"]).values[order]
            if "Date" in d.columns
            else np.arange(1, len(d) + 1)
        )
    }


# ============================================================
# REGRESSION SCATTERPLOT
# ============================================================

def plot_reg_scatter(d, resp, preds, logy):

    pk = _pk()

    num_all = globals().get("num_all", [])

    nums = [v for v in preds if v in num_all]

    if len(nums) < 1:
        raise ValueError(
            "Pilih minimal satu prediktor numerik untuk scatterplot."
        )

    yv = np.log(d[resp]) if logy else d[resp]

    n = len(nums)

    fig, axes = plt.subplots(
        1, n,
        figsize=(6 * n, 5),
        squeeze=False
    )

    for i, v in enumerate(nums):

        ax = axes[0, i]

        temp = pd.DataFrame({
            "x": d[v],
            "y": yv
        }).dropna()

        ax.scatter(
            temp["x"],
            temp["y"],
            color=pk["m"],
            alpha=0.35,
            s=20
        )

        sns.regplot(
            data=temp,
            x="x",
            y="y",
            scatter=False,
            ci=95,
            color=pk["d"],
            line_kws={"linewidth": 1.2},
            ax=ax
        )

        ax.set_title(v)
        ax.set_xlabel(v)
        ax.set_ylabel(
            f"ln({resp})" if logy else resp
        )

        _theme(ax)

    fig.suptitle(
        "Scatterplot respons terhadap prediktor numerik",
        fontweight="bold",
        color=pk["d"]
    )

    plt.tight_layout()

    return fig


# ============================================================
# REGRESSION CORRELATION MATRIX
# ============================================================

def plot_reg_cor(d, resp, preds):

    num_all = globals().get("num_all", [])

    vars = list(dict.fromkeys(
        [resp] + [v for v in preds if v in num_all]
    ))

    if len(vars) < 2:
        raise ValueError(
            "Pilih minimal satu prediktor numerik untuk matriks korelasi."
        )

    fig = plot_mv_cor(
        d,
        vars,
        "pearson"
    )

    return fig


# ============================================================
# RESIDUAL HISTOGRAM
# ============================================================

def plot_reg_hist(m):

    pk = _pk()

    df = reg_df(m)

    fig, ax = plt.subplots(figsize=(8, 6))

    sns.histplot(
        df["res"],
        bins=30,
        stat="density",
        color=pk["xl"],
        edgecolor="white",
        ax=ax
    )

    sns.kdeplot(
        df["res"],
        color=pk["d"],
        linewidth=1.3,
        ax=ax
    )

    x = np.linspace(
        df["res"].min(),
        df["res"].max(),
        200
    )

    ax.plot(
        x,
        stats.norm.pdf(
            x,
            loc=df["res"].mean(),
            scale=df["res"].std()
        ),
        color=pk["slate"],
        linestyle="--",
        linewidth=1.2
    )

    ax.set_title("Histogram residual")

    ax.set_xlabel("Residual")
    ax.set_ylabel("Densitas")

    ax.text(
        0.5, 1.02,
        "Garis tebal = density residual; putus-putus = kurva normal teoretis",
        transform=ax.transAxes,
        ha="center",
        fontsize=9
    )

    _theme(ax)
    plt.tight_layout()

    return fig


# ============================================================
# RESIDUAL TIME SERIES
# ============================================================

def plot_reg_ts(m, d):

    pk = _pk()

    o = reg_ord_res(m, d)

    n = len(o["res"])

    df = pd.DataFrame({
        "t": np.arange(1, n + 1),
        "res": o["res"]
    })

    fig, ax = plt.subplots(figsize=(10, 6))

    ax.plot(
        df["t"],
        df["res"],
        color=pk["l"],
        linewidth=0.6
    )

    ax.scatter(
        df["t"],
        df["res"],
        color=pk["m"],
        s=12,
        alpha=0.5
    )

    ax.axhline(
        0,
        color=pk["mauve"],
        linestyle="--"
    )

    smoothed = lowess(
        df["res"],
        df["t"],
        frac=0.6
    )

    ax.plot(
        smoothed[:, 0],
        smoothed[:, 1],
        color=pk["d"],
        linewidth=1.3
    )

    br = np.unique(
        np.round(np.linspace(0, n - 1, 6)).astype(int)
    )

    if "Date" in d.columns:
        labels = pd.to_datetime(o["date"][br]).strftime("%d-%m-%Y")
    else:
        labels = br + 1

    ax.set_xticks(br + 1)
    ax.set_xticklabels(labels, rotation=30)

    ax.set_title("Time series plot residual")

    ax.set_xlabel("Tanggal transaksi (urutan observasi)")
    ax.set_ylabel("Residual")

    ax.text(
        0.5, 1.02,
        "Observasi diurutkan menurut tanggal transaksi; garis tebal = tren (loess)",
        transform=ax.transAxes,
        ha="center",
        fontsize=9
    )

    _theme(ax)
    plt.tight_layout()

    return fig


# ============================================================
# RESIDUAL ACF
# ============================================================

def plot_reg_acf(m, d):

    pk = _pk()

    o = reg_ord_res(m, d)

    residuals = o["res"]

    n = len(residuals)

    max_lag = min(30, n - 1)

    acf_values = sm.tsa.acf(
        residuals,
        nlags=max_lag,
        fft=False
    )

    df = pd.DataFrame({
        "lag": np.arange(1, max_lag + 1),
        "acf": acf_values[1:]
    })

    ci = stats.norm.ppf(0.975) / np.sqrt(n)

    fig, ax = plt.subplots(figsize=(9, 6))

    ax.axhline(
        0,
        color=pk["mauve"]
    )

    ax.axhline(
        ci,
        color=pk["slate"],
        linestyle="--"
    )

    ax.axhline(
        -ci,
        color=pk["slate"],
        linestyle="--"
    )

    ax.vlines(
        df["lag"],
        0,
        df["acf"],
        color=pk["m"],
        linewidth=1.2
    )

    ax.scatter(
        df["lag"],
        df["acf"],
        color=pk["d"],
        s=25
    )

    ax.set_title("ACF plot residual")

    ax.set_xlabel("Lag")
    ax.set_ylabel("Autokorelasi")

    ax.text(
        0.5, 1.02,
        "Garis putus-putus = batas 95%; lag 0 tidak ditampilkan",
        transform=ax.transAxes,
        ha="center",
        fontsize=9
    )

    _theme(ax)
    plt.tight_layout()

    return fig


# ============================================================
# OBSERVED VS RESIDUAL
# ============================================================

def plot_reg_obsres(m):

    pk = _pk()

    df = reg_df(m)

    fig, ax = plt.subplots(figsize=(8, 6))

    ax.scatter(
        df["act"],
        df["res"],
        color=pk["m"],
        alpha=0.45,
        s=30
    )

    smoothed = lowess(
        df["res"],
        df["act"],
        frac=0.6
    )

    ax.plot(
        smoothed[:, 0],
        smoothed[:, 1],
        color=pk["d"],
        linewidth=1.4
    )

    ax.axhline(
        0,
        color=pk["mauve"],
        linestyle="--"
    )

    ax.set_title("Observed vs residual")

    ax.set_xlabel("Nilai observasi (aktual)")
    ax.set_ylabel("Residual")

    ax.text(
        0.5, 1.02,
        "Pola menyebar/mengerucut menandakan ragam residual tidak konstan",
        transform=ax.transAxes,
        ha="center",
        fontsize=9
    )

    _theme(ax)
    plt.tight_layout()

    return fig


# ============================================================
# REGRESSION INTERPRETATION
# ============================================================

def interp_reg(m, resp, logy):

    f = m.fvalue
    pF = m.f_pvalue

    n = int(m.nobs)

    r2 = m.rsquared
    ar2 = m.rsquared_adj

    kat = (
        "lemah" if ar2 < 0.3
        else "sedang" if ar2 < 0.6
        else "kuat"
    )

    response = f"ln({resp})" if logy else resp

    out = [
        (
            f"**Uji simultan (uji F):** F = {_fmt(f, 2)}, "
            f"p {_pq(pF)}. "
            + (
                "Minimal satu prediktor berpengaruh signifikan terhadap respons, sehingga model layak digunakan."
                if pF < 0.05
                else "Tidak ada prediktor yang berpengaruh signifikan secara simultan."
            )
        ),

        (
            f"**Kemampuan model:** R2 = {_fmt(r2, 3)} dan "
            f"Adj. R2 = {_fmt(ar2, 3)}, artinya sekitar "
            f"**{_fmt(100 * ar2, 1)}%** keragaman {response} "
            f"dijelaskan oleh prediktor dalam model "
            f"(kategori **{kat}**); sisanya dijelaskan faktor lain di luar model."
        )
    ]

    co = m.params.drop("Intercept", errors="ignore")

    significant = co[m.pvalues[co.index] < 0.05]

    if len(significant) == 0:

        out.append(
            "**Uji parsial (uji t):** tidak ada koefisien yang signifikan pada taraf 5%."
        )

    else:

        labels = [
            coef_label(m, name)
            for name in significant.index
        ]

        out.append(
            "**Uji parsial (uji t):** prediktor signifikan pada taraf 5%: "
            f"**{', '.join(labels)}**."
        )

        for name, b in significant.items():

            label = coef_label(m, name)

            term = term_of(m, name)

            is_categorical = (
                "C(" in name or "[" in name
            )

            if logy:

                effect = 100 * (np.exp(abs(b)) - 1)

                eff = f"{_fmt(effect, 1)}%"

            else:

                eff = _fmt(abs(b), 3)

            verb = "naik" if b > 0 else "turun"

            if is_categorical:

                out.append(
                    f"Koefisien kategori **{label}** menunjukkan "
                    f"perbedaan rata-rata respons dibandingkan kategori acuan. "
                    f"Interpretasi besarnya perubahan mengikuti koefisien model."
                )

            else:

                out.append(
                    f"Setiap kenaikan 1 satuan **{term}**, "
                    f"rata-rata **{resp}** {verb} sebesar **{eff}**, "
                    "dengan variabel lain dianggap tetap."
                )

    all_terms = list(dict.fromkeys(
        term_of(m, name)
        for name in co.index
    ))

    sig_terms = list(dict.fromkeys(
        term_of(m, name)
        for name in significant.index
    ))

    ns = [
        term for term in all_terms
        if term not in sig_terms
    ]

    if ns:

        out.append(
            "**Variabel tidak signifikan** "
            "(tidak ada bukti pengaruh setelah variabel lain dikendalikan): "
            f"{', '.join(ns)}."
        )

    residuals = m.resid

    sample_resid = (
        np.random.default_rng(42).choice(
            residuals,
            size=5000,
            replace=False
        )
        if len(residuals) > 5000
        else residuals
    )

    sh = shapiro(sample_resid).pvalue

    bp = bp_test(m)["p"]

    vt = vif_table(m)

    out.append(
        f"**Asumsi residual:** normalitas (Shapiro-Wilk) "
        f"p {_pq(sh)}, sehingga "
        + (
            "residual **tidak normal**"
            if sh < 0.05
            else "residual normal"
        )
        + f"; kesamaan ragam (Breusch-Pagan) p {_pq(bp)}, sehingga "
        + (
            "ragam residual **tidak konstan**."
            if bp < 0.05
            else "ragam konstan."
        )
    )

    if vt is not None and not vt.empty:

        max_vif = vt["VIF setara"].max()

        out.append(
            f"**Multikolinearitas:** VIF setara maksimum = "
            f"{_fmt(max_vif, 2)} "
            + (
                "(aman, di bawah 5)."
                if max_vif < 5
                else "(perlu diperhatikan, di atas 5)."
            )
        )

    return out

# =============================================================================
# KESIMPULAN OTOMATIS
# =============================================================================
def auto_conclusion(d, m, resp, logy):
    """
    Membuat kesimpulan otomatis berdasarkan data dan model regresi.
    Membutuhkan fungsi coef_label() dari blok kode regresi sebelumnya.
    """

    import numpy as np
    import pandas as pd

    from scipy.stats import skew

    # Ringkasan model dan koefisien selain intercept
    summary = m.summary2()
    coef = m.params.drop(labels=["Intercept"], errors="ignore")
    pvalues = m.pvalues.drop(labels=["Intercept"], errors="ignore")

    # Prediktor yang signifikan pada taraf 5%
    significant = pvalues[pvalues < 0.05].index.tolist()

    # Korelasi antarvariabel numerik
    corr_vars = ["Quantity", "Total", "Rating", "Hour"]
    corr_data = d[corr_vars].apply(pd.to_numeric, errors="coerce")
    corr_matrix = corr_data.corr()

    # Menghilangkan korelasi diagonal
    corr_matrix = corr_matrix.copy()
    np.fill_diagonal(corr_matrix.values, 0)

    # Menentukan pasangan variabel dengan korelasi terkuat
    corr_abs = corr_matrix.abs()
    i, j = np.unravel_index(
        np.nanargmax(corr_abs.to_numpy()),
        corr_abs.shape
    )

    corr_x = corr_matrix.index[i]
    corr_y = corr_matrix.columns[j]
    corr_value = corr_matrix.iloc[i, j]

    # Statistik deskriptif Total
    total = pd.to_numeric(d["Total"], errors="coerce").dropna()

    n = len(d)
    min_total = total.min()
    max_total = total.max()
    median_total = total.median()
    mean_total = total.mean()

    # Mengukur kemencengan distribusi
    skewness = skew(total, nan_policy="omit")

    distribution = (
        "distribusi miring ke kanan"
        if skewness > 0.5
        else "distribusi relatif simetris"
    )

    # Format angka
    def format_number(x, digits=2):
        if pd.isna(x):
            return "-"
        return f"{x:,.{digits}f}"

    # Menyusun nama prediktor signifikan
    if len(significant) == 0:
        significant_text = "tidak ada"
    else:
        significant_text = ", ".join(
            coef_label(m, name) for name in significant
        )

    # Mengambil Adjusted R-squared
    adj_r2 = m.rsquared_adj

    # Membuat kesimpulan
    conclusions = [
        (
            f"Data yang dianalisis berisi **{n} transaksi**; "
            f"nilai transaksi (Total) berkisar "
            f"{format_number(min_total)} sampai {format_number(max_total)} "
            f"dengan median **{format_number(median_total)}** dan "
            f"rata-rata {format_number(mean_total)} "
            f"({distribution})."
        ),

        (
            f"Pasangan variabel numerik dengan korelasi terkuat adalah "
            f"**{corr_x}** dan **{corr_y}** "
            f"(r = {format_number(corr_value, 3)})."
        ),

        (
            f"Model regresi (respons **{resp}**"
            f"{', transformasi ln' if logy else ''}) "
            f"memiliki Adj. R2 = **{format_number(adj_r2, 3)}**; "
            f"prediktor signifikan pada taraf 5%: "
            f"**{significant_text}**."
        )
    ]

    return conclusions


# ============================================================
# REKOMENDASI
# ============================================================

REKOMENDASI = [
    (
        "Jumlah unit per transaksi (Quantity) merupakan faktor utama "
        "nilai transaksi; promosi berbasis jumlah pembelian "
        "(paket bundling, potongan harga pembelian banyak) "
        "berpotensi menaikkan nilai transaksi."
    ),

    (
        "Cabang, tipe pelanggan, jenis kelamin, lini produk, dan metode "
        "pembayaran tidak menunjukkan pengaruh nyata terhadap nilai "
        "transaksi setelah Quantity dikendalikan, sehingga strategi "
        "dari sisi nilai transaksi tidak perlu dibedakan berdasarkan "
        "segmen tersebut."
    ),

    (
        "Rating tidak berkaitan dengan nilai transaksi maupun jumlah "
        "pembelian; peningkatan kepuasan pelanggan perlu ditelusuri "
        "melalui data tambahan seperti kualitas layanan dan waktu tunggu."
    ),

    (
        "Penelitian lanjutan disarankan menambahkan variabel harga satuan, "
        "periode data yang lebih panjang, serta model yang menangani "
        "ragam tidak konstan (transformasi atau regresi robust)."
    )
]


# ============================================================
# KETERBATASAN
# ============================================================

KETERBATASAN = [
    (
        "Data bersifat observasional dan hanya mencakup periode singkat "
        "(Januari-Maret) sehingga hasil tidak dapat digeneralisasi "
        "ke seluruh waktu."
    ),

    (
        "Regresi menunjukkan hubungan (asosiasi), bukan sebab-akibat."
    ),

    (
        "Variabel penting seperti harga satuan tidak tersedia pada data, "
        "sehingga sebagian keragaman nilai transaksi tidak dapat "
        "dijelaskan model."
    ),

    (
        "Residual cenderung tidak normal dan ragamnya tidak konstan "
        "(bertambah seiring Quantity); inferensi (uji t/F, selang "
        "kepercayaan) hanya pendekatan. Transformasi log atau regresi "
        "robust dapat dipertimbangkan."
    ),

    (
        "Model linear mengasumsikan efek aditif dan linear; interaksi "
        "antar variabel tidak dimodelkan."
    )
]
# =============================================================================
# TAMPILAN (CSS formal, tema pink)
# =============================================================================
css = """
@import url('https://fonts.googleapis.com/css2?family=Source+Sans+3:wght@400;600;700&family=Source+Serif+4:wght@600;700&display=swap');
:root{--p9:#880E4F;--p7:#C2185B;--p5:#F06292;--p3:#F8BBD0;--p1:#FFF5F8;--tx:#3B2A33;--ln:#EBD3DC;}
body{font-family:'Source Sans 3','Segoe UI',Arial,sans-serif;color:var(--tx);background:#fff;font-size:15px;}
h1,h2,h3,h4{font-family:'Source Serif 4',Georgia,'Times New Roman',serif;color:var(--p9);}
.app-header{border-top:6px solid var(--p7);border-bottom:1px solid var(--ln);background:var(--p1);padding:22px 30px 16px;margin:-15px -15px 22px;}
.app-header h1{margin:0 0 6px;font-size:25px;font-weight:700;}
.app-header .sub{color:#6D4C5E;font-size:14px;margin:0;}
.well{background:var(--p1);border:1px solid var(--ln);border-radius:4px;box-shadow:none;}
.well h4{font-size:15px;border-bottom:1px solid var(--ln);padding-bottom:6px;margin-top:0;}
.box-x{background:#fff;border:1px solid var(--ln);border-left:4px solid var(--p7);border-radius:3px;padding:14px 18px;margin-bottom:16px;}
.box-x h4{margin-top:0;font-size:16px;font-weight:700;}
.box-x p,.box-x li{line-height:1.6;}
.nav-tabs{border-bottom:2px solid var(--p3);}
.nav-tabs>li>a{color:#6D4C5E;font-weight:600;border:none;border-radius:0;margin-right:2px;}
.nav-tabs>li>a:hover{background:var(--p1);border:none;color:var(--p9);}
.nav-tabs>li.active>a,.nav-tabs>li.active>a:focus,.nav-tabs>li.active>a:hover{color:var(--p9);border:none;border-bottom:3px solid var(--p7);background:#fff;}
.btn-pink{background:var(--p7);border:1px solid var(--p7);color:#fff;border-radius:3px;font-weight:600;}
.btn-pink:hover,.btn-pink:focus{background:var(--p9);border-color:var(--p9);color:#fff;}
.btn-default{border-radius:3px;border-color:var(--ln);}
.control-label{font-weight:600;color:#6D4C5E;font-size:13px;}
.selectize-input.focus,.form-control:focus{border-color:var(--p5);box-shadow:0 0 0 2px rgba(240,98,146,.2);}
.checkbox input[type=checkbox]{accent-color:var(--p7);}
.radio input[type=radio]{accent-color:var(--p7);}
table.table{font-size:13.5px;}
table.table>thead>tr>th{background:var(--p7);color:#fff;border-bottom:none;font-weight:600;}
table.table-striped>tbody>tr:nth-of-type(odd){background:var(--p1);}
pre{background:var(--p1);border:1px solid var(--ln);color:var(--tx);font-size:12.5px;border-radius:3px;}
.nfilter{background:#fff;border:1px solid var(--ln);border-radius:3px;padding:8px 10px;margin-top:10px;font-size:13px;color:#6D4C5E;}
.nfilter b{color:var(--p9);font-size:16px;}
.irs--shiny .irs-bar{background:var(--p7);border-color:var(--p7);}
.help-note{color:#7d6672;font-size:13px;}
#uni_stem{white-space:pre;overflow:auto;max-height:400px;font-size:12px;}
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import date

# ============================================================
# KONFIGURASI HALAMAN
# ============================================================

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# FUNGSI BANTUAN
# ============================================================

def sel_cat(label, options, selected=None):
    return st.selectbox(
        label,
        options=options,
        index=options.index(selected)
        if selected in options else 0
    )


def bullets(items):
    for item in items:
        st.markdown(f"- {item}")


def show_table(data):
    if data is not None:
        st.dataframe(data, use_container_width=True, hide_index=True)


def show_plot(fig):
    if fig is not None:
        st.plotly_chart(fig, use_container_width=True)


def show_interpretation(content):
    if isinstance(content, (list, tuple)):
        for item in content:
            st.markdown(f"- {item}")
    elif content:
        st.markdown(content)


# ============================================================
# MEMBACA DATA AWAL
# ============================================================

try:
    d0 = read_sales(DATA_FILE)
except Exception:
    d0 = None


def ch(variable, default):
    if d0 is None or variable not in d0.columns:
        return default

    values = d0[variable].dropna().unique().tolist()
    return sorted(values, key=str)


# ============================================================
# CSS
# ============================================================

st.markdown(
    f"""
    <style>
    .stApp {{
        background-color: #FFF8FB;
    }}

    .app-header {{
        background: linear-gradient(135deg, #F7DCE8, #FCECF3);
        padding: 25px 30px;
        border-radius: 15px;
        margin-bottom: 20px;
        border: 1px solid #F0D2DF;
    }}

    .app-header h1 {{
        color: #8E4968;
        font-size: 30px;
        margin-bottom: 5px;
    }}

    .app-header p {{
        color: #8E4968;
        font-size: 13px;
    }}

    .box-x {{
        background-color: white;
        border: 1px solid #F0DCE5;
        border-radius: 12px;
        padding: 18px;
        margin-bottom: 15px;
    }}

    .box-x h3, .box-x h4 {{
        color: #A45D7D;
    }}

    div.stButton > button {{
        background-color: #D98BAE;
        color: white;
        border: none;
        border-radius: 8px;
    }}

    div.stButton > button:hover {{
        background-color: #BF6F95;
        color: white;
    }}

    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# HEADER APLIKASI
# ============================================================

st.markdown(
    f"""
    <div class="app-header">
        <h1>{APP_TITLE}</h1>
        <p>
            Ujian Tengah Semester Eksplorasi dan Visualisasi Data Multivariat
            | {NAMA} | {NIM}
        </p>
    </div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# SIDEBAR: SUMBER DATA DAN FILTER
# ============================================================

st.sidebar.header("Sumber data")

uploaded_file = st.sidebar.file_uploader(
    "Pilih file",
    type=["xlsx", "xls", "csv"],
    help="Excel/CSV (opsional)"
)

# Membaca file unggahan
if uploaded_file is not None:
    try:
        if uploaded_file.name.lower().endswith(".csv"):
            raw_data = pd.read_csv(uploaded_file)
        else:
            raw_data = pd.read_excel(uploaded_file)

        d0 = prepare_sales(raw_data)

        st.sidebar.success("Data berhasil dimuat.")

    except Exception as e:
        st.sidebar.error(f"Gagal membaca data: {e}")

# Jika tidak ada file unggahan, gunakan data awal
if d0 is None:
    st.error("Data belum tersedia. Silakan unggah file Excel atau CSV.")
    st.stop()


# Tombol reset filter
if st.sidebar.button("Atur ulang filter"):
    st.session_state.clear()
    st.rerun()


st.sidebar.header("Filter data")

# Filter Branch
branch_options = ch("Branch", ["A", "B", "C"])

f_branch = st.sidebar.multiselect(
    "Branch",
    options=branch_options,
    default=branch_options,
    key="f_branch"
)

# Filter Customer type
cust_options = ch("Customer type", ["Member", "Normal"])

f_cust = st.sidebar.multiselect(
    "Customer type",
    options=cust_options,
    default=cust_options,
    key="f_cust"
)

# Filter Gender
gender_options = ch("Gender", ["Female", "Male"])

f_gender = st.sidebar.multiselect(
    "Gender",
    options=gender_options,
    default=gender_options,
    key="f_gender"
)

# Filter Payment
pay_options = ch("Payment", ["Cash", "Credit card", "Ewallet"])

f_pay = st.sidebar.multiselect(
    "Payment",
    options=pay_options,
    default=pay_options,
    key="f_pay"
)

# Filter Product line
prod_options = ch("Product line", [])

f_prod = st.sidebar.multiselect(
    "Product line",
    options=prod_options,
    default=prod_options,
    key="f_prod"
)

# Filter tanggal
if "Date" in d0.columns:
    min_date = pd.to_datetime(d0["Date"]).min().date()
    max_date = pd.to_datetime(d0["Date"]).max().date()

    f_date = st.sidebar.date_input(
        "Rentang tanggal",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date,
        format="DD-MM-YYYY"
    )
else:
    f_date = None


# ============================================================
# MENERAPKAN FILTER DATA
# ============================================================

d = d0.copy()

if "Branch" in d.columns:
    d = d[d["Branch"].isin(f_branch)]

if "Customer type" in d.columns:
    d = d[d["Customer type"].isin(f_cust)]

if "Gender" in d.columns:
    d = d[d["Gender"].isin(f_gender)]

if "Payment" in d.columns:
    d = d[d["Payment"].isin(f_pay)]

if "Product line" in d.columns:
    d = d[d["Product line"].isin(f_prod)]

if f_date is not None and len(f_date) == 2:
    d["Date"] = pd.to_datetime(d["Date"])

    d = d[
        (d["Date"].dt.date >= f_date[0]) &
        (d["Date"].dt.date <= f_date[1])
    ]


# Informasi jumlah data
st.sidebar.markdown("---")
st.sidebar.metric(
    "Jumlah transaksi",
    f"{len(d):,}"
)

st.sidebar.caption(
    f"Data awal: {len(d0):,} transaksi"
)


# ============================================================
# DEFINISI VARIABEL DAN PILIHAN
# ============================================================

cat_all = [
    col for col in cat_all
    if col in d.columns
]

num_all = [
    col for col in num_all
    if col in d.columns
]

cat_base = [
    col for col in cat_all
    if col in d.columns
]

NONE = "Tidak ada"


# ============================================================
# TAB UTAMA
# ============================================================

tabs = st.tabs([
    "1. Masalah dan Data",
    "2. Univariat",
    "3. Bivariat",
    "4. Multivariat",
    "5. Regresi Linear Berganda",
    "6. Kesimpulan"
])


# ============================================================
# 1. MASALAH DAN DATA
# ============================================================

with tabs[0]:

    st.markdown("### Latar belakang")
    st.markdown(f'<div class="box-x">{LATAR_BELAKANG}</div>',
                unsafe_allow_html=True)

    st.markdown("### Tujuan analisis")
    bullets(TUJUAN)

    st.markdown("### Sumber dan karakteristik data")
    st.markdown(SUMBER_DATA)

    st.markdown("### Informasi data")

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric("Jumlah observasi", f"{len(d):,}")

    with col2:
        st.metric("Jumlah variabel", len(d.columns))

    with col3:
        st.metric("Data kosong", int(d.isna().sum().sum()))

    st.markdown("### Definisi variabel")

    if "KAMUS" in globals():
        show_table(KAMUS)

    st.markdown("### Pengolahan data")

    if "prep_info" in globals():
        show_interpretation(prep_info)

    st.markdown("### Cuplikan data")
    st.dataframe(d.head(10), use_container_width=True)


# ============================================================
# 2. UNIVARIAT
# ============================================================

with tabs[1]:

    st.markdown("### Analisis univariat")

    uni_options = cat_all + num_all + (
        ["Date"] if "Date" in d.columns else []
    )

    uni_var = st.selectbox(
        "Variabel",
        uni_options,
        index=uni_options.index("Total")
        if "Total" in uni_options else 0
    )

    col1, col2 = st.columns([7, 5])

    with col1:
        st.markdown("#### Visualisasi")

        # Menggunakan fungsi visualisasi univariat sebelumnya
        if "uni_plot" in globals():
            fig = uni_plot(d, uni_var)
            show_plot(fig)

    with col2:
        st.markdown("#### Statistik deskriptif")

        if "uni_stats" in globals():
            stats = uni_stats(d, uni_var)
            show_table(stats)

    st.markdown("### Interpretasi")

    if "uni_text" in globals():
        show_interpretation(uni_text(d, uni_var))


# ============================================================
# 3. BIVARIAT
# ============================================================

with tabs[2]:

    st.markdown("### Analisis bivariat")

    col1, col2, col3 = st.columns(3)

    with col1:
        bi_x = st.selectbox(
            "Variabel X",
            cat_all + num_all,
            index=(cat_all + num_all).index("Quantity")
            if "Quantity" in cat_all + num_all else 0
        )

    with col2:
        bi_y = st.selectbox(
            "Variabel Y",
            cat_all + num_all,
            index=(cat_all + num_all).index("Total")
            if "Total" in cat_all + num_all else 0
        )

    with col3:
        st.markdown("#### Jenis visualisasi")
        bi_type = None

    st.info(
        "Jenis grafik menyesuaikan tipe pasangan variabel: "
        "numerik-numerik, kategorik-numerik, atau kategorik-kategorik."
    )

    col1, col2 = st.columns([7, 5])

    with col1:
        st.markdown("#### Visualisasi")
        if "bi_plot" in globals():
            fig = bi_plot(d, bi_x, bi_y)
            show_plot(fig)

    with col2:
        st.markdown("#### Ringkasan dan uji")
        if "bi_summary" in globals():
            show_table(bi_summary(d, bi_x, bi_y))

    st.markdown("### Interpretasi")

    if "bi_text" in globals():
        show_interpretation(bi_text(d, bi_x, bi_y))


# ============================================================
# 4. MULTIVARIAT
# ============================================================

with tabs[3]:

    mv_tabs = st.tabs([
        "Scatter multivariat",
        "Heatmap rata-rata",
        "Matriks korelasi",
        "Tren waktu",
        "PCA",
        "Scatterplot matrix",
        "Pencilan Mahalanobis",
        "Q-Q plot Chi-Square"
    ])


    # --------------------------------------------------------
    # 4.1 SCATTER MULTIVARIAT
    # --------------------------------------------------------

    with mv_tabs[0]:

        st.markdown("### Scatter multivariat")

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            mv_x = st.selectbox("Sumbu X", num_all, key="mv_x")

        with col2:
            mv_y = st.selectbox("Sumbu Y", num_all, key="mv_y")

        with col3:
            mv_col = st.selectbox(
                "Warna",
                [NONE] + cat_all + num_all,
                index=(
                    [NONE] + cat_all + num_all
                ).index("Customer type")
                if "Customer type" in cat_all else 0,
                key="mv_col"
            )

        with col4:
            mv_size = st.selectbox(
                "Ukuran titik",
                [NONE] + num_all,
                index=(
                    [NONE] + num_all
                ).index("Rating")
                if "Rating" in num_all else 0,
                key="mv_size"
            )

        col1, col2 = st.columns(2)

        with col1:
            mv_facet = st.selectbox(
                "Panel (facet)",
                [NONE] + cat_all,
                key="mv_facet"
            )

        with col2:
            mv_lines = st.checkbox(
                "Tampilkan garis regresi per kelompok",
                value=True,
                key="mv_lines"
            )

        if "mv_scatter" in globals():
            fig = mv_scatter(
                d,
                mv_x,
                mv_y,
                mv_col,
                mv_size,
                mv_facet,
                mv_lines
            )
            show_plot(fig)

        st.markdown("### Interpretasi")

        if "mv_scatter_text" in globals():
            show_interpretation(
                mv_scatter_text(
                    d, mv_x, mv_y, mv_col, mv_size
                )
            )


    # --------------------------------------------------------
    # 4.2 HEATMAP RATA-RATA
    # --------------------------------------------------------

    with mv_tabs[1]:

        st.markdown("### Heatmap rata-rata")

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            mv_ha = sel_cat(
                "Kategorik 1 (sumbu X)",
                cat_all,
                "Branch"
            )

        with col2:
            mv_hb = sel_cat(
                "Kategorik 2 (sumbu Y)",
                cat_all,
                "Product line"
            )

        with col3:
            mv_hn = st.selectbox(
                "Variabel numerik",
                num_all,
                index=num_all.index("Total")
                if "Total" in num_all else 0
            )

        with col4:
            mv_hf = st.selectbox(
                "Ukuran",
                ["Rata-rata", "Median", "Jumlah", "Banyak transaksi"]
            )

        if "mv_heat" in globals():
            fig = mv_heat(d, mv_ha, mv_hb, mv_hn, mv_hf)
            show_plot(fig)

        st.markdown("### Interpretasi")

        if "mv_heat_text" in globals():
            show_interpretation(
                mv_heat_text(d, mv_ha, mv_hb, mv_hn, mv_hf)
            )


    # --------------------------------------------------------
    # 4.3 MATRIKS KORELASI
    # --------------------------------------------------------

    with mv_tabs[2]:

        st.markdown("### Matriks korelasi")

        col1, col2 = st.columns([2, 1])

        with col1:
            mv_cv = st.multiselect(
                "Variabel numerik",
                num_all,
                default=num_all
            )

        with col2:
            mv_cm = st.selectbox(
                "Metode",
                ["pearson", "spearman"]
            )

        if "mv_cor" in globals():
            fig = mv_cor(d, mv_cv, mv_cm)
            show_plot(fig)

        st.markdown("### Interpretasi")

        if "mv_cor_text" in globals():
            show_interpretation(
                mv_cor_text(d, mv_cv, mv_cm)
            )


    # --------------------------------------------------------
    # 4.4 TREN WAKTU
    # --------------------------------------------------------

    with mv_tabs[3]:

        st.markdown("### Tren waktu")

        col1, col2, col3 = st.columns(3)

        with col1:
            mv_tm = st.selectbox(
                "Metrik",
                [
                    "Jumlah transaksi",
                    "Total penjualan (jumlah)",
                    "Rata-rata Total",
                    "Rata-rata Rating",
                    "Total Quantity"
                ]
            )

        with col2:
            mv_tu = st.selectbox(
                "Satuan waktu",
                ["Harian", "Mingguan", "Bulanan"],
                index=1
            )

        with col3:
            mv_tg = st.selectbox(
                "Kelompok",
                [NONE] + cat_base
            )

        if "mv_trend" in globals():
            fig = mv_trend(d, mv_tm, mv_tu, mv_tg)
            show_plot(fig)

        st.markdown("### Interpretasi")

        if "mv_trend_text" in globals():
            show_interpretation(
                mv_trend_text(d, mv_tm, mv_tu, mv_tg)
            )


    # --------------------------------------------------------
    # 4.5 PCA
    # --------------------------------------------------------

    with mv_tabs[4]:

        st.markdown("### Principal Component Analysis (PCA)")

        col1, col2 = st.columns([2, 1])

        with col1:
            mv_pv = st.multiselect(
                "Variabel numerik (minimal 3)",
                num_all,
                default=num_all
            )

        with col2:
            mv_pc = st.selectbox(
                "Warna titik",
                [NONE] + cat_base
            )

        col1, col2 = st.columns([7, 5])

        with col1:
            if "mv_pca" in globals():
                fig = mv_pca(d, mv_pv, mv_pc)
                show_plot(fig)

        with col2:
            st.markdown("#### Keragaman dan loading")

            if "mv_pca_table" in globals():
                show_table(mv_pca_table(d, mv_pv))

        st.markdown("### Interpretasi")

        if "mv_pca_text" in globals():
            show_interpretation(
                mv_pca_text(d, mv_pv, mv_pc)
            )


    # --------------------------------------------------------
    # 4.6 SCATTERPLOT MATRIX
    # --------------------------------------------------------

    with mv_tabs[5]:

        st.markdown("### Scatterplot matrix")

        col1, col2 = st.columns([2, 1])

        with col1:
            mv_mvars = st.multiselect(
                "Variabel numerik (minimal 2)",
                num_all,
                default=num_all
            )

        with col2:
            mv_mcol = st.selectbox(
                "Warna titik",
                [NONE] + cat_base
            )

        if "mv_pairs" in globals():
            fig = mv_pairs(d, mv_mvars, mv_mcol)
            show_plot(fig)

        st.markdown("### Interpretasi")

        if "mv_pairs_text" in globals():
            show_interpretation(
                mv_pairs_text(d, mv_mvars, mv_mcol)
            )


    # --------------------------------------------------------
    # 4.7 PENENCILAN MAHALANOBIS
    # --------------------------------------------------------

    with mv_tabs[6]:

        st.markdown("### Pencilan Mahalanobis")

        mv_dvars = st.multiselect(
            "Variabel numerik (minimal 2)",
            num_all,
            default=num_all
        )

        col1, col2 = st.columns(2)

        if "mv_mahalanobis" in globals():

            result = mv_mahalanobis(d, mv_dvars)

            with col1:
                show_plot(result.get("plot1"))

            with col2:
                show_plot(result.get("plot2"))

            st.markdown("### Sepuluh observasi dengan jarak terbesar")
            show_table(result.get("table"))

            st.markdown("### Interpretasi")
            show_interpretation(result.get("interpretation"))


    # --------------------------------------------------------
    # 4.8 Q-Q PLOT CHI-SQUARE
    # --------------------------------------------------------

    with mv_tabs[7]:

        st.markdown("### Q-Q plot Chi-Square")

        mv_qvars = st.multiselect(
            "Variabel numerik (minimal 2)",
            num_all,
            default=num_all
        )

        if "mv_qq" in globals():
            fig = mv_qq(d, mv_qvars)
            show_plot(fig)

        st.markdown("### Interpretasi")

        if "mv_qq_text" in globals():
            show_interpretation(
                mv_qq_text(d, mv_qvars)
            )


# ============================================================
# 5. REGRESI LINEAR BERGANDA
# ============================================================

with tabs[4]:

    st.markdown("### Spesifikasi model")

    col1, col2, col3 = st.columns([3, 5, 2])

    with col1:
        rg_y = st.selectbox(
            "Variabel respons (Y)",
            ["Total", "Rating"]
        )

    with col2:
        predictor_options = (
            ["Quantity", "Rating", "Hour"] + cat_all
        )

        rg_x = st.multiselect(
            "Variabel prediktor (X)",
            predictor_options,
            default=[
                x for x in
                ["Quantity", "Rating", "Hour"] + cat_base
                if x != rg_y
            ]
        )

    with col3:
        rg_log = st.radio(
            "Transformasi respons",
            ["Tanpa", "ln(Y)"]
        )

    st.caption(
        "Respons Total dipilih karena merupakan ukuran utama nilai "
        "transaksi. Pilih minimal dua prediktor. Variabel kategorik "
        "otomatis diubah menjadi dummy dengan kategori pertama "
        "sebagai acuan."
    )

    # Membangun model
    if st.button("Bangun model", key="rg_go"):

        if len(rg_x) < 2:
            st.warning("Pilih minimal dua variabel prediktor.")

        elif rg_y in rg_x:
            st.warning(
                "Variabel respons tidak boleh menjadi prediktor."
            )

        elif rg_log == "ln(Y)" and (d[rg_y] <= 0).any():
            st.warning(
                "Transformasi ln(Y) hanya dapat dilakukan jika "
                "seluruh nilai Y lebih besar dari nol."
            )

        else:
            try:
                model = build_model(
                    d,
                    rg_y,
                    rg_x,
                    rg_log == "ln(Y)"
                )

                st.session_state["rg_model"] = model
                st.session_state["rg_y"] = rg_y
                st.session_state["rg_x"] = rg_x
                st.session_state["rg_log"] = (
                    rg_log == "ln(Y)"
                )

                st.success("Model regresi berhasil dibangun.")

            except Exception as e:
                st.error(f"Gagal membangun model: {e}")


    # Mengambil model yang sudah dibangun
    model = st.session_state.get("rg_model")

    if model is not None:

        rg_y = st.session_state["rg_y"]
        rg_x = st.session_state["rg_x"]
        logy = st.session_state["rg_log"]

        rg_tabs = st.tabs([
            "Hasil model",
            "Visualisasi hasil model",
            "Galeri visualisasi regresi",
            "Diagnostik asumsi",
            "Keterbatasan model"
        ])


        # ----------------------------------------------------
        # 5.1 HASIL MODEL
        # ----------------------------------------------------

        with rg_tabs[0]:

            st.markdown("#### Persamaan model")

            st.code(
                make_equation(model, rg_y, logy),
                language="text"
            )

            col1, col2 = st.columns([5, 7])

            with col1:
                st.markdown("#### Ukuran kecocokan model")
                show_table(fit_table(model))

            with col2:
                st.markdown("#### Uji signifikansi per variabel (ANOVA)")
                show_table(anova_table(model))

            st.markdown("#### Estimasi koefisien")
            show_table(coef_table(model))

            st.caption(
                "Signif.: *** p<0.001, ** p<0.01, "
                "* p<0.05, . p<0.1"
            )

            st.markdown("#### Interpretasi hasil")
            show_interpretation(
                interp_reg(model, rg_y, logy)
            )


        # ----------------------------------------------------
        # 5.2 VISUALISASI HASIL MODEL
        # ----------------------------------------------------

        with rg_tabs[1]:

            col1, col2 = st.columns(2)

            with col1:
                show_plot(plot_reg_avp(model))

            with col2:
                show_plot(plot_reg_coef(model))

            st.markdown("### Cara membaca")

            if "reg_vis_text" in globals():
                show_interpretation(reg_vis_text(model))


        # ----------------------------------------------------
        # 5.3 GALERI VISUALISASI REGRESI
        # ----------------------------------------------------

        with rg_tabs[2]:

            plots = [
                plot_reg_resid(model),
                plot_reg_qq(model),
                plot_reg_scale(model),
                plot_reg_cook(model),
                plot_reg_scatter(d, model),
                plot_reg_cor(d, rg_x),
                plot_reg_hist(model),
                plot_reg_obsres(model)
            ]

            for i in range(0, len(plots), 2):
                col1, col2 = st.columns(2)

                with col1:
                    show_plot(plots[i])

                with col2:
                    if i + 1 < len(plots):
                        show_plot(plots[i + 1])

            st.markdown("### Cara membaca")

            if "reg_gallery_text" in globals():
                show_interpretation(
                    reg_gallery_text(model)
                )


        # ----------------------------------------------------
        # 5.4 DIAGNOSTIK ASUMSI
        # ----------------------------------------------------

        with rg_tabs[3]:

            col1, col2 = st.columns(2)

            with col1:
                show_plot(plot_reg_resid(model))

            with col2:
                show_plot(plot_reg_qq(model))

            col1, col2 = st.columns(2)

            with col1:
                show_plot(plot_reg_scale(model))

            with col2:
                show_plot(plot_reg_cook(model))

            col1, col2 = st.columns(2)

            with col1:
                st.markdown("#### Uji asumsi")

                tests = {
                    "Breusch-Pagan": bp_test(model)
                }

                show_table(
                    pd.DataFrame(tests)
                    if isinstance(tests["Breusch-Pagan"], pd.DataFrame)
                    else pd.DataFrame([tests])
                )

            with col2:
                st.markdown("#### Multikolinearitas (VIF)")
                show_table(vif_table(model))


        # ----------------------------------------------------
        # 5.5 KETERBATASAN MODEL
        # ----------------------------------------------------

        with rg_tabs[4]:

            st.markdown("### Keterbatasan model yang digunakan")
            bullets(KETERBATASAN)

    else:
        st.info(
            "Pilih variabel respons dan prediktor, kemudian klik "
            "'Bangun model' untuk melihat hasil regresi."
        )


# ============================================================
# 6. KESIMPULAN
# ============================================================

with tabs[5]:

    st.markdown("### Kesimpulan utama")

    model = st.session_state.get("rg_model")

    if model is not None:

        conclusions = auto_conclusion(
            d,
            model,
            st.session_state["rg_y"],
            st.session_state["rg_log"]
        )

        show_interpretation(conclusions)

    else:
        st.info(
            "Bangun model regresi terlebih dahulu untuk menampilkan "
            "kesimpulan otomatis."
        )

    st.markdown("### Rekomendasi")
    bullets(REKOMENDASI)

# =============================================================================
# SERVER
# =============================================================================
# ============================================================
# SERVER / LOGIKA APLIKASI STREAMLIT
# ============================================================

import streamlit as st
import pandas as pd
import numpy as np

from scipy import stats
from statsmodels.tsa.stattools import acf


# ============================================================
# 1. MEMBACA DATA
# ============================================================

@st.cache_data
def load_sales_data(file):
    if file is None:
        try:
            return read_sales(DATA_FILE)
        except Exception:
            return None

    try:
        return read_sales(file)
    except Exception:
        return None


if uploaded_file is not None:
    data_raw = load_sales_data(uploaded_file)
else:
    data_raw = d0


if data_raw is None:
    st.error(
        "Data tidak ditemukan. Letakkan Data_Penjualan.xlsx "
        "di folder aplikasi atau unggah file melalui panel kiri. "
        "Kolom wajib: Branch, Customer type, Gender, Product line, "
        "Quantity, Total, Date, Time, Payment, Rating."
    )
    st.stop()


# ============================================================
# 2. MEMPERBARUI PILIHAN FILTER
# ============================================================

def get_levels(data, column):
    if column not in data.columns:
        return []

    if isinstance(data[column].dtype, pd.CategoricalDtype):
        return data[column].cat.categories.tolist()

    return sorted(data[column].dropna().unique().tolist(), key=str)


branch_levels = get_levels(data_raw, "Branch")
cust_levels = get_levels(data_raw, "Customer type")
gender_levels = get_levels(data_raw, "Gender")
pay_levels = get_levels(data_raw, "Payment")
prod_levels = get_levels(data_raw, "Product line")


# ============================================================
# 3. RESET FILTER
# ============================================================

if "reset_filter" not in st.session_state:
    st.session_state["reset_filter"] = False

if st.sidebar.button("Atur ulang filter"):
    st.session_state["reset_filter"] = True
    st.rerun()


# ============================================================
# 4. MENGATUR FILTER DATA
# ============================================================

def filter_options(key, levels):
    if st.session_state.get("reset_filter", False):
        return levels

    current = st.session_state.get(key)

    if current is None:
        return levels

    return [x for x in current if x in levels]


f_branch = filter_options("f_branch", branch_levels)
f_cust = filter_options("f_cust", cust_levels)
f_gender = filter_options("f_gender", gender_levels)
f_pay = filter_options("f_pay", pay_levels)
f_prod = filter_options("f_prod", prod_levels)

if "Date" in data_raw.columns:
    date_min = pd.to_datetime(data_raw["Date"]).min().date()
    date_max = pd.to_datetime(data_raw["Date"]).max().date()
else:
    date_min = None
    date_max = None


# Mengambil filter dari sidebar
with st.sidebar:

    st.markdown("### Filter data")

    f_branch = st.multiselect(
        "Branch",
        options=branch_levels,
        default=f_branch,
        key="filter_branch"
    )

    f_cust = st.multiselect(
        "Customer type",
        options=cust_levels,
        default=f_cust,
        key="filter_customer"
    )

    f_gender = st.multiselect(
        "Gender",
        options=gender_levels,
        default=f_gender,
        key="filter_gender"
    )

    f_pay = st.multiselect(
        "Payment",
        options=pay_levels,
        default=f_pay,
        key="filter_payment"
    )

    f_prod = st.multiselect(
        "Product line",
        options=prod_levels,
        default=f_prod,
        key="filter_product"
    )

    if date_min is not None:
        f_date = st.date_input(
            "Rentang tanggal",
            value=(date_min, date_max),
            min_value=date_min,
            max_value=date_max,
            format="DD-MM-YYYY"
        )
    else:
        f_date = None


st.session_state["reset_filter"] = False


# ============================================================
# 5. DATA SETELAH FILTER
# ============================================================

def apply_filters(
    data,
    branch,
    customer,
    gender,
    payment,
    product,
    date_range
):
    result = data.copy()

    result = result[
        result["Branch"].isin(branch)
        & result["Customer type"].isin(customer)
        & result["Gender"].isin(gender)
        & result["Payment"].isin(payment)
        & result["Product line"].isin(product)
    ]

    if date_range is not None and len(date_range) == 2:
        result["Date"] = pd.to_datetime(result["Date"])

        result = result[
            (result["Date"].dt.date >= date_range[0])
            & (result["Date"].dt.date <= date_range[1])
        ]

    return result.reset_index(drop=True)


data_use = apply_filters(
    data_raw,
    f_branch,
    f_cust,
    f_gender,
    f_pay,
    f_prod,
    f_date
)


# Validasi jumlah data
if len(data_use) < 30:
    st.error(
        "Data hasil filter terlalu sedikit (kurang dari 30 baris). "
        "Longgarkan filter pada panel kiri."
    )
    st.stop()


data_use = data_use.copy()


# ============================================================
# 6. INFORMASI JUMLAH DATA TERPILIH
# ============================================================

n0 = len(data_raw)
n1 = len(data_use)

percentage = 100 * n1 / n0 if n0 > 0 else 0

st.sidebar.markdown("---")

st.sidebar.markdown(
    f"""
    <div class="nfilter">
        Transaksi terpilih:<br>
        <b>{n1:,}</b> dari {n0:,}
        ({percentage:.1f}%)
    </div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# 7. TAB 1: MASALAH DAN DATA
# ============================================================

with tabs[0]:

    st.markdown("### Informasi data")

    period_start = pd.to_datetime(data_raw["Date"]).min()
    period_end = pd.to_datetime(data_raw["Date"]).max()

    data_info = [
        (
            f"Jumlah observasi: **{n0:,}** transaksi; "
            "jumlah variabel: **10** variabel asli "
            "(ditambah 3 variabel turunan: Hour, Bulan, Hari)."
        ),
        (
            f"Periode data: **{period_start:%d-%m-%Y}** "
            f"sampai **{period_end:%d-%m-%Y}**."
        ),
        (
            "Variabel kategorik: Branch, Customer type, Gender, "
            "Product line, Payment. Variabel numerik: Quantity, "
            "Total, Rating, serta Hour (turunan)."
        )
    ]

    show_interpretation(data_info)

    st.markdown("### Definisi variabel")

    show_table(kamus)

    st.markdown("### Pengolahan data")

    q = data_raw["Total"].quantile([0.25, 0.75])
    iqr = q.loc[0.75] - q.loc[0.25]

    lower = q.loc[0.25] - 1.5 * iqr
    upper = q.loc[0.75] + 1.5 * iqr

    outliers = (
        (data_raw["Total"] < lower)
        | (data_raw["Total"] > upper)
    ).sum()

    missing = int(data_raw.isna().sum().sum())
    duplicates = int(data_raw.duplicated().sum())

    prep_info = [
        f"**Data hilang:** {missing} sel kosong pada seluruh kolom analisis.",
        f"**Data ganda:** {duplicates} baris duplikat.",
        (
            f"**Pencilan:** terdapat {outliers} pencilan pada Total "
            "(aturan 1.5 x IQR). Pencilan dipertahankan karena merupakan "
            "transaksi bernilai besar yang valid."
        ),
        (
            "**Pengolahan:** variabel Hour, Bulan, dan Hari dibentuk "
            "dari Time dan Date; variabel kategorik diubah menjadi faktor."
        ),
        (
            "**Transformasi:** tidak dilakukan pada eksplorasi; pada "
            "regresi tersedia pilihan transformasi ln(Y) untuk mengatasi "
            "ragam yang tidak konstan."
        )
    ]

    show_interpretation(prep_info)

    st.markdown("### Cuplikan data")

    preview = data_raw.head(10).copy()
    preview["Date"] = preview["Date"].astype(str)

    show_table(preview)


# ============================================================
# 8. TAB 2: UNIVARIAT
# ============================================================

with tabs[1]:

    uni_type = uni_choices(uni_var)

    if uni_var not in data_use.columns:
        st.warning("Variabel tidak tersedia.")
    else:

        if "uni_type_selected" not in st.session_state:
            st.session_state["uni_type_selected"] = uni_type[0]

        selected_type = st.selectbox(
            "Jenis grafik",
            options=uni_type,
            key="uni_type_selected"
        )

        if selected_type == "Stem-and-leaf":

            x = data_use[uni_var]

            if pd.api.types.is_numeric_dtype(x):
                st.text("Stem-and-leaf plot: " + uni_var)

                values = x.dropna().to_numpy()

                if len(values) > 0:
                    stem_values = np.floor(values / 10).astype(int)
                    leaf_values = np.abs(values.astype(int)) % 10

                    stem_df = pd.DataFrame({
                        "Stem": stem_values,
                        "Leaf": leaf_values
                    })

                    for stem, group in stem_df.groupby("Stem"):
                        leaves = " ".join(
                            map(str, sorted(group["Leaf"].tolist()))
                        )
                        st.text(f"{stem} | {leaves}")

            else:
                st.warning(
                    "Stem-and-leaf hanya dapat digunakan "
                    "untuk variabel numerik."
                )

        else:

            st.markdown("### Visualisasi")

            fig = plot_uni(
                data_use,
                uni_var,
                selected_type
            )

            show_plot(fig)

        col1, col2 = st.columns([7, 5])

        with col2:
            st.markdown("### Statistik deskriptif")
            show_table(uni_table(data_use, uni_var))

        st.markdown("### Interpretasi")

        show_interpretation(
            interp_uni(data_use, uni_var)
        )


# ============================================================
# 9. TAB 3: BIVARIAT
# ============================================================

with tabs[2]:

    if bi_x == bi_y:

        st.warning("Pilih dua variabel yang berbeda.")

    else:

        selected_bi_type = bi_choices(bi_x, bi_y)

        bi_type_selected = st.selectbox(
            "Jenis grafik",
            options=selected_bi_type,
            key="bi_type_selected"
        )

        st.markdown("### Visualisasi")

        fig = plot_bi(
            data_use,
            bi_x,
            bi_y,
            bi_type_selected
        )

        show_plot(fig)

        bi_res = bi_text(
            data_use,
            bi_x,
            bi_y
        )

        col1, col2 = st.columns([7, 5])

        with col2:
            st.markdown("### Ringkasan dan uji")
            show_table(bi_res["table"])

        st.markdown("### Interpretasi")

        show_interpretation(bi_res["lines"])


# ============================================================
# 10. TAB 4: MULTIVARIAT
# ============================================================

with tabs[3]:

    # --------------------------------------------------------
    # 10.1 SCATTER MULTIVARIAT
    # --------------------------------------------------------

    with mv_tabs[0]:

        if mv_x == mv_y:

            st.warning("Pilih sumbu X dan Y yang berbeda.")

        else:

            fig = plot_mv_scatter(
                data_use,
                mv_x,
                mv_y,
                mv_col,
                mv_size,
                mv_facet,
                mv_lines
            )

            show_plot(fig)

            d_mv = data_use
            r = d_mv[mv_x].corr(d_mv[mv_y])

            ln = [
                (
                    f"Secara keseluruhan korelasi **{mv_x}** dan "
                    f"**{mv_y}** adalah **{fmt(r, 3)}** "
                    f"({strength(r)})."
                )
            ]

            if (
                mv_col != NONE
                and mv_col in d_mv.columns
                and not pd.api.types.is_numeric_dtype(d_mv[mv_col])
            ):

                grouped = (
                    d_mv.groupby(mv_col, observed=False)
                    .apply(
                        lambda g: g[mv_x].corr(g[mv_y]),
                        include_groups=False
                    )
                    .dropna()
                )

                group_text = "; ".join(
                    f"{k} = {fmt(v, 2)}"
                    for k, v in grouped.items()
                )

                ln.append(
                    f"Korelasi per **{mv_col}**: {group_text}."
                )

                if len(grouped) > 1:

                    if grouped.max() - grouped.min() < 0.1:
                        ln.append(
                            "Pola hubungan hampir sama di semua kelompok "
                            "sehingga variabel warna tidak mengubah "
                            "hubungan X dan Y."
                        )
                    else:
                        ln.append(
                            "Kekuatan hubungan berbeda antar kelompok, "
                            "menandakan kemungkinan interaksi."
                        )

            if mv_size != NONE:
                ln.append(
                    f"Ukuran titik menunjukkan **{mv_size}**; "
                    "jika titik besar dan kecil tersebar acak, "
                    "variabel tersebut tidak berkaitan dengan "
                    "posisi titik pada X dan Y."
                )

            show_interpretation(ln)


    # --------------------------------------------------------
    # 10.2 HEATMAP RATA-RATA
    # --------------------------------------------------------

    with mv_tabs[1]:

        if mv_ha == mv_hb:

            st.warning(
                "Pilih dua variabel kategorik yang berbeda."
            )

        else:

            fig = plot_mv_heat(
                data_use,
                mv_ha,
                mv_hb,
                mv_hn,
                mv_hf
            )

            show_plot(fig)

            result = mv_heat_data(
                data_use,
                mv_ha,
                mv_hb,
                mv_hn,
                mv_hf
            )

            hi = result.loc[result["val"].idxmax()]
            lo = result.loc[result["val"].idxmin()]

            difference = hi["val"] - lo["val"]

            percentage = (
                100 * difference / hi["val"]
                if hi["val"] != 0 else np.nan
            )

            show_interpretation([
                (
                    f"Nilai tertinggi pada kombinasi "
                    f"**{hi['A']} - {hi['B']}** "
                    f"({fmt(hi['val'], 1)}) dan terendah pada "
                    f"**{lo['A']} - {lo['B']}** "
                    f"({fmt(lo['val'], 1)})."
                ),
                (
                    f"Selisih tertinggi-terendah sebesar "
                    f"**{fmt(difference, 1)}** "
                    f"({fmt(percentage, 1)}% dari nilai tertinggi)."
                ),
                (
                    "Warna lebih pekat menandakan nilai lebih besar. "
                    "Perbedaan antar sel perlu dikonfirmasi dengan "
                    "uji statistik sebelum disimpulkan sebagai "
                    "perbedaan nyata."
                )
            ])


    # --------------------------------------------------------
    # 10.3 MATRIKS KORELASI
    # --------------------------------------------------------

    with mv_tabs[2]:

        if len(mv_cv) < 2:

            st.warning("Pilih minimal dua variabel.")

        else:

            fig = plot_mv_cor(
                data_use,
                mv_cv,
                mv_cm
            )

            show_plot(fig)

            corr = data_use[mv_cv].corr(method=mv_cm)

            upper = corr.where(
                np.triu(
                    np.ones(corr.shape),
                    k=1
                ).astype(bool)
            )

            corr_pairs = (
                upper.stack()
                .rename("Freq")
                .reset_index()
            )

            corr_pairs.columns = [
                "Var1", "Var2", "Freq"
            ]

            corr_pairs["abs"] = corr_pairs["Freq"].abs()

            corr_pairs = corr_pairs.sort_values(
                "abs",
                ascending=False
            )

            if len(corr_pairs) > 0:

                strongest = corr_pairs.iloc[0]
                weakest = corr_pairs.iloc[-1]

                show_interpretation([
                    (
                        f"Korelasi terkuat: **{strongest['Var1']} - "
                        f"{strongest['Var2']}** "
                        f"(r = {fmt(strongest['Freq'], 3)}; "
                        f"{strength(strongest['Freq'])})."
                    ),
                    (
                        f"Korelasi terlemah: **{weakest['Var1']} - "
                        f"{weakest['Var2']}** "
                        f"(r = {fmt(weakest['Freq'], 3)})."
                    ),
                    (
                        "Warna merah muda tua = korelasi positif, "
                        "biru keabuan = negatif, putih = mendekati nol."
                    )
                ])


    # --------------------------------------------------------
    # 10.4 TREN WAKTU
    # --------------------------------------------------------

    with mv_tabs[3]:

        fig = plot_mv_trend(
            data_use,
            mv_tm,
            mv_tu,
            mv_tg
        )

        show_plot(fig)

        td = trend_data(
            data_use,
            mv_tm,
            mv_tu,
            mv_tg
        )

        if len(td) > 0:

            hi = td.loc[td["val"].idxmax()]
            lo = td.loc[td["val"].idxmin()]

            group_text = ""

            if mv_tg != NONE:
                group_text = (
                    f" (kelompok {hi[mv_tg]})"
                )

            trend_lines = [
                (
                    f"Nilai **{mv_tm}** tertinggi pada periode "
                    f"**{hi['Periode']}**{group_text} sebesar "
                    f"{fmt(hi['val'], 1)}; terendah pada "
                    f"**{lo['Periode']}** sebesar "
                    f"{fmt(lo['val'], 1)}."
                ),
                (
                    "Bandingkan garis antar kelompok untuk melihat "
                    "apakah pola naik-turun sama."
                )
            ]

            if mv_tu == "Mingguan":
                trend_lines.append(
                    "Minggu pertama dan terakhir dapat berisi "
                    "lebih sedikit hari sehingga nilainya "
                    "tampak lebih rendah."
                )
            else:
                trend_lines.append(
                    "Data hanya mencakup beberapa bulan, sehingga "
                    "pola musiman tidak dapat disimpulkan."
                )

            show_interpretation(trend_lines)


    # --------------------------------------------------------
    # 10.5 PCA
    # --------------------------------------------------------

    with mv_tabs[4]:

        if len(mv_pv) < 3:

            st.warning(
                "Pilih minimal tiga variabel numerik."
            )

        else:

            fig = plot_mv_pca(
                data_use,
                mv_pv,
                mv_pc
            )

            show_plot(fig)

            pr = pca_fit(
                data_use,
                mv_pv
            )

            ve = (
                100 * pr.sdev ** 2
                / np.sum(pr.sdev ** 2)
            )

            loadings = pd.DataFrame(
                pr.rotation[:, :2],
                index=mv_pv,
                columns=["PC1", "PC2"]
            ).round(3)

            loading_table = loadings.reset_index()
            loading_table.columns = [
                "Variabel", "PC1", "PC2"
            ]

            st.markdown("### Keragaman dan loading")
            show_table(loading_table)

            l1 = pr.rotation[:, 0]
            max_index = np.argmax(np.abs(l1))

            show_interpretation([
                (
                    f"PC1 menjelaskan **{fmt(ve[0], 1)}%** "
                    f"dan PC2 **{fmt(ve[1], 1)}%** keragaman; "
                    f"bersama-sama **{fmt(ve[0] + ve[1], 1)}%**."
                ),
                (
                    f"Variabel dengan kontribusi terbesar pada PC1: "
                    f"**{mv_pv[max_index]}** "
                    f"(loading {fmt(l1[max_index], 2)})."
                ),
                (
                    "Panah yang searah menandakan variabel berkorelasi "
                    "positif; panah yang saling tegak lurus menandakan "
                    "variabel hampir tidak berkorelasi."
                )
            ])


    # --------------------------------------------------------
    # 10.6 SCATTERPLOT MATRIX
    # --------------------------------------------------------

    with mv_tabs[5]:

        if len(mv_mvars) < 2:

            st.warning(
                "Pilih minimal dua variabel numerik."
            )

        else:

            fig = plot_mv_pairs(
                data_use,
                mv_mvars,
                mv_mcol
            )

            show_plot(fig)

            corr = data_use[mv_mvars].corr()

            upper = corr.where(
                np.triu(
                    np.ones(corr.shape),
                    k=1
                ).astype(bool)
            )

            pairs = (
                upper.stack()
                .rename("Freq")
                .reset_index()
            )

            pairs.columns = [
                "Var1", "Var2", "Freq"
            ]

            pairs["abs"] = pairs["Freq"].abs()
            pairs = pairs.sort_values(
                "abs",
                ascending=False
            )

            if len(pairs) > 0:

                strongest = pairs.iloc[0]
                weakest = pairs.iloc[-1]

                show_interpretation([
                    (
                        "Diagonal: histogram tiap variabel; "
                        "segitiga bawah: scatterplot dengan kurva "
                        "lowess; segitiga atas: koefisien korelasi "
                        "Pearson."
                    ),
                    (
                        f"Pasangan dengan hubungan terkuat: "
                        f"**{strongest['Var1']} - {strongest['Var2']}** "
                        f"(r = {fmt(strongest['Freq'], 3)}; "
                        f"{strength(strongest['Freq'])}). "
                        f"Pasangan terlemah: **{weakest['Var1']} - "
                        f"{weakest['Var2']}** "
                        f"(r = {fmt(weakest['Freq'], 3)})."
                    ),
                    (
                        "Kurva lowess yang mendekati garis lurus "
                        "menandakan hubungan linear; kurva melengkung "
                        "menandakan hubungan nonlinear."
                    )
                ])


    # --------------------------------------------------------
    # 10.7 PENENCILAN MAHALANOBIS
    # --------------------------------------------------------

    with mv_tabs[6]:

        if len(mv_dvars) < 2:

            st.warning(
                "Pilih minimal dua variabel numerik."
            )

        else:

            mh_result = mahal_fit(
                data_use,
                mv_dvars
            )

            col1, col2 = st.columns(2)

            with col1:
                show_plot(
                    plot_mv_mahal_idx(mh_result)
                )

            with col2:
                show_plot(
                    plot_mv_mahal_sc(
                        mh_result,
                        mv_dvars
                    )
                )

            st.markdown(
                "### Sepuluh observasi dengan jarak terbesar"
            )

            show_table(
                mahal_table(
                    data_use,
                    mh_result,
                    mv_dvars
                )
            )

            f = mh_result
            n = len(f["df"])
            k = int(f["df"]["out"].sum())

            top = f["df"].loc[
                f["df"]["D2"].idxmax()
            ]

            z = {
                v: (
                    data_use[v].iloc[int(top["idx"]) - 1]
                    - data_use[v].mean()
                ) / data_use[v].std()
                for v in mv_dvars
            }

            vmax = max(z, key=lambda x: abs(z[x]))

            show_interpretation([
                (
                    f"Batas pencilan multivariat: Chi-Square 97.5% "
                    f"dengan df = {f['p']}, yaitu "
                    f"**{fmt(f['cut'], 2)}**."
                ),
                (
                    f"Ditemukan **{k} pencilan** "
                    f"(**{fmt(100 * k / n, 1)}%** dari "
                    f"{n} observasi)."
                ),
                (
                    f"Observasi paling ekstrem: baris "
                    f"**{int(top['idx'])}** "
                    f"(D2 = {fmt(top['D2'], 2)}). "
                    f"Variabel yang paling menyimpang: "
                    f"**{vmax}** (z = {fmt(z[vmax], 2)})."
                ),
                (
                    "Pencilan multivariat dapat berupa kombinasi "
                    "nilai yang tidak lazim walau tiap variabel "
                    "tampak wajar. Pencilan sebaiknya diperiksa "
                    "sebelum dihapus."
                )
            ])


    # --------------------------------------------------------
    # 10.8 Q-Q PLOT CHI-SQUARE
    # --------------------------------------------------------

    with mv_tabs[7]:

        if len(mv_qvars) < 2:

            st.warning(
                "Pilih minimal dua variabel numerik."
            )

        else:

            qq_result = mahal_fit(
                data_use,
                mv_qvars
            )

            fig = plot_mv_qqchi(qq_result)
            show_plot(fig)

            f = qq_result
            n = len(f["df"])

            sorted_d2 = np.sort(f["df"]["D2"].to_numpy())

            theoretical_q = stats.chi2.ppf(
                (np.arange(1, n + 1) - 0.5) / n,
                df=f["p"]
            )

            r = np.corrcoef(
                sorted_d2,
                theoretical_q
            )[0, 1]

            extreme_rows = (
                f["df"]
                .sort_values("D2")
                .tail(3)["idx"]
                .tolist()
            )

            show_interpretation([
                (
                    f"Korelasi antara D2 terurut dan kuantil "
                    f"Chi-Square = **{fmt(r, 3)}**; semakin "
                    "mendekati 1, semakin sesuai data dengan "
                    "normal multivariat."
                ),
                (
                    "Titik terjauh dari garis berada di ujung kanan "
                    f"(baris {', '.join(map(str, extreme_rows))})."
                ),
                (
                    "Titik yang melengkung ke atas di ujung kanan "
                    "menandakan ekor distribusi lebih berat atau "
                    "adanya pencilan. Asumsi normalitas multivariat "
                    "kurang terpenuhi jika penyimpangan terlihat jelas."
                )
            ])


# ============================================================
# 11. TAB 5: REGRESI LINEAR BERGANDA
# ============================================================

with tabs[4]:

    if st.button(
        "Bangun model",
        key="run_regression"
    ):

        preds = [
            x for x in rg_x
            if x != rg_y
        ]

        if len(preds) < 2:

            st.warning(
                "Pilih minimal dua variabel prediktor "
                "(selain variabel respons)."
            )

        elif rg_log == "ln(Y)" and not (
            data_use[rg_y] > 0
        ).all():

            st.warning(
                "Transformasi ln hanya untuk respons "
                "bernilai positif."
            )

        else:

            try:

                regression_model = build_model(
                    data_use,
                    rg_y,
                    preds,
                    rg_log == "ln(Y)"
                )

                st.session_state["regression_model"] = regression_model
                st.session_state["regression_data"] = data_use.copy()
                st.session_state["regression_resp"] = rg_y
                st.session_state["regression_preds"] = preds
                st.session_state["regression_log"] = (
                    rg_log == "ln(Y)"
                )

                st.success(
                    "Model regresi berhasil dibangun."
                )

            except Exception as e:

                st.error(
                    f"Gagal membangun model: {e}"
                )


    model = st.session_state.get(
        "regression_model"
    )

    if model is None:

        st.info(
            "Bangun model regresi terlebih dahulu "
            "untuk melihat hasil analisis."
        )

    else:

        m = model
        resp = st.session_state["regression_resp"]
        preds = st.session_state["regression_preds"]
        logy = st.session_state["regression_log"]
        reg_data = st.session_state["regression_data"]

        reg_tabs = st.tabs([
            "Hasil model",
            "Visualisasi hasil model",
            "Galeri visualisasi regresi",
            "Diagnostik asumsi",
            "Keterbatasan model"
        ])


        # ----------------------------------------------------
        # 11.1 HASIL MODEL
        # ----------------------------------------------------

        with reg_tabs[0]:

            st.markdown("### Persamaan model")

            st.code(
                make_equation(m, resp, logy),
                language="text"
            )

            col1, col2 = st.columns([5, 7])

            with col1:

                st.markdown(
                    "#### Ukuran kecocokan model"
                )

                show_table(
                    fit_table(m)
                )

            with col2:

                st.markdown(
                    "#### Uji signifikansi per variabel (ANOVA)"
                )

                show_table(
                    anova_table(m)
                )

            st.markdown("### Estimasi koefisien")

            show_table(
                coef_table(m)
            )

            st.caption(
                "Signif.: *** p<0.001, ** p<0.01, "
                "* p<0.05, . p<0.1"
            )

            st.markdown("### Interpretasi hasil")

            show_interpretation(
                interp_reg(m, resp, logy)
            )


        # ----------------------------------------------------
        # 11.2 VISUALISASI HASIL MODEL
        # ----------------------------------------------------

        with reg_tabs[1]:

            col1, col2 = st.columns(2)

            with col1:

                show_plot(
                    plot_reg_avp(m, logy, resp)
                )

            with col2:

                show_plot(
                    plot_reg_coef(m)
                )

            st.markdown("### Cara membaca")

            r2 = m.rsquared

            show_interpretation([
                (
                    f"**Aktual vs prediksi:** semakin rapat titik "
                    f"di sekitar garis putus-putus, semakin baik "
                    f"prediksi model (R2 = {fmt(r2, 3)})."
                ),
                (
                    "**Koefisien dan selang kepercayaan:** titik "
                    "merupakan estimasi, garis merupakan selang 95%. "
                    "Selang yang tidak memotong garis nol menandakan "
                    "prediktor signifikan pada taraf 5%."
                )
            ])


        # ----------------------------------------------------
        # 11.3 GALERI VISUALISASI REGRESI
        # ----------------------------------------------------

        with reg_tabs[2]:

            plots = [
                plot_reg_scatter(
                    reg_data, resp, preds, logy
                ),
                plot_reg_cor(
                    reg_data, resp, preds
                ),
                plot_reg_coef(m),
                plot_reg_qq(m),
                plot_reg_hist(m),
                plot_reg_ts(m, reg_data),
                plot_reg_acf(m, reg_data),
                plot_reg_obsres(m)
            ]

            for i in range(0, len(plots), 2):

                col1, col2 = st.columns(2)

                with col1:
                    show_plot(plots[i])

                with col2:
                    show_plot(plots[i + 1])

            st.markdown("### Cara membaca")

            ordered = reg_ord_res(
                m,
                reg_data
            )

            residuals = np.asarray(ordered["res"])
            n = len(residuals)

            acf_values = acf(
                residuals,
                nlags=min(30, n - 1),
                fft=False
            )

            ci = stats.norm.ppf(0.975) / np.sqrt(n)

            significant_lags = int(
                np.sum(
                    np.abs(acf_values[1:]) > ci
                )
            )

            dw = (
                np.sum(np.diff(residuals) ** 2)
                / np.sum(residuals ** 2)
            )

            sample_resid = (
                np.random.default_rng(123)
                .choice(
                    residuals,
                    size=5000,
                    replace=False
                )
                if n > 5000
                else residuals
            )

            sw = stats.shapiro(sample_resid).pvalue

            show_interpretation([
                (
                    "**Scatterplot:** menunjukkan bentuk hubungan "
                    "respons dengan tiap prediktor numerik."
                ),
                (
                    "**Correlation matrix:** korelasi respons dengan "
                    "prediktor menunjukkan kandidat prediktor kuat; "
                    "korelasi tinggi antar prediktor mengindikasikan "
                    "potensi multikolinearitas."
                ),
                (
                    "**Coefficient CI plot:** selang 95% yang tidak "
                    "memotong garis nol berarti koefisien signifikan "
                    "pada taraf 5%."
                ),
                (
                    f"**Normal probability plot dan histogram residual:** "
                    f"hasil Shapiro-Wilk p {pq(sw)}, sehingga residual "
                    f"{'**tidak normal**' if sw < 0.05 else 'normal'}."
                ),
                (
                    f"**Time series plot dan ACF:** statistik "
                    f"Durbin-Watson = {fmt(dw, 3)}; "
                    f"{significant_lags} dari {len(acf_values) - 1} "
                    "lag berada di luar batas 95%. "
                    f"{'Terdapat indikasi autokorelasi residual.' if significant_lags > 0 or abs(dw - 2) >= 0.4 else 'Tidak ada indikasi autokorelasi residual yang berarti.'}"
                ),
                (
                    "**Observed vs residual:** residual idealnya "
                    "tersebar acak di sekitar nol. Sebaran yang "
                    "melebar pada nilai observasi besar menandakan "
                    "ragam residual tidak konstan."
                )
            ])


        # ----------------------------------------------------
        # 11.4 DIAGNOSTIK ASUMSI
        # ----------------------------------------------------

        with reg_tabs[3]:

            col1, col2 = st.columns(2)

            with col1:
                show_plot(plot_reg_resid(m))

            with col2:
                show_plot(plot_reg_qq(m))

            col1, col2 = st.columns(2)

            with col1:
                show_plot(plot_reg_scale(m))

            with col2:
                show_plot(plot_reg_cook(m))

            # Uji Shapiro-Wilk
            residuals = np.asarray(m.resid)

            sample_resid = (
                np.random.default_rng(123)
                .choice(
                    residuals,
                    size=5000,
                    replace=False
                )
                if len(residuals) > 5000
                else residuals
            )

            sh = stats.shapiro(sample_resid)

            # Uji Breusch-Pagan
            bp = bp_test(m)

            if isinstance(bp, dict):
                bp_stat = bp.get("stat")
                bp_p = bp.get("p")
            elif isinstance(bp, pd.Series):
                bp_stat = bp.get("stat")
                bp_p = bp.get("p")
            else:
                bp_stat = np.nan
                bp_p = np.nan

            # Cook's distance
            influence = m.get_influence()
            cooks_d = influence.cooks_distance[0]

            cook_count = int(
                np.sum(cooks_d > 4 / len(cooks_d))
            )

            st.markdown("### Uji asumsi")

            tests = pd.DataFrame({
                "Uji": [
                    "Shapiro-Wilk (normalitas residual)",
                    "Breusch-Pagan (kesamaan ragam)",
                    "Observasi berpengaruh (Cook's D > 4/n)"
                ],

                "Statistik": [
                    fmt(sh.statistic, 4),
                    fmt(bp_stat, 3),
                    str(cook_count)
                ],

                "p-value": [
                    fmt_p(sh.pvalue),
                    fmt_p(bp_p),
                    "-"
                ],

                "Keputusan": [
                    (
                        "Tidak normal"
                        if sh.pvalue < 0.05
                        else "Normal"
                    ),
                    (
                        "Ragam tidak konstan"
                        if pd.notna(bp_p) and bp_p < 0.05
                        else "Ragam konstan"
                    ),
                    f"{cook_count} observasi"
                ]
            })

            show_table(tests)

            st.markdown("### Multikolinearitas (VIF)")

            v = vif_table(m)

            if v is not None:

                v = v.copy()

                if "VIF setara" in v.columns:

                    v["Keputusan"] = np.where(
                        pd.to_numeric(
                            v["VIF setara"],
                            errors="coerce"
                        ) < 5,
                        "Aman",
                        "Perlu perhatian"
                    )

                show_table(v)

            else:

                st.info(
                    "VIF membutuhkan minimal dua prediktor."
                )


        # ----------------------------------------------------
        # 11.5 KETERBATASAN MODEL
        # ----------------------------------------------------

        with reg_tabs[4]:

            st.markdown(
                "### Keterbatasan model yang digunakan"
            )

            bullets(KETERBATASAN)


# ============================================================
# 12. TAB 6: KESIMPULAN
# ============================================================

with tabs[5]:

    st.markdown("### Kesimpulan utama")

    model = st.session_state.get(
        "regression_model"
    )

    if model is not None:

        resp = st.session_state["regression_resp"]
        logy = st.session_state["regression_log"]

        conclusions = auto_conclusion(
            data_use,
            model,
            resp,
            logy
        )

        show_interpretation(conclusions)

    else:

        st.info(
            "Bangun model regresi terlebih dahulu untuk "
            "menampilkan kesimpulan otomatis."
        )

    st.markdown("### Rekomendasi")

    bullets(REKOMENDASI)
