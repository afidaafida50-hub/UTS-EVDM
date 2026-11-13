library(shiny)
library(ggplot2)
library(dplyr)
library(readxl)

# ---- KONFIGURASI: EDIT BAGIAN INI -------------------------------------------
DATA_FILE   <- "Data_Penjualan.xlsx"
NAMA        <- "Afida Mar'atus Solihah"
NIM         <- "231064005"
APP_TITLE   <- "Analisis Faktor yang Memengaruhi Nilai Transaksi Penjualan Supermarket"
SUMBER_DATA <- "Dataset sekunder data transaksi supermarket (Data_Penjualan.xlsx). Isi dengan tautan/sumber asli yang dapat dipertanggungjawabkan (mis. Kaggle: Supermarket Sales)."
LATAR_BELAKANG <- "Supermarket perlu memahami faktor yang berkaitan dengan besarnya nilai transaksi agar strategi promosi, penataan produk, dan pelayanan dapat dirancang berdasarkan data. Nilai transaksi dapat dipengaruhi oleh jumlah barang yang dibeli, karakteristik pelanggan, lini produk, cabang, metode pembayaran, dan waktu transaksi. Analisis eksplorasi dan regresi linear berganda digunakan untuk mengidentifikasi faktor-faktor tersebut."
TUJUAN <- c(
  "Menggambarkan karakteristik setiap variabel pada data transaksi (univariat).",
  "Mengidentifikasi hubungan dan perbedaan kelompok antar dua variabel (bivariat).",
  "Mengeksplorasi pola yang melibatkan tiga variabel atau lebih (multivariat).",
  "Membangun model regresi linear berganda untuk menjelaskan nilai transaksi (Total) dan mengevaluasi kemampuannya.",
  "Merumuskan kesimpulan dan rekomendasi bagi pengelola supermarket."
)

# ---- Konstanta variabel ------------------------------------------------------
cat_base <- c("Branch", "Customer type", "Gender", "Product line", "Payment")
cat_all  <- c(cat_base, "Bulan", "Hari")
num_all  <- c("Quantity", "Total", "Rating", "Hour")
NONE <- "(tidak ada)"
bulan_lv <- c("Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus",
              "September","Oktober","November","Desember")
hari_lv  <- c("Senin","Selasa","Rabu","Kamis","Jumat","Sabtu","Minggu")

kamus <- data.frame(
  Variabel = c("Branch","Customer type","Gender","Product line","Quantity","Total",
               "Date","Time","Payment","Rating","Hour (turunan)","Bulan (turunan)","Hari (turunan)"),
  Tipe = c("Kategorik","Kategorik","Kategorik","Kategorik","Numerik (diskret)","Numerik (kontinu)",
           "Tanggal","Waktu","Kategorik","Numerik (kontinu)","Numerik (diskret)","Kategorik","Kategorik"),
  Definisi = c("Cabang supermarket tempat transaksi (A, B, C).",
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
               "Hari dalam seminggu yang diambil dari Date."),
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

sel_cat  <- function(id, lab, sel) selectInput(id, lab, choices = cat_all, selected = sel)
d0 <- tryCatch(read_sales(DATA_FILE), error = function(e) NULL)
ch <- function(v, def) if (is.null(d0)) def else levels(d0[[v]])

ui <- fluidPage(
  tags$head(tags$meta(name = "viewport", content = "width=device-width, initial-scale=1"), tags$style(HTML(css))),
  div(class = "app-header", h1(APP_TITLE), p(class = "sub", "Ujian Tengah Semester Eksplorasi dan Visualisasi Data Multivariat  |  ", NAMA, "  |  ", NIM)),
  sidebarLayout(
    sidebarPanel(width = 3,
                 h4("Sumber data"),
                 fileInput("file", NULL, buttonLabel = "Pilih file", placeholder = "Excel/CSV (opsional)", accept = c(".xlsx", ".xls", ".csv")),
                 h4("Filter data"),
                 checkboxGroupInput("f_branch", "Branch", choices = ch("Branch", c("A","B","C")), selected = ch("Branch", c("A","B","C")), inline = TRUE),
                 checkboxGroupInput("f_cust", "Customer type", choices = ch("Customer type", c("Member","Normal")), selected = ch("Customer type", c("Member","Normal")), inline = TRUE),
                 checkboxGroupInput("f_gender", "Gender", choices = ch("Gender", c("Female","Male")), selected = ch("Gender", c("Female","Male")), inline = TRUE),
                 checkboxGroupInput("f_pay", "Payment", choices = ch("Payment", c("Cash","Credit card","Ewallet")), selected = ch("Payment", c("Cash","Credit card","Ewallet"))),
                 selectInput("f_prod", "Product line", choices = ch("Product line", NULL), selected = ch("Product line", NULL), multiple = TRUE),
                 dateRangeInput("f_date", "Rentang tanggal", start = if (is.null(d0)) NULL else min(d0$Date), end = if (is.null(d0)) NULL else max(d0$Date), format = "dd-mm-yyyy", language = "id"),
                 actionButton("f_reset", "Atur ulang filter", class = "btn-default btn-sm"),
                 uiOutput("n_info")
    ),
    mainPanel(width = 9,
              tabsetPanel(id = "tabs",
                          # ------------------------------------------------------------ 1. DATA
                          tabPanel("1. Masalah dan Data",
                                   br(),
                                   div(class = "box-x", h4("Latar belakang"), p(LATAR_BELAKANG)),
                                   div(class = "box-x", h4("Tujuan analisis"), tags$ol(lapply(TUJUAN, tags$li))),
                                   div(class = "box-x", h4("Sumber dan karakteristik data"), p(SUMBER_DATA), uiOutput("data_info")),
                                   div(class = "box-x", h4("Definisi variabel"), tableOutput("kamus")),
                                   div(class = "box-x", h4("Pengolahan data"), uiOutput("prep_info")),
                                   div(class = "box-x", h4("Cuplikan data"), tableOutput("preview"))
                          ),
                          # -------------------------------------------------------- 2. UNIVARIAT
                          tabPanel("2. Univariat",
                                   br(),
                                   fluidRow(column(4, selectInput("uni_var", "Variabel", choices = c(cat_all, num_all, "Date"), selected = "Total")), column(4, uiOutput("uni_type_ui"))),
                                   fluidRow(column(7, div(class = "box-x", uiOutput("uni_main"))), column(5, div(class = "box-x", h4("Statistik deskriptif"), tableOutput("uni_table")))),
                                   div(class = "box-x", h4("Interpretasi"), uiOutput("uni_interp"))
                          ),
                          # --------------------------------------------------------- 3. BIVARIAT
                          tabPanel("3. Bivariat",
                                   br(),
                                   fluidRow(column(3, selectInput("bi_x", "Variabel X", choices = c(cat_all, num_all), selected = "Quantity")),
                                            column(3, selectInput("bi_y", "Variabel Y", choices = c(cat_all, num_all), selected = "Total")),
                                            column(3, uiOutput("bi_type_ui"))),
                                   p(class = "help-note", "Jenis grafik menyesuaikan tipe pasangan variabel: numerik-numerik, kategorik-numerik, atau kategorik-kategorik."),
                                   fluidRow(column(7, div(class = "box-x", plotOutput("bi_plot", height = 400))), column(5, div(class = "box-x", h4("Ringkasan dan uji"), tableOutput("bi_table")))),
                                   div(class = "box-x", h4("Interpretasi"), uiOutput("bi_interp"))
                          ),
                          # ------------------------------------------------------- 4. MULTIVARIAT
                          tabPanel("4. Multivariat",
                                   br(),
                                   tabsetPanel(type = "pills",
                                               tabPanel("Scatter multivariat", br(),
                                                        fluidRow(column(3, selectInput("mv_x", "Sumbu X", num_all, "Quantity")), column(3, selectInput("mv_y", "Sumbu Y", num_all, "Total")),
                                                                 column(3, selectInput("mv_col", "Warna", c(NONE, cat_all, num_all), "Customer type")), column(3, selectInput("mv_size", "Ukuran titik", c(NONE, num_all), "Rating"))),
                                                        fluidRow(column(3, selectInput("mv_facet", "Panel (facet)", c(NONE, cat_all), NONE)), column(3, checkboxInput("mv_lines", "Tampilkan garis regresi per kelompok", TRUE))),
                                                        div(class = "box-x", plotOutput("mv_scatter", height = 440)), div(class = "box-x", h4("Interpretasi"), uiOutput("mv_scatter_i"))),
                                               tabPanel("Heatmap rata-rata", br(),
                                                        fluidRow(column(3, sel_cat("mv_ha", "Kategorik 1 (sumbu X)", "Branch")), column(3, sel_cat("mv_hb", "Kategorik 2 (sumbu Y)", "Product line")),
                                                                 column(3, selectInput("mv_hn", "Variabel numerik", num_all, "Total")), column(3, selectInput("mv_hf", "Ukuran", c("Rata-rata", "Median", "Jumlah", "Banyak transaksi"), "Rata-rata"))),
                                                        div(class = "box-x", plotOutput("mv_heat", height = 420)), div(class = "box-x", h4("Interpretasi"), uiOutput("mv_heat_i"))),
                                               tabPanel("Matriks korelasi", br(),
                                                        fluidRow(column(6, checkboxGroupInput("mv_cv", "Variabel numerik", num_all, num_all, inline = TRUE)), column(3, selectInput("mv_cm", "Metode", c("pearson", "spearman"), "pearson"))),
                                                        div(class = "box-x", plotOutput("mv_cor", height = 420)), div(class = "box-x", h4("Interpretasi"), uiOutput("mv_cor_i"))),
                                               tabPanel("Tren waktu", br(),
                                                        fluidRow(column(3, selectInput("mv_tm", "Metrik", c("Jumlah transaksi", "Total penjualan (jumlah)", "Rata-rata Total", "Rata-rata Rating", "Total Quantity"), "Total penjualan (jumlah)")),
                                                                 column(3, selectInput("mv_tu", "Satuan waktu", c("Harian", "Mingguan", "Bulanan"), "Mingguan")), column(3, selectInput("mv_tg", "Kelompok", c(NONE, cat_base), "Branch"))),
                                                        div(class = "box-x", plotOutput("mv_trend", height = 420)), div(class = "box-x", h4("Interpretasi"), uiOutput("mv_trend_i"))),
                                               tabPanel("PCA", br(),
                                                        fluidRow(column(6, checkboxGroupInput("mv_pv", "Variabel numerik (minimal 3)", num_all, num_all, inline = TRUE)), column(3, selectInput("mv_pc", "Warna titik", c(NONE, cat_base), "Customer type"))),
                                                        fluidRow(column(7, div(class = "box-x", plotOutput("mv_pca", height = 420))), column(5, div(class = "box-x", h4("Keragaman dan loading"), tableOutput("mv_pca_t")))),
                                                        div(class = "box-x", h4("Interpretasi"), uiOutput("mv_pca_i"))),
                                               tabPanel("Scatterplot matrix", br(),
                                                        fluidRow(column(6, checkboxGroupInput("mv_mvars", "Variabel numerik (minimal 2)", num_all, num_all, inline = TRUE)), column(3, selectInput("mv_mcol", "Warna titik", c(NONE, cat_base), NONE))),
                                                        div(class = "box-x", plotOutput("mv_pairs", height = 540)), div(class = "box-x", h4("Interpretasi"), uiOutput("mv_pairs_i"))),
                                               tabPanel("Pencilan Mahalanobis", br(),
                                                        fluidRow(column(8, checkboxGroupInput("mv_dvars", "Variabel numerik (minimal 2)", num_all, num_all, inline = TRUE))),
                                                        fluidRow(column(6, div(class = "box-x", plotOutput("mv_mh1", height = 380))), column(6, div(class = "box-x", plotOutput("mv_mh2", height = 380)))),
                                                        div(class = "box-x", h4("Sepuluh observasi dengan jarak terbesar"), tableOutput("mv_mh_t")), div(class = "box-x", h4("Interpretasi"), uiOutput("mv_mh_i"))),
                                               tabPanel("Q-Q plot Chi-Square", br(),
                                                        fluidRow(column(8, checkboxGroupInput("mv_qvars", "Variabel numerik (minimal 2)", num_all, num_all, inline = TRUE))),
                                                        div(class = "box-x", plotOutput("mv_qq", height = 440)), div(class = "box-x", h4("Interpretasi"), uiOutput("mv_qq_i")))
                                   )
                          ),
                          # ---------------------------------------------------------- 5. REGRESI
                          tabPanel("5. Regresi Linear Berganda",
                                   br(),
                                   div(class = "box-x", h4("Spesifikasi model"),
                                       fluidRow(column(3, selectInput("rg_y", "Variabel respons (Y)", c("Total", "Rating"), "Total")),
                                                column(5, checkboxGroupInput("rg_x", "Variabel prediktor (X)", c("Quantity", "Rating", "Hour", cat_all), c("Quantity", "Rating", "Hour", cat_base), inline = TRUE)),
                                                column(2, radioButtons("rg_log", "Transformasi respons", c("Tanpa" = "no", "ln(Y)" = "yes"), "no")),
                                                column(2, br(), actionButton("rg_go", "Bangun model", class = "btn-pink"))),
                                       p(class = "help-note", "Respons Total dipilih karena merupakan ukuran utama nilai transaksi. Pilih minimal dua prediktor. Variabel kategorik otomatis diubah menjadi dummy dengan kategori pertama sebagai acuan.")),
                                   tabsetPanel(type = "pills",
                                               tabPanel("Hasil model", br(),
                                                        div(class = "box-x", h4("Persamaan model"), verbatimTextOutput("rg_eq")),
                                                        fluidRow(column(5, div(class = "box-x", h4("Ukuran kecocokan model"), tableOutput("rg_fit"))), column(7, div(class = "box-x", h4("Uji signifikansi per variabel (ANOVA)"), tableOutput("rg_anova")))),
                                                        div(class = "box-x", h4("Estimasi koefisien"), tableOutput("rg_coef"), p(class = "help-note", "Signif.: *** p<0.001, ** p<0.01, * p<0.05, . p<0.1")),
                                                        div(class = "box-x", h4("Interpretasi hasil"), uiOutput("rg_interp"))),
                                               tabPanel("Visualisasi hasil model", br(),
                                                        fluidRow(column(6, div(class = "box-x", plotOutput("rg_p1", height = 380))), column(6, div(class = "box-x", plotOutput("rg_p2", height = 380)))),
                                                        div(class = "box-x", h4("Cara membaca"), uiOutput("rg_vis_i"))),
                                               tabPanel("Galeri visualisasi regresi", br(),
                                                        fluidRow(column(6, div(class = "box-x", plotOutput("rv1", height = 340))), column(6, div(class = "box-x", plotOutput("rv2", height = 340)))),
                                                        fluidRow(column(6, div(class = "box-x", plotOutput("rv3", height = 340))), column(6, div(class = "box-x", plotOutput("rv4", height = 340)))),
                                                        fluidRow(column(6, div(class = "box-x", plotOutput("rv5", height = 340))), column(6, div(class = "box-x", plotOutput("rv6", height = 340)))),
                                                        fluidRow(column(6, div(class = "box-x", plotOutput("rv7", height = 340))), column(6, div(class = "box-x", plotOutput("rv8", height = 340)))),
                                                        div(class = "box-x", h4("Cara membaca"), uiOutput("rv_i"))),
                                               tabPanel("Diagnostik asumsi", br(),
                                                        fluidRow(column(6, div(class = "box-x", plotOutput("rg_d1", height = 320))), column(6, div(class = "box-x", plotOutput("rg_d2", height = 320)))),
                                                        fluidRow(column(6, div(class = "box-x", plotOutput("rg_d3", height = 320))), column(6, div(class = "box-x", plotOutput("rg_d4", height = 320)))),
                                                        fluidRow(column(6, div(class = "box-x", h4("Uji asumsi"), tableOutput("rg_tests"))), column(6, div(class = "box-x", h4("Multikolinearitas (VIF)"), tableOutput("rg_vif"))))),
                                               tabPanel("Keterbatasan model", br(), div(class = "box-x", h4("Keterbatasan model yang digunakan"), bullets(KETERBATASAN)))
                                   )
                          ),
                          # ------------------------------------------------------- 6. KESIMPULAN
                          tabPanel("6. Kesimpulan",
                                   br(),
                                   div(class = "box-x", h4("Kesimpulan utama"), uiOutput("kes_auto")),
                                   div(class = "box-x", h4("Rekomendasi"), bullets(REKOMENDASI))
                          )
              )
    )
  )
)

# =============================================================================
# SERVER
# =============================================================================
server <- function(input, output, session) {
  
  data_raw <- reactive({
    d <- if (!is.null(input$file)) read_sales(input$file$datapath) else d0
    validate(need(!is.null(d), "Data tidak ditemukan. Letakkan Data_Penjualan.xlsx di folder aplikasi atau unggah file melalui panel kiri (kolom wajib: Branch, Customer type, Gender, Product line, Quantity, Total, Date, Time, Payment, Rating)."))
    d
  })
  
  # perbarui pilihan filter bila file baru diunggah
  observeEvent(input$file, {
    d <- data_raw()
    updateCheckboxGroupInput(session, "f_branch", choices = levels(d$Branch), selected = levels(d$Branch), inline = TRUE)
    updateCheckboxGroupInput(session, "f_cust", choices = levels(d$`Customer type`), selected = levels(d$`Customer type`), inline = TRUE)
    updateCheckboxGroupInput(session, "f_gender", choices = levels(d$Gender), selected = levels(d$Gender), inline = TRUE)
    updateCheckboxGroupInput(session, "f_pay", choices = levels(d$Payment), selected = levels(d$Payment))
    updateSelectInput(session, "f_prod", choices = levels(d$`Product line`), selected = levels(d$`Product line`))
    updateDateRangeInput(session, "f_date", start = min(d$Date), end = max(d$Date))
  })
  observeEvent(input$f_reset, {
    d <- data_raw()
    updateCheckboxGroupInput(session, "f_branch", selected = levels(d$Branch)); updateCheckboxGroupInput(session, "f_cust", selected = levels(d$`Customer type`))
    updateCheckboxGroupInput(session, "f_gender", selected = levels(d$Gender)); updateCheckboxGroupInput(session, "f_pay", selected = levels(d$Payment))
    updateSelectInput(session, "f_prod", selected = levels(d$`Product line`)); updateDateRangeInput(session, "f_date", start = min(d$Date), end = max(d$Date))
  })
  
  data_use <- reactive({
    d <- data_raw()
    if (!is.null(input$f_branch)) d <- d[d$Branch %in% input$f_branch, ] else d <- d[0, ]
    d <- d[d$`Customer type` %in% input$f_cust & d$Gender %in% input$f_gender & d$Payment %in% input$f_pay & d$`Product line` %in% input$f_prod, ]
    if (!is.null(input$f_date)) d <- d[d$Date >= input$f_date[1] & d$Date <= input$f_date[2], ]
    validate(need(nrow(d) >= 30, "Data hasil filter terlalu sedikit (kurang dari 30 baris). Longgarkan filter pada panel kiri."))
    d <- droplevels(d); d
  })
  
  output$n_info <- renderUI({
    n0 <- nrow(data_raw()); n1 <- tryCatch(nrow(data_use()), error = function(e) 0)
    div(class = "nfilter", "Transaksi terpilih:", br(), tags$b(format(n1, big.mark = ",")), paste0(" dari ", format(n0, big.mark = ","), " (", fmt(100 * n1 / n0, 1), "%)"))
  })
  
  # ---- 1. DATA ----
  output$data_info <- renderUI({
    d <- data_raw()
    bullets(c(sprintf("Jumlah observasi: **%s** transaksi; jumlah variabel: **10** variabel asli (ditambah 3 variabel turunan: Hour, Bulan, Hari).", format(nrow(d), big.mark = ",")),
              sprintf("Periode data: **%s** sampai **%s**.", min(d$Date), max(d$Date)),
              "Variabel kategorik: Branch, Customer type, Gender, Product line, Payment. Variabel numerik: Quantity, Total, Rating, serta Hour (turunan)."))
  })
  output$kamus <- renderTable(kamus, striped = TRUE, spacing = "s")
  output$prep_info <- renderUI({
    d <- data_raw(); q <- quantile(d$Total, c(.25, .75)); out <- sum(d$Total < q[1] - 1.5 * diff(q) | d$Total > q[2] + 1.5 * diff(q))
    bullets(c(sprintf("**Data hilang:** %d sel kosong pada seluruh kolom analisis.", sum(is.na(d))),
              sprintf("**Data ganda:** %d baris duplikat.", sum(duplicated(d))),
              sprintf("**Pencilan:** terdapat %d pencilan pada Total (aturan 1.5 x IQR). Pencilan dipertahankan karena merupakan transaksi bernilai besar yang valid.", out),
              "**Pengolahan:** variabel Hour, Bulan, dan Hari dibentuk dari Time dan Date; variabel kategorik diubah menjadi faktor.",
              "**Transformasi:** tidak dilakukan pada eksplorasi; pada regresi tersedia pilihan transformasi ln(Y) untuk mengatasi ragam yang tidak konstan."))
  })
  output$preview <- renderTable({ x <- head(data_raw(), 10); x$Date <- as.character(x$Date); x }, striped = TRUE, spacing = "s")
  
  # ---- 2. UNIVARIAT ----
  output$uni_type_ui <- renderUI(selectInput("uni_type", "Jenis grafik", uni_choices(input$uni_var)))
  uni_type_val <- reactive({ chs <- uni_choices(input$uni_var); if (is.null(input$uni_type) || !(input$uni_type %in% chs)) chs[1] else input$uni_type })
  output$uni_main <- renderUI(if (identical(uni_type_val(), "Stem-and-leaf")) verbatimTextOutput("uni_stem") else plotOutput("uni_plot", height = 380))
  output$uni_plot <- renderPlot(plot_uni(data_use(), input$uni_var, uni_type_val()), res = 96)
  output$uni_stem <- renderText({
    x <- data_use()[[input$uni_var]]; req(is.numeric(x))
    paste(c(paste("Stem-and-leaf plot:", input$uni_var), "", capture.output(stem(x))), collapse = "\n")
  })
  output$uni_table <- renderTable(uni_table(data_use(), input$uni_var), striped = TRUE, spacing = "s")
  output$uni_interp <- renderUI(bullets(interp_uni(data_use(), input$uni_var)))
  
  # ---- 3. BIVARIAT ----
  output$bi_type_ui <- renderUI(selectInput("bi_type", "Jenis grafik", bi_choices(input$bi_x, input$bi_y)))
  bi_type_val <- reactive({ chs <- bi_choices(input$bi_x, input$bi_y); if (is.null(input$bi_type) || !(input$bi_type %in% chs)) chs[1] else input$bi_type })
  output$bi_plot <- renderPlot({ validate(need(input$bi_x != input$bi_y, "Pilih dua variabel yang berbeda.")); plot_bi(data_use(), input$bi_x, input$bi_y, bi_type_val()) }, res = 96)
  bi_res <- reactive({ validate(need(input$bi_x != input$bi_y, "Pilih dua variabel yang berbeda.")); bi_text(data_use(), input$bi_x, input$bi_y) })
  output$bi_table <- renderTable(bi_res()$table, striped = TRUE, spacing = "s")
  output$bi_interp <- renderUI(bullets(bi_res()$lines))
  
  # ---- 4. MULTIVARIAT ----
  output$mv_scatter <- renderPlot({
    validate(need(input$mv_x != input$mv_y, "Pilih sumbu X dan Y yang berbeda."))
    plot_mv_scatter(data_use(), input$mv_x, input$mv_y, input$mv_col, input$mv_size, input$mv_facet, input$mv_lines)
  }, res = 96)
  output$mv_scatter_i <- renderUI({
    d <- data_use(); x <- input$mv_x; y <- input$mv_y; col <- input$mv_col; r <- cor(d[[x]], d[[y]])
    ln <- sprintf("Secara keseluruhan korelasi **%s** dan **%s** adalah **%s** (%s).", x, y, fmt(r, 3), strength(r))
    if (col != NONE && !is.numeric(d[[col]])) {
      g <- d %>% group_by(.data[[col]]) %>% summarise(r = cor(.data[[x]], .data[[y]]), .groups = "drop"); names(g)[1] <- "k"
      ln <- c(ln, sprintf("Korelasi per **%s**: %s. %s", col, paste0(g$k, " = ", fmt(g$r, 2), collapse = "; "),
                          if (diff(range(g$r)) < .1) "Pola hubungan hampir sama di semua kelompok sehingga variabel warna tidak mengubah hubungan X dan Y." else "Kekuatan hubungan berbeda antar kelompok, menandakan kemungkinan interaksi."))
    }
    if (input$mv_size != NONE) ln <- c(ln, sprintf("Ukuran titik menunjukkan **%s**; jika titik besar dan kecil tersebar acak, variabel tersebut tidak berkaitan dengan posisi titik pada X dan Y.", input$mv_size))
    bullets(ln)
  })
  
  output$mv_heat <- renderPlot({ validate(need(input$mv_ha != input$mv_hb, "Pilih dua variabel kategorik yang berbeda.")); plot_mv_heat(data_use(), input$mv_ha, input$mv_hb, input$mv_hn, input$mv_hf) }, res = 96)
  output$mv_heat_i <- renderUI({
    validate(need(input$mv_ha != input$mv_hb, "Pilih dua variabel kategorik yang berbeda."))
    r <- mv_heat_data(data_use(), input$mv_ha, input$mv_hb, input$mv_hn, input$mv_hf); hi <- r[which.max(r$val), ]; lo <- r[which.min(r$val), ]
    bullets(c(sprintf("Nilai tertinggi pada kombinasi **%s - %s** (%s) dan terendah pada **%s - %s** (%s).", hi$A, hi$B, fmt(hi$val, 1), lo$A, lo$B, fmt(lo$val, 1)),
              sprintf("Selisih tertinggi-terendah sebesar **%s** (%s%% dari nilai tertinggi).", fmt(hi$val - lo$val, 1), fmt(100 * (hi$val - lo$val) / hi$val, 1)),
              "Warna lebih pekat menandakan nilai lebih besar. Perbedaan antar sel perlu dikonfirmasi dengan uji statistik sebelum disimpulkan sebagai perbedaan nyata."))
  })
  
  output$mv_cor <- renderPlot({ validate(need(length(input$mv_cv) >= 2, "Pilih minimal dua variabel.")); plot_mv_cor(data_use(), input$mv_cv, input$mv_cm) }, res = 96)
  output$mv_cor_i <- renderUI({
    validate(need(length(input$mv_cv) >= 2, "Pilih minimal dua variabel."))
    m <- cor(data_use()[, input$mv_cv, drop = FALSE], method = input$mv_cm); m[lower.tri(m, diag = TRUE)] <- NA
    df <- as.data.frame(as.table(m)); df <- df[!is.na(df$Freq), ]; df <- df[order(-abs(df$Freq)), ]
    bullets(c(sprintf("Korelasi terkuat: **%s - %s** (r = %s; %s).", df$Var1[1], df$Var2[1], fmt(df$Freq[1], 3), strength(df$Freq[1])),
              sprintf("Korelasi terlemah: **%s - %s** (r = %s).", df$Var1[nrow(df)], df$Var2[nrow(df)], fmt(df$Freq[nrow(df)], 3)),
              "Warna merah muda tua = korelasi positif, biru keabuan = negatif, putih = mendekati nol."))
  })
  
  output$mv_trend <- renderPlot(plot_mv_trend(data_use(), input$mv_tm, input$mv_tu, input$mv_tg), res = 96)
  output$mv_trend_i <- renderUI({
    td <- trend_data(data_use(), input$mv_tm, input$mv_tu, input$mv_tg); hi <- td[which.max(td$val), ]; lo <- td[which.min(td$val), ]
    gtxt <- if (input$mv_tg != NONE) paste0(" (kelompok ", hi[[input$mv_tg]], ")") else ""
    bullets(c(sprintf("Nilai **%s** tertinggi pada periode **%s**%s sebesar %s; terendah pada **%s** sebesar %s.", input$mv_tm, hi$Periode, gtxt, fmt(hi$val, 1), lo$Periode, fmt(lo$val, 1)),
              "Bandingkan garis antar kelompok untuk melihat apakah pola naik-turun sama (pola serupa menandakan faktor waktu yang umum, bukan khas satu kelompok).",
              if (input$mv_tu == "Mingguan") "Minggu pertama dan terakhir dapat berisi lebih sedikit hari sehingga nilainya tampak lebih rendah." else "Perhatikan bahwa data hanya mencakup beberapa bulan, sehingga pola musiman tidak dapat disimpulkan."))
  })
  
  output$mv_pca <- renderPlot({ validate(need(length(input$mv_pv) >= 3, "Pilih minimal tiga variabel numerik.")); plot_mv_pca(data_use(), input$mv_pv, input$mv_pc) }, res = 96)
  output$mv_pca_t <- renderTable({
    validate(need(length(input$mv_pv) >= 3, "Pilih minimal tiga variabel numerik."))
    pr <- pca_fit(data_use(), input$mv_pv); ve <- 100 * pr$sdev^2 / sum(pr$sdev^2)
    ld <- round(pr$rotation[, 1:2], 3); data.frame(Variabel = rownames(ld), PC1 = ld[, 1], PC2 = ld[, 2], check.names = FALSE) %>%
      rbind(data.frame(Variabel = "Keragaman dijelaskan (%)", PC1 = round(ve[1], 1), PC2 = round(ve[2], 1)))
  }, striped = TRUE, spacing = "s")
  output$mv_pca_i <- renderUI({
    validate(need(length(input$mv_pv) >= 3, "Pilih minimal tiga variabel numerik."))
    pr <- pca_fit(data_use(), input$mv_pv); ve <- 100 * pr$sdev^2 / sum(pr$sdev^2); l1 <- pr$rotation[, 1]
    bullets(c(sprintf("PC1 menjelaskan **%s%%** dan PC2 **%s%%** keragaman; bersama-sama **%s%%**.", fmt(ve[1], 1), fmt(ve[2], 1), fmt(ve[1] + ve[2], 1)),
              sprintf("Variabel dengan kontribusi terbesar pada PC1: **%s** (loading %s).", names(which.max(abs(l1))), fmt(l1[which.max(abs(l1))], 2)),
              "Panah yang searah menandakan variabel berkorelasi positif; panah yang saling tegak lurus menandakan variabel hampir tidak berkorelasi. Jika keragaman terbagi merata ke banyak komponen, variabel-variabel tersebut membawa informasi yang berbeda-beda."))
  })
  
  # ---- 4b. SCATTERPLOT MATRIX, MAHALANOBIS, Q-Q CHI-SQUARE ----
  output$mv_pairs <- renderPlot({ validate(need(length(input$mv_mvars) >= 2, "Pilih minimal dua variabel numerik.")); plot_mv_pairs(data_use(), input$mv_mvars, input$mv_mcol) }, res = 96)
  output$mv_pairs_i <- renderUI({
    validate(need(length(input$mv_mvars) >= 2, "Pilih minimal dua variabel numerik."))
    m <- cor(data_use()[, input$mv_mvars, drop = FALSE]); m[lower.tri(m, diag = TRUE)] <- NA
    df <- as.data.frame(as.table(m)); df <- df[!is.na(df$Freq), ]; df <- df[order(-abs(df$Freq)), ]
    bullets(c("Diagonal: histogram tiap variabel; segitiga bawah: scatterplot dengan kurva lowess; segitiga atas: koefisien korelasi Pearson (semakin besar angka, semakin kuat hubungan).",
              sprintf("Pasangan dengan hubungan terkuat: **%s - %s** (r = %s; %s). Pasangan terlemah: **%s - %s** (r = %s).", df$Var1[1], df$Var2[1], fmt(df$Freq[1], 3), strength(df$Freq[1]), df$Var1[nrow(df)], df$Var2[nrow(df)], fmt(df$Freq[nrow(df)], 3)),
              "Kurva lowess yang mendekati garis lurus menandakan hubungan linear; kurva melengkung menandakan hubungan nonlinear. Titik yang jauh dari kumpulan utama adalah kandidat pencilan."))
  })
  
  mh <- reactive({ validate(need(length(input$mv_dvars) >= 2, "Pilih minimal dua variabel numerik.")); mahal_fit(data_use(), input$mv_dvars) })
  output$mv_mh1 <- renderPlot(plot_mv_mahal_idx(mh()), res = 96)
  output$mv_mh2 <- renderPlot(plot_mv_mahal_sc(mh(), input$mv_dvars), res = 96)
  output$mv_mh_t <- renderTable(mahal_table(data_use(), mh(), input$mv_dvars), striped = TRUE, spacing = "s")
  output$mv_mh_i <- renderUI({
    f <- mh(); d <- data_use(); vars <- input$mv_dvars; n <- nrow(f$df); k <- sum(f$df$out); top <- f$df[which.max(f$df$D2), ]
    z <- vapply(vars, function(v) (d[[v]][top$idx] - mean(d[[v]])) / sd(d[[v]]), 0); vmax <- names(z)[which.max(abs(z))]
    bullets(c(sprintf("Batas pencilan multivariat: Chi-Square 97.5%% dengan df = %d, yaitu **%s**. Observasi dengan D2 di atas batas dianggap pencilan.", f$p, fmt(f$cut, 2)),
              sprintf("Ditemukan **%d pencilan** (**%s%%** dari %d observasi). Pada data normal multivariat, proporsi yang diharapkan sekitar 2.5%%.", k, fmt(100 * k / n, 1), n),
              sprintf("Observasi paling ekstrem: baris **%d** (D2 = %s). Variabel yang paling menyimpang: **%s** (z = %s).", top$idx, fmt(top$D2, 2), vmax, fmt(z[vmax], 2)),
              "Pencilan multivariat dapat berupa kombinasi nilai yang tidak lazim walau tiap variabel tampak wajar. Pencilan sebaiknya diperiksa dulu keabsahannya sebelum dihapus."))
  })
  
  qq <- reactive({ validate(need(length(input$mv_qvars) >= 2, "Pilih minimal dua variabel numerik.")); mahal_fit(data_use(), input$mv_qvars) })
  output$mv_qq <- renderPlot(plot_mv_qqchi(qq()), res = 96)
  output$mv_qq_i <- renderUI({
    f <- qq(); n <- nrow(f$df); s <- sort(f$df$D2); q <- qchisq((seq_len(n) - .5) / n, df = f$p); r <- cor(s, q)
    bullets(c(sprintf("Korelasi antara D2 terurut dan kuantil Chi-Square = **%s**; semakin mendekati 1, semakin sesuai data dengan normal multivariat.", fmt(r, 3)),
              sprintf("Titik terjauh dari garis berada di ujung kanan (baris %s), yang menandakan observasi dengan jarak Mahalanobis besar.", paste(tail(f$df$idx[order(f$df$D2)], 3), collapse = ", ")),
              "Titik yang melengkung ke atas di ujung kanan menandakan ekor distribusi lebih berat atau adanya pencilan; pola berundak dapat muncul karena variabel diskret (Quantity, Hour). Asumsi normalitas multivariat dianggap kurang terpenuhi bila penyimpangan jelas."))
  })
  
  # ---- 5. REGRESI ----
  logy <- reactive(identical(input$rg_log, "yes"))
  model <- eventReactive(input$rg_go, {
    preds <- setdiff(input$rg_x, input$rg_y)
    validate(need(length(preds) >= 2, "Pilih minimal dua variabel prediktor (selain variabel respons)."))
    d <- data_use(); if (logy()) validate(need(all(d[[input$rg_y]] > 0), "Transformasi ln hanya untuk respons bernilai positif."))
    list(m = build_model(d, input$rg_y, preds, logy()), resp = input$rg_y, logy = logy(), preds = preds, d = d)
  }, ignoreNULL = FALSE)
  M <- reactive(model()$m)
  
  output$rg_eq <- renderText(make_equation(M(), model()$resp, model()$logy))
  output$rg_fit <- renderTable(fit_table(M()), striped = TRUE, spacing = "s")
  output$rg_anova <- renderTable(anova_table(M()), striped = TRUE, spacing = "s")
  output$rg_coef <- renderTable(coef_table(M()), striped = TRUE, spacing = "s")
  output$rg_interp <- renderUI(bullets(interp_reg(M(), model()$resp, model()$logy)))
  output$rg_p1 <- renderPlot(plot_reg_avp(M(), model()$logy, model()$resp), res = 96)
  output$rg_p2 <- renderPlot(plot_reg_coef(M()), res = 96)
  output$rg_vis_i <- renderUI({
    m <- M(); r2 <- summary(m)$r.squared
    bullets(c(sprintf("**Aktual vs prediksi:** semakin rapat titik di sekitar garis putus-putus, semakin baik prediksi model (R2 = %s). Titik yang menyebar lebar menunjukkan keragaman yang belum dijelaskan model.", fmt(r2, 3)),
              "**Koefisien dan selang kepercayaan:** titik = estimasi, garis = selang 95%. Selang yang tidak memotong garis nol (warna tua) menandakan prediktor signifikan pada taraf 5%. Skala koefisien bergantung pada satuan variabelnya, jadi bandingkan terutama ada-tidaknya pengaruh, bukan panjang garisnya."))
  })
  output$rg_d1 <- renderPlot(plot_reg_resid(M()), res = 96)
  output$rg_d2 <- renderPlot(plot_reg_qq(M()), res = 96)
  output$rg_d3 <- renderPlot(plot_reg_scale(M()), res = 96)
  output$rg_d4 <- renderPlot(plot_reg_cook(M()), res = 96)
  output$rg_tests <- renderTable({
    m <- M(); e <- resid(m); sh <- shapiro.test(if (length(e) > 5000) sample(e, 5000) else e); bp <- bp_test(m); cd <- cooks.distance(m)
    data.frame(Uji = c("Shapiro-Wilk (normalitas residual)", "Breusch-Pagan (kesamaan ragam)", "Observasi berpengaruh (Cook's D > 4/n)"),
               Statistik = c(fmt(unname(sh$statistic), 4), fmt(unname(bp["stat"]), 3), as.character(sum(cd > 4 / length(cd)))),
               `p-value` = c(fmt_p(sh$p.value), fmt_p(unname(bp["p"])), "-"),
               Keputusan = c(if (sh$p.value < .05) "Tidak normal" else "Normal", if (bp["p"] < .05) "Ragam tidak konstan" else "Ragam konstan", paste(sum(cd > 4 / length(cd)), "observasi")), check.names = FALSE)
  }, striped = TRUE, spacing = "s")
  output$rg_vif <- renderTable({ v <- vif_table(M()); validate(need(!is.null(v), "VIF membutuhkan minimal dua prediktor.")); v$`VIF setara` <- fmt(v$`VIF setara`, 3); v$Keputusan <- ifelse(as.numeric(v$`VIF setara`) < 5, "Aman", "Perlu perhatian"); v }, striped = TRUE, spacing = "s")
  
  # ---- 5b. GALERI VISUALISASI REGRESI ----
  output$rv1 <- renderPlot(plot_reg_scatter(model()$d, model()$resp, model()$preds, model()$logy), res = 96)
  output$rv2 <- renderPlot(plot_reg_cor(model()$d, model()$resp, model()$preds), res = 96)
  output$rv3 <- renderPlot(plot_reg_coef(M()), res = 96)
  output$rv4 <- renderPlot(plot_reg_qq(M()), res = 96)
  output$rv5 <- renderPlot(plot_reg_hist(M()), res = 96)
  output$rv6 <- renderPlot(plot_reg_ts(M(), model()$d), res = 96)
  output$rv7 <- renderPlot(plot_reg_acf(M(), model()$d), res = 96)
  output$rv8 <- renderPlot(plot_reg_obsres(M()), res = 96)
  output$rv_i <- renderUI({
    m <- M(); o <- reg_ord_res(m, model()$d); e <- o$res; n <- length(e)
    a <- acf(e, lag.max = min(30, n - 1), plot = FALSE); ci <- qnorm(.975) / sqrt(n); nsig <- sum(abs(as.numeric(a$acf)[-1]) > ci)
    dw <- sum(diff(e)^2) / sum(e^2); sw <- shapiro.test(if (n > 5000) sample(e, 5000) else e)$p.value
    bullets(c("**Scatterplot:** menunjukkan bentuk hubungan respons dengan tiap prediktor numerik; garis yang jelas menanjak/menurun menandakan hubungan linear.",
              "**Correlation matrix:** korelasi respons dengan prediktor menunjukkan kandidat prediktor kuat; korelasi tinggi antar prediktor mengindikasikan potensi multikolinearitas.",
              "**Coefficient CI plot:** selang 95% yang tidak memotong garis nol berarti koefisien signifikan pada taraf 5%.",
              sprintf("**Normal probability plot dan histogram residual:** titik mengikuti garis dan histogram mendekati kurva normal berarti residual normal. Hasil Shapiro-Wilk p %s, sehingga residual %s.", pq(sw), if (sw < .05) "**tidak normal**" else "normal"),
              sprintf("**Time series plot dan ACF:** statistik Durbin-Watson = %s (sekitar 2 berarti tidak ada autokorelasi); %d dari %d lag berada di luar batas 95%%. %s", fmt(dw, 3), nsig, length(a$acf) - 1,
                      if (nsig <= 0.05 * (length(a$acf) - 1) + 1 && abs(dw - 2) < .4) "Tidak ada indikasi autokorelasi residual yang berarti." else "Terdapat indikasi autokorelasi residual; periksa pola waktu."),
              "**Observed vs residual:** residual idealnya tersebar acak di sekitar nol. Sebaran yang melebar pada nilai observasi besar menandakan ragam residual tidak konstan (heteroskedastis)."))
  })
  
  # ---- 6. KESIMPULAN ----
  output$kes_auto <- renderUI(bullets(auto_conclusion(data_use(), M(), model()$resp, model()$logy)))
}

shinyApp(ui, server)
