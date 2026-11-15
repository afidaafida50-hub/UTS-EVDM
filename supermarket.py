# -*- coding: utf-8 -*-
# =============================================================================
# Konversi UTS_EVDM.R (R Shiny) -> Python (Streamlit)
# Jalankan:  streamlit run UTS_EVDM_streamlit.py
# Pustaka :  pip install streamlit plotnine pandas numpy scipy statsmodels matplotlib openpyxl
# Padanan  :  ggplot2 -> plotnine | dplyr -> pandas | readxl -> pandas.read_excel
#             lm -> statsmodels | prcomp -> numpy SVD | pairs() -> matplotlib
# =============================================================================
import os
import re
import math
import warnings

import numpy as np
import pandas as pd
from scipy import stats as sps
import statsmodels.api as sm
import statsmodels.formula.api as smf
from statsmodels.nonparametric.smoothers_lowess import lowess
from statsmodels.tsa.stattools import acf as sm_acf

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.colors import to_rgba

from plotnine import (
    ggplot, aes, labs, theme, theme_minimal, theme_void, element_text, element_line,
    element_rect, element_blank, geom_line, geom_point, geom_smooth, geom_col, geom_text,
    geom_histogram, geom_density, geom_vline, geom_hline, geom_boxplot, geom_violin,
    geom_jitter, geom_errorbar, geom_ribbon, geom_segment, geom_tile, geom_bin_2d,
    geom_abline, stat_qq, stat_qq_line, stat_function, facet_wrap, scale_fill_manual,
    scale_colour_manual, scale_fill_gradient, scale_fill_gradient2, scale_colour_gradient,
    scale_size_continuous, scale_size_manual, scale_x_discrete, scale_y_discrete,
    scale_x_continuous, scale_y_continuous, coord_flip, after_stat, arrow, position_dodge,
    position_stack,
)

import streamlit as st

st.set_page_config(page_title="Analisis Faktor yang Memengaruhi Nilai Transaksi Penjualan Supermarket", layout="wide")


class SafeException(Exception):
    pass

warnings.filterwarnings("ignore")

# ---- KONFIGURASI: EDIT BAGIAN INI -------------------------------------------
DATA_FILE = "Data_Penjualan.xlsx"
NAMA = "Afida Mar'atus Solihah"
NIM = "231064005"
APP_TITLE = "Analisis Faktor yang Memengaruhi Nilai Transaksi Penjualan Supermarket"
SUMBER_DATA = "Dataset sekunder data transaksi supermarket (Data_Penjualan.xlsx). Isi dengan tautan/sumber asli yang dapat dipertanggungjawabkan (mis. Kaggle: Supermarket Sales)."
LATAR_BELAKANG = "Supermarket perlu memahami faktor yang berkaitan dengan besarnya nilai transaksi agar strategi promosi, penataan produk, dan pelayanan dapat dirancang berdasarkan data. Nilai transaksi dapat dipengaruhi oleh jumlah barang yang dibeli, karakteristik pelanggan, lini produk, cabang, metode pembayaran, dan waktu transaksi. Analisis eksplorasi dan regresi linear berganda digunakan untuk mengidentifikasi faktor-faktor tersebut."
TUJUAN = [
    "Menggambarkan karakteristik setiap variabel pada data transaksi (univariat).",
    "Mengidentifikasi hubungan dan perbedaan kelompok antar dua variabel (bivariat).",
    "Mengeksplorasi pola yang melibatkan tiga variabel atau lebih (multivariat).",
    "Membangun model regresi linear berganda untuk menjelaskan nilai transaksi (Total) dan mengevaluasi kemampuannya.",
    "Merumuskan kesimpulan dan rekomendasi bagi pengelola supermarket.",
]

# ---- Konstanta variabel ------------------------------------------------------
cat_base = ["Branch", "Customer type", "Gender", "Product line", "Payment"]
cat_all = cat_base + ["Bulan", "Hari"]
num_all = ["Quantity", "Total", "Rating", "Hour"]
NONE = "(tidak ada)"
bulan_lv = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus",
            "September", "Oktober", "November", "Desember"]
hari_lv = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]

kamus = pd.DataFrame({
    "Variabel": ["Branch", "Customer type", "Gender", "Product line", "Quantity", "Total",
                 "Date", "Time", "Payment", "Rating", "Hour (turunan)", "Bulan (turunan)", "Hari (turunan)"],
    "Tipe": ["Kategorik", "Kategorik", "Kategorik", "Kategorik", "Numerik (diskret)", "Numerik (kontinu)",
             "Tanggal", "Waktu", "Kategorik", "Numerik (kontinu)", "Numerik (diskret)", "Kategorik", "Kategorik"],
    "Definisi": ["Cabang supermarket tempat transaksi (A, B, C).",
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
})


# ---- Membaca dan menyiapkan data --------------------------------------------
def _hour_of(t):
    if hasattr(t, "hour"):
        return int(t.hour)
    return int(str(t)[:2])


def read_sales(path):
    ext = os.path.splitext(path)[1].lower().lstrip(".")
    d = pd.read_excel(path) if ext in ("xlsx", "xls") else pd.read_csv(path)
    need = ["Branch", "Customer type", "Gender", "Product line", "Quantity", "Total", "Date", "Time", "Payment", "Rating"]
    if not all(n in d.columns for n in need):
        return None
    d = d.copy()
    d["Date"] = pd.to_datetime(d["Date"]).dt.normalize()
    d["Hour"] = [_hour_of(t) for t in d["Time"]]
    for v in ["Quantity", "Total", "Rating"]:
        d[v] = pd.to_numeric(d[v], errors="coerce")
    for v in cat_base:
        d[v] = pd.Categorical(d[v].astype(str), categories=sorted(d[v].astype(str).unique()))
    mon = d["Date"].dt.month.to_numpy()
    d["Bulan"] = pd.Categorical([bulan_lv[m - 1] for m in mon],
                                categories=[bulan_lv[m - 1] for m in sorted(set(mon))])
    d["Hari"] = pd.Categorical([hari_lv[w] for w in d["Date"].dt.weekday], categories=hari_lv)
    return d[cat_base + ["Bulan", "Hari", "Quantity", "Total", "Rating", "Hour", "Date"]].reset_index(drop=True)


# ---- Palet & tema grafik (pink formal) ---------------------------------------
pk = dict(d="#880E4F", m="#C2185B", s="#AD1457", l="#F48FB1", xl="#F8BBD0",
          bg="#FFF5F8", txt="#3B2A33", grid="#EFDDE4", mauve="#6D4C5E", slate="#4F6D8F")


def pal_cat(n):
    base = ["#AD1457", "#F48FB1", "#6D4C5E", "#F8BBD0", "#D81B60", "#B39DAB", "#880E4F"]
    return [base[i % len(base)] for i in range(n)]


def theme_pink(base=12):
    return theme_minimal(base_size=base) + theme(
        plot_title=element_text(weight="bold", color=pk["d"], size=base + 2),
        plot_subtitle=element_text(color=pk["mauve"]),
        axis_title=element_text(weight="bold", color=pk["mauve"]),
        axis_text=element_text(color=pk["txt"]),
        legend_title=element_text(weight="bold", color=pk["mauve"]),
        strip_text=element_text(weight="bold", color=pk["d"]),
        panel_grid_minor=element_blank(),
        panel_grid_major=element_line(color=pk["grid"]),
        plot_background=element_rect(fill="white", color="white"),
        panel_background=element_rect(fill="white", color="white"))


rot = theme(axis_text_x=element_text(angle=20, ha="right"))


# ---- Fungsi bantu ------------------------------------------------------------
def S(v):
    """Nama kolom aman untuk plotnine/statsmodels (spasi -> garis bawah)."""
    return v.replace(" ", "_")


def sd1(x):
    return float(np.std(np.asarray(x, dtype=float), ddof=1))


def fmt(x, k=2):
    if np.ndim(x) == 0:
        return f"{float(x):,.{k}f}"
    return [f"{float(v):,.{k}f}" for v in np.asarray(x).ravel()]


def fmt_p(p):
    if np.ndim(p) == 0:
        return "< 0.001" if p < 0.001 else f"{p:.3f}"
    return [fmt_p(v) for v in np.asarray(p).ravel()]


def pq(p):   # untuk kalimat: "p < 0.001" / "p = 0.019"
    return "< 0.001" if p < 0.001 else "= " + f"{p:.3f}"


def stars(p):
    if np.ndim(p) == 0:
        return "***" if p < 0.001 else "**" if p < 0.01 else "*" if p < 0.05 else "." if p < 0.1 else ""
    return [stars(v) for v in np.asarray(p).ravel()]


def skew_fun(x):
    x = np.asarray(x, dtype=float); x = x[~np.isnan(x)]; n = len(x)
    return n / ((n - 1) * (n - 2)) * np.sum(((x - x.mean()) / sd1(x)) ** 3)


def kurt_fun(x):
    x = np.asarray(x, dtype=float); x = x[~np.isnan(x)]; n = len(x); z = (x - x.mean()) / sd1(x)
    return n * (n + 1) / ((n - 1) * (n - 2) * (n - 3)) * np.sum(z ** 4) - 3 * (n - 1) ** 2 / ((n - 2) * (n - 3))


def bullets(lines):
    st.markdown("
".join(f"- {l}" for l in lines))


def strength(r):
    a = abs(r)
    if a < .2: return "sangat lemah"
    if a < .4: return "lemah"
    if a < .6: return "sedang"
    if a < .8: return "kuat"
    return "sangat kuat"


def validate(cond, msg):
    """Padanan validate(need(...)) di Shiny R."""
    if not cond:
        raise SafeException(msg)


def is_cat(d, v):
    return v in cat_all


def stem_text(x):
    """Port algoritma stem() R (scale = 1, width = 80, atom = 1e-8)."""
    x = np.sort(np.asarray(x, dtype=float)); n = len(x); scale = 1.0; width = 80; atom = 1e-8
    if n <= 1:
        return ""
    cdiv = lambda a, b: int(a / b)          # pembagian bulat C (pemotongan ke arah nol)
    out = [""]
    if x[-1] > x[0]:
        r = atom + (x[-1] - x[0]) / scale
        c = 10.0 ** int(1.0 - math.floor(math.log10(r) + 1e-7))
        mm = min(2, max(0, int(r * c / 25)))
        k = 3 * mm + 2 - 150 // (n + 50)
        if (k - 1) * (k - 2) * (k - 5) == 0:
            c *= 10.0
        x1 = max(abs(x[0]), abs(x[-1]))
        while x1 * c > 2147483647:
            c /= 10
    else:
        r = atom + abs(x[0]) / scale
        c = 10.0 ** int(11 - int(math.log10(r) + 10))
        k = 2
    mu = 10
    if k * (k - 4) * (k - 8) == 0: mu = 5
    if (k - 1) * (k - 5) * (k - 6) == 0: mu = 20
    lo = math.floor(x[0] * c / mu) * mu
    hi = math.floor(x[-1] * c / mu) * mu
    ldigits = int(math.floor(math.log10(-lo))) + 1 if lo < 0 else 0
    hdigits = int(math.floor(math.log10(hi))) if hi > 0 else 0
    ndigits = hdigits if ldigits < hdigits else ldigits
    if lo < 0 and math.floor(x[0] * c) == lo:
        lo = lo - mu
    hi = lo + mu
    if math.floor(x[0] * c + 0.5) > hi:
        lo = hi; hi = lo + mu
    pdigits = 1 - int(math.floor(math.log10(c) + 0.5))
    if pdigits == 0:
        out.append("  The decimal point is at the |")
    else:
        out.append("  The decimal point is %d digit(s) to the %s of the |" % (abs(pdigits), "right" if pdigits > 0 else "left"))
    out.append("")

    def stem_print(close, dist):
        if cdiv(close, 10) == 0 and dist < 0:
            return "  %*s | " % (ndigits, "-0")
        return "  %*d | " % (ndigits, cdiv(close, 10))

    i = 0
    while True:
        line = stem_print(int(hi), int(lo)) if lo < 0 else stem_print(int(lo), int(hi))
        j = 0
        while i < n:
            xi = int(x[i] * c - .5) if x[i] < 0 else int(x[i] * c + .5)
            if (hi == 0 and x[i] >= 0) or (lo < 0 and xi > hi) or (lo >= 0 and xi >= hi):
                break
            j += 1
            if j <= width - 12:
                line += "%1d" % (abs(xi) % 10)
            i += 1
        if j > width:
            line += "+%d" % (j - width)
        out.append(line)
        if i >= n:
            break
        hi += mu; lo += mu
    out.append("")
    return "\n".join(out)


def smooth_df(x, y, frac=0.75):
    """LOESS/LOWESS manual (dipakai untuk garis tren pada sumbu tanggal)."""
    r = lowess(np.asarray(y, float), np.asarray(x, float), frac=frac, return_sorted=True)
    return pd.DataFrame({"x": r[:, 0], "y": r[:, 1]})


# =============================================================================
# 2. UNIVARIAT
# =============================================================================
def uni_choices(v):
    if v == "Date":
        return ["Deret waktu harian"]
    elif v in cat_all:
        return ["Diagram batang", "Pareto", "Diagram lingkaran"]
    return ["Histogram", "Histogram + Density", "Boxplot", "Density", "Stem-and-leaf", "QQ-plot"]


def _freq_tab(d, v):
    t = d[v].value_counts(sort=False)
    tab = pd.DataFrame({"Kategori": [str(i) for i in t.index], "n": t.values.astype(int)})
    tab["pct"] = 100 * tab["n"] / tab["n"].sum()
    return tab


def plot_uni(d, v, type_):
    if v == "Date":
        dd = d.groupby("Date").size().reset_index(name="n")
        sm_ = smooth_df(dd["Date"].map(pd.Timestamp.toordinal), dd["n"])
        sm_["Date"] = [pd.Timestamp.fromordinal(int(i)) for i in sm_["x"]]
        return (ggplot(dd, aes("Date", "n")) + geom_line(color=pk["l"]) + geom_point(color=pk["m"], size=1.5)
                + geom_line(data=sm_, mapping=aes("Date", "y"), color=pk["d"], size=1.2)
                + labs(title="Jumlah transaksi per hari", subtitle="Garis tebal = tren (loess)", x="Tanggal", y="Jumlah transaksi")
                + theme_pink())
    if v in cat_all:
        tab = _freq_tab(d, v)
        if type_ == "Diagram lingkaran":
            fig, ax = plt.subplots(figsize=(6.4, 5.2))
            cols = pal_cat(len(tab))
            ax.pie(tab["n"], colors=cols, startangle=90, counterclock=False,
                   wedgeprops=dict(edgecolor="white"), autopct=lambda p: f"{round(p, 1)}%",
                   pctdistance=.6, textprops=dict(color="white", fontweight="bold"))
            ax.legend(tab["Kategori"], title=v, loc="center left", bbox_to_anchor=(1, .5), frameon=False)
            ax.set_title(f"Proporsi {v}", loc="left", fontweight="bold", color=pk["d"], fontsize=14)
            fig.tight_layout()
            return fig
        if type_ == "Pareto":
            tab = tab.sort_values("n", ascending=False, kind="stable").reset_index(drop=True)
            tot = tab["n"].sum(); kum = tab["n"].cumsum() / tot * 100
            fig, ax = plt.subplots(figsize=(7, 5))
            ax.bar(tab["Kategori"], tab["n"], color=pk["l"], width=.7)
            ax.set_ylim(0, tot * 1.1); ax.set_ylabel("Frekuensi", color=pk["mauve"], fontweight="bold")
            ax.set_xlabel(v, color=pk["mauve"], fontweight="bold")
            ax.plot(tab["Kategori"], kum * tot / 100, color=pk["d"], lw=2, marker="o")
            for xi, (a, b) in enumerate(zip(kum, kum * tot / 100)):
                ax.annotate(f"{round(a, 1)}%", (xi, b), textcoords="offset points", xytext=(0, 7), ha="center", fontsize=9, color=pk["mauve"])
            ax2 = ax.secondary_yaxis("right", functions=(lambda y: y / tot * 100, lambda y: y * tot / 100))
            ax2.set_ylabel("Kumulatif (%)", color=pk["mauve"], fontweight="bold")
            ax.set_title(f"Pareto {v}", loc="left", fontweight="bold", color=pk["d"], fontsize=14)
            plt.setp(ax.get_xticklabels(), rotation=20, ha="right")
            for s in ("top",): ax.spines[s].set_visible(False)
            fig.tight_layout()
            return fig
        tab["Kategori"] = pd.Categorical(tab["Kategori"], categories=list(tab["Kategori"]))
        return (ggplot(tab, aes("Kategori", "n", fill="Kategori")) + geom_col(show_legend=False, width=.7)
                + geom_text(aes(label="n"), va="bottom", nudge_y=0, color=pk["mauve"], fontweight="bold")
                + scale_fill_manual(values=pal_cat(len(tab))) + scale_y_continuous(expand=(0, 0, .1, 0))
                + labs(title=f"Frekuensi {v}", x=v, y="Frekuensi") + theme_pink() + rot)
    bw = v in ("Quantity", "Hour")

    def hist_l(dens):
        kw = dict(binwidth=1, boundary=.5) if bw else dict(bins=25)
        if dens:
            return geom_histogram(aes(y=after_stat("density")), fill=pk["xl"], color="white", **kw)
        return geom_histogram(fill=pk["xl"], color="white", **kw)

    ds = d.rename(columns=S); vv = S(v)
    g = ggplot(ds, aes(x=vv))
    if type_ == "Histogram + Density":
        return g + hist_l(True) + geom_density(color=pk["d"], size=1.1) + labs(title=f"Histogram dan Density {v}", x=v, y="Densitas") + theme_pink()
    if type_ == "Density":
        return (g + geom_density(fill=pk["l"], color=pk["d"], alpha=.5)
                + geom_vline(xintercept=float(d[v].mean()), linetype="dashed", color=pk["mauve"])
                + labs(title=f"Density {v}", subtitle="Garis putus-putus = rata-rata", x=v, y="Densitas") + theme_pink())
    if type_ == "Boxplot":
        dm = pd.DataFrame({"x": [""], "m": [float(d[v].mean())]})
        return (ggplot(ds.assign(_x=""), aes(x="_x", y=vv)) + geom_boxplot(fill=pk["xl"], color=pk["d"], width=.45, outlier_color=pk["m"])
                + geom_point(data=dm, mapping=aes("x", "m"), shape="D", size=3.5, fill="white", color=pk["d"], inherit_aes=False)
                + coord_flip() + labs(title=f"Boxplot {v}", subtitle="Berlian = rata-rata; titik = pencilan", x="", y=v) + theme_pink())
    if type_ == "QQ-plot":
        return (ggplot(ds, aes(sample=vv)) + stat_qq(color=pk["m"], alpha=.6) + stat_qq_line(color=pk["d"])
                + labs(title=f"QQ-plot {v}", subtitle="Titik mengikuti garis = mendekati normal", x="Kuantil teoretis", y="Kuantil sampel") + theme_pink())
    return g + hist_l(False) + labs(title=f"Distribusi {v}", x=v, y="Frekuensi") + theme_pink()


def uni_table(d, v):
    if v == "Date":
        nd = d["Date"].nunique()
        return pd.DataFrame({
            "Statistik": ["Tanggal awal", "Tanggal akhir", "Jumlah hari", "Jumlah hari dengan transaksi", "Rata-rata transaksi per hari"],
            "Nilai": [str(d["Date"].min().date()), str(d["Date"].max().date()), str((d["Date"].max() - d["Date"].min()).days + 1),
                      str(nd), fmt(len(d) / nd)]})
    if v in cat_all:
        t = _freq_tab(d, v)
        return pd.DataFrame({"Kategori": t["Kategori"], "Frekuensi": t["n"], "Persentase (%)": fmt(t["pct"])})
    x = d[v].to_numpy(dtype=float); q1, q3 = np.quantile(x, [.25, .75]); iqr = q3 - q1
    lo = q1 - 1.5 * iqr; hi = q3 + 1.5 * iqr
    return pd.DataFrame({
        "Statistik": ["N", "Rata-rata", "Median", "Simpangan baku", "Variansi", "Koef. variasi (%)", "Minimum", "Q1", "Q3", "Maksimum",
                      "IQR", "Skewness", "Kurtosis (excess)", "Batas pencilan bawah", "Batas pencilan atas", "Jumlah pencilan"],
        "Nilai": [str(len(x)), fmt(x.mean()), fmt(np.median(x)), fmt(sd1(x)), fmt(np.var(x, ddof=1)), fmt(100 * sd1(x) / x.mean()),
                  fmt(x.min()), fmt(q1), fmt(q3), fmt(x.max()), fmt(iqr), fmt(skew_fun(x), 3), fmt(kurt_fun(x), 3),
                  fmt(lo), fmt(hi), str(int(np.sum((x < lo) | (x > hi))))]})


def interp_uni(d, v):
    if v == "Date":
        dd = d.groupby("Date").size().reset_index(name="n"); top = dd.loc[dd["n"].idxmax()]
        return ["Data mencakup **%s** sampai **%s** (%d hari bertransaksi)." % (d["Date"].min().date(), d["Date"].max().date(), len(dd)),
                "Rata-rata **%s** transaksi per hari; puncak pada **%s** dengan **%d** transaksi." % (fmt(dd["n"].mean()), top["Date"].date(), top["n"]),
                "Garis tren (loess) menunjukkan apakah volume transaksi harian meningkat, menurun, atau relatif stabil sepanjang periode."]
    if v in cat_all:
        t = d[v].value_counts(sort=False).sort_values(ascending=False, kind="stable"); p = 100 * t / t.sum()
        sp = p.max() - p.min()
        return ["Kategori terbanyak: **%s** (%d transaksi; %s%%). Tersedikit: **%s** (%d; %s%%)." % (t.index[0], t.iloc[0], fmt(p.iloc[0], 1), t.index[-1], t.iloc[-1], fmt(p.iloc[-1], 1)),
                "Selisih proporsi tertinggi dan terendah **%s** poin persentase, sehingga sebaran kategori **%s**." % (fmt(sp, 1), "relatif seimbang" if sp < 5 else "tidak seimbang"),
                "Pareto menunjukkan apakah sedikit kategori mendominasi (garis kumulatif cepat mencapai 80%)."]
    x = d[v].to_numpy(dtype=float); m = x.mean(); me = np.median(x); s = sd1(x); sk = skew_fun(x); ku = kurt_fun(x)
    q1, q3 = np.quantile(x, [.25, .75]); iqr = q3 - q1; out = int(np.sum((x < q1 - 1.5 * iqr) | (x > q3 + 1.5 * iqr)))
    if abs(sk) < .5: shape = "hampir simetris"
    elif sk > 0: shape = "miring kuat ke kanan" if sk > 1 else "miring sedang ke kanan"
    else: shape = "miring kuat ke kiri" if sk < -1 else "miring sedang ke kiri"
    if abs(m - me) / s > .1:
        tail = "; rata-rata " + ("lebih besar" if m > me else "lebih kecil") + " dari median sehingga sebaran condong ke " + ("kanan" if m > me else "kiri")
    else:
        tail = "; keduanya berdekatan (sebaran seimbang)"
    kt = "datar (mendekati sebaran seragam)" if ku < -1 else "runcing" if ku > 1 else "mendekati normal"
    return ["Rata-rata **%s** dan median **%s**%s." % (fmt(m), fmt(me), tail),
            "Sebaran data: simpangan baku **%s**, rentang **%s** sampai **%s**, koefisien variasi **%s%%**." % (fmt(s), fmt(x.min()), fmt(x.max()), fmt(100 * s / m)),
            "Bentuk distribusi **%s** (skewness %s); kurtosis %s menunjukkan puncak yang **%s**." % (shape, fmt(sk, 2), fmt(ku, 2), kt),
            "Tidak ditemukan pencilan menurut aturan 1.5 x IQR." if out == 0 else
            "Terdapat **%d pencilan** menurut aturan 1.5 x IQR (%s%% data). Pencilan dipertahankan karena merupakan transaksi yang valid." % (out, fmt(100 * out / len(x), 1))]


# =============================================================================
# 3. BIVARIAT
# =============================================================================
def bi_choices(x, y):
    xc = x in cat_all; yc = y in cat_all
    if not xc and not yc:
        return ["Scatter plot + garis tren", "Scatter plot + kurva linear, kuadratik, nonparametrik", "Line plot rata-rata Y per X", "Correlation plot", "Heatmap kepadatan 2D"]
    elif xc != yc:
        return ["Grouped boxplot", "Grouped violin plot", "Grouped kernel density plot", "Grouped ridgeline plot", "Grouped strip plot",
                "Grouped mean/SEM plot", "Bar chart statistik ringkasan", "Bar chart rata-rata (+/- SE)", "Cleveland dot chart"]
    return ["Stacked bar chart", "Grouped bar chart", "Segmented bar chart (100%)", "Heatmap tabulasi silang"]


# Correlation plot: lingkaran (segitiga bawah) + angka r (segitiga atas)
def plot_corr_circle(d, vars_=None):
    vars_ = vars_ or num_all
    m = d[vars_].corr().to_numpy(); rows = []
    for a, v1 in enumerate(vars_):
        for b, v2 in enumerate(vars_):
            rows.append(dict(V1=v1, V2=v2, r=m[a, b], i=a + 1, j=b + 1))
    df = pd.DataFrame(rows)
    df["V1"] = pd.Categorical(df["V1"], categories=vars_); df["V2"] = pd.Categorical(df["V2"], categories=vars_)
    df["ar"] = df["r"].abs(); df["lab"] = df["r"].map(lambda z: f"{z:.2f}"); df["vl"] = df["V1"].astype(str)
    return (ggplot()
            + geom_point(data=df[df["i"] > df["j"]], mapping=aes("V1", "V2", size="ar", fill="r"), shape="o", color=pk["grid"])
            + geom_text(data=df[df["i"] < df["j"]], mapping=aes("V1", "V2", label="lab"), color=pk["txt"], fontweight="bold")
            + geom_text(data=df[df["i"] == df["j"]], mapping=aes("V1", "V2", label="vl"), color=pk["d"], fontweight="bold")
            + scale_size_continuous(range=(3, 22), limits=(0, 1), guide=None)
            + scale_fill_gradient2(low=pk["slate"], mid="white", high=pk["s"], limits=(-1, 1), midpoint=0, name="r")
            + scale_x_discrete(limits=vars_) + scale_y_discrete(limits=list(reversed(vars_)))
            + labs(title="Correlation plot (Pearson)", subtitle="Lingkaran: ukuran dan warna = kekuatan korelasi | Angka: nilai r", x="", y="")
            + theme_pink() + theme(panel_grid=element_blank()))


def plot_bi(d, x, y, type_):
    xc = x in cat_all; yc = y in cat_all
    # ---------- numerik x numerik ----------
    if not xc and not yc:
        if type_ == "Correlation plot":
            return plot_corr_circle(d, num_all)
        ds = d.rename(columns=S); xs, ys = S(x), S(y)
        if type_ == "Scatter plot + kurva linear, kuadratik, nonparametrik":
            cols = {"Linear": pk["d"], "Kuadratik": pk["slate"], "Nonparametrik (LOESS)": "#2E7D32"}
            sm_ = smooth_df(d[x], d[y])
            xs_ = np.sort(d[x].to_numpy(float)); q = np.polyval(np.polyfit(d[x], d[y], 2), xs_); l = np.polyval(np.polyfit(d[x], d[y], 1), xs_)
            cur = pd.concat([pd.DataFrame({"x": xs_, "y": l, "Kurva": "Linear"}),
                             pd.DataFrame({"x": xs_, "y": q, "Kurva": "Kuadratik"}),
                             pd.DataFrame({"x": sm_["x"], "y": sm_["y"], "Kurva": "Nonparametrik (LOESS)"})])
            cur["Kurva"] = pd.Categorical(cur["Kurva"], categories=list(cols))
            return (ggplot(ds, aes(xs, ys)) + geom_point(color=pk["l"], alpha=.5, size=1.8)
                    + geom_line(data=cur, mapping=aes("x", "y", color="Kurva"), size=1.1, inherit_aes=False)
                    + scale_colour_manual(values=cols, name="Kurva")
                    + labs(title=f"{y} terhadap {x}", subtitle="Perbandingan kurva linear, kuadratik, dan nonparametrik (LOESS)", x=x, y=y) + theme_pink())
        if type_ == "Line plot rata-rata Y per X":
            xv = d[x]; yv = d[y]
            grp = xv if xv.nunique() <= 12 else pd.cut(xv, bins=10)
            g = pd.DataFrame({"x": xv, "y": yv, "g": grp}).groupby("g", observed=True)
            agg = pd.DataFrame({"gx": g["x"].mean(), "my": g["y"].mean(), "se": g["y"].apply(lambda z: z.std(ddof=1) / np.sqrt(len(z)))}).reset_index(drop=True)
            agg["lo"] = agg["my"] - agg["se"]; agg["hi"] = agg["my"] + agg["se"]
            return (ggplot(agg, aes("gx", "my")) + geom_ribbon(aes(ymin="lo", ymax="hi"), fill=pk["xl"], alpha=.8)
                    + geom_line(color=pk["d"], size=1) + geom_point(color=pk["m"], size=2.5)
                    + labs(title=f"Rata-rata {y} menurut {x}", subtitle="Pita = rata-rata +/- 1 standard error", x=x, y=f"Rata-rata {y}") + theme_pink())
        if type_ == "Heatmap kepadatan 2D":
            return (ggplot(ds, aes(xs, ys)) + geom_bin_2d(bins=20) + scale_fill_gradient(low=pk["bg"], high=pk["d"], name="Jumlah")
                    + labs(title=f"Kepadatan {x} dan {y}", x=x, y=y) + theme_pink())
        return (ggplot(ds, aes(xs, ys)) + geom_point(color=pk["m"], alpha=.5, size=1.8)
                + geom_smooth(method="lm", formula="y ~ x", se=False, color=pk["d"], size=1)
                + labs(title=f"{y} terhadap {x}", subtitle="Scatterplot dengan garis regresi linear sederhana", x=x, y=y) + theme_pink())
    # ---------- kategorik x numerik ----------
    if xc != yc:
        cv = x if xc else y; nv = y if xc else x; k = d[cv].nunique(); cols = pal_cat(k)
        d2 = pd.DataFrame({"g": d[cv].to_numpy(), "v": d[nv].to_numpy()}); d2["g"] = pd.Categorical(d2["g"], categories=list(d[cv].cat.categories))
        gb = d2.groupby("g", observed=True)["v"]
        agg = pd.DataFrame({"m": gb.mean(), "se": gb.std(ddof=1) / np.sqrt(gb.size()), "md": gb.median(), "s": gb.std(ddof=1), "mn": gb.min(), "mx": gb.max()}).reset_index()
        agg["g"] = pd.Categorical(agg["g"].astype(str), categories=[str(c) for c in d[cv].cat.categories])
        d2["g"] = pd.Categorical(d2["g"].astype(str), categories=[str(c) for c in d[cv].cat.categories])
        agg["lo"] = agg["m"] - agg["se"]; agg["hi"] = agg["m"] + agg["se"]
        dm = agg[["g", "m"]]
        if type_ == "Grouped violin plot":
            return (ggplot(d2, aes("g", "v", fill="g")) + geom_violin(show_legend=False, alpha=.75, color=pk["mauve"])
                    + geom_boxplot(width=.12, fill="white", outlier_alpha=0) + scale_fill_manual(values=cols)
                    + labs(title=f"Grouped violin plot: {nv} menurut {cv}", x=cv, y=nv) + theme_pink() + rot)
        if type_ == "Grouped kernel density plot":
            return (ggplot(d2, aes("v", fill="g", color="g")) + geom_density(alpha=.35) + scale_fill_manual(values=cols) + scale_colour_manual(values=cols)
                    + labs(title=f"Grouped kernel density plot: {nv} per {cv}", x=nv, y="Densitas", fill=cv, color=cv) + theme_pink())
        if type_ == "Grouped ridgeline plot":
            lo_, hi_ = d2["v"].min(), d2["v"].max(); lv = list(d2["g"].cat.categories); xs_ = np.linspace(lo_, hi_, 256); parts = []
            for i, l in enumerate(lv, start=1):
                z = d2.loc[d2["g"] == l, "v"].to_numpy(float)
                parts.append(pd.DataFrame({"g": l, "i": i, "x": xs_, "y": sps.gaussian_kde(z)(xs_)}))
            rd = pd.concat(parts); rd["h"] = rd["y"] / rd["y"].max() * 1.6; rd["ymax"] = rd["i"] + rd["h"]
            rd["g"] = pd.Categorical(rd["g"], categories=lv)
            return (ggplot(rd, aes("x", group="g")) + geom_ribbon(aes(ymin="i", ymax="ymax", fill="g"), color=pk["mauve"], alpha=.8, show_legend=False)
                    + scale_fill_manual(values=cols) + scale_y_continuous(breaks=list(range(1, len(lv) + 1)), labels=lv)
                    + labs(title=f"Grouped ridgeline plot: {nv} per {cv}", x=nv, y=cv) + theme_pink())
        if type_ == "Grouped strip plot":
            return (ggplot(d2, aes("g", "v", color="g"))
                    + geom_jitter(width=.22, height=.15 if nv in ("Quantity", "Hour") else 0, alpha=.45, size=1.6, show_legend=False)
                    + geom_point(data=dm, mapping=aes("g", "m"), shape="D", size=4, fill="white", color=pk["d"], inherit_aes=False)
                    + scale_colour_manual(values=cols)
                    + labs(title=f"Grouped strip plot: {nv} menurut {cv}", subtitle="Titik = observasi; berlian = rata-rata", x=cv, y=nv) + theme_pink() + rot)
        if type_ == "Grouped mean/SEM plot":
            return (ggplot(agg, aes("g", "m", color="g")) + geom_errorbar(aes(ymin="lo", ymax="hi"), width=.15, size=.9, show_legend=False)
                    + geom_point(size=3.8, show_legend=False) + scale_colour_manual(values=cols)
                    + labs(title=f"Grouped mean/SEM plot: {nv} menurut {cv}", subtitle="Titik = rata-rata; garis = +/- 1 standard error (SEM)", x=cv, y=f"Rata-rata {nv}") + theme_pink() + rot)
        if type_ == "Bar chart statistik ringkasan":
            st = ["Rata-rata", "Median", "Simpangan baku", "Minimum", "Maksimum"]
            long = pd.DataFrame({"g": list(agg["g"].astype(str)) * 5, "stat": np.repeat(st, len(agg)),
                                 "val": np.concatenate([agg["m"], agg["md"], agg["s"], agg["mn"], agg["mx"]])})
            long["g"] = pd.Categorical(long["g"], categories=[str(c) for c in d[cv].cat.categories]); long["stat"] = pd.Categorical(long["stat"], categories=st)
            long["lab"] = long["val"].round(1)
            return (ggplot(long, aes("g", "val", fill="g")) + geom_col(show_legend=False, width=.7) + geom_text(aes(label="lab"), va="bottom", size=8, color=pk["txt"])
                    + facet_wrap("stat", scales="free_y") + scale_fill_manual(values=cols) + scale_y_continuous(expand=(0, 0, .15, 0))
                    + labs(title=f"Bar chart statistik ringkasan {nv} menurut {cv}", x=cv, y=nv) + theme_pink(11) + rot)
        if type_ == "Bar chart rata-rata (+/- SE)":
            agg["lab"] = agg["m"].round(2)
            return (ggplot(agg, aes("g", "m", fill="g")) + geom_col(show_legend=False, width=.7) + geom_errorbar(aes(ymin="lo", ymax="hi"), width=.2, color=pk["txt"])
                    + geom_text(aes(label="lab"), va="top", nudge_y=-0.02 * float(agg["m"].max()), color="white", fontweight="bold") + scale_fill_manual(values=cols)
                    + labs(title=f"Rata-rata {nv} menurut {cv}", subtitle="Garis = +/- 1 standard error", x=cv, y=f"Rata-rata {nv}") + theme_pink() + rot)
        if type_ == "Cleveland dot chart":
            a2 = agg.sort_values("m").reset_index(drop=True); lv = [str(g) for g in a2["g"]]
            cl = pd.DataFrame({"g": pd.Categorical(lv * 2, categories=lv), "stat": ["Rata-rata"] * len(a2) + ["Median"] * len(a2), "val": list(a2["m"]) + list(a2["md"])})
            pad = 0.15 * max(np.ptp(cl["val"]), abs(cl["val"].mean()) * .05); lo_ = cl["val"].min() - pad; hi_ = cl["val"].max() + pad
            sg = pd.DataFrame({"g": pd.Categorical(lv, categories=lv), "lo": lo_, "hi": hi_})
            return (ggplot(cl, aes("val", "g")) + geom_segment(data=sg, mapping=aes(x="lo", xend="hi", y="g", yend="g"), inherit_aes=False, color=pk["grid"], linetype="dotted")
                    + geom_point(aes(fill="stat"), shape="o", size=4, color=pk["d"]) + scale_fill_manual(values={"Rata-rata": pk["m"], "Median": "white"}, name="")
                    + scale_x_continuous(limits=(lo_, hi_)) + labs(title=f"Cleveland dot chart: {nv} menurut {cv}", subtitle="Kategori diurutkan menurut rata-rata", x=nv, y=cv)
                    + theme_pink() + theme(legend_position="bottom"))
        return (ggplot(d2, aes("g", "v", fill="g")) + geom_boxplot(show_legend=False, alpha=.85, color=pk["mauve"], outlier_color=pk["m"])
                + geom_point(data=dm, mapping=aes("g", "m"), shape="D", size=3, fill="white", color="black", inherit_aes=False) + scale_fill_manual(values=cols)
                + labs(title=f"Grouped boxplot: {nv} menurut {cv}", subtitle="Berlian = rata-rata", x=cv, y=nv) + theme_pink() + rot)
    # ---------- kategorik x kategorik ----------
    ny = d[y].nunique()
    ct = pd.crosstab(d[x], d[y])
    df = ct.stack().reset_index(); df.columns = ["X", "Y", "n"]
    df["X"] = pd.Categorical(df["X"].astype(str), categories=[str(c) for c in ct.index]); df["Y"] = pd.Categorical(df["Y"].astype(str), categories=[str(c) for c in ct.columns])
    df["pct"] = 100 * df["n"] / df.groupby("X", observed=True)["n"].transform("sum")
    if type_ == "Heatmap tabulasi silang":
        tab = df.rename(columns={"n": "Freq"}); mid = tab["Freq"].max() / 2
        tab["wh"] = np.where(tab["Freq"] > mid, "TRUE", "FALSE")
        return (ggplot(tab, aes("X", "Y", fill="Freq")) + geom_tile(color="white") + geom_text(aes(label="Freq", color="wh"), show_legend=False, fontweight="bold")
                + scale_colour_manual(values={"TRUE": "white", "FALSE": pk["txt"]}) + scale_fill_gradient(low=pk["bg"], high=pk["d"], name="Jumlah")
                + labs(title=f"{x} dan {y}", x=x, y=y) + theme_pink() + rot)
    if type_ == "Stacked bar chart":
        return (ggplot(df, aes("X", "n", fill="Y")) + geom_col(position="stack", color="white") + scale_fill_manual(values=pal_cat(ny))
                + labs(title=f"Stacked bar chart: {y} dalam {x}", subtitle="Tinggi total batang = jumlah transaksi per kategori", x=x, y="Jumlah transaksi", fill=y) + theme_pink() + rot)
    if type_ == "Segmented bar chart (100%)":
        return (ggplot(df, aes("X", "pct", fill="Y")) + geom_col(position="stack", color="white") + scale_fill_manual(values=pal_cat(ny))
                + scale_y_continuous(labels=lambda z: [f"{v:g}%" for v in z], expand=(0, 0, .02, 0))
                + labs(title=f"Segmented bar chart: {y} dalam {x}", subtitle="Setiap batang = 100%; bandingkan komposisi antar kategori", x=x, y="Persentase", fill=y) + theme_pink() + rot)
    return (ggplot(df, aes("X", "n", fill="Y")) + geom_col(position=position_dodge(.9)) + geom_text(aes(label="n"), position=position_dodge(.9), va="bottom", size=8, color=pk["txt"])
            + scale_fill_manual(values=pal_cat(ny)) + scale_y_continuous(expand=(0, 0, .1, 0))
            + labs(title=f"Grouped bar chart: {x} dan {y}", x=x, y="Jumlah transaksi", fill=y) + theme_pink() + rot)


def bi_text(d, x, y):
    xc = x in cat_all; yc = y in cat_all
    if not xc and not yc:
        r, pr = sps.pearsonr(d[x], d[y]); rho, ps = sps.spearmanr(d[x], d[y]); lr = sps.linregress(d[x], d[y]); b = lr.slope
        tab = pd.DataFrame({"Ukuran": ["Korelasi Pearson (r)", "p-value Pearson", "Korelasi Spearman (rho)", "p-value Spearman", "Kemiringan garis (b1)", "R-kuadrat regresi sederhana"],
                            "Nilai": [fmt(r, 3), fmt_p(pr), fmt(rho, 3), fmt_p(ps), fmt(b, 3), fmt(lr.rvalue ** 2, 3)]})
        lines = ["Hubungan linear **%s** dan **%s** (r = %s; p %s)." % (strength(r), "positif" if r > 0 else "negatif", fmt(r, 3), "< 0.001" if pr < .001 else "= " + fmt_p(pr)),
                 "Hubungan **signifikan** pada taraf 5%%: setiap kenaikan 1 satuan **%s**, rata-rata **%s** berubah sebesar **%s** satuan." % (x, y, fmt(b, 3)) if pr < .05
                 else "Hubungan **tidak signifikan** pada taraf 5%, sehingga tidak ada bukti hubungan linear antar kedua variabel.",
                 "Korelasi menunjukkan keterkaitan, bukan sebab-akibat."]
        return dict(table=tab, lines=lines)
    if xc != yc:
        cv = x if xc else y; nv = y if xc else x
        gb = d.groupby(cv, observed=True)[nv]
        g = pd.DataFrame({cv: [str(i) for i in gb.mean().index], "Rata-rata": fmt(gb.mean()), "Median": fmt(gb.median()), "Simpangan baku": fmt(gb.std(ddof=1)), "n": gb.size().to_numpy()})
        gm = pd.DataFrame({"grp": [str(i) for i in gb.mean().index], "m": gb.mean().to_numpy()})
        groups = [v.to_numpy(float) for _, v in gb]; allv = d[nv].to_numpy(float)
        ssb = sum(len(v) * (v.mean() - allv.mean()) ** 2 for v in groups); sst = np.sum((allv - allv.mean()) ** 2)
        Fv, p = sps.f_oneway(*groups); eta = ssb / sst; kw = sps.kruskal(*groups).pvalue
        if d[cv].nunique() == 2:
            tt = sps.ttest_ind(groups[0], groups[1], equal_var=False); p = tt.pvalue
            test = "uji t Welch (t = %s)" % fmt(tt.statistic, 2)
        else:
            test = "ANOVA (F = %s)" % fmt(Fv, 2)
        lines = ["Rata-rata **%s** tertinggi pada **%s** (%s) dan terendah pada **%s** (%s)." % (nv, gm["grp"][gm["m"].idxmax()], fmt(gm["m"].max()), gm["grp"][gm["m"].idxmin()], fmt(gm["m"].min())),
                 "Hasil %s: p %s; uji Kruskal-Wallis: p %s; ukuran efek eta-kuadrat = %s." % (test, pq(p), pq(kw), fmt(eta, 3)),
                 "Perbedaan rata-rata antar kelompok **signifikan** pada taraf 5%." if p < .05 else "Perbedaan rata-rata antar kelompok **tidak signifikan** pada taraf 5%; selisih yang tampak pada grafik dapat terjadi karena kebetulan."]
        return dict(table=g, lines=lines)
    tab = pd.crosstab(d[x], d[y]); ch = sps.chi2_contingency(tab.to_numpy()); chi2, pch, dof = ch[0], ch[1], ch[2]
    V = math.sqrt(chi2 / (tab.to_numpy().sum() * (min(tab.shape) - 1)))
    j, i = np.argwhere(tab.to_numpy().T == tab.to_numpy().max())[0]      # urutan kolom-utama seperti R
    out = tab.copy(); out.columns = [str(c) for c in out.columns]; out.insert(0, x, [str(r) for r in tab.index]); out = out.reset_index(drop=True)
    lines = ["Kombinasi terbanyak: **%s** x **%s** (%d transaksi)." % (tab.index[i], tab.columns[j], tab.to_numpy().max()),
             "Uji Chi-square: X2 = %s, df = %d, p %s; Cramer's V = %s." % (fmt(chi2, 2), dof, pq(pch), fmt(V, 3)),
             "Ada hubungan **signifikan** antar kedua variabel kategorik pada taraf 5%." if pch < .05 else "**Tidak ada** hubungan signifikan antar kedua variabel kategorik pada taraf 5%."]
    return dict(table=out, lines=lines)


# =============================================================================
# 4. MULTIVARIAT
# =============================================================================
def plot_mv_scatter(d, x, y, col, size, facet, lines):
    ds = d.rename(columns=S); xs, ys = S(x), S(y)
    g = ggplot(ds, aes(xs, ys))
    hasc = col != NONE; hass = size != NONE
    if hasc and hass:
        g = g + geom_point(aes(color=S(col), size=S(size)), alpha=.6)
    elif hasc:
        g = g + geom_point(aes(color=S(col)), alpha=.6, size=2)
    elif hass:
        g = g + geom_point(aes(size=S(size)), color=pk["m"], alpha=.5)
    else:
        g = g + geom_point(color=pk["m"], alpha=.6, size=2)
    if hasc:
        g = g + (scale_colour_gradient(low=pk["xl"], high=pk["d"], name=col) if col in num_all
                 else scale_colour_manual(values=pal_cat(d[col].nunique()), name=col))
    if hass:
        g = g + scale_size_continuous(range=(1, 5), name=size)
    if lines:
        if hasc and col not in num_all:
            g = g + geom_smooth(aes(color=S(col)), method="lm", formula="y ~ x", se=False, size=.9)
        else:
            g = g + geom_smooth(method="lm", formula="y ~ x", se=False, color=pk["d"], size=.9)
    if facet != NONE:
        g = g + facet_wrap(S(facet))
    return g + labs(title=f"{y} terhadap {x}", subtitle=f"Warna: {col} | Ukuran: {size} | Panel: {facet}", x=x, y=y) + theme_pink()


def mv_heat_data(d, a, b, num, fun):
    t = pd.DataFrame({"A": d[a].to_numpy(), "B": d[b].to_numpy()})
    t["A"] = pd.Categorical(t["A"], categories=list(d[a].cat.categories)); t["B"] = pd.Categorical(t["B"], categories=list(d[b].cat.categories))
    if fun == "Banyak transaksi":
        res = t.groupby(["A", "B"], observed=True).size().reset_index(name="val")
    else:
        t["v"] = d[num].to_numpy()
        f = {"Rata-rata": "mean", "Median": "median", "Jumlah": "sum"}[fun]
        res = t.groupby(["A", "B"], observed=True)["v"].agg(f).reset_index(name="val")
    return res


def plot_mv_heat(d, a, b, num, fun):
    res = mv_heat_data(d, a, b, num, fun); mid = (res["val"].max() + res["val"].min()) / 2
    lab = "Banyak transaksi" if fun == "Banyak transaksi" else f"{fun} {num}"
    res["A"] = pd.Categorical(res["A"].astype(str), categories=[str(c) for c in d[a].cat.categories])
    res["B"] = pd.Categorical(res["B"].astype(str), categories=[str(c) for c in d[b].cat.categories])
    res["lab"] = fmt(res["val"], 0 if fun == "Banyak transaksi" else 1); res["wh"] = np.where(res["val"] > mid, "TRUE", "FALSE")
    return (ggplot(res, aes("A", "B", fill="val")) + geom_tile(color="white") + geom_text(aes(label="lab", color="wh"), show_legend=False, fontweight="bold")
            + scale_colour_manual(values={"TRUE": "white", "FALSE": pk["txt"]}) + scale_fill_gradient(low=pk["bg"], high=pk["d"], name=lab)
            + labs(title=lab, subtitle=f"menurut {a} dan {b}", x=a, y=b) + theme_pink() + rot)


def plot_mv_cor(d, vars_, method):
    m = d[list(vars_)].corr(method=method).to_numpy(); rows = []
    for a, v1 in enumerate(vars_):
        for b, v2 in enumerate(vars_):
            rows.append(dict(V1=v1, V2=v2, r=m[a, b], lab=f"{m[a, b]:.2f}"))
    df = pd.DataFrame(rows); df["V1"] = pd.Categorical(df["V1"], categories=list(vars_)); df["V2"] = pd.Categorical(df["V2"], categories=list(vars_))
    return (ggplot(df, aes("V1", "V2", fill="r")) + geom_tile(color="white") + geom_text(aes(label="lab"), color=pk["txt"], fontweight="bold", size=11)
            + scale_fill_gradient2(low=pk["slate"], mid="white", high=pk["s"], limits=(-1, 1), midpoint=0, name="r")
            + labs(title=f"Matriks korelasi {method.title()}", x="", y="") + theme_pink() + theme(panel_grid=element_blank()))


def trend_data(d, metric, unit, grp):
    d = d.copy()
    if unit == "Harian":
        d["Periode"] = d["Date"]
    elif unit == "Mingguan":
        d["Periode"] = d["Date"] - pd.to_timedelta(d["Date"].dt.weekday, unit="D")
    else:
        d["Periode"] = d["Date"].dt.to_period("M").dt.to_timestamp()
    gl = ["Periode"] if grp == NONE else ["Periode", grp]
    gb = d.groupby(gl, observed=True)
    if metric == "Jumlah transaksi":
        out = gb.size()
    else:
        col, fn = {"Total penjualan (jumlah)": ("Total", "sum"), "Rata-rata Total": ("Total", "mean"),
                   "Rata-rata Rating": ("Rating", "mean"), "Total Quantity": ("Quantity", "sum")}[metric]
        out = gb[col].agg(fn)
    return out.reset_index(name="val")


def plot_mv_trend(d, metric, unit, grp):
    td = trend_data(d, metric, unit, grp)
    if grp == NONE:
        g = ggplot(td, aes("Periode", "val")) + geom_line(color=pk["d"], size=1) + geom_point(color=pk["m"])
    else:
        td = td.rename(columns={grp: "Kelompok"}); td["Kelompok"] = td["Kelompok"].astype(str)
        g = (ggplot(td, aes("Periode", "val", color="Kelompok", group="Kelompok")) + geom_line(size=1) + geom_point()
             + scale_colour_manual(values=pal_cat(td["Kelompok"].nunique())))
    return (g + labs(title=f"{metric} {unit} " + (f"menurut {grp}" if grp != NONE else ""),
                     subtitle="Minggu pertama/terakhir dapat tidak lengkap" if unit == "Mingguan" else None,
                     x="Periode", y=metric, color=grp if grp != NONE else "") + theme_pink())


def pca_fit(d, vars_):
    X = d[list(vars_)].to_numpy(float); Xs = (X - X.mean(axis=0)) / X.std(axis=0, ddof=1)
    U, s, Vt = np.linalg.svd(Xs, full_matrices=False); names = [f"PC{i + 1}" for i in range(len(s))]
    return dict(x=pd.DataFrame(Xs @ Vt.T, columns=names), sdev=s / np.sqrt(len(X) - 1), rotation=pd.DataFrame(Vt.T, index=list(vars_), columns=names))


def plot_mv_pca(d, vars_, colv):
    pr = pca_fit(d, vars_); sc = pr["x"][["PC1", "PC2"]].copy(); ve = 100 * pr["sdev"] ** 2 / np.sum(pr["sdev"] ** 2)
    ld = pr["rotation"][["PC1", "PC2"]].copy(); k = min(sc["PC1"].abs().max(), sc["PC2"].abs().max()) * 0.85
    ld["PC1"] = ld["PC1"] * k; ld["PC2"] = ld["PC2"] * k; ld["v"] = ld.index; ld["x0"] = 0.0; ld["y0"] = 0.0
    if colv != NONE:
        sc["Grup"] = d[colv].astype(str).to_numpy()
        g = (ggplot(sc, aes("PC1", "PC2", color="Grup")) + geom_point(alpha=.6, size=1.8)
             + scale_colour_manual(values=pal_cat(d[colv].nunique()), name=colv))
    else:
        g = ggplot(sc, aes("PC1", "PC2")) + geom_point(color=pk["m"], alpha=.5, size=1.8)
    return (g + geom_segment(data=ld, mapping=aes(x="x0", y="y0", xend="PC1", yend="PC2"), inherit_aes=False, arrow=arrow(length=.08), color=pk["d"], size=.8)
            + geom_text(data=ld, mapping=aes("PC1", "PC2", label="v"), inherit_aes=False, color=pk["d"], fontweight="bold", va="bottom")
            + labs(title="PCA biplot", x="PC1 (%.1f%%)" % ve[0], y="PC2 (%.1f%%)" % ve[1]) + theme_pink())


# ---- Scatterplot matrix, Mahalanobis, Q-Q Chi-Square ----
def plot_mv_pairs(d, vars_, colv):
    vars_ = list(vars_); p = len(vars_); grp = d[colv].astype(str).to_numpy() if colv != NONE else None
    levels = list(d[colv].cat.categories) if colv != NONE else []
    if grp is None:
        cols = [to_rgba(pk["m"], .55)] * len(d)
    else:
        pc = pal_cat(len(levels)); cmap = {str(l): to_rgba(pc[i], .55) for i, l in enumerate(levels)}; cols = [cmap[g] for g in grp]
    fig, axes = plt.subplots(p, p, figsize=(9, 8.2), squeeze=False)
    for i in range(p):
        for j in range(p):
            ax = axes[i, j]; ax.set_xticks([]); ax.set_yticks([])
            if i == j:
                ax.hist(d[vars_[i]], bins="sturges", color=pk["xl"], edgecolor="white")
                ax.text(.5, .93, vars_[i], transform=ax.transAxes, ha="center", va="top", fontweight="bold", fontsize=12)
            elif i > j:
                xv = d[vars_[j]].to_numpy(float); yv = d[vars_[i]].to_numpy(float)
                ax.scatter(xv, yv, s=6, c=cols, linewidths=0)
                lw = lowess(yv, xv, frac=2 / 3, return_sorted=True); ax.plot(lw[:, 0], lw[:, 1], color=pk["d"], lw=1.6)
            else:
                r = np.corrcoef(d[vars_[j]], d[vars_[i]])[0, 1]
                ax.text(.5, .5, f"r = {r:.2f}", transform=ax.transAxes, ha="center", va="center", color=pk["d"], fontweight="bold", fontsize=11 + 8 * abs(r))
    fig.suptitle("Scatterplot matrix (pairs plot)", color=pk["d"], fontsize=14, fontweight="bold")
    if grp is not None:
        pc = pal_cat(len(levels))
        handles = [plt.Line2D([], [], marker="o", ls="", color=pc[i], label=str(l)) for i, l in enumerate(levels)]
        fig.legend(handles=handles, loc="lower center", ncol=len(levels), frameon=False, title=colv)
    fig.tight_layout(rect=(0, .06 if grp is not None else 0, 1, .96))
    return fig


def mahal_fit(d, vars_):
    X = d[list(vars_)].to_numpy(float); p = len(vars_); D2 = None
    try:
        Sinv = np.linalg.inv(np.cov(X, rowvar=False).reshape(p, p)); df_ = X - X.mean(axis=0)
        D2 = np.einsum("ij,jk,ik->i", df_, Sinv, df_)
    except np.linalg.LinAlgError:
        D2 = None
    validate(D2 is not None, "Matriks kovarians singular (ada variabel yang bergantung linear). Kurangi variabel yang dipilih.")
    cut = sps.chi2.ppf(.975, df=p)
    return dict(df=pd.DataFrame({"idx": np.arange(1, len(X) + 1), "D2": D2, "pval": sps.chi2.sf(D2, p), "out": D2 > cut}), cut=cut, p=p, X=X)


def plot_mv_mahal_idx(f):
    df = f["df"].copy(); df["Status"] = pd.Categorical(np.where(df["out"], "Pencilan", "Normal"), categories=["Normal", "Pencilan"])
    o = df[df["out"]]; o = o.sort_values("D2", ascending=False).head(5)
    return (ggplot(df, aes("idx", "D2", color="Status")) + geom_point(alpha=.7, size=1.6) + geom_hline(yintercept=f["cut"], linetype="dashed", color=pk["d"])
            + geom_text(data=o, mapping=aes(label="idx"), va="bottom", nudge_y=0.02 * float(df["D2"].max()), size=8, show_legend=False)
            + scale_colour_manual(values={"Normal": pk["l"], "Pencilan": pk["d"]}, drop=False)
            + labs(title="Jarak Mahalanobis (D2) per observasi",
                   subtitle="Garis putus-putus = batas Chi-Square 97.5%% (df = %d): %.2f | Pencilan: %d (%.1f%%)" % (f["p"], f["cut"], df["out"].sum(), 100 * df["out"].mean()),
                   x="Indeks observasi", y="D2 (kuadrat jarak Mahalanobis)") + theme_pink(11))


def plot_mv_mahal_sc(f, vars_):
    df = pd.DataFrame({"x": f["X"][:, 0], "y": f["X"][:, 1], "Status": pd.Categorical(np.where(f["df"]["out"], "Pencilan", "Normal"), categories=["Normal", "Pencilan"])})
    return (ggplot(df, aes("x", "y", color="Status", size="Status")) + geom_point(alpha=.65)
            + scale_colour_manual(values={"Normal": pk["l"], "Pencilan": pk["d"]}, drop=False) + scale_size_manual(values={"Normal": 1.6, "Pencilan": 2.6}, guide=None)
            + labs(title="Pencilan multivariat pada ruang dua variabel", subtitle="Ditampilkan: " + vars_[0] + " dan " + vars_[1] + " (dua variabel pertama yang dipilih)", x=vars_[0], y=vars_[1]) + theme_pink(11))


def mahal_table(d, f, vars_):
    df = f["df"]; o = df.sort_values("D2", ascending=False).head(min(10, len(df)))
    t = pd.DataFrame({"Baris": o["idx"].to_numpy(), "D2": fmt(o["D2"].to_numpy(), 2), "p-value": fmt_p(o["pval"].to_numpy()), "Status": np.where(o["out"], "Pencilan", "Normal")})
    for v in vars_:
        t[v] = fmt(d[v].to_numpy()[o["idx"].to_numpy() - 1], 1)
    return t


def plot_mv_qqchi(f):
    n = len(f["df"]); o = np.argsort(f["df"]["D2"].to_numpy(), kind="stable")
    s = pd.DataFrame({"D2": f["df"]["D2"].to_numpy()[o], "idx": f["df"]["idx"].to_numpy()[o], "out": f["df"]["out"].to_numpy()[o]})
    s["q"] = sps.chi2.ppf((np.arange(1, n + 1) - .5) / n, df=f["p"])
    s["Status"] = pd.Categorical(np.where(s["out"], "Pencilan", "Normal"), categories=["Normal", "Pencilan"])
    return (ggplot(s, aes("q", "D2", color="Status")) + geom_point(alpha=.7, size=1.8) + geom_abline(slope=1, intercept=0, color=pk["d"], size=.9)
            + geom_text(data=s.tail(3), mapping=aes(label="idx"), va="bottom", nudge_y=0.02 * float(s["D2"].max()), size=8, show_legend=False)
            + scale_colour_manual(values={"Normal": pk["m"], "Pencilan": pk["d"]}, drop=False)
            + labs(title="Q-Q plot Chi-Square", subtitle="D2 terurut vs kuantil Chi-Square (df = %d); titik di sekitar garis = mendekati normal multivariat" % f["p"],
                   x="Kuantil teoretis Chi-Square", y="D2 terurut (sampel)") + theme_pink(11))


# =============================================================================
# 5. REGRESI LINEAR BERGANDA
# =============================================================================
UNS = {S(v): v for v in cat_all + num_all}          # nama aman -> nama asli


def clean(x):
    return x.replace("`", "")


def term_of(m, cn):
    return UNS.get(cn.split("[")[0], cn.split("[")[0])


def coef_label(m, cn):
    if cn == "Intercept":
        return "Intersep"
    tl = term_of(m, cn)
    mt = re.search(r"\[T\.(.*)\]$", cn)
    return tl if mt is None else f"{tl} = {mt.group(1)}"


def build_model(d, resp, preds, logy):
    """lm(resp ~ prediktor): matriks desain dibuat manual agar urutan suku sama dengan R
    (kategorik -> dummy, kategori pertama sebagai acuan)."""
    ds = d.rename(columns=S); X = pd.DataFrame({"Intercept": np.ones(len(ds))}, index=ds.index); slices = {}
    for p in preds:
        sp = S(p); start = X.shape[1]
        if p in cat_all:
            for l in [str(c) for c in d[p].cat.categories][1:]:
                X[f"{sp}[T.{l}]"] = (ds[sp].astype(str) == l).astype(float)
        else:
            X[sp] = ds[sp].astype(float)
        slices[sp] = (start, X.shape[1])
    y = np.log(ds[S(resp)].astype(float)) if logy else ds[S(resp)].astype(float)
    m = sm.OLS(y.to_numpy(), X).fit()
    m.xlevels = {S(p): [str(c) for c in d[p].cat.categories] for p in preds if p in cat_all}
    m.term_slices = slices; m.X = X; m.y = y.to_numpy()
    return m


def make_equation(m, resp, logy):
    b = m.params; lhs = f"ln({resp})" if logy else resp
    ln = ["%s-hat = %s" % (lhs, fmt(b.iloc[0], 4))]
    for n in b.index[1:]:
        ln.append("   %s %s x [%s]" % ("+" if b[n] >= 0 else "-", fmt(abs(b[n]), 4), coef_label(m, n)))
    return "\n".join(ln)


def vif_table(m):
    try:
        terms = list(m.term_slices)
        if len(terms) < 2:
            return None
        mm = m.X.to_numpy()[:, 1:]; R = np.corrcoef(mm, rowvar=False); dR = np.linalg.det(R); rows = []
        for t in terms:
            a, b = m.term_slices[t]; i = list(range(a - 1, b - 1)); rest = [k for k in range(R.shape[0]) if k not in i]
            gv = np.linalg.det(R[np.ix_(i, i)]) * np.linalg.det(R[np.ix_(rest, rest)]) / dR
            rows.append({"Variabel": UNS.get(t, t), "Df": len(i), "VIF setara": gv ** (1 / len(i))})
        return pd.DataFrame(rows)
    except Exception:
        return None


def bp_test(m):
    e = np.asarray(m.resid); u = e ** 2 / (np.sum(e ** 2) / len(e)); X = np.asarray(m.model.exog)
    beta = np.linalg.lstsq(X, u, rcond=None)[0]; fitted = X @ beta
    stat = np.sum((fitted - u.mean()) ** 2) / 2; df_ = X.shape[1] - 1
    return dict(stat=stat, df=df_, p=sps.chi2.sf(stat, df_))


def reg_df(m):
    infl = m.get_influence()
    return pd.DataFrame({"fit": np.asarray(m.fittedvalues), "act": np.asarray(m.model.endog), "res": np.asarray(m.resid),
                         "rs": infl.resid_studentized_internal, "cook": infl.cooks_distance[0], "idx": np.arange(1, int(m.nobs) + 1)})


def fit_table(m):
    e = np.asarray(m.resid)
    return pd.DataFrame({
        "Ukuran": ["Jumlah observasi (n)", "Derajat bebas residual", "R-kuadrat", "R-kuadrat terkoreksi (Adj. R2)", "Galat baku residual (sigma)", "RMSE", "MAE",
                   "Statistik F", "p-value (uji F simultan)", "AIC", "BIC"],
        "Nilai": [str(int(m.nobs)), str(int(m.df_resid)), fmt(m.rsquared, 4), fmt(m.rsquared_adj, 4), fmt(np.sqrt(m.mse_resid), 3), fmt(np.sqrt(np.mean(e ** 2)), 3), fmt(np.mean(np.abs(e)), 3),
                  "%s (df1 = %d, df2 = %d)" % (fmt(m.fvalue, 2), m.df_model, m.df_resid), fmt_p(m.f_pvalue), fmt(m.aic, 2), fmt(m.bic, 2)]})


def coef_table(m):
    ci = m.conf_int()
    return pd.DataFrame({"Koefisien": [coef_label(m, n) for n in m.params.index], "Estimasi": fmt(m.params.to_numpy(), 4), "Galat baku": fmt(m.bse.to_numpy(), 4),
                         "t": fmt(m.tvalues.to_numpy(), 3), "p-value": fmt_p(m.pvalues.to_numpy()), "CI 2.5%": fmt(ci[0].to_numpy(), 3), "CI 97.5%": fmt(ci[1].to_numpy(), 3),
                         "Sig.": stars(m.pvalues.to_numpy())})


def anova_table(m):
    """Padanan anova(m) di R: jumlah kuadrat sekuensial (Tipe I) menurut urutan suku pada model."""
    X = m.X.to_numpy(); y = m.y; rss_prev = np.sum((y - y.mean()) ** 2); end = 1; rows = []
    mse = m.mse_resid
    for t, (a, b) in m.term_slices.items():
        end = b; beta = np.linalg.lstsq(X[:, :end], y, rcond=None)[0]; rss = np.sum((y - X[:, :end] @ beta) ** 2)
        ss = rss_prev - rss; df_ = b - a; F = (ss / df_) / mse
        rows.append((UNS.get(t, t), df_, ss, ss / df_, F, sps.f.sf(F, df_, m.df_resid))); rss_prev = rss
    rows.append(("Residuals", int(m.df_resid), m.ssr, m.mse_resid, np.nan, np.nan))
    return pd.DataFrame({"Variabel": [r[0] for r in rows], "Df": [int(r[1]) for r in rows],
                         "Jumlah kuadrat": fmt([r[2] for r in rows], 1), "Kuadrat tengah": fmt([r[3] for r in rows], 1),
                         "F": ["" if np.isnan(r[4]) else fmt(r[4], 3) for r in rows], "p-value": ["" if np.isnan(r[5]) else fmt_p(r[5]) for r in rows]})


def plot_reg_avp(m, logy, resp):
    df = reg_df(m); r2 = m.rsquared; sc = f"ln({resp})" if logy else resp
    return (ggplot(df, aes("fit", "act")) + geom_point(color=pk["m"], alpha=.45, size=1.8) + geom_abline(slope=1, intercept=0, color=pk["d"], linetype="dashed", size=1)
            + labs(title="Nilai aktual vs prediksi", subtitle="Garis putus-putus = prediksi sempurna | R2 = %.3f" % r2, x="Prediksi " + sc, y="Aktual " + sc) + theme_pink())


def plot_reg_coef(m):
    ci = m.conf_int(); keep = [n for n in m.params.index if n != "Intercept"]
    df = pd.DataFrame({"lab": [coef_label(m, n) for n in keep], "est": m.params[keep].to_numpy(), "lo": ci.loc[keep, 0].to_numpy(), "hi": ci.loc[keep, 1].to_numpy(), "p": m.pvalues[keep].to_numpy()})
    df["sig"] = np.where(df["p"] < .05, "Signifikan (p < 0.05)", "Tidak signifikan"); df["lab"] = pd.Categorical(df["lab"], categories=list(reversed(df["lab"])))
    return (ggplot(df, aes("est", "lab", color="sig")) + geom_vline(xintercept=0, color=pk["mauve"], linetype="dashed")
            + geom_segment(aes(x="lo", xend="hi", y="lab", yend="lab"), size=.8) + geom_point(size=2.6)
            + scale_colour_manual(values={"Signifikan (p < 0.05)": pk["d"], "Tidak signifikan": "#B39DAB"}, name="")
            + labs(title="Estimasi koefisien dan selang kepercayaan 95%", subtitle="Selang yang memotong garis nol berarti tidak signifikan", x="Estimasi koefisien", y="")
            + theme_pink() + theme(legend_position="bottom"))


def _lo_line(x, y):
    return smooth_df(x, y)


def plot_reg_resid(m):
    df = reg_df(m); sm_ = _lo_line(df["fit"], df["res"])
    return (ggplot(df, aes("fit", "res")) + geom_point(color=pk["m"], alpha=.45, size=1.6) + geom_hline(yintercept=0, color=pk["mauve"], linetype="dashed")
            + geom_line(data=sm_, mapping=aes("x", "y"), color=pk["d"], size=.9, inherit_aes=False)
            + labs(title="Residual vs nilai prediksi", subtitle="Sebaran acak di sekitar nol = asumsi terpenuhi", x="Nilai prediksi", y="Residual") + theme_pink(11))


def plot_reg_qq(m):
    df = reg_df(m)
    return (ggplot(df, aes(sample="rs")) + stat_qq(color=pk["m"], alpha=.6) + stat_qq_line(color=pk["d"])
            + labs(title="Normal probability plot (QQ-plot residual terstandar)", subtitle="Titik mengikuti garis = residual mendekati normal", x="Kuantil teoretis", y="Kuantil sampel") + theme_pink(11))


def plot_reg_scale(m):
    df = reg_df(m); df["sq"] = np.sqrt(np.abs(df["rs"])); sm_ = _lo_line(df["fit"], df["sq"])
    return (ggplot(df, aes("fit", "sq")) + geom_point(color=pk["m"], alpha=.45, size=1.6) + geom_line(data=sm_, mapping=aes("x", "y"), color=pk["d"], size=.9, inherit_aes=False)
            + labs(title="Scale-Location", subtitle="Garis datar = ragam residual konstan (homoskedastis)", x="Nilai prediksi", y="Akar |residual terstandar|") + theme_pink(11))


def plot_reg_cook(m):
    df = reg_df(m); thr = 4 / len(df); top = df.sort_values("cook", ascending=False).head(3); df["z"] = 0.0
    return (ggplot(df, aes("idx", "cook")) + geom_segment(aes(xend="idx", yend="z"), color=pk["l"]) + geom_hline(yintercept=thr, color=pk["d"], linetype="dashed")
            + geom_text(data=top, mapping=aes(label="idx"), va="bottom", color=pk["d"], size=8)
            + labs(title="Cook's distance", subtitle="Garis putus-putus = 4/n = %.4f" % thr, x="Indeks observasi", y="Cook's distance") + theme_pink(11))


# ---- Visualisasi regresi tambahan ----
def reg_ord_res(m, d):
    o = np.argsort(d["Date"].to_numpy(), kind="stable")
    return dict(res=np.asarray(m.resid)[o], date=d["Date"].to_numpy()[o])


def plot_reg_scatter(d, resp, preds, logy):
    nums = [p for p in preds if p in num_all]
    validate(len(nums) >= 1, "Pilih minimal satu prediktor numerik untuk scatterplot.")
    yv = np.log(d[resp].to_numpy(float)) if logy else d[resp].to_numpy(float)
    long = pd.concat([pd.DataFrame({"Prediktor": v, "x": d[v].to_numpy(float), "y": yv}) for v in nums]); long["Prediktor"] = pd.Categorical(long["Prediktor"], categories=nums)
    return (ggplot(long, aes("x", "y")) + geom_point(color=pk["m"], alpha=.35, size=1.3) + geom_smooth(method="lm", formula="y ~ x", se=True, color=pk["d"], fill=pk["xl"], size=.9)
            + facet_wrap("Prediktor", scales="free_x")
            + labs(title="Scatterplot respons terhadap prediktor numerik", subtitle="Garis = regresi linear sederhana dengan pita kepercayaan 95%", x="Nilai prediktor", y=f"ln({resp})" if logy else resp) + theme_pink(11))


def plot_reg_cor(d, resp, preds):
    vars_ = list(dict.fromkeys([resp] + [p for p in preds if p in num_all]))
    validate(len(vars_) >= 2, "Pilih minimal satu prediktor numerik untuk matriks korelasi.")
    return plot_mv_cor(d, vars_, "pearson") + labs(title="Correlation matrix (respons dan prediktor numerik)")


def plot_reg_hist(m):
    df = reg_df(m); mu, sg = df["res"].mean(), sd1(df["res"])
    return (ggplot(df, aes("res")) + geom_histogram(aes(y=after_stat("density")), bins=30, fill=pk["xl"], color="white") + geom_density(color=pk["d"], size=1)
            + stat_function(fun=lambda z: sps.norm.pdf(z, mu, sg), color=pk["slate"], linetype="dashed", size=.9)
            + labs(title="Histogram residual", subtitle="Garis tebal = density residual; putus-putus = kurva normal teoretis", x="Residual", y="Densitas") + theme_pink(11))


def plot_reg_ts(m, d):
    o = reg_ord_res(m, d); n = len(o["res"]); df = pd.DataFrame({"t": np.arange(1, n + 1), "res": o["res"]})
    br = sorted(set(int(round(v)) for v in np.linspace(1, n, 6))); sm_ = _lo_line(df["t"], df["res"])
    return (ggplot(df, aes("t", "res")) + geom_line(color=pk["l"], size=.4) + geom_point(color=pk["m"], size=.8, alpha=.5) + geom_hline(yintercept=0, color=pk["mauve"], linetype="dashed")
            + geom_line(data=sm_, mapping=aes("x", "y"), color=pk["d"], size=.9, inherit_aes=False)
            + scale_x_continuous(breaks=br, labels=[pd.Timestamp(o["date"][b - 1]).strftime("%d-%m-%Y") for b in br])
            + labs(title="Time series plot residual", subtitle="Observasi diurutkan menurut tanggal transaksi; garis tebal = tren (loess)", x="Tanggal transaksi (urutan observasi)", y="Residual") + theme_pink(11))


def plot_reg_acf(m, d):
    o = reg_ord_res(m, d); n = len(o["res"]); a = sm_acf(o["res"], nlags=min(30, n - 1), fft=False)
    df = pd.DataFrame({"lag": np.arange(1, len(a)), "acf": a[1:]}); df["z"] = 0.0; ci = sps.norm.ppf(.975) / np.sqrt(n)
    return (ggplot(df, aes("lag", "acf")) + geom_hline(yintercept=0, color=pk["mauve"]) + geom_segment(aes(xend="lag", yend="z"), color=pk["m"], size=1) + geom_point(color=pk["d"], size=1.8)
            + geom_hline(yintercept=[-ci, ci], color=pk["slate"], linetype="dashed")
            + labs(title="ACF plot residual", subtitle="Garis putus-putus = batas 95%; lag 0 tidak ditampilkan", x="Lag", y="Autokorelasi") + theme_pink(11))


def plot_reg_obsres(m):
    df = reg_df(m); sm_ = _lo_line(df["act"], df["res"])
    return (ggplot(df, aes("act", "res")) + geom_point(color=pk["m"], alpha=.45, size=1.6) + geom_hline(yintercept=0, color=pk["mauve"], linetype="dashed")
            + geom_line(data=sm_, mapping=aes("x", "y"), color=pk["d"], size=.9, inherit_aes=False)
            + labs(title="Observed vs residual", subtitle="Pola menyebar/mengerucut menandakan ragam residual tidak konstan", x="Nilai observasi (aktual)", y="Residual") + theme_pink(11))


def _shapiro_p(e):
    e = np.asarray(e, float)
    if len(e) > 5000:
        e = np.random.choice(e, 5000, replace=False)
    return sps.shapiro(e)


def interp_reg(m, resp, logy):
    f = m.fvalue; pF = m.f_pvalue; r2 = m.rsquared; ar2 = m.rsquared_adj
    kat = "lemah" if ar2 < .3 else "sedang" if ar2 < .6 else "kuat"
    out = ["**Uji simultan (uji F):** F = %s, p %s. %s" % (fmt(f, 2), "< 0.001" if pF < .001 else "= " + fmt_p(pF),
                                                         "Minimal satu prediktor berpengaruh signifikan terhadap respons, sehingga model layak digunakan." if pF < .05 else "Tidak ada prediktor yang berpengaruh signifikan secara simultan."),
           "**Kemampuan model:** R2 = %s dan Adj. R2 = %s, artinya sekitar **%s%%** keragaman %s dijelaskan oleh prediktor dalam model (kategori **%s**); sisanya dijelaskan faktor lain di luar model." % (
               fmt(r2, 3), fmt(ar2, 3), fmt(100 * ar2, 1), f"ln({resp})" if logy else resp, kat)]
    cn = [n for n in m.params.index if n != "Intercept"]; sg = [n for n in cn if m.pvalues[n] < .05]
    if len(sg) == 0:
        out.append("**Uji parsial (uji t):** tidak ada koefisien yang signifikan pada taraf 5%.")
    else:
        out.append("**Uji parsial (uji t):** prediktor signifikan pada taraf 5%%: **%s**." % ", ".join(coef_label(m, n) for n in sg))
        for nm in sg:
            b = m.params[nm]; tl = term_of(m, nm); isnum = S(tl) not in m.xlevels
            eff = "%s%%" % fmt(100 * (math.exp(abs(b)) - 1), 1) if logy else fmt(abs(b), 3)
            verb = "naik" if b > 0 else "turun"
            if isnum:
                out.append("Setiap kenaikan 1 satuan **%s**, rata-rata **%s** %s sebesar **%s**, dengan variabel lain dianggap tetap." % (tl, resp, verb, eff))
            else:
                out.append("Dibanding kategori acuan **%s**, kategori **%s** memiliki rata-rata **%s** yang %s sebesar **%s**, dengan variabel lain dianggap tetap." % (
                    m.xlevels[S(tl)][0], re.search(r"\[T\.(.*)\]$", nm).group(1), resp, "lebih tinggi" if b > 0 else "lebih rendah", eff))
    tl_all = list(dict.fromkeys(term_of(m, n) for n in cn)); tl_sig = list(dict.fromkeys(term_of(m, n) for n in sg))
    ns = [t for t in tl_all if t not in tl_sig]
    if len(ns) > 0:
        out.append("Variabel **tidak signifikan** (tidak ada bukti pengaruh setelah variabel lain dikendalikan): %s." % ", ".join(ns))
    sh = _shapiro_p(m.resid).pvalue; bp = bp_test(m)["p"]; vt = vif_table(m)
    out.append("**Asumsi residual:** normalitas (Shapiro-Wilk) p %s, sehingga %s; kesamaan ragam (Breusch-Pagan) p %s, sehingga %s." % (
        pq(sh), "residual **tidak normal**" if sh < .05 else "residual normal", pq(bp), "ragam residual **tidak konstan**" if bp < .05 else "ragam konstan"))
    if vt is not None:
        mx = vt["VIF setara"].max()
        out.append("**Multikolinearitas:** VIF setara maksimum = %s (%s)." % (fmt(mx, 2), "aman, di bawah 5" if mx < 5 else "perlu diperhatikan, di atas 5"))
    return out


# =============================================================================
# KESIMPULAN OTOMATIS
# =============================================================================
def auto_conclusion(d, m, resp, logy):
    cn = [n for n in m.params.index if n != "Intercept"]; sg = [n for n in cn if m.pvalues[n] < .05]
    nm = ["Quantity", "Total", "Rating", "Hour"]; cm = d[nm].corr().to_numpy().copy(); np.fill_diagonal(cm, 0)
    j, i = np.argwhere(np.abs(cm).T == np.abs(cm).max())[0]       # urutan kolom-utama seperti R
    return ["Data yang dianalisis berisi **%d transaksi**; nilai transaksi (Total) berkisar %s sampai %s dengan median **%s** dan rata-rata %s (%s)." % (
                len(d), fmt(d["Total"].min()), fmt(d["Total"].max()), fmt(d["Total"].median()), fmt(d["Total"].mean()),
                "distribusi miring ke kanan" if skew_fun(d["Total"]) > .5 else "distribusi relatif simetris"),
            "Pasangan variabel numerik dengan korelasi terkuat adalah **%s** dan **%s** (r = %s)." % (nm[i], nm[j], fmt(cm[i, j], 3)),
            "Model regresi (respons **%s**%s) memiliki Adj. R2 = **%s**; prediktor signifikan pada taraf 5%%: **%s**." % (
                resp, ", transformasi ln" if logy else "", fmt(m.rsquared_adj, 3), "tidak ada" if len(sg) == 0 else ", ".join(coef_label(m, n) for n in sg))]


# EDIT: sesuaikan rekomendasi berikut dengan hasil analisis akhir Anda.
REKOMENDASI = [
    "Jumlah unit per transaksi (Quantity) merupakan faktor utama nilai transaksi; promosi berbasis jumlah pembelian (paket bundling, potongan harga pembelian banyak) berpotensi menaikkan nilai transaksi.",
    "Cabang, tipe pelanggan, jenis kelamin, lini produk, dan metode pembayaran tidak menunjukkan pengaruh nyata terhadap nilai transaksi setelah Quantity dikendalikan, sehingga strategi dari sisi nilai transaksi tidak perlu dibedakan berdasarkan segmen tersebut.",
    "Rating tidak berkaitan dengan nilai transaksi maupun jumlah pembelian; peningkatan kepuasan pelanggan perlu ditelusuri melalui data tambahan seperti kualitas layanan dan waktu tunggu.",
    "Penelitian lanjutan disarankan menambahkan variabel harga satuan, periode data yang lebih panjang, serta model yang menangani ragam tidak konstan (transformasi atau regresi robust).",
]
KETERBATASAN = [
    "Data bersifat observasional dan hanya mencakup periode singkat (Januari-Maret) sehingga hasil tidak dapat digeneralisasi ke seluruh waktu.",
    "Regresi menunjukkan hubungan (asosiasi), bukan sebab-akibat.",
    "Variabel penting seperti harga satuan tidak tersedia pada data, sehingga sebagian keragaman nilai transaksi tidak dapat dijelaskan model.",
    "Residual cenderung tidak normal dan ragamnya tidak konstan (bertambah seiring Quantity); inferensi (uji t/F, selang kepercayaan) hanya pendekatan. Transformasi log atau regresi robust dapat dipertimbangkan.",
    "Model linear mengasumsikan efek aditif dan linear; interaksi antar variabel tidak dimodelkan.",
]



# =============================================================================
# TAMPILAN STREAMLIT
# =============================================================================
import tempfile
from contextlib import contextmanager

CSS = """<style>
@import url('https://fonts.googleapis.com/css2?family=Source+Sans+3:wght@400;600;700&family=Source+Serif+4:wght@600;700&display=swap');
html, body, [class*="css"] {font-family:'Source Sans 3','Segoe UI',Arial,sans-serif;}
h1,h2,h3,h4{font-family:'Source Serif 4',Georgia,serif!important;color:#880E4F!important;}
.app-header{border-top:6px solid #C2185B;border-bottom:1px solid #EBD3DC;background:#FFF5F8;padding:18px 24px 12px;margin-bottom:18px;}
.app-header h1{margin:0 0 6px;font-size:26px;font-weight:700;padding:0;}
.app-header .sub{color:#6D4C5E;font-size:14px;margin:0;}
div[data-testid="stRadio"] label p{font-weight:600;color:#6D4C5E;}
.stButton>button{background:#C2185B;color:#fff;border:1px solid #C2185B;font-weight:600;border-radius:3px;}
.stButton>button:hover{background:#880E4F;border-color:#880E4F;color:#fff;}
.nfilter{background:#fff;border:1px solid #EBD3DC;border-radius:3px;padding:8px 10px;font-size:13px;color:#6D4C5E;}
.nfilter b{color:#880E4F;font-size:16px;}
</style>"""
st.markdown(CSS, unsafe_allow_html=True)


@contextmanager
def guard():
    try:
        yield
    except SafeException as e:
        st.warning(str(e))


def show(p, w=9, h=4.6):
    if hasattr(p, "savefig"):
        st.pyplot(p); plt.close(p)
    else:
        fig = (p + theme(figure_size=(w, h))).draw(); st.pyplot(fig); plt.close(fig)


def tbl(df):
    st.dataframe(df, hide_index=True)


def sb(cont, label, opts, default, key):
    opts = list(opts)
    return cont.selectbox(label, opts, index=opts.index(default) if default in opts else 0, key=key)


def lv(d, v):
    return [str(c) for c in d[v].cat.categories]


@st.cache_data(show_spinner=False)
def load_default():
    for p in [DATA_FILE, os.path.join(os.path.dirname(os.path.abspath(__file__)), DATA_FILE)]:
        if os.path.exists(p):
            return read_sales(p)
    return None


@st.cache_data(show_spinner=False)
def load_upload(name, content):
    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(name)[1]) as tf:
        tf.write(content); path = tf.name
    try:
        return read_sales(path)
    finally:
        os.remove(path)


def upper_pairs(m, vars_):
    p = len(vars_); out = [(vars_[i], vars_[j], m[i, j]) for j in range(p) for i in range(j)]
    return sorted(out, key=lambda t: -abs(t[2]))


st.markdown(f'<div class="app-header"><h1>{APP_TITLE}</h1><p class="sub">Ujian Tengah Semester Eksplorasi dan Visualisasi Data Multivariat &nbsp;|&nbsp; {NAMA} &nbsp;|&nbsp; {NIM}</p></div>', unsafe_allow_html=True)

ss = st.session_state
# ------------------------------------------------------------------ SIDEBAR
with st.sidebar:
    st.subheader("Sumber data")
    up = st.file_uploader("Excel/CSV (opsional)", type=["xlsx", "xls", "csv"])
    d_raw = None
    try:
        d_raw = load_upload(up.name, up.getvalue()) if up else load_default()
    except Exception:
        d_raw = None
    if d_raw is None:
        st.error("Data tidak ditemukan. Letakkan Data_Penjualan.xlsx di folder aplikasi (repositori) atau unggah file melalui panel ini (kolom wajib: Branch, Customer type, Gender, Product line, Quantity, Total, Date, Time, Payment, Rating).")
        st.stop()
    sig = (up.name if up else "default") + str(ss.get("rst", 0))
    st.subheader("Filter data")
    f_branch = st.multiselect("Branch", lv(d_raw, "Branch"), default=lv(d_raw, "Branch"), key="fb" + sig)
    f_cust = st.multiselect("Customer type", lv(d_raw, "Customer type"), default=lv(d_raw, "Customer type"), key="fc" + sig)
    f_gender = st.multiselect("Gender", lv(d_raw, "Gender"), default=lv(d_raw, "Gender"), key="fg" + sig)
    f_pay = st.multiselect("Payment", lv(d_raw, "Payment"), default=lv(d_raw, "Payment"), key="fp" + sig)
    f_prod = st.multiselect("Product line", lv(d_raw, "Product line"), default=lv(d_raw, "Product line"), key="fl" + sig)
    dmin, dmax = d_raw["Date"].min().date(), d_raw["Date"].max().date()
    f_date = st.date_input("Rentang tanggal", value=(dmin, dmax), min_value=dmin, max_value=dmax, format="DD-MM-YYYY", key="fd" + sig)
    if st.button("Atur ulang filter"):
        ss["rst"] = ss.get("rst", 0) + 1; st.rerun()

d = d_raw
d = d[d["Branch"].isin(f_branch) & d["Customer type"].isin(f_cust) & d["Gender"].isin(f_gender) & d["Payment"].isin(f_pay) & d["Product line"].isin(f_prod)]
if isinstance(f_date, (tuple, list)) and len(f_date) == 2:
    d = d[(d["Date"] >= pd.Timestamp(f_date[0])) & (d["Date"] <= pd.Timestamp(f_date[1]))]
n0, n1 = len(d_raw), len(d)
with st.sidebar:
    st.markdown(f'<div class="nfilter">Transaksi terpilih:<br><b>{n1:,}</b> dari {n0:,} ({fmt(100 * n1 / n0, 1)}%)</div>', unsafe_allow_html=True)
if n1 < 30:
    st.warning("Data hasil filter terlalu sedikit (kurang dari 30 baris). Longgarkan filter pada panel kiri.")
    st.stop()
d = d.copy()
for c_ in cat_all:
    d[c_] = d[c_].cat.remove_unused_categories()
d = d.reset_index(drop=True)
data_raw_, data_use_ = d_raw, d

SECS = ["1. Masalah dan Data", "2. Univariat", "3. Bivariat", "4. Multivariat", "5. Regresi Linear Berganda", "6. Kesimpulan"]
sec = st.radio("Bagian", SECS, horizontal=True, label_visibility="collapsed", key="sec")
st.divider()


def card(title):
    c = st.container(border=True); c.markdown(f"#### {title}"); return c


# ------------------------------------------------------------------ 1. DATA
if sec == SECS[0]:
    with card("Latar belakang"): st.write(LATAR_BELAKANG)
    with card("Tujuan analisis"):
        for i, t in enumerate(TUJUAN, 1): st.markdown(f"{i}. {t}")
    with card("Sumber dan karakteristik data"):
        st.write(SUMBER_DATA)
        bullets(["Jumlah observasi: **%s** transaksi; jumlah variabel: **10** variabel asli (ditambah 3 variabel turunan: Hour, Bulan, Hari)." % f"{len(d_raw):,}",
                 "Periode data: **%s** sampai **%s**." % (d_raw["Date"].min().date(), d_raw["Date"].max().date()),
                 "Variabel kategorik: Branch, Customer type, Gender, Product line, Payment. Variabel numerik: Quantity, Total, Rating, serta Hour (turunan)."])
    with card("Definisi variabel"): tbl(kamus)
    with card("Pengolahan data"):
        q1, q3 = np.quantile(d_raw["Total"], [.25, .75]); out = int(((d_raw["Total"] < q1 - 1.5 * (q3 - q1)) | (d_raw["Total"] > q3 + 1.5 * (q3 - q1))).sum())
        bullets(["**Data hilang:** %d sel kosong pada seluruh kolom analisis." % int(d_raw.isna().sum().sum()),
                 "**Data ganda:** %d baris duplikat." % int(d_raw.duplicated().sum()),
                 "**Pencilan:** terdapat %d pencilan pada Total (aturan 1.5 x IQR). Pencilan dipertahankan karena merupakan transaksi bernilai besar yang valid." % out,
                 "**Pengolahan:** variabel Hour, Bulan, dan Hari dibentuk dari Time dan Date; variabel kategorik diubah menjadi faktor.",
                 "**Transformasi:** tidak dilakukan pada eksplorasi; pada regresi tersedia pilihan transformasi ln(Y) untuk mengatasi ragam yang tidak konstan."])
    with card("Cuplikan data"):
        x = d_raw.head(10).copy(); x["Date"] = x["Date"].dt.strftime("%Y-%m-%d"); tbl(x)

# ------------------------------------------------------------------ 2. UNIVARIAT
elif sec == SECS[1]:
    c = st.columns([1, 1, 2]); v = sb(c[0], "Variabel", cat_all + num_all + ["Date"], "Total", "uni_var")
    t = c[1].selectbox("Jenis grafik", uni_choices(v), key="uni_type_" + v)
    a, b = st.columns([7, 5])
    with a, guard():
        with card("Grafik"):
            if t == "Stem-and-leaf":
                st.code("\n".join(["Stem-and-leaf plot: " + v, "", stem_text(d[v].to_numpy(float))]), language=None)
            else:
                show(plot_uni(d, v, t), 7, 4.6)
    with b:
        with card("Statistik deskriptif"): tbl(uni_table(d, v))
    with card("Interpretasi"): bullets(interp_uni(d, v))

# ------------------------------------------------------------------ 3. BIVARIAT
elif sec == SECS[2]:
    c = st.columns(3); x = sb(c[0], "Variabel X", cat_all + num_all, "Quantity", "bi_x"); y = sb(c[1], "Variabel Y", cat_all + num_all, "Total", "bi_y")
    t = c[2].selectbox("Jenis grafik", bi_choices(x, y), key=f"bi_type_{x}_{y}")
    st.caption("Jenis grafik menyesuaikan tipe pasangan variabel: numerik-numerik, kategorik-numerik, atau kategorik-kategorik.")
    with guard():
        validate(x != y, "Pilih dua variabel yang berbeda.")
        res = bi_text(d, x, y)
        a, b = st.columns([7, 5])
        with a, card("Grafik"): show(plot_bi(d, x, y, t), 7, 4.6)
        with b, card("Ringkasan dan uji"): tbl(res["table"])
        with card("Interpretasi"): bullets(res["lines"])

# ------------------------------------------------------------------ 4. MULTIVARIAT
elif sec == SECS[3]:
    MV = ["Scatter multivariat", "Heatmap rata-rata", "Matriks korelasi", "Tren waktu", "PCA", "Scatterplot matrix", "Pencilan Mahalanobis", "Q-Q plot Chi-Square"]
    sub = st.radio("Analisis", MV, horizontal=True, key="mv_sub"); st.write("")
    with guard():
        if sub == MV[0]:
            c = st.columns(4)
            x = sb(c[0], "Sumbu X", num_all, "Quantity", "mv_x"); y = sb(c[1], "Sumbu Y", num_all, "Total", "mv_y")
            col = sb(c[2], "Warna", [NONE] + cat_all + num_all, "Customer type", "mv_col"); size = sb(c[3], "Ukuran titik", [NONE] + num_all, "Rating", "mv_size")
            c = st.columns(4); facet = sb(c[0], "Panel (facet)", [NONE] + cat_all, NONE, "mv_facet"); lines = c[1].checkbox("Tampilkan garis regresi per kelompok", True, key="mv_lines")
            validate(x != y, "Pilih sumbu X dan Y yang berbeda.")
            with card("Grafik"): show(plot_mv_scatter(d, x, y, col, size, facet, lines), 9, 4.8)
            with card("Interpretasi"):
                r = d[x].corr(d[y]); ln = ["Secara keseluruhan korelasi **%s** dan **%s** adalah **%s** (%s)." % (x, y, fmt(r, 3), strength(r))]
                if col != NONE and col not in num_all:
                    gr = {str(k): g_[x].corr(g_[y]) for k, g_ in d.groupby(col, observed=True)}
                    ln.append("Korelasi per **%s**: %s. %s" % (col, "; ".join(f"{k} = {fmt(v_, 2)}" for k, v_ in gr.items()),
                              "Pola hubungan hampir sama di semua kelompok sehingga variabel warna tidak mengubah hubungan X dan Y." if (max(gr.values()) - min(gr.values())) < .1 else "Kekuatan hubungan berbeda antar kelompok, menandakan kemungkinan interaksi."))
                if size != NONE:
                    ln.append("Ukuran titik menunjukkan **%s**; jika titik besar dan kecil tersebar acak, variabel tersebut tidak berkaitan dengan posisi titik pada X dan Y." % size)
                bullets(ln)
        elif sub == MV[1]:
            c = st.columns(4)
            a_ = sb(c[0], "Kategorik 1 (sumbu X)", cat_all, "Branch", "mv_ha"); b_ = sb(c[1], "Kategorik 2 (sumbu Y)", cat_all, "Product line", "mv_hb")
            n_ = sb(c[2], "Variabel numerik", num_all, "Total", "mv_hn"); f_ = sb(c[3], "Ukuran", ["Rata-rata", "Median", "Jumlah", "Banyak transaksi"], "Rata-rata", "mv_hf")
            validate(a_ != b_, "Pilih dua variabel kategorik yang berbeda.")
            with card("Grafik"): show(plot_mv_heat(d, a_, b_, n_, f_), 9, 4.6)
            with card("Interpretasi"):
                r = mv_heat_data(d, a_, b_, n_, f_); hi = r.loc[r["val"].idxmax()]; lo = r.loc[r["val"].idxmin()]
                bullets(["Nilai tertinggi pada kombinasi **%s - %s** (%s) dan terendah pada **%s - %s** (%s)." % (hi["A"], hi["B"], fmt(hi["val"], 1), lo["A"], lo["B"], fmt(lo["val"], 1)),
                         "Selisih tertinggi-terendah sebesar **%s** (%s%% dari nilai tertinggi)." % (fmt(hi["val"] - lo["val"], 1), fmt(100 * (hi["val"] - lo["val"]) / hi["val"], 1)),
                         "Warna lebih pekat menandakan nilai lebih besar. Perbedaan antar sel perlu dikonfirmasi dengan uji statistik sebelum disimpulkan sebagai perbedaan nyata."])
        elif sub == MV[2]:
            c = st.columns([2, 1]); vs = c[0].multiselect("Variabel numerik", num_all, default=num_all, key="mv_cv"); me = sb(c[1], "Metode", ["pearson", "spearman"], "pearson", "mv_cm")
            validate(len(vs) >= 2, "Pilih minimal dua variabel.")
            with card("Grafik"): show(plot_mv_cor(d, vs, me), 7, 4.6)
            with card("Interpretasi"):
                pr = upper_pairs(d[vs].corr(method=me).to_numpy(), vs)
                bullets(["Korelasi terkuat: **%s - %s** (r = %s; %s)." % (pr[0][0], pr[0][1], fmt(pr[0][2], 3), strength(pr[0][2])),
                         "Korelasi terlemah: **%s - %s** (r = %s)." % (pr[-1][0], pr[-1][1], fmt(pr[-1][2], 3)),
                         "Warna merah muda tua = korelasi positif, biru keabuan = negatif, putih = mendekati nol."])
        elif sub == MV[3]:
            c = st.columns(3)
            tm = sb(c[0], "Metrik", ["Jumlah transaksi", "Total penjualan (jumlah)", "Rata-rata Total", "Rata-rata Rating", "Total Quantity"], "Total penjualan (jumlah)", "mv_tm")
            tu = sb(c[1], "Satuan waktu", ["Harian", "Mingguan", "Bulanan"], "Mingguan", "mv_tu"); tg = sb(c[2], "Kelompok", [NONE] + cat_base, "Branch", "mv_tg")
            with card("Grafik"): show(plot_mv_trend(d, tm, tu, tg), 9, 4.6)
            with card("Interpretasi"):
                td = trend_data(d, tm, tu, tg); hi = td.loc[td["val"].idxmax()]; lo = td.loc[td["val"].idxmin()]
                gtxt = f" (kelompok {hi[tg]})" if tg != NONE else ""
                bullets(["Nilai **%s** tertinggi pada periode **%s**%s sebesar %s; terendah pada **%s** sebesar %s." % (tm, hi["Periode"].date(), gtxt, fmt(hi["val"], 1), lo["Periode"].date(), fmt(lo["val"], 1)),
                         "Bandingkan garis antar kelompok untuk melihat apakah pola naik-turun sama (pola serupa menandakan faktor waktu yang umum, bukan khas satu kelompok).",
                         "Minggu pertama dan terakhir dapat berisi lebih sedikit hari sehingga nilainya tampak lebih rendah." if tu == "Mingguan" else "Perhatikan bahwa data hanya mencakup beberapa bulan, sehingga pola musiman tidak dapat disimpulkan."])
        elif sub == MV[4]:
            c = st.columns([2, 1]); vs = c[0].multiselect("Variabel numerik (minimal 3)", num_all, default=num_all, key="mv_pv"); pc = sb(c[1], "Warna titik", [NONE] + cat_base, "Customer type", "mv_pc")
            validate(len(vs) >= 3, "Pilih minimal tiga variabel numerik.")
            a, b = st.columns([7, 5]); pr = pca_fit(d, vs); ve = 100 * pr["sdev"] ** 2 / np.sum(pr["sdev"] ** 2)
            with a, card("Grafik"): show(plot_mv_pca(d, vs, pc), 7, 4.8)
            with b, card("Keragaman dan loading"):
                ld = pr["rotation"][["PC1", "PC2"]].round(3)
                t_ = pd.DataFrame({"Variabel": list(ld.index), "PC1": ld["PC1"].to_numpy(), "PC2": ld["PC2"].to_numpy()})
                tbl(pd.concat([t_, pd.DataFrame({"Variabel": ["Keragaman dijelaskan (%)"], "PC1": [round(ve[0], 1)], "PC2": [round(ve[1], 1)]})], ignore_index=True))
            with card("Interpretasi"):
                l1 = pr["rotation"]["PC1"]; k = l1.abs().idxmax()
                bullets(["PC1 menjelaskan **%s%%** dan PC2 **%s%%** keragaman; bersama-sama **%s%%**." % (fmt(ve[0], 1), fmt(ve[1], 1), fmt(ve[0] + ve[1], 1)),
                         "Variabel dengan kontribusi terbesar pada PC1: **%s** (loading %s)." % (k, fmt(l1[k], 2)),
                         "Panah yang searah menandakan variabel berkorelasi positif; panah yang saling tegak lurus menandakan variabel hampir tidak berkorelasi. Jika keragaman terbagi merata ke banyak komponen, variabel-variabel tersebut membawa informasi yang berbeda-beda."])
        elif sub == MV[5]:
            c = st.columns([2, 1]); vs = c[0].multiselect("Variabel numerik (minimal 2)", num_all, default=num_all, key="mv_mvars"); mc = sb(c[1], "Warna titik", [NONE] + cat_base, NONE, "mv_mcol")
            validate(len(vs) >= 2, "Pilih minimal dua variabel numerik.")
            with card("Grafik"): show(plot_mv_pairs(d, vs, mc))
            with card("Interpretasi"):
                pr = upper_pairs(d[vs].corr().to_numpy(), vs)
                bullets(["Diagonal: histogram tiap variabel; segitiga bawah: scatterplot dengan kurva lowess; segitiga atas: koefisien korelasi Pearson (semakin besar angka, semakin kuat hubungan).",
                         "Pasangan dengan hubungan terkuat: **%s - %s** (r = %s; %s). Pasangan terlemah: **%s - %s** (r = %s)." % (pr[0][0], pr[0][1], fmt(pr[0][2], 3), strength(pr[0][2]), pr[-1][0], pr[-1][1], fmt(pr[-1][2], 3)),
                         "Kurva lowess yang mendekati garis lurus menandakan hubungan linear; kurva melengkung menandakan hubungan nonlinear. Titik yang jauh dari kumpulan utama adalah kandidat pencilan."])
        elif sub == MV[6]:
            vs = st.multiselect("Variabel numerik (minimal 2)", num_all, default=num_all, key="mv_dvars")
            validate(len(vs) >= 2, "Pilih minimal dua variabel numerik."); f = mahal_fit(d, vs)
            a, b = st.columns(2)
            with a, card("Jarak Mahalanobis"): show(plot_mv_mahal_idx(f), 6, 4.2)
            with b, card("Ruang dua variabel"): show(plot_mv_mahal_sc(f, vs), 6, 4.2)
            with card("Sepuluh observasi dengan jarak terbesar"): tbl(mahal_table(d, f, vs))
            with card("Interpretasi"):
                n = len(f["df"]); k = int(f["df"]["out"].sum()); top = f["df"].loc[f["df"]["D2"].idxmax()]
                z = {v_: (d[v_].to_numpy()[int(top["idx"]) - 1] - d[v_].mean()) / sd1(d[v_]) for v_ in vs}; vmax = max(z, key=lambda v_: abs(z[v_]))
                bullets(["Batas pencilan multivariat: Chi-Square 97.5%% dengan df = %d, yaitu **%s**. Observasi dengan D2 di atas batas dianggap pencilan." % (f["p"], fmt(f["cut"], 2)),
                         "Ditemukan **%d pencilan** (**%s%%** dari %d observasi). Pada data normal multivariat, proporsi yang diharapkan sekitar 2.5%%." % (k, fmt(100 * k / n, 1), n),
                         "Observasi paling ekstrem: baris **%d** (D2 = %s). Variabel yang paling menyimpang: **%s** (z = %s)." % (top["idx"], fmt(top["D2"], 2), vmax, fmt(z[vmax], 2)),
                         "Pencilan multivariat dapat berupa kombinasi nilai yang tidak lazim walau tiap variabel tampak wajar. Pencilan sebaiknya diperiksa dulu keabsahannya sebelum dihapus."])
        else:
            vs = st.multiselect("Variabel numerik (minimal 2)", num_all, default=num_all, key="mv_qvars")
            validate(len(vs) >= 2, "Pilih minimal dua variabel numerik."); f = mahal_fit(d, vs)
            with card("Grafik"): show(plot_mv_qqchi(f), 8, 4.6)
            with card("Interpretasi"):
                n = len(f["df"]); s_ = np.sort(f["df"]["D2"].to_numpy()); q = sps.chi2.ppf((np.arange(1, n + 1) - .5) / n, df=f["p"]); r = np.corrcoef(s_, q)[0, 1]
                tail3 = f["df"]["idx"].to_numpy()[np.argsort(f["df"]["D2"].to_numpy(), kind="stable")][-3:]
                bullets(["Korelasi antara D2 terurut dan kuantil Chi-Square = **%s**; semakin mendekati 1, semakin sesuai data dengan normal multivariat." % fmt(r, 3),
                         "Titik terjauh dari garis berada di ujung kanan (baris %s), yang menandakan observasi dengan jarak Mahalanobis besar." % ", ".join(str(i) for i in tail3),
                         "Titik yang melengkung ke atas di ujung kanan menandakan ekor distribusi lebih berat atau adanya pencilan; pola berundak dapat muncul karena variabel diskret (Quantity, Hour). Asumsi normalitas multivariat dianggap kurang terpenuhi bila penyimpangan jelas."])


# ------------------------------------------------------------------ 5. REGRESI / 6. KESIMPULAN
def make_model(dd):
    resp = ss.get("rg_y", "Total"); xs = ss.get("rg_x", ["Quantity", "Rating", "Hour"] + cat_base); logy = ss.get("rg_log", "Tanpa") == "ln(Y)"
    preds = [p for p in xs if p != resp]
    validate(len(preds) >= 2, "Pilih minimal dua variabel prediktor (selain variabel respons).")
    if logy:
        validate(bool((dd[resp] > 0).all()), "Transformasi ln hanya untuk respons bernilai positif.")
    return dict(m=build_model(dd, resp, preds, logy), resp=resp, logy=logy, preds=preds, d=dd)


def get_model(force=False):
    if force or "model" not in ss:
        try:
            ss["model"] = make_model(d); ss["model_err"] = None
        except SafeException as e:
            ss["model"] = None; ss["model_err"] = str(e)
    if ss.get("model") is None:
        st.warning(ss.get("model_err") or "Model belum dibangun."); return None
    return ss["model"]


if sec == SECS[4]:
    with card("Spesifikasi model"):
        c = st.columns([2, 5, 2, 2])
        c[0].selectbox("Variabel respons (Y)", ["Total", "Rating"], key="rg_y")
        c[1].multiselect("Variabel prediktor (X)", ["Quantity", "Rating", "Hour"] + cat_all, default=["Quantity", "Rating", "Hour"] + cat_base, key="rg_x")
        c[2].radio("Transformasi respons", ["Tanpa", "ln(Y)"], key="rg_log")
        c[3].write(""); c[3].write("")
        go = c[3].button("Bangun model")
        st.caption("Respons Total dipilih karena merupakan ukuran utama nilai transaksi. Pilih minimal dua prediktor. Variabel kategorik otomatis diubah menjadi dummy dengan kategori pertama sebagai acuan.")
    MD = get_model(force=go)
    RG = ["Hasil model", "Visualisasi hasil model", "Galeri visualisasi regresi", "Diagnostik asumsi", "Keterbatasan model"]
    sub = st.radio("Tampilan", RG, horizontal=True, key="rg_sub"); st.write("")
    if sub == RG[4]:
        with card("Keterbatasan model yang digunakan"): bullets(KETERBATASAN)
    elif MD is not None:
        m, resp, logy, dm = MD["m"], MD["resp"], MD["logy"], MD["d"]
        with guard():
            if sub == RG[0]:
                with card("Persamaan model"): st.code(make_equation(m, resp, logy), language=None)
                a, b = st.columns([5, 7])
                with a, card("Ukuran kecocokan model"): tbl(fit_table(m))
                with b, card("Uji signifikansi per variabel (ANOVA)"): tbl(anova_table(m))
                with card("Estimasi koefisien"):
                    tbl(coef_table(m)); st.caption("Signif.: \\*\\*\\* p<0.001, \\*\\* p<0.01, \\* p<0.05, . p<0.1")
                with card("Interpretasi hasil"): bullets(interp_reg(m, resp, logy))
            elif sub == RG[1]:
                a, b = st.columns(2)
                with a, card("Aktual vs prediksi"): show(plot_reg_avp(m, logy, resp), 6, 4.2)
                with b, card("Koefisien"): show(plot_reg_coef(m), 6, 4.2)
                with card("Cara membaca"):
                    bullets(["**Aktual vs prediksi:** semakin rapat titik di sekitar garis putus-putus, semakin baik prediksi model (R2 = %s). Titik yang menyebar lebar menunjukkan keragaman yang belum dijelaskan model." % fmt(m.rsquared, 3),
                             "**Koefisien dan selang kepercayaan:** titik = estimasi, garis = selang 95%. Selang yang tidak memotong garis nol (warna tua) menandakan prediktor signifikan pada taraf 5%. Skala koefisien bergantung pada satuan variabelnya, jadi bandingkan terutama ada-tidaknya pengaruh, bukan panjang garisnya."])
            elif sub == RG[2]:
                gal = [lambda: plot_reg_scatter(dm, resp, MD["preds"], logy), lambda: plot_reg_cor(dm, resp, MD["preds"]), lambda: plot_reg_coef(m), lambda: plot_reg_qq(m),
                       lambda: plot_reg_hist(m), lambda: plot_reg_ts(m, dm), lambda: plot_reg_acf(m, dm), lambda: plot_reg_obsres(m)]
                for i in range(0, 8, 2):
                    a, b = st.columns(2)
                    for cc, fn in zip((a, b), gal[i:i + 2]):
                        with cc, guard(), card(" "): show(fn(), 6, 3.8)
                with card("Cara membaca"):
                    o = reg_ord_res(m, dm); e = o["res"]; n = len(e)
                    a_ = sm_acf(e, nlags=min(30, n - 1), fft=False); ci = sps.norm.ppf(.975) / np.sqrt(n); nsig = int(np.sum(np.abs(a_[1:]) > ci)); nl = len(a_) - 1
                    dw = np.sum(np.diff(e) ** 2) / np.sum(e ** 2); sw = _shapiro_p(e).pvalue
                    bullets(["**Scatterplot:** menunjukkan bentuk hubungan respons dengan tiap prediktor numerik; garis yang jelas menanjak/menurun menandakan hubungan linear.",
                             "**Correlation matrix:** korelasi respons dengan prediktor menunjukkan kandidat prediktor kuat; korelasi tinggi antar prediktor mengindikasikan potensi multikolinearitas.",
                             "**Coefficient CI plot:** selang 95% yang tidak memotong garis nol berarti koefisien signifikan pada taraf 5%.",
                             "**Normal probability plot dan histogram residual:** titik mengikuti garis dan histogram mendekati kurva normal berarti residual normal. Hasil Shapiro-Wilk p %s, sehingga residual %s." % (pq(sw), "**tidak normal**" if sw < .05 else "normal"),
                             "**Time series plot dan ACF:** statistik Durbin-Watson = %s (sekitar 2 berarti tidak ada autokorelasi); %d dari %d lag berada di luar batas 95%%. %s" % (
                                 fmt(dw, 3), nsig, nl, "Tidak ada indikasi autokorelasi residual yang berarti." if (nsig <= 0.05 * nl + 1 and abs(dw - 2) < .4) else "Terdapat indikasi autokorelasi residual; periksa pola waktu."),
                             "**Observed vs residual:** residual idealnya tersebar acak di sekitar nol. Sebaran yang melebar pada nilai observasi besar menandakan ragam residual tidak konstan (heteroskedastis)."])
            else:
                a, b = st.columns(2)
                with a, card("Residual vs prediksi"): show(plot_reg_resid(m), 6, 3.8)
                with b, card("QQ-plot residual"): show(plot_reg_qq(m), 6, 3.8)
                a, b = st.columns(2)
                with a, card("Scale-Location"): show(plot_reg_scale(m), 6, 3.8)
                with b, card("Cook's distance"): show(plot_reg_cook(m), 6, 3.8)
                a, b = st.columns(2)
                with a, card("Uji asumsi"):
                    e = m.resid; sh = _shapiro_p(e); bp = bp_test(m); cd = reg_df(m)["cook"].to_numpy(); ninf = int(np.sum(cd > 4 / len(cd)))
                    tbl(pd.DataFrame({"Uji": ["Shapiro-Wilk (normalitas residual)", "Breusch-Pagan (kesamaan ragam)", "Observasi berpengaruh (Cook's D > 4/n)"],
                                      "Statistik": [fmt(sh.statistic, 4), fmt(bp["stat"], 3), str(ninf)], "p-value": [fmt_p(sh.pvalue), fmt_p(bp["p"]), "-"],
                                      "Keputusan": ["Tidak normal" if sh.pvalue < .05 else "Normal", "Ragam tidak konstan" if bp["p"] < .05 else "Ragam konstan", f"{ninf} observasi"]}))
                with b, card("Multikolinearitas (VIF)"):
                    v = vif_table(m); validate(v is not None, "VIF membutuhkan minimal dua prediktor.")
                    v["Keputusan"] = np.where(v["VIF setara"].round(3) < 5, "Aman", "Perlu perhatian"); v["VIF setara"] = fmt(v["VIF setara"].to_numpy(), 3); tbl(v)

elif sec == SECS[5]:
    MD = get_model()
    with card("Kesimpulan utama"):
        if MD is not None:
            bullets(auto_conclusion(d, MD["m"], MD["resp"], MD["logy"]))
    with card("Rekomendasi"): bullets(REKOMENDASI)
