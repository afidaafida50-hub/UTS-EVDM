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

from shiny import App, reactive, render, ui
from shiny.types import SafeException, SilentException

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


def md(x):
    return ui.HTML(re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", x))


def bullets(lines):
    return ui.tags.ul(*[ui.tags.li(md(l)) for l in lines])


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


def inp(f, default=None):
    """Baca input yang mungkin belum terbentuk (UI dinamis)."""
    try:
        return f()
    except SilentException:
        return default


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
.sidebar{background:var(--p1)!important;border:1px solid var(--ln);border-radius:4px;box-shadow:none;}
.sidebar h4{font-size:15px;border-bottom:1px solid var(--ln);padding-bottom:6px;margin-top:0;}
.box-x{background:#fff;border:1px solid var(--ln);border-left:4px solid var(--p7);border-radius:3px;padding:14px 18px;margin-bottom:16px;}
.box-x h4{margin-top:0;font-size:16px;font-weight:700;}
.box-x p,.box-x li{line-height:1.6;}
.nav-tabs{border-bottom:2px solid var(--p3);}
.nav-tabs .nav-link{color:#6D4C5E;font-weight:600;border:none;border-radius:0;margin-right:2px;}
.nav-tabs .nav-link:hover{background:var(--p1);border:none;color:var(--p9);}
.nav-tabs .nav-link.active{color:var(--p9);border:none;border-bottom:3px solid var(--p7);background:#fff;}
.nav-pills .nav-link{color:#6D4C5E;font-weight:600;}
.nav-pills .nav-link.active{background:var(--p7);color:#fff;}
.btn-pink{background:var(--p7);border:1px solid var(--p7);color:#fff;border-radius:3px;font-weight:600;}
.btn-pink:hover,.btn-pink:focus{background:var(--p9);border-color:var(--p9);color:#fff;}
.btn-default{border-radius:3px;border-color:var(--ln);}
.control-label{font-weight:600;color:#6D4C5E;font-size:13px;}
.selectize-input.focus,.form-control:focus{border-color:var(--p5);box-shadow:0 0 0 2px rgba(240,98,146,.2);}
.form-check-input:checked{background-color:var(--p7);border-color:var(--p7);}
table.table{font-size:13.5px;}
table.table>thead>tr>th{background:var(--p7);color:#fff;border-bottom:none;font-weight:600;}
table.table>tbody>tr:nth-of-type(odd)>*{background:var(--p1);}
pre{background:var(--p1);border:1px solid var(--ln);color:var(--tx);font-size:12.5px;border-radius:3px;}
.nfilter{background:#fff;border:1px solid var(--ln);border-radius:3px;padding:8px 10px;margin-top:10px;font-size:13px;color:#6D4C5E;}
.nfilter b{color:var(--p9);font-size:16px;}
.irs--shiny .irs-bar{background:var(--p7);border-color:var(--p7);}
.help-note{color:#7d6672;font-size:13px;}
#uni_stem{white-space:pre;overflow:auto;max-height:400px;font-size:12px;}
"""


def sel_cat(id_, lab, sel):
    return ui.input_select(id_, lab, choices=cat_all, selected=sel)


try:
    d0 = read_sales(DATA_FILE)
except Exception:
    d0 = None


def ch(v, default):
    return default if d0 is None else [str(c) for c in d0[v].cat.categories]


def box(*children):
    return ui.div(*children, class_="box-x")


def col_(w, *children):
    return ui.column(w, *children)


app_ui = ui.page_fluid(
    ui.tags.head(ui.tags.meta(name="viewport", content="width=device-width, initial-scale=1"), ui.tags.style(ui.HTML(css))),
    ui.div(ui.h1(APP_TITLE), ui.p("Ujian Tengah Semester Eksplorasi dan Visualisasi Data Multivariat  |  ", NAMA, "  |  ", NIM, class_="sub"), class_="app-header"),
    ui.layout_sidebar(
        ui.sidebar(
            ui.h4("Sumber data"),
            ui.input_file("file", None, button_label="Pilih file", placeholder="Excel/CSV (opsional)", accept=[".xlsx", ".xls", ".csv"]),
            ui.h4("Filter data"),
            ui.input_checkbox_group("f_branch", "Branch", choices=ch("Branch", ["A", "B", "C"]), selected=ch("Branch", ["A", "B", "C"]), inline=True),
            ui.input_checkbox_group("f_cust", "Customer type", choices=ch("Customer type", ["Member", "Normal"]), selected=ch("Customer type", ["Member", "Normal"]), inline=True),
            ui.input_checkbox_group("f_gender", "Gender", choices=ch("Gender", ["Female", "Male"]), selected=ch("Gender", ["Female", "Male"]), inline=True),
            ui.input_checkbox_group("f_pay", "Payment", choices=ch("Payment", ["Cash", "Credit card", "Ewallet"]), selected=ch("Payment", ["Cash", "Credit card", "Ewallet"])),
            ui.input_select("f_prod", "Product line", choices=ch("Product line", []), selected=ch("Product line", []), multiple=True),
            ui.input_date_range("f_date", "Rentang tanggal", start=None if d0 is None else d0["Date"].min().date(), end=None if d0 is None else d0["Date"].max().date(),
                                format="dd-mm-yyyy", language="id"),
            ui.input_action_button("f_reset", "Atur ulang filter", class_="btn-default btn-sm"),
            ui.output_ui("n_info"),
            width=300,
        ),
        ui.navset_tab(
            # ------------------------------------------------------------ 1. DATA
            ui.nav_panel("1. Masalah dan Data",
                         ui.br(),
                         box(ui.h4("Latar belakang"), ui.p(LATAR_BELAKANG)),
                         box(ui.h4("Tujuan analisis"), ui.tags.ol(*[ui.tags.li(t) for t in TUJUAN])),
                         box(ui.h4("Sumber dan karakteristik data"), ui.p(SUMBER_DATA), ui.output_ui("data_info")),
                         box(ui.h4("Definisi variabel"), ui.output_table("kamus_tbl")),
                         box(ui.h4("Pengolahan data"), ui.output_ui("prep_info")),
                         box(ui.h4("Cuplikan data"), ui.output_table("preview"))),
            # -------------------------------------------------------- 2. UNIVARIAT
            ui.nav_panel("2. Univariat",
                         ui.br(),
                         ui.row(col_(4, ui.input_select("uni_var", "Variabel", choices=cat_all + num_all + ["Date"], selected="Total")), col_(4, ui.output_ui("uni_type_ui"))),
                         ui.row(col_(7, box(ui.output_ui("uni_main"))), col_(5, box(ui.h4("Statistik deskriptif"), ui.output_table("uni_tbl")))),
                         box(ui.h4("Interpretasi"), ui.output_ui("uni_interp"))),
            # --------------------------------------------------------- 3. BIVARIAT
            ui.nav_panel("3. Bivariat",
                         ui.br(),
                         ui.row(col_(3, ui.input_select("bi_x", "Variabel X", choices=cat_all + num_all, selected="Quantity")),
                                col_(3, ui.input_select("bi_y", "Variabel Y", choices=cat_all + num_all, selected="Total")),
                                col_(3, ui.output_ui("bi_type_ui"))),
                         ui.p("Jenis grafik menyesuaikan tipe pasangan variabel: numerik-numerik, kategorik-numerik, atau kategorik-kategorik.", class_="help-note"),
                         ui.row(col_(7, box(ui.output_plot("bi_plot", height="400px"))), col_(5, box(ui.h4("Ringkasan dan uji"), ui.output_table("bi_tbl")))),
                         box(ui.h4("Interpretasi"), ui.output_ui("bi_interp"))),
            # ------------------------------------------------------- 4. MULTIVARIAT
            ui.nav_panel("4. Multivariat",
                         ui.br(),
                         ui.navset_pill(
                             ui.nav_panel("Scatter multivariat", ui.br(),
                                          ui.row(col_(3, ui.input_select("mv_x", "Sumbu X", num_all, selected="Quantity")), col_(3, ui.input_select("mv_y", "Sumbu Y", num_all, selected="Total")),
                                                 col_(3, ui.input_select("mv_col", "Warna", [NONE] + cat_all + num_all, selected="Customer type")), col_(3, ui.input_select("mv_size", "Ukuran titik", [NONE] + num_all, selected="Rating"))),
                                          ui.row(col_(3, ui.input_select("mv_facet", "Panel (facet)", [NONE] + cat_all, selected=NONE)), col_(3, ui.input_checkbox("mv_lines", "Tampilkan garis regresi per kelompok", True))),
                                          box(ui.output_plot("mv_scatter", height="440px")), box(ui.h4("Interpretasi"), ui.output_ui("mv_scatter_i"))),
                             ui.nav_panel("Heatmap rata-rata", ui.br(),
                                          ui.row(col_(3, sel_cat("mv_ha", "Kategorik 1 (sumbu X)", "Branch")), col_(3, sel_cat("mv_hb", "Kategorik 2 (sumbu Y)", "Product line")),
                                                 col_(3, ui.input_select("mv_hn", "Variabel numerik", num_all, selected="Total")), col_(3, ui.input_select("mv_hf", "Ukuran", ["Rata-rata", "Median", "Jumlah", "Banyak transaksi"], selected="Rata-rata"))),
                                          box(ui.output_plot("mv_heat", height="420px")), box(ui.h4("Interpretasi"), ui.output_ui("mv_heat_i"))),
                             ui.nav_panel("Matriks korelasi", ui.br(),
                                          ui.row(col_(6, ui.input_checkbox_group("mv_cv", "Variabel numerik", num_all, selected=num_all, inline=True)), col_(3, ui.input_select("mv_cm", "Metode", ["pearson", "spearman"], selected="pearson"))),
                                          box(ui.output_plot("mv_cor", height="420px")), box(ui.h4("Interpretasi"), ui.output_ui("mv_cor_i"))),
                             ui.nav_panel("Tren waktu", ui.br(),
                                          ui.row(col_(3, ui.input_select("mv_tm", "Metrik", ["Jumlah transaksi", "Total penjualan (jumlah)", "Rata-rata Total", "Rata-rata Rating", "Total Quantity"], selected="Total penjualan (jumlah)")),
                                                 col_(3, ui.input_select("mv_tu", "Satuan waktu", ["Harian", "Mingguan", "Bulanan"], selected="Mingguan")), col_(3, ui.input_select("mv_tg", "Kelompok", [NONE] + cat_base, selected="Branch"))),
                                          box(ui.output_plot("mv_trend", height="420px")), box(ui.h4("Interpretasi"), ui.output_ui("mv_trend_i"))),
                             ui.nav_panel("PCA", ui.br(),
                                          ui.row(col_(6, ui.input_checkbox_group("mv_pv", "Variabel numerik (minimal 3)", num_all, selected=num_all, inline=True)), col_(3, ui.input_select("mv_pc", "Warna titik", [NONE] + cat_base, selected="Customer type"))),
                                          ui.row(col_(7, box(ui.output_plot("mv_pca", height="420px"))), col_(5, box(ui.h4("Keragaman dan loading"), ui.output_table("mv_pca_t")))),
                                          box(ui.h4("Interpretasi"), ui.output_ui("mv_pca_i"))),
                             ui.nav_panel("Scatterplot matrix", ui.br(),
                                          ui.row(col_(6, ui.input_checkbox_group("mv_mvars", "Variabel numerik (minimal 2)", num_all, selected=num_all, inline=True)), col_(3, ui.input_select("mv_mcol", "Warna titik", [NONE] + cat_base, selected=NONE))),
                                          box(ui.output_plot("mv_pairs", height="540px")), box(ui.h4("Interpretasi"), ui.output_ui("mv_pairs_i"))),
                             ui.nav_panel("Pencilan Mahalanobis", ui.br(),
                                          ui.row(col_(8, ui.input_checkbox_group("mv_dvars", "Variabel numerik (minimal 2)", num_all, selected=num_all, inline=True))),
                                          ui.row(col_(6, box(ui.output_plot("mv_mh1", height="380px"))), col_(6, box(ui.output_plot("mv_mh2", height="380px")))),
                                          box(ui.h4("Sepuluh observasi dengan jarak terbesar"), ui.output_table("mv_mh_t")), box(ui.h4("Interpretasi"), ui.output_ui("mv_mh_i"))),
                             ui.nav_panel("Q-Q plot Chi-Square", ui.br(),
                                          ui.row(col_(8, ui.input_checkbox_group("mv_qvars", "Variabel numerik (minimal 2)", num_all, selected=num_all, inline=True))),
                                          box(ui.output_plot("mv_qq", height="440px")), box(ui.h4("Interpretasi"), ui.output_ui("mv_qq_i"))),
                         )),
            # ---------------------------------------------------------- 5. REGRESI
            ui.nav_panel("5. Regresi Linear Berganda",
                         ui.br(),
                         box(ui.h4("Spesifikasi model"),
                             ui.row(col_(3, ui.input_select("rg_y", "Variabel respons (Y)", ["Total", "Rating"], selected="Total")),
                                    col_(5, ui.input_checkbox_group("rg_x", "Variabel prediktor (X)", ["Quantity", "Rating", "Hour"] + cat_all, selected=["Quantity", "Rating", "Hour"] + cat_base, inline=True)),
                                    col_(2, ui.input_radio_buttons("rg_log", "Transformasi respons", {"no": "Tanpa", "yes": "ln(Y)"}, selected="no")),
                                    col_(2, ui.br(), ui.input_action_button("rg_go", "Bangun model", class_="btn-pink"))),
                             ui.p("Respons Total dipilih karena merupakan ukuran utama nilai transaksi. Pilih minimal dua prediktor. Variabel kategorik otomatis diubah menjadi dummy dengan kategori pertama sebagai acuan.", class_="help-note")),
                         ui.navset_pill(
                             ui.nav_panel("Hasil model", ui.br(),
                                          box(ui.h4("Persamaan model"), ui.output_text_verbatim("rg_eq")),
                                          ui.row(col_(5, box(ui.h4("Ukuran kecocokan model"), ui.output_table("rg_fit"))), col_(7, box(ui.h4("Uji signifikansi per variabel (ANOVA)"), ui.output_table("rg_anova")))),
                                          box(ui.h4("Estimasi koefisien"), ui.output_table("rg_coef"), ui.p("Signif.: *** p<0.001, ** p<0.01, * p<0.05, . p<0.1", class_="help-note")),
                                          box(ui.h4("Interpretasi hasil"), ui.output_ui("rg_interp"))),
                             ui.nav_panel("Visualisasi hasil model", ui.br(),
                                          ui.row(col_(6, box(ui.output_plot("rg_p1", height="380px"))), col_(6, box(ui.output_plot("rg_p2", height="380px")))),
                                          box(ui.h4("Cara membaca"), ui.output_ui("rg_vis_i"))),
                             ui.nav_panel("Galeri visualisasi regresi", ui.br(),
                                          ui.row(col_(6, box(ui.output_plot("rv1", height="340px"))), col_(6, box(ui.output_plot("rv2", height="340px")))),
                                          ui.row(col_(6, box(ui.output_plot("rv3", height="340px"))), col_(6, box(ui.output_plot("rv4", height="340px")))),
                                          ui.row(col_(6, box(ui.output_plot("rv5", height="340px"))), col_(6, box(ui.output_plot("rv6", height="340px")))),
                                          ui.row(col_(6, box(ui.output_plot("rv7", height="340px"))), col_(6, box(ui.output_plot("rv8", height="340px")))),
                                          box(ui.h4("Cara membaca"), ui.output_ui("rv_i"))),
                             ui.nav_panel("Diagnostik asumsi", ui.br(),
                                          ui.row(col_(6, box(ui.output_plot("rg_d1", height="320px"))), col_(6, box(ui.output_plot("rg_d2", height="320px")))),
                                          ui.row(col_(6, box(ui.output_plot("rg_d3", height="320px"))), col_(6, box(ui.output_plot("rg_d4", height="320px")))),
                                          ui.row(col_(6, box(ui.h4("Uji asumsi"), ui.output_table("rg_tests"))), col_(6, box(ui.h4("Multikolinearitas (VIF)"), ui.output_table("rg_vif"))))),
                             ui.nav_panel("Keterbatasan model", ui.br(), box(ui.h4("Keterbatasan model yang digunakan"), bullets(KETERBATASAN))),
                         )),
            # ------------------------------------------------------- 6. KESIMPULAN
            ui.nav_panel("6. Kesimpulan",
                         ui.br(),
                         box(ui.h4("Kesimpulan utama"), ui.output_ui("kes_auto")),
                         box(ui.h4("Rekomendasi"), bullets(REKOMENDASI))),
            id="tabs",
        ),
    ),
)


# =============================================================================
# SERVER
# =============================================================================
from shiny import req

TBL = "table table-striped table-sm shiny-table w-auto"


def upper_pairs(m, vars_):
    """Pasangan segitiga atas, urutan kolom-utama (seperti as.table di R), diurutkan menurun menurut |r|."""
    p = len(vars_); out = [(vars_[i], vars_[j], m[i, j]) for j in range(p) for i in range(j)]
    return sorted(out, key=lambda t: -abs(t[2]))


def server(input, output, session):

    @reactive.calc
    def data_raw():
        f = input.file()
        d = read_sales(f[0]["datapath"]) if f else d0
        validate(d is not None, "Data tidak ditemukan. Letakkan Data_Penjualan.xlsx di folder aplikasi atau unggah file melalui panel kiri (kolom wajib: Branch, Customer type, Gender, Product line, Quantity, Total, Date, Time, Payment, Rating).")
        return d

    def lv(d, v):
        return [str(c) for c in d[v].cat.categories]

    # perbarui pilihan filter bila file baru diunggah
    @reactive.effect
    @reactive.event(input.file)
    def _upload():
        d = data_raw()
        ui.update_checkbox_group("f_branch", choices=lv(d, "Branch"), selected=lv(d, "Branch"), inline=True)
        ui.update_checkbox_group("f_cust", choices=lv(d, "Customer type"), selected=lv(d, "Customer type"), inline=True)
        ui.update_checkbox_group("f_gender", choices=lv(d, "Gender"), selected=lv(d, "Gender"), inline=True)
        ui.update_checkbox_group("f_pay", choices=lv(d, "Payment"), selected=lv(d, "Payment"))
        ui.update_select("f_prod", choices=lv(d, "Product line"), selected=lv(d, "Product line"))
        ui.update_date_range("f_date", start=d["Date"].min().date(), end=d["Date"].max().date())

    @reactive.effect
    @reactive.event(input.f_reset)
    def _reset():
        d = data_raw()
        ui.update_checkbox_group("f_branch", selected=lv(d, "Branch")); ui.update_checkbox_group("f_cust", selected=lv(d, "Customer type"))
        ui.update_checkbox_group("f_gender", selected=lv(d, "Gender")); ui.update_checkbox_group("f_pay", selected=lv(d, "Payment"))
        ui.update_select("f_prod", selected=lv(d, "Product line")); ui.update_date_range("f_date", start=d["Date"].min().date(), end=d["Date"].max().date())

    @reactive.calc
    def data_use():
        d = data_raw()
        fb = input.f_branch()
        d = d[d["Branch"].isin(list(fb))] if fb else d.iloc[0:0]
        d = d[d["Customer type"].isin(list(input.f_cust())) & d["Gender"].isin(list(input.f_gender())) & d["Payment"].isin(list(input.f_pay())) & d["Product line"].isin(list(input.f_prod()))]
        dr = input.f_date()
        if dr:
            d = d[(d["Date"] >= pd.Timestamp(dr[0])) & (d["Date"] <= pd.Timestamp(dr[1]))]
        validate(len(d) >= 30, "Data hasil filter terlalu sedikit (kurang dari 30 baris). Longgarkan filter pada panel kiri.")
        d = d.copy()
        for c in cat_all:                      # droplevels
            d[c] = d[c].cat.remove_unused_categories()
        return d.reset_index(drop=True)

    @render.ui
    def n_info():
        n0 = len(data_raw())
        try:
            n1 = len(data_use())
        except Exception:
            n1 = 0
        return ui.div("Transaksi terpilih:", ui.br(), ui.tags.b(f"{n1:,}"), f" dari {n0:,} ({fmt(100 * n1 / n0, 1)}%)", class_="nfilter")

    # ---- 1. DATA ----
    @render.ui
    def data_info():
        d = data_raw()
        return bullets(["Jumlah observasi: **%s** transaksi; jumlah variabel: **10** variabel asli (ditambah 3 variabel turunan: Hour, Bulan, Hari)." % f"{len(d):,}",
                        "Periode data: **%s** sampai **%s**." % (d["Date"].min().date(), d["Date"].max().date()),
                        "Variabel kategorik: Branch, Customer type, Gender, Product line, Payment. Variabel numerik: Quantity, Total, Rating, serta Hour (turunan)."])

    @render.table(classes=TBL)
    def kamus_tbl():
        return kamus

    @render.ui
    def prep_info():
        d = data_raw(); q1, q3 = np.quantile(d["Total"], [.25, .75]); out = int(((d["Total"] < q1 - 1.5 * (q3 - q1)) | (d["Total"] > q3 + 1.5 * (q3 - q1))).sum())
        return bullets(["**Data hilang:** %d sel kosong pada seluruh kolom analisis." % int(d.isna().sum().sum()),
                        "**Data ganda:** %d baris duplikat." % int(d.duplicated().sum()),
                        "**Pencilan:** terdapat %d pencilan pada Total (aturan 1.5 x IQR). Pencilan dipertahankan karena merupakan transaksi bernilai besar yang valid." % out,
                        "**Pengolahan:** variabel Hour, Bulan, dan Hari dibentuk dari Time dan Date; variabel kategorik diubah menjadi faktor.",
                        "**Transformasi:** tidak dilakukan pada eksplorasi; pada regresi tersedia pilihan transformasi ln(Y) untuk mengatasi ragam yang tidak konstan."])

    @render.table(classes=TBL)
    def preview():
        x = data_raw().head(10).copy(); x["Date"] = x["Date"].dt.strftime("%Y-%m-%d")
        return x

    # ---- 2. UNIVARIAT ----
    @render.ui
    def uni_type_ui():
        return ui.input_select("uni_type", "Jenis grafik", uni_choices(input.uni_var()))

    @reactive.calc
    def uni_type_val():
        chs = uni_choices(input.uni_var()); t = inp(input.uni_type)
        return chs[0] if (t is None or t not in chs) else t

    @render.ui
    def uni_main():
        if uni_type_val() == "Stem-and-leaf":
            return ui.output_text_verbatim("uni_stem")
        return ui.output_plot("uni_plot", height="380px")

    @render.plot
    def uni_plot():
        return plot_uni(data_use(), input.uni_var(), uni_type_val())

    @render.text
    def uni_stem():
        req(input.uni_var() in num_all)
        x = data_use()[input.uni_var()].to_numpy(float)
        return "\n".join(["Stem-and-leaf plot: " + input.uni_var(), "", stem_text(x)])

    @render.table(classes=TBL)
    def uni_tbl():
        return uni_table(data_use(), input.uni_var())

    @render.ui
    def uni_interp():
        return bullets(interp_uni(data_use(), input.uni_var()))

    # ---- 3. BIVARIAT ----
    @render.ui
    def bi_type_ui():
        return ui.input_select("bi_type", "Jenis grafik", bi_choices(input.bi_x(), input.bi_y()))

    @reactive.calc
    def bi_type_val():
        chs = bi_choices(input.bi_x(), input.bi_y()); t = inp(input.bi_type)
        return chs[0] if (t is None or t not in chs) else t

    @render.plot
    def bi_plot():
        validate(input.bi_x() != input.bi_y(), "Pilih dua variabel yang berbeda.")
        return plot_bi(data_use(), input.bi_x(), input.bi_y(), bi_type_val())

    @reactive.calc
    def bi_res():
        validate(input.bi_x() != input.bi_y(), "Pilih dua variabel yang berbeda.")
        return bi_text(data_use(), input.bi_x(), input.bi_y())

    @render.table(classes=TBL)
    def bi_tbl():
        return bi_res()["table"]

    @render.ui
    def bi_interp():
        return bullets(bi_res()["lines"])

    # ---- 4. MULTIVARIAT ----
    @render.plot
    def mv_scatter():
        validate(input.mv_x() != input.mv_y(), "Pilih sumbu X dan Y yang berbeda.")
        return plot_mv_scatter(data_use(), input.mv_x(), input.mv_y(), input.mv_col(), input.mv_size(), input.mv_facet(), input.mv_lines())

    @render.ui
    def mv_scatter_i():
        d = data_use(); x = input.mv_x(); y = input.mv_y(); col = input.mv_col(); r = d[x].corr(d[y])
        ln = ["Secara keseluruhan korelasi **%s** dan **%s** adalah **%s** (%s)." % (x, y, fmt(r, 3), strength(r))]
        if col != NONE and col not in num_all:
            gr = {str(k): grp[x].corr(grp[y]) for k, grp in d.groupby(col, observed=True)}
            ln.append("Korelasi per **%s**: %s. %s" % (col, "; ".join(f"{k} = {fmt(v, 2)}" for k, v in gr.items()),
                                                      "Pola hubungan hampir sama di semua kelompok sehingga variabel warna tidak mengubah hubungan X dan Y." if (max(gr.values()) - min(gr.values())) < .1
                                                      else "Kekuatan hubungan berbeda antar kelompok, menandakan kemungkinan interaksi."))
        if input.mv_size() != NONE:
            ln.append("Ukuran titik menunjukkan **%s**; jika titik besar dan kecil tersebar acak, variabel tersebut tidak berkaitan dengan posisi titik pada X dan Y." % input.mv_size())
        return bullets(ln)

    @render.plot
    def mv_heat():
        validate(input.mv_ha() != input.mv_hb(), "Pilih dua variabel kategorik yang berbeda.")
        return plot_mv_heat(data_use(), input.mv_ha(), input.mv_hb(), input.mv_hn(), input.mv_hf())

    @render.ui
    def mv_heat_i():
        validate(input.mv_ha() != input.mv_hb(), "Pilih dua variabel kategorik yang berbeda.")
        r = mv_heat_data(data_use(), input.mv_ha(), input.mv_hb(), input.mv_hn(), input.mv_hf()); hi = r.loc[r["val"].idxmax()]; lo = r.loc[r["val"].idxmin()]
        return bullets(["Nilai tertinggi pada kombinasi **%s - %s** (%s) dan terendah pada **%s - %s** (%s)." % (hi["A"], hi["B"], fmt(hi["val"], 1), lo["A"], lo["B"], fmt(lo["val"], 1)),
                        "Selisih tertinggi-terendah sebesar **%s** (%s%% dari nilai tertinggi)." % (fmt(hi["val"] - lo["val"], 1), fmt(100 * (hi["val"] - lo["val"]) / hi["val"], 1)),
                        "Warna lebih pekat menandakan nilai lebih besar. Perbedaan antar sel perlu dikonfirmasi dengan uji statistik sebelum disimpulkan sebagai perbedaan nyata."])

    @render.plot
    def mv_cor():
        validate(len(input.mv_cv()) >= 2, "Pilih minimal dua variabel.")
        return plot_mv_cor(data_use(), list(input.mv_cv()), input.mv_cm())

    @render.ui
    def mv_cor_i():
        validate(len(input.mv_cv()) >= 2, "Pilih minimal dua variabel.")
        vs = list(input.mv_cv()); pr = upper_pairs(data_use()[vs].corr(method=input.mv_cm()).to_numpy(), vs)
        return bullets(["Korelasi terkuat: **%s - %s** (r = %s; %s)." % (pr[0][0], pr[0][1], fmt(pr[0][2], 3), strength(pr[0][2])),
                        "Korelasi terlemah: **%s - %s** (r = %s)." % (pr[-1][0], pr[-1][1], fmt(pr[-1][2], 3)),
                        "Warna merah muda tua = korelasi positif, biru keabuan = negatif, putih = mendekati nol."])

    @render.plot
    def mv_trend():
        return plot_mv_trend(data_use(), input.mv_tm(), input.mv_tu(), input.mv_tg())

    @render.ui
    def mv_trend_i():
        td = trend_data(data_use(), input.mv_tm(), input.mv_tu(), input.mv_tg()); hi = td.loc[td["val"].idxmax()]; lo = td.loc[td["val"].idxmin()]
        gtxt = f" (kelompok {hi[input.mv_tg()]})" if input.mv_tg() != NONE else ""
        return bullets(["Nilai **%s** tertinggi pada periode **%s**%s sebesar %s; terendah pada **%s** sebesar %s." % (input.mv_tm(), hi["Periode"].date(), gtxt, fmt(hi["val"], 1), lo["Periode"].date(), fmt(lo["val"], 1)),
                        "Bandingkan garis antar kelompok untuk melihat apakah pola naik-turun sama (pola serupa menandakan faktor waktu yang umum, bukan khas satu kelompok).",
                        "Minggu pertama dan terakhir dapat berisi lebih sedikit hari sehingga nilainya tampak lebih rendah." if input.mv_tu() == "Mingguan"
                        else "Perhatikan bahwa data hanya mencakup beberapa bulan, sehingga pola musiman tidak dapat disimpulkan."])

    @render.plot
    def mv_pca():
        validate(len(input.mv_pv()) >= 3, "Pilih minimal tiga variabel numerik.")
        return plot_mv_pca(data_use(), list(input.mv_pv()), input.mv_pc())

    @render.table(classes=TBL)
    def mv_pca_t():
        validate(len(input.mv_pv()) >= 3, "Pilih minimal tiga variabel numerik.")
        pr = pca_fit(data_use(), list(input.mv_pv())); ve = 100 * pr["sdev"] ** 2 / np.sum(pr["sdev"] ** 2); ld = pr["rotation"][["PC1", "PC2"]].round(3)
        t = pd.DataFrame({"Variabel": list(ld.index), "PC1": ld["PC1"].to_numpy(), "PC2": ld["PC2"].to_numpy()})
        return pd.concat([t, pd.DataFrame({"Variabel": ["Keragaman dijelaskan (%)"], "PC1": [round(ve[0], 1)], "PC2": [round(ve[1], 1)]})], ignore_index=True)

    @render.ui
    def mv_pca_i():
        validate(len(input.mv_pv()) >= 3, "Pilih minimal tiga variabel numerik.")
        pr = pca_fit(data_use(), list(input.mv_pv())); ve = 100 * pr["sdev"] ** 2 / np.sum(pr["sdev"] ** 2); l1 = pr["rotation"]["PC1"]; k = l1.abs().idxmax()
        return bullets(["PC1 menjelaskan **%s%%** dan PC2 **%s%%** keragaman; bersama-sama **%s%%**." % (fmt(ve[0], 1), fmt(ve[1], 1), fmt(ve[0] + ve[1], 1)),
                        "Variabel dengan kontribusi terbesar pada PC1: **%s** (loading %s)." % (k, fmt(l1[k], 2)),
                        "Panah yang searah menandakan variabel berkorelasi positif; panah yang saling tegak lurus menandakan variabel hampir tidak berkorelasi. Jika keragaman terbagi merata ke banyak komponen, variabel-variabel tersebut membawa informasi yang berbeda-beda."])

    # ---- 4b. SCATTERPLOT MATRIX, MAHALANOBIS, Q-Q CHI-SQUARE ----
    @render.plot
    def mv_pairs():
        validate(len(input.mv_mvars()) >= 2, "Pilih minimal dua variabel numerik.")
        return plot_mv_pairs(data_use(), list(input.mv_mvars()), input.mv_mcol())

    @render.ui
    def mv_pairs_i():
        validate(len(input.mv_mvars()) >= 2, "Pilih minimal dua variabel numerik.")
        vs = list(input.mv_mvars()); pr = upper_pairs(data_use()[vs].corr().to_numpy(), vs)
        return bullets(["Diagonal: histogram tiap variabel; segitiga bawah: scatterplot dengan kurva lowess; segitiga atas: koefisien korelasi Pearson (semakin besar angka, semakin kuat hubungan).",
                        "Pasangan dengan hubungan terkuat: **%s - %s** (r = %s; %s). Pasangan terlemah: **%s - %s** (r = %s)." % (pr[0][0], pr[0][1], fmt(pr[0][2], 3), strength(pr[0][2]), pr[-1][0], pr[-1][1], fmt(pr[-1][2], 3)),
                        "Kurva lowess yang mendekati garis lurus menandakan hubungan linear; kurva melengkung menandakan hubungan nonlinear. Titik yang jauh dari kumpulan utama adalah kandidat pencilan."])

    @reactive.calc
    def mh():
        validate(len(input.mv_dvars()) >= 2, "Pilih minimal dua variabel numerik.")
        return mahal_fit(data_use(), list(input.mv_dvars()))

    @render.plot
    def mv_mh1():
        return plot_mv_mahal_idx(mh())

    @render.plot
    def mv_mh2():
        return plot_mv_mahal_sc(mh(), list(input.mv_dvars()))

    @render.table(classes=TBL)
    def mv_mh_t():
        return mahal_table(data_use(), mh(), list(input.mv_dvars()))

    @render.ui
    def mv_mh_i():
        f = mh(); d = data_use(); vs = list(input.mv_dvars()); n = len(f["df"]); k = int(f["df"]["out"].sum()); top = f["df"].loc[f["df"]["D2"].idxmax()]
        z = {v: (d[v].to_numpy()[int(top["idx"]) - 1] - d[v].mean()) / sd1(d[v]) for v in vs}; vmax = max(z, key=lambda v: abs(z[v]))
        return bullets(["Batas pencilan multivariat: Chi-Square 97.5%% dengan df = %d, yaitu **%s**. Observasi dengan D2 di atas batas dianggap pencilan." % (f["p"], fmt(f["cut"], 2)),
                        "Ditemukan **%d pencilan** (**%s%%** dari %d observasi). Pada data normal multivariat, proporsi yang diharapkan sekitar 2.5%%." % (k, fmt(100 * k / n, 1), n),
                        "Observasi paling ekstrem: baris **%d** (D2 = %s). Variabel yang paling menyimpang: **%s** (z = %s)." % (top["idx"], fmt(top["D2"], 2), vmax, fmt(z[vmax], 2)),
                        "Pencilan multivariat dapat berupa kombinasi nilai yang tidak lazim walau tiap variabel tampak wajar. Pencilan sebaiknya diperiksa dulu keabsahannya sebelum dihapus."])

    @reactive.calc
    def qq():
        validate(len(input.mv_qvars()) >= 2, "Pilih minimal dua variabel numerik.")
        return mahal_fit(data_use(), list(input.mv_qvars()))

    @render.plot
    def mv_qq():
        return plot_mv_qqchi(qq())

    @render.ui
    def mv_qq_i():
        f = qq(); n = len(f["df"]); s = np.sort(f["df"]["D2"].to_numpy()); q = sps.chi2.ppf((np.arange(1, n + 1) - .5) / n, df=f["p"]); r = np.corrcoef(s, q)[0, 1]
        tail3 = f["df"]["idx"].to_numpy()[np.argsort(f["df"]["D2"].to_numpy(), kind="stable")][-3:]
        return bullets(["Korelasi antara D2 terurut dan kuantil Chi-Square = **%s**; semakin mendekati 1, semakin sesuai data dengan normal multivariat." % fmt(r, 3),
                        "Titik terjauh dari garis berada di ujung kanan (baris %s), yang menandakan observasi dengan jarak Mahalanobis besar." % ", ".join(str(i) for i in tail3),
                        "Titik yang melengkung ke atas di ujung kanan menandakan ekor distribusi lebih berat atau adanya pencilan; pola berundak dapat muncul karena variabel diskret (Quantity, Hour). Asumsi normalitas multivariat dianggap kurang terpenuhi bila penyimpangan jelas."])

    # ---- 5. REGRESI ----
    @reactive.calc
    def logy():
        return input.rg_log() == "yes"

    @reactive.calc
    @reactive.event(input.rg_go, ignore_none=False)
    def model():
        preds = [p for p in input.rg_x() if p != input.rg_y()]
        validate(len(preds) >= 2, "Pilih minimal dua variabel prediktor (selain variabel respons).")
        d = data_use()
        if logy():
            validate(bool((d[input.rg_y()] > 0).all()), "Transformasi ln hanya untuk respons bernilai positif.")
        return dict(m=build_model(d, input.rg_y(), preds, logy()), resp=input.rg_y(), logy=logy(), preds=preds, d=d)

    @reactive.calc
    def M():
        return model()["m"]

    @render.text
    def rg_eq():
        return make_equation(M(), model()["resp"], model()["logy"])

    @render.table(classes=TBL)
    def rg_fit():
        return fit_table(M())

    @render.table(classes=TBL)
    def rg_anova():
        return anova_table(M())

    @render.table(classes=TBL)
    def rg_coef():
        return coef_table(M())

    @render.ui
    def rg_interp():
        return bullets(interp_reg(M(), model()["resp"], model()["logy"]))

    @render.plot
    def rg_p1():
        return plot_reg_avp(M(), model()["logy"], model()["resp"])

    @render.plot
    def rg_p2():
        return plot_reg_coef(M())

    @render.ui
    def rg_vis_i():
        r2 = M().rsquared
        return bullets(["**Aktual vs prediksi:** semakin rapat titik di sekitar garis putus-putus, semakin baik prediksi model (R2 = %s). Titik yang menyebar lebar menunjukkan keragaman yang belum dijelaskan model." % fmt(r2, 3),
                        "**Koefisien dan selang kepercayaan:** titik = estimasi, garis = selang 95%. Selang yang tidak memotong garis nol (warna tua) menandakan prediktor signifikan pada taraf 5%. Skala koefisien bergantung pada satuan variabelnya, jadi bandingkan terutama ada-tidaknya pengaruh, bukan panjang garisnya."])

    @render.plot
    def rg_d1():
        return plot_reg_resid(M())

    @render.plot
    def rg_d2():
        return plot_reg_qq(M())

    @render.plot
    def rg_d3():
        return plot_reg_scale(M())

    @render.plot
    def rg_d4():
        return plot_reg_cook(M())

    @render.table(classes=TBL)
    def rg_tests():
        m = M(); sh = _shapiro_p(m.resid); bp = bp_test(m); cd = reg_df(m)["cook"].to_numpy(); ninf = int(np.sum(cd > 4 / len(cd)))
        return pd.DataFrame({"Uji": ["Shapiro-Wilk (normalitas residual)", "Breusch-Pagan (kesamaan ragam)", "Observasi berpengaruh (Cook's D > 4/n)"],
                             "Statistik": [fmt(sh.statistic, 4), fmt(bp["stat"], 3), str(ninf)],
                             "p-value": [fmt_p(sh.pvalue), fmt_p(bp["p"]), "-"],
                             "Keputusan": ["Tidak normal" if sh.pvalue < .05 else "Normal", "Ragam tidak konstan" if bp["p"] < .05 else "Ragam konstan", f"{ninf} observasi"]})

    @render.table(classes=TBL)
    def rg_vif():
        v = vif_table(M()); validate(v is not None, "VIF membutuhkan minimal dua prediktor.")
        v["Keputusan"] = np.where(v["VIF setara"].round(3) < 5, "Aman", "Perlu perhatian"); v["VIF setara"] = fmt(v["VIF setara"].to_numpy(), 3)
        return v

    # ---- 5b. GALERI VISUALISASI REGRESI ----
    @render.plot
    def rv1():
        return plot_reg_scatter(model()["d"], model()["resp"], model()["preds"], model()["logy"])

    @render.plot
    def rv2():
        return plot_reg_cor(model()["d"], model()["resp"], model()["preds"])

    @render.plot
    def rv3():
        return plot_reg_coef(M())

    @render.plot
    def rv4():
        return plot_reg_qq(M())

    @render.plot
    def rv5():
        return plot_reg_hist(M())

    @render.plot
    def rv6():
        return plot_reg_ts(M(), model()["d"])

    @render.plot
    def rv7():
        return plot_reg_acf(M(), model()["d"])

    @render.plot
    def rv8():
        return plot_reg_obsres(M())

    @render.ui
    def rv_i():
        m = M(); o = reg_ord_res(m, model()["d"]); e = o["res"]; n = len(e)
        a = sm_acf(e, nlags=min(30, n - 1), fft=False); ci = sps.norm.ppf(.975) / np.sqrt(n); nsig = int(np.sum(np.abs(a[1:]) > ci)); nl = len(a) - 1
        dw = np.sum(np.diff(e) ** 2) / np.sum(e ** 2); sw = _shapiro_p(e).pvalue
        return bullets(["**Scatterplot:** menunjukkan bentuk hubungan respons dengan tiap prediktor numerik; garis yang jelas menanjak/menurun menandakan hubungan linear.",
                        "**Correlation matrix:** korelasi respons dengan prediktor menunjukkan kandidat prediktor kuat; korelasi tinggi antar prediktor mengindikasikan potensi multikolinearitas.",
                        "**Coefficient CI plot:** selang 95% yang tidak memotong garis nol berarti koefisien signifikan pada taraf 5%.",
                        "**Normal probability plot dan histogram residual:** titik mengikuti garis dan histogram mendekati kurva normal berarti residual normal. Hasil Shapiro-Wilk p %s, sehingga residual %s." % (pq(sw), "**tidak normal**" if sw < .05 else "normal"),
                        "**Time series plot dan ACF:** statistik Durbin-Watson = %s (sekitar 2 berarti tidak ada autokorelasi); %d dari %d lag berada di luar batas 95%%. %s" % (
                            fmt(dw, 3), nsig, nl, "Tidak ada indikasi autokorelasi residual yang berarti." if (nsig <= 0.05 * nl + 1 and abs(dw - 2) < .4) else "Terdapat indikasi autokorelasi residual; periksa pola waktu."),
                        "**Observed vs residual:** residual idealnya tersebar acak di sekitar nol. Sebaran yang melebar pada nilai observasi besar menandakan ragam residual tidak konstan (heteroskedastis)."])

    # ---- 6. KESIMPULAN ----
    @render.ui
    def kes_auto():
        return bullets(auto_conclusion(data_use(), M(), model()["resp"], model()["logy"]))


app = App(app_ui, server)

if __name__ == "__main__":
    app.run()
