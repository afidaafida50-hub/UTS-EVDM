import re
import html as _html
import zlib
import warnings
from io import BytesIO
from contextlib import contextmanager
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap, Normalize
from matplotlib.lines import Line2D
from scipy import stats
import statsmodels.api as sm
from statsmodels.nonparametric.smoothers_lowess import lowess
from statsmodels.tsa.stattools import acf as sm_acf

warnings.filterwarnings("ignore")

# KONFIGURASI
DATA_FILE   = "Data_Penjualan.xlsx"
NAMA        = "Afida Mar'atus Solihah"
NIM         = "231064005"
APP_TITLE   = "Analisis Faktor yang Memengaruhi Nilai Transaksi Penjualan Supermarket"
SUMBER_DATA = "Dataset sekunder data transaksi supermarket (Data_Penjualan.xlsx). Isi dengan tautan/sumber asli yang dapat dipertanggungjawabkan (mis. Kaggle: Supermarket Sales)."
LATAR_BELAKANG = "Supermarket perlu memahami faktor yang berkaitan dengan besarnya nilai transaksi agar strategi promosi, penataan produk, dan pelayanan dapat dirancang berdasarkan data. Nilai transaksi dapat dipengaruhi oleh jumlah barang yang dibeli, karakteristik pelanggan, lini produk, cabang, metode pembayaran, dan waktu transaksi. Analisis eksplorasi dan regresi linear berganda digunakan untuk mengidentifikasi faktor-faktor tersebut."
TUJUAN = [
    "Menggambarkan karakteristik setiap variabel pada data transaksi (univariat).",
    "Mengidentifikasi hubungan dan perbedaan kelompok antar dua variabel (bivariat).",
    "Mengeksplorasi pola yang melibatkan tiga variabel atau lebih (multivariat).",
    "Membangun model regresi linear berganda untuk menjelaskan nilai transaksi (Total) dan mengevaluasi kemampuannya.",
    "Merumuskan kesimpulan dan rekomendasi bagi pengelola supermarket.",
]

st.set_page_config(page_title=APP_TITLE, layout="wide")

# ---- Konstanta variabel ------------------------------------------------------
cat_base = ["Branch", "Customer type", "Gender", "Product line", "Payment"]
cat_all  = cat_base + ["Bulan", "Hari"]
num_all  = ["Quantity", "Total", "Rating", "Hour"]
NONE = "(tidak ada)"
bulan_lv = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus",
            "September", "Oktober", "November", "Desember"]
hari_lv  = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]

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
def hour_of(t):
    def f(x):
        if hasattr(x, "hour"):
            return x.hour
        try:
            return int(str(x)[:2])
        except Exception:
            return np.nan
    h = pd.Series([f(x) for x in t], dtype="float64")
    return h.astype("int64") if h.notna().all() else h


def read_sales(src, ext):
    try:
        d = pd.read_excel(src) if ext in ("xlsx", "xls") else pd.read_csv(src)
    except Exception:
        return None
    need = ["Branch", "Customer type", "Gender", "Product line", "Quantity", "Total", "Date", "Time", "Payment", "Rating"]
    if not all(c in d.columns for c in need):
        return None
    d = d.copy()
    dt = pd.to_datetime(d["Date"], errors="coerce")
    if dt.isna().mean() > .5:
        dt = pd.to_datetime(d["Date"], format="%m/%d/%Y", errors="coerce")
    d["Date"] = dt.dt.normalize()
    d["Hour"] = hour_of(d["Time"])
    for v in ["Quantity", "Total", "Rating"]:
        d[v] = pd.to_numeric(d[v], errors="coerce")
    for v in cat_base:
        s = d[v].astype(str)
        d[v] = pd.Categorical(s, categories=sorted(s.unique()))
    mon = d["Date"].dt.month
    wd = d["Date"].dt.weekday
    present = sorted(mon.dropna().unique())
    d["Bulan"] = pd.Categorical([bulan_lv[int(m) - 1] for m in mon], categories=[bulan_lv[int(m) - 1] for m in present])
    d["Hari"] = pd.Categorical([hari_lv[int(w)] for w in wd], categories=hari_lv)
    return d[cat_base + ["Bulan", "Hari", "Quantity", "Total", "Rating", "Hour", "Date"]].reset_index(drop=True)


@st.cache_data(show_spinner=False)
def load_default(path):
    p = Path(path)
    if not p.exists():
        p = Path(__file__).parent / path
    if not p.exists():
        return None
    return read_sales(str(p), p.suffix.lower().lstrip("."))


@st.cache_data(show_spinner=False)
def load_upload(raw, ext):
    return read_sales(BytesIO(raw), ext)


# ---- Palet & tema grafik ---------------------------------------
pk = dict(d="#880E4F", m="#C2185B", s="#AD1457", l="#F48FB1", xl="#F8BBD0",
          bg="#FFF5F8", txt="#3B2A33", grid="#EFDDE4", mauve="#6D4C5E", slate="#4F6D8F")
_PAL = ["#AD1457", "#F48FB1", "#6D4C5E", "#F8BBD0", "#D81B60", "#B39DAB", "#880E4F"]


def pal_cat(n):
    return [_PAL[i % len(_PAL)] for i in range(n)]


CM_SEQ = LinearSegmentedColormap.from_list("pinkseq", [pk["bg"], pk["d"]])
CM_DIV = LinearSegmentedColormap.from_list("pinkdiv", [pk["slate"], "white", pk["s"]])
CM_PTS = LinearSegmentedColormap.from_list("pinkpts", [pk["xl"], pk["d"]])

plt.rcParams.update({
    "figure.facecolor": "white", "axes.facecolor": "white", "axes.edgecolor": pk["grid"],
    "axes.grid": True, "grid.color": pk["grid"], "grid.linewidth": .8, "axes.axisbelow": True,
    "axes.spines.top": False, "axes.spines.right": False, "axes.spines.left": False, "axes.spines.bottom": False,
    "axes.labelcolor": pk["mauve"], "axes.labelweight": "bold", "xtick.color": pk["txt"], "ytick.color": pk["txt"],
    "xtick.major.size": 0, "ytick.major.size": 0, "text.color": pk["txt"], "font.size": 10,
    "legend.frameon": False, "figure.dpi": 100,
})


def ttl(ax, t, sub=None):
    ax.set_title(t, loc="left", fontweight="bold", color=pk["d"], fontsize=13, pad=22 if sub else 10)
    if sub:
        ax.text(0, 1.02, sub, transform=ax.transAxes, color=pk["mauve"], fontsize=9, va="bottom")


def fig_ttl(fig, t, sub=None):
    fig.text(0.01, 0.985, t, fontweight="bold", color=pk["d"], fontsize=13, va="top")
    if sub:
        fig.text(0.01, 0.94, sub, color=pk["mauve"], fontsize=9, va="top")


def show(fig, top=None):
    rect = getattr(fig, "_rect", None) or [0, 0, 1, top or getattr(fig, "_top", None) or 1]
    fig.tight_layout(rect=rect)
    st.pyplot(fig)
    plt.close(fig)


def rot(ax, angle=20):
    plt.setp(ax.get_xticklabels(), rotation=angle, ha="right")


def wrap_grid(n):
    ncol = int(np.ceil(np.sqrt(n)))
    return int(np.ceil(n / ncol)), ncol


# ---- Fungsi bantu ------------------------------------------------------------
def fmt(x, k=2):
    return f"{x:,.{k}f}"


def fmt_p(p):
    return "< 0.001" if p < 0.001 else f"{p:.3f}"


def pq(p):  # untuk kalimat: "p < 0.001" / "p = 0.019"
    return "< 0.001" if p < 0.001 else "= " + f"{p:.3f}"


def stars(p):
    return "***" if p < .001 else "**" if p < .01 else "*" if p < .05 else "." if p < .1 else ""


def skew_fun(x):
    x = np.asarray(x, float); x = x[~np.isnan(x)]; n = len(x)
    return n / ((n - 1) * (n - 2)) * np.sum(((x - x.mean()) / x.std(ddof=1)) ** 3)


def kurt_fun(x):
    x = np.asarray(x, float); x = x[~np.isnan(x)]; n = len(x); z = (x - x.mean()) / x.std(ddof=1)
    return n * (n + 1) / ((n - 1) * (n - 2) * (n - 3)) * np.sum(z ** 4) - 3 * (n - 1) ** 2 / ((n - 2) * (n - 3))


def md(x):
    return re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", _html.escape(x))


def bullets(lines):
    st.markdown("<ul>" + "".join(f"<li>{md(l)}</li>" for l in lines if l) + "</ul>", unsafe_allow_html=True)


def strength(r):
    a = abs(r)
    return "sangat lemah" if a < .2 else "lemah" if a < .4 else "sedang" if a < .6 else "kuat" if a < .8 else "sangat kuat"


class Need(Exception):
    """Padanan validate(need(...)) pada Shiny."""


def need(cond, msg):
    if not cond:
        raise Need(msg)


@contextmanager
def guard():
    try:
        yield
    except Need as e:
        st.warning(str(e))


def box(title=None):
    c = st.container(border=True)
    if title:
        c.markdown(f"#### {title}")
    return c


def table(df):
    st.dataframe(df.astype(str) if not isinstance(df, str) else df, hide_index=True)


def nrd0(x):
    x = np.asarray(x, float); s = x.std(ddof=1)
    iqr = np.subtract(*np.percentile(x, [75, 25]))
    lo = min(s, iqr / 1.34) if iqr > 0 else s
    if lo == 0:
        lo = abs(x[0]) if x[0] != 0 else 1.0
    return .9 * lo * len(x) ** -.2


def kde(x, grid=None, n=512):
    x = np.asarray(x, float); h = nrd0(x)
    if grid is None:
        grid = np.linspace(x.min() - 3 * h, x.max() + 3 * h, n)
    k = stats.gaussian_kde(x, bw_method=h / x.std(ddof=1))
    return grid, k(grid)


def lm_line(ax, x, y, color, lw=1.0, label=None, ls="-"):
    x = np.asarray(x, float); y = np.asarray(y, float)
    if len(np.unique(x)) < 2:
        return
    b = np.polyfit(x, y, 1); g = np.array([x.min(), x.max()])
    ax.plot(g, np.polyval(b, g), color=color, lw=lw, label=label, ls=ls)


def loess_line(ax, x, y, color, lw=1.0, label=None):
    x = np.asarray(x, float); y = np.asarray(y, float)
    if len(np.unique(x)) < 4:
        return
    r = lowess(y, x, frac=.75, return_sorted=True)
    ax.plot(r[:, 0], r[:, 1], color=color, lw=lw, label=label)


# =============================================================================
# 2. UNIVARIAT
# =============================================================================
def uni_choices(v):
    if v == "Date":
        return ["Deret waktu harian"]
    if v in cat_all:
        return ["Diagram batang", "Pareto", "Diagram lingkaran"]
    return ["Histogram", "Histogram + Density", "Boxplot", "Density", "Stem-and-leaf", "QQ-plot"]


def plot_uni(d, v, type_):
    if v == "Date":
        dd = d.groupby("Date").size()
        xs = dd.index.map(pd.Timestamp.toordinal).values.astype(float)
        fig, ax = plt.subplots(figsize=(7, 3.9))
        ax.plot(dd.index, dd.values, color=pk["l"]); ax.scatter(dd.index, dd.values, color=pk["m"], s=14, zorder=3)
        r = lowess(dd.values, xs, frac=.75, return_sorted=True)
        ax.plot(pd.to_datetime([pd.Timestamp.fromordinal(int(o)) for o in r[:, 0]]), r[:, 1], color=pk["d"], lw=2)
        ttl(ax, "Jumlah transaksi per hari", "Garis tebal = tren (loess)")
        ax.set_xlabel("Tanggal"); ax.set_ylabel("Jumlah transaksi")
        return fig
    if v in cat_all:
        vc = d[v].value_counts(sort=False)
        tab = pd.DataFrame({"Kategori": vc.index.astype(str), "n": vc.values})
        tab["pct"] = 100 * tab["n"] / tab["n"].sum()
        cols = pal_cat(len(tab))
        fig, ax = plt.subplots(figsize=(7, 3.9))
        if type_ == "Diagram lingkaran":
            ax.set_facecolor("white"); ax.grid(False)
            w, _, at = ax.pie(tab["n"], colors=cols, startangle=90, counterclock=False, autopct=lambda p: f"{p:.1f}%",
                              pctdistance=.65, wedgeprops=dict(edgecolor="white"), textprops=dict(color="white", fontweight="bold"))
            ax.legend(w, tab["Kategori"], title=v, loc="center left", bbox_to_anchor=(1, .5))
            ax.set_title(f"Proporsi {v}", loc="left", fontweight="bold", color=pk["d"], fontsize=13)
            return fig
        if type_ == "Pareto":
            tab = tab.sort_values("n", ascending=False).reset_index(drop=True)
            tab["kum"] = tab["n"].cumsum() / tab["n"].sum() * 100; tot = tab["n"].sum()
            ax.bar(tab["Kategori"], tab["n"], color=pk["l"], width=.7)
            ax.plot(tab["Kategori"], tab["kum"] * tot / 100, color=pk["d"], lw=2, marker="o", ms=5)
            for i, r in tab.iterrows():
                ax.text(i, r["kum"] * tot / 100 + tot * .02, f"{r['kum']:.1f}%", ha="center", fontsize=8.5, color=pk["mauve"])
            ax.set_ylim(0, tot * 1.1)
            ax.secondary_yaxis("right", functions=(lambda y: y / tot * 100, lambda y: y * tot / 100)).set_ylabel("Kumulatif (%)")
            ttl(ax, f"Pareto {v}"); ax.set_xlabel(v); ax.set_ylabel("Frekuensi"); rot(ax)
            return fig
        bars = ax.bar(tab["Kategori"], tab["n"], color=cols, width=.7)
        for b, n in zip(bars, tab["n"]):
            ax.text(b.get_x() + b.get_width() / 2, n + tab["n"].max() * .01, str(n), ha="center", va="bottom", color=pk["mauve"], fontweight="bold")
        ax.set_ylim(0, tab["n"].max() * 1.1)
        ttl(ax, f"Frekuensi {v}"); ax.set_xlabel(v); ax.set_ylabel("Frekuensi"); rot(ax)
        return fig
    x = d[v].astype(float).values
    bw = v in ("Quantity", "Hour")
    bins = np.arange(np.floor(x.min()) - .5, np.ceil(x.max()) + 1.5, 1) if bw else 25
    fig, ax = plt.subplots(figsize=(7, 3.9))
    if type_ == "Histogram + Density":
        ax.hist(x, bins=bins, density=True, color=pk["xl"], edgecolor="white")
        g, y = kde(x); ax.plot(g, y, color=pk["d"], lw=2)
        ttl(ax, f"Histogram dan Density {v}"); ax.set_xlabel(v); ax.set_ylabel("Densitas")
    elif type_ == "Density":
        g, y = kde(x); ax.fill_between(g, y, color=pk["l"], alpha=.5); ax.plot(g, y, color=pk["d"])
        ax.axvline(x.mean(), ls="--", color=pk["mauve"])
        ttl(ax, f"Density {v}", "Garis putus-putus = rata-rata"); ax.set_xlabel(v); ax.set_ylabel("Densitas")
    elif type_ == "Boxplot":
        ax.boxplot(x, vert=False, widths=.45, patch_artist=True, boxprops=dict(facecolor=pk["xl"], edgecolor=pk["d"]),
                   medianprops=dict(color=pk["d"]), whiskerprops=dict(color=pk["d"]), capprops=dict(color=pk["d"]),
                   flierprops=dict(marker="o", markeredgecolor=pk["m"], markersize=4))
        ax.scatter([x.mean()], [1], marker="D", s=55, facecolor="white", edgecolor=pk["d"], zorder=5)
        ax.set_yticks([]); ax.set_xlabel(v)
        ttl(ax, f"Boxplot {v}", "Berlian = rata-rata; titik = pencilan")
    elif type_ == "QQ-plot":
        (osm, osr), _ = stats.probplot(x, dist="norm")
        ax.scatter(osm, osr, color=pk["m"], alpha=.6, s=14)
        q = np.quantile(x, [.25, .75]); t = stats.norm.ppf([.25, .75]); sl = (q[1] - q[0]) / (t[1] - t[0]); ic = q[0] - sl * t[0]
        ax.plot([osm.min(), osm.max()], [ic + sl * osm.min(), ic + sl * osm.max()], color=pk["d"])
        ttl(ax, f"QQ-plot {v}", "Titik mengikuti garis = mendekati normal"); ax.set_xlabel("Kuantil teoretis"); ax.set_ylabel("Kuantil sampel")
    else:
        ax.hist(x, bins=bins, color=pk["xl"], edgecolor="white")
        ttl(ax, f"Distribusi {v}"); ax.set_xlabel(v); ax.set_ylabel("Frekuensi")
    return fig


def stem_text(x, scale=1.0, width=80, atom=1e-8):
    """Port algoritma stem() milik R (stem.c) agar keluarannya setara."""
    x = np.sort(np.asarray(x, float)); n = len(x)
    if n <= 1:
        return ""
    if x[-1] > x[0]:
        r = atom + (x[-1] - x[0]) / scale
        c = 10.0 ** int(11.0 - int(np.log10(r) + 10))
        mm = min(2, max(0, int(r * c / 25)))
        k = 3 * mm + 2 - 150 // (n + 50)
        if (k - 1) * (k - 2) * (k - 5) == 0:
            c *= 10.0
        x1 = max(abs(x[0]), abs(x[-1]))
        while x1 * c > 2147483647:
            c /= 10
    else:
        r = atom + abs(x[0]) / scale
        c = 10.0 ** int(11.0 - int(np.log10(r) + 10))
        k = 2
    mu = 10
    if k * (k - 4) * (k - 8) == 0: mu = 5
    if (k - 1) * (k - 5) * (k - 6) == 0: mu = 20
    lo = np.floor(x[0] * c / mu) * mu; hi = np.floor(x[-1] * c / mu) * mu
    ldigits = int(np.floor(np.log10(-lo))) + 1 if lo < 0 else 0
    hdigits = int(np.floor(np.log10(hi))) if hi > 0 else 0
    nd = hdigits if ldigits < hdigits else ldigits
    if lo < 0 and np.floor(x[0] * c) == lo: lo -= mu
    hi = lo + mu
    if np.floor(x[0] * c + .5) > hi:
        lo = hi; hi = lo + mu
    pd_ = 1 - int(np.floor(np.log10(c) + .5))
    out = ["", "  The decimal point is " + ("at the |" if pd_ == 0 else f"{abs(pd_)} digit(s) to the {'right' if pd_ > 0 else 'left'} of the |"), ""]

    def stem_print(close, dist):
        lab = "-0" if (int(close / 10) == 0 and dist < 0) else str(int(close / 10))
        return f"  {lab.rjust(nd)} | "
    i = 0
    while True:
        line = stem_print(hi, lo) if lo < 0 else stem_print(lo, hi); j = 0
        while i < n:
            xi = int(x[i] * c - .5) if x[i] < 0 else int(x[i] * c + .5)
            if (hi == 0 and x[i] >= 0) or (lo < 0 and xi > hi) or (lo >= 0 and xi >= hi):
                break
            j += 1
            if j <= width - 12: line += str(abs(xi) % 10)
            i += 1
        if j > width: line += f"+{j - width}"
        out.append(line)
        if i >= n: break
        hi += mu; lo += mu
    return "\n".join(out) + "\n"


def uni_table(d, v):
    if v == "Date":
        nd = d["Date"].nunique()
        return pd.DataFrame({"Statistik": ["Tanggal awal", "Tanggal akhir", "Jumlah hari", "Jumlah hari dengan transaksi", "Rata-rata transaksi per hari"],
                             "Nilai": [d["Date"].min().strftime("%Y-%m-%d"), d["Date"].max().strftime("%Y-%m-%d"),
                                       str((d["Date"].max() - d["Date"].min()).days + 1), str(nd), fmt(len(d) / nd)]})
    if v in cat_all:
        t = d[v].value_counts(sort=False)
        return pd.DataFrame({"Kategori": t.index.astype(str), "Frekuensi": t.values, "Persentase (%)": [fmt(100 * a / t.sum()) for a in t.values]})
    x = d[v].astype(float).values; q1, q3 = np.percentile(x, [25, 75]); iqr = q3 - q1
    lo, hi = q1 - 1.5 * iqr, q3 + 1.5 * iqr
    return pd.DataFrame({"Statistik": ["N", "Rata-rata", "Median", "Simpangan baku", "Variansi", "Koef. variasi (%)", "Minimum", "Q1", "Q3", "Maksimum",
                                       "IQR", "Skewness", "Kurtosis (excess)", "Batas pencilan bawah", "Batas pencilan atas", "Jumlah pencilan"],
                         "Nilai": [str(len(x)), fmt(x.mean()), fmt(np.median(x)), fmt(x.std(ddof=1)), fmt(x.var(ddof=1)), fmt(100 * x.std(ddof=1) / x.mean()),
                                   fmt(x.min()), fmt(q1), fmt(q3), fmt(x.max()), fmt(iqr), fmt(skew_fun(x), 3), fmt(kurt_fun(x), 3),
                                   fmt(lo), fmt(hi), str(int(np.sum((x < lo) | (x > hi))))]})


def interp_uni(d, v):
    if v == "Date":
        dd = d.groupby("Date").size(); top = dd.idxmax()
        return [f"Data mencakup **{d['Date'].min():%Y-%m-%d}** sampai **{d['Date'].max():%Y-%m-%d}** ({len(dd)} hari bertransaksi).",
                f"Rata-rata **{fmt(dd.mean())}** transaksi per hari; puncak pada **{top:%Y-%m-%d}** dengan **{dd.max()}** transaksi.",
                "Garis tren (loess) menunjukkan apakah volume transaksi harian meningkat, menurun, atau relatif stabil sepanjang periode."]
    if v in cat_all:
        t = d[v].value_counts(sort=True); p = 100 * t / t.sum(); sp = p.max() - p.min()
        return [f"Kategori terbanyak: **{t.index[0]}** ({t.iloc[0]} transaksi; {fmt(p.iloc[0], 1)}%). Tersedikit: **{t.index[-1]}** ({t.iloc[-1]}; {fmt(p.iloc[-1], 1)}%).",
                f"Selisih proporsi tertinggi dan terendah **{fmt(sp, 1)}** poin persentase, sehingga sebaran kategori **{'relatif seimbang' if sp < 5 else 'tidak seimbang'}**.",
                "Pareto menunjukkan apakah sedikit kategori mendominasi (garis kumulatif cepat mencapai 80%)."]
    x = d[v].astype(float).values; m, me, s = x.mean(), np.median(x), x.std(ddof=1); sk, ku = skew_fun(x), kurt_fun(x)
    q1, q3 = np.percentile(x, [25, 75]); iqr = q3 - q1; out = int(np.sum((x < q1 - 1.5 * iqr) | (x > q3 + 1.5 * iqr)))
    if abs(sk) < .5: shape = "hampir simetris"
    elif sk > 0: shape = "miring kuat ke kanan" if sk > 1 else "miring sedang ke kanan"
    else: shape = "miring kuat ke kiri" if sk < -1 else "miring sedang ke kiri"
    extra = ("; rata-rata " + ("lebih besar" if m > me else "lebih kecil") + " dari median sehingga sebaran condong ke " + ("kanan" if m > me else "kiri")
             ) if abs(m - me) / s > .1 else "; keduanya berdekatan (sebaran seimbang)"
    kt = "datar (mendekati sebaran seragam)" if ku < -1 else "runcing" if ku > 1 else "mendekati normal"
    return [f"Rata-rata **{fmt(m)}** dan median **{fmt(me)}**{extra}.",
            f"Sebaran data: simpangan baku **{fmt(s)}**, rentang **{fmt(x.min())}** sampai **{fmt(x.max())}**, koefisien variasi **{fmt(100 * s / m)}%**.",
            f"Bentuk distribusi **{shape}** (skewness {fmt(sk, 2)}); kurtosis {fmt(ku, 2)} menunjukkan puncak yang **{kt}**.",
            "Tidak ditemukan pencilan menurut aturan 1.5 x IQR." if out == 0 else
            f"Terdapat **{out} pencilan** menurut aturan 1.5 x IQR ({fmt(100 * out / len(x), 1)}% data). Pencilan dipertahankan karena merupakan transaksi yang valid."]


# =============================================================================
# 3. BIVARIAT
# =============================================================================
def bi_choices(x, y):
    xc, yc = x in cat_all, y in cat_all
    if not xc and not yc:
        return ["Scatter plot + garis tren", "Scatter plot + kurva linear, kuadratik, nonparametrik", "Line plot rata-rata Y per X", "Correlation plot", "Heatmap kepadatan 2D"]
    if xc != yc:
        return ["Grouped boxplot", "Grouped violin plot", "Grouped kernel density plot", "Grouped ridgeline plot", "Grouped strip plot",
                "Grouped mean/SEM plot", "Bar chart statistik ringkasan", "Bar chart rata-rata (+/- SE)", "Cleveland dot chart"]
    return ["Stacked bar chart", "Grouped bar chart", "Segmented bar chart (100%)", "Heatmap tabulasi silang"]


def plot_corr_circle(d, vars_=num_all):
    m = d[vars_].corr().values; n = len(vars_)
    fig, ax = plt.subplots(figsize=(7, 4.4)); ax.grid(False)
    nm = Normalize(-1, 1)
    for i in range(n):
        for j in range(n):
            r = m[i, j]; X, Y = i, n - 1 - j
            if i > j:
                dm = 3 + 19 * np.sqrt(abs(r)); ax.scatter([X], [Y], s=(dm * 2.0) ** 2, color=CM_DIV(nm(r)), edgecolor=pk["grid"], zorder=2)
            elif i < j:
                ax.text(X, Y, f"{r:.2f}", ha="center", va="center", color=pk["txt"], fontweight="bold")
            else:
                ax.text(X, Y, vars_[i], ha="center", va="center", color=pk["d"], fontweight="bold")
    ax.set_xlim(-.6, n - .4); ax.set_ylim(-.6, n - .4)
    ax.set_xticks(range(n)); ax.set_xticklabels(vars_); ax.set_yticks(range(n)); ax.set_yticklabels(vars_[::-1])
    sm_ = plt.cm.ScalarMappable(norm=nm, cmap=CM_DIV); fig.colorbar(sm_, ax=ax, label="r", shrink=.8)
    ttl(ax, "Correlation plot (Pearson)", "Lingkaran: ukuran dan warna = kekuatan korelasi | Angka: nilai r")
    return fig


def plot_bi(d, x, y, type_):
    xc, yc = x in cat_all, y in cat_all
    if not xc and not yc:
        if type_ == "Correlation plot":
            return plot_corr_circle(d, num_all)
        xv, yv = d[x].astype(float).values, d[y].astype(float).values
        fig, ax = plt.subplots(figsize=(7, 4.4))
        if type_ == "Scatter plot + kurva linear, kuadratik, nonparametrik":
            ax.scatter(xv, yv, color=pk["l"], alpha=.5, s=14)
            g = np.linspace(xv.min(), xv.max(), 200)
            if len(np.unique(xv)) >= 2: ax.plot(g, np.polyval(np.polyfit(xv, yv, 1), g), color=pk["d"], lw=2, label="Linear")
            if len(np.unique(xv)) >= 3: ax.plot(g, np.polyval(np.polyfit(xv, yv, 2), g), color=pk["slate"], lw=2, label="Kuadratik")
            loess_line(ax, xv, yv, "#2E7D32", 2, "Nonparametrik (LOESS)")
            ax.legend(title="Kurva")
            ttl(ax, f"{y} terhadap {x}", "Perbandingan kurva linear, kuadratik, dan nonparametrik (LOESS)")
        elif type_ == "Line plot rata-rata Y per X":
            grp = xv if len(np.unique(xv)) <= 12 else pd.cut(xv, bins=10)
            a = pd.DataFrame({"g": grp, "x": xv, "y": yv}).groupby("g", observed=True).agg(gx=("x", "mean"), my=("y", "mean"), se=("y", lambda z: z.std(ddof=1) / np.sqrt(len(z)))).reset_index(drop=True)
            ax.fill_between(a["gx"], a["my"] - a["se"], a["my"] + a["se"], color=pk["xl"], alpha=.8)
            ax.plot(a["gx"], a["my"], color=pk["d"], lw=2); ax.scatter(a["gx"], a["my"], color=pk["m"], s=28, zorder=3)
            ttl(ax, f"Rata-rata {y} menurut {x}", "Pita = rata-rata +/- 1 standard error"); y = f"Rata-rata {y}"
        elif type_ == "Heatmap kepadatan 2D":
            h = ax.hist2d(xv, yv, bins=20, cmap=CM_SEQ); fig.colorbar(h[3], ax=ax, label="Jumlah"); ax.grid(False)
            ttl(ax, f"Kepadatan {x} dan {y}")
        else:
            ax.scatter(xv, yv, color=pk["m"], alpha=.5, s=14); lm_line(ax, xv, yv, pk["d"], 2)
            ttl(ax, f"{y} terhadap {x}", "Scatterplot dengan garis regresi linear sederhana")
        ax.set_xlabel(x); ax.set_ylabel(y)
        return fig
    if xc != yc:
        cv, nv = (x, y) if xc else (y, x); cats = list(d[cv].cat.categories); k = len(cats); cols = pal_cat(k)
        grp = [d.loc[d[cv] == c, nv].astype(float).values for c in cats]
        m = np.array([g.mean() for g in grp]); se = np.array([g.std(ddof=1) / np.sqrt(len(g)) for g in grp])
        md_ = np.array([np.median(g) for g in grp]); sd_ = np.array([g.std(ddof=1) for g in grp])
        mn = np.array([g.min() for g in grp]); mx = np.array([g.max() for g in grp]); pos = np.arange(1, k + 1)
        fig, ax = plt.subplots(figsize=(7, 4.4))
        def cat_axis():
            ax.set_xticks(pos); ax.set_xticklabels(cats); rot(ax)
        if type_ == "Grouped violin plot":
            vp = ax.violinplot(grp, positions=pos, showextrema=False, widths=.8)
            for b, c in zip(vp["bodies"], cols): b.set_facecolor(c); b.set_edgecolor(pk["mauve"]); b.set_alpha(.75)
            ax.boxplot(grp, positions=pos, widths=.12, patch_artist=True, showfliers=False, boxprops=dict(facecolor="white"), medianprops=dict(color="black"))
            cat_axis(); ttl(ax, f"Grouped violin plot: {nv} menurut {cv}"); ax.set_xlabel(cv); ax.set_ylabel(nv)
        elif type_ == "Grouped kernel density plot":
            for c, g, cl in zip(cats, grp, cols):
                gx, gy = kde(g); ax.fill_between(gx, gy, color=cl, alpha=.35); ax.plot(gx, gy, color=cl, label=c)
            ax.legend(title=cv); ttl(ax, f"Grouped kernel density plot: {nv} per {cv}"); ax.set_xlabel(nv); ax.set_ylabel("Densitas")
        elif type_ == "Grouped ridgeline plot":
            allv = d[nv].astype(float).values; grid = np.linspace(allv.min(), allv.max(), 256)
            dens = [kde(g, grid)[1] for g in grp]; top = max(dd.max() for dd in dens)
            for i in range(k - 1, -1, -1):
                h = dens[i] / top * 1.6; ax.fill_between(grid, i + 1, i + 1 + h, color=cols[i], alpha=.8, edgecolor=pk["mauve"])
            ax.set_yticks(pos); ax.set_yticklabels(cats); ttl(ax, f"Grouped ridgeline plot: {nv} per {cv}"); ax.set_xlabel(nv); ax.set_ylabel(cv)
        elif type_ == "Grouped strip plot":
            rng = np.random.default_rng(1)
            for p, g, cl in zip(pos, grp, cols):
                jy = rng.uniform(-.15, .15, len(g)) if nv in ("Quantity", "Hour") else 0
                ax.scatter(p + rng.uniform(-.22, .22, len(g)), g + jy, color=cl, alpha=.45, s=12)
            ax.scatter(pos, m, marker="D", s=70, facecolor="white", edgecolor=pk["d"], zorder=5)
            cat_axis(); ttl(ax, f"Grouped strip plot: {nv} menurut {cv}", "Titik = observasi; berlian = rata-rata"); ax.set_xlabel(cv); ax.set_ylabel(nv)
        elif type_ == "Grouped mean/SEM plot":
            for p, mm, ss, cl in zip(pos, m, se, cols):
                ax.errorbar(p, mm, yerr=ss, fmt="o", color=cl, ms=8, capsize=4, lw=1.8)
            cat_axis(); ttl(ax, f"Grouped mean/SEM plot: {nv} menurut {cv}", "Titik = rata-rata; garis = +/- 1 standard error (SEM)"); ax.set_xlabel(cv); ax.set_ylabel(f"Rata-rata {nv}")
        elif type_ == "Bar chart statistik ringkasan":
            plt.close(fig)
            st_ = {"Rata-rata": m, "Median": md_, "Simpangan baku": sd_, "Minimum": mn, "Maksimum": mx}
            nr, nc = wrap_grid(5); fig, axs = plt.subplots(nr, nc, figsize=(10, 6)); axs = axs.ravel()
            for a_, (nm_, vals) in zip(axs, st_.items()):
                a_.bar(cats, vals, color=cols, width=.7)
                for i, vv in enumerate(vals): a_.text(i, vv + (vals.max() * .01), f"{vv:.1f}", ha="center", va="bottom", fontsize=8)
                a_.set_ylim(0, max(vals.max(), 1e-9) * 1.15); a_.set_title(nm_, fontweight="bold", color=pk["d"], fontsize=10); rot(a_)
            for a_ in axs[5:]: a_.axis("off")
            fig.supxlabel(cv, color=pk["mauve"], fontweight="bold"); fig.supylabel(nv, color=pk["mauve"], fontweight="bold")
            fig_ttl(fig, f"Bar chart statistik ringkasan {nv} menurut {cv}"); fig._top = .93
            return fig
        elif type_ == "Bar chart rata-rata (+/- SE)":
            ax.bar(cats, m, color=cols, width=.7); ax.errorbar(range(k), m, yerr=se, fmt="none", ecolor=pk["txt"], capsize=5)
            for i, mm in enumerate(m): ax.text(i, mm * .94, f"{mm:.2f}", ha="center", va="top", color="white", fontweight="bold")
            rot(ax); ttl(ax, f"Rata-rata {nv} menurut {cv}", "Garis = +/- 1 standard error"); ax.set_xlabel(cv); ax.set_ylabel(f"Rata-rata {nv}")
        elif type_ == "Cleveland dot chart":
            o = np.argsort(m); lv = [cats[i] for i in o]; mm, mdd = m[o], md_[o]
            allv = np.concatenate([mm, mdd]); pad = .15 * max(np.ptp(allv), abs(allv.mean()) * .05); lo, hi = allv.min() - pad, allv.max() + pad
            for i in range(k): ax.plot([lo, hi], [i, i], color=pk["grid"], ls=":", zorder=1)
            ax.scatter(mm, range(k), s=75, facecolor=pk["m"], edgecolor=pk["d"], zorder=3, label="Rata-rata")
            ax.scatter(mdd, range(k), s=75, facecolor="white", edgecolor=pk["d"], zorder=3, label="Median")
            ax.set_yticks(range(k)); ax.set_yticklabels(lv); ax.set_xlim(lo, hi); ax.legend(loc="upper center", bbox_to_anchor=(.5, -.15), ncol=2)
            ttl(ax, f"Cleveland dot chart: {nv} menurut {cv}", "Kategori diurutkan menurut rata-rata"); ax.set_xlabel(nv); ax.set_ylabel(cv)
        else:
            bp = ax.boxplot(grp, positions=pos, patch_artist=True, medianprops=dict(color=pk["mauve"]),
                            flierprops=dict(marker="o", markeredgecolor=pk["m"], markersize=4))
            for b, c in zip(bp["boxes"], cols): b.set_facecolor(c); b.set_edgecolor(pk["mauve"]); b.set_alpha(.85)
            ax.scatter(pos, m, marker="D", s=50, facecolor="white", edgecolor="black", zorder=5)
            cat_axis(); ttl(ax, f"Grouped boxplot: {nv} menurut {cv}", "Berlian = rata-rata"); ax.set_xlabel(cv); ax.set_ylabel(nv)
        return fig
    # ---------- kategorik x kategorik ----------
    tab = pd.crosstab(d[x], d[y]); xs = list(tab.index.astype(str)); ys = list(tab.columns.astype(str)); ny = len(ys); cols = pal_cat(ny)
    fig, ax = plt.subplots(figsize=(7, 4.4))
    if type_ == "Heatmap tabulasi silang":
        ax.grid(False); M = tab.values.T  # baris = Y, kolom = X
        im = ax.imshow(M, cmap=CM_SEQ, origin="lower", aspect="auto"); fig.colorbar(im, ax=ax, label="Jumlah"); mid = M.max() / 2
        for i in range(M.shape[0]):
            for j in range(M.shape[1]):
                ax.text(j, i, str(M[i, j]), ha="center", va="center", color="white" if M[i, j] > mid else pk["txt"], fontweight="bold")
        ax.set_xticks(range(len(xs))); ax.set_xticklabels(xs); ax.set_yticks(range(ny)); ax.set_yticklabels(ys); rot(ax)
        ttl(ax, f"{x} dan {y}"); ax.set_xlabel(x); ax.set_ylabel(y)
        return fig
    n_ = tab.values; pct = 100 * n_ / n_.sum(axis=1, keepdims=True)
    if type_ == "Stacked bar chart":
        bot = np.zeros(len(xs))
        for j in range(ny): ax.bar(xs, n_[:, j], bottom=bot, color=cols[j], edgecolor="white", label=ys[j]); bot += n_[:, j]
        ttl(ax, f"Stacked bar chart: {y} dalam {x}", "Tinggi total batang = jumlah transaksi per kategori"); ax.set_ylabel("Jumlah transaksi")
    elif type_ == "Segmented bar chart (100%)":
        bot = np.zeros(len(xs))
        for j in range(ny): ax.bar(xs, pct[:, j], bottom=bot, color=cols[j], edgecolor="white", label=ys[j]); bot += pct[:, j]
        ax.yaxis.set_major_formatter(lambda v, _: f"{v:.0f}%"); ax.set_ylim(0, 102)
        ttl(ax, f"Segmented bar chart: {y} dalam {x}", "Setiap batang = 100%; bandingkan komposisi antar kategori"); ax.set_ylabel("Persentase")
    else:
        w = .9 / ny; gx = np.arange(len(xs))
        for j in range(ny):
            bars = ax.bar(gx - .45 + w * (j + .5), n_[:, j], width=w, color=cols[j], label=ys[j])
            for b, v_ in zip(bars, n_[:, j]): ax.text(b.get_x() + b.get_width() / 2, v_, str(v_), ha="center", va="bottom", fontsize=8)
        ax.set_xticks(gx); ax.set_xticklabels(xs); ax.set_ylim(0, n_.max() * 1.1)
        ttl(ax, f"Grouped bar chart: {x} dan {y}"); ax.set_ylabel("Jumlah transaksi")
    ax.set_xlabel(x); ax.legend(title=y); rot(ax)
    return fig


def bi_text(d, x, y):
    xc, yc = x in cat_all, y in cat_all
    if not xc and not yc:
        xv, yv = d[x].astype(float).values, d[y].astype(float).values
        r, p = stats.pearsonr(xv, yv); rho, ps = stats.spearmanr(xv, yv); lr = stats.linregress(xv, yv)
        tab = pd.DataFrame({"Ukuran": ["Korelasi Pearson (r)", "p-value Pearson", "Korelasi Spearman (rho)", "p-value Spearman", "Kemiringan garis (b1)", "R-kuadrat regresi sederhana"],
                            "Nilai": [fmt(r, 3), fmt_p(p), fmt(rho, 3), fmt_p(ps), fmt(lr.slope, 3), fmt(lr.rvalue ** 2, 3)]})
        lines = [f"Hubungan linear **{strength(r)}** dan **{'positif' if r > 0 else 'negatif'}** (r = {fmt(r, 3)}; p {'< 0.001' if p < .001 else '= ' + fmt_p(p)}).",
                 f"Hubungan **signifikan** pada taraf 5%: setiap kenaikan 1 satuan **{x}**, rata-rata **{y}** berubah sebesar **{fmt(lr.slope, 3)}** satuan." if p < .05
                 else "Hubungan **tidak signifikan** pada taraf 5%, sehingga tidak ada bukti hubungan linear antar kedua variabel.",
                 "Korelasi menunjukkan keterkaitan, bukan sebab-akibat."]
        return tab, lines
    if xc != yc:
        cv, nv = (x, y) if xc else (y, x); cats = list(d[cv].cat.categories)
        grp = [d.loc[d[cv] == c, nv].astype(float).values for c in cats]
        g = pd.DataFrame({cv: cats, "Rata-rata": [fmt(a.mean()) for a in grp], "Median": [fmt(np.median(a)) for a in grp],
                          "Simpangan baku": [fmt(a.std(ddof=1)) for a in grp], "n": [len(a) for a in grp]})
        gm = np.array([a.mean() for a in grp])
        F, p = stats.f_oneway(*grp); allv = np.concatenate(grp); ssb = sum(len(a) * (a.mean() - allv.mean()) ** 2 for a in grp); eta = ssb / np.sum((allv - allv.mean()) ** 2)
        kw = stats.kruskal(*grp).pvalue
        if len(cats) == 2:
            tt = stats.ttest_ind(grp[0], grp[1], equal_var=False); p = tt.pvalue; test = f"uji t Welch (t = {fmt(tt.statistic, 2)})"
        else:
            test = f"ANOVA (F = {fmt(F, 2)})"
        lines = [f"Rata-rata **{nv}** tertinggi pada **{cats[int(np.argmax(gm))]}** ({fmt(gm.max())}) dan terendah pada **{cats[int(np.argmin(gm))]}** ({fmt(gm.min())}).",
                 f"Hasil {test}: p {pq(p)}; uji Kruskal-Wallis: p {pq(kw)}; ukuran efek eta-kuadrat = {fmt(eta, 3)}.",
                 "Perbedaan rata-rata antar kelompok **signifikan** pada taraf 5%." if p < .05 else
                 "Perbedaan rata-rata antar kelompok **tidak signifikan** pada taraf 5%; selisih yang tampak pada grafik dapat terjadi karena kebetulan."]
        return g, lines
    tab = pd.crosstab(d[x], d[y]); chi2, p, dof, _ = stats.chi2_contingency(tab.values)
    V = np.sqrt(chi2 / (tab.values.sum() * (min(tab.shape) - 1)))
    cj, ri = np.argwhere(tab.values.T == tab.values.max())[0]
    out = tab.copy(); out.columns = out.columns.astype(str); out.insert(0, x, tab.index.astype(str)); out = out.reset_index(drop=True)
    lines = [f"Kombinasi terbanyak: **{tab.index[ri]}** x **{tab.columns[cj]}** ({tab.values.max()} transaksi).",
             f"Uji Chi-square: X2 = {fmt(chi2, 2)}, df = {dof}, p {pq(p)}; Cramer's V = {fmt(V, 3)}.",
             "Ada hubungan **signifikan** antar kedua variabel kategorik pada taraf 5%." if p < .05 else
             "**Tidak ada** hubungan signifikan antar kedua variabel kategorik pada taraf 5%."]
    return out, lines


# =============================================================================
# 4. MULTIVARIAT
# =============================================================================
def plot_mv_scatter(d, x, y, col, size, facet, lines):
    hasc, hass = col != NONE, size != NONE
    catc = hasc and col in cat_all
    panels = [(None, d)] if facet == NONE else [(g, d[d[facet] == g]) for g in d[facet].cat.categories]
    if facet == NONE:
        fig, axs = plt.subplots(figsize=(8, 4.8)); axs = [axs]
    else:
        nr, nc = wrap_grid(len(panels)); fig, axs = plt.subplots(nr, nc, figsize=(9, 5.4), sharex=True, sharey=True); axs = list(np.atleast_1d(axs).ravel())
    if catc:
        cats = list(d[col].cat.categories); cmap = dict(zip(cats, pal_cat(len(cats))))
    if hasc and not catc:
        nm = Normalize(d[col].min(), d[col].max())
    if hass:
        smin, smax = d[size].min(), d[size].max()
        def sz(v):
            t = (v - smin) / (smax - smin) if smax > smin else 0 * v
            return (1.9 * (1 + 4 * np.sqrt(t))) ** 2
    for ax, (g, dd) in zip(axs, panels):
        xv, yv = dd[x].astype(float).values, dd[y].astype(float).values
        s = sz(dd[size].astype(float).values) if hass else (18 if hasc else 18)
        if catc:
            c = [cmap[v] for v in dd[col].astype(str)]
        elif hasc:
            c = CM_PTS(nm(dd[col].astype(float).values))
        else:
            c = pk["m"]
        ax.scatter(xv, yv, s=s, c=c, alpha=.6 if (hasc or hass) else .6)
        if lines:
            if catc:
                for k_, cl in cmap.items():
                    m_ = (dd[col].astype(str) == k_).values
                    if m_.sum() > 2: lm_line(ax, xv[m_], yv[m_], cl, 1.6)
            else:
                lm_line(ax, xv, yv, pk["d"], 1.6)
        if g is not None: ax.set_title(str(g), fontweight="bold", color=pk["d"], fontsize=10)
    for ax in axs[len(panels):]: ax.axis("off")
    if facet == NONE:
        axs[0].set_xlabel(x); axs[0].set_ylabel(y)
    else:
        fig.supxlabel(x, color=pk["mauve"], fontweight="bold"); fig.supylabel(y, color=pk["mauve"], fontweight="bold")
    bottom = 0
    if catc:
        hd = [Line2D([], [], marker="o", ls="", color=c_, label=k_) for k_, c_ in cmap.items()]
        if facet == NONE:
            axs[0].legend(handles=hd, title=col, loc="center left", bbox_to_anchor=(1.01, .5))
        else:
            fig.legend(handles=hd, title=col, loc="lower left", ncol=len(hd), fontsize=8); bottom = .1
    right = 1
    if hasc and not catc:
        cax = fig.add_axes([.93, .25, .015, .5]); fig.colorbar(plt.cm.ScalarMappable(norm=nm, cmap=CM_PTS), cax=cax, label=col); right = .9
    if hass:
        qs = np.unique(np.round(np.quantile(d[size].astype(float), [0, .5, 1]), 1))
        hd = [axs[0].scatter([], [], s=sz(np.array(q)), color=pk["m"], alpha=.5, label=str(q)) for q in qs]
        fig.legend(handles=hd, title=size, loc="lower right", ncol=len(hd), fontsize=8); bottom = .1
    fig_ttl(fig, f"{y} terhadap {x}", f"Warna: {col} | Ukuran: {size} | Panel: {facet}")
    fig._rect = [0, bottom, right, .9]
    return fig


def mv_heat_data(d, a, b, num, fun):
    g = d.groupby([a, b], observed=True)
    if fun == "Banyak transaksi":
        res = g.size()
    else:
        res = g[num].agg({"Rata-rata": "mean", "Median": "median", "Jumlah": "sum"}[fun])
    res = res.rename("val").reset_index(); res.columns = ["A", "B", "val"]
    return res


def plot_mv_heat(d, a, b, num, fun):
    res = mv_heat_data(d, a, b, num, fun); mid = (res["val"].max() + res["val"].min()) / 2
    lab = "Banyak transaksi" if fun == "Banyak transaksi" else f"{fun} {num}"
    xs, ys = list(d[a].cat.categories.astype(str)), list(d[b].cat.categories.astype(str))
    M = np.full((len(ys), len(xs)), np.nan)
    for _, r in res.iterrows(): M[ys.index(str(r["B"])), xs.index(str(r["A"]))] = r["val"]
    fig, ax = plt.subplots(figsize=(8, 4.8)); ax.grid(False)
    im = ax.imshow(np.ma.masked_invalid(M), cmap=CM_SEQ, origin="lower", aspect="auto"); fig.colorbar(im, ax=ax, label=lab)
    kd = 0 if fun == "Banyak transaksi" else 1
    for i in range(M.shape[0]):
        for j in range(M.shape[1]):
            if not np.isnan(M[i, j]): ax.text(j, i, fmt(M[i, j], kd), ha="center", va="center", color="white" if M[i, j] > mid else pk["txt"], fontweight="bold")
    ax.set_xticks(range(len(xs))); ax.set_xticklabels(xs); ax.set_yticks(range(len(ys))); ax.set_yticklabels(ys); rot(ax)
    ttl(ax, lab, f"menurut {a} dan {b}"); ax.set_xlabel(a); ax.set_ylabel(b)
    return fig


def plot_mv_cor(d, vars_, method, title=None):
    m = d[list(vars_)].corr(method=method).values; n = len(vars_)
    fig, ax = plt.subplots(figsize=(7, 4.6)); ax.grid(False)
    im = ax.imshow(m, cmap=CM_DIV, vmin=-1, vmax=1, origin="lower", aspect="auto"); fig.colorbar(im, ax=ax, label="r")
    for i in range(n):
        for j in range(n): ax.text(j, i, f"{m[i, j]:.2f}", ha="center", va="center", color=pk["txt"], fontweight="bold", fontsize=11)
    ax.set_xticks(range(n)); ax.set_xticklabels(vars_); ax.set_yticks(range(n)); ax.set_yticklabels(vars_)
    ttl(ax, title or f"Matriks korelasi {method.capitalize()}")
    return fig


def trend_data(d, metric, unit, grp):
    d = d.copy()
    if unit == "Harian": d["Periode"] = d["Date"]
    elif unit == "Mingguan": d["Periode"] = d["Date"] - pd.to_timedelta(d["Date"].dt.weekday, unit="D")
    else: d["Periode"] = d["Date"].dt.to_period("M").dt.to_timestamp()
    gl = ["Periode"] + ([] if grp == NONE else [grp]); g = d.groupby(gl, observed=True)
    if metric == "Jumlah transaksi":
        val = g.size()
    else:
        col, fn = {"Total penjualan (jumlah)": ("Total", "sum"), "Rata-rata Total": ("Total", "mean"), "Rata-rata Rating": ("Rating", "mean"), "Total Quantity": ("Quantity", "sum")}[metric]
        val = g[col].agg(fn)
    return val.rename("val").reset_index()


def plot_mv_trend(d, metric, unit, grp):
    td = trend_data(d, metric, unit, grp); fig, ax = plt.subplots(figsize=(8, 4.6))
    if grp == NONE:
        ax.plot(td["Periode"], td["val"], color=pk["d"], lw=2); ax.scatter(td["Periode"], td["val"], color=pk["m"], s=22, zorder=3)
    else:
        cats = list(d[grp].cat.categories)
        for c, cl in zip(cats, pal_cat(len(cats))):
            t = td[td[grp] == c]; ax.plot(t["Periode"], t["val"], color=cl, lw=2, label=c); ax.scatter(t["Periode"], t["val"], color=cl, s=18)
        ax.legend(title=grp)
    ttl(ax, f"{metric} {unit} " + (f"menurut {grp}" if grp != NONE else ""), "Minggu pertama/terakhir dapat tidak lengkap" if unit == "Mingguan" else None)
    ax.set_xlabel("Periode"); ax.set_ylabel(metric)
    return fig


def pca_fit(d, vars_):
    X = d[list(vars_)].astype(float).values; Z = (X - X.mean(0)) / X.std(0, ddof=1)
    U, s, Vt = np.linalg.svd(Z, full_matrices=False)
    return dict(x=Z @ Vt.T, sdev=s / np.sqrt(len(Z) - 1), rotation=pd.DataFrame(Vt.T, index=list(vars_), columns=[f"PC{i + 1}" for i in range(len(s))]))


def plot_mv_pca(d, vars_, colv):
    pr = pca_fit(d, vars_); sc = pr["x"][:, :2]; ve = 100 * pr["sdev"] ** 2 / np.sum(pr["sdev"] ** 2)
    ld = pr["rotation"].iloc[:, :2].values; k = min(np.abs(sc[:, 0]).max(), np.abs(sc[:, 1]).max()) * .85; ld = ld * k
    fig, ax = plt.subplots(figsize=(7.5, 4.8))
    if colv != NONE:
        cats = list(d[colv].cat.categories)
        for c, cl in zip(cats, pal_cat(len(cats))):
            m_ = (d[colv] == c).values; ax.scatter(sc[m_, 0], sc[m_, 1], color=cl, alpha=.6, s=16, label=c)
        ax.legend(title=colv)
    else:
        ax.scatter(sc[:, 0], sc[:, 1], color=pk["m"], alpha=.5, s=16)
    for (a, b), nm_ in zip(ld, vars_):
        ax.annotate("", xy=(a, b), xytext=(0, 0), arrowprops=dict(arrowstyle="->", color=pk["d"], lw=1.4))
        ax.text(a, b, nm_, color=pk["d"], fontweight="bold", ha="center", va="bottom")
    ttl(ax, "PCA biplot"); ax.set_xlabel(f"PC1 ({ve[0]:.1f}%)"); ax.set_ylabel(f"PC2 ({ve[1]:.1f}%)")
    return fig


def plot_mv_pairs(d, vars_, colv):
    vars_ = list(vars_); k = len(vars_); grp = d[colv] if colv != NONE else None
    cols = pk["m"] if grp is None else np.array(pal_cat(len(grp.cat.categories)))[grp.cat.codes.values]
    fig, axs = plt.subplots(k, k, figsize=(8.5, 7.8))
    for i in range(k):
        for j in range(k):
            ax = axs[i, j]; ax.grid(False); ax.set_xticks([]); ax.set_yticks([])
            for sp in ax.spines.values(): sp.set_visible(True); sp.set_color(pk["grid"])
            xv, yv = d[vars_[j]].astype(float).values, d[vars_[i]].astype(float).values
            if i == j:
                ax.hist(xv, bins="auto", color=pk["xl"], edgecolor="white"); ax.text(.5, .88, vars_[i], transform=ax.transAxes, ha="center", va="top", fontweight="bold", fontsize=12)
            elif i > j:
                ax.scatter(xv, yv, s=6, color=cols, alpha=.55); loess = lowess(yv, xv, frac=.6, return_sorted=True) if len(np.unique(xv)) > 3 else None
                if loess is not None: ax.plot(loess[:, 0], loess[:, 1], color=pk["d"], lw=1.6)
            else:
                r = np.corrcoef(xv, yv)[0, 1]; ax.text(.5, .5, f"r = {r:.2f}", transform=ax.transAxes, ha="center", va="center", color=pk["d"], fontweight="bold", fontsize=10 + 12 * abs(r))
    fig_ttl(fig, "Scatterplot matrix (pairs plot)")
    if grp is not None:
        cats = list(grp.cat.categories)
        fig.legend(handles=[Line2D([], [], marker="o", ls="", color=c, label=l) for l, c in zip(cats, pal_cat(len(cats)))], title=colv, loc="lower center", ncol=len(cats), fontsize=8)
        fig._rect = [0, .06, 1, .92]
    else:
        fig._top = .93
    return fig


def mahal_fit(d, vars_):
    X = d[list(vars_)].astype(float).values; p = X.shape[1]
    try:
        S = np.linalg.inv(np.cov(X, rowvar=False)); Z = X - X.mean(0); D2 = np.einsum("ij,jk,ik->i", Z, S, Z)
    except np.linalg.LinAlgError:
        raise Need("Matriks kovarians singular (ada variabel yang bergantung linear). Kurangi variabel yang dipilih.")
    cut = stats.chi2.ppf(.975, df=p)
    return dict(df=pd.DataFrame({"idx": np.arange(1, len(X) + 1), "D2": D2, "pval": stats.chi2.sf(D2, p), "out": D2 > cut}), cut=cut, p=p, X=X)


def _status(out):
    return np.where(out, "Pencilan", "Normal")


def plot_mv_mahal_idx(f):
    df = f["df"]; fig, ax = plt.subplots(figsize=(6.2, 3.9))
    for stt, cl in (("Normal", pk["l"]), ("Pencilan", pk["d"])):
        m_ = _status(df["out"]) == stt; ax.scatter(df["idx"][m_], df["D2"][m_], color=cl, alpha=.7, s=16, label=stt)
    ax.axhline(f["cut"], ls="--", color=pk["d"])
    for _, r in df[df["out"]].sort_values("D2", ascending=False).head(5).iterrows(): ax.text(r["idx"], r["D2"], f" {int(r['idx'])}", fontsize=8, va="bottom")
    ax.legend(title="Status")
    ttl(ax, "Jarak Mahalanobis (D2) per observasi", f"Garis putus-putus = batas Chi-Square 97.5% (df = {f['p']}): {f['cut']:.2f} | Pencilan: {int(df['out'].sum())} ({100 * df['out'].mean():.1f}%)")
    ax.set_xlabel("Indeks observasi"); ax.set_ylabel("D2 (kuadrat jarak Mahalanobis)")
    return fig


def plot_mv_mahal_sc(f, vars_):
    st_ = _status(f["df"]["out"]); fig, ax = plt.subplots(figsize=(6.2, 3.9))
    for s_, cl, sz in (("Normal", pk["l"], 16), ("Pencilan", pk["d"], 28)):
        m_ = st_ == s_; ax.scatter(f["X"][m_, 0], f["X"][m_, 1], color=cl, s=sz, alpha=.65, label=s_)
    ax.legend(title="Status")
    ttl(ax, "Pencilan multivariat pada ruang dua variabel", f"Ditampilkan: {vars_[0]} dan {vars_[1]} (dua variabel pertama yang dipilih)")
    ax.set_xlabel(vars_[0]); ax.set_ylabel(vars_[1])
    return fig


def mahal_table(d, f, vars_):
    df = f["df"]; o = df.sort_values("D2", ascending=False).head(10)
    out = pd.DataFrame({"Baris": o["idx"].values, "D2": [fmt(v) for v in o["D2"]], "p-value": [fmt_p(v) for v in o["pval"]], "Status": _status(o["out"].values)})
    for v in vars_: out[v] = [fmt(x, 1) for x in d[v].astype(float).values[o["idx"].values - 1]]
    return out


def plot_mv_qqchi(f):
    n = len(f["df"]); o = np.argsort(f["df"]["D2"].values); D2 = f["df"]["D2"].values[o]; idx = f["df"]["idx"].values[o]; out = f["df"]["out"].values[o]
    q = stats.chi2.ppf((np.arange(1, n + 1) - .5) / n, df=f["p"]); st_ = _status(out)
    fig, ax = plt.subplots(figsize=(7.5, 4.6))
    for s_, cl in (("Normal", pk["m"]), ("Pencilan", pk["d"])):
        m_ = st_ == s_; ax.scatter(q[m_], D2[m_], color=cl, alpha=.7, s=18, label=s_)
    lim = [min(q.min(), D2.min()), max(q.max(), D2.max())]; ax.plot(lim, lim, color=pk["d"], lw=1.6)
    for i in range(n - 3, n): ax.text(q[i], D2[i], f" {int(idx[i])}", fontsize=8, va="bottom")
    ax.legend(title="Status")
    ttl(ax, "Q-Q plot Chi-Square", f"D2 terurut vs kuantil Chi-Square (df = {f['p']}); titik di sekitar garis = mendekati normal multivariat")
    ax.set_xlabel("Kuantil teoretis Chi-Square"); ax.set_ylabel("D2 terurut (sampel)")
    return fig


# =============================================================================
# 5. REGRESI LINEAR BERGANDA
# =============================================================================
def build_model(d, resp, preds, logy):
    d = d.reset_index(drop=True); cols, term_cols, collabel = {}, {}, {}
    for v in preds:
        if v in num_all:
            cols[v] = d[v].astype(float); term_cols[v] = [v]; collabel[v] = v
        else:
            lv = list(d[v].cat.categories); need(len(lv) >= 2, f"Variabel {v} hanya memiliki satu kategori pada data terfilter; hapus dari prediktor atau longgarkan filter.")
            names = []
            for l in lv[1:]:
                nm = f"{v}{l}"; cols[nm] = (d[v] == l).astype(float); names.append(nm); collabel[nm] = f"{v} = {l}"
            term_cols[v] = names
    X = pd.DataFrame(cols); X.insert(0, "(Intercept)", 1.0)
    y = np.log(d[resp].astype(float)) if logy else d[resp].astype(float)
    res = sm.OLS(y.values, X.values).fit()
    colterm = {c: t for t, cs in term_cols.items() for c in cs}
    return dict(res=res, X=X, y=y.values, term_cols=term_cols, colterm=colterm, collabel=collabel, preds=list(preds), resp=resp, logy=logy, d=d,
                names=list(X.columns), xlevels={v: list(d[v].cat.categories) for v in preds if v not in num_all})


def coef_label(M, cn):
    return "Intersep" if cn == "(Intercept)" else M["collabel"][cn]


def make_equation(M):
    b = M["res"].params; nm = M["names"]; lhs = f"ln({M['resp']})" if M["logy"] else M["resp"]
    ln = [f"{lhs}-hat = {fmt(b[0], 4)}"]
    for i in range(1, len(b)):
        ln.append(f"   {'+' if b[i] >= 0 else '-'} {fmt(abs(b[i]), 4)} x [{M['collabel'][nm[i]]}]")
    return "\n".join(ln)


def vif_table(M):
    try:
        Xm = M["X"].drop(columns="(Intercept)"); terms = [t for t, cs in M["term_cols"].items() if cs]
        if len(terms) < 2: return None
        R = np.corrcoef(Xm.values, rowvar=False); R = np.atleast_2d(R); dR = np.linalg.det(R); rows = []
        for t in terms:
            i = [list(Xm.columns).index(c) for c in M["term_cols"][t]]; rest = [j for j in range(R.shape[0]) if j not in i]
            gv = np.linalg.det(R[np.ix_(i, i)]) * (np.linalg.det(R[np.ix_(rest, rest)]) if rest else 1.0) / dR
            rows.append((t, len(i), gv ** (1 / len(i))))
        return pd.DataFrame(rows, columns=["Variabel", "Df", "VIF setara"])
    except Exception:
        return None


def bp_test(M):
    e = M["res"].resid; u = e ** 2 / (np.sum(e ** 2) / len(e)); X = M["X"].values
    fit = X @ np.linalg.lstsq(X, u, rcond=None)[0]; stat = np.sum((fit - u.mean()) ** 2) / 2; df = X.shape[1] - 1
    return dict(stat=stat, df=df, p=stats.chi2.sf(stat, df))


def reg_df(M):
    r = M["res"]; inf = r.get_influence()
    return pd.DataFrame(dict(fit=r.fittedvalues, act=M["y"], res=r.resid, rs=inf.resid_studentized_internal, cook=inf.cooks_distance[0], idx=np.arange(1, len(r.resid) + 1)))


def fit_table(M):
    r = M["res"]; e = r.resid; k = len(r.params); n = int(r.nobs)
    aic = -2 * r.llf + 2 * (k + 1); bic = -2 * r.llf + np.log(n) * (k + 1)
    return pd.DataFrame({"Ukuran": ["Jumlah observasi (n)", "Derajat bebas residual", "R-kuadrat", "R-kuadrat terkoreksi (Adj. R2)", "Galat baku residual (sigma)", "RMSE", "MAE",
                                    "Statistik F", "p-value (uji F simultan)", "AIC", "BIC"],
                         "Nilai": [str(n), str(int(r.df_resid)), fmt(r.rsquared, 4), fmt(r.rsquared_adj, 4), fmt(np.sqrt(r.mse_resid), 3), fmt(np.sqrt(np.mean(e ** 2)), 3),
                                   fmt(np.mean(np.abs(e)), 3), f"{fmt(r.fvalue, 2)} (df1 = {int(r.df_model)}, df2 = {int(r.df_resid)})", fmt_p(r.f_pvalue), fmt(aic, 2), fmt(bic, 2)]})


def coef_table(M):
    r = M["res"]; ci = r.conf_int()
    return pd.DataFrame({"Koefisien": [coef_label(M, n) for n in M["names"]], "Estimasi": [fmt(v, 4) for v in r.params], "Galat baku": [fmt(v, 4) for v in r.bse],
                         "t": [fmt(v, 3) for v in r.tvalues], "p-value": [fmt_p(v) for v in r.pvalues], "CI 2.5%": [fmt(v, 3) for v in ci[:, 0]],
                         "CI 97.5%": [fmt(v, 3) for v in ci[:, 1]], "Sig.": [stars(v) for v in r.pvalues]})


def anova_table(M):
    """ANOVA sekuensial (Type I), padanan anova() pada R."""
    X, y = M["X"], M["y"]; r = M["res"]; rss_prev = np.sum((y - y.mean()) ** 2); cur = ["(Intercept)"]; rows = []
    mse = r.ssr / r.df_resid
    for t, cs in M["term_cols"].items():
        if not cs: continue
        cur = cur + cs; rss = sm.OLS(y, X[cur].values).fit().ssr; ss = rss_prev - rss; df = len(cs); F = (ss / df) / mse
        rows.append([t, df, fmt(ss, 1), fmt(ss / df, 1), fmt(F, 3), fmt_p(stats.f.sf(F, df, r.df_resid))]); rss_prev = rss
    rows.append(["Residuals", int(r.df_resid), fmt(r.ssr, 1), fmt(mse, 1), "", ""])
    return pd.DataFrame(rows, columns=["Variabel", "Df", "Jumlah kuadrat", "Kuadrat tengah", "F", "p-value"])


def plot_reg_avp(M):
    df = reg_df(M); r2 = M["res"].rsquared; sc = f"ln({M['resp']})" if M["logy"] else M["resp"]; fig, ax = plt.subplots(figsize=(6.2, 3.9))
    ax.scatter(df["fit"], df["act"], color=pk["m"], alpha=.45, s=16); lim = [min(df["fit"].min(), df["act"].min()), max(df["fit"].max(), df["act"].max())]
    ax.plot(lim, lim, ls="--", color=pk["d"], lw=1.6)
    ttl(ax, "Nilai aktual vs prediksi", f"Garis putus-putus = prediksi sempurna | R2 = {r2:.3f}"); ax.set_xlabel(f"Prediksi {sc}"); ax.set_ylabel(f"Aktual {sc}")
    return fig


def plot_reg_coef(M):
    r = M["res"]; ci = r.conf_int(); nm = M["names"]; keep = [i for i, n in enumerate(nm) if n != "(Intercept)"]
    est, lo, hi, p = r.params[keep], ci[keep, 0], ci[keep, 1], r.pvalues[keep]; lab = [M["collabel"][nm[i]] for i in keep]; n = len(keep)
    fig, ax = plt.subplots(figsize=(6.2, max(3.9, .3 * n + 1.8))); ypos = np.arange(n)[::-1]
    ax.axvline(0, ls="--", color=pk["mauve"])
    for yp, e, l_, h_, pp in zip(ypos, est, lo, hi, p):
        cl = pk["d"] if pp < .05 else "#B39DAB"; ax.errorbar(e, yp, xerr=[[e - l_], [h_ - e]], fmt="o", color=cl, capsize=3, ms=6, lw=1.6)
    ax.set_yticks(ypos); ax.set_yticklabels(lab)
    ax.legend(handles=[Line2D([], [], marker="o", ls="", color=pk["d"], label="Signifikan (p < 0.05)"), Line2D([], [], marker="o", ls="", color="#B39DAB", label="Tidak signifikan")],
              loc="upper center", bbox_to_anchor=(.5, -.12), ncol=2)
    ttl(ax, "Estimasi koefisien dan selang kepercayaan 95%", "Selang yang memotong garis nol berarti tidak signifikan"); ax.set_xlabel("Estimasi koefisien")
    return fig


def plot_reg_resid(M):
    df = reg_df(M); fig, ax = plt.subplots(figsize=(6.2, 3.6)); ax.scatter(df["fit"], df["res"], color=pk["m"], alpha=.45, s=14); ax.axhline(0, ls="--", color=pk["mauve"])
    loess_line(ax, df["fit"], df["res"], pk["d"], 1.6); ttl(ax, "Residual vs nilai prediksi", "Sebaran acak di sekitar nol = asumsi terpenuhi"); ax.set_xlabel("Nilai prediksi"); ax.set_ylabel("Residual")
    return fig


def plot_reg_qq(M):
    df = reg_df(M); x = df["rs"].values; fig, ax = plt.subplots(figsize=(6.2, 3.6)); (osm, osr), _ = stats.probplot(x, dist="norm")
    ax.scatter(osm, osr, color=pk["m"], alpha=.6, s=14); q = np.quantile(x, [.25, .75]); t = stats.norm.ppf([.25, .75]); sl = (q[1] - q[0]) / (t[1] - t[0]); ic = q[0] - sl * t[0]
    ax.plot([osm.min(), osm.max()], [ic + sl * osm.min(), ic + sl * osm.max()], color=pk["d"])
    ttl(ax, "Normal probability plot (QQ-plot residual terstandar)", "Titik mengikuti garis = residual mendekati normal"); ax.set_xlabel("Kuantil teoretis"); ax.set_ylabel("Kuantil sampel")
    return fig


def plot_reg_scale(M):
    df = reg_df(M); y = np.sqrt(np.abs(df["rs"].values)); fig, ax = plt.subplots(figsize=(6.2, 3.6)); ax.scatter(df["fit"], y, color=pk["m"], alpha=.45, s=14)
    loess_line(ax, df["fit"], y, pk["d"], 1.6); ttl(ax, "Scale-Location", "Garis datar = ragam residual konstan (homoskedastis)"); ax.set_xlabel("Nilai prediksi"); ax.set_ylabel("Akar |residual terstandar|")
    return fig


def plot_reg_cook(M):
    df = reg_df(M); thr = 4 / len(df); top = df.sort_values("cook", ascending=False).head(3); fig, ax = plt.subplots(figsize=(6.2, 3.6))
    ax.vlines(df["idx"], 0, df["cook"], color=pk["l"]); ax.axhline(thr, ls="--", color=pk["d"])
    for _, r in top.iterrows(): ax.text(r["idx"], r["cook"], f"{int(r['idx'])}", color=pk["d"], fontsize=8, ha="center", va="bottom")
    ttl(ax, "Cook's distance", f"Garis putus-putus = 4/n = {thr:.4f}"); ax.set_xlabel("Indeks observasi"); ax.set_ylabel("Cook's distance")
    return fig


# ---- Visualisasi regresi tambahan ----
def reg_ord_res(M):
    d = M["d"]; o = np.argsort(d["Date"].values, kind="stable")
    return dict(res=np.asarray(M["res"].resid)[o], date=d["Date"].values[o])


def plot_reg_scatter(M):
    d, resp, preds, logy = M["d"], M["resp"], M["preds"], M["logy"]; nums = [p for p in preds if p in num_all]
    need(len(nums) >= 1, "Pilih minimal satu prediktor numerik untuk scatterplot.")
    yv = np.log(d[resp].astype(float).values) if logy else d[resp].astype(float).values
    nr, nc = wrap_grid(len(nums)); fig, axs = plt.subplots(nr, nc, figsize=(6.4, 3.6)); axs = list(np.atleast_1d(axs).ravel())
    for ax, v in zip(axs, nums):
        xv = d[v].astype(float).values; ax.scatter(xv, yv, color=pk["m"], alpha=.35, s=8)
        if len(np.unique(xv)) >= 2:
            o = sm.OLS(yv, sm.add_constant(xv)).fit(); g = np.linspace(xv.min(), xv.max(), 100); pr = o.get_prediction(sm.add_constant(g)).conf_int()
            ax.fill_between(g, pr[:, 0], pr[:, 1], color=pk["xl"], alpha=.7); ax.plot(g, o.predict(sm.add_constant(g)), color=pk["d"], lw=1.6)
        ax.set_title(v, fontweight="bold", color=pk["d"], fontsize=10)
    for ax in axs[len(nums):]: ax.axis("off")
    fig.supxlabel("Nilai prediktor", color=pk["mauve"], fontweight="bold", fontsize=9); fig.supylabel(f"ln({resp})" if logy else resp, color=pk["mauve"], fontweight="bold", fontsize=9)
    fig_ttl(fig, "Scatterplot respons terhadap prediktor numerik", "Garis = regresi linear sederhana dengan pita kepercayaan 95%"); fig._top = .86
    return fig


def plot_reg_cor(M):
    vars_ = list(dict.fromkeys([M["resp"]] + [p for p in M["preds"] if p in num_all])); need(len(vars_) >= 2, "Pilih minimal satu prediktor numerik untuk matriks korelasi.")
    return plot_mv_cor(M["d"], vars_, "pearson", "Correlation matrix (respons dan prediktor numerik)")


def plot_reg_hist(M):
    e = reg_df(M)["res"].values; fig, ax = plt.subplots(figsize=(6.2, 3.6)); ax.hist(e, bins=30, density=True, color=pk["xl"], edgecolor="white")
    g, y = kde(e); ax.plot(g, y, color=pk["d"], lw=2); ax.plot(g, stats.norm.pdf(g, e.mean(), e.std(ddof=1)), ls="--", color=pk["slate"], lw=1.6)
    ttl(ax, "Histogram residual", "Garis tebal = density residual; putus-putus = kurva normal teoretis"); ax.set_xlabel("Residual"); ax.set_ylabel("Densitas")
    return fig


def plot_reg_ts(M):
    o = reg_ord_res(M); n = len(o["res"]); t = np.arange(1, n + 1); br = np.unique(np.round(np.linspace(1, n, 6)).astype(int)); fig, ax = plt.subplots(figsize=(6.2, 3.6))
    ax.plot(t, o["res"], color=pk["l"], lw=.5); ax.scatter(t, o["res"], color=pk["m"], s=5, alpha=.5); ax.axhline(0, ls="--", color=pk["mauve"]); loess_line(ax, t, o["res"], pk["d"], 1.6)
    ax.set_xticks(br); ax.set_xticklabels([pd.Timestamp(o["date"][b - 1]).strftime("%d-%m-%Y") for b in br], fontsize=8)
    ttl(ax, "Time series plot residual", "Observasi diurutkan menurut tanggal transaksi; garis tebal = tren (loess)"); ax.set_xlabel("Tanggal transaksi (urutan observasi)"); ax.set_ylabel("Residual")
    return fig


def plot_reg_acf(M):
    o = reg_ord_res(M); n = len(o["res"]); L = min(30, n - 1); a = sm_acf(o["res"], nlags=L, fft=False)[1:]; lag = np.arange(1, L + 1); ci = stats.norm.ppf(.975) / np.sqrt(n)
    fig, ax = plt.subplots(figsize=(6.2, 3.6)); ax.axhline(0, color=pk["mauve"]); ax.vlines(lag, 0, a, color=pk["m"], lw=2); ax.scatter(lag, a, color=pk["d"], s=14, zorder=3)
    for c in (-ci, ci): ax.axhline(c, ls="--", color=pk["slate"])
    ttl(ax, "ACF plot residual", "Garis putus-putus = batas 95%; lag 0 tidak ditampilkan"); ax.set_xlabel("Lag"); ax.set_ylabel("Autokorelasi")
    return fig


def plot_reg_obsres(M):
    df = reg_df(M); fig, ax = plt.subplots(figsize=(6.2, 3.6)); ax.scatter(df["act"], df["res"], color=pk["m"], alpha=.45, s=14); ax.axhline(0, ls="--", color=pk["mauve"])
    loess_line(ax, df["act"], df["res"], pk["d"], 1.6)
    ttl(ax, "Observed vs residual", "Pola menyebar/mengerucut menandakan ragam residual tidak konstan"); ax.set_xlabel("Nilai observasi (aktual)"); ax.set_ylabel("Residual")
    return fig


def shapiro_p(e):
    e = np.asarray(e)
    if len(e) > 5000: e = np.random.default_rng(1).choice(e, 5000, replace=False)
    return stats.shapiro(e)


def interp_reg(M):
    r = M["res"]; resp, logy = M["resp"], M["logy"]; pF = r.f_pvalue; r2, ar2 = r.rsquared, r.rsquared_adj
    kat = "lemah" if ar2 < .3 else "sedang" if ar2 < .6 else "kuat"; nm = M["names"]
    out = [f"**Uji simultan (uji F):** F = {fmt(r.fvalue, 2)}, p {'< 0.001' if pF < .001 else '= ' + fmt_p(pF)}. " +
           ("Minimal satu prediktor berpengaruh signifikan terhadap respons, sehingga model layak digunakan." if pF < .05 else "Tidak ada prediktor yang berpengaruh signifikan secara simultan."),
           f"**Kemampuan model:** R2 = {fmt(r2, 3)} dan Adj. R2 = {fmt(ar2, 3)}, artinya sekitar **{fmt(100 * ar2, 1)}%** keragaman {'ln(' + resp + ')' if logy else resp} dijelaskan oleh prediktor dalam model (kategori **{kat}**); sisanya dijelaskan faktor lain di luar model."]
    idx = [i for i in range(1, len(nm))]; sg = [i for i in idx if r.pvalues[i] < .05]
    if not sg:
        out.append("**Uji parsial (uji t):** tidak ada koefisien yang signifikan pada taraf 5%.")
    else:
        out.append("**Uji parsial (uji t):** prediktor signifikan pada taraf 5%: **" + ", ".join(M["collabel"][nm[i]] for i in sg) + "**.")
        for i in sg:
            b = r.params[i]; tl = M["colterm"][nm[i]]; isnum = tl not in M["xlevels"]
            eff = f"{fmt(100 * (np.exp(abs(b)) - 1), 1)}%" if logy else fmt(abs(b), 3)
            if isnum:
                out.append(f"Setiap kenaikan 1 satuan **{tl}**, rata-rata **{resp}** {'naik' if b > 0 else 'turun'} sebesar **{eff}**, dengan variabel lain dianggap tetap.")
            else:
                out.append(f"Dibanding kategori acuan **{M['xlevels'][tl][0]}**, kategori **{nm[i][len(tl):]}** memiliki rata-rata **{resp}** yang {'lebih tinggi' if b > 0 else 'lebih rendah'} sebesar **{eff}**, dengan variabel lain dianggap tetap.")
    tl_all = list(dict.fromkeys(M["colterm"][nm[i]] for i in idx)); tl_sig = set(M["colterm"][nm[i]] for i in sg); ns = [t for t in tl_all if t not in tl_sig]
    if ns: out.append("Variabel **tidak signifikan** (tidak ada bukti pengaruh setelah variabel lain dikendalikan): " + ", ".join(ns) + ".")
    sh = shapiro_p(r.resid).pvalue; bp = bp_test(M)["p"]; vt = vif_table(M)
    out.append(f"**Asumsi residual:** normalitas (Shapiro-Wilk) p {pq(sh)}, sehingga {'residual **tidak normal**' if sh < .05 else 'residual normal'}; kesamaan ragam (Breusch-Pagan) p {pq(bp)}, sehingga {'ragam residual **tidak konstan**' if bp < .05 else 'ragam konstan'}.")
    if vt is not None:
        mx = vt["VIF setara"].max(); out.append(f"**Multikolinearitas:** VIF setara maksimum = {fmt(mx, 2)} ({'aman, di bawah 5' if mx < 5 else 'perlu diperhatikan, di atas 5'}).")
    return out


# =============================================================================
# KESIMPULAN OTOMATIS
# =============================================================================
def auto_conclusion(d, M):
    r = M["res"]; resp, logy = M["resp"], M["logy"]; nm = M["names"]; sg = [nm[i] for i in range(1, len(nm)) if r.pvalues[i] < .05]
    v4 = ["Quantity", "Total", "Rating", "Hour"]; cm = d[v4].corr().values.copy(); np.fill_diagonal(cm, 0); mx = np.abs(cm).max()
    cj, ri = np.argwhere(np.abs(cm).T == mx)[0]  # urutan kolom-utama seperti which(arr.ind) di R
    tot = d["Total"].astype(float)
    return [f"Data yang dianalisis berisi **{len(d)} transaksi**; nilai transaksi (Total) berkisar {fmt(tot.min())} sampai {fmt(tot.max())} dengan median **{fmt(tot.median())}** dan rata-rata {fmt(tot.mean())} ({'distribusi miring ke kanan' if skew_fun(tot) > .5 else 'distribusi relatif simetris'}).",
            f"Pasangan variabel numerik dengan korelasi terkuat adalah **{v4[ri]}** dan **{v4[cj]}** (r = {fmt(cm[ri, cj], 3)}).",
            f"Model regresi (respons **{resp}**{', transformasi ln' if logy else ''}) memiliki Adj. R2 = **{fmt(r.rsquared_adj, 3)}**; prediktor signifikan pada taraf 5%: **{'tidak ada' if not sg else ', '.join(M['collabel'][n] for n in sg)}**."]


# EDIT
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
# TAMPILAN 
# =============================================================================
CSS = """
@import url('https://fonts.googleapis.com/css2?family=Source+Sans+3:wght@400;600;700&family=Source+Serif+4:wght@600;700&display=swap');
:root{--p9:#880E4F;--p7:#C2185B;--p5:#F06292;--p3:#F8BBD0;--p1:#FFF5F8;--tx:#3B2A33;--ln:#EBD3DC;}
html,body,[class*="css"],.stMarkdown,p,li,label{font-family:'Source Sans 3','Segoe UI',Arial,sans-serif;color:var(--tx);}
h1,h2,h3,h4{font-family:'Source Serif 4',Georgia,'Times New Roman',serif !important;color:var(--p9) !important;}
h4{font-size:16px !important;font-weight:700 !important;margin:0 0 .4rem 0 !important;padding:0 !important;}
.block-container{padding-top:1rem;}
.app-header{border-top:6px solid var(--p7);border-bottom:1px solid var(--ln);background:var(--p1);padding:22px 30px 16px;margin:0 0 22px;}
.app-header h1{margin:0 0 6px;font-size:25px;font-weight:700;padding:0;}
.app-header .sub{color:#6D4C5E;font-size:14px;margin:0;}
section[data-testid="stSidebar"]{background:var(--p1);border-right:1px solid var(--ln);}
section[data-testid="stSidebar"] h4{border-bottom:1px solid var(--ln);padding-bottom:6px !important;margin-top:.6rem !important;}
[class*="st-key-box"]{background:#fff;border:1px solid var(--ln) !important;border-left:4px solid var(--p7) !important;border-radius:3px;padding:14px 18px;margin-bottom:6px;}
[class*="st-key-box"] p,[class*="st-key-box"] li{line-height:1.6;}
button[kind="primary"]{background:var(--p7) !important;border:1px solid var(--p7) !important;color:#fff !important;border-radius:3px;font-weight:600;}
button[kind="primary"]:hover{background:var(--p9) !important;border-color:var(--p9) !important;}
button[kind="secondary"]{border-radius:3px;border-color:var(--ln);}
label[data-testid="stWidgetLabel"] p{font-weight:600;color:#6D4C5E;font-size:13px;}
.nfilter{background:#fff;border:1px solid var(--ln);border-radius:3px;padding:8px 10px;margin-top:10px;font-size:13px;color:#6D4C5E;}
.nfilter b{color:var(--p9);font-size:16px;}
.help-note{color:#7d6672;font-size:13px;}
pre{background:var(--p1) !important;border:1px solid var(--ln);font-size:12.5px;}
"""
st.markdown(f"<style>{CSS}</style>", unsafe_allow_html=True)
st.markdown(f"<div class='app-header'><h1>{_html.escape(APP_TITLE)}</h1><p class='sub'>Ujian Tengah Semester Eksplorasi dan Visualisasi Data Multivariat &nbsp;|&nbsp; {_html.escape(NAMA)} &nbsp;|&nbsp; {NIM}</p></div>", unsafe_allow_html=True)

_BOX = {"n": 0}


def box(title=None):  # noqa: F811  (versi dengan key agar bisa di-style CSS)
    _BOX["n"] += 1
    c = st.container(border=True, key=f"box{_BOX['n']}")
    if title:
        c.markdown(f"#### {title}")
    return c


def note(t):
    st.markdown(f"<p class='help-note'>{_html.escape(t)}</p>", unsafe_allow_html=True)


# ---- Widget dengan state yang bertahan antar-tab (padanan input$... Shiny) ---
def _h(opts):
    return zlib.crc32("|".join(map(str, opts)).encode())


def _wk(key, opts=None):
    return f"w_{key}" + (f"_{_h(opts)}" if opts is not None else "")


def w_select(label, options, key, default):
    options = list(options); wk = _wk(key, options)
    if wk not in st.session_state:
        s = st.session_state.get("P_" + key, default); st.session_state[wk] = s if s in options else options[0]
    v = st.selectbox(label, options, key=wk); st.session_state["P_" + key] = v
    return v


def w_multi(label, options, key, default):
    options = list(options); wk = _wk(key, options)
    if wk not in st.session_state:
        st.session_state[wk] = [x for x in st.session_state.get("P_" + key, default) if x in options]
    v = st.multiselect(label, options, key=wk); st.session_state["P_" + key] = v
    return v


def w_check(label, key, default):
    wk = _wk(key)
    if wk not in st.session_state:
        st.session_state[wk] = st.session_state.get("P_" + key, default)
    v = st.checkbox(label, key=wk); st.session_state["P_" + key] = v
    return v


def w_radio(label, options, labels, key, default):
    wk = _wk(key)
    if wk not in st.session_state:
        st.session_state[wk] = st.session_state.get("P_" + key, default)
    v = st.radio(label, options, format_func=lambda o: labels[o], key=wk, horizontal=True); st.session_state["P_" + key] = v
    return v


def w_pills(options, key):
    wk = _wk(key)
    if wk not in st.session_state:
        st.session_state[wk] = st.session_state.get("P_" + key, options[0])
    v = st.pills("sub", options, selection_mode="single", key=wk, label_visibility="collapsed") or options[0]
    st.session_state["P_" + key] = v
    return v


# =============================================================================
# SIDEBAR: sumber data & filter
# =============================================================================
d0 = load_default(DATA_FILE)
sb = st.sidebar
sb.markdown("#### Sumber data")
up = sb.file_uploader("Excel/CSV (opsional)", type=["xlsx", "xls", "csv"], label_visibility="collapsed")
if up is not None:
    data_raw = load_upload(up.getvalue(), up.name.rsplit(".", 1)[-1].lower()); sig = f"{up.name}_{up.size}"
else:
    data_raw = d0; sig = "default"
if data_raw is None:
    st.error("Data tidak ditemukan. Letakkan Data_Penjualan.xlsx di folder aplikasi atau unggah file melalui panel kiri (kolom wajib: Branch, Customer type, Gender, Product line, Quantity, Total, Date, Time, Payment, Rating).")
    st.stop()


def _reset():
    st.session_state["reset_n"] = st.session_state.get("reset_n", 0) + 1


ver = f"{sig}_{st.session_state.get('reset_n', 0)}"
sb.markdown("#### Filter data")
lv = lambda v: list(data_raw[v].cat.categories)
f_branch = sb.multiselect("Branch", lv("Branch"), default=lv("Branch"), key=f"f_branch_{ver}")
f_cust = sb.multiselect("Customer type", lv("Customer type"), default=lv("Customer type"), key=f"f_cust_{ver}")
f_gender = sb.multiselect("Gender", lv("Gender"), default=lv("Gender"), key=f"f_gender_{ver}")
f_pay = sb.multiselect("Payment", lv("Payment"), default=lv("Payment"), key=f"f_pay_{ver}")
f_prod = sb.multiselect("Product line", lv("Product line"), default=lv("Product line"), key=f"f_prod_{ver}")
dmin, dmax = data_raw["Date"].min().date(), data_raw["Date"].max().date()
f_date = sb.date_input("Rentang tanggal", value=(dmin, dmax), min_value=dmin, max_value=dmax, format="DD-MM-YYYY", key=f"f_date_{ver}")
sb.button("Atur ulang filter", on_click=_reset)
if not (isinstance(f_date, (tuple, list)) and len(f_date) == 2):
    st.info("Pilih tanggal akhir pada rentang tanggal (panel kiri)."); st.stop()

d = data_raw[data_raw["Branch"].isin(f_branch) & data_raw["Customer type"].isin(f_cust) & data_raw["Gender"].isin(f_gender)
             & data_raw["Payment"].isin(f_pay) & data_raw["Product line"].isin(f_prod)
             & (data_raw["Date"] >= pd.Timestamp(f_date[0])) & (data_raw["Date"] <= pd.Timestamp(f_date[1]))]
n0, n1 = len(data_raw), len(d)
sb.markdown(f"<div class='nfilter'>Transaksi terpilih:<br><b>{n1:,}</b> dari {n0:,} ({fmt(100 * n1 / n0, 1)}%)</div>", unsafe_allow_html=True)
if n1 < 30:
    st.error("Data hasil filter terlalu sedikit (kurang dari 30 baris). Longgarkan filter pada panel kiri."); st.stop()
d = d.copy().reset_index(drop=True)
for c in d.columns:
    if isinstance(d[c].dtype, pd.CategoricalDtype):
        d[c] = d[c].cat.remove_unused_categories()


# =============================================================================
# 1. DATA
# =============================================================================
def tab_data():
    with box("Latar belakang"):
        st.markdown(LATAR_BELAKANG)
    with box("Tujuan analisis"):
        st.markdown("\n".join(f"{i}. {t}" for i, t in enumerate(TUJUAN, 1)))
    with box("Sumber dan karakteristik data"):
        st.markdown(SUMBER_DATA)
        bullets([f"Jumlah observasi: **{n0:,}** transaksi; jumlah variabel: **10** variabel asli (ditambah 3 variabel turunan: Hour, Bulan, Hari).",
                 f"Periode data: **{data_raw['Date'].min():%Y-%m-%d}** sampai **{data_raw['Date'].max():%Y-%m-%d}**.",
                 "Variabel kategorik: Branch, Customer type, Gender, Product line, Payment. Variabel numerik: Quantity, Total, Rating, serta Hour (turunan)."])
    with box("Definisi variabel"):
        st.dataframe(kamus, hide_index=True)
    with box("Pengolahan data"):
        q1, q3 = np.percentile(data_raw["Total"], [25, 75]); out = int(((data_raw["Total"] < q1 - 1.5 * (q3 - q1)) | (data_raw["Total"] > q3 + 1.5 * (q3 - q1))).sum())
        bullets([f"**Data hilang:** {int(data_raw.isna().sum().sum())} sel kosong pada seluruh kolom analisis.",
                 f"**Data ganda:** {int(data_raw.duplicated().sum())} baris duplikat.",
                 f"**Pencilan:** terdapat {out} pencilan pada Total (aturan 1.5 x IQR). Pencilan dipertahankan karena merupakan transaksi bernilai besar yang valid.",
                 "**Pengolahan:** variabel Hour, Bulan, dan Hari dibentuk dari Time dan Date; variabel kategorik diubah menjadi faktor.",
                 "**Transformasi:** tidak dilakukan pada eksplorasi; pada regresi tersedia pilihan transformasi ln(Y) untuk mengatasi ragam yang tidak konstan."])
    with box("Cuplikan data"):
        x = data_raw.head(10).copy(); x["Date"] = x["Date"].dt.strftime("%Y-%m-%d"); st.dataframe(x, hide_index=True)


# =============================================================================
# 2. UNIVARIAT
# =============================================================================
def tab_uni():
    c1, c2, _ = st.columns([4, 4, 4])
    with c1: v = w_select("Variabel", cat_all + num_all + ["Date"], "uni_var", "Total")
    with c2: t = w_select("Jenis grafik", uni_choices(v), "uni_type", uni_choices(v)[0])
    l, r = st.columns([7, 5])
    with l:
        with box():
            if t == "Stem-and-leaf":
                st.code(f"Stem-and-leaf plot: {v}\n\n" + stem_text(d[v].astype(float).values), language=None)
            else:
                show(plot_uni(d, v, t))
    with r:
        with box("Statistik deskriptif"):
            table(uni_table(d, v))
    with box("Interpretasi"):
        bullets(interp_uni(d, v))


# =============================================================================
# 3. BIVARIAT
# =============================================================================
def tab_bi():
    c = st.columns(4)
    with c[0]: x = w_select("Variabel X", cat_all + num_all, "bi_x", "Quantity")
    with c[1]: y = w_select("Variabel Y", cat_all + num_all, "bi_y", "Total")
    with c[2]: t = w_select("Jenis grafik", bi_choices(x, y), "bi_type", bi_choices(x, y)[0])
    note("Jenis grafik menyesuaikan tipe pasangan variabel: numerik-numerik, kategorik-numerik, atau kategorik-kategorik.")
    if x == y:
        st.warning("Pilih dua variabel yang berbeda."); return
    l, r = st.columns([7, 5])
    with l:
        with box(): show(plot_bi(d, x, y, t))
    tab, lines = bi_text(d, x, y)
    with r:
        with box("Ringkasan dan uji"): table(tab)
    with box("Interpretasi"): bullets(lines)


# =============================================================================
# 4. MULTIVARIAT
# =============================================================================
def _top_pairs(m):
    n = m.shape[0]; rows = [(m.index[i], m.columns[j], m.iloc[i, j]) for i in range(n) for j in range(i + 1, n)]
    return sorted(rows, key=lambda t: -abs(t[2]))


def mv_scatter():
    c = st.columns(4)
    with c[0]: x = w_select("Sumbu X", num_all, "mv_x", "Quantity")
    with c[1]: y = w_select("Sumbu Y", num_all, "mv_y", "Total")
    with c[2]: col = w_select("Warna", [NONE] + cat_all + num_all, "mv_col", "Customer type")
    with c[3]: size = w_select("Ukuran titik", [NONE] + num_all, "mv_size", "Rating")
    c = st.columns(4)
    with c[0]: facet = w_select("Panel (facet)", [NONE] + cat_all, "mv_facet", NONE)
    with c[1]: lines = w_check("Tampilkan garis regresi per kelompok", "mv_lines", True)
    if x == y:
        st.warning("Pilih sumbu X dan Y yang berbeda."); return
    with box(): show(plot_mv_scatter(d, x, y, col, size, facet, lines))
    r = d[x].astype(float).corr(d[y].astype(float))
    ln = [f"Secara keseluruhan korelasi **{x}** dan **{y}** adalah **{fmt(r, 3)}** ({strength(r)})."]
    if col != NONE and col in cat_all:
        g = {k: d.loc[d[col] == k, x].astype(float).corr(d.loc[d[col] == k, y].astype(float)) for k in d[col].cat.categories}
        ln.append(f"Korelasi per **{col}**: " + "; ".join(f"{k} = {fmt(v, 2)}" for k, v in g.items()) + ". " +
                  ("Pola hubungan hampir sama di semua kelompok sehingga variabel warna tidak mengubah hubungan X dan Y." if np.ptp(list(g.values())) < .1
                   else "Kekuatan hubungan berbeda antar kelompok, menandakan kemungkinan interaksi."))
    if size != NONE:
        ln.append(f"Ukuran titik menunjukkan **{size}**; jika titik besar dan kecil tersebar acak, variabel tersebut tidak berkaitan dengan posisi titik pada X dan Y.")
    with box("Interpretasi"): bullets(ln)


def mv_heat():
    c = st.columns(4)
    with c[0]: a = w_select("Kategorik 1 (sumbu X)", cat_all, "mv_ha", "Branch")
    with c[1]: b = w_select("Kategorik 2 (sumbu Y)", cat_all, "mv_hb", "Product line")
    with c[2]: n = w_select("Variabel numerik", num_all, "mv_hn", "Total")
    with c[3]: f = w_select("Ukuran", ["Rata-rata", "Median", "Jumlah", "Banyak transaksi"], "mv_hf", "Rata-rata")
    if a == b:
        st.warning("Pilih dua variabel kategorik yang berbeda."); return
    with box(): show(plot_mv_heat(d, a, b, n, f))
    r = mv_heat_data(d, a, b, n, f); hi = r.loc[r["val"].idxmax()]; lo = r.loc[r["val"].idxmin()]
    with box("Interpretasi"):
        bullets([f"Nilai tertinggi pada kombinasi **{hi['A']} - {hi['B']}** ({fmt(hi['val'], 1)}) dan terendah pada **{lo['A']} - {lo['B']}** ({fmt(lo['val'], 1)}).",
                 f"Selisih tertinggi-terendah sebesar **{fmt(hi['val'] - lo['val'], 1)}** ({fmt(100 * (hi['val'] - lo['val']) / hi['val'], 1)}% dari nilai tertinggi).",
                 "Warna lebih pekat menandakan nilai lebih besar. Perbedaan antar sel perlu dikonfirmasi dengan uji statistik sebelum disimpulkan sebagai perbedaan nyata."])


def mv_cor():
    c = st.columns([6, 3, 3])
    with c[0]: vs = w_multi("Variabel numerik", num_all, "mv_cv", num_all)
    with c[1]: me = w_select("Metode", ["pearson", "spearman"], "mv_cm", "pearson")
    if len(vs) < 2:
        st.warning("Pilih minimal dua variabel."); return
    with box(): show(plot_mv_cor(d, vs, me))
    pr = _top_pairs(d[vs].astype(float).corr(method=me))
    with box("Interpretasi"):
        bullets([f"Korelasi terkuat: **{pr[0][0]} - {pr[0][1]}** (r = {fmt(pr[0][2], 3)}; {strength(pr[0][2])}).",
                 f"Korelasi terlemah: **{pr[-1][0]} - {pr[-1][1]}** (r = {fmt(pr[-1][2], 3)}).",
                 "Warna merah muda tua = korelasi positif, biru keabuan = negatif, putih = mendekati nol."])


def mv_trend():
    c = st.columns(4)
    with c[0]: m = w_select("Metrik", ["Jumlah transaksi", "Total penjualan (jumlah)", "Rata-rata Total", "Rata-rata Rating", "Total Quantity"], "mv_tm", "Total penjualan (jumlah)")
    with c[1]: u = w_select("Satuan waktu", ["Harian", "Mingguan", "Bulanan"], "mv_tu", "Mingguan")
    with c[2]: g = w_select("Kelompok", [NONE] + cat_base, "mv_tg", "Branch")
    with box(): show(plot_mv_trend(d, m, u, g))
    td = trend_data(d, m, u, g); hi = td.loc[td["val"].idxmax()]; lo = td.loc[td["val"].idxmin()]; gt = f" (kelompok {hi[g]})" if g != NONE else ""
    with box("Interpretasi"):
        bullets([f"Nilai **{m}** tertinggi pada periode **{hi['Periode']:%Y-%m-%d}**{gt} sebesar {fmt(hi['val'], 1)}; terendah pada **{lo['Periode']:%Y-%m-%d}** sebesar {fmt(lo['val'], 1)}.",
                 "Bandingkan garis antar kelompok untuk melihat apakah pola naik-turun sama (pola serupa menandakan faktor waktu yang umum, bukan khas satu kelompok).",
                 "Minggu pertama dan terakhir dapat berisi lebih sedikit hari sehingga nilainya tampak lebih rendah." if u == "Mingguan"
                 else "Perhatikan bahwa data hanya mencakup beberapa bulan, sehingga pola musiman tidak dapat disimpulkan."])


def mv_pca():
    c = st.columns([6, 3, 3])
    with c[0]: vs = w_multi("Variabel numerik (minimal 3)", num_all, "mv_pv", num_all)
    with c[1]: col = w_select("Warna titik", [NONE] + cat_base, "mv_pc", "Customer type")
    if len(vs) < 3:
        st.warning("Pilih minimal tiga variabel numerik."); return
    pr = pca_fit(d, vs); ve = 100 * pr["sdev"] ** 2 / np.sum(pr["sdev"] ** 2)
    l, r = st.columns([7, 5])
    with l:
        with box(): show(plot_mv_pca(d, vs, col))
    with r:
        with box("Keragaman dan loading"):
            ld = pr["rotation"].iloc[:, :2].round(3); t = pd.DataFrame({"Variabel": ld.index, "PC1": ld["PC1"].values, "PC2": ld["PC2"].values})
            t.loc[len(t)] = ["Keragaman dijelaskan (%)", round(ve[0], 1), round(ve[1], 1)]; table(t)
    l1 = pr["rotation"]["PC1"]; k = l1.abs().idxmax()
    with box("Interpretasi"):
        bullets([f"PC1 menjelaskan **{fmt(ve[0], 1)}%** dan PC2 **{fmt(ve[1], 1)}%** keragaman; bersama-sama **{fmt(ve[0] + ve[1], 1)}%**.",
                 f"Variabel dengan kontribusi terbesar pada PC1: **{k}** (loading {fmt(l1[k], 2)}).",
                 "Panah yang searah menandakan variabel berkorelasi positif; panah yang saling tegak lurus menandakan variabel hampir tidak berkorelasi. Jika keragaman terbagi merata ke banyak komponen, variabel-variabel tersebut membawa informasi yang berbeda-beda."])


def mv_pairs():
    c = st.columns([6, 3, 3])
    with c[0]: vs = w_multi("Variabel numerik (minimal 2)", num_all, "mv_mvars", num_all)
    with c[1]: col = w_select("Warna titik", [NONE] + cat_base, "mv_mcol", NONE)
    if len(vs) < 2:
        st.warning("Pilih minimal dua variabel numerik."); return
    with box(): show(plot_mv_pairs(d, vs, col))
    pr = _top_pairs(d[vs].astype(float).corr())
    with box("Interpretasi"):
        bullets(["Diagonal: histogram tiap variabel; segitiga bawah: scatterplot dengan kurva lowess; segitiga atas: koefisien korelasi Pearson (semakin besar angka, semakin kuat hubungan).",
                 f"Pasangan dengan hubungan terkuat: **{pr[0][0]} - {pr[0][1]}** (r = {fmt(pr[0][2], 3)}; {strength(pr[0][2])}). Pasangan terlemah: **{pr[-1][0]} - {pr[-1][1]}** (r = {fmt(pr[-1][2], 3)}).",
                 "Kurva lowess yang mendekati garis lurus menandakan hubungan linear; kurva melengkung menandakan hubungan nonlinear. Titik yang jauh dari kumpulan utama adalah kandidat pencilan."])


def mv_mahal():
    vs = w_multi("Variabel numerik (minimal 2)", num_all, "mv_dvars", num_all)
    if len(vs) < 2:
        st.warning("Pilih minimal dua variabel numerik."); return
    with guard():
        f = mahal_fit(d, vs)
        a, b = st.columns(2)
        with a:
            with box(): show(plot_mv_mahal_idx(f))
        with b:
            with box(): show(plot_mv_mahal_sc(f, vs))
        with box("Sepuluh observasi dengan jarak terbesar"): table(mahal_table(d, f, vs))
        n = len(f["df"]); k = int(f["df"]["out"].sum()); top = f["df"].loc[f["df"]["D2"].idxmax()]
        z = {v: (d[v].astype(float).values[int(top["idx"]) - 1] - d[v].astype(float).mean()) / d[v].astype(float).std(ddof=1) for v in vs}; vmax = max(z, key=lambda q: abs(z[q]))
        with box("Interpretasi"):
            bullets([f"Batas pencilan multivariat: Chi-Square 97.5% dengan df = {f['p']}, yaitu **{fmt(f['cut'], 2)}**. Observasi dengan D2 di atas batas dianggap pencilan.",
                     f"Ditemukan **{k} pencilan** (**{fmt(100 * k / n, 1)}%** dari {n} observasi). Pada data normal multivariat, proporsi yang diharapkan sekitar 2.5%.",
                     f"Observasi paling ekstrem: baris **{int(top['idx'])}** (D2 = {fmt(top['D2'], 2)}). Variabel yang paling menyimpang: **{vmax}** (z = {fmt(z[vmax], 2)}).",
                     "Pencilan multivariat dapat berupa kombinasi nilai yang tidak lazim walau tiap variabel tampak wajar. Pencilan sebaiknya diperiksa dulu keabsahannya sebelum dihapus."])


def mv_qq():
    vs = w_multi("Variabel numerik (minimal 2)", num_all, "mv_qvars", num_all)
    if len(vs) < 2:
        st.warning("Pilih minimal dua variabel numerik."); return
    with guard():
        f = mahal_fit(d, vs)
        with box(): show(plot_mv_qqchi(f))
        n = len(f["df"]); s = np.sort(f["df"]["D2"].values); q = stats.chi2.ppf((np.arange(1, n + 1) - .5) / n, df=f["p"]); r = np.corrcoef(s, q)[0, 1]
        last = f["df"]["idx"].values[np.argsort(f["df"]["D2"].values)][-3:]
        with box("Interpretasi"):
            bullets([f"Korelasi antara D2 terurut dan kuantil Chi-Square = **{fmt(r, 3)}**; semakin mendekati 1, semakin sesuai data dengan normal multivariat.",
                     f"Titik terjauh dari garis berada di ujung kanan (baris {', '.join(str(int(v)) for v in last)}), yang menandakan observasi dengan jarak Mahalanobis besar.",
                     "Titik yang melengkung ke atas di ujung kanan menandakan ekor distribusi lebih berat atau adanya pencilan; pola berundak dapat muncul karena variabel diskret (Quantity, Hour). Asumsi normalitas multivariat dianggap kurang terpenuhi bila penyimpangan jelas."])


def tab_mv():
    subs = {"Scatter multivariat": mv_scatter, "Heatmap rata-rata": mv_heat, "Matriks korelasi": mv_cor, "Tren waktu": mv_trend, "PCA": mv_pca,
            "Scatterplot matrix": mv_pairs, "Pencilan Mahalanobis": mv_mahal, "Q-Q plot Chi-Square": mv_qq}
    sel = w_pills(list(subs), "mv_sub"); subs[sel]()


# =============================================================================
# 5. REGRESI LINEAR BERGANDA
# =============================================================================
RG_CHOICES = ["Quantity", "Rating", "Hour"] + cat_all
RG_DEFAULT = ["Quantity", "Rating", "Hour"] + cat_base


def try_build(dd, y, xs, log):
    try:
        preds = [p for p in RG_CHOICES if p in xs and p != y]
        need(len(preds) >= 2, "Pilih minimal dua variabel prediktor (selain variabel respons).")
        if log: need((dd[y] > 0).all(), "Transformasi ln hanya untuk respons bernilai positif.")
        return build_model(dd, y, preds, log)
    except Need as e:
        return {"err": str(e)}


def get_model():
    if "MODEL" not in st.session_state:
        st.session_state["MODEL"] = try_build(d, "Total", RG_DEFAULT, False)
    return st.session_state["MODEL"]


def reg_hasil(M):
    with box("Persamaan model"): st.code(make_equation(M), language=None)
    a, b = st.columns([5, 7])
    with a:
        with box("Ukuran kecocokan model"): table(fit_table(M))
    with b:
        with box("Uji signifikansi per variabel (ANOVA)"): table(anova_table(M))
    with box("Estimasi koefisien"):
        table(coef_table(M)); note("Signif.: *** p<0.001, ** p<0.01, * p<0.05, . p<0.1")
    with box("Interpretasi hasil"): bullets(interp_reg(M))


def reg_visual(M):
    a, b = st.columns(2)
    with a:
        with box(): show(plot_reg_avp(M))
    with b:
        with box(): show(plot_reg_coef(M))
    with box("Cara membaca"):
        bullets([f"**Aktual vs prediksi:** semakin rapat titik di sekitar garis putus-putus, semakin baik prediksi model (R2 = {fmt(M['res'].rsquared, 3)}). Titik yang menyebar lebar menunjukkan keragaman yang belum dijelaskan model.",
                 "**Koefisien dan selang kepercayaan:** titik = estimasi, garis = selang 95%. Selang yang tidak memotong garis nol (warna tua) menandakan prediktor signifikan pada taraf 5%. Skala koefisien bergantung pada satuan variabelnya, jadi bandingkan terutama ada-tidaknya pengaruh, bukan panjang garisnya."])


def reg_galeri(M):
    fns = [plot_reg_scatter, plot_reg_cor, plot_reg_coef, plot_reg_qq, plot_reg_hist, plot_reg_ts, plot_reg_acf, plot_reg_obsres]
    for i in range(0, 8, 2):
        cols = st.columns(2)
        for c, fn in zip(cols, fns[i:i + 2]):
            with c:
                with box():
                    with guard(): show(fn(M))
    o = reg_ord_res(M); e = o["res"]; n = len(e); L = min(30, n - 1); a = sm_acf(e, nlags=L, fft=False); ci = stats.norm.ppf(.975) / np.sqrt(n)
    nsig = int(np.sum(np.abs(a[1:]) > ci)); dw = np.sum(np.diff(e) ** 2) / np.sum(e ** 2); sw = shapiro_p(e).pvalue
    with box("Cara membaca"):
        bullets(["**Scatterplot:** menunjukkan bentuk hubungan respons dengan tiap prediktor numerik; garis yang jelas menanjak/menurun menandakan hubungan linear.",
                 "**Correlation matrix:** korelasi respons dengan prediktor menunjukkan kandidat prediktor kuat; korelasi tinggi antar prediktor mengindikasikan potensi multikolinearitas.",
                 "**Coefficient CI plot:** selang 95% yang tidak memotong garis nol berarti koefisien signifikan pada taraf 5%.",
                 f"**Normal probability plot dan histogram residual:** titik mengikuti garis dan histogram mendekati kurva normal berarti residual normal. Hasil Shapiro-Wilk p {pq(sw)}, sehingga residual {'**tidak normal**' if sw < .05 else 'normal'}.",
                 f"**Time series plot dan ACF:** statistik Durbin-Watson = {fmt(dw, 3)} (sekitar 2 berarti tidak ada autokorelasi); {nsig} dari {L} lag berada di luar batas 95%. " +
                 ("Tidak ada indikasi autokorelasi residual yang berarti." if (nsig <= 0.05 * L + 1 and abs(dw - 2) < .4) else "Terdapat indikasi autokorelasi residual; periksa pola waktu."),
                 "**Observed vs residual:** residual idealnya tersebar acak di sekitar nol. Sebaran yang melebar pada nilai observasi besar menandakan ragam residual tidak konstan (heteroskedastis)."])


def reg_diag(M):
    for fa, fb in ((plot_reg_resid, plot_reg_qq), (plot_reg_scale, plot_reg_cook)):
        a, b = st.columns(2)
        with a:
            with box(): show(fa(M))
        with b:
            with box(): show(fb(M))
    a, b = st.columns(2)
    e = M["res"].resid; sh = shapiro_p(e); bp = bp_test(M); cd = reg_df(M)["cook"].values; nc = int(np.sum(cd > 4 / len(cd)))
    with a:
        with box("Uji asumsi"):
            table(pd.DataFrame({"Uji": ["Shapiro-Wilk (normalitas residual)", "Breusch-Pagan (kesamaan ragam)", "Observasi berpengaruh (Cook's D > 4/n)"],
                                "Statistik": [fmt(sh.statistic, 4), fmt(bp["stat"], 3), str(nc)], "p-value": [fmt_p(sh.pvalue), fmt_p(bp["p"]), "-"],
                                "Keputusan": ["Tidak normal" if sh.pvalue < .05 else "Normal", "Ragam tidak konstan" if bp["p"] < .05 else "Ragam konstan", f"{nc} observasi"]}))
    with b:
        with box("Multikolinearitas (VIF)"):
            v = vif_table(M)
            if v is None: st.warning("VIF membutuhkan minimal dua prediktor.")
            else:
                v["Keputusan"] = np.where(v["VIF setara"] < 5, "Aman", "Perlu perhatian"); v["VIF setara"] = [fmt(x, 3) for x in v["VIF setara"]]; table(v)


def tab_reg():
    with box("Spesifikasi model"):
        c = st.columns([3, 5, 2, 2])
        with c[0]: y = w_select("Variabel respons (Y)", ["Total", "Rating"], "rg_y", "Total")
        with c[1]: xs = w_multi("Variabel prediktor (X)", RG_CHOICES, "rg_x", RG_DEFAULT)
        with c[2]: lg = w_radio("Transformasi respons", ["no", "yes"], {"no": "Tanpa", "yes": "ln(Y)"}, "rg_log", "no")
        with c[3]:
            st.write(""); go = st.button("Bangun model", type="primary", key="rg_go")
        note("Respons Total dipilih karena merupakan ukuran utama nilai transaksi. Pilih minimal dua prediktor. Variabel kategorik otomatis diubah menjadi dummy dengan kategori pertama sebagai acuan.")
    if go or "MODEL" not in st.session_state:
        st.session_state["MODEL"] = try_build(d, y, xs, lg == "yes")
    M = st.session_state["MODEL"]
    sub = w_pills(["Hasil model", "Visualisasi hasil model", "Galeri visualisasi regresi", "Diagnostik asumsi", "Keterbatasan model"], "rg_sub")
    if sub == "Keterbatasan model":
        with box("Keterbatasan model yang digunakan"): bullets(KETERBATASAN)
        return
    if "err" in M:
        st.warning(M["err"]); return
    {"Hasil model": reg_hasil, "Visualisasi hasil model": reg_visual, "Galeri visualisasi regresi": reg_galeri, "Diagnostik asumsi": reg_diag}[sub](M)


# =============================================================================
# 6. KESIMPULAN
# =============================================================================
def tab_kes():
    M = get_model()
    with box("Kesimpulan utama"):
        if "err" in M: st.warning(M["err"])
        else: bullets(auto_conclusion(d, M))
    with box("Rekomendasi"): bullets(REKOMENDASI)


# =============================================================================
# UTAMA
# =============================================================================
TABS = {"1. Masalah dan Data": tab_data, "2. Univariat": tab_uni, "3. Bivariat": tab_bi, "4. Multivariat": tab_mv,
        "5. Regresi Linear Berganda": tab_reg, "6. Kesimpulan": tab_kes}
cur = st.segmented_control("Menu", list(TABS), default=list(TABS)[0], key="tabs_main", label_visibility="collapsed") or list(TABS)[0]
TABS[cur]()
