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
theme_pink <- function(base = 12) {
  theme_minimal(base_size = base) +
    theme(plot.title = element_text(face = "bold", colour = pk$d, size = base + 2),
          plot.subtitle = element_text(colour = pk$mauve),
          axis.title = element_text(face = "bold", colour = pk$mauve),
          axis.text = element_text(colour = pk$txt),
          legend.title = element_text(face = "bold", colour = pk$mauve),
          strip.text = element_text(face = "bold", colour = pk$d),
          panel.grid.minor = element_blank(),
          panel.grid.major = element_line(colour = pk$grid),
          plot.background = element_rect(fill = "white", colour = NA),
          panel.background = element_rect(fill = "white", colour = NA))
}
rot <- theme(axis.text.x = element_text(angle = 20, hjust = 1), plot.margin = margin(8, 12, 8, 34))

# ---- Fungsi bantu ------------------------------------------------------------
fmt   <- function(x, k = 2) formatC(x, format = "f", digits = k, big.mark = ",")
fmt_p <- function(p) ifelse(p < 0.001, "< 0.001", sprintf("%.3f", p))
pq    <- function(p) ifelse(p < 0.001, "< 0.001", paste("=", sprintf("%.3f", p)))   # untuk kalimat: "p < 0.001" / "p = 0.019"
stars <- function(p) ifelse(p < 0.001, "***", ifelse(p < 0.01, "**", ifelse(p < 0.05, "*", ifelse(p < 0.1, ".", ""))))
skew_fun <- function(x) { x <- x[!is.na(x)]; n <- length(x); n / ((n-1)*(n-2)) * sum(((x-mean(x))/sd(x))^3) }
kurt_fun <- function(x) {
  x <- x[!is.na(x)]; n <- length(x); z <- (x - mean(x)) / sd(x)
  n*(n+1)/((n-1)*(n-2)*(n-3)) * sum(z^4) - 3*(n-1)^2/((n-2)*(n-3))
}
md <- function(x) HTML(gsub("\\*\\*(.+?)\\*\\*", "<b>\\1</b>", x))
bullets <- function(lines) tags$ul(lapply(lines, function(l) tags$li(md(l))))
strength <- function(r) { a <- abs(r); if (a < .2) "sangat lemah" else if (a < .4) "lemah" else if (a < .6) "sedang" else if (a < .8) "kuat" else "sangat kuat" }

# =============================================================================
# 2. UNIVARIAT
# =============================================================================
uni_choices <- function(v) {
  if (v == "Date") "Deret waktu harian"
  else if (v %in% cat_all) c("Diagram batang", "Pareto", "Diagram lingkaran")
  else c("Histogram", "Histogram + Density", "Boxplot", "Density", "Stem-and-leaf", "QQ-plot")
}

plot_uni <- function(d, v, type) {
  if (v == "Date") {
    dd <- d %>% count(Date)
    return(ggplot(dd, aes(Date, n)) + geom_line(colour = pk$l) + geom_point(colour = pk$m, size = 1.5) +
             geom_smooth(method = "loess", formula = y ~ x, se = FALSE, colour = pk$d) +
             labs(title = "Jumlah transaksi per hari", subtitle = "Garis tebal = tren (loess)", x = "Tanggal", y = "Jumlah transaksi") + theme_pink())
  }
  if (v %in% cat_all) {
    tab <- as.data.frame(table(d[[v]])); names(tab) <- c("Kategori", "n")
    tab$pct <- 100 * tab$n / sum(tab$n)
    if (type == "Diagram lingkaran") {
      return(ggplot(tab, aes(x = "", y = n, fill = Kategori)) + geom_col(width = 1, colour = "white") + coord_polar("y") +
               geom_text(aes(label = paste0(round(pct, 1), "%")), position = position_stack(vjust = .5), colour = "white", fontface = "bold") +
               scale_fill_manual(values = pal_cat(nrow(tab))) + labs(title = paste("Proporsi", v), fill = v) + theme_void(base_size = 12) +
               theme(plot.title = element_text(face = "bold", colour = pk$d, size = 14), plot.background = element_rect(fill = "white", colour = NA)))
    }
    if (type == "Pareto") {
      tab <- tab[order(-tab$n), ]; tab$Kategori <- factor(tab$Kategori, levels = tab$Kategori)
      tab$kum <- cumsum(tab$n) / sum(tab$n) * 100; tot <- sum(tab$n)
      return(ggplot(tab, aes(Kategori)) + geom_col(aes(y = n), fill = pk$l, width = .7) +
               geom_line(aes(y = kum * tot / 100, group = 1), colour = pk$d, linewidth = 1) + geom_point(aes(y = kum * tot / 100), colour = pk$d, size = 2.4) +
               geom_text(aes(y = kum * tot / 100, label = paste0(round(kum, 1), "%")), vjust = -.9, size = 3.4, colour = pk$mauve) +
               scale_y_continuous(limits = c(0, tot * 1.1), sec.axis = sec_axis(~ . / tot * 100, name = "Kumulatif (%)")) +
               labs(title = paste("Pareto", v), x = v, y = "Frekuensi") + theme_pink() + rot)
    }
    return(ggplot(tab, aes(Kategori, n, fill = Kategori)) + geom_col(show.legend = FALSE, width = .7) +
             geom_text(aes(label = n), vjust = -.4, colour = pk$mauve, fontface = "bold") +
             scale_fill_manual(values = pal_cat(nrow(tab))) + scale_y_continuous(expand = expansion(mult = c(0, .1))) +
             labs(title = paste("Frekuensi", v), x = v, y = "Frekuensi") + theme_pink() + rot)
  }
  bw <- v %in% c("Quantity", "Hour")
  hist_l <- function(dens) {
    if (bw) { if (dens) geom_histogram(aes(y = after_stat(density)), binwidth = 1, boundary = .5, fill = pk$xl, colour = "white")
      else geom_histogram(binwidth = 1, boundary = .5, fill = pk$xl, colour = "white") }
    else { if (dens) geom_histogram(aes(y = after_stat(density)), bins = 25, fill = pk$xl, colour = "white")
      else geom_histogram(bins = 25, fill = pk$xl, colour = "white") }
  }
  g <- ggplot(d, aes(x = .data[[v]]))
  if (type == "Histogram + Density") return(g + hist_l(TRUE) + geom_density(colour = pk$d, linewidth = 1.1) + labs(title = paste("Histogram dan Density", v), x = v, y = "Densitas") + theme_pink())
  if (type == "Density") return(g + geom_density(fill = pk$l, colour = pk$d, alpha = .5) + geom_vline(xintercept = mean(d[[v]]), linetype = "dashed", colour = pk$mauve) +
                                  labs(title = paste("Density", v), subtitle = "Garis putus-putus = rata-rata", x = v, y = "Densitas") + theme_pink())
  if (type == "Boxplot") return(ggplot(d, aes(x = "", y = .data[[v]])) + geom_boxplot(fill = pk$xl, colour = pk$d, width = .45, outlier.colour = pk$m) +
                                  stat_summary(fun = mean, geom = "point", shape = 23, size = 3.5, fill = "white", colour = pk$d) + coord_flip() +
                                  labs(title = paste("Boxplot", v), subtitle = "Berlian = rata-rata; titik = pencilan", x = NULL, y = v) + theme_pink())
  if (type == "QQ-plot") return(ggplot(d, aes(sample = .data[[v]])) + stat_qq(colour = pk$m, alpha = .6) + stat_qq_line(colour = pk$d) +
                                  labs(title = paste("QQ-plot", v), subtitle = "Titik mengikuti garis = mendekati normal", x = "Kuantil teoretis", y = "Kuantil sampel") + theme_pink())
  g + hist_l(FALSE) + labs(title = paste("Distribusi", v), x = v, y = "Frekuensi") + theme_pink()
}

uni_table <- function(d, v) {
  if (v == "Date") {
    return(data.frame(Statistik = c("Tanggal awal", "Tanggal akhir", "Jumlah hari", "Jumlah hari dengan transaksi", "Rata-rata transaksi per hari"),
                      Nilai = c(as.character(min(d$Date)), as.character(max(d$Date)), as.character(as.integer(diff(range(d$Date))) + 1),
                                as.character(length(unique(d$Date))), fmt(nrow(d) / length(unique(d$Date))))))
  }
  if (v %in% cat_all) {
    t <- table(d[[v]]); out <- data.frame(Kategori = names(t), Frekuensi = as.integer(t), `Persentase (%)` = fmt(100 * as.numeric(t) / sum(t)), check.names = FALSE)
    return(out)
  }
  x <- d[[v]]; q <- quantile(x, c(.25, .75)); iqr <- q[2] - q[1]
  lo <- q[1] - 1.5 * iqr; hi <- q[2] + 1.5 * iqr
  data.frame(Statistik = c("N", "Rata-rata", "Median", "Simpangan baku", "Variansi", "Koef. variasi (%)", "Minimum", "Q1", "Q3", "Maksimum",
                           "IQR", "Skewness", "Kurtosis (excess)", "Batas pencilan bawah", "Batas pencilan atas", "Jumlah pencilan"),
             Nilai = c(length(x), fmt(mean(x)), fmt(median(x)), fmt(sd(x)), fmt(var(x)), fmt(100 * sd(x) / mean(x)), fmt(min(x)), fmt(q[1]), fmt(q[2]),
                       fmt(max(x)), fmt(iqr), fmt(skew_fun(x), 3), fmt(kurt_fun(x), 3), fmt(lo), fmt(hi), sum(x < lo | x > hi)))
}

interp_uni <- function(d, v) {
  if (v == "Date") {
    dd <- d %>% count(Date); top <- dd[which.max(dd$n), ]
    return(c(sprintf("Data mencakup **%s** sampai **%s** (%d hari bertransaksi).", min(d$Date), max(d$Date), nrow(dd)),
             sprintf("Rata-rata **%s** transaksi per hari; puncak pada **%s** dengan **%d** transaksi.", fmt(mean(dd$n)), top$Date, top$n),
             "Garis tren (loess) menunjukkan apakah volume transaksi harian meningkat, menurun, atau relatif stabil sepanjang periode."))
  }
  if (v %in% cat_all) {
    t <- sort(table(d[[v]]), decreasing = TRUE); p <- 100 * t / sum(t)
    sp <- max(p) - min(p)
    return(c(sprintf("Kategori terbanyak: **%s** (%d transaksi; %s%%). Tersedikit: **%s** (%d; %s%%).", names(t)[1], t[1], fmt(p[1], 1), names(t)[length(t)], t[length(t)], fmt(p[length(p)], 1)),
             sprintf("Selisih proporsi tertinggi dan terendah **%s** poin persentase, sehingga sebaran kategori **%s**.", fmt(sp, 1), if (sp < 5) "relatif seimbang" else "tidak seimbang"),
             "Pareto menunjukkan apakah sedikit kategori mendominasi (garis kumulatif cepat mencapai 80%)."))
  }
  x <- d[[v]]; m <- mean(x); me <- median(x); s <- sd(x); sk <- skew_fun(x); ku <- kurt_fun(x)
  q <- quantile(x, c(.25, .75)); iqr <- q[2] - q[1]; out <- sum(x < q[1] - 1.5 * iqr | x > q[2] + 1.5 * iqr)
  shape <- if (abs(sk) < .5) "hampir simetris" else if (sk > 0) (if (sk > 1) "miring kuat ke kanan" else "miring sedang ke kanan") else (if (sk < -1) "miring kuat ke kiri" else "miring sedang ke kiri")
  c(sprintf("Rata-rata **%s** dan median **%s**%s.", fmt(m), fmt(me), if (abs(m - me) / s > .1) paste0("; rata-rata ", if (m > me) "lebih besar" else "lebih kecil", " dari median sehingga sebaran condong ke ", if (m > me) "kanan" else "kiri") else "; keduanya berdekatan (sebaran seimbang)"),
    sprintf("Sebaran data: simpangan baku **%s**, rentang **%s** sampai **%s**, koefisien variasi **%s%%**.", fmt(s), fmt(min(x)), fmt(max(x)), fmt(100 * s / m)),
    sprintf("Bentuk distribusi **%s** (skewness %s); kurtosis %s menunjukkan puncak yang **%s**.", shape, fmt(sk, 2), fmt(ku, 2), if (ku < -1) "datar (mendekati sebaran seragam)" else if (ku > 1) "runcing" else "mendekati normal"),
    if (out == 0) "Tidak ditemukan pencilan menurut aturan 1.5 x IQR." else sprintf("Terdapat **%d pencilan** menurut aturan 1.5 x IQR (%s%% data). Pencilan dipertahankan karena merupakan transaksi yang valid.", out, fmt(100 * out / length(x), 1)))
}

# =============================================================================
# 3. BIVARIAT
# =============================================================================
bi_choices <- function(x, y) {
  xc <- x %in% cat_all; yc <- y %in% cat_all
  if (!xc && !yc) c("Scatter plot + garis tren", "Scatter plot + kurva linear, kuadratik, nonparametrik", "Line plot rata-rata Y per X", "Correlation plot", "Heatmap kepadatan 2D")
  else if (xc != yc) c("Grouped boxplot", "Grouped violin plot", "Grouped kernel density plot", "Grouped ridgeline plot", "Grouped strip plot",
                       "Grouped mean/SEM plot", "Bar chart statistik ringkasan", "Bar chart rata-rata (+/- SE)", "Cleveland dot chart")
  else c("Stacked bar chart", "Grouped bar chart", "Segmented bar chart (100%)", "Heatmap tabulasi silang")
}

# Correlation plot: lingkaran (segitiga bawah) + angka r (segitiga atas)
plot_corr_circle <- function(d, vars = num_all) {
  m <- cor(d[, vars, drop = FALSE]); df <- as.data.frame(as.table(m)); names(df) <- c("V1", "V2", "r")
  df$i <- match(as.character(df$V1), vars); df$j <- match(as.character(df$V2), vars)
  df$V1 <- factor(df$V1, levels = vars); df$V2 <- factor(df$V2, levels = vars)
  ggplot() +
    geom_point(data = df[df$i > df$j, ], aes(V1, V2, size = abs(r), fill = r), shape = 21, colour = pk$grid) +
    geom_text(data = df[df$i < df$j, ], aes(V1, V2, label = sprintf("%.2f", r)), colour = pk$txt, fontface = "bold") +
    geom_text(data = df[df$i == df$j, ], aes(V1, V2, label = as.character(V1)), colour = pk$d, fontface = "bold") +
    scale_size_continuous(range = c(3, 22), limits = c(0, 1), guide = "none") +
    scale_fill_gradient2(low = pk$slate, mid = "white", high = pk$s, limits = c(-1, 1), midpoint = 0, name = "r") +
    scale_x_discrete(limits = vars) + scale_y_discrete(limits = rev(vars)) +
    labs(title = "Correlation plot (Pearson)", subtitle = "Lingkaran: ukuran dan warna = kekuatan korelasi | Angka: nilai r", x = NULL, y = NULL) +
    theme_pink() + theme(panel.grid = element_blank())
}

plot_bi <- function(d, x, y, type) {
  xc <- x %in% cat_all; yc <- y %in% cat_all
  # ---------- numerik x numerik ----------
  if (!xc && !yc) {
    if (type == "Correlation plot") return(plot_corr_circle(d, num_all))
    if (type == "Scatter plot + kurva linear, kuadratik, nonparametrik") {
      cols <- c("Linear" = pk$d, "Kuadratik" = pk$slate, "Nonparametrik (LOESS)" = "#2E7D32")
      return(ggplot(d, aes(.data[[x]], .data[[y]])) + geom_point(colour = pk$l, alpha = .5, size = 1.8) +
               geom_smooth(aes(colour = "Linear"), method = "lm", formula = y ~ x, se = FALSE, linewidth = 1.1) +
               geom_smooth(aes(colour = "Kuadratik"), method = "lm", formula = y ~ poly(x, 2), se = FALSE, linewidth = 1.1) +
               geom_smooth(aes(colour = "Nonparametrik (LOESS)"), method = "loess", formula = y ~ x, se = FALSE, linewidth = 1.1) +
               scale_colour_manual(values = cols, name = "Kurva") +
               labs(title = paste(y, "terhadap", x), subtitle = "Perbandingan kurva linear, kuadratik, dan nonparametrik (LOESS)", x = x, y = y) + theme_pink())
    }
    if (type == "Line plot rata-rata Y per X") {
      xv <- d[[x]]; yv <- d[[y]]
      grp <- if (length(unique(xv)) <= 12) xv else cut(xv, breaks = 10)
      agg <- data.frame(gx = as.numeric(tapply(xv, grp, mean)), my = as.numeric(tapply(yv, grp, mean)), se = as.numeric(tapply(yv, grp, function(z) sd(z) / sqrt(length(z)))))
      return(ggplot(agg, aes(gx, my)) + geom_ribbon(aes(ymin = my - se, ymax = my + se), fill = pk$xl, alpha = .8) + geom_line(colour = pk$d, linewidth = 1) + geom_point(colour = pk$m, size = 2.5) +
               labs(title = paste("Rata-rata", y, "menurut", x), subtitle = "Pita = rata-rata +/- 1 standard error", x = x, y = paste("Rata-rata", y)) + theme_pink())
    }
    if (type == "Heatmap kepadatan 2D") return(ggplot(d, aes(.data[[x]], .data[[y]])) + geom_bin2d(bins = 20) + scale_fill_gradient(low = pk$bg, high = pk$d, name = "Jumlah") + labs(title = paste("Kepadatan", x, "dan", y), x = x, y = y) + theme_pink())
    return(ggplot(d, aes(.data[[x]], .data[[y]])) + geom_point(colour = pk$m, alpha = .5, size = 1.8) + geom_smooth(method = "lm", formula = y ~ x, se = FALSE, colour = pk$d, linewidth = 1) +
             labs(title = paste(y, "terhadap", x), subtitle = "Scatterplot dengan garis regresi linear sederhana", x = x, y = y) + theme_pink())
  }
  # ---------- kategorik x numerik ----------
  if (xc != yc) {
    cv <- if (xc) x else y; nv <- if (xc) y else x; k <- nlevels(d[[cv]]); cols <- pal_cat(k)
    d2 <- data.frame(g = d[[cv]], v = d[[nv]])
    agg <- d2 %>% group_by(g) %>% summarise(m = mean(v), se = sd(v) / sqrt(n()), md = median(v), s = sd(v), mn = min(v), mx = max(v), .groups = "drop")
    if (type == "Grouped violin plot")
      return(ggplot(d2, aes(g, v, fill = g)) + geom_violin(show.legend = FALSE, alpha = .75, colour = pk$mauve) + geom_boxplot(width = .12, fill = "white", outlier.shape = NA) +
               scale_fill_manual(values = cols) + labs(title = paste("Grouped violin plot:", nv, "menurut", cv), x = cv, y = nv) + theme_pink() + rot)
    if (type == "Grouped kernel density plot")
      return(ggplot(d2, aes(v, fill = g, colour = g)) + geom_density(alpha = .35) + scale_fill_manual(values = cols) + scale_colour_manual(values = cols) +
               labs(title = paste("Grouped kernel density plot:", nv, "per", cv), x = nv, y = "Densitas", fill = cv, colour = cv) + theme_pink())
    if (type == "Grouped ridgeline plot") {
      rng <- range(d2$v); lv <- levels(d2$g)
      rd <- do.call(rbind, lapply(seq_along(lv), function(i) { dd <- density(d2$v[d2$g == lv[i]], from = rng[1], to = rng[2], n = 256); data.frame(g = lv[i], i = i, x = dd$x, y = dd$y) }))
      rd$h <- rd$y / max(rd$y) * 1.6; rd$g <- factor(rd$g, levels = lv); rd <- rd[order(-rd$i), ]
      return(ggplot(rd, aes(x, group = g)) + geom_ribbon(aes(ymin = i, ymax = i + h, fill = g), colour = pk$mauve, alpha = .8, show.legend = FALSE) +
               scale_fill_manual(values = cols) + scale_y_continuous(breaks = seq_along(lv), labels = lv) +
               labs(title = paste("Grouped ridgeline plot:", nv, "per", cv), x = nv, y = cv) + theme_pink())
    }
    if (type == "Grouped strip plot")
      return(ggplot(d2, aes(g, v, colour = g)) + geom_jitter(width = .22, height = if (nv %in% c("Quantity", "Hour")) .15 else 0, alpha = .45, size = 1.6, show.legend = FALSE) +
               stat_summary(fun = mean, geom = "point", shape = 23, size = 4, fill = "white", colour = pk$d) + scale_colour_manual(values = cols) +
               labs(title = paste("Grouped strip plot:", nv, "menurut", cv), subtitle = "Titik = observasi; berlian = rata-rata", x = cv, y = nv) + theme_pink() + rot)
    if (type == "Grouped mean/SEM plot")
      return(ggplot(agg, aes(g, m, colour = g)) + geom_errorbar(aes(ymin = m - se, ymax = m + se), width = .15, linewidth = .9, show.legend = FALSE) + geom_point(size = 3.8, show.legend = FALSE) +
               scale_colour_manual(values = cols) + labs(title = paste("Grouped mean/SEM plot:", nv, "menurut", cv), subtitle = "Titik = rata-rata; garis = +/- 1 standard error (SEM)", x = cv, y = paste("Rata-rata", nv)) + theme_pink() + rot)
    if (type == "Bar chart statistik ringkasan") {
      st <- c("Rata-rata", "Median", "Simpangan baku", "Minimum", "Maksimum")
      long <- data.frame(g = rep(agg$g, 5), stat = factor(rep(st, each = nrow(agg)), levels = st), val = c(agg$m, agg$md, agg$s, agg$mn, agg$mx))
      return(ggplot(long, aes(g, val, fill = g)) + geom_col(show.legend = FALSE, width = .7) + geom_text(aes(label = round(val, 1)), vjust = -.3, size = 3.2, colour = pk$txt) +
               facet_wrap(vars(stat), scales = "free_y") + scale_fill_manual(values = cols) + scale_y_continuous(expand = expansion(mult = c(0, .15))) +
               labs(title = paste("Bar chart statistik ringkasan", nv, "menurut", cv), x = cv, y = nv) + theme_pink(11) + rot)
    }
    if (type == "Bar chart rata-rata (+/- SE)")
      return(ggplot(agg, aes(g, m, fill = g)) + geom_col(show.legend = FALSE, width = .7) + geom_errorbar(aes(ymin = m - se, ymax = m + se), width = .2, colour = pk$txt) +
               geom_text(aes(label = round(m, 2)), vjust = 1.6, colour = "white", fontface = "bold") + scale_fill_manual(values = cols) +
               labs(title = paste("Rata-rata", nv, "menurut", cv), subtitle = "Garis = +/- 1 standard error", x = cv, y = paste("Rata-rata", nv)) + theme_pink() + rot)
    if (type == "Cleveland dot chart") {
      a2 <- agg[order(agg$m), ]; lv <- as.character(a2$g)
      cl <- data.frame(g = factor(rep(lv, 2), levels = lv), stat = rep(c("Rata-rata", "Median"), each = nrow(a2)), val = c(a2$m, a2$md))
      pad <- 0.15 * max(diff(range(cl$val)), abs(mean(cl$val)) * .05); lo <- min(cl$val) - pad; hi <- max(cl$val) + pad
      sg <- data.frame(g = factor(lv, levels = lv), lo = lo, hi = hi)
      return(ggplot(cl, aes(val, g)) + geom_segment(data = sg, aes(x = lo, xend = hi, y = g, yend = g), inherit.aes = FALSE, colour = pk$grid, linetype = "dotted") +
               geom_point(aes(fill = stat), shape = 21, size = 4, colour = pk$d) + scale_fill_manual(values = c("Rata-rata" = pk$m, "Median" = "white"), name = NULL) +
               scale_x_continuous(limits = c(lo, hi)) + labs(title = paste("Cleveland dot chart:", nv, "menurut", cv), subtitle = "Kategori diurutkan menurut rata-rata", x = nv, y = cv) + theme_pink() + theme(legend.position = "bottom"))
    }
    return(ggplot(d2, aes(g, v, fill = g)) + geom_boxplot(show.legend = FALSE, alpha = .85, colour = pk$mauve, outlier.colour = pk$m) +
             stat_summary(fun = mean, geom = "point", shape = 23, size = 3, fill = "white") + scale_fill_manual(values = cols) +
             labs(title = paste("Grouped boxplot:", nv, "menurut", cv), subtitle = "Berlian = rata-rata", x = cv, y = nv) + theme_pink() + rot)
  }
  # ---------- kategorik x kategorik ----------
  ny <- nlevels(d[[y]])
  df <- as.data.frame(table(X = d[[x]], Y = d[[y]]), responseName = "n"); df$pct <- 100 * df$n / ave(df$n, df$X, FUN = sum)
  if (type == "Heatmap tabulasi silang") {
    tab <- as.data.frame(table(d[[x]], d[[y]])); names(tab) <- c("X", "Y", "Freq"); mid <- max(tab$Freq) / 2
    return(ggplot(tab, aes(X, Y, fill = Freq)) + geom_tile(colour = "white") + geom_text(aes(label = Freq, colour = Freq > mid), show.legend = FALSE, fontface = "bold") +
             scale_colour_manual(values = c(`TRUE` = "white", `FALSE` = pk$txt)) + scale_fill_gradient(low = pk$bg, high = pk$d, name = "Jumlah") +
             labs(title = paste(x, "dan", y), x = x, y = y) + theme_pink() + rot)
  }
  if (type == "Stacked bar chart")
    return(ggplot(df, aes(X, n, fill = Y)) + geom_col(position = "stack", colour = "white") + scale_fill_manual(values = pal_cat(ny)) +
             labs(title = paste("Stacked bar chart:", y, "dalam", x), subtitle = "Tinggi total batang = jumlah transaksi per kategori", x = x, y = "Jumlah transaksi", fill = y) + theme_pink() + rot)
  if (type == "Segmented bar chart (100%)")
    return(ggplot(df, aes(X, pct, fill = Y)) + geom_col(position = "stack", colour = "white") + scale_fill_manual(values = pal_cat(ny)) +
             scale_y_continuous(labels = function(z) paste0(z, "%"), expand = expansion(mult = c(0, .02))) +
             labs(title = paste("Segmented bar chart:", y, "dalam", x), subtitle = "Setiap batang = 100%; bandingkan komposisi antar kategori", x = x, y = "Persentase", fill = y) + theme_pink() + rot)
  ggplot(df, aes(X, n, fill = Y)) + geom_col(position = position_dodge(.9)) + geom_text(aes(label = n), position = position_dodge(.9), vjust = -.3, size = 3, colour = pk$txt) +
    scale_fill_manual(values = pal_cat(ny)) + scale_y_continuous(expand = expansion(mult = c(0, .1))) +
    labs(title = paste("Grouped bar chart:", x, "dan", y), x = x, y = "Jumlah transaksi", fill = y) + theme_pink() + rot
}

bi_text <- function(d, x, y) {
  xc <- x %in% cat_all; yc <- y %in% cat_all
  if (!xc && !yc) {
    ct <- cor.test(d[[x]], d[[y]]); sp <- suppressWarnings(cor.test(d[[x]], d[[y]], method = "spearman"))
    fit <- lm(d[[y]] ~ d[[x]]); b <- coef(fit)[2]; r <- unname(ct$estimate)
    tab <- data.frame(Ukuran = c("Korelasi Pearson (r)", "p-value Pearson", "Korelasi Spearman (rho)", "p-value Spearman", "Kemiringan garis (b1)", "R-kuadrat regresi sederhana"),
                      Nilai = c(fmt(r, 3), fmt_p(ct$p.value), fmt(unname(sp$estimate), 3), fmt_p(sp$p.value), fmt(b, 3), fmt(summary(fit)$r.squared, 3)))
    lines <- c(sprintf("Hubungan linear **%s** dan **%s** (r = %s; p %s).", strength(r), if (r > 0) "positif" else "negatif", fmt(r, 3), if (ct$p.value < .001) "< 0.001" else paste("=", fmt_p(ct$p.value))),
               if (ct$p.value < .05) sprintf("Hubungan **signifikan** pada taraf 5%%: setiap kenaikan 1 satuan **%s**, rata-rata **%s** berubah sebesar **%s** satuan.", x, y, fmt(b, 3)) else "Hubungan **tidak signifikan** pada taraf 5%, sehingga tidak ada bukti hubungan linear antar kedua variabel.",
               "Korelasi menunjukkan keterkaitan, bukan sebab-akibat.")
    return(list(table = tab, lines = lines))
  }
  if (xc != yc) {
    cv <- if (xc) x else y; nv <- if (xc) y else x
    g <- d %>% group_by(.data[[cv]]) %>% summarise(`Rata-rata` = fmt(mean(.data[[nv]])), Median = fmt(median(.data[[nv]])), `Simpangan baku` = fmt(sd(.data[[nv]])), n = n(), .groups = "drop"); names(g)[1] <- cv
    gm <- d %>% group_by(.data[[cv]]) %>% summarise(m = mean(.data[[nv]]), .groups = "drop"); names(gm)[1] <- "grp"
    a <- summary(aov(d[[nv]] ~ d[[cv]]))[[1]]; p <- a[["Pr(>F)"]][1]; eta <- a[["Sum Sq"]][1] / sum(a[["Sum Sq"]])
    kw <- kruskal.test(d[[nv]] ~ d[[cv]])$p.value
    test <- if (nlevels(d[[cv]]) == 2) { tt <- t.test(d[[nv]] ~ d[[cv]]); p <- tt$p.value; sprintf("uji t Welch (t = %s)", fmt(unname(tt$statistic), 2)) } else sprintf("ANOVA (F = %s)", fmt(a[["F value"]][1], 2))
    lines <- c(sprintf("Rata-rata **%s** tertinggi pada **%s** (%s) dan terendah pada **%s** (%s).", nv, gm$grp[which.max(gm$m)], fmt(max(gm$m)), gm$grp[which.min(gm$m)], fmt(min(gm$m))),
               sprintf("Hasil %s: p %s; uji Kruskal-Wallis: p %s; ukuran efek eta-kuadrat = %s.", test, pq(p), pq(kw), fmt(eta, 3)),
               if (p < .05) "Perbedaan rata-rata antar kelompok **signifikan** pada taraf 5%." else "Perbedaan rata-rata antar kelompok **tidak signifikan** pada taraf 5%; selisih yang tampak pada grafik dapat terjadi karena kebetulan.")
    return(list(table = g, lines = lines))
  }
  tab <- table(d[[x]], d[[y]]); ch <- suppressWarnings(chisq.test(tab)); V <- sqrt(unname(ch$statistic) / (sum(tab) * (min(dim(tab)) - 1)))
  ix <- which(tab == max(tab), arr.ind = TRUE)[1, ]
  out <- cbind(setNames(data.frame(rownames(tab)), x), as.data.frame.matrix(tab))
  lines <- c(sprintf("Kombinasi terbanyak: **%s** x **%s** (%d transaksi).", rownames(tab)[ix[1]], colnames(tab)[ix[2]], max(tab)),
             sprintf("Uji Chi-square: X2 = %s, df = %d, p %s; Cramer's V = %s.", fmt(unname(ch$statistic), 2), unname(ch$parameter), pq(ch$p.value), fmt(V, 3)),
             if (ch$p.value < .05) "Ada hubungan **signifikan** antar kedua variabel kategorik pada taraf 5%." else "**Tidak ada** hubungan signifikan antar kedua variabel kategorik pada taraf 5%.")
  list(table = out, lines = lines)
}

# =============================================================================
# 4. MULTIVARIAT
# =============================================================================
plot_mv_scatter <- function(d, x, y, col, size, facet, lines) {
  g <- ggplot(d, aes(.data[[x]], .data[[y]]))
  hasc <- col != NONE; hass <- size != NONE
  g <- g + if (hasc && hass) geom_point(aes(colour = .data[[col]], size = .data[[size]]), alpha = .6)
  else if (hasc) geom_point(aes(colour = .data[[col]]), alpha = .6, size = 2)
  else if (hass) geom_point(aes(size = .data[[size]]), colour = pk$m, alpha = .5)
  else geom_point(colour = pk$m, alpha = .6, size = 2)
  if (hasc) g <- g + (if (is.numeric(d[[col]])) scale_colour_gradient(low = pk$xl, high = pk$d, name = col) else scale_colour_manual(values = pal_cat(nlevels(d[[col]])), name = col))
  if (hass) g <- g + scale_size_continuous(range = c(1, 5), name = size)
  if (lines) g <- g + (if (hasc && !is.numeric(d[[col]])) geom_smooth(aes(colour = .data[[col]]), method = "lm", formula = y ~ x, se = FALSE, linewidth = .9)
                       else geom_smooth(method = "lm", formula = y ~ x, se = FALSE, colour = pk$d, linewidth = .9))
  if (facet != NONE) g <- g + facet_wrap(vars(.data[[facet]]))
  g + labs(title = paste(y, "terhadap", x), subtitle = paste0("Warna: ", col, " | Ukuran: ", size, " | Panel: ", facet), x = x, y = y) + theme_pink()
}

mv_heat_data <- function(d, a, b, num, fun) {
  f <- switch(fun, "Rata-rata" = mean, "Median" = median, "Jumlah" = sum, "Banyak transaksi" = length)
  v <- if (fun == "Banyak transaksi") d[[a]] else d[[num]]
  res <- aggregate(v, list(A = d[[a]], B = d[[b]]), f); names(res)[3] <- "val"; res
}
plot_mv_heat <- function(d, a, b, num, fun) {
  res <- mv_heat_data(d, a, b, num, fun); mid <- (max(res$val) + min(res$val)) / 2
  lab <- if (fun == "Banyak transaksi") "Banyak transaksi" else paste(fun, num)
  ggplot(res, aes(A, B, fill = val)) + geom_tile(colour = "white") + geom_text(aes(label = fmt(val, if (fun == "Banyak transaksi") 0 else 1), colour = val > mid), show.legend = FALSE, fontface = "bold") +
    scale_colour_manual(values = c(`TRUE` = "white", `FALSE` = pk$txt)) + scale_fill_gradient(low = pk$bg, high = pk$d, name = lab) +
    labs(title = lab, subtitle = paste("menurut", a, "dan", b), x = a, y = b) + theme_pink() + rot
}

plot_mv_cor <- function(d, vars, method) {
  m <- cor(d[, vars, drop = FALSE], method = method); df <- as.data.frame(as.table(m)); names(df) <- c("V1", "V2", "r")
  ggplot(df, aes(V1, V2, fill = r)) + geom_tile(colour = "white") + geom_text(aes(label = sprintf("%.2f", r)), colour = pk$txt, fontface = "bold", size = 4.5) +
    scale_fill_gradient2(low = pk$slate, mid = "white", high = pk$s, limits = c(-1, 1), midpoint = 0, name = "r") +
    labs(title = paste("Matriks korelasi", tools::toTitleCase(method)), x = NULL, y = NULL) + theme_pink() + theme(panel.grid = element_blank())
}

trend_data <- function(d, metric, unit, grp) {
  d$Periode <- switch(unit, "Harian" = d$Date, "Mingguan" = as.Date(cut(d$Date, "week")), "Bulanan" = as.Date(cut(d$Date, "month")))
  gl <- if (grp == NONE) "Periode" else c("Periode", grp)
  d %>% group_by(across(all_of(gl))) %>%
    summarise(val = switch(metric, "Jumlah transaksi" = n(), "Total penjualan (jumlah)" = sum(Total), "Rata-rata Total" = mean(Total), "Rata-rata Rating" = mean(Rating), "Total Quantity" = sum(Quantity)), .groups = "drop")
}
plot_mv_trend <- function(d, metric, unit, grp) {
  td <- trend_data(d, metric, unit, grp)
  g <- if (grp == NONE) ggplot(td, aes(Periode, val)) + geom_line(colour = pk$d, linewidth = 1) + geom_point(colour = pk$m)
  else ggplot(td, aes(Periode, val, colour = .data[[grp]])) + geom_line(linewidth = 1) + geom_point() + scale_colour_manual(values = pal_cat(nlevels(d[[grp]])))
  g + labs(title = paste(metric, unit, if (grp != NONE) paste("menurut", grp) else ""), subtitle = if (unit == "Mingguan") "Minggu pertama/terakhir dapat tidak lengkap" else NULL, x = "Periode", y = metric, colour = grp) + theme_pink()
}

pca_fit <- function(d, vars) prcomp(d[, vars, drop = FALSE], scale. = TRUE)
plot_mv_pca <- function(d, vars, colv) {
  pr <- pca_fit(d, vars); sc <- as.data.frame(pr$x[, 1:2]); ve <- 100 * pr$sdev^2 / sum(pr$sdev^2)
  ld <- as.data.frame(pr$rotation[, 1:2]); k <- min(max(abs(sc$PC1)), max(abs(sc$PC2))) * 0.85; ld$PC1 <- ld$PC1 * k; ld$PC2 <- ld$PC2 * k; ld$v <- rownames(ld)
  g <- ggplot(sc, aes(PC1, PC2))
  if (colv != NONE) { sc[[colv]] <- d[[colv]]; g <- ggplot(sc, aes(PC1, PC2, colour = .data[[colv]])) + geom_point(alpha = .6, size = 1.8) + scale_colour_manual(values = pal_cat(nlevels(d[[colv]])), name = colv) }
  else g <- g + geom_point(colour = pk$m, alpha = .5, size = 1.8)
  g + geom_segment(data = ld, aes(x = 0, y = 0, xend = PC1, yend = PC2), inherit.aes = FALSE, arrow = arrow(length = unit(.2, "cm")), colour = pk$d, linewidth = .8) +
    geom_text(data = ld, aes(PC1, PC2, label = v), inherit.aes = FALSE, colour = pk$d, fontface = "bold", vjust = -.6) +
    labs(title = "PCA biplot", x = sprintf("PC1 (%.1f%%)", ve[1]), y = sprintf("PC2 (%.1f%%)", ve[2])) + theme_pink()
}

# ---- Scatterplot matrix, Mahalanobis, Q-Q Chi-Square ----
plot_mv_pairs <- function(d, vars, colv) {
  grp <- if (colv != NONE) d[[colv]] else NULL
  cols <- if (is.null(grp)) pk$m else pal_cat(nlevels(grp))[as.integer(grp)]
  p_pts  <- function(x, y, ...) { points(x, y, pch = 16, cex = .7, col = adjustcolor(cols, .55)); lines(lowess(x, y), col = pk$d, lwd = 1.6) }
  p_cor  <- function(x, y, ...) { usr <- par("usr"); on.exit(par(usr = usr)); par(usr = c(0, 1, 0, 1)); r <- cor(x, y); text(.5, .5, sprintf("r = %.2f", r), cex = 1 + 1.4 * abs(r), col = pk$d, font = 2) }
  p_hist <- function(x, ...) { usr <- par("usr"); on.exit(par(usr = usr)); par(usr = c(usr[1:2], 0, 1.5)); h <- hist(x, plot = FALSE)
  rect(h$breaks[-length(h$breaks)], 0, h$breaks[-1], h$counts / max(h$counts), col = pk$xl, border = "white") }
  pairs(d[, vars, drop = FALSE], lower.panel = p_pts, upper.panel = p_cor, diag.panel = p_hist, gap = .4, cex.labels = 1.3, font.labels = 2,
        oma = c(6, 4, 6, 4), main = "Scatterplot matrix (pairs plot)", col.main = pk$d, cex.main = 1.2)
  if (!is.null(grp)) {
    par(fig = c(0, 1, 0, 1), oma = c(0, 0, 0, 0), mar = c(0, 0, 0, 0), new = TRUE); plot.new()
    legend("bottom", legend = levels(grp), pch = 16, col = pal_cat(nlevels(grp)), horiz = TRUE, bty = "n", title = colv, inset = c(0, .005))
  }
}

mahal_fit <- function(d, vars) {
  X <- as.matrix(d[, vars, drop = FALSE]); p <- length(vars)
  D2 <- tryCatch(mahalanobis(X, colMeans(X), cov(X)), error = function(e) NULL)
  validate(need(!is.null(D2), "Matriks kovarians singular (ada variabel yang bergantung linear). Kurangi variabel yang dipilih."))
  cut <- qchisq(.975, df = p)
  list(df = data.frame(idx = seq_len(nrow(X)), D2 = D2, pval = pchisq(D2, p, lower.tail = FALSE), out = D2 > cut), cut = cut, p = p, X = X)
}
plot_mv_mahal_idx <- function(f) {
  df <- f$df; df$Status <- factor(ifelse(df$out, "Pencilan", "Normal"), levels = c("Normal", "Pencilan"))
  o <- df[df$out, ]; o <- head(o[order(-o$D2), ], 5)
  ggplot(df, aes(idx, D2, colour = Status)) + geom_point(alpha = .7, size = 1.6) + geom_hline(yintercept = f$cut, linetype = "dashed", colour = pk$d) +
    geom_text(data = o, aes(label = idx), vjust = -.6, size = 3, show.legend = FALSE) +
    scale_colour_manual(values = c(Normal = pk$l, Pencilan = pk$d), drop = FALSE) +
    labs(title = "Jarak Mahalanobis (D2) per observasi", subtitle = sprintf("Garis putus-putus = batas Chi-Square 97.5%% (df = %d): %.2f | Pencilan: %d (%.1f%%)", f$p, f$cut, sum(df$out), 100 * mean(df$out)),
         x = "Indeks observasi", y = "D2 (kuadrat jarak Mahalanobis)") + theme_pink(11)
}
plot_mv_mahal_sc <- function(f, vars) {
  df <- data.frame(x = f$X[, 1], y = f$X[, 2], Status = factor(ifelse(f$df$out, "Pencilan", "Normal"), levels = c("Normal", "Pencilan")))
  ggplot(df, aes(x, y, colour = Status, size = Status)) + geom_point(alpha = .65) +
    scale_colour_manual(values = c(Normal = pk$l, Pencilan = pk$d), drop = FALSE) + scale_size_manual(values = c(Normal = 1.6, Pencilan = 2.6), guide = "none") +
    labs(title = "Pencilan multivariat pada ruang dua variabel", subtitle = paste0("Ditampilkan: ", vars[1], " dan ", vars[2], " (dua variabel pertama yang dipilih)"), x = vars[1], y = vars[2]) + theme_pink(11)
}
mahal_table <- function(d, f, vars) {
  df <- f$df; o <- df[order(-df$D2), ][seq_len(min(10, nrow(df))), ]
  cbind(data.frame(Baris = o$idx, D2 = fmt(o$D2, 2), `p-value` = fmt_p(o$pval), Status = ifelse(o$out, "Pencilan", "Normal"), check.names = FALSE),
        setNames(lapply(vars, function(v) fmt(d[[v]][o$idx], 1)), vars))
}
plot_mv_qqchi <- function(f) {
  n <- nrow(f$df); o <- order(f$df$D2)
  s <- data.frame(D2 = f$df$D2[o], idx = f$df$idx[o], out = f$df$out[o]); s$q <- qchisq((seq_len(n) - .5) / n, df = f$p)
  s$Status <- factor(ifelse(s$out, "Pencilan", "Normal"), levels = c("Normal", "Pencilan"))
  ggplot(s, aes(q, D2, colour = Status)) + geom_point(alpha = .7, size = 1.8) + geom_abline(slope = 1, intercept = 0, colour = pk$d, linewidth = .9) +
    geom_text(data = tail(s, 3), aes(label = idx), vjust = -.6, size = 3, show.legend = FALSE) +
    scale_colour_manual(values = c(Normal = pk$m, Pencilan = pk$d), drop = FALSE) +
    labs(title = "Q-Q plot Chi-Square", subtitle = sprintf("D2 terurut vs kuantil Chi-Square (df = %d); titik di sekitar garis = mendekati normal multivariat", f$p),
         x = "Kuantil teoretis Chi-Square", y = "D2 terurut (sampel)") + theme_pink(11)
}

# =============================================================================
# 5. REGRESI LINEAR BERGANDA
# =============================================================================
build_model <- function(d, resp, preds, logy) {
  f <- paste0(if (logy) paste0("log(`", resp, "`)") else paste0("`", resp, "`"), " ~ ", paste0("`", preds, "`", collapse = " + "))
  lm(as.formula(f), data = d)
}
clean <- function(x) gsub("`", "", x)
term_of <- function(m, cn) {
  tl <- clean(attr(terms(m), "term.labels")); cn <- clean(cn)
  cand <- tl[startsWith(cn, tl)]; cand[which.max(nchar(cand))]
}
coef_label <- function(m, cn) {
  tl <- term_of(m, cn); cn2 <- clean(cn)
  if (cn2 == tl) tl else paste0(tl, " = ", substring(cn2, nchar(tl) + 1))
}
make_equation <- function(m, resp, logy) {
  b <- coef(m); lab <- vapply(names(b), function(n) if (n == "(Intercept)") "" else coef_label(m, n), "")
  lhs <- if (logy) paste0("ln(", resp, ")") else resp
  ln <- c(sprintf("%s-hat = %s", lhs, fmt(b[1], 4)),
          vapply(seq_along(b)[-1], function(i) sprintf("   %s %s x [%s]", if (b[i] >= 0) "+" else "-", fmt(abs(b[i]), 4), lab[i]), ""))
  paste(ln, collapse = "\n")
}
vif_table <- function(m) {
  tryCatch({
    mm <- model.matrix(m); asg <- attr(mm, "assign"); mm <- mm[, -1, drop = FALSE]; asg <- asg[-1]
    tl <- clean(attr(terms(m), "term.labels")); if (length(unique(asg)) < 2) return(NULL)
    R <- cor(mm); dR <- det(R)
    do.call(rbind, lapply(unique(asg), function(t) {
      i <- which(asg == t); gv <- det(R[i, i, drop = FALSE]) * det(R[-i, -i, drop = FALSE]) / dR
      data.frame(Variabel = tl[t], Df = length(i), `VIF setara` = gv^(1 / length(i)), check.names = FALSE)
    }))
  }, error = function(e) NULL)
}
bp_test <- function(m) {
  e <- resid(m); u <- e^2 / (sum(e^2) / length(e)); X <- model.matrix(m); ft <- lm.fit(X, u)
  stat <- sum((ft$fitted.values - mean(u))^2) / 2; df <- ncol(X) - 1
  c(stat = stat, df = df, p = pchisq(stat, df, lower.tail = FALSE))
}
reg_df <- function(m) {
  data.frame(fit = fitted(m), act = model.frame(m)[[1]], res = resid(m), rs = rstandard(m), cook = cooks.distance(m), idx = seq_along(resid(m)))
}
fit_table <- function(m) {
  s <- summary(m); f <- s$fstatistic; e <- resid(m)
  data.frame(Ukuran = c("Jumlah observasi (n)", "Derajat bebas residual", "R-kuadrat", "R-kuadrat terkoreksi (Adj. R2)", "Galat baku residual (sigma)", "RMSE", "MAE",
                        "Statistik F", "p-value (uji F simultan)", "AIC", "BIC"),
             Nilai = c(nobs(m), m$df.residual, fmt(s$r.squared, 4), fmt(s$adj.r.squared, 4), fmt(s$sigma, 3), fmt(sqrt(mean(e^2)), 3), fmt(mean(abs(e)), 3),
                       sprintf("%s (df1 = %d, df2 = %d)", fmt(f[1], 2), f[2], f[3]), fmt_p(pf(f[1], f[2], f[3], lower.tail = FALSE)), fmt(AIC(m), 2), fmt(BIC(m), 2)))
}
coef_table <- function(m) {
  s <- summary(m)$coefficients; ci <- confint(m)
  data.frame(Koefisien = vapply(rownames(s), function(n) if (n == "(Intercept)") "Intersep" else coef_label(m, n), ""),
             Estimasi = fmt(s[, 1], 4), `Galat baku` = fmt(s[, 2], 4), t = fmt(s[, 3], 3), `p-value` = fmt_p(s[, 4]),
             `CI 2.5%` = fmt(ci[, 1], 3), `CI 97.5%` = fmt(ci[, 2], 3), Sig. = stars(s[, 4]), check.names = FALSE, row.names = NULL)
}
anova_table <- function(m) {
  a <- anova(m); data.frame(Variabel = clean(rownames(a)), Df = a$Df, `Jumlah kuadrat` = fmt(a$`Sum Sq`, 1), `Kuadrat tengah` = fmt(a$`Mean Sq`, 1),
                            F = ifelse(is.na(a$`F value`), "", fmt(a$`F value`, 3)), `p-value` = ifelse(is.na(a$`Pr(>F)`), "", fmt_p(a$`Pr(>F)`)), check.names = FALSE)
}

plot_reg_avp <- function(m, logy, resp) {
  df <- reg_df(m); r2 <- summary(m)$r.squared; sc <- if (logy) paste0("ln(", resp, ")") else resp
  ggplot(df, aes(fit, act)) + geom_point(colour = pk$m, alpha = .45, size = 1.8) + geom_abline(slope = 1, intercept = 0, colour = pk$d, linetype = "dashed", linewidth = 1) +
    labs(title = "Nilai aktual vs prediksi", subtitle = sprintf("Garis putus-putus = prediksi sempurna | R2 = %.3f", r2), x = paste("Prediksi", sc), y = paste("Aktual", sc)) + theme_pink()
}
plot_reg_coef <- function(m) {
  s <- summary(m)$coefficients; ci <- confint(m); keep <- rownames(s) != "(Intercept)"
  df <- data.frame(lab = vapply(rownames(s)[keep], function(n) coef_label(m, n), ""), est = s[keep, 1], lo = ci[keep, 1], hi = ci[keep, 2], p = s[keep, 4])
  df$sig <- ifelse(df$p < .05, "Signifikan (p < 0.05)", "Tidak signifikan"); df$lab <- factor(df$lab, levels = rev(df$lab))
  ggplot(df, aes(est, lab, colour = sig)) + geom_vline(xintercept = 0, colour = pk$mauve, linetype = "dashed") + geom_errorbarh(aes(xmin = lo, xmax = hi), height = .25, linewidth = .8) + geom_point(size = 2.6) +
    scale_colour_manual(values = c("Signifikan (p < 0.05)" = pk$d, "Tidak signifikan" = "#B39DAB"), name = NULL) +
    labs(title = "Estimasi koefisien dan selang kepercayaan 95%", subtitle = "Selang yang memotong garis nol berarti tidak signifikan", x = "Estimasi koefisien", y = NULL) + theme_pink() + theme(legend.position = "bottom")
}
plot_reg_resid <- function(m) {
  df <- reg_df(m)
  ggplot(df, aes(fit, res)) + geom_point(colour = pk$m, alpha = .45, size = 1.6) + geom_hline(yintercept = 0, colour = pk$mauve, linetype = "dashed") +
    geom_smooth(method = "loess", formula = y ~ x, se = FALSE, colour = pk$d, linewidth = .9) +
    labs(title = "Residual vs nilai prediksi", subtitle = "Sebaran acak di sekitar nol = asumsi terpenuhi", x = "Nilai prediksi", y = "Residual") + theme_pink(11)
}
plot_reg_qq <- function(m) {
  df <- reg_df(m)
  ggplot(df, aes(sample = rs)) + stat_qq(colour = pk$m, alpha = .6) + stat_qq_line(colour = pk$d) +
    labs(title = "Normal probability plot (QQ-plot residual terstandar)", subtitle = "Titik mengikuti garis = residual mendekati normal", x = "Kuantil teoretis", y = "Kuantil sampel") + theme_pink(11)
}
plot_reg_scale <- function(m) {
  df <- reg_df(m)
  ggplot(df, aes(fit, sqrt(abs(rs)))) + geom_point(colour = pk$m, alpha = .45, size = 1.6) + geom_smooth(method = "loess", formula = y ~ x, se = FALSE, colour = pk$d, linewidth = .9) +
    labs(title = "Scale-Location", subtitle = "Garis datar = ragam residual konstan (homoskedastis)", x = "Nilai prediksi", y = "Akar |residual terstandar|") + theme_pink(11)
}
plot_reg_cook <- function(m) {
  df <- reg_df(m); thr <- 4 / nrow(df); top <- df[order(-df$cook), ][1:3, ]
  ggplot(df, aes(idx, cook)) + geom_segment(aes(xend = idx, yend = 0), colour = pk$l) + geom_hline(yintercept = thr, colour = pk$d, linetype = "dashed") +
    geom_text(data = top, aes(label = idx), vjust = -.4, colour = pk$d, size = 3.2) +
    labs(title = "Cook's distance", subtitle = sprintf("Garis putus-putus = 4/n = %.4f", thr), x = "Indeks observasi", y = "Cook's distance") + theme_pink(11)
}

# ---- Visualisasi regresi tambahan ----
reg_ord_res <- function(m, d) { o <- order(d$Date, seq_len(nrow(d))); list(res = unname(resid(m))[o], date = d$Date[o]) }

plot_reg_scatter <- function(d, resp, preds, logy) {
  nums <- preds[preds %in% num_all]
  validate(need(length(nums) >= 1, "Pilih minimal satu prediktor numerik untuk scatterplot."))
  yv <- if (logy) log(d[[resp]]) else d[[resp]]
  long <- do.call(rbind, lapply(nums, function(v) data.frame(Prediktor = v, x = d[[v]], y = yv)))
  ggplot(long, aes(x, y)) + geom_point(colour = pk$m, alpha = .35, size = 1.3) +
    geom_smooth(method = "lm", formula = y ~ x, se = TRUE, colour = pk$d, fill = pk$xl, linewidth = .9) +
    facet_wrap(vars(Prediktor), scales = "free_x") +
    labs(title = "Scatterplot respons terhadap prediktor numerik", subtitle = "Garis = regresi linear sederhana dengan pita kepercayaan 95%", x = "Nilai prediktor", y = if (logy) paste0("ln(", resp, ")") else resp) + theme_pink(11)
}
plot_reg_cor <- function(d, resp, preds) {
  vars <- unique(c(resp, preds[preds %in% num_all])); validate(need(length(vars) >= 2, "Pilih minimal satu prediktor numerik untuk matriks korelasi."))
  plot_mv_cor(d, vars, "pearson") + labs(title = "Correlation matrix (respons dan prediktor numerik)")
}
plot_reg_hist <- function(m) {
  df <- reg_df(m)
  ggplot(df, aes(res)) + geom_histogram(aes(y = after_stat(density)), bins = 30, fill = pk$xl, colour = "white") + geom_density(colour = pk$d, linewidth = 1) +
    stat_function(fun = dnorm, args = list(mean = mean(df$res), sd = sd(df$res)), colour = pk$slate, linetype = "dashed", linewidth = .9) +
    labs(title = "Histogram residual", subtitle = "Garis tebal = density residual; putus-putus = kurva normal teoretis", x = "Residual", y = "Densitas") + theme_pink(11)
}
plot_reg_ts <- function(m, d) {
  o <- reg_ord_res(m, d); n <- length(o$res); df <- data.frame(t = seq_len(n), res = o$res); br <- unique(round(seq(1, n, length.out = 6)))
  ggplot(df, aes(t, res)) + geom_line(colour = pk$l, linewidth = .4) + geom_point(colour = pk$m, size = .8, alpha = .5) + geom_hline(yintercept = 0, colour = pk$mauve, linetype = "dashed") +
    geom_smooth(method = "loess", formula = y ~ x, se = FALSE, colour = pk$d, linewidth = .9) +
    scale_x_continuous(breaks = br, labels = format(o$date[br], "%d-%m-%Y")) +
    labs(title = "Time series plot residual", subtitle = "Observasi diurutkan menurut tanggal transaksi; garis tebal = tren (loess)", x = "Tanggal transaksi (urutan observasi)", y = "Residual") + theme_pink(11)
}
plot_reg_acf <- function(m, d) {
  o <- reg_ord_res(m, d); n <- length(o$res); a <- acf(o$res, lag.max = min(30, n - 1), plot = FALSE)
  df <- data.frame(lag = as.numeric(a$lag)[-1], acf = as.numeric(a$acf)[-1]); ci <- qnorm(.975) / sqrt(n)
  ggplot(df, aes(lag, acf)) + geom_hline(yintercept = 0, colour = pk$mauve) + geom_segment(aes(xend = lag, yend = 0), colour = pk$m, linewidth = 1) + geom_point(colour = pk$d, size = 1.8) +
    geom_hline(yintercept = c(-ci, ci), colour = pk$slate, linetype = "dashed") +
    labs(title = "ACF plot residual", subtitle = "Garis putus-putus = batas 95%; lag 0 tidak ditampilkan", x = "Lag", y = "Autokorelasi") + theme_pink(11)
}
plot_reg_obsres <- function(m) {
  df <- reg_df(m)
  ggplot(df, aes(act, res)) + geom_point(colour = pk$m, alpha = .45, size = 1.6) + geom_hline(yintercept = 0, colour = pk$mauve, linetype = "dashed") +
    geom_smooth(method = "loess", formula = y ~ x, se = FALSE, colour = pk$d, linewidth = .9) +
    labs(title = "Observed vs residual", subtitle = "Pola menyebar/mengerucut menandakan ragam residual tidak konstan", x = "Nilai observasi (aktual)", y = "Residual") + theme_pink(11)
}

interp_reg <- function(m, resp, logy) {
  s <- summary(m); f <- s$fstatistic; pF <- pf(f[1], f[2], f[3], lower.tail = FALSE); n <- nobs(m)
  r2 <- s$r2 <- s$r.squared; ar2 <- s$adj.r.squared
  kat <- if (ar2 < .3) "lemah" else if (ar2 < .6) "sedang" else "kuat"
  out <- c(sprintf("**Uji simultan (uji F):** F = %s, p %s. %s", fmt(f[1], 2), if (pF < .001) "< 0.001" else paste("=", fmt_p(pF)),
                   if (pF < .05) "Minimal satu prediktor berpengaruh signifikan terhadap respons, sehingga model layak digunakan." else "Tidak ada prediktor yang berpengaruh signifikan secara simultan."),
           sprintf("**Kemampuan model:** R2 = %s dan Adj. R2 = %s, artinya sekitar **%s%%** keragaman %s dijelaskan oleh prediktor dalam model (kategori **%s**); sisanya dijelaskan faktor lain di luar model.", fmt(r2, 3), fmt(ar2, 3), fmt(100 * ar2, 1), if (logy) paste0("ln(", resp, ")") else resp, kat))
  co <- coef(s)[-1, , drop = FALSE]; sg <- co[co[, 4] < .05, , drop = FALSE]
  if (nrow(sg) == 0) out <- c(out, "**Uji parsial (uji t):** tidak ada koefisien yang signifikan pada taraf 5%.")
  else {
    out <- c(out, sprintf("**Uji parsial (uji t):** prediktor signifikan pada taraf 5%%: **%s**.", paste(vapply(rownames(sg), function(n) coef_label(m, n), ""), collapse = ", ")))
    for (nm in rownames(sg)) {
      b <- sg[nm, 1]; tl <- term_of(m, nm); isnum <- !(tl %in% names(m$xlevels))
      eff <- if (logy) sprintf("%s%%", fmt(100 * (exp(abs(b)) - 1), 1)) else fmt(abs(b), 3)
      verb <- if (b > 0) "naik" else "turun"
      out <- c(out, if (isnum) sprintf("Setiap kenaikan 1 satuan **%s**, rata-rata **%s** %s sebesar **%s**, dengan variabel lain dianggap tetap.", tl, resp, verb, eff)
               else sprintf("Dibanding kategori acuan **%s**, kategori **%s** memiliki rata-rata **%s** yang %s sebesar **%s**, dengan variabel lain dianggap tetap.", m$xlevels[[tl]][1], substring(clean(nm), nchar(tl) + 1), resp, if (b > 0) "lebih tinggi" else "lebih rendah", eff))
    }
  }
  tl_all <- unique(vapply(rownames(co), function(n) term_of(m, n), "")); tl_sig <- unique(vapply(rownames(sg), function(n) term_of(m, n), ""))
  ns <- setdiff(tl_all, tl_sig)
  if (length(ns) > 0) out <- c(out, sprintf("Variabel **tidak signifikan** (tidak ada bukti pengaruh setelah variabel lain dikendalikan): %s.", paste(ns, collapse = ", ")))
  sh <- shapiro.test(if (nobs(m) > 5000) sample(resid(m), 5000) else resid(m))$p.value; bp <- bp_test(m)["p"]; vt <- vif_table(m)
  out <- c(out, sprintf("**Asumsi residual:** normalitas (Shapiro-Wilk) p %s, sehingga %s; kesamaan ragam (Breusch-Pagan) p %s, sehingga %s.", pq(sh), if (sh < .05) "residual **tidak normal**" else "residual normal", pq(bp), if (bp < .05) "ragam residual **tidak konstan**" else "ragam konstan"))
  if (!is.null(vt)) out <- c(out, sprintf("**Multikolinearitas:** VIF setara maksimum = %s (%s).", fmt(max(vt[["VIF setara"]]), 2), if (max(vt[["VIF setara"]]) < 5) "aman, di bawah 5" else "perlu diperhatikan, di atas 5"))
  out
}

# =============================================================================
# KESIMPULAN OTOMATIS
# =============================================================================
auto_conclusion <- function(d, m, resp, logy) {
  s <- summary(m); co <- coef(s)[-1, , drop = FALSE]; sg <- rownames(co)[co[, 4] < .05]
  cm <- cor(d[, c("Quantity", "Total", "Rating", "Hour")]); diag(cm) <- 0
  i <- which(abs(cm) == max(abs(cm)), arr.ind = TRUE)[1, ]
  c(sprintf("Data yang dianalisis berisi **%d transaksi**; nilai transaksi (Total) berkisar %s sampai %s dengan median **%s** dan rata-rata %s (%s).",
            nrow(d), fmt(min(d$Total)), fmt(max(d$Total)), fmt(median(d$Total)), fmt(mean(d$Total)), if (skew_fun(d$Total) > .5) "distribusi miring ke kanan" else "distribusi relatif simetris"),
    sprintf("Pasangan variabel numerik dengan korelasi terkuat adalah **%s** dan **%s** (r = %s).", rownames(cm)[i[1]], colnames(cm)[i[2]], fmt(cm[i[1], i[2]], 3)),
    sprintf("Model regresi (respons **%s**%s) memiliki Adj. R2 = **%s**; prediktor signifikan pada taraf 5%%: **%s**.", resp, if (logy) ", transformasi ln" else "", fmt(s$adj.r.squared, 3),
            if (length(sg) == 0) "tidak ada" else paste(vapply(sg, function(n) coef_label(m, n), ""), collapse = ", ")))
}
# EDIT: sesuaikan rekomendasi berikut dengan hasil analisis akhir Anda.
REKOMENDASI <- c(
  "Jumlah unit per transaksi (Quantity) merupakan faktor utama nilai transaksi; promosi berbasis jumlah pembelian (paket bundling, potongan harga pembelian banyak) berpotensi menaikkan nilai transaksi.",
  "Cabang, tipe pelanggan, jenis kelamin, lini produk, dan metode pembayaran tidak menunjukkan pengaruh nyata terhadap nilai transaksi setelah Quantity dikendalikan, sehingga strategi dari sisi nilai transaksi tidak perlu dibedakan berdasarkan segmen tersebut.",
  "Rating tidak berkaitan dengan nilai transaksi maupun jumlah pembelian; peningkatan kepuasan pelanggan perlu ditelusuri melalui data tambahan seperti kualitas layanan dan waktu tunggu.",
  "Penelitian lanjutan disarankan menambahkan variabel harga satuan, periode data yang lebih panjang, serta model yang menangani ragam tidak konstan (transformasi atau regresi robust)."
)
KETERBATASAN <- c(
  "Data bersifat observasional dan hanya mencakup periode singkat (Januari-Maret) sehingga hasil tidak dapat digeneralisasi ke seluruh waktu.",
  "Regresi menunjukkan hubungan (asosiasi), bukan sebab-akibat.",
  "Variabel penting seperti harga satuan tidak tersedia pada data, sehingga sebagian keragaman nilai transaksi tidak dapat dijelaskan model.",
  "Residual cenderung tidak normal dan ragamnya tidak konstan (bertambah seiring Quantity); inferensi (uji t/F, selang kepercayaan) hanya pendekatan. Transformasi log atau regresi robust dapat dipertimbangkan.",
  "Model linear mengasumsikan efek aditif dan linear; interaksi antar variabel tidak dimodelkan."
)

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