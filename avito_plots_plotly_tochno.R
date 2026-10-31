# =====================================================================
# Финмодель «Опыт владельцев» (Авито): графики plotly (R) по результатам Excel-модели.
# На выходе: PNG (для презентации) + HTML (интерактивные) + dashboard_all.html.
#
# Запуск:   Rscript avito_plots_plotly.R [путь_к_xlsx] [папка_вывода]
# По умолчанию: Avito_finmodel_Opyt_vladeltsev.xlsx и charts_R_plotly/
# Пакеты R: install.packages(c("readxl","dplyr","tidyr","plotly","htmlwidgets","htmltools"))
# PNG делается через установленный Chrome / Edge / Chromium (Python не нужен; найдётся сам,
# либо укажите путь: Sys.setenv(CHROME = "C:/Program Files/Google/Chrome/Application/chrome.exe")).
# Запасной вариант — Python + kaleido:  pip install "kaleido==0.2.1" "plotly==5.24.1"
# Если PNG не получились — HTML-файлы всё равно создаются (сохранить картинку можно кнопкой-камерой).
# Файл Excel должен быть пересчитан (сохранены значения формул).
# =====================================================================
suppressPackageStartupMessages({
  library(readxl); library(dplyr); library(tidyr); library(plotly); library(htmlwidgets); library(htmltools)
})


# Кириллица в названиях листов: нужна UTF-8 локаль (на Linux/macOS обычно уже есть)
invisible(suppressWarnings(try(Sys.setlocale("LC_CTYPE", "C.UTF-8"), silent = TRUE)))
args <- commandArgs(trailingOnly = TRUE)
xlsx <- if (length(args) >= 1) args[1] else "C:/Users/79161/Downloads/Avito_finmodel_Opyt_vladeltsev_tochno.xlsx"
out  <- if (length(args) >= 2) args[2] else "charts_R_plotly"
dir.create(out, showWarnings = FALSE, recursive = TRUE)

SC     <- c("Пессимист", "База", "Оптимист")
COLORS <- c("Пессимист" = "#D55E00", "База" = "#0072B2", "Оптимист" = "#009E73")  # Okabe-Ito
PAL    <- c("#4C78A8", "#F58518", "#54A24B", "#B279A2", "#9D755D", "#E45756")
LIMIT_Q <- 12                                                                     # лимит окупаемости по кейсу

# ---------------------------------------------------------------- чтение Excel
# Всё читаем как текст и ищем строки по подписи в колонке A — устойчиво к добавлению строк.
read_mat <- function(sheet) as.matrix(read_excel(xlsx, sheet = sheet, col_names = FALSE, col_types = "text", .name_repair = "minimal"))
M <- setNames(lapply(SC, function(s) read_mat(paste0("Модель_", s))), SC)

row_of <- function(m, label) {
  i <- which(m[, 1] == label)
  if (length(i) == 0) stop("Не найдена строка: ", label)
  i[1]
}
num     <- function(x) suppressWarnings(as.numeric(x))
q_row   <- function(m, label) num(m[row_of(m, label), 4:15])    # 12 кварталов: колонки D..O
y_row   <- function(m, label) num(m[row_of(m, label), 17:19])   # 3 года: колонки Q..S
scalar  <- function(m, label) m[row_of(m, label), 3]            # итоговый показатель: колонка C

periods <- M[["База"]][7, 4:15]                                 # «Q1 2027» ... «Q4 2029»
years   <- M[["База"]][7, 17:19]                                # «2027» ... «2029»

# ---------------------------------------------------------------- данные для графиков
quarterly <- bind_rows(lapply(SC, function(s) {
  m <- M[[s]]
  tibble(scenario = s, q = 1:12, period = periods,
         net = q_row(m, "Чистый денежный поток"), cum = q_row(m, "Накопленный денежный поток"),
         cumd = q_row(m, "Накопленный дисконтированный поток"))   # дисконтирование по околоключевой ставке (см. лист «Допущения»)
})) %>% mutate(scenario = factor(scenario, SC), period = factor(period, levels = periods))

kpi_labels <- c(
  pq   = "Окупаемость: квартал, когда накопленный поток ≥ 0",
  pqd  = "Дисконтированная окупаемость: квартал, когда накопл. дисконт. поток ≥ 0",
  peak = "Пиковая потребность в финансировании",
  trev = "Дополнительная выручка за 3 года",
  tcost = "Затраты за 3 года",
  tnet = "Накопленный чистый поток за 3 года",
  npv  = "NPV за 3 года",
  roi  = "ROI за 3 года (накопл. чистый поток / затраты)",
  be   = "Мин. прирост конверсии: простая окупаемость за 12 кв.",
  bed  = "Мин. прирост конверсии: дисконтированная окупаемость за 12 кв.",
  uplift = "Прирост конверсии в сделку в карточках с блоком (относительный)")

kpi <- bind_rows(lapply(SC, function(s) {
  v <- sapply(kpi_labels, function(l) num(scalar(M[[s]], l)))   # «не окупается» -> NA
  as_tibble(as.list(v)) %>% mutate(scenario = s, .before = 1)
})) %>% mutate(scenario = factor(scenario, SC))

stack_year <- function(mapping) {
  bind_rows(lapply(SC, function(s) {
    bind_rows(lapply(names(mapping), function(l)
      tibble(scenario = s, component = mapping[[l]], year = years, value = y_row(M[[s]], l))))
  })) %>%
    group_by(scenario, component, year) %>% summarise(value = sum(value), .groups = "drop") %>%
    mutate(scenario = factor(scenario, SC))
}
revenue <- stack_year(c(
  "Эффект 1: доход от конверсии (комиссия Доставки)" = "1. Конверсия в сделку",
  "Эффект 2: доход от продвижения новых объявлений"  = "2а. Продвижение объявлений",
  "Эффект 2: доход со сделок по новым объявлениям"   = "2б. Сделки по новым объявлениям",
  "Эффект 3: доход от рекламы и рекомендаций"        = "3. Визиты (реклама)",
  "Эффект 4: удержание (верхний сценарий)"           = "4. Удержание"))
costs <- stack_year(c(
  "Команда разработки" = "Команда", "ИИ-модерация" = "Модерация", "Ручная модерация спорных случаев" = "Модерация",
  "ИИ и инфраструктура" = "ИИ и инфраструктура", "Мотивация авторов (с годовым потолком)" = "Мотивация авторов",
  "Маркетинг (запуск + текущий)" = "Маркетинг", "Резерв" = "Резерв"))

# таблица чувствительности из листа «Сводка» (блок 3)
S  <- read_mat("Сводка")
r0 <- which(startsWith(ifelse(is.na(S[, 1]), "", S[, 1]), "3. Чувствительность"))[1] + 2
sens <- as_tibble(data.frame(uplift = num(S[r0:(r0 + 7), 1]), sapply(2:4, function(j) num(S[r0:(r0 + 7), j]))))
names(sens) <- c("uplift", SC)
sens <- sens %>% pivot_longer(-uplift, names_to = "scenario", values_to = "cum_net") %>% mutate(scenario = factor(scenario, SC))

# ---------------------------------------------------------------- оформление
FONT   <- list(family = "Arial", size = 13)
fmt_n  <- function(x, d = 0) formatC(x, format = "f", digits = d, big.mark = " ", decimal.mark = ",")
base_layout <- function(p, title, ..., margin = list(t = 70, b = 80, l = 70, r = 30)) layout(p, title = list(text = paste0("<b>", title, "</b>"), x = 0.01, xanchor = "left"),
  font = FONT, legend = list(orientation = "h", y = -0.18), margin = margin,
  paper_bgcolor = "white", plot_bgcolor = "white", ...)
grid <- list(gridcolor = "#E6E6E6", zerolinecolor = "#333333", zerolinewidth = 1.5)
col_of <- function(s) unname(COLORS[as.character(s)])

figs <- list(); dims <- list()
add_fig <- function(name, p, w = 1100, h = 600) { figs[[name]] <<- p; dims[[name]] <<- c(w, h) }

# 1. Накопленный денежный поток: сплошная — без дисконтирования, пунктир — дисконтированный, ромб — окупаемость
p <- plot_ly()
for (s in SC) {
  d <- filter(quarterly, scenario == s)
  p <- p %>%
    add_trace(data = d, x = ~period, y = ~cum, type = "scatter", mode = "lines+markers", name = paste0(s, " — без дисконт."),
              line = list(color = COLORS[[s]], width = 3), marker = list(color = COLORS[[s]], size = 6), legendgroup = s,
              hovertemplate = paste0(s, "<br>%{x}: %{y:.0f} млн ₽<extra></extra>")) %>%
    add_trace(data = d, x = ~period, y = ~cumd, type = "scatter", mode = "lines", name = paste0(s, " — дисконт. 15%"),
              line = list(color = COLORS[[s]], width = 2.5, dash = "dash"), legendgroup = s,
              hovertemplate = paste0(s, " (дисконт.)<br>%{x}: %{y:.0f} млн ₽<extra></extra>"))
}
pb <- pb_pts <- kpi %>% filter(!is.na(pq)) %>% select(scenario, q = pq) %>% left_join(quarterly, by = c("scenario", "q"))
p <- p %>% add_trace(data = pb_pts, x = ~period, y = ~cum, type = "scatter", mode = "markers", name = "Окупаемость",
                     marker = list(symbol = "diamond", size = 15, color = "white", line = list(color = "black", width = 2.5)),
                     hovertemplate = "Окупаемость: %{x}<extra></extra>") %>%
  base_layout("Накопленный чистый денежный поток по сценариям", xaxis = c(list(title = "", tickangle = -45, categoryorder = "array", categoryarray = periods), grid),
              yaxis = c(list(title = "млн ₽, нарастающим итогом"), grid), legend = list(orientation = "h", y = -0.3))
add_fig("cum_cashflow", p)

# 2. Ключевые показатели: NPV, накопленный поток, пик финансирования, ROI
kp <- list(list("npv", "NPV за 3 года, млн ₽", 1), list("tnet", "Накопленный поток, млн ₽", 1),
           list("peak", "Пиковая потребность, млн ₽", 1), list("roi", "ROI за 3 года, %", 100))
subs <- lapply(seq_along(kp), function(i) {
  k <- kp[[i]]; v <- kpi[[k[[1]]]] * k[[3]]
  plot_ly(x = as.character(kpi$scenario), y = v, type = "bar", marker = list(color = col_of(kpi$scenario)),
          text = fmt_n(v), textposition = "outside", cliponaxis = FALSE, showlegend = FALSE,
          hovertemplate = paste0("%{x}: %{y:.1f}<extra>", k[[2]], "</extra>")) %>%
    layout(xaxis = list(categoryorder = "array", categoryarray = SC),
           yaxis = c(list(title = ""), grid), annotations = list(list(text = paste0("<b>", k[[2]], "</b>"), x = 0.5, y = 1.1, xref = "paper", yref = "paper", showarrow = FALSE)))
})
p <- subplot(subs, nrows = 1, margin = 0.05, titleX = FALSE) %>% base_layout("Ключевые финансовые показатели за 3 года", showlegend = FALSE, margin = list(t = 120, b = 60, l = 60, r = 30))
add_fig("kpi", p, 1500, 520)

# 3. Окупаемость: простая и дисконтированная против лимита кейса
pay <- bind_rows(kpi %>% transmute(scenario, kind = "Простая", q = pq), kpi %>% transmute(scenario, kind = "Дисконтированная", q = pqd)) %>%
  mutate(q_plot = ifelse(is.na(q), 0, q), lab = ifelse(is.na(q), "нет", paste0(q, " кв.")))
p <- plot_ly()
for (kd in c("Простая", "Дисконтированная")) {
  d <- filter(pay, kind == kd)
  p <- p %>% add_trace(x = as.character(d$scenario), y = d$q_plot, type = "bar", name = kd, text = d$lab, textposition = "outside", cliponaxis = FALSE,
                       marker = list(color = col_of(d$scenario), opacity = ifelse(kd == "Простая", 1, 0.55), line = list(color = "white", width = 1)),
                       hovertemplate = paste0(kd, ": %{text}<extra></extra>"))
}
p <- p %>% base_layout("Срок окупаемости: простая (насыщенный цвет) и дисконтированная (бледный)", barmode = "group",
                       xaxis = list(categoryorder = "array", categoryarray = SC),
                       yaxis = c(list(title = "кварталов от старта", range = c(0, 14), dtick = 2), grid), showlegend = FALSE,
                       shapes = list(list(type = "line", x0 = 0, x1 = 1, xref = "paper", y0 = LIMIT_Q, y1 = LIMIT_Q, line = list(color = "red", dash = "dash", width = 2))),
                       annotations = list(list(x = 0, xref = "paper", y = LIMIT_Q + 0.6, text = "лимит кейса — 12 кв.", showarrow = FALSE, xanchor = "left", font = list(color = "red"))))
add_fig("payback", p, 900, 560)

# 4. Заложенный vs безубыточный прирост конверсии
p <- plot_ly()
for (t in list(list("uplift", "Заложено", "#BBBBBB"), list("be", "Минимум: простая окупаемость", "#222222"), list("bed", "Минимум: дисконтированная", "#777777"))) {
  v <- kpi[[t[[1]]]] * 100
  p <- p %>% add_trace(x = as.character(kpi$scenario), y = v, type = "bar", name = t[[2]], marker = list(color = t[[3]]),
                       text = paste0(fmt_n(v, 1), "%"), textposition = "outside", cliponaxis = FALSE,
                       hovertemplate = paste0(t[[2]], ": %{y:.1f}%<extra></extra>"))
}
p <- p %>% base_layout("Прирост конверсии: заложено vs минимум для окупаемости за 12 кв.", barmode = "group",
                       xaxis = list(categoryorder = "array", categoryarray = SC), yaxis = c(list(title = "%"), grid))
add_fig("breakeven", p, 1000, 560)

# 5. Чистый денежный поток по кварталам
p <- plot_ly()
for (s in SC) {
  d <- filter(quarterly, scenario == s)
  p <- p %>% add_trace(data = d, x = ~period, y = ~net, type = "bar", name = s, marker = list(color = COLORS[[s]]),
                       hovertemplate = paste0(s, "<br>%{x}: %{y:.1f} млн ₽<extra></extra>"))
}
p <- p %>% base_layout("Чистый денежный поток по кварталам", barmode = "group",
                       xaxis = list(title = "", tickangle = -45, categoryorder = "array", categoryarray = periods), yaxis = c(list(title = "млн ₽ за квартал"), grid))
add_fig("net_quarterly", p)

# 6–7. Структура выручки и затрат по годам (три панели — сценарии)
stacked <- function(df, title) {
  comps <- unique(df$component)
  panels <- lapply(seq_along(SC), function(i) {
    s <- SC[i]; pp <- plot_ly()
    for (j in seq_along(comps)) {
      d <- filter(df, scenario == s, component == comps[j])
      pp <- pp %>% add_trace(x = d$year, y = d$value, type = "bar", name = comps[j], legendgroup = comps[j], showlegend = (i == 1),
                             marker = list(color = PAL[(j - 1) %% length(PAL) + 1]),
                             hovertemplate = paste0(comps[j], ": %{y:.1f} млн ₽<extra>", s, "</extra>"))
    }
    pp %>% layout(barmode = "stack", yaxis = c(list(title = if (i == 1) "млн ₽ в год" else ""), grid),
                  annotations = list(list(text = paste0("<b>", s, "</b>"), x = 0.5, y = 1.06, xref = "paper", yref = "paper", showarrow = FALSE)))
  })
  subplot(panels, nrows = 1, shareY = TRUE, margin = 0.02, titleY = TRUE) %>% base_layout(title, barmode = "stack")
}
add_fig("revenue_structure", stacked(revenue, "Структура дополнительной выручки по эффектам"), 1300, 560)
add_fig("cost_structure",    stacked(costs,   "Структура затрат по статьям"), 1300, 560)

# 8. Чувствительность к приросту конверсии (кружки — принятые значения)
p <- plot_ly()
for (s in SC) {
  d <- filter(sens, scenario == s); k <- filter(kpi, scenario == s)
  p <- p %>% add_trace(x = d$uplift * 100, y = d$cum_net, type = "scatter", mode = "lines+markers", name = s, line = list(color = COLORS[[s]], width = 3),
                       marker = list(color = COLORS[[s]], size = 6), legendgroup = s,
                       hovertemplate = paste0(s, "<br>прирост %{x:.1f}%: %{y:.0f} млн ₽<extra></extra>")) %>%
    add_trace(x = k$uplift * 100, y = k$tnet, type = "scatter", mode = "markers+text", name = paste0(s, ": принято"), legendgroup = s, showlegend = FALSE,
              marker = list(size = 14, color = "white", line = list(color = COLORS[[s]], width = 3)),
              text = sprintf("%.0f%%", k$uplift * 100), textposition = "bottom center", hoverinfo = "skip")
}
p <- p %>% base_layout("Чувствительность накопленного потока к приросту конверсии",
                       xaxis = c(list(title = "Прирост конверсии в карточках с блоком, % (кружки — принятые в модели 4% / 5% / 8%)"), grid),
                       yaxis = c(list(title = "Накопленный поток за 3 года, млн ₽"), grid))
add_fig("sensitivity", p)

# ---------------------------------------------------------------- сохранение
# PNG: 1) через браузер Chrome/Edge/Chromium в headless-режиме (Python не нужен);
#      2) если браузер не найден — через Python + kaleido (pip install "kaleido==0.2.1" "plotly==5.24.1").
find_chrome <- function() {
  env <- Sys.getenv("CHROME", ""); if (nzchar(env) && file.exists(env)) return(env)
  cand <- c(Sys.which(c("google-chrome", "google-chrome-stable", "chromium", "chromium-browser", "chrome", "msedge")),
            "C:/Program Files/Google/Chrome/Application/chrome.exe",
            "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe",
            "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
            "C:/Program Files/Microsoft/Edge/Application/msedge.exe",
            file.path(Sys.getenv("LOCALAPPDATA"), "Google/Chrome/Application/chrome.exe"),
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
            "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
            "/Applications/Chromium.app/Contents/MacOS/Chromium")
  cand <- cand[nzchar(cand) & file.exists(cand)]
  if (length(cand)) unname(cand[1]) else ""
}
chrome <- find_chrome()
png_chrome <- function(nm, w, h) {
  w_ <- figs[[nm]] %>% layout(width = w, height = h) %>% config(displaylogo = FALSE, displayModeBar = FALSE)
  w_$width <- w; w_$height <- h                      # размер контейнера (иначе по умолчанию 400 px по высоте)
  w_$sizingPolicy$defaultWidth <- w; w_$sizingPolicy$defaultHeight <- h; w_$sizingPolicy$padding <- 0
  f  <- file.path(normalizePath(out), paste0(nm, "_tmp.html"))
  save_html(tagList(tags$head(tags$meta(charset = "utf-8"), tags$style("html,body{margin:0;padding:0;overflow:hidden}")), w_), f)
  png <- file.path(normalizePath(out), paste0(nm, ".png"))
  args <- c("--headless=new", "--disable-gpu", "--no-sandbox", "--hide-scrollbars", "--force-device-scale-factor=2",
            "--virtual-time-budget=8000", paste0("--window-size=", w, ",", h), paste0("--screenshot=", png),
            paste0("file:///", sub("^/", "", gsub("\\\\", "/", f))))
  system2(chrome, shQuote(args), stdout = FALSE, stderr = FALSE)
  unlink(f); unlink(sub("_tmp.html$", "_tmp_files", f), recursive = TRUE)
  file.exists(png)
}
png_kaleido <- function(nm, w, h) {
  tryCatch({ save_image(figs[[nm]], file.path(out, paste0(nm, ".png")), width = w, height = h, scale = 2); TRUE },
           error = function(e) FALSE)
}
if (!nzchar(chrome)) message("Chrome/Edge не найден — PNG попробую сделать через Python+kaleido. Можно задать путь: Sys.setenv(CHROME = '...')")
png_fail <- character()
for (nm in names(figs)) {
  w <- config(figs[[nm]], displaylogo = FALSE, toImageButtonOptions = list(format = "png", filename = nm, scale = 2))
  saveWidget(w, file.path(normalizePath(out), paste0(nm, ".html")), selfcontained = nzchar(Sys.which("pandoc")), title = nm)
  ok <- if (nzchar(chrome)) png_chrome(nm, dims[[nm]][1], dims[[nm]][2]) else FALSE
  if (!ok) ok <- png_kaleido(nm, dims[[nm]][1], dims[[nm]][2])
  if (!ok) png_fail <- c(png_fail, nm)
}
if (length(png_fail)) message("PNG не созданы для: ", paste(png_fail, collapse = ", "),
  "\nУстановите Google Chrome или Edge (скрипт найдёт сам) либо задайте путь: Sys.setenv(CHROME = 'путь/к/chrome').",
  "\nHTML-файлы созданы; PNG можно сохранить кнопкой с камерой на графике.")
page <- tagList(tags$head(tags$meta(charset = "utf-8"), tags$title("Финмодель: Опыт владельцев")),
                tags$div(style = "font-family:Arial;max-width:1500px;margin:auto", lapply(figs, function(w) tagList(w, tags$hr()))))
save_html(page, file.path(out, "dashboard_all.html"))
cat("Готово. Файлы в папке", normalizePath(out), ":\n"); print(list.files(out))
