# =====================================================================
# Финмодель «Опыт владельцев» (Авито): графики по результатам Excel-модели
# Статичные (ggplot2 -> PNG) и интерактивные (plotly -> HTML) версии.
#
# Запуск:   Rscript avito_plots.R [путь_к_xlsx] [папка_вывода]
# По умолчанию: Avito_finmodel_Opyt_vladeltsev.xlsx и charts_R/
# Пакеты:   install.packages(c("readxl","dplyr","tidyr","ggplot2","plotly",
#                              "scales","htmlwidgets","htmltools"))
# Файл должен быть пересчитан (в нём должны быть сохранены значения формул).
# =====================================================================
suppressPackageStartupMessages({
  library(readxl); library(dplyr); library(tidyr); library(ggplot2)
  library(plotly); library(scales); library(htmlwidgets); library(htmltools)
})

args <- commandArgs(trailingOnly = TRUE)
xlsx <- if (length(args) >= 1) args[1] else "Avito_finmodel_Opyt_vladeltsev.xlsx"
out  <- if (length(args) >= 2) args[2] else "charts_R"
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
  be   = "Безубыточный прирост конверсии (uplift), при котором накопл. поток Q12 = 0",
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
theme_set(theme_minimal(base_size = 12) +
            theme(plot.title = element_text(face = "bold"), legend.position = "bottom", panel.grid.minor = element_blank()))
sc_fill <- scale_fill_manual(values = COLORS, name = NULL)
sc_col  <- scale_color_manual(values = COLORS, name = NULL)
fmt_mln <- label_number(big.mark = " ", accuracy = 1)

# ---------------------------------------------------------------- графики (ggplot)
plots <- list()

# 1. Накопленный денежный поток + ромбы окупаемости
pb_pts <- kpi %>% filter(!is.na(pq)) %>% select(scenario, q = pq) %>% left_join(quarterly, by = c("scenario", "q"))
cum_long <- quarterly %>% select(scenario, period, `Без дисконтирования` = cum, `Дисконтированный (ставка 15%)` = cumd) %>%
  pivot_longer(c(`Без дисконтирования`, `Дисконтированный (ставка 15%)`), names_to = "kind", values_to = "cum") %>%
  mutate(grp = interaction(scenario, kind))
plots$cum_cashflow <- ggplot(cum_long, aes(period, cum, color = scenario, linetype = kind, group = grp,
                                           text = paste0(scenario, " (", kind, ")<br>", period, ": ", round(cum), " млн ₽"))) +
  geom_hline(yintercept = 0) + geom_line(linewidth = 1.2) +
  geom_point(data = filter(cum_long, kind == "Без дисконтирования"), size = 1.8) +
  geom_point(data = pb_pts, aes(period, cum), inherit.aes = FALSE, shape = 23, size = 5, stroke = 1.2, fill = NA) +
  scale_color_manual(values = COLORS, name = NULL) + scale_linetype_manual(values = c("solid", "dashed"), name = NULL) +
  scale_y_continuous(labels = fmt_mln) +
  labs(title = "Накопленный чистый денежный поток по сценариям", x = NULL, y = "млн ₽, нарастающим итогом",
       caption = "Ромб — квартал окупаемости (накопленный поток ≥ 0); пунктир — с дисконтированием") +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))

# 2. KPI: NPV, накопленный поток, пик финансирования, ROI
kpi_long <- kpi %>% transmute(scenario, `NPV за 3 года, млн ₽` = npv, `Накопленный поток, млн ₽` = tnet,
                              `Пиковая потребность, млн ₽` = peak, `ROI за 3 года, %` = roi * 100) %>%
  pivot_longer(-scenario, names_to = "metric", values_to = "value") %>%
  mutate(metric = factor(metric, unique(metric)))
plots$kpi <- ggplot(kpi_long, aes(scenario, value, fill = scenario, text = paste0(scenario, ": ", round(value, 1)))) +
  geom_hline(yintercept = 0) + geom_col(width = .65, show.legend = FALSE) +
  geom_text(aes(label = fmt_mln(value)), vjust = -0.4, size = 3.3) +
  facet_wrap(~metric, scales = "free_y", nrow = 1) + sc_fill +
  scale_y_continuous(expand = expansion(mult = c(.12, .15))) +
  labs(title = "Ключевые финансовые показатели за 3 года", x = NULL, y = NULL) +
  theme(axis.text.x = element_text(angle = 25, hjust = 1))

# 3. Окупаемость (простая и дисконтированная) против лимита кейса
pay <- bind_rows(
  kpi %>% transmute(scenario, kind = "Простая", q = pq),
  kpi %>% transmute(scenario, kind = "Дисконтированная", q = pqd)) %>%
  mutate(q_plot = ifelse(is.na(q), 0, q), lab = ifelse(is.na(q), "не окупается\nза 3 года", paste0(q, " кв.")),
         kind = factor(kind, c("Простая", "Дисконтированная")))
plots$payback <- ggplot(pay, aes(scenario, q_plot, fill = scenario, alpha = kind, text = paste0(kind, ": ", lab))) +
  geom_col(position = position_dodge(.75), width = .7, color = "white") +
  geom_text(aes(label = lab), position = position_dodge(.75), vjust = -0.3, size = 3.2, show.legend = FALSE) +
  geom_hline(yintercept = LIMIT_Q, linetype = "dashed", color = "red") +
  annotate("text", x = 0.55, y = LIMIT_Q + .5, label = "лимит кейса — 12 кв.", color = "red", hjust = 0, size = 3.5) +
  scale_fill_manual(values = COLORS, guide = "none") + scale_alpha_manual(values = c("Простая" = 1, "Дисконтированная" = 0.55), name = NULL) +
  scale_y_continuous(limits = c(0, 14), breaks = seq(0, 12, 2)) +
  labs(title = "Срок окупаемости: простая и дисконтированная", x = NULL, y = "кварталов от старта")

# 4. Заложенный vs безубыточный прирост конверсии
be_long <- kpi %>% transmute(scenario, `Заложено` = uplift * 100, `Нужно для безубыточности` = be * 100) %>%
  pivot_longer(-scenario, names_to = "type", values_to = "pct")
plots$breakeven <- ggplot(be_long, aes(scenario, pct, fill = type, text = paste0(type, ": ", round(pct, 1), "%"))) +
  geom_col(position = position_dodge(.7), width = .65) +
  geom_text(aes(label = sprintf("%.1f%%", pct)), position = position_dodge(.7), vjust = -0.3, size = 3.3) +
  scale_fill_manual(values = c("Заложено" = "#999999", "Нужно для безубыточности" = "#222222"), name = NULL) +
  scale_y_continuous(expand = expansion(mult = c(0, .12))) +
  labs(title = "Прирост конверсии: заложено vs нужно для выхода в ноль за 3 года", x = NULL, y = "%")

# 5. Чистый денежный поток по кварталам
plots$net_quarterly <- ggplot(quarterly, aes(period, net, fill = scenario,
                                             text = paste0(scenario, "<br>", period, ": ", round(net, 1), " млн ₽"))) +
  geom_hline(yintercept = 0) + geom_col(position = position_dodge(.8), width = .75) + sc_fill +
  scale_y_continuous(labels = fmt_mln) +
  labs(title = "Чистый денежный поток по кварталам", x = NULL, y = "млн ₽ за квартал") +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))

# 6–7. Структура выручки и затрат
stacked <- function(df, title) {
  ggplot(df, aes(year, value, fill = component, text = paste0(component, ": ", round(value, 1), " млн ₽"))) +
    geom_col(width = .7) + facet_wrap(~scenario) + scale_fill_manual(values = PAL, name = NULL) +
    scale_y_continuous(labels = fmt_mln) + labs(title = title, x = NULL, y = "млн ₽ в год")
}
plots$revenue_structure <- stacked(revenue, "Структура дополнительной выручки по эффектам")
plots$cost_structure    <- stacked(costs,   "Структура затрат по статьям")

# 8. Чувствительность к приросту конверсии
plots$sensitivity <- ggplot(sens, aes(uplift * 100, cum_net, color = scenario, group = scenario,
                                      text = paste0(scenario, "<br>uplift ", round(uplift * 100, 1), "%: ", round(cum_net), " млн ₽"))) +
  geom_hline(yintercept = 0) + geom_line(linewidth = 1.3) + geom_point(size = 2.2) + sc_col +
  scale_y_continuous(labels = fmt_mln) +
  labs(title = "Чувствительность накопленного потока к приросту конверсии",
       x = "Прирост конверсии в карточках с блоком, %", y = "Накопленный поток за 3 года, млн ₽")

# ---------------------------------------------------------------- сохранение
dev_png <- if (requireNamespace("ragg", quietly = TRUE)) ragg::agg_png else "png"
sizes <- list(kpi = c(14, 5), revenue_structure = c(12, 5), cost_structure = c(12, 5), breakeven = c(9, 5), payback = c(7, 5))
has_pandoc <- nzchar(Sys.which("pandoc")) || (requireNamespace("rmarkdown", quietly = TRUE) && rmarkdown::pandoc_available())

widgets <- list()
for (nm in names(plots)) {
  sz <- if (!is.null(sizes[[nm]])) sizes[[nm]] else c(10, 5.5)
  ggsave(file.path(out, paste0(nm, ".png")), plots[[nm]], width = sz[1], height = sz[2], dpi = 200, device = dev_png, bg = "white")
  w <- ggplotly(plots[[nm]], tooltip = "text") %>% layout(legend = list(orientation = "h", y = -0.2))
  widgets[[nm]] <- w
  tryCatch(saveWidget(w, file.path(getwd(), out, paste0(nm, ".html")), selfcontained = has_pandoc, title = nm),
           error = function(e) saveWidget(w, file.path(getwd(), out, paste0(nm, ".html")), selfcontained = FALSE, title = nm))
}

# единая страница со всеми интерактивными графиками (удобно открыть на встрече)
page <- tagList(tags$head(tags$meta(charset = "utf-8"), tags$title("Финмодель: Опыт владельцев")),
                tags$div(style = "font-family:Arial;max-width:1150px;margin:auto", lapply(widgets, function(w) tagList(w, tags$hr()))))
save_html(page, file.path(out, "dashboard_all.html"))

cat("Готово. Файлы в папке", normalizePath(out), ":\n"); print(list.files(out))
