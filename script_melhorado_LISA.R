
# 1. LIMPEZA DO AMBIENTE
rm(list = ls())

# 2. INSTALAÇÃO E CARREGAMENTO DE PACOTES
pacotes <- c("sf", "spdep", "ggplot2", "dplyr", "stringi", "openxlsx")
novos <- pacotes[!(pacotes %in% installed.packages()[, "Package"])]
if(length(novos)) install.packages(novos)
invisible(lapply(pacotes, library, character.only = TRUE))

# 3. CAMINHOS DOS ARQUIVOS
shapefile_path <- "C:/Users/Humberto Lago/Documents/autismo/BR_Municipios_2024/BR_Municipios_2024.shp"
csv_path       <- "C:/Users/Humberto Lago/Documents/autismo/autismo_municipios_simples_norm.csv"
saida_lisa_png <- "C:/Users/Humberto Lago/Documents/mapa_lisa_cluster_CORES_QUENTES.png"
saida_excel    <- "C:/Users/Humberto Lago/Documents/RELATORIO_ANALISE_AUTISMO_FINAL.xlsx"

# 4. VERIFICAÇÃO DE ARQUIVOS
if(!file.exists(csv_path)) stop("❌ CSV não encontrado: ", csv_path)
if(!file.exists(shapefile_path)) stop("❌ Shapefile não encontrado: ", shapefile_path)

# 5. LEITURA DOS DADOS
mapa  <- st_read(shapefile_path, quiet = TRUE)
dados <- read.csv(csv_path, stringsAsFactors = FALSE)

# 6. NORMALIZAÇÃO DE NOMES E CHAVE DE JUNÇÃO
dados <- dados %>%
  mutate(
    municipio_limpo = gsub(" \\(.*\\)$", "", municipio_norm),
    join_key = stri_trans_general(tolower(municipio_limpo), "Latin-ASCII")
  ) %>%
  distinct(join_key, .keep_all = TRUE)

mapa <- mapa %>%
  mutate(join_key = stri_trans_general(tolower(NM_MUN), "Latin-ASCII"))

# 7. JUNÇÃO ENTRE SHAPE E CSV
mapa_dados <- left_join(mapa, dados, by = "join_key")

# 8. DIAGNÓSTICO DO JOIN
total    <- nrow(mapa_dados)
com_dado <- sum(!is.na(mapa_dados$RP))
sem_dado <- total - com_dado
cat(sprintf("Total municípios: %d\n", total))
cat(sprintf("Com RP: %d\n", com_dado))
cat(sprintf("Sem RP: %d\n", sem_dado))
if(com_dado == 0) stop("❌ Nenhum município com RP após o join.")

# 9. FILTRAGEM PARA ANÁLISE
mapa_analise <- mapa_dados %>% filter(!is.na(RP) & is.finite(RP))

# 10. MATRIZ DE VIZINHANÇA
snap_dist <- 5000
nb <- poly2nb(mapa_analise, queen = TRUE, snap = snap_dist)
ilhas <- which(card(nb) == 0)
if(length(ilhas) > 0) warning("Existem ", length(ilhas), " ilhas; considere aumentar snap_dist.")
lw <- nb2listw(nb, style = "W", zero.policy = TRUE)

# 11. MORAN GLOBAL
moran_global <- moran.test(mapa_analise$RP, lw, zero.policy = TRUE)
print(moran_global)

# 12. LISA LOCAL
lisa_mat <- localmoran(mapa_analise$RP, lw, zero.policy = TRUE)
p_col <- grep("Pr\\(", colnames(lisa_mat), value = TRUE)

# 13. EXTRAÇÃO E ANÁLISE DE LISA
lag_rp  <- lag.listw(lw, mapa_analise$RP)
mean_rp <- mean(mapa_analise$RP, na.rm = TRUE)
li      <- lisa_mat[, "Ii"]
pvals   <- lisa_mat[, p_col]
mapa_analise <- mapa_analise %>%
  mutate(
    li      = li,
    lag_rp  = lag_rp,
    p_value = pvals,
    quadrante = case_when(
      RP > mean_rp & lag_rp > mean_rp & p_value <= 0.05 ~ "High-High",
      RP < mean_rp & lag_rp < mean_rp & p_value <= 0.05 ~ "Low-Low",
      RP > mean_rp & lag_rp < mean_rp & p_value <= 0.05 ~ "High-Low",
      RP < mean_rp & lag_rp > mean_rp & p_value <= 0.05 ~ "Low-High",
      TRUE ~ "Não-signif."
    )
  )

# 14. PLOTAGEM DO MAPA LISA
p <- ggplot(mapa_analise) +
  geom_sf(aes(fill = quadrante), color = NA) +
  scale_fill_manual(values = c(
    "High-High"   = "red",
    "Low-Low"     = "blue",
    "High-Low"    = "orange",
    "Low-High"    = "cyan",
    "Não-signif." = "grey80"
  )) +
  labs(
    title    = "LISA Cluster de RP - TEA",
    subtitle = paste0("Snap dist = ", snap_dist, " metros"),
    fill     = "Quadrante"
  ) +
  theme_minimal()

print(p)
ggsave(saida_lisa_png, plot = p, width = 10, height = 8, dpi = 300)

# 15. EXPORTAÇÃO PARA EXCEL
wb <- createWorkbook()

# Aba Resumo
addWorksheet(wb, "Resumo")
writeData(wb, "Resumo", data.frame(
  Moran_I = as.numeric(moran_global$estimate[["Moran I statistic"]]),
  p_value = moran_global$p.value
))

# Aba Dados LISA
addWorksheet(wb, "Dados LISA")
writeData(wb, "Dados LISA", mapa_analise %>%
  st_drop_geometry() %>%
  select(municipio = municipio_norm, RP, li, lag_rp, p_value, quadrante))

# Aba Municípios sem RP
addWorksheet(wb, "Sem RP")
writeData(wb, "Sem RP", mapa_dados %>%
  filter(is.na(RP)) %>%
  st_drop_geometry() %>%
  select(NM_MUN, join_key))

# Salvar planilha
saveWorkbook(wb, file = saida_excel, overwrite = TRUE)
cat("✅ Exportação concluída com sucesso para:", saida_excel, "\n")

