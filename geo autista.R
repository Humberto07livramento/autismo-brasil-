# 📦 Carregar bibliotecas
library(sf)
library(dplyr)
library(ggplot2)
library(readr)
library(stringi)
library(writexl)

# === Caminhos dos arquivos ===
shapefile_path <- "C:/Users/Humberto Lago/Documents/autismo/BR_Municipios_2024/BR_Municipios_2024.shp"
csv_path       <- "C:/Users/Humberto Lago/Documents/autismo/BR_Municipios_2024/municipios_com_indice_normalizado.csv"
saida_png      <- "C:/Users/Humberto Lago/Documents/mapa_calor_municipios_clusters.png"
saida_excel    <- "C:/Users/Humberto Lago/Documents/mapa_calor_municipios_clusters.xlsx"

# === 1. Carregar os dados ===
br_municipios <- st_read(shapefile_path, quiet = TRUE)
dados_indices <- read_csv(csv_path, show_col_types = FALSE)

# === 2. Função para normalizar os nomes ===
normalizar_nome <- function(x) {
  x %>%
    stri_trans_general("Latin-ASCII") %>%
    tolower() %>%
    trimws() %>%
    sub(" \\(.*\\)", "", .)
}

# === 3. Normalização dos nomes ===
br_municipios <- br_municipios %>%
  mutate(nome_mun_norm = normalizar_nome(NM_MUN))

dados_indices <- dados_indices %>%
  mutate(municipio_norm = normalizar_nome(municipio)) %>%
  distinct(municipio_norm, .keep_all = TRUE)

# === 4. Juntar dados shapefile + CSV ===
mapa_completo <- br_municipios %>%
  left_join(dados_indices, by = c("nome_mun_norm" = "municipio_norm"))

# Verifique se há NA (municipios não casados corretamente)
# summary(mapa_completo$RP)  # Pode descomentar para ver

# === 5. Criar clusters com base nos quintis ===
mapa_completo <- mapa_completo %>%
  mutate(
    cluster = ntile(RP, 5),
    cluster_label = case_when(
      cluster == 1 ~ "Muito Baixo",
      cluster == 2 ~ "Baixo",
      cluster == 3 ~ "Médio",
      cluster == 4 ~ "Alto",
      cluster == 5 ~ "Muito Alto",
      TRUE         ~ "Indefinido"
    )
  )

# === 6. Criar gráfico ===
grafico <- ggplot(mapa_completo) +
  geom_sf(aes(fill = cluster_label), color = "gray30", size = 0.1) +
  scale_fill_manual(
    values = c(
      "Muito Baixo" = "#f7fcf0",
      "Baixo"       = "#ccebc5",
      "Médio"       = "#a8ddb5",
      "Alto"        = "#43a2ca",
      "Muito Alto"  = "#0868ac"
    ),
    name = "Grupo de Incidência"
  ) +
  labs(
    title = "Mapa de Calor por Município com Clusters",
    subtitle = "Classificação de municípios por faixas de índice RP",
    caption = "Fonte: Base própria + IBGE shapefile"
  ) +
  theme_minimal() +
  theme(
    axis.text = element_blank(),
    axis.title = element_blank(),
    panel.grid = element_blank()
  )

# === 7. Exibir gráfico ===
print(grafico)

# === 8. Salvar imagem PNG ===
ggsave(filename = saida_png, plot = grafico, width = 10, height = 12, dpi = 300)

# === 9. Exportar tabela para Excel (sem geometria) ===
write_xlsx(st_drop_geometry(mapa_completo), path = saida_excel)

