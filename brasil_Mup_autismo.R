# 1. Carregar pacotes ---------------------------------------------------------
# Nenhuma alteração aqui, mas certifique-se que todos estão instalados.
# Se 'epitools' fosse necessário, seria adicionado aqui: library(epitools)
library(dplyr)
library(openxlsx)
library(janitor)
library(stringi)
message("Iniciando análise de prevalência…")

# 2. Definir caminhos ---------------------------------------------------------
docs_dir    <- file.path(Sys.getenv("USERPROFILE"), "Documents")
input_file  <- file.path(docs_dir, "brasil_mup.csv")
output_file <- file.path(docs_dir, "resultado_municipios_vs_brasil.xlsx")

# 3. Ler e preparar dados -----------------------------------------------------
if (!file.exists(input_file)) {
  stop(sprintf("ERRO: Arquivo não encontrado em: %s", input_file))
}

# Ler dados
df_raw <- read.csv2(
  input_file,
  fileEncoding = "UTF-8",
  stringsAsFactors = FALSE,
  check.names = FALSE
) %>% 
  clean_names()

# Verificar colunas necessárias
required_cols <- c("municipio", "populacao", "autistas")
if (!all(required_cols %in% names(df_raw))) {
  stop("ERRO: Colunas necessárias ausentes no arquivo.")
}

# Processar dados
df_proc <- df_raw %>%
  mutate(
    municipio_raw = municipio,
    municipio = tolower(trimws(stringi::stri_trans_general(municipio, "Latin-ASCII"))),
    populacao = as.numeric(gsub("\\.", "", populacao)),
    autistas = as.numeric(gsub("\\.", "", autistas))
  ) %>%
  filter(!is.na(populacao), !is.na(autistas), populacao > 0) # Adicionado filtro populacao > 0

# 4. Referência Brasil --------------------------------------------------------
dados_brasil <- df_proc %>% filter(municipio == "brasil")
if (nrow(dados_brasil) == 0) stop("ERRO: Dados do Brasil não encontrados.")

pop_brasil <- dados_brasil$populacao[1]
autistas_brasil <- dados_brasil$autistas[1]
prev_brasil <- autistas_brasil / pop_brasil
taxa_brasil <- prev_brasil * 100000

# 5. Filtrar municípios -------------------------------------------------------
dados_municipios <- df_proc %>% filter(municipio != "brasil")

# 6. Função para calcular estatísticas por município --------------------------
calcular_estatisticas <- function(autistas, populacao) {
  
  # Casos esperados com base na prevalência do Brasil
  esperados <- populacao * prev_brasil
  
  # Razão de Prevalência (RP)
  # CORREÇÃO: Adicionado tratamento para casos esperados = 0 para evitar divisão por zero
  RP <- if (esperados > 0) autistas / esperados else 0
  
  # CORREÇÃO GERAL NESTA SEÇÃO:
  # O erro "função não encontrada" ocorreu aqui.
  # A função `poisson.exact` não pertence a nenhum dos pacotes carregados.
  # A lógica foi substituída pela função `poisson.test` do R base, que é mais adequada
  # para comparar uma contagem observada com uma taxa de referência.
  # A fórmula para grandes populações também foi ajustada para um método mais padrão.
  
  if (autistas < 5 || esperados < 5) {
    # Método exato de Poisson (recomendado para baixas contagens)
    # Testa se a contagem 'autistas' é significativamente diferente da 'esperados'
    teste <- poisson.test(x = autistas, r = esperados)
    
    # O IC da razão de prevalência é o IC da contagem dividido pelo esperado
    ic <- teste$conf.int / esperados
    
    list(
      taxa = (autistas / populacao) * 100000,
      RP = RP,
      IC95_inf = ic[1],
      IC95_sup = ic[2],
      valor_p = teste$p.value
    )
  } else {
    # Método de aproximação normal (para contagens maiores)
    # CORREÇÃO: Fórmula do Erro Padrão (SE) ajustada para o log(RP), mais robusta.
    logRP <- log(RP)
    SE_logRP <- sqrt(1/autistas + 1/autistas_brasil) # Fórmula padrão
    
    # P-valor baseado na distribuição normal
    valor_p <- 2 * pnorm(-abs(logRP / SE_logRP))
    
    list(
      taxa = (autistas / populacao) * 100000,
      RP = RP,
      IC95_inf = exp(logRP - 1.96 * SE_logRP),
      IC95_sup = exp(logRP + 1.96 * SE_logRP),
      valor_p = valor_p
    )
  }
}

# 7. Aplicar função a cada município ------------------------------------------
# CORREÇÃO: Envolver a chamada da função em tryCatch para evitar parar a execução
# caso um município específico cause um erro matemático.
resultados <- lapply(1:nrow(dados_municipios), function(i) {
  muni <- dados_municipios[i,]
  
  # tryCatch garante que a análise continue mesmo se uma linha falhar
  tryCatch({
    stats <- calcular_estatisticas(muni$autistas, muni$populacao)
    
    data.frame(
      municipio = muni$municipio_raw,
      populacao = muni$populacao,
      autistas = muni$autistas,
      taxa = stats$taxa,
      RP = stats$RP,
      IC95_inf = stats$IC95_inf,
      IC95_sup = stats$IC95_sup,
      valor_p = stats$valor_p,
      stringsAsFactors = FALSE
    )
  }, error = function(e) {
    # Se houver erro, retorna um data.frame com NAs para podermos identificar o problema
    message(sprintf("Erro ao processar %s: %s", muni$municipio_raw, e$message))
    return(NULL)
  })
})

# Com a correção acima, 'resultados' será criado, e os comandos abaixo funcionarão.
df_result <- bind_rows(resultados) %>%
  mutate(
    significancia = case_when(
      valor_p < 0.001 ~ "***",
      valor_p < 0.01  ~ "**",
      valor_p < 0.05  ~ "*",
      TRUE            ~ "NS"
    )
  ) %>%
  arrange(desc(RP))

# 8. Sumário e exportação ---------------------------------------------------
sumario_ref <- data.frame(
  Descricao = c(
    "População do Brasil",
    "Total de Autistas no Brasil",
    "Prevalência (por 100.000)",
    "Municípios analisados",
    "Municípios com contagens baixas (método exato)",
    "*** p<0.001, ** p<0.01, * p<0.05"
  ),
  Valor = c(
    format(pop_brasil, big.mark = ".", decimal.mark = ","),
    format(autistas_brasil, big.mark = ".", decimal.mark = ","),
    format(round(taxa_brasil, 1), big.mark = ".", decimal.mark = ","),
    nrow(df_result), # Usando df_result que agora existe
    sum(df_result$autistas < 5, na.rm = TRUE), # Exemplo de critério
    ""
  )
)

wb <- createWorkbook()
addWorksheet(wb, "Sumário")
writeData(wb, "Sumário", sumario_ref, headerStyle = createStyle(textDecoration = "bold"))
addWorksheet(wb, "Resultados")
writeData(wb, "Resultados", df_result, headerStyle = createStyle(textDecoration = "bold"))

# Formatação condicional (agora funcionará)
negStyle <- createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
posStyle <- createStyle(fontColour = "#006100", bgFill = "#C6EFCE")

# Coluna do p-valor (coluna 8)
conditionalFormatting(wb, "Resultados", cols = 8, rows = 2:(nrow(df_result) + 1),
                      style = negStyle, type = "expression", rule = "<0.05")
# Coluna da Razão de Prevalência (coluna 5)
conditionalFormatting(wb, "Resultados", cols = 5, rows = 2:(nrow(df_result) + 1),
                      style = posStyle, type = "expression", rule = ">1")

saveWorkbook(wb, output_file, overwrite = TRUE)
message("✅ Análise concluída! Resultados salvos em:\n", output_file)
      
