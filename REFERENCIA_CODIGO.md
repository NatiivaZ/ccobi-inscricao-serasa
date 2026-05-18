# Referência do Código — Inscrição SERASA

Este documento descreve a estrutura do `automacao_sifama_integrada.py` para manutenção e novas implementações.

---

## Visão geral

- **Arquivo principal:** `automacao_sifama_integrada.py`
- **Funcionalidades:** (1) Consulta de Pagamento (Situação da Dívida), (2) **Inscrição na SERASA**
- **Fluxo Inscrição SERASA:** Login → Navegar ao formulário SERASA → Para cada auto: pesquisar → verificar resultado → marcar checkbox → preparar próximo auto → ao final salvar planilha e manter navegador aberto.

---

## Estrutura de classes

| Classe | Responsabilidade |
|--------|------------------|
| `Logger` | Log com timestamp e callback para a GUI. |
| `BaseAutomacao` | Driver Chrome, login SIFAMA, `aguardar_pausa`, `verificar_erro_servidor`, `tratar_erro_servidor`, `_obter_texto_body`, `fechar`. |
| `AutomacaoConsultaPagamento` | Navegação consulta, `consultar_auto`, `extrair_situacao_divida`, `processar_autos`, salvar resultado. |
| `AutomacaoInscricaoSerasa` | Toda a lógica de **Inscrição SERASA** (ver seção abaixo). |
| `InterfaceGrafica` | Tkinter: login, seleção de planilha, tipo (consulta vs inscrição), iniciar/pausar/parar, reprocessar erros, gerar planilha. |

---

## AutomacaoInscricaoSerasa — Onde está cada coisa

### Navegação e formulário
- **`navegar_para_formulario()`** — Hover no menu, clica em Portal de Sistemas (se existir), depois Serasa → Inscrição. Aguarda `Corpo_txbAutoInfracao`. IDs: `ContentPlaceHolderCorpo_...`, `Corpo_txbAutoInfracao`, `Corpo_btnPesquisar`, `Corpo_gdvAutoInfracao`, `Corpo_btnIncluirSerasa`, `Corpo_btnLimpar`.

### Pesquisa de auto
- **`pesquisar_auto(numero_auto)`** — Preenche `Corpo_txbAutoInfracao` via JS, clica em `Corpo_btnPesquisar` (com **retry 3x** por StaleElementReferenceException), espera checkbox ou "Nenhum registro" (WebDriverWait 10s), `time.sleep(1.0)` para grid estabilizar. Retorna `True`/`False`. Trata erro de servidor e `StaleElementReferenceException` no `except` geral.
- **`verificar_resultado_pesquisa()`** — Espera grid ou "Nenhum registro", busca checkboxes com retry 3x (stale), retorna `("encontrado", 1)`, `("multiplos", n)`, `("nao_encontrado", 0)`, `("erro_servidor", 0)`, `("erro", 0)`.

### Checkbox e seleção
- **`_obter_checkbox_primeira_linha_dados()`** — Retorna elemento do checkbox (ID `Corpo_gdvAutoInfracao_ckSelecionar_0` ou XPath tbody). Não é usado no fluxo principal; o fluxo usa JS em `_clicar_checkbox_auto`.
- **`_clicar_checkbox_auto()`** — `time.sleep(0.3)`, JS para clicar em `#Corpo_gdvAutoInfracao_ckSelecionar_0`, `time.sleep(0.4)`.
- **`_checkbox_foi_validado()`** — Até 3 tentativas com `intervalo=0.3` s, verifica via JS se `ckSelecionar_0` está `checked`.

### Inscrição e limpeza
- **`selecionar_e_inscrever()`** — Chama `_clicar_checkbox_auto` (ou fallback JS), depois clica em `Corpo_btnIncluirSerasa`.
- **`aguardar_inscricao_completa()`** — Espera até 30s pelo sumiço de `Corpo_gdvAutoInfracao_ckSelecionar_0` (NoSuchElementException = concluído).
- **`limpar_formulario()`** — Clica em `Corpo_btnLimpar`, `time.sleep(0.4)`.

### Processamento em lote
- **`processar_autos(autos, progress_callback, stats_callback, error_handler)`** — Para cada auto: `pesquisar_auto` (com **1 retentativa** se falhar), `verificar_resultado_pesquisa`, trata erro_servidor/nao_encontrado/multiplos; se "encontrado": `_clicar_checkbox_auto` + `_checkbox_foi_validado`, adiciona resultado, `_preparar_proximo_auto(autos[idx+1])`. Não chama `selecionar_e_inscrever` nem `aguardar_inscricao_completa` — apenas marca checkboxes; "Incluir na SERASA" fica manual ou para outro fluxo.
- **`_preparar_proximo_auto(proximo_auto)`** — Define valor do campo `Corpo_txbAutoInfracao` via JS para o próximo auto.

### Resultados e planilha
- **`salvar_resultados(caminho_original, sufixo_arquivo=None)`** — Gera CSV ou Excel com colunas "Autos de infração" e "Situação", nome com data/hora e opcional sufixo (ex.: "Reprocessamento").
- **`_aplicar_formatacao_excel()`** — Cabeçalho azul, bordas, tabela, centralização.

---

## Interface gráfica (InterfaceGrafica)

- **Login:** `fazer_login()` — thread com driver **headless** para validar; se OK, guarda `automacao_temp` e mostra tela principal.
- **Iniciar:** `iniciar_automacao()` — cria **novo** driver visível (`automacao.criar_driver(headless=False)`), login, `navegar_para_formulario()`, `processar_autos()`, `salvar_resultados()`, para Inscrição SERASA **não** fecha o navegador.
- **Reprocessar erros:** `reprocessar_apenas_erros()` — usa o **mesmo** `self.automacao` (navegador já aberto), `processar_autos(ultimos_autos_com_erro)`, salva com sufixo "Reprocessamento".
- **Gerar planilha:** `gerar_planilha_resultado()` — chama de novo `automacao.salvar_resultados(planilha_path)` com timestamp atual.

---

## IDs e seletores importantes (Inscrição SERASA)

| Elemento | ID ou seletor |
|----------|----------------|
| Campo auto | `Corpo_txbAutoInfracao` |
| Botão Pesquisar | `Corpo_btnPesquisar` |
| Checkbox primeira linha | `Corpo_gdvAutoInfracao_ckSelecionar_0` |
| Grid checkboxes | CSS `[id^='Corpo_gdvAutoInfracao_ckSelecionar_']` |
| Botão Incluir SERASA | `Corpo_btnIncluirSerasa` |
| Botão Limpar | `Corpo_btnLimpar` |
| Barra que atrapalha cliques | `wings_process_presentation_dashboard_bar` (oculta via JS) |

---

## Possíveis erros e onde corrigir

- **StaleElementReferenceException na pesquisa:** `pesquisar_auto` já tem retry no clique do botão e `except StaleElementReferenceException` no fim; `processar_autos` tem 1 retentativa se `pesquisar_auto` retornar False.
- **Stale ao verificar checkboxes:** `verificar_resultado_pesquisa` faz 3 tentativas e, na última, chama `_buscar_checkboxes()` dentro de try/except para não propagar.
- **Atraso / grid não estável:** Ajustar `time.sleep(1.0)` após o wait de checkbox em `pesquisar_auto`; em `_clicar_checkbox_auto` os sleeps 0.3 e 0.4.
- **Erro de servidor:** `verificar_erro_servidor()` + `tratar_erro_servidor(tentar_navegar_novamente=True)` + `navegar_para_formulario()` e retentar o mesmo auto.

---

## Onde implementar coisas novas

- **Novo passo após marcar checkbox (ex.: clicar Incluir SERASA em lote):** em `processar_autos`, no bloco `elif status == "encontrado":`, após `_checkbox_foi_validado()`, chamar a nova função (ex. `selecionar_e_inscrever` e `aguardar_inscricao_completa` se quiser inscrição automática).
- **Novo delay ou timeout:** procurar por `time.sleep` e `WebDriverWait(..., 10)` / `5` nos métodos acima.
- **Novo tipo de automação (terceiro fluxo):** nova classe herdando `BaseAutomacao`, novo valor em `tipo_automacao` na GUI e em `iniciar_automacao` / `_executar_automacao`.
- **Alterar coluna da planilha de autos:** em `selecionar_planilha`, a coluna é detectada por nome contendo "auto" e "infração"; alterar o loop em `df.columns` se o nome mudar.
