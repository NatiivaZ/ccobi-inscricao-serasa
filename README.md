# Automação SIFAMA (ANTT) — Consulta de pagamento e inscrição SERASA

Suite de automação para o portal **SIFAMA** da **ANTT**, com interface gráfica integrada. Cobre duas rotinas operacionais frequentes no projeto **CCOBI – SERASA**:

1. **Consulta de pagamento** — obtém a **situação da dívida** (ex.: quitada, pendente) para cada auto de infração listado em planilha.  
2. **Inscrição na SERASA** — pesquisa cada auto, **marca automaticamente o checkbox** quando há **exatamente um** resultado, gera planilha de status e **mantém o navegador aberto** para você confirmar manualmente **“Incluir na SERASA”**.

---

## Sumário

- [Funcionalidades](#funcionalidades)  
- [Requisitos](#requisitos)  
- [Instalação](#instalação)  
- [Como usar](#como-usar)  
- [Planilha de entrada](#planilha-de-entrada)  
- [Planilhas de saída](#planilhas-de-saída)  
- [Fluxos detalhados](#fluxos-detalhados)  
- [Arquivos do projeto](#arquivos-do-projeto)  
- [Troubleshooting](#troubleshooting)  
- [Segurança](#segurança)

---

## Funcionalidades

| Módulo | O que faz |
|--------|-----------|
| **Consulta de pagamento** | Navega até a tela de consulta, informa o número do documento (auto), trata pop-up “nenhum registro”, lê o texto da **situação da dívida** e grava resultado em Excel/CSV. |
| **Inscrição SERASA** | Navega ao fluxo Serasa, pesquisa cada auto, seleciona checkbox só com **um** match; múltiplos resultados ou vazio recebem status explícito na planilha. **Não** envia a inclusão final sozinha — fica a cargo do operador. |
| **Interface** | Login dedicado, escolha de planilha, logs coloridos em tempo real, estatísticas, **pausar / continuar / parar**. |
| **Navegador** | Chrome em modo **anônimo**, opções anti-detecção básicas; validação de login pode usar **headless** antes da sessão visível. |

---

## Requisitos

- **Python** 3.8+  
- **Google Chrome** atualizado  
- **Selenium 4** (ChromeDriver gerenciado automaticamente na maioria dos setups)  
- **pandas** e **openpyxl** para leitura/gravação de planilhas  

```bash
pip install -r requirements.txt
```

---

## Instalação

```bash
cd "Inscrição SERASA"
pip install -r requirements.txt
python automacao_sifama_integrada.py
```

Script legado apenas consulta pagamento: `automacao_sifama.py` (sem o fluxo integrado completo da GUI unificada).

---

## Como usar

### 1. Login

1. Execute `automacao_sifama_integrada.py`.  
2. Informe **usuário** e **senha** do SIFAMA.  
3. Confirme; o sistema valida e abre a sessão automatizada.

### 2. Escolha do modo e da planilha

1. Selecione **Consulta de Pagamento** ou **Inscrição na SERASA**.  
2. Use **Selecionar Planilha** (`.xlsx` ou `.csv`).  
3. **Iniciar automação** e acompanhe os logs.

### 3. Encerramento

- **Consulta de pagamento:** ao finalizar, resultados são salvos e o fluxo pode encerrar o navegador (conforme implementação atual).  
- **Inscrição SERASA:** após processar todos os autos, **revise a planilha de resultado** no disco e só então clique em **Incluir na SERASA** no portal, se estiver de acordo com a seleção automática.

---

## Planilha de entrada

- **Coluna A** deve conter o cabeçalho aceito pelo código (ex.: `auto de infração` ou `Autos de infração` — o parser normaliza variações conforme implementado).  
- **Linhas seguintes:** um **número de auto** por linha.

Exemplo:

| Autos de infração |
|-------------------|
| 28797736          |
| EPSB200060412014  |

---

## Planilhas de saída

### Consulta de pagamento

- Arquivo típico: `Planilha Consulta Pagamento Resultado.xlsx` (ou `.csv`).  
- Colunas conceituais: auto informado × **situação da dívida** (ou mensagens como não encontrado / erro).

### Inscrição SERASA

- Arquivo típico: `Planilha Inscrição Serasa Resultado.xlsx` (ou `.csv`).  
- Status possíveis incluem: **SELECIONADO**, **NÃO ENCONTRADO NA CAIXA**, **MÚLTIPLOS RESULTADOS**, variações de erro de pesquisa, etc.

> Os nomes exatos dos arquivos e colunas seguem o que estiver definido no código-fonte (`automacao_sifama_integrada.py`).

---

## Fluxos detalhados

### Consulta de pagamento (visão técnica)

1. Login no `Login.aspx` da ANTT.  
2. Navegação por menus / possível intermédio **Portal de Sistemas**.  
3. Para cada auto: preencher campo, acionar **Gerar**, tratar pop-ups, localizar o elemento que exibe a **situação da dívida** (implementação atual usa seletor por estilo/layout específico do HTML).  
4. Persistir linha a linha e exportar planilha final.

### Inscrição SERASA (visão técnica)

1. Mesma base de login e navegação até o módulo **Serasa**.  
2. Para cada auto: buscar, analisar quantidade de checkboxes encontrados:  
   - **1** → marcar e registrar **SELECIONADO**;  
   - **0** → **NÃO ENCONTRADO NA CAIXA**;  
   - **>1** → **MÚLTIPLOS RESULTADOS** (nenhum marcado, evita erro de seleção indevida).  
3. Limpar/preparar campo para o próximo auto (ex.: atalhos de teclado).  
4. Salvar relatório; **navegador permanece aberto** para ação humana final.

---

## Arquivos do projeto

| Arquivo | Descrição |
|---------|-----------|
| `automacao_sifama_integrada.py` | **Principal** — GUI + ambos os fluxos |
| `automacao_sifama.py` | Versão focada em consulta de pagamento |
| `requirements.txt` | Dependências |
| `README.md` | Esta documentação |
| `INSTALACAO.md`, `DICAS_MELHORIAS.md`, `REFERENCIA_CODIGO.md` | Notas complementares no repositório |

---

## Troubleshooting

| Problema | Verificação |
|----------|-------------|
| Elemento não encontrado | Site ANTT alterado — inspecionar HTML e atualizar XPath/IDs nos logs. |
| Chrome / Driver | Atualizar Chrome; conferir mensagem de exceção do Selenium. |
| Login inválido | Credenciais, CAPTCHA (se houver), política de rede. |
| Planilha não salva | Permissão de pasta, arquivo aberto no Excel. |
| Timeouts | Rede lenta ou portal instável — retentativas estão no código, mas pode ser necessário ajustar tempos. |

---

## Segurança

- **Nunca** commite usuários, senhas ou planilhas reais com dados pessoais.  
- Automação de sistemas governamentais deve respeitar **termos de uso** e políticas internas.  
- Prefira repositório **privado** se o código expuser detalhes sensíveis de integração.

---

## Créditos

Desenvolvido para apoiar operações com **autos de infração** no ecossistema **SIFAMA / ANTT**, projeto **CCOBI – SERASA**.
