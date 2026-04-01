# Automação SIFAMA - ANTT

Automação integrada para o sistema SIFAMA da ANTT (Agência Nacional de Transportes Terrestres) com duas funcionalidades principais.

## 🚀 Funcionalidades

### 1. Consulta de Pagamento (Situação da Dívida)
- Consulta automática da situação da dívida de cada auto de infração
- Extração da informação "Situação da Dívida" (Quitada, Pendente, etc.)
- Gera planilha com resultados: `Planilha Consulta Pagamento Resultado.xlsx` ou `.csv`

### 2. Inscrição de Autos na SERASA
- Pesquisa automática de autos para inscrição na SERASA
- Seleção automática de checkboxes (quando encontrado 1 resultado)
- Gera planilha com resultados antes de incluir na SERASA
- Navegador permanece aberto para você clicar manualmente em "Incluir na SERASA"

## ✨ Características

- ✅ Interface gráfica moderna e intuitiva
- ✅ Tela de login separada com validação
- ✅ Suporte para planilhas Excel (.xlsx) e CSV
- ✅ Navegação automática completa
- ✅ Tratamento robusto de erros
- ✅ Logs em tempo real com cores
- ✅ Estatísticas de sucessos/erros
- ✅ Controles de pausa, continuar e parar
- ✅ Modo anônimo do navegador
- ✅ Validação de login em modo headless (rápido e discreto)

## 📋 Requisitos

- Python 3.8 ou superior
- Google Chrome instalado
- ChromeDriver (gerenciado automaticamente pelo Selenium)

## 🔧 Instalação

1. Instale as dependências:
```bash
pip install -r requirements.txt
```

2. Execute o script principal:
```bash
python automacao_sifama_integrada.py
```

## 📖 Como Usar

### Passo 1: Login
1. Execute o script
2. Na tela de login, digite:
   - **Usuário**: Seu usuário do SIFAMA
   - **Senha**: Sua senha do SIFAMA
3. Clique em "Entrar"
4. O sistema validará suas credenciais (modo headless, rápido)

### Passo 2: Selecionar Automação e Planilha
1. Após login bem-sucedido, a tela principal será exibida
2. Selecione o tipo de automação:
   - **Consulta de Pagamento**: Para consultar situação da dívida
   - **Inscrição na SERASA**: Para selecionar autos para inscrição
3. Clique em "Selecionar Planilha" e escolha sua planilha
4. Clique em "Iniciar Automação"

### Passo 3: Acompanhar Processamento
- Acompanhe os logs em tempo real
- Veja estatísticas de sucessos e erros
- Use os botões para pausar, continuar ou parar se necessário

### Passo 4: Resultados

#### Consulta de Pagamento:
- Planilha gerada: `Planilha Consulta Pagamento Resultado.xlsx` (ou `.csv`)
- Coluna A: "Autos de infração"
- Coluna B: "Situação Divida" (Quitada, Pendente, NÃO ENCONTRADO, erro)

#### Inscrição na SERASA:
- Planilha gerada: `Planilha Inscrição Serasa Resultado.xlsx` (ou `.csv`)
- Coluna A: "Autos de infração"
- Coluna B: "Situação" (SELECIONADO, NÃO ENCONTRADO NA CAIXA, MÚLTIPLOS RESULTADOS, etc.)
- **Importante**: O navegador permanece aberto. Após verificar a planilha, clique manualmente em "Incluir na SERASA" quando quiser.

## 📊 Estrutura da Planilha de Entrada

A planilha deve ter:
- **Coluna A**: Cabeçalho "auto de infração" (ou "Autos de infração")
- **Linhas abaixo**: Números dos autos de infração (um por linha)

Exemplo:
```
| Autos de infração |
|-------------------|
| 28797736          |
| EPSB200060412014  |
| EPSA100015972014  |
```

## 🔄 Fluxo de Funcionamento

### Consulta de Pagamento:

1. **Login** → Validação headless → Login visível
2. **Navegação**:
   - Hover em menu → Clicar botão
   - Clicar "Portal de Sistemas" (se aparecer)
   - Repetir navegação inicial (se Portal foi clicado)
   - Hover em segundo menu → Clicar botão
3. **Para cada auto**:
   - Preencher campo "Nº do Documento"
   - Clicar em "Gerar"
   - Verificar popup "Nenhum registro encontrado" (se aparecer, clica em "Ok")
   - Extrair "Situação da Dívida" (busca div com estilo específico)
   - Limpar campo
4. **Salvar resultados** → Fechar navegador

### Inscrição na SERASA:

1. **Login** → Validação headless → Login visível
2. **Navegação**:
   - Hover em menu → Clicar botão
   - Clicar "Portal de Sistemas" (se aparecer)
   - Repetir navegação inicial (se Portal foi clicado)
   - Hover em "Serasa" → Clicar botão de inscrição
3. **Para cada auto**:
   - Preencher campo de busca
   - Clicar em "Pesquisar"
   - Verificar resultado:
     - **1 checkbox encontrado** → Clicar no checkbox → Marcar como "SELECIONADO"
     - **Múltiplos checkboxes** → Não clicar → Marcar como "MÚLTIPLOS RESULTADOS"
     - **Nenhum encontrado** → Marcar como "NÃO ENCONTRADO NA CAIXA"
   - Preparar próximo auto (CTRL+A → Digitar próximo auto)
4. **Salvar resultados** → **Navegador permanece aberto**
5. **Você clica manualmente** em "Incluir na SERASA" quando quiser

## ⚙️ Detalhes Técnicos

### Extração de Situação da Dívida
- Busca diretamente o `<div>` com estilo: `style="width:15.21mm;min-width: 15.21mm;"`
- Extrai o texto dentro do div (Quitada, Pendente, etc.)
- Não precisa buscar por "Situação da Dívida" primeiro

### Tratamento de Erros
- **Auto não encontrado**: Marca como "NÃO ENCONTRADO" ou "NÃO ENCONTRADO NA CAIXA"
- **Múltiplos resultados**: Marca como "MÚLTIPLOS RESULTADOS" (não seleciona nenhum)
- **Erro de conexão/timeout**: Marca como "ERRO AO PESQUISAR" ou "erro"
- **Popup detectado**: Clica automaticamente em "Ok" e marca como "NÃO ENCONTRADO"

### Navegação Inteligente
- Aguarda elementos aparecerem antes de clicar
- Delays configuráveis (0.8s para hovers, 2-3s para cliques principais)
- Timeout de 60 segundos para elementos críticos
- Tratamento de botão "Portal de Sistemas" opcional

## 📝 Observações Importantes

### Consulta de Pagamento:
- Se aparecer popup "Nenhum registro encontrado", o sistema clica automaticamente em "Ok"
- A extração busca o div específico com a situação da dívida
- Resultados são salvos em nova planilha para não sobrescrever a original

### Inscrição na SERASA:
- ⚠️ **O navegador NÃO fecha** após processar todos os autos
- Você deve verificar a planilha de resultados antes de clicar em "Incluir na SERASA"
- Apenas autos com exatamente 1 checkbox são selecionados automaticamente
- Autos com múltiplos resultados NÃO são selecionados (para evitar erros)
- Use CTRL+A e digite o próximo auto para navegar rapidamente entre autos

### Sistema Instável:
- O sistema SIFAMA pode ser instável
- A automação aguarda o carregamento das páginas
- Se houver timeout, tenta refazer a navegação
- Em caso de erro persistente, marca como erro e continua

## 🛠️ Solução de Problemas

### ChromeDriver não encontrado:
- O Selenium gerencia automaticamente o ChromeDriver
- Se houver problemas, verifique se o Chrome está atualizado

### Elementos não encontrados:
- O sistema pode ter sido atualizado (XPaths podem ter mudado)
- Verifique os logs para ver qual elemento não foi encontrado

### Login não funciona:
- Verifique se suas credenciais estão corretas
- O sistema valida primeiro em modo headless (rápido)
- Depois faz login no navegador visível

### Planilha não é salva:
- Verifique se você tem permissão de escrita na pasta
- Verifique se a planilha não está aberta em outro programa

## 📄 Arquivos do Projeto

- `automacao_sifama_integrada.py`: Script principal com interface gráfica integrada
- `automacao_sifama.py`: Script simples apenas para Consulta de Pagamento
- `requirements.txt`: Dependências do projeto
- `README.md`: Este arquivo

## 📞 Suporte

Em caso de problemas:
1. Verifique os logs na interface para identificar o erro
2. Verifique se o Chrome está atualizado
3. Verifique se sua conexão está estável
4. Verifique se os XPaths não mudaram (sistema pode ter sido atualizado)

---

**Desenvolvido para facilitar o trabalho com autos de infração no sistema SIFAMA da ANTT**
