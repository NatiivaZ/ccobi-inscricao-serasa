# Guia de Instalação - Automação SIFAMA

## Passo a Passo

### 1. Instalar Python
- Baixe Python 3.8 ou superior em: https://www.python.org/downloads/
- Durante a instalação, marque a opção "Add Python to PATH"

### 2. Instalar Dependências
Abra o terminal/PowerShell na pasta do projeto e execute:
```bash
pip install -r requirements.txt
```

### 3. Instalar ChromeDriver

O ChromeDriver é necessário para controlar o navegador Chrome.

#### Opção A: Instalação Automática (Recomendado)
O Selenium 4.15+ pode baixar automaticamente. Se não funcionar, use a Opção B.

#### Opção B: Instalação Manual

1. **Verificar versão do Chrome:**
   - Abra o Chrome
   - Vá em: Menu (3 pontos) → Ajuda → Sobre o Google Chrome
   - Anote o número da versão (ex: 120.0.6099.109)

2. **Baixar ChromeDriver:**
   - Acesse: https://googlechromelabs.github.io/chrome-for-testing/
   - Ou: https://chromedriver.chromium.org/downloads
   - Baixe a versão compatível com seu Chrome

3. **Colocar ChromeDriver no PATH:**
   - **Windows:**
     - Extraia o arquivo `chromedriver.exe`
     - Coloque na mesma pasta do script `automacao_sifama.py`
     - OU adicione ao PATH do sistema
   
   - **Alternativa Windows (mais fácil):**
     - Coloque `chromedriver.exe` na pasta: `C:\Windows\System32\`

### 4. Testar Instalação

Execute o script:
```bash
python automacao_sifama.py
```

Se aparecer a interface gráfica, está tudo certo!

## Solução de Problemas

### Erro: "chromedriver not found"
- Certifique-se de que o ChromeDriver está no PATH ou na mesma pasta do script
- Verifique se a versão do ChromeDriver é compatível com seu Chrome

### Erro: "selenium not found"
- Execute: `pip install selenium`

### Erro: "pandas not found"
- Execute: `pip install pandas openpyxl`

### Erro ao fazer login
- Verifique se os IDs dos campos não mudaram (o sistema pode ter sido atualizado)
- Verifique sua conexão com a internet

### Navegador não abre
- Verifique se o Chrome está instalado
- Tente atualizar o ChromeDriver para a versão mais recente

## Dúvidas?

Se encontrar problemas, verifique:
1. Versão do Python (deve ser 3.8+)
2. Versão do Chrome e ChromeDriver (devem ser compatíveis)
3. Se todas as dependências foram instaladas corretamente

