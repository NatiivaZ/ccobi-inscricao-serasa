"""
Automação para preenchimento de autos de infração no SIFAMA
Sistema: ANTT - Agência Nacional de Transportes Terrestres
"""

import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pandas as pd
import openpyxl
from openpyxl import load_workbook
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path


class AutomacaoSIFAMA:
    def __init__(self):
        self.driver = None
        self.wait = None
        self.planilha_path = None
        self.dados_planilha = None
        self.resultados = []
        
    def criar_driver(self, headless=False):
        """Cria o driver do Chrome em modo anônimo/privado"""
        options = webdriver.ChromeOptions()
        options.add_argument('--incognito')  # Modo anônimo
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
        options.add_experimental_option('useAutomationExtension', False)
        
        # Suprimir erros e logs desnecessários do Chrome
        options.add_argument('--log-level=3')  # Suprime logs (0=INFO, 1=WARNING, 2=ERROR, 3=FATAL)
        options.add_argument('--disable-logging')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--silent')
        options.add_argument('--disable-gpu-logging')
        
        if headless:
            options.add_argument('--headless')
            options.add_argument('--disable-gpu')
        else:
            options.add_argument('--start-maximized')
        
        try:
            self.driver = webdriver.Chrome(options=options)
            self.wait = WebDriverWait(self.driver, 60)  # Timeout de 60 segundos
            return True
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao iniciar o navegador:\n{str(e)}\n\nCertifique-se de que o ChromeDriver está instalado.")
            return False
    
    def fazer_login(self, usuario, senha):
        """Realiza o login no sistema SIFAMA"""
        try:
            self.driver.get("https://appweb1.antt.gov.br/sca/Site/Login.aspx")
            time.sleep(2)
            
            # Tentar encontrar campos de login (IDs podem variar)
            campo_usuario = None
            campo_senha = None
            botao_entrar = None
            
            # Estratégia 1: Buscar por XPath específico (atualizado)
            try:
                campo_usuario = self.wait.until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="ContentPlaceHolderCorpo_ContentPlaceHolderCorpo_ContentPlaceHolderCorpo_ContentPlaceHolderCorpo_TextBoxUsuario"]'))
                )
                campo_senha = self.driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolderCorpo_ContentPlaceHolderCorpo_ContentPlaceHolderCorpo_ContentPlaceHolderCorpo_TextBoxSenha"]')
                botao_entrar = self.driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolderCorpo_ContentPlaceHolderCorpo_ContentPlaceHolderCorpo_ContentPlaceHolderCorpo_ButtonOk"]')
            except:
                # Estratégia 2: Buscar por name ou placeholder
                try:
                    campos_input = self.driver.find_elements(By.TAG_NAME, "input")
                    for campo in campos_input:
                        campo_id = campo.get_attribute("id") or ""
                        campo_type = campo.get_attribute("type") or ""
                        if "usuario" in campo_id.lower() or "user" in campo_id.lower():
                            campo_usuario = campo
                        elif "senha" in campo_id.lower() or "password" in campo_id.lower() or campo_type == "password":
                            campo_senha = campo
                    
                    botoes = self.driver.find_elements(By.TAG_NAME, "input")
                    for botao in botoes:
                        botao_type = botao.get_attribute("type") or ""
                        botao_value = botao.get_attribute("value") or ""
                        if botao_type == "submit" or "entrar" in botao_value.lower() or "login" in botao_value.lower():
                            botao_entrar = botao
                            break
                except:
                    pass
            
            if not campo_usuario or not campo_senha or not botao_entrar:
                raise Exception("Não foi possível encontrar os campos de login. Verifique se a página carregou corretamente.")
            
            campo_usuario.clear()
            campo_usuario.send_keys(usuario)
            campo_senha.clear()
            campo_senha.send_keys(senha)
            botao_entrar.click()
            
            # Aguardar login completar
            time.sleep(3)
            
            # Verificar se está na página principal (Portal de Sistemas)
            if "PortalSistemas" in self.driver.current_url or "Portal" in self.driver.title:
                return True
            return True  # Assumir sucesso se não houver erro
            
        except TimeoutException:
            messagebox.showerror("Erro", "Timeout ao fazer login. Verifique sua conexão.")
            return False
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao fazer login:\n{str(e)}")
            return False
    
    def navegar_para_formulario(self):
        """Navega através dos menus até chegar no formulário de consulta"""
        try:
            # Passo 1: Hover no primeiro elemento
            elemento_hover_1 = self.wait.until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="ContentPlaceHolderCorpo_ContentPlaceHolderMenu_MenuSistemasn0"]/table/tbody/tr/td/a'))
            )
            ActionChains(self.driver).move_to_element(elemento_hover_1).perform()
            time.sleep(0.8)  # Delay para submenu aparecer
            
            # Passo 2: Clicar no botão que aparece após hover
            botao_1 = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="ContentPlaceHolderCorpo_ContentPlaceHolderMenu_MenuSistemasn5"]/td/table/tbody/tr/td/a'))
            )
            botao_1.click()
            time.sleep(2)  # Aguardar página carregar
            
            # Após clicar nos dois primeiros botões, clicar no botão "Portal de Sistemas" se aparecer
            try:
                botao_portal = self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="ContentPlaceHolderCorpo_ContentPlaceHolderCorpo_ContentPlaceHolderCorpo_ContentPlaceHolderCorpo_btnPortalSistemas"]'))
                )
                botao_portal.click()
                time.sleep(2)
                
                # Após clicar no Portal de Sistemas, clicar novamente nos dois botões
                elemento_hover_1_novo = self.wait.until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="ContentPlaceHolderCorpo_ContentPlaceHolderMenu_MenuSistemasn0"]/table/tbody/tr/td/a'))
                )
                ActionChains(self.driver).move_to_element(elemento_hover_1_novo).perform()
                time.sleep(0.8)
                
                botao_1_novo = self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="ContentPlaceHolderCorpo_ContentPlaceHolderMenu_MenuSistemasn5"]/td/table/tbody/tr/td/a'))
                )
                botao_1_novo.click()
                time.sleep(2)
            except:
                # Se o botão não aparecer, continua normalmente
                pass
            
            # Passo 3: Hover no segundo elemento
            elemento_hover_2 = self.wait.until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="ContentPlaceHolderCorpo_ContentPlaceHolderMenu_menun2"]/table/tbody/tr/td/a'))
            )
            ActionChains(self.driver).move_to_element(elemento_hover_2).perform()
            time.sleep(0.8)  # Delay para submenu aparecer
            
            # Passo 4: Clicar no botão que aparece após hover
            botao_2 = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="ContentPlaceHolderCorpo_ContentPlaceHolderMenu_menun19"]/td/table/tbody/tr/td/a'))
            )
            botao_2.click()
            time.sleep(2)  # Aguardar página carregar
            
            # Aguardar campo de auto aparecer (confirma que chegou na página correta)
            self.wait.until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="ContentPlaceHolderCorpo_ContentPlaceHolderCorpo_ContentPlaceHolderCorpo_ContentPlaceHolderCorpo_txbAutoInfracao"]'))
            )
            return True
            
        except TimeoutException:
            return False
        except Exception as e:
            print(f"Erro na navegação: {str(e)}")
            return False
    
    def consultar_auto(self, numero_auto):
        """Preenche o campo com o número do auto e clica em Gerar"""
        try:
            # Limpar e preencher campo
            campo_auto = self.wait.until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="ContentPlaceHolderCorpo_ContentPlaceHolderCorpo_ContentPlaceHolderCorpo_ContentPlaceHolderCorpo_txbAutoInfracao"]'))
            )
            campo_auto.clear()
            campo_auto.send_keys(str(numero_auto))
            
            # Clicar em Gerar
            botao_gerar = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="ContentPlaceHolderCorpo_ContentPlaceHolderCorpo_ContentPlaceHolderCorpo_ContentPlaceHolderCorpo_btnGerar"]'))
            )
            botao_gerar.click()
            
            # Aguardar um pouco para o sistema processar
            time.sleep(2)
            
            # Verificar se aparece o popup "Nenhum registro encontrado" e clicar em Ok
            try:
                # Usar timeout curto para não atrasar quando não houver popup
                botao_ok = WebDriverWait(self.driver, 3).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="MessageBox_ButtonOk"]'))
                )
                print(f"Popup de 'Nenhum registro encontrado' detectado para auto {numero_auto}, clicando em Ok...")
                botao_ok.click()
                time.sleep(1)
            except:
                # Se não aparecer o popup, continua normalmente
                pass
            
            time.sleep(1)
            return True
            
        except TimeoutException:
            return False
        except Exception as e:
            print(f"Erro ao consultar auto {numero_auto}: {str(e)}")
            return False
    
    def extrair_situacao_divida(self):
        """Extrai a situação da dívida da página buscando diretamente o div com estilo específico"""
        try:
            # Aguardar página carregar completamente
            time.sleep(2)
            
            # Buscar diretamente o div com o estilo específico que contém a situação
            # Padrão: <div style="width:15.21mm;min-width: 15.21mm;">Quitada</div>
            # ou <div style="width:15.21mm;min-width: 15.21mm;">Pendente</div>
            try:
                # Buscar todos os divs com esse estilo específico
                divs = self.driver.find_elements(
                    By.XPATH, 
                    "//div[contains(@style, 'width:15.21mm') and contains(@style, 'min-width: 15.21mm')]"
                )
                
                # Pegar o primeiro div que tiver texto (não vazio)
                for div in divs:
                    texto = div.text.strip()
                    if texto:  # Se tiver conteúdo (não vazio)
                        print(f"Situação encontrada: {texto}")
                        return texto
                
                # Se não encontrou nenhum com texto, tentar estratégia alternativa
                print("Nenhum div com conteúdo encontrado, tentando alternativa...")
                
            except Exception as e:
                print(f"Erro ao buscar div com estilo específico: {str(e)}")
            
            # Estratégia alternativa: buscar por XPath mais específico
            try:
                div_situacao = self.driver.find_element(
                    By.XPATH, 
                    "//div[contains(@style, 'width:15.21mm') and contains(@style, 'min-width: 15.21mm') and normalize-space(text()) != '']"
                )
                situacao = div_situacao.text.strip()
                if situacao:
                    return situacao
            except:
                pass
            
            return None
            
        except Exception as e:
            print(f"Erro ao extrair situação: {str(e)}")
            return None
    
    def ler_planilha(self, caminho):
        """Lê a planilha Excel ou CSV"""
        try:
            caminho = Path(caminho)
            
            if caminho.suffix.lower() == '.csv':
                df = pd.read_csv(caminho, encoding='utf-8')
            else:
                df = pd.read_excel(caminho, engine='openpyxl')
            
            # Procurar coluna "auto de infração" (case insensitive)
            coluna_auto = None
            for col in df.columns:
                if 'auto' in str(col).lower() and 'infração' in str(col).lower():
                    coluna_auto = col
                    break
            
            if coluna_auto is None:
                # Se não encontrou, usar primeira coluna
                coluna_auto = df.columns[0]
            
            # Filtrar linhas vazias e pegar apenas os autos
            autos = df[coluna_auto].dropna().astype(str).tolist()
            # Remover cabeçalho se estiver na lista
            autos = [a for a in autos if a.lower() != 'auto de infração' and a.strip() != '']
            
            return autos, caminho
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler planilha:\n{str(e)}")
            return None, None
    
    def salvar_resultado(self, auto, situacao):
        """Salva o resultado na planilha"""
        self.resultados.append({
            'auto': str(auto),
            'situacao': str(situacao) if situacao else 'erro'
        })
    
    def salvar_planilha_resultado(self):
        """Salva os resultados na planilha com cabeçalhos corretos"""
        try:
            caminho = Path(self.planilha_path)
            
            # Criar DataFrame com os resultados
            dados = []
            for resultado in self.resultados:
                dados.append({
                    'Autos de infração': resultado['auto'],
                    'Situação Divida': resultado['situacao']
                })
            
            df_resultado = pd.DataFrame(dados)
            
            if caminho.suffix.lower() == '.csv':
                # Para CSV, criar novo arquivo
                novo_caminho = caminho.parent / "Planilha Consulta Pagamento Resultado.csv"
                df_resultado.to_csv(novo_caminho, index=False, encoding='utf-8')
            else:
                # Para Excel
                novo_caminho = caminho.parent / "Planilha Consulta Pagamento Resultado.xlsx"
                df_resultado.to_excel(novo_caminho, index=False, engine='openpyxl')
            
            return novo_caminho
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar resultados:\n{str(e)}")
            return None
    
    def processar_autos(self, autos, progress_callback=None):
        """Processa todos os autos da planilha"""
        total = len(autos)
        sucessos = 0
        erros = 0
        
        for idx, auto in enumerate(autos, 1):
            auto = str(auto).strip()
            if not auto:
                continue
            
            try:
                if progress_callback:
                    progress_callback(f"Processando auto {idx}/{total}: {auto}")
                
                # Consultar auto
                if not self.consultar_auto(auto):
                    self.salvar_resultado(auto, 'erro')
                    erros += 1
                    # Tentar refazer navegação se falhou
                    if not self.navegar_para_formulario():
                        messagebox.showerror("Erro", "Não foi possível navegar até o formulário. Verifique sua conexão.")
                        break
                    continue
                
                # Aguardar resultado aparecer (sistema instável)
                max_tentativas = 10
                situacao = None
                for tentativa in range(max_tentativas):
                    time.sleep(2)
                    situacao = self.extrair_situacao_divida()
                    if situacao:
                        break
                    if tentativa == max_tentativas - 1:
                        # Última tentativa: verificar se página carregou
                        if "erro" in self.driver.page_source.lower() or "não encontrado" in self.driver.page_source.lower():
                            situacao = "erro"
                            break
                
                if not situacao:
                    situacao = 'erro'
                
                self.salvar_resultado(auto, situacao)
                
                if situacao != 'erro':
                    sucessos += 1
                else:
                    erros += 1
                
                # Limpar campo para próximo auto
                try:
                    campo_auto = self.driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolderCorpo_ContentPlaceHolderCorpo_ContentPlaceHolderCorpo_ContentPlaceHolderCorpo_txbAutoInfracao"]')
                    campo_auto.clear()
                except:
                    # Se não conseguir limpar, refazer navegação
                    if not self.navegar_para_formulario():
                        break
                
            except Exception as e:
                print(f"Erro ao processar auto {auto}: {str(e)}")
                self.salvar_resultado(auto, 'erro')
                erros += 1
                
                # Tentar refazer navegação
                try:
                    self.navegar_para_formulario()
                except:
                    pass
        
        return sucessos, erros
    
    def fechar(self):
        """Fecha o navegador"""
        if self.driver:
            self.driver.quit()


class InterfaceGrafica:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Automação SIFAMA - ANTT")
        self.root.geometry("450x350")  # Menor para tela de login
        self.root.resizable(False, False)
        
        self.automacao = AutomacaoSIFAMA()
        self.planilha_path = None
        self.autos = []
        self.logado = False
        self.usuario_logado = None
        self.senha_logada = None
        
        # Frames principais
        self.frame_login = None
        self.frame_principal = None
        
        self.criar_interface()
    
    def criar_interface(self):
        """Cria a interface gráfica com telas separadas"""
        # Container principal
        self.main_container = tk.Frame(self.root, bg='#f0f0f0')
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Criar tela de login
        self.criar_tela_login()
        
        # Criar tela principal (inicialmente escondida)
        self.criar_tela_principal()
        
        # Mostrar tela de login inicialmente
        self.mostrar_tela_login()
    
    def criar_tela_login(self):
        """Cria a tela de login - simples e compacta"""
        self.frame_login = tk.Frame(self.main_container, bg='#f0f0f0')
        
        # Título simples
        titulo = tk.Label(self.frame_login, text="Automação SIFAMA - ANTT", 
                         font=("Arial", 16, "bold"), bg='#f0f0f0', fg='#2c3e50')
        titulo.pack(pady=30)
        
        # Frame de login compacto
        frame_login_box = tk.Frame(self.frame_login, bg='#ffffff', relief=tk.RAISED, bd=1)
        frame_login_box.pack(padx=100, pady=20)
        
        login_inner = tk.Frame(frame_login_box, bg='#ffffff')
        login_inner.pack(padx=30, pady=25)
        
        # Usuário
        tk.Label(login_inner, text="Usuário:", font=("Arial", 10), 
                bg='#ffffff').grid(row=0, column=0, sticky="w", pady=8, padx=5)
        self.entry_usuario = tk.Entry(login_inner, width=25, font=("Arial", 10))
        self.entry_usuario.grid(row=0, column=1, pady=8, padx=5, ipady=3)
        self.entry_usuario.focus()
        
        # Senha
        tk.Label(login_inner, text="Senha:", font=("Arial", 10), 
                bg='#ffffff').grid(row=1, column=0, sticky="w", pady=8, padx=5)
        self.entry_senha = tk.Entry(login_inner, width=25, show="*", font=("Arial", 10))
        self.entry_senha.grid(row=1, column=1, pady=8, padx=5, ipady=3)
        
        # Botão Entrar
        btn_entrar_frame = tk.Frame(login_inner, bg='#ffffff')
        btn_entrar_frame.grid(row=2, column=0, columnspan=2, pady=15)
        
        self.btn_entrar = tk.Button(btn_entrar_frame, text="Entrar", 
                                    command=self.fazer_login, bg='#27ae60', fg='white',
                                    font=("Arial", 10, "bold"), width=15, height=1)
        self.btn_entrar.pack()
        
        # Bind Enter key
        self.entry_senha.bind('<Return>', lambda e: self.fazer_login())
        self.entry_usuario.bind('<Return>', lambda e: self.entry_senha.focus())
        
        # Status do login
        self.label_status_login = tk.Label(login_inner, text="", 
                                           font=("Arial", 9), bg='#ffffff', fg='#e74c3c')
        self.label_status_login.grid(row=3, column=0, columnspan=2, pady=5)
    
    def criar_tela_principal(self):
        """Cria a tela principal após login"""
        self.frame_principal = tk.Frame(self.main_container, bg='#f0f0f0')
        
        # Título
        titulo_frame = tk.Frame(self.frame_principal, bg='#2c3e50', height=60)
        titulo_frame.pack(fill=tk.X, pady=(0, 10))
        titulo_frame.pack_propagate(False)
        
        titulo_container = tk.Frame(titulo_frame, bg='#2c3e50')
        titulo_container.pack(fill=tk.BOTH, expand=True)
        
        titulo = tk.Label(titulo_container, text="Automação SIFAMA - ANTT", 
                         font=("Arial", 18, "bold"), bg='#2c3e50', fg='white')
        titulo.pack(side=tk.LEFT, padx=20, pady=15)
        
        # Label de usuário logado
        self.label_usuario_logado = tk.Label(titulo_container, text="", 
                                            font=("Arial", 10), bg='#2c3e50', fg='#ecf0f1')
        self.label_usuario_logado.pack(side=tk.RIGHT, padx=20)
        
        # Botão sair
        btn_sair = tk.Button(titulo_container, text="Sair", 
                            command=self.sair, bg='#e74c3c', fg='white',
                            font=("Arial", 9), relief=tk.RAISED, cursor='hand2', width=10)
        btn_sair.pack(side=tk.RIGHT, padx=10)
        
        # Frame de planilha
        frame_planilha = tk.LabelFrame(self.frame_principal, text="Planilha", 
                                       font=("Arial", 10, "bold"), bg='#f0f0f0')
        frame_planilha.pack(fill=tk.X, pady=10, padx=20)
        
        self.label_planilha = tk.Label(frame_planilha, text="Nenhuma planilha selecionada", 
                                       fg="gray", font=("Arial", 9), bg='#f0f0f0')
        self.label_planilha.pack(pady=10)
        
        btn_selecionar = tk.Button(frame_planilha, text="📁 Selecionar Planilha", 
                                   command=self.selecionar_planilha, bg='#3498db', fg='white',
                                   font=("Arial", 9), relief=tk.RAISED, cursor='hand2', width=20)
        btn_selecionar.pack(pady=10)
        
        # Barra de progresso
        self.progress_var = tk.StringVar(value="Aguardando...")
        self.label_progress = tk.Label(self.frame_principal, textvariable=self.progress_var, 
                                      fg="blue", font=("Arial", 9), bg='#f0f0f0')
        self.label_progress.pack(pady=10)
        
        self.progress_bar = ttk.Progressbar(self.frame_principal, mode='indeterminate')
        self.progress_bar.pack(pady=5, padx=20, fill="x")
        
        # Botões
        frame_botoes = tk.Frame(self.frame_principal, bg='#f0f0f0')
        frame_botoes.pack(pady=20)
        
        self.btn_iniciar = tk.Button(frame_botoes, text="▶ Iniciar Automação", 
                                     command=self.iniciar_automacao, bg='#27ae60', fg='white',
                                     font=("Arial", 10, "bold"), width=20, height=2)
        self.btn_iniciar.pack(side=tk.LEFT, padx=5)
        
        self.btn_fechar = tk.Button(frame_botoes, text="Fechar", command=self.fechar, 
                                    bg='#e74c3c', fg='white', width=15, height=2)
        self.btn_fechar.pack(side=tk.LEFT, padx=5)
    
    def mostrar_tela_login(self):
        """Mostra a tela de login"""
        if self.frame_principal:
            self.frame_principal.pack_forget()
        if self.frame_login:
            self.frame_login.pack(fill=tk.BOTH, expand=True)
    
    def mostrar_tela_principal(self):
        """Mostra a tela principal"""
        # Ajustar tamanho da janela para tela principal
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        if self.frame_login:
            self.frame_login.pack_forget()
        if self.frame_principal:
            self.frame_principal.pack(fill=tk.BOTH, expand=True)
    
    def fazer_login(self):
        """Realiza o login"""
        usuario = self.entry_usuario.get().strip()
        senha = self.entry_senha.get().strip()
        
        if not usuario or not senha:
            self.label_status_login.config(text="Preencha usuário e senha.", fg='#e74c3c')
            return
        
        # Desabilitar botão durante login
        self.btn_entrar.config(state=tk.DISABLED, text="Entrando...")
        self.label_status_login.config(text="Conectando...", fg='#3498db')
        self.root.update()
        
        # Testar login em thread separada para não travar interface
        def testar_login_thread():
            try:
                # Criar driver em modo headless (sem abrir janela) para validação
                if not self.automacao.criar_driver(headless=True):
                    self.root.after(0, lambda: self.label_status_login.config(
                        text="Erro ao iniciar navegador. Verifique o ChromeDriver.", fg='#e74c3c'))
                    self.root.after(0, lambda: self.btn_entrar.config(state=tk.NORMAL, text="Entrar"))
                    return
                
                if self.automacao.fazer_login(usuario, senha):
                    # Login bem-sucedido - NÃO FECHAR navegador (manter aberto)
                    self.logado = True
                    self.usuario_logado = usuario
                    self.senha_logada = senha
                    # Navegador headless continua aberto (não fechar)
                    
                    self.root.after(0, lambda: self.label_usuario_logado.config(text=f"Usuário: {usuario}"))
                    self.root.after(0, self.mostrar_tela_principal)
                else:
                    # Login falhou
                    self.automacao.fechar()
                    self.root.after(0, lambda: self.label_status_login.config(
                        text="Usuário ou senha inválidos. Tente novamente.", fg='#e74c3c'))
                    self.root.after(0, lambda: self.btn_entrar.config(state=tk.NORMAL, text="Entrar"))
                    
            except Exception as e:
                import traceback
                erro_completo = traceback.format_exc()
                print(f"ERRO NO LOGIN: {erro_completo}")  # Debug
                if self.automacao.driver:
                    self.automacao.fechar()
                self.root.after(0, lambda: self.label_status_login.config(
                    text=f"Erro: {str(e)}", fg='#e74c3c'))
                self.root.after(0, lambda: self.btn_entrar.config(state=tk.NORMAL, text="Entrar"))
        
        # Executar login em thread separada
        import threading
        thread_login = threading.Thread(target=testar_login_thread, daemon=True)
        thread_login.start()
    
    def sair(self):
        """Sai do sistema e volta para tela de login"""
        if self.automacao.driver:
            self.automacao.fechar()
        
        self.logado = False
        self.usuario_logado = None
        self.senha_logada = None
        self.planilha_path = None
        self.autos = []
        self.label_planilha.config(text="Nenhuma planilha selecionada", fg="gray")
        
        # Ajustar tamanho da janela para tela de login
        self.root.geometry("450x350")
        self.root.resizable(False, False)
        
        self.mostrar_tela_login()
        self.entry_usuario.delete(0, tk.END)
        self.entry_senha.delete(0, tk.END)
        self.label_status_login.config(text="")
    
    def selecionar_planilha(self):
        """Abre diálogo para selecionar planilha"""
        arquivo = filedialog.askopenfilename(
            title="Selecionar Planilha",
            filetypes=[("Planilhas", "*.xlsx *.xls *.csv"), ("Excel", "*.xlsx *.xls"), ("CSV", "*.csv"), ("Todos", "*.*")]
        )
        
        if arquivo:
            self.planilha_path = arquivo
            self.label_planilha.config(text=os.path.basename(arquivo), fg="black")
            
            # Ler planilha para validar
            autos, _ = self.automacao.ler_planilha(arquivo)
            if autos:
                self.autos = autos
                messagebox.showinfo("Sucesso", f"Planilha carregada com sucesso!\n{len(autos)} autos encontrados.")
            else:
                messagebox.showerror("Erro", "Não foi possível ler a planilha ou nenhum auto foi encontrado.")
    
    def atualizar_progresso(self, mensagem):
        """Atualiza a mensagem de progresso"""
        self.progress_var.set(mensagem)
        self.root.update()
    
    def iniciar_automacao(self):
        """Inicia o processo de automação"""
        # Validar
        if not self.logado:
            messagebox.showerror("Erro", "Você precisa fazer login primeiro.")
            return
        
        if not self.planilha_path or not self.autos:
            messagebox.showerror("Erro", "Por favor, selecione uma planilha válida.")
            return
        
        # Desabilitar botão
        self.btn_iniciar.config(state="disabled")
        self.progress_bar.start()
        
        # Executar em thread separada
        try:
            # Criar navegador visível (não headless) para a automação
            self.atualizar_progresso("Iniciando navegador...")
            if not self.automacao.criar_driver(headless=False):
                self.btn_iniciar.config(state="normal")
                self.progress_bar.stop()
                return
            
            self.atualizar_progresso("Fazendo login...")
            if not self.automacao.fazer_login(self.usuario_logado, self.senha_logada):
                self.automacao.fechar()
                self.btn_iniciar.config(state="normal")
                self.progress_bar.stop()
                return
            
            self.atualizar_progresso("Navegando até o formulário...")
            if not self.automacao.navegar_para_formulario():
                messagebox.showerror("Erro", "Não foi possível navegar até o formulário.")
                self.automacao.fechar()
                self.btn_iniciar.config(state="normal")
                self.progress_bar.stop()
                return
            
            self.atualizar_progresso("Processando autos...")
            sucessos, erros = self.automacao.processar_autos(self.autos, self.atualizar_progresso)
            
            self.atualizar_progresso("Salvando resultados...")
            caminho_resultado = self.automacao.salvar_planilha_resultado()
            
            self.progress_bar.stop()
            self.automacao.fechar()
            
            mensagem = f"Processamento concluído!\n\nSucessos: {sucessos}\nErros: {erros}\n\n"
            if caminho_resultado:
                mensagem += f"Resultados salvos em:\n{caminho_resultado}"
            
            messagebox.showinfo("Concluído", mensagem)
            self.progress_var.set("Concluído!")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro durante a automação:\n{str(e)}")
            if self.automacao.driver:
                self.automacao.fechar()
        
        finally:
            self.btn_iniciar.config(state="normal")
    
    def fechar(self):
        """Fecha a aplicação"""
        if self.automacao.driver:
            self.automacao.fechar()
        self.root.quit()
        self.root.destroy()
    
    def executar(self):
        """Executa a interface"""
        self.root.mainloop()


if __name__ == "__main__":
    app = InterfaceGrafica()
    app.executar()

