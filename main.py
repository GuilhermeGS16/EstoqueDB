import customtkinter as ctk
from tkinter import messagebox, simpledialog
import json
import tkinter as tk
import os
from datetime import datetime
from PIL import Image, ImageTk
import webbrowser
import urllib.parse
import unicodedata
import time
import getpass
import random

ARQUIVO_ESTOQUE = "estoque.json"
ARQUIVO_LOG = "log.txt"
SENHA_ADMIN = "ad"
ARQUIVO_REQUISICOES = "requisicoes.json"

def normalizar_nome(nome):
    return ''.join(c for c in unicodedata.normalize('NFD', nome.lower()) if unicodedata.category(c) != 'Mn')
    
def carregar_estoque():
    if not os.path.exists(ARQUIVO_ESTOQUE):
        return {}
    with open(ARQUIVO_ESTOQUE, "r") as f:
        return json.load(f)
def salvar_estoque(estoque):
    with open(ARQUIVO_ESTOQUE, "w") as f:
        json.dump(estoque, f, indent=4)

def registrar_log(produto, quantidade, modo):
    agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    usuario = getpass.getuser()
    linha = f"[{agora}] ({modo}) {produto}: {'+' if quantidade >= 0 else ''}{quantidade}  | usuário: {usuario}\n"
    
    with open(ARQUIVO_LOG, "a", encoding="utf-8") as f:
        f.write(linha)
        
    
def registrar_log_solicitacao(nome_usuario, itens):
    agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    lista = " | ".join(f"{qtd}/{produto}" for produto, qtd in itens.items())
    linha = f"[{agora}] [SOLICITAÇÃO] {nome_usuario} solicitou: {lista}\n"

    with open(ARQUIVO_LOG, "a", encoding="utf-8") as f:
        f.write(linha)
        
def carregar_requisicoes():
        if os.path.exists("requisicoes.json"):
            with open("requisicoes.json", "r", encoding="utf-8") as f:
                return json.load(f)
        return {}
    
def registrar_requisicao(produto, quantidade):
    # Lê o arquivo ou cria vazio
    if os.path.exists(ARQUIVO_REQUISICOES):
        with open(ARQUIVO_REQUISICOES, "r", encoding="utf-8") as f:
            requisicoes = json.load(f)
    else:
        requisicoes = {}

    # Incrementa a quantidade
    requisicoes[produto] = requisicoes.get(produto, 0) + quantidade

    # Salva o resultado
    with open(ARQUIVO_REQUISICOES, "w", encoding="utf-8") as f:
        json.dump(requisicoes, f, indent=4, ensure_ascii=False)

        
class AppEstoque(ctk.CTk):
    def __init__(self):
        super().__init__()
        # Carregar dados das impressoras
        with open("impressoras.json", "r", encoding="utf-8") as f:
            self.dados_impressoras = json.load(f)

        
        self.icone_ok = ctk.CTkImage(Image.open("assets/check.png").resize((16, 16)))
        self.img_moon = ctk.CTkImage(Image.open("assets/moon.png"), size=(24, 24))
        self.img_sun = ctk.CTkImage(Image.open("assets/sun.png"), size=(24, 24))
        self.icon_filtro_black = ctk.CTkImage(Image.open("assets/filterblack.png"), size=(20, 20))
        self.icon_filtro_white = ctk.CTkImage(Image.open("assets/filterwhite.png"), size=(20, 20))
        self.icone_alerta_amarelo = ctk.CTkImage(Image.open("assets/alert_yellow.png").resize((15,15)))
        self.icone_alerta_vermelho = ctk.CTkImage(Image.open("assets/alert_red.png").resize((15,15)))
        self.icones_status = {}

        self.title("Controle de Estoque TI")
        self.attributes("-fullscreen", True)
        self.bind("<Escape>", lambda e: self.attributes("-fullscreen", False))

        ctk.set_appearance_mode("dark")
        self.current_theme = "dark"
        self.toggle_size = 36
        self.toggle_margin = 2
        self.fonte_negrito = ctk.CTkFont(weight="bold")

        self.icone_lixeira = ctk.CTkImage(
            light_image=Image.open("assets/trash.png").resize((15,15)), size=(15,15)
        )

        self.estoque = carregar_estoque()
        self.labels = {}
        self.linhas_produtos = []
        self.labels_requisicao = {}
        self.criar_interface()
        self.lift()
        self.attributes('-topmost', True)
        self.after(100, lambda: self.attributes('-topmost', False))
        
    def abrir_interface_impressoras(self):
        janela_impressoras = ctk.CTkToplevel(self)
        janela_impressoras.title("Painel de Impressoras")
        janela_impressoras.geometry("600x500")
        janela_impressoras.lift()
        janela_impressoras.focus_force()
        janela_impressoras.grab_set()

        for impressora in self.dados_impressoras["impressoras"]:
            frame = ctk.CTkFrame(janela_impressoras)
            frame.pack(pady=10, padx=10, fill="x")

            # Cabeçalho com dados da impressora
            info = (
                f"{impressora['modelo']} (Andar {impressora['andar']})\n"
                f"N/S: {impressora['numero_serie']}  |  IP: {impressora['ip']}"
            )
            ctk.CTkLabel(frame, text=info, justify="left", anchor="w").pack(anchor="w", padx=10, pady=10)
        janela_impressoras.lift()
        janela_impressoras.attributes('-topmost', True)
        janela_impressoras.after(100, lambda: janela_impressoras.attributes('-topmost', False))
    
        
    def atualizar_status_linha(self, produto, index):
        dados = self.estoque[produto]
        qtd = dados["quantidade"]
        alerta = dados["alerta"]

        if qtd <= alerta // 2:
            novo_icone = self.icone_alerta_vermelho
        elif qtd <= alerta:
            novo_icone = self.icone_alerta_amarelo
        else:
            novo_icone = self.icone_ok

        # Atualiza o ícone apenas se já existir
        if produto in self.icones_status:
            self.icones_status[produto].configure(image=novo_icone)

        # Atualiza o fundo da linha
        cor_fundo = self.get_cores_fundo(index, qtd, alerta)
        self.linhas_produtos[index].configure(fg_color=cor_fundo)



    def distribuir_toners(self, itens_solicitados):
        import random

        distribuido = {}  # {numero_serie: {"modelo":..., "ip":..., "andar":..., "toners":[...]}}
        usados = set()

        for produto, qtd in itens_solicitados.items():
            # Encontra impressoras que usam esse toner
            candidatas = [
                imp for imp in self.dados_impressoras["impressoras"]
                if produto in imp["toners"]
            ]

            if not candidatas:
                raise Exception(f"Nenhuma impressora usa o toner '{produto}'.")

            random.shuffle(candidatas)

            for _ in range(qtd):
                for imp in candidatas:
                    chave = (imp["numero_serie"], produto)
                    if chave not in usados:
                        usados.add(chave)
                        ns = imp["numero_serie"]
                        if ns not in distribuido:
                            distribuido[ns] = {
                                "modelo": imp["modelo"],
                                "ip": imp["ip"],
                                "andar": imp["andar"],
                                "toners": []
                            }
                        distribuido[ns]["toners"].append(produto)
                        break
                else:
                    raise Exception(f"Sem impressora disponível para o toner '{produto}'.")

        return distribuido

    def carregar_estoque(self):
        ctk.CTkLabel(self.conteudo_principal, text="Tabela de Estoque", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=20)
        # Aqui você coloca a sua tabela de produtos com as cores e botões

    def acao_exemplo(self):
        print("Botão pressionado!")

        
    def aplicar_cor_por_tema(self, cor_dark, cor_light):
        return cor_dark if self.current_theme == "dark" else cor_light
    
    def filtrar_produtos(self, event=None):
        termo = self.entry_pesquisa.get().lower()
        for linha, produto in zip(self.linhas_produtos, self.estoque):
            if termo in produto.lower():
                linha.pack(pady=1, padx=10, fill="x")
            else:
                linha.pack_forget()


    def criar_interface(self):
        title = ctk.CTkLabel(
            self, text="Estoque TI", font=ctk.CTkFont(size=24, weight="bold")
        )
        title.pack(pady=10)

        botoes_frame = ctk.CTkFrame(self, fg_color="transparent")
        
        botoes_frame.pack(pady=10)

        # Botões principais
        for texto, comando, cor in [
            ("Solicitar Produtos", self.abrir_janela_solicitacao, "#70BAC5"),
            ("Adicionar Produto", self.adicionar_produto, "#90BE6D"),
            ("Modo Administrador", self.entrar_admin, "#F04C60"),
            ("Ver Relatório de Movimentações", self.ver_relatorio, "#7FBEAB"),
            ("Impressoras", self.abrir_interface_impressoras, "#9C7FBE"),
        ]:
            ctk.CTkButton(
                botoes_frame, text=texto, command=comando,
                fg_color=cor, hover_color=None,
                font=self.fonte_negrito, corner_radius=8, text_color="white"
            ).pack(side="left", padx=10, pady=10)

        # Botão toggle estilizado com ícone (sol/lua)
        self.btn_toggle_theme = ctk.CTkButton(
            botoes_frame,
            text="",
            width=36,
            height=36,
            corner_radius=18,
            fg_color="#383A55",  # cor inicial para modo dark
            hover_color="#2E3049",
            image=self.img_moon,
            command=self.toggle_theme
        )  
        self.btn_toggle_theme.pack(side="left", padx=8, pady=10)
        

        search_frame = ctk.CTkFrame(self)
        search_frame.pack(pady=(0,10))

        # Campo de busca
        self.entry_pesquisa = ctk.CTkEntry(
            search_frame, placeholder_text="Buscar material...", width=300
        )
        self.entry_pesquisa.pack(side="left", padx=(0,5))
        self.entry_pesquisa.bind("<KeyRelease>", lambda e: self.filtrar_produtos())

        # Menu de ordenação
        self.filtro_ativo = "A-Z"  # valor inicial

        self.btn_filtro = ctk.CTkButton(
        search_frame,
        image=self.icon_filtro_black if self.current_theme == "light" else self.icon_filtro_white,
        text="",
        width=40,
        height=36,
        fg_color=self.aplicar_cor_por_tema("#2c2c2c", "#e0e0e0"),  # invertido aqui
        hover_color=self.aplicar_cor_por_tema("#1e1e1e", "#d0d0d0"),  # invertido aqui
        command=self.abrir_menu_filtro
    )

        self.btn_filtro.pack(side="left", padx=5)
        # Área de produtos

        # Cabeçalho (fora do scrollable frame)
        self.cabecalho = ctk.CTkFrame(self, fg_color="transparent")
        self.cabecalho.pack(pady=(5, 0), padx=10, fill="x")

        ctk.CTkLabel(self.cabecalho, text="Materiais", font=ctk.CTkFont(size=20, weight="bold"), width=200, anchor="w").pack(side="left", padx=10)
        ctk.CTkLabel(self.cabecalho, text="Quantidade", font=ctk.CTkFont(size=20, weight="bold"), width=40, anchor="w").pack(side="left", padx=5)

        # Agora o scroll começa logo abaixo
        self.produtos_frame = ctk.CTkScrollableFrame(self, width=800, height=600)
        self.produtos_frame.pack(pady=(0, 10), fill="both", expand=True)

        # Linhas de produtos
        self.linhas_produtos.clear()
        produtos_visiveis = [p for p in self.estoque if p != "impressoras"]
        for idx, produto in enumerate(produtos_visiveis):
            linha = self.criar_linha_produto(produto, idx)
            self.linhas_produtos.append(linha)
            
    


    
    def abrir_menu_filtro(self):
        menu = tk.Menu(self, tearoff=0, bg=self.aplicar_cor_por_tema("#FFF7E7", "#212235"), 
                    fg=self.aplicar_cor_por_tema("#000000", "#FFFFFF"),
                    activebackground=self.aplicar_cor_por_tema("#f0e7c0", "#2e2f4a"),
                    activeforeground=self.aplicar_cor_por_tema("#000000", "#ffffff"))
        
        opcoes = ["A-Z", "Z-A"]
        for opcao in opcoes:
            menu.add_command(label=opcao, command=lambda c=opcao: self.ordenar(c))
        
        x = self.btn_filtro.winfo_rootx()
        y = self.btn_filtro.winfo_rooty() + self.btn_filtro.winfo_height()
        menu.tk_popup(x, y)

    def ordenar(self, criterio):
        # Reordena stocks e widgets
        pares = list(zip(self.estoque.items(), self.linhas_produtos))
        if criterio == "A-Z":
            pares.sort(key=lambda x: x[0][0].lower())
        elif criterio == "Z-A":
            pares.sort(key=lambda x: x[0][0].lower(), reverse=True)
        # Reexibe em ordem
        for (_, _), linha in pares:
            linha.pack_forget()
            linha.pack(pady=1, padx=10, fill="x")
        # Atualiza listas internas
        self.linhas_produtos = [linha for (_, linha) in pares]
        self.estoque = {nome:qtd for ((nome,qtd), _) in pares}
   
    def get_cores_fundo_dark(self, index):
        # Cores para modo dark
        return "#2c2c2c" if index % 2 == 0 else "#242424"    

    def get_cores_fundo_light(self, index):
        # Cores para modo light
        return "#f9f9f9" if index % 2 == 0 else "#eaeaea"

    def get_cores_fundo(self, index, qtd=None, alerta=None):
        if self.current_theme == "dark":
            return "#2c2c2c" if index % 2 == 0 else "#242424"
        else:
            return "#e0e0e0" if index % 2 == 0 else "#f5f5f5"


    def criar_linha_produto(self, produto, index):
        dados = self.estoque[produto]
        qtd = dados["quantidade"]
        alerta = dados["alerta"]
        cor_fundo = self.get_cores_fundo(index)

        linha = ctk.CTkFrame(self.produtos_frame, fg_color=cor_fundo)
        linha.pack(pady=1, padx=10, fill="x")

        # Define o ícone de status
        if qtd <= alerta // 2:
            icone = self.icone_alerta_vermelho
        elif qtd <= alerta:
            icone = self.icone_alerta_amarelo
        else:
            icone = self.icone_ok

        # Ícone direto no frame
        icone_label = ctk.CTkLabel(linha, image=icone, text="")
        icone_label.pack(side="left", padx=(10, 10), pady=5)
        self.icones_status[produto] = icone_label  # Salva para atualizar depois

        # Nome do produto
        ctk.CTkLabel(
            linha, text=produto, font=ctk.CTkFont(size=14, weight="bold"), anchor="w", width=200
        ).pack(side="left", pady=5)

        # Botão de remover
        linha.btn_menos = ctk.CTkButton(
            linha, text="–", width=36, height=36, corner_radius=8,
            command=lambda p=produto: self.alterar(p, -1),
            font=ctk.CTkFont(family="Arial", size=18, weight="bold"),
            fg_color="#F96F5D", hover_color="#F67280", text_color="white"
        )
        linha.btn_menos.pack(side="left", padx=5, pady=5)

        # Quantidade
        self.labels[produto] = ctk.CTkLabel(linha, text=str(qtd), width=40, font=self.fonte_negrito)
        self.labels[produto].pack(side="left")

        # Botão de adicionar
        linha.btn_mais = ctk.CTkButton(
            linha, text="+", width=36, height=36, corner_radius=8,
            command=lambda p=produto: self.alterar(p, 1),
            font=ctk.CTkFont(family="Arial", size=18, weight="bold"),
            fg_color="#90BE6D", hover_color="#3DDC97", text_color="white"
        )
        linha.btn_mais.pack(side="left", padx=5, pady=5)

        # Botão lixeira
        ctk.CTkButton(
            linha, image=self.icone_lixeira, text="", width=36, height=36, corner_radius=8,
            command=lambda p=produto, l=linha: self.remover_produto(p, l),
            fg_color="#7E998A", hover_color="#6F867A"
        ).pack(side="left", padx=5, pady=5)
        
        self.requisicoes = carregar_requisicoes()
        
        if produto in self.requisicoes:
            label = ctk.CTkLabel(
                linha, text=f"Solicitado {self.requisicoes[produto]}x",
                text_color="#EAC449", font=ctk.CTkFont(size=12)
            )
            label.pack(side="left", padx=(10, 5))
            self.labels_requisicao[produto] = label



        return linha

    def remover_produto(self, produto, linha_widget):
        resposta = messagebox.askyesno("Confirmar Remoção", f"Tem certeza que deseja remover o produto '{produto}'?")
        if resposta:
            self.estoque.pop(produto, None)
            salvar_estoque(self.estoque)
            linha_widget.destroy()
            del self.labels[produto]
            registrar_log(produto, 0, "REMOVIDO")

    def alterar(self, produto, delta):
        dados = self.estoque[produto]
        nova_qtd = max(0, dados["quantidade"] + delta)

        if nova_qtd == dados["quantidade"]:
            return

        dados["quantidade"] = nova_qtd
        self.labels[produto].configure(text=str(nova_qtd))
        salvar_estoque(self.estoque)
        registrar_log(produto, delta, "NORMAL")

        # Atualiza cor da linha
        index = list(self.estoque).index(produto)
        alerta = dados["alerta"]
        linha = self.labels[produto].master  # pega o frame pai do labe
        self.atualizar_status_linha(produto, index)



    def adicionar_produto(self):
        def confirmar():
            nome = entry_nome.get().strip()
            quantidade_str = entry_quantidade.get().strip()
            alerta_str = entry_alerta.get().strip()

            if not nome:
                messagebox.showerror("Erro", "Nome do produto não pode ser vazio!")
                return
            if nome in self.estoque:
                messagebox.showerror("Erro", "Este produto já existe!")
                return

            try:
                quantidade = int(quantidade_str)
                alerta = int(alerta_str)
                if quantidade < 0 or alerta < 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Erro", "Quantidade e Alerta devem ser números inteiros não negativos!")
                return

            # Salvar estrutura como dicionário com quantidade + alerta
            self.estoque[nome] = {
                "quantidade": quantidade,
                "alerta": alerta
            }

            salvar_estoque(self.estoque)
            registrar_log(nome, quantidade, "NOVO")
            linha = self.criar_linha_produto(nome, len(self.estoque) - 1)
            self.linhas_produtos.append(linha)
            janela.destroy()

        def cancelar():
            janela.destroy()

        janela = ctk.CTkToplevel(self)
        janela.title("Adicionar Novo Produto")
        largura = 300
        altura = 300

        janela.geometry(f"{largura}x{altura}")

        largura_tela = janela.winfo_screenwidth()
        altura_tela = janela.winfo_screenheight()

        pos_x = int((largura_tela / 2) - (largura / 2))
        pos_y = int((altura_tela / 2) - (altura / 2))

        janela.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")
        janela.grab_set()

        ctk.CTkLabel(janela, text="Nome do Produto:", font=self.fonte_negrito).pack(pady=(15, 5), anchor="w", padx=20)
        entry_nome = ctk.CTkEntry(janela, width=280)
        entry_nome.pack(padx=20)

        ctk.CTkLabel(janela, text="Quantidade Inicial:", font=self.fonte_negrito).pack(pady=(10, 5), anchor="w", padx=20)
        entry_quantidade = ctk.CTkEntry(janela, width=280)
        entry_quantidade.pack(padx=20)
        
        ctk.CTkLabel(janela, text="Quantidade de Alerta:", font=self.fonte_negrito).pack(pady=(10, 5), anchor="w", padx=20)
        entry_alerta = ctk.CTkEntry(janela, width=280)
        entry_alerta.pack(padx=20)


        botoes_frame = ctk.CTkFrame(janela, fg_color="transparent")
        botoes_frame.pack(pady=15)

        ctk.CTkButton(botoes_frame, text="Confirmar", command=confirmar, fg_color="#90BE6D", width=120).pack(side="left", padx=10)
        ctk.CTkButton(botoes_frame, text="Cancelar", command=cancelar, fg_color="#F04C60", width=120).pack(side="left", padx=10)

    def entrar_admin(self):
        senha = simpledialog.askstring("Admin", "Digite a senha:", show="*")
        if senha == SENHA_ADMIN:
            self.abrir_modo_admin()
        else:
            messagebox.showerror("Erro", "Senha incorreta!")

    def abrir_modo_admin(self):
        janela_admin = ctk.CTkToplevel(self)
        janela_admin.title("Modo Administrador")
        largura = 500
        altura = 900

        janela_admin.geometry(f"{largura}x{altura}")

        largura_tela = janela_admin.winfo_screenwidth()
        altura_tela = janela_admin.winfo_screenheight()

        pos_x = int((largura_tela / 2) - (largura / 2))
        pos_y = int((altura_tela / 2) - (altura / 2))

        janela_admin.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")
        janela_admin.grab_set()

        entradas_qtd = {}
        entradas_nome = {}
        entradas_alerta = {}

        # Área com rolagem
        scroll_area = ctk.CTkScrollableFrame(janela_admin, width=580, height=750)
        scroll_area.pack(padx=10, pady=(10, 0), fill="both", expand=True)

        for produto in self.estoque:
            if produto == "impressoras":
                continue

            dados = self.estoque[produto]

            frame = ctk.CTkFrame(scroll_area)
            frame.pack(pady=5, padx=10, fill="x")

            # Campo nome
            entry_nome = ctk.CTkEntry(frame, width=180)
            entry_nome.insert(0, produto)
            entry_nome.pack(side="left", padx=5)
            entradas_nome[produto] = entry_nome

            # Campo quantidade
            entry_qtd = ctk.CTkEntry(frame, width=60)
            entry_qtd.insert(0, str(dados["quantidade"]))
            entry_qtd.pack(side="left", padx=5)
            entradas_qtd[produto] = entry_qtd

            # Campo alerta
            entry_alerta = ctk.CTkEntry(frame, width=60)
            entry_alerta.insert(0, str(dados["alerta"]))
            entry_alerta.pack(side="left", padx=5)
            entradas_alerta[produto] = entry_alerta

        # Função salvar
        def salvar_todos():
            novo_estoque = {}
            nomes_novos = set()

            for produto_antigo in entradas_nome:
                nome_novo = entradas_nome[produto_antigo].get().strip()
                qtd_str = entradas_qtd[produto_antigo].get().strip()
                alerta_str = entradas_alerta[produto_antigo].get().strip()

                if not nome_novo:
                    messagebox.showerror("Erro", f"O nome do produto não pode estar vazio.")
                    return

                if nome_novo in nomes_novos:
                    messagebox.showerror("Erro", f"O nome '{nome_novo}' está duplicado.")
                    return
                nomes_novos.add(nome_novo)

                try:
                    nova_qtd = int(qtd_str)
                    novo_alerta = int(alerta_str)
                    if nova_qtd < 0 or novo_alerta < 0:
                        raise ValueError
                except ValueError:
                    messagebox.showerror("Erro", f"Quantidade ou alerta inválido para '{nome_novo}'")
                    return

                novo_estoque[nome_novo] = {
                    "quantidade": nova_qtd,
                    "alerta": novo_alerta
                }

                if nome_novo != produto_antigo:
                    registrar_log(produto_antigo, 0, f"RENOMEADO PARA {nome_novo}")
                else:
                    delta = nova_qtd - self.estoque[produto_antigo]["quantidade"]
                    if delta != 0:
                        registrar_log(nome_novo, delta, "ADMIN")

            # Mantém dados das impressoras, se houver
            if "impressoras" in self.estoque:
                novo_estoque["impressoras"] = self.estoque["impressoras"]

            self.estoque = novo_estoque
            salvar_estoque(self.estoque)
            self.atualizar_interface_total()
            janela_admin.destroy()


        # Botão de salvar (fora do scroll)
        btn_salvar = ctk.CTkButton(
            janela_admin,
            text="Salvar alterações",
            command=salvar_todos,
            fg_color="#90BE6D",
            hover_color="#7AA65A",
            font=self.fonte_negrito,
            corner_radius=20,
            width=200
        )
        btn_salvar.pack(pady=20)

    def atualizar_interface_total(self):
        for widget in self.produtos_frame.winfo_children():
            widget.destroy()

        self.labels.clear()
        self.linhas_produtos.clear()

        visiveis = [p for p in self.estoque if p != "impressoras"]
        for idx, produto in enumerate(visiveis):
            linha = self.criar_linha_produto(produto, idx)
            self.linhas_produtos.append(linha)


    def ver_relatorio(self):
        if not os.path.exists(ARQUIVO_LOG):
            messagebox.showinfo("Relatório", "Nenhuma movimentação registrada ainda.")
            return

        janela_log = ctk.CTkToplevel(self)
        janela_log.title("Relatório de Movimentações")
        largura = 800
        altura = 600

        janela_log.geometry(f"{largura}x{altura}")

        largura_tela = janela_log.winfo_screenwidth()
        altura_tela = janela_log.winfo_screenheight()

        pos_x = int((largura_tela / 2) - (largura / 2))
        pos_y = int((altura_tela / 2) - (altura / 2))

        janela_log.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")

        texto = ctk.CTkTextbox(janela_log, width=780, height=550, font=self.fonte_negrito)
        texto.pack(pady=10, padx=10)
        texto.configure(state="normal")

        # Ler e destacar conteúdo do log
        with open(ARQUIVO_LOG, "r", encoding="utf-8") as f:
            linhas = f.readlines()
            for i, linha in enumerate(linhas):
                linha_index = f"{i+1}.0"
                texto.insert(linha_index, linha)
                if "[SOLICITAÇÃO]" in linha:
                    texto.tag_add("solicitacao", linha_index, f"{i+1}.end")

        # Estilo da tag de solicitação
        texto.tag_config("solicitacao", foreground="#FCEC52")  # Amarelo

        texto.configure(state="disabled")
        janela_log.lift()
        janela_log.focus_force()
        janela_log.attributes("-topmost", True)
        janela_log.after(100, lambda: janela_log.attributes("-topmost", False))


    def toggle_theme(self):
        if self.current_theme == "dark":
            self.current_theme = "light"
            ctk.set_appearance_mode("light")
            self.btn_toggle_theme.configure(
                image=self.img_sun,
                fg_color="#FFF7E7",
                hover_color="#f0e7c0"
            )
        else:
            self.current_theme = "dark"
            ctk.set_appearance_mode("dark")
            self.btn_toggle_theme.configure(
                image=self.img_moon,
                fg_color="#383A55",
                hover_color="#2E3049"
            )

        # Atualizar botão de filtro
        self.btn_filtro.configure(
            fg_color=self.aplicar_cor_por_tema("#2c2c2c", "#e0e0e0"),
            hover_color=self.aplicar_cor_por_tema("#1e1e1e", "#d0d0d0"),
            image=self.icon_filtro_black if self.current_theme == "light" else self.icon_filtro_white
        )

        # Atualizar linhas da lista
        visiveis = [p for p in self.estoque if p != "impressoras"]
        for idx, produto in enumerate(visiveis):
            linha = self.linhas_produtos[idx]
            dados = self.estoque[produto]
            qtd = dados["quantidade"]
            alerta = dados["alerta"]
            cor = self.get_cores_fundo(idx, qtd, alerta)
            linha.configure(fg_color=cor)

            # Atualizar botões de + e –
            if hasattr(linha, "btn_mais") and hasattr(linha, "btn_menos"):
                if self.current_theme == "dark":
                    linha.btn_mais.configure(fg_color="#90BE6D", hover_color="#3DDC97")
                    linha.btn_menos.configure(fg_color="#F96F5D", hover_color="#F67280")
                else:
                    linha.btn_mais.configure(fg_color="#B9D8C2", hover_color="#A7C2AF")
                    linha.btn_menos.configure(fg_color="#FB9F89", hover_color="#E2907C")

        self.cabecalho.configure(fg_color=self.get_cores_fundo(0))

    def abrir_janela_solicitacao(self):
        janela = ctk.CTkToplevel(self)
        janela.title("Solicitar Produtos")
        largura = 500
        altura = 900

        janela.geometry(f"{largura}x{altura}")

        largura_tela = janela.winfo_screenwidth()
        altura_tela = janela.winfo_screenheight()

        pos_x = int((largura_tela / 2) - (largura / 2))
        pos_y = int((altura_tela / 2) - (altura / 2))

        janela.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")
        janela.grab_set()

        entradas = {}

        ctk.CTkLabel(janela, text="Preencha as quantidades desejadas:", font=self.fonte_negrito).pack(pady=10)

        # Frame com scroll
        scroll_frame = ctk.CTkScrollableFrame(janela, width=480, height=700)
        scroll_frame.pack(padx=10, pady=5, fill="both", expand=True)

        for produto in self.estoque:
            frame = ctk.CTkFrame(scroll_frame)
            frame.pack(pady=4, padx=10, fill="x")

            ctk.CTkLabel(frame, text=produto, width=200, anchor="w").pack(side="left", padx=5)
            entry = ctk.CTkEntry(frame, width=60)
            entry.pack(side="left")
            entradas[produto] = entry

        # Botão Toners
        ctk.CTkButton(
            janela, text="Solicitar Toners (MR Copiadoras)", command=lambda: enviar_email_toner(entradas),
            fg_color="#F04C60", corner_radius=30
        ).pack(pady=10)
        
        def salvar_requisicao(produto, mensagem):
            if not os.path.exists(ARQUIVO_REQUISICOES):
                requisicoes = {}
            else:
                with open(ARQUIVO_REQUISICOES, "r", encoding="utf-8") as f:
                    requisicoes = json.load(f)

            if produto not in requisicoes:
                requisicoes[produto] = []
            requisicoes[produto].append(mensagem)

            with open(ARQUIVO_REQUISICOES, "w", encoding="utf-8") as f:
                json.dump(requisicoes, f, indent=4, ensure_ascii=False)

        def confirmar_envio():
            itens_solicitados = {}
            for produto, entrada in entradas.items():
                valor = entrada.get().strip()
                if valor:
                    try:
                        qtd = int(valor)
                        if qtd > 0:
                            itens_solicitados[produto] = qtd
                    except ValueError:
                        continue

            if not itens_solicitados:
                messagebox.showwarning("Aviso", "Nenhum item foi solicitado.")
                return

            # Envia e registra a requisição
            self.enviar_email_solicitacao(itens_solicitados)
            usuario = getpass.getuser()
            registrar_log_solicitacao(usuario, itens_solicitados)

            # Registra no JSON de requisições e atualiza interface
            for produto, qtd in itens_solicitados.items():
                registrar_requisicao(produto, qtd)

                # Atualiza texto amarelo na interface
                if produto in self.labels_requisicao:
                    self.labels_requisicao[produto].configure(text=f"Solicitado {self.requisicoes.get(produto, qtd)}x")
                else:
                    label = ctk.CTkLabel(
                        self.labels[produto].master,
                        text=f"Solicitado {self.requisicoes.get(produto, qtd)}x",
                        text_color="#EAC449",
                        font=ctk.CTkFont(size=12)
                    )
                    label.pack(side="left", padx=(10, 5))
                    self.labels_requisicao[produto] = label

            # Fecha a janela
            janela.destroy()



        # Botão de Envio Normal
        ctk.CTkButton(
            janela, text="Enviar Solicitação", command=confirmar_envio,
            fg_color="#3B8ED0", corner_radius=30
        ).pack(pady=10)
        
        def enviar_email_toner(self, distribuicao):
            import urllib.parse
            import webbrowser
            import getpass
            from datetime import datetime

            destinatario = "empresa.terceirizada@exemplo.com"
            assunto = "Solicitação de Toner de Impressora"
            corpo = "Olá,\n\nSolicito os seguintes toners:\n\n"

            for ns, toners in distribuicao.items():
                impressora = next((imp for imp in self.dados_impressoras["impressoras"] if imp["numero_serie"] == ns), None)
                if not impressora:
                    continue

                corpo += f"N/S {ns}:\n"
                corpo += f"  Modelo: {impressora['modelo']}\n"
                corpo += f"  IP: {impressora['ip']}\n"
                corpo += f"  Andar: {impressora['andar']}\n"
                corpo += f"  Toners:\n"
                for toner in toners:
                    corpo += f"   - {toner}\n"
                corpo += "\n"

            corpo += "Entrega na matriz (sede).\n"
            corpo += "Agradeço o atendimento.\n\n"

            usuario = getpass.getuser()
            corpo += f"Att,\n{usuario}"

            # Codifica o corpo e assunto para mailto
            corpo_url = urllib.parse.quote(corpo)
            assunto_url = urllib.parse.quote(assunto)

            link = f"mailto:{destinatario}?subject={assunto_url}&body={corpo_url}"
            webbrowser.open(link)

    def enviar_email_solicitacao(self, itens_solicitados):
        destinatario = "handrikson.petzold@patrimar.com.br"  
        assunto = "Solicitação de Requisição de Materiais"
        
        corpo = "Olá,tudo bem?\n\nSolicito os seguintes itens para a requisição de TI:\n\n"
        for produto, qtd in itens_solicitados.items():
            corpo += f"- {produto}: {qtd} unidade(s)\n"
        corpo += "\nAguardo confirmação.\n\nAtenciosamente,\nGuilherme Gonçalves Salomé"

        corpo_url = urllib.parse.quote(corpo)
        assunto_url = urllib.parse.quote(assunto)
        
        link = f"mailto:{destinatario}?subject={assunto_url}&body={corpo_url}"
        webbrowser.open(link)

if __name__ == "__main__":
    app = AppEstoque()
    app.mainloop()