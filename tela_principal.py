import os
import sys
import zipfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
from threading import Thread
from processamento import processar_mbox_html

class Recuperacao_email:
    def __init__(self, root):
        # CORES E ATRIBUTOS
        self.azul_escuro = "#203464"
        self.azul_claro = "#04acec"
        branco = "white"
        self.fonte = "Arial"
        self.tamanho_fonte_cabecalho = 16
        self.tamanho_fonte_corpo = 13
        self.cursor = "hand2"
        base = os.path.abspath(sys.argv[0])
        self.caminhoExe = os.path.dirname(base)
        self.caminho_imagens = os.path.join(self.caminhoExe, "imagens")
        self.pasta_inicial = "G:\\Drives compartilhados\\SUPER. EXEC - UTI\\Google Workspace\\Backup emails desativados"
        self.pasta_inicial_homo = "C:\\Automações Fapec\\Relatórios\\Restaurador de backup email\\backup email"
        self.pasta_destino = "C:\\Downloads"

        # CONFIGURAÇÃO DA TELA
        self.root = root
        self.root.title("Restauração de backup de e-mails")
        self.root.geometry("800x600")
        self.root.configure(bg=self.azul_escuro)

        # LOGO
        imagemDeFundo = Image.open(f'{self.caminho_imagens}/Fapec-logo.png').resize((162,145))
        imagemDeFundo = ImageTk.PhotoImage(imagemDeFundo)
        self.root.iconbitmap(f'{self.caminho_imagens}/Fapec-logo.ico')
        self.label_logo = tk.Label(root, image=imagemDeFundo, bg=self.azul_escuro)
        self.label_logo.image = imagemDeFundo
        self.label_logo.pack(pady=10)

        # BOTÃO RECUPERAR E-MAIL
        self.titulo = tk.Label(root, text="Selecione um arquivo para recuperar os e-mails", font=(self.fonte, self.tamanho_fonte_cabecalho, "bold"), bg=self.azul_escuro, fg=branco)
        self.titulo.pack(pady=10)

        # FRAME PRINCIPAL
        frame_principal = tk.Frame(root, bg=self.azul_escuro)
        frame_principal.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.texto_saida = tk.Text(frame_principal, wrap=tk.WORD, state=tk.DISABLED, height=15)
        self.texto_saida.grid(row=0, column=0, sticky="nsew")

        # SCROLLBAR FRAME PRINCIPAL
        self.scroll = ttk.Scrollbar(frame_principal, command=self.texto_saida.yview)
        self.texto_saida.configure(yscrollcommand=self.scroll.set)
        self.scroll.grid(row=0, column=1, sticky="ns")

        # FRAME BOTÕES
        frame_botoes = tk.Frame(frame_principal, bg=self.azul_escuro)
        frame_botoes.grid(row=0, column=2, sticky="ns", padx=(10, 0))

        # BARRA DE PESQUISA
        # Rótulo do campo de filtro
        self.label_filtro = tk.Label(frame_botoes, text="Filtrar e-mails por:", font=(self.fonte, self.tamanho_fonte_corpo, "bold"), bg=self.azul_escuro, fg=branco)
        self.label_filtro.pack(pady=(5, 2))

        barra_pesquisa_img = Image.open(f'{self.caminho_imagens}/barra_pesquisa.png').resize((262,62))
        barra_pesquisa_img = ImageTk.PhotoImage(barra_pesquisa_img)

        self.barra_pesquisa = tk.Label(
            frame_botoes,
            image=barra_pesquisa_img,
            bg=self.azul_escuro,
            activebackground=self.azul_escuro,
            bd=0
        )
        self.barra_pesquisa.image = barra_pesquisa_img
        self.barra_pesquisa.pack(pady=5)

        # Campo de entrada
        self.entrada_filtro = tk.Entry(frame_botoes, font=(self.fonte, self.tamanho_fonte_corpo), width=30, bd=0, fg="gray")

        def on_focus_in(event):
            if self.entrada_filtro.get() == "Separe tópicos por ;":
                self.entrada_filtro.delete(0, tk.END)
                self.entrada_filtro.config(fg="black")

        def on_focus_out(event):
            if self.entrada_filtro.get() == "":
                self.entrada_filtro.insert(0, "Separe tópicos por ;")
                self.entrada_filtro.config(fg="gray")
        
        self.entrada_filtro.insert(0, "Separe tópicos por ;")
        self.entrada_filtro.bind("<FocusIn>", on_focus_in)
        self.entrada_filtro.bind("<FocusOut>", on_focus_out)
        self.entrada_filtro.place(x=25, y=50, width=180, height=30)

        # BOTÃO SELECIONAR
        icone_selecionar_normal = Image.open(f'{self.caminho_imagens}/selecionar_arquivo.png').resize((262,62))
        icone_selecionar_normal = ImageTk.PhotoImage(icone_selecionar_normal)
        icone_selecionar_active = Image.open(f'{self.caminho_imagens}/selecionar_arquivo_active.png').resize((262,62))
        icone_selecionar_active = ImageTk.PhotoImage(icone_selecionar_active)

        self.botao_selecionar = tk.Button(
            frame_botoes,
            command=self.selecionar_arquivo,
            image=icone_selecionar_normal,
            bg=self.azul_escuro,
            activebackground=self.azul_escuro,
            relief="flat",
            bd=0
        )
        self.botao_selecionar.image = icone_selecionar_normal
        self.botao_selecionar.bind("<Enter>", lambda e: self.botao_selecionar.config(image=icone_selecionar_active))
        self.botao_selecionar.bind("<Leave>", lambda e: self.botao_selecionar.config(image=icone_selecionar_normal))
        self.botao_selecionar.pack(pady=5)

        # BOTÃO INICIAR
        icone_iniciar_normal = Image.open(f'{self.caminho_imagens}/iniciar_processo.png').resize((262,62))
        icone_iniciar_normal = ImageTk.PhotoImage(icone_iniciar_normal)
        icone_iniciar_active = Image.open(f'{self.caminho_imagens}/iniciar_processo_active.png').resize((262,62))
        icone_iniciar_active = ImageTk.PhotoImage(icone_iniciar_active)
        self.botao_iniciar = tk.Button(
            frame_botoes,
            command=self.iniciar_processamento,
            image=icone_iniciar_normal,
            bg=self.azul_escuro,
            activebackground=self.azul_escuro,
            relief="flat",
            bd=0,
            state="normal"
        )
        self.botao_iniciar.image = icone_iniciar_normal
        self.botao_iniciar.bind("<Enter>", lambda e: self.botao_iniciar.config(image=icone_iniciar_active))
        self.botao_iniciar.bind("<Leave>", lambda e: self.botao_iniciar.config(image=icone_iniciar_normal))
        self.botao_iniciar.pack(pady=5)

        frame_principal.rowconfigure(0, weight=1)
        frame_principal.columnconfigure(0, weight=1)

    def selecionar_arquivo(self):
        arquivo = filedialog.askopenfilename(
            title="Selecione um arquivo",
            filetypes=[("Arquivos ZIP", "*.zip"), ("Todos os arquivos", "*.*")],
            initialdir= self.pasta_inicial
        )
        if arquivo:
            thread = Thread(target=self.processar_arquivo_zip, args=(arquivo,), daemon=True)
            thread.start()

    def processar_arquivo_zip(self, arquivo):
        try:
            self.adicionar_mensagem(f"Extraindo arquivos, por favor aguarde.")
            pasta_temp = os.path.join(os.getcwd(), "temp_extracao")
            if not os.path.exists(pasta_temp):
                os.makedirs(pasta_temp)
            
            if zipfile.is_zipfile(arquivo):
                with zipfile.ZipFile(arquivo, 'r') as zip_ref:
                    zip_ref.extractall(pasta_temp)
                    self.adicionar_mensagem(f"Arquivos extraidos com sucesso.")

                arquivo_mbox = None
                for root, _, files in os.walk(pasta_temp):
                    for file in files:
                        if file.endswith(".mbox"):
                            arquivo_mbox = os.path.join(root, file)
                            break
                    if arquivo_mbox:
                        break

                if arquivo_mbox:
                    self.arquivo_mbox = arquivo_mbox
                    self.pasta_temp = pasta_temp
                    self.botao_iniciar.config(state="normal")
                else:
                    messagebox.showerror("Erro", "Nenhum arquivo .mbox encontrado no ZIP.")
            else:
                messagebox.showerror("Erro", "O arquivo selecionado não é um ZIP válido.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar o arquivo: {e}")

    def iniciar_processamento(self, ):
        self.botao_iniciar.config(state="disabled")
        self.botao_selecionar.config(state="disabled")
        self.barra_pesquisa.config(state="disabled")
        self.entrada_filtro.config(state="disabled")
        filtro = self.entrada_filtro.get().strip()
        if filtro == "Separe por ; os tópicos":
            filtro = None
        thread = Thread(target=processar_mbox_html, args=(self, filtro), daemon=True)
        thread.start()

    def atualizar_titulo(self, texto):
        self.titulo.config(text=texto)

    def adicionar_mensagem(self, mensagem):
        self.texto_saida.tag_configure("fonte_grande", font=(self.fonte, self.tamanho_fonte_corpo))
        self.texto_saida.config(state=tk.NORMAL)
        self.texto_saida.insert(tk.END, mensagem + "\n", "fonte_grande")
        self.texto_saida.config(state=tk.DISABLED)
        self.texto_saida.see(tk.END)

if __name__ == "__main__":
    root = tk.Tk()
    tela = Recuperacao_email(root)
    root.mainloop()