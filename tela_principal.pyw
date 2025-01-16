import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog
from PIL import Image, ImageTk
from threading import Thread
from processamento import processar_mbox_html

class Recuperacao_email:
    def __init__(self, root):
        azul_escuro = "#203464"
        azul_claro = "#04acec"
        branco = "white"
        base = os.path.abspath(sys.argv[0])
        self.caminhoExe = os.path.dirname(base)
        imagemDeFundo = Image.open(f'{self.caminhoExe}\\imagens\\Fapec-logo.png').resize((162,145))
        imagemDeFundo = ImageTk.PhotoImage(imagemDeFundo)
        self.root = root
        self.root.title("Restauração de backup de e-mails")
        self.root.geometry("800x600")
        self.root.configure(bg=azul_escuro)

        self.root.iconbitmap(f'{self.caminhoExe}\\imagens\\Fapec-logo.ico')

        self.label_logo = tk.Label(root, image=imagemDeFundo, bg=azul_escuro)
        self.label_logo.image = imagemDeFundo
        self.label_logo.pack(pady=10)

        self.titulo = tk.Label(root, text="Selecione um arquivo MBOX para recuperar os e-mails", font=("Arial", 16), bg=azul_escuro, fg=branco)
        self.titulo.pack(pady=10)

        frame_principal = tk.Frame(root, bg=azul_escuro)
        frame_principal.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.texto_saida = tk.Text(frame_principal, wrap=tk.WORD, state=tk.DISABLED, height=15)
        self.texto_saida.grid(row=0, column=0, sticky="nsew")

        self.scroll = ttk.Scrollbar(frame_principal, command=self.texto_saida.yview)
        self.texto_saida.configure(yscrollcommand=self.scroll.set)
        self.scroll.grid(row=0, column=1, sticky="ns")

        frame_botoes = tk.Frame(frame_principal, bg=azul_escuro)
        frame_botoes.grid(row=0, column=2, sticky="ns", padx=(10, 0))

        icone_selecionar_normal = Image.open(f'{self.caminhoExe}\\imagens\\selecionar_arquivo.png').resize((254,54))
        icone_selecionar_normal = ImageTk.PhotoImage(icone_selecionar_normal)
        icone_selecionar_active = Image.open(f'{self.caminhoExe}\\imagens\\selecionar_arquivo_active.png').resize((254,54))
        icone_selecionar_active = ImageTk.PhotoImage(icone_selecionar_active)

        self.botao_selecionar = tk.Button(
            frame_botoes,
            command=self.selecionar_arquivo,
            image=icone_selecionar_normal,
            bg=azul_escuro,
            activebackground=azul_escuro,
            relief="flat",
            bd=0
        )
        self.botao_selecionar.image = icone_selecionar_normal
        self.botao_selecionar.bind("<Enter>", lambda e: self.botao_selecionar.config(image=icone_selecionar_active))
        self.botao_selecionar.bind("<Leave>", lambda e: self.botao_selecionar.config(image=icone_selecionar_normal))
        self.botao_selecionar.pack(pady=5)

        icone_iniciar_normal = Image.open(f'{self.caminhoExe}\\imagens\\iniciar_processo.png').resize((254,54))
        icone_iniciar_normal = ImageTk.PhotoImage(icone_iniciar_normal)
        icone_iniciar_active = Image.open(f'{self.caminhoExe}\\imagens\\iniciar_processo_active.png').resize((254,54))
        icone_iniciar_active = ImageTk.PhotoImage(icone_iniciar_active)
        self.botao_iniciar = tk.Button(
            frame_botoes,
            command=self.iniciar_processamento,
            image=icone_iniciar_normal,
            bg=azul_escuro,
            activebackground=azul_escuro,
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
            filetypes=[("Arquivos MBOX", "*.mbox"), ("Todos os arquivos", "*.*")]
        )
        if arquivo:
            self.arquivo_mbox = arquivo
            self.botao_iniciar.config(state="normal")
            self.adicionar_mensagem(f"Arquivo selecionado: {arquivo}")

    def iniciar_processamento(self, ):
        self.botao_iniciar.config(state="disabled")
        self.botao_selecionar.config(state="disabled")
        thread = Thread(target=processar_mbox_html, args=(self,))
        thread.start()

    def atualizar_titulo(self, texto):
        self.titulo.config(text=texto)

    def adicionar_mensagem(self, mensagem):
        self.texto_saida.tag_configure("fonte_grande", font=("Arial", 13))
        self.texto_saida.config(state=tk.NORMAL)
        self.texto_saida.insert(tk.END, mensagem + "\n", "fonte_grande")
        self.texto_saida.config(state=tk.DISABLED)
        self.texto_saida.see(tk.END)

if __name__ == "__main__":
    root = tk.Tk()
    tela = Recuperacao_email(root)
    root.mainloop()
