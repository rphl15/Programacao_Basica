import tkinter as tk
from tkinter import filedialog, messagebox
from funcao import gerar_relatorio
from PIL import Image, ImageTk
from pathlib import Path
import sys

class Aplicacao:

    def __init__(self):

        self.janela = tk.Tk()

        self.janela.title("Automação Financeiro")

        self.janela.geometry("700x450")

        self.janela.resizable(False, False)

        self.janela.configure(bg="#dad7cd")
        
        self.arquivo_hm = ""

        self.arquivo_processado = ""

        self.criar_interface()

        self.janela.mainloop()

    def criar_interface(self):
        
        # Frame para agrupar os logos
        frame_logos = tk.Frame(self.janela, bg="#dad7cd")
        frame_logos.pack(pady=10)

        # Logo esquerda
        img1 = Image.open(self.recurso("viva.png")).convert("RGBA")
        img1 = img1.resize((120, 120))
        self.logo1 = ImageTk.PhotoImage(img1)

        tk.Label(
            frame_logos,
            image=self.logo1,
            bg="#dad7cd"
        ).pack(side="left", padx=20)

        # Logo principal
        img2 = Image.open(self.recurso("cma.png")).convert("RGBA")
        img2 = img2.resize((90, 90))
        self.logo2 = ImageTk.PhotoImage(img2)

        tk.Label(
            frame_logos,
            image=self.logo2,
            bg="#dad7cd"
        ).pack(side="left", padx=20)

        # Logo direita
        img3 = Image.open(self.recurso("arqui.png")).convert("RGBA")
        img3 = img3.resize((90, 90))
        self.logo3 = ImageTk.PhotoImage(img3)

        tk.Label(
            frame_logos,
            image=self.logo3,
            bg="#dad7cd"
        ).pack(side="left", padx=20)
                
        titulo = tk.Label(
            self.janela,
            text="AUTOMAÇÃO - FINANCEIRO",
            font=("Segoe UI", 18, "bold"),
            bg="#dad7cd",
            fg="#220901"
        )

        titulo.pack(pady=20)

        # ===========================
        # HM
        # ===========================

        frame1 = tk.Frame(self.janela)

        frame1.pack(fill="x", padx=20, pady=5)

        self.label_hm = tk.Label(
            frame1,
            text="Nenhum arquivo selecionado",
            anchor="w",
            width=60
        )

        self.label_hm.pack(side="left")

        btn_hm = tk.Button(
            frame1,
            text="Selecionar HM",
            command=self.selecionar_hm,
            width=20,
            bg="#023047",
            fg="white",
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            cursor="hand2"
        )
        btn_hm.pack(side="right")

        # ===========================
        # PROCESSADO
        # ===========================

        frame2 = tk.Frame(self.janela)

        frame2.pack(fill="x", padx=20, pady=5)

        self.label_proc = tk.Label(
            frame2,
            text="Nenhum arquivo selecionado",
            anchor="w",
            width=60
        )

        self.label_proc.pack(side="left")

        btn_proc = tk.Button(
            frame2,
            text="Selecionar Processado",
            command=self.selecionar_processado,
            width=20,
            bg="#023047",
            fg="white",
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            cursor="hand2"
        )
        btn_proc.pack(side="right")

        # ===========================

        btn_processar = tk.Button(
            self.janela,
            text="▶ PROCESSAR",
            command=self.processar,
            width=25,
            height=2,
            bg="#023047",
            fg="white",
            font=("Segoe UI", 11, "bold"),
            relief="flat",
            cursor="hand2"
        )

        btn_processar.pack(pady=20)

        self.status = tk.Label(
            self.janela,
            text="Aguardando arquivos...",
            bg="#dad7cd"
        )

        self.status.pack()

    def selecionar_hm(self):

        arquivo = filedialog.askopenfilename(
            filetypes=[("Excel", "*.xlsx")]
        )

        if arquivo:

            self.arquivo_hm = arquivo

            self.label_hm.config(
                text=arquivo.split("/")[-1]
            )

    def selecionar_processado(self):

        arquivo = filedialog.askopenfilename(
            filetypes=[("Excel", "*.xlsx")]
        )

        if arquivo:

            self.arquivo_processado = arquivo

            self.label_proc.config(
                text=arquivo.split("/")[-1]
            )

    def recurso(self, nome):
        if getattr(sys, "frozen", False):
            base = Path(sys._MEIPASS)
        else:
            base = Path(__file__).parent

        return base / nome


    def processar(self):

        if self.arquivo_hm == "":

            messagebox.showwarning(
                "Aviso",
                "Selecione a planilha HM."
            )

            return

        if self.arquivo_processado == "":

            messagebox.showwarning(
                "Aviso",
                "Selecione a planilha Processado."
            )

            return

        self.status.config(
            text="Processando...",
            bg="#dad7cd"
        )

        self.janela.update()

        try:

            mensagem = gerar_relatorio(
                self.arquivo_hm,
                self.arquivo_processado
            )

            self.status.config(
                text="Processo finalizado.",
                bg="#dad7cd"
            )

            messagebox.showinfo(
                "Sucesso",
                mensagem
            )

        except Exception as erro:

            self.status.config(
                text="Erro."
            )

            messagebox.showerror(
                "Erro",
                str(erro)
            )


if __name__ == "__main__":

    Aplicacao()