import sys
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from automacao_sshd import DadosSolicitante, ErroAutomacao, gerar_fichas_sshd

ctk.set_appearance_mode("System")  
ctk.set_default_color_theme("blue") 


def caminho_recurso(nome_arquivo):
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / nome_arquivo


class AutomacaoFichas:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de Planilhas SSHD")
        
        self.root.geometry("620x520")

        icone = caminho_recurso("icone_estivas.ico")
        if icone.exists():
            try:
                self.root.iconbitmap(str(icone))
            except Exception:
                pass
        
        # Variavel do caminho da Planilha Geral.
        self.caminho_base_mae = ctk.StringVar()
        
        self.setup_ui()

    def setup_ui(self):
        main_frame = ctk.CTkFrame(self.root, corner_radius=15)
        main_frame.pack(pady=20, padx=20, fill="both", expand=True)

        titulo = ctk.CTkLabel(
            main_frame,
            text="Gerador de Fichas SSHD",
            font=ctk.CTkFont(size=20, weight="bold"),
        )
        titulo.pack(pady=(15, 10))

        subtitulo = ctk.CTkLabel(
            main_frame,
            text="Planilha Geral de colaboradores -> Modelo padrao da Prefeitura",
            text_color="gray",
        )
        subtitulo.pack(pady=(0, 15))

        # --- SECAO 1: ARQUIVO ---
        file_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        file_frame.pack(fill="x", padx=20, pady=(0, 15))
        
        entrada_arquivo = ctk.CTkEntry(
            file_frame,
            textvariable=self.caminho_base_mae,
            placeholder_text="Selecione a Planilha Geral (.xlsx)...",
            width=410,
            height=35,
        )
        entrada_arquivo.pack(side="left", padx=(0, 10))
        
        btn_procurar = ctk.CTkButton(file_frame, text="Procurar", command=self.selecionar_arquivo, width=100, height=35)
        btn_procurar.pack(side="right")

        # --- SECAO 2: DADOS DO SOLICITANTE ---
        dados_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        dados_frame.pack(fill="x", padx=20)
        
        ctk.CTkLabel(
            dados_frame,
            text="Dados do Solicitante (repetidos em todas as abas):",
            font=ctk.CTkFont(size=12, weight="bold"),
        ).pack(anchor="w", pady=(0, 5))
        
        # Campos de entrada
        self.entrada_nome = ctk.CTkEntry(dados_frame, placeholder_text="Nome do Solicitante (Ex: Alexandre Siqueira...)", height=30)
        self.entrada_nome.pack(fill="x", pady=3)
        
        self.entrada_sshd = ctk.CTkEntry(dados_frame, placeholder_text="SSHD (Ex: X0801681)", height=30)
        self.entrada_sshd.pack(fill="x", pady=3)
        
        self.entrada_cargo = ctk.CTkEntry(dados_frame, placeholder_text="Cargo (Ex: ANALISTA DE SUPORTE I)", height=30)
        self.entrada_cargo.pack(fill="x", pady=3)

        # --- BOTAO EXECUTAR ---
        self.btn_executar = ctk.CTkButton(
            main_frame,
            text="GERAR FICHAS",
            command=self.executar_processo,
            font=ctk.CTkFont(size=15, weight="bold"),
            height=45,
            fg_color="#28a745",
            hover_color="#218838",
        )
        self.btn_executar.pack(pady=(20, 10))

        self.status_label = ctk.CTkLabel(main_frame, text="Aguardando início...", text_color="gray")
        self.status_label.pack()

        observacao = ctk.CTkLabel(
            main_frame,
            text=(
                "Nesta primeira etapa o processamento considera o formato da "
                "Planilha Geral enviado nas imagens."
            ),
            text_color="gray",
            wraplength=540,
        )
        observacao.pack(pady=(10, 0))

    def selecionar_arquivo(self):
        arquivo = filedialog.askopenfilename(
            title="Selecionar Planilha Geral",
            filetypes=[("Arquivos Excel", "*.xlsx")]
        )
        if arquivo:
            self.caminho_base_mae.set(arquivo)

    def executar_processo(self):
        caminho_mae = self.caminho_base_mae.get()
        
        if not caminho_mae:
            messagebox.showwarning("Atenção", "Por favor, selecione a planilha mãe primeiro.")
            return

        caminho_template = caminho_recurso("TEMPLATE_NOVO.xlsx")
        if not caminho_template.exists():
            caminho_template = Path(__file__).resolve().parent / "TEMPLATE_NOVO.xlsx"
        if not caminho_template.exists():
            messagebox.showerror("Erro", "Arquivo 'TEMPLATE_NOVO.xlsx' não encontrado na pasta do programa.")
            return

        solicitante = DadosSolicitante(
            nome=self.entrada_nome.get(),
            sshd=self.entrada_sshd.get(),
            cargo=self.entrada_cargo.get(),
        )

        self.status_label.configure(text="Processando... por favor, aguarde.", text_color="#17a2b8")
        self.btn_executar.configure(state="disabled")
        self.root.update()

        try:
            resultado = gerar_fichas_sshd(
                caminho_planilha_geral=caminho_mae,
                solicitante=solicitante,
                caminho_template=caminho_template,
            )

            self.status_label.configure(text="Concluído com sucesso!", text_color="#28a745")
            messagebox.showinfo(
                "Sucesso",
                (
                    f"Automação concluída!\n"
                    f"Fichas geradas: {resultado.total_colaboradores}\n"
                    f"Arquivo salvo em:\n{resultado.caminho_saida}"
                ),
            )

        except ErroAutomacao as erro:
            self.status_label.configure(text="Erro no processamento.", text_color="red")
            messagebox.showerror("Erro", str(erro))
        except Exception as e:
            self.status_label.configure(text="Erro no processamento.", text_color="red")
            messagebox.showerror("Erro", f"Ocorreu um problema: {str(e)}")

        finally:
            self.btn_executar.configure(state="normal")

if __name__ == "__main__":
    root = ctk.CTk()
    app_gui = AutomacaoFichas(root)
    root.mainloop()