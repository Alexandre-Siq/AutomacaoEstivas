import sys
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from automacao_sshd import DadosSolicitante, ErroAutomacao, gerar_fichas_sshd

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("green")

COR_FUNDO = "#07130d"
COR_CARD = "#0f1f17"
COR_CARD_SECUNDARIO = "#10261b"
COR_BORDA = "#1f4f35"
COR_VERDE = "#22c55e"
COR_VERDE_ESCURO = "#15803d"
COR_TEXTO = "#f4f7f5"
COR_TEXTO_SECUNDARIO = "#9fb3a7"
COR_ALERTA = "#f97316"
COR_ERRO = "#ef4444"


def caminho_recurso(nome_arquivo):
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / nome_arquivo


class AutomacaoFichas:
    def __init__(self, root):
        self.root = root
        self.root.title("Automacao SSHD")
        self.root.geometry("780x680")
        self.root.minsize(720, 620)
        self.root.configure(fg_color=COR_FUNDO)

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
        main_frame = ctk.CTkFrame(self.root, fg_color=COR_FUNDO)
        main_frame.pack(pady=24, padx=24, fill="both", expand=True)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(3, weight=1)

        self._montar_cabecalho(main_frame)
        self._montar_passos(main_frame)
        self._montar_card_arquivo(main_frame)
        self._montar_conteudo_principal(main_frame)
        self._montar_rodape(main_frame)

    def _montar_cabecalho(self, parent):
        header_frame = ctk.CTkFrame(parent, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 18))
        header_frame.grid_columnconfigure(1, weight=1)

        marca = ctk.CTkFrame(
            header_frame,
            width=76,
            height=76,
            corner_radius=22,
            fg_color=COR_VERDE,
        )
        marca.grid(row=0, column=0, padx=(0, 18), sticky="n")
        marca.grid_propagate(False)

        ctk.CTkLabel(
            marca,
            text="SSHD",
            font=ctk.CTkFont(size=17, weight="bold"),
            text_color="#052e16",
        ).place(relx=0.5, rely=0.5, anchor="center")

        titulo = ctk.CTkLabel(
            header_frame,
            text="Gerador de Fichas SSHD",
            font=ctk.CTkFont(size=30, weight="bold"),
            text_color=COR_TEXTO,
            anchor="w",
        )
        titulo.grid(row=0, column=1, sticky="sw")

        subtitulo = ctk.CTkLabel(
            header_frame,
            text=(
                "Converta a fonte de dados em fichas padronizadas da Prefeitura, "
                "com layout preservado."
            ),
            font=ctk.CTkFont(size=14),
            text_color=COR_TEXTO_SECUNDARIO,
            anchor="w",
            wraplength=560,
        )
        subtitulo.grid(row=1, column=1, sticky="nw", pady=(6, 0))

    def _montar_passos(self, parent):
        steps_frame = ctk.CTkFrame(parent, fg_color="transparent")
        steps_frame.grid(row=1, column=0, sticky="ew", pady=(0, 18))
        for coluna in range(3):
            steps_frame.grid_columnconfigure(coluna, weight=1)

        passos = [
            ("1", "Selecione a fonte", "Arquivo .xlsx com os profissionais"),
            ("2", "Informe o solicitante", "Nome, SSHD e cargo"),
            ("3", "Gere o resultado", "Uma aba por profissional"),
        ]

        for coluna, (numero, titulo, descricao) in enumerate(passos):
            card = ctk.CTkFrame(
                steps_frame,
                fg_color=COR_CARD,
                border_width=1,
                border_color=COR_BORDA,
                corner_radius=16,
            )
            card.grid(row=0, column=coluna, sticky="ew", padx=(0 if coluna == 0 else 8, 0 if coluna == 2 else 8))
            card.grid_columnconfigure(1, weight=1)

            bolha = ctk.CTkFrame(card, width=34, height=34, corner_radius=17, fg_color=COR_VERDE)
            bolha.grid(row=0, column=0, rowspan=2, padx=14, pady=14, sticky="n")
            bolha.grid_propagate(False)
            ctk.CTkLabel(
                bolha,
                text=numero,
                font=ctk.CTkFont(size=14, weight="bold"),
                text_color="#052e16",
            ).place(relx=0.5, rely=0.5, anchor="center")

            ctk.CTkLabel(
                card,
                text=titulo,
                font=ctk.CTkFont(size=13, weight="bold"),
                text_color=COR_TEXTO,
                anchor="w",
            ).grid(row=0, column=1, sticky="ew", padx=(0, 12), pady=(13, 0))
            ctk.CTkLabel(
                card,
                text=descricao,
                font=ctk.CTkFont(size=11),
                text_color=COR_TEXTO_SECUNDARIO,
                anchor="w",
            ).grid(row=1, column=1, sticky="ew", padx=(0, 12), pady=(0, 13))

    def _montar_card_arquivo(self, parent):
        file_card = ctk.CTkFrame(
            parent,
            fg_color=COR_CARD,
            border_width=1,
            border_color=COR_BORDA,
            corner_radius=18,
        )
        file_card.grid(row=2, column=0, sticky="ew", pady=(0, 18))
        file_card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            file_card,
            text="Planilha fonte de dados",
            font=ctk.CTkFont(size=15, weight="bold"),
            text_color=COR_TEXTO,
            anchor="w",
        ).grid(row=0, column=0, columnspan=2, sticky="ew", padx=18, pady=(16, 4))

        ctk.CTkLabel(
            file_card,
            text="Selecione o arquivo Excel recebido com os dados dos profissionais.",
            font=ctk.CTkFont(size=12),
            text_color=COR_TEXTO_SECUNDARIO,
            anchor="w",
        ).grid(row=1, column=0, columnspan=2, sticky="ew", padx=18, pady=(0, 12))

        self.entrada_arquivo = ctk.CTkEntry(
            file_card,
            textvariable=self.caminho_base_mae,
            placeholder_text="Nenhum arquivo selecionado",
            height=42,
            corner_radius=12,
            fg_color="#07130d",
            border_color=COR_BORDA,
            text_color=COR_TEXTO,
            placeholder_text_color=COR_TEXTO_SECUNDARIO,
        )
        self.entrada_arquivo.grid(row=2, column=0, sticky="ew", padx=(18, 10), pady=(0, 18))
        
        btn_procurar = ctk.CTkButton(
            file_card,
            text="Procurar arquivo",
            command=self.selecionar_arquivo,
            width=150,
            height=42,
            corner_radius=12,
            fg_color=COR_VERDE,
            hover_color=COR_VERDE_ESCURO,
            text_color="#052e16",
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        btn_procurar.grid(row=2, column=1, padx=(0, 18), pady=(0, 18))

    def _montar_conteudo_principal(self, parent):
        content_frame = ctk.CTkFrame(parent, fg_color="transparent")
        content_frame.grid(row=3, column=0, sticky="nsew")
        content_frame.grid_columnconfigure(0, weight=3)
        content_frame.grid_columnconfigure(1, weight=2)

        self._montar_card_solicitante(content_frame)
        self._montar_card_regras(content_frame)

    def _montar_card_solicitante(self, parent):
        dados_frame = ctk.CTkFrame(
            parent,
            fg_color=COR_CARD,
            border_width=1,
            border_color=COR_BORDA,
            corner_radius=18,
        )
        dados_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 9))
        dados_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            dados_frame,
            text="Dados do solicitante",
            font=ctk.CTkFont(size=17, weight="bold"),
            text_color=COR_TEXTO,
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=18, pady=(18, 4))

        ctk.CTkLabel(
            dados_frame,
            text="Esses dados serao repetidos no cabecalho de todas as abas geradas.",
            font=ctk.CTkFont(size=12),
            text_color=COR_TEXTO_SECUNDARIO,
            anchor="w",
            wraplength=420,
        ).grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 14))
        
        self.entrada_nome = self._criar_campo(
            dados_frame,
            linha=2,
            rotulo="Nome do solicitante",
            placeholder="Ex: Alexandre Siqueira Souza Costa",
        )
        self.entrada_sshd = self._criar_campo(
            dados_frame,
            linha=3,
            rotulo="SSHD",
            placeholder="Ex: X0801681",
        )
        self.entrada_cargo = self._criar_campo(
            dados_frame,
            linha=4,
            rotulo="Cargo",
            placeholder="Ex: ANALISTA DE SUPORTE I",
        )

    def _montar_card_regras(self, parent):
        regras_frame = ctk.CTkFrame(
            parent,
            fg_color=COR_CARD_SECUNDARIO,
            border_width=1,
            border_color=COR_BORDA,
            corner_radius=18,
        )
        regras_frame.grid(row=0, column=1, sticky="nsew", padx=(9, 0))
        regras_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            regras_frame,
            text="Regras aplicadas",
            font=ctk.CTkFont(size=17, weight="bold"),
            text_color=COR_TEXTO,
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=18, pady=(18, 8))

        regras = [
            ("Template protegido", "A estrutura do TEMPLATE_NOVO.xlsx e preservada."),
            ("Campos sensiveis", "Genero e Orientacao Sexual saem como 'não informado'."),
            ("Empresa e unidade", "COMPLEXO HOSPITALAR DOS ESTIVADORES / SMS."),
            ("Resultado", "Uma aba preenchida para cada profissional."),
        ]

        for linha, (titulo, texto) in enumerate(regras, start=1):
            item = ctk.CTkFrame(regras_frame, fg_color="#0b1a12", corner_radius=14)
            item.grid(row=linha, column=0, sticky="ew", padx=18, pady=(0, 10))
            item.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(
                item,
                text=titulo,
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color=COR_VERDE,
                anchor="w",
            ).grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 0))
            ctk.CTkLabel(
                item,
                text=texto,
                font=ctk.CTkFont(size=11),
                text_color=COR_TEXTO_SECUNDARIO,
                anchor="w",
                wraplength=230,
            ).grid(row=1, column=0, sticky="ew", padx=12, pady=(2, 10))

    def _montar_rodape(self, parent):
        footer = ctk.CTkFrame(parent, fg_color="transparent")
        footer.grid(row=4, column=0, sticky="ew", pady=(18, 0))
        footer.grid_columnconfigure(0, weight=1)

        actions = ctk.CTkFrame(footer, fg_color="transparent")
        actions.grid(row=0, column=0, sticky="ew")
        actions.grid_columnconfigure(0, weight=1)

        self.btn_limpar = ctk.CTkButton(
            actions,
            text="Limpar campos",
            command=self.limpar_campos,
            height=46,
            width=150,
            corner_radius=14,
            fg_color="#173526",
            hover_color="#1f4f35",
            border_width=1,
            border_color=COR_BORDA,
            text_color=COR_TEXTO,
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.btn_limpar.grid(row=0, column=0, sticky="w")

        self.btn_executar = ctk.CTkButton(
            actions,
            text="Gerar fichas SSHD",
            command=self.executar_processo,
            font=ctk.CTkFont(size=15, weight="bold"),
            height=46,
            width=220,
            corner_radius=14,
            fg_color=COR_VERDE,
            hover_color=COR_VERDE_ESCURO,
            text_color="#052e16",
        )
        self.btn_executar.grid(row=0, column=1, sticky="e")

        status_card = ctk.CTkFrame(
            footer,
            fg_color=COR_CARD,
            border_width=1,
            border_color=COR_BORDA,
            corner_radius=16,
        )
        status_card.grid(row=1, column=0, sticky="ew", pady=(14, 0))
        status_card.grid_columnconfigure(0, weight=1)

        self.status_label = ctk.CTkLabel(
            status_card,
            text="Pronto para iniciar.",
            font=ctk.CTkFont(size=12),
            text_color=COR_TEXTO_SECUNDARIO,
            anchor="w",
        )
        self.status_label.grid(row=0, column=0, sticky="ew", padx=16, pady=(12, 4))

        self.progress_bar = ctk.CTkProgressBar(
            status_card,
            height=8,
            corner_radius=8,
            progress_color=COR_VERDE,
            fg_color="#07130d",
        )
        self.progress_bar.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 12))
        self.progress_bar.set(0)

    def _criar_campo(self, parent, linha, rotulo, placeholder):
        ctk.CTkLabel(
            parent,
            text=rotulo,
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=COR_TEXTO,
            anchor="w",
        ).grid(row=linha * 2, column=0, sticky="ew", padx=18, pady=(0 if linha == 2 else 10, 5))

        entrada = ctk.CTkEntry(
            parent,
            placeholder_text=placeholder,
            height=42,
            corner_radius=12,
            fg_color="#07130d",
            border_color=COR_BORDA,
            text_color=COR_TEXTO,
            placeholder_text_color=COR_TEXTO_SECUNDARIO,
        )
        entrada.grid(row=linha * 2 + 1, column=0, sticky="ew", padx=18, pady=(0, 2))
        return entrada

    def selecionar_arquivo(self):
        arquivo = filedialog.askopenfilename(
            title="Selecionar Planilha Geral",
            filetypes=[("Arquivos Excel", "*.xlsx")]
        )
        if arquivo:
            self.caminho_base_mae.set(arquivo)
            self.status_label.configure(
                text=f"Arquivo selecionado: {Path(arquivo).name}",
                text_color=COR_VERDE,
            )

    def limpar_campos(self):
        self.caminho_base_mae.set("")
        for entrada in (self.entrada_nome, self.entrada_sshd, self.entrada_cargo):
            entrada.delete(0, "end")
        self.progress_bar.stop()
        self.progress_bar.set(0)
        self.status_label.configure(text="Campos limpos. Pronto para iniciar.", text_color=COR_TEXTO_SECUNDARIO)

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

        self.status_label.configure(text="Processando arquivo... isso pode levar alguns segundos.", text_color=COR_VERDE)
        self.progress_bar.start()
        self.btn_executar.configure(state="disabled")
        self.btn_limpar.configure(state="disabled")
        self.root.update()

        try:
            resultado = gerar_fichas_sshd(
                caminho_planilha_geral=caminho_mae,
                solicitante=solicitante,
                caminho_template=caminho_template,
            )

            self.status_label.configure(
                text=f"Concluido: {resultado.total_colaboradores} ficha(s) gerada(s).",
                text_color=COR_VERDE,
            )
            messagebox.showinfo(
                "Sucesso",
                (
                    f"Automação concluída!\n"
                    f"Fichas geradas: {resultado.total_colaboradores}\n"
                    f"Arquivo salvo em:\n{resultado.caminho_saida}"
                ),
            )

        except ErroAutomacao as erro:
            self.status_label.configure(text="Nao foi possivel gerar as fichas.", text_color=COR_ALERTA)
            messagebox.showerror("Erro", str(erro))
        except Exception as e:
            self.status_label.configure(text="Erro inesperado no processamento.", text_color=COR_ERRO)
            messagebox.showerror("Erro", f"Ocorreu um problema: {str(e)}")

        finally:
            self.progress_bar.stop()
            self.progress_bar.set(0)
            self.btn_executar.configure(state="normal")
            self.btn_limpar.configure(state="normal")

if __name__ == "__main__":
    root = ctk.CTk()
    app_gui = AutomacaoFichas(root)
    root.mainloop()