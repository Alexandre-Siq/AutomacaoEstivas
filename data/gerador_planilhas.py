import os
import subprocess
import sys
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from automacao_sshd import DadosSolicitante, ErroAutomacao, gerar_fichas_sshd
from automacao_sshd.normalizacao import nome_arquivo_seguro


ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("green")

COR_FUNDO = "#07130d"
COR_CARD = "#0f1f17"
COR_BORDA = "#1f4f35"
COR_VERDE = "#22c55e"
COR_VERDE_ESCURO = "#15803d"
COR_TEXTO = "#f4f7f5"
COR_TEXTO_SECUNDARIO = "#9fb3a7"
COR_ALERTA = "#f97316"
COR_ERRO = "#ef4444"
FONTE_PADRAO = "Oswald"


def fonte(tamanho, peso="normal"):
    return ctk.CTkFont(family=FONTE_PADRAO, size=tamanho, weight=peso)


def caminho_recurso(nome_arquivo):
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / nome_arquivo


class AutomacaoFichas:
    def __init__(self, root):
        self.root = root
        self.root.title("Automação SSHD")
        self.root.geometry("680x660")
        self.root.resizable(False, False)
        self.root.configure(fg_color=COR_FUNDO)

        icone = caminho_recurso("icone_estivas.ico")
        if icone.exists():
            try:
                self.root.iconbitmap(str(icone))
            except Exception:
                pass

        self.caminho_base_mae = ctk.StringVar()
        self.pasta_saida = ctk.StringVar()
        self.ultimo_arquivo_gerado: Path | None = None
        self.setup_ui()

    def setup_ui(self):
        main_frame = ctk.CTkFrame(self.root, fg_color=COR_FUNDO)
        main_frame.pack(pady=18, padx=18, fill="both", expand=True)
        main_frame.grid_columnconfigure(0, weight=1)

        self._montar_cabecalho(main_frame)
        self._montar_card_arquivo(main_frame)
        self._montar_card_solicitante(main_frame)
        self._montar_acoes(main_frame)
        self._montar_status(main_frame)

    def _montar_cabecalho(self, parent):
        header = ctk.CTkFrame(parent, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 16))
        header.grid_columnconfigure(1, weight=1)

        marca = ctk.CTkFrame(header, width=56, height=56, corner_radius=16, fg_color=COR_VERDE)
        marca.grid(row=0, column=0, rowspan=2, padx=(0, 14), sticky="n")
        marca.grid_propagate(False)

        ctk.CTkLabel(
            marca,
            text="SSHD",
            font=fonte(13, "bold"),
            text_color="#052e16",
        ).place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(
            header,
            text="Gerador de Fichas SSHD",
            font=fonte(26, "bold"),
            text_color=COR_TEXTO,
            anchor="w",
        ).grid(row=0, column=1, sticky="sw")

        ctk.CTkLabel(
            header,
            text="Selecione a fonte, informe o solicitante e gere o arquivo final.",
            font=fonte(13),
            text_color=COR_TEXTO_SECUNDARIO,
            anchor="w",
        ).grid(row=1, column=1, sticky="nw", pady=(2, 0))

        ctk.CTkLabel(
            header,
            text="Desenvolvido por Alexandre Siqueira - Analista de Suporte",
            font=fonte(12),
            text_color=COR_TEXTO_SECUNDARIO,
            anchor="w",
        ).grid(row=2, column=1, sticky="nw", pady=(2, 0))

    def _montar_card_arquivo(self, parent):
        card = self._criar_card(parent, row=1)
        card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            card,
            text="Arquivos",
            font=fonte(15, "bold"),
            text_color=COR_TEXTO,
            anchor="w",
        ).grid(row=0, column=0, columnspan=2, sticky="ew", padx=16, pady=(14, 8))

        ctk.CTkLabel(
            card,
            text="Planilha fonte",
            font=fonte(12, "bold"),
            text_color=COR_TEXTO,
            anchor="w",
        ).grid(row=1, column=0, columnspan=2, sticky="ew", padx=16, pady=(0, 5))

        self.entrada_arquivo = ctk.CTkEntry(
            card,
            textvariable=self.caminho_base_mae,
            placeholder_text="Nenhum arquivo selecionado",
            height=38,
            corner_radius=12,
            fg_color=COR_FUNDO,
            border_color=COR_BORDA,
            text_color=COR_TEXTO,
            placeholder_text_color=COR_TEXTO_SECUNDARIO,
            font=fonte(13),
        )
        self.entrada_arquivo.grid(row=2, column=0, sticky="ew", padx=(16, 10), pady=(0, 12))

        ctk.CTkButton(
            card,
            text="Procurar",
            command=self.selecionar_arquivo,
            width=120,
            height=38,
            corner_radius=12,
            fg_color=COR_VERDE,
            hover_color=COR_VERDE_ESCURO,
            text_color="#052e16",
            font=fonte(14, "bold"),
        ).grid(row=2, column=1, padx=(0, 16), pady=(0, 12))

        ctk.CTkLabel(
            card,
            text="Pasta de saída",
            font=fonte(12, "bold"),
            text_color=COR_TEXTO,
            anchor="w",
        ).grid(row=3, column=0, columnspan=2, sticky="ew", padx=16, pady=(0, 5))

        self.entrada_saida = ctk.CTkEntry(
            card,
            textvariable=self.pasta_saida,
            placeholder_text="Usar mesma pasta da planilha fonte",
            height=38,
            corner_radius=12,
            fg_color=COR_FUNDO,
            border_color=COR_BORDA,
            text_color=COR_TEXTO,
            placeholder_text_color=COR_TEXTO_SECUNDARIO,
            font=fonte(13),
        )
        self.entrada_saida.grid(row=4, column=0, sticky="ew", padx=(16, 10), pady=(0, 16))

        ctk.CTkButton(
            card,
            text="Escolher",
            command=self.selecionar_pasta_saida,
            width=120,
            height=38,
            corner_radius=12,
            fg_color=COR_VERDE,
            hover_color=COR_VERDE_ESCURO,
            text_color="#052e16",
            font=fonte(14, "bold"),
        ).grid(row=4, column=1, padx=(0, 16), pady=(0, 16))

    def _montar_card_solicitante(self, parent):
        card = self._criar_card(parent, row=2)
        for coluna in range(3):
            card.grid_columnconfigure(coluna, weight=1)

        ctk.CTkLabel(
            card,
            text="Dados do solicitante",
            font=fonte(15, "bold"),
            text_color=COR_TEXTO,
            anchor="w",
        ).grid(row=0, column=0, columnspan=3, sticky="ew", padx=16, pady=(14, 8))

        self.entrada_nome = self._criar_campo(card, coluna=0, rotulo="Nome", placeholder="Solicitante")
        self.entrada_sshd = self._criar_campo(card, coluna=1, rotulo="SSHD", placeholder="X0000000")
        self.entrada_cargo = self._criar_campo(card, coluna=2, rotulo="Cargo", placeholder="Cargo")

    def _montar_acoes(self, parent):
        actions = ctk.CTkFrame(parent, fg_color="transparent")
        actions.grid(row=3, column=0, sticky="ew", pady=(4, 12))
        actions.grid_columnconfigure(0, weight=1)

        self.btn_limpar = ctk.CTkButton(
            actions,
            text="Limpar",
            command=self.limpar_campos,
            height=42,
            width=120,
            corner_radius=14,
            fg_color="#173526",
            hover_color="#1f4f35",
            border_width=1,
            border_color=COR_BORDA,
            text_color=COR_TEXTO,
            font=fonte(14, "bold"),
        )
        self.btn_limpar.grid(row=0, column=0, sticky="w")

        self.btn_abrir = ctk.CTkButton(
            actions,
            text="Abrir arquivo",
            command=self.abrir_arquivo_gerado,
            height=42,
            width=150,
            corner_radius=14,
            fg_color="#173526",
            hover_color="#1f4f35",
            border_width=1,
            border_color=COR_BORDA,
            text_color=COR_TEXTO,
            font=fonte(14, "bold"),
            state="disabled",
        )
        self.btn_abrir.grid(row=0, column=1, sticky="e", padx=(0, 14))

        self.btn_executar = ctk.CTkButton(
            actions,
            text="Gerar fichas SSHD",
            command=self.executar_processo,
            height=42,
            width=190,
            corner_radius=14,
            fg_color=COR_VERDE,
            hover_color=COR_VERDE_ESCURO,
            text_color="#052e16",
            font=fonte(16, "bold"),
        )
        self.btn_executar.grid(row=0, column=2, sticky="e")

    def _montar_status(self, parent):
        status_card = self._criar_card(parent, row=4, pady=(0, 0))
        status_card.grid_columnconfigure(0, weight=1)

        self.status_label = ctk.CTkLabel(
            status_card,
            text="Pronto para iniciar.",
            font=fonte(13),
            text_color=COR_TEXTO_SECUNDARIO,
            anchor="w",
        )
        self.status_label.grid(row=0, column=0, sticky="ew", padx=14, pady=(10, 4))

        self.progress_bar = ctk.CTkProgressBar(
            status_card,
            height=8,
            corner_radius=8,
            progress_color=COR_VERDE,
            fg_color=COR_FUNDO,
        )
        self.progress_bar.grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 10))
        self.progress_bar.set(0)

        ctk.CTkLabel(
            status_card,
            text="Template preservado | Gênero e orientação sexual: não informado",
            font=fonte(12),
            text_color=COR_TEXTO_SECUNDARIO,
            anchor="w",
        ).grid(row=2, column=0, sticky="ew", padx=14, pady=(0, 10))

    def _criar_card(self, parent, row, pady=(0, 12)):
        card = ctk.CTkFrame(
            parent,
            fg_color=COR_CARD,
            border_width=1,
            border_color=COR_BORDA,
            corner_radius=18,
        )
        card.grid(row=row, column=0, sticky="ew", pady=pady)
        return card

    def _criar_campo(self, parent, coluna, rotulo, placeholder):
        padx = (16 if coluna == 0 else 6, 16 if coluna == 2 else 6)

        ctk.CTkLabel(
            parent,
            text=rotulo,
            font=fonte(13, "bold"),
            text_color=COR_TEXTO,
            anchor="w",
        ).grid(row=1, column=coluna, sticky="ew", padx=padx, pady=(0, 5))

        entrada = ctk.CTkEntry(
            parent,
            placeholder_text=placeholder,
            height=38,
            corner_radius=12,
            fg_color=COR_FUNDO,
            border_color=COR_BORDA,
            text_color=COR_TEXTO,
            placeholder_text_color=COR_TEXTO_SECUNDARIO,
            font=fonte(13),
        )
        entrada.grid(row=2, column=coluna, sticky="ew", padx=padx, pady=(0, 16))
        return entrada

    def selecionar_arquivo(self):
        arquivo = filedialog.askopenfilename(
            title="Selecionar planilha fonte",
            filetypes=[("Arquivos Excel", "*.xlsx")],
        )
        if arquivo:
            self.caminho_base_mae.set(arquivo)
            if not self.pasta_saida.get().strip():
                self.pasta_saida.set(str(Path(arquivo).parent))
            self.status_label.configure(
                text=f"Arquivo selecionado: {Path(arquivo).name}",
                text_color=COR_VERDE,
            )

    def selecionar_pasta_saida(self):
        pasta = filedialog.askdirectory(title="Selecionar pasta de saída")
        if pasta:
            self.pasta_saida.set(pasta)
            self.status_label.configure(
                text=f"Pasta de saída: {Path(pasta).name}",
                text_color=COR_VERDE,
            )

    def limpar_campos(self):
        self.caminho_base_mae.set("")
        self.pasta_saida.set("")
        self.ultimo_arquivo_gerado = None
        for entrada in (self.entrada_nome, self.entrada_sshd, self.entrada_cargo):
            entrada.delete(0, "end")
        self.progress_bar.stop()
        self.progress_bar.set(0)
        self.btn_abrir.configure(state="disabled")
        self.status_label.configure(
            text="Campos limpos. Pronto para iniciar.",
            text_color=COR_TEXTO_SECUNDARIO,
        )

    def abrir_arquivo_gerado(self):
        if not self.ultimo_arquivo_gerado or not self.ultimo_arquivo_gerado.exists():
            messagebox.showwarning("Atenção", "Nenhum arquivo gerado foi encontrado para abrir.")
            return

        try:
            if sys.platform.startswith("win"):
                os.startfile(self.ultimo_arquivo_gerado)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(self.ultimo_arquivo_gerado)])
            else:
                subprocess.Popen(["xdg-open", str(self.ultimo_arquivo_gerado)])
        except Exception as erro:
            messagebox.showerror("Erro", f"Não foi possível abrir o arquivo: {erro}")

    def executar_processo(self):
        caminho_mae = self.caminho_base_mae.get()

        if not caminho_mae:
            messagebox.showwarning("Atenção", "Selecione a planilha fonte primeiro.")
            return

        caminho_template = caminho_recurso("TEMPLATE_NOVO.xlsx")
        if not caminho_template.exists():
            caminho_template = Path(__file__).resolve().parent / "TEMPLATE_NOVO.xlsx"
        if not caminho_template.exists():
            messagebox.showerror("Erro", "Arquivo 'TEMPLATE_NOVO.xlsx' não encontrado.")
            return

        solicitante = DadosSolicitante(
            nome=self.entrada_nome.get(),
            sshd=self.entrada_sshd.get(),
            cargo=self.entrada_cargo.get(),
        )

        pasta_saida = Path(self.pasta_saida.get().strip() or Path(caminho_mae).parent)
        nome_saida = f"{nome_arquivo_seguro(Path(caminho_mae).stem)}_SSHD.xlsx"
        caminho_saida = pasta_saida / nome_saida

        self.status_label.configure(text="Processando arquivo...", text_color=COR_VERDE)
        self.progress_bar.start()
        self.btn_executar.configure(state="disabled")
        self.btn_limpar.configure(state="disabled")
        self.btn_abrir.configure(state="disabled")
        self.root.update()

        try:
            resultado = gerar_fichas_sshd(
                caminho_planilha_geral=caminho_mae,
                solicitante=solicitante,
                caminho_template=caminho_template,
                caminho_saida=caminho_saida,
            )
            self.ultimo_arquivo_gerado = resultado.caminho_saida

            self.status_label.configure(
                text=f"Concluído: {resultado.total_colaboradores} ficha(s) gerada(s).",
                text_color=COR_VERDE,
            )
            self.btn_abrir.configure(state="normal")
            messagebox.showinfo(
                "Sucesso",
                (
                    "Automação concluída!\n"
                    f"Fichas geradas: {resultado.total_colaboradores}\n"
                    f"Arquivo salvo em:\n{resultado.caminho_saida}\n\n"
                    f"Relatório salvo em:\n{resultado.caminho_relatorio}"
                ),
            )

        except ErroAutomacao as erro:
            self.status_label.configure(text="Não foi possível gerar as fichas.", text_color=COR_ALERTA)
            messagebox.showerror("Erro", str(erro))
        except Exception as erro:
            self.status_label.configure(text="Erro inesperado no processamento.", text_color=COR_ERRO)
            messagebox.showerror("Erro", f"Ocorreu um problema: {erro}")

        finally:
            self.progress_bar.stop()
            self.progress_bar.set(0)
            self.btn_executar.configure(state="normal")
            self.btn_limpar.configure(state="normal")


if __name__ == "__main__":
    root = ctk.CTk()
    app_gui = AutomacaoFichas(root)
    root.mainloop()
