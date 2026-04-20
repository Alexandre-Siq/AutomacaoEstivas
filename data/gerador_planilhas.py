import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import xlwings as xw
import os

ctk.set_appearance_mode("System")  
ctk.set_default_color_theme("blue") 

class AutomacaoFichas:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de Planilhas SSHD")
        
        self.root.geometry("550x480")
        self.root.iconbitmap("icone_estivas.ico") # Mantendo o seu ícone
        
        # Variáveis dos caminhos e dos novos campos
        self.caminho_base_mae = ctk.StringVar()
        
        self.setup_ui()

    def setup_ui(self):
        main_frame = ctk.CTkFrame(self.root, corner_radius=15)
        main_frame.pack(pady=20, padx=20, fill="both", expand=True)

        titulo = ctk.CTkLabel(main_frame, text="Selecione a Planilha:", font=ctk.CTkFont(size=20, weight="bold"))
        titulo.pack(pady=(15, 10))

        # --- SEÇÃO 1: ARQUIVO ---
        file_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        file_frame.pack(fill="x", padx=20, pady=(0, 15))
        
        entrada_arquivo = ctk.CTkEntry(file_frame, textvariable=self.caminho_base_mae, placeholder_text="Selecione a planilha mãe...", width=320, height=35)
        entrada_arquivo.pack(side="left", padx=(0, 10))
        
        btn_procurar = ctk.CTkButton(file_frame, text="Procurar", command=self.selecionar_arquivo, width=100, height=35)
        btn_procurar.pack(side="right")

        # --- SEÇÃO 2: DADOS DO SOLICITANTE ---
        dados_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        dados_frame.pack(fill="x", padx=20)
        
        ctk.CTkLabel(dados_frame, text="Dados do Solicitante (Repetidos em todas as abas):", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", pady=(0, 5))
        
        # Campos de entrada
        self.entrada_nome = ctk.CTkEntry(dados_frame, placeholder_text="Nome do Solicitante (Ex: Alexandre Siqueira...)", height=30)
        self.entrada_nome.pack(fill="x", pady=3)
        
        self.entrada_sshd = ctk.CTkEntry(dados_frame, placeholder_text="SSHD (Ex: X0801681)", height=30)
        self.entrada_sshd.pack(fill="x", pady=3)
        
        self.entrada_cargo = ctk.CTkEntry(dados_frame, placeholder_text="Cargo (Ex: ANALISTA DE SUPORTE I)", height=30)
        self.entrada_cargo.pack(fill="x", pady=3)

        # --- BOTÃO EXECUTAR ---
        self.btn_executar = ctk.CTkButton(main_frame, text="GERAR FICHAS", command=self.executar_processo, 
                                          font=ctk.CTkFont(size=15, weight="bold"), height=45, fg_color="#28a745", hover_color="#218838")
        self.btn_executar.pack(pady=(20, 10))

        self.status_label = ctk.CTkLabel(main_frame, text="Aguardando início...", text_color="gray")
        self.status_label.pack()

    def selecionar_arquivo(self):
        arquivo = filedialog.askopenfilename(
            title="Selecionar Planilha Mãe",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
        )
        if arquivo:
            self.caminho_base_mae.set(arquivo)

    def executar_processo(self):
        caminho_mae = self.caminho_base_mae.get()
        
        if not caminho_mae:
            messagebox.showwarning("Atenção", "Por favor, selecione a planilha mãe primeiro.")
            return

        caminho_template = os.path.join(os.path.dirname(caminho_mae), "data", "TEMPLATE_NOVO.xlsx")
        
        if not os.path.exists(caminho_template):
            caminho_template = os.path.abspath("TEMPLATE_NOVO.xlsx")
            if not os.path.exists(caminho_template):
                messagebox.showerror("Erro", "Arquivo 'TEMPLATE_NOVO.xlsx' não encontrado na pasta do programa.")
                return

        # Captura os dados digitados na interface
        nome_solic = self.entrada_nome.get()
        sshd_solic = self.entrada_sshd.get()
        cargo_solic = self.entrada_cargo.get()

        self.status_label.configure(text="Processando... por favor, aguarde.", text_color="#17a2b8")
        self.root.update()

        app = xw.App(visible=False)
        wb = None

        try:
            df = pd.read_excel(caminho_mae)
            colunas_remover = ['Unidade', 'CBO', 'N°', 'Conselho', 'Telefone', 'Admissão', 'SSHD', 'DDD', 'Escala', 'TASY', 'MV', 'SENIOR']
            df = df.drop(columns=[c for c in colunas_remover if c in df.columns])

            wb = app.books.open(caminho_template)
            aba_template = wb.sheets['SSHD']

            mapeamento = {
                "Nome Colaborador": "B11", "Data Nascimento": "B12", "CPF": "B13",
                "Estado Civil": "B16", "Nome Completo da Mãe": "B17", "Nacionalidade": "B18",
                "Naturalidade": "B19", "E-MAIL": "B20", "Endereço": "B22",
                "Nº": "B23", "Complemento": "B24", "Bairro": "B25",
                "Cidade": "B26", "UF": "B27", "CEP": "B28",
                "Cargo": "B31", "Registro do Funcionário": "B32"
            }

            for index, row in df.iterrows():
                nome_colaborador = str(row.get('Nome Colaborador', f'Cadastro_{index}'))
                nome_aba = "".join(c for c in nome_colaborador if c not in r'/\*?[]:')[:31]
                
                nova_aba = aba_template.copy(after=wb.sheets[-1], name=nome_aba)
                
                for coluna_df, celula_excel in mapeamento.items():
                    if coluna_df in df.columns and pd.notna(row[coluna_df]):
                        nova_aba.range(celula_excel).value = row[coluna_df]
                
                nova_aba.range('B14').value = "Não informado"
                nova_aba.range('B15').value = "Não Informado"
                
                nova_aba.range('B30').value = "CLT"
                nova_aba.range('B34').value = "COMPLEXO HOSPITALAR DOS ESTIVADORES"
                nova_aba.range('B35').value = "SMS"
                nova_aba.range('B37').value = "COMPLEXO HOSPITALAR DOS ESTIVADORES"
                nova_aba.range('B38').value = "SMS"

                # === INSERINDO OS DADOS DO SOLICITANTE ===
                if nome_solic:
                    nova_aba.range('B5').value = nome_solic
                if sshd_solic:
                    nova_aba.range('B6').value = sshd_solic
                if cargo_solic:
                    nova_aba.range('B7').value = cargo_solic

            aba_template.delete()
            wb.sheets[0].activate()
            
            caminho_saida = os.path.join(os.path.dirname(caminho_mae), "SSHD_PREENCHIDO_FINAL.xlsx")
            wb.save(caminho_saida)
            
            self.status_label.configure(text="Concluído com sucesso!", text_color="#28a745")
            messagebox.showinfo("Sucesso", f"Automação concluída!\nArquivo salvo em:\n{caminho_saida}")

        except Exception as e:
            self.status_label.configure(text="Erro no processamento.", text_color="red")
            messagebox.showerror("Erro", f"Ocorreu um problema: {str(e)}")
        
        finally:
            if wb:
                wb.close()
            app.quit()

if __name__ == "__main__":
    root = ctk.CTk()
    app_gui = AutomacaoFichas(root)
    root.mainloop()