import customtkinter as ctk
from CTkMessagebox import CTkMessagebox
from tkinter import filedialog
import datetime
import pandas as pd
import os

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Conversor de Dados")
        self.geometry("300x200")

        self.file_path = ""

        self.create_widgets()

    # Creation of the graphical interface elements
    def create_widgets(self):
        self.btn_select = ctk.CTkButton(self, text="Selecionar Arquivo Excel", command=self.select_file)
        self.btn_select.pack(pady=20)

        self.lbl_file = ctk.CTkLabel(self, text="Nenhum arquivo selecionado")
        self.lbl_file.pack()

        self.btn_convert = ctk.CTkButton(self, text="Converter Arquivo", command=self.convert_file, state="disabled")
        self.btn_convert.pack(pady=15)

    # Function to select the Excel file
    def select_file(self):
        file_types = [("Excel files", "*xlsx *xls")]
        self.file_path = filedialog.askopenfilename(title="Selecione um arquivo Excel", filetypes=file_types)

        if self.file_path:
            self.lbl_file.configure(text=f"Arquivo Selecionado: {os.path.basename(self.file_path)}")
            self.btn_convert.configure(state="normal")

    # Main function to convert the file
    def convert_file(self):
        try:
            df = pd.read_excel(self.file_path, skiprows=2, dtype={'Matricula': str, 'CPF': str})
            df.columns = df.columns.str.replace(r'[\.\s]+', '', regex=True).str.strip()

            required_columns = [
                'Filial', 'Matricula', 'Nome', 'CPF', 'EmailPrinc',
                'DataAdmis', 'DescFuncao', 'CCMovto', 'DescCompl'
            ]
            
            if not all(col in df.columns for col in required_columns):
                missing = [col for col in required_columns if col not in df.columns]
                raise ValueError(f"Colunas faltantes: {', '.join(missing)}")
            
            df = df.astype({
                'EmailPrinc': str,
                'Nome': str,
                'DescFuncao': str,
                'CCMovto': str
            }).fillna('')

            
            df_output = pd.DataFrame()

            # Columns to the converted Excel
            df_output['Identificador'] = df['CPF'].str.replace(r'\D', '', regex=True)
            df_output['Email'] = df['EmailPrinc'].str.strip()
            df_output['Nome'] = df['Nome'].str.strip()
            df_output['Sobrenome'] = 'RE' + ' ' + df['Matricula'].astype(str)
            df_output['Estado Ativação'] = 3
            df_output['Identificador Razão Social'] = ''
            df_output['Usa Tolerancia Empresa'] = ''
            df_output['Empresa'] = ''
            df_output['Tolerancia (Minutos)'] = 0
            df_output['Data da contratacao'] = pd.to_datetime(df['DataAdmis']).dt.strftime('%d/%m/%Y')
            df_output['Data final da contratacao'] = ''
            df_output['Campo Personalizado 1'] = ''
            df_output['Campo Personalizado 2'] = ''
            df_output['Campo Personalizado 3'] = ''
            df_output['Posição'] = df['DescFuncao'].str.strip()
            df_output['PIS'] = ''
            df_output['direccion'] = ''
            df_output['latitud'] = ''
            df_output['longitud'] = ''
            df_output['telefono'] = ''
            df_output['Grupo'] = df['CCMovto'].str.strip()
            df_output['Perfil'] = 'Usuário'
            df_output['Marção Web'] = 'Sim'
            df_output['Marcação App'] = 'Sim'

            df_output['Data da contratacao'] = pd.to_datetime(
                df['DataAdmis'], errors='coerce', dayfirst=True
            ).dt.strftime('%d/%m/%Y')

            # Output of the formatted file
            default_name = f"ImportGeoVictoria_FORMATADO_{datetime.datetime.now().strftime('%H%M')}.xlsx"

            output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivo Excel", "*xlsx")], initialfile=default_name, title="Salvar Arquivo")

            if not output_path:
                return # If the user cancel

            df_output.to_excel(output_path, index=False, engine='openpyxl')

            CTkMessagebox(
                title="Sucesso!", 
                message=f"Arquivo salvo em:\n{output_path}",
                icon="check"
            )

        except Exception as e:
            CTkMessagebox(
                title="Erro!",
                message=f"Falha na conversão:\n{str(e)}",
                icon="cancel"
            )

if __name__ == "__main__":
    app = App()
    app.mainloop()