import os
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox
from ttkbootstrap import Style
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Side

def apply_styles(ws):
    """Aplica alinhamento central, bordas e ajusta a largura das colunas."""
    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))
    
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = thin_border
            except:
                pass
        ws.column_dimensions[col_letter].width = max_length + 2

def process_excel(file_path, status_label):
    """Processa o arquivo Excel e gera novas planilhas divididas pela coluna 'UNIDADE'."""
    if not file_path:
        messagebox.showwarning("Aviso", "Nenhum arquivo selecionado!")
        return
    
    file_name = os.path.splitext(os.path.basename(file_path))[0]
    output_dir = os.path.join(os.path.dirname(file_path), file_name)
    os.makedirs(output_dir, exist_ok=True)
    
    xls = pd.ExcelFile(file_path)
    sheets = xls.sheet_names
    unidades = set()
    headers = {}
    
    for sheet in sheets:
        df = pd.read_excel(xls, sheet_name=sheet, header=None)
        headers[sheet] = df.iloc[0].tolist()
        df.columns = df.iloc[1]
        df = df[2:].reset_index(drop=True)
        if 'UNIDADE' in df.columns:
            unidades.update(df['UNIDADE'].dropna().unique())
    
    for unidade in unidades:
        output_file = os.path.join(output_dir, f"{unidade}.xlsx")
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for sheet in sheets:
                df = pd.read_excel(xls, sheet_name=sheet, header=1)
                if 'UNIDADE' in df.columns:
                    df_filtered = df[df['UNIDADE'] == unidade]
                    if not df_filtered.empty:
                        header_df = pd.DataFrame([df_filtered.columns.tolist()], columns=df_filtered.columns)
                        first_row_df = pd.DataFrame([headers[sheet]], columns=df_filtered.columns)
                        df_final = pd.concat([first_row_df, header_df, df_filtered], ignore_index=True)
                        df_final.to_excel(writer, sheet_name=sheet, index=False, header=False)
        
        wb = load_workbook(output_file)
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            apply_styles(ws)
        wb.save(output_file)
    
    status_label.config(text=f"Processamento concluído! Arquivos em: {output_dir}")
    messagebox.showinfo("Sucesso", f"Processamento concluído! Arquivos salvos em: {output_dir}")

def select_file(entry_widget):
    """Abre a janela para seleção de arquivo e exibe o caminho no campo de entrada."""
    file_path = filedialog.askopenfilename(title="Selecione o arquivo Excel", filetypes=[("Excel files", "*.xlsx")])
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, file_path)

def create_gui():
    """Cria a interface gráfica usando tkintermodernthemes."""
    root = tk.Tk()
    style = Style(theme='cosmo')
    root.title("Processador de Planilhas Excel")
    root.geometry("430x250")
    root.resizable(False, False)
    
    frame = tk.Frame(root, padx=20, pady=20)
    frame.pack(fill=tk.BOTH, expand=True)
    
    tk.Label(frame, text="Selecione um arquivo Excel:").grid(row=0, column=0, sticky=tk.W, pady=5)
    
    file_entry = tk.Entry(frame, width=50)
    file_entry.grid(row=1, column=0, padx=5, pady=5)
    
    select_button = tk.Button(frame, text="Selecionar", command=lambda: select_file(file_entry))
    select_button.grid(row=1, column=1, padx=5, pady=5)
    
    process_button = tk.Button(frame, text="Processar", command=lambda: process_excel(file_entry.get(), status_label))
    process_button.grid(row=2, column=0, columnspan=2, pady=15)
    
    status_label = tk.Label(frame, text="", fg="blue")
    status_label.grid(row=3, column=0, columnspan=2)
    
    root.mainloop()

if __name__ == "__main__":
    create_gui()