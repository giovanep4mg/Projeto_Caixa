import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import easygui
import xlsxwriter

def obter_valor_numerico(mensagem):
    print("Iniciando obter_valor_numerico(mensagem)")  # Log de depuração
    while True:
        valor = easygui.enterbox(mensagem)
        if valor is None:
            return None
        try:
            return float(valor)
        except ValueError:
            messagebox.showerror("Erro", "Por favor, insira um valor numérico válido.")

def criar_arquivo_excel(df):
    print("Iniciando criar_arquivo_excel()")  # Log de depuração
    nome_arquivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivo Excel", "*.xlsx")])
    if nome_arquivo:
        try:
            writer = pd.ExcelWriter(nome_arquivo, engine='xlsxwriter')
            df.to_excel(writer, index=False)
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            formato_numero = workbook.add_format({'num_format': '#,##0.00'})
            worksheet.set_column('B:S', 15, formato_numero)
            writer.close()
            messagebox.showinfo("Sucesso", f"Arquivo '{nome_arquivo}' criado com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao criar arquivo: {e}")

def abrir_arquivo_excel():
    print("Iniciando abrir_arquivo_excel()")  # Log de depuração
    nome_arquivo = filedialog.askopenfilename(filetypes=[("Arquivo Excel", "*.xlsx")])
    if nome_arquivo:
        try:
            df = pd.read_excel(nome_arquivo)
            messagebox.showinfo("Sucesso", f"Arquivo '{nome_arquivo}' aberto com sucesso!")
            return df
        except FileNotFoundError:
            messagebox.showerror("Erro", f"O arquivo '{nome_arquivo}' não foi encontrado.")
            return None
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao abrir arquivo: {e}")
            return None

def editar_arquivo_excel():
    print("Iniciando editar_arquivo_excel()")  # Log de depuração
    nome_arquivo = filedialog.askopenfilename(filetypes=[("Arquivo Excel", "*.xlsx")])
    if nome_arquivo:
        try:
            df = pd.read_excel(nome_arquivo)
            dia = int(obter_valor_numerico("Digite o dia: "))
            dinhSalao = obter_valor_numerico("Dinheiro Salão")
            notas2 = obter_valor_numerico("Notas de 2")
            moedaSalao = obter_valor_numerico("Moeda Salão")
            dinhCasa = obter_valor_numerico("Dinheiro Casa")
            moedaCasa = obter_valor_numerico("Moeda Casa")
            gastoDia = obter_valor_numerico("Gasto Dia")
            sicoob = obter_valor_numerico("Sicoob")
            sumup = obter_valor_numerico("Sumup")
            nullbank = obter_valor_numerico("Nullbank")
            mercPago = obter_valor_numerico("MercPago")

            if None in (dia, dinhSalao, notas2, moedaSalao, dinhCasa, moedaCasa, gastoDia, sicoob, sumup, nullbank, mercPago):
                return

            casa = dinhCasa
            #
            novo_moedaCasa = moedaCasa
            #
            caixa = dinhSalao + notas2
            #
            totalbancos = sicoob + sumup + nullbank + mercPago

            moedaCasa_anterior = df['Moeda/Casa'].iloc[-1] if 'Moeda/Casa' in df.columns and not df.empty else obter_valor_numerico("Digite o valor anterior de Moeda Casa: ")
            if moedaCasa_anterior is None:
                return

            moedaCasa_atual = novo_moedaCasa + moedaCasa_anterior
            
            #
            totalsoma = casa + caixa + totalbancos + moedaSalao + moedaCasa_atual

            total_anterior = df['TotalSoma'].iloc[-1] if 'TotalSoma' in df.columns and not df.empty else obter_valor_numerico("Digite o valor do total anterior: ")
            if total_anterior is None:
                return

            novo_lucro = totalsoma - total_anterior
            totalDia = novo_lucro + gastoDia

            nova_linha = {
                'Dia': dia, 'Dinh/Salao': dinhSalao, 'Notas2': notas2, 'Moeda/Salao': moedaSalao,
                'Dinh/Casa': dinhCasa, 'Moeda/Casa': moedaCasa_atual, 'Gasto/Dia': gastoDia,
                'Sicoob': sicoob, 'Sumup': sumup, 'Nullbank': nullbank, 'MercPago': mercPago,
                'TotalBancos': totalbancos, 'Casa': casa, 'Caixa': caixa, 'TotalSoma': totalsoma,
                'Lucro': novo_lucro, 'TotalDia': totalDia, 'TotalAnterior': total_anterior
            }
            df_novo = pd.DataFrame([nova_linha])
            df = pd.concat([df, df_novo], ignore_index=True)

            writer = pd.ExcelWriter(nome_arquivo, engine='xlsxwriter')
            df.to_excel(writer, index=False)
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            formato_numero = workbook.add_format({'num_format': '#,##0.00'})
            worksheet.set_column('B:S', 15, formato_numero)
            writer.close()

            messagebox.showinfo("Sucesso", f"Arquivo '{nome_arquivo}' editado com sucesso!")
        except FileNotFoundError:
            messagebox.showerror("Erro", f"O arquivo '{nome_arquivo}' não foi encontrado.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao editar arquivo: {e}")

def criar_janela():
    print("Iniciando criar_janela()")  # Log de depuração
    janela = tk.Tk()
    janela.title("Gerenciador de Arquivos Excel")
    janela.geometry("600x400")

    dados = {
        'Dia': [], 'Dinh/Salao': [], 'Notas2': [], 'Moeda/Salao': [], 'Dinh/Casa': [],
        'Moeda/Casa': [], 'Gasto/Dia': [], 'TotalDia': [], 'Sicoob': [], 'Sumup': [],
        'Nullbank': [], 'MercPago': [], 'TotalBancos': [], 'Caixa': [], 'Casa': [],
        'TotalSoma': [], 'Lucro': [], 'TotalAnterior': []
    }
    df = pd.DataFrame(dados)

    botao_novo = tk.Button(janela, text="Criar Novo Excel", command=lambda: criar_arquivo_excel(df), width=20, height=2, font=("Arial", 12))
    botao_novo.pack(pady=10)

    botao_abrir = tk.Button(janela, text="Abrir Excel Existente", command=abrir_arquivo_excel, width=20, height=2, font=("Arial", 12))
    botao_abrir.pack(pady=10)

    botao_editar = tk.Button(janela, text="Editar Excel Existente", command=editar_arquivo_excel, width=20, height=2, font=("Arial", 12))
    botao_editar.pack(pady=10)

    janela.mainloop()

criar_janela()