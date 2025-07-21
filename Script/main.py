# ===================== BIBLIOTECAS
import pandas as pd
import os
from docxtpl import DocxTemplate
from datetime import datetime
import tkinter as tk
from tkinter import Tk
from datetime import timedelta

# ===================== CAMINHOS
CAMINHO_PLANILHA = 'clientes.xlsx'
CAMINHO_TEMPLATE = 'template_proposta.docx'
PASTA_SAIDA = 'Propostas'

# ===================== SCRIPT

df = pd.read_excel(CAMINHO_PLANILHA)

# Agrupa por cliente para gerar um único arquivo para cada
clientes = df.groupby(['Cliente', 'CNPJ'])

# Cria um arquivo para cada cliente
for (cliente, cnpj), grupo in clientes:
    itens = []
    total_bruto = 0
    total_desconto = 0

    for _, row in grupo.iterrows():
        qtde = row['Qtde']
        preco_unit = row['Preço Unit.']
        desc_pct = row['Desconto (%)']

        total = qtde * preco_unit
        desconto = total * (desc_pct)
        total_liquido = total - desconto

        itens.append({
            'produto': row['Produto'],
            'qtde': int(qtde),
            'preco_unit': f'{preco_unit:.2f}'.replace('.', ','),
            'total': f'{total:.2f}'.replace('.', ',')
        })

        total_bruto += total
        total_desconto += desconto

    valor_final = total_bruto - total_desconto

    data = datetime.today()
    data_limite = data + timedelta(days=10)

    # Gera o contexto para o template
    contexto = {
        'cliente': cliente,
        'cnpj': cnpj,
        'itens': itens,
        'valor_total_bruto': f'{total_bruto:.2f}'.replace('.', ','),
        'desconto_total': f'{total_desconto:.2f}'.replace('.', ','),
        'valor_final': f'{valor_final:.2f}'.replace('.', ','),
        'data_hoje': f'{data.strftime('%d-%m-%Y')}',
        'data_limite': f'{data_limite.strftime('%d-%m-%Y')}',
    }

    # Gera o arquivo
    doc = DocxTemplate(CAMINHO_TEMPLATE)
    doc.render(contexto)

    pasta_cliente = os.path.join(PASTA_SAIDA, cliente)
    os.makedirs(pasta_cliente, exist_ok=True)

    nome_arquivo = f'Proposta_{cliente.replace(' ', '_')}_{data.strftime('%d-%m-%Y')}.docx'
    caminho_saida = os.path.join(pasta_cliente, nome_arquivo)
    doc.save(caminho_saida)

# Cria a janela de conclusão
root = Tk()
root.title('Concluído')
root.attributes('-topmost', True)
root.deiconify()
root.lift()
root.focus_force() 

def centralizar_janela():
    root.update_idletasks()
    w = root.winfo_width()
    h = root.winfo_height()
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width // 2) - (w // 2)
    y = (screen_height // 2) - (h // 2)
    root.geometry(f'+{x}+{y}')

root.after(10, centralizar_janela)

label = tk.Label(root, text='Propostas geradas!', font=('Trebuchet MS', '16'), justify='center')
label.pack(pady=(15, 10), padx=20)

button_frame = tk.Frame(root)
button_frame.pack(pady=(10, 20))

def open_folder():
    root.destroy()
    os.startfile(PASTA_SAIDA)

open_button = tk.Button(button_frame, text='Ver arquivos', font=('Trebuchet MS', '12'), bg="#f36303", command=open_folder, padx=10, pady=2, bd=0)
open_button.pack(side='left', padx=10)

close_button = tk.Button(button_frame, text='Fechar', font=('Trebuchet MS', '12'), bg="#7a7a7a", command=root.destroy, padx=10, pady=2, bd=0)
close_button.pack(side='left', padx=10)

root.wait_window()
