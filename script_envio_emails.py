import os
import sys
import pandas as pd
import win32com.client as win32
import tkinter as tk
from tkinter import messagebox, filedialog

# Variável global para armazenar o caminho do arquivo
caminho_arquivo_excel = ""

# Função para obter o caminho do arquivo 'clientes.xlsx'
def obter_caminho_arquivo_excel():
    global caminho_arquivo_excel
    
    if not caminho_arquivo_excel:
        messagebox.showerror("Erro", "Por favor, selecione o arquivo 'clientes.xlsx'.")
    

    return caminho_arquivo_excel

# Função para permitir que o usuário selecione o arquivo 'clientes.xlsx'
def selecionar_arquivo():
    global caminho_arquivo_excel
    
    caminho_arquivo_excel = filedialog.askopenfilename(
        title="Selecione o arquivo 'clientes.xlsx'",
        filetypes=[("Arquivos Excel", "*.xlsx")]
    )

    if caminho_arquivo_excel:
        messagebox.showinfo("Arquivo Selecionado", f"Arquivo selecionado com sucesso: {caminho_arquivo_excel}")
    else:
        messagebox.showerror("Erro", "Nenhum arquivo foi selecionado.")

# Função para carregar dados dos clientes do Excel
def carregar_dados_clientes():
    try:
        clientes_df = pd.read_excel(obter_caminho_arquivo_excel())
        return clientes_df
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao carregar o arquivo Excel: {e}")

# Instância do Outlook
outlook = win32.Dispatch('Outlook.Application')

# Função para enviar e-mails
def enviar_emails(assunto, corpo_email):
    try:
        clientes_df = carregar_dados_clientes()  # Carregar dados aqui
        
        # Loop pelos clientes
        for _, cliente in clientes_df.iterrows():
            nome = cliente['Nome']
            email = cliente['Email']

            # Criar e enviar o e-mail
            email_item = outlook.CreateItem(0)
            email_item.To = email
            email_item.Subject = assunto.format(nome=nome)
            email_item.HTMLBody = corpo_email.format(nome=nome)

            email_item.Send()
            
            print(f"Email enviado para {nome}, ({email})")

        # Mensagem de sucesso após o envio de todos os e-mails
        messagebox.showinfo(f"Sucesso", "Todos os e-mails foram enviados com sucesso!",)
        
    except Exception as e:
        # Exibir mensagem de erro
        messagebox.showerror("Erro", f"Ocorreu um erro ao enviar os e-mails: {e}")

# Função que pega os dados da interface e chama a função de envio
def enviar_email():
    assunto = entry_assunto.get()
    corpo_email = text_corpo_email.get("1.0", "end-1c")  # Pega todo o conteúdo da área de texto

    if not assunto or not corpo_email:
        messagebox.showwarning("Aviso", "Por favor, preencha todos os campos.")
    else:
        # Chama a função de envio de e-mails com os dados fornecidos
        enviar_emails(assunto, corpo_email)

# Criando a interface gráfica
root = tk.Tk()
root.title("ENVIOS DE E-MAILS EM MASSA")

# Adicionando descrição explicativa
label_descricao = tk.Label(root, text="Essa ferramenta permite enviar e-mails em massa para todos os seus contatos cadastrados no arquivo 'clientes.xlsx'.\
    O e-mail de envio será o configurado no Outlook do seu computador.", wraplength=400)
label_descricao.pack(pady=10)

# Tamanho da janela
root.geometry("600x500")

# Label e campo para o assunto
label_assunto = tk.Label(root, text="Assunto do E-mail:")
label_assunto.pack(pady=5)
entry_assunto = tk.Entry(root, width=65)
entry_assunto.pack(pady=5)

# Label e campo para o corpo do e-mail
label_corpo_email = tk.Label(root, text="Corpo do E-mail (use {nome} para personalizar):")
label_corpo_email.pack(pady=5)
text_corpo_email = tk.Text(root, height=10, width=50)
text_corpo_email.pack(pady=5)

# Botão para selecionar o arquivo 'clientes.xlsx'
botao_selecionar_arquivo = tk.Button(root, text="Selecionar Arquivo 'clientes.xlsx'", command=selecionar_arquivo, height=2, width=30)
botao_selecionar_arquivo.pack(pady=10)

# Botão para enviar o e-mail
botao_enviar = tk.Button(root, text="Enviar E-mail para lista de contatos", command=enviar_email, height=3, width=50)
botao_enviar.pack(pady=40)

# Iniciar a interface
root.mainloop()