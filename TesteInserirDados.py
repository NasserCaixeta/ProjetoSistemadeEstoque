import tkinter as tk
from tkinter import ttk, messagebox
import mysql.connector

class TelaInsercaoItensCustom(tk.Tk):
    def __init__(self):
        super().__init__()

        # Configuração da conexão com o banco de dados MySQL
        self.conexao = mysql.connector.connect(
            host="localhost",
            user="root",
            password="senha",
            database="banco_dados_estoque"
        )
        self.cursor = self.conexao.cursor()

        # Configuração de estilo
        style = ttk.Style()
        style.configure('TLabel', background='black', foreground='white', font=('Arial', 12))
        style.configure('TEntry', background='black', foreground='white', font=('Arial', 12))
        style.configure('TButton', background='blue', foreground='white', font=('Arial', 12))

        self.criar_widgets()

    def criar_widgets(self):
        self.config(bg='black')  # Define o fundo da janela como preto

        # Criação dos rótulos e campos de entrada
        self.criar_label_entry("Nome:", row=0)
        self.criar_label_entry("Caixa:", row=1)
        self.criar_label_entry("Unidade:", row=2)
        self.criar_label_entry("Quilo:", row=3)
        self.criar_label_entry("Litro:", row=4)
        self.criar_label_entry("Grama:", row=5)
        self.criar_label_entry("Prioridade:", row=6)

        # Botão de inserção
        ttk.Button(self, text="Inserir Item", command=self.inserir_item, style='TButton').grid(row=7, column=0, columnspan=2, pady=10)

        # Configuração do fechamento da janela
        self.protocol("WM_DELETE_WINDOW", self.fechar_janela)

    def criar_label_entry(self, label_text, row):
        label = ttk.Label(self, text=label_text, style='TLabel')
        label.grid(row=row, column=0, sticky='w')

        entry = ttk.Entry(self, style='TEntry')
        entry.grid(row=row, column=1, padx=5, pady=5)

    def inserir_item(self):
        # Obtém os valores dos campos de entrada
        nome = self.entry_nome.get()
        caixa = self.entry_caixa.get()
        unidade = self.entry_unidade.get()
        quilo = self.entry_quilo.get()
        litro = self.entry_litro.get()
        grama = self.entry_grama.get()
        prioridade = self.entry_prioridade.get()

        # Verifica se a prioridade não é 0
        if int(prioridade) == 0:
            messagebox.showerror("Erro", "A prioridade do item não pode ser 0.")
            return

        # Inserir dados na tabela
        self.cursor.execute('''
            INSERT INTO itens (nome, caixa, unidade, quilo, litro, grama, prioridade)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
        ''', (nome, caixa, unidade, quilo, litro, grama, prioridade))

        # Commit para salvar as alterações no banco de dados
        self.conexao.commit()

        # Exibir mensagem de sucesso
        messagebox.showinfo("Sucesso", "Itens foram inseridos com sucesso.")

        # Limpar os campos de entrada após a inserção
        self.limpar_campos()

    def limpar_campos(self):
        self.entry_nome.delete(0, 'end')
        self.entry_caixa.delete(0, 'end')
        self.entry_unidade.delete(0, 'end')
        self.entry_quilo.delete(0, 'end')
        self.entry_litro.delete(0, 'end')
        self.entry_grama.delete(0, 'end')
        self.entry_prioridade.delete(0, 'end')

    def fechar_janela(self):
        # Fechar a conexão com o banco de dados antes de fechar a janela
        self.conexao.close()
        self.destroy()

# Criação da janela principal
app = TelaInsercaoItensCustom()
app.geometry("900x600+300+150")  # Defina o tamanho e posição da janela manualmente
app.mainloop()
