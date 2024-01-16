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
        self.style = ttk.Style()
        self.style.configure('TLabel', padding=(5, 5), font=('Arial', 10))
        self.style.configure('TEntry', padding=(5, 5), font=('Arial', 10))
        self.style.configure('TButton', padding=(10, 5), font=('Arial', 10))


            # Rótulos e campos de entrada estilizados
        ttk.Label(self, text="Nome:", style="TLabel").grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        self.entry_nome = ttk.Entry(self, style="TEntry")
        self.entry_nome.grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(self, text="Caixa:", style="TLabel").grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        self.entry_caixa = ttk.Entry(self, style="TEntry")
        self.entry_caixa.grid(row=1, column=2, padx=5, pady=5)

        ttk.Label(self, text="Unidade:", style="TLabel").grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)
        self.entry_unidade = ttk.Entry(self, style="TEntry")
        self.entry_unidade.grid(row=2, column=2, padx=5, pady=5)

        ttk.Label(self, text="Quilo:", style="TLabel").grid(row=3, column=1, padx=5, pady=5, sticky=tk.W)
        self.entry_quilo = ttk.Entry(self, style="TEntry")
        self.entry_quilo.grid(row=3, column=2, padx=5, pady=5)

        ttk.Label(self, text="Litro:", style="TLabel").grid(row=4, column=1, padx=5, pady=5, sticky=tk.W)
        self.entry_litro = ttk.Entry(self, style="TEntry")
        self.entry_litro.grid(row=4, column=2, padx=5, pady=5)

        ttk.Label(self, text="Grama:", style="TLabel").grid(row=5, column=1, padx=5, pady=5, sticky=tk.W)
        self.entry_grama = ttk.Entry(self, style="TEntry")
        self.entry_grama.grid(row=5, column=2, padx=5, pady=5)

        ttk.Label(self, text="Prioridade:", style="TLabel").grid(row=6, column=1, padx=5, pady=5, sticky=tk.W)
        self.entry_prioridade = ttk.Entry(self, style="TEntry")
        self.entry_prioridade.grid(row=6, column=2, padx=5, pady=5)

        # Botão de inserção
        ttk.Button(self, text="Inserir Item", command=self.inserir_item, style="TButton").grid(row=7, column=2, padx=5, pady=5, sticky=tk.W)

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
            tk.messagebox.showerror("Erro", "A prioridade do item não pode ser 0.")
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
        self.entry_nome.delete(0, tk.END)
        self.entry_caixa.delete(0, tk.END)
        self.entry_unidade.delete(0, tk.END)
        self.entry_quilo.delete(0, tk.END)
        self.entry_litro.delete(0, tk.END)
        self.entry_grama.delete(0, tk.END)
        self.entry_prioridade.delete(0, tk.END)

    def fechar_janela(self):
        # Fechar a conexão com o banco de dados antes de fechar a janela
        self.conexao.fechar_conexao()
        self.master.destroy()

# Criação da janela principal
app = TelaInsercaoItensCustom()
app.geometry("900x600+300+150")  # Defina o tamanho e posição da janela manualmente
app.mainloop()
