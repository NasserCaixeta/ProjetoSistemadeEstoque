import mysql.connector
from tkinter import Tk, Label, Entry, Button, Frame, ttk, messagebox
from tkinter import *
from ttkthemes import ThemedTk

class BancoDeDadosGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Banco de Dados Estoque")

        # Configuração do banco de dados
        self.conexao = mysql.connector.connect(
            host="localhost",
            user="root",
            password="senha",
            database="banco_dados_estoque"
        )
        self.cursor = self.conexao.cursor()

        # Componentes da interface

        self.button_buscar = Button(root, text="Buscar", command=self.buscar_dados)
        self.root.geometry("1000x700")  
        self.root.resizable(width = False, height = False)
        self.grid_frame = Frame(root)
        self.grid_frame.grid(row=1, column=0, columnspan=3)



        self.tree = ttk.Treeview(self.grid_frame, columns=("Caixa", "Unidade", "Quilo", "Litro", "Grama", "Prioridade"), height = 15)
        for column in ("Caixa", "Unidade", "Quilo", "Litro", "Grama", "Prioridade"):
            self.tree.heading(column, text=column)
        self.tree.heading("#0", text="Nome")
        self.tree.heading("Caixa", text="Caixa")
        self.tree.heading("Unidade", text="Unidade")
        self.tree.heading("Quilo", text="Quilo")
        self.tree.heading("Litro", text="Litro")
        self.tree.heading("Grama", text="Grama")
        self.tree.heading("Prioridade", text="Prioridade")

        for column in self.tree["columns"]:
            self.tree.column(column, width=100, anchor='center')

        self.label_nome = Label(root, text="Nome:")
        self.entry_nome = Entry(root)


        self.label_caixa = Label(root, text="Caixa:")
        self.entry_caixa = Entry(root)

        self.label_unidade = Label(root, text="Unidade:")
        self.entry_unidade = Entry(root)

        self.label_quilo = Label(root, text="Quilo:")
        self.entry_quilo = Entry(root)

        self.label_litro = Label(root, text="Litro:")
        self.entry_litro = Entry(root)

        self.label_grama = Label(root, text="Grama:")
        self.entry_grama = Entry(root)

        self.label_prioridade = Label(root, text="Prioridade:")
        self.entry_prioridade = Entry(root)

        self.button_atualizar = Button(root, text="Atualizar", command=self.atualizar_dados)
        self.button_excluir = Button(root, text="Excluir", command=self.excluir_dados)

        # Posicionamento dos componentes
        
        self.label_nome.grid(row=2, column=0, sticky="e", padx=(5))
        self.entry_nome.grid(row=2, column=1, pady=5, padx=(0, 5),sticky="w")

        # Se necessário, ajuste as outras colunas da Treeview
        self.tree.grid(row=0, column=0, columnspan=3, sticky="nsew", padx=5, pady=5)
        self.tree.bind("<ButtonRelease-1>", self.selecionar_item)

        self.label_caixa.grid(row=3, column=0, sticky="e", padx=(0, 5), pady=5)
        self.entry_caixa.grid(row=3, column=1, pady=5, sticky= "w")

        self.label_unidade.grid(row=4, column=0, sticky="e", padx=(0, 5), pady=5)
        self.entry_unidade.grid(row=4, column=1, pady=5, sticky= "w")

        self.label_quilo.grid(row=5, column=0, sticky="e", padx=(0, 5), pady=5)
        self.entry_quilo.grid(row=5, column=1, pady=5, sticky= "w")

        self.label_litro.grid(row=6, column=0, sticky="e", padx=(0, 5), pady=5)
        self.entry_litro.grid(row=6, column=1, pady=5, sticky= "w")

        self.label_grama.grid(row=7, column=0, sticky="e", padx=(0, 5), pady=5)
        self.entry_grama.grid(row=7, column=1, pady=5, sticky= "w")

        self.label_prioridade.grid(row=8, column=0, sticky="e", padx=(0, 5), pady=5)
        self.entry_prioridade.grid(row=8, column=1, pady=5, sticky= "w")


        self.button_atualizar.grid(row=9, column=1, pady=5, sticky="w")
        self.button_excluir.grid(row=10, column=1, pady=5, sticky="w")


        # Configuração do grid
        root.columnconfigure(0, weight=0)
        root.rowconfigure(1, weight=0)

        # Carregar dados iniciais no grid
        self.carregar_dados_iniciais()

    def carregar_dados_iniciais(self):
        # Limpar o Treeview antes de carregar os dados
        for item in self.tree.get_children():
            self.tree.delete(item)

        query = "SELECT Nome, Caixa, Unidade, Quilo, Litro, Grama, Prioridade FROM itens"
        self.cursor.execute(query)
        resultados = self.cursor.fetchall()

        for resultado in resultados:
            self.tree.insert("", "end", text=resultado[0], values=resultado[1:])

    def buscar_dados(self):
        nome = self.entry_nome.get()
        query = f"SELECT Caixa, Unidade, Quilo, Litro, Grama, Prioridade FROM itens WHERE Nome = '{nome}'"
        self.cursor.execute(query)
        resultado = self.cursor.fetchone()

        if resultado:
            self.entry_caixa.delete(0, "end")
            self.entry_caixa.insert(0, resultado[0])

            self.entry_unidade.delete(0, "end")
            self.entry_unidade.insert(0, resultado[1])

            self.entry_quilo.delete(0, "end")
            self.entry_quilo.insert(0, resultado[2])

            self.entry_litro.delete(0, "end")
            self.entry_litro.insert(0, resultado[3])

            self.entry_grama.delete(0, "end")
            self.entry_grama.insert(0, resultado[4])

            self.entry_prioridade.delete(0, "end")
            self.entry_prioridade.insert(0, resultado[5])

    def selecionar_item(self, event):
        item_selecionado = self.tree.selection()
        if item_selecionado:
            nome = self.tree.item(item_selecionado, "text")
            self.entry_nome.delete(0, "end")
            self.entry_nome.insert(0, nome)
            self.buscar_dados()

    def atualizar_dados(self):
        nome = self.entry_nome.get()
        caixa = self.entry_caixa.get()
        unidade = self.entry_unidade.get()
        quilo = self.entry_quilo.get()
        litro = self.entry_litro.get()
        grama = self.entry_grama.get()
        prioridade = self.entry_prioridade.get()

        query = f"UPDATE itens SET Caixa='{caixa}', Unidade='{unidade}', Quilo='{quilo}', Litro='{litro}', Grama='{grama}', Prioridade='{prioridade}' WHERE Nome='{nome}'"
        self.cursor.execute(query)
        self.conexao.commit()

        # Atualizar a exibição no grid
        self.carregar_dados_iniciais()

    def excluir_dados(self):
        nome = self.entry_nome.get()

        if nome:
            query = f"DELETE FROM itens WHERE Nome='{nome}'"
            self.cursor.execute(query)
            self.conexao.commit()

            # Limpar o Treeview
            for item in self.tree.get_children():
                self.tree.delete(item)

            # Carregar dados atualizados no Treeview
            self.carregar_dados_iniciais()

            # Limpar os campos após a exclusão
            self.entry_nome.delete(0, "end")
            self.entry_caixa.delete(0, "end")
            self.entry_unidade.delete(0, "end")
            self.entry_quilo.delete(0, "end")
            self.entry_litro.delete(0, "end")
            self.entry_grama.delete(0, "end")
            self.entry_prioridade.delete(0, "end")



if __name__ == "__main__":
    root = ThemedTk(theme="plastik")
    app = BancoDeDadosGUI(root)
    root.mainloop()
