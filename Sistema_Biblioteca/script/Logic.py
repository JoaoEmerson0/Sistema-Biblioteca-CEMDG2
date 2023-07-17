import Windows as w
import Data_Base as DB
from PyQt5 import QtWidgets
from datetime import datetime, timedelta, date
import win32com.client as win32

data_de_hoje = datetime.today()
Hoje = data_de_hoje.strftime("%d/%m/%Y")


def ValidaLogin():
    try:
        usuario = w.login.lineEdit.text()

        senha = w.login.lineEdit_2.text()        
        DB.cursor.execute("SELECT senha FROM usuarios WHERE nome = (?)", (usuario,))        
        senhaBD = DB.cursor.fetchone()[0]

        if senha == senhaBD:
            DB.cursor.execute("SELECT funcao FROM usuarios WHERE nome = (?)", (usuario,))        
            func = DB.cursor.fetchone()[0]
            if func == "ALUNO":
                FuncCatalogo()
                w.login.close()
            elif func == "FUNCIONARIO":
                FuncionarioReserva()
                w.login.close()
        else:
            w.login.label_5.setText("USUARIO OU SENHA INVALIDOS!")        
    except Exception as e:
        print(f"Ocorreu um erro: {str(e)}")
        w.login.label_5.setText("ERRO AO EFETUAR LOGIN! ERRO: N° 394919")

def Catalogo():
    try:
        w.catalogo.show()
        DB.cursor.execute("SELECT * FROM livros")
        livros = DB.cursor.fetchall()
        w.catalogo.tableWidget.setRowCount(len(livros))
        w.catalogo.tableWidget.setColumnCount(5)
        
        for i in range(0, len(livros)):
            for j in range(0, 5):
                w.catalogo.tableWidget.setItem(i,j,QtWidgets.QTableWidgetItem(str(livros[i][j])))
    except Exception as e:
        print(f"Ocorreu um erro: {str(e)}")
        w.catalogo.logue.setText("ERRO AO PUXAR CATALOGO! ERRO: Nº 10225")

def FuncCatalogo():
    try:
        Catalogo()
        w.catalogo.Bminhasreservas.clicked.connect(FuncMinhasReservas)
        w.catalogo.Bvisualizar.clicked.connect(Visualizar)
        w.catalogo.Breservar.clicked.connect(Reservar)
        w.catalogo.Bsuporte.clicked.connect(funcSuporte)
        w.catalogo.Bsair.clicked.connect(w.catalogo.close)
    except Exception as e:
        print(f"Ocorreu um erro: {str(e)}")
        print(0)    

def Reservar():
    try:
        linha = w.catalogo.tableWidget.currentRow()
        identificador = w.catalogo.tableWidget.item(linha, 0).text()
        DB.cursor.execute("SELECT ID FROM livros WHERE ID = (?)", (identificador,))
        livro_id = str(DB.cursor.fetchone()[0])

        DB.cursor.execute("SELECT nomeLivro FROM livros WHERE ID = (?)", (livro_id,))
        nomeLivro = str(DB.cursor.fetchone()[0])
        estatus = "OK"
        usuario = w.login.lineEdit.text()

        data_retirada = date.today()
        data_entrega = data_retirada + timedelta(days=7)
        entrega_formatada = data_entrega.strftime("%d/%m/%Y")
        retirada_formatada = data_retirada.strftime("%d/%m/%Y")

        DB.cursor.execute("SELECT quantidade FROM livros WHERE ID = (?)", (livro_id))
        quantidade = int(DB.cursor.fetchone()[0])

        if quantidade != 0:
            baixa = quantidade - 1
            DB.cursor.execute("UPDATE livros SET quantidade = ? WHERE ID = ?", (baixa, livro_id))
            DB.cursor.execute("INSERT INTO reservas (nomeAluno, nomeLivro, estatus, dataRetirada, dataEntrega) VALUES (?, ?, ?, ?, ?)", (usuario, nomeLivro, estatus, retirada_formatada, entrega_formatada))
            DB.DB.commit()
            Catalogo()
            w.catalogo.logue.setText(f"Livro: {nomeLivro} Reservado Com Sucesso! Data de Entrega: {entrega_formatada}")
        else:
            w.catalogo.logue.setText(f"O livro: {nomeLivro} nao esta disponivel no momento!")
        
    except Exception as e:
        print(f"Ocorreu um erro: {str(e)}")
        w.catalogo.logue.setText("ERRO AO RESERVAR LIVRO! ERRO: Nº 64946")


def Visualizar():
    try:
        w.visualizar.show()
        linha = w.catalogo.tableWidget.currentRow()
        identificador = w.catalogo.tableWidget.item(linha, 0).text()

        DB.cursor.execute("SELECT ID FROM livros WHERE ID = (?)", (identificador,))
        id = str(DB.cursor.fetchone()[0])

        DB.cursor.execute("SELECT nomeLivro FROM livros WHERE ID = (?)", (identificador,))
        nomeLivro = str(DB.cursor.fetchone()[0])

        DB.cursor.execute("SELECT sinopce FROM livros WHERE ID = (?)", (identificador,))
        sinopce = str(DB.cursor.fetchone()[0])

        w.visualizar.ID.setText(id)
        w.visualizar.NOMELIVRO.setText(nomeLivro)
        w.visualizar.RESUMOLIVRO.setText(sinopce)
    except Exception as e:
        print(f"Ocorreu um erro: {str(e)}")
        w.visualizar.logue.setText("ERRO AO VIZUALIZAR LIVRO! ERRO: Nº 559123")

# ARRUMAR SISTEMA DE PENDENCIAS
def MinhasReservas():
    try:
        usuario = w.login.lineEdit.text()
        DB.cursor.execute("SELECT * FROM reservas WHERE nomeAluno = (?)", (usuario,))
        reservas = DB.cursor.fetchall()
        w.minhas_reservas.tableWidget.setRowCount(len(reservas))
        w.minhas_reservas.tableWidget.setColumnCount(6)

        for i in range(len(reservas)):
            for j in range(6):
                w.minhas_reservas.tableWidget.setItem(i, j, QtWidgets.QTableWidgetItem(str(reservas[i][j])))

    except Exception as e:
        print(f"Ocorreu um erro: {str(e)}")
        w.minhas_reservas.logue.setText("ERRO AO CARREGAR RESERVAS! ERRO: Nº 35514")

def FuncMinhasReservas():
    w.minhas_reservas.show()
    MinhasReservas()

def FuncionarioReserva():
    w.reserva.show()
    reservas()
    w.reserva.Bsuporte.clicked.connect(funcSuporte)
    w.reserva.Batualizar.clicked.connect(reservas)
    w.reserva.Bretirar.clicked.connect(ConfDevolucao)

def ConfDevolucao():
    try:
        linha = w.reserva.tableWidget.currentRow()
        identificador = w.reserva.tableWidget.item(linha, 0).text()

        DB.cursor.execute("SELECT ID FROM livros WHERE ID = ?", (identificador,))
        id = str(DB.cursor.fetchone()[0])

        DB.cursor.execute("SELECT estatus FROM livros WHERE ID = ?", (identificador,))
        status = str(DB.cursor.fetchone()[0])

        DB.cursor.execute("SELECT quantidade FROM livros WHERE ID = ?", (identificador,))
        quantidade = int(DB.cursor.fetchone()[0])


        if status != "ENTREGUE":
            devolucao = quantidade + 1

            DB.cursor.execute("UPDATE reservas SET estatus = 'ENTREGUE' WHERE ID = ?", (identificador))
            DB.cursor.execute("UPDATE livros SET quantidade = ? WHERE ID = ?", (devolucao, identificador))
            DB.DB.commit()
        else:
            w.reserva.logue.setText("Hummmm....Parece que este Livro ja foi entregue!")  

    except Exception as e:
        print(f"Ocorreu um erro: {str(e)}")
        w.reserva.logue.setText("ERRO AO CARREGAR RESERVAS! ERRO: Nº 35514")  

def reservas():
    try:
        DB.cursor.execute("SELECT * FROM reservas")
        reservas = DB.cursor.fetchall()
        w.reserva.tableWidget.setRowCount(len(reservas))
        w.reserva.tableWidget.setColumnCount(6)

        for i in range(len(reservas)):
            for j in range(6):
                w.reserva.tableWidget.setItem(i, j, QtWidgets.QTableWidgetItem(str(reservas[i][j])))

    except Exception as e:
        print(f"Ocorreu um erro: {str(e)}")
        w.reserva.logue.setText("ERRO AO CARREGAR RESERVAS! ERRO: Nº 35514")

def funcSuporte():
    try:
        usuario = w.login.lineEdit.text()
        w.suporte.label_7.setText(usuario)
        w.suporte.label_8.setText(Hoje)
        w.suporte.Bsuporte.clicked.connect(NovoTicket)
        w.suporte.Bgravar.clicked.connect(SalvarTicket)
        w.suporte.Bvoltar.clicked.connect(w.suporte.close)

        w.suporte.show()
    except Exception as e:
        print(f"Ocorreu um erro: {str(e)}")
        w.suporte.logue.setText("ERRO AO INICIALIZAR SUPORTE! ERRO: Nº 65168")

def NovoTicket():
    try:
        ticket = False
        NumeroTicket = int(w.suporte.label_6.text())
        DB.cursor.execute("SELECT * FROM tickets WHERE numeroTicket = ?", (NumeroTicket,))
        if DB.cursor.fetchone() is not None:
            ticket = True
        if ticket:
            novoTicket = NumeroTicket + 1
            w.suporte.label_6.setText(str(novoTicket))
            w.suporte.logue.setText("")
        else:
            novoTicket = NumeroTicket
            w.suporte.label_6.setText(str(novoTicket))
    except Exception as e:
        print(f"Ocorreu um erro: {str(e)}")
        w.suporte.logue.setText("ERRO AO CRIAR NOVO TICKET! ERRO: Nº 11656")

def SalvarTicket():
    try:
        numeroTicket = w.suporte.label_6.text()
        solicitante = w.suporte.label_7.text()
        dataCriacao = w.suporte.label_8.text()
        numeroErro = w.suporte.comboBox.currentText()
        descricao = w.suporte.textEdit.toPlainText()
        status = "AGUARDANDO"
        DB.cursor.execute("INSERT INTO Tickets (numeroTicket, solicitante, dataCriacao, NumeroErro, explicacao, status) VALUES (?, ?, ?, ?, ?, ?)", (numeroTicket, solicitante, dataCriacao, numeroErro, descricao, status,))
        DB.DB.commit()
        w.suporte.logue.setText(f"O Ticket {numeroErro}, foi gravado e enviado com Sucesso!")
        w.suporte.textEdit.setText("")
        outlook = win32.Dispatch("outlook.application")
        email = outlook.CreateItem(0)
        email.To = "suporte.sistema10@gmail.com"
        email.Subject = f"Ticket {numeroErro}"
        email.HTMLBody = f'''
        <p>Olá o sistema Possui um erro Segue informações sobre!</p>

        <p>Numero Ticket: {numeroTicket}</p>
        <p>Solicitante: {solicitante}</p>
        <p>Erro encontrado: {numeroErro}</p>
        <p>Descrição: {descricao}</p>
        '''
        email.Send()
        print("Ticket Enviado!")
    except Exception as e:
        print(f"Ocorreu um erro: {str(e)}")
        w.suporte.logue.setText("ERRO AO SALVAR TICKET! ERRO: Nº 165315")