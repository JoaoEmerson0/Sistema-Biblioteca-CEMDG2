import Windows as w
import Data_Base as DB
from PyQt5 import QtWidgets
from datetime import datetime, timedelta, date
import win32com.client as win32

data_de_hoje = datetime.today()
Hoje = data_de_hoje.strftime("%d/%m/%Y")

# LOGICA VALIDAÇÃO LOGIN 

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

# LOGICA CATALOGO

def FuncCatalogo():
    try:
        Catalogo()
        w.catalogo.Bminhasreservas.clicked.connect(FuncMinhasReservas)
        w.catalogo.Bvisualizar.clicked.connect(confVizualizar)
        w.catalogo.Breservar.clicked.connect(Reservar)
        w.catalogo.Bsuporte.clicked.connect(funcSuporte)
        w.catalogo.Bsair.clicked.connect(w.catalogo.close)
    except Exception as e:
        print(f"Ocorreu um erro: {str(e)}")

def Catalogo():
    try:
        w.catalogo.show()
        DB.cursor.execute("SELECT * FROM livros")
        livros = DB.cursor.fetchall()
        w.catalogo.tableWidget.setRowCount(len(livros))
        w.catalogo.tableWidget.setColumnCount(5)
        
        usuario = w.login.lineEdit.text()
        DB.cursor.execute("SELECT id, dataEntrega, estatus FROM reservas WHERE nomeAluno = (?)", (usuario,))
        result = DB.cursor.fetchall()

        for row in result:
            reserva_id = row[0]
            dataEntrega = row[1]
            estatus = row[2]

            novo_status = "EM DIA"

            if dataEntrega < Hoje and estatus != "ENTREGUE":
                novo_status = "PENDENTE"
            
            elif estatus == "ENTREGUE":
                novo_status = "ENTREGUE"
            
            elif dataEntrega >= Hoje and estatus != "ENTREGUE":
                novo_status = "OK"

            # Atualizar o campo "estatus" no banco de dados
            DB.cursor.execute("UPDATE reservas SET estatus = (?) WHERE id = (?)", (novo_status, reserva_id,))

        # Certifique-se de commitar as mudanças após o loop
        DB.DB.commit()


        for i in range(0, len(livros)):
            for j in range(0, 5):
                w.catalogo.tableWidget.setItem(i,j,QtWidgets.QTableWidgetItem(str(livros[i][j])))
    except Exception as e:
        print(f"Ocorreu um erro: {str(e)}")
        w.catalogo.logue.setText("ERRO AO PUXAR CATALOGO! ERRO: Nº 10225")

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
            if quantidade == 0:
                DB.cursor.execute("UPDATE livros SET estatus = 'INDISPONIVEL' WHERE ID = ?", (livro_id,))
        else:
            w.catalogo.logue.setText(f"O livro: {nomeLivro} nao esta disponivel no momento!")
        
    except Exception as e:
        print(f"Ocorreu um erro: {str(e)}")
        w.catalogo.logue.setText("ERRO AO RESERVAR LIVRO! ERRO: Nº 64946")

# LOGICA VIZUALIZAR

def confVizualizar():
    w.visualizar.show()
    Visualizar()
    w.visualizar.Bminhasreservas.clicked.connect(FuncMinhasReservas)
    w.visualizar.Bsuporte.clicked.connect(funcSuporte)
    w.visualizar.Bvoltar.clicked.connect(w.visualizar.close)
    w.visualizar.Bsair.clicked.connect(w.visualizar.close)

def Visualizar():
    try:
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

# LOGICA MINHAS RESERVAS

def FuncMinhasReservas():
    w.minhas_reservas.show()
    w.minhas_reservas.Bsuporte.clicked.connect(funcSuporte)
    w.minhas_reservas.Bvoltar.clicked.connect(w.minhas_reservas.close)
    w.minhas_reservas.Bsair.clicked.connect(w.minhas_reservas.close)
    MinhasReservas()

# ARRUMAR SISTEMA DE PENDENCIAS
def MinhasReservas():
    try:
        usuario = w.login.lineEdit.text()
        DB.cursor.execute("SELECT * FROM reservas WHERE nomeAluno = (?)", (usuario,))
        reservas = DB.cursor.fetchall()
        w.minhas_reservas.tableWidget.setRowCount(len(reservas))
        w.minhas_reservas.tableWidget.setColumnCount(6)
        pendencias()
        for i in range(len(reservas)):
            for j in range(6):
                w.minhas_reservas.tableWidget.setItem(i, j, QtWidgets.QTableWidgetItem(str(reservas[i][j])))

    except Exception as e:
        print(f"Ocorreu um erro: {str(e)}")
        w.minhas_reservas.logue.setText("ERRO AO CARREGAR RESERVAS! ERRO: Nº 35514")

# LOGICA CONTROLE DE RESERVAS 

def FuncionarioReserva():
    w.reserva.show()
    reservas()
    w.reserva.Bsuporte.clicked.connect(funcSuporte)
    w.reserva.Batualizar.clicked.connect(reservas)
    w.reserva.Bretirar.clicked.connect(ConfDevolucao)
    w.reserva.Bsair.clicked.connect(w.reserva.close)

def ConfDevolucao():
    try:
        linha = w.reserva.tableWidget.currentRow()
        identificador = w.reserva.tableWidget.item(linha, 0).text()

        DB.cursor.execute("SELECT estatus FROM reservas WHERE ID = ?", (identificador,))
        status = DB.cursor.fetchone()[0]

        DB.cursor.execute("SELECT nomeLivro FROM reservas WHERE ID = ?", (identificador,))
        livro = str(DB.cursor.fetchone()[0])

        DB.cursor.execute("SELECT quantidade FROM livros WHERE nomeLivro = ?", (livro,))        
        quantidade = int(DB.cursor.fetchone()[0])

        if status != "ENTREGUE":
            devol = int(quantidade + 1)
            DB.cursor.execute("UPDATE reservas SET estatus = 'ENTREGUE' WHERE ID = ?", (identificador,))
            DB.cursor.execute("UPDATE livros SET quantidade = ? WHERE nomeLivro = ?", (devol, livro,))
            DB.DB.commit()
            w.reserva.logue.setText(f"Reserva de ID: {identificador} Entregue com Sucesso!")
        else:
            w.reserva.logue.setText("Hummmmmm.....Parece que esta reserva ja foi entregue!")


    except Exception as e:
        print(f"Ocorreu um erro: {str(e)}")
        w.reserva.logue.setText("ERRO AO REALIZAR BAIXA EM RESERVA! ERRO: Nº 5445623")  


#ARRUMAR SISTEMA DE PENDENCIAS 

def pendencias():
    try: 
        usuario = w.login.lineEdit.text()
        DB.cursor.execute("SELECT id, dataEntrega, estatus FROM reservas WHERE nomeAluno = (?)", (usuario,))
        result = DB.cursor.fetchall()

        for row in result:
            reserva_id = row[0]
            dataEntrega = row[1]
            estatus = row[2]

            novo_status = "EM DIA"

            if dataEntrega < Hoje and estatus != "ENTREGUE":
                novo_status = "PENDENTE"
            
            elif estatus == "ENTREGUE":
                novo_status = "ENTREGUE"
            
            elif dataEntrega >= Hoje and estatus != "ENTREGUE":
                novo_status = "OK"

            # Atualizar o campo "estatus" no banco de dados
            DB.cursor.execute("UPDATE reservas SET estatus = (?) WHERE id = (?)", (novo_status, reserva_id,))

        # Certifique-se de commitar as mudanças após o loop
        DB.DB.commit()
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

# LOGICA SUPORTE

def funcSuporte():
    try:
        RecTicket()
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

def RecTicket():
    DB.cursor.execute("SELECT numeroTicket FROM Tickets ORDER BY numeroTicket DESC LIMIT 1")
    UltimoTiket = DB.cursor.fetchone()

    if UltimoTiket:
        UltimoTiket = UltimoTiket[0]
        
    else:
        UltimoTiket = None
    w.suporte.label_6.setText(UltimoTiket)

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
        # email.To = "suporte.sistema10@gmail.com"
        email.To = "je4740091@gmail.com"
        email.Subject = f"Ticket {numeroErro}"
        email.HTMLBody = f'''
        <p>Ola Parece que um de Nossos Usuarios Encontrou um ERRO !</p>
        <p>Numero Ticket: {numeroTicket}    Data Criação: {dataCriacao} </p>
        <p>Solicitante: {solicitante}</p>
        <p>Numero do Erro: {numeroErro}</p>
        <p>Descrição: {descricao}</p>
        '''

        email.Send()
        print("Ticket Enviado!")
    except Exception as e:
        print(f"Ocorreu um erro: {str(e)}")
        w.suporte.logue.setText("ERRO AO SALVAR TICKET! ERRO: Nº 165315")
        