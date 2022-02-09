import sqlite3
from PyQt5 import uic, QtWidgets
import pandas as pd
import shutil
import os
import win32print
import win32api
import pdfkit
import datetime
from datetime import datetime


def salvar_dados():

    nom = tela.nome.text()
    raz = tela.razao.text()
    end = tela.endereco.text()
    cp = tela.cep.text()
    telefon = tela.telefone.text()
    cel = tela.celular.text()
    em = tela.email.text()
    prodv = tela.produto.text()
    val = tela.valor.text()
    qntpr = tela.quantidade.text()
    primcom4 = tela.datafechamento.text()
    descprod = tela.descricao.toPlainText()

    primcom2 = datetime.strptime(primcom4, '%d/%m/%Y %H:%M')
    primcom = primcom2.strftime('%d/%m/%Y %H:%M')


    # talvez alterar futuramente para aderir aqueles que pagaram por exemplo dinheiro e cartão ao mesmo tempo
    try:
        banco = sqlite3.connect('dbgilv02.db')
        cursor = banco.cursor()

        cursor.execute('CREATE TABLE IF NOT EXISTS CLIENTES (NOME VARCHAR(50), RAZAO VARCHAR(100), '
                       'ENDERECO VARCHAR(100),  CEP VARCHAR(30), TELEFONE VARCHAR(20), '
                       'CELULAR VARCHAR(30), EMAIL VARCHAR(50),  PRODUTO VARCHAR(100), '
                       'DESCRICAO_PROD TEXT, QTD_PROD INTEGER(50), VALOR INTEGER(50),  '
                       'COND_PAG VARCHAR(50), DATA_FECH DATETIME);')

        if tela.rbCartao.isChecked():
            if len(telefon) != 0:
                if len(cel) != 0:
                    cursor.execute("INSERT INTO CLIENTES(NOME, RAZAO, ENDERECO, CEP, TELEFONE, CELULAR, "
                                   "EMAIL, PRODUTO, DESCRICAO_PROD, QTD_PROD, VALOR, "
                                   "COND_PAG, DATA_FECH) VALUES ('"+nom.upper()+"', '"+raz.upper()+"', '"+end.upper()+"', '"+cp.upper()+"',  '"+telefon+"', '" + cel + "', '"+em.upper()+"', '"+prodv.upper()+"', '"+descprod.upper()+"', '"+qntpr+"', '"+val+"', 'CARTÃO', '"+primcom+"')")
                    banco.commit()
                if len(cel) == 0:
                    cursor.execute("INSERT INTO CLIENTES(NOME, RAZAO, ENDERECO, CEP, TELEFONE, CELULAR,"
                                   "EMAIL, PRODUTO, DESCRICAO_PROD, QTD_PROD, VALOR, "
                                   "COND_PAG, DATA_FECH) VALUES ('" + nom.upper() + "', '" + raz.upper() + "', '" + end.upper() + "', '" + cp.upper() + "',  '" + telefon + "', 'N/A', '" + em.upper() + "', '" + prodv.upper() + "', '" + descprod.upper() + "', '" + qntpr + "', '" + val + "', 'CARTÃO', '" + primcom + "')")
                    banco.commit()
            if len(telefon) == 0:
                if len(cel) != 0:
                    cursor.execute("INSERT INTO CLIENTES(NOME, RAZAO, ENDERECO, CEP, TELEFONE, CELULAR, "
                                   "EMAIL, PRODUTO, DESCRICAO_PROD, QTD_PROD, VALOR, "
                                   "COND_PAG, DATA_FECH) VALUES ('"+nom.upper()+"', '"+raz.upper()+"', '"+end.upper()+"', '"+cp.upper()+"', 'N/A', '" + cel + "', '"+em.upper()+"', '"+prodv.upper()+"', '"+descprod.upper()+"', '"+qntpr+"', '"+val+"', 'CARTÃO', '"+primcom+"')")
                    banco.commit()

                if len(cel) == 0:
                    cursor.execute("INSERT INTO CLIENTES(NOME, RAZAO, ENDERECO, CEP, TELEFONE, CELULAR,"
                                   "EMAIL, PRODUTO, DESCRICAO_PROD, QTD_PROD, VALOR, "
                                   "COND_PAG, DATA_FECH) VALUES ('" + nom.upper() + "', '" + raz.upper() + "', '" + end.upper() + "', '" + cp.upper() + "',  'N/A', 'N/A', '" + em.upper() + "', '" + prodv.upper() + "', '" + descprod.upper() + "', '" + qntpr + "', '" + val + "', 'CARTÃO', '" + primcom + "')")
                    banco.commit()

        if tela.rbDinheiro.isChecked():
            if len(telefon) != 0:
                if len(cel) != 0:
                    cursor.execute("INSERT INTO CLIENTES(NOME, RAZAO, ENDERECO, CEP, TELEFONE, CELULAR, "
                                   "EMAIL, PRODUTO, DESCRICAO_PROD, QTD_PROD, VALOR, "
                                   "COND_PAG, DATA_FECH) VALUES ('" + nom.upper() + "', '" + raz.upper() + "', '" + end.upper() + "', '" + cp.upper() + "',  '" + telefon + "', '" + cel + "', '" + em.upper() + "', '" + prodv.upper() + "', '" + descprod.upper() + "', '" + qntpr + "', '" + val + "', 'DINHEIRO', '" + primcom + "')")
                    banco.commit()
                if len(cel) == 0:
                    cursor.execute("INSERT INTO CLIENTES(NOME, RAZAO, ENDERECO, CEP, TELEFONE, CELULAR,"
                                   "EMAIL, PRODUTO, DESCRICAO_PROD, QTD_PROD, VALOR, "
                                   "COND_PAG, DATA_FECH) VALUES ('" + nom.upper() + "', '" + raz.upper() + "', '" + end.upper() + "', '" + cp.upper() + "',  '" + telefon + "', 'N/A', '" + em.upper() + "', '" + prodv.upper() + "', '" + descprod.upper() + "', '" + qntpr + "', '" + val + "', 'DINHEIRO', '" + primcom + "')")
                    banco.commit()

            if len(telefon) == 0:
                if len(cel) != 0:
                    cursor.execute("INSERT INTO CLIENTES(NOME, RAZAO, ENDERECO, CEP, TELEFONE, CELULAR, "
                                   "EMAIL, PRODUTO, DESCRICAO_PROD, QTD_PROD, VALOR, "
                                   "COND_PAG, DATA_FECH) VALUES ('" + nom.upper() + "', '" + raz.upper() + "', '" + end.upper() + "', '" + cp.upper() + "', 'N/A', '" + cel + "', '" + em.upper() + "', '" + prodv.upper() + "', '" + descprod.upper() + "', '" + qntpr + "', '" + val + "', 'DINHEIRO', '" + primcom + "')")
                    banco.commit()

                if len(cel) == 0:
                    cursor.execute("INSERT INTO CLIENTES(NOME, RAZAO, ENDERECO, CEP, TELEFONE, CELULAR,"
                                   "EMAIL, PRODUTO, DESCRICAO_PROD, QTD_PROD, VALOR, "
                                   "COND_PAG, DATA_FECH) VALUES ('" + nom.upper() + "', '" + raz.upper() + "', '" + end.upper() + "', '" + cp.upper() + "',  'N/A', 'N/A', '" + em.upper() + "', '" + prodv.upper() + "', '" + descprod.upper() + "', '" + qntpr + "', '" + val + "', 'DINHEIRO', '" + primcom + "')")
                    banco.commit()

        if tela.rbTransferencia.isChecked():
            if len(telefon) != 0:
                if len(cel) != 0:
                    cursor.execute("INSERT INTO CLIENTES(NOME, RAZAO, ENDERECO, CEP, TELEFONE, CELULAR, "
                                   "EMAIL, PRODUTO, DESCRICAO_PROD, QTD_PROD, VALOR, "
                                   "COND_PAG, DATA_FECH) VALUES ('" + nom.upper() + "', '" + raz.upper() + "', '" + end.upper() + "', '" + cp.upper() + "',  '" + telefon + "', '" + cel + "', '" + em.upper() + "', '" + prodv.upper() + "', '" + descprod.upper() + "', '" + qntpr + "', '" + val + "', 'TRANSFERENCIA', '" + primcom + "')")
                    banco.commit()
                if len(cel) == 0:
                    cursor.execute("INSERT INTO CLIENTES(NOME, RAZAO, ENDERECO, CEP, TELEFONE, CELULAR,"
                                   "EMAIL, PRODUTO, DESCRICAO_PROD, QTD_PROD, VALOR, "
                                   "COND_PAG, DATA_FECH) VALUES ('" + nom.upper() + "', '" + raz.upper() + "', '" + end.upper() + "', '" + cp.upper() + "',  '" + telefon + "', 'N/A', '" + em.upper() + "', '" + prodv.upper() + "', '" + descprod.upper() + "', '" + qntpr + "', '" + val + "', 'TRANSFERENCIA', '" + primcom + "')")
                    banco.commit()

            if len(telefon) == 0:
                if len(cel) != 0:
                    cursor.execute("INSERT INTO CLIENTES(NOME, RAZAO, ENDERECO, CEP, TELEFONE, CELULAR, "
                                   "EMAIL, PRODUTO, DESCRICAO_PROD, QTD_PROD, VALOR, "
                                   "COND_PAG, DATA_FECH) VALUES ('" + nom.upper() + "', '" + raz.upper() + "', '" + end.upper() + "', '" + cp.upper() + "', 'N/A', '" + cel + "', '" + em.upper() + "', '" + prodv.upper() + "', '" + descprod.upper() + "', '" + qntpr + "', '" + val + "', 'TRANSFERENCIA', '" + primcom + "')")
                    banco.commit()

                if len(cel) == 0:
                    cursor.execute("INSERT INTO CLIENTES(NOME, RAZAO, ENDERECO, CEP, TELEFONE, CELULAR,"
                                   "EMAIL, PRODUTO, DESCRICAO_PROD, QTD_PROD, VALOR, "
                                   "COND_PAG, DATA_FECH) VALUES ('" + nom.upper() + "', '" + raz.upper() + "', '" + end.upper() + "', '" + cp.upper() + "',  'N/A', 'N/A', '" + em.upper() + "', '" + prodv.upper() + "', '" + descprod.upper() + "', '" + qntpr + "', '" + val + "', 'TRANSFERENCIA', '" + primcom + "')")
                    banco.commit()
        if tela.rbCheque.isChecked():
            if len(telefon) != 0:
                if len(cel) != 0:
                    cursor.execute("INSERT INTO CLIENTES(NOME, RAZAO, ENDERECO, CEP, TELEFONE, CELULAR, "
                                   "EMAIL, PRODUTO, DESCRICAO_PROD, QTD_PROD, VALOR, "
                                   "COND_PAG, DATA_FECH) VALUES ('" + nom.upper() + "', '" + raz.upper() + "', '" + end.upper() + "', '" + cp.upper() + "',  '" + telefon + "', '" + cel + "', '" + em.upper() + "', '" + prodv.upper() + "', '" + descprod.upper() + "', '" + qntpr + "', '" + val + "', 'CHEQUE', '" + primcom + "')")
                    banco.commit()
                if len(cel) == 0:
                    cursor.execute("INSERT INTO CLIENTES(NOME, RAZAO, ENDERECO, CEP, TELEFONE, CELULAR,"
                                   "EMAIL, PRODUTO, DESCRICAO_PROD, QTD_PROD, VALOR, "
                                   "COND_PAG, DATA_FECH) VALUES ('" + nom.upper() + "', '" + raz.upper() + "', '" + end.upper() + "', '" + cp.upper() + "',  '" + telefon + "', 'N/A', '" + em.upper() + "', '" + prodv.upper() + "', '" + descprod.upper() + "', '" + qntpr + "', '" + val + "', 'CHEQUE', '" + primcom + "')")
                    banco.commit()

            if len(telefon) == 0:
                if len(cel) != 0:
                    cursor.execute("INSERT INTO CLIENTES(NOME, RAZAO, ENDERECO, CEP, TELEFONE, CELULAR, "
                                   "EMAIL, PRODUTO, DESCRICAO_PROD, QTD_PROD, VALOR, "
                                   "COND_PAG, DATA_FECH) VALUES ('" + nom.upper() + "', '" + raz.upper() + "', '" + end.upper() + "', '" + cp.upper() + "', 'N/A', '" + cel + "', '" + em.upper() + "', '" + prodv.upper() + "', '" + descprod.upper() + "', '" + qntpr + "', '" + val + "', 'CHEQUE', '" + primcom + "')")
                    banco.commit()

                if len(cel) == 0:
                    cursor.execute("INSERT INTO CLIENTES(NOME, RAZAO, ENDERECO, CEP, TELEFONE, CELULAR,"
                                   "EMAIL, PRODUTO, DESCRICAO_PROD, QTD_PROD, VALOR, "
                                   "COND_PAG, DATA_FECH) VALUES ('" + nom.upper() + "', '" + raz.upper() + "', '" + end.upper() + "', '" + cp.upper() + "',  'N/A', 'N/A', '" + em.upper() + "', '" + prodv.upper() + "', '" + descprod.upper() + "', '" + qntpr + "', '" + val + "', 'CHEQUE', '" + primcom + "')")
                    banco.commit()
        if tela.rbna.isChecked():
            if len(telefon) != 0:
                if len(cel) != 0:
                    cursor.execute("INSERT INTO CLIENTES(NOME, RAZAO, ENDERECO, CEP, TELEFONE, CELULAR, "
                                   "EMAIL, PRODUTO, DESCRICAO_PROD, QTD_PROD, VALOR, "
                                   "COND_PAG, DATA_FECH) VALUES ('" + nom.upper() + "', '" + raz.upper() + "', '" + end.upper() + "', '" + cp.upper() + "',  '" + telefon + "', '" + cel + "', '" + em.upper() + "', '" + prodv.upper() + "', '" + descprod.upper() + "', '" + qntpr + "', '" + val + "', 'N/A', '" + primcom + "')")
                    banco.commit()
                if len(cel) == 0:
                    cursor.execute("INSERT INTO CLIENTES(NOME, RAZAO, ENDERECO, CEP, TELEFONE, CELULAR,"
                                   "EMAIL, PRODUTO, DESCRICAO_PROD, QTD_PROD, VALOR, "
                                   "COND_PAG, DATA_FECH) VALUES ('" + nom.upper() + "', '" + raz.upper() + "', '" + end.upper() + "', '" + cp.upper() + "',  '" + telefon + "', 'N/A', '" + em.upper() + "', '" + prodv.upper() + "', '" + descprod.upper() + "', '" + qntpr + "', '" + val + "', 'N/A', '" + primcom + "')")
                    banco.commit()

            if len(telefon) == 0:
                if len(cel) != 0:
                    cursor.execute("INSERT INTO CLIENTES(NOME, RAZAO, ENDERECO, CEP, TELEFONE, CELULAR, "
                                   "EMAIL, PRODUTO, DESCRICAO_PROD, QTD_PROD, VALOR, "
                                   "COND_PAG, DATA_FECH) VALUES ('" + nom.upper() + "', '" + raz.upper() + "', '" + end.upper() + "', '" + cp.upper() + "', 'N/A', '" + cel + "', '" + em.upper() + "', '" + prodv.upper() + "', '" + descprod.upper() + "', '" + qntpr + "', '" + val + "', 'N/A', '" + primcom + "')")
                    banco.commit()

                if len(cel) == 0:
                    cursor.execute("INSERT INTO CLIENTES(NOME, RAZAO, ENDERECO, CEP, TELEFONE, CELULAR,"
                                   "EMAIL, PRODUTO, DESCRICAO_PROD, QTD_PROD, VALOR, "
                                   "COND_PAG, DATA_FECH) VALUES ('" + nom.upper() + "', '" + raz.upper() + "', '" + end.upper() + "', '" + cp.upper() + "',  'N/A', 'N/A', '" + em.upper() + "', '" + prodv.upper() + "', '" + descprod.upper() + "', '" + qntpr + "', '" + val + "', 'N/A', '" + primcom + "')")
                    banco.commit()
        banco.close()

    except sqlite3.Error as error:
        print('aqui', error)

    tela.nome.setText("")
    tela.razao.setText("")
    tela.endereco.setText("")
    tela.cep.setText("")
    tela.telefone.setText("")
    tela.celular.setText("")
    tela.email.setText("")
    tela.produto.setText("")
    tela.valor.setText("")
    tela.quantidade.setText("")
    tela.descricao.setText("")


def listar_dados():
    banco = sqlite3.connect('dbgilv02.db')
    cursor = banco.cursor()

    try:
        sql = 'SELECT * FROM CLIENTES'
        cursor.execute(sql)
        dados_lidos = cursor.fetchall()
        tela.tableWidget.setRowCount(len(dados_lidos))
        tela.tableWidget.setColumnCount(13)

        for i in range(0, len(dados_lidos)):
            for j in range(13):
                tela.tableWidget.setItem(i, j, QtWidgets.QTableWidgetItem(str(dados_lidos[i][j])))
        cursor.close()
        banco.commit()
        banco.close()
    except sqlite3.Error as error:
        print("Erro ao inserir os dados: 2", error)


def listar_dados_2():

    banco = sqlite3.connect('dbgilv02.db')
    cursor = banco.cursor()
    nom5 = tela.txtNome.text()
    prodv5 = tela.txtProduto.text()
    data2 = tela.dtDate.text()
    data3 = datetime.strptime(data2, '%d/%m/%Y %H:%M') # '%d/%m/%Y %H:%M'
    data = data3.strftime('%d/%m/%Y %H:%M')
    print(data)
    # data = str(data4)

    if tela.rbData.isChecked():
        if len(nom5) != 0:  # quer dizer que quer consultar pelo nome:
            if len(prodv5) != 0:
                sql = "SELECT * FROM CLIENTES WHERE NOME LIKE '"+nom5.upper()+"%' AND PRODUTO LIKE '"+prodv5.upper()+"%' AND DATA_FECH >= '"+data+"'"
                cursor.execute(sql)
                dados_lidos = cursor.fetchall()
                tela.tableWidget_2.setRowCount(len(dados_lidos))
                tela.tableWidget_2.setColumnCount(13)
                print(1)
                for i in range(0, len(dados_lidos)):
                    for j in range(13):
                        tela.tableWidget_2.setItem(i, j, QtWidgets.QTableWidgetItem(str(dados_lidos[i][j])))

            if len(prodv5) == 0:
                sql = "SELECT * FROM CLIENTES WHERE NOME LIKE '"+nom5.upper()+"%' AND DATA_FECH >= '"+data+"'"
                cursor.execute(sql)
                dados_lidos = cursor.fetchall()
                tela.tableWidget_2.setRowCount(len(dados_lidos))
                tela.tableWidget_2.setColumnCount(13)
                print(2)
                for i in range(0, len(dados_lidos)):
                    for j in range(13):
                        tela.tableWidget_2.setItem(i, j, QtWidgets.QTableWidgetItem(str(dados_lidos[i][j])))

        if len(prodv5) != 0:
            if len(nom5) == 0:
                sql = "SELECT * FROM CLIENTES WHERE PRODUTO LIKE '"+prodv5.upper()+"%' AND DATA_FECH = '"+data+"'"
                cursor.execute(sql)
                dados_lidos = cursor.fetchall()
                tela.tableWidget_2.setRowCount(len(dados_lidos))
                tela.tableWidget_2.setColumnCount(13)
                print(3)
                for i in range(0, len(dados_lidos)):
                    for j in range(13):
                        tela.tableWidget_2.setItem(i, j, QtWidgets.QTableWidgetItem(str(dados_lidos[i][j])))

        if len(nom5) == 0:
            if len(prodv5) == 0:
                sql = "SELECT * FROM CLIENTES WHERE DATA_FECH >= '" + data + "'"
                cursor.execute(sql)
                dados_lidos = cursor.fetchall()
                tela.tableWidget_2.setRowCount(len(dados_lidos))
                tela.tableWidget_2.setColumnCount(13)
                print('data')
                for i in range(0, len(dados_lidos)):
                    for j in range(13):
                        tela.tableWidget_2.setItem(i, j, QtWidgets.QTableWidgetItem(str(dados_lidos[i][j])))

    if tela.rbData.isChecked() is not True:
        if len(nom5) != 0:  # quer dizer que quer consultar pelo nome:
            if len(prodv5) != 0:
                sql = "SELECT * FROM CLIENTES WHERE NOME LIKE '" + nom5.upper() + "%' AND PRODUTO LIKE '" + prodv5.upper() + "%'"
                cursor.execute(sql)
                dados_lidos = cursor.fetchall()
                tela.tableWidget_2.setRowCount(len(dados_lidos))
                tela.tableWidget_2.setColumnCount(13)
                print(4)
                for i in range(0, len(dados_lidos)):
                    for j in range(13):
                        tela.tableWidget_2.setItem(i, j, QtWidgets.QTableWidgetItem(str(dados_lidos[i][j])))

            if len(prodv5) == 0:
                sql = "SELECT * FROM CLIENTES WHERE NOME LIKE '" + nom5.upper() + "%'"
                cursor.execute(sql)
                dados_lidos = cursor.fetchall()
                tela.tableWidget_2.setRowCount(len(dados_lidos))
                tela.tableWidget_2.setColumnCount(13)
                print(5)
                for i in range(0, len(dados_lidos)):
                    for j in range(13):
                        tela.tableWidget_2.setItem(i, j, QtWidgets.QTableWidgetItem(str(dados_lidos[i][j])))

        if len(prodv5) != 0:
            if len(nom5) == 0:
                sql = "SELECT * FROM CLIENTES WHERE PRODUTO LIKE '" + prodv5.upper() + "%'"
                cursor.execute(sql)
                dados_lidos = cursor.fetchall()
                tela.tableWidget_2.setRowCount(len(dados_lidos))
                tela.tableWidget_2.setColumnCount(13)
                print(6)
                for i in range(0, len(dados_lidos)):
                    for j in range(13):
                        tela.tableWidget_2.setItem(i, j, QtWidgets.QTableWidgetItem(str(dados_lidos[i][j])))

    cursor.close()
    banco.commit()
    banco.close()


def apagar_dados():
    banco = sqlite3.connect('dbgilv02.db')
    cursor = banco.cursor()
    produto1 = tela.produtoApagar.text()
    nome1 = tela.nomeApagar.text()

    try:
        if nome1 is not None:
            banco = sqlite3.connect('dbgilv02.db')
            cursor = banco.cursor()
            cursor.execute("DELETE FROM CLIENTES WHERE PRODUTO = '"+produto1.upper()+"' AND NOME = '"+nome1.upper()+"'")
        banco.commit()
        banco.close()

    except sqlite3.Error as error:
        print("Erro ao inserir os dados: 3", error)

    tela.produtoApagar.setText("")
    tela.nomeApagar.setText("")


def imprime_HTML():
    banco = sqlite3.connect('dbgilv02.db')
    cursor = banco.cursor()

    try:
        frame = pd.read_sql_query('SELECT * FROM CLIENTES', banco)
        html = frame.to_html()
        text_file = open("datagilv08.html", "w")
        text_file.write(html)
        text_file.close()
        src2 = os.getcwd() + r"\datagilv08.html"
        dest2 = os.getcwd() + r"\templates\files\datagilv08.html"
        shutil.move(src2, dest2)
        banco.commit()
        banco.close()

    except sqlite3.Error as error:
        print("Erro ao inserir os dados: 4", error)


def imprimir_dados():
    banco = sqlite3.connect('dbgilv02.db')
    cursor = banco.cursor()

    try:
        frame = pd.read_sql_query('SELECT * FROM CLIENTES', banco)
        html = frame.to_html()
        text_file = open("datagilv08.html", "w")
        text_file.write(html)
        text_file.close()
        config = pdfkit.configuration(wkhtmltopdf=r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe")
        pdfkit.from_file('datagilv08.html', 'datagilv08.pdf', configuration=config)
        source = os.getcwd()+r"\datagilv08.pdf"
        destination = os.getcwd()+r"\templates\datagilv08.pdf"
        shutil.move(source, destination)
        src2 = os.getcwd()+r"\datagilv08.html"
        dest2 = os.getcwd()+r"\templates\files\datagilv08.html"
        shutil.move(src2, dest2)
        banco.commit()
        banco.close()

        lista_impressoras = win32print.EnumPrinters(2)
        impressora = lista_impressoras[4]
        win32print.SetDefaultPrinter(impressora[2])  # 'L4160 Series(Rede)' -> Indice 2 que é o nome da impressora

        caminho = "templates"
        lista_arquivos = os.listdir(caminho)
        for arquivo in lista_arquivos:
            win32api.ShellExecute(0, "print", arquivo, None, caminho, 0)
        print(lista_arquivos)

    except sqlite3.Error as error:
        print("Erro ao inserir os dados: 4", error)


def imprimir_dados_2():
    banco = sqlite3.connect('dbgilv02.db')
    cursor = banco.cursor()
    nomeim = tela.nomeImprimir.text()
    pagamentoim = tela.pagamentoImprimir.text()
    dataim = tela.datade.text()
    dataim2 = tela.datate.text()
    produtoim = tela.produtoImprimir.text()
    maiorim = tela.maiorImprimir.text()
    valoruprim = tela.valorSupImprimir.text()
    valorinim = tela.valorInfImprimir.text()
    menorim = tela.menorImprimir.text()
    try:
        if nomeim is not None:
            frame = pd.read_sql_query("SELECT * FROM CLIENTES WHERE NOME = '" +nomeim.upper()+"'", banco)
        if pagamentoim is not None:
            frame = pd.read_sql_query("SELECT * FROM CLIENTES WHERE COND_PAG LIKE '" +pagamentoim.upper()+"%'", banco)
        if produtoim is not None:
            frame = pd.read_sql_query("SELECT * FROM CLIENTES WHERE PRODUTO LIKE '"+produtoim.upper()+"%'", banco)
        if menorim is not None:
            if maiorim is None:
                frame = pd.read_sql_query(
                    "SELECT * FROM CLIENTES WHERE QTD_PROD <= '"+menorim.upper()+"'",
                    banco)
            else:
                frame = pd.read_sql_query(
                    "SELECT * FROM CLIENTES WHERE QTD_PROD >='" + maiorim.upper() + "'",
                    banco)
        if maiorim is not None:
            if menorim is None:
                frame = pd.read_sql_query(
                    "SELECT * FROM CLIENTES WHERE QTD_PROD >='"+maiorim.upper()+"'",
                    banco)
            else:
                frame = pd.read_sql_query(
                    "SELECT * FROM CLIENTES WHERE QTD_PROD <= '" + menorim.upper() + "'",
                    banco)
        if maiorim and menorim is None:
            frame = pd.read_sql_query(
                "SELECT * FROM CLIENTES WHERE QTD_PROD BETWEEN '" + menorim.upper() + "' AND '" + maiorim.upper() + "'",
                banco)
        if valoruprim is not None:
            if valorinim is None:
                frame = pd.read_sql_query(
                    "SELECT * FROM CLIENTES WHERE VALOR >='" + valoruprim + "'",
                    banco)
        if valorinim is not None:
            if valoruprim is None:
                frame = pd.read_sql_query(
                    "SELECT * FROM CLIENTES WHERE VALOR <= '" + valorinim + "'",
                    banco)
        if valoruprim and valorinim is not None:
            frame = pd.read_sql_query(
                "SELECT * FROM CLIENTES WHERE VALOR BETWEEN '" + valorinim + "' AND '" + valoruprim + "'",
                banco)
        if dataim is not None:
            if dataim2 is None:
                frame = pd.read_sql_query(
                    "SELECT * FROM CLIENTES WHERE DATA_FECH <= '" + dataim + "'",
                    banco)
        if dataim2 is not None:
            if dataim is None:
                frame = pd.read_sql_query(
                    "SELECT * FROM CLIENTES WHERE DATA_FECH >='" + dataim2 + "'",
                    banco)
        if dataim and dataim2 is not None:
            frame = pd.read_sql_query("SELECT * FROM CLIENTES WHERE DATA_FECH BETWEEN '" +dataim+"' AND '"+dataim2+"'", banco)
        banco.commit()
        tela.nomeImprimir.setText("")
        tela.pagamentoImprimir.setText("")
        tela.produtoImprimir.setText("")
        tela.maiorImprimir.setText("")
        tela.valorSupImprimir.setText("")
        tela.valorInfImprimir.setText("")
        tela.menorImprimir.setText("")

        html = frame.to_html()
        text_file = open("datagilv02.html", "w")
        text_file.write(html)
        text_file.close()
        banco.close()
    # ZERAR OS CAMPOS
    # fazer if
    except sqlite3.Error as error:
        print("Erro ao inserir os dados: 5", error)


def limpar_dados():
    tela.nomeImprimir.setText("")
    tela.pagamentoImprimir.setText("")
    tela.produtoImprimir.setText("")
    tela.valorSupImprimir.setText("")
    tela.valorInfImprimir.setText("")


def imprimir_nome():
    banco = sqlite3.connect('dbgilv02.db')
    cursor = banco.cursor()
    nomeim = tela.nomeImprimir.text()

    try:
        if nomeim is not None:
            frame = pd.read_sql_query("SELECT * FROM CLIENTES WHERE NOME = '" + nomeim.upper() + "'", banco)
        tela.nomeImprimir.setText("")

        html = frame.to_html()
        text_file = open("datagilv08EDIT.html", "w")
        text_file.write(html)
        text_file.close()
        config = pdfkit.configuration(wkhtmltopdf=r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe")
        pdfkit.from_file('datagilv08EDIT.html', 'datagilv08EDIT.pdf', configuration=config)
        source = os.getcwd()+r"\datagilv08EDIT.pdf"
        destination = os.getcwd()+r"\templates\edit\datagilv08EDIT.pdf"
        shutil.move(source, destination)
        src2 = os.getcwd()+r"\datagilv08EDIT.html"
        dest2 = os.getcwd()+r"\templates\files\datagilv08EDIT.html"
        shutil.move(src2, dest2)

        lista_impressoras = win32print.EnumPrinters(2)
        impressora = lista_impressoras[4]
        win32print.SetDefaultPrinter(impressora[2])  # 'L4160 Series(Rede)' -> Indice 2 que é o nome da impressora

        caminho = r"templates\edit"
        lista_arquivos = os.listdir(caminho)
        for arquivo in lista_arquivos:
            win32api.ShellExecute(0, "print", arquivo, None, caminho, 0)
        print(lista_arquivos)

        banco.commit()
        banco.close()
    # ZERAR OS CAMPOS
    # fazer if
    except sqlite3.Error as error:
        print("Erro ao inserir os dados: 5", error)


def imprimir_data():
    banco = sqlite3.connect('dbgilv02.db')
    cursor = banco.cursor()
    dataim = tela.datade.text()
    dataim2 = tela.datate.text()

    try:
        if dataim is not None:
            if dataim2 is None:
                frame = pd.read_sql_query(
                    "SELECT * FROM CLIENTES WHERE DATA_FECH <= '" + dataim + "'",
                    banco)
        if dataim2 is not None:
            if dataim is None:
                frame = pd.read_sql_query(
                    "SELECT * FROM CLIENTES WHERE DATA_FECH >='" + dataim2 + "'",
                    banco)
        if dataim and dataim2 is not None:
            frame = pd.read_sql_query(
                "SELECT * FROM CLIENTES WHERE DATA_FECH BETWEEN '" + dataim + "' AND '" + dataim2 + "'", banco)
        banco.commit()
        html = frame.to_html()
        text_file = open("datagilv08EDIT.html", "w")
        text_file.write(html)
        text_file.close()
        config = pdfkit.configuration(wkhtmltopdf=r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe")
        pdfkit.from_file('datagilv08EDIT.html', 'datagilv08EDIT.pdf', configuration=config)
        source = os.getcwd()+r"\datagilv08EDIT.pdf"
        destination = os.getcwd()+r"\templates\edit\datagilv08EDIT.pdf"
        shutil.move(source, destination)
        src2 = os.getcwd() + r"\datagilv08EDIT.html"
        dest2 = os.getcwd() + r"\templates\files\datagilv08EDIT.html"
        shutil.move(src2, dest2)
        caminho = r"templates\edit"
        lista_arquivos = os.listdir(caminho)
        for arquivo in lista_arquivos:
            win32api.ShellExecute(0, "print", arquivo, None, caminho, 0)
        print(lista_arquivos)

        banco.close()

    except sqlite3.Error as error:
        print("Erro ao inserir os dados: 5", error)


def imprime_nome_data():
    banco = sqlite3.connect('dbgilv02.db')
    cursor = banco.cursor()
    nomeim2 = tela.nomeImprimir_2.text()
    dataim2 = tela.datade_2.text()
    dataim22 = tela.datate_2.text()

    try:
        if nomeim2 is not None:
            frame = pd.read_sql_query("SELECT * FROM CLIENTES WHERE NOME = '" + nomeim2.upper() + "' AND DATA_FECH BETWEEN '" + dataim2 + "' AND '" + dataim22 + "'", banco)
        tela.nomeImprimir_2.setText("")

        banco.commit()
        html = frame.to_html()
        text_file = open("datagilv08EDIT.html", "w")
        text_file.write(html)
        text_file.close()
        config = pdfkit.configuration(wkhtmltopdf=r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe")
        pdfkit.from_file('datagilv08EDIT.html', 'datagilv08EDIT.pdf', configuration=config)
        source = os.getcwd()+r"\datagilv08EDIT.pdf"
        destination = os.getcwd()+r"\templates\edit\datagilv08EDIT.pdf"
        shutil.move(source, destination)
        src2 = os.getcwd() + r"\datagilv08EDIT.html"
        dest2 = os.getcwd() + r"\templates\files\datagilv08EDIT.html"
        shutil.move(src2, dest2)

        caminho = r"templates\edit"
        lista_arquivos = os.listdir(caminho)
        for arquivo in lista_arquivos:
            win32api.ShellExecute(0, "print", arquivo, None, caminho, 0)
        print(lista_arquivos)
        banco.close()

    except sqlite3.Error as error:
        print("Erro ao inserir os dados: 5", error)


def imprime_produto():
    banco = sqlite3.connect('dbgilv02.db')
    cursor = banco.cursor()
    prod2 = tela.produtoImprimir.text()

    if prod2 is not None:
        frame = pd.read_sql_query("SELECT * FROM CLIENTES WHERE PRODUTO LIKE '"+prod2.upper()+"%'", banco)
        tela.produtoImprimir.setText("")

    banco.commit()
    html = frame.to_html()
    text_file = open("datagilv08EDIT.html", "w")
    text_file.write(html)
    text_file.close()
    config = pdfkit.configuration(wkhtmltopdf=r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe")
    pdfkit.from_file('datagilv08EDIT.html', 'datagilv08EDIT.pdf', configuration=config)
    source = os.getcwd()+r"\datagilv08EDIT.pdf"
    destination = os.getcwd()+r"\templates\edit\datagilv08EDIT.pdf"
    shutil.move(source, destination)
    src2 = os.getcwd() + r"\datagilv08EDIT.html"
    dest2 = os.getcwd() + r"\templates\files\datagilv08EDIT.html"
    shutil.move(src2, dest2)

    caminho = r"templates\edit"
    lista_arquivos = os.listdir(caminho)
    for arquivo in lista_arquivos:
        win32api.ShellExecute(0, "print", arquivo, None, caminho, 0)
    print(lista_arquivos)
    banco.close()


"""def imprime_valor():
    banco = sqlite3.connect('dbgilv02.db')
    cursor = banco.cursor()
    valorsuperior = tela.valorSupImprimir.text()
    valorinferior = tela.valorInfImprimir.text()

    if valorsuperior is not None:
        frame = pd.read_sql_query("SELECT * FROM CLIENTES WHERE VALOR > '" + valorsuperior + "'", banco)
        tela.valorSupImprimir.setText("")
    elif valorinferior is not None:
        frame = pd.read_sql_query("SELECT * FROM CLIENTES WHERE VALOR < '" + valorinferior + "'", banco)
        tela.valorInfImprimir.setText("")
    else:
        frame = pd.read_sql_query("SELECT * FROM CLIENTES WHERE VALOR BETWEEN '" + valorinferior + "' AND '"+valorsuperior+"'", banco)

    banco.commit()
    html = frame.to_html()
    text_file = open("datagilv08EDIT.html", "w")
    text_file.write(html)
    text_file.close()
    config = pdfkit.configuration(wkhtmltopdf=r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe")
    pdfkit.from_file('datagilv08EDIT.html', 'datagilv08EDIT.pdf', configuration=config)
    source = os.getcwd()+r"\datagilv08EDIT.pdf"
    destination = os.getcwd()+r"\templates\edit\datagilv08EDIT.pdf"
    shutil.move(source, destination)
    src2 = os.getcwd() + r"\datagilv08EDIT.html"
    dest2 = os.getcwd() + r"\templates\files\datagilv08EDIT.html"
    shutil.move(src2, dest2)

    caminho = r"templates\edit"
    lista_arquivos = os.listdir(caminho)
    for arquivo in lista_arquivos:
        win32api.ShellExecute(0, "print", arquivo, None, caminho, 0)
    print(lista_arquivos)
    banco.close()"""


def imprime_condicao():
    banco = sqlite3.connect('dbgilv02.db')
    cursor = banco.cursor()
    cond1 = tela.pagamentoImprimir.text()
    if cond1 is not None:
        frame = pd.read_sql_query("SELECT * FROM CLIENTES WHERE COND_PAG LIKE '" +cond1.upper()+"%'", banco)
        tela.pagamentoImprimir.setText("")

    banco.commit()
    html = frame.to_html()
    text_file = open("datagilv08EDIT.html", "w")
    text_file.write(html)
    text_file.close()
    config = pdfkit.configuration(wkhtmltopdf=r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe")
    pdfkit.from_file('datagilv08EDIT.html', 'datagilv08EDIT.pdf', configuration=config)
    source = os.getcwd()+r"\datagilv08EDIT.pdf"
    destination = os.getcwd()+r"\templates\edit\datagilv08EDIT.pdf"
    shutil.move(source, destination)
    src2 = os.getcwd() + r"\datagilv08EDIT.html"
    dest2 = os.getcwd() + r"\templates\files\datagilv08EDIT.html"
    shutil.move(src2, dest2)

    caminho = r"templates\edit"
    lista_arquivos = os.listdir(caminho)
    for arquivo in lista_arquivos:
        win32api.ShellExecute(0, "print", arquivo, None, caminho, 0)
    print(lista_arquivos)
    banco.close()


app = QtWidgets.QApplication([])
tela = uic.loadUi("giv04.ui")
tela.btnEnviar.clicked.connect(salvar_dados)
tela.btnConsultar.clicked.connect(listar_dados)
tela.btnConsultarCondicao.clicked.connect(listar_dados_2)
tela.btnApagar.clicked.connect(apagar_dados)
tela.btnImprimir.clicked.connect(imprimir_dados)
#tela.btnImprimir_2.clicked.connect(imprimir_dados_2)
tela.btnLimparDados.clicked.connect(limpar_dados)
tela.btnImprimeNome.clicked.connect(imprimir_nome)
tela.btnImprimeData.clicked.connect(imprimir_data)
tela.btnImprimeNomeData.clicked.connect(imprime_nome_data)
tela.btnImprimeProduto.clicked.connect(imprime_produto)
tela.btnImprimeCondicao.clicked.connect(imprime_condicao)
tela.btnImprimirHTML.clicked.connect(imprime_HTML)
tela.show()
app.exec()
