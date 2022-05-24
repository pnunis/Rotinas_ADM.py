import win32com.client
import subprocess
import pandas as pd
from UserID2 import login,senha
import os
import glob
import shutil
import time
import win32clipboard
import pyperclip
import re
import numpy as np
import sys
from datetime import datetime
# ACESSO AS PASTAS PARA RATEIO
path = r'C:\Users\paulo.souza\Desktop\Rateio\Maio'
temp_path = r'C:\Users\paulo.souza\Desktop\Rateio\Maio\Subindo'
path_success = r'C:\Users\paulo.souza\Desktop\Rateio\Maio\1 - Sucesso'
path_error = r'C:\Users\paulo.souza\Desktop\Rateio\Maio\2 - Arquivo com erro(checar)'
path_fail = r'C:\Users\paulo.souza\Desktop\Rateio\Maio\3 - Não processado'
localarquivos = glob.glob(path + '/*.csv')

# ACESSO DE ABRIR APP SAP PARA TESTE
subprocess.Popen("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")
# FOI CRIADO TEMPO PARA OS BUG DO SISTEMA
time.sleep(3)
SapGuiAuto = win32com.client.GetObject('SAPGUI')
application = SapGuiAuto.GetScriptingEngine
connection = application.OpenConnection("ECC NOVO PRD")

session = connection.Children(0)
# DIGITAÇÃO DE USUARIO E SENHA
session.findById("wnd[0]").resizeWorkingPane(164,40,False)
session.findById("wnd[0]/usr/txtRSYST-BNAME").text = login
session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = senha
session.findById("wnd[0]/usr/pwdRSYST-BCODE").setFocus()
session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 14
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]").maximize()

def subir_arquivos():

    # Acessando transação zfir103
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nzfir103"
    session.findById("wnd[0]").sendVKey(0)

    # Laço para percorrer todos arquivos dentro da pasta com arquivos para subir
    for arquivo in localarquivos:
        # Pega nome do arquivo e coloca no formato de caminho "\nome_arquivo.csv"
        arqv = '\\' + os.path.basename(arquivo)
        # print(os.path.dirname(arquivo))

        # Verifica se no arquivo o numero da divisão está diferente do Centro de custo e coloca na variavel "teste"
        # "Ok" para não e "Erro" para sim
        df = pd.read_csv(arquivo, sep=';', decimal=',', encoding='ISO-8859-1', dtype=object).fillna(0)
        df['Centro'] = df['Centro de Custo'].str[5:7].str.lstrip('0').fillna(0)
        df["Teste"] = np.where(df['Divisao'] == df['Centro'], "OK", "Erro")
        lista = pd.Series(df['Teste']).unique()
        teste = "Erro" in lista
        # Move arquivo para dentro de uma pasta temporária, pois o SAP necessita de autorização para arquivos novos.
        # Mantendo o nome padrão (SUBIR RATEIO.csv), se evita esse problema na hora de subir o arquivo
        shutil.move(path + arqv, temp_path + '\\SUBIR RATEIO.csv')

        # Se arquivo tiver erro na verificação de Divisão = Centro de custo, move o arquivo para pasta Erro sem subi-lo.
        # incluindo o motivo do erro no final do nome do arquivo.
        if teste == "Erro":
            shutil.move(temp_path + '\\SUBIR RATEIO.csv',
                        path_error + arqv[:-4] + " " + " Divisão diferente do Centro de custo" + '.csv')

        # Seleciona arquivo de rateio dentro da pasta
        else:
            # noinspection PyBroadException
            try:
                session.findById("wnd[0]").maximize()
                session.findById("wnd[0]/usr/ctxtP_FILE").text = temp_path + '\\SUBIR RATEIO.csv'
                session.findById("wnd[0]/usr/ctxtP_FILE").caretPosition = 73
                session.findById("wnd[0]").sendVKey(8)
                # Copia codigo de lançamento '@08@' '@0A@' para o clipboard e salva na variavel "cod_texto"
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItemByPosition(0)
                win32clipboard.OpenClipboard()
                texto = win32clipboard.GetClipboardData()
                cod_texto = texto[:4]
                win32clipboard.CloseClipboard()
                time.sleep(0.2)
                # Copia texto da mensagem e salva na variavel "texto2"
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = "MESSAGE"
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn("MESSAGE")
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").deselectColumn("TP_MENS")
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItemByPosition(0)
                texto2 = pd.read_clipboard(sep='\t', header=None)
                # Coloca textos na mesma linha separado por um hifen
                # txt_linha = texto2[0].str.cat(sep=' - ')
                # remover caracteres ":" do texto
                # txt_linha = re.sub(':', "", txt_linha)
                # txt_linha = re.sub(r'([\/:*?"<>|])', "", txt_linha)
                txt_linha = texto2[0].str.cat(sep='\n')
                # expressão regular para remover espaços em excesso
                # txt_linha = re.sub(r"\s+", " ", txt_linha)
                session.findById("wnd[0]").sendVKey(3)
                # Verifica se codigo de lançamento for @08@ mova arquivo para pasta sucesso
                if cod_texto == '@08@':
                    shutil.move(temp_path + '\\SUBIR RATEIO.csv', path_success + arqv)
                # Verifica se cod de lançamento for @0A@ mova arquivo para pasta falha e escreva motivo no nome do arquivo
                if cod_texto == '@0A@':
                    print(txt_linha, file=open(path_fail + arqv[:-4] + ' Descricao do Erro' + '.txt', 'w'))
                    shutil.move(temp_path + '\\SUBIR RATEIO.csv', path_fail + arqv)
                    # shutil.move(temp_path + '\\SUBIR RATEIO.csv', path_fail + arqv[:-4] + " " + txt_linha + '.csv')
            # Exceção para identificar arquivo que a transação nao conseguiu subir. É um arquivo com erro
            except:
                shutil.move(temp_path + '\\SUBIR RATEIO.csv',
                            path_error + arqv[:-4] + " " + " Arquivo com erro não identificado" + '.csv')


subir_arquivos()
