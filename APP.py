import win32com.client
import time
from datetime import datetime
import traceback
import pandas as pd
from pywinauto.application import Application
import os
from dotenv import load_dotenv # Importa a função load_dotenv

# Carrega as variáveis do arquivo .env
load_dotenv()

def fechar_sap_existente():
    """
    Tenta fechar todas as instâncias e sessões existentes do SAP GUI.
    Esta função é robusta e não depende de uma sessão pré-existente.
    """
    print("Tentando fechar instâncias existentes do SAP...")
    try:
        # Tenta obter o objeto SAPGUI. Se não houver SAP GUI aberto, isso falhará.
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine

        # Itera sobre todas as conexões abertas
        for i in range(application.Connections.Count):
            connection = application.Children(i)
            # Itera sobre todas as sessões em cada conexão
            for j in range(connection.Sessions.Count):
                session = connection.Children(j)
                try:
                    # Tenta maximizar e fechar a janela principal da sessão
                    session.findById("wnd[0]").maximize() # Corrigido: .maximize() é um método
                    session.findById("wnd[0]").close()
                    # Tenta lidar com o diálogo de logoff, se aparecer
                    try:
                        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press() # Pressiona "Sim" ou "OK"
                        print(f"Diálogo de logoff tratado para sessão {j} da conexão {i}.")
                    except Exception:
                        # Se o diálogo não aparecer, apenas ignora
                        print(f"Nenhum diálogo de logoff encontrado para sessão {j} da conexão {i}.")
                    print(f"Sessão {j} da conexão {i} fechada com sucesso.")
                except Exception as e:
                    print(f"Erro ao fechar sessão {j} da conexão {i}: {e}")
    except Exception as e:
        # Captura o erro se o objeto SAPGUI não for encontrado (SAP não está aberto)
        print("Nenhuma instância do SAP GUI encontrada ou erro ao acessar: ", e)
    print("Tentativa de fechamento de instâncias SAP concluída.")

def fechar_pastas_trabalho_excel():
    """
    Fecha todas as pastas de trabalho abertas do Excel.
    """
    print("Fechando pastas de trabalho do Excel...")
    try:
        xl = win32com.client.Dispatch("Excel.Application")
        xl.DisplayAlerts = False # Desativa alertas para evitar pop-ups
        for wb in xl.Workbooks:
            print(f"Fechando pasta de trabalho: {wb.Name}")
            wb.Close(SaveChanges=False) # Fecha sem salvar alterações
        xl.Quit() # Fecha o aplicativo Excel
        print("Todas as pastas de trabalho do Excel foram fechadas.")
    except Exception as e:
        print("Erro ao fechar as pastas de trabalho do Excel:", e)

print("Iniciando processo...")
# Chama a função para fechar qualquer instância existente do SAP antes de iniciar uma nova
fechar_sap_existente()
fechar_pastas_trabalho_excel()
            
data_atual = datetime.now()
data_convertida = data_atual.strftime('%d.%m.%Y')
print(f"Data atual: {data_convertida}")

# Caminho para o executável SAP Logon
sap_logon_path = r'C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe'

print("Iniciando SAP Logon...")
app = Application(backend="uia").start(sap_logon_path)
time.sleep(5) # Espera o SAP Logon abrir
print("SAP Logon aberto com sucesso.")

SapGuiAuto = win32com.client.GetObject('SAPGUI')
application = SapGuiAuto.GetScriptingEngine

print("Conectando ao S/4HANA PS4...")
# Abre a conexão com o sistema SAP especificado
connection = application.OpenConnection('S/4HANA PS4', True)
time.sleep(3) # Espera a conexão ser estabelecida
session = connection.Children(0) # Obtém a primeira sessão da conexão
session.findById('wnd[0]').maximize() # Corrigido: .maximize() é um método
print("Conexão estabelecida com sucesso.")

# O bloco WScript foi comentado, pois pode estar causando o erro "Cadeia de caracteres de classe inválida".
# Se a manipulação de eventos SAP GUI for necessária, verifique a configuração do WScript em seu sistema.
# try:
#     # Conecta objetos para manipulação de eventos (se necessário)
#     WScript = win32com.client.Dispatch("WScript")
#     WScript.ConnectObject(session, "on")
#     WScript.ConnectObject(application, "on") # Usar 'application' em vez de 'App'
# except Exception as e:
#     print(f"Erro ao conectar objetos WScript: {e}")

session.FindById("wnd[0]").Maximize() # Corrigido: .Maximize() é um método

# Obtém usuário e senha das variáveis de ambiente (agora lidas do .env)
sap_usuario = os.getenv('SAP_USER')
sap_senha = os.getenv('SAP_PASSWORD')

# Verifica se as variáveis de ambiente foram carregadas
if not sap_usuario:
    print("Erro: Variável de ambiente 'SAP_USER' não encontrada ou vazia. Verifique seu arquivo .env.")
    exit()
if not sap_senha:
    print("Erro: Variável de ambiente 'SAP_PASSWORD' não encontrada ou vazia. Verifique seu arquivo .env.")
    exit()

print("Realizando login no SAP...")
# Preenche os campos de usuário e senha
session.FindById("wnd[0]/usr/txtRSYST-BNAME").Text = sap_usuario
session.FindById("wnd[0]/usr/pwdRSYST-BCODE").Text = sap_senha
session.FindById("wnd[0]/usr/pwdRSYST-BCODE").SetFocus() # Corrigido: .SetFocus() é um método
session.FindById("wnd[0]/usr/pwdRSYST-BCODE").CaretPosition = 8
session.findById("wnd[0]").sendVKey(0) # Pressiona Enter para logar
print("Login realizado com sucesso.")

print("Acessando a transação ZPMMT_287...")
session.findById("wnd[0]").maximize()
session.findById("wnd[0]/tbar[0]/okcd").text = "ZPMMT_287"
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = "01.01.2024"
session.findById("wnd[0]/usr/ctxtS_DATA-HIGH").text = data_convertida
session.findById("wnd[0]/usr/ctxtS_DATA-HIGH").setFocus()
session.findById("wnd[0]/usr/ctxtS_DATA-HIGH").caretPosition = 10
session.findById("wnd[0]/usr/btn%_S_CENT1_%_APP_%-VALU_PUSH").press()
session.findById("wnd[1]/tbar[0]/btn[23]").press()
session.findById("wnd[2]/usr/ctxtDY_PATH").text = r"C:\Users\3976339\Desktop\ONTIME"
session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "CODIGO BASES.txt"
session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 16
session.findById("wnd[2]/tbar[0]/btn[0]").press()
session.findById("wnd[1]/tbar[0]/btn[8]").press()
session.findById("wnd[0]/usr/btn%_S_CENT2_%_APP_%-VALU_PUSH").press()
session.findById("wnd[1]/tbar[0]/btn[23]").press()
session.findById("wnd[2]/usr/ctxtDY_PATH").text = r"C:\Users\3976339\Desktop\ONTIME"
session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "CODIGO BASES.txt"
session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 16
session.findById("wnd[2]/tbar[0]/btn[0]").press()
session.findById("wnd[1]/tbar[0]/btn[8]").press()
session.findById("wnd[0]/tbar[1]/btn[8]").press()
print("Exportando dados para Excel...")
session.findById("wnd[0]/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
session.findById("wnd[0]/shellcont/shell").selectContextMenuItem("&XXL")
session.findById("wnd[1]/usr/ctxtDY_PATH").text = r"C:\Users\3976339\Desktop\ONTIME"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZPMMT.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[11]").press()
print("Dados exportados para ZPMMT.xlsx")

print("Processando arquivo ZPMMT.xlsx...")
Requisicao = pd.read_excel(r'C:\Users\3976339\Desktop\ONTIME\ZPMMT.xlsx')
Requisicao_zp = Requisicao.loc[:,['Requisição de Compras']]

caminho_pasta_req = r"C:\Users\3976339\Desktop\ONTIME"
# Corrigido: Usando os.path.join para construir o caminho do arquivo
Nome_Arquivo_zpmmt = os.path.join(caminho_pasta_req, 'ZPMMT_REQ.txt')
Requisicao_zp.to_csv(Nome_Arquivo_zpmmt, index=False)
print("Arquivo ZPMMT_REQ.txt criado com sucesso.")

print("Acessando a transação SE16N para tabela EBAN...")
session.findById("wnd[0]").maximize()
session.findById("wnd[0]/tbar[0]/okcd").text = "/NSE16N"
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/usr/ctxtGD-TAB").text = "EBAN"
session.findById("wnd[0]/usr/ctxtGD-TAB").setFocus()
session.findById("wnd[0]/usr/ctxtGD-TAB").caretPosition = 4
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").setFocus()
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").press()
session.findById("wnd[1]/tbar[0]/btn[21]").press()
session.findById("wnd[2]/usr/ctxtDY_PATH").text = r"C:\Users\3976339\Desktop\ONTIME"
session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = r"ZPMMT_REQ.txt"
session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[2]/tbar[0]/btn[0]").press() 
session.findById("wnd[1]/tbar[0]/btn[8]").press()
session.findById("wnd[0]/usr/ctxtGD-VARIANT").text = "/LOG_ONTIME"
session.findById("wnd[0]/usr/txtGD-MAX_LINES").text = ""
session.findById("wnd[0]/usr/txtGD-MAX_LINES").setFocus()
session.findById("wnd[0]/usr/txtGD-MAX_LINES").caretPosition = 0
session.findById("wnd[0]/tbar[1]/btn[8]").press()
print("Exportando dados da tabela EBAN...")
session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&XXL")
session.findById("wnd[1]/usr/ctxtDY_PATH").text = r"C:\Users\3976339\Desktop\ONTIME"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "EBAN.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
session.findById("wnd[1]/tbar[0]/btn[11]").press()
session.findById("wnd[0]/tbar[0]/btn[3]").press()
print("Dados da tabela EBAN exportados para EBAN.xlsx")

print("Processando tabela EKET...")
session.findById("wnd[0]/usr/ctxtGD-TAB").text = "EKET"
session.findById("wnd[0]/usr/ctxtGD-TAB").caretPosition = 4
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]").sendVKey(71)
session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").text = "BANFN"
session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
session.findById("wnd[1]").sendVKey(0)
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,0]").setFocus()
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,0]").press()
session.findById("wnd[1]/tbar[0]/btn[21]").press()
session.findById("wnd[2]/usr/ctxtDY_PATH").text = r"C:\Users\3976339\Desktop\ONTIME"
session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = r"ZPMMT_REQ.txt"
session.findById("wnd[2]/tbar[0]/btn[0]").press()
session.findById("wnd[1]/tbar[0]/btn[8]").press()
session.findById("wnd[0]/usr/ctxtGD-VARIANT").setFocus()
session.findById("wnd[0]/usr/ctxtGD-VARIANT").caretPosition = 0
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/usr/ctxtGD-VARIANT").text = "/LOG_ONTIME"
session.findById("wnd[0]/usr/txtGD-MAX_LINES").text = ""
session.findById("wnd[0]/usr/txtGD-MAX_LINES").setFocus()
session.findById("wnd[0]/usr/txtGD-MAX_LINES").caretPosition = 0
session.findById("wnd[0]/tbar[1]/btn[8]").press()
print("Exportando dados da tabela EKET...")
session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&XXL")
session.findById("wnd[1]/usr/ctxtDY_PATH").text = r"C:\Users\3976339\Desktop\ONTIME"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "EKET.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
session.findById("wnd[1]/tbar[0]/btn[11]").press()
session.findById("wnd[0]/tbar[0]/btn[3]").press()
print("Dados da tabela EKET exportados para EKET.XLSX")

print("Lendo e consolidando dados...")
base_eket = pd.read_excel(r'C:\Users\3976339\Desktop\ONTIME\EKET.xlsx')
base_eban = pd.read_excel(r'C:\Users\3976339\Desktop\ONTIME\EBAN.xlsx')

coluna_pedido_eket = base_eket['Documento de compras']
coluna_pedido_eban = base_eban['Pedido']

df_pedido_consolidado = pd.concat([coluna_pedido_eket, coluna_pedido_eban], axis=0).drop_duplicates().reset_index(drop=True)
df_pedido_consolidado = df_pedido_consolidado.dropna().astype(int)

df_pedido_consolidado.to_csv(r"C:\Users\3976339\Desktop\ONTIME\PEDIDOS_CONSOLIDADO.txt", index=False, header=False)
print("Arquivo PEDIDOS_CONSOLIDADO.txt criado com sucesso.")

print("Processando tabela LIPS...")
session.findById("wnd[0]").maximize()
session.findById("wnd[0]/usr/ctxtGD-TAB").text = "LIPS"
session.findById("wnd[0]/usr/ctxtGD-TAB").setFocus()
session.findById("wnd[0]/usr/ctxtGD-TAB").caretPosition = 4
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]").sendVKey(71)
session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").text = "VGBEL"
session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,0]").setFocus()
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,0]").press()
session.findById("wnd[1]/tbar[0]/btn[21]").press()
session.findById("wnd[2]/usr/ctxtDY_PATH").text = r"C:\Users\3976339\Desktop\ONTIME"
session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "PEDIDOS_CONSOLIDADO.txt"
session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 23
session.findById("wnd[2]/tbar[0]/btn[0]").press()
session.findById("wnd[1]/tbar[0]/btn[8]").press()
session.findById("wnd[0]/usr/ctxtGD-VARIANT").text = "/LOG_ONTIME"
session.findById("wnd[0]/usr/txtGD-MAX_LINES").text = ""
session.findById("wnd[0]/usr/txtGD-MAX_LINES").setFocus()
session.findById("wnd[0]/usr/txtGD-MAX_LINES").caretPosition = 0
session.findById("wnd[0]/tbar[1]/btn[8]").press()
print("Exportando dados da tabela LIPS...")
session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&XXL")
session.findById("wnd[1]/usr/ctxtDY_PATH").text = r"C:\Users\3976339\Desktop\ONTIME"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "LIPS.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
session.findById("wnd[1]/tbar[0]/btn[11]").press()
session.findById("wnd[0]/tbar[0]/btn[3]").press()
print("Dados da tabela LIPS exportados para LIPS.XLSX")

print("Processando arquivo LIPS.xlsx...")
remessa = pd.read_excel(r'C:\Users\3976339\Desktop\ONTIME\LIPS.xlsx')
remessa_zp = remessa.loc[:,['Remessa']]
remessa_zp.to_csv(r"C:\Users\3976339\Desktop\ONTIME\REMESSA.txt", index=False, header=False)
print("Arquivo REMESSA.txt criado com sucesso.")

print("Processando tabela VBFA...")
session.findById("wnd[0]").maximize()
session.findById("wnd[0]/usr/ctxtGD-TAB").text = "VBFA"
session.findById("wnd[0]/usr/ctxtGD-TAB").setFocus()
session.findById("wnd[0]/usr/ctxtGD-TAB").caretPosition = 4
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,2]").setFocus()
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,2]").press()
session.findById("wnd[1]/tbar[0]/btn[21]").press()
session.findById("wnd[2]/usr/ctxtDY_PATH").text = r"C:\Users\3976339\Desktop\ONTIME"
session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "REMESSA.txt"
session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 11
session.findById("wnd[2]/tbar[0]/btn[0]").press()
session.findById("wnd[1]/tbar[0]/btn[8]").press()
session.findById("wnd[0]").sendVKey(71)
session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").text = "BWART"
session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
session.findById("wnd[1]").sendVKey(0)
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,0]").setFocus()
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,0]").press()
session.findById("wnd[1]/usr/tblSAPLSE16NMULTI_TC/ctxtGS_MULTI_SELECT-LOW[1,0]").text = "101"
session.findById("wnd[1]/usr/tblSAPLSE16NMULTI_TC/ctxtGS_MULTI_SELECT-LOW[1,1]").text = "862"
session.findById("wnd[1]/usr/tblSAPLSE16NMULTI_TC/ctxtGS_MULTI_SELECT-LOW[1,2]").text = "861"
session.findById("wnd[1]/usr/tblSAPLSE16NMULTI_TC/ctxtGS_MULTI_SELECT-LOW[1,2]").setFocus()
session.findById("wnd[1]/usr/tblSAPLSE16NMULTI_TC/ctxtGS_MULTI_SELECT-LOW[1,2]").caretPosition = 3
session.findById("wnd[1]/tbar[0]/btn[8]").press()
session.findById("wnd[0]/usr/ctxtGD-VARIANT").text = "/LOG_ONTIME"
session.findById("wnd[0]/usr/txtGD-MAX_LINES").text = ""
session.findById("wnd[0]/usr/txtGD-MAX_LINES").setFocus()
session.findById("wnd[0]/usr/txtGD-MAX_LINES").caretPosition = 0
session.findById("wnd[0]/tbar[1]/btn[8]").press()
print("Exportando dados da tabela VBFA...")
session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&XXL")
session.findById("wnd[1]/usr/ctxtDY_PATH").text = r"C:\Users\3976339\Desktop\ONTIME"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "VBFA.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
session.findById("wnd[1]/tbar[0]/btn[11]").press()
print("Dados da tabela VBFA exportados para VBFA.XLSX")

print("Processando arquivo VBFA.xlsx...")
base_vbfa = pd.read_excel(r'C:\Users\3976339\Desktop\ONTIME\VBFA.xlsx')
# Corrigido: O filtro deve ser aplicado diretamente em base_vbfa
base_filtrada = base_vbfa[base_vbfa['Tipo de movimento'].isin([101, 862])]
base_filtrada['Concatenado'] = base_filtrada['Doc.subsequente'].astype(str) + base_filtrada['Ano doc.material'].astype(str)
base_filtrada['Concatenado'].to_csv(r'C:\Users\3976339\Desktop\ONTIME\VBFA_CONSOLIDADO.txt', index=False, header=False)
print("Arquivo VBFA_CONSOLIDADO.txt criado com sucesso.")

print("Processando tabela J_1BNFLIN...")
session.findById("wnd[0]").maximize()
session.findById("wnd[0]/usr/txtGD-TAB").setFocus()
session.findById("wnd[0]/usr/txtGD-TAB").caretPosition = 0
session.findById("wnd[0]/tbar[0]/btn[3]").press()
session.findById("wnd[0]/usr/ctxtGD-TAB").text = "J_1BNFLIN"
session.findById("wnd[0]/usr/ctxtGD-TAB").caretPosition = 9
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]").sendVKey(71)
session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").text = "REFKEY"
session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 6
session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,0]").setFocus()
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,0]").press()
session.findById("wnd[1]/tbar[0]/btn[21]").press()
session.findById("wnd[2]/usr/ctxtDY_PATH").text = r"C:\Users\3976339\Desktop\ONTIME"
session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "VBFA_CONSOLIDADO.txt"
session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 20
session.findById("wnd[2]/tbar[0]/btn[0]").press()
session.findById("wnd[1]/tbar[0]/btn[8]").press()
session.findById("wnd[0]/usr/ctxtGD-VARIANT").text = "/LOG_ONTIME"
session.findById("wnd[0]/usr/txtGD-MAX_LINES").text = ""
session.findById("wnd[0]/usr/txtGD-MAX_LINES").setFocus()
session.findById("wnd[0]/usr/txtGD-MAX_LINES").caretPosition = 0
session.findById("wnd[0]/tbar[1]/btn[8]").press()
print("Exportando dados da tabela J_1BNFLIN...")
session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&XXL")
session.findById("wnd[1]/usr/ctxtDY_PATH").text = r"C:\Users\3976339\Desktop\ONTIME"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "J_1BNFLIN.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 14
session.findById("wnd[1]/tbar[0]/btn[11]").press()
print("Dados da tabela J_1BNFLIN exportados para J_1BNFLIN.xlsx")

print("Processando arquivo J_1BNFLIN.xlsx...")
jlin = pd.read_excel(r'C:\Users\3976339\Desktop\ONTIME\J_1BNFLIN.xlsx')
jlin_zp = jlin.loc[:,['Nº documento']]
jlin_zp.to_csv(r"C:\Users\3976339\Desktop\ONTIME\JLIN.txt", index=False, header=False)
print("Arquivo JLIN.txt criado com sucesso.")

print("Processando tabela J_1BNFDOC...")
session.findById("wnd[0]").maximize()
session.findById("wnd[0]/tbar[0]/btn[3]").press()
session.findById("wnd[0]/usr/ctxtGD-TAB").text = "J_1BNFDOC"
session.findById("wnd[0]/usr/ctxtGD-TAB").caretPosition = 9
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").setFocus()
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").press()
session.findById("wnd[1]/tbar[0]/btn[21]").press()
session.findById("wnd[2]/usr/ctxtDY_PATH").text = r"C:\Users\3976339\Desktop\ONTIME"
session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "JLIN.txt"
session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[2]/tbar[0]/btn[0]").press()
session.findById("wnd[1]/tbar[0]/btn[8]").press()
session.findById("wnd[0]/usr/ctxtGD-VARIANT").text = "/LOG_ONTIME"
session.findById("wnd[0]/usr/txtGD-MAX_LINES").text = ""
session.findById("wnd[0]/usr/txtGD-MAX_LINES").setFocus()
session.findById("wnd[0]/usr/txtGD-MAX_LINES").caretPosition = 0
session.findById("wnd[0]/tbar[1]/btn[8]").press()
print("Exportando dados da tabela J_1BNFDOC...")
session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&XXL")
session.findById("wnd[1]/usr/ctxtDY_PATH").text = r"C:\Users\3976339\Desktop\ONTIME"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "J_1BNFDOC.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 14
session.findById("wnd[1]/tbar[0]/btn[11]").press()
print("Dados da tabela J_1BNFDOC exportados para J_1BNFDOC.XLSX")

print("Processando arquivo ZPMMT.xlsx para tabela MARA...")
mara = pd.read_excel(r'C:\Users\3976339\Desktop\ONTIME\ZPMMT.xlsx')
mara_zp = mara.loc[:,['Material']]
mara_zp.to_csv(r"C:\Users\3976339\Desktop\ONTIME\MARA.txt", index=False, header=False)
print("Arquivo MARA.txt criado com sucesso.")

print("Processando tabela MARA...")
session.findById("wnd[0]").maximize()
session.findById("wnd[0]/tbar[0]/btn[3]").press()
session.findById("wnd[0]/usr/ctxtGD-TAB").text = "MARA"
session.findById("wnd[0]/usr/ctxtGD-TAB").caretPosition = 4
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").setFocus()
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").press()
session.findById("wnd[1]/tbar[0]/btn[21]").press()
session.findById("wnd[2]/usr/ctxtDY_PATH").text = r"C:\Users\3976339\Desktop\ONTIME"
session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "MARA.txt"
session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[2]/tbar[0]/btn[0]").press()
session.findById("wnd[1]/tbar[0]/btn[8]").press()
session.findById("wnd[0]/usr/ctxtGD-VARIANT").text = "/LOG_ONTIME"
session.findById("wnd[0]/usr/txtGD-MAX_LINES").text = ""
session.findById("wnd[0]/usr/txtGD-MAX_LINES").setFocus()
session.findById("wnd[0]/usr/txtGD-MAX_LINES").caretPosition = 0
session.findById("wnd[0]/tbar[1]/btn[8]").press()
print("Exportando dados da tabela MARA...")
session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&XXL")
session.findById("wnd[1]/usr/ctxtDY_PATH").text = r"C:\Users\3976339\Desktop\ONTIME"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "MARA.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
session.findById("wnd[1]/tbar[0]/btn[11]").press()
print("Dados da tabela MARA exportados para MARA.XLSX")

print("Extrações SAP Conluídas!")