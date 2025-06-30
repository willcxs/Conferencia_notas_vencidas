
import os
import pandas as pd
import win32com.client
import psutil
import signal
import time
import pyautogui
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import pyautogui


# Data de hoje
today = datetime.today()
# Subtrai 30 dias
ha_30_dias = today - timedelta(days=30)
#formato datas
today = today.strftime('%d.%m.%Y')
ha_30_dias = ha_30_dias.strftime('%d.%m.%Y')

# CONFIGURAÇÕES (ajuste conforme seu ambiente local)
# Caminho do SAP Logon
saplogon_path = r"C:\Path\To\SAP\SAPLogon.exe"

#Credenciais
credentials_df = pd.read_excel(r"C:\Path\To\Credentials\SAP_Credentials.xlsx")
user =  credentials_df['usuario'][0]
password = credentials_df['senha'][0]

#Caso exista copias anteriores apagar
try:
    os.remove(r"C:\Path\To\merc_nao_entregue.XLSX")
    
except:
    print('Arquivo não encontrado')


# Fecha todos os processos do SAP GUI
os.system("taskkill /F /IM saplogon.exe")

# Inicia o SAP Logon
os.startfile(saplogon_path)
time.sleep(10)


# Conecta ao SAP GUI
SapGui = win32com.client.GetObject("SAPGUI")
application = SapGui.GetScriptingEngine
connection = application.OpenConnection("pe1", True)  

time.sleep(5)
connection = application.Children(0)
session = connection.Children(0)
session.findById("wnd[0]").maximize



# Preenche os campos de login
session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "400"  # Cliente
session.findById("wnd[0]/usr/txtRSYST-BNAME").text = user  # Usuário
session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password  # Senha
session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "PT"  # Idioma

# Entrar
session.findById("wnd[0]").sendVKey(0)

print("Login realizado com sucesso!")

#Acessar transação
session.StartTransaction(Transaction="FBL5N")

#Chamar variante
session.findById("wnd[0]").sendVKey(17)
session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
session.findById("wnd[1]/tbar[0]/btn[8]").press()


# %%
# Acessa a tabela ALV
alv_table = session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell")

# Valor a ser buscado
valor_busca = "MERC_N_ENTREGU"  # Substitua pelo valor que você está procurando

# Número de linhas na tabela
num_linhas = alv_table.RowCount

# Itera pelas linhas da tabela para encontrar o valor desejado
for i in range(num_linhas):
    if alv_table.GetCellValue(i, 'VARIANT') == valor_busca:
        alv_table.selectedRows = str(i)  # Seleciona a linha encontrada
        break

session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows =  alv_table.selectedRows
session.findById("wnd[1]").sendVKey(2)

#Roda a transação
session.findById("wnd[0]").sendVKey(8)

#exportar os Dados
session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Path\To\Mercadorias não entregues"
session.findById("wnd[1]/tbar[0]/btn[0]").press()

# %%
time.sleep(15)
# Fechar excel
[os.kill(proc.pid, signal.SIGTERM) for proc in psutil.process_iter(['name']) if "EXCEL.EXE" in proc.info['name']]
time.sleep(3)

os.rename("C:\Path\To\export.XLSX",rf"C:\Path\To\merc_nao_entregue.XLSX")

#Ler notas não entregues

df_nao_entregues = pd.read_excel(r"C:\Path\To\merc_nao_entregue.XLSX", skipfooter=1)

qtd = df_nao_entregues['Referência'].count()

if qtd >0:

    # remove tudo a partir do hífen (inclusive) e depois retira zeros à esquerda
    df_nao_entregues['Ref_limpa'] = (
        df_nao_entregues['Referência']
        .str.replace(r'-.*', '', regex=True)   # tudo após o '-' some
        .str.lstrip('0')                        # remove zeros à esquerda
    )

    notas = list(df_nao_entregues['Ref_limpa'].unique())
    in_clause = ",".join(f"'{n}'" for n in notas)

    #Conferir notas

    import pyodbc
    
    # Conectar ao banco de dados
    conn = pyodbc.connect('Driver=SQL Server;'
                        'Server=00.0.00.000;'
                        'Database=PRD;'
                        'Trusted_Connection=yes;')
    
    # Criar um objeto cursor
    cursor = conn.cursor()
    
    # Comando SQL a executar
    query = f"""
        SELECT * FROM MV_PBI_DATASLOGISTICAS
        WHERE NF IN ({in_clause})
    """
    
    # Executar a query e armazenar os resultados em um DataFrame
    df_notas_logs = pd.read_sql(query, conn)
    
    # Fechar a conexão
    conn.close()

    df_notas_logs['referencia'] = df_notas_logs['NF'].astype(str).apply(lambda x: ("0" * (9-len(x))) + x)
    df_notas_logs['referencia'] = df_notas_logs['referencia'] + '-' + df_notas_logs['SERIE'].astype(str)

    #Selecionar Colunas
    colunas = ['Referência','FrmPgto','Atribuição']
    df_nao_entregues = df_nao_entregues[colunas]
    colunas = ['referencia','PREV LOGIST','DT ENT']
    df_notas_logs = df_notas_logs[colunas]

    #Juntar Bases
    df_merge = pd.merge(df_nao_entregues, df_notas_logs, how= 'left', left_on= 'Referência' , right_on= 'referencia')
    df_merge.to_excel(r"C:\Path\To\Notas Fiscais.xlsx") 
    
    #Verificar se não localizou alguma NF
    df_merge['PREV LOGIST'] = df_merge['PREV LOGIST'].fillna('Nao localizada')
    df_nao_localizada = df_merge[df_merge['PREV LOGIST']=='Nao localizada']
    qtd_nao_localizada = df_nao_localizada['Referência'].count()

    if qtd_nao_localizada > 0:
        df_nao_localizada.to_excel(r"C:\Path\To\Notas Não Localizadas.xlsx")
    else:
        print('todas Notas encontradas')

    #Verificar se alguma Nota Fiscal foi entregue

    df_merge['DT ENT'] = df_merge['DT ENT'].str.replace(' ','Sem Entrega')
    df_merge['DT ENT'] = df_merge['DT ENT'].fillna('Sem Entrega')
    df_nfs_entregues = df_merge[df_merge['DT ENT']!= 'Sem Entrega']
    qtd_nfs_entregues = df_nfs_entregues['Referência'].count()

    if qtd_nfs_entregues > 0:
        df_nfs_entregues.to_excel(r"C:\Path\To\Nfs entregues.xlsx")
        print('Executar ajuste no SAP')
        nfs = list(df_nfs_entregues['Referência'].unique())

        for nf in nfs:
            session.findById("wnd[0]/usr/lbl[51,8]").setFocus()
            session.findById("wnd[0]").sendVKey(2)
            session.findById("wnd[0]/tbar[1]/btn[38]").press()
            session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN002-LOW").text = nf
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            session.findById("wnd[0]/usr/chk[1,10]").selected = True
            session.findById("wnd[0]/tbar[1]/btn[44]").press()
            session.findById("wnd[0]/usr/txtBSEG-ZUONR").text = ""
            session.findById("wnd[0]/tbar[0]/btn[11]").press()


    else:
        print('nenhuma nota entregue')     

      

else:
    print('sem notas')


 
# Nome do processo do SAP Logon 740
sap_process = "saplogon.exe"
 
# Percorre todos os processos em execução
for process in psutil.process_iter(attrs=['pid', 'name']):
    if process.info['name'].lower() == sap_process:
        print(f"Fechando {sap_process} (PID: {process.info['pid']})")
        psutil.Process(process.info['pid']).terminate()

print('concluido')



