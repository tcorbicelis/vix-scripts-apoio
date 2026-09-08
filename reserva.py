import win32com.client as win32
import openpyxl
import os
import pyautogui as pa
import time
import subprocess
from datetime import date, timedelta


# ==========================================================
# CONFIGURAÇÕES
# ==========================================================

usuario = os.environ["USERNAME"]

caminho = rf"C:\Users\{usuario}\OneDrive - PETROBRAS\Área de Trabalho\Importacao"

# Data atual
hoje = date.today()

# Período da consulta
data_inicial = hoje - timedelta(days=60)
data_final = hoje + timedelta(days=30)

# Formato das datas para o SAP
data_inicial_sap = data_inicial.strftime("%d.%m.%Y")
data_final_sap = data_final.strftime("%d.%m.%Y")

# Nome do arquivo YSATENDRES
nome_arquivo = f"ysatendres-{hoje.strftime('%d.%m.%Y')}.xlsx"

# Caminho completo dos arquivos
ysatendres = os.path.join(caminho, nome_arquivo)
mb25 = os.path.join(caminho, "mb25.xlsx")


# ==========================================================
# DELETAR ARQUIVOS ANTIGOS
# ==========================================================

if os.path.exists(ysatendres):

    try:
        os.remove(ysatendres)

        print(
            f"Arquivo antigo deletado: "
            f"{os.path.basename(ysatendres)}"
        )

    except PermissionError:
        raise Exception(
            f"Não foi possível deletar "
            f"{os.path.basename(ysatendres)}. "
            f"Verifique se o arquivo está aberto no Excel."
        )


if os.path.exists(mb25):

    try:
        os.remove(mb25)

        print(
            f"Arquivo antigo deletado: "
            f"{os.path.basename(mb25)}"
        )

    except PermissionError:
        raise Exception(
            f"Não foi possível deletar "
            f"{os.path.basename(mb25)}. "
            f"Verifique se o arquivo está aberto no Excel."
        )


print("==========================================")
print("ARQUIVOS ANTIGOS REMOVIDOS")
print("Iniciando exportação pelo SAP...")
print("==========================================")


# ==========================================================
# CONEXÃO COM O SAP
# ==========================================================

sapguiauto = win32.GetObject("SAPGUI")
application = sapguiauto.GetScriptingEngine

connection = application.Children(0)
session = connection.Children(0)

print("Conectado ao SAP.")


# ==========================================================
# MAXIMIZAR SAP
# ==========================================================

session.findById("wnd[0]").maximize()

time.sleep(1)


# ==========================================================
# ACESSAR YSATENDRES
# ==========================================================

session.findById(
    "wnd[0]/tbar[0]/okcd"
).text = "/nysatendres"

session.findById(
    "wnd[0]"
).sendVKey(0)

time.sleep(2)


# ==========================================================
# ABRIR SELEÇÃO DE VARIANTE
# ==========================================================

session.findById(
    "wnd[0]/tbar[1]/btn[17]"
).press()

time.sleep(2)


# ==========================================================
# INFORMAR USUÁRIO DA VARIANTE
# ==========================================================

session.findById(
    "wnd[1]/usr/txtENAME-LOW"
).text = "cp1w"

time.sleep(0.5)


# Executar busca
session.findById(
    "wnd[1]/tbar[0]/btn[8]"
).press()

time.sleep(2)

# ==========================================================
# SELECIONAR VARIANTE
# ==========================================================

variant = session.findById(
    "wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell"
)

variant.selectedRows = "0"

time.sleep(0.5)

variant.doubleClickCurrentCell()

time.sleep(2)


# ==========================================================
# PREENCHER CAMPOS
# ==========================================================

# Data inicial
session.findById(
    "wnd[0]/usr/ctxtSD_BDTER-LOW"
).text = data_inicial_sap


# Data final
session.findById(
    "wnd[0]/usr/ctxtSD_BDTER-HIGH"
).text = data_final_sap


# Layout
session.findById(
    "wnd[0]/usr/ctxtALV_DEF"
).text = "/ARM-MACAE"

time.sleep(1)


# ==========================================================
# EXECUTAR YSATENDRES
# ==========================================================

session.findById(
    "wnd[0]/tbar[1]/btn[8]"
).press()

time.sleep(4)


# ==========================================================
# EXPORTAÇÃO YSATENDRES
# ==========================================================

print("Iniciando exportação YSATENDRES...")


# Lista -> Exportar -> Arquivo local
session.findById(
    "wnd[0]/mbar/menu[0]/menu[1]/menu[2]"
).select()

time.sleep(2)


# Selecionar "Arquivo local"
session.findById(
    "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/"
    "sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]"
).select()

time.sleep(0.5)


# Confirmar
session.findById(
    "wnd[1]/tbar[0]/btn[0]"
).press()

time.sleep(2)


# ==========================================================
# SALVAR YSATENDRES
# ==========================================================

session.findById(
    "wnd[1]/usr/ctxtDY_PATH"
).text = caminho

session.findById(
    "wnd[1]/usr/ctxtDY_FILENAME"
).text = nome_arquivo

time.sleep(0.5)

session.findById(
    "wnd[1]/tbar[0]/btn[0]"
).press()

time.sleep(3)


print("YSATENDRES exportado com sucesso.")


# ==========================================================
# VOLTAR PARA A TELA INICIAL
# ==========================================================

session.findById("wnd[0]").maximize()

session.findById(
    "wnd[0]/tbar[0]/btn[15]"
).press()

time.sleep(1)

session.findById(
    "wnd[0]/tbar[0]/btn[15]"
).press()

time.sleep(2)


# ==========================================================
# ABRIR MB25
# ==========================================================

print("Abrindo MB25...")

session.findById(
    "wnd[0]/tbar[0]/okcd"
).text = "mb25"

session.findById(
    "wnd[0]"
).sendVKey(0)

time.sleep(2)


# ==========================================================
# CONFIGURAR MB25
# ==========================================================

session.findById(
    "wnd[0]/usr/ctxtWERKS-LOW"
).text = ""

session.findById(
    "wnd[0]/usr/ctxtALV_DEF"
).text = "/ARM.MACAE"

session.findById(
    "wnd[0]/usr/ctxtALV_DEF"
).setFocus

session.findById(
    "wnd[0]/usr/ctxtALV_DEF"
).caretPosition = 10

session.findById(
    "wnd[0]"
).sendVKey(0)

time.sleep(2)

# ==========================================================
# ABRIR SELEÇÃO MÚLTIPLA
# ==========================================================

session.findById(
    "wnd[0]/tbar[1]/btn[16]"
).press()

time.sleep(1)


# ==========================================================
# LOCALIZAR CAMPO DE RESERVA
# ==========================================================

tree = session.findById(
    "wnd[0]/usr/"
    "ssub%SUBSCREEN%_SUB%_CONTAINER:SAPLSSEL:2001/"
    "ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/"
    "cntlSUB_CONTAINER/shellcont/shellcont/shell/"
    "shellcont[1]/shell"
)

tree.expandNode("          1")

tree.selectNode("          2")

tree.topNode = "          1"

tree.doubleClickNode("          2")

time.sleep(1)


# ==========================================================
# ABRIR SELEÇÃO MÚLTIPLA DAS RESERVAS
# ==========================================================

session.findById(
    "wnd[0]/usr/"
    "ssub%SUBSCREEN%_SUB%_CONTAINER:SAPLSSEL:2001/"
    "ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/"
    "ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/"
    "btn%%%DYN001%APP%-VALU_PUSH"
).press()

time.sleep(1)


# ==========================================================
# LER RESERVAS DA COLUNA B DO ysatendres.XLSX
# ==========================================================

print("Lendo reservas do arquivo ysatendres...")

if not os.path.exists(ysatendres):
    raise FileNotFoundError(
        f"O arquivo ysatendres não foi encontrado: {ysatendres}"
    )


excel = win32.Dispatch("Excel.Application")
excel.Visible = False

wb = excel.Workbooks.Open(ysatendres)
ws = wb.Worksheets(1)


# Última linha preenchida na coluna B
ultima_linha = ws.Cells(
    ws.Rows.Count,
    2
).End(-4162).Row


reservas = []

for linha in range(7, ultima_linha + 1):

    valor = ws.Cells(
        linha,
        2
    ).Value

    if valor is not None:

        # Caso o Excel entregue o número como float
        if isinstance(valor, float):
            valor = str(int(valor))

        else:
            valor = str(valor).strip()

        if valor:
            reservas.append(valor)


# Fechar Excel
wb.Close(False)
excel.Quit()


print(
    f"Total de reservas encontradas: "
    f"{len(reservas)}"
)


# ==========================================================
# VERIFICAR SE ENCONTROU RESERVAS
# ==========================================================

if not reservas:

    raise Exception(
        "Nenhuma reserva foi encontrada na coluna B "
        "do arquivo ysatendres.xlsx."
    )


# ==========================================================
# COPIAR RESERVAS PARA O CLIPBOARD DO WINDOWS
# ==========================================================

texto_reservas = "\r\n".join(reservas)


processo = subprocess.Popen(
    ["clip"],
    stdin=subprocess.PIPE,
    text=True
)

processo.communicate(texto_reservas)


print("Reservas copiadas para a área de transferência.")


# ==========================================================
# COLAR RESERVAS NO SAP
# ==========================================================

session.findById(
    "wnd[1]/tbar[0]/btn[24]"
).press()

time.sleep(1)


# ==========================================================
# CONFIRMAR SELEÇÃO DAS RESERVAS
# ==========================================================

session.findById(
    "wnd[1]/tbar[0]/btn[8]"
).press()

time.sleep(1)


print("Reservas inseridas no SAP com sucesso.")

#==========================================================
# EXECUTAR MB25
#==========================================================

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[1]/btn[8]").press()

#=========================================================
# SALVAR 
#=========================================================

session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[1]/usr/ctxtDY_PATH").text = caminho
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "mb25.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
session.findById("wnd[1]/tbar[0]/btn[0]").press()

time.sleep(1)

# ==========================================================
# TRATAMENTO FINAL DA PLANILHA MB25
# ==========================================================

print("==========================================")
print("Iniciando tratamento final da planilha MB25...")
print("==========================================")


if not os.path.exists(mb25):
    raise FileNotFoundError(
        f"O arquivo MB25 não foi encontrado: {mb25}"
    )


encontrou_qtd = False


# ==========================================================
# ABRIR O ARQUIVO DIRETAMENTE, SEM ABRIR O EXCEL
# ==========================================================

wb = openpyxl.load_workbook(mb25)


# ==========================================================
# PERCORRER TODAS AS PLANILHAS
# ==========================================================

for ws in wb.worksheets:

    print(f"Verificando planilha: {ws.title}")

    for row in ws.iter_rows():

        for celula in row:

            if celula.value is not None:

                if str(celula.value).strip() == "Qtd.necess.":

                    print(
                        f"Alterado em {ws.title} "
                        f"{celula.coordinate}: "
                        f"'Qtd.necess.' -> 'Qtdnecess.'"
                    )

                    celula.value = "Qtdnecess."

                    encontrou_qtd = True


# ==========================================================
# SALVAR ALTERAÇÃO
# ==========================================================

if encontrou_qtd:

    wb.save(mb25)

    print(
        "Planilha MB25 alterada e salva com sucesso."
    )

else:

    print(
        "Atenção: 'Qtd.necess.' não foi encontrado "
        "na planilha MB25."
    )


print("Tratamento final concluído.")