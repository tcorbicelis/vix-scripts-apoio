import win32com.client as win32
import os
import pyautogui as pa
import time
from datetime import date
from dateutil.relativedelta import relativedelta
from pathlib import Path

usuario = os.environ["USERNAME"]
caminho = rf"C:\Users\{usuario}\OneDrive - PETROBRAS\Área de Trabalho\importação"
vl06g = os.path.join(caminho, "vl06g.xlsx")
relatrem = os.path.join(caminho, "relatrem.xlsx")
mb52 = os.path.join(caminho, "mb52.xlsx")
depositos = os.path.join(caminho, "depositos.xlsx")

# ==========================================================
# DELETAR ARQUIVOS ANTIGOS
# ==========================================================

arquivos_antigos = [
    vl06g,
    relatrem,
    mb52
]

for arquivo in arquivos_antigos:

    if os.path.exists(arquivo):

        try:
            os.remove(arquivo)

            print(
                f"Arquivo antigo deletado: "
                f"{os.path.basename(arquivo)}"
            )

        except PermissionError:
            raise Exception(
                f"Não foi possível deletar "
                f"{os.path.basename(arquivo)}. "
                f"Verifique se o arquivo está aberto no Excel."
            )

print("==========================================")
print("ARQUIVOS ANTIGOS REMOVIDOS")
print("Iniciando exportação pelo SAP...")
print("==========================================")

# =========================================================
# ==========================================================
# ======================================================

sapguiauto = win32.GetObject("SAPGUI")
application = sapguiauto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)

print(type(session))

# exportação vl06g

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "vl06g"
session.findById("wnd[0]").sendVKey (0)
session.findById("wnd[0]/usr/ctxtIT_WADAT-LOW").text = "01.01.2012"
session.findById("wnd[0]/usr/ctxtIT_WADAT-HIGH").text = ""
session.findById("wnd[0]/usr/ctxtIT_WADAT-LOW").setFocus
session.findById("wnd[0]/usr/ctxtIT_WADAT-LOW").caretPosition = 10
session.findById("wnd[0]/usr/ctxtIT_WADAT-LOW").showContextMenu
session.findById("wnd[0]/usr").selectContextMenuItem ("&018")
session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell (1,"TEXT")
session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "1"
session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[0]/tbar[1]/btn[8]").press()
session.findById("wnd[0]/tbar[1]/btn[18]").press()
session.findById("wnd[0]/tbar[1]/btn[33]").press()
session.findById("wnd[1]/usr").verticalScrollbar.position = 18
session.findById("wnd[1]/usr/lbl[1,12]").setFocus
session.findById("wnd[1]/usr/lbl[1,12]").caretPosition = 10
session.findById("wnd[1]").sendVKey (2)
session.findById("wnd[0]/mbar/menu[0]/menu[4]/menu[2]").select()
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[1]/usr/ctxtDY_PATH").text = caminho
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "vl06g.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
session.findById("wnd[1]/tbar[0]/btn[0]").press()

# ==========================
# ESPERAR ARQUIVO EXISTIR
# ==========================

print("Esperando arquivo...")

for tentativa in range(20):
    if os.path.exists(vl06g):
        print("Arquivo encontrado!")
        break

    time.sleep(1)

else:
    raise Exception(f"Arquivo não encontrado: {vl06g}")


# ==========================
# ABRIR VL06G
# ==========================

print("Abrindo VL06G...")

excel = win32.Dispatch("Excel.Application")
excel.Visible = True
excel.DisplayAlerts = False

wb_vl06g = excel.Workbooks.Open(str(Path(vl06g).resolve()))

# Aguarda o Excel terminar de abrir
time.sleep(3)

print(f"Arquivo aberto: {wb_vl06g.Name}")


# ==========================
# PEGAR A PLANILHA
# ==========================

ws_vl06g = wb_vl06g.Worksheets(1)

print(f"Planilha selecionada: {ws_vl06g.Name}")


# ==========================
# TRATAR VL06G
# ==========================

print("Tratando VL06G...")

# Excluir linha 3 primeiro
ws_vl06g.Rows(3).Delete()

# Depois excluir linha 1
ws_vl06g.Rows(1).Delete()

print("Linhas 3 e 1 excluídas com sucesso!")


# ==========================
# COPIAR NOVA COLUNA A
# ==========================

ultima_linha = ws_vl06g.Cells(
    ws_vl06g.Rows.Count, 1
).End(-4162).Row   # xlUp

print(f"Última linha encontrada: {ultima_linha}")

if ultima_linha >= 2:

    ws_vl06g.Range(
        f"A2:A{ultima_linha}"
    ).Copy()

    print(f"Coluna A copiada: A2:A{ultima_linha}")

else:
    print("Não existem dados suficientes na coluna A.")

# ==========================
# ABRIR YSRELATREM
# ==========================

session.findById("wnd[0]/tbar[0]/btn[3]").press()
time.sleep(0.5)

session.findById("wnd[0]/tbar[0]/btn[3]").press()
time.sleep(0.5)

session.findById("wnd[0]/tbar[0]/okcd").text = "ysrelatrem"
session.findById("wnd[0]").sendVKey(0)

time.sleep(1)

# ==========================
# PREENCHER FILTROS
# ==========================

session.findById(
    "wnd[0]/usr/ctxtS_VSTEL-LOW"
).text = "**"

# Abrir seleção múltipla de remessas
session.findById(
    "wnd[0]/usr/btn%S_LVBELN%APP%-VALU_PUSH"
).press()

time.sleep(0.5)

# Colar valores
session.findById("wnd[1]/tbar[0]/btn[24]").press()

# Confirmar
session.findById("wnd[1]/tbar[0]/btn[8]").press()

time.sleep(0.5)

# ==========================
# EXECUTAR RELATÓRIO
# ==========================

session.findById("wnd[0]/tbar[1]/btn[8]").press()

time.sleep(2)

# ==========================
# SELECIONAR LAYOUT
# ==========================

session.findById("wnd[0]/tbar[1]/btn[33]").press()

time.sleep(1)

layout = session.findById(
    "wnd[1]/usr/"
    "subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/"
    "cntlD500_CONTAINER/shellcont/shell"
)

layout.currentCellRow = 23
layout.firstVisibleRow = 11
layout.selectedRows = "23"

layout.clickCurrentCell()

time.sleep(2)

print("\n==============================")
print("LAYOUT SELECIONADO")
print("==============================")
print("Janela ativa:", session.ActiveWindow.Name)
print("Título:", session.ActiveWindow.Text)
print("Quantidade de janelas:", session.Children.Count)

for i in range(session.Children.Count):
    try:
        wnd = session.Children(i)
        print(f"wnd[{i}] = {wnd.Text}")
    except:
        pass

# ==========================
# EXPORTAR
# ==========================

session.findById("wnd[0]").maximize
session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").select()
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press()

session.findById("wnd[1]/usr/ctxtDY_PATH").text = caminho
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "relatrem.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 12
session.findById("wnd[1]/tbar[0]/btn[0]").press()

# RETORNAR PARA O MENU

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/btn[3]").press()
session.findById("wnd[0]/tbar[0]/btn[3]").press()

# ABRIR MB52

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "mb52"
session.findById("wnd[0]").sendVKey (0)

session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/ctxtP_VARI").text = "//F7A5"
session.findById("wnd[0]/usr/ctxtP_VARI").caretPosition = 6
session.findById("wnd[0]").sendVKey (0)

# COPIAR ITENS DA VL06G E COLAR NA MB52

ultima_linha_e = ws_vl06g.Cells(ws_vl06g.Rows.Count, 5).End(-4162).Row

ws_vl06g.Range(f"E2:E{ultima_linha_e}").Copy()

print(f"Coluna E copiada: E2:E{ultima_linha_e}")

session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/btn%MATNR%APP%-VALU_PUSH").press()
session.findById("wnd[1]/tbar[0]/btn[24]").press()
session.findById("wnd[1]/tbar[0]/btn[8]").press()

ultima_linha_e = ws_vl06g.Cells(ws_vl06g.Rows.Count, 11).End(-4162).Row

ws_vl06g.Range(f"K2:K{ultima_linha_e}").Copy()

print(f"Coluna K copiada: K2:K{ultima_linha_e}")

session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/btn%WERKS%APP%-VALU_PUSH").press()
session.findById("wnd[1]/tbar[0]/btn[24]").press()
session.findById("wnd[1]/tbar[0]/btn[8]").press()

# ==========================
# ABRIR PLANILHA DEPÓSITOS
# ==========================

print("Abrindo depósitos...")
print(depositos)

if not os.path.exists(depositos):
    raise Exception(
        f"Arquivo depósitos não encontrado: {depositos}"
    )

# Usar a mesma instância do Excel que já está sendo usada pela VL06G
print("Abrindo depósitos na instância atual do Excel...")

wb_depositos = None

# Verificar se o arquivo já está aberto
for wb_aberto in excel.Workbooks:

    try:
        if os.path.abspath(wb_aberto.FullName).lower() == \
           os.path.abspath(depositos).lower():

            wb_depositos = wb_aberto
            print("Depósitos já estava aberto!")
            break

    except Exception:
        pass


# Se não estiver aberto, abrir
if wb_depositos is None:

    print("Depósitos não estava aberto. Abrindo agora...")

    wb_depositos = excel.Workbooks.Open(depositos)

    time.sleep(2)

    print("Depósitos aberto com sucesso!")


# ==========================
# PEGAR A PLANILHA
# ==========================

ws_depositos = wb_depositos.Worksheets(1)

print("Arquivo conectado:", wb_depositos.Name)
print("Planilha conectada:", ws_depositos.Name)

# ==========================
# COPIAR DEPÓSITOS
# ==========================

ultima_linha_depositos = ws_depositos.Cells(
    ws_depositos.Rows.Count, 1
).End(-4162).Row

ws_depositos.Range(
    f"A1:A{ultima_linha_depositos}"
).Copy()

print(
    f"Depósitos - Coluna A copiada: "
    f"A1:A{ultima_linha_depositos}"
)

# COLAR E RELÓGINHO

session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/btn%LGORT%APP%-VALU_PUSH").press()
session.findById("wnd[1]/tbar[0]/btn[24]").press()
session.findById("wnd[1]/tbar[0]/btn[8]").press()
session.findById("wnd[0]/tbar[1]/btn[8]").press()

session.findById("wnd[0]").maximize
session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[1]/usr/ctxtDY_PATH").text = caminho
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "mb52.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]/tbar[0]/btn[0]").press()

# ==========================
# FECHAR DEPÓSITOS
# ==========================

wb_depositos.Close(SaveChanges=False)

print("Depósitos fechado.")

# ==========================================================
# FECHAR VL06G
# ==========================================================

wb_vl06g.Close(SaveChanges=False)

print("VL06G fechado.")

# ==========================================================
# FIM DA ETAPA 1 - EXPORTAÇÃO SAP
# INÍCIO DA ETAPA 2 - TRATAMENTO DAS PLANILHAS
# ==========================================================

print("==========================================")
print("EXPORTAÇÃO SAP CONCLUÍDA!")
print("Iniciando tratamento das planilhas...")
print("==========================================")

# ============================================================
# ============================================================
# ============================================================

import win32com.client as win32
import os
import pyautogui as pa
import time

usuario = os.environ["USERNAME"]
caminho = rf"C:\Users\{usuario}\OneDrive - PETROBRAS\Área de Trabalho\importação"

vl06g = os.path.join(caminho, "vl06g.xlsx")
relatrem = os.path.join(caminho, "relatrem.xlsx")
mb52 = os.path.join(caminho, "mb52.xlsx")
depositos = os.path.join(caminho, "depositos.xlsx")


# Abre o Excel
excel = win32.Dispatch("Excel.Application")

# Deixa o Excel visível
excel.Visible = True

# Abre as 3 planilhas
wb_vl06g = excel.Workbooks.Open(vl06g)
wb_relatrem = excel.Workbooks.Open(relatrem)
wb_mb52 = excel.Workbooks.Open(mb52)

print("As três planilhas foram abertas com sucesso!")

# ==========================================================
# TRATAMENTO DA VL06G
# ==========================================================

ws_vl06g = wb_vl06g.Worksheets(1)

# Deleta a linha 3
ws_vl06g.Rows(3).Delete()

# Deleta a linha 1
ws_vl06g.Rows(1).Delete()

# ==========================================================
# TRATAMENTO DA MB52
# ==========================================================

ws_mb52 = wb_mb52.Worksheets(1)

# Deleta a linha 3
ws_mb52.Rows(3).Delete()

# Deleta a linha 1
ws_mb52.Rows(1).Delete()


print("Tratamento da VL06G e MB52 concluído!")

# ==========================================================
# TRATAMENTO DA RELATREM
# ==========================================================

ws_relatrem = wb_relatrem.Worksheets(1)

# Deleta a linha 5 primeiro
ws_relatrem.Rows(6).Delete()

# Depois deleta as linhas 1, 2 e 3
ws_relatrem.Rows(1).Delete()
ws_relatrem.Rows(1).Delete()
ws_relatrem.Rows(1).Delete()
ws_relatrem.Rows(1).Delete()

print("Tratamento da RELATREM concluído!")

# ==========================================================
# TRATAMENTO DA MB52
# ==========================================================

# PREENCHER L2 E M2 DA MB52
ws_mb52.Range("L1").Value = "S"
ws_mb52.Range("M1").Value = "S"

# ATIVAR FILTRO NA MB52
ultima_coluna = ws_mb52.Cells(1, ws_mb52.Columns.Count).End(-4159).Column

ws_mb52.Range(
    ws_mb52.Cells(1, 1),
    ws_mb52.Cells(1, ultima_coluna)
).AutoFilter()

# REMOVER "RP" DA COLUNA F DA MB52
ws_mb52.Columns("F").Replace(
    What="RP",
    Replacement="",
    LookAt=2
)

# FILTRAR O VALOR 0 EXATO NA COLUNA I E EXCLUIR AS LINHAS
ultima_linha = ws_mb52.Cells(ws_mb52.Rows.Count, "A").End(-4162).Row

ws_mb52.Range(
    ws_mb52.Cells(1, 1),
    ws_mb52.Cells(ultima_linha, ultima_coluna)
).AutoFilter(
    Field=9,
    Criteria1="=0"
)

# EXCLUIR AS LINHAS VISÍVEIS DO FILTRO, SEM EXCLUIR O CABEÇALHO
try:
    ws_mb52.Range(
        "A2:A" + str(ultima_linha)
    ).SpecialCells(12).EntireRow.Delete()
except:
    pass

# REMOVER O FILTRO
if ws_mb52.AutoFilterMode:
    ws_mb52.ShowAllData()

# CONCATENAR C + A + F EM TODA A COLUNA L DA MB52
ultima_linha = ws_mb52.Cells(ws_mb52.Rows.Count, "A").End(-4162).Row

ws_mb52.Range("L2").Formula = "=C2&A2&F2"

ws_mb52.Range("L2:L" + str(ultima_linha)).FillDown()

# IGUALAR M2 À B2
ultima_linha = ws_mb52.Cells(ws_mb52.Rows.Count, "A").End(-4162).Row

ws_mb52.Range("M2").Formula = "=B2"
ws_mb52.Range("M2:M" + str(ultima_linha)).FillDown()

# ==========================================================
# TRATAMENTO DA VL06G
# ==========================================================

# PREENCHER O1, P1 E Q1 DA VL06G
ws_vl06g.Range("O1").Value = "S"
ws_vl06g.Range("P1").Value = "S"
ws_vl06g.Range("Q1").Value = "S"


# CONCATENAR E + K + M EM TODA A COLUNA O DA VL06G
ultima_linha = ws_vl06g.Cells(ws_vl06g.Rows.Count, "A").End(-4162).Row

ws_vl06g.Range("O2").Formula = "=E2&K2&M2"
ws_vl06g.Range("O2:O" + str(ultima_linha)).FillDown()

# PROCV EM TODA A COLUNA P DA VL06G
ultima_linha = ws_vl06g.Cells(
    ws_vl06g.Rows.Count, "A"
).End(-4162).Row

nome_aba_mb52 = ws_mb52.Name

ws_vl06g.Range("P2").Formula = (
    f"=VLOOKUP(O2,'[mb52.xlsx]{nome_aba_mb52}'!$L:$M,2,FALSE)"
)

ws_vl06g.Range(
    "P2:P" + str(ultima_linha)
).FillDown()

# ==========================================================
# FILTRAR E DELETAR LINHAS COM #N/D NA COLUNA P DA VL06G
# ==========================================================

# Última linha da tabela
ultima_linha = ws_vl06g.Cells(
    ws_vl06g.Rows.Count, "P"
).End(-4162).Row

# Última coluna utilizada
ultima_coluna = ws_vl06g.Cells(
    1, ws_vl06g.Columns.Count
).End(-4159).Column

# Ativa o filtro
ws_vl06g.Range(
    ws_vl06g.Cells(1, 1),
    ws_vl06g.Cells(ultima_linha, ultima_coluna)
).AutoFilter(
    Field=16,          # Coluna P
    Criteria1="#N/D"
)

# Seleciona somente as linhas visíveis após o cabeçalho
try:
    linhas_filtradas = ws_vl06g.Range(
        f"A2:A{ultima_linha}"
    ).SpecialCells(12)  # xlCellTypeVisible

    linhas_filtradas.EntireRow.Delete()

except:
    # Caso não existam #N/D
    pass

# Remove o filtro
if ws_vl06g.AutoFilterMode:
    ws_vl06g.AutoFilterMode = False

# ==========================================================
# ATIVAR FILTRO NA VL06G
# ==========================================================

ultima_linha = ws_vl06g.Cells(
    ws_vl06g.Rows.Count, "A"
).End(-4162).Row

ultima_coluna = ws_vl06g.Cells(
    1, ws_vl06g.Columns.Count
).End(-4159).Column

ws_vl06g.Range(
    ws_vl06g.Cells(1, 1),
    ws_vl06g.Cells(ultima_linha, ultima_coluna)
).AutoFilter()

# IGUALAR C2 À P2 NA VL06G
ultima_linha = ws_vl06g.Cells(
    ws_vl06g.Rows.Count, "P"
).End(-4162).Row

ws_vl06g.Range("C2").Formula = "=P2"

ws_vl06g.Range(
    "C2:C" + str(ultima_linha)
).FillDown()

# ==========================================================
# PROCV EM TODA A COLUNA Q DA VL06G
# ==========================================================

ultima_linha = ws_vl06g.Cells(
    ws_vl06g.Rows.Count, "A"
).End(-4162).Row

nome_aba_relatrem = ws_relatrem.Name

ws_vl06g.Range("Q2").Formula = (
    f"=VLOOKUP(A2,'[relatrem.xlsx]{nome_aba_relatrem}'!$A:$H,8,FALSE)"
)

ws_vl06g.Range(
    "Q2:Q" + str(ultima_linha)
).FillDown()

# ==========================================================
# FILTRAR NÚMEROS PUROS DIFERENTES DE ZERO NA COLUNA Q
# E DEPOIS FILTRAR MACAE NA COLUNA H
# ==========================================================

ultima_linha = ws_vl06g.Cells(
    ws_vl06g.Rows.Count, "Q"
).End(-4162).Row


# ==========================================================
# LIMPAR FILTRO ANTERIOR
# ==========================================================

try:
    if ws_vl06g.FilterMode:
        ws_vl06g.ShowAllData()
except:
    pass

try:
    ws_vl06g.AutoFilterMode = False
except:
    pass


# ==========================================================
# COLUNA AUXILIAR R
# ==========================================================

ws_vl06g.Range("R1").Value = "AUXILIAR"


# Número puro diferente de zero = SIM
# Zero = NAO
# Texto = NAO

ws_vl06g.Range("R2").Formula = (
    '=IFERROR('
    'IF(AND(ISNUMBER(Q2),Q2<>0),"SIM","NAO"),'
    '"NAO")'
)


# Preencher até a última linha
ws_vl06g.Range(
    f"R2:R{ultima_linha}"
).FillDown()


# Calcular
excel.Calculate()


# ==========================================================
# PRIMEIRO FILTRO:
# COLUNA R = SIM
# ==========================================================

ws_vl06g.Range(
    f"A1:R{ultima_linha}"
).AutoFilter(
    Field=18,
    Criteria1="SIM"
)


# ==========================================================
# SEGUNDO FILTRO:
# COLUNA H = MACAE
#
# IMPORTANTE:
# O filtro anterior da coluna R CONTINUA ATIVO.
# ==========================================================

ws_vl06g.Range(
    f"A1:R{ultima_linha}"
).AutoFilter(
    Field=8,
    Criteria1="MACAE"
)


# ==========================================================
# DELETAR AS LINHAS QUE PASSARAM NOS DOIS FILTROS
# ==========================================================

try:

    linhas_visiveis = ws_vl06g.Range(
        f"A2:A{ultima_linha}"
    ).SpecialCells(12)  # xlCellTypeVisible

    linhas_visiveis.EntireRow.Delete()

except:
    pass


# ==========================================================
# REMOVER A COLUNA AUXILIAR R
# ==========================================================

ws_vl06g.Columns("R").Delete()


# ==========================================================
# LIMPAR O FILTRO
# ==========================================================

try:
    if ws_vl06g.FilterMode:
        ws_vl06g.ShowAllData()
except:
    pass

try:
    ws_vl06g.AutoFilterMode = False
except:
    pass

# ==========================================================
# FILTRAR A COLUNA Q
# MANTER TUDO QUE NÃO SEJA NÚMERO PURO
# ==========================================================

ultima_linha = ws_vl06g.Cells(
    ws_vl06g.Rows.Count, "Q"
).End(-4162).Row


# ==========================================================
# COLUNA AUXILIAR R
# ==========================================================

ws_vl06g.Range("R1").Value = "AUX"

# Se Q for texto → SIM
# Se Q for número → NAO
# Se Q estiver vazio → NAO

ws_vl06g.Range("R2").Formula = (
    '=IF(AND(Q2<>"",ISTEXT(Q2)),"SIM","NAO")'
)

ws_vl06g.Range(
    f"R2:R{ultima_linha}"
).FillDown()


# ==========================================================
# CALCULAR AS FÓRMULAS
# ==========================================================

excel.Calculate()


# ==========================================================
# LIMPAR FILTRO ANTERIOR
# ==========================================================

if ws_vl06g.AutoFilterMode:
    try:
        ws_vl06g.ShowAllData()
    except:
        pass

    ws_vl06g.AutoFilterMode = False


# ==========================================================
# APLICAR FILTRO
# ==========================================================

ws_vl06g.Range(
    f"A1:R{ultima_linha}"
).AutoFilter(
    Field=18,
    Criteria1="SIM"
)

# ==========================================================
# IGUALAR A COLUNA H À COLUNA Q NA VL06G
# ==========================================================

ultima_linha = ws_vl06g.Cells(
    ws_vl06g.Rows.Count, "Q"
).End(-4162).Row

ws_vl06g.Range("H2").Formula = "=Q2"

ws_vl06g.Range(
    "H2:H" + str(ultima_linha)
).FillDown()

# ==========================================================
# RESETAR FILTRO DA VL06G
# ==========================================================

try:
    if ws_vl06g.FilterMode:
        ws_vl06g.ShowAllData()
except:
    pass

try:
    ws_vl06g.AutoFilterMode = False
except:
    pass


# ==========================================================
# FILTRO NA COLUNA H DA VL06G
#
# DELETAR:
# RAM1
# NAM1
# NAM2
# MOM0
# JU17
# HP18
# HM09
# A098
# GMP1
# todos os números puros
# todos os valores que terminam em M0
# EXCETO 26M0
#
# NÃO DELETAR:
# valores que começam com NS
# 26M0
# demais valores
# ==========================================================

ultima_linha = ws_vl06g.Cells(
    ws_vl06g.Rows.Count, "H"
).End(-4162).Row


# ==========================================================
# COLUNA AUXILIAR R
# ==========================================================

ws_vl06g.Range("R1").Value = "AUX"


# ==========================================================
# IDENTIFICAR O QUE DEVE SER DELETADO
# ==========================================================

ws_vl06g.Range("R2").Formula = (
    '=IF('
    'OR('
        # Códigos específicos
        'H2="RAM1",'
        'H2="NAM1",'
        'H2="NAM2",'
        'H2="MOM0",'
        'H2="JU17",'
        'H2="HP18",'
        'H2="HM09",'
        'H2="A098",'
        'H2="GPM1",'
        'H2="55ME",'

        # Todos os números puros
        'ISNUMBER(H2),'

        # Termina em M0, mas NÃO é 26M0
        'RIGHT(H2,2)="M0"'
    '),'
    '"SIM",'
    '"NAO"'
    ')'
)


# ==========================================================
# PREENCHER A FÓRMULA ATÉ O FINAL
# ==========================================================

ws_vl06g.Range(
    f"R2:R{ultima_linha}"
).FillDown()


# ==========================================================
# CALCULAR
# ==========================================================

excel.Calculate()


# ==========================================================
# LIMPAR FILTRO ANTERIOR
# ==========================================================

try:
    if ws_vl06g.FilterMode:
        ws_vl06g.ShowAllData()
except:
    pass

try:
    ws_vl06g.AutoFilterMode = False
except:
    pass


# ==========================================================
# FILTRAR AS LINHAS QUE DEVEM SER DELETADAS
# ==========================================================

ws_vl06g.Range(
    f"A1:R{ultima_linha}"
).AutoFilter(
    Field=18,
    Criteria1="SIM"
)


# ==========================================================
# DELETAR AS LINHAS FILTRADAS
# ==========================================================

try:

    linhas_visiveis = ws_vl06g.Range(
        f"A2:A{ultima_linha}"
    ).SpecialCells(12)  # xlCellTypeVisible

    linhas_visiveis.EntireRow.Delete()

except:
    pass


# ==========================================================
# REMOVER COLUNA AUXILIAR R
# ==========================================================

ws_vl06g.Columns("R").Delete()


# ==========================================================
# LIMPAR FILTRO
# ==========================================================

try:
    if ws_vl06g.FilterMode:
        ws_vl06g.ShowAllData()
except:
    pass

try:
    ws_vl06g.AutoFilterMode = False
except:
    pass


# ==========================================================
# CONVERTER A:N EM VALORES NA VL06G
# ==========================================================

ultima_linha = ws_vl06g.Cells(
    ws_vl06g.Rows.Count, "A"
).End(-4162).Row

intervalo = ws_vl06g.Range(
    f"A2:N{ultima_linha}"
)

# Copia e cola somente os valores
intervalo.Value = intervalo.Value

# ==========================================================
# DELETAR LINHAS QUE CONTÊM #N/D
# ==========================================================

ultima_linha = ws_vl06g.Cells(
    ws_vl06g.Rows.Count, "A"
).End(-4162).Row

ultima_coluna = ws_vl06g.Cells(
    1, ws_vl06g.Columns.Count
).End(-4159).Column

for linha in range(ultima_linha, 1, -1):

    valores = ws_vl06g.Range(
        ws_vl06g.Cells(linha, 1),
        ws_vl06g.Cells(linha, ultima_coluna)
    ).Value

    if valores and valores[0]:

        encontrou_nd = any(
            valor is not None and str(valor).upper() in ("#N/D", "#N/A")
            for valor in valores[0]
        )

        if encontrou_nd:
            ws_vl06g.Rows(linha).Delete()

# ==========================================================
# CRIAR A PLANILHA REMESSAS
# ==========================================================

remessas = os.path.join(caminho, "remessas.xlsx")

# Se já existir uma remessas antiga, excluir
if os.path.exists(remessas):
    try:
        os.remove(remessas)
        print("Remessas antiga excluída.")
    except PermissionError:
        raise Exception(
            "Não foi possível excluir remessas.xlsx. "
            "Verifique se ela está aberta no Excel."
        )


# ==========================================================
# CRIAR NOVA PASTA DE TRABALHO
# ==========================================================

wb_remessas = excel.Workbooks.Add()

ws_remessas = wb_remessas.Worksheets(1)

ws_remessas.Name = "Remessas"


# ==========================================================
# CABEÇALHOS
# ==========================================================

ws_remessas.Range("B1").Value = "Fornmto"
ws_remessas.Range("C1").Value = "SaídaMerc"
ws_remessas.Range("D1").Value = "Dep."
ws_remessas.Range("E1").Value = "Item"
ws_remessas.Range("F1").Value = "Material"
ws_remessas.Range("G1").Value = "Denominação"
ws_remessas.Range("H1").Value = "Doc.ref."
ws_remessas.Range("I1").Value = "Cidade recebedor ordem"
ws_remessas.Range("J1").Value = "Criado por"
ws_remessas.Range("K1").Value = "Qtd.remessa"
ws_remessas.Range("L1").Value = "Cen."
ws_remessas.Range("M1").Value = "Lote"
ws_remessas.Range("N1").Value = "Receb.merc"


# ==========================================================
# ÚLTIMA LINHA DA VL06G
# ==========================================================

ultima_linha = ws_vl06g.Cells(
    ws_vl06g.Rows.Count,
    "A"
).End(-4162).Row


# ==========================================================
# COPIAR VL06G A:N
# PARA REMESSAS B:O
# SOMENTE VALORES
# ==========================================================

if ultima_linha >= 2:

    origem = ws_vl06g.Range(
        f"A2:N{ultima_linha}"
    )

    destino = ws_remessas.Range(
        f"B2:O{ultima_linha}"
    )

    destino.Value = origem.Value

    print(
        f"Dados da VL06G copiados: "
        f"A2:N{ultima_linha} → B2:O{ultima_linha}"
    )

else:

    print("VL06G não possui dados para copiar.")


# ==========================================================
# RECORTAR COLUNA O E COLAR NA COLUNA N
# ==========================================================

if ultima_linha >= 2:

    ws_remessas.Range(
        f"O2:O{ultima_linha}"
    ).Cut(
        Destination=ws_remessas.Range(
            f"N2:N{ultima_linha}"
        )
    )

    print("Coluna O recortada e colada na coluna N.")
    
# ==========================================================
# GARANTIR QUE NÃO EXISTAM FÓRMULAS
# ==========================================================

if ultima_linha >= 2:

    area_dados = ws_remessas.Range(
        f"B2:N{ultima_linha}"
    )

    area_dados.Value = area_dados.Value


# ==========================================================
# REMOVER FILTROS
# ==========================================================

try:

    if ws_remessas.FilterMode:
        ws_remessas.ShowAllData()

except:

    pass


try:

    ws_remessas.AutoFilterMode = False

except:

    pass


# ==========================================================
# SALVAR REMESSAS
# ==========================================================

excel.DisplayAlerts = False

wb_remessas.SaveAs(
    remessas,
    FileFormat=51
)

excel.DisplayAlerts = True

print("==========================================")
print("REMESSAS CRIADA COM SUCESSO!")
print(f"Arquivo: {remessas}")
print("==========================================")

input("Pressione Enter para fechar...")