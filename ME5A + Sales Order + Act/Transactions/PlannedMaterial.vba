Option Explicit
Public SapGuiAuto, SAP_APP, Connection, Session, Wscript


Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Integer

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Sub PM(control As IRibbonControl)

'Declara variáveis do código
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("ME5A")
    Dim Arquivo
    ws.Activate 'planilha ativa (SAPGUIEXPO)
    Dim msgValue
    Dim lastRow As Long
    Dim PGR As String 'Variável do PGR
    Dim colIndex As Integer
    Dim headerCell As Range
    Dim copyrange As Range
    
'PGR
    ThisWorkbook.Sheets("PGR").Activate
    Range("A1").Select
    ActiveCell.Offset(2, 0).Select
    PGR = ActiveCell.Value
    
'Checar se arquivo "Sheet1" já existe e deletar caso exista


    Dim FPath1 As String            'Target output path
    Dim FName1 As String            'Target output filename
    Dim wsfile As String
    
    FPath1 = "C:\SAP"             'Diretório
    FName1 = "PlannedOrder.xlsx"   'Nome do arquivo que vai ser salvo
    wsfile = ThisWorkbook.Name
    
    Dim FNameLocation As String: FNameLocation = FPath1 & "\" & FName1 'Juntando nome do arquivo e diretório
    
    If Len(Dir(FNameLocation)) <> 0 Then
        SetAttr FNameLocation, vbNormal
        Kill FNameLocation
    End If
        
    '-------------------------------
    ' Download Open Order Report from SAP
    
    '****************************
    '* Open connection with SAP *
    '****************************
    If Not IsObject(SAP_APP) Then
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SAP_APP = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(Connection) Then
    Set Connection = SAP_APP.Children(0)
    End If
    If Not IsObject(Session) Then
    Set Session = Connection.Children(0)
    End If
    If IsObject(Wscript) Then
    Wscript.ConnectObject Session, "on"
    Wscript.ConnectObject SAP_APP, "on"
    End If
    
    ' Ativar a planilha ME5A
    ws.Activate

    ' Encontrar a coluna "Material"
    colIndex = 0 ' Resetar o valor
    For Each headerCell In ws.Rows(3).Cells
        If Trim(headerCell.Value) = "Material" Then
            colIndex = headerCell.Column
            Exit For
        End If
    Next headerCell

    ' Encontrar a última linha preenchida na coluna identificada
    lastRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row

    ' Definir o intervalo para cópia (da linha 2 até a última linha preenchida)
    Set copyrange = ws.Range(ws.Cells(2, colIndex), ws.Cells(lastRow, colIndex))

    ' Copiar os valores da coluna "Material"
    copyrange.copy
    'Run T-Code
    Session.findById("wnd[0]/tbar[0]/okcd").Text = "/NZPTP_MPLN"
    Session.findById("wnd[0]").sendVKey 0
    Session.findById("wnd[0]").maximize
    Session.findById("wnd[0]/usr/ctxtS_PLWRK-LOW").Text = "A712"
    Session.findById("wnd[0]/usr/ctxtS_DISPO-LOW").Text = "*"
    Session.findById("wnd[0]/usr/ctxtS_EKGRP-LOW").SetFocus
    Session.findById("wnd[0]/usr/ctxtS_EKGRP-LOW").caretPosition = 0
    Session.findById("wnd[0]/usr/ctxtS_EKGRP-LOW").Text = PGR
    Session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").SetFocus
    Session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").caretPosition = 0
    Session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press
    Session.findById("wnd[1]/tbar[0]/btn[24]").press
    Session.findById("wnd[1]/tbar[0]/btn[0]").press
    Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    '*** Run Report/Execute/executa***
    Session.findById("wnd[0]/tbar[1]/btn[8]").press

    
    '*** Setup Report Layout ***
    Session.findById("wnd[0]/tbar[1]/btn[33]").press        'escolhe standard layout
    Session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellRow = -1
    Session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectColumn "VARIANT"
    Session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").contextMenu
    Session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectContextMenuItem "&FIND"
    Session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").Text = "/ZPTP_CTRL"                                                            'insere nome do layout escolhido
    Session.findById("wnd[2]/tbar[0]/btn[0]").press                                                                             'Execute Find
    Session.findById("wnd[2]/tbar[0]/btn[12]").press                                                                            'Close Find Window
    Session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell  'Hit Enter -> Use found layout

    Session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select
    Session.findById("wnd[1]/usr/ctxtDY_PATH").Text = FPath1
    Session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "PlannedOrder.xlsx"

    Session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
    Session.findById("wnd[1]/tbar[0]/btn[0]").press
    
    'Go back to SAP main page
    Session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    Session.findById("wnd[0]").sendVKey 0
    
    Dim sourceWorkbook As Workbook
    Dim destinationWorkbook As Workbook
    Dim sourceRange As Range 'Planned Open Date
    Dim PGRR As Range
    'Dim lastRow As Long
    Dim PlannedOrderRg As String
    Dim PlannedDate As Range
    Dim PlannedNO As Range
    Dim MRPCtrl As Range
    Dim MRPdate As Range
    Dim Materialno As Range
    Dim MaterialDesc As Range
    Dim PlannedQty As Range
    Dim Cur As Range
    Dim MRPMessage As Range
    Dim SO As Range
    Dim SOprojectname As Range
    Dim newname As String
    Dim SOitem As Range
    Dim Sourcefilename As String
    Dim destinationSheet As Worksheet
    
    Sourcefilename = "PlannedOrder.xlsx"
    'newname = "Sheet1"
    
    ' Caminho para o arquivo de origem
    Dim sourceFilePath As String
    sourceFilePath = "C:\SAP\PlannedOrder.xlsx" ' Ajuste o caminho conforme necessário
    Set sourceWorkbook = Workbooks.Open(sourceFilePath)
    Set sourceWorkbook = Workbooks(Sourcefilename)

    
    ' Definir os intervalos de origem
    Set sourceRange = sourceWorkbook.Sheets("Sheet1").Range("A2:A" & sourceWorkbook.Sheets("Sheet1").Range("A" & Rows.Count).End(xlUp).Row)
    sourceRange.Interior.Color = RGB(237, 125, 49) 'Planned Open Date Copy
    
    Set PGRR = sourceWorkbook.Sheets("Sheet1").Range("B2:B" & sourceWorkbook.Sheets("Sheet1").Range("B" & Rows.Count).End(xlUp).Row)
    PGRR.Interior.Color = RGB(237, 125, 49) 'PGR Copy
    
    Set PlannedDate = sourceWorkbook.Sheets("Sheet1").Range("C2:C" & sourceWorkbook.Sheets("Sheet1").Range("C" & Rows.Count).End(xlUp).Row)
    PlannedDate.Interior.Color = RGB(237, 125, 49) 'Planned Start Date
    
    Set PlannedNO = sourceWorkbook.Sheets("Sheet1").Range("D2:D" & sourceWorkbook.Sheets("Sheet1").Range("D" & Rows.Count).End(xlUp).Row)
    PlannedNO.Interior.Color = RGB(237, 125, 49) 'Planned Order No
    
    Set MRPCtrl = sourceWorkbook.Sheets("Sheet1").Range("F2:F" & sourceWorkbook.Sheets("Sheet1").Range("F" & Rows.Count).End(xlUp).Row)
    MRPCtrl.Interior.Color = RGB(237, 125, 49) 'MRP Controller
    
    Set MRPdate = sourceWorkbook.Sheets("Sheet1").Range("E2:E" & sourceWorkbook.Sheets("Sheet1").Range("E" & Rows.Count).End(xlUp).Row)
    MRPdate.Interior.Color = RGB(237, 125, 49) 'MRP REQ DATE
    
    Set Materialno = sourceWorkbook.Sheets("Sheet1").Range("G2:G" & sourceWorkbook.Sheets("Sheet1").Range("G" & Rows.Count).End(xlUp).Row)
    Materialno.Interior.Color = RGB(237, 125, 49) 'Material No.
    
    Set MaterialDesc = sourceWorkbook.Sheets("Sheet1").Range("H2:H" & sourceWorkbook.Sheets("Sheet1").Range("H" & Rows.Count).End(xlUp).Row)
    MaterialDesc.Interior.Color = RGB(237, 125, 49) 'Material Description
    
    Set PlannedQty = sourceWorkbook.Sheets("Sheet1").Range("I2:I" & sourceWorkbook.Sheets("Sheet1").Range("I" & Rows.Count).End(xlUp).Row)
    PlannedQty.Interior.Color = RGB(237, 125, 49) 'Planned Order QTY
    
    Set Cur = sourceWorkbook.Sheets("Sheet1").Range("J2:J" & sourceWorkbook.Sheets("Sheet1").Range("J" & Rows.Count).End(xlUp).Row)
    Cur.Interior.Color = RGB(237, 125, 49)
    
    Set MRPMessage = sourceWorkbook.Sheets("Sheet1").Range("K2:K" & sourceWorkbook.Sheets("Sheet1").Range("K" & Rows.Count).End(xlUp).Row)
    MRPMessage.Interior.Color = RGB(237, 125, 49)
    
    Set SO = sourceWorkbook.Sheets("Sheet1").Range("L2:L" & sourceWorkbook.Sheets("Sheet1").Range("L" & Rows.Count).End(xlUp).Row)
    SO.Interior.Color = RGB(237, 125, 49)
    
    Set SOitem = sourceWorkbook.Sheets("Sheet1").Range("M2:M" & sourceWorkbook.Sheets("Sheet1").Range("M" & Rows.Count).End(xlUp).Row)
    SOitem.Interior.Color = RGB(237, 125, 49)
    
    Set SOprojectname = sourceWorkbook.Sheets("Sheet1").Range("N2:N" & sourceWorkbook.Sheets("Sheet1").Range("N" & Rows.Count).End(xlUp).Row)
    SOprojectname.Interior.Color = RGB(237, 125, 49)

    
    ' Definir a planilha de destino
    Set destinationWorkbook = ThisWorkbook ' Supondo que este código esteja em SAPGUIEXPO.xlsm
    Set destinationSheet = destinationWorkbook.Sheets("ME5A")
    
lastRow = destinationSheet.Cells.Find(What:="*", _
                After:=destinationSheet.Cells(4, 4), _
                LookIn:=xlFormulas, _
                LookAt:=xlPart, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious, _
                MatchCase:=False).Row + 1


Dim pasteRange As Range
Set pasteRange = destinationSheet.Range("D" & lastRow)
Dim B As Range
Set B = destinationSheet.Range("B" & lastRow)

'Set the value for the first cell in column A
pasteRange.Value = "Planned Order"

' Fill down the value to all cells below the last row
destinationSheet.Range(pasteRange, destinationSheet.Cells(destinationSheet.Rows.Count, "D")).FillDown
pasteRange.Interior.Color = RGB(237, 125, 49)

    ' Copiar e colar o intervalo
    sourceRange.copy
    destinationSheet.Range("C" & lastRow).PasteSpecial Paste:=xlPasteAll

    PGRR.copy
    destinationSheet.Range("B" & lastRow).PasteSpecial Paste:=xlPasteAll 'ValuesAndNumberFormats

    PlannedDate.copy
    destinationSheet.Range("A" & lastRow).PasteSpecial Paste:=xlPasteAll

    PlannedNO.copy
    destinationSheet.Range("D" & lastRow).PasteSpecial Paste:=xlPasteAll

    MRPCtrl.copy
    destinationSheet.Range("Z" & lastRow).PasteSpecial Paste:=xlPasteAll
    
    MRPdate.copy
    destinationSheet.Range("C" & lastRow).PasteSpecial Paste:=xlPasteAll
    
    Materialno.copy
    destinationSheet.Range("M" & lastRow).PasteSpecial Paste:=xlPasteAll
    
    MaterialDesc.copy
    destinationSheet.Range("N" & lastRow).PasteSpecial Paste:=xlPasteAll
    
    PlannedQty.copy
    destinationSheet.Range("P" & lastRow).PasteSpecial Paste:=xlPasteAll
    
    Cur.copy
    destinationSheet.Range("R" & lastRow).PasteSpecial Paste:=xlPasteAll
    
    SO.copy
    destinationSheet.Range("AM" & lastRow).PasteSpecial Paste:=xlPasteAll
    
    SOitem.copy
    destinationSheet.Range("AN" & lastRow).PasteSpecial Paste:=xlPasteAll
    
    SOprojectname.copy
    destinationSheet.Range("AO" & lastRow).PasteSpecial Paste:=xlPasteAll
    

    ' Fechar o arquivo de origem
End Sub


