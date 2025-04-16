VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Insert your PGR"
   ClientHeight    =   2790
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6790
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()

End Sub
Private Sub CommandButton1_Click()
'Dim ws As Worksheet
' Set ws = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to the name of your sheet
'ws.Rows("3:" & ws.Rows.Count).Clear

ThisWorkbook.Sheets("PGR").Activate
linha = Range("A2").End(xlDown).Row
Cells(linha, 1) = CaixaPGR
    
If CheckBox1.Value = True Then
Call ME5A.closeUserform 'Fecha Userform
Call ME5A.Book1_DOWN 'ME5A
Call ME5A.copy 'ME5A
Call SalesOrder.Se16n 'EBKN = Sales Order
Call SalesOrder.copy
Call SalesOrder.UpdateFormula 'Lookup de Sales order entre Sheet1 e Sheet2
Call SalesOrder.UpdateFormulaitem
Call Ato.Ato 'VBAK = Ato consessório e nome do projeto
Call Ato.copy2
Call Ato.UpdateFormulaato
Call Ato.UpdateFormulaPJ 'VLookup entre Sales Order e Nome do Projeto
Call Contract.Contrato
Call Contract.copy2
Call Contract.Macro1
Call Contract.UpdateFormulaAgmt
Call Contract.UpdateFormulaContrato
Call Contract.UpdateFormulaitem
Call ReschDate.ReschDate
Call ReschDate.copy3
Call ME2N.ME2N
Call ME2N.teste
Call ME2N.AplicarVLOOKUP_ME5A
Call ReschDate.UpdateFormulaMSG
Call PlannedOrder.ZPTP_MPLN 'Planned Order
Call PlannedOrder.teste 'Cria pasta para Planned Order e copia dados de dentro dela
Call Module1.AdicionarIconesBandeira
Call Module1.DecreaseDecimalPlaces
Call Module1.copy

'Call ReschDate.UpdateFormulaato 'VLookup entre Sales Order e Ato concessório

Else

Call ME5A.closeUserform 'Fecha Userform
Call ME5A.Book1_DOWN 'ME5A
Call ME5A.copy 'ME5A
Call SalesOrder.Se16n 'EBKN = Sales Order
Call SalesOrder.copy
Call SalesOrder.UpdateFormula 'Lookup de Sales order entre Sheet1 e Sheet2
Call SalesOrder.UpdateFormulaitem
Call Ato.Ato 'VBAK = Ato consessório e nome do projeto
Call Ato.copy2
Call Ato.UpdateFormulaPJ 'VLookup entre Sales Order e Nome do Projeto
Call Ato.UpdateFormulaato
Call Contract.Contrato
Call Contract.copy2
Call ReschDate.ReschDate
Call ReschDate.copy3
Call ReschDate.UpdateFormulaMSG
Call ME2N.ME2N
Call ME2N.teste
Call ME2N.AplicarVLOOKUP_ME5A
Call Contract.Macro1
Call Contract.UpdateFormulaAgmt
Call Contract.UpdateFormulaContrato
Call Contract.UpdateFormulaitem
Call Module1.copy
Call Module1.AdicionarIconesBandeira
Call Module1.DecreaseDecimalPlaces
Call Module1.copy
'Call ReschDate.UpdateFormulaato 'VLookup entre Sales Order e Ato concessório

End If
End Sub

Private Sub UserForm_Click()

End Sub
