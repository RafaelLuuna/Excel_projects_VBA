VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfrmFecharCaixa 
   Caption         =   "Fechamento do caixa"
   ClientHeight    =   4830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6060
   OleObjectBlob   =   "usfrmFecharCaixa.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfrmFecharCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Planilha5.Range("B5") = Now
    Planilha5.Range("B8") = CDbl(TextBox1)
    Planilha5.Range("C2") = ""
    
    Planilha5.Range("B1:B15").Copy
    
    Planilha6.Activate
    
    Dim iRow As Integer
    iRow = Planilha6.Range("A:A").SpecialCells(xlCellTypeConstants).Count + 1
    
    Planilha6.Range("A" & iRow).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    
    Planilha5.Activate
    
    Planilha5.Range("B1").ClearContents
    Planilha5.Range("B4:B6").ClearContents
    Planilha5.Range("B8").ClearContents
    
    Planilha5.Range("D2:G1000000").ClearContents
    
    ThisWorkbook.Save
    
    Planilha1.Activate
    
    Unload Me
    
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call FormatarNumeros.FormatarValor(TextBox1, KeyCode)
End Sub


Private Sub UserForm_Initialize()
    Planilha5.Range("C2") = "Cancelar fechamento do caixa"
    lbResponsavel.Caption = Planilha5.Range("B1")
    lbDtAbertura.Caption = Format(Planilha5.Range("B4"), "dd/mm/yyyy hh:mm")
    lbFundoCaixa.Caption = Format(Planilha5.Range("B6"), "R$ 0.00")
    lbVendaDinheiro.Caption = Format(Planilha5.Range("B9"), "R$ 0.00")
    lbVendaDebito.Caption = Format(Planilha5.Range("B10"), "R$ 0.00")
    lbVendaCredito.Caption = Format(Planilha5.Range("B11"), "R$ 0.00")
    lbVendaVR.Caption = Format(Planilha5.Range("B12"), "R$ 0.00")
    lbVendaPix.Caption = Format(Planilha5.Range("B13"), "R$ 0.00")
    lbEntradas.Caption = Format(Planilha5.Range("B14"), "R$ 0.00")
    lbSaidas.Caption = Format(Planilha5.Range("B15"), "R$ 0.00")
    lbValorFechamento.Caption = Format(Planilha5.Range("B7"), "R$ 0.00")
End Sub
