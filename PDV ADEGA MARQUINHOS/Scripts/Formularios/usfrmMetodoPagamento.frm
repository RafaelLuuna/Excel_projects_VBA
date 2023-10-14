VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfrmMetodoPagamento 
   Caption         =   "Método de pagamento"
   ClientHeight    =   2970
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3435
   OleObjectBlob   =   "usfrmMetodoPagamento.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfrmMetodoPagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()
    Select Case ComboBox1
        Case "Dinheiro"
            lbTroco.Visible = True
            lbTituloTroco.Visible = True
        Case Else
            lbTroco.Visible = False
            lbTituloTroco.Visible = False
    End Select
End Sub

Private Sub ComboBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    KeyCode = 0
End Sub

Private Sub CommandButton1_Click()
    If ComboBox1 = "" Then
        MsgBox "Selceione um método de pagamento"
        Exit Sub
    End If
    If CDbl(tbValorPago.Value) = 0 Then
        MsgBox "Digite um valor para continuar"
        Exit Sub
    End If
    
    Dim iRow As Integer
    iRow = Planilha5.Range("K:K").SpecialCells(xlCellTypeConstants).Count + 1
    Planilha5.Cells(iRow, 11) = CDbl(tbValorPago.Value)
    Planilha5.Cells(iRow, 12) = ComboBox1
    
    If Planilha5.Range("M2") > 0 Then
        lbTotal_a_Pagar.Caption = Format(Planilha5.Range("M2").Value, "R$ 0.00")
        tbValorPago = "0,00"
        Exit Sub
    End If
    
    
    Functions.MetodoPagamento = ComboBox1
    
    ThisWorkbook.Save
    
    Unload Me
    
End Sub

Private Sub CommandButton2_Click()
    Functions.MetodoPagamento = "Cancelar"
    Unload Me
End Sub


Private Sub tbValorPago_Change()
    Dim VlTroco As Double
    VlTroco = CDbl(tbValorPago) - Planilha5.Range("M2")
    If VlTroco < 0 Then VlTroco = 0
    lbTroco.Caption = Format(VlTroco, "R$ 0.00")
End Sub

Private Sub tbValorPago_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call FormatarValor(tbValorPago, KeyCode)
End Sub


Private Sub UserForm_Initialize()
    lbTotal_a_Pagar.Caption = Format(Planilha5.Range("M2").Value, "R$ 0.00")
    
    lbTroco.Visible = False
    lbTituloTroco.Visible = False
    Functions.MetodoPagamento = "Cancelar"
    ComboBox1.AddItem "Dinheiro"
    ComboBox1.AddItem "Débito"
    ComboBox1.AddItem "Crédito"
    ComboBox1.AddItem "VR"
    ComboBox1.AddItem "Pix"
End Sub

