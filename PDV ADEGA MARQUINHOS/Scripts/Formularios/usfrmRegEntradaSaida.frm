VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfrmRegEntradaSaida 
   Caption         =   "Registrar entrada / saída"
   ClientHeight    =   2835
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "usfrmRegEntradaSaida.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfrmRegEntradaSaida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    KeyCode = 0
End Sub

Private Sub CommandButton1_Click()
    If ComboBox1 = "" Then
        MsgBox "Escolha o tipo."
        Exit Sub
    End If
    If TextBox1 = "" Then
        MsgBox "Digite uma escrição"
        Exit Sub
    End If
    
    Dim iRow As Integer
    iRow = Planilha5.Range("F:F").SpecialCells(xlCellTypeConstants).Count + 1
    
    Planilha5.Cells(iRow, 6) = ComboBox1.Value
    Planilha5.Cells(iRow, 7) = CDbl(TextBox2.Value)
    
    iRow = Planilha4.Range("A:A").SpecialCells(xlCellTypeConstants).Count + 1
    
    Planilha4.Cells(iRow, 1) = Now
    Planilha4.Cells(iRow, 2) = ComboBox1.Value
    Planilha4.Cells(iRow, 3) = TextBox1.Value
    Planilha4.Cells(iRow, 4) = CDbl(TextBox2.Value)
    
    ThisWorkbook.Save
    
    Unload Me
    
End Sub

Private Sub TextBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call FormatarValor(TextBox2, KeyCode)
End Sub

Private Sub UserForm_Initialize()
    ComboBox1.AddItem "Entrada"
    ComboBox1.AddItem "Saída"
End Sub
