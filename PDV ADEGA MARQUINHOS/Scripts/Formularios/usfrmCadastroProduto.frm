VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfrmCadastroProduto 
   Caption         =   "Cadastro de produto"
   ClientHeight    =   2730
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3960
   OleObjectBlob   =   "usfrmCadastroProduto.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfrmCadastroProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    If CodBarras = "" Then
        MsgBox "Digite o código do produto"
        Exit Sub
    End If
    If DescProduto = "" Then
        MsgBox "Digite a descrição do produto"
        Exit Sub
    End If
    If ValorVenda = "0,00" Then
        MsgBox "Digite o valor de venda"
        Exit Sub
    End If
    
    Dim iRow As Integer
    
    iRow = Planilha3.Range("A:A").SpecialCells(xlCellTypeConstants).Count + 1
    
    With Planilha3
        .Cells(iRow, 1) = CodBarras
        .Cells(iRow, 2) = DescProduto
        .Cells(iRow, 3) = CDbl(ValorCusto)
        .Cells(iRow, 4) = CDbl(ValorVenda)
    End With
    
    ThisWorkbook.Save
    
    MsgBox "Produto cadastrado com sucesso!"
    
    Unload Me
    
End Sub

Private Sub UserForm_Initialize()
    BloquearCadastrto = False
End Sub

Private Sub ValorCusto_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call FormatarValor(ValorCusto, KeyCode)
End Sub

Private Sub ValorVenda_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call FormatarValor(ValorVenda, KeyCode)
End Sub
