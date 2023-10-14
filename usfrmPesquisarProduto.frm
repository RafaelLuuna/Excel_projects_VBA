VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfrmPesquisarProduto 
   Caption         =   "Pesquisar prouto"
   ClientHeight    =   4410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6735
   OleObjectBlob   =   "usfrmPesquisarProduto.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfrmPesquisarProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i As Integer
    For i = 1 To ListBox1.ListCount
        If ListBox1.Selected(i) Then
            On Error Resume Next
            Dim Qnt As Double
            Qnt = InputBox("Digite a quantidade", "Quantidade")
            If Err.Number = 13 Then Exit Sub
            Call Functions.RegistrarItem(Planilha3.Cells(ListBox1.List(i, 0), 1), Qnt)
            Unload Me
            Exit Sub
        End If
    Next
End Sub

Private Sub AtualizarListBox()
    ListBox1.Clear
    ListBox1.AddItem ""
    
    ListBox1.List(0, 0) = "#"
    ListBox1.List(0, 1) = "Produto"
    ListBox1.List(0, 2) = "Valor"
    Dim i, iList As Integer
    Dim textProduto As String
    iList = 1
    For i = 2 To Planilha3.UsedRange.Rows.Count
        If InStr(1, UCase(Planilha3.Cells(i, 2)), UCase(TextBox1)) > 0 Or TextBox1 = "" Then
            ListBox1.AddItem ""
            ListBox1.List(iList, 0) = i
            ListBox1.List(iList, 1) = Planilha3.Cells(i, 2)
            ListBox1.List(iList, 2) = Format(Planilha3.Cells(i, 4), "R$ 0.00")
            iList = iList + 1
        End If
    Next
End Sub

Private Sub TextBox1_Change()
    Call AtualizarListBox
End Sub

Private Sub UserForm_Initialize()
    Call AtualizarListBox
End Sub
