VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfrmCancelarItem 
   Caption         =   "Cancelar item"
   ClientHeight    =   4410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6735
   OleObjectBlob   =   "usfrmCancelarItem.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfrmCancelarItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i As Integer
    For i = 1 To ListBox1.ListCount
        If ListBox1.Selected(i) Then
            Dim Linha As Integer
            Linha = ListBox1.List(i, 0)
            Planilha1.Range("B" & Linha & ":F" & Linha).ClearContents
            
            ThisWorkbook.Save
            
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
    ListBox1.List(0, 2) = "Qnt."
    ListBox1.List(0, 3) = "Valor unitário"
    ListBox1.List(0, 4) = "Valor total"
    Dim i, iList As Integer
    Dim textProduto As String
    iList = 1
    For i = 7 To Planilha1.Cells(Planilha1.ListObjects("tabPDV").ListRows.Count + 6, 2).End(xlUp).Row
        ListBox1.AddItem ""
        ListBox1.List(iList, 0) = i
        ListBox1.List(iList, 1) = Planilha1.Cells(i, 3)
        ListBox1.List(iList, 2) = Planilha1.Cells(i, 4)
        ListBox1.List(iList, 3) = Format(Planilha1.Cells(i, 5), "R$ 0.00")
        ListBox1.List(iList, 4) = Format(Planilha1.Cells(i, 6), "R$ 0.00")
        iList = iList + 1
    Next
End Sub

Private Sub TextBox1_Change()
    Call AtualizarListBox
End Sub

Private Sub UserForm_Initialize()
    Call AtualizarListBox
End Sub
