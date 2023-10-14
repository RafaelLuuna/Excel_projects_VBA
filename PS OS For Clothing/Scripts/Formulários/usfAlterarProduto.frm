VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfAlterarProduto 
   Caption         =   "Selecione o Produto"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6195
   OleObjectBlob   =   "usfAlterarProduto.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfAlterarProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call CommandButton1_Click
End Sub

Private Sub CommandButton1_Click()
    If Not lbNome.Caption = "" Then
        On Error Resume Next
        With usfCadastrarProduto
            .Linha = Planilha3.Range("A:G").Find(lbNome.Caption).Row
            .tbDescricao.Text = Planilha3.Cells(.Linha, 2).Value
            .combTipoProduto.Text = Planilha3.Cells(.Linha, 1).Value
            .combFornecedor.Text = Planilha3.Cells(.Linha, 3).Value
            .tbEstoqueMinimo.Text = Planilha3.Cells(.Linha, 4).Value
            .tbValorVenda = Format(Planilha3.Cells(.Linha, 9).Value, "#,##0.00")
            .tbEstoqueInicial.Enabled = False
            .tbTamanho1.Enabled = False
            .tbTamanho2.Enabled = False
            .tbTamanho3.Enabled = False
            .tbTamanho4.Enabled = False
            .tbTamanho5.Enabled = False
            .tbTamanho6.Enabled = False
            .Image1.Picture = LoadPicture(Planilha3.Cells(.Linha, 27))
            If Planilha3.Cells(.Linha, 27) = "Null" Then
                    .Foto = ""
            Else: .Foto = Planilha3.Cells(.Linha, 27)
            End If
            .Show
        End With
    Else: MsgBox "Selecione um produto"
    End If
End Sub


Private Sub ListBox1_Change()
    Dim i As Integer
    For i = 1 To UBound(ListBox1.List)
        If ListBox1.Selected(i) Then
            lbTipo.Caption = ListBox1.List(i, 0)
            lbNome.Caption = ListBox1.List(i, 1)
            Exit For
        Else
            lbTipo.Caption = ""
            lbNome.Caption = ""
        End If
    Next
End Sub





Private Sub combFiltroTipo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    KeyCode = 0
End Sub

Private Sub combFornecedores_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    KeyCode = 0
End Sub

Private Sub combFiltroTipo_Change()
    Call Filtrar
End Sub

Private Sub combFornecedores_Change()
    Call Filtrar
End Sub

Private Sub tbPesquisa_Change()
    tbPesquisa = UCase(tbPesquisa)
    Call Filtrar
End Sub





Private Sub UserForm_Initialize()
        
    Image2.Picture = LoadPicture(Application.ThisWorkbook.Path & "\image source\logo.bmp")
    
    With Planilha3.ListObjects("tabESTOQUE").Range  '|LIMPAR FILTROS tabESTOQUE
        .AutoFilter Field:=1                        '|
        .AutoFilter Field:=2                        '|
        .AutoFilter Field:=3                        '|
        .AutoFilter Field:=4                        '|
        .AutoFilter Field:=5                        '|
        .AutoFilter Field:=6                        '|
        .AutoFilter Field:=7                        '|
    End With                                        '|
'----------------------------------------------------|
    

    
    Dim i As Integer
    
    combFiltroTipo.AddItem "*[TODOS]*"
    combFornecedores.AddItem "*[TODOS]*"
    
    For i = 2 To Planilha3.Range("K:K").End(xlDown).Row
        combFiltroTipo.AddItem Planilha3.Cells(i, 11)
    Next
    For i = 2 To Planilha7.Range("A:A").End(xlDown).Row
        combFornecedores.AddItem Planilha7.Cells(i, 1)
    Next
    
    combFiltroTipo.ListIndex = 0
    combFornecedores.ListIndex = 0
    
    Call Filtrar
    
End Sub




Private Sub Filtrar()
    Dim i, iList As Integer
    
    Dim FilterIndex(2) As Boolean
    FilterIndex(0) = False
    FilterIndex(1) = False
    FilterIndex(2) = False
    If Not combFiltroTipo.Text = "*[TODOS]*" Then FilterIndex(0) = True
    If Not combFornecedores.Text = "*[TODOS]*" Then FilterIndex(1) = True
    If Not tbPesquisa.Text = "" Then FilterIndex(2) = True

    ListBox1.Clear
    lbTipo.Caption = ""
    lbNome.Caption = ""
    
    
    With ListBox1                       '| PREPARAR CABEÇALHO
        .AddItem                        '|
        .List(0, 0) = "[TIPO]"          '|
        .List(0, 1) = "[DESCRIÇÃO]"     '|
        .List(0, 2) = "[FORNECEDOR]"    '|
    End With                            '|
    iList = 1                           '|
'----------------------------------------|
    

    
    For i = 2 To Planilha3.Range("A:G").End(xlDown).Row
FiltroNome:
        If Not FilterIndex(2) Then GoTo FiltroTipo
        
        
        If InStr(1, Planilha3.Range("A:G").Cells(i, 2).Value, tbPesquisa.Text) > 0 Then
            
            If Not FilterIndex(0) And Not FilterIndex(1) Then
                GoTo GetList
            Else: GoTo FiltroTipo
            End If
            
        Else: GoTo NextItem
        End If
        
        
        
FiltroTipo:
        If Not FilterIndex(0) Then GoTo FiltroFornecedor

        If Planilha3.Range("A:G").Cells(i, 1) = combFiltroTipo.Text Then

            If FilterIndex(1) Then
                GoTo FiltroFornecedor
            Else: GoTo GetList
            End If

        Else: GoTo NextItem
        End If


FiltroFornecedor:
        If Not FilterIndex(1) Then GoTo GetList

        If Planilha3.Range("A:G").Cells(i, 3) = combFornecedores.Text Then
            GoTo GetList
        Else: GoTo NextItem
        End If


GetList:
        With Me.ListBox1
            .AddItem
            .List(iList, 0) = Planilha3.Cells(i, 1)
            .List(iList, 1) = Planilha3.Cells(i, 2)
            .List(iList, 2) = Planilha3.Cells(i, 3)
        End With
        iList = iList + 1
NextItem:
        Next i
    


End Sub








