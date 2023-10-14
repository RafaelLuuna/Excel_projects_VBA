Attribute VB_Name = "Functions"
Public MetodoPagamento As String
Public DtVenda As Date
Public NumVenda As Long

Public LinProduto As Integer

Public Sub RegistrarItem(Codigo As String, Qnt As Double)
    On Error Resume Next
    Dim ProdExiste As Boolean
    ProdExiste = False
    If WorksheetFunction.VLookup(Codigo, Planilha3.Range("tabProdutos[Código]"), 1, False) = Codigo Then ProdExiste = True
    If Err.Number = 1004 Then
        MsgBox "Produto não encontrado"
        Exit Sub
    End If

    Planilha1.Activate
    Range("B5") = Codigo
    Range("D5") = Qnt
    
    Range("C5").Select
    Selection.Copy
    Range("C4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

    Range("B5:F5").Select
    Selection.Copy

    Dim i As Integer
    For i = 7 To 1007
        If Planilha1.Cells(i, 2) = "" Then
            Exit For
        End If
    Next

    Planilha1.Cells(i, 2).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

    Range("B5").Value = ""
    Range("D5").Value = ""

    Planilha1.BarCode = ""
    Planilha1.Qnt = 1
    
    ThisWorkbook.Save

    Planilha1.BarCode.Activate
    
    



End Sub

Public Sub FecharVenda()
    If WorksheetFunction.CountIf(Planilha1.Range("tabPDV[Código de barra]"), "<>") = 0 Then
        MsgBox "A venda atual está vazia, por favor insira algum produto para finalizar a venda."
        Exit Sub
    End If
    
    usfrmMetodoPagamento.Show
    If MetodoPagamento = "Cancelar" Then
        Planilha5.Range("K:L").ClearContents
        Planilha5.Range("K1") = "Valor pago (temp.)"
        Planilha5.Range("L1") = "Tipo Pagamento (temp.)"
        Exit Sub
    End If
    
    Planilha1.Activate
    'Classifica a coluna de código de barras
    'Isso serve para não copiar linhas em branco
    Range("tabPDV[Código de barra]").Select
    Planilha1.ListObjects("tabPDV").Sort.SortFields.Clear
    Planilha1.ListObjects("tabPDV").Sort.SortFields.Add2 _
        Key:=Range("tabPDV[Código de barra]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With Planilha1.ListObjects("tabPDV").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Dim iRow, iRow2 As Integer
    
    NumVenda = Planilha1.Range("A5").Value
    DtVenda = Now
    
    'Insere o número do pedido na aba "Info. do caixa atual"
    iRow = Planilha5.Range("D:D").SpecialCells(xlCellTypeConstants).Count + 1
    Planilha5.Activate
    Planilha5.Cells(iRow, 4) = NumVenda
    
    'Copia a tabela do PDV
    Planilha1.Activate
    Planilha1.Range("B7:F" & Range("tabPDV[Código de barra]").SpecialCells(xlCellTypeConstants).Count + 6).Select
    Selection.Copy
    
    'Colar na tabVendas
    iRow = Planilha2.Range("A:A").SpecialCells(xlCellTypeConstants).Count + 1
    Planilha2.Activate
    Planilha2.Cells(iRow, 3).Select
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    iRow2 = Planilha2.Range("C:C").SpecialCells(xlCellTypeConstants).Count

    iRow = Planilha2.Range("A:A").SpecialCells(xlCellTypeConstants).Count + 1
    Planilha2.Range("A" & iRow & ":A" & iRow2).Value = NumVenda
    
    iRow = Planilha2.Range("B:B").SpecialCells(xlCellTypeConstants).Count + 1
    Planilha2.Range("B" & iRow & ":B" & iRow2).Value = DtVenda
    
    
    Planilha5.Range("P1:P6").Copy
    
    Planilha8.Activate
    iRow = Planilha8.Range("A:A").SpecialCells(xlCellTypeConstants).Count + 1
    Planilha8.Cells(iRow, 1) = NumVenda
    Planilha8.Cells(iRow, 2).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    
    Planilha5.Range("K:L").ClearContents
    Planilha5.Range("K1") = "Valor pago (temp.)"
    Planilha5.Range("L1") = "Tipo Pagamento (temp.)"
    
    
    Planilha1.Activate
    Planilha1.Range("B6").Select
    Planilha1.BarCode.Activate
    
    Planilha1.ListObjects("tabPDV").DataBodyRange.SpecialCells(xlCellTypeConstants).ClearContents
    
    ThisWorkbook.Save
    
    
End Sub

Public Sub Hide_Menus()
    Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"",False)"
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = False
    ActiveWindow.DisplayWorkbookTabs = False
    ActiveWindow.DisplayHorizontalScrollBar = False
    
End Sub

Public Sub Show_Menus()
    Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"",True)"
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
    ActiveWindow.DisplayWorkbookTabs = True
    ActiveWindow.DisplayHorizontalScrollBar = True
End Sub
