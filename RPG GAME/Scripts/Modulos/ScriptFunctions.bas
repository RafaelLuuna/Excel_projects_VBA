Attribute VB_Name = "ScriptFunctions"
'Com a palavra "Script" nesse código entende-se roteiro
Public Function NextScriptID(ByVal ScriptSequence As Collection) As Integer
    Dim i As Integer
    NextScriptID = ScriptSequence.Item(1)
    For i = 1 To ScriptSequence.Count
        If CheckScriptID(ScriptSequence.Item(i)) Then
            If i = ScriptSequence.Count Then
                NextScriptID = ScriptSequence.Item(i)
            Else
                NextScriptID = ScriptSequence.Item(i + 1)
            End If
        End If
    Next
End Function

Public Function CheckScriptID(ScriptID As Integer) As Boolean
    Dim i As Integer
    CheckScriptID = False
    For i = 1 To DATA.ScriptCheck.Count
        If DATA.ScriptCheck.Item(i) = ScriptID Then CheckScriptID = True
    Next
End Function

Public Sub RenderScript(ScriptLine As Integer)
    With ScriptData
        usfrmTalk.lbMsg = WorksheetFunction.VLookup(DATA.ActualScript & "," & ScriptLine, .Range("C:E"), 3, False)
        usfrmTalk.Caption = WorksheetFunction.VLookup(DATA.ActualScript & "," & ScriptLine, .Range("C:F"), 4, False)
        usfrmTalk.imgTalker.Picture = LoadPicture(ThisWorkbook.Path & "\Texture\Entity\" & WorksheetFunction.VLookup(DATA.ActualScript & "," & ScriptLine, .Range("C:G"), 5, False))
    End With
End Sub


Public Sub DoAction(ScriptLine As Integer)
    Dim NewLine As Integer
    Dim ActionID As String
    ActionID = WorksheetFunction.VLookup(DATA.ActualScript & "," & ScriptLine, ScriptData.Range("C:L"), 6, False)
    
    Select Case ActionID
        Case "GaveItem"
            Dim Var1 As Integer
            Dim Var2 As String
            Dim Var3 As Integer
            Dim Var4 As Integer
            Var1 = WorksheetFunction.VLookup(DATA.ActualScript & "," & ScriptLine, ScriptData.Range("C:L"), 7, False)
            Var2 = WorksheetFunction.VLookup(DATA.ActualScript & "," & ScriptLine, ScriptData.Range("C:L"), 8, False)
            Var3 = WorksheetFunction.VLookup(DATA.ActualScript & "," & ScriptLine, ScriptData.Range("C:L"), 9, False)
            Var4 = WorksheetFunction.VLookup(DATA.ActualScript & "," & ScriptLine, ScriptData.Range("C:L"), 10, False)
            Call InventoryFunctions.AddItem(Var1, Var2, Var3, Var4)
            If InventoryFunctions.InventoryFullTest Then
                MsgBox "Inventory is Full"
            Else
                MsgBox "You get " & Var3 & " " & Var2
            End If
        Case "OptionMode"
            Dim i As Integer
            Set DATA.Talk_OptionTexts = New Collection
            For i = 1 To 20
                If Not WorksheetFunction.VLookup(DATA.ActualScript & "," & ScriptLine, ScriptData.Range("C:AB"), i + 6, False) = "" Then
                    DATA.Talk_OptionTexts.Add WorksheetFunction.VLookup(DATA.ActualScript & "," & ScriptLine, ScriptData.Range("C:AB"), i + 6, False)
                    DATA.Talk_OptionNum = i
                End If
            Next
            usfrmTalk_OptionMode.Show
        Case "OptionSelected"
            Dim OptionValue As String
            OptionValue = WorksheetFunction.VLookup(DATA.ActualScript & "," & ScriptLine, ScriptData.Range("C:AB"), DATA.Talk_SelectedOption + 6, False)
            If Left(OptionValue, 5) = "GoTo:" Then
                NewLine = Right(OptionValue, Len(OptionValue) - 5)
                usfrmTalk.iTalk = NewLine
                Call RenderScript(usfrmTalk.iTalk)
                Call DoAction(usfrmTalk.iTalk)
            ElseIf Left(OptionValue, 7) = "SetVar:" Then
                Dim Line As Integer
                Dim VarNum As Integer
                Dim Value
                'Define os parâmetros
                OptionValue = Right(OptionValue, Len(OptionValue) - 7)
                Line = Left(OptionValue, InStr(1, OptionValue, ",") - 1)
                OptionValue = Right(OptionValue, Len(OptionValue) - InStr(1, OptionValue, ","))
                VarNum = Left(OptionValue, InStr(1, OptionValue, ",") - 1)
                Value = Right(OptionValue, Len(OptionValue) - InStr(1, OptionValue, ","))
                
                'Atualiza as variaveis
                ScriptData.Cells(WorksheetFunction.Match(DATA.ActualScript & "," & Line, ScriptData.Range("C:C"), 0), 8 + VarNum) = Value
                
                usfrmTalk.iTalk = usfrmTalk.iTalk + 1
                Call RenderScript(usfrmTalk.iTalk)
                Call DoAction(usfrmTalk.iTalk)
            Else
                ScriptData.Cells(WorksheetFunction.Match(DATA.ActualScript & "," & ScriptLine, ScriptData.Range("C:C"), 0), 5) = OptionValue
                Call RenderScript(usfrmTalk.iTalk)
            End If
        Case "GoTo"
            NewLine = CInt(WorksheetFunction.VLookup(DATA.ActualScript & "," & ScriptLine, ScriptData.Range("C:I"), 7, False))
            usfrmTalk.iTalk = NewLine
            Call RenderScript(usfrmTalk.iTalk)
            Call DoAction(usfrmTalk.iTalk)
        Case "UpdateWallet"
            Dim WalletID As Integer
            Dim WalletValue As Integer
            Dim OnErrorLine As Integer
            WalletID = WorksheetFunction.VLookup(DATA.ActualScript & "," & ScriptLine, ScriptData.Range("C:L"), 7, False)
            WalletValue = WorksheetFunction.VLookup(DATA.ActualScript & "," & ScriptLine, ScriptData.Range("C:L"), 8, False)
            OnErrorLine = WorksheetFunction.VLookup(DATA.ActualScript & "," & ScriptLine, ScriptData.Range("C:L"), 9, False)
            If WalletData.Cells(WalletID + 1, 2) + WalletValue < 0 Then
                usfrmTalk.iTalk = OnErrorLine
                Call RenderScript(usfrmTalk.iTalk)
                Call DoAction(usfrmTalk.iTalk)
            Else
                WalletData.Cells(WalletID + 1, 2) = WalletData.Cells(WalletID + 1, 2) + WalletValue
                Call Render.UpdateHud
                usfrmTalk.iTalk = usfrmTalk.iTalk + 1
                Call RenderScript(usfrmTalk.iTalk)
                Call DoAction(usfrmTalk.iTalk)
            End If
            
            
        Case Else
    End Select
End Sub
