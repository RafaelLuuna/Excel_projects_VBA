Attribute VB_Name = "FormatarNumeros"
Sub FormatarData(tb, Key)
    Select Case Key
        Case 8, 46, 13
        Case 48 To 57, 96 To 107
            If Len(tb.Value) = 2 Then tb.Value = tb.Value & "/"
            If Len(tb.Value) = 5 Then tb.Value = tb.Value & "/"
            If Len(tb.Value) >= 10 Then Key = 0
        Case Else
            Key = 0
    End Select
End Sub

Sub FormatarCPF(tb, Key)
    Select Case Key
        Case 8, 46, 13
        Case 48 To 57, 96 To 107
            If Len(tb.Value) = 3 Then tb.Value = tb.Value & "."
            If Len(tb.Value) = 7 Then tb.Value = tb.Value & "."
            If Len(tb.Value) = 11 Then tb.Value = tb.Value & "-"
            If Len(tb.Value) >= 14 Then Key = 0
        Case Else
            Key = 0
    End Select
End Sub

Sub FormatarTelefone(tb, Key)
    Select Case Key
        Case 8, 46, 13
        Case 48 To 57, 96 To 107
            If Len(tb.Value) = 0 Then tb.Value = "("
            If Len(tb.Value) = 3 Then tb.Value = tb.Value & ") "
            If Len(tb.Value) = 9 Then tb.Value = tb.Value & "-"
            If Len(tb.Value) = 14 Then tb.Value = Left(tb.Value, 6) & " " & Mid(tb.Value, 7, 3) & Mid(tb.Value, 11, 1) & "-" & Right(tb.Value, 3)
            If Len(tb.Value) = 16 Then Key = 0
        Case Else
            Key = 0
    End Select
End Sub

Sub LimitarTamanho(tb, Key, Tamanho)
    Select Case Key
        Case 8, 46, 13
        Case 48 To 57, 96 To 107
            If Len(tb.Value) >= Tamanho Then Key = 0
            Debug.Print Tamanho
        Case Else
            Key = 0
    End Select
End Sub

Sub FormatarValor(tb, Key)
If tb.Text = "" Then tb.Text = "0,00"

Select Case Key

    'PERMIÇÕES
    Case 13 'ENTER
    Case 112 'F1
    Case 113 'F2
    Case 114 'F3
    Case 115 'F4
    Case 27 'ESC

    'VALOR NEGATIVO
    Case 109
        Key = 0
        If Not Left(tb, 1) = "-" Then
            tb.Value = "-" & tb
        End If

    'BACKSPACE
    Case 8
        Key = 0
        'NEGATIVO
        If tb = "-0,00" Then
            tb.Value = "0,00"
        End If
        If Left(tb, 1) = "-" Then
            tb.Value = Right(tb, Len(tb) - 1)
            On Error Resume Next
            If Len(tb) > 4 Then
                tb.Value = Left(tb, Len(tb) - 4) & "," & Mid(tb, Len(tb) - 3, 1) & Mid(tb, Len(tb) - 1, 1)
            ElseIf Not tb.Value = "0," & Left(tb, 1) & Mid(tb, Len(tb) - 1, 1) Then
                tb.Value = "0," & Left(tb, 1) & Mid(tb, Len(tb) - 1, 1)
            ElseIf Not tb.Value = "0,0" & Mid(tb, Len(tb) - 1, 1) Then
                tb.Value = "0,0" & Mid(tb, Len(tb) - 1, 1)
            ElseIf Not tb.Value = "0,00" Then
                tb.Value = "0,00"
            End If
            tb.Value = "-" & tb
        'POSITIVO
        Else
            On Error Resume Next
            If Len(tb) > 4 Then
                tb.Value = Left(tb, Len(tb) - 4) & "," & Mid(tb, Len(tb) - 3, 1) & Mid(tb, Len(tb) - 1, 1)
            ElseIf Not tb.Value = "0," & Left(tb, 1) & Mid(tb, Len(tb) - 1, 1) Then
                tb.Value = "0," & Left(tb, 1) & Mid(tb, Len(tb) - 1, 1)
            ElseIf Not tb.Value = "0,0" & Mid(tb, Len(tb) - 1, 1) Then
                tb.Value = "0,0" & Mid(tb, Len(tb) - 1, 1)
            ElseIf Not tb.Value = "0,00" Then
                tb.Value = "0,00"
            End If
        End If


    'TECLADO NUMERICO
    Case 48 To 57, 96 To 105
        'NEGATIVO
        If Left(tb, 1) = "-" Then
            tb.Value = Right(tb, Len(tb) - 1)
            If Left(tb, 4) = "0,00" Then
                tb.Value = "0,0"
            ElseIf Left(tb, 3) = "0,0" Then
                tb.Value = "0," & Right(tb, 1)
            ElseIf Left(tb, 2) = "0," Then
                tb.Value = Mid(tb, 3, 1) & "," & Right(tb, 1)
            Else
                tb.Value = Left(tb, Len(tb) - 3) & Mid(tb, Len(tb) - 1, 1) & "," & Right(tb, 1)
            End If
            tb.Value = "-" & tb
        'POSITIVO
        Else
            If Left(tb, 4) = "0,00" Then
                tb.Value = "0,0"
            ElseIf Left(tb, 3) = "0,0" Then
                tb.Value = "0," & Right(tb, 1)
            ElseIf Left(tb, 2) = "0," Then
                tb.Value = Mid(tb, 3, 1) & "," & Right(tb, 1)
            Else
                tb.Value = Left(tb, Len(tb) - 3) & Mid(tb, Len(tb) - 1, 1) & "," & Right(tb, 1)
            End If
        End If

    Case Else
        Debug.Print Key
        Key = 0
End Select
End Sub

