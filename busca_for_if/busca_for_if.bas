Attribute VB_Name = "Módulo1"
Sub estrutura_for()

    Dim nome As String
    Dim nota As Double
    Dim i As Integer

    nome = UCase(Range("C10").Value)
    
    Range("F10").Select
    
    For i = 1 To 100
        If ActiveCell.Value = nome Then
            nota = ActiveCell.Offset(0, 1).Value
            MsgBox "A nota de " & nome & " é: " & nota & "."
    Exit For
        Else
            ActiveCell.Offset(1, 0).Select
        End If
    Next
    
    If ActiveCell.Value = "" Then
        Range("F10").Select
        MsgBox "Nome não encontrado!"
    End If

End Sub
