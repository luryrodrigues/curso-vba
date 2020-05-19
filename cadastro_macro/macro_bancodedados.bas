Attribute VB_Name = "Módulo1"
Sub adc_bancodados()
Attribute adc_bancodados.VB_ProcData.VB_Invoke_Func = " \n14"
'
' adc_bancodados Macro
'

'
    Application.ScreenUpdating = False

    Sheets("EXERCÍCIOS").Select
    Range("B11:E11").Select
    Selection.Copy
    Sheets("CADASTRADOS").Select
    Range("B2").Select
    Selection.End(xlDown).Select
    ' Range("B18").Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Sheets("EXERCÍCIOS").Select
    
    Application.ScreenUpdating = True

End Sub

Sub ordem_alfabetica()
Attribute ordem_alfabetica.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ordem_alfabetica Macro
'

'
    Application.ScreenUpdating = False

    Sheets("CADASTRADOS").Select
    Range("B3").Select
    Set R1 = Range(Selection, Selection.End(xlDown))
    Set R2 = Range(R1, Selection.End(xlToRight))
    Set BancoDados = Union(R1, R2)
    BancoDados.Select
    ' Range(Selection, Selection.End(xlDown)).Select
    ' Range(Selection, Selection.End(xlToRight)).Select
    ActiveWorkbook.Worksheets("CADASTRADOS").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("CADASTRADOS").Sort.SortFields.Add2 Key:=Range("B3" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("CADASTRADOS").Sort
        .SetRange BancoDados
        '.SetRange Range("B3:E18")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("EXERCÍCIOS").Select
    
    Application.ScreenUpdating = True
    
End Sub
Sub duplic()
Attribute duplic.VB_ProcData.VB_Invoke_Func = " \n14"
'
' duplic Macro
'

'
    Application.ScreenUpdating = False

    Sheets("CADASTRADOS").Select
    Range("B3").Select
    Set R1 = Range(Selection, Selection.End(xlDown))
    Set R2 = Range(R1, Selection.End(xlToRight))
    Set BancoDados = Union(R1, R2)
    BancoDados.Select
    ' ActiveSheet.Range("B2:E19").RemoveDuplicates Columns:=Array(1, 2, 3, 4), _
        Header:=xlYes
    BancoDados.RemoveDuplicates Columns:=Array(1, 2, 3, 4), Header:=xlYes
    Sheets("EXERCÍCIOS").Select
    
    Application.ScreenUpdating = True
    
End Sub
