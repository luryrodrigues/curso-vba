VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formestoque 
   Caption         =   "Controle de Estoque"
   ClientHeight    =   3555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4455
   OleObjectBlob   =   "formestoque.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formestoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_buscar_Click()
    
    ' declarar variáveis
    Dim cod As String
    Dim prod As String
    Dim qtd As String
    Dim i As Integer
    
    ' atribuir variáveis
    cod = txt_codigo
    prod = txt_produto
    qtd = txt_qtd
    
    ' fazer a lógica
    Worksheets("PERMISSÕES").visible = True
    Sheets("EXERCÍCIOS").Select
    Range("B10").Select
    
    For i = 1 To 100
        
        If cod = ActiveCell.Value Then
            txt_produto = ActiveCell.Offset(0, 1).Value
            txt_qtd = ActiveCell.Offset(0, 2)
            Exit For
        Else
            ActiveCell.Offset(1, 0).Select
        End If
    Next
    
    If ActiveCell.Value = "" Then
        Range("B10").Select
        MsgBox "Código não encontrado!"
    End If
    
    Worksheets("PERMISSÕES").visible = False
    
End Sub

Private Sub btn_editar_Click()

    formaviso.Show

End Sub

Private Sub btn_fechar_Click()
  
    Unload formestoque

End Sub

Private Sub btn_limpar_Click()
    
    txt_codigo.Value = ""
    txt_produto.Value = ""
    txt_qtd.Value = ""
    
End Sub

Private Sub btn_usuario_Click()

    MsgBox "Para cadastrar um novo usuário, confirme seu login!"
    formlogin2.Show
    
End Sub
