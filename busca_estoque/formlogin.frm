VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formlogin 
   Caption         =   "Login"
   ClientHeight    =   2835
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "formlogin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_canc_Click()

   Unload formlogin
   formaviso.Hide
    
End Sub

Private Sub btn_cancel_Click()

    Unload formlogin
    formaviso.Hide

End Sub

Private Sub btn_login_Click()

    Dim usuario As String
    Dim senha As String
    Dim i As Integer
       
    usuario = txt_usuario.Value
    senha = txt_senha.Value
    
    Worksheets("PERMISSÕES").visible = True
    Sheets("PERMISSÕES").Select
    Range("C3").Select
    
    For i = 1 To 5
        
       If ActiveCell.Value = "" And ActiveCell.Offset(0, 1).Value = "" Then
            Sheets("EXERCÍCIOS").Select
            MsgBox "Usuário/Senha incorreto(s)!"
            txt_usuario.Value = ""
            txt_senha.Value = ""
            Exit For
        End If
        
        If ActiveCell.Value = usuario And ActiveCell.Offset(0, 1).Value = senha Then
            Sheets("EXERCÍCIOS").Select
            ActiveCell.Offset(0, 2).Value = formestoque.txt_qtd
            Unload formlogin
            Unload formaviso
            Unload formestoque
            MsgBox "Estoque atualizado com sucesso!"
            Exit For
        Else
            ActiveCell.Offset(1, 0).Select
        End If
    Next
    
    Worksheets("PERMISSÕES").visible = False
        
End Sub

Private Sub UserForm_Click()

End Sub
