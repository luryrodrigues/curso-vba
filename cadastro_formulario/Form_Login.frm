VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_Login 
   Caption         =   "Solicitação de Acesso"
   ClientHeight    =   1950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4935
   OleObjectBlob   =   "Form_Login.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_cancelar_Click()

    Form_Login.Hide

    Unload Form_Login

End Sub

Private Sub btn_entrar_Click()

    Dim login As String
    Dim senha As String

    login = txt_login.Value
    senha = txt_senha.Value

    If login = "admin" And senha = "123" Then
        Form_Login.Hide
        Unload Form_Login
        Form_cadastro.Show
    Else
        MsgBox ("Usuário/senha incorreto(s)!")
        Unload Form_Login
    End If

End Sub
