VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formlogin2 
   Caption         =   "Login"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "formlogin2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formlogin2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_cancel_Click()

    Unload formlogin2

End Sub

Private Sub btn_login_Click()


    Dim usuario As String
    Dim senha As String
    Dim i As Integer
       
    usuario2 = txt_usuario2.Value
    senha2 = txt_senha2.Value
    
    Worksheets("PERMISSÕES").visible = True
    Sheets("PERMISSÕES").Select
    Range("C3").Select
    
    For i = 1 To 100
        
       If ActiveCell.Value = "" And ActiveCell.Offset(0, 1).Value = "" Then
            Unload formlogin2
            Sheets("EXERCÍCIOS").Select
            MsgBox "Usuário/Senha incorreto(s)! Não foi possível cadastrar novo usuário."
            Exit For
        Else
            If ActiveCell.Value = usuario2 And ActiveCell.Offset(0, 1).Value = senha2 Then
                Sheets("EXERCÍCIOS").Select
                formcadastro.Show
            Exit For
            Else
                ActiveCell.Offset(1, 0).Select
            End If
        End If
    Next

    Worksheets("PERMISSÕES").visible = False
    
End Sub
