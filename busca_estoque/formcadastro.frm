VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formcadastro 
   Caption         =   "Cadastro de novo usuário"
   ClientHeight    =   2790
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "formcadastro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formcadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_cadastrar_Click()
    
    Dim usuario As String
    Dim senha As String
    
    usuario = txt_novousuario.Value
    senha = txt_novasenha.Value
    
    If txt_novousuario = "" Or txt_novasenha = "" Then
        MsgBox "Preencha os campos requeridos"
    Else
        Worksheets("PERMISSÕES").visible = True
        Sheets("PERMISSÕES").Select
        Range("C3").Select
        Selection.End(xlDown).Select
        ActiveCell.Offset(1, 0).Select
        If Range("C:C").Find(txt_novousuario.Value) Is Nothing And Range("D:D").Find(txt_novasenha.Value) Is Nothing Then
            ActiveCell.Value = txt_novousuario
            ActiveCell.Offset(0, 1).Value = txt_novasenha
            Unload formcadastro
            Unload formlogin2
            Unload formestoque
            Sheets("EXERCÍCIOS").Select
            MsgBox "Usuário cadastrado com sucesso!"
        Else
            Sheets("EXERCÍCIOS").Select
            txt_novousuario.Value = ""
            txt_novasenha.Value = ""
            MsgBox "Nome ou senha de usuário já cadastrado. Tente um diferente!"
        End If
    End If
        Worksheets("PERMISSÕES").visible = False
End Sub

Private Sub btn_canc_Click()

    Unload formcadastro
    Unload formlogin2

End Sub

Private Sub UserForm_Click()

End Sub
