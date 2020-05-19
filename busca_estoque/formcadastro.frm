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
    
    usuario = txt_novousuario
    senha = txt_novasenha
    
    If txt_novousuario = "" Or txt_novasenha = "" Then
        MsgBox "Preencha os campos requeridos"
    Else
        Sheets("PERMISSÕES").Select
        Range("C3").Select
        Selection.End(xlDown).Select
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Value = txt_novousuario
        ActiveCell.Offset(0, 1).Value = txt_novasenha
        Unload formcadastro
        Sheets("EXERCÍCIOS").Select
    End If
           
End Sub

Private Sub btn_canc_Click()

    Unload formcadastro

End Sub
