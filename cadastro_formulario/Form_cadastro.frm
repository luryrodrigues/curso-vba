VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_cadastro 
   Caption         =   "Cadastro"
   ClientHeight    =   4920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4155
   OleObjectBlob   =   "Form_cadastro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_fechar_Click()

    Form_cadastro.Hide
    
    Unload Form_cadastro

End Sub

Private Sub btn_limpar_Click()

    txt_nome = ""
    txt_cidade = ""
    txt_fruta = ""
    txt_cor = ""

End Sub

Private Sub btn_cadastrar_Click()

    If txt_cidade.Value = "Lorena" Or txt_cidade.Value = "Itajubá" Or txt_cidade.Value = "SJC" Then

        Sheets("CADASTRADOS").Select
        Range("B3").Select
        Selection.End(xlDown).Select
        ActiveCell.Offset(1, 0).Select
        
        ActiveCell.Value = txt_nome
        ActiveCell.Offset(0, 1).Value = txt_cidade
        ActiveCell.Offset(0, 2).Value = txt_fruta
        ActiveCell.Offset(0, 3).Value = txt_cor
        
        Form_cadastro.Hide
        
        txt_nome = ""
        txt_cidade = ""
        txt_fruta = ""
        txt_cor = ""
    Else
        MsgBox "Cidade incorreta!"
    End If
        
End Sub
