VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formaviso 
   Caption         =   "AVISO!"
   ClientHeight    =   1275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4725
   OleObjectBlob   =   "formaviso.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formaviso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_nao_Click()

    formaviso.Hide

End Sub

Private Sub btn_sim_Click()

    formlogin.Show
    
End Sub
