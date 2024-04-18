VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_GameMenu 
   Caption         =   "GameMenu"
   ClientHeight    =   4665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "form_GameMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_GameMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnIniciar_Click()
    Call Init
    form_GameMenu.Hide
End Sub
Private Sub btnReiniciar_Click()
    Call Init
    form_GameMenu.Hide
End Sub

Private Sub btnEncerrar_Click()
    Call EndGame
    form_GameMenu.Hide
End Sub
