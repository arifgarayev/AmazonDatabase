VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Pass 
   Caption         =   "System Log in"
   ClientHeight    =   3495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7125
   OleObjectBlob   =   "Pass.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Pass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbx_user_Change()

End Sub

Private Sub CommandButton1_Click()
    Call CheckUser
End Sub

Private Sub UserForm_Initialize()
    Call FillUser
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = 0 Then
        MsgBox ("Access is denied. For system login, please enter Login and Password! "), vbCritical, "Caution!"
        Application.Quit
        ActiveWorkbook.Save
    End If

End Sub
