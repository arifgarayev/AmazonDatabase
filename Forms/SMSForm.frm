VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SMSForm 
   Caption         =   "SMS-Notification"
   ClientHeight    =   5265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11460
   OleObjectBlob   =   "SMSForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SMSForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub btn_sendSMS_Click()
    If Me.lable_check.Caption = "P" Then
        MsgBox ("Success! E-Mail has been successfully sent!")
    ElseIf Me.lable_check.Caption = "" Then
        MsgBox ("Please enter an E-Mail adress!")
    Else
        MsgBox ("Invalid E-Mail Adress!")
    End If
    Call SMSSender
End Sub




Private Sub ClinetInfo_Click()

End Sub

Private Sub listbox_clientSMS_Click()
    Dim i As Long
    
    For i = 0 To Me.listbox_clientSMS.ListCount - 1
        If Me.listbox_clientSMS.Selected(i) Then
            Me.txb_orderSMS = Me.listbox_clientSMS.List(i)
            Me.txb_nameSMS = Me.listbox_clientSMS.List(i, 1)
            Me.txb_telSMS = Me.listbox_clientSMS.List(i, 2)
            Me.txb_priceSMS = Me.listbox_clientSMS.List(i, 3)
            Me.txb_timeSMS = Me.listbox_clientSMS.List(i, 4)
        End If
    Next

End Sub

Private Sub txb_priceSMS_Enter()
    If Len(Me.txb_priceSMS) = 0 Then
        Me.txb_priceSMS = "$"
        
    End If
End Sub

Private Sub txb_telSMS_Change()
    If Len(Me.txb_telSMS) = 0 Then
        Me.lable_check.Caption = ""
        Exit Sub
    End If
    
    If InStr(Me.txb_telSMS, "@") > 0 Then
        Me.lable_check.ForeColor = VBA.RGB(84, 130, 53)
        Me.lable_check.Caption = "P"
        
        
    Else
        Me.lable_check.ForeColor = VBA.RGB(255, 0, 0)
        Me.lable_check.Caption = "O"
    End If
    
End Sub

Private Sub UserForm_Initialize()
    Call LastOrderForSMS
    Call DownloadClient
End Sub

