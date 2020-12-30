VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sales 
   Caption         =   "Sales"
   ClientHeight    =   8895.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8415.001
   OleObjectBlob   =   "Sales.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btn_dlt_Click()
    Call DeleteSales
End Sub

Private Sub btn_updt_Click()
    Call UpdateSales
End Sub

Private Sub cbx_category_Change()
    Call FillGroup
End Sub

Private Sub chbx_pick_Click()
    If chbx_pick.Value Then
        Me.cbx_driver.Value = ""
        Me.cbx_driver.Enabled = False
        Me.txb_delivprice.Enabled = False
        Me.cbx_driver.Value = ""
        Me.txb_delivprice.Value = ""
        Me.cbx_city.Enabled = False
        Me.cbx_city.Value = ""
    Else
    
        Me.cbx_driver.Enabled = True
            
    
    End If

End Sub

Private Sub CommandButton1_Click()
    Call AddSales
End Sub


Private Sub CommandButton4_Click()

End Sub

Private Sub spin_quantity_Change()
    Me.txb_quantity.Value = Me.spin_quantity.Value
End Sub

Private Sub txb_price_Change()
    Call SummCalculate
End Sub

Private Sub txb_quantity_Change()
    Call SummCalculate
End Sub


Private Sub txb_quantity_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 48 To 57, 8 'numbers from 0-9
        Case 44, 46, 47 'comma dote or slash
            KeyAscii = 44
        
            If InStr(Me.txb_quantity, ",") Then
                KeyAscii = 0
            End If
        Case Else
            KeyAscii = 0 'Nothing will happen
            
    End Select
End Sub

Private Sub txb_vendor_Change()
    Call SearchVendorCode
End Sub

Private Sub UserForm_Initialize()
    Call FillCategory
    Me.spin_quantity.Value = 1
    Me.cbx_status.Value = "New"
    Me.opt_client.Value = True
    If ThisWorkbook.Worksheets("Settings").Range("UpdateForm") = 1 Then
    
        Me.btn_updt.Enabled = True
        Me.btn_dlt.Enabled = True
        Me.CommandButton1.Enabled = False
        
        Call UpdateFill
    End If
    
End Sub
