Attribute VB_Name = "Module1"
Option Explicit

Dim ShSales As Worksheet 'gobal parameter for slaes
Dim ShList As Worksheet 'gobal parameter for list
Dim ShStock As Worksheet '// for stock
Dim ShDelivery As Worksheet '// logistics list
Dim ShSmsTemplate As Worksheet '// Sms-settings list
Dim ShClients As Worksheet '// Customers list
Dim ShSettings As Worksheet
Dim SalesListObj As ListObject
Dim RazdelListObj As ListObject
Dim StockListObj As ListObject
Dim DeliveryListObj As ListObject
Dim SetListObj As ListObject
Dim SmsTempListObj As ListObject
Dim ClientsListObj As ListObject
Dim SalesListRow As ListRow
Dim RazdelListRow As ListRow
Dim DeliveryListRow As ListRow
Dim SetListRow As ListRow
Dim SmsTempListRow As ListRow
Dim StockListRow As ListRow
Dim ClientsListRow As ListRow


Sub ShowSales() 'click macros for sales
Attribute ShowSales.VB_ProcData.VB_Invoke_Func = "S\n14"
    ThisWorkbook.Worksheets("Settings").Range("UpdateForm") = ""
    Sales.Show
End Sub

Sub ShowSMS() 'click macros for sms

    SMSForm.Show

End Sub

Sub AddSales() 'adding a new order
    
    
    
    Set ShSales = ThisWorkbook.Worksheets("Sales") 'appoint and global parameters to sales folder in excel
    Set SalesListObj = ShSales.ListObjects("Sales_tb") 'appoint sheet of sales as object to a sales logicaltable
    
    If Sales.txb_order.Value = "" Then
        Sales.Label1_info.Caption = "Please enter an Order Number!"
        Exit Sub
    End If
    
    If Sales.txb_vendor.Value = "" Then
        Sales.Label1_info.Caption = "Please enter a Vendor Code!"
        Exit Sub
    End If
    
    If Sales.chbx_pick.Value = False Then 'check if pickup then work, if not then skip
        If Sales.cbx_driver.Value = "" Then
            Sales.Label1_info.Caption = "Please enter a Driver Number!"
            Exit Sub
        End If
        If Sales.txb_delivprice.Value = "" Then
            Sales.Label1_info.Caption = "Please enter a Delivery Price!"
            Exit Sub
        End If
        If Sales.cbx_city.Value = "" Then
            Sales.Label1_info.Caption = "Please enter a City!"
            Exit Sub
            End If
    End If
    
    
    ThisWorkbook.Worksheets("Sales").Unprotect Password:=(PassSales())
    
    
    Set SalesListRow = SalesListObj.ListRows.Add 'add a row in a "sales" logical table
    
    
    SalesListRow.Range(1) = Sales.txb_order.Value
    SalesListRow.Range(2) = Sales.cbx_status.Value
    SalesListRow.Range(3) = Sales.cbx_category.Value
    SalesListRow.Range(4) = Sales.cbx_group.Value
    SalesListRow.Range(5) = Sales.txb_vendor.Value
    SalesListRow.Range(6) = Sales.txb_description.Value
    SalesListRow.Range(7) = CDbl(Sales.txb_quantity.Value)
    SalesListRow.Range(8) = Sales.txb_price.Value
    SalesListRow.Range(10) = CDbl(Val(Replace(Trim(Sales.txb_sum.Value), ",", ".")))
    SalesListRow.Range(11) = PurchasedPrice(Sales.txb_vendor.Value)
    SalesListRow.Range(12) = Margin(Sales.txb_price.Value, SalesListRow.Range(11))
    SalesListRow.Range(13) = Profit(Sales.txb_quantity.Value, SalesListRow.Range(12))
    
    
    If Sales.opt_client.Value Then
        
        SalesListRow.Range(9) = "Private Person"
    
    End If
    If Sales.opt_org.Value Then
        
        SalesListRow.Range(9) = "Organization"
    
    End If
    
    SalesListRow.Range(14) = Sales.txb_delivprice.Value
    SalesListRow.Range(15) = Sales.cbx_customer.Value
    SalesListRow.Range(16) = Sales.txb_producer.Value
    
    If Sales.chbx_pick Then
        SalesListRow.Range(17) = "PickUp"
    Else
        SalesListRow.Range(17) = Sales.cbx_driver.Value
        
    End If
    
    SalesListRow.Range(18) = ThisWorkbook.Worksheets("Settings").Range("LastUser")
    
    Call AddDelivery
    
    ThisWorkbook.Worksheets("Sales").Protect Password:=(PassSales())
    
    Sales.Label1_info.Caption = "Information has been added!" & "   Order ¹ " & Sales.txb_order.Value


End Sub
Public Function InTable() As Boolean 'sales sheet call sales form when cliced inside the table
    Set ShSales = ThisWorkbook.Worksheets("Sales")
    Set SalesListObj = ShSales.ListObjects("Sales_tb")

    If Intersect(ActiveCell, SalesListObj.Range) Is Nothing Then
        InTable = False
    ElseIf ActiveCell.Row < 2 Then
    
        InTable = False
    Else
        InTable = True
    End If
End Function
Sub UpdateFill() 'filling sales form while updating sales table
    Set ShSales = ThisWorkbook.Worksheets("Sales")
    Set SalesListObj = ShSales.ListObjects("Sales_tb")
    
    
    
    Set SalesListRow = SalesListObj.ListRows(ActiveCell.Row - 1)
    Sales.txb_order = SalesListRow.Range(1)
    Sales.cbx_status = SalesListRow.Range(2)
    Sales.cbx_category = SalesListRow.Range(3)
    Sales.cbx_group = SalesListRow.Range(4)
    Sales.txb_vendor = SalesListRow.Range(5)
    Sales.txb_description = SalesListRow.Range(6)
    Sales.txb_quantity = SalesListRow.Range(7)
    Sales.txb_price = SalesListRow.Range(8)
    Sales.txb_sum = SalesListRow.Range(10)
    Sales.txb_delivprice = SalesListRow.Range(14)
    Sales.cbx_driver = SalesListRow.Range(17)

    
    
    
    
    
End Sub
Sub UpdateSales() 'Sales Form update the new info
    Set ShSales = ThisWorkbook.Worksheets("Sales")
    Set SalesListObj = ShSales.ListObjects("Sales_tb")
    Set SalesListRow = SalesListObj.ListRows(ActiveCell.Row - 1)
    ThisWorkbook.Worksheets("Sales").Unprotect Password:=(PassSales())
    
    SalesListRow.Range(1) = Sales.txb_order.Value
    SalesListRow.Range(2) = Sales.cbx_status.Value
    SalesListRow.Range(3) = Sales.cbx_category.Value
    SalesListRow.Range(4) = Sales.cbx_group.Value
    SalesListRow.Range(5) = Sales.txb_vendor.Value
    SalesListRow.Range(6) = Sales.txb_description.Value
    SalesListRow.Range(7) = Sales.txb_quantity.Value
    SalesListRow.Range(8) = Sales.txb_price.Value
    SalesListRow.Range(10) = CDbl(Sales.txb_sum)

    
    SalesListRow.Range(11) = PurchasedPrice(Sales.txb_vendor.Value)
    SalesListRow.Range(12) = Margin(Sales.txb_price.Value, SalesListRow.Range(11))
    SalesListRow.Range(13) = Profit(Val(Replace(Trim(Sales.txb_quantity.Value), ",", ".")), SalesListRow.Range(12))
    
    ThisWorkbook.Worksheets("Sales").Protect Password:=(PassSales())
    
    Sales.Label1_info.Caption = "Update has done successfully!" & " " & "Order: " & Sales.txb_order.Value

End Sub
Sub DeleteSales() 'deleting sales info
    Set ShSales = ThisWorkbook.Worksheets("Sales")
    Set SalesListObj = ShSales.ListObjects("Sales_tb")
    Set SalesListRow = SalesListObj.ListRows(ActiveCell.Row - 1)
    
    ThisWorkbook.Worksheets("Sales").Unprotect Password:=(PassSales())
    
    If MsgBox("Are you sure to DELETE Order ¹: " & SalesListRow.Range(1), vbYesNo, "Deleting Order!") = vbYes Then
        SalesListRow.Delete
    Else
        Exit Sub
        
    End If
    
    ThisWorkbook.Worksheets("Sales").Protect Password:=(PassSales())
    
    Sales.Label1_info.Caption = "This Information has been successfully DELETED!"



End Sub

Sub FillCategory() 'Filling category in forrm
    
    Dim i As Long
    

    Set ShList = ThisWorkbook.Worksheets("List") 'fillint a row in a "sales" logical table
    Set RazdelListObj = ShList.ListObjects("Category_tb")
    
    i = 1
    
    For Each RazdelListRow In RazdelListObj.ListRows 'check if same object are derected then skip (not to copy same thing two times)
    
       If RazdelListRow.Range.Cells(i, 1) <> RazdelListRow.Range.Cells(i + 1, 1) Then
            Sales.cbx_category.AddItem RazdelListRow.Range.Cells(i, 1)
            
       
        
       
        End If
    Next RazdelListRow
End Sub

    
Sub FillGroup() 'Filling group in form
    Set ShList = ThisWorkbook.Worksheets("List") 'fillint a row in a "sales" logical table
    Set RazdelListObj = ShList.ListObjects("Category_tb")
    Sales.cbx_group.Clear
    
    
    For Each RazdelListRow In RazdelListObj.ListRows 'check if same object are derected then skip (not to copy same thing two times)
        
        If RazdelListRow.Range(1) = Sales.cbx_category.Value Then
        
            Sales.cbx_group.AddItem RazdelListRow.Range(2)
        End If
    Next RazdelListRow
End Sub


Sub SearchVendorCode() 'Search vendor code of a good at stock
    Dim Cell As Range
    
    Set ShStock = ThisWorkbook.Worksheets("Stock")
    Set StockListObj = ShStock.ListObjects("Stock_tb")
    
    If Sales.txb_vendor.Value = "" Then
        Sales.txb_description.Value = ""
        Sales.txb_availability.Value = ""
        Sales.txb_price.Value = ""
        Sales.cbx_category.Value = ""
        Sales.cbx_group.Value = ""
        Sales.txb_producer.Value = ""
    End If
    
    Set Cell = StockListObj.ListColumns.Item(1).Range.Find(Sales.txb_vendor.Value, LookAt:=xlWhole) 'take a whole number of vendor code
    
    If Not Cell Is Nothing Then 'if vendor code is found
        Sales.txb_description.Value = Cell.Cells(1, 2)
        Sales.txb_availability.Value = Cell.Cells(1, 4)
        Sales.txb_price.Value = Cell.Cells(1, 6)
        Sales.cbx_category.Value = Cell.Cells(1, 7)
        Sales.cbx_group.Value = Cell.Cells(1, 8)
        Sales.txb_producer.Value = Cell.Cells(1, 9)
        
    End If
End Sub

Function PurchasedPrice(VendorCode As Double) As Double 'fuction which copies vedor code from stock takes price and puts into sales sheet to purchased price cell
    Dim Cell As Range
    
    Set ShStock = ThisWorkbook.Worksheets("Stock")
    Set StockListObj = ShStock.ListObjects("Stock_tb")
    
    Set Cell = StockListObj.ListColumns.Item(1).Range.Find(VendorCode, LookAt:=xlWhole) 'take a whole number of vendor code

    If Not Cell Is Nothing Then 'if vendor code is found
         PurchasedPrice = Cell.Cells(1, 3)
         
         
    End If
    



End Function

Function Margin(Price As Double, PurchasedPrice As Double) As Double

    Margin = Price - PurchasedPrice


End Function

Function Profit(Quantity As Double, Margin As Double) As Double

    Profit = Quantity * Margin


End Function


Sub SummCalculate() 'Calculating summ in  Sales sheet
    
    On Error Resume Next

    Dim Price As Double
    
    Dim Quantity As Double
    
    Dim Summ As Double
    
    
    Price = Sales.txb_price.Value
    Quantity = Sales.txb_quantity
    
    Summ = Price * Quantity
    Sales.txb_sum.Value = Summ
    
'    If Err Then
'        MsgBox ("Error occured in value!") & vbCrLf & Err.Description & vbCrLf & Err.Number 'Error message
'    End If

End Sub


Sub AddDelivery() 'filling logistics table
    Dim Cell As Range
    
    Set ShDelivery = ThisWorkbook.Worksheets("Logistics")
    Set DeliveryListObj = ShDelivery.ListObjects("Logistics_tb")
    
    Set Cell = DeliveryListObj.ListColumns.Item(1).Range.Find(Sales.txb_order.Value, LookAt:=xlWhole)
    
    If Not Cell Is Nothing Then
        
        Exit Sub
    
    End If
    
    
    Set DeliveryListRow = DeliveryListObj.ListRows.Add
    
    DeliveryListRow.Range(1) = Sales.txb_order.Value
    If Sales.chbx_pick.Value = True Then
        DeliveryListRow.Range(2) = "PickUp"
        
    Else
    
        DeliveryListRow.Range(2) = Sales.cbx_driver.Value
    
        
    
    End If
    DeliveryListRow.Range(3) = Sales.cbx_city.Value
    DeliveryListRow.Range(4) = Sales.txb_name.Value
    DeliveryListRow.Range(5) = Sales.txb_number.Value
    DeliveryListRow.Range(6) = Sales.txb_delivprice.Value
    
End Sub



Sub FillUser() 'Filling user form
    
    Set ShSettings = ThisWorkbook.Worksheets("Settings")
    Set SetListObj = ShSettings.ListObjects("User_tb")
    
    For Each SetListRow In SetListObj.ListRows
        Pass.cbx_user.AddItem SetListRow.Range(1)
    Next SetListRow

End Sub


Sub CheckUser() 'checking users
    
    Set ShSettings = ThisWorkbook.Worksheets("Settings")
    Set SetListObj = ShSettings.ListObjects("User_tb")
    
    If Pass.cbx_user.Value = "" Then
        Pass.lable_info.Caption = "User has not been selected. Please select a User!"
        Exit Sub
    End If
    If Pass.txb_password.Value = "" Then
        Pass.lable_info.Caption = "Please enter a Password!"
        Exit Sub
    
    End If
    
    For Each SetListRow In SetListObj.ListRows
        If Pass.cbx_user.Value = SetListRow.Range(1) Then
           If Pass.txb_password.Value = CStr(SetListRow.Range(2)) Then
                ShSettings.Range("LastUser") = Pass.cbx_user.Value
                ShSettings.Range("LastDate") = VBA.Now
                Call AccessCheck(Pass.cbx_user.Value)
                Unload Pass
                Exit Sub
            Else
                Pass.lable_info.Caption = "Invalid Password!"
                Exit Sub
            End If
            
                
            
        
        End If
    Next SetListRow
    


End Sub

Sub AccessCheck(User As String)

    Set ShSettings = ThisWorkbook.Worksheets("Settings")
    Set SetListObj = ShSettings.ListObjects("User_tb")
     For Each SetListRow In SetListObj.ListRows
            If User = SetListRow.Range(1) Then
                If SetListRow.Range(3) = 1 Then
                
                    ThisWorkbook.Worksheets("Settings").Visible = xlSheetVisible
                    
                
                End If
            End If
     Next SetListRow



End Sub


Function PassSales() 'password for sales table changing
    PassSales = ThisWorkbook.Worksheets("Settings").Range("Password_sales")


End Function


Sub DownloadClient() 'ListBox list of customers
    Set ShClients = ThisWorkbook.Worksheets("Customers")
    Set ClientsListObj = ShClients.ListObjects("Customers_tb")
    
    SMSForm.listbox_clientSMS.Clear
    For Each ClientsListRow In ClientsListObj.ListRows
    
        SMSForm.listbox_clientSMS.AddItem ClientsListRow.Range(1)
        SMSForm.listbox_clientSMS.List(SMSForm.listbox_clientSMS.ListCount - 1, 1) = ClientsListRow.Range(2)
        SMSForm.listbox_clientSMS.ColumnWidths = "150;150;150"
        SMSForm.listbox_clientSMS.List(SMSForm.listbox_clientSMS.ListCount - 1, 2) = ClientsListRow.Range(3)
        SMSForm.listbox_clientSMS.ColumnWidths = "150;150;150"
        SMSForm.listbox_clientSMS.List(SMSForm.listbox_clientSMS.ListCount - 1, 3) = ClientsListRow.Range(4)
        SMSForm.listbox_clientSMS.ColumnWidths = "150;150;150"
        SMSForm.listbox_clientSMS.List(SMSForm.listbox_clientSMS.ListCount - 1, 4) = ClientsListRow.Range(5)
        
    Next ClientsListRow
    
    



End Sub




Sub SMSSender()


    Dim TextSms As String 'main sms text
    Dim NumberOrder As String 'order number
    Dim NameClient As String ' Customer name
    Dim PriceOrder As String 'Price of oder
    Dim Telephone As String ' MObile numberr
    Dim DateDelivery As Date 'Date
    Dim TimeDelivery As String 'Time
    
    
    Set ShSmsTemplate = ThisWorkbook.Worksheets("SMS-Settings")
    Set SmsTempListObj = ShSmsTemplate.ListObjects("SmsTitle_tb")
    
    Set ShClients = ThisWorkbook.Worksheets("Customers")
    Set ClientsListObj = ShClients.ListObjects("Customers_tb")
    
    NumberOrder = SMSForm.txb_orderSMS
    NameClient = SMSForm.txb_nameSMS
    PriceOrder = SMSForm.txb_priceSMS
    Telephone = SMSForm.txb_telSMS
    TimeDelivery = SMSForm.txb_timeSMS
    
    'formatting mesage text from smstitle_tb
    For Each SmsTempListRow In SmsTempListObj.ListRows
        If SMSForm.chbx_pickupSMS.Value = True Then
            TextSms = NameClient & ", " & SmsTempListRow.Range(2) & ". " & SmsTempListRow.Range(3) & ": " & PriceOrder & ". " & "Pick-Up" & ". " & SmsTempListRow.Range(5)
            
        Else
            TextSms = NameClient & ", " & SmsTempListRow.Range(2) & ". " & SmsTempListRow.Range(3) & ": " & PriceOrder & ". " & SmsTempListRow.Range(4) & ": " & TimeDelivery & ". " & SmsTempListRow.Range(5)
        
        End If
        
        
    
    Next SmsTempListRow
    

    
    If SMSForm.txb_telSMS = Telephone Then
        Dim OutApp As Object
        Dim OutMail As Object
    
        Set OutApp = CreateObject("Outlook.Application")
        Set OutMail = OutApp.CreateItem(0)
    
        On Error Resume Next
        With OutMail
            .To = Telephone
            .CC = ""
            .BCC = ""
            .Subject = "Amazon.com"
            .Body = TextSms
            .Send
            
        End With
        
        Set ClientsListRow = ClientsListObj.ListRows.Add
        ClientsListRow.Range(1) = NumberOrder
        ClientsListRow.Range(2) = NameClient
        ClientsListRow.Range(3) = Telephone
        ClientsListRow.Range(4) = PriceOrder
        ClientsListRow.Range(5) = TimeDelivery
        Call DownloadClient
        On Error GoTo 0
        
        
    
        Set OutMail = Nothing
        Set OutApp = Nothing

    
    
    End If
    
    
    
        
    
End Sub

Sub LastOrderForSMS() 'fill last order for SMSForm nottification

    Dim LastRow As Long
    Dim LastOrder As Long


    Set ShClients = ThisWorkbook.Worksheets("Customers")
    Set ClientsListObj = ShClients.ListObjects("Customers_tb")

    
    LastRow = ClientsListObj.ListRows.Count

    LastOrder = ClientsListObj.Range.Cells(LastRow + 1, 1)


    SMSForm.txb_orderSMS = LastOrder

End Sub



