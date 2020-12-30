VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If InTable() Then 'if inside table clicked then enable sales form
        ThisWorkbook.Worksheets("Settings").Range("UpdateForm") = 1
        Sales.Show
    End If
End Sub

