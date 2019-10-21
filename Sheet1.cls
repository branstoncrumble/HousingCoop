VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Select Case Target.Cells(1, 1)
        Case "CLICK To RUN": CreateRentalStatement
        Case "CLICK To FORMAT FILE": FormatRentalStatement
        Case "CLICK For Bank Transactions": CreateBankStatement
    End Select
     
End Sub
