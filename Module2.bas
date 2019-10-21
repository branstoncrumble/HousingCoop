Attribute VB_Name = "Module2"
Option Compare Database

Function export_rental_statements()
Dim ActiveTenants As Recordset
Dim Exporter As QueryDef
Dim CurrentFolder As String
    
    CurrentFolder = Application.CurrentProject.Path

    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, "Rent_Balance_Summaries", CurrentFolder & "\Rental Statement " & Format(Now(), "mmm-yy") '& ".xls"
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, "Bank_Balance_Summaries", CurrentFolder & "\Rental Statement " & Format(Now(), "mmm-yy") '& ".xls"
    
    Set ActiveTenants = CurrentDb.OpenRecordset("ActiveTenants")
    
    ActiveTenants.MoveFirst
        
    Do
        Set Exporter = New QueryDef
        Exporter.SQL = "SELECT Date, Payee, Category, Sub_Category, Amount, '=IF(C2=""Invoice"",-E2,E2)+F3' as Balance " & _
                                                "FROM ALL_TX " & _
                                                "WHERE (((ALL_TX.[Payee])=""" & ActiveTenants.Fields("Name") & """))" & _
                                                "ORDER BY Date DESC, Category ;"
        Exporter.Name = ActiveTenants.Fields("Name")
        CurrentDb.QueryDefs.Append Exporter
        
        DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, ActiveTenants.Fields("Name"), CurrentFolder & "\Rental Statement " & Format(Now(), "mmm-yy") '& ".xls"
        CurrentDb.QueryDefs.Delete ActiveTenants.Fields("Name")
        
        ActiveTenants.MoveNext
    Loop Until ActiveTenants.EOF
    
    MsgBox "Rental Statement created :- " & vbCrLf & vbCrLf & "\Rental Statement " & Format(Now(), "mmm-yy"), vbOKOnly + vbInformation, "Finished!"
    
End Function

Function SwitchWarnings(WarningsOn As Boolean)
    
    DoCmd.SetWarnings WarningsOn
    
End Function

