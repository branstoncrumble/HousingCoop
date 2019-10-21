Attribute VB_Name = "Module1"
Sub FormatRentalStatement()
Dim fName As String
Dim FileToOpen As FileDialog
Dim wkbkRentalStatements As Workbook
Dim wkshtRentStatement As Worksheet
Dim rngTable As Range
Dim iLastColumn As Integer

    Set FileToOpen = Application.FileDialog(msoFileDialogOpen)
    
    FileToOpen.Show
    fName = FileToOpen.SelectedItems(1)

    Set wkbkRentalStatements = Workbooks.Open(fName)
    wkbkRentalStatements.Activate
    
    
    
    For Each wkshtRentStatement In wkbkRentalStatements.Worksheets
    
        wkshtRentStatement.Select
        Select Case wkshtRentStatement.Name
            Case "Rent_Balance_Summaries":
                'select all data and format
                Sheets("Rent_Balance_Summaries").Select
                Range("A1:" & Range("C2").End(xlDown).Address).Select
                
                'pivot data
                ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
                    Selection.Address, Version:=6).CreatePivotTable _
                    TableDestination:="Rent_Balance_Summaries!R1C5", TableName:="PivotTable1", _
                    DefaultVersion:=6
                Cells(1, 5).Select
                With ActiveSheet.PivotTables("PivotTable1").PivotFields("Period")
                    .Orientation = xlColumnField
                    .Position = 1
                End With
                With ActiveSheet.PivotTables("PivotTable1").PivotFields("Payee")
                    .Orientation = xlRowField
                    .Position = 1
                End With
                ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
                    "PivotTable1").PivotFields("Amnt"), "Sum of Amnt", xlSum
                
                'copy pivot
                Range("E1:" & Range(Range("F2").End(xlToRight).Address).End(xlDown).Address).Select
                Selection.Copy
                
                'paste copt of pivot values
                Range("Q1").Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                    
                'delete pivot
                Columns("A:p").Select
                Range("N1").Activate
                Application.CutCopyMode = False
                Selection.Delete Shift:=xlToLeft
                Rows("1:1").Select
                Selection.Delete Shift:=xlUp
                
                'format as table
                Set rngTable = Range("A1:" & Range(Range("A1").End(xlToRight).Address).End(xlDown).Address)
                rngTable.Select
                ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = _
                    "Table1"
                
                ActiveSheet.ListObjects("Table1").TableStyle = "TableStyleMedium3"
                rngTable.Select
                Selection.NumberFormat = "$#,##0_);[Red]($#,##0)"
                Range("A1").Select
 
            Case "Bank_Balance_Summaries":
                'select all data and format
                Sheets("Bank_Balance_Summaries").Select
                Range("A1:" & Range("D2").End(xlDown).Address).Select
                
                'create pivot
                ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
                    Selection.Address, Version:=6).CreatePivotTable _
                    TableDestination:="Bank_Balance_Summaries!R1C6", TableName:="PivotTable1", _
                    DefaultVersion:=6
                
                Cells(1, 6).Select
                With ActiveSheet.PivotTables("PivotTable1").PivotFields("period")
                    .Orientation = xlColumnField
                    .Position = 1
                End With
                With ActiveSheet.PivotTables("PivotTable1").PivotFields("Category")
                    .Orientation = xlRowField
                    .Position = 1
                End With
                With ActiveSheet.PivotTables("PivotTable1").PivotFields("Sub_Category")
                    .Orientation = xlRowField
                    .Position = 2
                End With
                ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
                    "PivotTable1").PivotFields("SumOfAmount"), "Sum of SumOfAmount", xlSum
                ActiveSheet.PivotTables("PivotTable1").PivotFields("Sub_Category").LayoutForm _
                    = xlTabular
                With ActiveSheet.PivotTables("PivotTable1").PivotFields("Sub_Category")
                    .LayoutForm = xlOutline
                    .LayoutCompactRow = False
                End With
                ActiveSheet.PivotTables("PivotTable1").PivotFields("Category"). _
                    LayoutCompactRow = False
                    
                'copy pivot
                Set rngTable = Range("F1:" & Range(Range("H2").End(xlToRight).Address).End(xlDown).Address)
                rngTable.Select
                Selection.Copy
                Range("H2").End(xlToRight).Select
                iLastColumn = Selection.Column
                ActiveSheet.Cells(1, iLastColumn + 2).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                
                'delete pivot & raw data
                Range(ActiveSheet.Cells(1, 1), ActiveSheet.Cells(1, iLastColumn + 1)).EntireColumn.Select
                Application.CutCopyMode = False
                Selection.Delete Shift:=xlToLeft
                
                'select all data and format
                Rows("1:1").EntireRow.Delete
                Set rngTable = Range("A1:" & Range(Range("C1").End(xlToRight).Address).End(xlDown).Address)
                rngTable.Select
                'Add (SourceType, Source, LinkSource, XlListObjectHasHeaders, Destination, TableStyleName)
                ActiveSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=rngTable, XlListObjectHasHeaders:=xlYes).Name = _
                    "Table9"
                Range("Table9[#All]").Select
                ActiveSheet.ListObjects("Table9").TableStyle = "TableStyleMedium3"
                Rows("2:2").Select
                Selection.Copy
                Rows("13:13").Select
                Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                    SkipBlanks:=False, Transpose:=False
                Application.CutCopyMode = False
                'Range("Table9[[#Headers],[Row Labels]]").Select
            
            'worksheet is a rent statement
            Case Else:
            
                'make the formula in col F work
                Range("F2").Select
                TheFormula = ActiveCell.Value
                ActiveCell.Formula = TheFormula
                Range("F2").Select
                Selection.AutoFill Destination:=Range("F2:" & Range("F2").End(xlDown).Address)
                
                'select all data and format
                Range("A1:" & Range("F2").End(xlDown).Address).Select
                With Selection.Font
                    .Name = "Verdana"
                    .Size = 11
                    .Strikethrough = False
                    .Superscript = False
                    .Subscript = False
                    .OutlineFont = False
                    .Shadow = False
                    .Underline = xlUnderlineStyleNone
                    .ThemeColor = xlThemeColorLight1
                    .TintAndShade = 0
                    .ThemeFont = xlThemeFontNone
                End With
                ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$F$63"), , xlYes).Name = _
                    "Table2"
                ActiveSheet.ListObjects("Table2").TableStyle = "TableStyleMedium3"
                Columns("A:F").Select
                Columns("A:F").EntireColumn.AutoFit
                Columns("A:A").Select
                Selection.ColumnWidth = 18.14
                With Selection
                    .HorizontalAlignment = xlLeft
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                    .MergeCells = False
                End With
            
                'insert header
                Rows("1:4").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                
                'company name
                Range("A1").Select
                ActiveCell.FormulaR1C1 = "Blackcurrent Housing Co-operative"
                
                With Selection.Font
                    .Name = "Verve"
                    .Size = 11
                End With
                Range("A1:D1").Select
                Selection.Merge
                With Selection
                    .HorizontalAlignment = xlLeft
                    .VerticalAlignment = xlCenter
                    .MergeCells = True
                End With
                Selection.Font.Size = 18
                
                'rental statement title, members name, date range
                Range("A2:C2").Select
                Selection.Merge
                With Selection
                    .HorizontalAlignment = xlLeft
                    .VerticalAlignment = xlCenter
                    .WrapText = True
                    .MergeCells = True
                End With
                    With Selection.Font
                    .Name = "Constantia"
                    .Size = 12
                    .Bold = True
                End With
            
                ActiveCell.FormulaR1C1 = _
                    "=""Rental Statement - ""&R[4]C[1]&CHAR(13)&CHAR(10)&TEXT(MIN(R[4]C:R[204]C),""d mmm yy "")&""  To  ""&TEXT(MAX(R[4]C:R[204]C),""d mmm yy "")"
                Rows("2:2").RowHeight = 33.75
                
                'add total
                Range("E2").Select
                ActiveCell.Formula = "=F6"
                With Selection.Font
                    .Name = "Constantia"
                    .Size = 14
                    .Bold = True
                End With
                Selection.NumberFormat = "$#,##0_);[Red]($#,##0)"
                Range("F2").Select
        End Select
    Next wkshtRentStatement
       
End Sub

Sub CreateRentalStatement()

    Dim RentalStatements As Worksheet
    Dim Rates   As Range
    Dim LastMonday, StartDate, TxDate As Date
    Dim Row As Range
    
    Sheet1.Activate
    ActiveSheet.Range("A3").Select
    Selection.End(xlDown).Select
    Selection.End(xlToRight).Select
    
    Set Rates = Sheet1.Range("$A$3:" & Sheet1.Range(Range("A3").End(xlDown).Address).End(xlToRight).Address)
    
    'MsgBox Rates.Address
    
    Set RentalStatements = Sheet2
    RentalStatements.Activate
    
    'create a date range for transactions to be created
    LastMonday = (Now() - Weekday(Now(), vbMonday) + 1)
    StartDate = RentalStatements.Range("StartDate")
    
    'move to the first available line on the rental statements
    'RentalStatements.Range("B4").Select
    'Selection.End(xlDown).Select
    'RentalStatements.Range(Selection.Row + 1 & ":" & Selection.Row + 1).Select
    
    'delete previous data
    RentalStatements.Range("A5").Select
    Do
        Selection.Cells(1, 1).EntireRow.Delete
        RentalStatements.Range("A5").Select
    Loop While Len(Selection) <> 0
    
    For TxDate = StartDate To LastMonday
        If Format(TxDate, "dddd") = "Monday" Then
            For Each Row In Rates.Rows
                If TxDate > Row.Cells(1, 4) And TxDate < Row.Cells(1, 5) Then
                '=IF(B6<DATEVALUE("01-" & TEXT(B6,"mmm-yyyy"))-WEEKDAY(  DATEVALUE("01-" & TEXT(B6,"mmm-yyyy")),15)+21,TEXT(B6,"mm-yy"),TEXT(B6+28,"mmm-yy"))
                    If TxDate < DateValue("01-" & Format(TxDate, "mmm-yyyy")) - Weekday(DateValue("01-" & Format(TxDate, "mmm-yyyy")), vbFriday) + 21 Then
                        Selection.Cells(1, 1) = "'" & Format(TxDate, "yy-mm")
                    Else
                        Selection.Cells(1, 1) = "'" & Format(TxDate + 28, "yy-mm") 'period col
                    End If
                    Selection.Cells(1, 2) = TxDate 'date column
                    Selection.Cells(1, 3) = Row.Cells(1, 1) 'payee column
                    Selection.Cells(1, 4) = "Invoice"   'category
                    Selection.Cells(1, 5) = Row.Cells(1, 2) ' sub_category
                    Selection.Cells(1, 6) = Row.Cells(1, 3) 'amount
                    Selection.Cells(1, 7) = Format(TxDate, "0") & Row.Cells(1, 1) & Row.Cells(1, 2) & Row.Cells(1, 3) 'tx_id
                    RentalStatements.Range(Selection.Row + 1 & ":" & Selection.Row + 1).Select
                End If
            Next Row
        End If
    Next TxDate
    
    'define the name for the linked table and save
    Range("$A$4:" & Selection.Cells(1, 7).Address).Select
    ThisWorkbook.Names.Add "MEMBERS_TX", "='Rental Statement'!" & Selection.Address, True
    ThisWorkbook.Save
    
End Sub

Sub CreateBankStatement()

    Dim StatementDownload, BankStatement As Worksheet
    Dim EndOfDownload As Boolean
    Dim InputRow As Range
    
    EndOfDownload = False
    
    
    
    Set StatementDownload = Worksheets(Sheet1.Range("I2").Value)
    Set BankStatement = Sheet4
    
    'move to the first available line on the bank statements
    BankStatement.Activate
    'BankStatement.Range("A1").Select
    'Selection.End(xlDown).Select
    'BankStatement.Range(Selection.Row + 1 & ":" & Selection.Row + 1).Select
    
    'delete previous data
    BankStatement.Range("A2").Select
    Do
        Selection.Cells(1, 1).EntireRow.Delete
        BankStatement.Range("A2").Select
    Loop While Len(Selection) <> 0
    Set InputRow = Selection.Range("2:2")
    
    'move to statement download and select 1st row to begin procesing
    StatementDownload.Activate
    StatementDownload.Range("1:1").Select
    Do
        'col H contain a balance. Testing a value means page hdr/ftr ignored
        If VBA.IsNumeric(Selection.Cells(1, 8)) And Selection.Cells(1, 8) > 0 Then
            
            'add values from download sheet to bank statement
            '=IF(B6<DATEVALUE("01-" & TEXT(B6,"mmm-yyyy"))-WEEKDAY(  DATEVALUE("01-" & TEXT(B6,"mmm-yyyy")),15)+21,TEXT(B6,"mm-yy"),TEXT(B6+28,"mmm-yy"))
                If Selection(1, 1) < DateValue("01-" & Format(Selection(1, 1), "mmm-yyyy")) - Weekday(DateValue("01-" & Format(Selection(1, 1), "mmm-yyyy")), vbFriday) + 21 Then
                    InputRow.Cells(1, 1) = "'" & Format(Selection(1, 1), "yy-mm")
                Else
                    InputRow.Cells(1, 1) = "'" & Format(Selection(1, 1) + 28, "yy-mm") 'period col
                End If
            InputRow.Cells(1, 2) = Selection(1, 1) 'date column
            InputRow.Cells(1, 3) = Selection(1, 2) 'desc column
            InputRow.Cells(1, 4) = Selection(1, 3) 'bank ref
            InputRow.Cells(1, 5) = Selection(1, 4) 'customer ref
            InputRow.Cells(1, 6) = Selection(1, 5) - Selection(1, 6) ' amount
            InputRow.Cells(1, 7) = Selection(1, 7) 'additional info
            
            'split NBS benefit payment
            If Selection(1, 4) = "NBC BENEFITS" Then
                InputRow.Cells(1, 4) = "Laura"
                InputRow.Cells(1, 6) = 260
            End If
            
            'move input to next row
            Set InputRow = BankStatement.Range(InputRow.Row + 1 & ":" & InputRow.Row + 1)
            
            'split NBS benefit payment
            If Selection(1, 4) = "NBC BENEFITS" Then
            '=IF(B6<DATEVALUE("01-" & TEXT(B6,"mmm-yyyy"))-WEEKDAY(  DATEVALUE("01-" & TEXT(B6,"mmm-yyyy")),15)+21,TEXT(B6,"mm-yy"),TEXT(B6+28,"mmm-yy"))
                If Selection(1, 1) < DateValue("01-" & Format(Selection(1, 1), "mmm-yyyy")) - Weekday(DateValue("01-" & Format(Selection(1, 1), "mmm-yyyy")), vbFriday) + 21 Then
                    InputRow.Cells(1, 1) = "'" & Format(Selection(1, 1), "yy-mm")
                Else
                    InputRow.Cells(1, 1) = "'" & Format(Selection(1, 1) + 28, "yy-mm") 'period col
                End If
                InputRow.Cells(1, 2) = Selection(1, 1) 'date column
                InputRow.Cells(1, 3) = Selection(1, 2) 'desc column
                InputRow.Cells(1, 4) = "Ben Jovi"
                InputRow.Cells(1, 5) = Selection(1, 4) 'customer ref
                InputRow.Cells(1, 6) = Selection(1, 5) - 260
                InputRow.Cells(1, 7) = Selection(1, 7) 'additional info
                
                Set InputRow = BankStatement.Range(InputRow.Row + 1 & ":" & InputRow.Row + 1)
            End If
            
        End If
        StatementDownload.Rows(Selection.Row + 1).Select
            
    Loop Until Selection.Cells(1, 1) = "" And Selection.Cells(1, 3) = ""
    Set InputRow = BankStatement.Range(InputRow.Row - 1 & ":" & InputRow.Row - 1)
    
    'define the name for the linked table and save
    BankStatement.Activate
    Range("$A$1:" & InputRow.Cells(1, 7).Address).Select
    ThisWorkbook.Names.Add "BANK_TX", "='Bank Statement'!" & Selection.Address, True
    ThisWorkbook.Save
    
End Sub
