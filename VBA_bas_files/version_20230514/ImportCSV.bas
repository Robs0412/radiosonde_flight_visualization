Attribute VB_Name = "ImportCSV"
Option Explicit

'ImportCSV
Sub ImportCSV()
    
    Dim file            As Object
    Dim files           As Range
    Dim FileName        As String
    Dim ws_cur          As Worksheet
    Dim Sheet           As Worksheet
    Dim TempSheetName   As String
    Dim WorksheetExists As Boolean
    Dim sheets_add      As Long
    Dim csv_num         As Long
    Dim Click           As Integer
    Dim skipped         As Long
    
    On Error GoTo Error_Msg
    
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    
    'Check how many csv file path are listed - if 0 then message
    csv_num = WorksheetFunction.CountA(Range("A:A")) - 1        '-1 because of headline
    
    If csv_num = 0 Then
        
        MsgBox "No folder or link for CSV import added?", vbQuestion
        
        GoTo EndSub
        
    End If
    
    'Check if just one csv is getting imported and if so go to Jump_To
    'Problem to solve: if just one file is in the list, "Range(Selection, Selection.End(xlDown)).Select" will select whole column
    'and produce an extra empty sheet
    If csv_num = 1 Then
        
        Range("A2").Select
        Set files = Selection
        
        GoTo Jump_To
        
    End If
    
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Set files = Selection
    
Jump_To:
    
    For Each file In files
        
        'https://stackoverflow.com/questions/27923926/file-name-without-extension-name-vba/27924854
        'https://www.mrexcel.com/board/threads/getting-filename-without-extension.17181/
        With CreateObject("Scripting.FileSystemObject")
            FileName = .getFileName(file)
            FileName = Replace(FileName, ".csv", "")
            FileName = Replace(FileName, ".CSV", "")
        End With
        
        'Check if File/Sheet Name already exists
        'https://www.howtoexcel.org/vba/how-to-check-if-a-worksheet-exists-using-vba/
        TempSheetName = UCase(FileName)
        
        For Each Sheet In Worksheets
            
            If TempSheetName = UCase(Sheet.Name) Then
                
                skipped = skipped + 1
                GoTo SkipImport
                
            End If
            
        Next Sheet
        
        Sheets.Add After:=Sheets(Sheets.Count)
        
        'count added sheets for status
        sheets_add = sheets_add + 1
        
        Set ws_cur = ActiveWorkbook.Worksheets(ActiveWorkbook.Sheets.Count)
        
        'Deactivation of automatic calculation of new sheet
        ws_cur.EnableCalculation = False
        
        'Deactivation of format calculation
        'https://docs.microsoft.com/de-de/office/vba/api/excel.worksheet
        ws_cur.EnableFormatConditionsCalculation = False
        
        'Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
        'https://www.techonthenet.com/excel/formulas/rnd.php
        'Color Index from 1-56
        ws_cur.Tab.ColorIndex = Int((56 - 1 + 1) * Rnd + 1)
        
        'Naming:  https://www.automateexcel.com/vba/add-and-name-worksheets/
        ws_cur.Name = FileName
        
        Range("A1").Select
        
        'Query - Import csv, all columns formated as text
        'https://docs.microsoft.com/de-de/office/vba/api/excel.querytable
        With ws_cur.QueryTables.Add(Connection:="TEXT;" & file.Value, Destination:=ws_cur.Range("$A$1"))
            .FieldNames = True
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .RefreshStyle = xlOverwriteCells
            .SavePassword = False
            .SaveData = True
            .AdjustColumnWidth = True
            .RefreshPeriod = 0
            .TextFilePromptOnRefresh = False
            .TextFilePlatform = 850
            .TextFileStartRow = 1
            .TextFileParseType = xlDelimited
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
            .TextFileConsecutiveDelimiter = False
            .TextFileTabDelimiter = False
            .TextFileSemicolonDelimiter = True
            .TextFileCommaDelimiter = False
            .TextFileSpaceDelimiter = False
            .TextFileColumnDataTypes = Array(2, 2, 2, 2, 2, 2, 2, 2, 2)
            'https://docs.microsoft.com/en-us/office/vba/api/excel.xlcolumndatatype
            .TextFileTrailingMinusNumbers = True
            .Refresh BackgroundQuery:=False
        End With
        
        'Formatting of certain columns
        'https://docs.microsoft.com/de-de/office/vba/api/excel.range.numberformat
        'https://www.herber.de/forum/archiv/268to272/269487_NumberFormat.html
        'https://stackoverflow.com/questions/20648149/what-are-numberformat-options-in-excel-vba
        'https://www.excelanytime.com/excel/index.php?option=com_content&view=article&id=121:custom-number-formats-date-a-time-formats-in-excel-a-vba-numberformat-property&catid=79&Itemid=475
        'https://stackoverflow.com/questions/18375034/converting-a-certain-range-of-columns-from-text-to-number-format-with-vba/18376758
        Range("D2:E2").Select
        Range(Selection, Selection.End(xlDown)).Select
        With Selection
            .NumberFormat = "0.00000"
            .Value = .Value
        End With
        
        Range("F2:H2").Select
        Range(Selection, Selection.End(xlDown)).Select
        With Selection
            .NumberFormat = "0"
            .Value = .Value
        End With
        
        Range("C2").Select
        Range(Selection, Selection.End(xlDown)).Select
        With Selection
            .NumberFormat = "yyyy-mm-dd  hh:mm:ss"
            .Value = .Value
        End With
        
        Range("A1").Select
        
        'AutoFit of columns
        Columns("A:I").EntireColumn.AutoFit
        
        'Clear ClipBoard
        Application.CutCopyMode = False
        
        'Deletes connections & query from csv import
        'https://www.ms-office-forum.net/forum/showthread.php?t=223639
        While ActiveWorkbook.Connections.Count > 0
            ActiveWorkbook.Connections.Item(1).Delete
        Wend
        
        While ActiveSheet.QueryTables.Count > 0
            ActiveSheet.QueryTables(1).Delete
        Wend
        
        'Deactivation of automatic calculation of new sheet
        ws_cur.EnableCalculation = True
        
        'Deactivation of format calculation
        'https://docs.microsoft.com/de-de/office/vba/api/excel.worksheet
        ws_cur.EnableFormatConditionsCalculation = True
        
        'https://stackoverflow.com/questions/5181164/how-can-i-create-a-progress-bar-in-excel-vba
        Application.StatusBar = "Status: " & sheets_add & " of " & (csv_num - skipped) & " CSVs imported  -  " & Format(sheets_add / (csv_num - skipped), "0%") & " completed"
        DoEvents
        
SkipImport:
        
    Next
    
EndSub:
    
    Import.Activate
    Range("A1").Select
    Application.CutCopyMode = False
    Application.Calculation = xlAutomatic
    Application.StatusBar = False
    'Application.ScreenUpdating = True   not needed, will be active anyway after macro is finished
    
    Exit Sub
    
Error_Msg:
    
    MsgBox prompt:="Something went wrong!" & vbLf & _
           "Sorry - debugging seems to be necessary..." & vbLf & vbLf & _
           "Alt+F11 for debugging features.", Buttons:=vbInformation, Title:="Information"
    
    Import.Activate
    Range("A1").Select
    
    Application.CutCopyMode = False
    Application.Calculation = xlAutomatic
    Application.StatusBar = False
    'Application.ScreenUpdating = True   not needed, will be active anyway after macro is finished
    
    End
    
End Sub

