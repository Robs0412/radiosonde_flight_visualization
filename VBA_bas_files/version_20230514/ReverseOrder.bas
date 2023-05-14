Attribute VB_Name = "ReverseOrder"
Option Explicit

Sub ReverseOrder()
    
    On Error GoTo Error_Msg
    
    Application.ScreenUpdating = False
    
    'Insert column left of A
    'https://excelchamps.com/vba/insert-column/
    Range("A1").EntireColumn.Insert
    
    'AutoFillHelpNumber
    Range("A1").Value = 0
    Range("A1").AutoFill Destination:=Range("A1:A" & (Range("B" & Rows.Count).End(xlUp).Row)), Type:=xlFillSeries
    
    'Sort descending order
    'https://trumpexcel.com/sort-data-vba/
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range("A1"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange ActiveSheet.UsedRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("A1").EntireColumn.Delete
    
    Range("A1").Select
    
    'Application.ScreenUpdating = True   not needed, will be active anyway after macro is finished
    
    Exit Sub
    
Error_Msg:
    
    MsgBox prompt:="Something went wrong!" & vbLf & _
           "Sorry - debugging seems to be necessary..." & vbLf & vbLf & _
           "Alt+F11 for debugging features.", Buttons:=vbInformation, Title:="Information"
    
    Import.Activate
    Range("A1").Select
    
    Application.Calculation = xlAutomatic
    Application.StatusBar = False
    'Application.ScreenUpdating = True   not needed, will be active anyway after macro is finished
    
    End
    
End Sub

