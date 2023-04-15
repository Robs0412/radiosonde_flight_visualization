Attribute VB_Name = "SearchSheet"
Option Explicit

'https://www.extendoffice.com/documents/excel/3158-excel-search-by-sheet-name.html
Sub SearchSheet()
    
    Dim xName       As String
    Dim xFound      As Boolean
    
    On Error GoTo Error_Msg
    
    Application.ScreenUpdating = False
    
    xName = InputBox("Enter sheet name to find in workbook:", "Sheet search")
    
    If xName = "" Then Exit Sub
    
    On Error Resume Next
    
    ActiveWorkbook.Sheets(xName).Select
    
    xFound = (Err = 0)
    
    On Error GoTo 0
    
    If xFound Then
        
        MsgBox xName & " has been found and will be selected."
        
    Else
        
        MsgBox xName & " could not be found in this workbook."
        
    End If
    
    'Application.ScreenUpdating = True   not needed, will be active anyway after macro is finished
    
Exit Sub

Error_Msg:

    MsgBox prompt:="Something went wrong!" & vbLf & _
           "Sorry - debugging seems to be necessary..." & vbLf & vbLf & _
           "Alt+F11 for debugging features.", Buttons:=vbInformation, Title:="Information"
    
    Import.Activate
    Range("A1").Select
    
    'Application.Calculation = xlAutomatic
    'Application.StatusBar = False
    'Application.ScreenUpdating = True   not needed, will be active anyway after macro is finished
    
    End
    
End Sub
