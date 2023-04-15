Attribute VB_Name = "DeleteAllWorksheetsExceptActive"
Option Explicit

'DeleteAllWorksheetsExceptActive

Sub DeleteAllWorksheetsExceptActive()
    'https://www.excelhowto.com/macros/delete-all-worksheets-except-active-one/
    
    Dim ws          As Worksheet
    
    On Error GoTo Error_Msg
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> ThisWorkbook.ActiveSheet.Name Then
            ws.Delete
        End If
    Next ws
    
    Range("A1").Select
    Application.DisplayAlerts = True
    'Application.ScreenUpdating = True   not needed, will be active anyway after macro is finished
    
Exit Sub
    

Error_Msg:

    Application.DisplayAlerts = True

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

