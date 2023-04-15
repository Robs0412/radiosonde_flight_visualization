Attribute VB_Name = "DeleteFilePath"
Option Explicit

'DeleteFilePath

Sub DeleteFilePath()

    On Error GoTo Error_Msg
    
    Application.ScreenUpdating = False
    
    Columns("A:B").ClearContents
    Range("A1").Value = "File paths"
    Range("B1").Value = "File name"
    Range("A1").Select
    
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
