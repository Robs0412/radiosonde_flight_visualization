Attribute VB_Name = "TimeFromZero"
Option Explicit

'Convert yyyy-mm-dd hh:mm:ss  to seconds  ->  1s = 1 / (24 * 60 * 60)

Sub TimeFromZero()
    
    Dim Time        As Double
    Dim Rows        As Long
    Dim i           As Long
    
    On Error GoTo Error_Msg
   
    Application.ScreenUpdating = False
    
    'Count number of rows to know number of values
    Rows = WorksheetFunction.CountA(Range("C:C"))
    
    'Insert column left of Latitude
    Range("D1").EntireColumn.Insert
    
    Range("D1").Value = "Time since start s"
    
    'Change cell format to number
    Range("D2:D" & Rows).NumberFormat = "0"
    
    'Set start time value to 0 as new reference point
    Range("D2").Value = 0
    
    'First original time value as master
    Time = Range("C2").Value
    
    'Loop to modify time values, start with third value, as first is headline, second is 0
    For i = 3 To Rows
        
        Range("D" & i).Value = (Range("C" & i).Value - Time) * 86400
       
    Next i
    
    'Columns("A:I").EntireColumn.AutoFit  --> done in ProcessImportedSheets
    
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

