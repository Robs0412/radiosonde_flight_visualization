Attribute VB_Name = "ProcessImportedSheets"
Option Explicit

Sub ProcessImportedSheets()
    
    'https://support.microsoft.com/en-us/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0
    
    Dim WS_Count        As Long
    Dim Counter         As Long
    Dim WS_Name         As String
    Dim Click           As Integer
    Dim step            As Long
    Dim ws_cur          As Worksheet
    Dim skipped         As Long
        
    On Error GoTo Error_Msg
        
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    
    ' Set WS_Count equal to the number of worksheets in the active workbook.
    WS_Count = ActiveWorkbook.Worksheets.Count
    
    If WS_Count = 1 Then
        
        Import.Activate
        Range("A1").Select
        
        Application.ScreenUpdating = True
        
        MsgBox "No CSV imported to be processed.", vbInformation
        
        Exit Sub
        
    End If
    
    ' Begin the loop with worksheet 2, as IMPORT will not be manipulated
    For Counter = 2 To WS_Count
        
        Worksheets(Counter).Activate
        WS_Name = ActiveSheet.Name
        
        'Check if sheet has a plausible value in A1, if empty show message and continue next
        If Range("A1").Value = "" Then
       
                Click = MsgBox( _
                    prompt:=WS_Name & " seems to contain non valid data" & vbLf & _
                   "(cell 'A1' checked), skip to next with OK or cancle...", Buttons:=vbOKCancel, Title:="Information")
                If Click = vbOK Then
                    skipped = skipped + 1
                    GoTo Skip
                Else
                    GoTo Cancle
                End If
            
        End If
        
        'Check if sheet has already been manipulated (see cell W2), if so skip this sheet
        If Range("W2").Value = "TRUE" Then

            skipped = skipped + 1
            GoTo Skip
           
        End If

        'https://docs.microsoft.com/de-de/office/vba/api/excel.worksheet
        'Deactivation of automatic calculation of new sheet
        ActiveSheet.EnableCalculation = False
        'Deactivation of format calculation
        ActiveSheet.EnableFormatConditionsCalculation = False

        
        'Call Sub ReverseOrder to swap order
        ReverseOrder.ReverseOrder
        
        'Call Sub DescriptionManipulation to extract information from Description
        DescriptionManipulation.DescriptionManipulation
        
        'Call Sub TimeFromZero to change time format
        TimeFromZero.TimeFromZero
        
        'Call Sub TimeUntilLanding to count down until landing
        TimeUntilLanding.TimeUntilLanding
        
        Range("A1").Value = "Station"
        Range("B1").Value = "Sonde"
        'Range("D1").Value = "Time since start s"  --> already written in TimeFromZero
        'Range("E1").Value = "Time until landing s"  --> already written in TimeUntilLanding
        Range("F1").Value = "Latitude °"
        Range("G1").Value = "Longitude °"
        Range("H1").Value = "Course °"
        Range("I1").Value = "Speed km/h"
        Range("J1").Value = "Altitude m"
        'all other headers are already declared in DescriptionManipulation
            
        Range("W1").Value = "Sheet processed"
        Range("W2").NumberFormat = "@"
        Range("W2").Value = "TRUE"
        
        Columns("A:W").EntireColumn.AutoFit
      
        'https://docs.microsoft.com/de-de/office/vba/api/excel.worksheet
        'Deactivation of automatic calculation of new sheet
        ActiveSheet.EnableCalculation = True
        'Deactivation of format calculation
        ActiveSheet.EnableFormatConditionsCalculation = True
        
        Range("A1").Select
        
        step = step + 1
        
        'https://stackoverflow.com/questions/5181164/how-can-i-create-a-progress-bar-in-excel-vba
        Application.StatusBar = "Status: " & step & " of " & (WS_Count - 1 - skipped) & " CSVs processed  -  " & Format(step / (WS_Count - 1 - skipped), "0%") & " completed"
        DoEvents
        
Skip:
        
    Next Counter
    
Cancle:
    
    Import.Activate
    Range("A1").Select
    
    Application.Calculation = xlAutomatic
    Application.StatusBar = False
    'Application.ScreenUpdating = True   not needed, will be active anyway after macro is finished
    
Exit Sub
    

Error_Msg:

    MsgBox prompt:="Something went wrong in " & WS_Name & "!" & vbLf & _
           "Sorry - debugging seems to be necessary..." & vbLf & vbLf & _
           "Alt+F11 for debugging features.", Buttons:=vbInformation, Title:="Information"
    
    Import.Activate
    Range("A1").Select
    
    Application.Calculation = xlAutomatic
    Application.StatusBar = False
    'Application.ScreenUpdating = True   not needed, will be active anyway after macro is finished
    
    End
    
End Sub

