Attribute VB_Name = "ChartDraw"
Option Explicit

Sub ChartDraw()
    
    'https://support.microsoft.com/en-us/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0
    
    Dim WS_Count        As Long
    Dim Counter         As Long
    Dim WS_Name         As String
    Dim Rows            As Long
    Dim Chart           As Long
    Dim Click           As Integer
    Dim Click2          As Integer
    Dim skipped         As Long
    Dim step            As Long
    Dim RowBurst        As Long   'Just needed for DOT in chart TRACK
    Dim sh              As Shape
    
    On Error GoTo Error_Msg
    
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    
    'Set WS_Count equal to the number of worksheets in the active workbook.
    WS_Count = ActiveWorkbook.Worksheets.Count
    
    If WS_Count = 1 Then
        
        Import.Activate
        Range("A1").Select
        
        Application.ScreenUpdating = True
        
        MsgBox "No CSV imported to draw charts.", vbInformation
        
        GoTo Cancle
        
    End If
    
    'Activate preconfigured charts to get copied later, if not done error will occure as chart objects would not be available to be copied
    'https://www.mrexcel.com/board/threads/chartobjects-count-method.595136/
    For Chart = 1 To Import.ChartObjects.Count
       Import.ChartObjects(Import.ChartObjects(Chart).Name).Activate
    Next Chart
    
    ' Begin the loop with worksheet 2, as IMPORT will not be manipulated
    For Counter = 2 To WS_Count
        
        Worksheets(Counter).Activate
        WS_Name = ActiveSheet.Name
        
        'Check if Sheet has a plausible value in A1, if empty show message and continue next
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
                
        'Check if Sheet has already been manipulated (see cell W2), show message and continue next
        If Range("W2").Value = "" Then
       
                If Click2 = 0 Then
       
                    Click2 = MsgBox(prompt:="Some CSVs need to be processed first," & vbLf & _
                    "continue with OK or cancle drawing charts...", Buttons:=vbOKCancel, Title:="Information")
                   
                    If Click2 = vbOK Then
                        skipped = skipped + 1
                        GoTo Skip
                    Else
                        GoTo Cancle
                    End If
                        
                Else
                    skipped = skipped + 1
                    GoTo Skip
                End If
            
        End If
                
        'Check if Sheet has already charts inserted (see cell X2), if so skip this sheet
        If Range("X2").Value = "TRUE" Then
 
            skipped = skipped + 1
            GoTo Skip
            
        End If
                
        'Count number of rows to know number of values for chart
        Rows = WorksheetFunction.CountA(Range("C:C"))
                
        'CLIMB SPEED
        'Copy from IMPORT sheet
        Import.ChartObjects("Chart_Alt_Climb").Copy
        ActiveSheet.Paste Destination:=ActiveSheet.Range("B3")
        ActiveSheet.ChartObjects("Chart_Alt_Climb").Activate
        'Altitude
        ActiveChart.FullSeriesCollection("Altitude").XValues = WS_Name & "!$C$2:$C$" & Rows
        ActiveChart.FullSeriesCollection("Altitude").Values = WS_Name & "!$J$2:$J$" & Rows
        'Climb speed
        ActiveChart.FullSeriesCollection("Climb speed").XValues = WS_Name & "!$C$2:$C$" & Rows
        ActiveChart.FullSeriesCollection("Climb speed").Values = WS_Name & "!$K$2:$K$" & Rows
         'Burst
        ActiveChart.FullSeriesCollection("Burst").XValues = WS_Name & "!$V$2," & WS_Name & "!$V$2"
        ActiveChart.FullSeriesCollection("Burst").Values = Array(-100000, 100000)
        'Size
        ActiveSheet.Shapes("Chart_Alt_Climb").ScaleHeight 1.25, msoFalse, msoScaleFromTopLeft
        
        'PRESSURE
        'Copy from IMPORT sheet
        Import.ChartObjects("Chart_Alt_Press").Copy
        ActiveSheet.Paste Destination:=ActiveSheet.Range("I3")
        ActiveSheet.ChartObjects("Chart_Alt_Press").Activate
        'Altitude
        ActiveChart.FullSeriesCollection("Altitude").XValues = WS_Name & "!$C$2:$C$" & Rows
        ActiveChart.FullSeriesCollection("Altitude").Values = WS_Name & "!$J$2:$J$" & Rows
        'Pressure
        ActiveChart.FullSeriesCollection("Pressure").XValues = WS_Name & "!$C$2:$C$" & Rows
        ActiveChart.FullSeriesCollection("Pressure").Values = WS_Name & "!$L$2:$L$" & Rows
         'Burst
        ActiveChart.FullSeriesCollection("Burst").XValues = WS_Name & "!$V$2," & WS_Name & "!$V$2"
        ActiveChart.FullSeriesCollection("Burst").Values = Array(-100000, 100000)
        'Size
        ActiveSheet.Shapes("Chart_Alt_Press").ScaleHeight 1.25, msoFalse, msoScaleFromTopLeft
        
        'SPEED (WIND)
        'Copy from IMPORT sheet
        Import.ChartObjects("Chart_Alt_Speed").Copy
        ActiveSheet.Paste Destination:=ActiveSheet.Range("B25")
        ActiveSheet.ChartObjects("Chart_Alt_Speed").Activate
        'Altitude
        ActiveChart.FullSeriesCollection("Altitude").XValues = WS_Name & "!$C$2:$C$" & Rows
        ActiveChart.FullSeriesCollection("Altitude").Values = WS_Name & "!$J$2:$J$" & Rows
        'Speed
        ActiveChart.FullSeriesCollection("Speed (wind)").XValues = WS_Name & "!$C$2:$C$" & Rows
        ActiveChart.FullSeriesCollection("Speed (wind)").Values = WS_Name & "!$I$2:$I$" & Rows
         'Burst
        ActiveChart.FullSeriesCollection("Burst").XValues = WS_Name & "!$V$2," & WS_Name & "!$V$2"
        ActiveChart.FullSeriesCollection("Burst").Values = Array(-100000, 100000)
        'Size
        ActiveSheet.Shapes("Chart_Alt_Speed").ScaleHeight 1.25, msoFalse, msoScaleFromTopLeft
        
        'TEMPERATURE
        'Copy from IMPORT sheet
        Import.ChartObjects("Chart_Alt_Temp").Copy
        ActiveSheet.Paste Destination:=ActiveSheet.Range("I25")
        ActiveSheet.ChartObjects("Chart_Alt_Temp").Activate
        'Altitude
        ActiveChart.FullSeriesCollection("Altitude").XValues = WS_Name & "!$C$2:$C$" & Rows
        ActiveChart.FullSeriesCollection("Altitude").Values = WS_Name & "!$J$2:$J$" & Rows
        'Temperature
        ActiveChart.FullSeriesCollection("Temperature").XValues = WS_Name & "!$C$2:$C$" & Rows
        ActiveChart.FullSeriesCollection("Temperature").Values = WS_Name & "!$M$2:$M$" & Rows
         'Burst
        ActiveChart.FullSeriesCollection("Burst").XValues = WS_Name & "!$V$2," & WS_Name & "!$V$2"
        ActiveChart.FullSeriesCollection("Burst").Values = Array(-100000, 100000)
        'Size
        ActiveSheet.Shapes("Chart_Alt_Temp").ScaleHeight 1.25, msoFalse, msoScaleFromTopLeft
        
        'HUMIDITY
        'Copy from IMPORT sheet
        Import.ChartObjects("Chart_Alt_Humi").Copy
        'Sheets(WS_Name).Select
        ActiveSheet.Paste Destination:=ActiveSheet.Range("B47")
        ActiveSheet.ChartObjects("Chart_Alt_Humi").Activate
        'Altitude
        ActiveChart.FullSeriesCollection("Altitude").XValues = WS_Name & "!$C$2:$C$" & Rows
        ActiveChart.FullSeriesCollection("Altitude").Values = WS_Name & "!$J$2:$J$" & Rows
        'Humidity
        ActiveChart.FullSeriesCollection("Humidity").XValues = WS_Name & "!$C$2:$C$" & Rows
        ActiveChart.FullSeriesCollection("Humidity").Values = WS_Name & "!$N$2:$N$" & Rows
         'Burst
        ActiveChart.FullSeriesCollection("Burst").XValues = WS_Name & "!$V$2," & WS_Name & "!$V$2"
        ActiveChart.FullSeriesCollection("Burst").Values = Array(-100000, 100000)
        'Size
        ActiveSheet.Shapes("Chart_Alt_Humi").ScaleHeight 1.25, msoFalse, msoScaleFromTopLeft
        
        'COURSE
        'Copy from IMPORT sheet
        Import.ChartObjects("Chart_Alt_Course").Copy
        ActiveSheet.Paste Destination:=ActiveSheet.Range("I47")
        ActiveSheet.ChartObjects("Chart_Alt_Course").Activate
        'Altitude
        ActiveChart.FullSeriesCollection("Altitude").XValues = WS_Name & "!$C$2:$C$" & Rows
        ActiveChart.FullSeriesCollection("Altitude").Values = WS_Name & "!$J$2:$J$" & Rows
        'Course
        ActiveChart.FullSeriesCollection("Course").XValues = WS_Name & "!$C$2:$C$" & Rows
        ActiveChart.FullSeriesCollection("Course").Values = WS_Name & "!$H$2:$H$" & Rows
         'Burst
        ActiveChart.FullSeriesCollection("Burst").XValues = WS_Name & "!$V$2," & WS_Name & "!$V$2"
        ActiveChart.FullSeriesCollection("Burst").Values = Array(-100000, 100000)
        'Size
        ActiveSheet.Shapes("Chart_Alt_Course").ScaleHeight 1.25, msoFalse, msoScaleFromTopLeft
        
        'TRACK
        'Copy from IMPORT sheet
        Import.ChartObjects("Chart_Track").Copy
        ActiveSheet.Paste Destination:=ActiveSheet.Range("B69")
        ActiveSheet.ChartObjects("Chart_Track").Activate
        'Track
        ActiveChart.FullSeriesCollection("Track").XValues = WS_Name & "!$G$2:$G$" & Rows
        ActiveChart.FullSeriesCollection("Track").Values = WS_Name & "!$F$2:$F$" & Rows
         'Burst
        RowBurst = Range("J2:J" & Rows).Find(Range("U2").Value).Row   'get row value at moment of burst
        ActiveChart.FullSeriesCollection("Burst").XValues = "=" & WS_Name & "!$G$" & RowBurst   '"=" necessary here, unknown why
        ActiveChart.FullSeriesCollection("Burst").Values = "=" & WS_Name & "!$F$" & RowBurst   '"=" necessary here, unknown why
        'Size
        ActiveSheet.Shapes("Chart_Track").ScaleHeight 1.25, msoFalse, msoScaleFromTopLeft
        
        Range("X1").Value = "Charts inserted"
        Range("X2").NumberFormat = "@"
        Range("X2").Value = "TRUE"
        Columns("X").EntireColumn.AutoFit
        
        Range("A1").Select
        
        step = step + 1
       
        'https://stackoverflow.com/questions/5181164/how-can-i-create-a-progress-bar-in-excel-vba
        Application.StatusBar = "Status: " & step & " of " & (WS_Count - 1 - skipped) & " sheets completed with charts  -  " & Format(step / (WS_Count - 1 - skipped), "0%") & " completed"
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

           
    For Each sh In ActiveSheet.Shapes

        sh.Delete
    
    Next sh

    MsgBox prompt:="Error while drawing charts in " & WS_Name & "!" & vbLf & _
           "(Excel struggles sometimes  with drawing chart objects)" & vbLf & vbLf & _
           "Please try 'Draw charts' function again.", Buttons:=vbInformation, Title:="Information"
    
    Import.Activate
    Range("A1").Select
    
    Application.Calculation = xlAutomatic
    Application.StatusBar = False
    'Application.ScreenUpdating = True   not needed, will be active anyway after macro is finished
    
    End
    
End Sub

