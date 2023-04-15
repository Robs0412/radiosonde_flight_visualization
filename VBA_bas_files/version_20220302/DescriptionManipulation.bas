Attribute VB_Name = "DescriptionManipulation"
Option Explicit

Sub DescriptionManipulation()
    
    Dim Rows            As Long
    Dim X               As Long
    Dim Description     As String
    Dim Position        As Integer
    Dim Length          As Integer
    Dim Climb           As String
    Dim Pressure        As String
    Dim Temperature     As String
    Dim Humidity        As String
    Dim Frequency       As String
    Dim Sonde           As String
    Dim Battery         As String
    Dim TxOff           As String
    Dim TxOff_h         As Integer
    Dim TxOff_m         As Integer
    Dim TxOff_s         As Integer
    Dim PowerUp         As String
    Dim PowerUp_h       As Integer
    Dim PowerUp_m       As Integer
    Dim PowerUp_s       As Integer
    Dim RowBurst        As Long
        
    On Error GoTo Error_Msg
        
    Application.ScreenUpdating = False
    
    Range("1:1").NumberFormat = "@"
    Range("1:1").Font.Bold = True
    
    Range("J1").Value = "Climb speed m/s"
    Range("K1").Value = "Pressure hPa"
    Range("L1").Value = "Temperature °C"
    Range("M1").Value = "Humidity %"
    Range("N1").Value = "Frequency MHz"
    Range("O1").Value = "Sonde type"
    Range("P1").Value = "Battery V"
    Range("Q1").Value = "TxOff hh:mm:ss"
    Range("R1").Value = "PowerUp hh:mm:ss"
    
    'Count number of rows to know number of values
    Rows = WorksheetFunction.CountA(Range("I:I"))
    
    Range("J2:M" & Rows).NumberFormat = "0.0"
    
    Range("N2:N" & Rows).NumberFormat = "0.00"
    
    Range("O2:O" & Rows).NumberFormat = "@"
    
    Range("P2:P" & Rows).NumberFormat = "0.0"
    
    Range("Q2:Q" & Rows).NumberFormat = "h:mm:ss;@"
    
    Range("R2:R" & Rows).NumberFormat = "h:mm:ss;@"
    
    For X = 2 To Rows
        
        Description = Range("I" & X).Value
        
        'ClimbSpeed
        Position = InStr(Description, "m/s ") - 1
        If Position = -1 Then GoTo Pressure
        Climb = Left(Description, Position)
        Length = Len(Climb)
        Position = InStrRev(Climb, "Clb=") + 3
        If Position = 3 Then
            Climb = ""
            GoTo Pressure
        End If
        Climb = Right(Climb, Length - Position)
        Range("J" & X).Value = Climb
        
        'Pressure
Pressure:
        Position = InStr(Description, "hPa") - 1
        If Position = -1 Then GoTo Temperature
        Pressure = Left(Description, Position)
        Length = Len(Pressure)
        Position = InStrRev(Pressure, " p=") + 2
        If Position = 1 Then
            Pressure = ""
            GoTo Temperature
        End If
        Pressure = Right(Pressure, Length - Position)
        If Pressure = "-1.0" Then
            GoTo Temperature
        End If
        Range("K" & X).Value = Pressure
        
        'Temperature
Temperature:
        Position = InStr(Description, "C ") - 1
        If Position = -1 Then GoTo Humidity
        Temperature = Left(Description, Position)
        Length = Len(Temperature)
        Position = InStrRev(Temperature, " t=") + 2
        If Position = 2 Then
            Temperature = ""
            GoTo Humidity
        End If
        Temperature = Right(Temperature, Length - Position)
        If Temperature = "-273.0" Then
            GoTo Humidity
        End If
        Range("L" & X).Value = Temperature
        
        'Humidity
Humidity:
        Position = InStr(Description, "% ") - 1
        If Position = -1 Then GoTo Frequency
        Humidity = Left(Description, Position)
        Length = Len(Humidity)
        Position = InStrRev(Humidity, " h=") + 2
        If Position = 2 Then
            Humidity = ""
            GoTo Frequency
        End If
        Humidity = Right(Humidity, Length - Position)
        If Humidity = "-1.0" Then
            GoTo Frequency
        End If
        Range("M" & X).Value = Humidity
        
        'Frequency
Frequency:
        Position = InStr(Description, "MHz ") - 1
        If Position = -1 Then GoTo Sonde
        Frequency = Left(Description, Position)
        Length = Len(Frequency)
        Position = InStrRev(Frequency, " 40")
        If Position = 0 Then
            Frequency = ""
            GoTo Sonde
        End If
        Frequency = Right(Frequency, Length - Position)
        Range("N" & X).Value = Frequency
        
        'Sonde type
Sonde:
        Position = InStr(Description, "Type=") + 5
        If Position = 5 Then GoTo Battery
        Length = Len(Description)
        Sonde = Right(Description, Length - Position + 1)
        Position = InStr(Sonde, " ") - 1
        If Position = -1 Then
            Sonde = ""
            GoTo Skip
        End If
        Sonde = Left(Sonde, Position)
        Range("O" & X).Value = Sonde
        
        'Battery
Battery:
        Position = InStr(Description, "batt=") + 4
        If Position = 4 Then GoTo TxOff
        Length = Len(Description)
        Battery = Right(Description, Length - Position)
        Position = InStr(Battery, "V") - 1
        If Position = -1 Then
            Battery = ""
            GoTo Skip
        End If
        Battery = Left(Battery, Position)
        Range("P" & X).Value = Battery
        
        'TxOff
TxOff:
        Position = InStr(Description, "TxOff=") + 5
        If Position = 5 Then GoTo PowerUp
        Length = Len(Description)
        TxOff = Right(Description, Length - Position)
        Position = InStr(TxOff, "h")
        If Position = 0 Then GoTo TxOff_m        'in case no hour time stamp is present
        If Position > 3 Then GoTo TxOff_m        'in case no hour time stamp is present but other "h" in string
        TxOff_h = CInt(Left(TxOff, Position - 1))
        Length = Len(TxOff)
        TxOff = Right(TxOff, Length - Position)
TxOff_m:
        Position = InStr(TxOff, "m")
        If Position = 0 Then GoTo TxOff_s        'in case no minute time stamp is present
        If Position > 3 Then GoTo TxOff_s        'in case no minute time stamp is present but other "m" in string
        TxOff_m = CInt(Left(TxOff, Position - 1))
        Length = Len(TxOff)
        TxOff = Right(TxOff, Length - Position)
TxOff_s:
        Position = InStr(TxOff, "s ")
        If Position = 0 Then GoTo TxOff_save        'in case no second time stamp is present, save TxOff and write to cell
        If Position > 3 Then GoTo TxOff_save        'in case no second time stamp is present but other "s" in string, save TxOff and write to cell
        TxOff_s = CInt(Left(TxOff, Position - 1))
TxOff_save:
        TxOff = TxOff_h & ":" & TxOff_m & ":" & TxOff_s
        Range("Q" & X).Value = CDate(TxOff)
              
        'PowerUp
PowerUp:
        Position = InStr(Description, "powerup=") + 7
        If Position = 7 Then GoTo Skip
        Length = Len(Description)
        PowerUp = Right(Description, Length - Position)
        Position = InStr(PowerUp, "h")
        If Position = 0 Then GoTo PowerUp_m        'in case no hour time stamp is present
        If Position > 3 Then GoTo PowerUp_m        'in case no hour time stamp is present but other "h" in string
        PowerUp_h = CInt(Left(PowerUp, Position - 1))
        Length = Len(PowerUp)
        PowerUp = Right(PowerUp, Length - Position)
PowerUp_m:
        Position = InStr(PowerUp, "m")
        If Position = 0 Then GoTo PowerUp_s        'in case no minute time stamp is present
        If Position > 3 Then GoTo PowerUp_s        'in case no minute time stamp is present but other "m" in string
        PowerUp_m = CInt(Left(PowerUp, Position - 1))
        Length = Len(PowerUp)
        PowerUp = Right(PowerUp, Length - Position)
PowerUp_s:
        Position = InStr(PowerUp, "s ")
        If Position = 0 Then GoTo PowerUp_save        'in case no second time stamp is present, save PowerUp and write to cell
        If Position > 3 Then GoTo PowerUp_save        'in case no second time stamp is present but other "s" in string, save PowerUp and write to cell
        PowerUp_s = CInt(Left(PowerUp, Position - 1))
PowerUp_save:
        PowerUp = PowerUp_h & ":" & PowerUp_m & ":" & PowerUp_s
        Range("R" & X).Value = CDate(PowerUp)
        
Skip:
        
    Next X
   
    Columns("I:I").Cut
    Columns("S:S").Insert Shift:=xlToRight
    
    'Burst altitude and time (for charts)
    Range("S1").Value = "Burst altitude m"
    Range("S2").NumberFormat = "0"
    Range("S2").Value = WorksheetFunction.Max(Range("H2:H" & Rows))
    
    Range("T1").Value = "Burst time hh:mm:ss"
    Range("T2").NumberFormat = "h:mm:ss;@"
    RowBurst = Range("H2:H" & Rows).Find(Range("S2").Value).Row   'get row value of moment of burst
    Range("T2").Value = Range("C" & RowBurst).Value
    
    'Clear ClipBoard
    Application.CutCopyMode = False
    
    'Columns("A:S").EntireColumn.AutoFit  --> done in ProcessImportedSheets
    
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



