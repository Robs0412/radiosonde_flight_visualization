Attribute VB_Name = "DescriptionManipulation"
Option Explicit

Sub DescriptionManipulation()
    
    Dim Rows                As Long
    Dim X                   As Long
    Dim Description         As String
    Dim DescriptionLength   As Integer
    Dim Position            As Integer
    Dim Length              As Integer
    Dim ClimbSpeed          As String
    Dim Pressure            As String
    Dim Temperature         As String
    Dim Humidity            As String
    Dim Frequency           As String
    Dim Sonde               As String
    Dim Battery             As String
    Dim TxOff               As String
    Dim TxOff_h             As Integer
    Dim TxOff_m             As Integer
    Dim TxOff_s             As Integer
    Dim PowerUp             As String
    Dim PowerUp_h           As Integer
    Dim PowerUp_m           As Integer
    Dim PowerUp_s           As Integer
    Dim O3Pressure          As String
    Dim O3Temperature       As String
    Dim O3PumpCurrent       As String
    Dim RowBurst            As Long
    
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
    Range("S1").Value = "O3 pressure mPa"
    Range("T1").Value = "O3 analyser temp °C"
    Range("U1").Value = "O3 pump current mA"
    
    'Count number of rows to know number of values
    Rows = WorksheetFunction.CountA(Range("I:I"))
    
    Range("J2:M" & Rows).NumberFormat = "0.0"
    Range("N2:N" & Rows).NumberFormat = "0.00"
    Range("O2:O" & Rows).NumberFormat = "@"
    Range("P2:P" & Rows).NumberFormat = "0.0"
    Range("Q2:Q" & Rows).NumberFormat = "h:mm:ss;@"
    Range("R2:R" & Rows).NumberFormat = "h:mm:ss;@"
    Range("S2:T" & Rows).NumberFormat = "0.0"
    Range("U2:U" & Rows).NumberFormat = "0"
    
    For X = 2 To Rows
        
        Description = Range("I" & X).Value
        DescriptionLength = Len(Description)
        
        'ClimbSpeed
ClimbSpeed:
        Position = InStr(Description, "Clb=")
        If Position = 0 Then GoTo Pressure
        ClimbSpeed = Right(Description, DescriptionLength - Position - 3)
        Position = InStr(ClimbSpeed, "m/s ")
        If Position = 0 Then GoTo Pressure
        ClimbSpeed = Left(ClimbSpeed, Position - 1)
        If ClimbSpeed = "-9999.0" Then GoTo Pressure
        Range("J" & X).Value = ClimbSpeed
        
        'Pressure
Pressure:
        Position = InStr(Description, " p=")
        If Position = 0 Then GoTo Temperature
        Pressure = Right(Description, DescriptionLength - Position - 2)
        Position = InStr(Pressure, "hPa ")
        If Position = 0 Then GoTo Temperature
        Pressure = Left(Pressure, Position - 1)
        If Pressure = "-1.0" Then GoTo Temperature
        Range("K" & X).Value = Pressure
        
        'Temperature
Temperature:
        Position = InStr(Description, " t=")
        If Position = 0 Then GoTo Humidity
        Temperature = Right(Description, DescriptionLength - Position - 2)
        Position = InStr(Temperature, "C ")
        If Position = 0 Then GoTo Humidity
        Temperature = Left(Temperature, Position - 1)
        If Temperature = "-273.0" Then GoTo Humidity
        Range("L" & X).Value = Temperature
        
        'Humidity
Humidity:
        Position = InStr(Description, " h=")
        If Position = 0 Then GoTo Frequency
        Humidity = Right(Description, DescriptionLength - Position - 2)
        Position = InStr(Humidity, "% ")
        If Position = 0 Then GoTo Frequency
        Humidity = Left(Humidity, Position - 1)
        If Humidity = "-1.0" Then GoTo Frequency
        Range("M" & X).Value = Humidity
        
        'Frequency
Frequency:
        Position = InStr(Description, "MHz ")
        If Position = 0 Then GoTo Sonde
        Frequency = Left(Description, Position - 1)
        Length = Len(Frequency)
        Position = InStrRev(Frequency, " 40")
        If Position = 0 Then GoTo Sonde
        Frequency = Right(Frequency, Length - Position)
        Range("N" & X).Value = Frequency
        
        'Sonde type
Sonde:
        Position = InStr(Description, " Type=")
        If Position = 0 Then GoTo Battery
        Sonde = Right(Description, DescriptionLength - Position - 5)
        Position = InStr(Sonde, " ")
        If Position = 0 Then GoTo Battery
        Sonde = Left(Sonde, Position - 1)
        Range("O" & X).Value = Sonde
        
        'Battery
Battery:
        Position = InStr(Description, " batt=")
        If Position = 0 Then GoTo TxOff
        Battery = Right(Description, DescriptionLength - Position - 5)
        Position = InStr(Battery, "V ")
        If Position = 0 Then GoTo TxOff
        Battery = Left(Battery, Position - 1)
        Range("P" & X).Value = Battery
        
        'TxOff
TxOff:
        Position = InStr(Description, " TxOff=")
        If Position = 0 Then GoTo PowerUp
        TxOff = Right(Description, DescriptionLength - Position - 6)
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
        Position = InStr(Description, "powerup=")
        If Position = 0 Then GoTo O3Pressure
        PowerUp = Right(Description, DescriptionLength - Position - 7)
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
        
        'O3 partial pressure
O3Pressure:
        Position = InStr(Description, " o3=")
        If Position = 0 Then GoTo O3Temperature
        O3Pressure = Right(Description, DescriptionLength - Position - 3)
        Position = InStr(O3Pressure, "mPa ")
        If Position = 0 Then GoTo O3Temperature
        O3Pressure = Left(O3Pressure, Position - 1)
        Range("S" & X).Value = O3Pressure
        
        'O3 analyser temperature
O3Temperature:
        Position = InStr(Description, " ti=")
        If Position = 0 Then GoTo O3PumpCurrent
        O3Temperature = Right(Description, DescriptionLength - Position - 3)
        Position = InStr(O3Temperature, "C ")
        If Position = 0 Then GoTo O3PumpCurrent
        O3Temperature = Left(O3Temperature, Position - 1)
        If O3Temperature = "-273.0" Then GoTo O3PumpCurrent
        Range("T" & X).Value = O3Temperature
        
        'O3 pump current
O3PumpCurrent:
        Position = InStr(Description, " Pump=")
        If Position = 0 Then GoTo Skip
        O3PumpCurrent = Right(Description, DescriptionLength - Position - 5)
        Position = InStr(O3PumpCurrent, "mA")
        If Position = 0 Then GoTo Skip
        O3PumpCurrent = Left(O3PumpCurrent, Position - 1)
        Range("U" & X).Value = O3PumpCurrent
        
Skip:
        
    Next X
    
    Columns("I:I").Cut
    Columns("V:V").Insert Shift:=xlToRight
    
    'Burst altitude and time (for charts)
    Range("V1").Value = "Burst altitude m"
    Range("V2").NumberFormat = "0"
    Range("V2").Value = WorksheetFunction.Max(Range("H2:H" & Rows))
    
    Range("W1").Value = "Burst time hh:mm:ss"
    Range("W2").NumberFormat = "h:mm:ss;@"
    RowBurst = Range("H2:H" & Rows).Find(Range("V2").Value).Row        'get row value of moment of burst
    Range("W2").Value = Range("C" & RowBurst).Value
    
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

