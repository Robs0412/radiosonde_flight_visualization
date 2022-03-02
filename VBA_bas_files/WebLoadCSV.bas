Attribute VB_Name = "WebLoadCSV"
Option Explicit

Sub WebLoadCSV()
    
    Dim Radiosonde      As String
    Dim FirstLetter     As String
    Dim Link            As String
    Dim csv_num         As Long
    Dim Request         As Object
    Dim rc              As Variant
    
    Application.ScreenUpdating = False
    
    On Error GoTo Error_Msg
    
    'Check how many csv file path are listed
    csv_num = WorksheetFunction.CountA(Range("A:A")) - 1        '-1 because of headline
    
    Radiosonde = InputBox("Enter radiosonde identifier:" & vbLf & _
                 "- without file extension (no '.csv')" & vbLf & _
                 "- without space and tab" & vbLf & _
                 "- only finished flights (archived)" & vbLf & _
                 "- uppercase for letters" & vbLf & vbLf & _
                 "Example: T1231139", "Web-Load CSV from radiosondy.info archive")
    
    If Radiosonde = "" Then Exit Sub
    
    FirstLetter = Left(Radiosonde, 1)
    
    Link = "https://radiosondy.info/sonde-data/CSV/" & FirstLetter & "/" & Radiosonde & ".csv"
    
    'Check if CSV Link is available
    'https://www.mrexcel.com/board/threads/check-if-url-exists-is-so-then-return-true.567315/
    'https://stackoverflow.com/questions/25428611/vba-check-if-file-from-website-exists/25428811
    
    Set Request = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    With Request
        .Open "GET", Link, False
        .send
        rc = .StatusText
    End With
    Set Request = Nothing
    
    If rc = "OK" Then
        
        MsgBox "Radiosonde identifier found in radiosondy.info archive," & vbLf & _
               "link to CSV will be added to file paths.", vbInformation
        
    Else
        
        MsgBox "Radiosonde identifier not found in radiosondy.info archive," & vbLf & _
               "flight not finished yet, naming errors or bad internet connection?", vbQuestion
        
        Exit Sub
        
    End If
    
    'Writes in column A complete folder-file path
    Range("A" & csv_num + 2).Value = Link
    
    'Writes in column B name of file only
    Range("B" & csv_num + 2).Value = Radiosonde & ".csv"
   
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


