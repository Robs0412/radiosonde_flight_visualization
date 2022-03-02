Attribute VB_Name = "DownloadSaveCSV"
Option Explicit

Sub DownloadSaveCSV()

    Dim Radiosonde      As String
    Dim FirstLetter     As String
    Dim Link            As String
    Dim Request         As Object
    Dim rc              As Variant
    Dim WinHttpReq      As Object
    Dim oStream         As Object

    On Error GoTo Error_Msg

    Application.ScreenUpdating = False

    Radiosonde = InputBox("Enter radiosonde identifier:" & vbLf & _
                 "- without file extension (no '.csv')" & vbLf & _
                 "- without space and tab" & vbLf & _
                 "- only finished flights (archived)" & vbLf & _
                 "- uppercase for letters" & vbLf & vbLf & _
                 "Example: T1231139", "Download CSV from radiosondy.info archive")
    
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
        
        MsgBox "Radiosonde identifier found in radiosondy.info archive." & vbLf & _
               "CSV flight data will be downloaded." & vbLf & _
               "(same folder as this workbook)", vbInformation
        
    Else
        
        MsgBox "Radiosonde identifier not found in radiosondy.info archive," & vbLf & _
               "flight not finished yet, naming errors or bad internet connection?", vbQuestion
        
        Exit Sub
        
    End If
  
    'https://stackoverflow.com/questions/52757325/excel-vba-download-text-file-from-the-internet-that-update-every-5-minutes
    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    With WinHttpReq
        .Open "GET", Link, False
        .send
    End With

    If WinHttpReq.StatusText = "OK" Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.responseBody
        oStream.SaveToFile ThisWorkbook.Path & "\" & Radiosonde & ".csv", 2        '1 = no overwrite, 2 = overwrite
        oStream.Close
    End If
    
    Set WinHttpReq = Nothing

    Range("A1").Select

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

