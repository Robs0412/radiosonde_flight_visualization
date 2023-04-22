Attribute VB_Name = "OpenMultipleFiles"
Option Explicit

'https://riptutorial.com/excel-vba/example/29904/open-file-dialog---multiple-files

Sub OpenMultipleFiles()
    
    Dim FD              As FileDialog
    Dim FileChosen      As Long
    Dim i               As Long
    Dim FSO             As Variant
    Dim FileName        As String
    Dim csv_num         As Long
    
    On Error GoTo Error_Msg
    
    Application.ScreenUpdating = False
    
    'Check how many csv file path are listed
    csv_num = WorksheetFunction.CountA(Range("A:A")) - 1        '-1 because of headline
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Set FD = Application.FileDialog(msoFileDialogFilePicker)
    
    With FD
        .Title = "Select multiple files (format type: radiosondy.info CSV)"
        .InitialFileName = ActiveWorkbook.Path        'Set Default Location to the Active Workbook Path, optional:  Environ$("USERPROFILE") & "\Desktop\"
        .InitialView = msoFileDialogViewList
        .ButtonName = "Add selected CSV files"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "radiosondy.info CSV", "*.csv"
    End With
    
    FileChosen = FD.Show
    
    If FileChosen = -1 Then
        
        'open each of the files chosen
        For i = 1 To FD.SelectedItems.Count
            
            'Writes in column A complete folder-file path
            Range("A" & i + 1 + csv_num).Value = FD.SelectedItems(i)
            
            'Writes in column B name of file only
            FileName = FSO.getFileName(FD.SelectedItems(i))
            Range("B" & i + 1 + csv_num).Value = FileName
            
        Next i
        
    End If
    
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

