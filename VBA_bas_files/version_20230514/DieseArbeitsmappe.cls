VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DieseArbeitsmappe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_Open()
    
    'https://www.herber.de/forum/archiv/1164to1168/1164854_Dezimalzeichen_Komma_Punkt_Tabellenblattabhaengig.html
    'https://docs.microsoft.com/de-de/office/vba/api/excel.worksheet.activate(even)
    
    With Application
        .DecimalSeparator = "."
        .ThousandsSeparator = ","
        .UseSystemSeparators = False
    End With
    
    'Size of window
    'With Application
    '.WindowState = xlNormal
    '.Width = 1200
    '.Height = 800
    'End With
    
    'https://www.teachexcel.com/excel-tutorial/2417/hide-the-entire-excel-interface-ribbon-menu-quick-access-toolbar-status?nav=yt
    ' Hide the Row/Column Headings
    'ActiveWindow.DisplayHeadings = False
    ' Hide the Formula Bar
    'Application.DisplayFormulaBar = False
    ' Hide the Ribbon Menu and Quick Access Toolbar
    'Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"",False)"
    ' Hide the Status Bar (bottom of the window)
    'Application.DisplayStatusBar = False
    
    'Activate preconfigured charts to get copied later, if not done error will occure as chart objects would not be available to be copied
    'https://www.mrexcel.com/board/threads/chartobjects-count-method.595136/
    'Sometimes error will occure anyway, but at second try to load charts it will work - optional pre-load the charts by scrolling down the IMPORT page
    For i = 1 To Import.ChartObjects.Count
        Import.ChartObjects(Import.ChartObjects(i).Name).Activate
    Next i
    
    MsgBox prompt:="In this workbook:" & vbLf & vbLf & _
           "DOT '.' is used as decimal separator!" & vbLf & _
           "After closing, setting is reversed to system default." & vbLf & vbLf & _
           "During use the calculation mode will change between" & vbLf & _
           "automatic & manual - please restore personal setting" & vbLf & _
           "yourself after. Setting won't be restored automatically." & vbLf & _
           "Check 'Formulas' ribbon." & vbLf & vbLf & _
           "Excel user interface can be hidden by 'Excel UI' buttons." & vbLf & vbLf & _
           "The message and function can be deactived in VBA editor:" & vbLf & _
           "Alt+F11 -> 'ThisWorkbook: Workbook_Open()' or 'ExcelUI'", Buttons:=vbInformation, Title:="Information"
    
End Sub

Private Sub Workbook_Close()
    
    Application.UseSystemSeparators = True
    
    'https://www.teachexcel.com/excel-tutorial/2417/hide-the-entire-excel-interface-ribbon-menu-quick-access-toolbar-status?nav=yt
    ' Show the Row/Column Headings
    'ActiveWindow.DisplayHeadings = True
    ' Show the Formula Bar
    'Application.DisplayFormulaBar = True
    ' Show the Ribbon Menu and Quick Access Toolbar
    'Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"",True)"
    ' Show the Status Bar (bottom of the window)
    'Application.DisplayStatusBar = True
    
    'Application.WindowState = xlMaximized
    
End Sub

Private Sub Workbook_Activate()
    
    With Application
        .DecimalSeparator = "."
        .ThousandsSeparator = ","
        .UseSystemSeparators = False
    End With
    
End Sub

Private Sub Workbook_Deactivate()
    
    Application.UseSystemSeparators = True
    
End Sub

