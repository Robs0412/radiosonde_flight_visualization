Attribute VB_Name = "ExcelUI"
Option Explicit

Sub HideAll()

'https://www.teachexcel.com/excel-tutorial/2417/hide-the-entire-excel-interface-ribbon-menu-quick-access-toolbar-status?nav=yt
    ' Hide the Row/Column Headings
    ActiveWindow.DisplayHeadings = False
    ' Hide the Formula Bar
    Application.DisplayFormulaBar = False
    ' Hide the Ribbon Menu and Quick Access Toolbar
    Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"",False)"
    ' Hide the Status Bar (bottom of the window)
    'Application.DisplayStatusBar = False
 
End Sub


Sub ShowAll()

'https://www.teachexcel.com/excel-tutorial/2417/hide-the-entire-excel-interface-ribbon-menu-quick-access-toolbar-status?nav=yt
    ' Show the Row/Column Headings
    ActiveWindow.DisplayHeadings = True
    ' Show the Formula Bar
    Application.DisplayFormulaBar = True
    ' Show the Ribbon Menu and Quick Access Toolbar
    Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"",True)"
    ' Show the Status Bar (bottom of the window)
    'Application.DisplayStatusBar = True
 
End Sub
