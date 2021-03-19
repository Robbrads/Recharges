Sub SplitTabsNewWorkBooks()
Dim xPath As String
Dim mth As String
mth = MonthName(Month(Now), True)
xPath = Application.ActiveWorkbook.Path
Application.ScreenUpdating = False
Application.DisplayAlerts = False
For Each xWs In ThisWorkbook.Sheets
    xWs.Copy
    Application.ActiveWorkbook.SaveAs Filename:=xPath & "\" & xWs.Name & " recharges " & mth & " " & Year(Date) & ".xlsx"
    Application.ActiveWorkbook.Close False
Next
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub

