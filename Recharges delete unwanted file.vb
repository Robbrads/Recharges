Sub DeleteFile()
mth = MonthName(Month(Now), True)
Dim Path As String
Path = ThisWorkbook.Path
Kill (Path & "\" & "Cost Codes recharges" & " " & mth & " " & Year(Date) & ".xlsx")
End Sub

