Sub NewTabByCostCode()
    Const strCol = "N" ' the column with categories
    Dim wshS As Worksheet
    Dim wshT As Worksheet
    Dim r As Long
    Dim r0 As Long
    Dim m As Long
    Dim ID
    Application.ScreenUpdating = False
    Set wshS = ActiveSheet
    m = wshS.Range(strCol & Rows.Count).End(xlUp).Row
    wshS.Range("1:" & m).Sort Key1:=wshS.Range(strCol & "1"), Header:=xlYes
    r = 2
    Do
        If wshS.Range(strCol & r).Value <> wshS.Range(strCol & r - 1).Value Then
            Set wshT = Worksheets.Add(After:=Worksheets(Worksheets.Count))
            wshT.Name = wshS.Range(strCol & r).Value
            wshT.Range("1:1").Value = wshS.Range("1:1").Value
            ID = wshS.Range(strCol & r).Value
            r0 = r
            Do While wshS.Range(strCol & r + 1).Value = ID
                r = r + 1
            Loop
            wshT.Range("2:" & r - r0 + 2).Value = wshS.Range(r0 & ":" & r).Value
        End If
        r = r + 1
    Loop Until r > m
    Application.ScreenUpdating = True
End Sub
