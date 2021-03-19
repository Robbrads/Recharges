Sub CreateEmail()
    On Error GoTo ErrHandler
    
    ' SET Outlook APPLICATION OBJECT.
    Dim objOutlook As Object
    Set objOutlook = CreateObject("Outlook.Application")
	Dim Path As String
	Path = ThisWorkbook.Path   
   mth = MonthName(Month(Now))
    ' CREATE EMAIL OBJECT.
    Dim objEmail As Object
    Set objEmail = objOutlook.CreateItem(olMailItem)

    With objEmail
        .to = "bernadette.grant@uhl-tr.nhs.uk; bill.tracey@uhl-tr.nhs.uk; daniel.macswiney@uhl-tr.nhs.uk"
        .CC = "paul.dunn@uhl-tr.nhs.uk; stephen.weston@uhl-tr.nhs.uk;"
        .Subject = "Transplant Lab recharges " & mth & " " & Year(Date)
        .Body = "Hi," & Chr(13) & Chr(13) & "Please find attached Transplant Lab recharges for January.  LCN cost code to be charged to North Lincs and Goole as normal.  Note there are several referrals with unknown cost codes." & Chr(13) & Chr(13) & "Thanks" & Chr(13) & Chr(13) & "Rob" & Chr(13) & Chr(13) & "Robert Bradshaw" & Chr(13) & "Clinical Scientist" & Chr(13) & "Transplant Laboratory" & Chr(13) & "University Hospitals of Leicester NHS Trust" & Chr(13) & "Gwendolen Road" & Chr(13) & "Leicester" & Chr(13) & "LE5 4PW" & Chr(13) & "0116 258 4607" & Chr(13) & "robert.bradshaw@uhl-tr.nhs.uk" & Chr(13) & "robertbradshaw@nhs.net" & Chr(13) & "uho-tr.eastmidlandstransplantlab@nhs.net"
        .Attachments.add (Path "\"*.xlsx"
    End With
    
    
    Set objEmail = Nothing:    Set objOutlook = Nothing
        
ErrHandler:
    '
End Sub
