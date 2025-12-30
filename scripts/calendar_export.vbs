Option Explicit

' =============================================================
' SCA Time Automation - Calendar Export v2
' Exports calendar events to JSON with external attendee domains
' =============================================================

Public Sub ExportCalendarToJSON()
    
    On Error GoTo ErrHandler
    
    ' --- Configuration ---
    Const WEEKS_BACK As Integer = 4
    Const OUTPUT_PATH As String = "C:\Users\1028120\Documents\Scripts\sca-time-automation\data\input\calendar_export.json"
    
    ' Internal domains to exclude
    Const INTERNAL_DOMAINS As String = "blueyonder.com,jda.com,microsoft.com"
    
    ' --- Variables ---
    Dim fso As Object
    Dim outFile As Object
    Dim stream As Object
    Dim ns As Outlook.NameSpace
    Dim calendarFolder As Outlook.MAPIFolder
    Dim items As Outlook.Items
    Dim filteredItems As Outlook.Items
    Dim appt As Outlook.AppointmentItem
    Dim recip As Outlook.Recipient
    
    Dim startDate As Date
    Dim endDate As Date
    Dim filterStr As String
    Dim jsonStr As String
    Dim comma As String
    Dim category As String
    Dim externalDomains As String
    Dim recipEmail As String
    Dim recipDomain As String
    Dim cats() As String
    Dim cat As Variant
    Dim internalArr() As String
    Dim eventCount As Long
    Dim i As Long
    Dim isInternal As Boolean
    
    ' --- Setup dates ---
    endDate = DateAdd("d", 1, Date)
    startDate = DateAdd("ww", -WEEKS_BACK, Date)
    
    ' Parse internal domains
    internalArr = Split(INTERNAL_DOMAINS, ",")
    
    ' --- Get calendar ---
    Set ns = Application.GetNamespace("MAPI")
    Set calendarFolder = ns.GetDefaultFolder(olFolderCalendar)
    Set items = calendarFolder.items
    
    ' --- Filter and sort ---
    items.Sort "[Start]", False
    items.IncludeRecurrences = True
    filterStr = "[Start] >= '" & Format(startDate, "MM/DD/YYYY") & "' AND [Start] < '" & Format(endDate, "MM/DD/YYYY") & "'"
    Set filteredItems = items.Restrict(filterStr)
    
    ' --- Build JSON string ---
    jsonStr = "{""events"": ["
    comma = ""
    eventCount = 0
    
    ' --- Process events ---
    For Each appt In filteredItems
        
        ' Extract category (with . prefix)
        category = ""
        If Len(appt.Categories) > 0 Then
            cats = Split(appt.Categories, ",")
            For Each cat In cats
                cat = Trim(cat)
                If Left(cat, 1) = "." Then
                    category = Mid(cat, 2)
                    Exit For
                End If
            Next cat
        End If
        
        ' Skip if no tracked category
        If category = "" Then GoTo NextAppt
        
        ' Get external domains from recipients
        externalDomains = ""
        On Error Resume Next
        For Each recip In appt.Recipients
            recipEmail = ""
            
            ' Try to get SMTP address
            If recip.AddressEntry.AddressEntryUserType = olExchangeUserAddressEntry Or _
               recip.AddressEntry.AddressEntryUserType = olExchangeRemoteUserAddressEntry Then
                recipEmail = recip.AddressEntry.GetExchangeUser.PrimarySmtpAddress
            ElseIf recip.AddressEntry.AddressEntryUserType = olSmtpAddressEntry Then
                recipEmail = recip.AddressEntry.Address
            End If
            
            ' Extract domain
            If InStr(recipEmail, "@") > 0 Then
                recipDomain = LCase(Mid(recipEmail, InStr(recipEmail, "@") + 1))
                
                ' Check if external
                isInternal = False
                For i = LBound(internalArr) To UBound(internalArr)
                    If recipDomain = Trim(internalArr(i)) Then
                        isInternal = True
                        Exit For
                    End If
                Next i
                
                ' Add if external and not already in list
                If Not isInternal And recipDomain <> "" Then
                    If InStr(externalDomains, recipDomain) = 0 Then
                        If externalDomains <> "" Then externalDomains = externalDomains & ","
                        externalDomains = externalDomains & recipDomain
                    End If
                End If
            End If
        Next recip
        On Error GoTo ErrHandler
        
        ' Build JSON for this event
        jsonStr = jsonStr & comma & "{" & vbCrLf
        jsonStr = jsonStr & "  ""start"": """ & Format(appt.Start, "YYYY-MM-DD HH:NN") & """," & vbCrLf
        jsonStr = jsonStr & "  ""end"": """ & Format(appt.End, "YYYY-MM-DD HH:NN") & """," & vbCrLf
        jsonStr = jsonStr & "  ""category"": """ & UCase(category) & """," & vbCrLf
        jsonStr = jsonStr & "  ""title"": """ & CleanString(appt.Subject) & """," & vbCrLf
        jsonStr = jsonStr & "  ""minutes"": " & DateDiff("n", appt.Start, appt.End) & "," & vbCrLf
        jsonStr = jsonStr & "  ""all_day"": " & LCase(appt.AllDayEvent) & "," & vbCrLf
        jsonStr = jsonStr & "  ""external_domains"": """ & externalDomains & """," & vbCrLf
        jsonStr = jsonStr & "  ""location"": """ & CleanString(appt.Location) & """," & vbCrLf
        jsonStr = jsonStr & "  ""recipients"": " & appt.Recipients.Count & "," & vbCrLf
        jsonStr = jsonStr & "  ""busy_status"": " & appt.BusyStatus & vbCrLf
        jsonStr = jsonStr & "}"
        
        comma = ","
        eventCount = eventCount + 1
        
NextAppt:
    Next appt
    
    ' --- Close JSON ---
    jsonStr = jsonStr & vbCrLf & "]," & vbCrLf
    jsonStr = jsonStr & """export_date"": """ & Format(Now, "YYYY-MM-DD HH:NN:SS") & """," & vbCrLf
    jsonStr = jsonStr & """weeks_back"": " & WEEKS_BACK & "," & vbCrLf
    jsonStr = jsonStr & """event_count"": " & eventCount & vbCrLf
    jsonStr = jsonStr & "}"
    
    ' --- Write as UTF-8 ---
    Set stream = CreateObject("ADODB.Stream")
    stream.Open
    stream.Type = 2  ' Text
    stream.Charset = "utf-8"
    stream.WriteText jsonStr
    stream.SaveToFile OUTPUT_PATH, 2  ' Overwrite
    stream.Close
    
    MsgBox "Export complete!" & vbCrLf & vbCrLf & _
           "Events: " & eventCount & vbCrLf & _
           "File: " & OUTPUT_PATH, vbInformation, "SCA Calendar Export"
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Export Error"
    
End Sub

Private Function CleanString(ByVal str As String) As String
    str = Replace(str, "\", "\\")
    str = Replace(str, """", "\""")
    str = Replace(str, vbCrLf, " ")
    str = Replace(str, vbCr, " ")
    str = Replace(str, vbLf, " ")
    str = Replace(str, vbTab, " ")
    CleanString = str
End Function