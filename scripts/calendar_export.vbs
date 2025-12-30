  Option Explicit

  ' =============================================================
  ' SCA Time Automation - Calendar Export
  ' Exports your calendar events to JSON for time tracking
  ' =============================================================

  Public Sub ExportCalendarToJSON()
      
      On Error GoTo ErrHandler
      
      ' --- Configuration ---
      Const WEEKS_BACK As Integer = 4
      Const OUTPUT_PATH As String = "C:\Users\1028120\Documents\Scripts\sca-time-automation\data\input\calendar_export.json"
      
      ' --- Variables ---
      Dim fso As Object
      Dim outFile As Object
      Dim ns As Outlook.NameSpace
      Dim calendarFolder As Outlook.MAPIFolder
      Dim items As Outlook.Items
      Dim filteredItems As Outlook.Items
      Dim appt As Outlook.AppointmentItem
      Dim organizer As Outlook.AddressEntry
      
      Dim startDate As Date
      Dim endDate As Date
      Dim filterStr As String
      Dim comma As String
      Dim category As String
      Dim organizerDomain As String
      Dim organizerEmail As String
      Dim cats() As String
      Dim cat As Variant
      Dim eventCount As Long
      
      ' --- Setup dates ---
      endDate = DateAdd("d", 1, Date)
      startDate = DateAdd("ww", -WEEKS_BACK, Date)
      
      ' --- Get calendar ---
      Set ns = Application.GetNamespace("MAPI")
      Set calendarFolder = ns.GetDefaultFolder(olFolderCalendar)
      Set items = calendarFolder.items
      
      ' --- Filter and sort ---
      items.Sort "[Start]", False
      items.IncludeRecurrences = True
      filterStr = "[Start] >= '" & Format(startDate, "MM/DD/YYYY") & "' AND [Start] < '" & Format(endDate, "MM/DD/YYYY") & "'"
      Set filteredItems = items.Restrict(filterStr)
      
      ' --- Create output file ---
      Set fso = CreateObject("Scripting.FileSystemObject")
      Set outFile = fso.CreateTextFile(OUTPUT_PATH, True, True)
      
      outFile.Write "{""events"": ["
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
                      category = Mid(cat, 2)  ' Remove . prefix
                      Exit For
                  End If
              Next cat
          End If
          
          ' Skip if no tracked category
          If category = "" Then GoTo NextAppt
          
          ' Get organizer domain
          organizerDomain = ""
          organizerEmail = ""
          
          On Error Resume Next
          If appt.MeetingStatus <> olNonMeeting Then
              organizerEmail = appt.Organizer
              
              ' Try to get SMTP address
              Set organizer = ns.CreateRecipient(appt.Organizer).AddressEntry
              If Not organizer Is Nothing Then
                  If organizer.AddressEntryUserType = olExchangeUserAddressEntry Or _
                    organizer.AddressEntryUserType = olExchangeRemoteUserAddressEntry Then
                      organizerEmail = organizer.GetExchangeUser.PrimarySmtpAddress
                  ElseIf organizer.AddressEntryUserType = olSmtpAddressEntry Then
                      organizerEmail = organizer.Address
                  End If
              End If
              
              ' Extract domain
              If InStr(organizerEmail, "@") > 0 Then
                  organizerDomain = LCase(Mid(organizerEmail, InStr(organizerEmail, "@") + 1))
              End If
          End If
          On Error GoTo ErrHandler
          
          ' Write JSON
          outFile.WriteLine comma & "{"
          outFile.WriteLine "  ""start"": """ & Format(appt.Start, "YYYY-MM-DD HH:NN") & ""","
          outFile.WriteLine "  ""end"": """ & Format(appt.End, "YYYY-MM-DD HH:NN") & ""","
          outFile.WriteLine "  ""category"": """ & UCase(category) & ""","
          outFile.WriteLine "  ""title"": """ & CleanString(appt.Subject) & ""","
          outFile.WriteLine "  ""minutes"": " & DateDiff("n", appt.Start, appt.End) & ","
          outFile.WriteLine "  ""all_day"": " & LCase(appt.AllDayEvent) & ","
          outFile.WriteLine "  ""organizer_domain"": """ & organizerDomain & ""","
          outFile.WriteLine "  ""location"": """ & CleanString(appt.Location) & ""","
          outFile.WriteLine "  ""recipients"": " & appt.Recipients.Count & ","
          outFile.WriteLine "  ""busy_status"": " & appt.BusyStatus
          outFile.Write "}"
          
          comma = ","
          eventCount = eventCount + 1
          
  NextAppt:
      Next appt
      
      ' --- Close JSON ---
      outFile.WriteLine ""
      outFile.WriteLine "],"
      outFile.WriteLine """export_date"": """ & Format(Now, "YYYY-MM-DD HH:NN:SS") & ""","
      outFile.WriteLine """weeks_back"": " & WEEKS_BACK & ","
      outFile.WriteLine """event_count"": " & eventCount
      outFile.Write "}"
      
      outFile.Close
      
      MsgBox "Export complete!" & vbCrLf & vbCrLf & _
            "Events: " & eventCount & vbCrLf & _
            "File: " & OUTPUT_PATH, vbInformation, "SCA Calendar Export"
      
      Exit Sub
      
  ErrHandler:
      MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Export Error"
      If Not outFile Is Nothing Then outFile.Close
      
  End Sub

  Private Function CleanString(ByVal str As String) As String
      ' Clean string for JSON output
      str = Replace(str, "\", "\\")
      str = Replace(str, """", "\""")
      str = Replace(str, vbCrLf, " ")
      str = Replace(str, vbCr, " ")
      str = Replace(str, vbLf, " ")
      str = Replace(str, vbTab, " ")
      CleanString = str
  End Function