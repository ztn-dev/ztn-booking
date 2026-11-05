'===============================================================================
'  ZTN STUDIO BOOKING SYSTEM  -  single self-building module
'  Paste into a new standard module, then run BuildZTNWorkbook once (Alt+F8).
'  Source is pure ASCII on purpose; dashes, arrows and dots are built with ChrW
'  so the module pastes cleanly into any VBA editor locale.
'===============================================================================
Option Explicit

'------------------------------- CONFIG ----------------------------------------
' Engineer who receives new requests. Blank = no engineer emails are sent.
Public Const ENGINEER_EMAIL As String = ""
Public Const SCHED_FIRST_ROW As Long = 5
Public Const WEEK_FIRST_ROW As Long = 5

Private Const SCHED_PREP_LAST_ROW As Long = 200

Private Const SH_FORM As String = "Booking Form"
Private Const SH_SCHED As String = "Studio Schedule"
Private Const SH_WEEK As String = "Weekly Overview"
Private Const SH_HOWTO As String = "How to Use"
Private Const SH_SHOWS As String = "Weekly Shows"
Private Const SHOWS_FIRST_ROW As Long = 5
Private Const SHOWS_PREP_LAST_ROW As Long = 120
Private Const DEFAULT_SHOW_STUDIO As String = "Studio 1"

' Studio Schedule columns (O..Q are hidden engine columns)
Private Const C_REF As Long = 1
Private Const C_DATE As Long = 2
Private Const C_DAY As Long = 3
Private Const C_START As Long = 4
Private Const C_END As Long = 5
Private Const C_STUDIO As Long = 6
Private Const C_TITLE As Long = 7
Private Const C_REQ As Long = 8
Private Const C_PROD As Long = 9
Private Const C_LIVE As Long = 10
Private Const C_CREW As Long = 11
Private Const C_STATUS As Long = 12
Private Const C_PRIO As Long = 13
Private Const C_NOTES As Long = 14
Private Const C_LASTAPPLIED As Long = 15
Private Const C_SETUP As Long = 16
Private Const C_KEY As Long = 17

Private mNextAutoClose As Date

Public Function StatusList() As Variant
    StatusList = Array("Requested", "Pending Approval", "Conflict", "Confirmed", _
                       "Declined", "Rescheduled", "Cancelled", "Completed")
End Function

Public Function BlockingStatuses() As Variant
    BlockingStatuses = Array("Requested", "Pending Approval", "Conflict", "Confirmed", _
                             "Rescheduled", "Completed")
End Function

Public Function LiveStatuses() As Variant
    LiveStatuses = Array("Requested", "Pending Approval", "Conflict", "Confirmed", "Rescheduled")
End Function

Public Function DisplayStatuses() As Variant
    DisplayStatuses = Array("Confirmed", "Completed")
End Function

Public Function StudioList() As Variant
    StudioList = Array("Studio 1", "Studio 2", "Audio Studio", "MCR", "Other")
End Function

Public Function PriorityList() As Variant
    PriorityList = Array("High", "Urgent")
End Function

' Overview rows in minutes since midnight: 06:00-08:00, half-hour rows
' 08:00 ... 19:30, then 19:30-22:00.
Public Function SlotBounds() As Variant
    Dim b() As Variant, i As Long, n As Long
    ReDim b(0 To 25)
    b(0) = 360
    For i = 480 To 1170 Step 30
        n = n + 1
        b(n) = i
    Next i
    b(n + 1) = 1320
    SlotBounds = b
End Function

' Regular programmes seeded into the Weekly Shows sheet:
' show | days | start | end | studio
Private Function SeedShows() As Variant
    SeedShows = Array( _
        Array("Kwayedza Show", "Fri", "09:00", "09:30", DEFAULT_SHOW_STUDIO), _
        Array("GAME CHANGERS", "Tue", "10:30", "11:00", DEFAULT_SHOW_STUDIO), _
        Array("INSIDE THE SCRUM", "Wed", "10:30", "11:00", DEFAULT_SHOW_STUDIO), _
        Array("Beyond the Boundary", "Thu", "10:30", "11:00", DEFAULT_SHOW_STUDIO), _
        Array("NEWSBLITZ / NHAU / INDABA / WORLD IN 5", "Mon-Fri", "12:00", "13:00", DEFAULT_SHOW_STUDIO), _
        Array("NEWSBLITZ", "Mon-Fri", "16:00", "16:30", DEFAULT_SHOW_STUDIO), _
        Array("ULTIMATE SPORTS SHOW", "Mon", "18:00", "18:30", DEFAULT_SHOW_STUDIO), _
        Array("NEWSPLUS", "Mon-Fri", "19:00", "19:30", DEFAULT_SHOW_STUDIO))
End Function

Private Function EngineOwnedLabels() As Variant
    EngineOwnedLabels = Array("Booking reference", "Booking status", _
                              "Studio/Technical approval", "Confirmed date/time", "Date submitted")
End Function

Private Function DashEm() As String
    DashEm = ChrW(8212)
End Function

Private Function DashEn() As String
    DashEn = ChrW(8211)
End Function

'===============================================================================
'  INSTALLER
'===============================================================================
Public Sub BuildZTNWorkbook()
    Dim wb As Workbook, ws As Worksheet, tmp As Worksheet, i As Long, ok As Boolean

    Set wb = ThisWorkbook
    If MsgBox("This deletes EVERY sheet in """ & wb.Name & """ and rebuilds the ZTN booking system." & _
              vbLf & vbLf & "Continue?", vbExclamation + vbYesNo, "ZTN Builder") <> vbYes Then Exit Sub

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    ' A workbook must always keep one sheet, so park a temporary one first.
    Set tmp = wb.Worksheets.Add(Before:=wb.Sheets(1))
    For i = wb.Sheets.Count To 1 Step -1
        If Not wb.Sheets(i) Is tmp Then wb.Sheets(i).Delete
    Next i

    Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count)): ws.Name = SH_FORM
    BuildBookingForm ws
    Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count)): ws.Name = SH_SCHED
    BuildSchedule ws
    Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count)): ws.Name = SH_WEEK
    BuildWeekly ws
    Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count)): ws.Name = SH_SHOWS
    BuildWeeklyShows ws
    Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count)): ws.Name = SH_HOWTO
    BuildHowTo ws
    tmp.Delete

    ViewSetup wb.Worksheets(SH_HOWTO), 0, False
    ViewSetup wb.Worksheets(SH_SHOWS), 4, True
    ViewSetup wb.Worksheets(SH_WEEK), 4, False
    ViewSetup wb.Worksheets(SH_SCHED), 4, True
    ViewSetup wb.Worksheets(SH_FORM), 0, False

    ProtectForm wb.Worksheets(SH_FORM)
    ProtectSheet wb.Worksheets(SH_SCHED)
    ProtectSheet wb.Worksheets(SH_WEEK)
    ProtectSheet wb.Worksheets(SH_SHOWS)
    ProtectSheet wb.Worksheets(SH_HOWTO)

    RebuildWeeklyOverview

    wb.Worksheets(SH_FORM).Activate
    wb.Worksheets(SH_FORM).Range("B5").Select
    ok = True

Done:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If ok Then MsgBox "ZTN booking workbook built." & vbLf & _
        "Save it as .xlsm. Run ScheduleAutoClose once per session for the hourly auto-close.", _
        vbInformation, "ZTN Builder"
    Exit Sub
Fail:
    MsgBox "Build failed: " & Err.Description, vbCritical, "ZTN Builder"
    Resume Done
End Sub

Private Sub BuildBookingForm(ByVal ws As Worksheet)
    Const YN As String = "Yes,No,To be confirmed"
    Dim widths As Variant, i As Long

    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    widths = Array(27, 23, 24, 22, 22, 22, 3, 24)
    For i = 0 To UBound(widths)
        ws.Columns(i + 1).ColumnWidth = widths(i)
    Next i

    With ws.Range("A1:F1")
        .Merge
        .Value = "ZTN STUDIO BOOKING REQUEST FORM"
        .Font.Size = 16: .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 56, 100)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .RowHeight = 32
    End With
    With ws.Range("A2:F2")
        .Merge
        .Value = "Complete all relevant fields and click Submit Request. " & _
                 "A booking is confirmed only after Studio/Technical approval."
        .Font.Italic = True: .WrapText = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .RowHeight = 30
    End With

    AddSection ws, 4, "1. REQUESTER AND PROGRAMME DETAILS"
    AddField ws, 5, 1, "Requester name"
    AddField ws, 5, 3, "Department / Unit"
    AddField ws, 5, 5, "Contact number / email"
    AddField ws, 6, 1, "Programme / production title"
    AddField ws, 6, 3, "Programme type", , _
        "News bulletin,Interview,Talk show,Podcast,Commercial,Livestream,Special event,Other"
    AddField ws, 6, 5, "Producer / coordinator"
    AddMergedField ws, 7, 1, "Purpose / brief"
    AddField ws, 8, 1, "Target audience / platform"
    AddField ws, 8, 3, "Live or recorded", , "Live,Recorded,Live-to-tape"
    AddField ws, 8, 5, "Language(s)", , "English,Shona,Ndebele,English/Shona,Other"
    AddField ws, 9, 1, "Requested date", "date"
    AddField ws, 9, 3, "Start time", "time"
    AddField ws, 9, 5, "End time", "time"
    AddField ws, 10, 1, "Setup access time", "time"
    AddField ws, 10, 3, "Expected duration"
    AddField ws, 10, 5, "No. of participants"
    AddMergedField ws, 11, 1, "Alternative date/time"

    AddSection ws, 13, "2. BOOKING REQUIREMENTS"
    AddField ws, 14, 1, "Studio requested", , StudioList()
    AddField ws, 14, 3, "Booking priority", , PriorityList()
    AddField ws, 14, 5, "Booking status", , StatusList(), True
    AddField ws, 15, 1, "Set / background required"
    AddField ws, 15, 3, "Seating / layout"
    AddField ws, 15, 5, "Audience required?", , YN
    AddField ws, 16, 1, "Camera requirement", , "1 camera,2 cameras,3 cameras,PTZ / remote,To be confirmed"
    AddField ws, 16, 3, "Audio / microphones", , _
        "Handheld mics,Lapel mics,Wireless mics,Studio microphones,Mixed / other"
    AddField ws, 16, 5, "Lighting requirement", , "Standard,Special lighting,To be confirmed"
    AddField ws, 17, 1, "Graphics / CG"
    AddField ws, 17, 3, "VT / playback", , YN
    AddField ws, 17, 5, "Teleprompter", , YN
    AddField ws, 18, 1, "Livestream platform(s)"
    AddField ws, 18, 3, "Remote guest / video call", , YN
    AddField ws, 18, 5, "Recording required?", , YN
    AddMergedField ws, 19, 1, "Special equipment / notes"
    AddField ws, 20, 1, "Rehearsal / sound check required?", , YN
    AddField ws, 20, 3, "Rehearsal date/time"

    AddSection ws, 22, "3. TECHNICAL AND PRODUCTION REQUIREMENTS"
    AddField ws, 23, 1, "Production crew requested"
    AddField ws, 23, 3, "Studio camera operator", , YN
    AddField ws, 23, 5, "Audio operator", , YN
    AddField ws, 24, 1, "Vision mixer / switcher operator", , YN
    AddField ws, 24, 3, "VT / playback operator", , YN
    AddField ws, 24, 5, "CG operator", , YN
    AddField ws, 25, 1, "Streaming / encoder operator", , YN
    AddMergedField ws, 25, 3, "Technical support notes"
    AddField ws, 26, 1, "Content / media supplied by"
    AddField ws, 26, 3, "Media deadline", "date"

    AddSection ws, 28, "4. APPROVAL AND CONFIRMATION"
    AddField ws, 29, 1, "Requester signature / name"
    AddField ws, 29, 3, "Date submitted", "datetime", , True
    AddField ws, 29, 5, "Producer approval"
    AddField ws, 30, 1, "Studio/Technical approval", , , True
    AddField ws, 30, 3, "Confirmed date/time", , , True
    AddField ws, 30, 5, "Booking reference", , , True
    AddMergedField ws, 31, 1, "Conditions / remarks"

    AddButton ws, ws.Columns("H").Left, ws.Rows(4).Top, 150, 30, "Submit Request", "SubmitBookingForm"
    AddButton ws, ws.Columns("H").Left, ws.Rows(4).Top + 40, 150, 30, "Reset Form", "ResetBookingForm"
End Sub

Private Sub BuildSchedule(ByVal ws As Worksheet)
    Dim hdr As Variant, widths As Variant, i As Long, lastR As Long, x As Double

    lastR = SCHED_PREP_LAST_ROW
    hdr = Array("Booking Ref", "Date", "Day", "Start Time", "End Time", "Studio", _
                "Programme / Production", "Department / Requester", "Producer / Coordinator", _
                "Live / Recorded", "Crew / Technical Support", "Status", "Priority", "Notes", _
                "Last Applied", "Setup Access", "Change Key")
    widths = Array(14, 15, 13, 12, 12, 18, 29, 24, 22, 16, 30, 20, 14, 40, 14, 12, 26)

    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    For i = 0 To UBound(hdr)
        ws.Cells(4, i + 1).Value = hdr(i)
        ws.Columns(i + 1).ColumnWidth = widths(i)
    Next i

    With ws.Range("A1:N1")
        .Merge
        .Value = "ZTN STUDIO BOOKING & PRODUCTION SCHEDULE"
        .Font.Size = 16: .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 56, 100)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .RowHeight = 30
    End With
    With ws.Range("A2:N2")
        .Merge
        .Value = "One booking per row. Change Status in column L, then click Apply All Changes. " & _
                 "Setup access time is counted when checking for clashes."
        .Font.Italic = True
        .HorizontalAlignment = xlCenter
    End With
    ws.Rows(3).RowHeight = 32

    With ws.Range("A4:Q4")
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(47, 84, 150)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .WrapText = True
        .RowHeight = 30
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(255, 255, 255)
    End With

    With ws.Range(ws.Cells(SCHED_FIRST_ROW, 1), ws.Cells(lastR, C_NOTES))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(217, 217, 217)
        .VerticalAlignment = xlTop
    End With
    ws.Range(ws.Cells(SCHED_FIRST_ROW, C_TITLE), ws.Cells(lastR, C_PROD)).WrapText = True
    ws.Range(ws.Cells(SCHED_FIRST_ROW, C_CREW), ws.Cells(lastR, C_CREW)).WrapText = True
    ws.Range(ws.Cells(SCHED_FIRST_ROW, C_NOTES), ws.Cells(lastR, C_NOTES)).WrapText = True

    ws.Range(ws.Cells(SCHED_FIRST_ROW, C_DAY), ws.Cells(lastR, C_DAY)).Formula = _
        "=IF(B" & SCHED_FIRST_ROW & "="""","""",TEXT(B" & SCHED_FIRST_ROW & ",""dddd""))"
    ws.Range(ws.Cells(SCHED_FIRST_ROW, C_DATE), ws.Cells(lastR, C_DATE)).NumberFormat = "dd-mmm-yyyy"
    ws.Range(ws.Cells(SCHED_FIRST_ROW, C_START), ws.Cells(lastR, C_END)).NumberFormat = "hh:mm"
    ws.Range(ws.Cells(SCHED_FIRST_ROW, C_SETUP), ws.Cells(lastR, C_SETUP)).NumberFormat = "hh:mm"

    ApplyList ws.Range(ws.Cells(SCHED_FIRST_ROW, C_STUDIO), ws.Cells(lastR, C_STUDIO)), StudioList(), True
    ApplyList ws.Range(ws.Cells(SCHED_FIRST_ROW, C_LIVE), ws.Cells(lastR, C_LIVE)), "Live,Recorded,Live-to-tape", True
    ApplyList ws.Range(ws.Cells(SCHED_FIRST_ROW, C_STATUS), ws.Cells(lastR, C_STATUS)), StatusList(), True
    ApplyList ws.Range(ws.Cells(SCHED_FIRST_ROW, C_PRIO), ws.Cells(lastR, C_PRIO)), PriorityList(), True

    ws.Cells.Locked = True
    ws.Range(ws.Cells(SCHED_FIRST_ROW, C_REF), ws.Cells(lastR, C_DATE)).Locked = False
    ws.Range(ws.Cells(SCHED_FIRST_ROW, C_START), ws.Cells(lastR, C_NOTES)).Locked = False
    ws.Range(ws.Columns(C_LASTAPPLIED), ws.Columns(C_KEY)).EntireColumn.Hidden = True

    x = ws.Range("A3").Left + 2
    AddButton ws, x, ws.Rows(3).Top + 4, 140, 24, "Apply All Changes", "ApplyScheduleChanges"
    AddButton ws, x + 150, ws.Rows(3).Top + 4, 170, 24, "Rebuild Weekly Overview", "RebuildWeeklyOverview"
    AddButton ws, x + 330, ws.Rows(3).Top + 4, 140, 24, "Re-check Conflicts", "RecheckAllConflicts"
    AddButton ws, x + 480, ws.Rows(3).Top + 4, 140, 24, "Close Completed", "CloseCompletedBookings"
End Sub

Private Sub BuildWeekly(ByVal ws As Worksheet)
    Dim days As Variant, b As Variant, i As Long, nSlots As Long, grid As Range

    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    ws.Columns(1).ColumnWidth = 18
    ws.Range("B:H").ColumnWidth = 25

    With ws.Range("A1:H1")
        .Merge
        .Value = "ZTN WEEKLY STUDIO OVERVIEW"
        .Font.Size = 16: .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 56, 100)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .RowHeight = 30
    End With

    ws.Range("A2").Value = "Week commencing:"
    StyleLabel ws.Range("A2")
    ws.Range("B2").Value = MondayOf(Date)
    StyleInput ws.Range("B2"), "date"
    ws.Range("D2").Value = "Prepared by:"
    StyleLabel ws.Range("D2")
    ws.Range("E2:H2").Merge
    ws.Range("E2").Value = Application.UserName
    StyleInput ws.Range("E2:H2")
    ws.Rows(2).RowHeight = 22
    ws.Rows(3).RowHeight = 32

    days = Array("Time", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")
    For i = 0 To 7
        ws.Cells(4, i + 1).Value = days(i)
    Next i
    With ws.Range("A4:H4")
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(47, 84, 150)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .RowHeight = 22
    End With

    b = SlotBounds()
    nSlots = UBound(b) - LBound(b)
    For i = 1 To nSlots
        ws.Cells(WEEK_FIRST_ROW + i - 1, 1).Value = Mm(b(LBound(b) + i - 1)) & DashEn() & Mm(b(LBound(b) + i))
    Next i
    With ws.Range(ws.Cells(WEEK_FIRST_ROW, 1), ws.Cells(WEEK_FIRST_ROW + nSlots - 1, 1))
        .Font.Bold = True
        .Interior.Color = RGB(242, 242, 242)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    Set grid = ws.Range(ws.Cells(WEEK_FIRST_ROW, 1), ws.Cells(WEEK_FIRST_ROW + nSlots - 1, 8))
    With grid
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(166, 166, 166)
        .RowHeight = 24
    End With
    With ws.Range(ws.Cells(WEEK_FIRST_ROW, 2), ws.Cells(WEEK_FIRST_ROW + nSlots - 1, 8))
        .WrapText = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .Font.Size = 9
    End With
    With ws.Range("C3")
        .Value = "Green = regular show   " & ChrW(183) & "   Blue = booking   " & ChrW(183) & _
                 "   Red = same studio used twice"
        .Font.Italic = True
        .Font.Color = RGB(89, 89, 89)
        .VerticalAlignment = xlCenter
    End With

    ws.Cells(WEEK_FIRST_ROW + nSlots, 1).Font.Italic = True
    ws.Cells(WEEK_FIRST_ROW + nSlots, 1).Font.Color = RGB(118, 118, 118)
    With ws.Range(ws.Cells(WEEK_FIRST_ROW + nSlots + 1, 1), ws.Cells(WEEK_FIRST_ROW + nSlots + 1, 8))
        .Merge
        .Value = "Operational reminders: allow time for studio setup and reset; confirm crew and equipment; " & _
                 "test audio, graphics, playback and streaming before transmission; " & _
                 "update the main Studio Schedule after any change."
        .WrapText = True
        .Font.Italic = True
        .VerticalAlignment = xlTop
        .RowHeight = 30
    End With

    ws.Cells.Locked = True
    ws.Range("B2").Locked = False
    ws.Range("E2:H2").Locked = False

    AddButton ws, ws.Range("A3").Left + 2, ws.Rows(3).Top + 4, 120, 24, "Refresh Now", "RebuildWeeklyOverview"
End Sub

Private Sub BuildWeeklyShows(ByVal ws As Worksheet)
    Dim hdr As Variant, widths As Variant, seed As Variant, i As Long, r As Long, lastR As Long

    lastR = SHOWS_PREP_LAST_ROW
    hdr = Array("Show / Programme", "Days", "Start Time", "End Time", "Studio", "Live / Recorded", _
                "Producer / Presenter", "Active From", "Active To", "Notes")
    widths = Array(40, 16, 12, 12, 22, 16, 26, 14, 14, 36)

    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    For i = 0 To UBound(hdr)
        ws.Cells(4, i + 1).Value = hdr(i)
        ws.Columns(i + 1).ColumnWidth = widths(i)
    Next i

    With ws.Range("A1:J1")
        .Merge
        .Value = "ZTN WEEKLY SHOWS (REPEAT EVERY WEEK)"
        .Font.Size = 16: .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 56, 100)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .RowHeight = 30
    End With
    With ws.Range("A2:J2")
        .Merge
        .Value = "One row per regular programme. Days: Mon, Mon-Fri, Tue,Thu, Weekdays, Weekends or Daily. " & _
                 "Leave Active From/To blank for an ongoing show. Changes show on the Weekly Overview after Refresh."
        .Font.Italic = True
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .RowHeight = 30
    End With

    With ws.Range("A4:J4")
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(84, 130, 53)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .WrapText = True
        .RowHeight = 30
    End With
    With ws.Range(ws.Cells(SHOWS_FIRST_ROW, 1), ws.Cells(lastR, 10))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(217, 217, 217)
        .VerticalAlignment = xlTop
        .Locked = False
    End With
    ws.Range(ws.Cells(SHOWS_FIRST_ROW, 3), ws.Cells(lastR, 4)).NumberFormat = "hh:mm"
    ws.Range(ws.Cells(SHOWS_FIRST_ROW, 8), ws.Cells(lastR, 9)).NumberFormat = "dd-mmm-yyyy"
    ApplyList ws.Range(ws.Cells(SHOWS_FIRST_ROW, 5), ws.Cells(lastR, 5)), StudioList(), True
    ApplyList ws.Range(ws.Cells(SHOWS_FIRST_ROW, 6), ws.Cells(lastR, 6)), "Live,Recorded,Live-to-tape", True

    seed = SeedShows()
    For i = LBound(seed) To UBound(seed)
        r = SHOWS_FIRST_ROW + i - LBound(seed)
        ws.Cells(r, 1).Value = seed(i)(0)
        ws.Cells(r, 2).Value = seed(i)(1)
        ws.Cells(r, 3).Value = ToMinutes(seed(i)(2)) / 1440
        ws.Cells(r, 4).Value = ToMinutes(seed(i)(3)) / 1440
        ws.Cells(r, 5).Value = seed(i)(4)
    Next i

    ws.Range(ws.Cells(1, 1), ws.Cells(4, 10)).Locked = True
End Sub

Private Sub BuildHowTo(ByVal ws As Worksheet)
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    ws.Columns(1).ColumnWidth = 27
    ws.Columns(2).ColumnWidth = 38
    ws.Columns(3).ColumnWidth = 30
    ws.Columns(4).ColumnWidth = 30

    With ws.Range("A1:D1")
        .Merge
        .Value = "ZTN STUDIO BOOKING WORKFLOW"
        .Font.Size = 16: .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 56, 100)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .RowHeight = 30
    End With

    AddHowRow ws, 3, "1. Request", "Requester completes the Booking Form (date, start/end time, setup access, " & _
        "studio, technical needs, rehearsal) and clicks Submit Request. The engine gives it a ZTN-YYYY-NNNN " & _
        "reference, adds it to the Studio Schedule and emails the studio engineer."
    AddHowRow ws, 4, "2. Check availability", "Every submission is checked against the Studio Schedule for the same " & _
        "date and studio, including setup access time. A clash is written as Conflict with the clashing references in Notes."
    AddHowRow ws, 5, "3. Approve", "Producer / department head and Studio/Technical team decide. Set Status in column L " & _
        "and click Apply All Changes. A clashing booking cannot be confirmed; it is reverted to Conflict."
    AddHowRow ws, 6, "4. Record", "Confirmed bookings get an approval stamp on the Booking Form (when it still shows that " & _
        "reference) and the requester is emailed. Put the requester's email in 'Contact number / email'."
    AddHowRow ws, 7, "5. Prepare", "Confirm scripts, media, graphics, playback files, microphones, camera setup, lighting, " & _
        "internet/streaming needs and any remote guests before the session."
    AddHowRow ws, 8, "6. Close out", "Confirmed bookings move to Completed automatically once their end time passes " & _
        "(Close Completed button, or hourly after running ScheduleAutoClose). Record incidents in Notes."
    AddHowRow ws, 9, "Conflict warning", "Clashes = same date + same studio + overlapping times (setup access extends the start). " & _
        "Requested, Pending Approval, Conflict, Confirmed, Rescheduled and Completed all block a slot. " & _
        "Crew conflicts still need a manual check."
    AddHowRow ws, 10, "Status definitions", "Requested = received; Pending Approval = awaiting decision; " & _
        "Conflict = clashes with another booking; Confirmed = approved and scheduled; Declined = not approved; " & _
        "Rescheduled = moved to another slot; Cancelled = withdrawn; Completed = production finished."

    With ws.Range("A12:D12")
        .Merge
        .Value = "BUTTONS"
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(47, 84, 150)
        .IndentLevel = 1
    End With
    AddHowRow ws, 13, "Submit Request", "Booking Form: validate, reference, conflict check, add to schedule, email engineer."
    AddHowRow ws, 14, "Reset Form", "Booking Form: clear every field for a new request."
    AddHowRow ws, 15, "Apply All Changes", "Studio Schedule: process every row whose Status changed since the last apply, " & _
        "re-check rows whose date/time/studio was edited, send emails, refresh the weekly overview."
    AddHowRow ws, 16, "Rebuild Weekly Overview", "Redraw the weekly grid (Confirmed and Completed only)."
    AddHowRow ws, 17, "Re-check Conflicts", "Re-run the clash check on every open booking."
    AddHowRow ws, 18, "Close Completed", "Move Confirmed bookings whose end time has passed to Completed."
    AddHowRow ws, 19, "Refresh Now", "Weekly Overview: change 'Week commencing' (any date in the week) and click to redraw."

    With ws.Range("A21:D21")
        .Merge
        .Value = "REGULAR WEEKLY SHOWS"
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(84, 130, 53)
        .IndentLevel = 1
    End With
    AddHowRow ws, 22, "Weekly Shows sheet", "List programmes that repeat every week (show, days, start/end, studio). " & _
        "They appear on the Weekly Overview every week in green, without being entered as bookings."
    AddHowRow ws, 23, "Overlaps with a show", "A request that overlaps a regular show in the same studio is NOT rejected. " & _
        "It gets an 'Overlaps regular show' note and the engineer decides whether to approve it. " & _
        "Confirming it records 'Confirmed over regular show ... at engineer's discretion'."

    ws.Cells.Locked = True
End Sub

Private Sub AddHowRow(ByVal ws As Worksheet, ByVal r As Long, ByVal head As String, ByVal txt As String)
    ws.Cells(r, 1).Value = head
    StyleLabel ws.Cells(r, 1)
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 4)).Merge
    ws.Cells(r, 2).Value = txt
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 4))
        .WrapText = True
        .VerticalAlignment = xlTop
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(166, 166, 166)
    End With
    ws.Rows(r).RowHeight = IIf(Len(txt) > 180, 54, IIf(Len(txt) > 95, 40, 22))
End Sub

Private Sub AddSection(ByVal ws As Worksheet, ByVal r As Long, ByVal title As String)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 6))
        .Merge
        .Value = title
        .Font.Bold = True: .Font.Size = 11: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(47, 84, 150)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .RowHeight = 22
        .Locked = True
    End With
End Sub

Private Sub AddField(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal label As String, _
                     Optional ByVal kind As String = "", Optional ByVal listItems As Variant, _
                     Optional ByVal engineOwned As Boolean = False)
    ws.Cells(r, c).Value = label
    StyleLabel ws.Cells(r, c)
    StyleInput ws.Cells(r, c + 1), kind, engineOwned
    If Not IsMissing(listItems) Then ApplyList ws.Cells(r, c + 1), listItems, False
    If ws.Rows(r).RowHeight < 28 Then ws.Rows(r).RowHeight = 28
End Sub

Private Sub AddMergedField(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal label As String, _
                           Optional ByVal kind As String = "")
    Dim v As Range
    ws.Cells(r, c).Value = label
    StyleLabel ws.Cells(r, c)
    Set v = ws.Range(ws.Cells(r, c + 1), ws.Cells(r, 6))
    v.Merge
    StyleInput v, kind
    ws.Rows(r).RowHeight = 36
End Sub

Private Sub StyleLabel(ByVal rng As Range)
    With rng
        .Font.Bold = True
        .Font.Color = RGB(31, 56, 100)
        .Interior.Color = RGB(242, 242, 242)
        .WrapText = True
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(166, 166, 166)
        .Locked = True
    End With
End Sub

Private Sub StyleInput(ByVal rng As Range, Optional ByVal kind As String = "", _
                       Optional ByVal engineOwned As Boolean = False)
    With rng
        .Interior.Color = IIf(engineOwned, RGB(230, 230, 230), RGB(255, 255, 255))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(166, 166, 166)
        .WrapText = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
        .Locked = engineOwned
        Select Case LCase$(kind)
            Case "date": .NumberFormat = "dd-mmm-yyyy"
            Case "time": .NumberFormat = "hh:mm"
            Case "datetime": .NumberFormat = "dd-mmm-yyyy hh:mm"
        End Select
    End With
End Sub

Private Sub AddButton(ByVal ws As Worksheet, ByVal l As Double, ByVal t As Double, ByVal w As Double, _
                      ByVal h As Double, ByVal cap As String, ByVal macroName As String)
    Dim b As Object
    Set b = ws.Buttons.Add(l, t, w, h)
    b.Caption = cap
    b.OnAction = macroName
    b.Placement = xlFreeFloating
    On Error Resume Next
    b.Font.Bold = True
End Sub

Private Sub ApplyList(ByVal rng As Range, ByVal items As Variant, Optional ByVal strict As Boolean = True)
    Dim csv As String
    If IsArray(items) Then csv = Join(items, ",") Else csv = CStr(items)
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=csv
        .IgnoreBlank = True
        .InCellDropdown = True
        .ErrorTitle = "Invalid selection"
        .ErrorMessage = "Please select a value from the list."
        .ShowError = strict
    End With
End Sub

' Lock everything except value cells: label cells (odd columns), section rows,
' blank filler cells and engine-written cells stay locked.
Private Sub ProtectForm(ByVal ws As Worksheet)
    Dim r As Long, c As Long, cell As Range, lbl As Variant, hit As Range

    ws.Cells.Locked = True
    ws.Range("A4:F31").Locked = False
    For r = 4 To 31
        For c = 1 To 6
            Set cell = ws.Cells(r, c)
            If cell.MergeCells Then
                If cell.MergeArea.Column Mod 2 = 1 And cell.Address = cell.MergeArea.Cells(1, 1).Address Then
                    cell.MergeArea.Locked = True
                End If
            ElseIf c Mod 2 = 1 Then
                cell.Locked = True
            ElseIf Len(SafeStr(cell.Offset(0, -1).Value)) = 0 And Not cell.Offset(0, -1).MergeCells Then
                cell.Locked = True
            End If
        Next c
    Next r
    For Each lbl In EngineOwnedLabels()
        Set hit = FindLabelCell(ws, CStr(lbl))
        If Not hit Is Nothing Then hit.Offset(0, 1).MergeArea.Locked = True
    Next lbl
    ProtectSheet ws
End Sub

' UserInterfaceOnly is not saved with the file, so every runtime macro re-arms it.
Private Sub ProtectSheet(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, Scenarios:=True, _
               UserInterfaceOnly:=True, AllowFormattingCells:=True, _
               AllowFormattingColumns:=True, AllowFormattingRows:=True
End Sub

Private Sub ViewSetup(ByVal ws As Worksheet, ByVal freezeRows As Long, ByVal showGrid As Boolean)
    ws.Activate
    With ActiveWindow
        .FreezePanes = False
        .ScrollRow = 1
        .ScrollColumn = 1
        If freezeRows > 0 Then
            .SplitColumn = 0
            .SplitRow = freezeRows
            .FreezePanes = True
        End If
        .DisplayGridlines = showGrid
    End With
End Sub

'===============================================================================
'  RUNTIME - BOOKING FORM
'===============================================================================
Public Sub SubmitBookingForm()
    Dim wsS As Worksheet, hits As Collection, shows As Collection, showTxt As String
    Dim reqName As String, dept As String, contact As String, title As String, studio As String
    Dim dt As Variant, sv As Variant, ev As Variant
    Dim sMin As Long, eMin As Long, effStart As Long, r As Long
    Dim missing As String, ref As String, status As String, prio As String, existingRef As String
    Dim subj As String, body As String

    If Not SheetsReady() Then Exit Sub
    ArmSheets
    Set wsS = Sh(SH_SCHED)

    reqName = FV("Requester name")
    dept = FV("Department / Unit")
    contact = FV("Contact number / email")
    title = FV("Programme / production title")
    studio = FV("Studio requested")
    dt = ToDateVal(GetFormValue("Requested date"))
    sv = ToMinutes(GetFormValue("Start time"))
    ev = ToMinutes(GetFormValue("End time"))

    If reqName = "" Then missing = missing & vbLf & "  - Requester name"
    If title = "" Then missing = missing & vbLf & "  - Programme / production title"
    If IsEmpty(dt) Then missing = missing & vbLf & "  - Requested date (a valid date)"
    If IsEmpty(sv) Then missing = missing & vbLf & "  - Start time (e.g. 10:00)"
    If IsEmpty(ev) Then missing = missing & vbLf & "  - End time (e.g. 12:00)"
    If studio = "" Then missing = missing & vbLf & "  - Studio requested"
    If Len(missing) > 0 Then
        MsgBox "Please complete the following before submitting:" & missing, vbExclamation, "ZTN Booking"
        Exit Sub
    End If

    sMin = sv: eMin = ev
    If eMin <= sMin Then
        MsgBox "End time must be later than start time (bookings cannot run past midnight).", vbExclamation, "ZTN Booking"
        Exit Sub
    End If
    If Not IsInArray(studio, StudioList()) Then
        MsgBox "Choose a studio from the drop-down list.", vbExclamation, "ZTN Booking"
        Exit Sub
    End If

    existingRef = FV("Booking reference")
    If existingRef <> "" Then
        If MsgBox("This form was already submitted as " & existingRef & "." & vbLf & _
                  "Submit it again as a NEW booking?", vbQuestion + vbYesNo, "ZTN Booking") <> vbYes Then Exit Sub
    End If

    effStart = SetupStartMin(GetFormValue("Setup access time"), sMin)
    Set hits = FindConflicts(wsS, dt, effStart, eMin, studio, "")
    Set shows = FindShowOverlaps(dt, effStart, eMin, studio)
    showTxt = RefsJoined(shows)
    status = IIf(hits.Count > 0, "Conflict", "Requested")
    ref = NextBookingRef()
    r = LastSchedRow(wsS) + 1
    prio = FV("Booking priority")

    Application.ScreenUpdating = False
    FormatSchedRow wsS, r
    wsS.Cells(r, C_REF).Value = ref
    wsS.Cells(r, C_DATE).Value = CDate(dt)
    wsS.Cells(r, C_START).Value = sMin / 1440
    wsS.Cells(r, C_END).Value = eMin / 1440
    wsS.Cells(r, C_STUDIO).Value = studio
    wsS.Cells(r, C_TITLE).Value = title
    wsS.Cells(r, C_REQ).Value = JoinFilter(Array(reqName, dept, contact), " " & DashEm() & " ")
    wsS.Cells(r, C_PROD).Value = FV("Producer / coordinator")
    wsS.Cells(r, C_LIVE).Value = FV("Live or recorded")
    wsS.Cells(r, C_CREW).Value = BuildCrewSummary()
    wsS.Cells(r, C_STATUS).Value = status
    wsS.Cells(r, C_PRIO).Value = prio
    If effStart < sMin Then wsS.Cells(r, C_SETUP).Value = effStart / 1440
    wsS.Cells(r, C_NOTES).Value = JoinFilter(Array( _
        IIf(effStart < sMin, "Setup access from " & Mm(effStart), ""), _
        FV("Special equipment / notes"), _
        IIf(LCase$(FV("Rehearsal / sound check required?")) = "yes", _
            "Rehearsal requested " & FV("Rehearsal date/time"), "")), " | ")
    If hits.Count > 0 Then StampNote wsS, r, "CONFLICT with " & RefsJoined(hits)
    If shows.Count > 0 Then StampNote wsS, r, ShowOverlapNote(showTxt)
    wsS.Cells(r, C_LASTAPPLIED).Value = status
    wsS.Cells(r, C_KEY).Value = RowKey(wsS, r)
    ColorStatus wsS, r

    SetFormValue "Booking reference", ref
    SetFormValue "Booking status", status
    SetFormValue "Date submitted", Now
    SetFormValue "Studio/Technical approval", ""
    SetFormValue "Confirmed date/time", ""
    Application.ScreenUpdating = True

    If hits.Count > 0 Then
        subj = "[ZTN Studio] CONFLICT " & DashEm() & " " & ref & " " & DashEm() & " " & title
        body = "A new booking request CLASHES with: " & RefsJoined(hits) & vbLf & vbLf
    Else
        subj = "[ZTN Studio] New booking request " & DashEm() & " " & ref & " " & DashEm() & " " & title
        body = "A new studio booking request has been submitted." & vbLf & vbLf
    End If
    If shows.Count > 0 Then
        body = body & "NOTE: overlaps regular show(s) " & showTxt & " in " & studio & _
               ". Approve only if the show can move or does not need the studio." & vbLf & vbLf
    End If
    SendMail ENGINEER_EMAIL, subj, body & BookingSummary(wsS, r)

    If hits.Count > 0 Then
        MsgBox "Submitted as " & ref & " but it CLASHES with " & RefsJoined(hits) & "." & vbLf & _
               "Status set to Conflict. The studio team will contact you.", vbExclamation, "ZTN Booking"
    ElseIf shows.Count > 0 Then
        MsgBox "Submitted as " & ref & " (Requested)." & vbLf & _
               "Note: this time overlaps the regular show " & showTxt & " in " & studio & "." & vbLf & _
               "The studio engineer will decide whether it can be approved.", vbInformation, "ZTN Booking"
    Else
        MsgBox "Submitted as " & ref & " (Requested)." & vbLf & _
               "The booking is confirmed only after Studio/Technical approval.", vbInformation, "ZTN Booking"
    End If
End Sub

Public Sub ResetBookingForm()
    Dim ws As Worksheet, r As Long, c As Long, cell As Range

    If Not SheetsReady() Then Exit Sub
    If MsgBox("Clear every field on the Booking Form?", vbQuestion + vbYesNo, "Reset form") <> vbYes Then Exit Sub
    ArmSheets
    Set ws = Sh(SH_FORM)
    For r = 5 To 31
        For c = 1 To 5 Step 2
            Set cell = ws.Cells(r, c)
            If Not cell.MergeCells And Len(SafeStr(cell.Value)) > 0 Then
                cell.Offset(0, 1).MergeArea.ClearContents
            End If
        Next c
    Next r
    ws.Activate
    ws.Range("B5").Select
End Sub

Public Function GetFormValue(ByVal label As String) As Variant
    Dim ws As Worksheet, hit As Range
    GetFormValue = Empty
    Set ws = Sh(SH_FORM)
    If ws Is Nothing Then Exit Function
    Set hit = FindLabelCell(ws, label)
    If Not hit Is Nothing Then GetFormValue = hit.Offset(0, 1).Value
End Function

Public Function SetFormValue(ByVal label As String, ByVal v As Variant) As Boolean
    Dim ws As Worksheet, hit As Range
    Set ws = Sh(SH_FORM)
    If ws Is Nothing Then Exit Function
    Set hit = FindLabelCell(ws, label)
    If hit Is Nothing Then Exit Function
    hit.Offset(0, 1).Value = v
    SetFormValue = True
End Function

Private Function FindLabelCell(ByVal ws As Worksheet, ByVal label As String) As Range
    Dim r As Long, c As Long, key As String, cell As Range
    key = NormalizeStr(label)
    If key = "" Then Exit Function
    For r = 1 To 60
        For c = 1 To 6
            Set cell = ws.Cells(r, c)
            If Not cell.MergeCells Then
                If NormalizeStr(SafeStr(cell.Value)) = key Then
                    Set FindLabelCell = cell
                    Exit Function
                End If
            End If
        Next c
    Next r
End Function

Private Function FV(ByVal label As String) As String
    FV = Trim$(SafeStr(GetFormValue(label)))
End Function

Private Function FormRefIs(ByVal ref As String) As Boolean
    If Len(ref) = 0 Then Exit Function
    FormRefIs = (NormalizeStr(FV("Booking reference")) = NormalizeStr(ref))
End Function

Private Function BuildCrewSummary() As String
    Dim roles As Variant, i As Long, yesList As String
    roles = Array("Studio camera operator", "Audio operator", "Vision mixer / switcher operator", _
                  "VT / playback operator", "CG operator", "Streaming / encoder operator")
    For i = LBound(roles) To UBound(roles)
        If LCase$(FV(CStr(roles(i)))) = "yes" Then yesList = JoinFilter(Array(yesList, roles(i)), ", ")
    Next i
    BuildCrewSummary = JoinFilter(Array(FV("Production crew requested"), yesList, _
                                        FV("Technical support notes")), " | ")
End Function

' Setup access may be a clock time ("09:30") or a lead time ("30 min", "1 hr").
Private Function SetupStartMin(ByVal raw As Variant, ByVal startMin As Long) As Long
    Dim s As String, n As Double, t As Variant
    SetupStartMin = startMin
    If IsError(raw) Or IsEmpty(raw) Or IsNull(raw) Then Exit Function
    s = LCase$(Trim$(CStr(raw)))
    If s = "" Then Exit Function
    If VarType(raw) = vbString And (InStr(s, "min") > 0 Or InStr(s, "hr") > 0 Or InStr(s, "hour") > 0) Then
        n = Val(s)
        If InStr(s, "min") = 0 Then n = n * 60
        If n > 0 Then SetupStartMin = IIf(startMin - n < 0, 0, startMin - CLng(n))
        Exit Function
    End If
    t = ToMinutes(raw)
    If Not IsEmpty(t) Then
        If t < startMin Then SetupStartMin = t
    End If
End Function

'===============================================================================
'  RUNTIME - STUDIO SCHEDULE
'===============================================================================
Public Sub ApplyScheduleChanges()
    Dim ws As Worksheet, r As Long, cur As String, prev As String
    Dim nNew As Long, nStatus As Long, nRecheck As Long

    If Not SheetsReady() Then Exit Sub
    ArmSheets
    Set ws = Sh(SH_SCHED)
    Application.ScreenUpdating = False

    nNew = AssignMissingRefs(ws)
    For r = SCHED_FIRST_ROW To LastSchedRow(ws)
        If RowHasData(ws, r) Then
            cur = Trim$(SafeStr(ws.Cells(r, C_STATUS).Value))
            prev = Trim$(SafeStr(ws.Cells(r, C_LASTAPPLIED).Value))
            If StrComp(cur, prev, vbBinaryCompare) <> 0 Then
                HandleStatusChange r, cur
                nStatus = nStatus + 1
            ElseIf IsInArray(CanonicalStatus(cur), LiveStatuses()) And _
                   SafeStr(ws.Cells(r, C_KEY).Value) <> RowKey(ws, r) Then
                RecheckRowConflict r
                nRecheck = nRecheck + 1
            End If
        End If
    Next r

    RebuildWeeklyOverview
    Application.ScreenUpdating = True
    MsgBox "Apply complete." & vbLf & _
           "  New references assigned: " & nNew & vbLf & _
           "  Status changes processed: " & nStatus & vbLf & _
           "  Edited rows re-checked: " & nRecheck, vbInformation, "ZTN Schedule"
End Sub

Public Sub HandleStatusChange(ByVal r As Long, ByVal newStatus As String)
    Dim ws As Worksheet, st As String, prev As String, ref As String, title As String, studio As String
    Dim d As Variant, s As Variant, e As Variant, hits As Collection, toAddr As String, tag As String

    Set ws = Sh(SH_SCHED)
    If ws Is Nothing Then Exit Sub
    st = CanonicalStatus(newStatus)
    prev = CanonicalStatus(SafeStr(ws.Cells(r, C_LASTAPPLIED).Value))
    ref = Trim$(SafeStr(ws.Cells(r, C_REF).Value))
    title = SafeStr(ws.Cells(r, C_TITLE).Value)
    studio = Trim$(SafeStr(ws.Cells(r, C_STUDIO).Value))

    If st = "" Then
        If prev = "" Then prev = "Requested"
        StampNote ws, r, "Invalid status '" & newStatus & "' reverted to " & prev
        SetRowStatus ws, r, prev
        Exit Sub
    End If

    d = ToDateVal(ws.Cells(r, C_DATE).Value)
    s = ToMinutes(ws.Cells(r, C_START).Value)
    e = ToMinutes(ws.Cells(r, C_END).Value)

    If st = "Confirmed" Then
        If IsEmpty(d) Or IsEmpty(s) Or IsEmpty(e) Or studio = "" Then
            If prev = "" Or prev = "Confirmed" Then prev = "Requested"
            SetRowStatus ws, r, prev
            StampNote ws, r, "Cannot confirm " & DashEm() & " date, time or studio missing"
            Exit Sub
        End If
        If e <= s Then
            If prev = "" Or prev = "Confirmed" Then prev = "Requested"
            SetRowStatus ws, r, prev
            StampNote ws, r, "Cannot confirm " & DashEm() & " end time is not after start time"
            Exit Sub
        End If
        Set hits = FindConflicts(ws, d, EffectiveStart(ws, r, CLng(s)), CLng(e), studio, ref)
        If hits.Count > 0 Then
            SetRowStatus ws, r, "Conflict"
            StampNote ws, r, "Cannot confirm " & DashEm() & " clash with " & RefsJoined(hits)
            ws.Cells(r, C_KEY).Value = RowKey(ws, r)
            Exit Sub
        End If
        ' Regular shows never block: the engineer's decision is recorded instead.
        Set hits = FindShowOverlaps(d, EffectiveStart(ws, r, CLng(s)), CLng(e), studio)
        If hits.Count > 0 Then
            StampNote ws, r, "Confirmed over regular show " & RefsJoined(hits) & " at engineer's discretion"
        End If
    End If

    SetRowStatus ws, r, st
    If st <> "Confirmed" And IsInArray(st, LiveStatuses()) Then
        RecheckRowConflict r
        st = CanonicalStatus(SafeStr(ws.Cells(r, C_STATUS).Value))
    End If
    ws.Cells(r, C_KEY).Value = RowKey(ws, r)

    toAddr = ExtractEmail(SafeStr(ws.Cells(r, C_REQ).Value))
    tag = " " & DashEm() & " " & ref & " " & DashEm() & " " & title
    Select Case st
        Case "Confirmed"
            If FormRefIs(ref) Then
                SetFormValue "Studio/Technical approval", "Approved " & Format$(Now, "dd-mmm-yyyy hh:mm")
                SetFormValue "Confirmed date/time", Format$(d, "dd-mmm-yyyy") & " " & Mm(s) & DashEn() & Mm(e)
            End If
            SendMail toAddr, "[ZTN Studio] Booking confirmed" & tag, _
                "Your studio booking has been CONFIRMED." & vbLf & vbLf & BookingSummary(ws, r) & _
                vbLf & vbLf & "ZTN Studio / Technical Department"
        Case "Declined", "Cancelled"
            If FormRefIs(ref) Then
                SetFormValue "Studio/Technical approval", st & " " & Format$(Now, "dd-mmm-yyyy hh:mm")
                SetFormValue "Confirmed date/time", ""
            End If
            SendMail toAddr, "[ZTN Studio] Booking " & LCase$(st) & tag, _
                "Your studio booking has been " & UCase$(st) & "." & vbLf & vbLf & _
                "Reason / notes:" & vbLf & SafeStr(ws.Cells(r, C_NOTES).Value) & vbLf & vbLf & _
                BookingSummary(ws, r) & vbLf & vbLf & "ZTN Studio / Technical Department"
            ReleaseClashes ws, r
        Case "Rescheduled"
            SendMail toAddr, "[ZTN Studio] Booking rescheduled" & tag, _
                "Your studio booking has been rescheduled. Current details:" & vbLf & vbLf & _
                BookingSummary(ws, r) & vbLf & vbLf & "ZTN Studio / Technical Department"
    End Select
End Sub

Public Function FindConflicts(ByVal ws As Worksheet, ByVal dt As Variant, ByVal startMin As Long, _
                              ByVal endMin As Long, ByVal studio As String, ByVal excludeRef As String, _
                              Optional ByVal firmOnly As Boolean = False) As Collection
    Dim hits As New Collection, r As Long, d0 As Variant, d As Variant, s As Variant, e As Variant
    Dim ref As String, pool As Variant, stu As String

    Set FindConflicts = hits
    d0 = ToDateVal(dt)
    stu = NormalizeStr(studio)
    If IsEmpty(d0) Or stu = "" Or endMin <= startMin Then Exit Function
    If firmOnly Then pool = DisplayStatuses() Else pool = BlockingStatuses()

    For r = SCHED_FIRST_ROW To LastSchedRow(ws)
        ref = Trim$(SafeStr(ws.Cells(r, C_REF).Value))
        If ref <> "" And StrComp(ref, Trim$(excludeRef), vbTextCompare) <> 0 Then
            If IsInArray(CanonicalStatus(SafeStr(ws.Cells(r, C_STATUS).Value)), pool) Then
                If NormalizeStr(SafeStr(ws.Cells(r, C_STUDIO).Value)) = stu Then
                    d = ToDateVal(ws.Cells(r, C_DATE).Value)
                    If Not IsEmpty(d) Then
                        If d = d0 Then
                            s = ToMinutes(ws.Cells(r, C_START).Value)
                            e = ToMinutes(ws.Cells(r, C_END).Value)
                            If Not IsEmpty(s) And Not IsEmpty(e) Then
                                s = EffectiveStart(ws, r, CLng(s))
                                If s < endMin And startMin < e Then hits.Add ref
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next r
End Function

' Regular shows in the same studio that overlap the slot, as "NAME hh:mm-hh:mm".
' These are warnings for the engineer, never automatic conflicts.
Public Function FindShowOverlaps(ByVal dt As Variant, ByVal startMin As Long, ByVal endMin As Long, _
                                 ByVal studio As String) As Collection
    Dim hits As New Collection, ws As Worksheet, r As Long, d0 As Variant, s As Variant, e As Variant

    Set FindShowOverlaps = hits
    Set ws = Sh(SH_SHOWS)
    d0 = ToDateVal(dt)
    If ws Is Nothing Or IsEmpty(d0) Or NormalizeStr(studio) = "" Or endMin <= startMin Then Exit Function

    For r = SHOWS_FIRST_ROW To LastShowRow(ws)
        If ShowActiveOn(ws, r, CDate(d0)) Then
            If NormalizeStr(ShowStudio(ws, r)) = NormalizeStr(studio) Then
                s = ToMinutes(ws.Cells(r, 3).Value)
                e = ToMinutes(ws.Cells(r, 4).Value)
                If s < endMin And startMin < e Then
                    hits.Add Trim$(SafeStr(ws.Cells(r, 1).Value)) & " " & Mm(s) & DashEn() & Mm(e)
                End If
            End If
        End If
    Next r
End Function

Private Function ShowOverlapNote(ByVal shows As String) As String
    ShowOverlapNote = "Overlaps regular show " & shows & " " & DashEm() & " engineer to decide"
End Function

' True when the Weekly Shows row runs on this date (valid times, day matches, within Active From/To).
Private Function ShowActiveOn(ByVal ws As Worksheet, ByVal r As Long, ByVal d As Date) As Boolean
    Dim days As Variant, s As Variant, e As Variant, af As Variant, at As Variant
    If Len(Trim$(SafeStr(ws.Cells(r, 1).Value))) = 0 Then Exit Function
    s = ToMinutes(ws.Cells(r, 3).Value)
    e = ToMinutes(ws.Cells(r, 4).Value)
    If IsEmpty(s) Or IsEmpty(e) Then Exit Function
    If e <= s Then Exit Function
    days = ParseDays(SafeStr(ws.Cells(r, 2).Value))
    If Not days(Weekday(d, vbMonday)) Then Exit Function
    af = ToDateVal(ws.Cells(r, 8).Value)
    at = ToDateVal(ws.Cells(r, 9).Value)
    If Not IsEmpty(af) Then
        If d < af Then Exit Function
    End If
    If Not IsEmpty(at) Then
        If d > at Then Exit Function
    End If
    ShowActiveOn = True
End Function

Private Function ShowStudio(ByVal ws As Worksheet, ByVal r As Long) As String
    ShowStudio = Trim$(SafeStr(ws.Cells(r, 5).Value))
    If ShowStudio = "" Then ShowStudio = DEFAULT_SHOW_STUDIO
End Function

Private Function LastShowRow(ByVal ws As Worksheet) As Long
    LastShowRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If LastShowRow < SHOWS_FIRST_ROW Then LastShowRow = SHOWS_FIRST_ROW - 1
End Function

' "Mon", "Mon-Fri", "Tue, Thu", "Weekdays", "Weekends", "Daily" -> Boolean(1 To 7), Monday = 1.
Public Function ParseDays(ByVal txt As String) As Variant
    Dim out(1 To 7) As Boolean, parts As Variant, p As Variant, i As Long, a As Long, b As Long, dash As Long
    Dim s As String

    s = LCase$(Trim$(txt))
    s = Replace(Replace(s, ChrW(8211), "-"), ChrW(8212), "-")
    s = Replace(Replace(Replace(s, " to ", "-"), "&", ","), "/", ",")
    s = Replace(Replace(s, " and ", ","), ";", ",")
    parts = Split(s, ",")
    For Each p In parts
        p = Trim$(p)
        Select Case p
            Case "daily", "everyday", "every day", "all", "mon-sun"
                For i = 1 To 7: out(i) = True: Next i
            Case "weekdays", "weekday"
                For i = 1 To 5: out(i) = True: Next i
            Case "weekends", "weekend"
                out(6) = True: out(7) = True
            Case Else
                dash = InStr(p, "-")
                If dash > 0 Then
                    a = DayIndex(Trim$(Left$(p, dash - 1)))
                    b = DayIndex(Trim$(Mid$(p, dash + 1)))
                    If a > 0 And b > 0 Then
                        i = a
                        Do
                            out(i) = True
                            If i = b Then Exit Do
                            i = i Mod 7 + 1
                        Loop
                    End If
                Else
                    a = DayIndex(CStr(p))
                    If a > 0 Then out(a) = True
                End If
        End Select
    Next p
    ParseDays = out
End Function

Private Function DayIndex(ByVal s As String) As Long
    Dim names As Variant, i As Long
    names = Array("mon", "tue", "wed", "thu", "fri", "sat", "sun")
    s = Left$(LCase$(Trim$(s)), 3)
    For i = 0 To 6
        If s = names(i) Then DayIndex = i + 1: Exit Function
    Next i
End Function

' A Confirmed row is only knocked back by another Confirmed/Completed booking;
' any other open row becomes Conflict on any blocking overlap.
Public Sub RecheckRowConflict(ByVal r As Long)
    Dim ws As Worksheet, st As String, ref As String, studio As String
    Dim d As Variant, s As Variant, e As Variant, eff As Long, hits As Collection

    Set ws = Sh(SH_SCHED)
    If ws Is Nothing Then Exit Sub
    st = CanonicalStatus(SafeStr(ws.Cells(r, C_STATUS).Value))
    ref = Trim$(SafeStr(ws.Cells(r, C_REF).Value))
    studio = SafeStr(ws.Cells(r, C_STUDIO).Value)
    d = ToDateVal(ws.Cells(r, C_DATE).Value)
    s = ToMinutes(ws.Cells(r, C_START).Value)
    e = ToMinutes(ws.Cells(r, C_END).Value)

    If IsInArray(st, LiveStatuses()) And Not IsEmpty(d) And Not IsEmpty(s) And Not IsEmpty(e) Then
        eff = EffectiveStart(ws, r, CLng(s))
        If st = "Confirmed" Then
            Set hits = FindConflicts(ws, d, eff, CLng(e), studio, ref, True)
            If hits.Count > 0 Then
                SetRowStatus ws, r, "Conflict"
                StampNote ws, r, "Clash with confirmed " & RefsJoined(hits)
            Else
                Set hits = FindConflicts(ws, d, eff, CLng(e), studio, ref)
                If hits.Count > 0 Then StampNote ws, r, "Overlaps unconfirmed " & RefsJoined(hits)
            End If
        Else
            Set hits = FindConflicts(ws, d, eff, CLng(e), studio, ref)
            If hits.Count > 0 Then
                If st <> "Conflict" Then SetRowStatus ws, r, "Conflict"
                StampNote ws, r, "Clash with " & RefsJoined(hits)
            ElseIf st = "Conflict" Then
                SetRowStatus ws, r, "Requested"
                StampNote ws, r, "Clash cleared"
            End If
        End If
        Set hits = FindShowOverlaps(d, eff, CLng(e), studio)
        If hits.Count > 0 And st <> "Confirmed" Then StampNote ws, r, ShowOverlapNote(RefsJoined(hits))
    End If
    ws.Cells(r, C_KEY).Value = RowKey(ws, r)
End Sub

Public Sub RecheckAllConflicts()
    Dim ws As Worksheet, r As Long, before As String, n As Long

    If Not SheetsReady() Then Exit Sub
    ArmSheets
    Set ws = Sh(SH_SCHED)
    Application.ScreenUpdating = False
    AssignMissingRefs ws
    For r = SCHED_FIRST_ROW To LastSchedRow(ws)
        If RowHasData(ws, r) Then
            before = SafeStr(ws.Cells(r, C_STATUS).Value)
            If IsInArray(CanonicalStatus(before), LiveStatuses()) Then
                RecheckRowConflict r
                If SafeStr(ws.Cells(r, C_STATUS).Value) <> before Then n = n + 1
            End If
        End If
    Next r
    RebuildWeeklyOverview
    Application.ScreenUpdating = True
    MsgBox "Conflict re-check complete. " & n & " row(s) changed status.", vbInformation, "ZTN Schedule"
End Sub

' Once a booking is declined/cancelled, conflicts it caused on the same day and studio may clear.
Private Sub ReleaseClashes(ByVal ws As Worksheet, ByVal r As Long)
    Dim rr As Long, d As Variant, stu As String, d2 As Variant
    d = ToDateVal(ws.Cells(r, C_DATE).Value)
    If IsEmpty(d) Then Exit Sub
    stu = NormalizeStr(SafeStr(ws.Cells(r, C_STUDIO).Value))
    For rr = SCHED_FIRST_ROW To LastSchedRow(ws)
        If rr <> r Then
            If CanonicalStatus(SafeStr(ws.Cells(rr, C_STATUS).Value)) = "Conflict" Then
                If NormalizeStr(SafeStr(ws.Cells(rr, C_STUDIO).Value)) = stu Then
                    d2 = ToDateVal(ws.Cells(rr, C_DATE).Value)
                    If Not IsEmpty(d2) Then
                        If d2 = d Then RecheckRowConflict rr
                    End If
                End If
            End If
        End If
    Next rr
End Sub

Public Sub CloseCompletedBookings()
    Dim n As Long
    If Not SheetsReady() Then Exit Sub
    ArmSheets
    n = CloseCompletedCore()
    RebuildWeeklyOverview
    MsgBox n & " booking(s) moved to Completed.", vbInformation, "ZTN Schedule"
End Sub

Private Function CloseCompletedCore() As Long
    Dim ws As Worksheet, r As Long, d As Variant, e As Variant, n As Long
    Set ws = Sh(SH_SCHED)
    If ws Is Nothing Then Exit Function
    For r = SCHED_FIRST_ROW To LastSchedRow(ws)
        If CanonicalStatus(SafeStr(ws.Cells(r, C_STATUS).Value)) = "Confirmed" Then
            d = ToDateVal(ws.Cells(r, C_DATE).Value)
            e = ToMinutes(ws.Cells(r, C_END).Value)
            If Not IsEmpty(d) And Not IsEmpty(e) Then
                If CDbl(CDate(d)) + e / 1440 < CDbl(Now) Then
                    SetRowStatus ws, r, "Completed"
                    StampNote ws, r, "Auto-completed"
                    n = n + 1
                End If
            End If
        End If
    Next r
    CloseCompletedCore = n
End Function

Public Sub ScheduleAutoClose()
    On Error Resume Next
    If mNextAutoClose > Now Then
        Application.OnTime mNextAutoClose, "'" & ThisWorkbook.Name & "'!ScheduleAutoClose", , False
    End If
    On Error GoTo 0
    If SheetsReady(True) Then
        ArmSheets
        If CloseCompletedCore() > 0 Then RebuildWeeklyOverview
    End If
    mNextAutoClose = Now + TimeSerial(1, 0, 0)
    Application.OnTime mNextAutoClose, "'" & ThisWorkbook.Name & "'!ScheduleAutoClose"
End Sub

Public Sub StopAutoClose()
    On Error Resume Next
    If mNextAutoClose > Now Then
        Application.OnTime mNextAutoClose, "'" & ThisWorkbook.Name & "'!ScheduleAutoClose", , False
    End If
    mNextAutoClose = 0
End Sub

'===============================================================================
'  RUNTIME - WEEKLY OVERVIEW
'===============================================================================
Public Sub RebuildWeeklyOverview()
    Dim wsW As Worksheet, wsS As Worksheet, v As Variant, wkStart As Date, b As Variant
    Dim wsH As Worksheet, nSlots As Long, r As Long, k As Long, j As Long, m As Long
    Dim txt() As String, cnt() As Long, studs() As String, clash() As Boolean, owner() As String, booked() As Boolean
    Dim d As Variant, s As Variant, e As Variant, dayIdx As Long, nm As String, stu As String
    Dim shown As Long, showSlots As Long, cell As Range, grid As Range

    If Not SheetsReady(True) Then Exit Sub
    ArmSheets
    Set wsW = Sh(SH_WEEK)
    Set wsS = Sh(SH_SCHED)
    Set wsH = Sh(SH_SHOWS)

    v = ToDateVal(wsW.Range("B2").Value)
    If IsEmpty(v) Then v = Date
    wkStart = MondayOf(CDate(v))
    wsW.Range("B2").Value = wkStart

    b = SlotBounds()
    nSlots = UBound(b) - LBound(b)
    ReDim txt(1 To nSlots, 1 To 7)
    ReDim cnt(1 To nSlots, 1 To 7)
    ReDim studs(1 To nSlots, 1 To 7)
    ReDim clash(1 To nSlots, 1 To 7)
    ReDim owner(1 To nSlots, 1 To 7)
    ReDim booked(1 To nSlots, 1 To 7)

    ' Regular weekly shows
    For r = SHOWS_FIRST_ROW To LastShowRow(wsH)
        For dayIdx = 0 To 6
            If ShowActiveOn(wsH, r, wkStart + dayIdx) Then
                nm = Trim$(SafeStr(wsH.Cells(r, 1).Value))
                stu = ShowStudio(wsH, r)
                s = ToMinutes(wsH.Cells(r, 3).Value)
                e = ToMinutes(wsH.Cells(r, 4).Value)
                showSlots = showSlots + PlaceOnGrid(txt, cnt, studs, clash, owner, booked, b, dayIdx + 1, _
                    CLng(s), CLng(e), stu, nm & vbLf & Mm(s) & DashEn() & Mm(e) & " " & ChrW(183) & " " & stu, _
                    ChrW(8627) & " " & nm & " (cont.)", "SHOW" & r, False)
            End If
        Next dayIdx
    Next r

    ' Confirmed / Completed bookings
    For r = SCHED_FIRST_ROW To LastSchedRow(wsS)
        If IsInArray(CanonicalStatus(SafeStr(wsS.Cells(r, C_STATUS).Value)), DisplayStatuses()) Then
            d = ToDateVal(wsS.Cells(r, C_DATE).Value)
            s = ToMinutes(wsS.Cells(r, C_START).Value)
            e = ToMinutes(wsS.Cells(r, C_END).Value)
            If Not IsEmpty(d) And Not IsEmpty(s) And Not IsEmpty(e) Then
                dayIdx = CLng(CDate(d) - wkStart)
                If dayIdx >= 0 And dayIdx <= 6 And e > s Then
                    stu = SafeStr(wsS.Cells(r, C_STUDIO).Value)
                    If PlaceOnGrid(txt, cnt, studs, clash, owner, booked, b, dayIdx + 1, CLng(s), CLng(e), stu, _
                        SafeStr(wsS.Cells(r, C_REF).Value) & " " & ChrW(183) & " " & stu & vbLf & _
                        SafeStr(wsS.Cells(r, C_TITLE).Value) & vbLf & Mm(s) & DashEn() & Mm(e), _
                        ChrW(8627) & " " & SafeStr(wsS.Cells(r, C_REF).Value) & " (cont.)", _
                        "BK" & r, True) > 0 Then shown = shown + 1
                End If
            End If
        End If
    Next r

    Set grid = wsW.Range(wsW.Cells(WEEK_FIRST_ROW, 2), wsW.Cells(WEEK_FIRST_ROW + nSlots - 1, 8))
    With grid
        .UnMerge
        .ClearContents
        .Interior.Color = RGB(255, 255, 255)
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    For k = 1 To nSlots
        For j = 1 To 7
            If cnt(k, j) > 0 Then
                Set cell = wsW.Cells(WEEK_FIRST_ROW + k - 1, j + 1)
                cell.Value = txt(k, j)
                If clash(k, j) Then
                    cell.Interior.Color = RGB(255, 199, 206)
                ElseIf booked(k, j) Then
                    cell.Interior.Color = RGB(221, 235, 247)
                Else
                    cell.Interior.Color = RGB(226, 239, 218)
                End If
            End If
        Next j
    Next k
    wsW.Range(wsW.Cells(WEEK_FIRST_ROW, 1), wsW.Cells(WEEK_FIRST_ROW + nSlots - 1, 8)).Rows.AutoFit
    For k = WEEK_FIRST_ROW To WEEK_FIRST_ROW + nSlots - 1
        If wsW.Rows(k).RowHeight < 24 Then wsW.Rows(k).RowHeight = 24
    Next k

    ' Merge consecutive slots that hold one and the same item, like a hand-drawn grid.
    For j = 1 To 7
        k = 1
        Do While k <= nSlots
            m = k
            If cnt(k, j) = 1 Then
                Do While m < nSlots
                    If cnt(m + 1, j) <> 1 Or owner(m + 1, j) <> owner(k, j) Then Exit Do
                    m = m + 1
                Loop
            End If
            If m > k Then
                wsW.Range(wsW.Cells(WEEK_FIRST_ROW + k, j + 1), wsW.Cells(WEEK_FIRST_ROW + m - 1, j + 1)).ClearContents
                wsW.Range(wsW.Cells(WEEK_FIRST_ROW + k - 1, j + 1), wsW.Cells(WEEK_FIRST_ROW + m - 1, j + 1)).Merge
            End If
            k = m + 1
        Loop
    Next j

    wsW.Cells(WEEK_FIRST_ROW + nSlots, 1).Value = "Week " & Format$(wkStart, "dd-mmm-yyyy") & " " & DashEn() & " " & _
        Format$(wkStart + 6, "dd-mmm-yyyy") & "  " & ChrW(183) & "  " & shown & " booking(s)  " & ChrW(183) & _
        "  regular shows from the Weekly Shows sheet  " & ChrW(183) & "  refreshed " & Format$(Now, "dd-mmm hh:mm")
End Sub

' Adds one item to every overview slot it touches. Returns the number of slots used.
' A cell is flagged as a clash when the same studio appears in it twice.
Private Function PlaceOnGrid(ByRef txt() As String, ByRef cnt() As Long, ByRef studs() As String, _
                             ByRef clash() As Boolean, ByRef owner() As String, ByRef booked() As Boolean, _
                             ByVal b As Variant, ByVal dayCol As Long, ByVal s As Long, ByVal e As Long, _
                             ByVal studio As String, ByVal firstTxt As String, ByVal contTxt As String, _
                             ByVal key As String, ByVal isBooking As Boolean) As Long
    Dim k As Long, n As Long, stuKey As String, sep As String
    sep = vbLf & DashEm() & " " & DashEm() & " " & DashEm() & vbLf
    stuKey = "|" & NormalizeStr(studio) & "|"
    For k = 1 To UBound(b) - LBound(b)
        If s < b(LBound(b) + k) And e > b(LBound(b) + k - 1) Then
            If cnt(k, dayCol) > 0 Then txt(k, dayCol) = txt(k, dayCol) & sep
            txt(k, dayCol) = txt(k, dayCol) & IIf(n = 0, firstTxt, contTxt)
            cnt(k, dayCol) = cnt(k, dayCol) + 1
            owner(k, dayCol) = IIf(cnt(k, dayCol) = 1, key, "")
            If InStr(studs(k, dayCol), stuKey) > 0 Then clash(k, dayCol) = True
            studs(k, dayCol) = studs(k, dayCol) & stuKey
            If isBooking Then booked(k, dayCol) = True
            n = n + 1
        End If
    Next k
    PlaceOnGrid = n
End Function

'===============================================================================
'  SCHEDULE HELPERS
'===============================================================================
Private Sub SetRowStatus(ByVal ws As Worksheet, ByVal r As Long, ByVal st As String)
    ws.Cells(r, C_STATUS).Value = st
    ws.Cells(r, C_LASTAPPLIED).Value = st
    ColorStatus ws, r
    If FormRefIs(SafeStr(ws.Cells(r, C_REF).Value)) Then SetFormValue "Booking status", st
End Sub

Private Sub ColorStatus(ByVal ws As Worksheet, ByVal r As Long)
    Dim c As Range, bg As Long, fg As Long
    Set c = ws.Cells(r, C_STATUS)
    Select Case CanonicalStatus(SafeStr(c.Value))
        Case "Confirmed": bg = RGB(198, 239, 206): fg = RGB(0, 97, 0)
        Case "Completed": bg = RGB(217, 217, 217): fg = RGB(89, 89, 89)
        Case "Pending Approval": bg = RGB(255, 235, 156): fg = RGB(156, 101, 0)
        Case "Conflict", "Declined", "Cancelled": bg = RGB(255, 199, 206): fg = RGB(156, 0, 6)
        Case "Rescheduled": bg = RGB(222, 235, 247): fg = RGB(31, 78, 121)
        Case Else
            c.Interior.Pattern = xlNone
            c.Font.ColorIndex = xlAutomatic
            c.Font.Bold = False
            Exit Sub
    End Select
    c.Interior.Color = bg
    c.Font.Color = fg
    c.Font.Bold = True
End Sub

Private Function EffectiveStart(ByVal ws As Worksheet, ByVal r As Long, ByVal startMin As Long) As Long
    Dim p As Variant
    EffectiveStart = startMin
    p = ToMinutes(ws.Cells(r, C_SETUP).Value)
    If Not IsEmpty(p) Then
        If p < startMin Then EffectiveStart = p
    End If
End Function

Private Function RowKey(ByVal ws As Worksheet, ByVal r As Long) As String
    Dim d As Variant
    d = ToDateVal(ws.Cells(r, C_DATE).Value)
    If Not IsEmpty(d) Then d = CLng(CDate(d))
    RowKey = CStr(d) & "|" & CStr(ToMinutes(ws.Cells(r, C_START).Value)) & "|" & _
             CStr(ToMinutes(ws.Cells(r, C_END).Value)) & "|" & _
             NormalizeStr(SafeStr(ws.Cells(r, C_STUDIO).Value)) & "|" & _
             CStr(ToMinutes(ws.Cells(r, C_SETUP).Value))
End Function

Private Function AssignMissingRefs(ByVal ws As Worksheet) As Long
    Dim r As Long, n As Long
    For r = SCHED_FIRST_ROW To LastSchedRow(ws)
        If RowHasData(ws, r) Then
            If Len(Trim$(SafeStr(ws.Cells(r, C_REF).Value))) = 0 Then
                FormatSchedRow ws, r
                ws.Cells(r, C_REF).Value = NextBookingRef()
                n = n + 1
            End If
            If Len(Trim$(SafeStr(ws.Cells(r, C_STATUS).Value))) = 0 Then ws.Cells(r, C_STATUS).Value = "Requested"
        End If
    Next r
    AssignMissingRefs = n
End Function

' Rows past the pre-built range need their Day formula, formats and lists.
Private Sub FormatSchedRow(ByVal ws As Worksheet, ByVal r As Long)
    If r <= SCHED_PREP_LAST_ROW Then Exit Sub
    ws.Cells(r, C_DAY).Formula = "=IF(B" & r & "="""","""",TEXT(B" & r & ",""dddd""))"
    ws.Cells(r, C_DATE).NumberFormat = "dd-mmm-yyyy"
    ws.Range(ws.Cells(r, C_START), ws.Cells(r, C_END)).NumberFormat = "hh:mm"
    ws.Cells(r, C_SETUP).NumberFormat = "hh:mm"
    ApplyList ws.Cells(r, C_STATUS), StatusList(), True
    ApplyList ws.Cells(r, C_STUDIO), StudioList(), True
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, C_NOTES))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(217, 217, 217)
        .VerticalAlignment = xlTop
    End With
    ws.Range(ws.Cells(r, C_REF), ws.Cells(r, C_DATE)).Locked = False
    ws.Range(ws.Cells(r, C_START), ws.Cells(r, C_NOTES)).Locked = False
End Sub

Private Function LastSchedRow(ByVal ws As Worksheet) As Long
    Dim c As Variant, r As Long
    LastSchedRow = SCHED_FIRST_ROW - 1
    For Each c In Array(C_REF, C_DATE, C_TITLE)
        r = ws.Cells(ws.Rows.Count, CLng(c)).End(xlUp).Row
        If r > LastSchedRow Then LastSchedRow = r
    Next c
End Function

Private Function RowHasData(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    RowHasData = Len(Trim$(SafeStr(ws.Cells(r, C_REF).Value))) + _
                 Len(Trim$(SafeStr(ws.Cells(r, C_DATE).Value))) + _
                 Len(Trim$(SafeStr(ws.Cells(r, C_TITLE).Value))) > 0
End Function

Private Function CanonicalStatus(ByVal s As String) As String
    Dim v As Variant, key As String
    key = NormalizeStr(s)
    If key = "" Then Exit Function
    For Each v In StatusList()
        If NormalizeStr(CStr(v)) = key Then CanonicalStatus = CStr(v): Exit Function
    Next v
End Function

Private Function BookingSummary(ByVal ws As Worksheet, ByVal r As Long) As String
    Dim d As Variant, s As Variant, e As Variant, whenTxt As String
    d = ToDateVal(ws.Cells(r, C_DATE).Value)
    s = ToMinutes(ws.Cells(r, C_START).Value)
    e = ToMinutes(ws.Cells(r, C_END).Value)
    If Not IsEmpty(d) Then whenTxt = Format$(d, "dddd dd mmmm yyyy")
    If Not IsEmpty(s) And Not IsEmpty(e) Then whenTxt = whenTxt & ", " & Mm(s) & " - " & Mm(e)
    BookingSummary = "Reference:    " & SafeStr(ws.Cells(r, C_REF).Value) & vbLf & _
        "Date / time:  " & whenTxt & vbLf & _
        "Studio:       " & SafeStr(ws.Cells(r, C_STUDIO).Value) & vbLf & _
        "Programme:    " & SafeStr(ws.Cells(r, C_TITLE).Value) & vbLf & _
        "Requester:    " & SafeStr(ws.Cells(r, C_REQ).Value) & vbLf & _
        "Producer:     " & SafeStr(ws.Cells(r, C_PROD).Value) & vbLf & _
        "Live/Rec:     " & SafeStr(ws.Cells(r, C_LIVE).Value) & vbLf & _
        "Crew:         " & SafeStr(ws.Cells(r, C_CREW).Value) & vbLf & _
        "Priority:     " & SafeStr(ws.Cells(r, C_PRIO).Value) & vbLf & _
        "Status:       " & SafeStr(ws.Cells(r, C_STATUS).Value) & vbLf & _
        "Notes:        " & SafeStr(ws.Cells(r, C_NOTES).Value)
End Function

'===============================================================================
'  GENERAL HELPERS
'===============================================================================
Private Function Sh(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set Sh = ThisWorkbook.Worksheets(nm)
End Function

Private Function SheetsReady(Optional ByVal silent As Boolean = False) As Boolean
    Dim nm As Variant
    For Each nm In Array(SH_FORM, SH_SCHED, SH_WEEK, SH_SHOWS, SH_HOWTO)
        If Sh(CStr(nm)) Is Nothing Then
            If Not silent Then MsgBox "Sheet '" & nm & "' is missing. Run BuildZTNWorkbook first.", vbExclamation, "ZTN"
            Exit Function
        End If
    Next nm
    SheetsReady = True
End Function

Private Sub ArmSheets()
    Dim nm As Variant, ws As Worksheet
    For Each nm In Array(SH_FORM, SH_SCHED, SH_WEEK, SH_SHOWS, SH_HOWTO)
        Set ws = Sh(CStr(nm))
        If Not ws Is Nothing Then ProtectSheet ws
    Next nm
End Sub

Private Function SafeStr(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    If IsObject(v) Then Exit Function
    SafeStr = CStr(v)
End Function

Private Function IsAllDigits(ByVal s As String) As Boolean
    IsAllDigits = Len(s) > 0 And Not (s Like "*[!0-9]*")
End Function

Public Function ToDateVal(ByVal v As Variant) As Variant
    Dim d As Date
    ToDateVal = Empty
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    On Error GoTo Bad
    If VarType(v) = vbDate Then
        d = v
    ElseIf VarType(v) <> vbString And IsNumeric(v) Then
        If CDbl(v) < 1 Then Exit Function
        d = CDate(Int(CDbl(v)))
    ElseIf IsDate(CStr(v)) Then
        d = CDate(CStr(v))
    Else
        Exit Function
    End If
    If CDbl(d) < 1 Then Exit Function
    ToDateVal = DateSerial(Year(d), Month(d), Day(d))
    Exit Function
Bad:
    ToDateVal = Empty
End Function

' Minutes since midnight (Long) or Empty. Accepts Excel times, date-times,
' numbers like 10 / 10.3 / 1000, and text like "10:00", "1000", "10.00", "10 am", "10:00 pm".
Public Function ToMinutes(ByVal v As Variant) As Variant
    Dim d As Double, s As String, am As Boolean, pm As Boolean
    Dim h As Long, m As Long, p As Long, hs As String, ms As String

    ToMinutes = Empty
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function

    If VarType(v) = vbDate Then
        d = CDbl(v)
        If d = 1 Then
            ToMinutes = 1440
        Else
            ToMinutes = CLng(Round((d - Int(d)) * 1440, 0))
        End If
        Exit Function
    End If

    If VarType(v) <> vbString And IsNumeric(v) Then
        d = CDbl(v)
        If d >= 0 And d < 1 Then
            ToMinutes = CLng(Round(d * 1440, 0))
        ElseIf d >= 1 And d <= 24 Then
            h = Int(d)
            m = CLng(Round((d - h) * 100, 0))
            If m < 60 And h * 60 + m <= 1440 Then ToMinutes = h * 60 + m
        ElseIf d >= 100 And d <= 2400 And d = Int(d) Then
            h = CLng(d) \ 100
            m = CLng(d) Mod 100
            If m < 60 And h * 60 + m <= 1440 Then ToMinutes = h * 60 + m
        ElseIf d > 2400 Then
            ToMinutes = CLng(Round((d - Int(d)) * 1440, 0))
        End If
        Exit Function
    End If

    s = Replace(LCase$(Trim$(CStr(v))), " ", "")
    s = Replace(Replace(s, "a.m.", "am"), "p.m.", "pm")
    If Right$(s, 2) = "pm" Then
        pm = True: s = Left$(s, Len(s) - 2)
    ElseIf Right$(s, 2) = "am" Then
        am = True: s = Left$(s, Len(s) - 2)
    End If
    s = Replace(Replace(s, ".", ":"), "h", ":")
    If s = "" Then Exit Function

    p = InStr(s, ":")
    If p > 0 Then
        hs = Left$(s, p - 1)
        ms = Mid$(s, p + 1)
        If InStr(ms, ":") > 0 Then ms = Left$(ms, InStr(ms, ":") - 1)
        If ms = "" Then ms = "0"
    ElseIf Len(s) <= 2 Then
        hs = s: ms = "0"
    ElseIf Len(s) <= 4 Then
        hs = Left$(s, Len(s) - 2): ms = Right$(s, 2)
    Else
        Exit Function
    End If
    If Not IsAllDigits(hs) Or Not IsAllDigits(ms) Then Exit Function
    If Len(hs) > 2 Or Len(ms) > 2 Then Exit Function

    h = CLng(hs): m = CLng(ms)
    If (am Or pm) And (h < 1 Or h > 12) Then Exit Function
    If pm And h < 12 Then h = h + 12
    If am And h = 12 Then h = 0
    If m > 59 Or h * 60 + m > 1440 Then Exit Function
    ToMinutes = h * 60 + m
End Function

Public Function Mm(ByVal minutes As Long) As String
    Mm = Format$(minutes \ 60, "00") & ":" & Format$(minutes Mod 60, "00")
End Function

Public Function MondayOf(ByVal d As Date) As Date
    MondayOf = DateSerial(Year(d), Month(d), Day(d)) - (Weekday(d, vbMonday) - 1)
End Function

Public Function NextBookingRef() As String
    Dim ws As Worksheet, r As Long, pfx As String, s As String, n As Long, mx As Long
    pfx = "ZTN-" & Year(Date) & "-"
    Set ws = Sh(SH_SCHED)
    If Not ws Is Nothing Then
        For r = SCHED_FIRST_ROW To LastSchedRow(ws)
            s = UCase$(Trim$(SafeStr(ws.Cells(r, C_REF).Value)))
            If Left$(s, Len(pfx)) = pfx Then
                n = CLng(Val(Mid$(s, Len(pfx) + 1)))
                If n > mx Then mx = n
            End If
        Next r
    End If
    NextBookingRef = pfx & Format$(mx + 1, "0000")
End Function

Public Function IsInArray(ByVal v As Variant, ByVal arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If StrComp(SafeStr(v), SafeStr(arr(i)), vbTextCompare) = 0 Then
            IsInArray = True
            Exit Function
        End If
    Next i
End Function

Public Function JoinFilter(ByVal arr As Variant, Optional ByVal sep As String = ", ") As String
    Dim i As Long, s As String, out As String
    For i = LBound(arr) To UBound(arr)
        s = Trim$(SafeStr(arr(i)))
        If Len(s) > 0 Then
            If Len(out) > 0 Then out = out & sep
            out = out & s
        End If
    Next i
    JoinFilter = out
End Function

Public Function RefsJoined(ByVal refs As Collection) As String
    Dim v As Variant, out As String
    For Each v In refs
        If Len(out) > 0 Then out = out & ", "
        out = out & CStr(v)
    Next v
    RefsJoined = out
End Function

Public Sub StampNote(ByVal ws As Worksheet, ByVal r As Long, ByVal note As String)
    Dim cur As String, stamp As String
    cur = SafeStr(ws.Cells(r, C_NOTES).Value)
    If InStr(1, cur, note, vbTextCompare) > 0 Then Exit Sub
    stamp = "[" & Format$(Now, "dd-mmm hh:mm") & "] " & note
    If Len(cur) = 0 Then
        ws.Cells(r, C_NOTES).Value = stamp
    Else
        ws.Cells(r, C_NOTES).Value = cur & vbLf & stamp
    End If
End Sub

Public Function ExtractEmail(ByVal s As String) As String
    Dim p As Long, a As Long, b As Long, cand As String, at As Long
    p = InStr(s, "@")
    Do While p > 0
        a = p - 1
        Do While a >= 1
            If Not (Mid$(s, a, 1) Like "[A-Za-z0-9._%+-]") Then Exit Do
            a = a - 1
        Loop
        b = p + 1
        Do While b <= Len(s)
            If Not (Mid$(s, b, 1) Like "[A-Za-z0-9.-]") Then Exit Do
            b = b + 1
        Loop
        cand = Mid$(s, a + 1, b - a - 1)
        Do While Right$(cand, 1) = "."
            cand = Left$(cand, Len(cand) - 1)
        Loop
        at = InStr(cand, "@")
        If at > 1 And InStr(at + 2, cand, ".") > 0 Then
            ExtractEmail = cand
            Exit Function
        End If
        p = InStr(p + 1, s, "@")
    Loop
End Function

Public Function NormalizeStr(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    s = LCase$(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[a-z0-9]" Then out = out & ch
    Next i
    NormalizeStr = out
End Function

' Outlook via late-bound COM. No Outlook, no address, or any send error = silent no-op.
' Outlook is only started when a mail profile exists; without one it opens a
' modal "New Profile" wizard that would freeze Excel.
Public Sub SendMail(ByVal toAddr As String, ByVal subject As String, ByVal body As String)
    Dim ol As Object, mi As Object
    If Len(Trim$(toAddr)) = 0 Or InStr(toAddr, "@") = 0 Then Exit Sub
    On Error Resume Next
    Set ol = GetObject(, "Outlook.Application")
    If ol Is Nothing Then
        If Not OutlookHasProfile() Then Exit Sub
        Set ol = CreateObject("Outlook.Application")
    End If
    If ol Is Nothing Then Exit Sub
    Set mi = ol.CreateItem(0)
    If mi Is Nothing Then Exit Sub
    mi.To = toAddr
    mi.Subject = subject
    mi.Body = body
    mi.Send
    Err.Clear
End Sub

Private Function OutlookHasProfile() As Boolean
    Dim v As Variant
    On Error Resume Next
    v = CreateObject("WScript.Shell").RegRead("HKCU\Software\Microsoft\Office\" & _
                                              Application.Version & "\Outlook\DefaultProfile")
    OutlookHasProfile = (Len(CStr(v)) > 0)
End Function
