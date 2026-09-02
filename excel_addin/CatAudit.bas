Attribute VB_Name = "CatAudit"
'==============================================================================
' Cat Asset Management V2 - the write log and the before-images
'
' The README named this gap outright: nothing here writes a CCAT_AUDIT row, and
' manual edits are the ones most worth having a record of. This closes it.
'
' Every row a Run sends gets one line: who, when, from which workbook and sheet,
' what was sent, what came back - AND the state of the record BEFORE it was
' touched. Those are the same line on purpose. A log that says what happened and
' a snapshot that says what it was are the same fact recorded twice if kept
' apart, and they drift.
'
' The before-image is what makes an Expire survivable. Expire removes an
' ownership record; without a before-image the only route back is somebody's
' memory. CatUndo reads these lines back into an Add/Update sheet.
'
' WHERE IT LIVES: %APPDATA%\CatAssetTools\cat-write-log-YYYY-MM.csv
' Per user, one file a month, outside the workbook. Not a hidden sheet: the log
' has to outlive the workbook that happened to be open, and has to be there when
' someone asks about a run from three weeks ago in a file that has since been
' renamed. CCAT > Write Log opens the folder.
'
' EVERY FIELD IS QUOTED. A DCN name with a comma in it - "SMS RENTAL (WA) PTY
' LTD, INC" - would otherwise shift every later column by one, and a log that
' misreports which record was touched is worse than no log.
'
' Depends on CatAssetLookup (HeaderArray) - the before-image columns are built
' from it, so the log and the lookup can never disagree about field order.
'==============================================================================
Option Explicit

' Whether a before-image was captured, which is NOT the same question as
' whether it is blank:
'     FOUND       - the record existed and is recorded below
'     NONE        - looked, and CCAT held no such record (nothing to restore)
'     UNAVAILABLE - the lookup failed, so we do not know
' Undo must tell these apart. Restoring "all blanks" because a lookup timed out
' would write emptiness over a live record.
Public Const BEFORE_FOUND       As String = "FOUND"
Public Const BEFORE_NONE        As String = "NONE"
Public Const BEFORE_UNAVAILABLE As String = "UNAVAILABLE"

'==============================================================================
' PUBLIC: where it all goes
'==============================================================================
Public Function AuditFolder() As String
    Dim p As String
    p = Environ$("APPDATA") & "\CatAssetTools"
    If Len(Dir$(p, vbDirectory)) = 0 Then MkDir p
    AuditFolder = p
End Function

' One file per month. Long enough to keep a run together, short enough that the
' file stays openable in Excel.
Public Function AuditPath() As String
    AuditPath = AuditFolder() & "\cat-write-log-" & Format$(Now, "yyyy-mm") & ".csv"
End Function

' Ribbon: show someone the log without making them know where it is.
'
' Says WHY the log is empty rather than opening an empty folder. The log covers
' writes only, and someone who has just done two lookups has every reason to
' expect a "write log" to have something in it. An empty Explorer window
' answers nothing - it reads as broken.
Public Sub CatOpenAuditFolder()
    On Error GoTo Fail

    Dim p As String: p = AuditPath()

    If Len(Dir$(p)) = 0 Then
        MsgBox "Nothing has been logged yet this month." & vbCrLf & vbCrLf & _
               "The write log records what RUN sends to CCAT - add / update, " & _
               "expire and transfer. Lookups are deliberately not logged: a " & _
               "search changes nothing, so there is nothing to record and " & _
               "nothing to undo." & vbCrLf & vbCrLf & _
               "The file appears the first time you Run a sheet:" & vbCrLf & _
               p, _
               vbInformation, "Cat Asset Tools - Write Log"
        Exit Sub
    End If

    ' /select, opens the folder with the file already highlighted, which beats
    ' dropping someone into a directory listing to go hunting.
    Dim runs As Collection: Set runs = AuditRunIds()
    Shell "explorer.exe /select,""" & p & """", vbNormalFocus

    If runs.Count > 0 Then
        Application.StatusBar = runs.Count & " run(s) logged this month - " & p
    End If
    Exit Sub

Fail:
    MsgBox "The write log lives in:" & vbCrLf & vbCrLf & AuditFolder() & vbCrLf & vbCrLf & _
           "(Could not open it automatically: " & Err.Description & ")", _
           vbInformation, "Cat Asset Tools - Write Log"
End Sub

'==============================================================================
' PUBLIC: one run
'==============================================================================

' Called once per Run. Returns the id that ties that run's lines together, and
' is what you quote to CatUndo.
Public Function AuditBegin() As String
    AuditBegin = Format$(Now, "yyyymmdd-hhnnss")
    EnsureHeader
End Function

' One line per row actually sent.
'
' `before` is the RecordValues array as it stood before the send, or Empty when
' beforeState is not FOUND.
Public Sub AuditRow(ByVal runId As String, ByVal ws As Worksheet, ByVal r As Long, _
                    ByVal op As String, ByVal serial As String, ByVal dcn As String, _
                    ByVal outcome As String, ByVal beforeState As String, _
                    ByVal before As Variant)
    On Error Resume Next          ' logging must never break a run

    Dim line As String
    line = CsvField(runId) & "," & _
           CsvField(Format$(Now, "yyyy-mm-dd hh:nn:ss")) & "," & _
           CsvField(Environ$("USERNAME")) & "," & _
           CsvField(ws.Parent.Name) & "," & _
           CsvField(ws.Name) & "," & _
           CsvField(CStr(r)) & "," & _
           CsvField(op) & "," & _
           CsvField(serial) & "," & _
           CsvField(dcn) & "," & _
           CsvField(outcome) & "," & _
           CsvField(beforeState)

    Dim h As Variant: h = HeaderArray()
    Dim i As Long
    For i = LBound(h) To UBound(h)
        If beforeState = BEFORE_FOUND And IsArray(before) Then
            line = line & "," & CsvField(CStr(before(i)))
        Else
            line = line & "," & CsvField("")
        End If
    Next i

    AppendLine line
End Sub

'==============================================================================
' PUBLIC: reading it back (used by Undo)
'==============================================================================

' Every logged line for one run id, as a Collection of string arrays split back
' out of the CSV. Empty collection when the run isn't in this month's file.
Public Function AuditLinesFor(ByVal runId As String) As Collection
    Set AuditLinesFor = New Collection

    Dim p As String: p = AuditPath()
    If Len(Dir$(p)) = 0 Then Exit Function

    Dim f As Integer, line As String, fields As Variant
    f = FreeFile
    On Error GoTo Done
    Open p For Input As #f
    Do Until EOF(f)
        Line Input #f, line
        fields = SplitCsv(line)
        If IsArray(fields) Then
            If UBound(fields) >= 0 Then
                If fields(0) = runId Then AuditLinesFor.Add fields
            End If
        End If
    Loop
Done:
    On Error Resume Next        ' Close on a handle Open never got to
    Close #f
End Function

' The run ids in this month's file, newest first - so Undo can offer a list
' rather than asking someone to remember a timestamp.
Public Function AuditRunIds() As Collection
    Set AuditRunIds = New Collection

    Dim p As String: p = AuditPath()
    If Len(Dir$(p)) = 0 Then Exit Function

    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim ordered As Collection: Set ordered = New Collection

    Dim f As Integer, line As String, fields As Variant
    f = FreeFile
    On Error GoTo Done
    Open p For Input As #f
    Do Until EOF(f)
        Line Input #f, line
        fields = SplitCsv(line)
        If IsArray(fields) Then
            If UBound(fields) >= 6 Then
                If fields(0) <> "RunId" And Not seen.Exists(fields(0)) Then
                    seen.Add fields(0), True
                    ordered.Add Array(fields(0), fields(1), fields(6), fields(4))
                End If
            End If
        End If
    Loop
Done:
    On Error Resume Next        ' Close on a handle Open never got to
    Close #f

    ' Newest first.
    Dim i As Long
    For i = ordered.Count To 1 Step -1
        AuditRunIds.Add ordered(i)
    Next i
End Function

'==============================================================================
' PRIVATE
'==============================================================================

' Header written once, when the month's file is created. Built from
' HeaderArray so the before-image columns can never drift from the lookup.
Private Sub EnsureHeader()
    On Error Resume Next
    Dim p As String: p = AuditPath()
    If Len(Dir$(p)) > 0 Then Exit Sub

    Dim line As String
    line = "RunId,When,User,Workbook,Sheet,Row,Operation,Serial,DCN,Outcome,BeforeState"

    Dim h As Variant: h = HeaderArray()
    Dim i As Long
    For i = LBound(h) To UBound(h)
        line = line & ",Before_" & CStr(h(i))
    Next i

    AppendLine line
End Sub

' Opened and closed per line on purpose. A run is dominated by HTTP calls, so
' the file cost is noise - and a log written as it goes survives a crash, which
' is exactly the run you would most want a record of.
Private Sub AppendLine(ByVal line As String)
    On Error Resume Next
    Dim f As Integer
    f = FreeFile
    Open AuditPath() For Append As #f
    Print #f, line
    Close #f
End Sub

Private Function CsvField(ByVal s As String) As String
    CsvField = """" & Replace$(s, """", """""") & """"
End Function

' Splits one CSV line, honouring quotes and doubled quotes. Written out rather
' than using Split(",") because a DCN name containing a comma is ordinary data
' here, and Split would silently shift every column after it.
Private Function SplitCsv(ByVal line As String) As Variant
    Dim out() As String, n As Long
    ReDim out(0 To 63)

    Dim i As Long, ch As String, cur As String
    Dim inQ As Boolean

    For i = 1 To Len(line)
        ch = Mid$(line, i, 1)
        If inQ Then
            If ch = """" Then
                If Mid$(line, i + 1, 1) = """" Then
                    cur = cur & """"
                    i = i + 1
                Else
                    inQ = False
                End If
            Else
                cur = cur & ch
            End If
        Else
            Select Case ch
                Case """": inQ = True
                Case ",":  out(n) = cur: cur = "": n = n + 1
                           If n > UBound(out) Then ReDim Preserve out(0 To n + 15)
                Case Else: cur = cur & ch
            End Select
        End If
    Next i

    out(n) = cur
    ReDim Preserve out(0 To n)
    SplitCsv = out
End Function
