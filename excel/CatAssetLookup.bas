Attribute VB_Name = "CatAssetLookup"
'==============================================================================
' Cat Asset Management V2 - Excel lookup
'
' Exposes =CatLookupSerial("<serial>") as a worksheet function that calls the
' Caterpillar Asset Management V2 search endpoint and spills a row per matching
' ownership record. Handles OAuth2 client-credentials auth (Entra ID) with a
' cached token.
'
' DEPENDENCIES:
'   1. VBA-JSON (JsonConverter.bas)  -> import into the same workbook.
'      https://github.com/VBA-tools/VBA-JSON
'   2. Reference "Microsoft Scripting Runtime" (Tools > References) - required
'      by VBA-JSON for Scripting.Dictionary.
'
' SETUP: add a sheet named "Config" with these labels in column A and your
' values in column B:
'     ClientId        <your client id>
'     ClientSecret    <your client secret>
'     Scope           <your client id>/.default
'     TenantId        ceb177bf-013b-49ab-8a9c-4abce32afc1e
'     PartyNumber     ZZIO
'     BaseUrl         https://services.cat.com/catDigital/assetManagement/v2
'     TokenUrl        (optional - overrides TenantId)
'
' USAGE in a cell:   =CatLookupSerial(B1)     where B1 holds a serial number
'==============================================================================
Option Explicit

Private Const CFG_SHEET As String = "Config"

' --- token cache (module-level) ---
Private mToken As String
Private mTokenExpiry As Double          ' Date serial; refresh ~60s early

'==============================================================================
' PUBLIC: worksheet function
'==============================================================================
Public Function CatLookupSerial(ByVal Serial As String) As Variant
    On Error GoTo Fail

    If Len(Trim$(Serial)) = 0 Then
        CatLookupSerial = "Enter a serial number"
        Exit Function
    End If

    Dim respText As String
    respText = CatSearch(CStr(Serial), "")          ' filter by serialNumber

    Dim root As Object
    Set root = JsonConverter.ParseJson(respText)

    If root Is Nothing Then CatLookupSerial = "No data": Exit Function
    If Not root.Exists("ownershipRecords") Then CatLookupSerial = "No data": Exit Function

    Dim recs As Object
    Set recs = root("ownershipRecords")
    Dim n As Long: n = recs.Count
    If n = 0 Then CatLookupSerial = "No records found for " & Serial: Exit Function

    Dim cols As Variant: cols = HeaderArray()

    Dim out() As Variant
    ReDim out(0 To n, 0 To UBound(cols))

    Dim c As Long
    For c = 0 To UBound(cols): out(0, c) = cols(c): Next c   ' header row

    Dim i As Long
    For i = 1 To n
        Dim vals As Variant: vals = RecordValues(recs(i))
        For c = 0 To UBound(cols): out(i, c) = vals(c): Next c
    Next i

    CatLookupSerial = out
    Exit Function
Fail:
    CatLookupSerial = "Error: " & Err.Description
End Function

' Clear the cached token (handy for demos / after changing credentials).
Public Sub CatClearTokenCache()
    mToken = ""
    mTokenExpiry = 0
    MsgBox "Cat token cache cleared.", vbInformation
End Sub

'==============================================================================
' PUBLIC: batch macro - prompts for a column of serials, writes a results sheet
' (one row per ownership record). Run it from the Macros dialog or a button.
'==============================================================================
Public Sub CatBatchLookup()
    On Error GoTo Fail

    ' 1) Ask the user to select the column of serial numbers
    Dim rng As Range
    On Error Resume Next
    Set rng = Application.InputBox( _
        Prompt:="Select the column of serial numbers (a header row is OK):", _
        Title:="Cat Asset Lookup", Type:=8)
    On Error GoTo Fail
    If rng Is Nothing Then Exit Sub                 ' user cancelled

    ' Limit to the used area so selecting a whole column isn't a million cells
    Set rng = Application.Intersect(rng.Columns(1), rng.Worksheet.UsedRange)
    If rng Is Nothing Then MsgBox "That selection has no data.", vbExclamation: Exit Sub

    ' 2) Collect serials (skip blanks and an obvious header cell)
    Dim serials As Collection: Set serials = New Collection
    Dim cell As Range, s As String
    For Each cell In rng.Cells
        s = Trim$(CStr(cell.Value))
        If Len(s) > 0 Then
            If Not (serials.Count = 0 And LCase$(s) Like "*serial*") Then serials.Add s
        End If
    Next cell
    If serials.Count = 0 Then MsgBox "No serial numbers found in the selection.", vbExclamation: Exit Sub

    ' 3) Loop the serials, accumulate result rows
    Dim headers As Variant: headers = HeaderArray()
    Dim nFields As Long: nFields = UBound(headers) + 1
    Dim totalCols As Long: totalCols = 1 + nFields + 1     ' QuerySerial + fields + Note

    Dim rowsOut As Collection: Set rowsOut = New Collection
    Application.ScreenUpdating = False

    Dim idx As Long
    For idx = 1 To serials.Count
        s = serials(idx)
        Application.StatusBar = "Looking up " & idx & " of " & serials.Count & "  (" & s & ")"

        Dim respText As String, errMsg As String
        errMsg = ""
        On Error Resume Next
        respText = CatSearch(s, "")
        If Err.Number <> 0 Then errMsg = Err.Description: Err.Clear
        On Error GoTo Fail

        If Len(errMsg) > 0 Then
            rowsOut.Add MakeRow(s, Nothing, "ERROR: " & errMsg, totalCols, nFields)
        Else
            Dim root As Object, recs As Object
            Set root = JsonConverter.ParseJson(respText)
            Set recs = Nothing
            If Not root Is Nothing Then
                If root.Exists("ownershipRecords") Then Set recs = root("ownershipRecords")
            End If
            If recs Is Nothing Then
                rowsOut.Add MakeRow(s, Nothing, "no data", totalCols, nFields)
            ElseIf recs.Count = 0 Then
                rowsOut.Add MakeRow(s, Nothing, "no records found", totalCols, nFields)
            Else
                Dim k As Long
                For k = 1 To recs.Count
                    rowsOut.Add MakeRow(s, recs(k), "", totalCols, nFields)
                Next k
            End If
        End If
    Next idx

    ' 4) Write everything to a fresh sheet in one shot
    Dim ws As Worksheet: Set ws = MakeOutputSheet()
    Dim grid() As Variant
    ReDim grid(1 To rowsOut.Count + 1, 1 To totalCols)

    grid(1, 1) = "QuerySerial"
    Dim c As Long
    For c = 0 To nFields - 1: grid(1, 2 + c) = headers(c): Next c
    grid(1, totalCols) = "Note"

    Dim rr As Long, rowArr As Variant
    For rr = 1 To rowsOut.Count
        rowArr = rowsOut(rr)
        For c = 0 To totalCols - 1: grid(rr + 1, c + 1) = rowArr(c): Next c
    Next rr

    ws.Range(ws.Cells(1, 1), ws.Cells(rowsOut.Count + 1, totalCols)).Value = grid
    ws.Rows(1).Font.Bold = True
    ws.Range("A1").AutoFilter
    ws.Columns.AutoFit
    ws.Activate

    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox serials.Count & " serial(s) looked up -> " & rowsOut.Count & _
           " record row(s) on '" & ws.Name & "'.", vbInformation
    Exit Sub
Fail:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

'==============================================================================
' HTTP: search + token
'==============================================================================
Private Function CatSearch(ByVal serialNumber As String, ByVal dcn As String) As String
    Dim baseUrl As String, party As String, url As String
    baseUrl = Cfg("BaseUrl")
    If Len(baseUrl) = 0 Then baseUrl = "https://services.cat.com/catDigital/assetManagement/v2"
    party = Cfg("PartyNumber")
    url = baseUrl & "/ownershipRecords/search?partyNumber=" & UrlEncode(party)

    Dim filters As String
    If Len(dcn) > 0 Then filters = filters & FilterJson("dcn", dcn) & ","
    If Len(serialNumber) > 0 Then filters = filters & FilterJson("serialNumber", serialNumber) & ","
    If Len(filters) > 0 Then filters = Left$(filters, Len(filters) - 1)

    Dim payload As String
    payload = "{""filters"":[" & filters & "]}"

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", url, False
    http.SetRequestHeader "Authorization", "Bearer " & GetToken()
    http.SetRequestHeader "Content-Type", "application/json"
    http.SetRequestHeader "Accept", "application/json"
    http.Send payload

    If http.Status <> 200 Then
        Err.Raise vbObjectError + 2, , "API " & http.Status & ": " & Left$(http.responseText, 300)
    End If
    CatSearch = http.responseText
End Function

Private Function GetToken() As String
    If Len(mToken) > 0 And Now < mTokenExpiry Then
        GetToken = mToken
        Exit Function
    End If

    Dim tokenUrl As String
    tokenUrl = Cfg("TokenUrl")
    If Len(tokenUrl) = 0 Then
        tokenUrl = "https://login.microsoftonline.com/" & Cfg("TenantId") & "/oauth2/v2.0/token"
    End If

    Dim body As String
    body = "grant_type=client_credentials" & _
           "&client_id=" & UrlEncode(Cfg("ClientId")) & _
           "&client_secret=" & UrlEncode(Cfg("ClientSecret")) & _
           "&scope=" & UrlEncode(Cfg("Scope"))

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", tokenUrl, False
    http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.Send body

    If http.Status <> 200 Then
        Err.Raise vbObjectError + 1, , "Token " & http.Status & ": " & Left$(http.responseText, 300)
    End If

    Dim j As Object
    Set j = JsonConverter.ParseJson(http.responseText)
    mToken = CStr(j("access_token"))

    Dim expiresIn As Double: expiresIn = 3600
    If j.Exists("expires_in") Then expiresIn = CDbl(j("expires_in"))
    mTokenExpiry = Now + (expiresIn - 60) / 86400#

    GetToken = mToken
End Function

Private Function FilterJson(ByVal prop As String, ByVal val As String) As String
    FilterJson = "{""type"":""stringEquals"",""propertyName"":""" & prop & _
                 """,""values"":[""" & JsonEsc(val) & """]}"
End Function

'==============================================================================
' Config sheet access
'==============================================================================
Private Function Cfg(ByVal label As String) As String
    Dim ws As Worksheet, found As Range
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CFG_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Err.Raise vbObjectError + 9, , "Missing '" & CFG_SHEET & "' sheet"
    Set found = ws.Columns(1).Find(What:=label, LookAt:=xlWhole, MatchCase:=False)
    If Not found Is Nothing Then Cfg = Trim$(CStr(ws.Cells(found.Row, 2).Value))
End Function

'==============================================================================
' Record shaping (shared by the function and the batch macro)
'==============================================================================
Private Function HeaderArray() As Variant
    HeaderArray = Array("SerialNumber", "MakeCode", "MakeName", "Model", "ModelYear", _
                        "AssetName", "DealerCode", "DealerName", "DCN", "DcnName", _
                        "CCID", "CcidName", "OwnershipType", "Status")
End Function

' Returns a 1-D array of the 14 field values for one ownership record.
Private Function RecordValues(ByVal rec As Object) As Variant
    Dim md As Object, ow As Object, da As Object, mk As Object, ot As Object, rs As Object
    Set md = SafeObj(rec, "metadata")
    Set ow = SafeObj(rec, "ownership")
    Set da = SafeObj(ow, "dealerAssociation")
    Set mk = SafeObj(md, "makeInfo")
    Set ot = SafeObj(da, "dcnOwnershipType")
    Set rs = SafeObj(da, "dcnRelationStatus")

    Dim v(0 To 13) As Variant
    v(0) = SafeStr(md, "serialNumber")
    v(1) = SafeStr(mk, "code")
    v(2) = SafeStr(mk, "name")
    v(3) = SafeStr(md, "model")
    v(4) = SafeStr(md, "modelYear")
    v(5) = SafeStr(md, "assetName")
    v(6) = SafeStr(da, "dealerCode")
    v(7) = SafeStr(da, "dealerName")
    v(8) = SafeStr(da, "dcn")
    v(9) = SafeStr(da, "dcnName")
    v(10) = SafeStr(ow, "ccid")
    v(11) = SafeStr(ow, "ccidName")
    v(12) = SafeStr(ot, "code")
    v(13) = SafeStr(rs, "code")
    RecordValues = v
End Function

' Builds one output row: QuerySerial + 14 fields (if rec given) + Note.
Private Function MakeRow(ByVal serial As String, ByVal rec As Object, _
                         ByVal note As String, ByVal totalCols As Long, _
                         ByVal nFields As Long) As Variant
    Dim r() As Variant
    ReDim r(0 To totalCols - 1)
    r(0) = serial
    If Not rec Is Nothing Then
        Dim vals As Variant: vals = RecordValues(rec)
        Dim c As Long
        For c = 0 To nFields - 1: r(1 + c) = vals(c): Next c
    End If
    r(totalCols - 1) = note
    MakeRow = r
End Function

'==============================================================================
' Output sheet
'==============================================================================
Private Function MakeOutputSheet() As Worksheet
    Dim base As String: base = "Asset Lookup Results"
    Dim nm As String: nm = base
    Dim i As Long: i = 1
    Do While SheetExists(nm)
        i = i + 1
        nm = base & " " & i
    Loop
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = nm
    Set MakeOutputSheet = ws
End Function

Private Function SheetExists(ByVal nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

'==============================================================================
' Small helpers
'==============================================================================
Private Function SafeObj(ByVal parent As Object, ByVal key As String) As Object
    If parent Is Nothing Then Exit Function
    On Error Resume Next
    If parent.Exists(key) Then
        If IsObject(parent(key)) Then Set SafeObj = parent(key)
    End If
End Function

Private Function SafeStr(ByVal parent As Object, ByVal key As String) As String
    If parent Is Nothing Then Exit Function
    On Error Resume Next
    If parent.Exists(key) Then
        If Not IsObject(parent(key)) Then SafeStr = CStr(parent(key))
    End If
End Function

Private Function UrlEncode(ByVal s As String) As String
    Dim i As Long, ch As String, code As Long, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) _
           Or (code >= 97 And code <= 122) Or InStr("-_.~", ch) > 0 Then
            out = out & ch
        Else
            out = out & "%" & Right$("0" & Hex$(code And &HFF), 2)
        End If
    Next i
    UrlEncode = out
End Function

Private Function JsonEsc(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    JsonEsc = s
End Function
