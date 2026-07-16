Attribute VB_Name = "CatAssetLookup"
'==============================================================================
' Cat Asset Management V2 - core + worksheet function (via PROXY)
'
' Talks to the Asset Management PROXY (an Azure Function) instead of calling the
' Cat API directly. The proxy holds the Cat OAuth credentials server-side, so
' THIS WORKBOOK NEVER HOLDS A CLIENT SECRET - only the proxy URL and a function
' key. Share the .xlsm freely with anyone allowed to have the function key.
'
' Exposes =CatLookupSerial("<serial>") and =CatLookupDCN("<dcn>") plus the
' shared building blocks (CatSearch, CatProxyCall, CatAddUpdate/CatExpire/
' CatTransfer, HeaderArray, RecordValues, OwnershipRecords) used by the
' CatBatchLookup and CatActions modules.
'
' DEPENDENCIES:
'   1. VBA-JSON (JsonConverter.bas)  -> import into the same workbook.
'      https://github.com/VBA-tools/VBA-JSON
'   2. Reference "Microsoft Scripting Runtime" (Tools > References) - required
'      by VBA-JSON for Scripting.Dictionary.
'
' SETUP: add a sheet named "Config" with these labels in column A and your
' values in column B:
'     ProxyUrl      https://<app>.azurewebsites.net/api
'     FunctionKey   <the function key from the proxy>
'     PartyNumber   (optional - usually blank; the proxy supplies the dealer code)
'
' USAGE in a cell:   =CatLookupSerial(B1)     where B1 holds a serial number
'                    =CatLookupDCN(B2)        where B2 holds a DCN
' For a whole column of serials, run the CatBatchLookup macro (other module).
'==============================================================================
Option Explicit

Private Const CFG_SHEET As String = "Config"

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
    respText = CatSearch(CStr(Serial), "")          ' filter by serialNumber (raises on non-200)

    Dim recs As Object
    Set recs = OwnershipRecords(respText)
    If recs Is Nothing Then CatLookupSerial = "No data": Exit Function

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

' Same spilled table, filtered by DCN instead of serial. A DCN typically has
' many assets, so expect many rows (one per ownership record).
Public Function CatLookupDCN(ByVal Dcn As String) As Variant
    On Error GoTo Fail

    If Len(Trim$(Dcn)) = 0 Then
        CatLookupDCN = "Enter a DCN"
        Exit Function
    End If

    Dim respText As String
    respText = CatSearch("", CStr(Dcn))             ' filter by dcn (raises on non-200)

    Dim recs As Object
    Set recs = OwnershipRecords(respText)
    If recs Is Nothing Then CatLookupDCN = "No data": Exit Function

    Dim n As Long: n = recs.Count
    If n = 0 Then CatLookupDCN = "No records found for DCN " & Dcn: Exit Function

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

    CatLookupDCN = out
    Exit Function
Fail:
    CatLookupDCN = "Error: " & Err.Description
End Function

'==============================================================================
' PUBLIC: shared building blocks (used here and by CatBatchLookup / CatActions)
'==============================================================================

' GET /search - returns the proxy envelope JSON text. Raises on non-200 so the
' batch macro and worksheet function can classify success vs. failure.
Public Function CatSearch(ByVal serialNumber As String, ByVal dcn As String) As String
    serialNumber = CleanId(serialNumber)
    dcn = CleanId(dcn)
    Dim q As String
    If Len(serialNumber) > 0 Then q = q & "serial=" & UrlEncode(serialNumber) & "&"
    If Len(dcn) > 0 Then q = q & "dcn=" & UrlEncode(dcn) & "&"
    If Len(q) > 0 Then q = Left$(q, Len(q) - 1)

    Dim status As Long, txt As String
    txt = CatProxyCall("GET", "search", q, "", status)
    If status <> 200 Then Err.Raise vbObjectError + 2, , "API " & status & ": " & Left$(ProxyError(txt, status), 300)
    CatSearch = txt
End Function

' POST /ownership - add or update an ownership record. Returns the proxy
' envelope text and sets statusOut to the HTTP status.
Public Function CatAddUpdate(ByVal serialNumber As String, ByVal dcn As String, _
        ByVal makeCode As String, ByVal dealerMakeCode As String, _
        ByVal ownershipTypeCode As String, ByVal model As String, _
        ByVal modelYear As String, ByVal productFamilyCode As String, _
        ByVal productFamilyName As String, ByVal baseAssetName As String, _
        ByVal customAssetName As String, ByRef statusOut As Long) As String
    serialNumber = CleanId(serialNumber)
    dcn = CleanId(dcn)
    Dim b As String
    b = JF("serial_number", serialNumber) & JF("dcn", dcn) & _
        JF("make_code", makeCode) & JF("dealer_make_code", dealerMakeCode) & _
        JF("ownership_type_code", ownershipTypeCode) & JF("model", model) & _
        JF("model_year", modelYear) & JF("product_family_code", productFamilyCode) & _
        JF("product_family_name", productFamilyName) & JF("base_asset_name", baseAssetName) & _
        JF("custom_asset_name", customAssetName)
    CatAddUpdate = CatProxyCall("POST", "ownership", "", "{" & TrimComma(b) & "}", statusOut)
End Function

' POST /expire - expire an ownership record (success = 204).
Public Function CatExpire(ByVal serialNumber As String, ByVal dcn As String, _
        ByVal makeCode As String, ByVal dealerMakeCode As String, _
        ByRef statusOut As Long) As String
    serialNumber = CleanId(serialNumber)
    dcn = CleanId(dcn)
    Dim b As String
    b = JF("serial_number", serialNumber) & JF("dcn", dcn) & _
        JF("make_code", makeCode) & JF("dealer_make_code", dealerMakeCode)
    CatExpire = CatProxyCall("POST", "expire", "", "{" & TrimComma(b) & "}", statusOut)
End Function

' POST /transfer - approve/reject a pending transfer (success = 204).
' No DCN is used for this endpoint.
Public Function CatTransfer(ByVal serialNumber As String, ByVal makeCode As String, _
        ByVal dealerMakeCode As String, ByVal statusValue As String, _
        ByVal reason As String, ByRef statusOut As Long) As String
    serialNumber = CleanId(serialNumber)
    Dim b As String
    b = JF("serial_number", serialNumber) & JF("make_code", makeCode) & _
        JF("dealer_make_code", dealerMakeCode) & JF("status", UCase$(statusValue)) & _
        JF("reason", reason)
    CatTransfer = CatProxyCall("POST", "transfer", "", "{" & TrimComma(b) & "}", statusOut)
End Function

' Low-level proxy call. method = "GET"/"POST"; route = "search"/"ownership"/
' "expire"/"transfer"; query is appended to the URL; jsonBody is sent for POST.
' Returns the response text and sets statusOut to the HTTP status. Does NOT raise
' on non-2xx - the caller decides. Sends the function key as the x-functions-key
' header (kept out of the URL).
Public Function CatProxyCall(ByVal method As String, ByVal route As String, _
        ByVal query As String, ByVal jsonBody As String, ByRef statusOut As Long) As String
    Dim baseUrl As String: baseUrl = Cfg("ProxyUrl")
    If Len(baseUrl) = 0 Then Err.Raise vbObjectError + 10, , "Config!ProxyUrl is empty."
    If Right$(baseUrl, 1) = "/" Then baseUrl = Left$(baseUrl, Len(baseUrl) - 1)

    Dim url As String: url = baseUrl & "/" & route
    If Len(query) > 0 Then url = url & "?" & query

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open method, url, False
    http.SetTimeouts 5000, 10000, 10000, 30000   ' resolve, connect, send, receive (ms)

    Dim key As String: key = Cfg("FunctionKey")
    If Len(key) > 0 Then http.SetRequestHeader "x-functions-key", key
    http.SetRequestHeader "Accept", "application/json"
    If UCase$(method) = "POST" Then http.SetRequestHeader "Content-Type", "application/json"

    http.Send jsonBody
    statusOut = http.status
    CatProxyCall = http.responseText
End Function

'==============================================================================
' PUBLIC: envelope + record parsing
'==============================================================================

' Pull ownershipRecords out of the proxy envelope: { ok, data:{ ownershipRecords:[...] } }
Public Function OwnershipRecords(ByVal envelopeText As String) As Object
    On Error Resume Next
    Dim root As Object: Set root = JsonConverter.ParseJson(envelopeText)
    If root Is Nothing Then Exit Function
    Dim data As Object: Set data = SafeObj(root, "data")
    If data Is Nothing Then Exit Function
    If data.Exists("ownershipRecords") Then Set OwnershipRecords = data("ownershipRecords")
End Function

' Read a single field from the proxy envelope's data object (e.g. "status").
Public Function FieldOf(ByVal txt As String, ByVal key As String) As String
    On Error Resume Next
    Dim root As Object: Set root = JsonConverter.ParseJson(txt)
    If root Is Nothing Then Exit Function
    Dim data As Object: Set data = SafeObj(root, "data")
    If Not data Is Nothing Then If data.Exists(key) Then FieldOf = CStr(data(key))
End Function

' Turn the proxy error envelope ({ ok:false, error:{ code, description } }) into a
' short message; falls back to the raw text.
Public Function ProxyError(ByVal txt As String, ByVal status As Long) As String
    On Error Resume Next
    Dim root As Object: Set root = JsonConverter.ParseJson(txt)
    If Not root Is Nothing Then
        Dim er As Object: Set er = SafeObj(root, "error")
        If Not er Is Nothing Then
            If er.Exists("code") Then ProxyError = CStr(er("code")) & " "
            If er.Exists("description") Then ProxyError = ProxyError & CStr(er("description"))
        End If
    End If
    If Len(Trim$(ProxyError)) = 0 Then ProxyError = Left$(txt, 200)
End Function

' Normalize a pasted identifier (serial / DCN): convert non-breaking spaces,
' tabs and line breaks to regular spaces, strip zero-width characters, then trim
' the ends. Internal regular spaces and hyphens are preserved (both are valid in
' CCAT serials) so the value still matches exactly - this only removes the
' invisible junk that pasted data tends to carry.
Public Function CleanId(ByVal s As String) As String
    If Len(s) = 0 Then Exit Function
    s = Replace(s, ChrW$(160), " ")     ' non-breaking space
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, ChrW$(8203), "")     ' zero-width space
    s = Replace(s, ChrW$(8204), "")     ' zero-width non-joiner
    s = Replace(s, ChrW$(8205), "")     ' zero-width joiner
    s = Replace(s, ChrW$(65279), "")    ' zero-width no-break space / BOM
    CleanId = Trim$(s)
End Function

' Column headers for the 14 record fields.
Public Function HeaderArray() As Variant
    HeaderArray = Array("SerialNumber", "MakeCode", "MakeName", "Model", "ModelYear", _
                        "AssetName", "DealerCode", "DealerName", "DCN", "DcnName", _
                        "CCID", "CcidName", "OwnershipType", "Status")
End Function

' Returns a 1-D array of the 14 field values for one ownership record.
Public Function RecordValues(ByVal rec As Object) As Variant
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

'==============================================================================
' PRIVATE: JSON building helpers
'==============================================================================

' One JSON "name":"value", pair (trailing comma) - empty if val is blank.
Private Function JF(ByVal name As String, ByVal val As String) As String
    If Len(val) = 0 Then Exit Function
    JF = """" & name & """:""" & JsonEsc(val) & ""","
End Function

Private Function TrimComma(ByVal s As String) As String
    If Len(s) > 0 Then If Right$(s, 1) = "," Then s = Left$(s, Len(s) - 1)
    TrimComma = s
End Function

'==============================================================================
' PRIVATE: config + small helpers
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
