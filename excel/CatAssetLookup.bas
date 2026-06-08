Attribute VB_Name = "CatAssetLookup"
'==============================================================================
' Cat Asset Management V2 - core + worksheet function
'
' Exposes =CatLookupSerial("<serial>") and the shared building blocks
' (CatSearch, HeaderArray, RecordValues) used by the CatBatchLookup module.
' Handles OAuth2 client-credentials auth (Entra ID) with a cached token.
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
' For a whole column of serials, run the CatBatchLookup macro (other module).
'==============================================================================
Option Explicit

Private Const CFG_SHEET As String = "Config"

' Read-only switch. While True, the write endpoints (CatAddUpdate / CatExpire /
' CatTransfer) are disabled and raise an error if called. Set to False to
' re-enable writes, then re-run CatSetupActionsSheet to show the write buttons.
Public Const CAT_READ_ONLY As Boolean = True

' --- token cache (module-level; shared across the function and the batch macro
'     because the batch macro calls CatSearch -> GetToken in this module) ---
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
' PUBLIC: shared building blocks (used here and by the CatBatchLookup module)
'==============================================================================

' POST /ownershipRecords/search - returns raw JSON text (raises on non-200).
Public Function CatSearch(ByVal serialNumber As String, ByVal dcn As String) As String
    Dim filters As String
    If Len(dcn) > 0 Then filters = filters & FilterJson("dcn", dcn) & ","
    If Len(serialNumber) > 0 Then filters = filters & FilterJson("serialNumber", serialNumber) & ","
    If Len(filters) > 0 Then filters = Left$(filters, Len(filters) - 1)

    Dim q As String
    q = "/ownershipRecords/search?partyNumber=" & UrlEncode(Cfg("PartyNumber"))

    Dim status As Long, txt As String
    txt = CatPost(q, "{""filters"":[" & filters & "]}", status)
    If status <> 200 Then Err.Raise vbObjectError + 2, , "API " & status & ": " & Left$(txt, 300)
    CatSearch = txt
End Function

' Low-level authenticated POST. Returns response text and sets statusOut to the
' HTTP status. Does NOT raise on non-2xx - the caller decides what to do.
Public Function CatPost(ByVal pathWithQuery As String, ByVal jsonBody As String, _
                        ByRef statusOut As Long) As String
    Dim baseUrl As String
    baseUrl = Cfg("BaseUrl")
    If Len(baseUrl) = 0 Then baseUrl = "https://services.cat.com/catDigital/assetManagement/v2"

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", baseUrl & pathWithQuery, False
    http.SetTimeouts 5000, 10000, 10000, 30000   ' resolve, connect, send, receive (ms)
    http.SetRequestHeader "Authorization", "Bearer " & GetToken()
    http.SetRequestHeader "Accept", "application/json"
    If Len(jsonBody) > 0 Then http.SetRequestHeader "Content-Type", "application/json"
    http.Send jsonBody
    statusOut = http.Status
    CatPost = http.responseText
End Function

' POST /ownershipRecords - add or update an ownership record. Returns the
' response body (AddOwnershipResponse JSON) and sets statusOut.
Public Function CatAddUpdate(ByVal serialNumber As String, ByVal dcn As String, _
        ByVal makeCode As String, ByVal dealerMakeCode As String, _
        ByVal ownershipTypeCode As String, ByVal model As String, _
        ByVal modelYear As String, ByVal productFamilyCode As String, _
        ByVal productFamilyName As String, ByVal baseAssetName As String, _
        ByVal customAssetName As String, ByRef statusOut As Long) As String
    If CAT_READ_ONLY Then Err.Raise vbObjectError + 100, , "Writes are disabled (read-only mode)."
    Dim q As String
    q = "/ownershipRecords?partyNumber=" & UrlEncode(Cfg("PartyNumber")) & _
        AssetQuery(serialNumber, dcn, makeCode, dealerMakeCode)

    Dim b As String
    b = b & JF("ownershipTypeCode", ownershipTypeCode)
    b = b & JF("model", model)
    b = b & JF("modelYear", modelYear)
    b = b & JF("productFamilyCode", productFamilyCode)
    b = b & JF("productFamilyName", productFamilyName)
    b = b & JF("baseAssetName", baseAssetName)
    b = b & JF("customAssetName", customAssetName)
    If Len(b) > 0 Then b = Left$(b, Len(b) - 1)        ' strip trailing comma

    CatAddUpdate = CatPost(q, "{" & b & "}", statusOut)
End Function

' POST /ownershipRecords/expire - expire an ownership record (success = 204).
Public Function CatExpire(ByVal serialNumber As String, ByVal dcn As String, _
        ByVal makeCode As String, ByVal dealerMakeCode As String, _
        ByRef statusOut As Long) As String
    If CAT_READ_ONLY Then Err.Raise vbObjectError + 100, , "Writes are disabled (read-only mode)."
    Dim q As String
    q = "/ownershipRecords/expire?partyNumber=" & UrlEncode(Cfg("PartyNumber")) & _
        AssetQuery(serialNumber, dcn, makeCode, dealerMakeCode)
    CatExpire = CatPost(q, "", statusOut)
End Function

' POST /ownershipRequests/transfer - approve/reject a pending transfer (204).
' No DCN is used for this endpoint.
Public Function CatTransfer(ByVal serialNumber As String, ByVal makeCode As String, _
        ByVal dealerMakeCode As String, ByVal statusValue As String, _
        ByVal reason As String, ByRef statusOut As Long) As String
    If CAT_READ_ONLY Then Err.Raise vbObjectError + 100, , "Writes are disabled (read-only mode)."
    Dim q As String
    q = "/ownershipRequests/transfer?partyNumber=" & UrlEncode(Cfg("PartyNumber")) & _
        AssetQuery(serialNumber, "", makeCode, dealerMakeCode)
    Dim b As String
    b = JF("status", UCase$(statusValue))
    b = b & JF("reason", reason)
    If Len(b) > 0 Then b = Left$(b, Len(b) - 1)
    CatTransfer = CatPost(q, "{" & b & "}", statusOut)
End Function

' Build the shared asset-identifier query (&serialNumber=..&makeCode=..&dcn=..).
Private Function AssetQuery(ByVal serialNumber As String, ByVal dcn As String, _
        ByVal makeCode As String, ByVal dealerMakeCode As String) As String
    Dim s As String
    If Len(serialNumber) > 0 Then s = s & "&serialNumber=" & UrlEncode(serialNumber)
    If Len(makeCode) > 0 Then s = s & "&makeCode=" & UrlEncode(makeCode)
    If Len(dealerMakeCode) > 0 Then s = s & "&dealerMakeCode=" & UrlEncode(dealerMakeCode)
    If Len(dcn) > 0 Then s = s & "&dcn=" & UrlEncode(dcn)
    AssetQuery = s
End Function

' One JSON "name":"value", pair (trailing comma) - empty if val is blank.
Private Function JF(ByVal name As String, ByVal val As String) As String
    If Len(val) = 0 Then Exit Function
    JF = """" & name & """:""" & JsonEsc(val) & ""","
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
' PRIVATE: token + request internals
'==============================================================================
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
    http.SetTimeouts 5000, 10000, 10000, 30000   ' resolve, connect, send, receive (ms)
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
