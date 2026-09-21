<%
' ============================================================================
'  mailgateway.inc.asp  -  Classic ASP / VBScript client for MailGateway
'
'  Requires in global.asa (or set before including):
'     Application("MG_URL")     = "https://apps.central-office.info/mailgateway/api/mail/send"
'     Application("MG_API_KEY") = "....."
'
'  Usage (no attachments, identical to the old call):
'     ok = MG_Send("member@x.org", "Subject", "<p>html</p>", "", "", "", "")
'
'  Usage (with attachments from disk on THIS web server - any server, not just the gateway's):
'     atts = MG_AttachFile(Server.MapPath("/receipts/12345.pdf"), "Receipt-12345.pdf", "")
'     atts = atts & "," & MG_AttachFile(Server.MapPath("/img/logo.png"), "logo.png", "logo")   ' inline: <img src="cid:logo">
'     ok = MG_Send("member@x.org", "Your receipt", html, "", "noreply@pasmeeting.org", "PAS Meeting", atts)
'
'  Usage (attachment already in memory as a byte array, e.g. generated PDF):
'     atts = MG_AttachBytes(pdfBytes, "invoice.pdf", "")
'
'  After the call:  MG_LastStatus (HTTP code), MG_LastResponse (JSON text), MG_LastRequestId, MG_LastError
' ============================================================================

Dim MG_LastStatus, MG_LastResponse, MG_LastRequestId, MG_LastError

' ---- JSON string escaping ------------------------------------------------
Function MG_JsonEscape(s)
    Dim r, i, c, code
    If IsNull(s) Then s = ""
    r = ""
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        code = AscW(c)
        Select Case c
            Case """": r = r & "\"""
            Case "\":  r = r & "\\"
            Case vbCr: r = r & "\r"
            Case vbLf: r = r & "\n"
            Case vbTab: r = r & "\t"
            Case Else
                If code < 32 Then
                    r = r & "\u" & Right("0000" & Hex(code), 4)
                Else
                    r = r & c
                End If
        End Select
    Next
    MG_JsonEscape = r
End Function

' ---- base64 helpers ------------------------------------------------------
Function MG_BytesToBase64(bytes)
    Dim xml, node
    Set xml = Server.CreateObject("MSXML2.DOMDocument.6.0")
    Set node = xml.createElement("b64")
    node.dataType = "bin.base64"
    node.nodeTypedValue = bytes
    ' MSXML wraps base64 at 76 chars; raw line breaks are illegal inside a JSON string, so strip them here
    MG_BytesToBase64 = Replace(Replace(node.text, vbCr, ""), vbLf, "")
    Set node = Nothing
    Set xml = Nothing
End Function

Function MG_FileToBase64(path)
    Dim stm, bytes
    Set stm = Server.CreateObject("ADODB.Stream")
    stm.Type = 1                            ' adTypeBinary
    stm.Open
    stm.LoadFromFile path
    bytes = stm.Read
    stm.Close
    Set stm = Nothing
    MG_FileToBase64 = MG_BytesToBase64(bytes)
End Function

' ---- attachment JSON fragments (comma-join several) -----------------------
' contentId: "" for a normal attachment, or the cid to reference from htmlBody for an inline image.
Function MG_AttachBase64(base64, fileName, contentId)
    Dim j
    j = "{""name"":""" & MG_JsonEscape(fileName) & """,""contentBase64"":""" & base64 & """"
    If Len(contentId) > 0 Then
        j = j & ",""contentId"":""" & MG_JsonEscape(contentId) & """,""isInline"":true"
    End If
    MG_AttachBase64 = j & "}"
End Function

Function MG_AttachFile(path, fileName, contentId)
    If Len(fileName) = 0 Then fileName = Mid(path, InStrRev(path, "\") + 1)
    MG_AttachFile = MG_AttachBase64(MG_FileToBase64(path), fileName, contentId)
End Function

Function MG_AttachBytes(bytes, fileName, contentId)
    MG_AttachBytes = MG_AttachBase64(MG_BytesToBase64(bytes), fileName, contentId)
End Function

' ---- the call --------------------------------------------------------------
' toList: "a@x.org; b@y.org" (semicolon or comma separated)
' fromAddr / fromName: "" to use the gateway's default sender
' attachmentsJson: "" or comma-joined fragments from MG_Attach*()
Function MG_Send(toList, subject, htmlBody, textBody, fromAddr, fromName, attachmentsJson)
    MG_Send = MG_SendEx(toList, "", "", "", subject, htmlBody, textBody, fromAddr, fromName, attachmentsJson, "")
End Function

Function MG_SendEx(toList, ccList, bccList, replyTo, subject, htmlBody, textBody, fromAddr, fromName, attachmentsJson, appName)
    Dim http, body, url, key
    url = Application("MG_URL")
    key = Application("MG_API_KEY")

    body = "{"
    body = body & """appName"":""" & MG_JsonEscape(appName) & ""","
    body = body & """to"":""" & MG_JsonEscape(toList) & ""","
    If Len(ccList) > 0 Then body = body & """cc"":""" & MG_JsonEscape(ccList) & ""","
    If Len(bccList) > 0 Then body = body & """bcc"":""" & MG_JsonEscape(bccList) & ""","
    If Len(replyTo) > 0 Then body = body & """replyTo"":""" & MG_JsonEscape(replyTo) & ""","
    If Len(fromAddr) > 0 Then body = body & """from"":""" & MG_JsonEscape(fromAddr) & ""","
    If Len(fromName) > 0 Then body = body & """fromName"":""" & MG_JsonEscape(fromName) & ""","
    body = body & """subject"":""" & MG_JsonEscape(subject) & ""","
    body = body & """htmlBody"":""" & MG_JsonEscape(htmlBody) & ""","
    body = body & """textBody"":""" & MG_JsonEscape(textBody) & """"
    If Len(attachmentsJson) > 0 Then body = body & ",""attachments"":[" & attachmentsJson & "]"
    body = body & "}"

    MG_LastStatus = 0: MG_LastResponse = "": MG_LastRequestId = "": MG_LastError = ""

    On Error Resume Next
    Set http = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 15000, 15000, 120000, 300000      ' resolve, connect, send, receive (large attachments take time)
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.setRequestHeader "X-API-Key", key
    http.Send body
    If Err.Number <> 0 Then
        MG_LastError = "HTTP error: " & Err.Description
        Err.Clear
        On Error GoTo 0
        MG_SendEx = False
        Exit Function
    End If
    On Error GoTo 0

    MG_LastStatus = http.status
    MG_LastResponse = http.responseText
    Set http = Nothing

    MG_LastRequestId = MG_JsonField(MG_LastResponse, "requestId")
    If MG_LastStatus >= 200 And MG_LastStatus < 300 Then
        MG_SendEx = True
    Else
        MG_LastError = MG_JsonField(MG_LastResponse, "error")
        If Len(MG_LastError) = 0 Then MG_LastError = "HTTP " & MG_LastStatus
        MG_SendEx = False
    End If
End Function

' Tiny extractor for flat string fields in the gateway's JSON response (no full parser needed).
Function MG_JsonField(json, name)
    Dim p, q
    MG_JsonField = ""
    p = InStr(json, """" & name & """")
    If p = 0 Then Exit Function
    p = InStr(p, json, ":")
    If p = 0 Then Exit Function
    p = InStr(p, json, """")
    If p = 0 Then Exit Function
    q = InStr(p + 1, json, """")
    If q = 0 Then Exit Function
    MG_JsonField = Mid(json, p + 1, q - p - 1)
End Function
%>
