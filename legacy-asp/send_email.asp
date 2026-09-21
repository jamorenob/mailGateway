<%
' ============================================================================
'  send_email.asp  -  legacy ASP pages: send a plain email through the MailGateway
'
'  Drop-in replacement for the old CDO/SMTP include (no attachment version).
'  Reads the same session variables as before:
'     session("subject_line")   subject
'     session("from_line")      "Display Name <addr@domain>"  (or just the address)
'     session("bcc_line")       bcc address(es), ; or , separated   (optional)
'     session("email")          To address(es), ; or , separated
'     session("message")        HTML body
'  Sets, exactly like before:
'     session("smtperror")      "Email sent successfully!" or "Error: ..."
'  and writes the same text to the response.
'
'  Never attaches anything (use send_email_ad.asp for that).
'  Deliberately contains NO Const / Function / Sub and declares its variables via
'  ExecuteGlobal, so a page may
'  #include it several times without "Name redefined".
' ============================================================================

' ---- declarations that survive being included twice (Option Explicit-safe) ----
' ExecuteGlobal turns a repeated Dim into a runtime error we can ignore, instead of a compile error.
On Error Resume Next
ExecuteGlobal "Dim se_url, se_key, se_ok, se_err, se_fromLine, se_fromName, se_fromAddr, se_p, se_q, " & _
              "se_fields, se_i, se_j, se_s, se_r, se_c, se_body, se_http, se_status, se_resp, se_rid"
Err.Clear
On Error GoTo 0

' ---- gateway settings (Application() values from global.asa win when present) ----
se_url = Application("MG_URL")
If Len(se_url & "") = 0 Then se_url = "https://apps-thanos.central-office.info/mailgateway/api/mail/send"
se_key = Application("MG_API_KEY")
If Len(se_key & "") = 0 Then se_key = "PUT-THE-CUSTOMER-API-KEY-HERE"

se_ok = True
se_err = ""

' split "Name <addr>" (tolerates a missing closing ">")
se_fromLine = Trim(session("from_line") & "")
se_fromName = ""
se_fromAddr = se_fromLine
se_p = InStr(se_fromLine, "<")
If se_p > 0 Then
    se_fromName = Trim(Left(se_fromLine, se_p - 1))
    se_fromAddr = Trim(Replace(Mid(se_fromLine, se_p + 1), ">", ""))
End If

' JSON-escape each text field
se_fields = Array(session("email") & "", session("bcc_line") & "", se_fromAddr, se_fromName, session("subject_line") & "", session("message") & "")
For se_i = 0 To UBound(se_fields)
    se_s = se_fields(se_i)
    se_r = ""
    For se_j = 1 To Len(se_s)
        se_c = Mid(se_s, se_j, 1)
        Select Case se_c
            Case """":  se_r = se_r & "\"""
            Case "\":   se_r = se_r & "\\"
            Case vbCr:  se_r = se_r & "\r"
            Case vbLf:  se_r = se_r & "\n"
            Case vbTab: se_r = se_r & "\t"
            Case Else
                If AscW(se_c) < 32 Then
                    se_r = se_r & "\u" & Right("0000" & Hex(AscW(se_c)), 4)
                Else
                    se_r = se_r & se_c
                End If
        End Select
    Next
    se_fields(se_i) = se_r
Next

se_body = "{"
se_body = se_body & """appName"":""legacy " & Request.ServerVariables("SERVER_NAME") & ""","
se_body = se_body & """to"":""" & se_fields(0) & ""","
If Len(se_fields(1)) > 0 Then se_body = se_body & """bcc"":""" & se_fields(1) & ""","
If Len(se_fields(2)) > 0 Then se_body = se_body & """from"":""" & se_fields(2) & ""","
If Len(se_fields(3)) > 0 Then se_body = se_body & """fromName"":""" & se_fields(3) & ""","
se_body = se_body & """subject"":""" & se_fields(4) & ""","
se_body = se_body & """htmlBody"":""" & se_fields(5) & """"
se_body = se_body & "}"

On Error Resume Next
Set se_http = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
se_http.setTimeouts 15000, 15000, 60000, 120000
se_http.Open "POST", se_url, False
se_http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
se_http.setRequestHeader "X-API-Key", se_key
se_http.Send se_body
If Err.Number <> 0 Then
    se_err = "Error: " & Err.Description
    Err.Clear
    se_ok = False
Else
    se_status = se_http.status
    se_resp = se_http.responseText
    If se_status < 200 Or se_status >= 300 Then
        se_err = ""
        se_p = InStr(se_resp, """error""")
        If se_p > 0 Then
            se_p = InStr(se_p, se_resp, ":")
            se_p = InStr(se_p, se_resp, """")
            se_q = InStr(se_p + 1, se_resp, """")
            If se_p > 0 And se_q > se_p Then se_err = Mid(se_resp, se_p + 1, se_q - se_p - 1)
        End If
        If Len(se_err) = 0 Then se_err = "HTTP " & se_status
        se_rid = ""
        se_p = InStr(se_resp, """requestId""")
        If se_p > 0 Then
            se_p = InStr(se_p, se_resp, ":")
            se_p = InStr(se_p, se_resp, """")
            se_q = InStr(se_p + 1, se_resp, """")
            If se_p > 0 And se_q > se_p Then se_rid = Mid(se_resp, se_p + 1, se_q - se_p - 1)
        End If
        se_err = "Error: " & se_err & " (request " & se_rid & ")"
        se_ok = False
    End If
End If
On Error GoTo 0
Set se_http = Nothing

' ---- report exactly like the old include did ----
If se_ok Then
    Response.Write "Email sent successfully!"
    session("smtperror") = "Email sent successfully!"
Else
    Response.Write se_err
    session("smtperror") = se_err
End If
%>
