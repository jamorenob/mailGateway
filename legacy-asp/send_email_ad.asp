<%
' ============================================================================
'  send_email_ad.asp  -  legacy ASP pages: send mail through the MailGateway
'
'  Drop-in replacement for the old CDO/SMTP include. Same contract:
'     session("subject_line")   subject
'     session("from_line")      "Display Name <addr@domain>"  (or just the address)
'     session("bcc_line")       bcc address(es), ; or , separated   (optional)
'     session("email")          To address(es)
'     session("message")        HTML body
'     session("pdf_localpath")  absolute path of the uploaded file on this server (preferred), or
'     session("pdf1")           its site-relative path, e.g. ad-recruit\file.pdf
'                               (both empty = send without an attachment)
'  Sets, exactly like before:
'     session("smtperror")      "Email sent successfully!" or "Error: ..."
'  and writes the same text to the response.
'
'  Deliberately contains NO Const / Function / Sub and declares its variables via
'  ExecuteGlobal, so a page may
'  #include it several times (upload-pdf.asp does) without "Name redefined".
' ============================================================================

' ---- declarations that survive being included twice (Option Explicit-safe) ----
' ExecuteGlobal turns a repeated Dim into a runtime error we can ignore, instead of a compile error.
On Error Resume Next
ExecuteGlobal "Dim sea_url, sea_key, sea_ok, sea_err, sea_attJson, sea_relPath, sea_localPath, sea_fileName, sea_fileNameJ, " & _
              "sea_fso, sea_stm, sea_bytes, sea_xml, sea_node, sea_b64, sea_fromLine, sea_fromName, sea_fromAddr, sea_p, sea_q, " & _
              "sea_fields, sea_i, sea_j, sea_s, sea_r, sea_c, sea_body, sea_http, sea_status, sea_resp, sea_rid"
Err.Clear
On Error GoTo 0

' ---- gateway settings (Application() values from global.asa win when present) ----
sea_url = Application("MG_URL")
If Len(sea_url & "") = 0 Then sea_url = "https://apps-thanos.central-office.info/mailgateway/api/mail/send"
sea_key = Application("MG_API_KEY")
If Len(sea_key & "") = 0 Then sea_key = "PUT-THE-CUSTOMER-API-KEY-HERE"

sea_ok = True
sea_err = ""
sea_attJson = ""

' ---- optional attachment: absolute path in session("pdf_localpath"), else site-relative session("pdf1") ----
sea_localPath = ""
If Len(Trim(session("pdf_localpath") & "")) > 0 Then
    sea_localPath = Trim(session("pdf_localpath"))
    sea_relPath = sea_localPath
ElseIf Len(Trim(session("pdf1") & "")) > 0 Then
    sea_relPath = Replace(Trim(session("pdf1")), "\", "/")
    If Left(sea_relPath, 1) <> "/" Then sea_relPath = "/" & sea_relPath
    sea_localPath = Server.MapPath(sea_relPath)
End If

If Len(sea_localPath) > 0 Then
    sea_fileName = Mid(sea_localPath, InStrRev(sea_localPath, "\") + 1)
    sea_fileName = Mid(sea_fileName, InStrRev(sea_fileName, "/") + 1)

    Set sea_fso = Server.CreateObject("Scripting.FileSystemObject")
    If Not sea_fso.FileExists(sea_localPath) Then
        sea_err = "Error: attachment not found on server (" & sea_relPath & ")"
        sea_ok = False
    End If
    Set sea_fso = Nothing

    If sea_ok Then
        On Error Resume Next
        Set sea_stm = Server.CreateObject("ADODB.Stream")
        sea_stm.Type = 1
        sea_stm.Open
        sea_stm.LoadFromFile sea_localPath
        sea_bytes = sea_stm.Read
        sea_stm.Close
        Set sea_stm = Nothing

        Set sea_xml = Server.CreateObject("MSXML2.DOMDocument.6.0")
        Set sea_node = sea_xml.createElement("b64")
        sea_node.dataType = "bin.base64"
        sea_node.nodeTypedValue = sea_bytes
        sea_b64 = sea_node.text
        ' MSXML wraps base64 at 76 chars; raw line breaks are illegal inside a JSON string
        sea_b64 = Replace(Replace(sea_b64, vbCr, ""), vbLf, "")
        Set sea_node = Nothing
        Set sea_xml = Nothing
        If Err.Number <> 0 Then
            sea_err = "Error: could not read attachment - " & Err.Description
            Err.Clear
            sea_ok = False
        End If
        On Error GoTo 0
    End If

    If sea_ok Then
        ' JSON-escape the file name (quotes and backslashes are the only realistic risks)
        sea_fileNameJ = Replace(Replace(sea_fileName, "\", "\\"), """", "\""")
        sea_attJson = ",""attachments"":[{""name"":""" & sea_fileNameJ & """,""contentBase64"":""" & sea_b64 & """}]"
    End If
End If

' ---- build and send ----
If sea_ok Then
    ' split "Name <addr>" (tolerates a missing closing ">")
    sea_fromLine = Trim(session("from_line") & "")
    sea_fromName = ""
    sea_fromAddr = sea_fromLine
    sea_p = InStr(sea_fromLine, "<")
    If sea_p > 0 Then
        sea_fromName = Trim(Left(sea_fromLine, sea_p - 1))
        sea_fromAddr = Trim(Replace(Mid(sea_fromLine, sea_p + 1), ">", ""))
    End If

    ' JSON-escape each text field: order matters (backslash first)
    sea_fields = Array(session("email") & "", session("bcc_line") & "", sea_fromAddr, sea_fromName, session("subject_line") & "", session("message") & "")
    For sea_i = 0 To UBound(sea_fields)
        sea_s = sea_fields(sea_i)
        sea_r = ""
        For sea_j = 1 To Len(sea_s)
            sea_c = Mid(sea_s, sea_j, 1)
            Select Case sea_c
                Case """":  sea_r = sea_r & "\"""
                Case "\":   sea_r = sea_r & "\\"
                Case vbCr:  sea_r = sea_r & "\r"
                Case vbLf:  sea_r = sea_r & "\n"
                Case vbTab: sea_r = sea_r & "\t"
                Case Else
                    If AscW(sea_c) < 32 Then
                        sea_r = sea_r & "\u" & Right("0000" & Hex(AscW(sea_c)), 4)
                    Else
                        sea_r = sea_r & sea_c
                    End If
            End Select
        Next
        sea_fields(sea_i) = sea_r
    Next

    sea_body = "{"
    sea_body = sea_body & """appName"":""legacy " & Request.ServerVariables("SERVER_NAME") & ""","
    sea_body = sea_body & """to"":""" & sea_fields(0) & ""","
    If Len(sea_fields(1)) > 0 Then sea_body = sea_body & """bcc"":""" & sea_fields(1) & ""","
    If Len(sea_fields(2)) > 0 Then sea_body = sea_body & """from"":""" & sea_fields(2) & ""","
    If Len(sea_fields(3)) > 0 Then sea_body = sea_body & """fromName"":""" & sea_fields(3) & ""","
    sea_body = sea_body & """subject"":""" & sea_fields(4) & ""","
    sea_body = sea_body & """htmlBody"":""" & sea_fields(5) & """"
    sea_body = sea_body & sea_attJson
    sea_body = sea_body & "}"

    On Error Resume Next
    Set sea_http = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
    sea_http.setTimeouts 15000, 15000, 120000, 300000
    sea_http.Open "POST", sea_url, False
    sea_http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    sea_http.setRequestHeader "X-API-Key", sea_key
    sea_http.Send sea_body
    If Err.Number <> 0 Then
        sea_err = "Error: " & Err.Description
        Err.Clear
        sea_ok = False
    Else
        sea_status = sea_http.status
        sea_resp = sea_http.responseText
        If sea_status < 200 Or sea_status >= 300 Then
            ' pull "error" and "requestId" out of the gateway's JSON reply
            sea_err = ""
            sea_p = InStr(sea_resp, """error""")
            If sea_p > 0 Then
                sea_p = InStr(sea_p, sea_resp, ":")
                sea_p = InStr(sea_p, sea_resp, """")
                sea_q = InStr(sea_p + 1, sea_resp, """")
                If sea_p > 0 And sea_q > sea_p Then sea_err = Mid(sea_resp, sea_p + 1, sea_q - sea_p - 1)
            End If
            If Len(sea_err) = 0 Then sea_err = "HTTP " & sea_status
            sea_rid = ""
            sea_p = InStr(sea_resp, """requestId""")
            If sea_p > 0 Then
                sea_p = InStr(sea_p, sea_resp, ":")
                sea_p = InStr(sea_p, sea_resp, """")
                sea_q = InStr(sea_p + 1, sea_resp, """")
                If sea_p > 0 And sea_q > sea_p Then sea_rid = Mid(sea_resp, sea_p + 1, sea_q - sea_p - 1)
            End If
            sea_err = "Error: " & sea_err & " (request " & sea_rid & ")"
            sea_ok = False
        End If
    End If
    On Error GoTo 0
    Set sea_http = Nothing
End If

' ---- report exactly like the old include did ----
If sea_ok Then
    Response.Write "Email sent successfully!"
    session("smtperror") = "Email sent successfully!"
Else
    Response.Write sea_err
    session("smtperror") = sea_err
End If
%>
