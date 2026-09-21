<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="clsUpload.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" class="view-type-desktop client-type-desktop client-browser-chrome client-browser-group-chrome client-browser-version-major-41 client-browser-version-minor-0" id="page-">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
	<link href="/includes/aps/ad-recruit/layout-1.css" rel="stylesheet" type="text/css" id="cc-layout-css"/>
	<link href="/includes/aps/ad-recruit/base.css" rel="stylesheet" type="text/css" id="cc-base-css"/>
	<link href="/includes/aps/ad-recruit/default.css" rel="stylesheet" type="text/css" id="cc-theme-css"/>
	<link href='https://fonts.googleapis.com/css?family=Open+Sans' rel='stylesheet' type='text/css'>
	<title>Career Center</title>
</head>
<BODY>
<div id="cc-container" class="cc-document">
	<div id="cc-content-outer">
		<div id="cc-content-inner">
			<!--#include virtual="/includes/aps/ad-recruit/sidebar.asp"-->
			<div id="cc-content" class="cc-panel">
				<h1>APS Career Center</h1>
				<br><br>
				<font size="3">
				Your listing needs a PDF file that includes all the relevant information for your recruitment ad. Please upload.
				</font>
				<br><br><br><br>
				<center>
				<FORM ACTION = "uploads.asp" ENCTYPE="multipart/form-data" METHOD="POST">
				File Name: <INPUT required TYPE=FILE NAME="txtFile"><P>
				<table>
				<tr>
				<td><INPUT class="cc-btn-primary" style="background:#A3276C;font-weight: bold; color:white;" TYPE="SUBMIT" NAME="cmdSubmit" VALUE=" Upload "></td>
				<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
				<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
				<td><font size="4"><%=session("pdfuploadresult")%></font></td>
				<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
				<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
				<td><input class="cc-btn-primary" style="background:#A3276C;font-weight: bold; color:white;" type="button" onclick="location.href = 'https://www.aps1888.org/'" value=" EXIT " name="B1"></td>
				</tr>
				</table>
				</FORM>
				<P>
				<%
				session("pdfuploadresult")=""
				set o = new clsUpload
				if o.Exists("cmdSubmit") then
				'get client file name without path
				sFileSplit = split(o.FileNameOf("txtFile"), "\")
				sFile = sFileSplit(Ubound(sFileSplit))
				o.FileInputName = "txtFile"
				o.FileFullPath = Server.MapPath(".") & "\" & sFile
				o.save
				 if o.Error = "" then
					session("file-ext")=right(o.FileFullPath,3)
					session("len-pdf")=len(o.FileFullPath)
					' site-relative path (e.g. ad-recruit\PDF\file.pdf), computed from the real web root
					' instead of a fixed character offset, so the stored URL is right on any server
					session("pdf1")=mid(o.FileFullPath, len(Server.MapPath("/")) + 2)
					session("pdf")="https://www.aps1888.org/" & replace(session("pdf1"), "\", "/")
					session("pdfuploadresult")="Success. File uploaded.<br>Thank you. Your position will post soon."
					save_pdf = "UPDATE tblAPS_Recruit_ads SET PDFfile='" & session("pdf") & "' WHERE email='"&session("email")&"' and timestamp='" &session("timestamp")&"'"
					set my_Conn= Server.CreateObject("ADODB.Connection")
					my_Conn.ConnectionTimeout=Session("ConnectionTimeOut")
					my_Conn.CommandTimeOut=Session("CommandTimeout")
					my_Conn.Open Session("ConnectionMembers")
					set RS=my_conn.Execute (save_pdf)
					session("uploadedpdf")="Y"
					'
					if session("amount")=995 then
						session("prod") = "<strong>Basic Package</strong><br>"
					end if
					if session("amount")=1850 then
						session("prod") = "<strong>Basic+ Package</strong><br>"
					end if
					if session("amount")=3250 then
						session("prod") = "<strong>Advanced Connect Package</strong><br>"
					end if

					' ---- notify staff through the MailGateway, with the uploaded file attached ----
					' The include reads session("email") as the To address, so the buyer's address is
					' parked in session("email_buyer") while the staff address takes its place.
					session("subject_line") = "Attachment for APS Recruitment Center"
					session("from_line")    = "APS Central Office <noreply@aps1888.org>"
					session("bcc_line")     = ""
					session("message")      = "Attachment for APS Recruitment Center<br><br><br>" & sFile & "<br>" & session("prod") & "<br>" & session("email") & "<br>" & Session("organization") & "<br>" & session("title") & "<br>"
					session("pdf_localpath") = o.FileFullPath
					session("email_buyer")  = session("email")
					session("email")        = "aps.recruit.ads@aps-spr.org"
					%>
					<!--#include virtual="/includes/send_email_ad.asp"-->
					<%
					session("email")         = session("email_buyer")
					session("pdf_localpath") = ""
					' response.redirect "https://www.aps1888.org/ad-recruit/endupload.asp"
				 else
					response.write "Failed due to the following error: " & o.Error
					session("pdfuploadresult")="Failed due to the following error: " & o.Error
				 end if
				end if
				set o = nothing
				%>
			</div>
		</div>
	</div>
</div>
</BODY>
</html>
