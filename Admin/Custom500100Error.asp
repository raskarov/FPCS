<%@ Language=VBScript %>
<% Response.Clear %>
<html>
<head>
	<LINK rel="stylesheet" type="text/css" href="/FPCSdev/css/honestyle.css">
</head>
<body>

<%
on error resume next
Dim oASPErr		'ASP 3.0 Server Error Object
dim strASPCode
dim intNumber
dim strSource
dim strCategory
dim strFile
dim intLine
dim intColumn
dim strDescription
dim strASPDescription
dim strAdminEmail			'Administrators' email addresses
dim strMessage		'Email body
dim strIntro


	strAdminEmail = Application.Value("strAdminEmail")

	Set oASPErr = Server.GetLastError
	with oASPErr
		strASPCode			= .ASPCode
		intNumber			= .Number
		strSource			= .Source
		strCategory			= .Category
		strFile				= .File
		intLine				= .Line
		intColumn			= .Column
		strDescription		= .Description
		strASPDescription = .ASPDescription
	end with

	If strASPCode <> "" Then
		strMessage = strMessage & "<tr><td width='20%' align='left'><b>Error #:</b></td>" & vbCR
		strMessage = strMessage & "<td>" & strASPCode & "</td></tr>"
	End If
	
	If intrNumber <> 0 Then
		strMessage = strMessage & "<tr><td width='20%' align='left'><b>COM Error #:</b></td>" & vbCR
		strMessage = strMessage & "<td>" & intNumber & "</td></tr>"
	End If
	
	If strSource <> "" Then
		strMessage = strMessage & "<tr><td width='20%' align='left'><b>Error Source:</b></td>" & vbCR
		strMessage = strMessage & "<td>" & strSource & "</td></tr>"
	End If
	
	If strCategory <> "" Then
		strMessage = strMessage & "<tr><td width='20%' align='left'><b>Error Category:</b></td>" & vbCR
		strMessage = strMessage & "<td>" & strCategory & "</td></tr>"
	End If
	
	If strFile <> "" Then
		strMessage = strMessage & "<tr><td width='20%' align='left'><b>File:</b></td>" & vbCR
		strMessage = strMessage & "<td>//" & Request.ServerVariables	("SERVER_NAME") & strFile & "</td></tr>"
	End If

	If intLine <> 0 Then
		strMessage = strMessage & "<tr><td width='20%' align='left'><b>Line, Column:</b></td>" & vbCR
		strMessage = strMessage & "<td>" & intLine & ", " & intColumn & "</td></tr>"
	End If
	
	If strDescription <> "" Then
		strMessage = strMessage & "<tr><td width='20%' align='left'><b>Description:</b></td>" & vbCR
		strMessage = strMessage & "<td><font color='red'><b>" & strDescription & "</b></font></td></tr>"
	End If
	
	If strASPDescription <> "" Then
		strMessage = strMessage & "<tr><td width='20%' align='left'><b>ASPDescription:</b></td>" & vbCR
		strMessage = strMessage & "<td>" & strASPDescription & "</td></tr>"
	End If

	strMessage = strMessage & "<tr><td width='20%' align='left'><b>Credentials:</b></td>" & vbCR
	strMessage = strMessage & "<td>" & Request.ServerVariables("REMOTE_USER") & "</td></tr>"
	

	'***********************************************************
	'Section:	Send Email
	'Purpose:	Sends an email to the Central Dispatcher alerting
	'				them of a new service request and providing them a link
	'				to click on. We do not send an email if this request is
	'				being made by the Central Dispatcher
	'***********************************************************
	Dim iMsg		'As New CDO.Message
	Dim iConf	'As New CDO.Configuration
	  
	set iMsg = Server.CreateObject("CDO.Message")
	set iConf = Server.CreateObject("CDO.Configuration")
	  
	With iConf.Fields
		.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = Application("strSMTPserver")
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30 ' quick timeout
		.Update
	End With

	strIntro = "<html><body><font face='Tahoma'>ASP Error occurred on <b>" & now() & _
		"</b>.<br><br>" & vbCR & "<table cols='2'>"

	With iMsg
		Set .Configuration = iConf
		.To = strAdminEmail
		.From = "3Shapes AutoBot <donotreply@3Shapes>"
		.Subject = "ASP Server Error: " & strDescription
		.HTMLBody = strIntro & strMessage & "</table></body></html>"
		.Send
		select case Err.number
			'case 80040212 'means the SMTP sever listed in Application("strSMTPserver") is not available
			'	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "161.99.66.31"	'AMANCX1 Exchange server
			'	.Send
			case else
				if Err.number <> 0 then
					strMessage = strMessage & "<tr><td width='20%' align='left'><b>Email ERROR:</b></td>" & vbCR
					strMessage = strMessage & "<td>" & Err.number & ": - " & Err.description & " EMAIL NOT SENT</td></tr>"
				end if		
		end select
	End With
	
	set iConf = nothing
	set iMsg = nothing
	'***********************************************************
	'end EMAIL section
	'***********************************************************
		
%>
FPCS: An error has occurred in this application!<br><br>
<b><% = now() %></b><BR><BR>
<table border="0">
	<tr>
		<td align="left" >
			The error was not caused by anything that you did.
		</td>
	</tr>
	<tr>
		<td align="left" >
			A copy of this screen has been emailed to the FPCS(3Shapes) Analyst.
		</td>
	</tr>
</table>
<br></br>
<table border="1" cols="2">
<% 
	Response.Write strMessage
%>
</table><BR><BR>
<%
dim strMethod
dim lngPos
Const lngMaxFormBytes = 200

 strMethod = Request.ServerVariables("REQUEST_METHOD")

  Response.Write strMethod & " "

  If strMethod = "POST" Then
    Response.Write Request.TotalBytes & " bytes to "
  End If

  Response.Write Request.ServerVariables("SCRIPT_NAME")

  lngPos = InStr(Request.QueryString, "|")

  If lngPos > 1 Then
    Response.Write "?" & Left(Request.QueryString, (lngPos - 1))
  End If

  Response.Write "</li>"

  If strMethod = "POST" Then
    Response.Write "<p><li>POST Data:<br>"
    If Request.TotalBytes > lngMaxFormBytes Then
       Response.Write Server.HTMLEncode(Left(Request.Form, lngMaxFormBytes)) & " . . ."
    Else
      Response.Write Server.HTMLEncode(Request.Form)
    End If
    Response.Write "</li>"
  End If
%>
</body>
</html>
