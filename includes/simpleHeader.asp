<%@ Language=VBScript %>
<%
'**********************************************************************
'Name:      simpleHeader.asp
'Purpose:   HTML header. This include provides the opening tags. 
'
'Note**:	This page is  called indirectly by way of Server.Execute.  Thus,
'           all variables used here are not available to other pages.  Use
'           Session and Applicaiton variables instead.
'
'Author:    Scott Bacon, ThreeShapes LLC
'Date:      04/24/2002
'**********************************************************************
' No seeing webpages unless your are in SSL mode (should be port 443)
'if Request.ServerVariables("SERVER_PORT_SECURE") <> 1 then
'	Response.Redirect "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
'end if

' No seeing webpages until you log in
if Session("bolUserLoggedIn") = false and inStr(1,Request.ServerVariables("URL"),"EmailPassword.asp") < 1 _
	and inStr(1,Request.ServerVariables("URL"),"login.asp") < 1 then
	'session.Value("strURL") = "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
	'Response.Redirect(Application("strSSLWebRoot") & "UserAdmin/login.asp")
	Session.Value("strURL") = Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
	Response.Redirect(Application("strWebRoot") & "UserAdmin/login.asp")
end if 

' No seeing students info that doesn't belong to you
if ucase(session.Contents("strRole")) <> "ADMIN" then
	if request("intStudent_ID") <> "" then
		if instr(1,session.Contents("student_list"),"~" & request("intStudent_ID") & "~") < 1 then
			
			dim cnCrack
			dim insertCrack
			set cnCrack = server.CreateObject("ADODB.CONNECTION")
			cnCrack.Open(Application("cnnFPCS"))
			insertCrack = "insert into tblCrack_Attempts(szUser_ID,szIP_Address,szURL,szALL,szUser_Create) " & _
						" values (" & _
						"'" & session.Contents("strUserID") & "'," & _
						"'" & request.ServerVariables("REMOTE_ADDR") & "'," & _
						"'" & replace(Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING"),"'","''") & "'," & _
						"'" & replace(request.ServerVariables("ALL_HTTP"),"'","''") & "'," & _
						"'" & session.Contents("strUserID") & "')" 
			cnCrack.Execute(insertCrack)
			cnCrack.Close
			set cnCrack = nothing
			response.Write "<H1>Security Breach Attempt. Your user information has been captured and sent to the Admin.</h1>" & _
						   "<br>This action can cause your account to be terminated and legal steps may also be taken."
			response.End
		end if
	end if
end if

' Force action if required
if session.Value("bolActionNeeded") = true and Request.QueryString("bolForced") = "" _
	and Request.QueryString("exempt") = "" then
	for faI = 0 to ubound(session.Value("arActions"))
		if session.Value("arActions")(faI,1) <> "" then
			Response.Redirect(Application("strWebRoot") & session.Value("arActions")(faI,1))
		end if 
	next 
end if 
'for each i in session.Contents
'	response.Write i & " - " & session.Contents(i) & "<BR>"
'next
%>	
<html>
<head>
<title><% = session.Value("strTitle") %></title>
<link rel="stylesheet" href="<% =  Application.Value("strWebRoot") %>CSS/homestyle.css">
	<!-- This doesn't seem to work on a mac IE 5 -->
<script language=javascript src="<%= Application.Value("strWebRoot") %>includes/formCheck.js"></script>	
<script language=javascript src="<%= Application.Value("strWebRoot") %>includes/appJSFunctions.js"></script>	
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/1/jquery.min.js" type="text/javascript"></script>
    <script src="http://ajax.googleapis.com/ajax/libs/jqueryui/1/jquery-ui.min.js" type="text/javascript"></script>
    <link href="http://ajax.googleapis.com/ajax/libs/jqueryui/1/themes/base/jquery-ui.css"
        rel="Stylesheet" type="text/css" />
<STYLE>
p { page-break-before: always }
</STYLE>
</head>
<!-- oncontextmenu="javascript:return false;"  add this to body to disable right click menus -->
<body bgcolor="#ffffff" onLoad="<% = session.Value("simpleOnLoad") %>" background="<% = session.contents("strBGImagePath") %>">
