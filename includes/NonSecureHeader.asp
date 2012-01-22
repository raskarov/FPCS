<%@ Language=VBScript %>
<%
Option Explicit
'**********************************************************************
'Name:      NonSecurHeader.asp
'Purpose:   HTML header. This include provides the opening tags. 
'
'Note**:	This page is  called indirectly by way of Server.Execute.  Thus,
'           all variables used here are not available to other pages.  Use
'           Session and Applicaiton variables instead.
'
'Author:    Scott Bacon
'Date:      11 July 2005
'**********************************************************************
dim strLastUpdate
dim strRightTitle
dim faI
 
strLastUpdate = "Today"

 'place signin/signout buttons in the top row

 	strRightTitle = Session.Value("strUserID") & _
 		"&nbsp;<button onclick=" & """" & _
 		"window.location.href='" & Application("strURL") & "';" & _
 		"""" & ">Exit</button>"
 
 dim oFunc
 set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
 
 ' Security check to make sure logged in user only accesses student info that they have rights to 
 call oFunc.OpenCN()
 call oFunc.CheckAuth()
 dim i
' for each i in session.Contents
'	response.Write i & " - " & session.Contents(i) & "<BR>"
'next

%>
<!-- Begin header.asp -->
<html>
<head>	
	<link rel="stylesheet" type="text/css" href="<% = Application.Value("strWebRoot") %>css/homestyle.css">
	<title><% = Session.Value("strTitle") %></title>
	<!-- This doesn't seem to work on a mac IE 5-->
	<script language=javascript src="<%= Application.Value("strWebRoot") %>includes/formCheck.js"></script>	
	<script language=javascript src="<%= Application.Value("strWebRoot") %>includes/appJSFunctions.js"></script>	
<STYLE>
p { page-break-before: always }
</STYLE>
</head>
<body leftmargin="0" topmargin="0" onload="window.focus();" alink=red vlink=navy link=navy onunload="//window.open('<% = Application.Value("strWebRoot") %>includes/cleanup.asp');">
<table bgcolor="navy" width="100%" cellpadding="0" cellspacing="1" border="0" ID="Table1">
   <tr style="FONT-FAMILY: Verdana; COLOR: White;FONT-SIZE: 70.5%; FONT-WEIGHT: bolder">
      <td align="left" valign="middle" >
         &nbsp;&nbsp;FPCS Online System Version 3.0 
      </td>
      <td align="right" valign="middle">
         <% = strRightTitle %>
      </td>
   </tr>
</table>
<table border="0" cellpadding=0 cellspacing=0 width=100% ID="Table2">
     <tr>
      <td colspan=10 id="tdContent" style="PADDING-LEFT: 9px; PADDING-TOP: 10px" valign="top">
<!-- End NonSecurHeader.asp -->		


