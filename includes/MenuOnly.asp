<%@ Language=VBScript %>
<%
Option Explicit
'**********************************************************************
'Name:      header.asp
'Purpose:   HTML header. This include provides the opening tags. 
'
'Note**:	This page is  called indirectly by way of Server.Execute.  Thus,
'           all variables used here are not available to other pages.  Use
'           Session and Applicaiton variables instead.
'
'Author:    Bryan K Mofley
'Date:      22 Feb 2002
'**********************************************************************
dim strLastUpdate
dim strRightTitle
dim faI
' No seeing webpages unless your are in SSL mode (should be port 443)
'if Request.ServerVariables("SERVER_PORT_SECURE") <> 1 then
'	Response.Redirect "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
'end if

' No seeing webpages until you log in
if Session("bolUserLoggedIn") = false  and inStr(1,Request.ServerVariables("URL"),"EmailPassword.asp") < 1 then
	'session.Value("strURL") = "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
	'Response.Redirect(Application("strSSLWebRoot") & "UserAdmin/login.asp")
	Session.Value("strURL") = Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
	Response.Redirect(Application("strWebRoot") & "UserAdmin/login.asp")
end if 

' Force action if required
if session.Value("bolActionNeeded") = true and Request.QueryString("bolForced") = "" then
	for faI = 0 to ubound(session.Value("arActions"))
		if session.Value("arActions")(faI,1) <> "" then
			Response.Redirect(Application("strWebRoot") & session.Value("arActions")(faI,1))
		end if 
	next 
end if 
strLastUpdate = "Today" 'ShowFileDate(Server.MapPath(Request.ServerVariables("PATH_INFO")) )

Function ShowFileDate(filespec)
 'from google.com
 Dim dfso, df, ds
 	Set dfso = CreateObject("Scripting.FileSystemObject")
 	Set df = dfso.GetFile(filespec)
 	ds = Day(df.DateLastModified) & " " & MonthName(Month(df.DateLastModified), true) & " " & Year(df.DateLastModified)
 	ShowFileDate = ds
End Function


 'place signin/signout buttons in the top row
 if Session.Value("bolUserLoggedIn") = false then
 	strRightTitle = "<button onclick=" & """" & _
 		"window.open('" & Application.Value("strWebRoot") & "UserAdmin/login.asp', '_self');" & _
 		"""" & ">Sign In</button>"
 else
 	strRightTitle = Session.Value("strUserID") & _
 		"&nbsp;<button onclick=" & """" & _
 		"window.open('" & Application.Value("strWebRoot") & "UserAdmin/logout.asp', '_self');" & _
 		"""" & ">Sign Out</button>"
 end if
 
 dim oFunc
 set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
 
 ' Security check to make sure logged in user only accesses student info that they have rights to 
 call oFunc.OpenCN()
 call oFunc.CheckAuth()
%>
<!-- Begin header.asp -->
<html>
<head>	
	<link rel="stylesheet" type="text/css" href="<% = Application.Value("strWebRoot") %>css/homestyle.css">
	<title><% = Session.Value("strTitle") %></title>
	<!-- This doesn't seem to work on a mac IE 5-->
	<script language=javascript src="<%= Application.Value("strWebRoot") %>includes/formCheck.js"></script>	
	<script language=javascript src="<%= Application.Value("strWebRoot") %>includes/appJSFunctions.js"></script>	
	<script language=javascript>

if (window.name != "<% = Session.Value("strAppWindow")%>"){
	window.open("<% = Application.Value("strWebRoot") %>default.asp",'<% = Session.Value("strAppWindow")%>', config='toolbar=no,menubar=no,scrollbars=yes,resizable=yes,location=no,directories=no,status=yes'); 
	window.location.replace("https://<% = Request.ServerVariables("SERVER_NAME")%>");
}
</script>
<STYLE>
p { page-break-before: always }
</STYLE>
</head>
<body leftmargin="0" topmargin="0" onload="window.focus();" alink=red vlink=navy link=navy onunload="//window.open('<% = Application.Value("strWebRoot") %>includes/cleanup.asp');">
<table bgcolor="navy" width="100%" cellpadding="0" cellspacing="1" border="0" ID="Table1">
   <tr style="FONT-FAMILY: Verdana; COLOR: White;FONT-SIZE: 70.5%; FONT-WEIGHT: bolder">
      <td align="left" valign="middle" >
         &nbsp;&nbsp;Your Gateway to the FPCS Online System
      </td>
      <td align="right" valign="middle">
         <% = strRightTitle %>
      </td>
   </tr>
</table>
<table border="0" cellpadding=0 cellspacing=0 width=100% ID="Table2">
   <tr bgcolor=e6e6e6>
	  <td width=0>
		&nbsp;
	  </td>
      <td valign="top">         
         <table width=100% cellpadding="0" cellspacing="0" border="0" style="border-right-width: 1px" ID="Table3">
            <tr>               
                 <td>
                  <table cellpadding="0" cellspacing="0" border="0" ID="Table4">
                     <tr>
                        <td class="flyoutLink" onclick="window.location.href='https://<% = Request.ServerVariables("SERVER_NAME") & left(Application.Value("strWebRoot"),len(Application.Value("strWebRoot"))-1) %>';"> 
                          <B>Home</B>
                        </td>
                     </tr>
                  </table>
                 </td>
                 <td>
                  <table ID="Table5">
					<tr>
                 <%                                    
                  response.Write oFunc.AddMenu("addILP")                  
                 ' response.Write oFunc.AddMenu("enrollmentInfo")                  
                  response.Write oFunc.AddMenu("familyManager")
                  response.Write oFunc.AddMenu("ilpBank")
                  response.Write oFunc.AddMenu("manageClasses")
                  response.Write oFunc.AddMenu("shortForm")
                  response.Write oFunc.AddMenu("statement")
                  response.Write oFunc.AddMenu("budget")
                  response.Write oFunc.AddMenu("transfer")
                  response.Write oFunc.AddMenu("forms")
                  response.Write oFunc.AddMenu("teacherProfile")
                  response.Write oFunc.AddMenu("teacherBio")
                  response.Write oFunc.AddMenu("teacherBioEdit")                  
                  response.Write oFunc.AddMenu("vendorAuth")
                  response.Write oFunc.AddMenu("vendorProfile")
                  response.Write "</tr></table>  "
%>
				</td>
				<td width=100% class=svplain8>
					&nbsp;&nbsp;School Year <% = oFunc.SchoolYearRange%>
				</td>				
            </tr>
            
         </table>  
      </td>
     <tr>
      <td colspan=10 id="tdContent" style="PADDING-LEFT: 9px; PADDING-TOP: 10px" valign="top">
<!-- End header.asp -->		


