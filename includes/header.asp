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
	Response.Redirect(Application("strSSLWebRoot") & "UserAdmin/login.asp")
end if 

if (application.Contents("dtSchool_Year_Start" & session.Contents("intSchool_Year")) = "" or _
	application.Contents("dtSchool_Year_End" & session.Contents("intSchool_Year")) = "")  and _
	ucase(session.Contents("strRole")) = "ADMIN" and inStr(1,Request.ServerVariables("URL"),"globalvariables.asp") < 1 then
	response.Redirect(Application("strSSLWebRoot") & "Admin/globalvariables.asp")
elseif (application.Contents("dtSchool_Year_Start" & session.Contents("intSchool_Year")) = "" or _
	application.Contents("dtSchool_Year_End" & session.Contents("intSchool_Year")) = "")  and _
	ucase(session.Contents("strRole")) <> "ADMIN" then
	
	response.Write "The " & session.Contents("intSchool_Year") & " school year has not been set up for access by the Principal. " & _
				   " Until this school year is set up for use it can not be accessed. "
	response.End
end if
	

' Force action if required
'if session.Value("bolActionNeeded") = true and Request.QueryString("bolForced") = "" then
'	for faI = 0 to ubound(session.Value("arActions"))
'		if session.Value("arActions")(faI,1) <> "" then
'			Response.Redirect(Application("strWebRoot") & session.Value("arActions")(faI,1))
'		end if 
'	next 
'end if 
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
 		"window.open('" & Application.Value("strSSLWebRoot") & "UserAdmin/login.asp', '_self');" & _
 		"""" & ">Sign In</button>"
 else
 	strRightTitle = Session.Value("strUserID") & _
 		"&nbsp;<button onclick=" & """" & _
 		"window.open('" & Application.Value("strSSLWebRoot") & "UserAdmin/logout.asp', '_self');" & _
 		"""" & ">Sign Out</button>"
 end if
 
 dim oFunc
 set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
 
 ' Security check to make sure logged in user only accesses student info that they have rights to 
 call oFunc.OpenCN()
 call oFunc.CheckAuth()
 dim oMenu
 set oMenu = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/Menu.wsc"))
' for each i in session.Contents
'	response.Write i & " - " & session.Contents(i) & "<BR>"
'next

%>
<!-- Begin header.asp -->
<html>
<head>	
	<link rel="stylesheet" type="text/css" href="<% = Application.Value("strSSLWebRoot") %>css/homestyle.css">
	<title><% = Session.Value("strTitle") %></title>
	<!-- This doesn't seem to work on a mac IE 5-->
	<script language=javascript src="<%= Application.Value("strSSLWebRoot") %>includes/formCheck.js"></script>	
	<script language=javascript src="<%= Application.Value("strSSLWebRoot") %>includes/appJSFunctions.js"></script>	
	<script language=javascript>

if (window.name != "<% = Session.Value("strAppWindow")%>"){
	window.open("<% = Application.Value("strSSLWebRoot") %>default.asp",'<% = Session.Value("strAppWindow")%>', config='toolbar=no,menubar=no,scrollbars=yes,resizable=yes,location=no,directories=no,status=yes'); 
	window.location.replace("https://<% = Request.ServerVariables("SERVER_NAME")%>");
}
</script>
<STYLE>
p { page-break-before: always }

/* Default Menu Style (Opera inspired) */
div.domMenu_menuBar {
    border: solid #7E7E7E;  
    border-width: 1px 0 0 1px;
}
div.domMenu_menuElement {
    font-family: Arial, sans-serif; 
    font-size: 10px;
    border: solid #7E7E7E;  
    border-width: 0 1px 1px 0;
    background: url(<% = Application("strSSLWebRoot") %>images\gradient.png) repeat-x; 
    color: #0F0F0F;
    text-align: center;
    height: 20px;
    line-height: 20px;
    vertical-align: middle;
    font-weight: bold;
}
div.domMenu_menuElementHover {
    background: url(<% = Application("strSSLWebRoot") %>images\gradient_hover.png) repeat-x;
}
div.domMenu_subMenuBar {
    border: solid #7E7E7E 0px;
    background-color: #FFFFFF;
    padding-bottom: 1px;
    opacity: .9;
    filter: alpha(opacity=90);
}
div.domMenu_subMenuElement {
    font-family: Arial, sans-serif; 
    font-size: 10px;
    border: solid #CCCCCC 1px;
    margin: 1px 1px 0 1px;
    color: #0F0F0F;
    padding: 2px 2px;
}
div.domMenu_subMenuElementHover {
    background-color: #c0c0c0;
}
</STYLE>
</head>
<body leftmargin="0" topmargin="0" onload="window.focus();" alink=red vlink=navy link=navy >
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
   <tr bgcolor=e6e6e6>
	  <td width=0>
		&nbsp;
	  </td>
      <td valign="top">         
         <table width=100% cellpadding="2" cellspacing="0" border="0" style="border-right-width: 1px" ID="Table3">
            <tr>               
                <td>
					<div id="domMenu_main" style="background-color: #FFFFFF;"></div>
				</td>
				<td class=svplain8 nowrap align="left">
					&nbsp;&nbsp;School Year <% = oFunc.SchoolYearRange%>
				</td>				
            </tr>
            <%	if request("intStudent_ID") <> "" then 
					dim strStudentList
					dim strSQL
					if ucase(session.Contents("strRole")) = "ADMIN" then
						strSQL = "SELECT     s.intSTUDENT_ID, (CASE ss.intReEnroll_State WHEN 86 THEN s.szLAST_NAME + ',' + s.szFIRST_NAME + ': Withdrawn (' + CASE isNull(ss.dtWithdrawn, " & _ 
											" 1) WHEN 1 THEN 'No Date Entered' ELSE CONVERT(varChar(100), ss.dtWithdrawn)  " & _ 
											" END + ')' WHEN 123 THEN s.szLAST_NAME + ',' + s.szFIRST_NAME + ': Graduated (' + CONVERT(varChar(20), ss.dtModify)  " & _ 
											" + ')' ELSE s.szLAST_NAME + ',' + s.szFIRST_NAME END) AS Name " & _ 
											"FROM tblSTUDENT s INNER JOIN " & _ 
											" tblStudent_States ss ON s.intSTUDENT_ID = ss.intStudent_id " & _ 
											"WHERE (ss.intReEnroll_State IN (" & Application.Contents("strEnrollmentList") & ")) AND (ss.intSchool_Year = " & Session.Value("intSchool_Year") & ") " & _ 
											"ORDER BY Name "
					elseif ucase(session.Contents("strRole")) = "TEACHER" then
						strSQL = "SELECT tblSTUDENT.intStudent_ID, tblSTUDENT.szLAST_NAME + ', ' + tblSTUDENT.szFIRST_NAME AS Name " & _
								"FROM tblENROLL_INFO INNER JOIN tblSTUDENT " & _
								" ON tblENROLL_INFO.intSTUDENT_ID = tblSTUDENT.intSTUDENT_ID INNER JOIN " & _
								" tblFAMILY ON tblSTUDENT.intFamily_ID = tblFAMILY.intFamily_ID " & _
								"WHERE (tblENROLL_INFO.intSponsor_Teacher_ID = " & session.Contents("instruct_ID") & _
								") AND (tblENROLL_INFO.sintSCHOOL_YEAR = " & session.Contents("intSchool_Year") & ") " & _
								"ORDER BY tblSTUDENT.szLAST_NAME"
					elseif ucase(session.Contents("strRole")) = "GUARD" then
						' Guards can ONLY see ACTIVE students
						strSQL = "SELECT s.intSTUDENT_ID, s.szLAST_NAME + ',' + s.szFIRST_NAME AS Name " & _ 
									"FROM tblSTUDENT s INNER JOIN " & _ 
									"tblFAMILY f ON s.intFamily_ID = f.intFamily_ID INNER JOIN " & _ 
									"tblStudent_States ss ON s.intSTUDENT_ID = ss.intStudent_id " & _ 
									"WHERE     (f.intFamily_ID = " & Session.Value("intFamily_ID") & ") AND (ss.intSchool_Year = " & Session.Value("intSchool_Year") & ") " & _ 
									"AND ss.intReEnroll_State  IN (" & Application.Contents("ActiveEnrollList") & ")  " & _ 										
									"ORDER BY Name" 
					end if
					strStudentList = oFunc.MakeListSQL(strSQL,"","",request("intStudent_ID"))
					if oFunc.makeListRecordCount > 1 then
					%>
					<tr>
						<td colspan=10>
							<table ID="Table6">
								<tr>
									<td	class=svplain8>
										&nbsp;<B>Change Student:</B>
									</td>
									<td>
										<select name=intStudent_ID onchange="window.location.href='<%=request.servervariables("SERVER_NAME ") & request.servervariables("SCRIPT_NAME") & "?intStudent_ID="%>'+this.value;" ID="Select1">
											<option></option>
											<% = strStudentList %>
										</select>
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<%
					end if
				end if
				' check to see if year is locked
				if  oFunc.LockYear then
					%>
					<tr>
						<td colspan=10>
							&nbsp;&nbsp;<span class="error10"><b>Please Note: This school year has been locked and no modifications can be made.</b></span>
						</td>
					</tr>
					<%
				elseif  oFunc.LockSpending  then
					%>
					<tr>
						<td colspan=10>
							&nbsp;&nbsp;<span class="error10"><b>Please Note: Spending has been locked and no budget modifications can be made.</b></span>
						</td>
					</tr>
					<%
				elseif  oFunc.ShowLockMsg then
					%>
					<tr>
						<td colspan=10>
							&nbsp;&nbsp;<span class="error10"><b>Please Note: Spending has been locked.</b></span>
						</td>
					</tr>
					<%
				end if
					%>
         </table>  
      </td>
     <tr>
      <td colspan=10 id="tdContent" style="PADDING-LEFT: 9px; PADDING-TOP: 10px" valign="top">
<!-- End header.asp -->		
<%	
	response.Write oMenu.GetMenu() 
	set oMenu = nothing 
	oFunc.CloseCN()
	set oFunc = nothing
%>

