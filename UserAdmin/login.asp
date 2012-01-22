<%@ Language=VBScript %>
<%
'*******************************************
'Name:		UserAdmin\login.asp
'Purpose:	Displays a form to validate the user
'
'NOTES:	settings for cookies: (subtract 1 letter)
'			qbt =	pas	(szPassword)
'			hc =	id	(szUserID)
'
'Author:	Bryan K Mofley
'			3Shapes is Scott Bacon, Bryan Mofley, Guy Mofley
'Date:   15-May-2000
'*******************************************
'option explicit
dim strTitle		'HTML Title
dim cnnAccess		'ADO Connection Object
dim strUserID		'From request.form OR request.cookie
dim strPassword	'From request.form
dim strURL			'Session variable
dim mstrLoginMsg	'If Login Fails, this variable advises user

'*******************************************************************************
dim strUA			'User Agent
dim strBrowserMsg	'HTML to send back to client regarding instructions for IE

	strBrowserMsg = "<HTML><HEAD><TITLE>FPCS Online Office Version 3.0</TITLE>" & _
		"<link rel='stylesheet' type='text/css' href='" & Application.Value("strWebRoot") & _
		"css/homestyle.css'></HEAD><BODY>" & _
		"<CENTER>Your browser does not meet the minimum qualifications as required " & _
		"by the FPCS Online Office Application.<BR>You must have a generation 5 or greater browser and can not use Internet Explorer 5.x on a mac.<BR><BR>"

	strUA= Request.ServerVariables("HTTP_USER_AGENT") 
	strUA= lcase(strUA)  
	ie5 = instr(strUA,"msie 5") 
	ie6 = instr(strUA,"msie 6") 
	ie7 = instr(strUA,"msie 7") 
	ie8 = instr(strUA,"msie 7") 
	moz = instr(strUA,"mozilla/5.0")
	mac = instr(strUA,"mac")
	ie = instr(strUA,"msie")
	
	if (ie > 0 and mac > 0)  then
		strBrowserMsg  = strBrowserMsg  &  "<a href='http://www.mozilla.org/products/firefox/' target='_new'>Download Firefox</a>"

		strBrowserMsg = strBrowserMsg & "<BR><BR>Your browser is reporting the following:<BR>" & _
			Request.ServerVariables("HTTP_USER_AGENT") & "</CENTER></BODY></HTML>"
		Response.Write strBrowserMsg 
		Response.End
	elseif (ie or moz) then
		'do nothing
	else
		strBrowserMsg  = strBrowserMsg  &  "<a href='http://www.mozilla.org/products/firefox/' target='_new'>Download Firefox</a>"

		strBrowserMsg = strBrowserMsg & "<BR><BR>Your browser is reporting the following:<BR>" & _
			Request.ServerVariables("HTTP_USER_AGENT") & "</CENTER></BODY></HTML>"
		Response.Write strBrowserMsg 
		Response.End
	end if
'*******************************************************************************
	 strLoginMsg = "&nbsp;"
	 strPassword = replace(replace(Request.Form("szPassword"),"'",""),";","") 
	 strURL = Session.Contents("strURL")
	'if ucase(session.Contents("strUserId")) = "SCOTT" then response.Write strURL
	'JD if Request.Form("strURL") <> "" then
	'JD 	strURL = Request.Form("strURL")
	'JD 	Session.Contents("strURL") = strURL
	'JD end if
	
	'if ucase(session.Contents("strUserId")) = "SCOTT" then response.Write strURL
	
	'JD if strURL = "" then
	'JD 	strURL = Application.Value("strSSLWebRoot")
	'JD end if
	
	'if ucase(session.Contents("strUserId")) = "SCOTT" then response.Write strURL
	
	strUserID = replace(replace(Request.Form("szUserID"),"'",""),";","")
	
	'JD if strUserID & "" = "" then 
	'JD 	response.end
	'JD end if
	
	Response.Expires = -1000 'Makes the browser not cache this page
	Response.Buffer = True 'Buffers the content so our Response.Redirect will work


	Session.Contents("bolUserLoggedIn") = False
	If Request.Form("hLogin") = "True" Then 'only true if this page is a redirection of itself
	    call subCheckLogin 
	Else
		session.Abandon
	    call subShowLogin 
	End If 


Sub subShowLogin 
'*******************************************
'Name:		subShowLogin (procedure)
'Purpose:	Shows the Login Form
'
'Author:		Bryan K Mofley
'				3Shapes is Scott Bacon, Bryan Mofley, Guy Mofley
'Date:      27-Nov-2001
'*******************************************
Session.Contents("strTitle") = "FPCS Online Office Version 3.0 Log in"
session.Contents("simpleOnLoad") = "frmLogin.szUserID.focus();"
Server.Execute(Application.Value("strWebRoot") & "Includes/simpleheader.asp")
' ALWAYS CLEAR session.Contents("simpleOnLoad") OR ELSE IT WILL PERSIT TO OTHER PAGES!!!
session.Contents("simpleOnLoad") = ""
%>
<script language="javascript">	
function jfCheckValue(form) {
	if (form.szUserID.value == "") {
		alert("All fields are required [User ID]");
		form.szUserID.focus();
		return;
	}
	if (form.szPassword.value == "") {
		alert("All fields are required [Password]");
		form.szPassword.focus();
		return;
	}
	form.submit();
}	

function OpenCertDetails()
	{
	thewindow = window.open('https://www.thawte.com/cgi/server/certdetails.exe?code=USFPCS2', 'anew', config='height=400,width=450,toolbar=no,menubar=no,scrollbars=yes,resizable=no,location=no,directories=no,status=yes');
	} 
	
if (window.name.indexOf("app") > -1){
	if (!window.opener.closed){
		window.opener.location.replace('<% = Application.Value("strSSLWebRoot")%>UserAdmin/login.asp');
	}
	window.close();
}

</script>

<center>
<!--
<font face=arial size=3 color=red size=3><b>
The Online Office is currently down due to software upgrades.<br> Please try back around 9am today (9-30-2003).</b></font>
<br>-->
<table cellpadding="3" border="0" cellspacing="0" ID="Table1">
	<tr>
		<td>
			<% = mstrLoginMsg %>
		</td>
	</tr>
	<!--	<tr>		<td align="left">			<font face=tahoma size=-1>				<b>New Member?<br></b>				<b><a href="<%=Application("strWebRoot")%>UserAdmin/register.asp">Click here to register for FREE!</a></b>			</font>		</td>	</tr>	-->
	<tr>
		<td align="left">&nbsp;</td>
	</tr>
	<tr>		
		<td>
			<img src="../images/fpcsLogo.gif">
		</td>
	</tr>
	<tr>
		<td align="left" class="yellowHeader">
			&nbsp;<b>FPCS Online Office Vers. 3.0</b>			
		</td>
	</tr>
	<tr align="left">
		<td>
			<table cellpadding="1" cellspacing="1" ID="Table2">
				<tr>
					<td colspan="2">
						<font size="-1" face="tahoma"><b>Please sign in.</b></font>
					</td>
				</tr>
				<form action="<% = Application.Value("strSSLWebRoot") %>UserAdmin/login.asp" method="post" name="frmLogin" onSubmit="return false;" ID="Form1">
				<input type="hidden" name="strURL" value="<%=strURL%>" ID="Hidden1">
				<tr valign="middle">
					<td>
						<font size="-1" face="tahoma"><b>User ID:</b></font>
					</td>
					<td>
						<input type="text" size="20" name="szUserID" value="<%=strUserID%>" ID="Text1">
					</td>
				</tr>
				<tr valign="middle">
					<td>
						<font size="-1" face="tahoma"><b>Password:</b></font>
					</td>
					<td>
						<input type="password" size="20" name="szPassword" ID="Password1">
					</td>
				</tr>
				<%
					set oFunc = GetObject("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
					dim dtCurYear
					dtCurYear = oFunc.SchoolYear 
					set oFunc = nothing						 
				 %>
				<!-- <input type=hidden name=intSchool_Year value="<% = dtCurYear %>" ID="Hidden2">				-->
				<tr valign="middle">
					<td>
						<font size="-1" face="tahoma"><b>School Year:</b></font>
					</td>
					<td>
						<select name="intSchool_Year">
							<%
							for i = (application.Value("dtYearAppStarted")) to (dtCurYear +1) 							
								if i = dtCurYear then 
									strSelected = " selected "
								else
									strSelected = ""
								end if 
								Response.Write "<option value=""" & i & """" & strSelected & ">" & right(i-1,2) & "' - " & right(i,2) & "'" & chr(13)
							next					 
							%>
						</select>
					</td>
				</tr>
				<tr>
					<td colspan="2" bgcolor="f0f0f0" align="right">
						<input type="hidden" name="hLogin" value="True" ID="Hidden3">
						<input type="submit" value="Sign In" class="btSmallGray" id="cmdLogin" name="cmdLogin" onClick="jfCheckValue(this.form);">
					</td>
				</tr>
			</table>
			<br><BR>
			<% if request("IsVendor") <> "" then %>
			<span class="svError">PLEASE NOTE: Vendors, if you already have a user name and password log in above.<br>
			<B>If you are new to FPCS and would like to apply to become a vendor within our school click <a href="<% = Application.Value("strSSLWebRoot") %>/forms/vis/vendoradmin.asp?xsuggestvendor=true">HERE</a>.</b></span>
			<% else %>
			<!--
			JD 041311 Comment out the 04-05 link 
			<span class="svError">PLEASE NOTE: <a href="<% = Application("strURL") %>/APP2/UserAdmin/login.asp" class="svError">To log into the 04' - 05' School Year please click here.</a></span>-->
			<% end if %>
			<!--<table cellspacing=0 border=1 bordercolor=marroon width=250>
				<tr>
					<td>
						<table>
							<tr>
								<Td>
									<font size="-1" face="tahoma"><b>SYSTEM STATUS</b></font>
								</td>
							</tr>
							<tr>
								<td>
									<font size="-1" face="tahoma">
									The FPCS Online System is being upgraded 
									and should be back online around 8am
									Friday Feb 27, 2004. 
									</font>
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>	-->
		</td>
	</tr>
</table><br><br>
<a href="javascript:OpenCertDetails()">
<img SRC="<% = Application.Value("strImageRoot") %>stamp.gif" BORDER="0" ALT="Click here for more details"></a>
</center>
</body>
</html>
<%
End Sub


Sub subCheckLogin
'*******************************************
'Name:		subCheckLogin (procedure)
'Purpose:	Validates the supplied credintials
'
'Author:	Bryan K Mofley
'			3Shapes is Scott Bacon, Bryan Mofley, Guy Mofley
'Date:      27-Nov-2001
'*******************************************
dim oCrypto		'wsc object
dim oFunc		'wsc object
dim strEncPwd
dim rsValidate		'ADO RecordSet - Used to validate UserID
dim strSQLvalidate	'SQL for rsValidate	

	set oCrypto = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/Crypto.wsc"))
		'encrypt password for database compare
		'oCrypto.Key = "something"	'actual key is not shown here
		oCrypto.Text = strPassword
		Call oCrypto.Encypttext
		strEncPwd = oCrypto.EncryptedText
	set oCrypto = nothing

	set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
	call oFunc.OpenCN()
	
	' RESET SESSION VARIABLES
	oFunc.ResetSelectSessionVariables()
	Session.Contents("intStudent_ID") = ""
	Session.Contents("intFamily_id") = ""
	Session.Contents("strFamily_Name") = ""
	Session.Contents("intGuardian_ID") = ""
	Session.Contents("instruct_id") = ""
	session.Contents("intVendor_ID") = ""
	session.Contents("szVendor_Name") = ""
	session.Contents("szVendor_Email") = ""
	
	set rsValidate = Server.CreateObject("ADODB.RecordSet")
	rsValidate.CursorLocation = 3 'adUseClient 

	'bkm 20-jun-02 added tblUsers.blnActive = 1:  only allow active users to log in
	strSQLvalidate =	"SELECT tblUsers.szUser_ID, tblUsers.szName_First, tblUsers.szName_Last, " & _
							" tblUsers.szEmail, tblUsers.blnActive, tblRoles.szRole_CD, tascInstr_User.intInstructor_ID " & _
							"FROM tascUserRoles INNER JOIN " & _
							"tblRoles ON tascUserRoles.szRole_CD = tblRoles.szRole_CD RIGHT OUTER JOIN " & _
							"tblUsers ON tascUserRoles.szUser_ID = tblUsers.szUser_ID " & _
							" LEFT OUTER JOIN tascInstr_User ON " & _
							"tblUsers.szUser_ID = tascInstr_User.szUser_ID " & _
							"WHERE tblUsers.blnActive = 1 AND tblUsers.szUser_ID = '" & trim(strUserID) & _
							"' AND szPassword = '" & strEncPwd & "'"
				
	rsValidate.Open strSQLvalidate, oFunc.FPCScnn
	''''WARNING This code does not yet handle multiple roles for a single user!!!!!
	if not rsValidate.BOF and not rsValidate.EOF then

	
		Session.Contents("bolUserLoggedIn") = True
		Session.Contents("strName") = trim(rsValidate("szName_First"))
		Session.Contents("strFullName") = trim(rsValidate("szName_First")) & " " & trim(rsValidate("szName_Last"))
		Session.Contents("strUserID") = trim(strUserID)
		Session.Contents("strEmail") = trim(rsValidate("szEmail"))
		Session.Contents("strRole") = trim(rsValidate("szRole_CD"))
		Session.Contents("blnActive") = cbool(rsValidate("blnActive"))
		
		if Session.Contents("strRole") = "TEACHER" then
			Session.Contents("instruct_id") = trim(rsValidate("intInstructor_ID"))
			rsValidate.Close
			sql = "select intStudent_ID from tblEnroll_Info " & _
				  " where intSponsor_Teacher_ID = " & Session.Contents("instruct_id") & _
				  " and sintSchool_Year = " & Request.Form("intSchool_Year")
			rsValidate.Open sql, oFunc.FPCScnn
			' This list is used to ensure that this teacher can only access 
			' their own students and not try to bypass the system and get into
			' other accounts
			if rsValidate.RecordCount > 0 then
				do while not rsValidate.EOF
					session.Contents("student_list") = session.Contents("student_list") & "~" & rsValidate(0) & "~"
					rsValidate.MoveNext
				loop
			else
				session.Contents("student_list") = ""
			end if						
		end if 
		Session.Contents("intSchool_Year") = Request.Form("intSchool_Year")
	
		if Session.Contents("strRole") = "GUARD" then
			sqlGetFamilies = "SELECT f.intFamily_ID, f.szFamily_Name, gu.intGuardian_ID " & _
							 "FROM tascFAM_GUARD fg, tblGUARDIAN g, " & _
							 "   tascGUARD_USERS gu, tblUsers u, tblFAMILY f " & _
							 "WHERE u.szUser_ID = '" &  Session.Contents("strUserID") & "' AND " & _
							 "   u.szUser_ID = gu.szUser_ID AND " & _
							 "   gu.intGuardian_ID = g.intGuardian_id AND " & _
							 "   g.intGuardian_id = fg.intGuardian_id AND " & _
							 "   fg.intFamily_id = f.intFamily_id "
							 
			set rsFamilies = server.CreateObject("ADODB.RECORDSET")
			rsFamilies.CursorLocation = 3
			rsFamilies.Open sqlGetFamilies,oFunc.FPCScnn
			
			' NEED TO PUT CODE IN HERE TO  DEAL WITH MULTIPLE FAMILES.  PROBLAY REDIRECT TO A
			' SCREEN THAT ASKS WHICH FAMILY THEY WOULD LIKE TO TWORK WITH.
			if rsFamilies.RecordCount > 0 then
				Session.Contents("intFamily_id") = rsFamilies("intFamily_id")
				Session.Contents("strFamily_Name") = rsFamilies("szFamily_Name")
				Session.Contents("intGuardian_ID") = rsFamilies("intGuardian_ID")				
				sql = "select intStudent_ID from tblStudent where intFamily_ID = " & rsFamilies("intFamily_id")
				rsFamilies.Close
				rsFamilies.Open sql, oFunc.FPCScnn
				
				' This list is used to ensure that this guardian can only access 
				' their own students and not try to bypass the system and get into
				' other accounts
				if rsFamilies.RecordCount > 0 then
					do while not rsFamilies.EOF
						session.Contents("student_list") = session.Contents("student_list") & "~" & rsFamilies(0) & "~"
						rsFamilies.MoveNext
					loop
				else
					session.Contents("student_list") = ""
				end if
			end if 
			
			rsFamilies.Close
			set rsFamilies= nothing
			
			'bkm - need to finish this - idea is to store all of the guardian's associated student_ID's in
			'a session array or string to limit them to these specific students.  Currently, they can type any ID 
			'in the URL
			dim strSQLGuardStudent
			dim rsGuardStudent
			strSQLGuardStudent = "SELECT s.intSTUDENT_ID " & _
					"FROM tascGUARD_USERS gu INNER JOIN " & _
					"tascFAM_GUARD fg ON gu.intGUARDIAN_ID = fg.intGUARDIAN_ID INNER JOIN " & _
					"tblSTUDENT s ON fg.intFamily_ID = s.intFamily_ID " & _
					"WHERE (gu.szUser_ID = '" & Session.Contents("strUserID") & "')"
		end if 
	    'JD:Invalidate login if VENDOR is not approved	
		if Session.Contents("strRole") = "VENDOR" then
			set rsVend = server.CreateObject("ADODB.RECORDSET")
			rsVend.CursorLocation = 3
			'sql = "SELECT     v.intVendor_ID, v.szVendor_Name, v.szVendor_Email " & _ 
			'		"FROM         tblVendors v INNER JOIN " & _ 
			'		"                      tascVendor_User vu ON v.intVendor_ID = vu.intVendor_ID " & _ 
			'		"WHERE     (vu.szUser_ID = '" & Session.Contents("strUserID") & "') "
			sql = "SELECT v.intVendor_ID  " &_
                    ", v.szVendor_Name  " &_
                    ", v.szVendor_Email  " &_
                    ", (SELECT     TOP 1 upper(szVendor_Status_CD)  " &_
	                "    FROM          tblVendor_Status vs  " &_
	                "    WHERE      vs.intVendor_ID = v.intVendor_ID AND vs.intSchool_Year <=  " & Request.Form("intSchool_Year") &_
	                "    ORDER BY intSchool_Year DESC,intVendor_Status_ID DESC) AS szVendor_Status_CD  " &_
                    "FROM tblVendors v  " &_
                    "INNER JOIN tascVendor_User vu  " &_
                    "ON v.intVendor_ID = vu.intVendor_ID " &_
                    "WHERE vu.szUser_ID = '" & Session.Contents("strUserID") & "'"

            

			rsVend.Open sql, oFunc.FpcsCnn
			
			if rsVend.RecordCount > 0 then
			    if rsVend("szVendor_Status_CD") <> "APPR" then
			        Session.Contents("bolUserLoggedIn") = false
			        mstrLoginMsg = "<TABLE border=1 borderColor=#666699 cellPadding=10 cellSpacing=0 width='95%'>" & _
	                "<TBODY><TR><TD align=middle bgColor=#666699 vAlign=center width=40>" & _
	                "<IMG border=0 src='" & Application("strImageRoot") & "vbExclamation.gif' VALIGN='bottom'></TD>" & _
	                "<TD bgColor=#ffffff vAlign=top><font size=2 face=tahoma color=red><b>" & _
	                "Invalid UserName/Password</b></font></TD></TR></TBODY></TABLE>" & _
	                "<BR><font face='verdana' size=-1><B><A href='" & Application("strAdminRoot") & "EmailPassword.asp?szUserID=" & trim(strUserID) & "'>" & _
	                "EMail My FPCS Password to Me</A></B><BR>" & _
	                "You may have supplied a valid FPCS ID<BR> " & _
	                "but have forgotten your password. <BR></font>"
	                subShowLogin
	                response.End
			    end if
				session.Contents("intVendor_ID") = rsVend("intVendor_ID")
				session.Contents("szVendor_Name") = rsVend("szVendor_Name")
				session.Contents("szVendor_Email") = rsVend("szVendor_Email")
			end if
			rsVend.Close
			set rsVend = nothing
		end if
		'JD
		
		call vbsForceAction
		'the "Session.Contents" line below is currently not used - leaving it here for an idea I had
		'Session.Contents("simpleOnLoad") = "window.open('" & Session.Contents("strURL") & "', 'app', config='toolbar=no,menubar=no,scrollbars=yes,resizable=yes,location=no,directories=no,status=yes');"
		'response.Write Application("strSSLWebRoot") & "userAdmin/launch.asp"
		'response.End
		if session.Contents("strRole") = "VENDOR" then
			session.Contents("strUrl") = Application.Value("strSSLWebRoot") & "vendorHome.asp"		
		end if
		Response.Redirect(Application("strSSLWebRoot") & "userAdmin/launch.asp")		
		
		'Response.Redirect(strURL)		
	else
		mstrLoginMsg = "<TABLE border=1 borderColor=#666699 cellPadding=10 cellSpacing=0 width='95%'>" & _
		"<TBODY><TR><TD align=middle bgColor=#666699 vAlign=center width=40>" & _
		"<IMG border=0 src='" & Application("strImageRoot") & "vbExclamation.gif' VALIGN='bottom'></TD>" & _
		"<TD bgColor=#ffffff vAlign=top><font size=2 face=tahoma color=red><b>" & _
		"Invalid UserName/Password</b></font></TD></TR></TBODY></TABLE>" & _
		"<BR><font face='verdana' size=-1><B><A href='" & Application("strAdminRoot") & "EmailPassword.asp?szUserID=" & trim(strUserID) & "'>" & _
		"EMail My FPCS Password to Me</A></B><BR>" & _
		"You may have supplied a valid FPCS ID<BR> " & _
		"but have forgotten your password. <BR></font>"
		subShowLogin
	end if
	rsValidate.Close
	set rsValidate = nothing
	call oFunc.CloseCN()
	set oFunc = nothing
End Sub

Sub vbsForceAction
	' This sub checks to see if any forced action is required by the user.
	' If so strURL is set to the script requiring action.
	dim rsCheckAction
	dim sqlCheckAction
	dim fa				'Forced Action looping variable
	dim strURLPeice		'URL fragment that will be added to strAction_URL
	
	set oFunc2 = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
	call oFunc2.OpenCN()
	
	set rsCheckAction = server.CreateObject("ADODB.RECORDSET")
	rsCheckAction.CursorLocation = 3
	
	sqlCheckAction = "select fa.intAction_ID, fa.szAction_URL, ua.intUser_Action_ID " & _
					 "from tascUsers_Action ua, tblForce_Action fa " & _
					 "where ua.szUser_ID = '" & Session.Contents("strUserID") & _
					 "' and ua.intAction_ID = fa.intAction_ID " & _
					 "order by ua.intOrder_ID " 
					 
	rsCheckAction.Open sqlCheckAction, oFunc2.FPCScnn
	
	if rsCheckAction.RecordCount > 0 then
		if Session.Contents("strRole") = "TEACHER" then
			strURLPeice = "&intInstructor_id=" & Session.Contents("instruct_id")
		end if 
		
		redim arActions(rsCheckAction.RecordCount,2)
	
		for fa = 0 to rsCheckAction.RecordCount - 1			
			arActions(fa,0) = rsCheckAction("intAction_ID")
			arActions(fa,1) = rsCheckAction("szAction_URL") & strURLPeice & "&intCount=" & fa
			arActions(fa,2) = rsCheckAction("intUser_Action_ID")
			rsCheckAction.MoveNext
		next
	
		Session.Contents("arActions") = arActions
		Session.Contents("bolActionNeeded") = true
		strURL = Application("strSSLWebRoot") & arActions(0,1)
	else
		Session.Contents("bolActionNeeded") = false
	end if
	
	rsCheckAction.Close
	set rsCheckAction = nothing
	call oFunc2.CloseCN()
	set oFunc2 = nothing
End Sub
%>

