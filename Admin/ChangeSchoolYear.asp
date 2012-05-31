<%@ Language=VBScript %>
<%
if Request.Form("intSchool_Year") <> "" then
	session.Value("intSchool_Year") = Request.Form("intSchool_Year")
	if ucase(Session.Contents("strRole")) <> "ADMIN" then
		set oFunc = GetObject("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
		call oFunc.OpenCN()
		set rsValidate = server.CreateObject("ADODB.RECORDSET")
		rsValidate.CursorLocation = 3
			
		if Session.Contents("strRole") = "TEACHER" then
				
				sql = "select intStudent_ID from tblEnroll_Info " & _
					" where intSponsor_Teacher_ID = " & Session.Contents("instruct_id") & _
					" and sintSchool_Year = " & session.Value("intSchool_Year")
				rsValidate.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
				' This list is used to ensure that this teacher can only access 
				' their own students and not try to bypass the system and get into
				' other accounts
				session.Contents("student_list") = ""
				
				if rsValidate.RecordCount > 0 then
					do while not rsValidate.EOF
						session.Contents("student_list") = session.Contents("student_list") & "~" & rsValidate(0) & "~"
						rsValidate.MoveNext
					loop
				else
					session.Contents("student_list") = ""
				end if						
			end if 
			
		elseif ucase(Session.Contents("strRole")) = "GUARD" then
			sqlGetFamilies = "SELECT f.intFamily_ID, f.szFamily_Name, gu.intGuardian_ID " & _
							"FROM tascFAM_GUARD fg, tblGUARDIAN g, " & _
							"   tascGUARD_USERS gu, tblUsers u, tblFAMILY f " & _
							"WHERE u.szUser_ID = '" &  Session.Contents("strUserID") & "' AND " & _
							"   u.szUser_ID = gu.szUser_ID AND " & _
							"   gu.intGuardian_ID = g.intGuardian_id AND " & _
							"   g.intGuardian_id = fg.intGuardian_id AND " & _
							"   fg.intFamily_id = f.intFamily_id "
				
			rsValidate.Open sqlGetFamilies, Application("cnnFPCS")'oFunc.FPCScnn
			
			' NEED TO PUT CODE IN HERE TO  DEAL WITH MULTIPLE FAMILES.  PROBALLY REDIRECT TO A
			' SCREEN THAT ASKS WHICH FAMILY THEY WOULD LIKE TO TWORK WITH.
			if rsValidate.RecordCount > 0 then
				Session.Contents("intFamily_id") = rsValidate("intFamily_id")
				Session.Contents("strFamily_Name") = rsValidate("szFamily_Name")
				Session.Contents("intGuardian_ID") = rsValidate("intGuardian_ID")				
				sql = "select intStudent_ID from tblStudent where intFamily_ID = " & rsValidate("intFamily_id")
				rsValidate.Close
				rsValidate.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
				
				' This list is used to ensure that this guardian can only access 
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
			rsValidate.Close
			set rsValidate = nothing
			call oFunc.CloseCN()
			set oFunc = nothing
		end if
	Server.Execute(Application.Value("strWebRoot") & "default.asp")
else
	Server.Execute(Application.Value("strWebRoot") & "Includes/simpleheader.asp")
%>
<center>
<table cellpadding="3" border="0" cellspacing="0">
	<tr>		
		<td>
			<img src="../images/fpcsLogo.gif" >
		</td>
	</tr>
	<tr>
		<td align="left"  class="yellowHeader">
			&nbsp;<b>FPCS Online Office</b>			
		</td>
	</tr>
	<tr align="left">
		<td>
			<table cellpadding="1" cellspacing="1">
				<tr>
					<td colspan="2">
						<font size="-1" face=tahoma><b>Change School Year.</b></font>
					</td>
				</tr>
				<form action="ChangeSchoolYear.ASP" method=post>
				<tr valign="middle">
					<td>
						<font size="-1" face=tahoma><b>School Year:</b></font>
					</td>
					<td>
						<select name="intSchool_Year">
						 <%
							dim dtCurYear
							dim strSelected
							dtCurYear = datePart("yyyy",now())
							if ucase(session.Contents("strRole")) = "ADMIN" then
								intNum = dtCurYear +2
							else
								intNum = dtCurYear +1
							end if 
							
							for i = (application.Value("dtYearAppStarted")) to (intNum +1) 							
								if i = cint(session.Value("intSchool_Year")) then 
									strSelected = " selected "
								else
									strSelected = ""
								end if 
								Response.Write "<option value=""" & i & """" & strSelected & ">" & i & chr(13)
							next					 
						 %>
						 </select>
					</td>
				</tr>
				<tr>
					<td colspan=2 bgcolor=f0f0f0 align=right>
						<input type="submit" value="submit" >
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</center>
</body>
</html>
<%
end if
%>
