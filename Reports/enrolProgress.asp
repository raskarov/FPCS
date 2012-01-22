<%@ Language=VBScript %>
<%
dim intStudent_id
dim sql
dim bolShow
dim strStudentName
dim strGuardInfo

Session.Value("strSimpleTitle") = "Enrollment Progres Report"
Session.Value("strLastUpdate") = "26 Aug 2002"
Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")
%>
<table width=100%>
	<tr>	
		<Td class=yellowHeader>
				&nbsp;<b>Enrollment Progress Report</b>&nbsp;&nbsp;
				(Listed students are incomplete. 'X' refers to action not taken)&nbsp;&nbsp; 
				<input type=button value="Home" onClick="window.location.href='<%=Application("strSSLWebRoot")%>';" id=btSmallGray>
		</td>
	</tr>
</table>
<table>
	<tr>
		<td class=gray>
			<b>Student Name</b>
		</td>
		<td class=gray>
			<b>Guradians</b>
		</td>
		<td class=gray>
			<b>Change Password</b>
		</td>
		<td class=gray>
			<b>Family Profile</b>
		</td>
		<td class=gray>
			<b>Re-Enroll Questionaire</b>
		</td>
	</tr>
<%
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '' Get Student Name 
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

sql = "SELECT tblSTUDENT.intStudent_ID,tblSTUDENT.szLast_Name + ' ' + tblSTUDENT.szFIRST_NAME studentName, " & _
		"    tblFAMILY.szHome_Phone, tblGUARDIAN.szFIRST_NAME + ' ' + tblGUARDIAN.szLAST_NAME guardName, " & _
		"    tblGUARDIAN.szBUSINESS_PHONE, " & _
		"    tascGUARD_USERS.szUser_ID " & _
		"FROM tblSTUDENT INNER JOIN" & _
		"    tblStudent_States ON " & _
		"    tblSTUDENT.intSTUDENT_ID = tblStudent_States.intStudent_id INNER" & _
		"     JOIN" & _
		"    tblFAMILY ON " & _
		"    tblSTUDENT.intFamily_ID = tblFAMILY.intFamily_ID INNER JOIN" & _
		"    tascFAM_GUARD ON " & _
		"    tblFAMILY.intFamily_ID = tascFAM_GUARD.intFamily_ID INNER JOIN" & _
		"    tblGUARDIAN ON " & _
		"    tascFAM_GUARD.intGUARDIAN_ID = tblGUARDIAN.intGUARDIAN_ID" & _
		"     INNER JOIN" & _
		"    tascGUARD_USERS ON " & _
		"    tblGUARDIAN.intGUARDIAN_ID = tascGUARD_USERS.intGUARDIAN_ID " & _
		"WHERE tblStudent_States.intReEnroll_State  IN (" & Application.Contents("ActiveEnrollList") & ") AND " & _
		"    (tblStudent_States.intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _
		"ORDER BY tblSTUDENT.szLAST_NAME, " & _
		"    tblSTUDENT.szFIRST_NAME"
		
set rs = server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3
rs.Open sql, oFunc.FPCScnn

set rsCheck = server.CreateObject("ADODB.Recordset")
rsCheck.CursorLocation = 3

' Initalize variables
intStudent_id = rs("intStudent_ID")
bolShow = true
strStudentName = ""
strGuardInfo = ""
bolPassword = false
bolFamProfile = false
bolShortForm = false
bolEnroll = false
		
do while not rs.EOF
	if cint(intStudent_ID) <> cint(rs("intStudent_ID")) then
		if bolShow = true then
			call vbfPrintResults
		end if 
		bolShow = true
		intStudent_ID = rs("intStudent_ID")
		strStudentName = ""
		strGuardInfo = ""
		bolPassword = false
		bolFamProfile = false
		bolShortForm = false
		bolEnroll = false
	end if 
	
	if bolShow = true then
		sql = "SELECT tblForce_Action.intAction_ID " & _
				"FROM tascUsers_Action INNER JOIN " & _
				"    tblForce_Action ON  " & _
				"    tascUsers_Action.intAction_ID = tblForce_Action.intAction_ID  " & _
				"WHERE (tascUsers_Action.szUser_ID = '" & rs("szUser_ID") & "') " & _
				"ORDER BY tblForce_Action.intAction_ID " 
		rsCheck.Open sql,oFunc.FPCScnn
	
		if rsCheck.RecordCount < 1 then
			bolShow = false
		else
			strStudentName = rs("studentName")
			strGuardInfo = strGuardInfo & rs("guardName") & " (hm:" & rs("szHome_Phone") & _
						   ", wk:" & rs("szBUSINESS_PHONE") & ")<BR>"
			
			do while not rsCheck.EOF
				select case rsCheck("intAction_ID")
					case "2" 
						bolPassword = true
					case "3"
						bolFamProfile = true
					case "4" 
						bolShortForm = true
					case "5"
						bolEnroll = true
				end select
				rsCheck.MoveNext
			loop		
		end if
		rsCheck.Close
	end if 
	rs.MoveNext
loop
rs.Close
set rs = nothing
function vbfPrintResults
%>
	<tr>
		<td class=gray>
			<% = strStudentName %>
		</td>
		<td class=gray>
			<% = left(strGuardInfo,len(strGuardInfo)-4) %>
		</td>
		<td class=gray align=center>
			<% if bolPassword then Response.Write "X" %>
		</td>
		<td class=gray align=center>
			<% if bolFamProfile then Response.Write "X" %>
		</td>
		<td class=gray align=center>
			<% if bolEnroll then Response.Write "X" %>
		</td>
	</tr>

<%
end function 
%>
</table>
<%
call oFunc.CloseCN()
set oFunc = nothing
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")
%>
