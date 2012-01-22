<%@ Language=VBScript %>
<%
dim oFunc
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()
   
dim strTitle
strTitle = "Class Enrollment List"

dim sql
dim intCount	'Number of students in a class
dim strInfo		'contains guardian info for mouse over display
dim sqlTeacher

'JD 052611 include the sponsor teacher and their email

'sql = "SELECT     s.szLAST_NAME + ',' + s.szFIRST_NAME AS Name, c.szClass_Name, i.intStudent_ID, c.intInstructor_ID, c.intGuardian_ID, c.intVendor_ID, " & _
'"                     c.dtClass_Start, c.dtClass_End, c.szDays_Meet_On, c.intMin_Students, c.intMax_Students, f.szFamily_Name, f.szDesc, f.szHome_Phone,  " & _
'"                      f.szEMAIL " & _
'" FROM         tblILP i INNER JOIN " & _
'"                      tblClasses c ON i.intClass_ID = c.intClass_ID INNER JOIN " & _
'"                      tblSTUDENT s ON i.intStudent_ID = s.intSTUDENT_ID LEFT OUTER JOIN " & _
'"                      tblFAMILY f ON s.intFamily_ID = f.intFamily_ID " & _
'" WHERE     (c.intClass_ID = " & Request.QueryString("intClass_ID") & ") " & _
'" ORDER BY Name"     

sql = "SELECT     s.szLAST_NAME + ',' + s.szFIRST_NAME AS Name, c.szClass_Name, i.intStudent_ID, c.intInstructor_ID, c.intGuardian_ID, c.intVendor_ID, " & _
" ins.szFIRST_NAME + ' ' + ins.szLAST_NAME AS SPONSOR_NAME, iNS.szEmail AS SPONSOR_EMAIL,  " & _
"                     c.dtClass_Start, c.dtClass_End, c.szDays_Meet_On, c.intMin_Students, c.intMax_Students, f.szFamily_Name, f.szDesc, f.szHome_Phone,   " & _
"                      f.szEMAIL  " & _
" FROM         tblILP i INNER JOIN  " & _
"                      tblClasses c ON i.intClass_ID = c.intClass_ID INNER JOIN  " & _
"                      tblSTUDENT s ON i.intStudent_ID = s.intSTUDENT_ID LEFT OUTER JOIN  " & _
"                      tblFAMILY f ON s.intFamily_ID = f.intFamily_ID  " & _
"                      LEFT OUTER JOIN  " & _
"                      tblINSTRUCTOR ins RIGHT OUTER JOIN   " & _
"                      tblENROLL_INFO ei ON ins.intINSTRUCTOR_ID = ei.intSponsor_Teacher_ID AND (ei.sintSCHOOL_YEAR = 2011) ON " & _    
"                      s.intSTUDENT_ID = ei.intSTUDENT_ID AND (ei.sintSCHOOL_YEAR = "& Session.Contents("intSchool_Year") & ")" & _
" WHERE     (c.intClass_ID = " & Request.QueryString("intClass_ID") & ") " & _
" ORDER BY Name"     

set rsReport = server.CreateObject("ADODB.RECORDSET")
rsReport.CursorLocation = 3
rsReport.Open sql,oFunc.FPCScnn
intCount = rsReport.RecordCount

if intCount > 0 then
	session.Value("simpleTitle") = "Class Enrollment List"
	Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")
	
	if rsReport("intInstructor_ID") & "" <> "" then
		sqlTeacher = "select szFirst_Name,szLast_Name,szEmail,szHome_Phone," & _
			  "szBusiness_Phone " & _
			  "from tblInstructor " & _
			  "where intInstructor_ID = " & rsReport("intInstructor_ID")
	elseif rsReport("intGuardian_ID") & "" <> "" then
		sqlTeacher = "select szFirst_Name,szLast_Name,szEmail,szHome_Phone," & _
			  "szBusiness_Phone " & _
			  "from tblGuardian " & _
			  "where intGuardian_ID = " & rsReport("intGuardian_ID")
	elseif rsReport("intVendor_ID") & "" <> "" then
		sqlTeacher = "select '',szVendor_Name,szVendor_Email,'',szVendor_Phone + ' or ' + szVendor_Phone as phone " & _
			  "from tblVendors " & _
			  "where intVendor_ID = " & rsReport("intVendor_ID")	
	end if

	set rsTInfo = server.CreateObject("ADODB.RECORDSET")
	rsTInfo.CursorLocation = 3
	rsTInfo.Open sqlTeacher,oFunc.FPCScnn 

	strTeacherInfo = "Home Phone: " & 	rsTInfo(3) & _
					 " Work Phone: " & rsTInfo(4)
	strClassTeacher = "<a href='mailto:" &  rsTInfo(2) & "'>" & rsTInfo(0) & " " & rsTInfo(1) & "</a>" 

	rsTInfo.Close
	set rsTInfo = nothing			 
	set rsGuardian = server.CreateObject("ADODB.RECORDSET")
	rsGuardian.CursorLocation = 3
	session.Value("simpleTitle") = strTitle
	Server.Execute(Application.Value("strWebRoot") & "Includes/simpleHeader.asp")		
%>
<script language=javascript>
	function jfManageEmail(pEmail){
		var sList = document.main.strEmailList;
		
		if (sList.value.indexOf(";"+pEmail+";") == -1 ) {
			// Email is not in list so add it
			sList.value = sList.value + pEmail + ";";
		}else{
			// Email is in list so remove it
			var re = new RegExp(pEmail + ";",'gi');
			sList.value = sList.value.replace(re,'');
		}
	}
	
	function jfOpenMailClient(){
		var sList = document.main.strEmailList;
		window.location.href = "mailto:" + sList.value;
	}
</script>
<form name="main">
<input type="hidden" name="strEmailList" value=";" ID="Hidden1">

<table width=100%>
	<tr>	
		<Td class=svbold11>
				<hr width="100%" size=1>
				&nbsp;<b><% = rsReport("szClass_Name") %></b>
				<hr width="100%" size=1>
		</td>
	</tr>
	<tr>
		<td bgcolor=f7f7f7>
			<table cellpadding="2">
				<tr>	
					<Td class="gray">
						<b>Instructed By </b>
					</td>
					<td class="TableCell">
					<span title="<% = strTeacherInfo%>"><% = strClassTeacher %></span>
					</td>
				</tr>
				<tr>	
					<Td class="gray">
						<b>Start & End Dates </b>
					</td>
					<td class="TableCell">
					<% = rsReport("dtClass_Start") & " - " & rsReport("dtClass_End") %>
					</td>
				</tr>
				<tr>	
					<Td class="gray">
						<b>Meets On </b>
					</td>
					<td class="TableCell">
					<% = rsReport("szDays_Meet_On") %>
					</td>
				</tr>
				<tr>	
					<Td class="gray">
						<b>Min # of Students </b>
					</td>
					<td class="TableCell">
					<% = rsReport("intMin_Students") %>
					</td>
				</tr>
				<tr>	
					<Td class="gray">
						<b>Max # of Students </b>
					</td>
					<td class="TableCell">
					<% = rsReport("intMax_Students") %>
					</td>
				</tr>
				<tr>	
					<Td class="gray">
						&nbsp;<B># of Students in Class</b>
					</td>
					<td class="TableCell">
						<% = intCount %>
					</td>
				</tr>
			</table>
			<br>
			<table>
				<tr>	
					<Td class=gray>
						&nbsp;<B>STUDENT'S ENROLLED</b>&nbsp;
					</td>
					<Td class=gray>
						&nbsp;<B>STUDENT'S GUARDIANS</b>&nbsp;
					</td>
					<% if not oFunc.IsGuardian then %>
					<td class=gray>
						<input type=button value="email checked" class="btSmallWhite" onclick="jfOpenMailClient();" ID="Button12" NAME="Button12">
					</td>
					<% end if %>
					<Td class=gray>
						&nbsp;<B>HOME PHONE</b>&nbsp;
					</td>
					<%'JD 052611 student sponsor %>
					<td class="gray">
						&nbsp;<B>STUDENT'S SPONSOR</b>&nbsp;
					</td>
					
				</tr>
				<% 
					do while not rsReport.EOF
				%>
				<tr>
					<Td class=gray>
						<% = rsReport("Name") %>
					</td>
					<td class=gray>
						<a href="mailto:<% = rsReport("szEMAIL") %>"><% = rsReport("szDesc") & " " & rsReport("szFamily_Name") %></a>									
					</td>	
					<td align="center"  class=gray>
					<% 
					if instr(1,sMailList,rsReport("szEMAIL") & ";") < 1 then 		
						if not oFunc.IsGuardian then %>						
							<input type="checkbox" value="<% = rsReport("szEMAIL") %>" onChange="jfManageEmail('<% = rsReport("szEMAIL") %>');" ID="Checkbox1" NAME="Checkbox1">								
						<% end if 
					end if 
					sMailList = sMailList & rsReport("szEMAIL") & ";"	
					%>
					</td>
					<td class=gray>
						<% = rsReport("szHome_Phone") %>			
					</td>
				    <%'JD 052611 student sponsor %>
					<td class="gray">
						<a href="mailto:<% = rsReport("SPONSOR_EMAIL") %>"><% = rsReport("SPONSOR_NAME") %></a>									
					</td>

				</tr>
				<% 
						rsReport.MoveNext
				   loop 
				%>
			</table>
			<input type=button value="Close" onClick="window.opener.focus();window.close();" id=button2 name=button2>			
		</td>
	</tr>
</table>
<span class=svplain>&nbsp;Mouse over the instructor or guardian for more information.</span>
</body>
</html>
<%
	set rsGuardian = nothing
	rsReport.Close
else
%>
<html>
<head>
<title>Students Enrolled</title>
<link rel="stylesheet" href="<% = Application("strSSLWebRoot") %>/CSS/homestyle.css">
</head>
<body background=c0c0c0>
<form id=form1 name=form1>
<table width=100% height=100%>
	<tr>
		<Td align=center valign=middle>
			<table>
				<tr>
					<Td class=svplain10>
						No students are currently enrolled in this class<br><BR>
						<center>
						<input type=button value="Close" onCLick="window.opener.focus();window.close();" id=button1 name=button1>
						</center>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</form>
</body>
</html>
<%
end if
set rsReport = nothing
oFunc.CloseCN
set oFunc = nothing
%>