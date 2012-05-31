<%@ Language=VBScript %>
<%
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()
Server.Execute(Application.Value("strWebRoot") & "includes/header.asp")

%>
<form name=main method=post action="enrollmentReportByStudent.asp" >
<table width="100%">
	<tr>
		<Td class="yellowHeader">
			&nbsp;<b>FPCS Re-Enrollment Matrix Report</b>
		</Td>
	</tr>
	<tr>
		<td bgcolor="f7f7f7">
			<table>
				<tr>
					<td class=Gray>
						<nobr>&nbsp;<b>Select Student(s)</b></nobr><br>
					</td>
					<td width=100%>
						<select name=strStudents multiple size=5>
							<option value="all">ALL STUDENTS
						<%
							dim sqlStudent
							sqlStudent = "Select intStudent_ID,szLast_Name + ',' + szFirst_Name as Name " & _
											 "from tblStudent order by szLast_Name"
							Response.Write oFunc.MakeListSQL(sqlStudent,intStudent_ID,Name,strTeachers)	
						%>
						</select>
					</td>
				</tr>
				<tr>
					<td class="gray">
							<nobr>&nbsp;<b>School Year:</b></nobr>
					</td>
					<Td>						
						<select name="intSchool_Year">
							<%
								= oFunc.MakeYearList(2,1,session.Contents("intSchool_Year"))
							%>
						</select>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>	
&nbsp;<input type=submit value="submit">
</form>
<br>
<%
dim strStudents
dim sql 
dim strWhere
dim strStateRow

strStudents = Request.Form("strStudents")

if strStudents <> "" then
%>
<table>
	<tr>
		<td class=Gray>
			<b>Student Name</b>
		</td>
		<td class=Gray>
			<b>Re-Enrollment Sent</b>
		</td>
		<td class=Gray>
			<b>Re-Enrollment Received</b>
		</td>
		<td class=Gray>
			<b>Re-Enroll (Yes)</b>
		</td>
		<td class=Gray>
			<b>1st Phone Call</b>
		</td>
		<td class=Gray>
			<b>2nd Phone Call</b>
		</td>
		<td class=Gray>
			<b>Letter of Dismissal</b>
		</td>
		<td class=Gray>
			<b>Exit Candidate</b>
		</td>
		<td class=Gray>
			<b>Value</b>
		</td>
	</tr>
<%
	'If we have a list of students coming from the header we break them up and 
	'dynamicaly create the where clause
	strStudents = replace(strStudents," ","")
	if instr(1,strStudents,",") > 0 then
		arStudentList = split(strStudents,",")
		strWhere = " where intStudent_ID = '" & arStudentList(0) & "' "
		
		for w = 1 to ubound(arStudentList)
			strWhere = strWhere & " or intStudent_ID = '" & arStudentList(w) & "' "
		next 
		
	elseif strStudents = "all" then
		strWhere = ""
	else
		' Only a single selection was made
		strWhere = " where intStudent_ID = '" & strStudents & "' "
	end if
	

	set rsReport = server.CreateObject("ADODB.RECORDSET")
	rsReport.CursorLocation = 3
	
	sql= "SELECT szLast_Name,szFirst_Name,intStudent_ID " & _
		 "from tblStudent " & _
		 strWhere & _
		 "order by szLast_Name "              
	rsReport.Open sql,Application("cnnFPCS")'oFunc.FPCScnn
	
	set rsStates = server.CreateObject("ADODB.RECORDSET")
	rsStates.CursorLocation = 3
	
	do while not rsReport.EOF 
		sql = "select intReEnroll_State " & _
			  "from tblStudent_States " & _
			  "where intStudent_ID = " & rsReport("intStudent_ID") & _
			  " and intSchool_Year = " & request("intSchool_Year")
		rsStates.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
		
		' Reset State Values 		
		str1 = ""
		str2 = ""
		str4 = ""
		str8 = ""
		str16 = ""	
		str32 = ""			
		str64 = ""				
		
		if rsStates.RecordCount > 0 then
			select case rsStates("intReEnroll_State")
				case "1"
					str1 = "X"
				case "7"
					str1 = "X"
					str2 = "X" 
					str4 = "X"
				case "9"
					str1 = "X"
					str8 = "X"
				case "15"
					str1 = "X"
					str2 = "X" 
					str4 = "X"
					str8 = "X"
				case "25"
					str1 = "X"
					str8 = "X"
					str16 = "X"
				case "31"
					str1 = "X"
					str2 = "X" 
					str4 = "X"
					str8 = "X"
					str16 = "X"
				case "57"
					str1 = "X"
					str8 = "X"
					str16 = "X"
					str32 = "X"
				case "67"
					str1 = "X"
					str2 = "X" 
					str64 = "X"
				case "75"
					str1 = "X"
					str2 = "X" 
					str8 = "X"
					str64 = "X"
				case "91"
					str1 = "X"
					str2 = "X" 
					str8 = "X"
					str16 = "X"				
					str64 = "X"
				case "121"
					str1 = "X"
					str8 = "X"
					str16 = "X"	
					str32 = "X"			
					str64 = "X"
			end select
				
%>
		<Tr>
			<td class=gray> 
				<% = rsReport("szLast_Name") & "," &  rsReport("szFirst_Name") %>
			</td>
			<td class=gray align=center>  
				<% = str1 %>
			</td>
			<td class=gray align=center>  
				<% = str2 %>
			</td>
			<td class=gray align=center>  
				<% = str4 %>
			</td>
			<td class=gray align=center>  
				<% = str8 %>
			</td>
			<td class=gray align=center>  
				<% = str16 %>
			</td>
			<td class=gray align=center>  
				<% = str32 %>
			</td>
			<td class=gray align=center>  
				<% = str64 %>
			</td>
			<td class=gray align=center>  
				<% = rsStates("intReEnroll_State") %>
			</td>
		</tr>	
<%				
		end if
			
		rsReport.MoveNext
		rsStates.Close
	loop
	rsReport.Close
	set rsReport = nothing
	set rsStates = nothing
	Response.Write "</table>"
end if 
set oFunc = nothing
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")
%>
