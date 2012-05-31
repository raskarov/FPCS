<%@ Language=VBScript %>
<%
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()
Server.Execute(Application.Value("strWebRoot") & "includes/header.asp")

dim intReEnroll_State
dim sql 

intReEnroll_State = Request.Form("intReEnroll_State")
%>
<form name=main method=post action="enrollmentReportByCase.asp" >
<table width="100%">
	<tr>
		<Td class="yellowHeader">
			&nbsp;<b>FPCS Re-Enrollment by Case Report</b>
		</Td>
	</tr>
	<tr>
		<td bgcolor="f7f7f7">
			<table>
				<tr>
					<td class=Gray>
						<nobr>&nbsp;<b>Select Case:</b></nobr>
					</td>
					<td width=100%>
						<select name="intReEnroll_State" >
							<option value="">
						<%
							dim sqlStates
							sqlStates = "Select intReEnroll_State, strCase from trefReEnroll_States order by strCase"
							Response.Write oFunc.MakeListSQL(sqlStates,"intReEnroll_State","strCase",intReEnroll_State)	
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
&nbsp;<input type=submit value="submit" >
</form>
<br>
<%
if intReEnroll_State <> "" then
	'If we have a list of students coming from the header we break them up and 
	'dynamicaly create the where clause
	
	set rsReport = server.CreateObject("ADODB.RECORDSET")
	rsReport.CursorLocation = 3    
	
	sql = "SELECT s.szLAST_NAME, s.szFIRST_NAME, s.intSTUDENT_ID, f.szHome_Phone, f.szFamily_Name + ': ' + f.szDesc AS FamilyName " & _
			"FROM tblSTUDENT s INNER JOIN " & _
			" tblStudent_States ss ON s.intSTUDENT_ID = ss.intStudent_id INNER JOIN " & _
			" tblFAMILY f ON s.intFamily_ID = f.intFamily_ID " & _
			"WHERE     (ss.intReEnroll_State = " & intReEnroll_State & ") " & _
			" and ss.intSchool_Year = " & session.Contents("intSchool_Year") & _
			" ORDER BY s.szLAST_NAME"	    
	rsReport.Open sql,Application("cnnFPCS")'oFunc.FPCScnn
	
	
%>
<table>
	<tr>
		<td colspan=3 class=gray>
			Total Records Returned: <% = rsReport.RecordCount %>
		</td>
	</tr>
	<tr>
		<td class=Gray>
			&nbsp;<b>Student Name</b>
		</td>
		<td class=Gray>
			&nbsp;<b>Family</b>
		</td>
		<td class=Gray>
			&nbsp;<b>Contact #</b>
		</td>
	</tr>
<%
	
	do while not rsReport.EOF 
				
%>
		<Tr>
			<td class=gray> 
				&nbsp;<% = rsReport("szLast_Name") & "," &  rsReport("szFirst_Name") %>
			</td>
			<td class=gray >  
				&nbsp;<% = rsReport("FamilyName") %>
			</td>
			<td class=gray >  
				&nbsp;<% = rsReport("szHome_Phone")  %>
			</td>
		</tr>	
<%				
		rsReport.MoveNext
	loop
	rsReport.Close
	set rsReport = nothing
	Response.Write "</table>"
end if 
set oFunc = nothing
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")
%>
