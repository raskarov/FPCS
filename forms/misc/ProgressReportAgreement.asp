<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		ProgressReportAgreement.asp
'Purpose:	Form that stores digital signature of guardians for 
'Date:		10 Aug 2005
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim oFunc				'windows scripting component generalized functions
dim oStudent			' student object providing various properties
dim intFamily_ID			' must be defined in order for script to fully run
dim sql

Session.Value("strTitle")		= "Progress Report Agreement"
Session.Value("strLastUpdate")	= "29 July 2005"

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

if request("intFamily_ID") <> "" then
	intFamily_ID = request("intFamily_ID")
elseif request("intStudent_ID") <> "" then
	set oStudent = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/StudentBudgetInfo.wsc"))
	'oStudent.PopulateStudentFunding oFunc.FpcsCnn,request("intStudent_ID"), session.Contents("intSchool_Year")
	oStudent.PopulateStudentFunding Application("cnnFPCS"),request("intStudent_ID"), session.Contents("intSchool_Year")
	intFamily_ID = oStudent.FamilyId
	set oStudent = nothing
else
	response.Write "<h1>Page Improperly Called</h1>"
	response.End
end if 

if request.Form("ids") <> "" then
	call SaveSignature()
end if

Server.Execute(Application.Value("strWebRoot") & "includes/Header.asp")

sql = "SELECT     s.intSTUDENT_ID, s.szFIRST_NAME, s.szLAST_NAME, tblENROLL_INFO.bolProgress_Agreement, " & _ 
		" tblENROLL_INFO.dtProgress_Signed, tblENROLL_INFO.szUser_Progress_Signed, tblENROLL_INFO.intEnroll_INFO_ID " & _
		"FROM tblSTUDENT s INNER JOIN " & _ 
		"	tblStudent_States ON s.intSTUDENT_ID = tblStudent_States.intStudent_id INNER JOIN " & _ 
		"	tblENROLL_INFO ON s.intSTUDENT_ID = tblENROLL_INFO.intSTUDENT_ID " & _ 
		"WHERE	(s.intFamily_ID = " & intFamily_ID & ") AND (tblStudent_States.intSchool_Year = " & session.Contents("intSchool_Year") & ") AND (tblStudent_States.intReEnroll_State IN (" & Application.Contents("ActiveEnrollList") & ") ) AND  " & _ 
		"	(tblENROLL_INFO.sintSCHOOL_YEAR = " & session.Contents("intSchool_Year") & ") " & _
		"ORDER BY s.szLAST_NAME, s.szFIRST_NAME " 
		
set rs = server.CreateObject("ADODB.RECORDSET")
rs.CursorLocation = 3
rs.Open sql, Application("cnnFPCS")'oFunc.FpcsCnn

%>
	
<form action="ProgressReportAgreement.asp" method=post name="main" ID="Form1">
<input type="hidden" name="intFamily_ID" value="<% = intFamily_ID %>" ID="Hidden1">
<input type="hidden" name="intStudent_ID" value="<% = request("intStudent_ID") %>" ID="Hidden2">
<table width=100% ID="Table1">
	<tr>	
		<Td class=yellowHeader >
				&nbsp;<b>Progress Report Agreement</b> 
		</td>
	</tr>
	<tr>
		<td class="svplain10"><br>
		Parent/sponsor teacher communication is a critical component of our 
		charter school.  Quarterly progress reports and <b>monthly face-to-face</b> meetings for each student  
		should be completed throughout the school year.  Semester grades are to be submitted by <b>December 16th for 1st Semester 
		and May 17th for all students for 2nd semester and final end of the year grades</b>.
<BR><BR>
*Note:  Students eligible for sports in high school and middle school must meet different deadlines for quarterly grade submissions.

<%
if rs.RecordCount > 0 then
%>
<BR><BR>Guardian signatures are required ...<br><br>
			<table ID="Table2">
<%
do while not rs.EOF
%>
				<tr class="tablecell">
					<td class="svplain8">
						I am aware that progress report forms for <b><% = rs("szFirst_Name") & " " & rs("szLast_Name") %></b>
						<BR>will be available as a link within the Student Online System.<br><br>
					</td>
					<td valign="top">
						|<br>|
					</td>
					<% if rs("dtProgress_Signed") & "" <> "" then %>
					<td class="svplain8" valign="top">
						<i>(Signed on <% = rs("dtProgress_Signed") %> By
						<% = rs("szUser_Progress_Signed") %>)</i>
					</td>
					<% elseif ucase(session.Contents("strRole")) = "GUARD" then %>
					<td class="svplain8" valign="top">
					<input type="checkbox" name="<% = rs("intEnroll_INFO_ID") %>Selection" value="1" ID="Checkbox1">
					(checking this box acts as your signature)
					</td>
					<%
					strIds = strIds & rs("intEnroll_INFO_ID") & ","
					else
					%>
					<td class="svplain8">
						(guardian has not yet signed for this student)
					</td>
					<% end if %>
				</tr>
<%
	rs.MoveNext
loop

	if len(strIds) > 0 then
		strIds = left(strIds,len(strIds)-1)	
%>
				<tr>
					<td colspan="4">
						<BR>
						<input type="submit" class="navsave" value="Sign and Save" ID="Submit1" NAME="Submit1">
						<input type="hidden" name="Ids" value="<% = strIds%>" ID="Hidden3">
					</td>
				</tr>
<%
	end if


%>			
			</table>
		</td>
	</tr>
<% end if %>
</table>
</form>		
<input type=button value="return to packet" class="btSmallgray" onClick="window.location.href='<%=Application.Value("strWebRoot")%>forms/packet/packet.asp?intStudent_ID=<%= request("intStudent_ID") %>'" ID="Button1" NAME="Button1">
<%
Server.Execute(Application.Value("strWebRoot") & "Includes/footer.asp")
call oFunc.CloseCn()
set oFunc = nothing


sub SaveSignature
	dim update
	arIds = split(request("ids"),",")
	
	if isArray(arIds) then		
		for i = 0 to ubound(arIds)
			if request(arIds(i) & "Selection") & "" <> "" then
				update = "update tblEnroll_Info set bolProgress_Agreement = " & request(arIds(i) & "Selection") & _
						 ", dtProgress_Signed = CURRENT_TIMESTAMP, szUser_Progress_Signed = '" & session.Contents("strUserID") & "' " & _
						 " where intEnroll_Info_ID = " & arIds(i)
				oFunc.ExecuteCn(update)
			end if
		next
	end if
end sub
%>