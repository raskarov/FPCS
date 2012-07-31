<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		TestingAgreement.asp
'Purpose:	Form that stores digital signature of guardians for 
'Date:		29 July 2005
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

sql = "SELECT     s.intSTUDENT_ID, s.szFIRST_NAME, s.szLAST_NAME, tblENROLL_INFO.bolASD_Testing, " & _ 
		" tblENROLL_INFO.dtASD_Signed, tblENROLL_INFO.szUser_ASD_Signed, tblENROLL_INFO.intEnroll_INFO_ID " & _
		"FROM tblSTUDENT s INNER JOIN " & _ 
		"	tblStudent_States ON s.intSTUDENT_ID = tblStudent_States.intStudent_id INNER JOIN " & _ 
		"	tblENROLL_INFO ON s.intSTUDENT_ID = tblENROLL_INFO.intSTUDENT_ID " & _ 
		"WHERE	(s.intFamily_ID = " & intFamily_ID & ") AND (tblStudent_States.intSchool_Year = " & session.Contents("intSchool_Year") & ") AND (tblStudent_States.intReEnroll_State IN (" & Application.Contents("ActiveEnrollList") & ")) AND  " & _ 
		"	(tblENROLL_INFO.sintSCHOOL_YEAR = " & session.Contents("intSchool_Year") & ") " & _
		"ORDER BY s.szLAST_NAME, s.szFIRST_NAME " 
		
set rs = server.CreateObject("ADODB.RECORDSET")
rs.CursorLocation = 3
rs.Open sql, Application("cnnFPCS")'oFunc.FpcsCnn

%>
	
<form action="TestingAgreement.asp" method=post name="main">
<input type="hidden" name="intFamily_ID" value="<% = intFamily_ID %>">
<input type="hidden" name="intStudent_ID" value="<% = request("intStudent_ID") %>">
<table width=100%>
	<tr>	
		<Td class=yellowHeader >
				&nbsp;<b>ASD Testing Agreement</b> 
		</td>
	</tr>
	<tr>
		<td class="svplain10"><br>
		Family Partnership Charter School is required to test all enrolled students in grades 3-9 in mandatory state testing to comply with No Child Left Behind legislation.
<BR><BR>FPCS has only reached Adequate Yearly Progress (AYP) in test participation and graduation rates twice in its 10 year history. 
The school is also required to participate in a universal screening for a reading assessment for students in Grades K-6. We need your help!
<BR><BR><b>*All 10th Grade Students need to take and pass the HSGQE (High School Graduation Qualifying Exam).
<BR><BR>*All 11th Grade Students need take the WorkKeys testing to qualify for the Alaska Performance Scholarships.
<BR><BR>*Please agree to this even if your student is in grades K-2 or in 12th Grade and will not need to test this year.</b>
<%
if rs.RecordCount > 0 then
%>
<BR><BR>By selecting a check box below, I am signing 
acknowledgment that...<br><br>
			<table>
<%
do while not rs.EOF 
%>
				<tr>
					<td class="svplain8">
						<b><% = rs("szFirst_Name") & " " & rs("szLast_Name") %></b>
					</td>
					<% if rs("dtASD_Signed") & "" <> "" then %>
					<td class="svplain8">
						<% if rs("bolASD_Testing") then 
								response.Write " will participate "	
						   'else
							'	response.Write " will not participate "
						   end if
						%>
						in all mandatory testing. <i>(Signed on <% = rs("dtASD_Signed") %> By
						<% = rs("szUser_ASD_Signed") %>)</i>
					</td>
					<% elseif ucase(session.Contents("strRole")) = "GUARD" then %>
					<td class="svplain8">
					(<input type="checkbox" name="<% = rs("intEnroll_INFO_ID") %>Selection" value="1" id='chk<%= rs("intEnroll_INFO_ID") %>'>
					<label for='chk<%= rs("intEnroll_INFO_ID") %>'>Will Participate</label>
                    <%If False Then %> | 
					<input type="checkbox" name="<% = rs("intEnroll_INFO_ID") %>Selection" value="0" ID="Checkbox1">
					Will Not Participate
                    <%End If %> 
                    ) in all mandatory testing. The date you "sign" will be recorded.
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
						<input type="submit" class="navsave" value="Sign and Save">
						<input type="hidden" name="Ids" value="<% = strIds%>">
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
<span class="svplain8"><a href="https://docs.google.com/viewer?a=v&pid=explorer&chrome=true&srcid=0B5w5Wuf_btjhMjRlOTg3YTctZTVmMS00NDU2LWJkZWMtOTExYTdjZmEwYTFk&hl=en_US" target="_blank">View Testing Dates</a></span>&nbsp;
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
				update = "update tblEnroll_Info set bolASD_Testing = " & request(arIds(i) & "Selection") & _
						 ", dtASD_Signed = CURRENT_TIMESTAMP, szUser_ASD_Signed = '" & session.Contents("strUserID") & "' " & _
						 " where intEnroll_Info_ID = " & arIds(i)
				oFunc.ExecuteCn(update)
			end if
		next
	end if
end sub
%>