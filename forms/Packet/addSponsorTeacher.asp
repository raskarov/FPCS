<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		addSponsorTeacher.asp
'Purpose:	Form for adding/inserting a sponsor teacher that will aid a 
'				parent in creating students ciriculum.			
'Date:		9-04-01
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim oFunc
dim intStudent_ID
dim strStudentName
dim intSponsor_Teacher_ID
dim bolGotoPacket

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()
if Session.Contents("intStudent_ID") <> "" then
	intStudent_ID = Session.Contents("intStudent_ID") 
elseif request("intStudent_ID") <> "" then
	intStudent_ID = request("intStudent_ID")
	arStudentInfo = oFunc.StudentInfo(intStudent_ID,8)
	strStudentName = arStudentInfo(2)
	intSponsor_Teacher_ID = arStudentInfo(7)	
	bolGotoPacket = true
end if

if request("intSponsor_Teacher_ID") <> "" then	
	dim intSchoolYear
	dim update
		
	intSchool_Year = session.Contents("intSchool_Year")
	
	set rsSponsor = server.CreateObject("ADODB.RECORDSET")
	rsSponsor.CursorLocation = 3
	sqlSponsor = "select intSponsor_Teacher_ID from tblEnroll_Info " & _
				 " where intStudent_ID = " & intStudent_ID & _
				 " and sintSchool_Year = " & intSchool_Year
	rsSponsor.Open sqlSponsor, Application("cnnFPCS")'oFunc.FPCScnn	
	
	if rsSponsor.RecordCount < 1 then
	' Insert a new Sponsor Teacher Record
		dim insert
		
		insert = "insert into tblEnroll_Info(intStudent_ID,sintSchool_Year," & _
				"intSponsor_Teacher_ID, szUSER_CREATE) " & _
				"values ( " & _
				intStudent_ID & "," & _
				"'" & intSchool_Year & "'," & _
				"'" & oFunc.EscapeTick(request("intSponsor_Teacher_ID")) & "', '" & Session.Value("strUserID")	& "')"
		oFunc.ExecuteCN(insert)
	else
		update = "update tblEnroll_Info " & _
				 "set intSponsor_Teacher_ID = " & request("intSponsor_Teacher_ID") & _
				 ", szUser_Modify = '" & oFunc.EscapeTick(session.Contents("strUserID")) & "'," & _
				 "dtModify = '" & now() & "' " & _
				 "where intStudent_ID = " & intStudent_ID & _
				 " AND sintSchool_Year = " & intSchool_Year
		oFunc.ExecuteCN(update)
	end if 
	rsSponsor.Close
	set rsSponsor = nothing
	Session("intSponsorTeacherID" & intStudent_ID & intSchool_Year) = request("intSponsor_Teacher_ID")
	if bolGotoPacket then
		response.Redirect(Application.Value("strWebRoot") & "forms/packet/packet.asp?intStudent_ID=" & intStudent_ID)
	else
		Response.Redirect(Application.Value("strWebRoot") & "forms/ilp/ILP1.asp?intStudent_ID="&intStudent_ID&"&intShort_ILP_ID=" & request("intShort_ILP_ID"))
	end if
end if 

Session.Value("strTitle") = "Add a Sponsor Teacher"
Session.Value("strLastUpdate") = "12 May 2002"
Server.Execute(Application.Value("strWebRoot") & "includes/header.asp")
%>
<form action="addSponsorTeacher.asp" name=main method=post>
<input type=hidden name=intStudent_ID value="<% = request("intStudent_ID") %>">
<table >
	<tr>	
		<Td class=yellowHeader >
				&nbsp;<b>Select a Sponsor Teacher for <% = Session.Value("strStudentName") & strStudentName%></b>
		</td>
	</tr>
	<tr>
		<td bgcolor=f7f7f7>
			<table>
				<tr>	
					<Td colspan=2>
						<font class=svplain11>
							<b>Note: Before an ILP can be filled out a FPCS Sponsor Teacher must be
							selected for your student.</b>
						</font>
						<font class=svplain>
						</font>
					</td>
				</tr>
			</table>
			<table>
				<tr>	
					<Td>
						<font class=svplain11>
							&nbsp;Please select a Sponsor Teacher: &nbsp;
						</font>
						<font class=svplain>
						</font>
					</td>					
					<td>
						<select name="intSponsor_Teacher_ID">
							<option value="">
							<option value="1">UAA
							<option value="2">ASD School
						<%
							dim sqlInstructor
							sqlInstructor = "Select distinct i.intInstructor_ID,i.szLast_Name + ',' + i.szFirst_Name as Name " & _
											 "from tblInstructor i " & _
											 "inner join tblInstructor_Pay_Data pd on pd.intInstructor_ID = i.intInstructor_ID " & _
											 "and (current_timestamp between pd.dtEffective_start and pd.dtEffective_End or pd.dtEffective_End is null) " & _
											 "where i.intPay_Type_ID <> 5 and pd.bolActive = 1 " & _
											 "order by Name, i.intInstructor_ID"											 
							Response.Write oFunc.MakeListSQL(sqlInstructor,"intInstructor_ID","Name",intSponsor_Teacher_ID)												 
						%>
						</select>	
					</td>
				</tr>	
				<tr>
					<td colspan=2>
						<input type=hidden name="intShort_ILP_ID" value="<%=request("intShort_ILP_ID")%>">
						<input type=submit value="SUBMIT" class="btSmallGray">
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
oFunc.CloseCN
set oFunc = nothing
%>