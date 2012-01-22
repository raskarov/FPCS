<%@ Language=VBScript %>
<%
dim oFunc
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()
dim sqlGetILPs
dim sql
dim intCount
dim strStudent_Name
dim strMessage

' If submit was clicked then we update records
if request("update") <> "" then
	call vbfUpdateDates()
end if 

' Print the HTML header
Session.Value("strTitle") = "Student Enrollment Percentages"
Session.Value("strLastUpdate") = "03 June 2002"
Server.Execute(Application.Value("strWebRoot") & "includes/Header.asp")
%>
<form name="frmStudentHead" method="GET" onsubmit="return false;">
<script language="javascript">
	function jfChangeStudent(form){
		//reloads page with newly selected student
		var strURL = "<% = Application("strWebRoot")%>forms/misc/enrollmentDate.asp?intStudent_ID=" + form.selintStudent_ID.value;
		window.open(strURL, "_self");
	}
</script>

<% if strMessage <> "" then %>
&nbsp;<font class=svPlain11 color=red><b><% = strMessage %></b></font><br>
<% end if %>

<table width=100%>
	<tr>	
		<Td class=yellowHeader>
				&nbsp;<b>Enrollment Dates For: <% = strStudent_Name %></b>&nbsp;&nbsp;&nbsp;
						
						<select name="selintStudent_ID" onchange="jfChangeStudent(this.form);">
							<option value="">
						<%
							'this change was requested by Val/Scott - partialy implemented by bkm 20-Dec-2001
							dim sqlStudent
							sqlStudent = "Select intStudent_ID,szLast_Name + ',' + szFirst_Name as Name " & _
											 "from tblStudent order by szLast_Name"
							Response.Write oFunc.MakeListSQL(sqlStudent,"intStudent_ID","Name",Request.QueryString("intStudent_ID"))												 
						%>
						</select>
		</td>
	</tr>

<%

if Request.QueryString("intStudent_ID") <> "" then
	'Get Name of Student
	set rsGetName = server.CreateObject("ADODB.Recordset")
	rsGetName.CursorLocation = 3
	sql = "select szFirst_Name + ' ' + szLast_name as name " & _
		  "from tblStudent " & _
		  "where intStudent_ID = " & Request.QueryString("intStudent_ID")
	rsGetName.Open sql,oFunc.FPCScnn
	strStudent_Name = rsGetName(0)
	rsGetName.Close
	set rsGetName = nothing

	'Get all ILP's for student
	set rsGetILPs = server.CreateObject("ADODB.Recordset")
	rsGetILPs.CursorLocation = 3
	sqlGetILPs = "select c.intClass_ID, c.szClass_Name,i.intILP_ID,dtStudent_Enrolled  " & _
				 "from tblILP i, tblClasses c " & _
				 "where intStudent_ID = " & Request.QueryString("intStudent_ID") & _
				 " and sintSchool_Year = " & session.Contents("intSchool_Year") & _
				 " and i.intClass_ID = c.intClass_ID order by c.szClass_Name " 		 
	rsGetILPs.Open sqlGetILPs,oFunc.FPCScnn
else
	'close html. nothing else to show until we get a student id
%>
	<tr>
		<Td class=svplain11>
			<b><i>Please select a student.</I></B> 
		</TD>
	</TR>
</table>
</form>
</BODY>
</HTML>
<%
	Response.End
end if

if rsGetILPs.RecordCount > 0 then
%>
	<tr>
		<Td class=svplain11>
			<b><i>View/Edit Existing Student Enrollment Dates.</I></B> <br>			
		</TD>
	</TR>
	<tr>
		<td class=svred8>
			<b>All dates must be in the dd/mm/yyyy format.</b>
		</td>
	</tr>
</form>	
</table>
<table>
<form name=main action=enrollmentDate.asp method=post>
<input type=hidden name=intStudent_id value="<%=Request.QueryString("intStudent_id")%>">
<input type=hidden name=sintSchool_Year value="<%= session.Contents("intSchool_Year")%>">
	<tr>
		<td class=gray>
			<b>Class Name</b>
		</td>
		<td class=gray>
			<b>Enrollment Date</b>
		</td>
	</tr>
<%
	intCount = 0 
	do while not rsGetILPs.EOF
%>
	<tr>
		<td class=gray>			
			<%=rsGetILPs("szClass_Name")%>
		</td>
		<td class=gray>			
			<input type=text name="dtEnrolled<% = intCount%>" value="<%=rsGetILPs("dtStudent_Enrolled")%>">
			<input type=hidden name="intILP_ID<% = intCount%>" value="<%=rsGetILPs("intILP_ID")%>">
		</td>
	</tr>	
<%
		intCount = intCount + 1
		rsGetILPs.MoveNext
	loop

%>
	<tr>
		<td>
			<input type=hidden name=intCount value="<% = (intCount-1)%>">
			<input type=submit value="Save" name="update" id="btSmallGray">
		</td>
	</tr>
</table>
</form>
<%
else
%>
	</table>
	<BR>
	<font class=svplain10><B>No classes found for this student.</b></font><br><br>
<%
end if
rsGetILPs.Close
set rsGetILPs = nothing

call oFunc.CloseCN()
set oFunc = nothing
' Conclude html 
response.Write "</table>"
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")

function vbfUpdateDates()
	dim update
	for i = 0 to request("intCount")
		if request("dtEnrolled"&i) <> "" then
			update = "update tblILP set dtStudent_Enrolled = '" & replace(request("dtEnrolled"&i),"'","") & _
					 "', szUSER_MODIFY = '" & Session.Value("strUserID") & "' where intILP_ID = " & request("intILP_ID"&i) 
			oFunc.ExecuteCN(update)
		end if
	next
	strMessage = "Update Complete"
end function

%>

