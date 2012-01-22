<%@ Language=VBScript %>
<%
dim sql
dim intCount	'Number of Teacher 
dim strInfo		'contains Teacher info for mouse over display
dim sqlTeacher

if session.Contents("strRole") <> "ADMIN" and  session.Contents("strRole") <> "TEACHER" then
	response.Write "<h1>Improper Request</h1>"
	response.End
end if

dim oFunc		'wsc object
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

set rsReport = server.CreateObject("ADODB.RECORDSET")
rsReport.CursorLocation = 3

sql= "SELECT i.szLast_Name,i.szFirst_Name,p.curPay_Rate,i.intInstructor_id, " & _
	 "i.szEmail,i.szHome_Phone,i.szBusiness_Phone " & _
	 "from tblInstructor i, tblInstructor_Pay_Data p " & _
	 "where i.intInstructor_ID = p.intInstructor_ID " & _
	 "and p.dtEffective_End is null order by i.szLast_Name"
               
rsReport.Open sql,oFunc.FPCScnn
intCount = rsReport.RecordCount

' Print the HTML header
Session.Value("strTitle") = "Teacher Per Diem Report"
Session.Value("strLastUpdate") = "03 June 2002"
Server.Execute(Application.Value("strWebRoot") & "includes/Header.asp")

if intCount > 0 then
%>
<table width=100%>
	<tr>	
		<Td class=yellowHeader>
				&nbsp;<b>Teacher Per Diems</b>
		</td>
	</tr>
	<tr>
		<td bgcolor=f7f7f7>
			<table>
				<tr>	
					<Td class=svplain10 colspan=2>
						&nbsp;<B>Total Number of Teachers:</b> <% = intCount %>&nbsp;
					</td>
				</tr>
				<tr>	
					<Td class=gray>
						&nbsp;<B>Teacher</b>&nbsp;
					</td>
					<Td class=gray>
						&nbsp;<B>Per Diem</b>&nbsp;
					</td>
					<Td class=gray>
						&nbsp;<B>Base Hrly</b>&nbsp;
					</td>
					<Td class=gray>
						&nbsp;<B>W/Benefits</b>&nbsp;
					</td>
				</tr>
				<% 
					do while not rsReport.EOF
				%>
				<tr>			
				<%
						
					strInfo = "Home Phone: " & 	rsReport("szHome_Phone") & _
						      " Work Phone: " & rsReport("szBusiness_Phone")
				%>
					<td class=gray>
						<span title="<% = strInfo %>">
						<a href="javascript:" onClick="jfGetProfile('<%=rsReport("intInstructor_id")%>');">
						<% = rsReport("szLast_Name") & ", " & rsReport("szFirst_Name") & ": " & rsReport("intInstructor_id")%></a>&nbsp;</span>
					</td>
					<td class=gray>
						&nbsp;$<% = rsReport("curPay_Rate") %>
					</td>
					<td class=gray>
						&nbsp;$<% Response.Write formatNumber(cDBL(rsReport("curPay_Rate"))/cdbl(7.5),2)%>
					</td>
					<td class=gray>
						&nbsp;$<%
								arRate = oFunc.InstructorCosts(rsReport("intInstructor_ID"))							
								if isArray(arRate) then
									Response.Write formatNumber(cdbl(arRate(9)),2) 
								end if 								
							   %>
					</td>
				</tr>
				<% 						
						rsReport.MoveNext
				   loop 
				%>
			</table>
			<input type=button value="< Back" onCLick="window.location.href='<%=Application.Value("strWebRoot")%>';"  id=button1 name=button1>
		</td>
	</tr>
</table>
<span class=svplain>&nbsp;Mouse over the instructor or guardian for more information.</span>
<script language=javascript>
	function jfGetProfile(id){
		var winProfile;
		winProfile = window.open("../forms/Teachers/addTeacher.asp?bolWin=True&intInstructor_ID="+id,"winProfile","width=800,height=550,scrollbars=yes");
		winProfile.focus();
		winProfile.moveTo(0,0);
	}
</script>
</body>
</html>
<%
	set rsGuardian = nothing
	rsReport.Close
else
%>
<form id=form1 name=form1>
<table width=100% height=100%>
	<tr>
		<Td align=center valign=middle>
			<table>
				<tr>
					<Td class=svplain10>
						No students are currently enrolled in this class<br><BR>
						<center>
						<input type=button value="< Back" onCLick="window.location.href='<%=strPath%>';"  id=button2 name=button2>
						</center>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<%
end if
set rsReport = nothing

call oFunc.CloseCN()
set oFunc = nothing
' Conclude html 
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")
%>