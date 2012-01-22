<%@ Language=VBScript %>
<%
if session.Contents("strRole") <> "ADMIN" and  session.Contents("strRole") <> "TEACHER" then
	response.Write "<h1>Improper Request</h1>"
	response.End
end if
server.ScriptTimeout = 10000
dim oFunc
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

dim sql
dim intCount	'Number of Teacher 
dim strInfo		'contains Teacher info for mouse over display
dim sqlTeacher
dim intILPCount	'Number of ILP's per Class 
dim intTotalPlanning
dim intTotalInstruction

set rsReport = server.CreateObject("ADODB.RECORDSET")
rsReport.CursorLocation = 3

sql= "SELECT i.szLast_Name,i.szFirst_Name,i.intInstructor_id, " & _
	 "i.szEmail,i.szHome_Phone,i.szBusiness_Phone " & _
	 "from tblInstructor i " & _
	 "order by szLast_Name "
               
rsReport.Open sql,oFunc.FPCScnn
intCount = rsReport.RecordCount

Session.Value("strTitle") = "Teacher Hour Breakdown"
Session.Value("strLastUpdate") = "17 June 2002"
Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")

if intCount > 0 then
	set rsClasses = server.CreateObject("ADODB.RECORDSET")
	rsClasses.CursorLocation = 3
	set rsILP = server.CreateObject("ADODB.RECORDSET")
	rsILP.CursorLocation = 3
%>
<table width=100%>
	<tr>	
		<Td class=yellowHeader>
				&nbsp;<b>Teacher Hours Detailed Breakdown</b>
		</td>
	</tr>
	<tr>
		<td bgcolor=f7f7f7>
			<table>
				<tr>
					<Td colspan=3>
						<input type=button value="< Back" onCLick="window.location.href='<%=Application.Value("strWebRoot")%>';"  id=btSmallGray name=button3>
					</td>
				</tr>
				<tr>	
					<Td class=svplain10 colspan=3>
						&nbsp;<B>Total Number of Teachers:</b> <% = intCount %>&nbsp;
					</td>
				</tr>
		<% 
			do while not rsReport.EOF
				strInfo = "Home Phone: " & 	rsReport("szHome_Phone") & _
						  " Work Phone: " & rsReport("szBusiness_Phone")
		%>				
				<tr>
					<Td colspan=3>
						<br>
					</td>
				</tr>
				<tr>						
					<Td class=gray12 colspan=3>
						&nbsp;<B>Teacher: </b>&nbsp;
						<span title="<% = strInfo %>">
						<a href="javascript:" onClick="jfGetProfile('<%=rsReport("intInstructor_id")%>');">
						<% = rsReport("szLast_Name") & " " & rsReport("szFirst_Name")%></a>&nbsp;</span>
					</td>					
				</tr>
				<%
				
				sql = "select intClass_ID, szClass_Name,decHours_Student,decHours_Planning,decHours_Student " & _
					  "from tblClasses " & _
					  "where intInstructor_ID = " & rsReport("intInstructor_ID") & _
					  " and (c.intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _
					  " order by szClass_Name"
				rsClasses.Open sql, oFUnc.FPCScnn
				
				if rsClasses.RecordCount > 0 then
				
				do while not rsClasses.EOF		
				%>
				<tr>
					<td >
						&nbsp;&nbsp;&nbsp;
					</td>
					<td colspan=2>
						<table>
							<tr>
								<td class=gray colspan=4>
									&nbsp;<b>Class Name:</b> <% = rsClasses("szClass_Name") %>
								</td>	
							</tr>
				<%					
					sql = "select s.intStudent_id, s.szFirst_Name,s.szLast_Name, " & _
						  "i.intILP_ID " & _
						  "from tblStudent s, tblILP i " &  _
						  "where i.intClass_ID = " & rsClasses("intClass_ID") & _
						  " and i.intStudent_id = s.intStudent_ID " & _
						  " order by s.szLast_Name " 
					rsILP.Open sql,oFUnc.FPCScnn
					intILPCount = rsILP.RecordCount
					if intILPCount > 0 then
				%>
							<tr>
								<td>
										&nbsp;&nbsp;&nbsp;
								</td>			
								<Td>
									<table border=1 cellpadding=0 cellspacing=0>
										<tr>		
											<td class=gray>
												&nbsp;Student Name&nbsp;
											</td>
											<td class=gray>
												&nbsp;Planning Hrs&nbsp;
											</td>
											<td class=gray>
												&nbsp;Instruction Hrs&nbsp;
											</td>
											<td class=gray>
												&nbsp;Total Hrs&nbsp;
											</td>
										</tr>								
				<%					
					do while not rsILP.EOF				
				%>
										<tr>		
											<td class=svplain10>
												&nbsp;<% = rsILP("szLast_Name") & "," & rsILP("szFirst_Name") %>
											</td>
											<td class=svplain10>
												&nbsp;<% Response.Write formatNumber(cDBL(rsClasses("decHours_Planning"))/cdbl(intILPCount),2)%>
											</td>
											<td class=svplain10>
												&nbsp;<% Response.Write formatNumber(cDBL(rsClasses("decHours_Student"))/cdbl(intILPCount),2) %>
											</td>
											<td class=svplain10>
												&nbsp;<% Response.Write formatNumber((cDBL(rsClasses("decHours_Student"))/cdbl(intILPCount)) + (cDBL(rsClasses("decHours_Planning"))/cdbl(intILPCount)),2)  %>
											</td>
										</tr>
				<% 						
						rsILP.MoveNext						
				   loop 
				   rsILP.Close
				%>
										<tr>		
											<td class=svplain10>
												&nbsp;<b>Totals</b>
											</td>
											<td class=svplain10>
												&nbsp;<b><%  = rsClasses("decHours_Planning")%>&nbsp;</b>
											</td>
											<td class=svplain10>
												&nbsp;<b><% =  rsClasses("decHours_Student") %>&nbsp;</b>
											</td>
											<td class=svplain10>
												&nbsp;<b><% =  cdbl(rsClasses("decHours_Student")) + cdbl(rsClasses("decHours_Planning")) %>&nbsp;</b>
											</td>
										</tr>
									</table>
								</td>
							</tr>							
				<%
						intTotalPlanning = cdbl(intTotalPlanning) + cdbl(rsClasses("decHours_Planning"))
						intTotalInstruction = cdbl(intTotalInstruction) + cdbl(rsClasses("decHours_Student"))
					else
				%>
						
							<tr>
								<td>
										&nbsp;&nbsp;&nbsp;
								</td>	
								<td class=svplain10>
									<B> No Students Enrolled in this Class.</b>
								</td>
							</tr>
				<%					
						rsILP.Close						
					end if 'intILPCount > 0 
				%>								
						</table>
					</td>
				</tr>						
				
				<%										
					rsClasses.MoveNext
				loop
				arRate = oFunc.InstructorCosts(rsReport("intInstructor_ID"))
				if isArray(arRate) then
					dblPayRate = formatNumber(cdbl(arRate(9)),2)	
				end if 						
				%>
				<tr>					
					<td colspan=3>
						<br>
						<table cellpadding=0 cellspacing=0 border=1> 
							<tr>
								<td class=svplain10>
									&nbsp;&nbsp;
								</td>
								<td class=svplain10>
									&nbsp;<b>Total Planning Hrs</b>&nbsp;
								</td>
								<td class=svplain10>
									&nbsp;<b>Total Instruction Hrs</b>&nbsp;
								</td>
								<td class=svplain10>
									&nbsp;<b>Total Hrs</b>&nbsp;
								</td>
							</tr>
							<tr>
								<td class=svplain10>
									&nbsp;<B>Hours</b>&nbsp;
								</td>
								<td class=svplain10>
									&nbsp;<%=intTotalPlanning%>&nbsp;
								</td>
								<td class=svplain10>
									&nbsp;<%=intTotalInstruction%>&nbsp;
								</td>
								<td class=svplain10>
									&nbsp;<%=(intTotalPlanning + intTotalInstruction)%>&nbsp;
								</td>
							</tr>
							<tr>
								<td class=svplain10>
									&nbsp;<B>Pay</b>&nbsp;
								</td>
								<td class=svplain10>
									&nbsp;<%=dblPayRate * intTotalPlanning%>&nbsp;
								</td>
								<td class=svplain10>
									&nbsp;<%=dblPayRate * intTotalInstruction%>&nbsp;
								</td>
								<td class=svplain10>
									&nbsp;<%= dblPayRate * (intTotalPlanning + intTotalInstruction)%>&nbsp;
								</td>
							</tr>							
						</table>
					</td>
				</tr>							
				<%							
			else
			%>
				<tr>						
					<Td class=gray12 colspan=3>
						&nbsp;This Teacher Currently has no Classes.
					</td>					
				</tr>
			<%		
			end if
				rsClasses.Close
				rsReport.MoveNext		
				intTotalPlanning = 0 
				intTotalInstruction = 0 
			loop
			rsReport.Close
			set rsReport = nothing
			set rsClasses = nothing
			set rsILP = nothing
				%>
			</table>			
		</td>
	</tr>
</table>
<input type=button value="< Back" onCLick="window.location.href='<%=Application.Value("strWebRoot")%>';"  id=btSmallGray name=button1>
<span class=svplain>&nbsp;Mouse over the instructor  for more information.</span>
<script language=javascript>
	function jfGetProfile(id){
		var winProfile;
		winProfile = window.open("../forms/Teachers/addTeacher.asp?bolWin=True&intInstructor_ID="+id,"winProfile","width=800,height=550,scrollbars=yes");
		winProfile.focus();
		winProfile.moveTo(0,0);
	}
</script>
<%
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
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")
%>