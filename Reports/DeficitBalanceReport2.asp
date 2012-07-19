<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		DeficitBalanceReport.asp
'Purpose:	Returns Students who have a negative balance
'Date:		03 Jan 2005
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim intStudent_ID 
dim sql
dim mError		'conitains our error messages after validation is complete
dim strDiasbled 
dim strStudentName
dim arInfo
dim oHtml

server.ScriptTimeout = 2200
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

set oHtml = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/htmlFuncs.wsc"))

'Initialize some key variables
if ucase(session.Contents("strRole")) = "ADMIN" then
	intReporting_Period_ID = request("intReporting_Period_ID")
else
	'terminate page since page was improperly called.
	response.Write "<html><body><H1>Page improperly called.</h1></body></html>"
	response.End
end if

if request.Form.Count > 0 then
	' Transfers all of the post http header variables into vbs variables
	' so we can more readily access them
	for each i in request.Form
		execute("dim " & i)
		execute(i & " = """ & replace(replace(replace(request.Form(i),"""","'"),chr(13),""),chr(10),"") & """")
	next 
end if 

if sepDate <> "" then
	if isdate(sepDate) then
		sepDate = cdate(sepDate)
		sepDate = month(sepDate) & "/" & day(sepDate) & "/" & year(sepDate)
	else
		strMessage = "<span style='color:red;' class='svplain8'><B>Count Deadline date is invalid. Please enter a valid date.</b></span><br>"
	end if	
end if 
'prevent page from cacheing
Response.AddHeader "Cache-Control","No-Cache"
Response.Expires = -1

'Print the header
Session.Value("strTitle") = "Student Progress Report"
Session.Value("strLastUpdate") = "08 Dec 2004"
Server.Execute(Application.Value("strWebRoot") & "includes/Header.asp")
start = now()

%>
<form name=main method=post action="DeficitBalanceReport2.asp" ID="Form1">
<input type=hidden name=intStudent_ID value="<% = intStudent_ID %>" ID="Hidden1">
<input type=hidden name="studentList" value="," ID="Hidden2">
<table style="width:100%;" ID="Table3">
	<tr>
		<td class="yellowHeader">
			&nbsp;<b>Student Balance Reports</b>
		</td>
	</tr>	
	<tr>
		<td>
			<table ID="Table1" cellpadding="3">
				<tr>
					<td class='gray' align="center">
						<b>Budget Report Type</b>
					</td>
					<td class='gray' align="center">
						<b>Show Inactive Students</b> 
					</td>
					<td class='gray' align="center">
						<b>Count Deadline</b> 
					</td>
					<td rowspan="2" valign="middle">
						<input type="submit" value="get report" class="btSmallGray" ID="Submit1" NAME="Submit1">
					</td>
				</tr>
				<tr>
					<td>
						<select name='selReportType' ID="Select1">
							<option value="deficit" <% if request("selReportType") = "deficit" then response.Write " Selected " %>>Deficit Balance Report</option>
							<option value="all" <% if request("selReportType") = "all" then response.Write " Selected " %>>Students Balance Report</option>
						</select>
					</td>
					<td class='svplain8' align="center">
						Yes: <input type=checkbox name='bolInactive' value="true" <% if request("bolInactive") & "" <> "" then response.Write " checked " %> ID="Checkbox1">						
					</td>
					<td align="center" class="svplain8">
						<b><% = application.Contents("dtCount_Deadline" & session.Contents("intSchool_Year")) %></b>
						<input type="hidden" name="sepDate" value="<% = application.Contents("dtCount_Deadline" & session.Contents("intSchool_Year")) %>" ID="Text1">
					</td>
					
				</tr>
			</table>
		</td>
	</tr>	
	
<% 
	if strMessage <> "" then
		response.Write strMessage 
	end if
	
	if request("selReportType") <> "" and strMessage = "" then %>
	<tr>
		<td>			
		
<%

	if request("bolInactive") = "" then 
		strShowActive = " AND ss.intReEnroll_State IN (" & Application.Contents("ActiveEnrollList") & ") "
	else
		strShowActive = " AND ss.intReEnroll_State  IN (" & Application.Contents("strEnrollmentList") & ") "
	end if 
				
	sql = "SELECT     s.szLAST_NAME, s.szFIRST_NAME, s.intSTUDENT_ID, i.szFIRST_NAME + ' ' + i.szLAST_NAME AS TEACHERS_NAME, i.szEmail AS TEACHER_EMAIL,  " & _ 
				"                      i.szHOME_PHONE AS TEACHER_PHONE, f.szEMAIL, f.szHome_Phone, f.szDesc, f.szFamily_Name, ss.intReEnroll_State, ss.dtWithdrawn, ss.szGrade, " & _ 
				"                       DM_PERCENT_ENROLLED.TotalCoreHours, DM_PERCENT_ENROLLED.TotalElectiveHours, DM_PERCENT_ENROLLED.TotalHours,  " & _ 
				"                      DM_PERCENT_ENROLLED.CoreCredits, DM_PERCENT_ENROLLED.ElectiveCredits, DM_PERCENT_ENROLLED.ActualEnrolledPercent,  " & _ 
				"                      DM_STUDENT_EXPENSES.TotalGSBudgeted, DM_STUDENT_EXPENSES.TotalGSSpent, DM_STUDENT_EXPENSES.ElectiveGSBudgeted,  " & _ 
				"                      DM_STUDENT_EXPENSES.ElectiveGSSpent, DM_STUDENT_EXPENSES.ElectiveTeacherCost, DM_STUDENT_EXPENSES.TotalElectiveBudget,  " & _ 
				"                      DM_STUDENT_EXPENSES.TotalElectiveSpent, DM_STUDENT_EXPENSES.TotalTeacherCosts, DM_STUDENT_EXPENSES.TotalTeacherHours,  " & _ 
				"                      DM_STUDENT_EXPENSES.Deposits, DM_STUDENT_EXPENSES.Withdrawls, DM_STUDENT_EXPENSES.TotalTransfers,  " & _ 
				"                      tblFunding.curFund_Amount, ei.intPercent_Enrolled_Fpcs " & _ 
				"FROM         tblFunding INNER JOIN " & _ 
				"                      tblStudent_States ss INNER JOIN " & _ 
				"                      tblSTUDENT s ON ss.intStudent_id = s.intSTUDENT_ID ON tblFunding.szGrade = ss.szGrade LEFT OUTER JOIN " & _ 
				"                      DM_STUDENT_EXPENSES ON s.intSTUDENT_ID = DM_STUDENT_EXPENSES.StudentId LEFT OUTER JOIN " & _ 
				"                      DM_PERCENT_ENROLLED ON s.intSTUDENT_ID = DM_PERCENT_ENROLLED.StudentID LEFT OUTER JOIN " & _ 
				"                      tblFAMILY f ON s.intFamily_ID = f.intFamily_ID LEFT OUTER JOIN " & _ 
				"                      tblINSTRUCTOR i RIGHT OUTER JOIN " & _ 
				"                      tblENROLL_INFO ei ON i.intINSTRUCTOR_ID = ei.intSponsor_Teacher_ID AND i.intINSTRUCTOR_ID = ei.intSponsor_Teacher_ID ON  " & _ 
				"                      s.intSTUDENT_ID = ei.intSTUDENT_ID " & _ 
				"WHERE     (ss.intSchool_Year = " & session.Contents("intSchool_Year") & ") " & strShowActive & _ 
				"		    AND (ei.sintSCHOOL_YEAR = " & session.Contents("intSchool_Year") & ") AND (DM_PERCENT_ENROLLED.SchoolYear = " & session.Contents("intSchool_Year") & ") AND  " & _ 
				"                      (DM_STUDENT_EXPENSES.SchoolYear = " & session.Contents("intSchool_Year") & ") AND (tblFunding.intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _ 
				"ORDER BY s.szLAST_NAME, s.szFIRST_NAME "
				
	dim rs 
	set rs = server.CreateObject("ADODB.RECORDSET")
	rs.CursorLocation = 3
	rs.Open sql, Application("cnnFPCS")'oFunc.FPCScnn

	if rs.RecordCount > 0 then			
%>	
		<table  ID="Table2" cellpadding="3">
			
<%		
		intCount = 0 
		do while not rs.EOF
			' Print headers when needed
			if intCount mod 29 = 0 then 
				response.Write vbfTableHeaders("")
				intCount = 1
			end if
			
			fltPlannedFunding = (cdbl(rs("curFund_Amount")) * cdbl((rs("intPercent_Enrolled_Fpcs")/100))) + cdbl(rs("totalTransfers"))
			fltPlannedExpenses = cdbl(rs("TotalTeacherCosts")) + cdbl(rs("TotalGSBudgeted"))
			fltPlannedBalance = fltPlannedFunding - fltPlannedExpenses
			fltActualFunding = (cdbl(rs("curFund_Amount")) * cdbl((rs("ActualEnrolledPercent")/100))) + cdbl(rs("totalTransfers"))  
			fltActualExpenses = cdbl(rs("TotalTeacherCosts")) + cdbl(rs("TotalGSSpent"))
			fltActualBalance = fltActualFunding - fltActualExpenses
			
			if (request("selReportType") = "deficit" and (fltPlannedBalance < -.001 or fltActualBalance < -.001)) or _
			    request("selReportType") = "all" then
			    intCount = intCount + 1
			    if rs("intReEnroll_State") <> "7" and rs("intReEnroll_State") <> "15" and _
						rs("intReEnroll_State") <> "31" then					
					strInactive = "<span style='color:red;'><b>inactive</b></span>&nbsp;"
					strIADate = "<br><b>Date Inactivated: </b>" & rs("dtWithdrawn")
				else
					strInactive = ""
					strIADate = ""
				end if 
				
				bolExclude = false
				if isDate(sepDate) and rs("intReEnroll_State") = "86" then
					if cdate(rs("dtWithdrawn")) < dateadd("d",1,cdate(sepDate)) then
						bolExclude = true
					end if
				end if				
				
				if not bolExclude then 
%>
			<tr>
				<td class="TableCell" >
					<a href="javascript:" onclick="jfViewReport('<% = rs("intStudent_ID")%>');">
					<% response.Write oHtml.ToolTip("<b>" & rs("szLAST_NAME") & ", " & rs("szFIRST_NAME") & "</b>", _
										"<table><tr><td class='svplain8' nowrap><b>Sponsor Teacher:</b> <a href=""mailto:" & rs("TEACHER_EMAIL") & """>" & rs("TEACHERS_NAME") & "</a><BR>" & _
										"<b>Guardians:</b> <a href=""mailto:" & rs("szEMAIL") & """>" & rs("szDesc") & "</a><BR>" & _
										"<b>Family Phone:</b> " & rs("szHome_Phone") & strIADate & "</td></tr></table>", _
										false,"",true,"ToolTip","","",false,false)  %></a>&nbsp;
					<% = strInactive %>
				</td>
				<td class="TableCell"  align="right">
					$<% = formatnumber(fltPlannedFunding,2) %>&nbsp;
				</td>
				<td class="TableCell"  align="right">
					$<% = formatnumber(fltPlannedExpenses,2) %>&nbsp;
				</td>			
				<td class="TableCell"  align="right">
					<% if fltPlannedBalance < -.0001 then%>
					<span style='color:red;'>$<% = formatnumber(fltPlannedBalance,2) %></span>&nbsp;
					<% else %>
						$<% = formatnumber(fltPlannedBalance,2) %>&nbsp;
					<% end if %>
				</td>
				<td class="TableCell"  align="right">
					$<% = formatnumber(fltActualFunding,2) %>&nbsp;
				</td>
				<td class="TableCell"  align="right">
					$<% = formatnumber(fltActualExpenses,2) %>&nbsp;
				</td>
				<td class="TableCell" align="right">
					<% if fltActualBalance < -.0001 then%>
					<span style='color:red;'>$<% = formatnumber(fltActualBalance,2) %></span>&nbsp;
					<% else %>
						$<% = formatnumber(fltActualBalance,2) %>&nbsp;
					<% end if %>
				</td>
			</tr>						
<%																	  					
					
					if (rs("intReEnroll_State") = "86" or rs("intReEnroll_State") = "123")  then
						dblInactiveTotalBudget = dblInactiveTotalBudget + fltPlannedBalance
						dblInactiveTotalActual = dblInactiveTotalActual + fltActualBalance
						dblInactivePlanFund = dblInactivePlanFund + fltPlannedFunding
						dblInactiveActualFund = dblInactiveActualFund + fltActualFunding
						dblInactivePlanExp = dblInactivePlanExp + fltPlannedExpenses
						dblInactiveActualExp = dblInactiveActualExp + fltActualExpenses
					else
						dblTotalBudget = dblTotalBudget + fltPlannedBalance
						dblTotalActual = dblTotalActual + fltActualBalance
						dblPlanFund = dblPlanFund + fltPlannedFunding
						dblActualFund = dblActualFund + fltActualFunding
						dblPlanExp = dblPlanExp + fltPlannedExpenses 
						dblActualExp = dblActualExp + fltActualExpenses
					end if
				else
					strExclude = strExclude & "<tr> " & _
								"<td class='TableCell' >" & _
								"<a href=""javascript:"" onclick=""jfViewReport('" & rs("intStudent_ID") & "');"">" & _
								oHtml.ToolTip("<b>" & rs("szLAST_NAME") & ", " & rs("szFIRST_NAME") & "</b>", _
										"<table><tr><td class='svplain8' nowrap><b>Sponsor Teacher:</b> <a href=""mailto:" & rs("TEACHER_EMAIL") & """>" & rs("TEACHERS_NAME") & "</a><BR>" & _
										"<b>Guardians:</b> <a href=""mailto:" & rs("szEMAIL") & """>" & rs("szDesc") & "</a><BR>" & _
										"<b>Family Phone:</b> " & rs("szHome_Phone") & strIADate & "</td></tr></table>", _
										false,"",true,"ToolTip","","",false,false)  & _
										"</a>&nbsp;" & _
								strInactive & _
								"</td>" & _
								"<td class='TableCell'  align='right'>" & _
								"	$" & formatnumber(fltPlannedFunding,2) & "&nbsp;" & _
								"</td>" & _
								"<td class='TableCell'  align='right'>" & _
								"	$" &  formatnumber(fltPlannedExpenses,2) & "&nbsp;" & _
								"</td>	" & _		
								"<td class='TableCell'  align='right'>" 
				
					if fltPlannedBalance < -.0001 then
						strExclude = strExclude & "<span style='color:red;'>$<" & formatnumber(fltPlannedBalance,2) & "</span>&nbsp;" 
					else 
						strExclude = strExclude & "$" & formatnumber(fltPlannedBalance,2) & "&nbsp;"
					end if 
					
					strExclude = strExclude & "</td>" & _
								"<td class='TableCell'  align='right'>" & _
								"	$" & formatnumber(fltActualFunding,2) & "&nbsp;" & _
								"</td>" & _
								"<td class='TableCell'  align='right'>" & _
								"	$" & formatnumber(fltActualExpenses,2) & "&nbsp;" & _
								"</td>" & _
								"<td class='TableCell' align='right'>" 
								
					if fltActualBalance < -.0001 then
						strExclude = strExclude & "	<span style='color:red;'>$" & formatnumber(fltActualBalance,2) & "</span>&nbsp;" 
					else 
						strExclude = strExclude & "$" & formatnumber(fltActualBalance,2) & "&nbsp;"
					end if 
					
					strExclude = strExclude & "</td>" & _
											"</tr>	"
					dblExcludeTotalBudget = dblExcludeTotalBudget + fltPlannedBalance
					dblExcludeTotalActual = dblExcludeTotalActual + fltActualBalance
					dblExcludePlanFund = dblExcludePlanFund + fltPlannedFunding
					dblExcludeActualFund = dblExcludeActualFund + fltActualFunding
					dblExcludePlanExp = dblExcludePlanExp + fltPlannedExpenses 
					dblExcludeActualExp = dblExcludeActualExp + fltActualExpenses
				end if
			end if
			
			select case ucase(rs("szGrade"))
				case "K"
					intK = intK + 1
				case "1"
					int1 = int1 + 1
				case "2"
					int2 = int2 + 1
				case "3"
					int3 = int3 + 1
				case "4"
					int4 = int4 + 1
				case "5"
					int5 = int5 + 1
				case "6"
					int6 = int6 + 1
				case "7"
					int7 = int7 + 1
				case "8"
					int8 = int8 + 1
				case "9"
					int9 = int9 + 1
				case "10"
					int10 = int10 + 1
				case "11"
					int11 = int11 + 1
				case "12"
					int12 = int12 + 1
			end select
			
			if rs("intReEnroll_State") = "86" or rs("intReEnroll_State") = "123" then
				intInactiveCount = intInactiveCount + 1
			end if
			
			rs.MoveNext										
		loop
		response.Write "<tr><td align='right' class='svplain8'><b>All Student Totals:</b></td>" & _
					   "<td class='svplain8' align='right'>$" & formatnumber(dblPlanFund + dblInactivePlanFund,2) & "</td>" & _
					   "<td class='svplain8' align='right'>$" & formatnumber(dblPlanExp + dblInactivePlanExp,2) & "</td>" & _
					   "<td class='svplain8' align='right'>$" & formatnumber(dblTotalBudget + dblInactiveTotalBudget,2) & "</td>" & _
					   "<td class='svplain8' align='right'>$" & formatnumber(dblActualFund + dblInactiveActualFund,2) & "</td>" & _
					   "<td class='svplain8' align='right'>$" & formatnumber(dblActualExp + dblInactiveActualExp,2) & "</td>" & _
					   "<td class='svplain8' align='right'>$" & formatnumber(dblTotalActual + dblInactiveTotalActual,2) & "</td></tr>"& _
					   "<tr><td align='right' class='svplain8'><b>Active Student Totals:</b></td>" & _
					   "<td class='svplain8' align='right'>$" & formatnumber(dblPlanFund,2) & "</td>" & _
					   "<td class='svplain8' align='right'>$" & formatnumber(dblPlanExp,2) & "</td>" & _
					   "<td class='svplain8' align='right'>$" & formatnumber(dblTotalBudget,2) & "</td>" & _
					   "<td class='svplain8' align='right'>$" & formatnumber(dblActualFund,2) & "</td>" & _
					   "<td class='svplain8' align='right'>$" & formatnumber(dblActualExp,2) & "</td>" & _
					   "<td class='svplain8' align='right'>$" & formatnumber(dblTotalActual,2) & "</td></tr>" & _
					   "<tr><td align='right' class='svplain8'><b>Inactive Student Totals:</b></td>" & _
					   "<td class='svplain8' align='right'>$" & formatnumber(dblInactivePlanFund,2) & "</td>" & _
					   "<td class='svplain8' align='right'>$" & formatnumber(dblInactivePlanExp,2) & "</td>" & _
					   "<td class='svplain8' align='right'>$" & formatnumber(dblInactiveTotalBudget,2) & "</td>" & _
					   "<td class='svplain8' align='right'>$" & formatnumber(dblInactiveActualFund,2) & "</td>" & _
					   "<td class='svplain8' align='right'>$" & formatnumber(dblInactiveActualExp,2) & "</td>" & _
					   "<td class='svplain8' align='right'>$" & formatnumber(dblInactiveTotalActual,2) & "</td></tr>" 
					   
		call vbsPrintFPCSAccountTotals
		
		if strExclude <> "" then	
			response.Write "<tr><td colspan='7'><p></p></td></tr><tr><td colspan='7' class='svplain8'><BR><BR><B>Budget Information for Students that Withdrew Prior to the Count</b></td></tr>" 
			RESPONSE.Write vbfTableHeaders("true") & strExclude		
			response.Write "<tr><td align='right' class='svplain8'><b>Non Count Totals:</b></td>" & _
					   "<td class='svplain8' align='right'>$" & formatnumber(dblExcludePlanFund,2) & "</td>" & _
					   "<td class='svplain8' align='right'>$" & formatnumber(dblExcludePlanExp,2) & "</td>" & _
					   "<td class='svplain8' align='right'>$" & formatnumber(dblExcludeTotalBudget,2) & "</td>" & _
					   "<td class='svplain8' align='right'>$" & formatnumber(dblExcludeActualFund,2) & "</td>" & _
					   "<td class='svplain8' align='right'>$" & formatnumber(dblExcludeActualExp,2) & "</td>" & _
					   "<td class='svplain8' align='right'>$" & formatnumber(dblExcludeTotalActual,2) & "</td></tr>" 					   					
		end if
		
%>
		<tr>
			<td colspan="7" class="svplain8">
				<br><br>
				<b>Student Counts</b>
			</td>
		</tr>
		<tr>
			<td colspan="7">
				<table ID="Table4">
					<tr>
						<td class='TableHeader' align="center">
							K
						</td>
						<td class='TableHeader' align="center">
							1
						</td>
						<td class='TableHeader' align="center">
							2
						</td>
						<td class='TableHeader' align="center">
							3
						</td>
						<td class='TableHeader' align="center">
							4
						</td>
						<td class='TableHeader' align="center">
							5
						</td>
						<td class='TableHeader' align="center">
							6
						</td>
						<td class='TableHeader' align="center">
							7
						</td>
						<td class='TableHeader' align="center">
							8
						</td>
						<td class='TableHeader' align="center">
							9
						</td>
						<td class='TableHeader' align="center">
							10
						</td>
						<td class='TableHeader' align="center">
							11
						</td>
						<td class='TableHeader' align="center">
							12
						</td>	
						<td class='TableHeader' align="center">
							Total
						</td>		
						<td class='TableHeader' align="center">
							Total Inactive
						</td>		
						<td class='TableHeader' align="center">
							Total Active
						</td>	
					</tr>
					<tr>
						<td align="center" class="svplain8">
							<% = intK %>
						</td>
						<td align="center" class="svplain8">
							<% = int1 %>
						</td>
						<td align="center" class="svplain8">
							<% = int2 %>
						</td>
						<td align="center" class="svplain8">
							<% = int3 %>
						</td>
						<td align="center" class="svplain8">
							<% = int4 %>
						</td>
						<td align="center" class="svplain8">
							<% = int5 %>
						</td>
						<td align="center" class="svplain8">
							<% = int6 %>
						</td>
						<td align="center" class="svplain8">
							<% = int7 %>
						</td>
						<td align="center" class="svplain8">
							<% = int8 %>
						</td>
						<td align="center" class="svplain8">
							<% = int9 %>
						</td>
						<td align="center" class="svplain8">
							<% = int10 %>
						</td>
						<td align="center" class="svplain8">
							<% = int11 %>
						</td>
						<td align="center" class="svplain8">
							<% = int12 %>
						</td>	
						<td align="center" class="svplain8">
							<% = intK + int1 + int2 + int3 + int4 + int5 + int6 + int7 + int8 + int9 + int10 + int11 + int12 %>
						</td>	
						<td align="center" class="svplain8">
							<% = intInactiveCount %>
						</td>	
						<td align="center" class="svplain8">
							<% = ( intK + int1 + int2 + int3 + int4 + int5 + int6 + int7 + int8 + int9 + int10 + int11 + int12) - intInactiveCount %>
						</td>					
					</tr>
				</table>
			</td>
		</tr>
				
<%
		response.Write "</table>"
	end if ' end if recordcount > 0
	rs.Close
	set rs = nothing
	response.Write oHtml.ToolTipDivs
%>
		</td>
	</tr>
<% end if %>
</table>
</form>
<script language="javascript">
	function jfViewReport(pStudentID) {
		var winSPR;
				
		strURL = "<%=Application.Value("strWebRoot")%>forms/packet/packet.asp?SimpleHeader=true&intStudent_id=" + pStudentID;
		winSPR = window.open(strURL,"winSPR","width=800,height=500,scrollbars=yes,resizable=yes");
		winSPR.moveTo(0,0);
		winSPR.focus();
	}
	

	function jfPrint(){
		var winPrint;
		var studentList = document.main.studentList.value;	
		strURL = "<%=Application.Value("strWebRoot")%>Reports/StudentProgressReport.asp?print=true&intStudent_id=" + studentList + "&intReporting_Period_ID=<%=intReporting_Period_ID%>";
		winPrint = window.open(strURL,"winPrint","width=710,height=500,scrollbars=yes,resizable=yes");
		winPrint.moveTo(0,0);
		winPrint.focus();
	}
</script>
<%
if request("selReportType") <> "" then 
	response.Write "<span class='svplain'>" & (datediff("s",start,now())/60) & "</span>"
end if

call oFunc.CloseCN()
set oFunc = nothing
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")

function vbfTableHeaders(pHideBreak)
		if pHideBreak = "" then
%>
			<tr>
				<td colspan='7'>
					<p></p>
				</td>
			</tr>
<%
		end if
%>	
			<tr>
				<td class="TableHeader" align="center">
					<b>Student Name/Packet Link</b>
				</td>
				<td class="TableHeader"  align="center">
					<b>Planned Funding</b>
				</td>
				<td class="TableHeader"  align="center">
					<b>Planned Expenses</b>
				</td>
				<td class="TableHeader"  align="center">
					<b>Planned Budget Balance</b>
				</td>	
				<td class="TableHeader"  align="center">
					<b>Actual Funding</b>
				</td>
				<td class="TableHeader"  align="center">
					<b>Actual Expenses</b>
				</td>
				<td class="TableHeader"  align="center">
					<b>Actual Balance</b>
				</td>							
			</tr>
<%
end function

sub vbsPrintFPCSAccountTotals
	set rs2 = server.CreateObject("ADODB.RECORDSET")
	rs2.CursorLocation = 3
	
	if oFunc.IsAdmin and Application.Contents("SchoolBudgetAccount") <> "" then
		SBA = Application.Contents("SchoolBudgetAccount")
		sql = "SELECT bt.intFrom_Student_ID, bt.intTo_Student_ID, bt.curAmount, s1.szLAST_NAME + ', ' + s1.szFIRST_NAME AS FromStudent,  " & _ 
				" s2.szLAST_NAME + ', ' + s2.szFIRST_NAME AS ToStudent, bt.dtCREATE, bt.szComment, bt.szUSER_CREATE " & _ 
				"FROM tblBudget_Transfers bt INNER JOIN " & _ 
				"	tblSTUDENT s1 ON bt.intFrom_Student_ID = s1.intSTUDENT_ID INNER JOIN " & _ 
				"	tblSTUDENT s2 ON bt.intTo_Student_ID = s2.intSTUDENT_ID " & _ 
				"WHERE (bt.intSchool_Year = " & session.Contents("intSchool_Year")& ") AND (bt.intFrom_Student_ID = " & SBA  & ") OR " & _ 
				"	(bt.intSchool_Year = " & session.Contents("intSchool_Year")& ") AND (bt.intTo_Student_ID = " & SBA & ") " & _ 
				"ORDER BY bt.dtCREATE "	
		rs2.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
		
		if rs2.RecordCount > 0 then
%>
<tr>
	<td colspan="7" class="svplain8">
		<BR><BR>
		<span class="svplain10">&nbsp;<b>FPCS Account Summary</b> <BR>(includes all transactions for all students for <% = oFunc.SchoolYearRange %> school year.)</span>
		<table cellpadding="3" ID="Table5">
			<tr>
				<td class="TableHeader">Date</td>
				<td class="TableHeader">Comments</td>
				<td class="TableHeader">Created By</td>
				<td class="TableHeader">From Student</td>
				<td class="TableHeader">To Student</td>
				<td class="TableHeader">Amount</td>
			</tr>
			<%
			dim FpcsTo, FpcsFrom
			FpcsTo = 0
			FpcsFrom = 0
			do while not rs2.EOF
%>
			<tr>
				<td class="TableCell">
					<% = formatDateTime(rs2("dtCREATE"),2) %>
				</td>
				<td class="TableCell">
					&nbsp;<% = rs2("szComment") %>
				</td>	
				<td class="TableCell">
					<% = rs2("szUser_Create") %>
				</td>
				<td class="TableCell">
					<% = rs2("FromStudent") %>
				</td>
				<td class="TableCell">
					<% = rs2("ToStudent") %>
				</td>
				<td class="TableCell" align="right">
					$<% = formatNUmber(rs2("curAmount"),2) %>
				</td>
			</tr>
<%			
			if SBA & ""  = cstr(rs2("intFrom_Student_ID")) & "" then
				FpcsFrom = FpcsFrom + cdbl(rs2("curAmount"))
			end if
			
			if SBA & "" = cstr(rs2("intTo_Student_ID")) & "" then
				FpcsTo = FpcsTo + cdbl(rs2("curAmount"))
			end if
			
			rs2.MoveNext						
		loop
%>
		</table>
		Total Withdrew from FPCS Account: $<% = formatnumber(FpcsFrom,2) %><br>
		Total Deposited into FPCS Account: $<% = formatnumber(FpcsTo,2) %><br>
		<b>Net: $<% = formatnumber(FpcsTo - FpcsFrom,2)%> </b>
	</td>
</tr>
<%			
		end if
		rs2.Close
	end if
	
	set rs2 = nothing

end sub			
%>