<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		packet.asp
'Purpose:	Main information page contaning Course management, budgets,
'			and student status information
'Date:		26 oCt 2004
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim intStudent_ID, intShort_ILP_ID , strWhere
dim intPreviousID		' Used to determine when a course is changed in rsBudget
dim strColor			' Used to define table row colors
dim intColor			' Used as a marker to alternate row colors
dim intActualEnroll		' Actual enrollment percentage
dim intTargetEnroll		' Target enrollment percentage
dim oFunc				' wsc object
dim dblFunding			' funding amount set for grade level
dim dblTargetBalance	' Target start - all budgeted expenses
dim dblActualBalance	' Actual start - all actual expenses
dim dblWithdraw			' Amount to reduce budget funding by due to Budget Transfer withdrawal 
dim dblDeposit			' Amount to reduce budget funding by due to Budget Transfer deposit 
dim dblBudgetCost		' Calculated cost for a budgeted item
dim dblUnitCost			' Used to handle teachers cost vs budgeted goods/services
dim dblShipping			' Used to track shipping costs
dim dblCharge 
dim dblAdjBudget 
dim dblClassCharge 
dim dblClassBudget 
dim dblTotalCharge 
dim dblTotalBudget 
dim mDivCount
dim mLablelCount
dim strBList
dim strDateField
dim bStatus				' budgeted item status
mLablelCount = 0
mDivCount = 0


set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()
oFunc.ResetSelectSessionVariables


'Initialize some key variables
if request("intStudent_ID") <> "" then
	intStudent_ID = request("intStudent_ID") 
	intShort_ILP_ID = request("intShort_ILP_ID")
	intActualEnroll = oFunc.StudentPercentage(intStudent_ID)
	arTargetFunding = oFunc.TargetFundingInfo(intStudent_ID)
	
	arStudentEnroll = oFunc.arStudentEnroll 
	intTotalHrs = (oFunc.CoreHours + oFunc.ElectiveHours)

	' Figure out funding figures
	intTargetEnroll = arTargetFunding(0)
	dblFunding = cdbl(arTargetFunding(1))
	dblTargetBalance = formatNumber(cdbl(arTargetFunding(2)),2)
	dblActualBalance = formatNumber((cdbl(intActualEnroll * .01) * dblFunding),2)
	
else
	'terminate page since page was improperly called.
	response.Write "<html><body><H1>Page improperly called.</h1></body></html>"
	response.End
end if

'prevent page from cacheing
Response.AddHeader "Cache-Control","No-Cache"
Response.Expires = -1

'Print the header
Session.Value("strTitle") = "Budget Worksheet"
Session.Value("strLastUpdate") = "26 Oct 2004"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get student Information
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
set rsStudentInfo = server.CreateObject("ADODB.RECORDSET")
with rsStudentInfo
	.CursorLocation = 3
		sql = "SELECT s.intStudent_ID, s.szFirst_Name, s.szLast_Name,e.intPercent_Enrolled_Locked, " & _
			"s.intGrad_year, s.szGrade, e.intPercent_Enrolled_FPCS, e.intEnroll_Info_ID, " & _
			"i.szFirst_Name + ' ' + i.szLast_Name as Sponsor, i.szEmail as SponsorEmail, e.bolASD_Testing, e.bolProgress_Agreement " & _
			"FROM tblStudent s left outer join " & _
			" tblIEP ON s.intSTUDENT_ID = tblIEP.intStudent_ID LEFT OUTER JOIN " & _
			"tblEnroll_info e on s.intStudent_ID = e.intStudent_ID  LEFT OUTER JOIN "  & _
			"tblInstructor i on e.intSponsor_Teacher_ID = i.intInstructor_ID inner join " & _
			" tblStudent_States ss on s.intStudent_ID = ss.intStudent_ID and " & _
			"ss.intSchool_year = " & session.Contents("intSchool_Year") & " " & _
			"where s.intStudent_ID = " & intStudent_ID & _
			" and e.sintSchool_Year = " & session.Value("intSchool_Year") & _
			" AND (tblIEP.intSchool_Year = " & session.Value("intSchool_Year") & ")"
	'response.Write sql 
	.Open sql, oFunc.FPCScnn
	
	if .RecordCount > 0 then
		'This for loop dimentions and defines all the columns we selected in sqlClass
		'and we use the variables created here to populate the form.
		for each item in .Fields
			execute("dim " & item.Name)
			execute(item.Name & " = item")		
		next  
		if Sponsor = "" then Sponsor = "No Sponsor Selected": SponsorEmail=""
	else
		Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")
		%>
			<table cellspacing=0 cellpadding=4 border=1 width=85% ID="Table4">
				<tr>
					<td class=svplain10>
						<b>Before you can plan any courses you must update
						your students information for SY <% = session.Contents("intSchool_Year")%>.
						To do this click on the 'Family Manager' link on the menu above
						follow the instructions found on that page.
						
						</b>
					</td>
				</tr>
			</table>
		<%
		
		call oFunc.CloseCN()
		set oFunc = nothing
		Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")
		response.End
	end if
	.Close
end with
set rsStudentInfo = nothing

'Find out if student is in High School
if isNumeric(szGrade) then
	if cint(szGrade) >= 9 then
		bolHighSchool = true
	else
		bolHighSchool = false
	end if
end if

strStudentInfo = "<table cellspacing=1 cellpadding=2 >" & _	
					"<tr><td class='TableCell'>Core Unit:</td>" & _			
					"<td class='TableCell' align='right'>" & formatNumber((arStudentEnroll(0)/90),1) & "</td></tr>" & _
					"<tr><td class='TableCell'>Elective Unit:</td>" & _
					"<td class='TableCell' align='right'>" & formatNumber((arStudentEnroll(1)/90),1) & "</td></tr>" & _
					"<tr><td class='TableCell'>ASD Contracted Hrs:</td>" & _
					"<td class='TableCell' align='right'>" &  arStudentEnroll(2) & "</td></tr>" & _
					"<tr><td class='TableCell'>Total Hrs:</td>" & _
					"<td class='TableCell' align='right'>" & intTotalHrs & "</td>" & _
					"</tr></table>" 


	if bolASD_Testing then
		strTestForm = "Yes"
	else
		strTestForm = "No"
	end if
	
	if bolProgress_Agreement then
		strProgressForm = "Yes"
	else
		strProgressForm = "No"
	end if
			
strFormsTable =		"<table cellspacing=1 cellpaddin='2' style='height:100%;width:100%'>" & _
					"<tr><td class='TableCell' colspan=2>MANDATORY SIGNED FORMS</td></tr>" & _
					"<tr><td class='TableCell'>Student Testing:</td>" & _
					"<td class='TableCell'>" & strTestForm & "</td></tr>" & _
					"<tr><td class='TableCell'>Student Progress Report:</td>" & _
					"<td class='TableCell'>" & strProgressForm & "</td>" & _
					"</tr></table>" 					
	
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")
	%>
	<link rel="stylesheet" href="<% =  Application.Value("strWebRoot") %>CSS/printStyle.css">
	<script language=javascript>
		function jfPrint(){
			if (window.print){
			window.print()
			}
			else {
			alert("Mac users: please press Apple-P to print this form.\nWindows users: Please press ctrl-P to print this form.")
			}
		}		
		jfPrint();
	</script>
	<%

' Print top section of worksheet
%>
<script language=javascript>	
	
	
	function jfToggle(pList,pID){
		// toggles display of objs in pList on and off
		var arList = pList.split(",");
		var i;
		var obj;
		var sText;
		for(i=0;i< arList.length;i++){
			if (arList[i] != '') {
				obj = document.getElementById('div'+arList[i]);
				if (obj.style.display == 'none'){
					obj.style.display = 'block';
					sText = 'hide';
				}else{
					obj.style.display = 'none';
					sText = 'show';
				}
			}
		}
		if (pID != ''){
			obj = document.getElementById(pID);
			obj.innerHTML = sText;
		}
	}
	

</script>
<form name="main" action="<%=Application("strSSLWebRoot")%>forms/packet/packet.asp" method="post" ID="Form1">
<input type="hidden" name="intStudent_ID" value="<%=intStudent_ID%>" ID="Hidden2">
<input type="hidden" name="bolHighSchool" value="<%=bolHighSchool%>" ID="Hidden3">
<input type="hidden" name="courseTitleData" value="" ID="Hidden4">
<input type=hidden name="simpleHeader" value="<% = request("simpleHeader") %>" ID="Hidden5">
<input type=hidden name="lastIndex" value="" ID="Hidden6">
<table style="width:640px;" ID="Table3">
	<tr>
		<td style="width:100%;">
			<table style="width:100%;">
				<tr>
					<td align=left>
						<img src="<% = Application("strImageRoot")%>fpcsLogo.gif">
					</td>
					<td align=right class=svplain10 width=100% nowrap>
						3339 Fairbanks St.<br>
						Anchorage, AK 99503<br>
						Ph: 907-742-3700<br>
						Fax: 907-742-3710
					</td>
				</tr>
				<tr class=yellowHeader>	
					<Td colspan=2>
						<table align=right valign="middle" ID="Table27"><tr><td align=right><font face=arial size=2 color=white><% = date()%></font></td></tr></table>
						&nbsp;<span style='font-size:10pt;'><b>Student Packet/Budget</b> 
						for <% = oFunc.StudentInfo(intStudent_ID,3) %></span>								
					</td>					
				</tr>
			</table>
		</td>
	</tr>	
	<tr>
		<td>
			<table ID="Table2" style="width:100%;">
					<tr>
						<td class="TableSubHeader" colspan=4>
							<B>&nbsp;Student Enrollment Information</B>
						</td>
					</tr>
					<tr>
						<td valign='top'  style='height:100%;'>
							<table ID="Table29" cellspacing='1' cellpadding='2' style='height:100%;width:100%;'>
									<tr>	
										<td nowrap class="TableCell" title="The enrollment goal you have chosen determines the eligible funding amount for SY <% = oFunc.SchoolYearRange%>.">					
											<b>Planned Enrollment:</b>
										</td>
										<td class="TableCell" valign=middle>
											&nbsp;<% if intPercent_Enrolled_Locked <> "" then response.Write intPercent_Enrolled_Locked else response.Write intPercent_Enrolled_FPCS end if%>% 
									</tr>							
									<tr>
										<td class="TableCell" nowrap valign=middle title="This is the enrollment level you currently have based on the amount of the plan you've implemented.">
											<b>Actual Enrollment:</b>					
										</td>
										<td class="TableCell" valign=middle>
											&nbsp;<% = intActualEnroll%>%	
										</td>
									</tr>
									<tr>
										<td class="TableCell" valign=middle nowrap>																				
											Sponsor Teacher:															
										</td>
										<td class="TableCell" valign=middle nowrap>
											&nbsp;<% = Sponsor%>&nbsp;									
										</td>
									</tr>
									<tr>
									<td class="TableCell" valign=top title="This is the enrollment level you currently have based on the amount of the plan you've implemented.">
										Student Grade: 						
									</td>
									<td class="TableCell" valign=top>
										&nbsp;<% = szGrade%>
									</td>
								</tr>
								</table>
							</td>
							<td valign=top stlye="width:33%;" align=center>
								<% = strStudentInfo %>
							</td>
							<td valign=top style='height:100%;width:33%;'>
								<% = strFormsTable %>
							</td>
						</tr>
					</table>
		</td>
	</tr>
	<tr>
		<td>
			<table ID="Table1">				
				<tr >					
					<td bgcolor=white colspan="11" >
						
					</td>
					<td class=TableHeader align=center>
						Planned Budget
					</td>
					<td class=TableHeader align=center>
						Actual
					</td>
				</tr>
				<tr >
					<td colspan="10" align=right class="svplain8">
						Beginning Balance:	
					</td>
					<td bgcolor=white>
					</td>
					<td class=TableHeader align=right>
						$<%=dblTargetBalance%>
					</td>
					<td class=TableHeader align=right>
						$<%=dblActualBalance%>
					</td>
				</tr>
<%
	dblDeposits = oFunc.TransferAdd(intStudent_ID)
	dblWithdraw = oFunc.TransferDeduct(intStudent_ID)
	dblActualBalance = (cdbl(dblActualBalance) + cdbl(dblDeposits)) - cdbl(dblWithdraw)
	dblTargetBalance = (cdbl(dblTargetBalance) + cdbl(dblDeposits)) - cdbl(dblWithdraw)
%>				
				<tr >
					<td colspan="10" align=right class="svplain8">
						Budget Transfer Deposits:	
					</td>
					<td bgcolor=white>
						&nbsp;&nbsp;&nbsp;
					</td>
					<td class=TableHeader align=right>
						<nobr>$<%=dblDeposits%></nobr>
					</td>
					<td class=TableHeader align=right>
						<nobr>$<%=dblDeposits%></nobr>
					</td>
				</tr>
				<tr >
					<td colspan="10" align=right class="svplain8">
						Budget Transfer Withdrawals:	
					</td>
					<td bgcolor=white>
						&nbsp;&nbsp;&nbsp;
					</td>
					<td class=TableHeader align=right>
						<nobr>- $<%=dblWithdraw%></nobr>
					</td>
					<td class=TableHeader align=right>
						<nobr>- $<%=dblWithdraw%></nobr>
					</td>
				</tr>
<%

'Define Where clause.  This logic determines if we show a budget worksheet
'for all courses for a given student or only for a given course

if intShort_ILP_ID <> "" then
	'Show only for a specific course
	strWhere = " (ISF.intShort_ILP_ID = " & intShort_ILP_ID & ")"
else
	'show all courses
	strWhere = " (ISF.intStudent_ID = " & intStudent_ID & _
			   ") AND (ISF.intSchool_Year = " & session.Contents("intSchool_Year") & ")"
end if 

sql = "SELECT ISF.szCourse_Title, POS.txtCourseTitle, ISF.intShort_ILP_ID, I.szName, tblILP.intILP_ID,tblILP.bolApproved as aStatus, tblILP.bolSponsor_Approved as sStatus,oi.bolApproved, oi.bolSponsor_Approved,  " & _ 
		" CASE ISF.intPOS_Subject_ID WHEN 22 THEN 0 ELSE 1 END AS isSponsor, oi.intQty, oi.curUnit_Price, oi.curShipping,ISF.intCourse_Hrs, tblILP.decCourse_Hours,  " & _ 
		" oi.intQty * oi.curUnit_Price + oi.curShipping AS total, oi.intOrdered_Item_ID, tblClasses.intInstructor_ID, tps.szSubject_Name, tblClasses.intClass_ID, " & _ 
		" tblClasses.intInstruct_Type_ID, tblILP.intContract_Guardian_ID,tblClasses.intGuardian_ID,tblClasses.intVendor_ID, " & _
		" tblClasses.szClass_Name, CASE WHEN tblClasses.intInstructor_ID IS NOT NULL THEN ins.szFirst_Name + ' ' + ins.szLast_Name   " & _
		" WHEN tblClasses.intGuardian_ID IS NOT NULL THEN g.szFirst_Name + ' ' + g.szLast_Name END AS teacherName, " & _
		" tblILP.szAdmin_Comments, tblILP.szSponsor_Comments, tblILP.bolReady_For_Review, tblILP.dtReady_For_Review, " & _
		"          (SELECT top 1 oa2.szValue " & _
        "             FROM          tblOrd_Attrib oa2 " & _
        "             WHERE      oa2.intOrdered_Item_Id = oi.intOrdered_Item_Id AND " & _
        "			 (oa2.intItem_Attrib_ID = 9 OR " & _
        "              oa2.intItem_Attrib_ID = 5 OR " & _
        "              oa2.intItem_Attrib_ID = 6 OR " & _
		"              oa2.intItem_Attrib_ID = 22 or oa2.intItem_Attrib_ID = 33) order by oa2.intOrd_Attrib_ID) AS oiDesc, bolClosed " & _
		", oi.bolReimburse, I.intItem_Group_ID, oi.szDeny_Reason " & _
		"FROM tblClasses INNER JOIN " & _ 
		" tblILP ON tblClasses.intClass_ID = tblILP.intClass_ID LEFT OUTER JOIN " & _ 
		" trefItems I INNER JOIN " & _ 
		" tblOrdered_Items oi ON I.intItem_ID = oi.intItem_ID ON tblILP.intILP_ID = oi.intILP_ID RIGHT OUTER JOIN " & _ 
		" tblILP_SHORT_FORM ISF ON tblILP.intShort_ILP_ID = ISF.intShort_ILP_ID LEFT OUTER JOIN " & _ 
		" tblProgramOfStudies POS ON ISF.lngPOS_ID = POS.lngPOS_ID inner join " & _
		"  trefPOS_SUBJECTS tps ON tps.intPOS_SUBJECT_ID = ISF.intPOS_SUBJECT_ID LEFT OUTER JOIN " & _
		" tblINSTRUCTOR INS ON tblClasses.intInstructor_ID = INS.intINSTRUCTOR_ID left outer join" & _
		" tblGUARDIAN g ON tblClasses.intGuardian_ID = g.intGUARDIAN_ID " & _
		"WHERE " & strWhere & _
		" ORDER BY isSponsor, POS.txtCourseTitle, ISF.szCourse_Title, ISF.intShort_ILP_ID "
	
set rsBudget = server.CreateObject("ADODB.RECORDSET")
rsBudget.CursorLocation = 3
rsBudget.Open sql,oFunc.FPCScnn

intPreviousID = 0

do while not rsBudget.EOF
	' We check to see if the course has changed within the recordset
	' If so we will need to reprint the table headers.
	
	if intPreviousID <> rsBudget("intShort_ILP_ID") then		
		intPreviousID = rsBudget("intShort_ILP_ID")		
		
		if intColor > 0 then 
			rsBudget.MovePrevious
			call vbsShowTotals()
			rsBudget.MoveNext
		end if		
		
		' Handle Course Hours
		if isNumeric(rsBudget("decCourse_Hours")) then
			intHours = rsBudget("decCourse_Hours")
		elseif isNumeric(rsBudget("intCourse_Hrs")) then 
			intHours = rsBudget("intCourse_Hrs")
		else
			intHours = 0 
		end if
		
		if rsBudget("intInstructor_ID") & "" <> "" then
			strContractSchedule = "Contract"
		else
			strContractSchedule = "Schedule"
		end if
		
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''		
' Handle ILP Status
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''				
		strStatus = ""
		bolLock = true 
		' Handles status for guardians
		select case rsBudget("aStatus")
			case true
				strStatus = "a-appr"
			case false
				strStatus = "a-must fix"
		end select
		
		if strStatus = "" then
			select case rsBudget("sStatus")
				case true
					strStatus = "s-appr"
				case false
					strStatus = "s-must fix"
			end select							
		end if
		
		' unlocks Packet so it can be deleted
		if strStatus = "" then
			if  (ucase(session.Contents("strRole")) <> "ADMIN" _
			and cint(application.Contents("intYear_Locked")) < cint(session.Contents("intSchool_Year")))then
				bolLock = false	
			end if
		end if
		
		if rsBudget("intILP_ID") & "" <> "" AND strStatus = "" then	
			if rsBudget("bolREady_For_Review") = true then
				strStatus = "ready for review"
			else				
				strStatus = "implemented"												
			end if
		elseif strStatus = "" then
			strStatus = "planned"	
		end if
	    
		mDivCount = mDivCount + 1
		strBList = strBList & mDivCount & ","
%>
				<tr>
					<td colspan="10" style="width:100%;">	
						<table style="width:100%;" cellpadding='2' cellspacing='1' ID="Table7">
							<tr class="TableHeader">
								<td align=left style="">
									&nbsp;<b>Course Title</b>
								</td>
								<td align='center'>
									<b>Subject</b>
								</td>
								<td align='center' nowrap style="width:0%;">
									<b>Hrs</b>
								</td>
								<td align='center' nowrap>
									<b>ILP Status</b>
								</td>
								<td align='center' >
									<b>Comments</b>
								</td>
							</tr>
							<tr>
								<td class="TableCell" valign="top"  >
									 <% = rsBudget("txtCourseTitle") & rsBudget("szCourse_Title")%>
								</td>
								<td class="TableCell" valign="top">
									 <% = rsBudget("szSubject_Name") %>
								</td>
								<td class="TableCell" align='center' valign="top">
									 <% = intHours %>
								</td>
								<td class="TableCell" align='center' valign="top" style="width:0%;">
									 <% = strStatus %>
								</td>	
								<td class="TableCell" valign="top" style="width:200px;">
									 <% 										
									if rsBudget("szAdmin_Comments") <> "" then
										response.Write "Admin Comments: " & rsBudget("szAdmin_Comments") 
									end if 
									
									if rsBudget("szSponsor_Comments") <> "" then
										if rsBudget("szAdmin_Comments") <> "" then response.Write "<BR>"
										response.Write "Sponsor Comments: " & rsBudget("szSponsor_Comments") 
									end if 																	
									 %>&nbsp;
								</td>															
							</tr>
						</table>				
						<nobr>										
					</td>
					<td   colspan="3">
						&nbsp;
					</td>
				</tr>				
				<tr id="div<% = mDivCount%>">
					<td class="TableSubHeader" align="center">
						Budget Item
					</td>
					<td class="TableSubHeader" align="center">
						Status
					</td>
					<td class="TableSubHeader">
						Description
					</td>
					<td class="TableSubHeader" align="center">
						QTY
					</td>
					<td class="TableSubHeader" align="center">
						Unit Cost
					</td>
					<td class="TableSubHeader" align="center" title="Shipping and Handling">
						S/H
					</td>
					<td class="TableSubHeader" align="center">
						Budget Total
					</td>
					<td class="TableSubHeader" align="center">
						Actual Charges
					</td>
					<td class="TableSubHeader" align="center">
						Budget Adjust
					</td>
					<td class="TableSubHeader" align="center">
						Budget Balance
					</td>
					<td  >
						&nbsp;
					</td>
					<td  >
						&nbsp;
					</td>
					<td  >
						&nbsp;
					</td>
				</tr>						
<%
		'Set alternating row color
		call vbsAlternateColor
		strClass = "TableCell"  ' default class setting
		if len(rsBudget("intInstructor_ID")) > 0 then
			' display teacher cost
			arASDCostInfo = oFunc.InstructionCostInfo(rsBudget("intILP_ID"))
			dblClassCharge = arASDCostInfo(0)
			dblClassBudget = arASDCostInfo(0)		
			mDivCount = mDivCount + 1
			strBList = strBList & mDivCount & ","
				%>	

				<tr id="div<%=mDivCount%>">
					<td class="<% = strClass %>">
						Instruction
					</td>
					<td class="<% = strClass %>"  align="center">
						n/a
					</td>
					<td class="<% = strClass %>">
						Instruction by: <% = arASDCostInfo(2)%> 
					</td>
					<td class="<% = strClass %>" align="center">
						<%= formatNumber(arASDCostInfo(1),2)%>
					</td>
					<td class="<% = strClass %>" align="right">
						$<%= formatNumber(arASDCostInfo(3),2)%>
					</td>
					<td class="<% = strClass %>" align="center">
						n/a
					</td>
					<td class="<% = strClass %>" align="right">
						$<%= formatNumber(arASDCostInfo(0))%>
					</td>
					<td class="<% = strClass %>" align="right">
						$<%= formatNumber(arASDCostInfo(0))%>
					</td>
					<td class="<% = strClass %>" align="right">
						$0.00
					</td>
					<td class="<% = strClass %>" align="right">
						$0.00
					</td>
					<td   colspan="1">
						&nbsp;
					</td>
					<td class="<% = strClass %>" align="right" nowrap>
						-$<%= formatNumber(arASDCostInfo(0))%>
					</td>
					<td class="<% = strClass %>" align="right" nowrap>
						-$<%= formatNumber(arASDCostInfo(0))%>
					</td>
				</tr>
	<% end if 			
	end if
			
	if rsBudget("intOrdered_Item_ID") & "" <> "" then 
		
		' Set the budgeted cost for this item
		dblShipping = 0
		if rsBudget("curShipping") & "" <> "" then
			if isNumeric(rsBudget("curShipping")) then
				dblShipping = formatNumber(rsBudget("curShipping"),2)
			end if
		end if
			
		dblBudgetCost = formatNumber(rsBudget("Total"),2)
		'Get Line Item info
		liInfo = LineItemInfo(rsBudget("intOrdered_Item_ID"),dblBudgetCost, rsBudget("bolClosed"), oFunc.FPCScnn,strClass)
		bStatus = GetBudgetStatus(rsBudget("intItem_Group_ID"),rsBudget("bolApproved"),liInfo(4),rsBudget("bolReimburse"))
		if bStatus = "rejc" then
			strClass = "TableCellStrike"
		else
			strClass = "TableCell"
		end if
		
		dblCharge = formatNumber(liInfo(1),2)
		dblAdjBudget = formatNumber(dblBudgetCost + cdbl(liInfo(2)),2)
		dblClassCharge = dblClassCharge + cdbl(dblCharge)
		dblClassBudget = dblClassBudget + cdbl(dblAdjBudget)
		mDivCount = mDivCount + 1
		strBList = strBList & mDivCount & ","
		
		if rsBudget("szDeny_Reason") <> "" then
			strReason = "<BR><b>Note:</b> " & rsBudget("szDeny_Reason")
		else
			strReason = ""
		end if
		
		if rsBudget("bolReimburse")  then
			strItemType = "Reimburse #" & rsBudget("intOrdered_Item_ID") & ": "
		else
			strItemType = "Requisition #" & rsBudget("intOrdered_Item_ID") & ": "
		end if
		' Print row with budget info		
%>																		
				<tr id="div<% = mDivCount %>">
					<td class=<% = strClass %>>
						<% = rsBudget("szName") %>
					</td>
					<td class="<% = strClass %>" nowrap align="center">
						<% = bStatus %>
					</td>
					<td class=<% = strClass %>>
						<% = strItemType & rsBudget("oiDesc") & strReason %>&nbsp;
					</td>
					<td class=<% = strClass %> align="center" nowrap>
						<% = rsBudget("intQTY") %>
					</td>
					<td class=<% = strClass %> align=right nowrap>
						$<% = formatNumber(rsBudget("curUnit_Price"),2) %>
					</td>
					<td class=<% = strClass %> align=right nowrap>
						&nbsp;$<% = dblShipping %>
					</td>
					<td class=<% = strClass %> align=right nowrap>
						$<% = dblBudgetCost %>
					</td>
					<td class=<% = strClass %> align=right nowrap>
						$<% = formatNumber(liInfo(1),2)%>
					</td>
					<td class=<% = strClass %> align=right nowrap>
						$<% = formatNumber(liInfo(2),2) %>
					</td>
					<td class=<% = strClass %> align=right nowrap>
						$<% = formatNumber((dblBudgetCost - cdbl(liInfo(1))) + cdbl(liInfo(2)),2)%>
					</td>
					<td bgcolor=white>
					</td>
					<td class="<% = strClass %>" align="right" nowrap>
						-$<% = dblAdjBudget %>
					</td>
					<td class="<% = strClass %>" align="right" nowrap>
						-$<% = dblCharge %>
					</td>
				</tr>
				<% = liInfo(0) %>
<%
	else
	mDivCount = mDivCount + 1
	strBList = strBList & mDivCount & ","
%>
				<tr bgcolor="<% = strColor%>" id="div<%=mDivCount%>">
					<td class=svplain8 colspan=9>
						No items have been budgeted for this course.	
					</td>
					<td bgcolor=white>
					</td>
					<td  >
						&nbsp;
					</td>
					<td  >
						&nbsp;
					</td>
				</tr>
<%
	end if 
	rsBudget.MoveNext
loop	

'Print last course totals
if rsBudget.RecordCount > 0 then
	rsBudget.MoveLast
	call vbsShowTotals()
	
	dblTargetBalance = dblTargetBalance - dblTotalBudget
	dblActualBalance = dblActualBalance - dblTotalCharge
%>

			  <tr bgcolor="<% = strColor%>">
					<td class=svplain10 colspan="10" align=right>
						Available Remaining Funds:	
					</td>
					<td bgcolor=white>
						&nbsp;&nbsp;&nbsp;
					</td>
					<td class="TableHeader" align=right>
						$<%=formatNumber(dblTargetBalance,2)%>
						<input type=hidden name="budgetBalance" value="<%=formatNumber(dblTargetBalance,2)%>" ID="Hidden8">
					</td>
					<td class="TableHeader" align=right>
						$<%=formatNumber(dblActualBalance,2)%>
					</td>
				</tr>
<script language=javascript>
	function jfBudget(id,bid,courseTitle){
		// Opens up edit add window
		var winBudgetTool;
		var URL = "budgetItemTool.asp?intStudent_ID=<%=intStudent_ID%>&intBudget_ID="+bid;
		URL += "&intShort_ILP_ID=" + id+"&budgetBalance=<%=formatNumber(dblTargetBalance,2)%>";	
		URL += "&szCourseTitle=" + courseTitle;
		winBudgetTool = window.open(URL,"winBudgetTool","width=640,height=270,scrollbars=yes,resizable=yes");
		winBudgetTool.moveTo(0,0);
		winBudgetTool.focus();
	}
	function jfToggleBudget(pMe){
		jfToggle('<%=strBList%>','');		
		
		if (pMe.value == "Show Budget") {
			pMe.value = "Hide Budget";
		}else{
			pMe.value = "Show Budget";
		}
	}
</script>				
<%
end if

set rsBudget = nothing					
%>							
			</table>
		</td>
	</tr>
</table>
</form>
<%

call oFunc.CloseCN()
set oFunc = nothing
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")

sub vbsShowTotals()
	
	if ilp_ID & "" = "" then		
		'dblClassCharge = "0.00" 
	end if
	
	if ilpShortID & "" = "" then
		'dblClassBudget = "0.00"
	end if 
	mDivCount = mDivCount + 1
	strBList = strBList & mDivCount & ","
%>
				<tr class=svplain8 bgcolor="<% = strColor%>" id="div<%=mDivCount%>">
					<td colspan="10" align=right>					
						&nbsp;<b>Course Totals:</b>
					</td>
					<td bgcolor=white>
						&nbsp;&nbsp;&nbsp;
					</td>
					<td class=TableHeader align=right>
						<nobr>
						<% if instr(1,dblBudgetCost,"-") > 0 then
								response.Write "+ $" & formatNumber(replace(dblClassBudget,"-",""),2)
						   else
								response.Write "- $" & formatNumber(dblClassBudget,2)
						   end if						
						%></nobr>
					</td>
					<td class=TableHeader align=right>
						<nobr>
						<% if instr(1,dblActualCost,"-") > 0 then
								response.Write "+ $" & formatNumber(replace(dblClassCharge,"-",""),2)
						   else
								response.Write "- $" & formatNumber(dblClassCharge,2)
						   end if						
						%></nobr>
					</td>
				</tr>	
<%
	
%>				
				<tr bgcolor=white >
					<td colspan="10">
						&nbsp;
					</td>
					<td  >
						&nbsp;
					</td>
					<td  >
						&nbsp;
					</td>
				</tr>	
<%
	dblTotalCharge = cdbl(dblTotalCharge) + cdbl(dblClassCharge)
	dblTotalBudget = cdbl(dblTotalBudget) + cdbl(dblClassBudget)
	dblClassBudget = 0 
	dblClassCharge = 0 
end sub

sub vbsAlternateColor()
	'Set alternating row color
	if intColor mod 2 = 0 then
		strColor = "white"
	else
		strColor="f7f7f7"
	end if
	intColor = intColor + 1
end sub

function LineItemInfo(pOrderedID,pBudget,pClosed,pCn,pCellClass)
	' Checks for line item entries and returns the following array if they exist...
	' ar(0) = html table of all line items
	' ar(1) = Total amount Charged (sum of all line items)
	' ar(2) = Budget Adjustment (deifined if budget is closed or is negative)
	' ar(3) = Div List"  Table row id's used to hide or show line item html row
	' ar(4) = If true Line Items do exist else no line items exist
	dim sql
	dim tCharged
	dim tBudget
	dim sHtml
	dim rs
	dim dAdjust
	dim strDivList
	dim strClosed
	dim bolLineItem
	
	tCharged = 0
	dAdjust = 0
	bolLineItem = false
	
	if pClosed then 
		strClosed = "Budget is Closed"
	else
		strClosed = "Budget is Open"
	end if
	
	sql = "SELECT intLine_Item_ID, dtLine_Item, szLine_Item_desc, curUnit_Price, intQuantity, curShipping, " & _ 
			" (curUnit_Price * intQuantity) + curShipping as Total, dtCREATE " & _
			"FROM tblLine_Items " & _ 
			"WHERE (intOrdered_Item_ID = " & pOrderedID & ") " & _
			" Order by intLine_Item_ID "
				
	set rs = server.CreateObject("ADODB.RECORDSET")
	rs.CursorLocation = 3
	rs.Open sql, pCn
	
	do while not rs.EOF
		bolLineItem = true
		mDivCount = mDivCount + 1
		strDivList = strDivList & mDivCount & ","
		tCharged = tCharged + formatNumber(rs("Total"),2)		
		sHtml = sHtml & "<tr id='div" & mDivCount & "' style='display:none;'>" & _
				"<td>&nbsp;</td><td colspan='2'  class='TableCellContrast'>Entered: " & formatDateTime(rs("dtCREATE"),2) & "</td>" & _
				"<td class='TableCellContrast' >" & rs("szLine_Item_desc") & "</td>" & _
				"<td class='TableCellContrast' align='center' valign='middle'>" & rs("intQuantity") & "</td>" & _
				"<td class='TableCellContrast' align='right' valign='middle'>$" & formatNumber(rs("curUnit_Price"),2) & "&nbsp;</td>" & _
				"<td class='TableCellContrast' align='right' valign='middle'>$" & formatNumber(rs("curShipping"),2) & "</td><td class='TableCellRed' >&nbsp;</td>" & _
				"<td class='TableCellContrast' align='right' valign='middle'>$" & formatNumber(rs("Total"),2) & "</td>" & _
				"<td colspan='3'>&nbsp;</td><td colspan='2' class='TableCellContrast' align='center'>" & strClosed & "</td></tr>" 
		rs.MoveNext		
	loop	
	rs.Close
	set rs = nothing
	
	tBudget = pBudget - tCharged
	if tBudget < 0 or pClosed then
		dAdjust = tBudget * -1
	end if
	
	dim ar(4)
	ar(0) = sHtml
	ar(1) = formatNumber(tCharged,2)
	ar(2) = formatNumber(dAdjust,2)
	ar(3) = strDivList
	ar(4) = bolLineItem
	LineItemInfo = ar
end function

sub vbsDelete(id)
	'Deletes existing Short Form
	dim delete	
	
	' This check keeps us from requesting to delete records that do not exist.
	' This is possible if a request to delete a course had been made and 
	' a user refreshes the browser. This will cause the request to delete
	' a course to be recalled because of the information in the http header (querystring)
	' We don't want this to happen because we don't want to send a message to the user
	' that a course has been deleted when it has not. 
	set rs = server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	sql = "Select * from tblILP_Short_Form where intShort_ILP_ID = " & id
	rs.Open sql,oFunc.FPCScnn
	
	if rs.RecordCount > 0 then
		oFunc.BeginTransCN
		' Delete Budget records first
		delete = "delete from tblBudget " & _
				"WHERE intShort_ILP_ID = " & id 
		oFunc.ExecuteCN(delete)
		
		' Now delete the Short Form
		delete = "delete from tblILP_Short_Form " & _
				"WHERE intShort_ILP_ID = " & id & _ 
				"AND intStudent_ID = " & intStudent_ID
		oFunc.ExecuteCN(delete)	 
		
		oFunc.CommitTransCN
		strMessage = "Course Deleted"
	end if 
end sub

sub vbsUpdateEnrollPercent(percent,ID)
	update = "update tblEnroll_Info set " & _
		     "intPercent_Enrolled_Fpcs = " & percent & ", " & _
		     "szUser_Modify = '" & session.Contents("strUserID") & "'," & _
		     "dtModify = '" & now() & "' " & _
		     "Where intEnroll_Info_ID = " & ID
	oFunc.ExecuteCN(update)
end sub

sub vbsUpdateILPStatus(pstrILPList)
	arList = split(pstrILPList,",")	
	if isArray(arList) then
		for i = 0 to ubound(arList)
			if arList(i) <> "" then
				call vbsApprovedStatus(arList(i),request("bolApproved" & arList(i)))
				call vbsUpdateComments(request("szComments" & arList(i)),arList(i))
			end if
		next
	end if
end sub

sub vbsApprovedStatus(ilp_id,bolApproved)
	dim strResetAdmin
	' Sets the Approval status for a specific ILP
	if bolApproved = "ready for review" then
		update = "update tblILP set bolApproved=Null, bolSponsor_Approved=NULL, " & _
				 "bolReady_For_Review = 1,dtReady_For_Review = CURRENT_TIMESTAMP " & _
				 " Where intILP_ID = " & ilp_ID
	elseif bolApproved = "implemented" then
		update = "update tblILP set bolApproved=Null, bolSponsor_Approved=NULL, " & _
				 "bolReady_For_Review = null  " & _
				 " Where intILP_ID = " & ilp_ID
	else
		if instr(1,bolApproved,"s-") > 0 then 
			strApproved = "bolSponsor_Approved"
			strDateField = "dtSponsor_Approved"
			if ucase(session.Contents("strRole")) = "ADMIN" then
				strResetAdmin = " ,bolApproved = NULL , dtApproved = CURRENT_TIMESTAMP "			
			end if
		else
			strResetAdmin = ""
		end if
		
		if bolApproved = "" or bolApproved = "implemented" then bolApproved = "Null" 
		if instr(1,bolApproved,"appr") > 0 then bolApproved = 1
		if instr(1,bolApproved,"must fix") > 0 then bolApproved = 0
			
		update = "update tblILP set " & strApproved & " = " & bolApproved & _
				", " & strDateField & " = CURRENT_TIMESTAMP " &  strResetAdmin & _
				" Where intILP_ID = " & ilp_ID
	end if
    oFunc.ExecuteCN(update)
end sub

sub vbsUpdateComments(comments,ilp_id)
	' Updates comments that can only be made by an Admin
	update = "update tblILP set " & strCommentField & " = '" & oFunc.EscapeTick(replace(comments,"""","''")) & "' " & _
			 " where intILP_ID = " & ilp_ID  
	oFunc.ExecuteCN(update)			 
end sub

sub vbsUpdateTestForm(pTF,pEnrollID)
	if pEnrollID <> "" then
		if pTF <> "false" then
			pTF = 1
		else
			pTF = 0 
		end if
		
		update = "update tblEnroll_Info set " & _
				"bolASD_Testing = " & pTF & ", " & _
				"szUser_Modify = '" & session.Contents("strUserID") & "'," & _
				"dtModify = '" & now() & "' " & _
				"Where intEnroll_Info_ID = " & pEnrollID
		oFunc.ExecuteCN(update)	
	end if
end sub

sub vbsUpdateProgressForm(pTF,pEnrollID)
	if pEnrollID <> "" then
		if pTF <> "false" then
			pTF = 1
		else
			pTF = 0 
		end if
		
		update = "update tblEnroll_Info set " & _
				"bolProgress_Agreement = " & pTF & ", " & _
				"szUser_Modify = '" & session.Contents("strUserID") & "'," & _
				"dtModify = '" & now() & "' " & _
				"Where intEnroll_Info_ID = " & pEnrollID
		oFunc.ExecuteCN(update)	
	end if
end sub

sub vbsLockEnrollLevel(pintEnrollID,pintPercent)
	' Locks student funding so it can not excede level defined by pintPercent
	if pintEnrollID <> "" and pintPercent <> "" then
		update = "update tblEnroll_Info set intPercent_Enrolled_Locked=" & pintPercent & _
				 " where intEnroll_Info_ID = " & pintEnrollID
		oFunc.ExecuteCN(update)
	end if
end sub

function GetBudgetStatus(pItemGroup,pBappr,pBolLineItems,pIsReimburse)	
	if pBappr & "" = "" and pBolLineItems = false then
		GetBudgetStatus = "pend"
	elseif pBappr = false then
		GetBudgetStatus = "rejc"
	elseif pBolLineItems = true and pIsReimburse = false and pItemGroup = 2 then
		GetBudgetStatus = "pick up"
	elseif pBolLineItems = true and pIsReimburse = false and pItemGroup = 1 then
		GetBudgetStatus = "pymt made"
	elseif pBappr = true and pBolLineItems = false and pIsReimburse = false and pItemGroup = 2 then
		GetBudgetStatus = "ordered"
	elseif pBappr = true and pBolLineItems = false and pIsReimburse = false and pItemGroup = 1 then
		GetBudgetStatus = "vend appr"
	elseif pBappr = true and pBolLineItems = true and pIsReimburse = true then
		GetBudgetStatus = "check cut"
	end if
end function
%>