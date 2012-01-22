<%@ Language=VBScript %>
<%
'JD Edit: In class time and planning time are shared collectively among the students. Calculate In Class Time Total, Planning Time Total, and Total Teacher Hours Total as such.

if session.Contents("strRole") <> "ADMIN" and  session.Contents("strRole") <> "TEACHER" then
	response.Write "<h1>Improper Request</h1>"
	response.End
end if

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

dim sql
dim strInfo				'contains Teacher info for mouse over display
dim sqlTeacher
dim intILPCount			'Number of ILP's per Class 
dim intTotalPlanning
dim intTotalInstruction
dim dblTotalHrs
dim bolTR				'toggles printing an html table <tr> tag
dim intCurrentTeacher	'tracks whether we have changed teachers
dim bolNoPayData	
dim dblInClassTimeTotal 
dim dblPlanningTimeTotal 
dim dblTotalTeacherHrsTotal	
'JD
dim dblInClassTimeSlice
dim dblPlanningTimeSlice
'JD
dim dblStudentChargedHrs
dim dblWages
dim dblTRS
dim dblPERS
dim dblFICA
dim dblMedicare
dim dblHealth
dim dblWorkComp
dim dblTotal
dim dblLifeInsurance
dim dblUnemployment
dim dblStudentChargedHrsTotal
dim dblWagesTotal
dim dblTRSTotal
dim dblPERSTotal
dim dblFICATotal
dim dblMedicareTotal
dim dblHealthTotal
dim dblWorkCompTotal
dim dblLifeInsuranceTotal
dim dblUnemploymentTotal
dim dblTotalTotal
dim strWhere
dim strTeachers
dim bolNoShow
dim dblALLInClassTimeTotal
dim dblALLPlanningTimeTotal
dim dblALLTotalTeacherHrsTotal
dim dblALLStudentChargedHrsTotal
dim dblALLWagesTotal
dim dblALLTRSTotal
dim dblALLPERSTotal
dim dblALLFICATotal
dim dblALLMedicareTotal
dim dblALLHealthTotal 
dim dblALLWorkCompTotal
dim dblALLLifeInsuranceTotal
dim dblALLUnemploymentTotal 
dim dblALLTotalTotal
dim fltTRS ,fltMedicare,fltWorkmans_Comp,fltPERS ,curHealth_Cost 
dim fltFICA,fltUnemployment,curLife_Insurance,curFICA_Cap
dim intTERS_Base_Percent,intPERS_Base_Percent 
'JD
dim flatRate
dim dblPlannedCost
dim dblPlannedCostTotal
dim dblAllPlannedCostTotal

bolNoShow = false

if session.Contents("strRole") = "ADMIN" and request("intInstructor_ID") = "" then
	strTeachers = Request("intInstructor_ids")
elseif session.Contents("instruct_id") <> "" then
	' User is logged in as a teacher and will only be allowed to see their own report
	strTeachers = session.Contents("instruct_id")
elseif session.Contents("strRole") = "ADMIN" and request("intInstructor_ID") <> "" then
	strTeachers = request("intInstructor_ID")
end if 

Session.Value("strTitle") = "Teacher Payroll Report"
Session.Value("strLastUpdate") = "17 June 2002"
if Request("chkFormat") = "View in Excel" then	
	Response.ContentType = "application/x-msexcel"	
	Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")
else	
	Server.Execute(Application.Value("strWebRoot") & "includes/Header.asp")
	call vbfSelectTeacher
end if

strClass = "svplain8"
strHdrClass = "TableHeader"

if strTeachers <> "" then
	'If we have a list of teachers coming from the header we break them up and 
	'dynamicaly create the where clause
	strTeachers = replace(strTeachers," ","")
	if instr(1,strTeachers,",") > 0 then
		arTeacherList = split(strTeachers,",")
		strWhere = " where intInstructor_ID = '" & arTeacherList(0) & "' "
		
		for w = 1 to ubound(arTeacherList)
			strWhere = strWhere & " or intInstructor_ID = '" & arTeacherList(w) & "' "
		next 
		
	elseif strTeachers = "all" then
		strWhere = "WHERE EXISTS " & _ 
								" (SELECT intClass_id " & _ 
								" FROM tblClasses c " & _ 
								" WHERE c.intInstructor_Id = i.intInstructor_ID AND c.intSchool_Year = " & session.Contents("intSchool_Year") & ") " 
	else
		' Only a single selection was made
		strWhere = " where intInstructor_ID = '" & strTeachers & "' "
	end if
	
	
	

	set rsReport = server.CreateObject("ADODB.RECORDSET")
	rsReport.CursorLocation = 3

	sql= "SELECT i.szLast_Name,i.szFirst_Name,i.intInstructor_id, " & _
		 "i.szEmail,i.szHome_Phone,i.szBusiness_Phone, pt.szPay_Type_Name, i.dtCert_Expire " & _
		 "from  tblINSTRUCTOR i LEFT OUTER JOIN " & _
         "trefPay_Types pt ON i.intPay_Type_id = pt.intPay_Type_ID " & _
		 strWhere & _
		 "order by szLast_Name "     		 
		     
	rsReport.Open sql,oFunc.FPCScnn
	
    'JD: Get the flat rate
    if Session.Contents("intSchool_Year") => 2012 then
	    set teacherFlatRate = server.CreateObject("ADODB.RECORDSET")
	    teacherFlatRate.CursorLocation = 3
	    sql = "SELECT flatRate from tblInstructor_Flat_Rate where intSchool_year = " & session.Contents("intSchool_year")
	    teacherFlatRate.Open sql, oFunc.FPCScnn
	    flatRate = teacherFlatRate("flatRate")
	    teacherFlatRate.Close()
	end if
		
	if rsReport.RecordCount > 0 then
		set rsClasses = server.CreateObject("ADODB.RECORDSET")
		rsClasses.CursorLocation = 3
		set rsILP = server.CreateObject("ADODB.RECORDSET")
		rsILP.CursorLocation = 3
		
		sql = "SELECT fltTRS, fltMedicare, fltWorkmans_Comp, fltPERS, curHealth_Cost, fltFICA,  " & _ 
				" fltUnemployment, curLife_Insurance, curFICA_Cap, intTERS_Base_Percent,  " & _ 
				" intPERS_Base_Percent " & _ 
				"FROM tblBenefit_Tax_Rates " & _ 
				"WHERE (intSchool_Year = " & session.Contents("intSchool_Year") & ") "
				
		rsClasses.Open sql, oFunc.FPCScnn
		
		
		if rsClasses.RecordCount > 0 then
			fltTRS = rsClasses("fltTRS")
			fltMedicare = rsClasses("fltMedicare")
			fltWorkmans_Comp = rsClasses("fltWorkmans_Comp")
			fltPERS = rsClasses("fltPERS")
			curHealth_Cost = rsClasses("curHealth_Cost")
			fltFICA = rsClasses("fltFICA")
			fltUnemployment = rsClasses("fltUnemployment")
			curLife_Insurance = rsClasses("curLife_Insurance")
			curFICA_Cap = rsClasses("curFICA_Cap")
			intTERS_Base_Percent = rsClasses("intTERS_Base_Percent")
			intPERS_Base_Percent = rsClasses("intPERS_Base_Percent")
				
		else
			response.Write "<h3>The tax and benefit figures have not been entered for this school year.<BR>" & _
						   "These numbers must be entered by the Business Manager or Principal before this report can be run.</h3>"
			response.End
		end if		
		rsClasses.Close
		intCurrentTeacher = -1
		do while not rsReport.EOF	
			set oTeacher = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/TeacherInfo.wsc"))
			oTeacher.PopulateObject oFunc.FPCScnn, rsReport("intInstructor_ID"), session.Contents("intSchool_Year")
				
			' Get all the classes for a specific teacher								
			sql = "select intClass_ID, szClass_Name,decHours_Student,decHours_Planning,intMin_Students, intMax_Students " & _
				  "from tblClasses  " & _
				  "where intInstructor_ID = " & rsReport("intInstructor_ID") & _
				  " and (intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _
				  " order by szClass_Name"
			
			rsClasses.Open sql,oFunc.FPCScnn							
			
			if rsClasses.RecordCount > 0 then														
			
				if intCurrentTeacher <> rsReport("intInstructor_ID") then
					'Print Table header
					if request("totals") = "" or bolNoShow = false then						
						call vbfHeaderRow(oTeacher.BaseHourlyRate)						
					end if
					if request("totals") <> "" then
						bolNoShow = true
					end if 
					'Teacher Contact Info
					strInfo = "Home Phone: " & 	rsReport("szHome_Phone") & _
							  " Work Phone: " & rsReport("szBusiness_Phone")
					'Prints row with teachers name and hourly rate			
					'if isArray(arCosts) then
					'	call vbfPrintTeacher(1,arCosts(0))	
					'else
					'	call vbfPrintTeacher(1,0)
					'	bolNoPayData = true
					'end if 
				end if 		
				intCurrentTeacher = rsReport("intInstructor_ID")
				
				if bolNoPayData = false then
					do while not rsClasses.EOF					
						sql = "select s.intStudent_id, s.szFirst_Name,s.szLast_Name, " & _
							  "i.intILP_ID " & _
							  "from tblStudent s, tblILP i " &  _
							  "where i.intClass_ID = " & rsClasses("intClass_ID") & _
							  " and i.intStudent_id = s.intStudent_ID " & _	
							  " order by s.szLast_Name " 
						rsILP.Open sql,oFunc.FPCScnn	
						intILPCount = rsILP.RecordCount
						bolTR = false
						if intILPCount > 0 then
							if request("totals") = "" then
								call vbfPrintClass(intILPCount)
							end if
							
							do while not rsILP.EOF	
								
								dblStudentChargedHrs = formatNumber((cDBL(rsClasses("decHours_Student"))/cdbl(intILPCount)) + (cDBL(rsClasses("decHours_Planning"))/cdbl(intILPCount)),3)
								'JD
                                dblInClassTimeSlice = formatNumber((cDBL(rsClasses("decHours_Student"))/cdbl(intILPCount)), 3)
                                dblPlanningTimeSlice = formatNumber((cDBL(rsClasses("decHours_Planning"))/cdbl(intILPCount)), 3)
                                'JD

								 		
								if request("totals") = "" then
									if bolTR = true then Response.Write "<TR>"
									bolTR = true
								end if
								
								if intILPCount < rsClasses("intMin_Students")  then
									if request("totals") =  "" then
%>	
		<td class=<% = strClass %> >
			<% = rsILP("szLast_Name") & ", " & rsILP("szFirst_Name") %>
		</td>
		<td class=<% = strClass %> colspan="14" >	
			<span style="color:red;"><b>Minimum Class Enrollment of <% =rsClasses("intMin_Students") %> has not been met. Currently <% =intILPCount %> student(s) enrolled.</b></span>
		</td>
	</tr>
<%			
								end if
								else		
									dblInClassTime = cdbl(formatNumber(rsClasses("decHours_Student"),2))
									dblPlanningTime = cdbl(formatNumber(rsClasses("decHours_Planning"),2))
									dblTotalTeacherHrs = cdbl(formatNumber(cDBL(rsClasses("decHours_Student")) + cDBL(rsClasses("decHours_Planning")),2))
									dblStudentChargedHrs = cdbl(dblStudentChargedHrs)
                                    'JD
                                    dblInClassTimeSlice = cdbl(dblInClassTimeSlice)
                                    dblPlanningTimeSlice = cdbl(dblPlanningTimeSlice)
                                    dblPlannedCost = cdbl(formatNumber(flatRate * dblStudentChargedHrs, 2))
									'JD
									dblWages = cdbl(formatNumber(oTeacher.BaseHourlyRate * dblStudentChargedHrs,2))
									dblTRS = cdbl(formatNumber(oTeacher.TersCostPerHour * dblStudentChargedHrs,2))
									dblPERS = cdbl(formatNumber(oTeacher.PersCostPerHour * dblStudentChargedHrs,2))
									dblFICA = cdbl(formatNumber(oTeacher.FicaCostPerHour * dblStudentChargedHrs,2))
									dblMedicare = cdbl(formatNumber(oTeacher.MedicareCostPerHour * dblStudentChargedHrs,2))
									dblHealth = cdbl(formatNumber(oTeacher.HealthInsuranceCostPerHour * dblStudentChargedHrs,2))
									dblWorkComp = cdbl(formatNumber(oTeacher.WorkersCompCostPerHour * dblStudentChargedHrs,2))
									dblLifeInsurance = cdbl(formatNumber(oTeacher.LifeInsuranceCostPerHour * dblStudentChargedHrs,2))
									dblUnemployment = cdbl(formatNumber(oTeacher.UnemploymentCostPerHour * dblStudentChargedHrs,2))
									dblTotal = dblWages + dblTRS + dblPERS + dblFICA + dblMedicare + dblHealth + dblWorkComp + dblLifeInsurance + dblUnemployment
									
									'JD
									'dblInClassTimeTotal = dblInClassTimeTotal + dblInClassTime
									'dblPlanningTimeTotal = dblPlanningTimeTotal + dblPlanningTime
									'dblTotalTeacherHrsTotal = dblTotalTeacherHrsTotal + dblTotalTeacherHrs
							        dblInClassTimeTotal = dblInClassTimeTotal + dblInClassTimeSlice
									dblPlanningTimeTotal = dblPlanningTimeTotal + dblPlanningTimeSlice
									dblTotalTeacherHrsTotal = dblTotalTeacherHrsTotal + dblStudentChargedHrs
									dblPlannedCostTotal = dblPlannedCostTotal + dblPlannedCost
                                    'JD
									dblStudentChargedHrsTotal = dblStudentChargedHrsTotal + dblStudentChargedHrs
									dblWagesTotal = dblWagesTotal + dblWages
									dblTRSTotal = dblTRSTotal + dblTRS
									dblPERSTotal = dblPERSTotal + dblPERS
									dblFICATotal = dblFICATotal + dblFICA
									dblMedicareTotal = dblMedicareTotal + dblMedicare
									dblHealthTotal = dblHealthTotal + dblHealth
									dblWorkCompTotal = dblWorkCompTotal + dblWorkComp
									dblLifeInsuranceTotal = dblLifeInsuranceTotal + dblLifeInsurance
									dblUnemploymentTotal = formatNumber(dblUnemploymentTotal + dblUnemployment,2)
									dblTotalTotal = dblTotalTotal + dblTotal
								end if
								
								if request("totals") = "" and intILPCount >= rsClasses("intMin_Students")then											
%>	
		<td class=<% = strClass %> >
			<% = rsILP("szLast_Name") & ", " & rsILP("szFirst_Name") %>
		</td>
		<td class=<% = strClass %> align=right>
			<% = dblInClassTime %>
		</td>
		<td class=<% = strClass %> align=right>
			<% = dblPlanningTime %>
		</td>											
		<td class=<% = strClass %> align=right>
			<% = dblTotalTeacherHrs %> 
		</td>
		<td class=<% = strClass %> align=right>
			<% = dblStudentChargedHrs %> 
		</td>
		<td class=<% = strClass %> align=right>
			<% = dblWages %> 
		</td>
		<%'JD flatrate * studentchargedhrs %>
		<td class=<% = strClass %> align=right>
			<% if Session.Contents("intSchool_Year") < 2012 then%>
                 n/a
             <%else%>
		        <% = dblPlannedCost %> 
			<% end if%>
		</td>

		<td class=<% = strClass %> align=right>
			<% = dblTRS %> 
		</td>
		<td class=<% = strClass %> align=right>
			<% = dblPERS %>
		</td>
		<td class=<% = strClass %> align=right>
			<% = dblFICA %> 
		</td>
		<td class=<% = strClass %> align=right>
			<% = dblMedicare %>
		</td>
		<td class=<% = strClass %> align=right>
			<% = dblHealth %>
		</td>
		<td class=<% = strClass %> align=right>
			<% = dblLifeInsurance %>
		</td>
		<td class=<% = strClass %> align=right>
			<% = dblWorkComp %> 
		</td>
		<td class=<% = strClass %> align=right>
			<% = dblUnemployment %> 
		</td>
		<td class=<% = strClass %> align=right>
			<% = dblTotal %> 
		</td>
	</tr>
<% 				
								end if
								rsILP.MoveNext						
						loop  ' do while not rsILP.EOF
						rsILP.Close
							
					else
						if intCurrentTeacher <> rsReport("intInstructor_ID") then
							call vbfPrintTeacher(100,"")	
							intCurrentTeacher = rsReport("intInstructor_ID")
						end if
						if request("totals") = "" then
							vbfPrintClass(1)
						
%>		
		<Td class="TableCell" colspan=16>
			&nbsp;No Students Enrolled  in this Class.
		</td>					
	</tr>
<%		
						end if 
						rsILP.Close			
					end if 'intILPCount > 0 							
																					
					rsClasses.MoveNext
					Response.Write strError
				loop	'do while not rsClasses.EOF								
			end if	'if bolNoPayData = false		
			call vbfPrintTotals
			bolNoPayData = false
		else
			if request("totals") = "" then						
				response.Write "<font class='svplain8'><B>" & oTeacher.FirstName & " " & oTeacher.LastName & " has not created any " & _
							"classes for the " & session.Contents("intSchool_Year") & _
							" school year.</b></font>"
			end if
		end if	
	rsClasses.Close
	set oTeacher = nothing
	rsReport.MoveNext		
	intTotalPlanning = 0 
	intTotalInstruction = 0 

		if request("totals") = "" then	
%>
</table>
<p>
<%	
		end if 
	loop

	rsReport.Close
	set rsReport = nothing
	set rsClasses = nothing
	set rsILP = nothing
	if Request("chkFormat") = "" then
%>		

<script language=javascript>
	function jfGetProfile(id){
		var winProfile;
		winProfile = window.open("../forms/Teachers/addTeacher.asp?bolWin=True&intInstructor_ID="+id,"winProfile","width=800,height=550,scrollbars=yes");
		winProfile.focus();
		winProfile.moveTo(0,0);
	}
</script>
<%
	end if 
	
	if request("totals") <> "" then
		call vbfPrintAllTotals()
	end if 
%>
</body>
</html>
<%
	else
%>
<html>
<head>
<title></title>
<link rel="stylesheet" href="<% = strPath %>/CSS/homestyle.css">
</head>
<body background=c0c0c0>
<form id=form1 name=form1>
<table width=100% height=100%>
	<tr>
		<Td align=center valign=middle>
			<table>
				<tr>
					<Td class=<% = strClass %>>
						There are currenlty 0 Teachers in the Teacher Inforamtion System (TIS).<br><BR>
						<center>
						<input type=button value="< Back" onCLick="window.location.href='<%=Application.Value("strWebRoot")%>';"  class="btSmallGray" name=button2>
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
end if 

Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")
call oFunc.CloseCN()
set oFunc = nothing

function vbfHeaderRow(pCost)
	'This function prints the column headers for each teachers payroll report
%>
<br>
<% if request("totals") = "" then 
		if isDate(rsReport("dtCert_Expire")) then
			if cdate(rsReport("dtCert_Expire")) < now() then
				strCert = "<font color='red'><b>Certificate Expiration:</b> " & rsReport("dtCert_Expire") & "</font>"
			else
				strCert = "<b>Certificate Expiration:</b> " & rsReport("dtCert_Expire") 
			end if 
		else
			strCert = "<font color='red'><b>Certificate Expiration:</b> " & rsReport("dtCert_Expire") & "</font>"
		end if
%>
<span title="<% = strInfo %>" class="svplain8">
<b><a href="javascript:" onClick="jfGetProfile('<%=rsReport("intInstructor_id")%>');">
<% = rsReport("szLast_Name") & ", " & rsReport("szFirst_Name")%></a></b>
&nbsp;&nbsp; <% if Session.Contents("intSchool_Year") => 2012 then %><b>Flat Rate: $<%=formatNumber(flatRate, 2) %></b><%end if %>&nbsp;&nbsp;<b>Base Rate Per Hour: $<%=pCost%></b> &nbsp;&nbsp;<b>Pay Type:</b> <% = rsReport("szPay_Type_Name")%>&nbsp;&nbsp;<% = strCert %></span>
<% end if %>
<table  cellpadding=3 cellspacing=0 border=1 bordercolor="#e6e6e6">
	<tr>		
		<% if request("totals") = "" then %>
		<td class=<% = strHdrClass %> valign=top align=center>
			Class
		</td>
		<td class=<% = strHdrClass %> valign=top align=center>
			Student
		</td>
		<td class=<% = strHdrClass %> valign=top align=center>
			In-Class<br>Time
		</td>
		<td class=<% = strHdrClass %> valign=top align=center>
			Plan<BR>Time
		</td>
		<td class=<% = strHdrClass %> valign=top align=center>
			Total<br>Hrs
		</td>
		<td class=<% = strHdrClass %> valign=top align=center>
			Student Charged <br/> Hrs
		</td>	
		<% else %>
		<td class=<% = strHdrClass %> valign=top align=center>
			Teacher 
		</td>
		<td class=<% = strHdrClass %> valign=top align=center>
			Pay Type
		</td>
		<% end if %>
		<td class=<% = strHdrClass %> valign=top align=center>
			Teacher <br/>Wage
		</td>
		<%'Added column for flatRate * hrs %>
		<td class=<% = strHdrClass %> valign=top align=center>
		    Flat Rate Cost <br />
		    or<br />
		    Student Cost
		</td>
		<td class=<% = strHdrClass %> valign=top align=center>
			TRS
			<br>
			<% = fltTRS %>
		</td>
		<td class=<% = strHdrClass %> valign=top align=center>
			PERS
			<br>
			<% = fltPERS %>
		</td>
		<td class=<% = strHdrClass %> valign=top align=center>
			FICA
			<br>
			<% = fltFICA %>
		</td>
		<td class=<% = strHdrClass %> valign=top align=center>
			Medicare
			<br>
			<% = fltMedicare %>
		</td>
		<td class=<% = strHdrClass %> valign=top align=center>
			Health Ins.
			<br>
			<% = curHealth_Cost %>
		</td>
		<td class=<% = strHdrClass %> valign=top align=center>
			Life Ins.
			<br>
			<% = curLife_Insurance %>
		</td>
		<td class=<% = strHdrClass %> valign=top align=center>
			W/C
			<br>
			<% = fltWorkmans_Comp %>
		</td>		
		<td class=<% = strHdrClass %> valign=top align=center>
			Unemplmt
			<br>
			<% = fltUnemployment %>
		</td>
		<td class=<% = strHdrClass %> valign=top align=center>
			FPCS Cost<br /> or Teacher Wage <br /> with Benefits
			
		</td>
	</tr>
<%
end function

function vbfPrintTeacher(rows,hourlyRate)
	' Just prints Teachers Name and takes 2 parameters first one determines how many
	' HTML table rows to span second gives us the teachers hourly rate
	if request("totals") = "" then
%>			
	<tr>						
		<Td class="TableCell" rowspan=<%=rows%> valign=top>
			<span title="<% = strInfo %>">
			<a href="javascript:" onClick="jfGetProfile('<%=rsReport("intInstructor_id")%>');">
			<% = rsReport("szLast_Name") & ", " & rsReport("szFirst_Name")%></a>
			&nbsp;&nbsp;<b>Per Diem: </b></span>
		</td>	
		<td>
		</td>		
		<td>
		</td>
		<td>
		</td>
		<td>
		</td>
		<td>
		</td>
		<td>
		</td>
		<td class=svplain10 rowspan=<%=rows%> valign=top>
			<%=hourlyRate%>
		</td>	
		<td>
		</td>
		<td>
		</td>
		<td>
		</td>
		<td>
		</td>
		<td>
		</td>
		<td>
		</td>
		<td>
		</td>
		<td>
		</td>	
		<td>
		</td>
		<td>
		</td>	
	</tr>
<%
	else
%>
	<tr>						
		<Td class=gray12 valign=top>
			<span title="<% = strInfo %>">
			<a href="javascript:" onClick="jfGetProfile('<%=rsReport("intInstructor_id")%>');">
			<% = rsReport("szLast_Name") & ", " & rsReport("szFirst_Name")%></a></span>
		</td>	
<%
	end if
end function 

function vbfPrintClass(rows)
	' Just prints Class Name and takes a parameter that determines how many
	' HTML table rows to span
%>
	<tr>
		<td class="TableCell" rowspan=<%=rows%> valign=top>
			<% = rsClasses("szClass_Name") %>
		</td>	
<%
end function

function vbfPrintTotals
	if request("totals") = "" then 
%>	
	<tr>		
		<Td class="TableCell" valign=top colspan=2 align=right>
			<b>Totals:</b>
		</td>	
		<td class="TableCell" align=right>
			<% = formatNumber(dblInClassTimeTotal,2) %>
		</td>
		<td class="TableCell" align=right>
			<% = formatNumber(dblPlanningTimeTotal,2) %>
		</td>
		<td class="TableCell" align=right>
			<% = formatNumber(dblTotalTeacherHrsTotal,2) %>
		</td>
		<td class="TableCell" align=right>
			<% = formatNumber(dblStudentChargedHrsTotal,2) %>
		</td>
		<% else %>
		<td class="TableCell">
			<% = rsReport("szLast_Name") & ", " & rsReport("szFirst_Name")%>
		</td>
		<td class="TableCell" align=center>
			<% = rsReport("szPay_Type_Name") %>
		</td>
		<% end if %>
		<td class="TableCell" align=right>
			<% = formatNumber(dblWagesTotal,2) %> 
		</td>
		
		<%'Flat rate total %>
		<td class="TableCell" align=right>
		    <% = formatNumber(dblPlannedCostTotal, 2) %>
		</td>
		
		<td class="TableCell" align=right>
			<% = formatNumber(dblTRSTotal,2) %> 
		</td>
		<td class="TableCell" align=right>
			<% = formatNumber(dblPERSTotal,2) %> 
		</td>
		<td class="TableCell" align=right>
			<% = formatNumber(dblFICATotal,2) %> 
		</td>
		<td class="TableCell" align=right>
			<% = formatNumber(dblMedicareTotal,2) %> 
		</td>
		<td class="TableCell" align=right>
			<% = formatNumber(dblHealthTotal,2) %> 
		</td>
		<td class="TableCell" align=right>
			<% = formatNumber(dblLifeInsuranceTotal,2) %> 
		</td>
		<td class="TableCell" align=right>
			<% = formatNumber(dblWorkCompTotal,2) %> 
		</td>
		<td class="TableCell" align=right>
			<% = formatNumber(dblUnemploymentTotal,2) %>
		</td>
		<td class="TableCell" align=right>
			<% = formatNumber(dblTotalTotal,2) %>
		</td>	
	</tr>
<%
	if request("totals") <> "" then
		dblALLInClassTimeTotal = dblALLInClassTimeTotal + dblInClassTimeTotal
		dblALLPlanningTimeTotal = dblALLPlanningTimeTotal + dblPlanningTimeTotal
		dblALLTotalTeacherHrsTotal = dblALLTotalTeacherHrsTotal + dblTotalTeacherHrsTotal
		dblALLStudentChargedHrsTotal = dblALLStudentChargedHrsTotal + dblStudentChargedHrsTotal
		dblAllPlannedCostTotal = dblAllPlannedCostTotal + dblPlannedCostTotal
		dblALLWagesTotal = dblALLWagesTotal + dblWagesTotal
		dblALLTRSTotal = dblALLTRSTotal + dblTRSTotal
		dblALLPERSTotal = dblALLPERSTotal + dblPERSTotal
		dblALLFICATotal = dblALLFICATotal + dblFICATotal
		dblALLMedicareTotal = dblALLMedicareTotal + dblMedicareTotal
		dblALLHealthTotal = dblALLHealthTotal + dblHealthTotal
		dblALLWorkCompTotal = dblALLWorkCompTotal + dblWorkCompTotal
		dblALLLifeInsuranceTotal = dblALLLifeInsuranceTotal + dblLifeInsuranceTotal
		dblALLUnemploymentTotal = cdbl(dblALLUnemploymentTotal) + cdbl(dblUnemploymentTotal)
		dblALLTotalTotal = dblALLTotalTotal + dblTotalTotal
	end if
	
	dblInClassTimeTotal = 0
	dblPlanningTimeTotal = 0
	dblTotalTeacherHrsTotal = 0
	dblStudentChargedHrsTotal = 0
	dblPlannedCostTotal = 0
	dblWagesTotal = 0
	dblTRSTotal = 0
	dblPERSTotal = 0
	dblFICATotal = 0 
	dblMedicareTotal = 0
	dblHealthTotal = 0
	dblWorkCompTotal = 0
	dblLifeInsuranceTotal = 0 
	dblUnemploymentTotal = 0
	dblTotalTotal = 0
	
end function

function vbfPrintAllTotals
%>
	<tr>
		<td class="TableCell" align=right colspan=2>
			<b>Totals:</b>
		</td>
		<td class="TableCell" align=right>
			<% = formatNumber(dblALLWagesTotal,2) %> 
		</td>
		<%'Flat rate total %>
		<td class="TableCell" align=right>
		    <% = formatNumber(dblAllPlannedCostTotal, 2) %>
		</td>
		<td class="TableCell" align=right>
			<% = formatNumber(dblALLTRSTotal,2) %> 
		</td>
		<td class="TableCell" align=right>
			<% = formatNumber(dblALLPERSTotal,2) %> 
		</td>
		<td class="TableCell" align=right>
			<% = formatNumber(dblALLFICATotal,2) %> 
		</td>
		<td class="TableCell" align=right>
			<% = formatNumber(dblALLMedicareTotal,2) %> 
		</td>
		<td class="TableCell" align=right>
			<% = formatNumber(dblALLHealthTotal,2) %> 
		</td>
		<td class="TableCell" align=right>
			<% = formatNumber(dblALLLifeInsuranceTotal,2) %> 
		</td>
		<td class="TableCell" align=right>
			<% = formatNumber(dblALLWorkCompTotal,2) %> 
		</td>
		<td class="TableCell" align=right>
			<% = formatNumber(dblALLUnemploymentTotal,2) %>
		</td>
		<td class="TableCell" align=right>
			<% = formatNumber(dblALLTotalTotal,2) %>
		</td>	
	</tr>
</table>
<%
	
end function

function vbfSelectTeacher
%>
<table width=100%>	
	<% if session.Contents("strRole") = "ADMIN" then %>
	<form action=teacherPayRoll.asp method=get id=form2 name=form2>	
	<tr>	
		<Td class=yellowHeader colspan=2>
				&nbsp;<b>Teacher Payroll Report</b>	(includes only teachers with class contracts in the system)		
		</td>
	</tr>
	<tr>
		<td class=Gray valign=top >
			<nobr><b>Select Teacher(s)</b></nobr><br>
			&nbsp;<input type=submit value="View in  Html"  class="btSmallGray" >
			&nbsp;<input type=submit name=chkFormat value="View in Excel"  class="btSmallGray" >
			<br>
			<nobr>Show Totals Only:<input type=checkbox name="totals" value="true" <% if request("totals") <> "" then response.Write " checked "%>></nobr>
			
		</td>
		<td width=100%>
			<select name=intInstructor_ids multiple size=5>
				<option value="all">ALL TEACHERS
			<%
				dim sqlInstructor
				sqlInstructor = "SELECT intINSTRUCTOR_ID, szLAST_NAME + ',' + szFIRST_NAME AS Name " & _ 
								"FROM tblINSTRUCTOR i " & _ 
								"WHERE EXISTS " & _ 
								" (SELECT intClass_id " & _ 
								" FROM tblClasses c " & _ 
								" WHERE c.intInstructor_Id = i.intInstructor_ID AND c.intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _ 
								"ORDER BY szLAST_NAME "
				Response.Write oFunc.MakeListSQL(sqlInstructor,intStudent_ID,Name,strTeachers)	
			%>
			</select>
		</td>
	</tr>
	<% else %>
	<form action=teacherPayRoll.asp method=get id="Form3" name=form2 target="_new">	
	<tr>	
		<Td class=yellowHeader colspan=2>
				&nbsp;<b>Teacher Payroll Report</b>
				<input type=button value="Home Page" onCLick="window.location.href='<%=Application.Value("strWebRoot")%>';"  class="btSmallGray" name=button3>				
				&nbsp;<input type=submit name=chkFormat value="View in Excel" class="btSmallGray" >
		</td>
	</tr>
	<% end if %>
	</form>
</table>
<p>
<% 
end function
%>