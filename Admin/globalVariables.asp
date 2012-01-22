<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		globalVariables.asp
'Purpose:	Admin page to set various global variables that greatly impact 
'			the application
'Date:		17 Mar 2005
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim sql
dim oFunc
dim rs1,rs2,rs3, rs4 'JD: rs4 to store the instructor flat rate.
dim dtLock_Spending, dtSem_One_Progress_Deadline,dtSem_Two_Progress_Deadline, bolLock_School_Year
dim dtSchool_Year_Start, dtSchool_Year_End, dtCount_Deadline
' Security Check. Must be an Admin
if ucase(session.Contents("strRole")) <> "ADMIN" then
	response.Write "<H1>PAGE ILLEGALLY CALLED</H1>"
	response.End
end if

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

oFunc.CheckSuperAdmin 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Date Inserts and Update Logic							    ''

if request.Form("benefits") <> "" and request.Form("intBenefit_Tax_Rate_ID") <> "" then
	Call UpdateBenefits
elseif request.Form("benefits") <> "" then
	call InsertBenefits
end if 

if request.Form("Funding") <> "" then
	call UpdateFunding
end if 

if request.Form("bolLock_School_Year") <> "" then
	bolLock_School_Year = 1
else
	bolLock_School_Year = 0
end if
	
if request.Form("intGlobal_Variable_ID") <> "" and request.Form("Locks") <> "" then
	call UpdateGlobalVars
elseif request.Form("Locks") <> "" then
	call InsertGlobalVars
end if

'JD 041411 call update if value change
if request.Form("InstructorFlatRate") <> "" and Request.Form("intFlat_Inst_Id")<> "" then
    call UpdateInstructorFlatRate
elseif request.Form("InstructorFlatRate") <> "" then 
    call InsertInstructorFlatRate    
end if

if request.Form("pwd") <> "" then
	call UpdatePWD
end if

if Request.Form("HasUpdated") <> "" then call UpdateLockedStudentAccounts

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Sql Statements and Records Sets							    ''

sql =   "SELECT     intBenefit_Tax_Rate_ID, fltTRS, fltMedicare, fltWorkmans_Comp, fltPERS,  " & _ 
		" curHealth_Cost, fltFICA, fltUnemployment, curLife_Insurance,  " & _ 
		" curFICA_Cap, intSchool_Year,intTERS_Base_Percent, " & _
		" intPERS_Base_Percent,dtCREATE, dtMODIFY,  " & _ 
		" szUSER_CREATE, szUSER_MODIFY " & _ 
		"FROM tblBenefit_Tax_Rates " & _ 
		"WHERE (intSchool_Year = " & session.Contents("intSchool_Year") & ") "

set rs1 = server.CreateObject("ADODB.RECORDSET")
rs1.CursorLocation = 3
rs1.Open sql, oFunc.FPCScnn				

sql2 =  "SELECT     a.gradeK, a.intFunding_ID as IDk, b.grade1, b.intFunding_ID as ID1, c.grade2, c.intFunding_ID as ID2, " & _
		" d.grade3,  d.intFunding_ID as ID3, e.grade4, e.intFunding_ID as ID4, f.grade5, f.intFunding_ID as ID5, " & _
		"g.grade6, g.intFunding_ID as ID6, h.grade7, h.intFunding_ID as ID7, i.grade8, i.intFunding_ID as ID8, " & _
		"j.grade9, j.intFunding_ID as ID9, k.grade10, k.intFunding_ID as ID10, l.grade11, l.intFunding_ID as ID11, m.grade12, m.intFunding_ID as ID12 " & _ 
		"FROM         (SELECT     curFund_Amount AS gradeK, intFunding_ID " & _ 
		"                       FROM          tblFunding " & _ 
		"                       WHERE      (szGrade = 'K') AND (intSchool_Year = '" & session.Contents("intSchool_Year") & "')) a , " & _ 
		"                          (SELECT     curFund_Amount AS grade1, intFunding_ID " & _ 
		"                            FROM          tblFunding " & _ 
		"                            WHERE      (szGrade = '1') AND (intSchool_Year = '" & session.Contents("intSchool_Year") & "')) b , " & _ 
		"                          (SELECT     curFund_Amount AS grade2, intFunding_ID " & _ 
		"                            FROM          tblFunding " & _ 
		"                            WHERE      (szGrade = '2') AND (intSchool_Year = '" & session.Contents("intSchool_Year") & "')) c , " & _ 
		"                          (SELECT     curFund_Amount AS grade3, intFunding_ID " & _ 
		"                            FROM          tblFunding " & _ 
		"                            WHERE      (szGrade = '3') AND (intSchool_Year = '" & session.Contents("intSchool_Year") & "')) d , " & _ 
		"                          (SELECT     curFund_Amount AS grade4, intFunding_ID " & _ 
		"                            FROM          tblFunding " & _ 
		"                            WHERE      (szGrade = '4') AND (intSchool_Year = '" & session.Contents("intSchool_Year") & "')) e , " & _ 
		"                          (SELECT     curFund_Amount AS grade5, intFunding_ID " & _ 
		"                            FROM          tblFunding " & _ 
		"                            WHERE      (szGrade = '5') AND (intSchool_Year = '" & session.Contents("intSchool_Year") & "')) f , " & _ 
		"                          (SELECT     curFund_Amount AS grade6, intFunding_ID " & _ 
		"                            FROM          tblFunding " & _ 
		"                            WHERE      (szGrade = '6') AND (intSchool_Year = '" & session.Contents("intSchool_Year") & "')) g , " & _ 
		"                          (SELECT     curFund_Amount AS grade7, intFunding_ID " & _ 
		"                            FROM          tblFunding " & _ 
		"                            WHERE      (szGrade = '7') AND (intSchool_Year = '" & session.Contents("intSchool_Year") & "')) h , " & _ 
		"                          (SELECT     curFund_Amount AS grade8, intFunding_ID " & _ 
		"                            FROM          tblFunding " & _ 
		"                            WHERE      (szGrade = '8') AND (intSchool_Year = '" & session.Contents("intSchool_Year") & "')) i , " & _ 
		"                          (SELECT     curFund_Amount AS grade9, intFunding_ID " & _ 
		"                            FROM          tblFunding " & _ 
		"                            WHERE      (szGrade = '9') AND (intSchool_Year = '" & session.Contents("intSchool_Year") & "')) j , " & _ 
		"                          (SELECT     curFund_Amount AS grade10, intFunding_ID " & _ 
		"                            FROM          tblFunding " & _ 
		"                            WHERE      (szGrade = '10') AND (intSchool_Year = '" & session.Contents("intSchool_Year") & "')) k , " & _ 
		"                          (SELECT     curFund_Amount AS grade11, intFunding_ID " & _ 
		"                            FROM          tblFunding " & _ 
		"                            WHERE      (szGrade = '11') AND (intSchool_Year = '" & session.Contents("intSchool_Year") & "')) l , " & _ 
		"                          (SELECT     curFund_Amount AS grade12, intFunding_ID " & _ 
		"                            FROM          tblFunding " & _ 
		"                            WHERE      (szGrade = '12') AND (intSchool_Year = '" & session.Contents("intSchool_Year") & "')) m "		

set rs2 = server.CreateObject("ADODB.RECORDSET")
rs2.CursorLocation = 3
rs2.Open sql2, oFunc.FPCScnn				


sql3 =  "SELECT     intGlobal_Variable_ID, dtLock_Spending, dtSem_One_Progress_Deadline,  " & _ 
		" dtSem_Two_Progress_Deadline, bolLock_School_Year, intSchool_Year,  " & _ 
		" dtCREATE, dtMODIFY, szUSER_CREATE, szUSER_MODIFY,dtSchool_Year_Start,dtSchool_Year_End, " & _ 
		" (select dtSchool_Year_End from tblGlobal_Variables where intSchool_Year = " & session.Contents("intSchool_Year") - 1 & ") as MustStartAfter, " & _
		" (select dtSchool_Year_Start from tblGlobal_Variables where intSchool_Year = " & session.Contents("intSchool_Year") + 1 & ") as MustEndBefore " & _	
		" , dtCount_Deadline " & _	
		"FROM tblGlobal_Variables " & _ 
		"WHERE (intSchool_Year = " & session.Contents("intSchool_Year") & ") "

set rs3 = server.CreateObject("ADODB.RECORDSET")
rs3.CursorLocation = 3
rs3.Open sql3, oFunc.FPCScnn	
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

if rs3.RecordCount > 0 then
	MustStartAfter = rs3("MustStartAfter")
	MustEndBefore = rs3("MustEndBefore")
end if

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'JD 041411: create the recordset to get the flat instructor rate
sql4 ="select intFlat_Inst_Id, flatRate from tblInstructor_Flat_Rate where intSchool_year = " & session.Contents("intSchool_Year")
set rs4 = server.CreateObject("ADODB.RECORDSET")
rs4.CursorLocation = 3
rs4.Open sql4, oFunc.FPCScnn
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Print the header
Session.Value("strTitle") = "Global Settings Administration Page"
Session.Value("strLastUpdate") = "17 Mar 2005"
Server.Execute(Application.Value("strWebRoot") & "includes/Header.asp")		
%>
<script language=javascript src="<%= Application.Value("strWebRoot") %>includes/CalendarPopup.js"></script>	
<script language=javascript>
	var cal = new CalendarPopup('divCal');
	cal.showNavigationDropdowns();
	
	var cal2 = new CalendarPopup('divCal');
	cal2.showNavigationDropdowns();
	
	var cal3 = new CalendarPopup('divCal');
	cal3.showNavigationDropdowns();
	
	var cal4 = new CalendarPopup('divCal');
	cal4.showNavigationDropdowns();
	
	var cal5 = new CalendarPopup('divCal');
	cal5.showNavigationDropdowns();
	
	var calStart = new CalendarPopup('divCal');
	calStart.showNavigationDropdowns();
	
	var calEnd = new CalendarPopup('divCal');
	calEnd.showNavigationDropdowns();
	
	function jfValidate(pForm){
		var strError = "";
		var strItems = "";
		for (i=0; i< pForm.selRestrictedStudents.length; i++) {
			strItems = strItems + pForm.selRestrictedStudents.options[i].value + ",";
		}
		pForm.RestrictedStudents.value = strItems.substr(0, strItems.length - 1);
		pForm.HasUpdated.value = "true";
		
		if (pForm.benefits.value != "") {
			if (pForm.fltTRS.value == "" || pForm.fltMedicare.value == "" || pForm.fltWorkmans_Comp.value == "" ||
			    pForm.fltPERS.value == "" || pForm.curHealth_Cost.value == "" || pForm.fltFICA.value == "" ||
			    pForm.fltUnemployment.value == "" || pForm.curLife_Insurance.value == "" || pForm.curFICA_Cap.value == "" || 
			    pForm.intTERS_Base_Percent.value == "" || pForm.intPERS_Base_Percent.value == "" ) {
					strError = "All 'Teacher Benefit and Tax Rates' fields must be filled in.\n";
			}
		
		}
		
		if (pForm.Funding.value != "") {
			if (pForm.gradeK.value == "" || pForm.grade1.value == "" || pForm.grade2.value == "" ||
			    pForm.grade3.value == "" || pForm.grade4.value == "" || pForm.grade5.value == "" ||
			    pForm.grade6.value == "" || pForm.grade7.value == "" || pForm.grade8.value == "" || 
			    pForm.grade9.value == "" || pForm.grade10.value == "" || pForm.grade11.value == "" || pForm.grade12.value == "" ) {
					strError = strError + "All 'Student Funding' fields must be filled in.\n";
			}
			//alert(document.main.Funding.value);
		}
		
		if (pForm.dtSchool_Year_Start.value == "" || pForm.dtSchool_Year_End.value == "") {
			strError = strError + "First Day of School and Last Day of School can not be blank.\n";
		}else{
			var dtStart = Date.parse(pForm.dtSchool_Year_Start.value);
			var dtEnd = Date.parse(pForm.dtSchool_Year_End.value);
			if ('<% = MustStartAfter %>' != "") {
				var mustStartAfter = Date.parse('<% = MustStartAfter %>');				
				if (dtStart <= mustStartAfter) {
					strError = strError + "First Day of School must start AFTER <% = MustStartAfter%>.\n";
				}				
			}
			
			if ('<% = MustEndBefore %>' != "") {
				var mustEndBefore = Date.parse('<% = MustEndBefore %>');				
				if (dtEnd >= mustEndBefore) {
					strError = strError + "Last Day of School must end BEFORE <% = MustEndBefore%>.\n";
				}				
			}
			
			if (dtEnd <= dtStart) {
				strError = strError + "The First Day of School must come BEFORE the Last Day of School.\n";
			}
		}
		
		if (pForm.pwd.value != pForm.pwd2.value) {
			strError = "The Password and Confirm Password are not the same. Please re-enter.\n"
		}
		
		if (strError == "") {
			pForm.submit();
		}else{
			alert(strError);
		}
	}
	
	function jfUpdateList(id) {
		// if an item as been changed log it only once.  We will use this list
		// to determine which grades should be modified
		if (document.main.Funding.value.indexOf(id+",") == -1 ) {
			document.main.Funding.value = document.main.Funding.value + id + ",";
		}
	}	
</script>
<form name=main method=post action="globalVariables.asp" ID="Form1" onsubmit="return false;">
<input type="hidden" name="benefits" value=""/>
<input type="hidden" name="Funding" value="" ID="Hidden1"/>
<input type="hidden" name="Locks" value="" ID="Hidden2"/>
<input type="hidden" name="RestrictedStudents"/>
<input type="hidden" name="HasUpdated" />
<input type="hidden" name="InstructorFlatRate" value="" /> <%'JD 041411 added check change var for instructor rate %>
<table style="width:100%;" ID="Table3">
	<tr>
		<td class="yellowHeader">
			&nbsp;<b>Principal/Business Manager Administrative Page</b>
		</td>
	</tr>		
	<tr>
		<td class="svplain8">
			<br>
			This page allows system administration to modify Global settings that 
			impact many aspects of the application. <br>
			<b>Please Note: These settings are applicable for the <u>entire</u> <% = oFunc.SchoolYearRange %> school year.</b>
			<br><br>
		</td>
	</tr>
		<tr>
		<td class="TableHeader">
			&nbsp;<b>Define School Year</b>
		</td>
	</tr>
	<tr>
		<td class="svplain8">
			<b>First Day of School:</b> 
			<input type="text" size="10" name="dtSchool_Year_Start" value="<%if rs3.RecordCount > 0 then response.Write rs3("dtSchool_Year_Start")%>" ID="Text22" onChange="document.forms[0].Locks.value='dirty';">
						<a href="#" onclick="calStart.select(document.forms[0].dtSchool_Year_Start,'aStart','M/dd/yyyy','<%if rs3.RecordCount > 0 then response.Write rs3("dtSchool_Year_Start")%>');document.forms[0].Locks.value='dirty';return false;" id="aStart" name="aStart">calendar</a>
			&nbsp;&nbsp;
			<b>Last Day of School:</b> 
			<input type="text" size="10" name="dtSchool_Year_End" value="<%if rs3.RecordCount > 0 then response.Write rs3("dtSchool_Year_End")%>" ID="Text27" onChange="document.forms[0].Locks.value='dirty';">
						<a href="#" onclick="calEnd.select(document.forms[0].dtSchool_Year_End,'aEnd','M/dd/yyyy','<%if rs3.RecordCount > 0 then response.Write rs3("dtSchool_Year_End")%>');document.forms[0].Locks.value='dirty';return false;" id="aEnd" name="aEnd">calendar</a>
			<br><br>
		</td>
	</tr>
	<% 'JD 041411 new entry for flat instructor pay %>
    <tr>
		<td class="TableHeader">
			&nbsp;<b>Teacher Flat rate</b>
		</td>
	</tr>
	<tr>
	    <td class="svplain8">
	        <input type="hidden" name="intFlat_Inst_Id" value="<% if rs4.RecordCount > 0 then response.Write rs4("intFlat_Inst_Id")%>" ID="Hidden16" />
	        <input type="text" size="10" name="flatRate" value="<% if rs4.RecordCount > 0 then response.Write rs4("flatRate") %>" ID="Text30" onchange="document.forms[0].InstructorFlatRate.value = 'dirty';" />
	    </td>
	</tr>
	<% 'JD end new entry for flat instructor pay %>

	<tr>
		<td class="TableHeader">
			&nbsp;<b>Teacher Benefit and Tax Rates</b>
		</td>
	</tr>
	<tr>
		<td>
			<table cellpadding="3">
				<input type="hidden" name="intBenefit_Tax_Rate_ID" value="<% if rs1.RecordCount > 0 then response.Write rs1("intBenefit_Tax_Rate_ID")%>">
				<tr>
					<td class="TableCell" align="center">
						Group Life Insurance	<br>
						(in $'s)					
					</td>
					<td class="TableCell" align="center">
						Group Medical Insurance		<br>
						(in $'s)			
					</td>
					<td class="TableCell" align="center">
						Workers' Comp			<br>
						(in decimal form)	
					</td>
					<td class="TableCell" align="center">
						Unemployment	<br>
						(in decimal form)				
					</td>
					<td class="TableCell" align="center">
						FICA				<br>
						(in decimal form)	
					</td>
					<td class="TableCell" align="center">
						FICA Cap				<br>
						(in $)	
					</td>
					<td class="TableCell" align="center">
						Medicare			<br>
						(in decimal form)		
					</td>					
				</tr>
				<tr>
					<td align="center">
						<input type="text" size="10" name="curLife_Insurance" value="<% if rs1.RecordCount > 0 then response.Write rs1("curLife_Insurance")%>" onchange="this.form.benefits.value='dirty';">
					</td>
					<td align="center">
						<input type="text" size="10" name="curHealth_Cost" value="<%if rs1.RecordCount > 0 then response.Write rs1("curHealth_Cost")%>" ID="Text1" onchange="this.form.benefits.value='dirty';">
					</td>
					<td align="center">
						<input type="text" size="10" name="fltWorkmans_Comp" value="<%if rs1.RecordCount > 0 then response.Write rs1("fltWorkmans_Comp")%>" ID="Text2" onchange="this.form.benefits.value='dirty';">
					</td>
					<td align="center">
						<input type="text" size="10" name="fltUnemployment" value="<%if rs1.RecordCount > 0 then response.Write rs1("fltUnemployment")%>" ID="Text3" onchange="this.form.benefits.value='dirty';">
					</td>
					<td align="center">
						<input type="text" size="10" name="fltFICA" value="<%if rs1.RecordCount > 0 then response.Write rs1("fltFICA")%>" ID="Text4" onchange="this.form.benefits.value='dirty';">
					</td>
					<td align="center">
						<input type="text" size="10" name="curFICA_Cap" value="<%if rs1.RecordCount > 0 then response.Write rs1("curFICA_Cap")%>" ID="Text28" onchange="this.form.benefits.value='dirty';">
					</td>
					<td align="center">
						<input type="text" size="10" name="fltMedicare" value="<%if rs1.RecordCount > 0 then response.Write rs1("fltMedicare")%>" ID="Text5" onchange="this.form.benefits.value='dirty';">
					</td>
				</tr>
				<tr>
					<td class="TableCell" align="center">
						TERS Base Contract %					
					</td>
					<td class="TableCell" align="center">
						PERS Base Contract %			
					</td>
					<td class="TableCell" align="center">
						TERS Retirement			<br>
						(in decimal form)		
					</td>
					<td class="TableCell" align="center">
						PERS Retirement			<br>
						(in decimal form)	
					</td>
					<td>
					</td>
				</tr>
				<tr>
					<td align="center">
						<input type="text" size="10" name="intTERS_Base_Percent" value="<% if rs1.RecordCount > 0 then response.Write rs1("intTERS_Base_Percent")%>" onchange="this.form.benefits.value='dirty';" ID="Text25">
					</td>
					<td align="center">
						<input type="text" size="10" name="intPERS_Base_Percent" value="<%if rs1.RecordCount > 0 then response.Write rs1("intPERS_Base_Percent")%>" ID="Text26" onchange="this.form.benefits.value='dirty';">
					</td>
					<td align="center">
						<input type="text" size="10" name="fltTRS" value="<%if rs1.RecordCount > 0 then response.Write rs1("fltTRS")%>" ID="Text6" onchange="this.form.benefits.value='dirty';">
					</td>
					<td align="center">
						<input type="text" size="10" name="fltPERS" value="<%if rs1.RecordCount > 0 then response.Write rs1("fltPERS")%>" ID="Text7" onchange="this.form.benefits.value='dirty';">
					</td>
					<td>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>
			&nbsp;
		</td>
	</tr>
	<tr>
		<td class="TableHeader">
			&nbsp;<b>Student Funding Amounts</b> (by grade)
		</td>
	</tr>
	<tr>
		<td>
			<table cellpadding="3" ID="Table1">
				<tr>
					<td class="TableCell" align="center">
						K					
					</td>
					<td class="TableCell" align="center">
						1			
					</td>
					<td class="TableCell" align="center">
						2
					</td>
					<td class="TableCell" align="center">
						3				
					</td>
					<td class="TableCell" align="center">
						4
					</td>
					<td class="TableCell" align="center">
						5	
					</td>
					<td class="TableCell" align="center">
						6		
					</td>
					<td class="TableCell" align="center">
						7	
					</td>
					<td class="TableCell" align="center">
						8				
					</td>
					<td class="TableCell" align="center">
						9
					</td>
					<td class="TableCell" align="center">
						10	
					</td>
					<td class="TableCell" align="center">
						11		
					</td>
					<td class="TableCell" align="center">
						12	
					</td>
				</tr>
				<tr>
					<td align="center">
						<input type="text" size="5" name="gradeK" value="<%if rs2.RecordCount > 0 then response.Write rs2("gradeK")%>" ID="Text8" onchange="jfUpdateList(this.name);">						
						<input type="hidden" name="gradeKID" value="<%if rs2.RecordCount > 0 then response.Write rs2("IDk")%>">
					</td>
					<td align="center">
						<input type="text" size="5" name="grade1" value="<%if rs2.RecordCount > 0 then response.Write rs2("grade1")%>" ID="Text9" onchange="jfUpdateList(this.name);">
						<input type="hidden" name="grade1ID" value="<%if rs2.RecordCount > 0 then response.Write rs2("ID1")%>" ID="Hidden3">
					</td>
					<td align="center">
						<input type="text" size="5" name="grade2" value="<%if rs2.RecordCount > 0 then response.Write rs2("grade2")%>" ID="Text10" onchange="jfUpdateList(this.name);">
						<input type="hidden" name="grade2ID" value="<%if rs2.RecordCount > 0 then response.Write rs2("ID2")%>" ID="Hidden4">
					</td>
					<td align="center">
						<input type="text" size="5" name="grade3" value="<%if rs2.RecordCount > 0 then response.Write rs2("grade3")%>" ID="Text11" onchange="jfUpdateList(this.name);">
						<input type="hidden" name="grade3ID" value="<%if rs2.RecordCount > 0 then response.Write rs2("ID3")%>" ID="Hidden5">
					</td>
					<td align="center">
						<input type="text" size="5" name="grade4" value="<%if rs2.RecordCount > 0 then response.Write rs2("grade4")%>" ID="Text12" onchange="jfUpdateList(this.name);">
						<input type="hidden" name="grade4ID" value="<%if rs2.RecordCount > 0 then response.Write rs2("ID4")%>" ID="Hidden6">
					</td>
					<td align="center">
						<input type="text" size="5" name="grade5" value="<%if rs2.RecordCount > 0 then response.Write rs2("grade5")%>" ID="Text13" onchange="jfUpdateList(this.name);">
						<input type="hidden" name="grade5ID" value="<%if rs2.RecordCount > 0 then response.Write rs2("ID5")%>" ID="Hidden7">
					</td>
					<td align="center">
						<input type="text" size="5" name="grade6" value="<%if rs2.RecordCount > 0 then response.Write rs2("grade6")%>" ID="Text14" onchange="jfUpdateList(this.name);">
						<input type="hidden" name="grade6ID" value="<%if rs2.RecordCount > 0 then response.Write rs2("ID6")%>" ID="Hidden8">
					</td>
					<td align="center">
						<input type="text" size="5" name="grade7" value="<%if rs2.RecordCount > 0 then response.Write rs2("grade7")%>" ID="Text15" onchange="jfUpdateList(this.name);">
						<input type="hidden" name="grade7ID" value="<%if rs2.RecordCount > 0 then response.Write rs2("ID7")%>" ID="Hidden9">
					</td>
					<td align="center">
						<input type="text" size="5" name="grade8" value="<%if rs2.RecordCount > 0 then response.Write rs2("grade8")%>" ID="Text16" onchange="jfUpdateList(this.name);">
						<input type="hidden" name="grade8ID" value="<%if rs2.RecordCount > 0 then response.Write rs2("ID8")%>" ID="Hidden10">
					</td>
					<td align="center">
						<input type="text" size="5" name="grade9" value="<%if rs2.RecordCount > 0 then response.Write rs2("grade9")%>" ID="Text17" onchange="jfUpdateList(this.name);">
						<input type="hidden" name="grade9ID" value="<%if rs2.RecordCount > 0 then response.Write rs2("ID9")%>" ID="Hidden11">
					</td>
					<td align="center">
						<input type="text" size="5" name="grade10" value="<%if rs2.RecordCount > 0 then response.Write rs2("grade10")%>" ID="Text18" onchange="jfUpdateList(this.name);">
						<input type="hidden" name="grade10ID" value="<%if rs2.RecordCount > 0 then response.Write rs2("ID10")%>" ID="Hidden12">
					</td>
					<td align="center">
						<input type="text" size="5" name="grade11" value="<%if rs2.RecordCount > 0 then response.Write rs2("grade11")%>" ID="Text19" onchange="jfUpdateList(this.name);">
						<input type="hidden" name="grade11ID" value="<%if rs2.RecordCount > 0 then response.Write rs2("ID11")%>" ID="Hidden13">
					</td>
					<td align="center">
						<input type="text" size="5" name="grade12" value="<%if rs2.RecordCount > 0 then response.Write rs2("grade12")%>" ID="Text20" onchange="jfUpdateList(this.name);">
						<input type="hidden" name="grade12ID" value="<%if rs2.RecordCount > 0 then response.Write rs2("ID12")%>" ID="Hidden14">						
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>
			&nbsp;
		</td>
	</tr>
	<tr>
		<td class="TableHeader">
			&nbsp;<b>Deadlines and Locks</b>
		</td>
	</tr>
	<tr>
		<td>
			<table cellpadding="3" ID="Table2">
				<tr>
					<td class="TableCell" align="center">
						Progress Report Deadline Sem 1	
					</td>
					<td class="TableCell" align="center">
						Progress Report Deadline Sem 2			
					</td>			
					<td class="TableCell" align="center">
						Spending Lock Date		<br> (applies to Teachers & Guardians)			
					</td>
					<td class="TableCell" align="center">
						Lock School Year <% = oFunc.SchoolYearRange %>  <br> (applies to Teachers & Guardians)
					</td>	
					<td class="TableCell" align="center">
						Count Deadline
					</td>				
				</tr>
				<tr>
					<input type="hidden" name="intGlobal_Variable_ID" value="<% if rs3.RecordCount > 0 then response.Write rs3("intGlobal_Variable_ID")%>" ID="Hidden15">
					<td align="center" class="svplain8">
						<input type="text" size="15" name="dtSem_One_Progress_Deadline" value="<%if rs3.RecordCount > 0 then response.Write rs3("dtSem_One_Progress_Deadline")%>" ID="Text23" onChange="document.forms[0].Locks.value='dirty';">
						<a href="#" onclick="cal3.select(document.forms[0].dtSem_One_Progress_Deadline,'aSem1','M/dd/yyyy','<%if rs3.RecordCount > 0 then response.Write rs3("dtSem_One_Progress_Deadline")%>');document.forms[0].Locks.value='dirty';return false;" id="aSem1" name="aSem1">calendar</a>
					</td>
					<td align="center" class="svplain8">
						<input type="text" size="15" name="dtSem_Two_Progress_Deadline" value="<%if rs3.RecordCount > 0 then response.Write rs3("dtSem_Two_Progress_Deadline")%>" ID="Text24"  onChange="document.forms[0].Locks.value='dirty';">
						<a href="#" onclick="cal4.select(document.forms[0].dtSem_Two_Progress_Deadline,'aSem2','M/dd/yyyy','<%if rs3.RecordCount > 0 then response.Write rs3("dtSem_Two_Progress_Deadline")%>');document.forms[0].Locks.value='dirty';return false;" id="aSem2" name="aSem2">calendar</a>
					</td>	
					<td align="center" class="svplain8">
						<input type="text" size="10" name="dtLock_Spending" value="<%if rs3.RecordCount > 0 then response.Write rs3("dtLock_Spending")%>" ID="Text21" onChange="document.forms[0].Locks.value='dirty';">
						<a href="#" onclick="cal.select(document.forms[0].dtLock_Spending,'aSpend','M/dd/yyyy','<%if rs3.RecordCount > 0 then response.Write rs3("dtLock_Spending")%>');document.forms[0].Locks.value='dirty';return false;" id="aSpend" name="aSpend">calendar</a>
					</td>
					<td align="center" class="svplain8">
						<input type="checkbox" name="bolLock_School_Year" value="1" <% if rs3.RecordCount > 0 then if rs3("bolLock_School_Year") then response.Write " checked "%> onclick="document.forms[0].Locks.value='dirty';"><b>Yes</b>
					</td>	
					<td align="center" class="svplain8">
						<input type="text" size="15" name="dtCount_Deadline" value="<%if rs3.RecordCount > 0 then response.Write rs3("dtCount_Deadline")%>" ID="Text29"  onChange="document.forms[0].Locks.value='dirty';">
						<a href="#" onclick="cal5.select(document.forms[0].dtCount_Deadline,'aCountDeadline','M/dd/yyyy','<%if rs3.RecordCount > 0 then response.Write rs3("dtCount_Deadline")%>');document.forms[0].Locks.value='dirty';return false;" id="aCountDeadline" name="aCountDeadline">calendar</a>
					</td>								
				</tr>
			</table>
			<br>
		</td>		
	</tr>
	<tr>
		<td>
			<table>
				<tr>
					<td class="TableCell" align="center" colspan=2>
						Lock an Individual Students Packet/Spending<br>
						Effects Teachers and Guardians
					</td>	
				</tr>
				<TR>			
					<TD valign="top">
						<SELECT name="selStudents"  multiple size="6" style="FONT-SIZE:xx-small;width: 250px" ID="Select1">
							<option>----------						
							<%
							sqlStudent = "SELECT     s.intSTUDENT_ID, (CASE ss.intReEnroll_State WHEN 86 THEN s.szLAST_NAME + ',' + s.szFIRST_NAME + ': Withdrawn (' + CASE isNull(ss.dtWithdrawn, " & _ 
											" 1) WHEN 1 THEN 'No Date Entered' ELSE CONVERT(varChar(100), ss.dtWithdrawn)  " & _ 
											" END + ')' WHEN 123 THEN s.szLAST_NAME + ',' + s.szFIRST_NAME + ': Graduated (' + CONVERT(varChar(20), ss.dtModify)  " & _ 
											" + ')' ELSE s.szLAST_NAME + ',' + s.szFIRST_NAME END) AS Name, ss.intReEnroll_State, ss.dtWithdrawn " & _ 
											"FROM tblSTUDENT s INNER JOIN " & _ 
											" tblStudent_States ss ON s.intSTUDENT_ID = ss.intStudent_id " & _ 
											"WHERE (ss.intReEnroll_State IN (" & Application.Contents("strEnrollmentList") & ")) AND (ss.intSchool_Year = " & Session.Value("intSchool_Year") & ") " & _ 
											" AND NOT EXISTS(Select 'x' from STUDENT_LOCKED_ACCOUNTS sla where sla.StudentId = s.intStudent_Id and sla.SchoolYear = ss.intSchool_Year) " & _
											"ORDER BY Name "						
							Response.Write oFunc.MakeListSQL(sqlStudent,intStudent_ID,Name,"")												 
							%>
						</SELECT>
					</td>
					<TD valign="top" width=0%>
						<SELECT name="selRestrictedStudents"  multiple size="6" style="FONT-SIZE:xx-small;width:250px" ID="Select2">
							<%
							sql = "SELECT STUDENT_LOCKED_ACCOUNTS.LockedAccountId, STUDENT_LOCKED_ACCOUNTS.StudentId,   " & _ 
									"	tblSTUDENT.szLAST_NAME + ', ' + tblSTUDENT.szFIRST_NAME as StudentName " & _ 
									"FROM	STUDENT_LOCKED_ACCOUNTS INNER JOIN " & _ 
									"	tblSTUDENT ON STUDENT_LOCKED_ACCOUNTS.StudentId = tblSTUDENT.intSTUDENT_ID " & _ 
									"WHERE	(STUDENT_LOCKED_ACCOUNTS.SchoolYear = " & session.Contents("intSchool_Year") & ") " & _ 
									"ORDER BY tblSTUDENT.szLAST_NAME, tblSTUDENT.szFIRST_NAME "
							Response.Write oFunc.MakeListSQL(sql,"StudentId","StudentName","")			 
							%>
						</SELECT>
					</TD>
				<tr>
					<td align=right>
						<input type=button value="Lock Account >>" title="Add selected Family" class="btSmallGray"
						onclick="jfSelectItemFromTo('selStudents', 'selRestrictedStudents');" align=right NAME="Button2" ID="Button1">
					</td>
					<TD valign=middle align=right>
						<input type=button value="Unlock Account" style="position:relative"  class="btSmallGray" title="Remove selected Family or Families" onclick="jfRemoveItems('selRestrictedStudents');" ID="Button2" NAME="Button2">
					</TD>
				</tr>			
			</table>
		</td>
	</tr>
	<tr>
		<td class="TableHeader">
			&nbsp;<b>Change Super Admin Password</b>
		</td>
	</tr>
	<tr>
		<td class="svplain8">
			<b>Password:</b> <input type="password" name="pwd" size="15">&nbsp;&nbsp;
			<b>Confirm Password:</b> <input type="password" size="15" ID="Password1" NAME="pwd2">
		</td>
	</tr>
	<tr>
		<td>
			<input type="submit" value="Save Changes" class="NavSave" onclick="jfValidate(this.form);">
		</td>
	</tr>
</table>
</form>
<DIV ID="divCal" STYLE="position:absolute;visibility:hidden;background-color:white;layer-background-color:white;"></DIV>
<%
rs1.Close
set rs1 = nothing
rs2.Close
set rs2 = nothing
rs3.Close
set rs3 = nothing
rs4.Close
set rs4 = nothing
call oFunc.CloseCN()
set oFunc = nothing
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")

sub UpdateBenefits
	dim update
	update = "update tblBenefit_Tax_Rates set " & _
			 "fltTRS = " & oFUnc.EscapeTick(request.Form("fltTRS")) & ", " & _
			 "fltPERS = " & oFUnc.EscapeTick(request.Form("fltPERS")) & ", " & _
			 "fltMedicare = " & oFUnc.EscapeTick(request.Form("fltMedicare")) & ", " & _
			 "fltWorkmans_Comp = " & oFUnc.EscapeTick(request.Form("fltWorkmans_Comp")) & ", " & _
			 "curHealth_Cost = " & oFUnc.EscapeTick(request.Form("curHealth_Cost")) & ", " & _
			 "fltFICA = " & oFUnc.EscapeTick(request.Form("fltFICA")) & ", " & _
			 "fltUnemployment = " & oFUnc.EscapeTick(request.Form("fltUnemployment")) & ", " & _
			 "curLife_Insurance = " & oFUnc.EscapeTick(request.Form("curLife_Insurance")) & ", " & _
			 "curFICA_Cap = " & oFUnc.EscapeTick(request.Form("curFICA_Cap")) & ", " & _
			 "intTERS_Base_Percent = " & oFUnc.EscapeTick(request.Form("intTERS_Base_Percent")) & ", " & _
			 "intPERS_Base_Percent = " & oFUnc.EscapeTick(request.Form("intPERS_Base_Percent")) & ", " & _
			 "dtModify = CURRENT_TIMESTAMP, szUSER_Modify = '" & oFunc.EscapeTick(session.Contents("strUserID")) & "' " & _
			 "WHERE  intBenefit_Tax_Rate_ID = " & request.Form("intBenefit_Tax_Rate_ID")
	oFunc.ExecuteCN(update)			 
end sub

sub InsertBenefits
	if session.Contents("Benefits" & session.Contents("intSchool_Year")) = "" then
		dim insert
		insert = "Insert into tblBenefit_Tax_Rates (fltTRS, fltMedicare, fltWorkmans_Comp, fltPERS,  " & _ 
			" curHealth_Cost, fltFICA, fltUnemployment, curLife_Insurance,  " & _ 
			" curFICA_Cap,intSchool_Year,intTERS_Base_Percent, " & _
			" intPERS_Base_Percent,szUSER_CREATE, dtCREATE) " & _
			"Values (" & _
			oFUnc.EscapeTick(request.Form("fltTRS")) & ", " & _
			oFUnc.EscapeTick(request.Form("fltMedicare")) & ", " & _
			oFUnc.EscapeTick(request.Form("fltWorkmans_Comp")) & ", " & _
			oFUnc.EscapeTick(request.Form("fltPERS")) & ", " & _
			oFUnc.EscapeTick(request.Form("curHealth_Cost")) & ", " & _
			oFUnc.EscapeTick(request.Form("fltFICA")) & ", " & _
			oFUnc.EscapeTick(request.Form("fltUnemployment")) & ", " & _
			oFUnc.EscapeTick(request.Form("curLife_Insurance")) & ", " & _
			oFUnc.EscapeTick(request.Form("curFICA_Cap")) & ", " & _
			session.Contents("intSchool_Year") & "," & _
			oFUnc.EscapeTick(request.Form("intTERS_Base_Percent")) & ", " & _
			oFUnc.EscapeTick(request.Form("intPERS_Base_Percent")) & ", " & _
			"'" & oFunc.EscapeTick(session.Contents("strUserID")) & "', " & _
			"CURRENT_TIMESTAMP)"
		oFunc.ExecuteCN(insert)
	end if
	session.Contents("Benefits" & session.Contents("intSchool_Year")) = "true"
end sub

sub UpdateFunding
	arFunding = split(request.Form("Funding"),",")
	oFunc.BeginTransCn
	for i = 0 to ubound(arFunding)
		if arFunding(i) <> "" then
			if request.Form(arFunding(i) + "ID") <> "" then
				update = "update tblFunding set curFund_Amount = " & request.Form(arFunding(i)) & "," & _
						 "dtModify = CURRENT_TIMESTAMP, szUSER_Modify = '" & oFunc.EscapeTick(session.Contents("strUserID")) & "' " & _
						 " where intFunding_ID = " & request.Form(arFunding(i) + "ID")
						' response.Write update & "<BR>"
				oFunc.ExecuteCN(update)
			else
				if session.Contents("Funding" & session.Contents("intSchool_Year")) = "" then
					if isNumeric(request.Form(arFunding(i))) then
						iGrade = cint(request.Form(arFunding(i)))
						
						if iGrade <= 5 then
							fundTypeID = 1
						elseif iGrade >=6 and iGrade <= 8 then
							fundTypeID = 2
						elseif iGrade > 8 then
							fundTypeID = 3
						end if
					else
						fundTypeID = 1
					end if
					
					insert = "Insert into tblFunding(szGrade,intFund_Type_ID,curFund_Amount,intSchool_Year,szUSER_CREATE, dtCreate) " & _
							" VALUES (" & _
							"'" & ucase(replace(arFunding(i),"grade","")) & "'," & _
							fundTypeID & "," & _
							oFunc.EscapeTick(request.Form(arFunding(i))) & "," & _
							session.Contents("intSchool_Year") & "," & _
							"'" & oFunc.EscapeTick(session.Contents("strUserID")) & "', " & _
							"CURRENT_TIMESTAMP)"	
					oFunc.ExecuteCN(insert)				 
				end if
			end if			
		end if
	next
	 session.Contents("Funding" & session.Contents("intSchool_Year")) = "true"
	 
	 oFunc.ExecuteCN("ts_UpdateallStudentExpenses " & session.Contents("intSchool_Year"))
	oFunc.CommitTransCn
end sub

sub UpdatePWD 
	dim update	
	update = "update tblSecurity set " & _
			 " szPassword = '" & oFunc.EscapeTick(request.Form("pwd")) & "',"  & _
			 " dtModify = CURRENT_TIMESTAMP, szUSER_Modify = '" & oFunc.EscapeTick(session.Contents("strUserID")) & "' " & _
			 "WHERE intSecurity_ID = 1"
	oFunc.ExecuteCN(update)
end sub

sub UpdateGlobalVars
	dim update
	call CheckDates	
	update = "update tblGlobal_Variables set " & _
			 " dtLock_Spending = " & dtLock_Spending & ","  & _
			 " dtSem_One_Progress_Deadline = " & dtSem_One_Progress_Deadline & ","  & _
			 " dtSem_Two_Progress_Deadline = " & dtSem_Two_Progress_Deadline & ","  & _
			 " bolLock_School_Year = " & bolLock_School_Year & ","  & _
			 " dtSchool_Year_Start = " & dtSchool_Year_Start & "," & _
			 " dtSchool_Year_End = " & dtSchool_Year_End & "," & _
			 " dtCount_Deadline = " & dtCount_Deadline & ", " & _
			 " dtModify = CURRENT_TIMESTAMP, szUSER_Modify = '" & oFunc.EscapeTick(session.Contents("strUserID")) & "' " & _
			 "WHERE intGlobal_Variable_ID = " & request.Form("intGlobal_Variable_ID") 
	oFunc.ExecuteCN(update)
end sub

sub UpdateLockedStudentAccounts
	dim delete, insert
	delete = "delete from STUDENT_LOCKED_ACCOUNTS where SchoolYear = " & session.Contents("intSchool_Year")
	oFunc.ExecuteCN(delete)
		
	arList = split(request("RestrictedStudents"),",")
	for i = 0 to ubound(arList)
		if isNumeric(arList(i)) then
			insert = "insert into STUDENT_LOCKED_ACCOUNTS(StudentID, SchoolYear,UserCreated) " & _
					" values (" & arList(i) & "," & session.Contents("intSchool_Year") & ",'" & session.Contents("strUserID") & "')"
			oFunc.ExecuteCN(insert)
		end if
	next
	Application.Contents("LockedStudentAccounts" & session.Contents("intSchool_Year")) = "," & request("RestrictedStudents") & ","

end sub
sub InsertGlobalVars
	dim insert
	call CheckDates
	
	if session.Contents("GlobalVars" & session.Contents("intSchool_Year")) = "" then		
		insert = "Insert into tblGlobal_Variables(dtLock_Spending, dtSem_One_Progress_Deadline,  " & _ 
			" dtSem_Two_Progress_Deadline, bolLock_School_Year, intSchool_Year, dtSchool_Year_Start, dtSchool_Year_End, " & _ 
			" szUSER_CREATE,dtCREATE, dtCount_Deadline) " & _
			" VALUES(" & _
			"" & dtLock_Spending & ","  & _
			"" & dtSem_One_Progress_Deadline & ","  & _
			"" & dtSem_Two_Progress_Deadline & ","  & _
			"" & bolLock_School_Year & ","  & _
			"'" &session.Contents("intSchool_Year") & "',"  & _
			"" & dtSchool_Year_Start & ","  & _
			"" & dtSchool_Year_End & ","  & _
			"'" & oFunc.EscapeTick(session.Contents("strUserID")) & "', " & _
			"CURRENT_TIMESTAMP, " & dtCount_Deadline & ")"
		oFunc.ExecuteCN(insert)
		
		session.Contents("GlobalVars" & session.Contents("intSchool_Year")) = "true"										
	end if
end sub

sub CheckDates()
	dim iYear
	iYear = session.Contents("intSchool_Year")
	if isDate(request.Form("dtLock_Spending")) then
		dtLock_Spending = "'" & cdate(request.Form("dtLock_Spending")) & "'"
		Application.Contents("dtLock_Spending" & iYear) = cdate(request.Form("dtLock_Spending"))
	else
		dtLock_Spending = " NULL "
		Application.Contents("dtLock_Spending" & iYear) = ""
	end if
	
	if isDate(request.Form("dtSem_One_Progress_Deadline")) then
		dtSem_One_Progress_Deadline = "'" & cdate(request.Form("dtSem_One_Progress_Deadline")) & "'"
		Application.Contents("dtSem_One_Progress_Deadline" & iYear) = cdate(request.Form("dtSem_One_Progress_Deadline"))
	else
		dtSem_One_Progress_Deadline = " NULL "
		Application.Contents("dtSem_One_Progress_Deadline" & iYear) = ""
	end if
	
	if isDate(request.Form("dtCount_Deadline")) then
		dtCount_Deadline = "'" & cdate(request.Form("dtCount_Deadline")) & "'"
		Application.Contents("dtCount_Deadline" & iYear) = cdate(request.Form("dtCount_Deadline"))
	else
		dtCount_Deadline = " NULL "
		Application.Contents("dtCount_Deadline" & iYear) = ""
	end if
	
	if isDate(request.Form("dtSem_Two_Progress_Deadline")) then
		dtSem_Two_Progress_Deadline = "'" & cdate(request.Form("dtSem_Two_Progress_Deadline")) & "'"
		Application.Contents("dtSem_Two_Progress_Deadline" & iYear) = cdate(request.Form("dtSem_Two_Progress_Deadline"))
	else
		dtSem_Two_Progress_Deadline = " NULL "
		Application.Contents("dtSem_Two_Progress_Deadline" & iYear) = ""
	end if
	
	if bolLock_School_Year = 1 then 
		Application.Contents("bolLock_School_Year" & iYear) = true
	else
		Application.Contents("bolLock_School_Year" & iYear) = 0
	end if
	
	if isDate(request.Form("dtSchool_Year_Start")) then
		dtSchool_Year_Start = "'" & cdate(request.Form("dtSchool_Year_Start")) & "'"
		Application.Contents("dtSchool_Year_Start" & iYear) = cdate(request.Form("dtSchool_Year_Start"))
	else
		dtSchool_Year_Start = " NULL "
		response.Write "ERROR: First Day of School was not a valid date.  Go <a href='#' onClick=""history.go(-1);"">Back</a> and correct."
		response.End
	end if
	
	if isDate(request.Form("dtSchool_Year_End")) then
		dtSchool_Year_End = "'" & cdate(request.Form("dtSchool_Year_End")) & "'"
		Application.Contents("dtSchool_Year_End" & iYear) = cdate(request.Form("dtSchool_Year_End"))
	else
		dtSchool_Year_End = " NULL "
		response.Write "ERROR: Last Day of School was not a valid date.  Go <a href='#' onClick=""history.go(-1);"">Back</a> and correct."
		response.End
	end if
end sub

'JD 041411 sub to update instructor flat rate
sub UpdateInstructorFlatRate
    dim update
    update = "update tblInstructor_Flat_Rate set flatRate = " & Request.Form("flatRate") & " where intSchool_Year = " & session.Contents("intSchool_Year")
    oFunc.ExecuteCN(update)
end sub
'JD 041411 sub to insert instructor flat rate
sub InsertInstructorFlatRate

    dim insert
    insert = "Insert into dbo.tblInstructor_Flat_Rate (flatRate, intSchool_year) Values (" & Request.Form("flatRate") & ", " & session.Contents("intSchool_Year") & ")"
    oFunc.ExecuteCN(insert)
end sub
%>