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
dim strItemType			' tells user if item is requiestion or reimbursement
dim oHtml

mLablelCount = 0
mDivCount = 0		

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()
oFunc.ResetSelectSessionVariables

set oHtml = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/htmlFuncs.wsc"))
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Handle Data Modifications
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Handle deletion if needed
if request.QueryString("bolDelete") <> "" then
	call vbsDelete(request.QueryString("intShort_ILP_ID"),request.QueryString("intStudent_ID"))
elseif request("intEnroll_Info_ID") <> "" and request("changePercent") <> "" then
	'Handle updating of Percent Enrolled
	call vbsUpdateEnrollPercent(request("intPercent_Enrolled_Fpcs"),request("intEnroll_Info_ID"))
end if
 strMessage = request.QueryString("strMessage")
' Since we have simular functionality for admins and teachers define what
' fields we should be working with based on role
if ucase(session.Contents("strRole")) = "ADMIN" then
	strCommentField = "szAdmin_Comments"
	
	if request("bolLock") & "" <> "" then
		call vbsLockEnrollLevel(request("intEnroll_Info_ID"),request("intPercent_Enrolled_Fpcs"))
	end if
	
	if request("updateTestForm") <> "" then
		call vbsUpdateTestForm(request("bolASD_Testing"), request("intEnroll_Info_ID"))
	end if
	
	if request("updateProgressForm") <> "" then
		call vbsUpdateProgressForm(request("bolProgress"), request("intEnroll_Info_ID"))
	end if
elseif ucase(session.Contents("strRole")) = "TEACHER" then
	strCommentField = "szSponsor_Comments"
elseif ucase(session.Contents("strRole")) = "GUARD" then
	' saves status change made by guardian preventing non guardian status' from 
	' being processed
	if request("bolChangeStatus") <> "" and (request("bolApproved") = "implemented" _
		or request("bolApproved") = "ready for sponsor") then
		call vbsApprovedStatus(request("intILP_ID"),request("bolApproved"))
	end if
end if

if (ucase(session.Contents("strRole")) = "ADMIN" or ucase(session.Contents("strRole")) = "TEACHER") _
	 and request("strILPList") <> "" then
	call vbsUpdateILPStatus(request("strILPList"))
end if

if request("bolUpdateComments") <> "" and strCommentField <> "" then
	call vbsUpdateComments(oFunc.EscapeTick(request("szComments")),request("intILP_ID"))
end if
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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
			"s.intGrad_year, ss.szGrade, e.intPercent_Enrolled_FPCS, e.intEnroll_Info_ID, " & _
			"i.szFirst_Name + ' ' + i.szLast_Name as Sponsor, i.szEmail as SponsorEmail, e.bolASD_Testing, e.bolProgress_Agreement " & _
			"FROM tblStudent s left outer join " & _
			" tblIEP ON s.intSTUDENT_ID = tblIEP.intStudent_ID LEFT OUTER JOIN " & _
			"tblEnroll_info e on s.intStudent_ID = e.intStudent_ID  LEFT OUTER JOIN "  & _
			"tblInstructor i on e.intSponsor_Teacher_ID = i.intInstructor_ID inner join " & _
			"tblStudent_States ss on ss.intStudent_ID = s.intStudent_ID and ss.intSchool_Year = " & session.Contents("intSchool_Year") & " " & _
			"where s.intStudent_ID = " & intStudent_ID & _
			" and e.sintSchool_Year = " & session.Value("intSchool_Year") & _
			" AND (tblIEP.intSchool_Year = " & session.Value("intSchool_Year") & ")"
		if ucase(session.Contents("strUserID")) = "SCOTT" then
			'response.Write sql
		end if		
	
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
		Server.Execute(Application.Value("strWebRoot") & "includes/Header.asp")
		%>
			<table cellspacing=0 cellpadding=4 width=85% ID="Table4">
				<tr>
					<td class=svplain10>
						<b>Before you can plan any courses please update
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

strStudentInfo = "<table cellspacing=1 cellpadding=2 style='height:100%;'>" & _	
					"<tr><td class='TableCell'>Core Unit:</td>" & _			
					"<td class='TableCell' align='right'>" & formatNumber((arStudentEnroll(0)/90),1) & "</td></tr>" & _
					"<tr><td class='TableCell'>Elective Unit:</td>" & _
					"<td class='TableCell' align='right'>" & formatNumber((arStudentEnroll(1)/90),1) & "</td></tr>" & _
					"<tr><td class='TableCell'>ASD Contracted Hrs:</td>" & _
					"<td class='TableCell' align='right'>" &  arStudentEnroll(2) & "</td></tr>" & _
					"<tr><td class='TableCell'>Total Hrs:</td>" & _
					"<td class='TableCell' align='right'>" & intTotalHrs & "</td>" & _
					"</tr></table>" 

if ucase(session.Contents("strRole")) = "ADMIN" then
	if bolASD_Testing then
		strChecked = " checked "
	else
		strChecked = ""
	end if
	
	strTestForm = "<input type=checkbox name=""bolASD_Testing"" " & strChecked & "value=""true"" onClick=""jfUpdateTestForm(this.checked);"">"


	if bolProgress_Agreement then
		strChecked = " checked "
	else
		strChecked = ""
	end if
	
	strProgressForm = "<input type=checkbox name=""bolProgress_Agreement"" " & strChecked & "value=""true"" onClick=""jfUpdateProgressForm(this.checked);"">"
else
	if bolASD_Testing then
		strTestForm = "Yes"
	else
		strTestForm = "Please click <a href=""javascript:"" onClick=""jfPrintTestForm();"">HERE</a> to print form."
	end if
	
	if bolProgress_Agreement then
		strProgressForm = "Yes"
	else
		strProgressForm = "Please click <a href=""javascript:"" onClick=""jfPrintProgressForm();"">HERE</a> to print form."
	end if
end if
			
strFormsTable =		"<table cellspacing=1 cellpaddin='2' style='height:100%;'>" & _
					"<tr><td class='TableCell' colspan=2>MANDATORY SIGNED FORMS</td></tr>" & _
					"<tr><td class='TableCell'>Student Testing:</td>" & _
					"<td class='TableCell'>" & strTestForm & "</td></tr>" & _
					"<tr><td class='TableCell'>Student Progress Report:</td>" & _
					"<td class='TableCell'>" & strProgressForm & "</td>" & _
					"</tr></table>" 	
					
strColorTable = "<table>" & _
			    "<tr><td class='SubHeader' nowrap>&nbsp;&nbsp;</td><td class='svplain8'><b>planned</b></td></tr>" & _
			    "<tr><td class='TableheaderPurple' nowrap>&nbsp;&nbsp;</td><td class='svplain8'><b>ready for sponsor</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>" & _
			    "<tr><td class='TableheaderBlue' nowrap>&nbsp;&nbsp;</td><td class='svplain8'><b>Implemented</b></td></tr>" & _
			    "<tr><td class='Tableheader' nowrap>&nbsp;&nbsp;</td><td class='svplain8'><b>s-appr</b></td></tr>" & _
			    "<tr><td class='TableheaderRed' nowrap>&nbsp;&nbsp;</td><td class='svplain8'><b>must amend</b></td></tr>" & _			    
			    "<tr><td class='TableHeaderGreen' nowrap>&nbsp;&nbsp;</td><td class='svplain8'><b>a-appr</b></td></tr>" & _
			    "</table>" 
			    								
	
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

if request("print") <> "" or request("simpleHeader") <> "" then
	Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")
	strSimpleHeader = "simpleHeader=true&"
	if request("print") <> "" then
	%>
	<link rel="stylesheet" href="<% =  Application.Value("strWebRoot") %>CSS/printStyle.css">
	<script language=javascript>
		function jfPrint(){
			if (window.print){
			window.print()
			var obj = document.getElementById("btPrint");
			obj.style.display = "none";
			}
			else {
			alert("Mac users: please press Apple-P to print this form.\nWindows users: Please press ctrl-P to print this form.")
			}
		}		
	</script>
	<input name="btPrint" type=button value="Print This Page" class="btPrint" onclick="jfPrint();" ID="Button1">
	<%
	end if
else
	Server.Execute(Application.Value("strWebRoot") & "includes/Header.asp")
end if

' Print top section of worksheet
%>
<script language=javascript>	
	
	function jfOrder(ilpID){
	var strURL = "<%=Application.Value("strWebRoot")%>forms/Requisitions/req1.asp?intStudent_ID=<%=intStudent_ID%>&intILP_ID=" + ilpID;
		costsWin = window.open(strURL,"costsWin","width=710,height=500,scrollbars=yes,resize=yes,resizable=yes");
		costsWin.moveTo(0,0);
		costsWin.focus();
	}
	
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
					obj.style.display = '';
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
	
	function jfCallAction(id){
		var strAction = eval("document.main.action"+id+".value;");
		
		if (strAction == "edit") {
			jfOpen(id);
		}else if(strAction == "delete"){
			jfDelete(id);
		}else if (strAction == "budget"){
			jfBudget(id,'');
		}else{		
			eval(strAction);
		}
	}
	
	function jfPrintTestForm(){
		var strURL = "<%=Application.Value("strWebRoot")%>forms/PrintableForms/allPrintable.asp?strAction=T&intStudent_ID=<% = intStudent_ID %>";
		var testWin = window.open(strURL,"testWin","width=710,height=500,scrollbars=yes,resizable=yes");
		testWin.moveTo(0,0);
		testWin.focus();
	}
	
	function jfPrintProgressForm(){
		var strURL = "<%=Application.Value("strWebRoot")%>forms/PrintableForms/allPrintable.asp?strAction=P&intStudent_ID=<% = intStudent_ID %>";
		var progressWin = window.open(strURL,"progressWin","width=710,height=500,scrollbars=yes,resizable=yes");
		progressWin.moveTo(0,0);
		progressWin.focus();
	}
	
	function jfOpen(id){
		// Opens up edit/add course window
		var winSF;
		var URL = "<%=Application.Value("strWebRoot")%>forms/packet/addCourse.asp?intStudent_ID=<%=intStudent_ID%>&bolHighSchool=<%=bolHighSchool%>&intShort_ILP_ID=" + id;
		winSF = window.open(URL,"winSF","width=640,height=250,scrollbars=yes,resizable=yes");
		winSF.moveTo(0,0);
		winSF.focus();
	}

	function jfDelete(id){
		window.location.href="<%=Application.Value("strWebRoot")%>forms/packet/packet.asp?bolDelete=true&intStudent_ID=<%=intStudent_ID%>&intShort_ILP_ID="+id;	
	}
	
	<%	 
	if strMessage <> "" then
		response.write "alert('" & strMessage & "');"
		strMessage = "" 
	end if		
	%>
	
	function jfViewAnotherStudent(id){
		var strURL = "<%=Application.Value("strWebRoot")%>forms/packet/packet.asp?intStudent_ID=" + id.value;
		window.location.href = strURL;
	}
	
	function jfChangePercent(enrollId,percent){
		var bolConfirm = confirm("Are you sure you want to change the 'Target Enrollment' percentage?");
		if (bolConfirm){
			var URL = "<%=Application.Value("strWebRoot")%>forms/packet/packet.asp?intStudent_ID=<%=intStudent_ID%>&";
			URL += "intEnroll_Info_ID=" + enrollId;
			URL += "&changePercent=true&intPercent_Enrolled_Fpcs=" + percent; 
			window.location.href = URL;
		}else{
			// returns selection back to original value
			document.main.intPercent_Enrolled_Fpcs.selectedIndex = document.main.lastIndex.value;
		}
	}
	
	function jfUpdateTestForm(pValue){
		var URL = "<%=Application.Value("strWebRoot")%>forms/packet/packet.asp?intStudent_ID=<%=intStudent_ID%>&";
			URL += "updateTestForm=true&intEnroll_Info_ID=<%= intEnroll_Info_ID%>";
			URL += "&bolASD_Testing=" + pValue; 
			window.location.href = URL;
	}
	
	function jfUpdateProgressForm(pValue){
		var URL = "<%=Application.Value("strWebRoot")%>forms/packet/packet.asp?intStudent_ID=<%=intStudent_ID%>&";
			URL += "updateProgressForm=true&intEnroll_Info_ID=<%= intEnroll_Info_ID%>";
			URL += "&bolProgress=" + pValue; 
			window.location.href = URL;
	}
	
	function jfGetIndex(obj){
		// Stores original value
		document.main.lastIndex.value = obj.selectedIndex;	
	}
	
	function jfViewCosts(studentID,ilpID,classID){
		var strURL = "<%=Application.Value("strWebRoot")%>forms/Requisitions/req1.asp?intClass_ID="+classID;
		strURL += "&intStudent_ID=" + studentID + "&intILP_ID=" + ilpID;
		costsWin = window.open(strURL,"costsWin","width=710,height=500,scrollbars=yes,resize=yes,resizable=yes");
		costsWin.moveTo(0,0);
		costsWin.focus();
	}
	
	function jfViewILP(ilp_id,class_ID,class_name,cg,vendor,teacherName) {
		var ilpWin;
		var strURL;
		var strILP;
				
		strURL = "<%=Application.Value("strWebRoot")%>forms/ILP/ilpMain.asp?plain=yes&intILP_ID=" + ilp_id + "&intClass_id=" + class_ID;
		strURL += "&szClass_Name=" + class_name;
		strURL += "&intVendor_ID=" + vendor;
		strURL += "&strTeacherName=" + teacherName;
		strURL += "&intContract_Guardian_ID=" + cg;
		ilpWin = window.open(strURL,"ilpWin","width=710,height=500,scrollbars=yes,resizable=yes");
		ilpWin.moveTo(0,0);
		ilpWin.focus();
	}
	
	function jfContractSchedule(class_id,instructor_id,instruct_type,intContract_Guardian_ID,intGuardian_ID,intVendor_ID) {
		var classWin;
		var strURL = "<%=Application.Value("strWebRoot")%>forms/teachers/classAdmin.asp?bolInWindow=true&plain=yes<%=strDisabled%>&intClass_id="+class_id;
		strURL += "&intInstructor_id="+instructor_id+"&intInstruct_Type_ID="+instruct_type;
		strURL += "&intContract_Guardian_ID="+intContract_Guardian_ID;
		strURL += "<% = strHideGoodService %>&intGuardian_id="+intGuardian_ID;
		strURL += "&intVendor_ID="+intVendor_ID;
		classWin = window.open(strURL,"classWin","width=750,height=500,scrollbars=yes,resizable=yes");
		classWin.moveTo(0,0);
		classWin.focus();
	}
	
	function jfDeleteILP(ilp_id) {
		var answer;
		answer = confirm("Are you sure you want to delete this class? (All Goods and Services and the ILP for this class will be deleted as well)");
		if (answer) {
			var winDel;
			winDel = window.open("<%=Application.Value("strWebRoot")%>forms/teachers/deleteClass.asp?intILP_id="+ilp_id,"winDel","width=200,height=200,scrollbars=yes,resizable=yes");
			winDel.moveTo(0,0);
			winDel.focus();			
		}
	}	
	
	function jfViewRoll(class_id) {
		var winRoll;
		winRoll = window.open("<%=Application.Value("strWebRoot")%>Reports/studentsInClass.asp?intClass_id="+class_id,"winRoll","width=640,height=480,scrollbars=yes,resizable=yes");
		winRoll.moveTo(0,0);
		winRoll.focus();					
	}
	
	function jfAddClass(studentID,shortID){
		// Opens ilp1.asp that will create a class 
		var URL = "<%=Application.Value("strWebRoot")%>forms/ilp/ILP1.asp?intStudent_ID="+studentID+"&intShort_ILP_ID="+shortID;
		window.location.href = URL;
	}
	
	
	function jfChangeSponsor(studentID){
		window.location.href = "<%=Application.Value("strWebRoot")%>forms/packet/addSponsorTeacher.asp?intStudent_ID=" + studentID;
	}
	
	function jfShowComments(acom,scom){
		if (acom != "") { acom = "AA Comments: " + acom + "\n";}
		if (scom != "") { scom = "Sponsor Teacher Comments: " + scom;}
		alert(acom + scom);
	}
	
	function jfILPStatus(pID){
		var sList = document.main.strILPList;
		
		if (sList.value.indexOf(","+pID+",") == -1 ) {
			sList.value = sList.value + pID + ",";
		}
	}
	
	function jfUpdateStatus(myVal,ilp_id){
		var URL = "<%=Application.Value("strWebRoot")%>forms/packet/packet.asp?bolChangeStatus=true&bolApproved="+myVal.value;
		URL += "&intILP_ID="+ilp_id+"&intStudent_ID=<%=intStudent_ID%>"
		window.location.href = URL;
	}
	
	function jfPrintPacket(){
		var strURL;
		<% IF Request.QueryString("intStudent_id") <> "" then %>
		strURL = "&intStudent_ID=<%=Request("intStudent_id")%>";
		<% end if %>		
		//var winPrintPacket = window.open("<%=Application.Value("strWebRoot")%>forms/PrintableForms/allPrintable.asp?strAction=S" + strURL,"winPrintPacket","width=700,height=500,scrollbars=yes,resize=yes,resizable=yes");
		var winPrintPacket = window.open("<%=Application.Value("strWebRoot")%>forms/PrintableForms/printPacket.asp?strAction=S" + strURL,"winPrintPacket","width=700,height=500,scrollbars=yes,resize=yes,resizable=yes");
		winPrintPacket.moveTo(0,0);
		winPrintPacket.focus();
	}
</script>
<form name="main" action="<%=Application("strSSLWebRoot")%>forms/packet/packet.asp" method="post" ID="Form1">
<input type="hidden" name="intStudent_ID" value="<%=intStudent_ID%>" ID="Hidden2">
<input type="hidden" name="bolHighSchool" value="<%=bolHighSchool%>" ID="Hidden3">
<input type="hidden" name="courseTitleData" value="" ID="Hidden4">
<input type=hidden name="simpleHeader" value="<% = request("simpleHeader") %>" ID="Hidden5">
<input type=hidden name="lastIndex" value="" ID="Hidden6">
<table width="100%" ID="Table3" >
	<tr>
		<td class="yellowHeader">
			<table width=100% cellpadding=0 cellspacing=0 ID="Table8">
				<tr  class="yellowHeader">
					<td align=left >
						&nbsp;<b> Student Packet/Budget for <% = szFirst_Name & " " &  szLast_Name %> </b> 
					</td>
					<td align=right>
						<% 
							if ucase(session.Contents("strRole")) <> "GUARD" then
								dim strFamEmail							
								strFamEmail = oFunc.FamilyInfo("1",intStudent_ID,"4")
								if strFamEmail & "" = "" then
									response.Write "<B>Family Email:</b> No Email Provided"
								else
									response.Write "<B>Family Email:</b> <a href=""mailto:" & strFamEmail & "?cc=" & SponsorEmail & """ style=""color:white;"">" & strFamEmail & "</a>"
								end if								
							end if
						%>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>
			<table ID="Table2">
					<tr>
						<td class="TableSubHeader" colspan=4>
							<B>&nbsp;Student Enrollment Information</B>
						</td>
					</tr>
					<tr>
						<td valign='top'  style='height:100%;'>
							<table ID="Table5" cellspacing='1' cellpadding='2' style='height:100%;'>
								<tr>	
									<td class="TableCell" title="The enrollment goal you have chosen determines the eligible funding amount for SY <% = oFunc.SchoolYearRange%>.">					
										Planned Enrollment: 					
									</td>
									<td class="TableCell" valign=middle>
										<% if (intPercent_Enrolled_Locked <> "" and Ucase(session.Contents("strRole")) <> "ADMIN") or oFunc.LockYear then %>
											&nbsp;<% if intPercent_Enrolled_Locked <> "" then 
														response.Write intPercent_Enrolled_Locked
													else
														response.Write intPercent_Enrolled_Fpcs
													end if%>% 
											<input type=hidden name=intPercent_Enrolled_Fpcs value="<%=intPercent_Enrolled_Fpcs%>" ID="Hidden1" >
										<% else %>			
										&nbsp;<select name="intPercent_Enrolled_Fpcs" style="font:arial;font-size=10;" onclick="jfGetIndex(this);" onchange="jfChangePercent('<% = intEnroll_Info_ID %>',this.value);" ID="Select1">
										<%
										Response.Write oFunc.MakeList("25,50,75,100","25%,50%,75%,100%",intPercent_Enrolled_FPCS)
										%>
										</select>
										<% end if %>
										<% if ucase(session.Contents("strRole")) = "ADMIN" then %>
										<script language=javascript>
											function jfLock(){
												var intSel = document.main.intPercent_Enrolled_Fpcs.selectedIndex;
												var intLevel = document.main.intPercent_Enrolled_Fpcs.options[intSel].value;
												var URL = "<%=Application.Value("strWebRoot")%>forms/Packet/Packet.asp?";
												URL += "intStudent_ID=<%=intStudent_ID%>&intPercent_Enrolled_Fpcs="+intLevel;
												URL += "&bolLock=true&intEnroll_Info_ID=<%=intEnroll_Info_ID%>";
												window.location.href = URL;
											}
										</script>
										<input type=button class="btSmallGray" value="lock" onClick="jfLock();" title="Lock enrollment level so level can never excede locked level." NAME="Button2">
										<% end if %>
										<% if intPercent_Enrolled_Locked <> "" then %>
										<img src="<%=Application("strImageRoot")%>lock.gif" title="Enrollment level has been locked at <% = intPercent_Enrolled_Locked%>%">
										<% end if %>
								</tr>							
								<tr>
									<td class="TableCell" valign=top title="This is the enrollment level you currently have based on the amount of the plan you've implemented.">
										Actual Enrollment: 						
									</td>
									<td class="TableCell" valign=top>
										&nbsp;<% = intActualEnroll%>%	
									</td>
								</tr>
								<tr>
									<td class="TableCell" valign=middle>																				
										Sponsor Teacher:															
									</td>
									<td class="TableCell" valign=middle>
										&nbsp;<a href="mailto:<% = SponsorEmail %>"><% = Sponsor%></a>&nbsp;
										<input type=button value="edit" onclick="jfChangeSponsor('<%=intStudent_ID%>');" class="btSmallGray" <% if oFunc.LockYear then response.Write " disabled "%>>
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
						<td valign=top>
							<% = strStudentInfo %>
						</td>
						<td valign=top style='height:100%;'>
							<% = strFormsTable %>
						</td>
					</tr>
				</table>
		</td>
	</tr>
	<tr>
		<td>
			<table ID="Table1" cellspacing="2">				
				<tr >					
					<td bgcolor=white colspan="12" >
						<table width=95% align=left ID="Table6">
							<tr>
								<td colspan="10">
									<font class="svplain10">
									<% if not oFunc.LockYear then
									%>
									<input type="button" value="Plan New Course" class="NavLink" onclick="jfOpen('');" NAME="Button2">                          
									<input type="button" value="Hide Budget" class="NavLink" onclick="jfToggleBudget(this);" NAME="Button2"> 
									<input type="button" value="Print Packet" onclick="jfPrintPacket();" class="NavLink"> 
									<%
										end if
										if ucase(session.Contents("strRole")) = "ADMIN" or ucase(session.Contents("strRole")) = "TEACHER" then %>
									<input type=hidden name="strILPList" value="," ID="Hidden7">
									<input type=submit value="Save Status & Comments" class="NavSave" style="width:165px;">									
									<% end if %>
									</font>									
								</td>
							</tr>    
						</table>
						
					</td>
					<td class=gray align=center>
						Planned Budget
					</td>
					<td class=gray align=center>
						Actual
					</td>
				</tr>
				<tr >
					<td rowspan=3 colspan=8 class="svplain">
						<%
							if ucase(session.Contents("strRole")) = "GUARD" then
						%>
							<ul>														
							
							<li>Once you are finished working on a class and are ready for your sponsor
							to review it select 'ready for sponsor' under ILP Status to alert the 
							sponsor teacher.</li>
							</ul>
						<%
							elseif ucase(session.Contents("strRole")) = "TEACHER" then
						%>
							<ul>
							<li>To learn more about how each budget column is calculated simply
							mouse over the column in question.</li>														
							<li><b>PLEASE NOTE:</b> Once you change an ILP Status or modify a comment you
							must click the "Save Status & Comments" or "Save" button in order for the changes to be saved.</li>						
							</ul>
						<%
							end if
						%>
						<br>
						<% response.Write oHtml.ToolTip("<B><a href='#' class='svplain8'>Click to View Course Color Key</a></b>",strColorTable,true,"Course Color Key",false,"tooltip","250px","",false,true) %>
					</td>
					<td align=right class="svplain8" nowrap colspan="3">
						Beginning Balance:	
					</td>
					<td bgcolor=white style="width:0%;">
						&nbsp;
					</td>
					<td class=gray align=right>
						$<%=dblTargetBalance%>
					</td>
					<td class=gray align=right>
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
					<td align=right class="svplain8" nowrap colspan="3">
						Budget Transfer Deposits:	
					</td>
					<td bgcolor=white style="width:0%;">
						&nbsp;
					</td>
					<td class=gray align=right>
						<nobr>$<%=dblDeposits%></nobr>
					</td>
					<td class=gray align=right>
						<nobr>$<%=dblDeposits%></nobr>
					</td>
				</tr>
				<tr >
					<td align=right class="svplain8" nowrap colspan="3">
						Budget Transfer Withdrawals:	
					</td>
					<td bgcolor=white  style="width:0%;">
						&nbsp;
					</td>
					<td class=gray align=right>
						<nobr>- $<%=dblWithdraw%></nobr>
					</td>
					<td class=gray align=right>
						<nobr>- $<%=dblWithdraw%></nobr>
					</td>
				</tr>
<%

'Define Where clause.  This logic determines if we show a budget worksheet
'for all courses for a given student or only for a given course

'if intShort_ILP_ID <> "" then
	'Show only for a specific course
'	strWhere = " (ISF.intShort_ILP_ID = " & intShort_ILP_ID & ")"
'else
	'show all courses
	strWhere = " (ISF.intStudent_ID = " & intStudent_ID & _
			   ") AND (ISF.intSchool_Year = " & session.Contents("intSchool_Year") & ")"
'end if 

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
		", oi.bolReimburse, I.intItem_Group_ID, oi.szDeny_Reason , tblVendors.szVendor_Name, " & _
        "             tblVendors.szVendor_Phone, tblVendors.szVendor_Fax, tblVendors.szVendor_Email, tblVendors.szVendor_Website, oi.dtCreate as oiCreate " & _
		"FROM tblClasses INNER JOIN " & _ 
		" tblILP ON tblClasses.intClass_ID = tblILP.intClass_ID LEFT OUTER JOIN " & _ 
		" trefItems I INNER JOIN " & _ 
		" tblOrdered_Items oi ON I.intItem_ID = oi.intItem_ID ON tblILP.intILP_ID = oi.intILP_ID RIGHT OUTER JOIN " & _ 
		" tblILP_SHORT_FORM ISF ON tblILP.intShort_ILP_ID = ISF.intShort_ILP_ID LEFT OUTER JOIN " & _ 
		" tblProgramOfStudies POS ON ISF.lngPOS_ID = POS.lngPOS_ID inner join " & _
		"  trefPOS_SUBJECTS tps ON tps.intPOS_SUBJECT_ID = ISF.intPOS_SUBJECT_ID LEFT OUTER JOIN " & _
		"  tblVendors ON oi.intVendor_ID = tblVendors.intVendor_ID LEFT OUTER JOIN " & _             
		" tblINSTRUCTOR INS ON tblClasses.intInstructor_ID = INS.intINSTRUCTOR_ID left outer join" & _
		" tblGUARDIAN g ON tblClasses.intGuardian_ID = g.intGUARDIAN_ID " & _
		"WHERE " & strWhere & _
		" ORDER BY isSponsor, POS.txtCourseTitle, ISF.szCourse_Title, ISF.intShort_ILP_ID "
	'if ucase(session.Contents("strUserID")) = "SCOTT" then
'		response.Write sql 
'	end if	
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
				strClassHeader = "TableHeaderGreen"
			case false
				strStatus = "a-must amend"
				strClassHeader = "TableHeaderRed"
		end select
		
		if strStatus = "" then
			select case rsBudget("sStatus")
				case true
					strStatus = "s-appr"
					strClassHeader = "TableHeader"
				case false
					strStatus = "s-must amend"
					strClassHeader = "TableHeaderRed"
			end select							
		end if
				
		' unlocks Packet so it can be deleted
		if strStatus = ""  then
			if not oFunc.LockYear then
				bolLock = false	
			end if
		end if
		
		if rsBudget("intILP_ID") & "" <> "" AND strStatus = "" then	
			if rsBudget("bolREady_For_Review") = true then
				strStatus = "ready for sponsor"
				strClassHeader = "TableHeaderPurple"
			else				
				strStatus = "implemented"	
				strClassHeader = "TableHeaderBlue"											
			end if
		elseif strStatus = "" then
			strStatus = "planned"	
			strClassHeader = "SubHeader"
		end if
		strStatus2 = ""
		' now handle status for admins and teachers
		if (ucase(session.Contents("strRole")) = "ADMIN" or ucase(session.Contents("strRole")) = "TEACHER") _
			and strStatus <> "planned" then
			bolMatch = "" 	
			if ucase(session.Contents("strRole")) = "ADMIN" then
				strApprList = "implemented,ready for sponsor,s-must amend,s-appr,a-must amend,a-appr"
				bolMatch = oFunc.TrueFalse(rsBudget("aStatus"))				
				if bolMatch = "1" then bolMatch = "a-appr"
				if bolMatch = "0" then bolMatch = "a-must amend"
				if bolMatch = "" then bolMatch = oFunc.TrueFalse(rsBudget("sStatus"))
				if bolMatch = "1" then bolMatch = "s-appr"
				if bolMatch = "0" then bolMatch = "s-must amend"
				if bolMatch = "" then bolMatch = strStatus								
			elseif strStatus <> "a-appr" then				
				strApprList = ",implemented,ready for sponsor,s-must amend,s-appr"
				bolMatch = oFunc.TrueFalse(rsBudget("sStatus"))
				if bolMatch = "1" then bolMatch = "s-appr"
				if bolMatch = "0" then bolMatch = "s-must amend"
				if bolMatch = "" then bolMatch = strStatus
				if strStatus = "a-must amend" then
					strStatus2 = "<font color='red'><b>" & strStatus & "</b></font><BR>"
					bolMatch = "no match"
				else
					strStatus2 = ""
				end if			
			end if	
			
			if bolMatch <> "" then
				' bolMatch remains "" when the status is a-apprv or a-must amend and the role
				' is a teacher.  In that case we will not provide a drop down but simply display the 
				' status
				strStatus = "<select name=""bolApproved" & rsBudget("intILP_ID") & """ onChange=""jfILPStatus('" & rsBudget("intILP_ID") & "');"" style=""font:arial;font-size:10;"">" & _
							oFunc.MakeList(strApprList,strApprList,bolMatch) & _
							"</select>"	
			end if
	    elseif ucase(session.Contents("strRole")) = "GUARD" and (strStatus = "implemented" or strStatus = "ready for sponsor" or instr(1,strStatus,"must amend") > 0) then					
				if strStatus = "a-must amend" or 	strStatus = "s-must amend" then
					strStatus2 = "<font color='red'><b>" & strStatus & "</b></font><BR>"
					bolMatch = "not match"
				else
					strStatus2 = ""
					bolMatch = strStatus
				end if
				strStatus = "<select name=""bolApproved" & rsBudget("intILP_ID") & """ onChange=""jfUpdateStatus(this,'" & rsBudget("intILP_ID") & "');"" style=""font:arial;font-size:10;"">" & _
							oFunc.MakeList(",implemented,ready for sponsor",",implemented,ready for sponsor",bolMatch) & _
							"</select>"										
		end if						    
	    
		mDivCount = mDivCount + 1
		strBList = strBList & mDivCount & ","
		'if intShort_ILP_ID & "" = rsBudget("intShort_ILP_ID") & "" then
		'	strClassHeader = "SubHeader"
		'else
		'	strClassHeader = "TableHeader"
		'end if 
%>	
				<a name="<% = rsBudget("intShort_ILP_ID") %>"></a>
				<tr>
					<td colspan="11" style="width:100%;">	
						<table style="width:100%;" cellpadding='2' cellspacing='1' ID="Table7">
							<tr class="<% = strClassHeader %>">
								<td align=left style="width:50%;">
									&nbsp;<b>Course Title</b>
								</td>
								<td align='center' style="width:30%;">
									<b>Subject</b>
								</td>
								<td align='center' nowrap style="width:0%;">
									&nbsp;<b>Hrs</b>&nbsp;
								</td>								
								<td align='center' style="width:0%;">
									<b>Actions</b>
								</td>
							</tr>
							<tr>
								<td class="TableCell" valign="top"  style="width:50%;padding-left:8px;BACKGROUND-COLOR:#E9E9FF;color:#404040;">
									 <b><% = rsBudget("txtCourseTitle") & rsBudget("szCourse_Title")%></b>
								</td>
								<td class="TableCell" valign="top" style="width:30%;">
									 <% = rsBudget("szSubject_Name") %>
								</td>
								<td class="TableCell" align='center' valign="top" style="width:0%;">
									 <% = intHours %>
								</td>								
								<td class="TableCell" align='right' nowrap valign="top">
									 <select name='action<%= rsBudget("intShort_ILP_ID")%>' style="font:arial;font-size:10;" ID="Select2">
									<option value="">
									<option value="">- - - - - - - - 
									<% if rsBudget("intILP_ID") & "" = "" then %>
									<option value="delete">Delete Plan</option>	
									<option value="edit">Edit Plan</option>						
									<option value="jfAddClass('<% = intStudent_ID%>','<%=rsBudget("intShort_ILP_ID")%>');">Implement Plan</option>
									<% else %>						
									<option value="jfContractSchedule('<%=rsBudget("intClass_ID")%>','<%=rsBudget("intInstructor_ID")%>','<%=rsBudget("intInstruct_Type_ID")%>','<%=rsBudget("intContract_Guardian_ID")%>','<%=rsBudget("intGuardian_ID")%>','<%=rsBudget("intVendor_ID")%>');"><% = strContractSchedule %></option>						
									<% if not bolLock or ucase(session.Contents("strRole")) = "ADMIN" then%>
									<option value="jfDeleteILP('<% =rsBudget("intILP_ID")%>');">Delete Class</option>						
									<% end if %>
									<option value="jfViewRoll('<%=rsBudget("intClass_ID")%>');">Enrollment List</option>	
									<option value="jfViewCosts('<%= intStudent_ID %>','<%=rsBudget("intILP_ID")%>','<%=rsBudget("intClass_ID")%>');">Goods/Services</option>
									<option value="jfViewILP('<%=rsBudget("intILP_ID")%>','<%=rsBudget("intClass_ID")%>','<%=replace(rsBudget("szClass_Name"),"'","\'")%>','<%=rsBudget("intContract_Guardian_ID")%>','<%=rsBudget("intVendor_ID")%>','<%=rsBudget("teacherName")%>');">ILP</option>																		
									<% end if %>												
								</select>
								<input type=button value="go" onclick="jfCallAction('<%= rsBudget("intShort_ILP_ID")%>');" class="btSmallGray" NAME="Button1">&nbsp;
								</td>								
							</tr>
							<tr>
								<td colspan="4" style="width:100%;">
									<table style="width:100%;" cellpadding=0 cellspacing=1>
										<tr>
											<td class="TableCell" align='center' nowrap style="width:10%;BACKGROUND-COLOR:#E9E9FF;color:#404040;"">
												&nbsp;<b>ILP Status</b>&nbsp;
											</td>
											<td class="TableCell" align='center' style="width:90%;BACKGROUND-COLOR:#E9E9FF;color:#404040;"">
												<table cellpadding="0" cellspacing="0" style="width:100%;">
													<tr>
														<% if ucase(session.Contents("strRole")) <> "GUARD" then %>
														<td style="width:0%;">
															<input type=submit value="Save" class="NavLinkltPurple" ID="Submit1" NAME="Submit1">	
														</td>																
														<% end if %>
														<td style="width:100%;" align="center" class="svplain8">
															<b>ILP Comments</b>
														</td>
													</tr>
												</table>																								
											</td>
										</tr>
										<tr>
											<td class="TableCell" align='center' valign="middle" nowrap>
												<% = strStatus2 & strStatus %>
											</td>	
											<td class="TableCell" valign="top" style="width:90%;" align="center">
												<% 
													if ucase(session.Contents("strRole")) = "ADMIN" then
														' admins view of comments
														response.Write "<textarea style='width:99%;' rows='1' wrap='virtual' name='szComments" & rsBudget("intILP_ID") & "' onfocus='this.rows=4;' onblur='this.rows=1;' onKeyDown='jfMaxSize(1999,this);' " &_
																		" onChange=""jfILPStatus('" & rsBudget("intILP_ID") & "');"">" & _
																		rsBudget("szAdmin_Comments") & "</textarea>"
																	    
														if rsBudget("szSponsor_Comments") <> "" then 
															response.Write "<BR>"
															response.Write "Sponsor Comments: " & rsBudget("szSponsor_Comments") 
														end if 
													elseif ucase(session.Contents("strRole")) = "TEACHER" then
														' teachers view of comments
														response.Write "<textarea style='width:99%;' rows='1' wrap='virtual' name='szComments" & rsBudget("intILP_ID") & "' onfocus='this.rows=4;' onblur='this.rows=1;' onKeyDown='jfMaxSize(1999,this);' " & _
																		" onChange=""jfILPStatus('" & rsBudget("intILP_ID") & "');"">" & _
																		rsBudget("szSponsor_Comments") & "</textarea>"
														if rsBudget("szAdmin_Comments") <> "" then 
															response.Write "<BR>"
															response.Write "Admin Comments: " & rsBudget("szAdmin_Comments") 
														end if 
													else
														' Guardians view of comments
														if rsBudget("szAdmin_Comments") <> "" then
															response.Write "Admin Comments: " & rsBudget("szAdmin_Comments") 
														end if 
														
														if rsBudget("szSponsor_Comments") <> "" then
															if rsBudget("szAdmin_Comments") <> "" then response.Write "<BR>"
															response.Write "Sponsor Comments: " & rsBudget("szSponsor_Comments") 
														end if 
													end if																				
												%>&nbsp;
											</td>	
										</tr>
									</table>
								</td>
							</tr>
						</table>				
						<nobr>										
					</td>
					<td  class="ltGray" colspan="3"  style="width:0%;">
						&nbsp;
					</td>
				</tr>				
				<tr id="div<% = mDivCount%>">
					<td class="TableSubHeader" align="center" title="If line items have been charged against your budget they can be viewed by clicking on 'show'.">
						Line Items
					</td>	
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
					<td class="TableSubHeader" align="center" title="(QTY * Unit Cost) + S/H">
						Budget Total
					</td>
					<td class="TableSubHeader" align="center" title="Sum of all line items (charged expeneses) entered by the office for a specific budget.">
						Actual Charges
					</td>
					<td class="TableSubHeader" align="center" title="Adjustments are needed to handle over expenditures and to release unused budgeted funds once the budget is closed.">
						Budget Adjust
					</td>
					<td class="TableSubHeader" align="center" title="(Budget Total - Actual Charges) + Budget Adjust">
						Budget Balance
					</td>
					<td  class="ltGray"  style="width:0%;">
						&nbsp;
					</td>
					<td  class="ltGray">
						&nbsp;
					</td>
					<td  class="ltGray">
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
					<td class="TableCell">
						&nbsp;
					</td>
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
					<td class="<% = strClass %>" align="right" title="Teachers Hourly Rate">
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
					<td  class="ltGray" style="width:0%;">
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
		
		dblCharge = formatNumber(liInfo(1),2)
		dblAdjBudget = formatNumber(dblBudgetCost + cdbl(liInfo(2)),2)		
		mDivCount = mDivCount + 1
		strBList = strBList & mDivCount & ","
		
		if bStatus = "rejc" then
			strClass = "TableCellStrike"
		else
			strClass = "TableCell"
			dblClassCharge = dblClassCharge + cdbl(dblCharge)
			dblClassBudget = dblClassBudget + cdbl(dblAdjBudget)
		end if
		
		if rsBudget("szDeny_Reason") <> "" then
			strReason = "<BR><b>Comment:</b> " & rsBudget("szDeny_Reason")
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
					<td class="<% = strClass %>" align="center" nowrap>
						<% if liInfo(3) <> "" then%>
						<a href="javascript:" onclick="jfToggle('<%=liInfo(3)%>','a<%=mLablelCount%>');" id="a<%=mLablelCount%>">show</a>
						<% 
							mLablelCount = mLablelCount + 1
						else
							response.Write "&nbsp;"
						end if 
						%>
					</td>
					<td class=<% = strClass %>>
						<% = rsBudget("szName") %>
					</td>
					<td class="<% = strClass %>" align="center">
						<% = bStatus %>
					</td>
					<td class=<% = strClass %>>
						<% response.Write oHtml.ToolTip(strItemType & rsBudget("oiDesc") & strReason, _
							  "<table cellpadding='2'><tr><td class='svplain8' valign='top'><b>Vendor Name:</b></td><td class='svplain8' nowrap>" & rsBudget("szVendor_Name") & "</td></tr>" & _
													 "<tr><td class='svplain8' nowrap><b>Phone Number:</b></td><td class='svplain8' nowrap>" & oFunc.Reformat(rsBudget("szVendor_Phone") , Array("(", 3, ") ", 3, "-", 4)) & "</td></tr>" & _
													 "<tr><td class='svplain8' nowrap><b>Fax Number:</b></td><td class='svplain8' nowrap>" & oFunc.Reformat(rsBudget("szVendor_Fax") , Array("(", 3, ") ", 3, "-", 4))  & "</td></tr>" & _
													 "<tr><td class='svplain8' nowrap><b>Vendor Email:</b></td><td class='svplain8' nowrap>" & rsBudget("szVendor_Email") & "</td></tr>" & _
													 "<tr><td class='svplain8' nowrap><b>Budget Created:</b></td><td class='svplain8' nowrap>" & rsBudget("oiCreate") & "</td></tr></table>", _
													 false, "",false,"tooltip","","",false,false)%>&nbsp;
					</td>
					<td class=<% = strClass %> align="center" nowrap>
						<% = rsBudget("intQTY") %>
					</td>
					<td class=<% = strClass %> align=right nowrap>
						$<% = formatNumber(rsBudget("curUnit_Price"),2) %>
					</td>
					<td class=<% = strClass %> align=right nowrap title="Shipping and Handling">
						&nbsp;$<% = dblShipping %>
					</td>
					<td class=<% = strClass %> align=right nowrap title="(QTY * Unit Cost) + S/H">
						$<% = dblBudgetCost %>
					</td>
					<td class=<% = strClass %> align=right nowrap title="Sum of all line items (charged expeneses) entered by the office for a specific budget.">
						$<% = formatNumber(liInfo(1),2)%>
					</td>
					<td class=<% = strClass %> align=right nowrap title="Adjustments are needed to handle over expendatures and to release unused budgeted funds once the budget is closed.">
						$<% = formatNumber(liInfo(2),2) %>
					</td>
					<td class=<% = strClass %> align=right nowrap title="(Budget Total - Actual Charges) + Budget Adjust">
						$<% = formatNumber((dblBudgetCost - cdbl(liInfo(1))) + cdbl(liInfo(2)),2)%>
					</td>
					<td bgcolor=white style="width:0%;">
						&nbsp;
					</td>
					<td class="<% = strClass %>" align="right" nowrap title="Budget Total - Budget Adjust">
						-$<% = dblAdjBudget %>
					</td>
					<td class="<% = strClass %>" align="right" nowrap title="Actual Charges">
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
					<td class=svplain10 colspan=10>
						No Goods or Services have been budgeted for this course.	
					</td>
					<td bgcolor=white style="width:0%;">
						&nbsp;&nbsp;&nbsp;
					</td>
					<td  class="ltGray">
						&nbsp;
					</td>
					<td  class="ltGray">
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
					<td class=svplain10 colspan="11" align=right>
						Available Remaining Funds:	
					</td>
					<td bgcolor=white  style="width:0%;">
						&nbsp;&nbsp;&nbsp;
					</td>
					<td class="Gray" align=right>
						$<%=formatNumber(dblTargetBalance,2)%>
						<input type=hidden name="budgetBalance" value="<%=formatNumber(dblTargetBalance,2)%>" ID="Hidden8">
					</td>
					<td class="Gray" align=right>
						$<%=formatNumber(dblActualBalance,2)%>
					</td>
				</tr>
<script language=javascript>
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
response.Write oHtml.ToolTipDivs
call oFunc.CloseCN()
set oFunc = nothing
set oHtml = nothing
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")

sub vbsShowTotals()
	
	if ilp_ID & "" = "" then		
		'dblClassCharge = "0.00" 
	end if
	
	if ilpShortID & "" = "" then
		'dblClassBudget = "0.00"
	end if 
	' ADD THESE LINES TO MAKE COURSE TOTALS HIDDEN
	'mDivCount = mDivCount + 1
	'strBList = strBList & mDivCount & ","
	'id="div<%=mDivCount" (NEEDS TO BE ADDED TO <TR> TAG IN HTML BELOW)
%>
				<tr class=svplain10 bgcolor="<% = strColor%>" >
					<td colspan="11" align=right>					
						&nbsp;<b>Course Totals:</b>
					</td>
					<td bgcolor=white  style="width:0%;">
						&nbsp;&nbsp;&nbsp;
					</td>
					<td class=gray align=right>
						<nobr>
						<% if instr(1,dblBudgetCost,"-") > 0 then
								response.Write "+ $" & formatNumber(replace(dblClassBudget,"-",""),2)
						   else
								response.Write "- $" & formatNumber(dblClassBudget,2)
						   end if						
						%></nobr>
					</td>
					<td class=gray align=right>
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
					<td colspan="20">
						&nbsp;
					</td>
				</tr>	
<%
	'response.Write dblTotalCharge & " - " &  dblClassCharge
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
			" (curUnit_Price * intQuantity) + curShipping as Total, dtCREATE, szCheck_Number " & _
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
		if rs("szCheck_Number") & "" <> "" then
			szCheck_Number = "Check #: " & rs("szCheck_Number")
			if rs("szLine_Item_desc") & "" <> "" then
				szCheck_Number = "<BR>" & szCheck_Number
			end if
		else
			szCheck_Number = ""
		end if
			
		sHtml = sHtml & "<tr id='div" & mDivCount & "' style='display:none;'>" & _
				"<td>&nbsp;</td><td colspan='2'  class='TableCellContrast'>Entered: " & formatDateTime(rs("dtCREATE"),2) & "</td>" & _
				"<td class='TableCellContrast' >" & rs("szLine_Item_desc") & szCheck_Number & "</td>" & _
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

sub vbsDelete(id,pStudent_ID)
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
		' Now delete the Short Form
		delete = "delete from tblILP_Short_Form " & _
				"WHERE intShort_ILP_ID = " & id & _ 
				" AND intStudent_ID = " & pStudent_ID
				
		oFunc.ExecuteCN(delete)	 
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
	if bolApproved = "ready for sponsor" then
		update = "update tblILP set bolApproved=Null, bolSponsor_Approved=NULL, " & _
				 "bolReady_For_Review = 1,dtReady_For_Review = CURRENT_TIMESTAMP " & _
				 " Where intILP_ID = " & ilp_ID
		strApproved = "go"	
	elseif bolApproved = "implemented" then
		update = "update tblILP set bolApproved=Null, bolSponsor_Approved=NULL, " & _
				 "bolReady_For_Review = NULL,dtReady_For_Review = NULL " & _
				 " Where intILP_ID = " & ilp_ID
		strApproved = "go"		 
	else
		if instr(1,bolApproved,"s-") > 0 then 
			strApproved = "bolSponsor_Approved"
			strDateField = "dtSponsor_Approved"
			if ucase(session.Contents("strRole")) = "ADMIN" then
				strResetAdmin = " ,bolApproved = NULL , dtApproved = NULL "			
			end if
		elseif instr(1,bolApproved,"a-") > 0 then 
			strApproved = "bolApproved"
			strDateField = "dtApproved"
		else
			strResetAdmin = ""
		end if
		
		if bolApproved = "" or bolApproved = "implemented" then bolApproved = "Null" 
		if instr(1,bolApproved,"appr") > 0 then bolApproved = 1
		if instr(1,bolApproved,"must amend") > 0 then bolApproved = 0
			
		update = "update tblILP set " & strApproved & " = " & bolApproved & _
				", " & strDateField & " = CURRENT_TIMESTAMP " &  strResetAdmin & _
				" Where intILP_ID = " & ilp_ID
	end if
	if strApproved <> "" then
		oFunc.ExecuteCN(update)
    end if
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
		GetBudgetStatus = "<font color='green'><b>pick up</b></font>"
	elseif pBolLineItems = true and pIsReimburse = false and pItemGroup = 1 then
		GetBudgetStatus = "pymt made"
	elseif pBappr = true and pBolLineItems = false and pIsReimburse = false and pItemGroup = 2 then
		GetBudgetStatus = "ordered"
	elseif pBappr = true and pBolLineItems = false and pIsReimburse = false and pItemGroup = 1 then
		GetBudgetStatus = "vend appr"
	elseif pBappr = true and pBolLineItems = true and pIsReimburse = true then
		GetBudgetStatus = "check cut"
	else
		GetBudgetStatus = "pend"
	end if
end function
%>