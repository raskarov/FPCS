<%@  language="VBScript" %>
<%
Response.LCID=1033


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
dim oFunc				' wsc object
dim dblTargetBalance	' Target start - all budgeted expenses
dim dblActualBalance	' Actual start - all actual expenses
dim dblWithdraw			' Amount to reduce budget fundling by due to Budget Transfer withdrawal 
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

Dim dblGlobalReserve    'Global Reserve for the current School Year

mLablelCount = 0
mDivCount = 0		

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()
oFunc.ResetSelectSessionVariables
session.Contents("intStudent_ID") = ""

set oBudget = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/StudentBudgetInfo.wsc"))

set oHtml = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/htmlFuncs.wsc"))

'Initialize some key variables
if request("intStudent_ID") <> "" then
	intStudent_ID = request("intStudent_ID") 
	intShort_ILP_ID = request("intShort_ILP_ID")
	
	' Crucial updates to make prior to getting student funding info
	if request.QueryString("bolDelete") <> "" then
		' Handle deletion if needed
		call vbsDelete(request.QueryString("intShort_ILP_ID"),request.QueryString("intStudent_ID"))
	elseif request("intEnroll_Info_ID") <> "" and request("changePercent") <> "" then
		'Handle updating of Percent Enrolled
		call vbsUpdateEnrollPercent(request("intPercent_Enrolled_Fpcs"),request("intEnroll_Info_ID"))
	end if
	
	if ucase(session.Contents("strRole")) = "ADMIN" then
		if request("bolLock") & "" <> "" then
			call vbsLockEnrollLevel(request("intEnroll_Info_ID"),request("intPercent_Enrolled_Fpcs"))
		end if
		
		if request("updateTestForm") <> "" then
			call vbsUpdateTestForm(request("bolASD_Testing"), request("intEnroll_Info_ID"))
		end if
		
		if request("updateProgressForm") <> "" then
			call vbsUpdateProgressForm(request("bolProgress"), request("intEnroll_Info_ID"))
		end if
	end if
	
	oBudget.PopulateStudentFunding Application("cnnFPCS"),intStudent_ID,session.Contents("intSchool_Year")
	'oBudget.PopulateStudentFunding oFunc.FPCScnn,intStudent_ID,session.Contents("intSchool_Year")
	
	dblDeposits = oBudget.Deposits
	dblWithdraw = oBudget.Withdrawls
	dblActualBalance = oBudget.ActualFunding
	dblTargetBalance = oBudget.BudgetFunding 	
	intEnroll_Info_ID = oBudget.EnrollInfoId	
	myBudgetBalance = oBudget.BudgetBalance
	myActualBalance = oBudget.ActualBalance

    Dim sqlGR
    sqlGR="select curFund_Amount from dbo.tblFunding where szGrade='G' and intSchool_Year=" & session.Contents("intSchool_Year")
    Dim rsGR
    Set rsGR = Server.CreateObject("ADODB.RecordSet")
    rsGR.CursorLocation = 3
	rsGR.Open sqlGR, Application("cnnFPCS")'
    If Not rsGR.EOF and Not rsGR.BOF Then
    dblGlobalReserve = rsGR("curFund_Amount")
    End If
    rsGR.Close
    Set rsGR = Nothing
else
	'terminate page since page was improperly called.
	response.Write "<html><body><H1>Page improperly called.</h1></body></html>"
	response.End
end if

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Handle Data Modifications
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 strMessage = request.QueryString("strMessage")
 
if request("strILPList") <> "" then	
	call vbsUpdateILPStatus(request("strILPList"))
end if

if request("strAlertList") <> "" then	
	call vbsUpdateAlerts(request("strAlertList"), "bolSponsorAlert", "Alert")
end if

if request("strParentList") <> "" then	
	call vbsUpdateAlerts(request("strParentList"), "bolParentAlert", "ParentAlert")
end if
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'prevent page from cacheing
Response.AddHeader "Cache-Control","No-Cache"
Response.Expires = -1

'Print the header
Session.Value("strTitle") = "Student Packet/Budget Worksheet"
Session.Value("strLastUpdate") = "26 Oct 2004"

if oBudget.FamilyId & "" = "" or oBudget.StudentGrade & "" = ""then
	Server.Execute(Application.Value("strWebRoot") & "includes/Header.asp")
%>
<table cellspacing="0" cellpadding="4" width="85%" id="Table12">
    <tr>
        <td class="svplain10">
            <% if oBudget.FamilyId & "" = "" then %>
            <b>This student does not belong to a family in the Student Information System.</b>
            <br/>
            An Administrator will need to add the student to a family before work on the packet
            can begin.
            <% else %>
            <b>A grade has not been selected for this student.</b><br/>
            Before work can begin on the packet you will need to go to the student profile and
            enter the students' current grade.
            <%end if%>
        </td>
    </tr>
</table>
<%
	call oFunc.CloseCN()
	set oFunc = nothing
	set oBudget = nothing
	Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")
	response.End
elseif isNumeric(oBudget.EnrollmentId) then 'and isNumeric(oBudget.IepId) then
	' Student Profiles have been updated by family. Now we check to see if a sponsor has been selected.
	if not isNumeric(oBudget.SponsorID)  then Sponsor = "No Sponsor Selected": SponsorEmail=""
else
	if request("print") <> "" or request("simpleHeader") <> "" then
		Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")
		strSimpleHeader = "simpleHeader=true&"			
	else
		Server.Execute(Application.Value("strWebRoot") & "includes/Header.asp")
	end if
%>
<table cellspacing="0" cellpadding="4" width="85%" id="Table4">
    <tr>
        <td class="svplain10">
            <b>Before you can plan a course you must update your students information for SY
                <% = session.Contents("intSchool_Year")%>. To do this click <a href="<% = Application.Contents("strSSLWebRoot") & "admin/familyManager.asp?intFamily_ID=" & oBudget.FamilyId %>">
                    HERE</a> to go to the Family Manager (also found on the above menu under 'Profiles').
            </b>
        </td>
    </tr>
</table>
<%
	
	call oFunc.CloseCN()
	set oFunc = nothing
	set oBudget = nothing
	Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")
	response.End
end if

set rsStudentInfo = nothing

'Find out if student is in High School
if isNumeric(oBudget.StudentGrade) then
	if cint(oBudget.StudentGrade) >= 9 then
		bolHighSchool = true
	else
		bolHighSchool = false
	end if
end if

strColorTable = "<table>" & _
			    "<tr><td class='SubHeader' nowrap>&nbsp;&nbsp;</td><td class='svplain8'><b>Planned</b></td></tr>" & _
			    "<tr><td class='TableheaderBlue' nowrap>&nbsp;&nbsp;</td><td class='svplain8'><b>Not Fully Signed</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>" & _
			    "<tr><td class='TableHeaderBlack' nowrap>&nbsp;&nbsp;</td><td class='svplain8'><b>Rejected</b></td></tr>" & _
			    "<tr><td class='TableheaderRed' nowrap>&nbsp;&nbsp;</td><td class='svplain8'><b>Must Amend<b></td></tr>" & _			    
			    "<tr><td class='TableHeaderGreen' nowrap>&nbsp;&nbsp;</td><td class='svplain8'><b>Fully Signed</b></td></tr>" & _
			    "<tr><td class='TableHeaderGrape' nowrap>&nbsp;&nbsp;</td><td class='svplain8'><b>Sponsor Alert</b></td></tr>" & _
			    "<tr><td class='TableHeaderTeal' nowrap>&nbsp;&nbsp;</td><td class='svplain8'><b>Parent Alert</b></td></tr>" & _
			    "</table>" 
			    								
	
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

if request("print") <> "" or request("simpleHeader") <> "" then
	Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")
	session.Contents("strSimpleHeader") = "&simpleHeader=true&"
	if request("print") <> "" then
%>
<link rel="stylesheet" href="<% =  Application.Value("strWebRoot") %>CSS/printStyle.css">

<script language="javascript">
    function jfPrint() {
        if (window.print) {
            window.print()
            var obj = document.getElementById("btPrint");
            obj.style.display = "none";
        }
        else {
            alert("Mac users: please press Apple-P to print this form.\nWindows users: Please press ctrl-P to print this form.")
        }
    }		
				
</script>

<input name="btPrint" type="button" value="Print This Page" class="btPrint" onclick="jfPrint();"
    id="Button1"/>
<%
	end if
else
	Server.Execute(Application.Value("strWebRoot") & "includes/Header.asp")
	session.Contents("strSimpleHeader") = ""
end if

' Print top section of worksheet
%>

<script language="javascript" type="text/javascript">	
	function jfOrder(ilpID){
		SetPageScroll('<% = intStudent_id %>Packet');
		var strURL = '<%=Application.Value("strWebRoot")%>forms/Requisitions/req1.asp?intStudent_ID=<%=intStudent_ID%>&intILP_ID=' + ilpID;
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
	</script>
    <script type="text/javascript">
	function jfCallAction(id){
		SetPageScroll('<% = intStudent_id %>Packet');
		var strAction = eval('document.forms["main"].action'+id+".value;");
		
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
	
	function jfOpen(id){
		// Opens up edit/add course window
		var winSF;
		var URL = '<%=Application.Value("strWebRoot")%>forms/packet/addCourse.asp?intStudent_ID=<%=intStudent_ID%>&GRADE=<%= oBudget.StudentGrade %>&bolHighSchool=<%=bolHighSchool%>&intShort_ILP_ID=' + id;
		winSF = window.open(URL,"winSF","width=640,height=250,scrollbars=yes,resizable=yes");
		winSF.moveTo(0,0);
		winSF.focus();
	}

	function jfDelete(id){
		SetPageScroll('<% = intStudent_id %>Packet');
		window.location.href='<%=Application.Value("strWebRoot")%>forms/packet/packet.asp?bolDelete=true&intStudent_ID=<%=intStudent_ID%>&intShort_ILP_ID='+id+'<%=session.Contents("strSimpleHeader")%>';	
	}
	
	<%	 
	if strMessage <> "" then
		response.write "alert('" & strMessage & "');"
		strMessage = "" 
	end if		
	%>
	
	function jfViewAnotherStudent(id){
		SetPageScroll('<% = intStudent_id %>Packet');
		var strURL = '<%=Application.Value("strWebRoot")%>forms/packet/packet.asp?intStudent_ID=' + id.value;
		window.location.href = strURL;
	}
	
	function jfChangePercent(enrollId,percent){
		SetPageScroll('<% = intStudent_id %>Packet');
		var bolConfirm = window.confirm("Are you sure you want to change the 'Target Enrollment' percentage?");
		if (bolConfirm){
			var URL = '<%=Application.Value("strWebRoot")%>forms/packet/packet.asp?intStudent_ID=<%=intStudent_ID%>&';
			URL += "intEnroll_Info_ID=" + enrollId;
			URL += "&changePercent=true&intPercent_Enrolled_Fpcs=" + percent; 
			window.location.href = URL;
		}else{
			// returns selection back to original value
			document.forms["main"].intPercent_Enrolled_Fpcs.selectedIndex = document.forms["main"].lastIndex.value;
		}
	}
	
	function jfGetIndex(obj){
		// Stores original value
		document.forms["main"].lastIndex.value = obj.selectedIndex;	
	}
	
	function jfViewCosts(studentID,ilpID,classID){
		SetPageScroll('<% = intStudent_id %>Packet');
		var strURL = '<%=Application.Value("strWebRoot")%>forms/Requisitions/req1.asp?intClass_ID=' + classID;
		strURL += "&intStudent_ID=" + studentID + "&intILP_ID=" + ilpID;
		costsWin = window.open(strURL,"costsWin","width=710,height=500,scrollbars=yes,resize=yes,resizable=yes");
		costsWin.moveTo(0,0);
		costsWin.focus();
	}
	
	function jfViewILP(ilp_id,class_ID,class_name,cg,vendor,teacherName) {
		SetPageScroll('<% = intStudent_id %>Packet');
		var ilpWin;
		var strURL;
		var strILP;
				
		strURL = '<%=Application.Value("strWebRoot")%>forms/ILP/ilpMain.asp?isPopUp=yes&intILP_ID=' + ilp_id + "&intClass_id=" + class_ID;
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
		var strURL = '<%=Application.Value("strWebRoot")%>forms/teachers/classAdmin.asp?bolInWindow=true&isPopUp=yes<%=strDisabled%>&intClass_id='+class_id;
		strURL += "&intInstructor_id="+instructor_id+"&intInstruct_Type_ID="+instruct_type;
		strURL += "&intContract_Guardian_ID="+intContract_Guardian_ID;
		strURL += "<% = strHideGoodService %>&intGuardian_id="+intGuardian_ID;
		strURL += "&intVendor_ID="+intVendor_ID;
		classWin = window.open(strURL,"classWin","width=750,height=500,scrollbars=yes,resizable=yes");
		classWin.moveTo(0,0);
		classWin.focus();
	}
	
	function jfDeleteILP(ilp_id) {
		SetPageScroll('<% = intStudent_id %>Packet');
		var answer;
		answer = confirm("Are you sure you want to delete this class? (All Goods and Services and the ILP for this class will be deleted as well)");
		if (answer) {
			var winDel;
			winDel = window.open('<%=Application.Value("strWebRoot")%>forms/teachers/deleteClass.asp?intILP_id='+ilp_id+'<% =session.Contents("strSimpleHeader")%>',"winDel","width=200,height=200,scrollbars=yes,resizable=yes");
			winDel.moveTo(0,0);
			winDel.focus();			
		}
	}	
	
	function jfViewRoll(class_id) {
		var winRoll;
		winRoll = window.open('<%=Application.Value("strWebRoot")%>Reports/studentsInClass.asp?intClass_id='+class_id,"winRoll","width=640,height=480,scrollbars=yes,resizable=yes");
		winRoll.moveTo(0,0);
		winRoll.focus();					
	}
	
	function jfAddClass(studentID,shortID){
		// Opens ilp1.asp that will create a class 
		SetPageScroll('<% = intStudent_id %>Packet');
		var URL = '<%=Application.Value("strWebRoot")%>forms/ilp/ILP1.asp?intStudent_ID='+studentID+"&intShort_ILP_ID="+shortID;
		window.location.href = URL;
	}
	
	
	function jfChangeSponsor(studentID){
		//window.location.href = '<%=Application.Value("strWebRoot")%>forms/packet/addSponsorTeacher.asp?intStudent_ID=' + studentID;
		window.location.href = 'addSponsorTeacher.asp?intStudent_ID=' + studentID;
	}
	
	function jfShowComments(acom,scom){
		if (acom != "") { acom = "AA Comments: " + acom + "\n";}
		if (scom != "") { scom = "Sponsor Teacher Comments: " + scom;}
		alert(acom + scom);
	}
	
	function jfSponsorAlert(pID){
		// Tracks the sponsor alert list so we will know which 
		// courses to turn the alert on/off for.
		var sList = document.forms["main"].strAlertList;
		
		if (sList.value.indexOf(","+pID+",") == -1 ) {
			sList.value = sList.value + pID + ",";			
		}
		SetPageScroll('<% = intStudent_id %>Packet');
		document.forms["main"].submit();
	}
	
	function jfParentAlert(pID){
		// Tracks the sponsor alert list so we will know which 
		// courses to turn the alert on/off for.
		var sList = document.forms["main"].strParentList;
		
		if (sList.value.indexOf(","+pID+",") == -1 ) {
			sList.value = sList.value + pID + ",";			
		}
		SetPageScroll('<% = intStudent_id %>Packet');
		document.forms["main"].submit();
	}
	
	function jfILPStatus(pID,pClassName,pObj){
		var sList = document.forms["main"].strILPList;
		var sClassName = document.forms["main"].ClassName;
		
		if (sList.value.indexOf(","+pID+",") == -1 ) {
			sList.value = sList.value + pID + ",";
			if (pClassName != "") {
			sClassName.value = sClassName.value + pClassName + ",";
			}
		}
		var re = new RegExp(pClassName + ",",'gi');
		if (pObj.type != 'textarea') {
			if (pObj.type == 'checkbox') {			
				if (pObj.checked == false) {
				sClassName.value = sClassName.value.replace(re,'');
				}
			}
			else if(pObj.value != 1){
				sClassName.value = sClassName.value.replace(re,'');
			}
		}
	}
	
	function jfUpdateStatus(myVal,ilp_id){
		SetPageScroll('<% = intStudent_id %>Packet');
		var URL = '<%=Application.Value("strWebRoot")%>forms/packet/packet.asp?bolChangeStatus=true&bolApproved='+myVal.value;
		URL += "&intILP_ID="+ilp_id+"&intStudent_ID=<%=intStudent_ID%>"
		window.location.href = URL;
	}
	
	function jfPrintPacket(){
		var strURL;
		<% IF Request("intStudent_id") <> "" then %>
		strURL = "&intStudent_ID=<%=Request("intStudent_id")%>";
		<% end if %>		
		//var winPrintPacket = window.open("<%=Application.Value("strWebRoot")%>forms/PrintableForms/allPrintable.asp?strAction=S" + strURL,"winPrintPacket","width=700,height=500,scrollbars=yes,resize=yes,resizable=yes");
		var winPrintPacket = window.open('<%=Application.Value("strWebRoot")%>forms/PrintableForms/printPacket2.asp?strAction=S' + strURL,"winPrintPacket","width=700,height=500,scrollbars=yes,resize=yes,resizable=yes");
		winPrintPacket.moveTo(0,0);
		winPrintPacket.focus();
	}
	
	function ConfirmSignatures(){
	jfCheckBudget();
		SetPageScroll('<% = intStudent_id %>Packet');
		if (document.forms["main"].ClassName.value != ''){
			var bConfirm = confirm("You are about to sign a digital signature for the following courses ...\n" + document.forms["main"].ClassName.value.replace(/\,/gi,'\n') + "Once signed these courses can not be unsigned. Do you want to continue?");
			if (bConfirm == true){
				document.forms["main"].submit();		
			}	
		}else{
			document.forms["main"].submit();	
		}
	}
		function jfCheckBudget(){
			var total = <%= myBudgetBalance %>;
			if (total < 0){	
				var strMessage;
				strMessage="Warning: Your remaining balance is negative (";
				strMessage+=total;
				strMessage+= ")";
				alert(strMessage);
			}
		}
		function round(number,X) {
			// rounds number to X decimal places, defaults to 2
			X = (!X ? 2 : X);
			return Math.floor(number*Math.pow(10,X))/Math.pow(10,X);
		}
</script>


<form name="main" action="<%=Application("strSSLWebRoot")%>forms/packet/packet.asp"
method="post" id="Form1">
<input type="hidden" name="intStudent_ID" value="<%=intStudent_ID%>" id="Hidden2"/>
<input type="hidden" name="bolHighSchool" value="<%=bolHighSchool%>" id="Hidden3"/>
<input type="hidden" name="courseTitleData" value="" id="Hidden4"/>
<input type="hidden" name="simpleHeader" value="<% = request("simpleHeader") %>"
    id="Hidden5"/>
<input type="hidden" name="lastIndex" value="" id="Hidden6"/>
<input type="hidden" name="ClassName" value=""/>
<table width="100%" id="Table3">
    <tr>
        <td class="yellowHeader">
        <% response.write "" & application.Contents("dtCount_Deadline") 'response.write session.Contents("intSchool_Year")%>
            <table width="100%" cellpadding="0" cellspacing="0" id="Table8">
                <tr class="yellowHeader">
                    <td align="left">
                        &nbsp;<b> Student Packet/Budget for
                            <% = oBudget.StudentName %>
                        </b>&nbsp;&nbsp;Grade:
                        <% = oBudget.StudentGrade %>
                    </td>
                    <td align="right">
                        <% 
							if ucase(session.Contents("strRole")) <> "GUARD" then
								if oBudget.FamilyName & "" = "" then
									response.Write "<B>Family Email:</b> No Email Provided"
								else
									response.Write "<B>Family Email:</b> <a href=""mailto:" & oBudget.FamilyEmail & "?cc=" & oBudget.SponsorEmail & """ style=""color:white;"">" & oBudget.FamilyEmail & "</a>"
								end if								
							end if
                        %>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
    <% if oBudget.ReEnrollState = "129" then 
		sql = "SELECT NeededEnrollInfoCD, Description, " & _ 
				"	(SELECT	StudentEnrollNeedID " & _ 
				"	FROM	STUDENT_ENROLL_INFO_NEEDED se " & _ 
				"	WHERE	se.NeededEnrollInfoCD = ne.NeededEnrollInfoCD AND " & _
				"	se.StudentID = " &  intStudent_ID & " AND se.SchoolYear = " & session.Contents("intSchool_Year") & ") AS Needed " & _ 
				"FROM	trefNEEDED_ENROLL_INFO ne " & _ 
				"ORDER BY OrderID "
		dim rsNeeded
		set rsNeeded = server.CreateObject("ADODB.RECORDSET")
		rsNeeded.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
		
		strMoreConditional = "<table><tr><td class='svplain8'>Common reasons for Conditional Enrollment: (items still needed in red)<br/>" & _
			"<ul>" 
		
		do while not rsNeeded.EOF
			if rsNeeded("Needed") & "" <> "" then
				sNeedStyle = " style='color:red;' "
			else
				sNeedStyle = ""
			end if 
			strMoreConditional = strMoreConditional & "<li " & sNeedStyle & ">" & rsNeeded("Description") & "</li>"
			rsNeeded.MoveNext
		loop
		
		rsNeeded.Close
		set rsNeeded = nothing
		
		strMoreConditional = strMoreConditional &"</ul>" & _
			"Call the FPCS office (742-3705) for the specific reason(s) why this student 			" & _
			"is conditionally enrolled.</td></tr></table>"
		
    %>
    <tr>
        <td colspan="13" class="svplain10" bgcolor='#ff8080'>
            <br/>
            <b>Please Note: </b>This student is currently <b>Conditionally Enrolled</b>.<br/>
            Spending has been locked until the Conditional Enrollment status is lifted.
            <% response.Write oHtml.ToolTip("<a href='#' class='svplain10'>Click HERE to find out more.</a>",strMoreConditional,true,"Addition Enrollment Status Information",false,"tooltip","250px","",false,true) %>
            <br/>
            <br/>
        </td>
    </tr>
    <% end if %>
    <tr>
        <td bgcolor="#ededed">
            <table width="95%" align="left" id="Table6">
                <tr>
                    <td colspan="10">
                        <font class="svplain10">
                            <% if not oFunc.LockYear then
                            %>
                            <%'JD 052511 move plan new course button %>
                            <!--<input type="button" value="Plan New Course" class="NavLink" onclick="jfOpen('');"
                                name="Button2" id="Button4">-->
                            <input type="button" value="Hide Detail" class="NavLink" onclick="jfToggleBudget(this);"
                                name="Button2" id="Button5"/>
                            <input type="button" value="Print Packet" onclick="jfPrintPacket();" class="NavLink"
                                id="Button6" name="Button6"/>
                            <%
							end if
                            %>
                            <input type="hidden" name="strILPList" value="," id="Hidden7"/>
                            <input type="hidden" name="strAlertList" value=","/>
                            <input type="hidden" name="strParentList" value="," id="Hidden9"/>
                            <input type="button" value="Save Status & Comments" onclick="ConfirmSignatures();"
                                class="NavSave" style="width: 165px;" id="Submit1" name="Submit1"/>
                        </font>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
    <tr>
        <td>
            <table>
                <tr>
                    <td valign="top">
                        <table id="Table2">
                            <tr>
                                <td valign='top' style='height: 100%;'>
                                    <table id="Table5" cellspacing='1' cellpadding='4' style='height: 100%;'>
                                        <tr>
                                            <td class="TableHeader" align="center">
                                                <b>Progress<br/>
                                                    Chart</b>
                                            </td>
                                            <td class="gray" align="center">
                                                <b>Enrollment</b>
                                            </td>
                                            <td class="gray" align="center">
                                                <b>Core<br/>
                                                    Units</b>
                                            </td>
                                            <td class="gray" align="center">
                                                <b>Elective<br/>
                                                    Units</b>
                                            </td>
                                            <td class="gray" align="center">
                                                <b>ILP<br/>
                                                    Hrs</b>
                                            </td>
                                            <td class="gray" align="center">
                                                <b>Contract<br/>
                                                    Hrs</b>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="gray" align="center">
                                                <b>Goal</b>
                                            </td>
                                            <td class="TableCell" valign="middle" nowrap>
                                                <% 'if (oBudget.PercentEnrolledLocked <> "" and Ucase(session.Contents("strRole")) <> "ADMIN") or oFunc.LockYear then %>
                                                <% if (Ucase(session.Contents("strRole")) <> "ADMIN") or oFunc.LockYear then %>
                                                &nbsp;<% if oBudget.PercentEnrolledLocked <> "" then 
																response.Write oBudget.PercentEnrolledLocked
															else
																response.Write oBudget.PlannedEnrollment
															end if%>%
                                                <input type="hidden" name="intPercent_Enrolled_Fpcs" value="<%=oBudget.PlannedEnrollment%>"
                                                    id="Hidden1"/>
                                                <% else %>
                                                &nbsp;<select name="intPercent_Enrolled_Fpcs" style="font: arial; font-size:10px;"
                                                    onclick="jfGetIndex(this);" onchange="jfChangePercent('<% = intEnroll_Info_ID %>',this.value);"
                                                    id="Select1">
                                                    <%
												Response.Write oFunc.MakeList("25,50,75,100","25%,50%,75%,100%",oBudget.PlannedEnrollment)
                                                    %>
                                                </select>
                                                <% end if %>
                                                <% if ucase(session.Contents("strRole")) = "ADMIN" then %>

                                                <script language="javascript" type="text/javascript">
													function jfLock(){
														var intSel = document.forms["main"].intPercent_Enrolled_Fpcs.selectedIndex;
														var intLevel = document.forms["main"].intPercent_Enrolled_Fpcs.options[intSel].value;
														var URL = '<%=Application.Value("strWebRoot")%>forms/Packet/Packet.asp?';
														URL += "intStudent_ID=<%=intStudent_ID%>&intPercent_Enrolled_Fpcs="+intLevel;
														URL += "&bolLock=true&intEnroll_Info_ID=<%=intEnroll_Info_ID%>";
														window.location.href = URL;
													}
                                                </script>

                                                <input type="button" class="btSmallGray" value="lock" onclick="jfLock();" title="Lock enrollment level so level can never excede locked level."
                                                    name="Button2" id="Button2"/>
                                                <% end if %>
                                                <% if oBudget.PercentEnrolledLocked <> "" then %>
                                                <img src="<%=Application("strImageRoot")%>lock.gif" title="Enrollment level has been locked at <% = oBudget.PercentEnrolledLocked%>%"/>
                                                <% end if %>
                                            </td>
                                            <td class="TableCell" align="center" colspan="2">
                                                <% = oBudget.GoalCoreCredits %>
                                                Core /
                                                <% = oBudget.GoalCoreCredits + oBudget.GoalElectiveCredits %>
                                                Total
                                            </td>
                                            <td class="TableCell" align="center">
                                                <% = oBudget.GoalClassTime %>
                                            </td>
                                            <td class="TableCell" align="center">
                                                <% = oBudget.GoalContractHours %>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="gray" align="center">
                                                <b>Achieved</b>
                                            </td>
                                            <td class="<% 
												if oBudget.ActualEnrollment < oBudget.PlannedEnrollment then 
													response.Write "ErrorCell" 													
												else 
													response.Write "TableCell" 
												end if
													   %>" valign="middle" align="center">
                                                <b>
                                                    <% = oBudget.ActualEnrollment %>%</b>
                                            </td>
                                            <td class="<%
												if oBudget.CoreUnits < oBudget.GoalCoreCredits then 
													response.Write "ErrorCell" 
													packetHelper = packetHelper & "<li>" & round(oBudget.GoalCoreCredits - oBudget.CoreUnits,1) & " more Core Units</li>"
												else 
													response.Write "TableCell" 
												end if
													   %>" align="center">
                                                <b>
                                                    <% = round(oBudget.CoreUnits,1) %></b>
                                            </td>
                                            <td class="<%
												if oBudget.CoreUnits < oBudget.GoalCoreCredits or (oBudget.ElectiveUnits + oBudget.CoreUnits) < (oBudget.GoalCoreCredits + oBudget.GoalElectiveCredits)  then 
													response.Write "ErrorCell" 
													packetHelper = packetHelper  & "<li>" & round((oBudget.GoalCoreCredits+ oBudget.GoalElectiveCredits) - (oBudget.ElectiveUnits + oBudget.CoreUnits),1) & " more Units overall</li>"
												else 
													response.Write "TableCell" 
												end if
												%>" align="center">
                                                <b>
                                                    <% = round(oBudget.ElectiveUnits,1) %></b>
                                            </td>
                                            <td class="<%
												if oBudget.TotalHours < oBudget.GoalClassTime then 
													response.Write "ErrorCell" 
												else 
													response.Write "TableCell" 
												end if
													  %>" align="center">
                                                <b>
                                                    <% = oBudget.TotalHours %></b>
                                            </td>
                                            <td class="<%
												if oBudget.ContractHours < oBudget.GoalContractHours then 
													response.Write "ErrorCell" 
													packetHelper = packetHelper  & "<li>" & oBudget.GoalContractHours - oBudget.ContractHours & " more Contract Hours</li>"
												else 
													response.Write "TableCell" 
												end if
												
												'if packetHelper <> "" then 
												'	packetHelper = left(packetHelper,len(packetHelper)-1)												
												'end if
													  %>" align="center">
                                                <b>
                                                    <% = round(oBudget.ContractHours,1) %></b>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                        </table>
                    </td>
                    <% 
                    oBudget.PopulateFamilyBudgetInfo Application("cnnFPCS"), oBudget.FamilyId, session.Contents("intSchool_Year") 
                    'oBudget.PopulateFamilyBudgetInfo oFunc.FpcsCnn, oBudget.FamilyId, session.Contents("intSchool_Year") 
                    %>
                    <% 
                    oBudget.PopulateIndividualBudgetInfo Application("cnnFPCS"), intStudent_ID,session.Contents("intSchool_Year") 
                    'oBudget.PopulateIndividualBudgetInfo oFunc.FpcsCnn, intStudent_ID,session.Contents("intSchool_Year") 
                    %>
                    <td valign="top">
                    <%'JD 052611 change family budget to refer to individual budget %>
                        <table cellpadding="4" style="width: 400px;">
                            <tr>
                                <td class="TableHeader" align="center" nowrap>
                                    <b>*Individual Elective<br/>
                                        Spending Limits </b>
                                </td>
                                <td class="gray" align="center">
                                    <b>Budget Limit</b>
                                </td>
                                <td class="gray" align="center">
                                    <b>Amount Budgeted</b>
                                </td>
                                <td class="gray" align="center">
                                    <b>Elective<br/>
                                        Balance</b>
                                </td>
                                <td class="gray" align="center">
                                    <b>Global<br/>
                                        Reserve</b>
                                </td>
                            </tr>
                            <tr>
                                <td class="gray">
                                    <b>Individual Budget</b>
                                </td>
                                <td class="TableCell" align="right">
                                    <% 'JD 051911 Budget limit is half of the beginning balance %>
                                    <%'' = formatNumber(oBudget.FamilyBudgetFunding,2) %>
                                    $<%=formatNumber(oBudget.BasePlannedFunding/2,2)%>
                                </td>
                                <td class="TableCell" align="right">
                                    $<% = formatNumber(oBudget.IndividualElectiveBudget,2) %>
                                </td>
                                <td class="TableCell" align="right">
                                    <% 'JD elective budget is budget limit - family elective budget 
                                        availElectiveBudget = oBudget.BasePlannedFunding/2 - oBudget.IndividualElectiveBudget
                                        if oBudget.AvailableElectiveBudget >= 0 then                                        
									        response.Write "$" & formatNumber(availElectiveBudget,2)
									    else
											response.Write "<span class='sverror'>$" & formatNumber(availElectiveBudget,2) & "</span>"	
									    end if
                                    
                                       'if oBudget.AvailableElectiveBudget >= 0 then 
                                       '
									   '    response.Write "$" & formatNumber(oBudget.AvailableElectiveBudget,2)
									   'else
									   '    response.Write "<span class='sverror'>$" & formatNumber(oBudget.AvailableElectiveBudget,2) & "</span>"	
									   'end if
                                    %>
                                </td>
                                <td class="TableCell" align="right">
                                    $<% = formatNumber(dblGlobalReserve,2) %>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="4" class="svplain8">
                                    <b>*</b> Students cannot spend more than 50% of the student's budget
                                    on Music, Art, and/or P.E. classes. 
                                </td>
                            </tr>
                        </table>
                    </td>

                </tr>
            </table>
        </td>
    </tr>
    <tr>
        <td>
            <table id="Table1" cellspacing="2">
                <tr>
                    <td bgcolor="white" colspan="12">
                        <span class="svplain10"><b>Sponsor Teacher:</b>&nbsp;<a href="mailto:<% = oBudget.SponsorEmail %>"><% = oBudget.SponsorName%></a>&nbsp;
                            <input type="button" value="edit" onclick="jfChangeSponsor('<%=intStudent_ID%>');"
                                class="btSmallGray" <% if oFunc.LockYear then response.Write " disabled "%> id="Button3"
                                name="Button3"></span>
                    <br />            
                    <%'JD 052511 move Plan New Course and make it bigger %>
                    
                    <% if not oFunc.LockYear then %>
                        <input type="button" value="Plan New Course" class="NavLinkLarger" onclick="jfOpen('');" name="Button2" id="Button8">
                    <%	end if %>
                    
                    </td>
                    <td class="gray" align="center" title="<% = myBudgetBalance%>">
                        Budget
                    </td>
                    <% 'JD 041111: Take out the Spent Column %>
                    <% 'JD 052511: Return the Spent Column %>
                    <% 'NG 11142011: Take out the Spent Column %>
                    
                    <% if Session.Contents("strRole") = "ADMIN" then 'column Spent shown only for admin user' %>
					<td class="gray" align="center" title="<% = myActualBalance%>">
						Spent
					</td>
					 <% end if %>
                
                </tr>
                <tr>
                    <td rowspan="4" colspan="8" class="svplain8" valign="bottom">
                        <%
							if oBudget.TSTestingSigned < 0 then
								packetHelper = packetHelper & "<li>ASD Testing Agreement must be signed.  " & _
											   "<a href='" & application.Contents("strSSLWebRoot") & "forms/misc/testingAgreement.asp?intStudent_ID=" & intStudent_ID & "'>(Click Here)</a></li>"
							end if
							
							if not oBudget.IsProgressSigned then
								packetHelper = packetHelper & "<li>Progress Report Agreement must be signed.  " & _
											   "<a href='" & application.Contents("strSSLWebRoot") & "forms/misc/ProgressReportAgreement.asp?intStudent_ID=" & intStudent_ID & "'>(Click Here)</a></li>"
							end if
														
							if not oBudget.IsPhilosophyFilled then
								packetHelper = packetHelper & "<li>Must provide an ILP Philosophy. " & _
											   "<a href='" & application.Contents("strSSLWebRoot") & "forms/ilp/ilpPhilosophy.asp?intStudent_ID=" & intStudent_ID & "'>(Click Here)</a></li>"
							end if
								
							if not oBudget.HasSponsorCourse then
								packetHelper = packetHelper & "<li>Packet must include an ASD Sponsor/Oversight class with at least 1 contract hour.</li>"
							end if
													
							If packetHelper <> "" then
                        %>
                        <table class="svplain8" id="Table11">
                            <tr>
                                <td style='width: 140px;' class="TableHeaderRed">
                                    &nbsp;<b>Packet Helper</b>
                                </td>
                                <td>
                                    Items still needed to complete this packet ...
                                    <ul>
                                        <% = packetHelper %>
                                        <li>Course Signatures </li>
                                    </ul>
                                </td>
                            </tr>
                        </table>
                        <%			
							elseif oBudget.AdminPacketSigned then
                        %>
                        <table class="svplain8" id="Table13">
                            <tr>
                                <td style='width: 140px;' class="TableHeaderGreen">
                                    &nbsp;<b>Packet Helper</b>
                                </td>
                                <td>
                                    <b>Congratulations! This packet has been SIGNED and APPROVED.</b>
                                </td>
                            </tr>
                        </table>
                        <%			
							else
                        %>
                        <table class="svplain8" id="Table14">
                            <tr>
                                <td style='width: 140px;' class="TableHeaderBlue">
                                    &nbsp;<b>Packet Helper</b>
                                </td>
                                <td>
                                    Goals from the 'Progress Chart' above have been achieved. AFTER all courses have
                                    been signed by the guardian, sponsor and instructors THEN the packet may be sent
                                    to the principal for review.
                                </td>
                            </tr>
                        </table>
                        <%
							end if
                        %>
                        <br/>
                        <% response.Write oHtml.ToolTip("<a href='#' class='svplain8'>Course Color Key</a>",strColorTable,true,"Course Color Key",false,"tooltip","250px","",false,true) %>
                    </td>
                    <td align="right" class="svplain8" nowrap colspan="3">
                        Beginning Balance:
                    </td>
                    <td bgcolor="white" style="width: 0%;">
                        &nbsp;
                    </td>
                    <td class="gray" align="right">
                        $<%=formatNumber(oBudget.BasePlannedFunding,2)%>
                    </td>
                    <% 'JD 041111: Take out the Spent Column %>
                    <% 'JD 052511: Return the Spent Column %>
                    <% 'NG 11142011: Take out the Spent Column %>
                    <% if Session.Contents("strRole") = "ADMIN" then 'column shown only for admin user' %>
					<td class=gray align=right>
						$<%=formatNumber(oBudget.BaseActualFunding,2)%>
					</td>
					<% end if%>
                </tr>
                <tr>
                    <td align="right" class="svplain8" nowrap colspan="3">
                        Budget Transfer Deposits:
                    </td>
                    <td bgcolor="white" style="width: 0%;">
                        &nbsp;
                    </td>
                    <td class="gray" align="right">
                        <nobr>$<%=formatNumber(oBudget.Deposits,2)%></nobr>
                    </td>
                    <% 'JD 041111: Take out the Spent Column %>
                    <% 'JD 052511: Return the Spent Column %>
                    <% 'NG 11142011: Take out the Spent Column %>
                    <% if Session.Contents("strRole") = "ADMIN" then 'column shown only for admin user' %>
					<td class="gray" align="right">
						<nobr>$<%=formatNumber(oBudget.Deposits,2)%></nobr>
					</td>
					<% end if%>
                </tr>
                <tr>
                    <td align="right" class="svplain8" nowrap colspan="3">
                        Budget Transfer Withdrawals:
                    </td>
                    <td bgcolor="white" style="width: 0%;">
                        &nbsp;
                    </td>
                    <td class="gray" align="right">
                        <nobr>- $<%=formatNumber(oBudget.Withdrawls,2)%></nobr>
                    </td>
                    <% 'JD 041111: Take out the Spent Column %>
                    <% 'JD 052511: Return the Spent Column %>
                    <% 'NG 11142011: Take out the Spent Column %>
                    <% if Session.Contents("strRole") = "ADMIN" then 'column shown only for admin user' %>
					<td class=gray align=right>
						<nobr>- $<%=formatNumber(oBudget.Withdrawls,2)%></nobr>
					</td>
					<% end if%>
					
                </tr>
                <tr>
                    <td align="right" class="svplain8" nowrap colspan="3">
                        Available Remaining Funds:
                    </td>
                    <td bgcolor="white" style="width: 0%;">
                        &nbsp;
                    </td>
                    <td class="gray" align="right">
                        <nobr>$<%=formatNumber(myBudgetBalance-dblGlobalReserve,2)%></nobr>
                    </td>
                    <% 'JD 041111: Take out the Spent Column %>
                    <% 'JD 052511: Return the Spent Column %>
					<% 'NG 11142011: Take out the Spent Column %>
					<% if Session.Contents("strRole") = "ADMIN" then 'column shown only for admin user' %>
					<td class=gray align=right>
						<nobr>$<%=formatNumber(myActualBalance,2)%></nobr>
					</td>
					<% end if%>
					
                </tr>
                <%

'Define Where clause.  This logic determines if we show a budget worksheet
'for all courses for a given student or only for a given course
'JD Add select for the Instructor flat rate
sql = "SELECT     ISF.szCourse_Title  " & _
", POS.txtCourseTitle  " & _
", ISF.intShort_ILP_ID  " & _
", I.szName  " & _
", tblILP.intILP_ID  " & _
", tblILP.bolApproved AS aStatus  " & _
", tblILP.bolSponsor_Approved AS sStatus  " & _
", oi.bolApproved  " & _
", oi.bolSponsor_Approved  " & _
", CASE isNull(tblClasses.intPOS_Subject_ID,1)   " & _
"	when 1 then   " & _
"		case ISF.intPOS_Subject_ID   " & _
"			WHEN 22 THEN 0   " & _
"			ELSE 1 END  " & _
"	ELSE case tblClasses.intPOS_Subject_ID   " & _
"			WHEN 22 THEN 0 ELSE 1 END  " & _
"  END AS isSponsor  " & _
", oi.intQty  " & _
", oi.curUnit_Price  " & _
", oi.curShipping  " & _
", ISF.intCourse_Hrs  " & _
", tblILP.decCourse_Hours  " & _
", oi.intQty * oi.curUnit_Price + oi.curShipping AS total  " & _
", oi.intOrdered_Item_ID  " & _
", tblClasses.intInstructor_ID  " & _
", CASE isNull(tps2.szSubject_Name,'a')   " & _
"	when 'a' then tps.szSubject_Name   " & _
"	else tps2.szSubject_Name   " & _
"  end as szSubject_Name  " & _
", tblClasses.intClass_ID  " & _
", tblClasses.intInstruct_Type_ID  " & _
", tblILP.intContract_Guardian_ID  " & _
", tblClasses.intGuardian_ID  " & _
", tblClasses.intVendor_ID  " & _
", tblClasses.szClass_Name  " & _
", CASE WHEN tblClasses.intInstructor_ID IS NOT NULL THEN ins.szFirst_Name + ' ' + ins.szLast_Name   " & _
"	   WHEN tblClasses.intGuardian_ID IS NOT NULL THEN g.szFirst_Name + ' ' + g.szLast_Name   " & _
"  END AS teacherName  " & _
", tblILP.szAdmin_Comments  " & _
", tblILP.szSponsor_Comments  " & _
", tblILP.bolReady_For_Review  " & _
", tblILP.dtReady_For_Review  " & _
", (SELECT TOP 1 oa2.szValue     " & _
"   FROM tblOrd_Attrib oa2     " & _
"   WHERE oa2.intOrdered_Item_Id = oi.intOrdered_Item_Id   " & _
"   AND (oa2.intItem_Attrib_ID = 9   " & _
"		OR oa2.intItem_Attrib_ID = 5   " & _
"		OR oa2.intItem_Attrib_ID = 6   " & _
"		OR oa2.intItem_Attrib_ID = 22   " & _
"		OR oa2.intItem_Attrib_ID = 33)     " & _
"   ORDER BY oa2.intOrd_Attrib_ID) AS oiDesc  " & _
", oi.bolClosed  " & _
", oi.bolReimburse  " & _
", I.intItem_Group_ID  " & _
", oi.szDeny_Reason  " & _
", tblVendors.szVendor_Name  " & _
", tblVendors.szVendor_Phone  " & _
", tblVendors.szVendor_Fax  " & _
", tblVendors.szVendor_Email  " & _
", tblVendors.szVendor_Website  " & _
", oi.dtCREATE AS oiCreate  " & _
", DM_TEACHER_CLASS_COST.TeacherCostPerStudent  " & _
", DM_TEACHER_RATES.HourlyRateTaxBen  " & _
", DM_TEACHER_CLASS_COST.HoursChargedPerStudent  " & _
", (SELECT top 1 flatRate from tblInstructor_Flat_Rate where intSchool_year = " & session.Contents("intSchool_Year") & ") As InstructorFlatRate  " & _
", DM_TEACHER_CLASS_COST.HoursChargedPerStudent * (SELECT top 1 flatRate from tblInstructor_Flat_Rate where intSchool_year = " & session.Contents("intSchool_Year") & ")  AS FlatRateByHours  " & _
", tblILP.GuardianStatusId  " & _
", tblILP.SponsorStatusId  " & _
", tblILP.InstructorStatusId  " & _
", tblILP.AdminStatusId  " & _
", tblILP.GuardianStatusDate  " & _
", tblILP.SponsorStatusDate  " & _
", tblILP.InstructorStatusDate  " & _
", tblILP.AdminStatusDate  " & _
", tblILP.GuardianComments, tblILP.InstructorComments  " & _
", (select u.szName_First + ' ' + u.szName_Last from tblUsers u where u.szUser_ID = tblILP.GuardianUser) as GuardianUser  " & _
", (select u.szName_First + ' ' + u.szName_Last from tblUsers u where u.szUser_ID = tblILP.SponsorUser) as SponsorUser  " & _
", (select u.szName_First + ' ' + u.szName_Last from tblUsers u where u.szUser_ID = tblILP.InstructorUser) as InstructorUser  " & _
", (select u.szName_First + ' ' + u.szName_Last from tblUsers u where u.szUser_ID = tblILP.AdminUser) as AdminUser  " & _
", (select u.szName_First + ' ' + u.szName_Last from tblUsers u where u.szUser_ID = tblClasses.szUser_Approved) as AdminUser2  " & _
", tblClasses.intInstructor_ID  " & _
", tblClasses.intContract_Status_ID  " & _
", tblClasses.dtApproved  " & _
", tblClasses.szUser_Approved  " & _
", tblILP.bolSponsorAlert  " & _
", tblILP.bolParentAlert  " & _
", CASE isNull(tblClasses.szClass_Name,'a')   " & _
"	WHEN 'a' then   " & _
"		CASE isNull(POS.txtCourseTitle,'a')   " & _
"			WHEN 'a' then ISF.szCourse_Title    " & _
"			else POS.txtCourseTitle end  " & _
"	else tblClasses.szClass_Name   " & _
"   end as ClassLabel  " & _
", tblClasses.szASD_COURSE_ID  " & _
", POS.txtCourseNbr  " & _
", ico.CheckedOutInventoryID    " & _
"FROM tblClasses   " & _
"INNER JOIN tblILP ON tblClasses.intClass_ID = tblILP.intClass_ID   " & _
"LEFT OUTER JOIN trefItems I   " & _
"INNER JOIN tblOrdered_Items oi ON I.intItem_ID = oi.intItem_ID ON tblILP.intILP_ID = oi.intILP_ID   " & _
"RIGHT OUTER JOIN tblILP_SHORT_FORM ISF ON tblILP.intShort_ILP_ID = ISF.intShort_ILP_ID   " & _
"LEFT OUTER JOIN tblProgramOfStudies POS ON ISF.lngPOS_ID = POS.lngPOS_ID   " & _
"INNER JOIN trefPOS_Subjects tps ON tps.intPOS_Subject_ID = ISF.intPOS_Subject_ID   " & _
"LEFT OUTER JOIN trefPOS_Subjects tps2 ON tps2.intPOS_Subject_ID = tblClasses.intPOS_Subject_ID   " & _
"LEFT OUTER JOIN DM_TEACHER_RATES ON tblClasses.intInstructor_ID = DM_TEACHER_RATES.InstructorId AND DM_TEACHER_RATES.StartSchoolYear = " & session.Contents("intSchool_Year") & " " & _
"LEFT OUTER JOIN DM_TEACHER_CLASS_COST ON tblClasses.intClass_ID = DM_TEACHER_CLASS_COST.ClassId   " & _
"LEFT OUTER JOIN tblVendors ON oi.intVendor_ID = tblVendors.intVendor_ID   " & _
"LEFT OUTER JOIN tblINSTRUCTOR INS ON tblClasses.intInstructor_ID = INS.intINSTRUCTOR_ID   " & _
"LEFT OUTER JOIN tblGUARDIAN g ON tblClasses.intGuardian_ID = g.intGUARDIAN_ID    " & _
"LEFT OUTER JOIN INVENTORY_CHECKED_OUT ico ON oi.intOrdered_Item_ID = ico.OrderedItemID    " & _
"WHERE (ISF.intStudent_ID = " & intStudent_ID & ")   " & _
"AND (ISF.intSchool_Year = " & session.Contents("intSchool_Year") & ")     " & _
"ORDER BY isSponsor  " & _
", ClassLabel  " & _
", ISF.intShort_ILP_ID   "
set rsBudget = server.CreateObject("ADODB.RECORDSET")
rsBudget.CursorLocation = 3
rsBudget.Open sql, Application("cnnFPCS")'oFunc.FPCScnn

intPreviousID = 0

if rsBudget.RecordCount < 1 then
                %>
                <tr>
                    <td colspan="20" align="center" class="svplain10">
                        <br/>
                        <b>No courses have been planned yet. To get started click the 'Plan New Course' button
                            above.</b>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
</table>
<%
	response.Write oHtml.ToolTipDivs
	rsBudget.Close
	set rsBudget = nothing
	call oFunc.CloseCN()
	set oFunc = nothing
	set oHtml = nothing
	set oBudget = nothing
	Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")
	response.End
end if
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
				
		' handle Header Color based on status
		if rsBudget("AdminStatusId") = "3" or rsBudget("SponsorStatusId") = "3" or _
			rsBudget("InstructorStatusId") = "3" then
				'Rejected 
				strClassHeader = "TableHeaderBlack"
				CourseHelper = " This course has been rejected.  The Guardian or the Sponsor must delete this course. The funds budgeted by this course will not be released until the course is deleted."			
		elseif  rsBudget("AdminStatusId")  = "2" or rsBudget("SponsorStatusId") = "2" then
			' Needs Work
			strClassHeader = "TableHeaderRed"
			CourseHelper = " This course needs work before it can be signed off on. Please fix any problems and re-sign the contract after any issues have been resolved."
		elseif rsBudget("intILP_ID") & "" = "" then
			strClassHeader = "SubHeader"
			CourseHelper = " This course is in the <b>planned stage</b>. The next step is to implement the plan.  This can be done by selecting 'Implement Plan' under 'Actions' and then click the 'go' button."
		'JD 051911 don't include the admin in determining that the color is green
		'elseif rsBudget("GuardianStatusId") & "" <> "1" or rsBudget("SponsorStatusId") & "" <> "1" or _
		'	(rsBudget("AdminStatusId") & "" <> "1" 
		'	and rsBudget("intContract_Status_Id") & "" <> "5") or _
		'	(rsBudget("intInstructor_ID") & "" <> "" and rsBudget("intInstructor_ID") & "" <> oBudget.SponsorId & "" and rsBudget("InstructorStatusId") & ""  <> "1") then
		'	 strClassHeader = "TableheaderBlue"
		'	 CourseHelper = " This course has not yet been signed by all parties. In order for this course to be complete all parties must sign."
	    elseif rsBudget("GuardianStatusId") & "" <> "1" or rsBudget("SponsorStatusId") & "" <> "1" or _
			(rsBudget("intInstructor_ID") & "" <> "" and rsBudget("intInstructor_ID") & "" <> oBudget.SponsorId & "" and rsBudget("InstructorStatusId") & ""  <> "1") then
			 strClassHeader = "TableheaderBlue"
			 CourseHelper = " This course has not yet been signed by all parties. In order for this course to be complete all parties must sign."

		else
			strClassHeader = "TableHeaderGreen"			
			CourseHelper = "Congratulations! This course has been approved."
		end if 				
		
		if rsBudget("bolSponsorAlert") then
			strClassHeader = "TableHeaderGrape"
		end if 
		
		if rsBudget("bolParentAlert") then
			strClassHeader = "TableHeaderTeal"
		end if 
		
		if rsBudget("AdminStatusId") = 3 or rsBudget("SponsorStatusId") = 3 or _
			rsBudget("InstructorStatusId") = 3 then
			' ILP can be deleted since the course has been rejected
			bolLock = false
		elseif rsBudget("AdminStatusId") = 1 or rsBudget("SponsorStatusId") = 1 _
			or rsBudget("GuardianStatusId") = 1 or rsBudget("InstructorStatusId") = 1  then
			' Prevent ILP from being deleted
			bolLock = true
		else 
			bolLock = false
		end if
		
		if mDivcount > 1 then
			mDivCount = mDivCount + 1
			strBList = strBList & mDivCount & ","
		end if
		
		'if rsBudget("szClass_Name") & "" = "" then 
		'	if rsBudget("txtCourseTitle") & "" <> "" then
		'		myClassName = replace(rsBudget("txtCourseTitle"),"'","\'")
		'	else
		'		myClassName = replace(rsBudget("szCourse_Title"),"'","\'")
		'	end if
		'else 
		'	myClassName = replace(rsBudget("szClass_Name"),"'","\'")
		'end if
		
		'if rsBudget("szClass_Name") & "" <> "" then
		'	myClassName = replace(replace(rsBudget("szClass_Name"),"'","\'"),"""","")
		'end if
		
		if rsBudget("ClassLabel") & "" <> "" then
			myClassName = replace(replace(rsBudget("ClassLabel"),"'","\'"),"""","")
		end if
		
		if rsBudget("teacherName") & "" <> ""  then
			myTeacherName = replace(rsBudget("teacherName"), "'"," ")
		else
			myTeacherName = ""
		end if
			
		'response.Write szClass_Name & "<<<"
%>
<a name="<% = rsBudget("intShort_ILP_ID") %>"></a>
<tr>
    <td colspan="11" style="width: 100%;">
        <table style="width: 100%;" cellpadding='2' cellspacing='1' id="Table7">
            <tr class="<% = strClassHeader %>" <% if mDivcount > 1 then response.Write "id=""div" & mDivCount & """"%>>
                <td align="left" style="width: 50%;">
                    &nbsp;<b>Course Title</b>
                </td>
                <td align='center' style="width: 30%;">
                    <b>Subject</b>
                </td>
                <td align='center' nowrap style="width: 0%;">
                    <b>ILP Hrs</b>
                </td>
                <td align='center' style="width: 0%;">
                    <b>Actions</b>
                </td>
            </tr>
            <tr>
                <td valign="middle" class="<% = strClassHeader%>" style="width: 50%; padding-left: 8px;">
                    <b>
                        <% = ucase(myClassName) %>
                        <% if rsBudget("szASD_COURSE_ID") & "" <> "" and rsBudget("txtCourseNbr") & "" <> "" then
											if rsBudget("szASD_COURSE_ID") & "" <> "" then
												response.Write ": " & rsBudget("szASD_COURSE_ID")
											else	
												response.Write ": " & sBudget("txtCourseNbr")
											end if
										end if
                        %>
                    </b>
                </td>
                <td class="TableCell" valign="top" style="width: 30%;">
                    <% = rsBudget("szSubject_Name") %>
                </td>
                <td class="TableCell" align='center' valign="top" style="width: 0%;">
                    <% = intHours %>
                </td>
                <td class="TableCell" align='right' nowrap valign="top">
                    <select name='action<%= rsBudget("intShort_ILP_ID")%>' style="font: arial; font-size: 10;"
                        id="Select2">
                        <option value="">
                            <option value="">- - - - - - - -
                                <% if rsBudget("intILP_ID") & "" = "" then %>
                                <option value="delete">Delete Plan</option>
                                <option value="edit">Edit Plan</option>
                                <option value="jfAddClass('<% = intStudent_ID%>','<%=rsBudget("intShort_ILP_ID")%>');">
                                    Implement Plan</option>
                                <% else %>
                                <option value="jfContractSchedule('<%=rsBudget("intClass_ID")%>','<%=rsBudget("intInstructor_ID")%>','<%=rsBudget("intInstruct_Type_ID")%>','<%=rsBudget("intContract_Guardian_ID")%>','<%=rsBudget("intGuardian_ID")%>','<%=rsBudget("intVendor_ID")%>');">
                                    <% = strContractSchedule %></option>
                                <% if not bolLock or ucase(session.Contents("strRole")) = "ADMIN" then%>
                                <option value="jfDeleteILP('<% =rsBudget("intILP_ID")%>');">Delete Class</option>
                                <% end if %>
                                <option value="jfViewRoll('<%=rsBudget("intClass_ID")%>');">Enrollment List</option>
                                <option value="jfViewCosts('<%= intStudent_ID %>','<%=rsBudget("intILP_ID")%>','<%=rsBudget("intClass_ID")%>');">
                                    Goods/Services</option>
                                <option value="jfViewILP('<%=rsBudget("intILP_ID")%>','<%=rsBudget("intClass_ID")%>','<%= myClassName%>','<%=rsBudget("intContract_Guardian_ID")%>','<%=rsBudget("intVendor_ID")%>','<%=myTeacherName%>');">
                                    ILP</option>
                                <% end if %>
                    </select>
                    <input type="button" value="go" onclick="jfCallAction('<%= rsBudget("intShort_ILP_ID")%>');"
                        class="btSmallGray" name="Button1" id="Button7">&nbsp;
                </td>
            </tr>
            <% 
							mDivCount = mDivCount + 1
							strBList = strBList & mDivCount & ","
							strSmallList = mDivCount & ","
            %>
            <tr id="div<% = mDivCount%>">
                <td colspan="4">
                    <% if  rsBudget("intILP_ID") & "" <> "" then  %>
                    <table style="width: 100%;" cellspacing="1" cellpadding="0" id="Table10">
                        <tr class="svplain">
                            <td valign="middle" rowspan="2" align="center" class="TableCell" valign="middle"
                                style="background-color: #E9E9FF; color: #404040; width: 130px;">
                                <nobr><b>Course Signatures</b></nobr>
                            </td>
                            <td align="center">
                                Guardian
                            </td>
                            <td align="center">
                                Sponsor<% if rsBudget("intInstructor_ID") & ""  = oBudget.SponsorId & "" then response.Write "/Instructor <input type='hidden' name='IsInstruct" & rsBudget("intILP_ID") & "' value='1'>" %>
                            </td>
                            <% if rsBudget("intInstructor_ID") & "" <> "" and rsBudget("intInstructor_ID") & "" <> oBudget.SponsorId & "" THEN %>
                            <td align="center">
                                Instructor
                            </td>
                            <% end if %>
                            <% 
                            'JD 051911 hide admin signature
                            if (ucase(session.Contents("strRole")) = "ADMIN") then %>
                            <td align="center">
                                Admin
                            </td>
                            <% end if %>
                        </tr>
                        <tr class="svplain">
                            <td align="center">
                                <% 
													' If the logged in user is this students guardian then show drop down
													if ucase(session.Contents("strRole")) = "GUARD" and _
													  rsBudget("GuardianStatusId") & "" <> "1" then												
                                %>
                                <!--<select style="font-size:7pt;" ID="Select3" NAME="status<% = rsBudget("intILP_ID")%>"  onchange="jfILPStatus('<% = rsBudget("intILP_ID")%>');">
													<option></option>
													<option value="1">Sign</option>													
												</select> -->
                                <input type="checkbox" name="status<% = rsBudget("intILP_ID")%>" value="1" onclick="jfILPStatus('<% = rsBudget("intILP_ID")%>','<% = myClassName %>',this);">
                                <% else 
													  if rsBudget("GuardianStatusId") & "" = "" then
                                %>
                                not signed
                                <% else %>
                                <span title="signed on: <% = rsBudget("GuardianStatusDate")%>">
                                    <% = rsBudget("GuardianUser")%></span>
                                <% if oFunc.IsGuardian then %>
                                <input type="hidden" name="status<% = rsBudget("intILP_ID")%>" value="1">
                                <% end if %>
                                <% end if %>
                                <% end if %>
                            </td>
                            <td valign="middle" align="center">
                                <% 
												' If the user logged in is a teacher and this students sponsor then show dropdown
												
												if (Session.Contents("instruct_id") & "" = oBudget.SponsorId & "") and _
													  rsBudget("SponsorStatusId") & "" <> "1" then												
                                %>
                                <select id="Select4" name="status<% = rsBudget("intILP_ID")%>" style="font-size: 7pt;"
                                    onchange="jfILPStatus('<% = rsBudget("intILP_ID")%>','<% = myClassName %>',this);">
                                    <option></option>
                                    <option <% if rsBudget("SponsorStatusId")& "" = "2" then response.Write " selected " %>
                                        value="2">Must Amend</option>
                                    <option <% if rsBudget("SponsorStatusId")& "" = "3" then response.Write " selected " %>
                                        value="3">Rejected</option>
                                    <option value="1">Sign</option>
                                </select>
                                <% else 
													  if rsBudget("SponsorStatusId") & ""  = "1" then %>
                                <span title="signed on: <% = rsBudget("SponsorStatusDate")%>">
                                    <% = rsBudget("SponsorUser")%></span>
                                <% if Session.Contents("instruct_id") & "" = oBudget.SponsorId & "" then %>
                                <input type="hidden" name="status<% = rsBudget("intILP_ID")%>" value="1" id="Hidden10">
                                <% end if %>
                                <% else
															response.Write InterpretStatus(rsBudget("SponsorStatusId"))
														end if%>
                                <% end if %>
                            </td>
                            <% if rsBudget("intInstructor_ID") & "" <> "" and rsBudget("intInstructor_ID") & "" <> oBudget.SponsorId & "" THEN %>
                            <td valign="middle" align="center">
                                <% 
												' If the user logged in is instructor of the class 
												if (session.Contents("instruct_id") & "" = rsBudget("intInstructor_ID") & "" ) and _
													  (rsBudget("InstructorStatusId") & "" <> "1") then												
                                %>
                                <select id="Select5" name="status<% = rsBudget("intILP_ID")%>" style="font-size: 7pt;"
                                    onchange="jfILPStatus('<% = rsBudget("intILP_ID")%>','<% = myClassName %>',this);">
                                    <option></option>
                                    <option <% if rsBudget("InstructorStatusId") & ""  = "3" then response.Write " selected " %>
                                        value="3">Rejected</option>
                                    <option value="1">Sign</option>
                                </select>
                                <% else 
													  if rsBudget("InstructorStatusId") & ""  = "1" then %>
                                <span title="signed on: <% = rsBudget("InstructorStatusDate")%>">
                                    <% = rsBudget("InstructorUser")%></span>
                                <% if Session.Contents("instruct_id") & "" = rsBudget("intInstructor_ID") & "" then %>
                                <input type="hidden" name="status<% = rsBudget("intILP_ID")%>" value="1" id="Hidden11">
                                <% end if %>
                                <% else
															response.Write InterpretStatus(rsBudget("InstructorStatusId"))
														end if%>
                                <% end if %>
                            </td>
                            <% end if %>
                            <td valign="middle" align="center">
                            <%'JD 051911 hide admin signature %>
                            <% if (ucase(session.Contents("strRole")) = "ADMIN") then %>

                                <% if rsBudget("intInstructor_ID") & "" <> "" and rsBudget("intContract_Status_ID") & "" = "5" then 
													' This is ASD course is pre-approved via the principal class approval admin
                                %>
                                <span title="signed on: <% = rsBudget("dtApproved")%>">
                                    <% = rsBudget("AdminUser2")%></span>
                                <% if oFunc.IsAdmin then %>
                                <input type="hidden" name="status<% = rsBudget("intILP_ID")%>" value="1" id="Hidden12">
                                <% end if %>
                                <% elseif rsBudget("AdminStatusId") & "" = "1" then 
													' Signed Schedule
                                %>
                                <span title="signed on: <% = rsBudget("AdminStatusDate")%>">
                                    <% = rsBudget("AdminUser") %></span>
                                <% if oFunc.IsAdmin then %>
                                <input type="hidden" name="status<% = rsBudget("intILP_ID")%>" value="1" id="Hidden13">
                                <% end if %>
                                <% elseif (ucase(session.Contents("strRole")) = "ADMIN") then %>
                                <select name="status<% = rsBudget("intILP_ID")%>" style="font-size: 7pt;" onchange="jfILPStatus('<% = rsBudget("intILP_ID")%>','<% = myClassName %>',this);">
                                    <option></option>
                                    <option <% if rsBudget("AdminStatusId") & "" = "2" then response.Write " selected " %>
                                        value="2">Must Amend</option>
                                    <option <% if rsBudget("AdminStatusId") & "" = "3" then response.Write " selected " %>
                                        value="3">Rejected</option>
                                    <option value="1">Sign</option>
                                </select>
                                <% else 
                                    'JD 051911 hide admin signature
									'response.Write InterpretStatus(rsBudget("AdminStatusId"))												
								end if %>
								
								<% end if %>
                            </td>
                        </tr>
                    </table>
                    <% end if ' ends if ilp_ID <> "" %>
                </td>
            </tr>
            <% 
							mDivCount = mDivCount + 1
							strBList = strBList & mDivCount & ","
							strSmallList = strSmallList & mDivCount & ","
							' We need to know if a Sponsor or Admin has set the course status to Must Amend
							if rsBudget("AdminStatusId") & "" = "2" or rsBudget("SponsorStatusId") & "" = "2" then
            %>
            <input type="hidden" name="MustAmend<% = rsBudget("intILP_ID")%>" value="1">
            <%
							end if 
							
            %>
            <tr id="div<% = mDivCount%>">
                <td colspan="4" style="width: 100%;">
                    <table style="width: 100%;" cellpadding="0" cellspacing="1" id="Table9">
                        <% if rsBudget("intILP_ID") & "" <> "" then %>
                        <tr>
                            <td class="TableCell" valign="middle" align="center" style="background-color: #E9E9FF;
                                color: #404040; width: 130px;" nowrap>
                                <b>ILP Comments</b><br/>
                                <table cellpadding="0" cellspacing="0">
                                    <tr>
                                        <td class="svplain8">
                                            Sponsor Alert&nbsp;
                                        </td>
                                        <td>
                                            <input type="checkbox" name="alert<% = rsBudget("intILP_ID")%>" value="1" <% if rsBudget("bolSponsorAlert") then response.Write " checked " %>
                                                onclick="jfSponsorAlert('<% = rsBudget("intILP_ID")%>');" id="Checkbox2">
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="svplain8">
                                            Parent Alert&nbsp;
                                        </td>
                                        <td>
                                            <input type="checkbox" name="ParentAlert<% = rsBudget("intILP_ID")%>" value="1" <% if rsBudget("bolParentAlert") then response.Write " checked " %>
                                                onclick="jfParentAlert('<% = rsBudget("intILP_ID")%>');" id="Checkbox3">
                                        </td>
                                    </tr>
                                </table>
                            </td>
                            <td valign="top" style="width: 90%;" align="center">
                                <% 
													select case ucase(session.Contents("strRole"))
														case "ADMIN"
															roleComments = rsBudget("szAdmin_Comments")
														case "TEACHER"
															if session.Contents("instruct_id") & "" = oBudget.SponsorId & "" then
																roleComments = rsBudget("szSponsor_Comments")
															elseif session.Contents("instruct_id") & ""  = rsBudget("intInstructor_ID") & "" then
																roleComments = rsBudget("InstructorComments")														
															end if
														case "GUARD"
															roleComments = rsBudget("GuardianComments")	
													end select
													strCommentTable = ""		
													if rsBudget("szAdmin_Comments") & "" <> "" then
														strCommentTable = strCommentTable & "<tr>" & _
																			"<td class='TableCell' style='width:140px;' align=left valign='top'><b>Admin Comments</b></td>" & _
																			"<td class='TableCell'>" & rsBudget("szAdmin_Comments") & "</td></tr>"
													end if
													
													if rsBudget("szSponsor_Comments") & "" <> "" then
														strCommentTable = strCommentTable & "<tr>" & _
																			"<td class='TableCell' style='width:140px;' align=left valign='top'><b>Sponsor Comments</b></td>" & _
																			"<td class='TableCell'>" & rsBudget("szSponsor_Comments") & "</td></tr>"
													end if
													
													if rsBudget("InstructorComments") & "" <> "" then
														strCommentTable = strCommentTable & "<tr>" & _
																			"<td class='TableCell' style='width:140px;' align=left valign='top'><b>Instructor Comments</b></td>" & _
																			"<td class='TableCell'>" & rsBudget("InstructorComments") & "</td></tr>"
													end if
													
													if rsBudget("GuardianComments") & "" <> "" then
														strCommentTable = strCommentTable & "<tr>" & _
																			"<td class='TableCell' style='width:140px;' align=left valign='top'><b>Guardian Comments</b></td>" & _
																			"<td class='TableCell'>" & rsBudget("GuardianComments") & "</td></tr>"
													end if
													
													strCommentTable = strCommentTable & "<tr >" & _
																	"<td  class='TableCell' style='width:140px;background-color:#F0F0F0;' align='center' valign='middle'>&nbsp;<b>Course Helper</b></td>" & _
																	"<td class='TableCell' >" & CourseHelper & "</td></tr>"
																									
													strCommentTable = "<table cellpadding='2' style='width:100%;'>" & strCommentTable & "</table>"																			
                                %>
                                <textarea style='width: 99%;' rows='1' wrap='virtual' name="szComments<% = rsBudget("intILP_ID") %>"
                                    onfocus='this.rows=4;' onblur='this.rows=1;' onkeydown='jfMaxSize(1999,this);'
                                    onchange="jfILPStatus('<% = rsBudget("intILP_ID") %>','',this);" id="Textarea1"><% = roleComments %></textarea>&nbsp;
                                <% = strCommentTable %>
                            </td>
                            <td style="width: 0%;" valign="top">
                                <input type="button" value="Save" onclick="ConfirmSignatures();" class="NavLinkltPurple"
                                    id="Submit2" name="Submit1">
                            </td>
                        </tr>
                        <% else
											response.write "<tr >" & _
														   "<td  class='TableCell' colspan='4' style='width:140px;background-color:#F0F0F0;' align='center' valign='middle'>&nbsp;<b>Course Helper</b></td>" & _
														   "<td class='TableCell' >" & CourseHelper & "</td></tr>"
										 end if %>
                    </table>
                </td>
            </tr>
        </table>
        <nobr>
    </td>
    <td class="ltGray" colspan="3" style="width: 0%;">
        &nbsp;
    </td>
</tr>
<% 
					mDivCount = mDivCount + 1
					strBList = strBList & mDivCount & ","	
					strSmallList = strSmallList & mDivCount & ","
%>
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
    <td class="TableSubHeader" align="center" title="Adjustments are needed to handle over expendatures and to release unused budgeted funds once the budget is closed.">
        Budget Adjust
    </td>
    <td class="TableSubHeader" align="center" title="(Budget Total - Actual Charges) + Budget Adjust">
        Budget Balance
    </td>
    <td class="ltGray" style="width: 0%;">
        &nbsp;
    </td>
    <td class="ltGray">
        &nbsp;
    </td>
    <td class="ltGray">
        &nbsp;
    </td>
</tr>
<%
		'Set alternating row color
		call vbsAlternateColor
		strClass = "TableCell"  ' default class setting
		if len(rsBudget("intInstructor_ID")) > 0 then
			' display teacher cost				
			mDivCount = mDivCount + 1
			strBList = strBList & mDivCount & ","
			strSmallList = strSmallList & mDivCount & ","
			'JD calculate with instructor flat rate
			'JD edit 052611 fix script error
			if Session.Contents("intSchool_Year") < 2013 then
    			dblClassCharge = round(cdbl(rsBudget("TeacherCostPerStudent")),2)
		    	dblClassBudget = round(cdbl(rsBudget("TeacherCostPerStudent")),2)
			else
			    dblClassCharge = round(cdbl(rsBudget("FlatRateByHours")),2)
			    dblClassBudget = round(cdbl(rsBudget("FlatRateByHours")),2)
            end if
%>
<tr id="div<%=mDivCount%>">
    <td class="TableCell">
        &nbsp;
    </td>
    <td class="<% = strClass %>">
        Instruction
    </td>
    <td class="<% = strClass %>" align="center">
        n/a
    </td>
    <td class="<% = strClass %>">
        Instruction by:
        <% = rsBudget("teacherName") %>
    </td>
    <td class="<% = strClass %>" align="center">
        <%= round(rsBudget("HoursChargedPerStudent"),3)%>
    </td>
    <td class="<% = strClass %>" align="right" title="Teachers Hourly Rate">
        <%'JD Change to the Instructor flat rate
						'$ formatNumber(round(rsBudget("HourlyRateTaxBen"),3),3)%>
		<%if Session.Contents("intSchool_Year") < 2013 then %>
		    $<%= formatNumber(round(rsBudget("HourlyRateTaxBen"),3),3)%>
		<%else %>
            $<%= formatNumber(round(rsBudget("InstructorFlatRate"),3),3)%>
        <%end if %>
    </td>
    <td class="<% = strClass %>" align="center">
        n/a
    </td>
    <td class="<% = strClass %>" align="right">
        <%'JD Change to the instructor flat rate
						'$formatNumber(round(rsBudget("TeacherCostPerStudent"),2),2)%>
	    <%if Session.Contents("intSchool_Year") < 2013 then %>
		    $<%= formatNumber(round(rsBudget("TeacherCostPerStudent"),3),3)%>
		<%else %>
            $<%= formatNumber(round(rsBudget("FlatRateByHours"),3),3)%>
        <%end if %>
    </td>
    <td class="<% = strClass %>" align="right">
        <%'JD Change to the instructor flat rate
						'$formatNumber(round(rsBudget("TeacherCostPerStudent"),2),2)%>
	    <%if Session.Contents("intSchool_Year") < 2013 then %>
		    $<%= formatNumber(round(rsBudget("TeacherCostPerStudent"),3),3)%>
		<%else %>
            $<%= formatNumber(round(rsBudget("FlatRateByHours"),3),3)%>
        <%end if %>
    </td>
    <td class="<% = strClass %>" align="right">
        $0.00
    </td>
    <td class="<% = strClass %>" align="right">
        $0.00
    </td>
    <td class="ltGray" style="width: 0%;">
        &nbsp;
    </td>
    <td class="<% = strClass %>" align="right" nowrap>
        <%'JD Change to the instructor flat rate
						'-$formatNumber(round(rsBudget("TeacherCostPerStudent"),2),2)%>
		<%if Session.Contents("intSchool_Year") < 2013 then %>
		    -$<%= formatNumber(round(rsBudget("TeacherCostPerStudent"),3),3)%>
		<%else %>
            -$<%= formatNumber(round(rsBudget("FlatRateByHours"),3),3)%>
        <%end if %>
    </td>
    <% 'JD 041111: Take out the Spent Column %>
    <% 'JD 052511: Return the Spent Column %>
    <% if Session.Contents("strRole") = "ADMIN" then 'column shown only for admin user' %>
					<td class="<% = strClass %>" align="right" nowrap>
						-$<%= formatNumber(round(rsBudget("TeacherCostPerStudent"),2),2)%> 
					</td>
    <% end if%>
					
</tr>
<% end if 			
	end if ' end first time through a given course
			
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
		liInfo = LineItemInfo(rsBudget("intOrdered_Item_ID"),dblBudgetCost, rsBudget("bolClosed"), Application("cnnFPCS"),strClass)
		'liInfo = LineItemInfo(rsBudget("intOrdered_Item_ID"),dblBudgetCost, rsBudget("bolClosed"), oFunc.FPCScnn,strClass)
		bStatus = GetBudgetStatus(rsBudget("intItem_Group_ID"),rsBudget("bolApproved"),liInfo(4),rsBudget("bolReimburse"),rsBudget("CheckedOutInventoryID"))		
		
		dblCharge = formatNumber(liInfo(1),2)
		dblAdjBudget = formatNumber(dblBudgetCost + cdbl(liInfo(2)),2)		
		mDivCount = mDivCount + 1
		strBList = strBList & mDivCount & ","
		strSmallList = strSmallList & mDivCount & ","
		
		if bStatus = "rejc" then
			strClass = "TableCellStrike"
		else
			strClass = "TableCell"
			dblClassCharge = dblClassCharge + cdbl(dblCharge)
			dblClassBudget = dblClassBudget + cdbl(dblAdjBudget)
		end if
		
		if rsBudget("szDeny_Reason") <> "" then
			strReason = "<br/><b>Comment:</b> " & rsBudget("szDeny_Reason")
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
        <a href="javascript:" onclick="jfToggle('<%=liInfo(3)%>','a<%=mLablelCount%>');"
            id="a<%=mLablelCount%>">show</a>
        <% 
							mLablelCount = mLablelCount + 1
						else
							response.Write "&nbsp;"
						end if 
        %>
    </td>
    <td class="<% = strClass %>">
        <% = rsBudget("szName") %>
    </td>
    <td class="<% = strClass %>" align="center">
        <% = bStatus %>
    </td>
    <td class="<% = strClass %>">
        <% response.Write oHtml.ToolTip(strItemType & rsBudget("oiDesc") & strReason, _
							  "<table cellpadding='2'><tr><td class='svplain8' valign='top'><b>Vendor Name:</b></td><td class='svplain8' nowrap>" & rsBudget("szVendor_Name") & "</td></tr>" & _
													 "<tr><td class='svplain8' nowrap><b>Phone Number:</b></td><td class='svplain8' nowrap>" & oFunc.Reformat(rsBudget("szVendor_Phone") , Array("(", 3, ") ", 3, "-", 4)) & "</td></tr>" & _
													 "<tr><td class='svplain8' nowrap><b>Fax Number:</b></td><td class='svplain8' nowrap>" & oFunc.Reformat(rsBudget("szVendor_Fax") , Array("(", 3, ") ", 3, "-", 4))  & "</td></tr>" & _
													 "<tr><td class='svplain8' nowrap><b>Vendor Email:</b></td><td class='svplain8' nowrap>" & rsBudget("szVendor_Email") & "</td></tr>" & _
													 "<tr><td class='svplain8' nowrap><b>Budget Created:</b></td><td class='svplain8' nowrap>" & rsBudget("oiCreate") & "</td></tr></table>", _
													 false, "",false,"tooltip","","",false,false)%>&nbsp;
    </td>
    <td class="<% = strClass %>" align="center" nowrap>
        <% = rsBudget("intQTY") %>
    </td>
    <td class="<% = strClass %>" align="right" nowrap>
        $<% = formatNumber(rsBudget("curUnit_Price"),2) %>
    </td>
    <td class="<% = strClass %>" align="right" nowrap title="Shipping and Handling">
        &nbsp;$<% = dblShipping %>
    </td>
    <td class="<% = strClass %>" align="right" nowrap title="(QTY * Unit Cost) + S/H">
        $<% = dblBudgetCost %>
    </td>
    <td class="<% = strClass %>" align="right" nowrap title="Sum of all line items (charged expeneses) entered by the office for a specific budget.">
        $<% = formatNumber(liInfo(1),2)%>
    </td>
    <td class="<% = strClass %>" align="right" nowrap title="Adjustments are needed to handle over expendatures and to release unused budgeted funds once the budget is closed.">
        $<% = formatNumber(liInfo(2),2) %>
    </td>
    <td class="<% = strClass %>" align="right" nowrap title="(Budget Total - Actual Charges) + Budget Adjust">
        $<% = formatNumber((dblBudgetCost - cdbl(liInfo(1))) + cdbl(liInfo(2)),2)%>
    </td>
    <td bgcolor="white" style="width: 0%;">
        &nbsp;
    </td>
    <td class="<% = strClass %>" align="right" nowrap title="Budget Total - Budget Adjust">
        -$<% = dblAdjBudget %>
    </td>
    <% 'JD 041111: Take out the Spent Column %>
    <% 'JD 052511: Return the Spent Column %>
    <% if Session.Contents("strRole") = "ADMIN" then 'column shown only for admin user' %>
		<td class="<% = strClass %>" align="right" nowrap title="Actual Charges">
		-$<% = dblCharge %>
		</td>
    <% end if%>

</tr>
<% = liInfo(0) %>
<%
	else
	mDivCount = mDivCount + 1
	strBList = strBList & mDivCount & ","
	strSmallList = strSmallList & mDivCount & ","
%>
<tr bgcolor="<% = strColor%>" id="div<%=mDivCount%>">
    <td class="svplain10" colspan="10">
        No Goods or Services have been budgeted for this course.
    </td>
    <td bgcolor="white" style="width: 0%;">
        &nbsp;&nbsp;&nbsp;
    </td>
    <td class="ltGray">
        &nbsp;
    </td>
    <td class="ltGray">
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
    <td class="svplain10" colspan="11" align="right">
        Available Remaining Funds:
    </td>
    <td bgcolor="white" style="width: 0%;">
        &nbsp;&nbsp;&nbsp;
    </td>
   <td class="Gray" align="right">
        <%'formatNumber(dblTargetBalance,2)%>
        $<%=formatNumber(myBudgetBalance,2)%>
        <input type="hidden" name="budgetBalance" value="<%=formatNumber(dblTargetBalance,2)%>"
            id="Hidden8">
    </td>
    <% 'JD 041111: Take out the Spent Column %>
    <% 'JD 052511: Return the Spent Column %>
    
    <% if Session.Contents("strRole") = "ADMIN" then 'column shown only for admin user' %>
				<td class="Gray" align=right>
						<%'formatNumber(dblActualBalance,2)%>
						$<%=formatNumber(myActualBalance,2)%>
					</td>	
    <% end if%>
    			    
</tr>

<script language="javascript">
    function jfToggleBudget(pMe) {
        jfToggle('<%=strBList%>', '');

        if (pMe.value == "Show Detail") {
            pMe.value = "Hide Detail";
        } else {
            pMe.value = "Show Detail";
        }
    }	
</script>

<%
end if

set rsBudget = nothing					
%>
</table> </td> </tr> </table>
</form>
<%
response.Write oHtml.ToolTipDivs
call oFunc.CloseCN()
set oFunc = nothing
set oHtml = nothing
set oBudget = nothing
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
	mDivCount = mDivCount + 1
	strDivList = strDivList & mDivCount & ","
	strSmallList = strSmallList & mDivCount & ","
%>
<tr class="svplain10" bgcolor="<% = strColor%>">
    <td colspan="11" align="right">
        <table style="width: 100%;">
            <tr>
                <td class="svplain8" align="left">
                    <a href="javascript:" onclick="jfToggle('<%=strSmallList%>','');">hide/show</a>
                </td>
                <td align="right" class="svplain10">
                    <b>Course Totals:</b>
                </td>
            </tr>
        </table>
    </td>
    <td bgcolor="white" style="width: 0%;">
        &nbsp;&nbsp;&nbsp;
    </td>
    <td class="gray" align="right">
        <nobr>
						<% if instr(1,dblBudgetCost,"-") > 0 then
								response.Write "+ $" & formatNumber(replace(dblClassBudget,"-",""),2)
						   else
								response.Write "- $" & formatNumber(dblClassBudget,2)
						   end if						
						%></nobr>
    </td>
    <% 'JD 041111: Take out the Spent Column %>
    <% 'JD 052511: Return the Spent Column %>
    <% if Session.Contents("strRole") = "ADMIN" then 'column shown only for admin user' %>
								<td class="gray" align="right">
						<nobr>
						<% if instr(1,dblActualCost,"-") > 0 then
								response.Write "+ $" & formatNumber(replace(dblClassCharge,"-",""),2)
						   else
								response.Write "- $" & formatNumber(dblClassCharge,2)
						   end if						
						%></nobr>
					</td>		
    <% end if%>

</tr>
<%
		
%>
<tr bgcolor="white" id="div<% = mDivCount%>">
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
		mDivCount = mDivCount + 1
		strDivList = strDivList & mDivCount & ","
		tCharged = tCharged + formatNumber(rs("Total"),2)	
		if rs("szCheck_Number") & "" <> "" then
			szCheck_Number = "Check #: " & rs("szCheck_Number")
			if rs("szLine_Item_desc") & "" <> "" then
				szCheck_Number = "<br/>" & szCheck_Number
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
	
	if tCharged <> 0 then
		bolLineItem = true
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
	rs.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
	
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
				call vbsIlpStatus(arList(i))
			end if
		next
	end if
end sub

sub vbsUpdateAlerts(pAlertList,pType,pFieldName)
	' This sub updates the Sponsor Alert Status
	arList = split(pAlertList,",")	
	if isArray(arList) then
		dim update, myVal
		for i = 0 to ubound(arList)		
			if arList(i) <> "" then
				if request(pFieldName & arList(i)) & "" <> "" then
					myVal = 1
				else
					myVal = 0
				end if
				
				update = "update tblILP set " & pType & " = " & myVal & "," & _
					     "szUser_Modify = '" & session.Contents("strUserId") & "', " & _
					     "dtModify = CURRENT_TIMESTAMP " & _
					     "where intILP_ID = " & arList(i) & _
					     " and intStudent_ID = " & request("intStudent_ID")
				oFunc.ExecuteCN(update)
			end if
		next
	end if
end sub

sub vbsIlpStatus(pIlpId)
	' update ILP Status and comments based on user Role
	dim update, myStatus
	
	if request("status" & pIlpId) & "" = "" then
		myStatus = " NULL "
	else
		myStatus = request("status" & pIlpId)
	end if
	update = "update tblILP set "
	if ucase(session.Contents("strRole")) = "ADMIN" then
		update = update & " AdminStatusId = " & myStatus & ", " & _
						  " AdminStatusDate = CURRENT_TIMESTAMP, " & _
						  " AdminUser = '" & session.Contents("strUserId") & "', " & _
						  " szAdmin_Comments = '" & oFunc.EscapeTick(request("szComments" & pIlpId)) & "' " 
		if request("status" & pIlpId) & "" = "2" then
			update = update & " ,InstructorStatusId = null,SponsorStatusId = null, GuardianStatusId = null " 					
		end if
	elseif request("IsInstruct" & pIlpId) & "" = "1" and Session.Contents("instruct_id") & "" = oBudget.SponsorId & "" then
		' User is both the Instructor and the Sponsor Teacher
		' Session.Contents("instruct_id") is only defined at log in if the user us a teacher
		' request("IsInstruct" & pIlpId) is only defined if the sponsor teacher is also the instructor for the class that relates to pIlpId
		update = update & " SponsorStatusId = " & myStatus & ", " & _
						  " SponsorStatusDate = CURRENT_TIMESTAMP, " & _
						  " SponsorUser = '" & session.Contents("strUserId") & "', " & _
						  " szSponsor_Comments = '" & oFunc.EscapeTick(request("szComments" & pIlpId)) & "' " 
		
		update = update & ", InstructorStatusId = " & myStatus  & ", " & _
						  " InstructorStatusDate = CURRENT_TIMESTAMP, " & _
						  " InstructorUser = '" & session.Contents("strUserId") & "' "  		
						  				  
		if request("status" & pIlpId) & "" = "2" then
			update = update & " , GuardianStatusId = null " 
		end if
	elseif ucase(session.Contents("strRole")) = "TEACHER" and Session.Contents("instruct_id") & "" = oBudget.SponsorId & "" then
		' Sponsor Teacher
		update = update & " SponsorStatusId = " & myStatus & ", " & _
						  " SponsorStatusDate = CURRENT_TIMESTAMP, " & _
						  " SponsorUser = '" & session.Contents("strUserId") & "', " & _
						  " szSponsor_Comments = '" & oFunc.EscapeTick(request("szComments" & pIlpId)) & "' " 
		'if myStatus & "" = "1" and request("IsInstruct" & pIlpId) & "" = "" then
		'	update = update & ", InstructorStatusId = " & myStatus  & ", " & _
		'				  " InstructorStatusDate = CURRENT_TIMESTAMP, " & _
		'				  " InstructorUser = '" & session.Contents("strUserId") & "' " 
		'end if
		
		if myStatus & "" = "2" then
			update = update & " , GuardianStatusId = null " 
		elseif request("MustAmend" & pIlpId) & "" <> "" and request("status" & pIlpId) & "" = "1" then
			update = update & " , AdminStatusId = null "			
		end if
	'elseif ucase(session.Contents("strRole")) = "TEACHER" then
		'update = update & " InstructorStatusId = " & myStatus  & ", " & _
		'				  " InstructorStatusDate = CURRENT_TIMESTAMP, " & _
		'				  " InstructorUser = '" & session.Contents("strUserId") & "', " & _
		'				  " InstructorComments = '" & oFunc.EscapeTick(request("szComments" & pIlpId)) & "' " 		
	elseif ucase(session.Contents("strRole")) = "GUARD" then
		if myStatus <> " NULL " then
			' we are signing
			update = update & " GuardianStatusId = " & myStatus & ", " & _
							" GuardianStatusDate = CURRENT_TIMESTAMP, " & _
							" GuardianUser = '" & session.Contents("strUserId") & "', " & _
							" GuardianComments = '" & oFunc.EscapeTick(request("szComments" & pIlpId)) & "' " 
		else
			' just saving comments
			update = update & " GuardianComments = '" & oFunc.EscapeTick(request("szComments" & pIlpId)) & "' " 
		end if
		
		if request("MustAmend" & pIlpId) & "" <> "" and request("status" & pIlpId) & "" = "1" then
			update = update & " , AdminStatusId = null, SponsorStatusId = null "
		end if
	else
		exit sub
	end if 
	
	update = update & " where intILP_ID = " &  pIlpId
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

function GetBudgetStatus(pItemGroup,pBappr,pBolLineItems,pIsReimburse,pCheckedOut)	

	if pBappr & "" = "" and pBolLineItems = false then
		GetBudgetStatus = "pend"
	elseif pBappr = false then
		GetBudgetStatus = "rejc"
	elseif pBolLineItems = true and pIsReimburse = false and pItemGroup = 2 then
		if pCheckedOut & "" = "" then 
			GetBudgetStatus = "<font color='red'><b>pick up</b></font>"
		else
			GetBudgetStatus = "<font color='green'>delivered</font>"
		end if
	elseif pBolLineItems = true and pIsReimburse = false and pItemGroup = 1 then
		GetBudgetStatus = "pymt made"
	elseif pBappr = true and pBolLineItems = false and pIsReimburse = false and pItemGroup = 2 then
		GetBudgetStatus = "ordered"
	elseif pBappr = true and pBolLineItems = false and pIsReimburse = false and pItemGroup = 1 then
		GetBudgetStatus = "vend appr"
	elseif pBappr = true and pBolLineItems = true and pIsReimburse = true then
		'GetBudgetStatus = "check cut"
		GetBudgetStatus ="processing"
	else
		GetBudgetStatus = "pend"
	end if
end function

function InterpretStatus(pStatusId)
	' simply takes the statusId and gives us the corresponding label so 
	' we don't have to make 4 more sub queries to get the label for each role
	' This of course stinks if the labels need to be changed. 
	select case pStatusId
		case "1"
			InterpretStatus = "Signed"
		case "2"
			InterpretStatus = "Must Amend"
		case "3"
			InterpretStatus = "Rejected"
		case else
			InterpretStatus = "not signed"
	end select
end function
%>

<script language="javascript">
    RestoreScroll('<% = intStudent_id %>Packet');
</script>

