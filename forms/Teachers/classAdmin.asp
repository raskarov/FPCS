<%@ Language=VBScript %>
<%
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Dimention Variables, Make db Connection, print HTML header.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim intClass_Id
dim intInstructor_ID
dim sqlInstructor
dim sqlClass
dim sqlMaterials
dim intCount
dim strMaterials
dim strClassTitle
dim strInstructMessage
dim intStudent_id
dim strAddSQL				'Dynamic peice of sql defined depending on instructor,guardian or vendor
dim curInstructionRate		'Holds the hourly rate of instruction including taxes and benefits
dim strStudentName			'Contains the current students name.
dim strCalcType				'Determines if we run a javascript function that figures based on Instructor
							'fields or Vendor
dim strDisabled				'This string is used in form elements to disable them when we are adding ILP's
dim strFamilyList			'Contains list of families that this class is restricted to 	
dim strFamilyValues			'This is used to keep track of whether the families pulldown is populated.
							'If in edit mode it was populated and the admin decided to make it open
							'to everyone we needed some way of nowing that all family restrictions 
							'for this class dhould be deleted and not replaced with others.		
dim strFormType

dim oFunc	'wsc object
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

'set default dates
dim regMonth
dim regDay
dim regYear
dim monthStart
dim dayStart
dim yearStart
dim dayEnd
dim monthEnd
dim yearEnd
dim strASDWrite

regMonth = "9"
regDay = "1"
regYear = Session.Contents("intSchool_Year") - 1
	
monthStart = "7"
dayStart = "1"
yearStart = Session.Contents("intSchool_Year") - 1

monthEnd = 6
dayEnd = 1
yearEnd = Session.Contents("intSchool_Year") 

' This session variable needs to be cleared each time this script runs to prevent problems
' of mainting improper state when used with viewClass.asp.  This variable will be defined
' in classInsert.asp and is designed to be used in ilpMain.asp to define the pos subject
' id for the generic ILP
Session.Contents("intPOS_Subject_ID_from_class") = ""
' ####NOTE#### Session variables for intInstructor_ID are used in ilpMain.asp
' Define instructor information
if Request("intInstructor_ID") <> "" then	
	intInstructor_ID = request("intInstructor_ID")
	Session.Value("intInstructor_ID") = intInstructor_ID
	
	set oTeacher = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/TeacherInfo.wsc"))
	'oTeacher.PopulateObject oFunc.FPCScnn, intInstructor_ID, session.Contents("intSchool_Year")
	oTeacher.PopulateObject Application("cnnFPCS"), intInstructor_ID, session.Contents("intSchool_Year")
	
	strTeacherName = oTeacher.FirstName & " " & oTeacher.LastName	
	
	Session.Value("strTeacherName") = strTeacherName
	strClassTitle = "Add a Contract "
	strFormType = "Contract"
	strCalcType = "jfAddHRS();"
else
	' Ensures Session.Value("intInstructor_ID") is erased if we are not dealing with an instructor
	strClassTitle = "Course Schedule"
	strFormType = "Schedule"
	Session.Value("intInstructor_ID") = ""
	Session.Value("strTeacherName") = ""
	intInstructor_ID = ""
	if session.Contents("intPOS_Subject_ID") & "" <> "" then
		' Get the subject name (This is possible from the info provided by
		' ilp1.asp)
		set rsSubject = server.CreateObject("ADODB.RECORDSET")
		rsSubject.CursorLocation = 3
		sql = "Select szSubject_Name " & _
			  "from trefPOS_Subjects " & _
			  "where intPOS_Subject_ID = " & session.Contents("intPOS_Subject_ID")
		rsSubject.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
		szSubject_Name = rsSubject("szSubject_Name")
		intPOS_Subject_ID = session.Contents("intPOS_Subject_ID")
		rsSubject.Close
		set rsSubject = nothing
	end if
end if 

' Initualize variables from their different sources.  We do not want the form elements to be enabled
' when we are dealing with students. Our rule is that classes can only be edited by the instructor or
' fpcs admin. So a parent can not edit a class that effects other students. 
if request("strDisabled") <> "" then
	'Coming from viewClasses.asp with a defined student_id. 
	intClass_Id = Request("intClass_Id")
	if intInstructor_ID & "" <> "" then
		strDisabled = " disabled "
	else
		strGuardDisabled = " disabled " 		
	end if 
elseif Request.QueryString("intClass_Id") <> "" then
	'Coming from veiw/delete class (viewClasses.asp from the teachers version)
	intClass_Id = Request.QueryString("intClass_Id")
elseif Request.Form("intClass_Id") <> ""  then
	'Coming from add an ILP (ILP1.asp) 
	intClass_Id = Request.Form("intClass_Id")
	if intInstructor_ID & "" <> "" then	
		strDisabled = " disabled "
	else
		strGuardDisabled = " disabled " 		
	end if
end if


'  This session variable tells ilpMain.asp how to work in relation to this page.
'  (it will use a little different logic as opposed to how ilpMain would work if it was
'   called from viewClasses.asp). We only alert ilpMain that we are coming from this script 
' if we are creating a contract from scratch or if we are adding a contract to a student.  
' Otherwise we are in edit mode from viewClasses page.
'Note: bolInWindow is from viewClasses.asp and we use it to tell us what header to include.
' We use it here as well because if it is defined we are not creating a class we are viewing it
' and it allows us to distinguish in the second part of our if clause if we are
' adding an existing contract to a student's course list or if we are just viewing an already 
' added contract.

if intClass_Id = "" or (intClass_Id <> "" and Session.Value("intStudent_ID") <> "" and Request.QueryString("bolInWindow") = "") _
	or (intClass_ID <> "" and Request.QueryString("strThisIsACopy") <> "") then
	session.Value("blnFromClassAdmin") = true
else
	session.Value("blnFromClassAdmin") = false
end if

intCount = 0

' Session.Value("intStudent_ID") may not always be destroyed when coming direclty from root.
' (This script is executed coming from ilp1.asp and default.asp)
' If it's coming from default AND from a teacher ADD A CLASS request 
' request("bolFromTeacher") will be defined and we can not have intStudent_ID populated

if Request.QueryString("bolFromTeacher") = "" then
	intStudent_id = Session.Value("intStudent_ID")
else
	Session.Value("intStudent_ID") = ""
	Session.Value("studentFirstName") = ""
end if 

if Session.Value("strStudentName") <> "" then
	strStudentName = "<B>Current Student:</b> " & Session.Value("strStudentName") 
end if 

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' This select block sets what parts of the html form we show 
'' depending on the Instruction type.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
select case request("intInstruct_Type_ID")
	case "4" 'Contract ASD Teacher
		strInstructMessage = "<B> Instructed by:</b> " & strTeacherName
		' JD: EDIT get the flat rate, instead:
		if session.Contents("intSchool_Year") < 2012 then
		   curInstructionRate = oTeacher.FullHourlyRate
		else
		sql4 ="select intFlat_Inst_Id, flatRate from tblInstructor_Flat_Rate where intSchool_year = " & session.Contents("intSchool_Year")
        set rs4 = server.CreateObject("ADODB.RECORDSET")
        rs4.CursorLocation = 3
        rs4.Open sql4, Application("cnnFPCS")'oFunc.FPCScnn
        curinstructionrate = formatNumber(rs4("flatRate"), 2)
        rs4.Close()
        end if
		' This next bit of code creates the html needed to allow a user to
		' copy an existing contract.
		if intClass_ID = "" then
			dim sqlContracts
			dim strParams
			dim strNames
			dim strIDS
			sqlContracts = "select c.intClass_ID,gi.intILP_ID,c.szClass_Name " & _
				  "from tblInstructor i,tblClasses c left outer join tblILP_Generic gi " & _
				  " ON c.intClass_ID = gi.intClass_ID " & _
				  "where i.intInstructor_ID = c.intInstructor_ID and " & _
			      "i.intInstructor_ID =" &  intInstructor_ID & _ 
			      " order by c.szClass_Name "
			 
			set rsIDS = server.CreateObject("ADODB.RECORDSET")
			rsIDS.CursorLocation = 3
			
			rsIDS.Open sqlContracts, Application("cnnFPCS")'oFunc.FPCScnn									

			strSelectContract = "<table><tr><td class=gray>&nbsp;Copy an Existing Contract:</td>" & _
								"<td><select name=intContract_ID onChange='jfGetContract(this);'>" & _
								"<option>Select a Contract" 
			if rsIDS.recordcount > 0 then
				do while not rsIDS.eof
					strSelectContract = strSelectContract & "<option value=""" & rsIDS("intClass_ID") & "|" & rsIDS("intILP_ID") & """>" & rsIDS("szClass_Name") & chr(13)
					rsIDS.moveNext
				loop			
			end if 
			rsIDS.Close
			set rsIDS = nothing
			strSelectContract = strSelectContract & "</select></td></tr></table>"
			strParams = "?intInstruct_Type_ID=" & Request.QueryString("intInstruct_Type_ID") & _
						"&bolFromTeacher=True&intInstructor_ID=" & 	intInstructor_ID & _
						"&strThisIsACopy=true"	
			set oTeacher = nothing					
		end if								
end select

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' This next section will fill the form in with class info 
'' if we have a valid class id passed to this script.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if intClass_Id <> "" then
	'sqlClass gets most of the class information
	set rsClass = server.CreateObject("ADODB.RECORDSET")	
	rsClass.CursorLocation = 3 
	if intInstructor_ID <> "" then
		strAddSQL = "i.intInstructor_ID,i.curPay_Rate , pos.szSubject_Name " & _
					"FROM tblClasses c INNER JOIN " & _
                    "tblINSTRUCTOR i ON c.intInstructor_ID = i.intINSTRUCTOR_ID INNER JOIN " & _
                    "trefPOS_Subjects pos ON c.intPOS_Subject_ID = pos.intPOS_Subject_ID " & _
					"where c.intClass_ID = " & intClass_Id 
	elseif Request.Form("intGuardian_id") <> "" or Request.QueryString("intGuardian_id") <> "" then
		strAddSQL = "g.intGuardian_ID, pos.szSubject_Name, c.intDuration_ID, c.intSession_Minutes " & _
					"FROM tblClasses c INNER JOIN " & _
                    "tblGUARDIAN g ON c.intGuardian_ID = g.intGUARDIAN_ID INNER JOIN " & _
                    "trefPOS_Subjects pos ON c.intPOS_Subject_ID = pos.intPOS_Subject_ID " & _
					"where c.intClass_ID = " & intClass_Id 
	end if 
	
	sqlClass = "select c.intPOS_Subject_ID,c.intInstructor_ID,c.szClass_Name,c.szASD_Course_ID," & _
			   "c.szLocation,c.dtReg_Deadline,c.intMin_Students,c.intMax_Students," & _
			   "c.sGrade_Level,c.sGrade_Level2,c.dtClass_Start,c.dtClass_End,c.szStart_Time,c.szEnd_Time," & _
			   "c.szSchedule_Comments,c.decHours_Student,c.decHours_Planning,c.szDays_Meet_On, " & _
			   "c.decOriginal_Student_Hrs, c.intContract_Status_ID, c.decOriginal_Planning_hrs, dtHrs_Last_Updated, " & _
			   	strAddSQL 
			   	
	rsClass.Open sqlClass, Application("cnnFPCS")'oFunc.FPCScnn		

	'This for loop dimentions and defines all the columns we selected in sqlClass
	'and we use the variables created here to populate the form.
	for each item in rsClass.Fields
		execute("dim " & rsClass.Fields(intCount).Name)
		execute(rsClass.Fields(intCount).Name & " = item")		
		intCount = intCount + 1
	next 
	if strParams = "" then 
		'strParams is only defined if we are coping a contract. If this is a copy we do not want to 
		'define the following
		Session.Value("szClass_Name") = szClass_Name
		Session.Value("intClass_Id") = intClass_Id
	end if

	rsClass.Close
	set rsClass = nothing
	
	' See if this class is limited to select familes and if so get them in a comma seperated list
	' so we can auto populate them on the form
	dim sqlRestricted
	sqlRestricted = "select a.intFamily_ID, Name = " & _
					"CASE " & _
					"WHEN f.szDesc is null then f.szFamily_Name " & _
					"WHEN f.szDesc is not null then f.szFamily_Name + ', ' + f.szDesc " & _
					"END " & _
				    "from tascClass_Family a, tblFamily f " & _
					"where a.intClass_ID = " & intClass_ID & _
					" and a.intFamily_ID = f.intFamily_ID " & _
					" order by f.szFamily_Name "  
	strRestrictedFamList = oFunc.MakeListSQL(sqlRestricted,"intFamily_ID","Name","")				
	
	'This next section breaks up date information that is stored in single columns in the 
	'database because they are displayed as individual drop downs in the HTML form.
	'So we need the individual values to populate the drop downs.
	
	dim hourStart
	dim minuteStart
	dim amPmStart
	dim hourEnd
	dim minuteEnd
	dim amPmEnd
	 
	regMonth = datePart("m",dtReg_Deadline)
	regDay = datePart("d",dtReg_Deadline)
	regYear = datePart("yyyy",dtReg_Deadline)
	 
	monthStart = datePart("m",dtClass_Start)
	dayStart = datePart("d",dtClass_Start)
	yearStart = datePart("yyyy",dtClass_Start)

	monthEnd = datePart("m",dtClass_End)
	dayEnd = datePart("d",dtClass_End)
	yearEnd = datePart("yyyy",dtClass_End)
	 
	if szStart_Time <> "" then
		arStartTime = split(szStart_Time,":")
		hourStart = arStartTime(0)
		arStartTime2 = split(arStartTime(1)," ")
		minuteStart = arStartTime2(0)
		amPmStart = arStartTime2(1)
	end if 
	
	if szEnd_Time <> "" then
		arEndTime = split(szEnd_Time,":")
		hourEnd = arEndTime(0)
		arEndTime2 = split(arEndTime(1)," ")
		minuteEnd = arEndTime2(0)
		amPmEnd = arEndTime2(1)
	end if 
	
	strClassTitle = "Scheduled Class"	
'elseif session.Contents("intFamily_ID") <> "" then
	' This will make the default restricted family the family
	' of the guardian that is loged in.
	' session.Contents("intFamily_ID") is only defined when a guardian logs 
	' into the system.
'	sqlRestricted = "select a.intFamily_ID, Name = " & _
'					"CASE " & _
'					"WHEN f.szDesc is null then f.szFamily_Name " & _
'					"WHEN f.szDesc is not null then f.szFamily_Name + ', ' + f.szDesc " & _
'					"END " & _
'				    "from tascClass_Family a, tblFamily f " & _
'					"where f.intFamily_ID = " & session.Contents("intFamily_ID")
'	strRestrictedFamList = oFunc.MakeListSQL(sqlRestricted,"intFamily_ID","Name","")			
end if 	
Session.Value("strTitle") = strFormType
Session.Value("strLastUpdate") = "22 Feb 2002"

'This next section prints the side menu from our header ONLY if this page
'will NOT be in a spawned window. (Due to window size is smaller in spawned window
' and our work flow will close the spawned window at task completion.)
'if Request.QueryString("bolInWindow") <> "" then
'	session.Value("simpleTitle") = strFormType
	Server.Execute(Application.Value("strWebRoot") & "Includes/simpleHeader.asp")	
'else
'	Server.Execute(Application.Value("strWebRoot") & "Includes/header.asp")
'end if 

if intInstructor_ID <> "" and intContract_Status_ID & "" = "5" and ucase(session.Contents("strRole")) <> "ADMIN" then
	bolLock = true	
	strDisabled = " disabled "
end if

%>
<script language=javascript>
	function jfCheckCG(){
		<% if request("intGuardian_ID") <> "" then %>		
		location.href="../ILP/ILPMain.asp?intGuardian_ID=<%=request("intGuardian_ID")%>";
		<% else %>
		if (document.main.intGuardian_ID2.value == "")	{
			alert("You must select a Guardian that will be on the 'Parent Teacher Contract' for this course.");
		}
		else{
			// Check to see if adding teacher cost will blow our budget
			if (!jfCheckFunds()) {
				return false;								
			}
			var maxCost = document.main.maxCost.value;
			location.href="../ILP/ILPMain.asp?maxCost="+maxCost+"&intContract_Guardian_ID=" + document.main.intGuardian_ID2.value;
		}
		<% end if%>
	}

	function jfSubmit(objForm) {
		if (jfValidate(objForm) == true) {	
			objForm.submit();
		}
	}

	function jfValidate(objForm) {
	//added bkm 26-Apr-2002
	//Ensure all approriate fields have been filled out
		var strErrMsg				= '';
		var szClass_Name			= objForm.szClass_Name.value;
		var intPOS_Subject_ID		= objForm.intPOS_Subject_ID.value;
		var intMin_Students			= objForm.intMin_Students.value;
		var intMax_Students			= objForm.intMax_Students.value;			
		var szDays_Meet_On			= objForm.szDays_Meet_On;	
		var szSchedule_Comments		= objForm.szSchedule_Comments.value;
		var strItems				= "";
		var isChecked				= false;
		
		
		for (i=0; i< szDays_Meet_On.length; i++) {			
			if (szDays_Meet_On[i].checked == true) {
				isChecked = true;
				break;
			}
		}
		
		for (i=0; i< objForm.selFamilies.length; i++) {
			strItems = strItems + objForm.selFamilies.options[i].value + ",";
		}
		objForm.intFamily_ID.value = strItems.substr(0, strItems.length - 1); 
		
		<% if intInstructor_id <> "" then %>			
			var sGrade_Level		= objForm.sGrade_Level.value;
			var intStartMonth		= objForm.monthStart.value;
			var intStartDay		= objForm.dayStart.value;
			var intStartYear		= objForm.yearStart.value;
			var intEndMonth		= objForm.monthEnd.value;
			var intEndDay			= objForm.dayEnd.value;
			var intEndYear			= objForm.yearEnd.value;
			var intRegMonth		= objForm.regMonth.value;
			var intRegDay			= objForm.regDay.value;
			var intRegYear			= objForm.regYear.value;
			var decHours_Student = objForm.decHours_Student.value;
			var decHours_Planning = objForm.decHours_Planning.value;
			if(sGrade_Level.length == 0) {strErrMsg += 'Grade Level\n';}
			if(intStartMonth.length == 0) {strErrMsg += 'Start Month\n';}
			if(intStartDay.length == 0) {strErrMsg += 'Start Day\n';}
			if(intStartYear.length == 0) {strErrMsg += 'Start Year\n';}
			if(intEndMonth.length == 0) {strErrMsg += 'End Month\n';}
			if(intEndDay.length == 0) {strErrMsg += 'End Day\n';}
			if(intEndYear.length == 0) {strErrMsg += 'End Year\n';}
			if(intRegMonth.length == 0) {strErrMsg += 'Registration Month\n';}
			if(intRegDay.length == 0) {strErrMsg += 'Registration Day\n';}
			if(intRegYear.length == 0) {strErrMsg += 'Registration Year\n';}
			if(intPOS_Subject_ID.length == 0) {strErrMsg += 'Subject\n';}
			
			if (parseFloat(decHours_Planning) / parseFloat(decHours_Student) > .50){
				strErrMsg += 'Hours spent on planning can not excede 1/2 the hours spent with student.\n';
			}
			
			if(decHours_Student.length == 0 && decHours_Planning.length == 0) {
				strErrMsg += 'Either \'Teacher Hours With Student\' or \'Hours For Teacher Planning\' must be filled in\n';
			}else{
				if(decHours_Student.length != 0){
					if (isNaN(decHours_Student) == true) {strErrMsg += 'Teacher Hours With Student must be a number\n';}
				}
				if(decHours_Planning.length != 0){
					if (isNaN(decHours_Planning) == true) {strErrMsg += 'Hours For Teacher Planning must be a number\n';}
				}
			}
			
			if (strErrMsg.length == 0 ) {
				//if all of the required fields are populated then we test the values
				//in some of them.  Additionaly, if some UnRequired fields are populated,
				//we test their values as well
				if (checkDate(objForm.regYear, objForm.regMonth, objForm.regDay, "Registration Date") == false) {return false;}
				if (checkDate(objForm.yearStart, objForm.monthStart, objForm.dayStart, "Start Date") == false) {return false;}
				var strDate = objForm.monthStart.value+'/'+objForm.dayStart.value+'/'+objForm.yearStart.value;
				var dtStart = Date.parse(strDate);
				if (dtStart < Date.parse('<% = application.contents("dtSchool_Year_Start" & session.contents("intSchool_Year")) %>') || dtStart > Date.parse('<% = application.contents("dtSchool_Year_End" & session.contents("intSchool_Year")) %>')){
					alert("Start date cannot be less than the start of the current School Year\n<% = application.contents("dtSchool_Year_Start" & session.contents("intSchool_Year")) %> or greater than <% = application.contents("dtSchool_Year_End" & session.contents("intSchool_Year")) %>");
					return false;
				}
				if (checkDate(objForm.yearEnd, objForm.monthEnd, objForm.dayEnd, "End Date") == false) {return false;}						
				strDate = objForm.monthEnd.value+'/'+objForm.dayEnd.value+'/'+objForm.yearEnd.value;
				var dtEnd = Date.parse(strDate);
				if (dtEnd < Date.parse('<% = application.contents("dtSchool_Year_Start" & session.contents("intSchool_Year")) %>') || dtEnd > Date.parse('<% = application.contents("dtSchool_Year_End" & session.contents("intSchool_Year")) %>')){
					alert("End date cannot be less than the start of the current School Year\n<% = application.contents("dtSchool_Year_Start" & session.contents("intSchool_Year")) %> or greater than the end of the school year, <% = application.contents("dtSchool_Year_End" & session.contents("intSchool_Year")) %>");
					return false;
				}
				if (dtStart > dtEnd){
					alert("The Class Start Date must come before the Class End Date.");
					return false;
				}
			}
		<% else %>
			var intDuration_ID	= objForm.intDuration_ID.value;
			var intSession_Minutes	= objForm.intSession_Minutes.value;
			if(intDuration_ID.length == 0) {strErrMsg += 'Class Duration\n';}
			if(intSession_Minutes.length == 0) {strErrMsg += 'Session Length\n';}
		<% end if %>
		//these are the required fields - they must be populated
		if(szClass_Name.length == 0) {strErrMsg += 'Class Name\n';}
		
		if(intMin_Students.length == 0) {
			strErrMsg += 'Min No. of Students\n';
		}else{
			if (isInteger(intMin_Students) == false) {strErrMsg += 'Min No. of Students must be a number\n';}
		}
		if(intMax_Students.length == 0) {
			strErrMsg += 'Max No. of Students\n';
		}else{
			if (isInteger(intMax_Students) == false) {strErrMsg += 'Max No. of Students must be a number\n';}
		}
		
		if(!isChecked && szSchedule_Comments.length == 0) {strErrMsg += 'Either \'Meets Every\' or \'Schedule Comments\' must be filled in\n';}
			
		if (strErrMsg.length == 0 ) {
			return true;
		}else{
			strErrMsg = 'Please Enter/Correct the Following:\n \n' + strErrMsg;
			alert(strErrMsg);
			return false;
		}
	}
	
	function jfNameChange(){
		// This function is used to tell viewClasses.asp to refresh so the 
		// class name list will be correct.
		document.all.item('new_name').value = "true";
	}
</script>
<form action="classInsert.asp" method=Post name=main onSubmit="return false;">
<input type=hidden name=changed value="">
<input type=hidden name=bolValidated value="<% = request("bolValidated") %>">
<input type=hidden name=intInstructor_ID value="<% = intInstructor_ID %>">
<input type=hidden name=intGuardian_ID value="<% = Request.QueryString("intGuardian_ID") %>">
<input type=hidden name=intILPGenericID value="<% = Request.QueryString("intILPGenericID") %>">
<input type=hidden name=resourceList value="">
<input type=hidden name=new_name value="">
<input type=hidden name=maxCost value="">
<input type="hidden" name="hdnDays_Meet_On" value="">
<% if session.contents("intSchool_Year") = 2007 then %>
<span class="sverror"><b>Final ASD teacher cost per student and final deduction per student account
							will change subject to the completion of negotiations between the teachers 
							and management for SY 11/12.</b></span>
<% end if %>
<table width=100%>
	<tr>	
		<Td class=yellowHeader valign=top>
				&nbsp;<b><% = strClassTitle %></b><%= strInstructMessage%> &nbsp;&nbsp;	
				<%=strStudentName%>
		</td>
	</tr>
	<tr>
		<td bgcolor=f7f7f7>
			<table style="width:200px;"><tr><td>
		<%  if request("viewing") <> "" then %>
		<table cellpadding="4">									
				<tr class="TableHeader"> 	
					<%
						if (request("fromILP") <> "" and request("intGuardian_ID") = "") or request("intContract_Guardian_ID") <> "" then
					%>	
					<td>
						<b>GUARDIAN ON CONTRACT</b>
					</td>	
					<td >
						<select name="intGuardian_ID2" onChange="jfChanged();">
							<option value="">
							<%
							'JD: Show only 'active' guardians
								dim sqlGaurdian
								sqlGaurdian = "Select g.intGuardian_ID,g.szLast_Name + ',' + g.szFirst_Name as Name " & _
												 "from tblGuardian g, tascStudent_Guardian sg " & _
												 "where sg.intStudent_id = " & intStudent_ID & _
												 " and g.intGuardian_id = sg.intGuardian_ID " & _
												 " and g.blnDeleted = 0 " & _
												 " order by szLast_Name"										 
								Response.Write oFunc.MakeListSQL(sqlGaurdian,"intGuardian_ID","Name",request("intContract_Guardian_ID"))												 
							%>
						</select>																	
					</td>
					<% end if %>
					<td>
						<b>ACCEPT OR REJECT THIS <%= ucase(strFormType)%>: </b>
					</td>	
					<td align=center>					
						<input type=button name=accept value="ACCEPT" class="btSmallGray" onCLick="jfCheckCG();">
						<input type=button name=accept value="REJECT" class="btSmallGray" onClick="window.location.href='<%=Application.Value("strWebRoot")%>forms/ilp/ilp1.asp?<%=replace(ucase(Request.ServerVariables("QUERY_STRING")),"INTCLASS_ID","intClass_ID2")%>';">
					</td>
				</tr>										
			</table>
			<% end if %>
			<script language=javascript>
				function jfGetContract(id){
					var strIDS = id.value;
					var arIds = strIDS.split("|");
					var url = "<%=Application("strSSLWebRoot")%>forms/teachers/classAdmin.asp<%=strParams%>";
					url += "&bolHideGoodsServices=true&intClass_id=" + arIds[0] + "&intILPGenericID=" + arIds[1];
					window.location.href = url;
				}
			</script>
			<% = strSelectContract %>
			<table style="width:100%;">
				<tr>	
					<Td colspan=2>
						<font class=svplain11>
							<b><i><% = strFormType %> Information</I></B> 
						</font>
						<font class=svplain>
						</font>
					</td>
				</tr>
				<tr>
					<td class=gray valign=middle>
							&nbsp;Subject
					</td>
					<td class=gray valign=middle nowrap>
							&nbsp;Name of Class&nbsp;
					</td>
					<td class=gray valign=middle align="center">
							<nobr>&nbsp;ASD Course ID &nbsp;</nobr>
					</td>					
					<td class=gray align="center" valign=middle nowrap>
						&nbsp;Location &nbsp;
					</td>											
				</tr>
				<tr>									
					<td class=svplain10>
						<% 
							if (intInstructor_ID <> "" and intStudent_ID = "") or oFunc.IsAdmin then 
								' We are in Teacher Mode 
						%>
						<select name="intPOS_Subject_ID" onChange="jfChanged();"  <% = strDisabled  %> ID="Select1">
							<option value="">
						<%
							sql = "select intPOS_Subject_ID, upper(szSubject_Name) Name from trefPOS_Subjects where bolShow = '1' order by szSubject_Name"
							Response.Write oFunc.MakeListSQL(sql,"intPOS_Subject_ID","Name",intPOS_Subject_ID)	
						%>
						</select>			
						<% 
							else  
								'We are in Parent mode		
								response.Write  szSubject_Name & "&nbsp;"
						%>
						<input type=hidden name=intPOS_Subject_ID value="<% = intPOS_Subject_ID %>">
						<%  end if %>
					
					</td>
					<td class=svplain10>
					<!--Commneted out to allow editing of parent tought classes -->
						<% 'if intInstructor_ID & "" = "" then 
							'	if session.Contents("intShort_ILP_ID") <> "" then
							'		myClassName = oFunc.CourseInfo(session.Contents("intShort_ILP_ID"),3)
						%>
							<% '= myClassName
						%>
						<!--<input type=hidden name="szClass_Name"  maxlength=64 value="<% '= myClassName %>" -->
						<%'		else %>
						<%' = szClass_Name %> 
						
						<!--<input type=hidden name="szClass_Name" value="<% = szClass_Name %>" ID="Hidden1"  maxlength=64 > -->
						<%'		end if
						  ' elseif bolLock then
						%>
						<%' = szClass_Name %>
						<%'else %>
						<input type=text name="szClass_Name" value="<% = szClass_Name%>" maxlength=64 size=20 onChange="jfChanged();jfNameChange();" <% = strDisabled%> ID="Text2">
						<%' end if %>
					</td>	
					<td class=svplain10>
						<% 
						if szASD_Course_ID & "" <> "" then
								strASDWrite =  szASD_Course_ID
						else
							if session.Contents("intShort_ILP_ID") <> "" then
								strASDWrite =  oFunc.CourseInfo(session.Contents("intShort_ILP_ID"),5)  
							end if
						end if 
						if not bolLock then
						%>
						<input type=text name="szASD_Course_ID" value="<% = strASDWrite%>" maxlength=13 size=13 onChange="jfChanged();" <% = strDisabled%> ID="Text1">
						<%else%>
						<% = strASDWrite %>
						<%end if %>
					</td>
					<td class=svplain10 align="center">
						<% if not bolLock then%>
						<input type=text name="szLocation" value="<% if szLocation & "" = "" and intInstructor_ID & "" = "" then response.write "HOME" else response.Write szLocation end if%>" maxlength=50 size=17 onChange="jfChanged();"  <% = strDisabled%> >
						<%else%>
						<% = szLocation %>
						<%end if %>
					</td>									
				</tr>
			</table>
			<% 
				' Determine which table to show
				if intInstructor_ID & "" <> "" then
					response.Write vbfInstructorFields 
				else
					response.Write vbfGuardianFields
				end if
				
				if intInstructor_ID <> "" and session.Contents("intStudent_ID") = "" _
					or intInstructor_ID & "" = "" then
			%>
			<table cellpadding=4 style="width:100%;">
				<tr>	
					<Td colspan=4>
						<font class=svplain11>
							<b><i><% = strFormType %>  Restrictions</I></B> 
						</font>
					</td>
				</tr>
				<tr>
					<td class=gray width=225 valign=top>
						&nbsp;To limit enrollment to specific families select a family
						you wish to limit the class to and then click the <b>'Add Family >>'</b> button.											
					</td>	
					<td class=gray width=240 valign=top>
						&nbsp;If the box below is empty then this class 
						will be open to all families otherwise the class will only be
						accessible to the listed families.<br>
						&nbsp; To delete a family from the list, click on the family in the
						list and then click the <b>'X'</b> button.
					</td>								
				</tr>
				<TR>			
					<TD valign="top">
						<SELECT name="selRestrictedFamilies"  multiple size="6" style="FONT-SIZE:xx-small;width: 250px" onFocus="this.size=20;" onblur="this.size=6;">
							<option>----------						
							<%
							sqlFamilies = "SELECT DISTINCT  " & _ 
										" tblFAMILY.intFamily_ID, CASE WHEN szDesc IS NULL THEN upper(szFamily_Name) WHEN szDesc IS NOT NULL  " & _ 
										" THEN upper(szFamily_Name + ', ' + szDesc) END AS Name " & _ 
										" FROM tblFAMILY INNER JOIN " & _ 
										" tblSTUDENT ON tblFAMILY.intFamily_ID = tblSTUDENT.intFamily_ID INNER JOIN " & _ 
										" tblStudent_States ON tblSTUDENT.intSTUDENT_ID = tblStudent_States.intStudent_id " & _ 
										" WHERE (tblStudent_States.intSchool_Year = " & session.Contents("intSchool_Year") & ") AND (tblStudent_States.intReEnroll_State  IN (" & Application.Contents("strEnrollmentList") & ") ) " & _ 
										" ORDER BY Name "
										response.Write sqlFamilies
							Response.Write oFunc.MakeListSQL(sqlFamilies,"intFamily_ID","Name","")
							%>
						</SELECT>
					</td>
					<TD valign="top" width=0%>
						<SELECT name="selFamilies"  multiple size="6" style="FONT-SIZE:xx-small;width:250px" ID="Select2">
							<%
							if strRestrictedFamList = "" and intStudent_ID <> "" and intInstructor_ID & "" = "" then
								arFamInfo = oFunc.FamilyInfo("1",intStudent_ID,"6")
								strRestrictedFamList = "<option value='" & arFamInfo(0) &"'>" & arFamInfo(1) & ", " & arFamInfo(2)
							end if 
							response.Write	strRestrictedFamList			 
							%>
						</SELECT>
					</TD>
				<tr>
					<td align=right>
						<input type=button value="Add Family >>" title="Add selected Family" class="btSmallGray"
						onclick="jfSelectItemFromTo('selRestrictedFamilies', 'selFamilies');" align=right NAME="Button2">
					</td>
					<TD valign=middle align=left>
						<input type=hidden name=intFamily_ID ID="Hidden2">
						<input type=button value="X" style="position:relative"  class="btSmallGray" title="Remove selected Family or Families" onclick="jfRemoveItems('selFamilies');">
					</TD>
				</tr>			
			</table>	
			<BR>	
<% 
			end if ' for show only when no student id
dim dblGrandTotal
dblGrandTotal = 0

if intClass_ID <> "" and request.QueryString("bolHideGoodsServices") = "" _
	AND intInstructor_ID <> "" then	
		'sqlItems = "SELECT (CASE ci.intItem_ID WHEN 3 THEN " & _
		'			" (SELECT vs.szVend_Service_Name " & _
		'			" FROM trefVendor_Services vs, tblClass_Attrib ca2 " & _
		'			" WHERE ci.intClass_Item_ID = ca2.intClass_Item_ID AND ca2.intItem_Attrib_Id = 26 AND vs.intVend_Service_ID = ca2.szValue)  " & _
		'			" ELSE ca.szValue END) AS Description, ci.intQty, ci.curUnit_Price,ci.curShipping, i.szName, ci.intClass_Item_ID AS ExistingItemID, ig.szName AS ItemType,  " & _
		'			" i.intItem_Group_ID " & _
		'			"FROM tblClass_Items ci INNER JOIN " & _
		'			"	tblClass_Attrib ca ON ci.intClass_Item_ID = ca.intClass_Item_ID INNER JOIN " & _
		'			"	trefItems i ON ci.intItem_ID = i.intItem_ID INNER JOIN " & _
		'			"	trefItem_Groups ig ON i.intItem_Group_ID = ig.intItem_Group_ID " & _
		'			"WHERE (ca.intOrder = 1) AND  " & _
		'			"    (ci.intClass_ID = " & intClass_Id & ") " & _
		'			"ORDER BY i.intItem_Group_ID, i.szName"
		
		
		sqlItems =  "SELECT ci.intClass_Item_ID as ExistingItemID,v.szVendor_Name, ig.intItem_Group_ID,ig.szName AS ItemType, i.szName, " & _
			   "ci.intQty, ci.curUnit_Price,ci.curShipping,((ci.intQty * ci.curUnit_Price)+ci.curShipping) as Total, " & _
               "           (SELECT     ca2.szValue " & _
               "            FROM          tblClass_Attrib ca2 " & _
               "            WHERE      ca2.intClass_Item_Id = ci.intClass_Item_Id AND " & _
               "			ca2.intItem_Attrib_ID = 15) AS Consumable, " & _
               "           (SELECT     ca2.szValue " & _
               "            FROM          tblClass_Attrib ca2 " & _
               "            WHERE      ca2.intClass_Item_Id = ci.intClass_Item_Id AND " & _
               "			ca2.intItem_Attrib_ID = 3) + ' - ' + " & _
               "           (SELECT     ca2.szValue " & _
               "            FROM          tblClass_Attrib ca2 " & _
               "            WHERE      ca2.intClass_Item_Id = ci.intClass_Item_Id AND " & _
               "            ca2.intItem_Attrib_ID = 4) AS Dates, " & _
               "          (SELECT  top 1   ca2.szValue " & _
               "             FROM          tblClass_Attrib ca2 " & _
               "             WHERE      ca2.intClass_Item_Id = ci.intClass_Item_Id AND " & _
               "			 (ca2.intItem_Attrib_ID = 9 OR " & _
               "              ca2.intItem_Attrib_ID = 5 OR " & _
               "              ca2.intItem_Attrib_ID = 6 OR " & _
			   "              ca2.intItem_Attrib_ID = 22 or ca2.intItem_Attrib_ID = 33) order by ca2.intItem_Attrib_ID ) AS Description, '' as szDeny_Reason, 1 as bolApproved, intContract_Status_ID " & _
		       " FROM tblClass_Items ci INNER JOIN " & _
               "       tblVendors v ON ci.intVendor_ID = v.intVendor_ID INNER JOIN " & _
               "       trefItems i ON ci.intItem_ID = i.intItem_ID INNER JOIN " & _
               "       trefItem_Groups ig ON i.intItem_Group_ID = ig.intItem_Group_ID inner join " & _
               "	   tblClasses c ON c.intClass_ID = ci.intClass_ID " & _
			   " WHERE (ci.intClass_ID = " & intClass_ID & ")  and ci.bolRequired = 1 " & _
			   "order by i.szName "
			   		
					
	'else
		'This sql needs to join ord_items to tblILP to tblClasses where inClass_ID = Val
	'	sqlItems = "SELECT trefItems.szName, tblOrd_Attrib.szValue, " & _
	'			   "tblOrdered_Items.intQty, tblOrdered_Items.curUnit_Price, " & _
	'			   "tblOrdered_Items.intOrdered_Item_ID AS ExistingItemID," & _
	'			   " trefItem_Groups.szName AS ItemType, trefItems.intItem_Group_ID " & _
	'			   "FROM tblOrdered_Items INNER JOIN " & _
     '              "tblOrd_Attrib ON " & _
      '             "tblOrdered_Items.intOrdered_Item_ID = tblOrd_Attrib.intOrdered_Item_ID " & _
     '              "INNER JOIN " & _
      '             "trefItems ON tblOrdered_Items.intItem_ID = trefItems.intItem_ID " & _
     '              "INNER JOIN " & _
     '              "trefItem_Groups ON " & _
     '              "trefItems.intItem_Group_ID = trefItem_Groups.intItem_Group_ID " & _
     '              "INNER JOIN " & _
     '              "tblILP ON tblOrdered_Items.intILP_ID = tblILP.intILP_ID " & _
     '              "INNER JOIN " & _
     '              "tblClasses ON tblILP.intClass_ID = tblClasses.intClass_ID " & _
	'			   "WHERE (tblOrd_Attrib.intOrder = 1) " & _
	'			   " AND (tblClasses.intClass_ID = " & intClass_ID & ") " & _
	'			   "ORDER BY trefItems.intItem_Group_ID, trefItems.szName "				
	'end if
	set rsItems = server.CreateObject("ADODB.Recordset")
	rsItems.CursorLocation = 3
	rsItems.Open sqlItems, Application("cnnFPCS")'oFunc.FPCScnn
	
	
	if rsItems.RecordCount < 1 then		
	%>
	<table ID="Table1" style="width:100%;">
		<tr>
			<Td class=gray>
				&nbsp;No Goods or Services have been added to this class.
			</td>
		</tr>
	</table>
	<br>
	<%
	else
	%>
	<table cellpadding=3 ID="Table2">
		<tr>
			<td class=svplain11 colspan=6>
				<B><i>Required Goods and Services</i></b>
			</td>
		</tr>
		<tr>
			<Td class=gray align=center>
				<b>Type</b>
			</td>
			<td class=gray align=center>
				<b>Category</b>
			</td>
			<Td class=gray align=center>
				<b>Name</b>
			</td>
			<td class=gray align=center>
				<b>Qty</b>
			</td>
			<Td class=gray align=center>
				<b>Unit Price</b>
			</td>
			<Td class=gray align=center>
				<b>Shipping</b>
			</td>
			<td class=gray align=center>
				<b>Total</b>
			</td>
		</tr>	
	<%		
		dim dblTotal
		
		do while not rsItems.EOF
	%>
		<tr>
			<Td class=gray align=center>
				<% = rsItems("ItemType") %>
			</td>
			<td class=gray align=center>
				<% = rsItems("szName") %>
			</td>
			<Td class=gray align=center>
				<% if rsItems("Description") & "" = "" then
						response.Write rsItems("szName")
				else
						response.Write rsItems("Description")
				end if			
				%>
			</td>
			<td class=gray align=center>
				<% = rsItems("intQty") %>
			</td>
			<Td class=gray align=center>
				$<% = rsItems("curUnit_Price") %>
			</td>
			<Td class=gray align=center>
				$<% = rsItems("curShipping") %>
			</td>
			<td class=gray align=right>
				<% 
					dblTotal = round((cdbl(rsItems("intQty")) * cdbl(rsItems("curUnit_Price"))) + cdbl(rsItems("curShipping")),2)
					dblGrandTotal = dblGrandTotal + dblTotal
					Response.Write "$" & formatNumber(dblTotal,2)
				%>
			</td>
		</tr>
	<%
			rsItems.MoveNext
		loop
	%>
		<tr>
			<td colspan=6 class=gray align=right>
				<B>Grand Total:</b>
			</td>
			<td class=gray align=right>
				$<%  if isNumeric(dblGrandTotal) then response.Write formatNumber(dblGrandTotal,2) end if %>
			</td>	
		</tr>
	</table>	
	<br>
	<%
	end if
	rsItems.Close
	set rsItems = nothing	
end if 

if intInstructor_ID <> "" then
	' Show teacher costs table
	call vbfClassDetailsForASD
end if
%>			</td></tr></table>
		</td>
	</tr>
</table>
<% 
' first check to see if year is locked
if not oFunc.LockYear then
	
						
if request("viewing") = "" and request("isPopUp") = "" then 
' Class Creation Buttons%>
<input type=button value="Close without saving" onClick="window.location.href='<%=Application.Value("strWebRoot")%>';" class="NavLink" >
<!--<input type=submit value="ADD CLASS" id="btSmallGray" onClick="jfSubmit(this.form);">-->
<input type=submit value="SAVE (ilp is next)" class="NavSave" onClick="jfSubmit(this.form);">

<% elseif (ucase(session.Contents("strRole")) = "ADMIN" or ucase(session.Contents("strRole")) = "TEACHER") and (intStudent_ID = "") then  
	if ucase(session.Contents("strRole")) = "TEACHER" then
		
		if intContract_Status_ID & "" = "5" then
			%>
			<span class="svplain8">
				Since an Admin has already signed off on this contract it is only editable by the FPCS Office.<br>
				Please contact the FPCS office for changes. <br>
				<input type=button value="Close without saving" onClick="window.opener.focus();window.close();" class="NavLink" name=button1>
			</span>
			<% 
		else
			' Edit buttons for Instructors%>
			<input type=button value="Close without saving" onClick="window.opener.focus();window.close();" class="btSmallGray" name=button1>
			<input type=submit value="SAVE" class="NavSave" name=submit1 onClick="jfSubmit(this.form);">
			<input type=hidden name="edit" value="yes">
			<input type=hidden name="intClass_ID" value="<% = intClass_Id %>">
<% 
		end if
	else
%>
			<input type=button value="Close without saving" onClick="window.opener.focus();window.close();" class="btSmallGray" name=button1>
			<input type=submit value="SAVE" class="NavSave" name=submit1 onClick="jfSubmit(this.form);">
			<input type=hidden name="edit" value="yes" ID="Hidden3">
			<input type=hidden name="intClass_ID" value="<% = intClass_Id %>" ID="Hidden4">
<%
	end if
elseif intInstructor_ID & "" = "" and request("isPopUp") & "" <> ""  then 
' Edit buttons for Guardians%>

<input type=button value="Close without saving" class="navLink" onClick="window.opener.focus();window.close();" id=button1 name=button1>
<input type=submit value="SAVE" class="NavSave" onClick="jfSubmit(this.form);">
<input type=hidden name="edit" value="yes">
<input type=hidden name="intClass_ID" value="<% = intClass_Id %>">
<%elseif request("viewing") = "" then%>
<input type=button value="Close Window" onclick="opener.window.focus();window.close();">
<% end if 
else
%>
	<span class="svplain8"><b>This school year has been locked. No modifications can be made.</b></span><br>
	<br><input type=button class="Navlink" value="Close Window" onclick="opener.window.focus();window.close();" ID="Button2" NAME="Button2">
<%
end if
%>
</form>
<script language=javascript>	
	var mstrVenSel = "";
	function jfAddHRS(){
		var intStudentHours = document.main.decHours_Student.value;
		var intHRS_Planning = document.main.decHours_Planning.value;
		var intMinStudent = document.main.intMin_Students.value;
		var intMaxStudent = document.main.intMax_Students.value;
		var intTotalHours = parseFloat(intStudentHours) + parseFloat(intHRS_Planning);
		var intRate = document.main.curRate.value;
		var curMiscTotal = 0;
		var curItemTotal;
		if (intStudentHours == "" || intHRS_Planning == "" || intMinStudent == ""
				|| intMaxStudent == "") {
			var strMessage;
			strMessage = "To Calculate totals you must provide a value for \n";
			strMessage += "'Min # of Students'\n'Max # of Students'\n";
			strMessage += "'Number of teacher hours with student'\n";
			strMessage += "'Number of hours for teacher planning'.";
			alert(strMessage);
			return;
		}
		
		intRate = intRate.replace("$","");
		document.main.totalHours.value = intTotalHours;
		document.main.intMax_Charged.value = round(intTotalHours / parseFloat(intMinStudent),4);
		document.main.intMin_Charged.value = round(intTotalHours / parseFloat(intMaxStudent),4);
		document.main.intMaxTeacherCost.value  =  "$" + round((parseFloat(document.main.intMax_Charged.value) * parseFloat(intRate)));
		document.main.intMinTeacherCost.value  =  "$" + round((parseFloat(document.main.intMin_Charged.value) * parseFloat(intRate)));	
		var max1 = document.main.intMaxTeacherCost.value;
		var min1 = document.main.intMinTeacherCost.value
		var materials = '<% = dblGrandTotal %>';
		
		max1 = max1.replace("$","");
		min1 = min1.replace("$","");
		materials = materials.replace("$","");
		document.main.intMinTotalCost.value  =  "$" + round((parseFloat(min1) + parseFloat(materials)));
		document.main.intMaxTotalCost.value  =  "$" + round((parseFloat(max1) + parseFloat(materials)));	
	}
	<% if intClass_id <> ""   then 
		response.write strCalcType 
	 end if %>
	 
	function round(number,X) {
		// rounds number to X decimal places, defaults to 2
		X = (!X ? 2 : X);
		//return Math.floor(number*Math.pow(10,X))/Math.pow(10,X);
		return Math.round(number*Math.pow(10,X))/Math.pow(10,X);
	}

</script>
<% 
call oFunc.CloseCN
Server.Execute(Application.Value("strWebRoot") & "Includes/footer.asp")
 
function vbfClassDetailsForASD
%>
	<input type=hidden name="dtEffective">
	<input type=hidden name="startStudentHrs" value="<% = decHours_Student %>">
	<input type=hidden name="startPlanningHrs" value="<% = decHours_Planning %>">
	<input type=hidden name="strReason">
	
	<script language=javascript>
		function jfHourChange(item){
			var bolChange;
			var bolStop;
			bolStop = false;
			// This if logic is needed if the user changes an hour field and then clicks 
			//'cancel' in the confirm dialog. The Hours will be reset to the starting hour
			// and since it's changed again the onChange event will fire. This logic prevents the 
			// onChange event from firing when the value is reset to starting value via cancel.
			if (item.name == "decHours_Student") {
				if (document.main.startStudentHrs.value == item.value){
					bolChange = false;
				}else{
					bolChange = confirm("Are you sure you want to change the hours for this class?");
				}
			} else {
				if (document.main.startPlanningHrs.value == item.value) {
					bolChange = false;
				}else{
					bolChange = confirm("Are you sure you want to change the hours for this class?");	
				}
			}
						
			
			if (bolChange){
				// Attempts to get a valid date unless action is canceled. 
				// A valid date (if given) is stored using jfGetDate
				if (document.main.dtEffective.value == "") {
					while (!bolStop){					
						bolStop = jfGetDate();
						if (bolStop == 'cancel') {
							jfResetHrs(item);
							bolStop = true;						
						}
					}					
				}
			} else {		
				// Hours are rest since cancel was clicked		
				jfResetHrs(item);
			}			
		}
		
		function jfResetHrs(item)	{
			// Resets hour field
			if (item.name == "decHours_Student") {
				item.value = document.main.startStudentHrs.value;
				document.main.decHours_Planning.focus();
			} else {
				item.value = document.main.startPlanningHrs.value;
				document.main.decHours_Student.focus();
			}
		}
		
		function jfGetDate(){					
			var dtChange;	
			dtChange = prompt("Enter the date this change is/was effective. (mm/dd/yyyy format.)\nOr type '0' to cancel.","");
			if (dtChange != null) {	
			
				if (dtChange == 0) {
					var strReturn = "cancel";
					return strReturn;
				}
				 
				arDate = dtChange.split("/");
				if (isDate(arDate[2],arDate[0],arDate[1])) {
					var dtOrigEffective = "<%= dtHrs_Last_Updated%>";
					if (dtOrigEffective != null) {
						// Date has to be newer than last effective date
						var dtOld = new Date(dtOrigEffective);
						var dtNew = new Date(dtChange);
						if (dtNew <= dtOld) {
							alert("The date you enter must be greater than " + dtOrigEffective + ".\nMake sure it is in mm/dd/yyyy format.");
							return false;
						}
					}
					// Have valid date so we store it
					document.main.dtEffective.value = dtChange;
					document.main.strReason.value = prompt("Please enter the reason for change in hours.","");
					jfAddHRS();
					return true;
				}
				else	{
					alert("The date you entered is invalid. Make sure it is in mm/dd/yyyy format");
					return false;
				}
			}	
			else	{
				alert("The date you entered is invalid. Make sure it is in mm/dd/yyyy format");
				return false;
			}						
		}
		
		<% if intStudent_ID <> "" and session.contents("intShort_ILP_ID") <> "" then %>
		function jfCheckFunds(){
			var strAlert;						
			<%
				set oBudget = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/StudentBudgetInfo.wsc"))
				set oClass = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/ClassInfo.wsc"))
				
				'oClass.PopulateObject oFunc.FpcsCnn, intClass_ID
				oClass.PopulateObject Application("cnnFPCS"), intClass_ID
				'oBudget.PopulateStudentFunding oFunc.FpcsCnn, intStudent_ID, session.contents("intSchool_Year")
				oBudget.PopulateStudentFunding Application("cnnFPCS"), intStudent_ID, session.contents("intSchool_Year")
	
				myBudget = oBudget.BudgetBalance
				bolLimit = false
				if oFunc.IsSpendingLimitSubject(intPOS_SUBJECT_ID) then
					'oBudget.PopulateFamilyBudgetInfo oFunc.FpcsCnn, oBudget.FamilyId, session.contents("intSchool_Year") 
					oBudget.PopulateFamilyBudgetInfo Application("cnnFPCS"), oBudget.FamilyId, session.contents("intSchool_Year") 
					if oBudget.BudgetBalance > oBudget.AvailableElectiveBudget then
						myBudget = oBudget.AvailableElectiveBudget
						bolLimit = true
					end if						
				end if				
				
				dblMaxCost = formatNumber(round(oClass.ProjectedTeacherCost,2) + dblGrandTotal,2)
				dblBudgetAfterCost = formatNumber(myBudget - dblMaxCost,2)
				set oBudget = nothing
				set oClass = nothing
			%>
				strAlert =  " ";
				strAlert += "The cost of this class can range from " + document.main.intMinTotalCost.value ;
				strAlert += " to " + document.main.intMaxTotalCost.value;
				strAlert += " based on actual enrollment which is determined the first day of class.";	
			<%
				if dblBudgetAfterCost >= 0 or oFunc.IsAdmin then 
					if dblBudgetAfterCost < 0 and oFunc.IsAdmin then 
						%>
						strAlert += '\n\nADMIN ALERT! You are about to give this account a negative balance!\n\n';
						<%
					end if
				if session.contents("intSchool_Year") = 2006 then
			%>		
				strAlert += " Final ASD teacher cost per student and final deduction per student account " +
							"will change subject to the completion of negotiations between the teachers " +
							"and management for SY 11/12.\n\n";
			<%
				end if
			%>								
				strAlert += " This cost will fluxuate as enrollment changes and will automatically be reflected on your students budget and statement."
				strAlert += " Your budget will be automatically adjusted to refect the teacher cost.";
				strAlert += " ";
				var bolContinue = confirm(strAlert);
				if (bolContinue) {
					document.main.maxCost.value = "<% = dblMaxCost %>";
					return true; 
				}else{
					return false;
				}		   
			<% else %>
				
				strAlert = "<% if bolLimit then %>This Class is subject to the Famliy Elective Spending Limit.\nThis family has $<% = round(myBudget,2) %> left for elective spending.\n<% end if %>Currently you do not have the needed funds to enroll in this class.";
				strAlert += " Adding this class would give you ";
				strAlert += "a <% if bolLimit then %>Family Elective Spending <% end if %>balance of -$<% = formatNumber((dblBudgetAfterCost*-1),2)%>. You can not add this class ";
				strAlert += "until you free up some funding (by transfering funds or deleting an existing expense) or until more students enroll bringing the total class cost down. "
				alert(strAlert);
				window.location.href = "<% =Application.Value("strWebRoot")%>forms/packet/packet.asp?intStudent_ID=<% = intStudent_ID %>";
				return false;	
			<% end if %>
		}
		<% end if %>
	</script>
			<table>
				<tr>	
					<Td colspan=2>
						<font class=svplain11>
							<b><i>Class Costs</I></B> 
						</font>
					</td>
				</tr>
				<tr>
					<td class=gray>
						<input type=text name="decHours_Student" value="<% = decHours_Student %>" size=8 maxlength=5 onChange="jfChanged();<% if intClass_ID <> "" and Request.QueryString("strThisIsACopy") = "" then Response.Write("jfHourChange(this);") %>"  <% = strDisabled%> >
					</td>		
					<td class=gray>
						&nbsp;Number of teacher hours with student.
					</td>							
				</tr>		
				<tr>
					<td class=gray>
						<input type=text name="decHours_Planning" value="<% = decHours_Planning %>" size=8 maxlength=5 onChange="jfChanged();<% if intClass_ID <> ""  and Request.QueryString("strThisIsACopy") = "" then Response.Write("jfHourChange(this);") %>" <% = strDisabled%> >
					</td>		
					<td class=gray>
						&nbsp;Number of hours for teacher planning.
					</td>							
				</tr>	
				<% if strDisabled = "" then %>
				<tr>
					<td class=gray align=center>
						&nbsp;=&nbsp;
					</td>
					<td class=gray>
						<input type=button value="calculate totals" onClick="<%=strCalcType%>" class="btSmallGray">
					</td>							
				</tr>	
				<% end if %>
				<tr>
					<td class=gray>
						<input type=text name="totalHours"  size=8 maxlength=4 disabled>
					</td>		
					<td class=gray>
						&nbsp;<B>Total teacher hours.</b>
					</td>							
				</tr>	
 				<tr>
					<td class=gray>
						<input type=text name="intMin_Charged" value="" size=8 maxlength=10 disabled>
					</td>		
					<td class=gray>
						&nbsp;Minimum number of hours to be charged to each student.
					</td>							
				</tr>	
				<tr>
					<td class=gray>
						<input type=text name="intMax_Charged" value="" size=8 maxlength=10 disabled>
					</td>		
					<td class=gray>
						&nbsp;Maximum number of hours to be charged to each student.
					</td>							
				</tr>	
				<tr>
					<td class=gray>
					
						<input type=text name="curRate" value="$<% = curInstructionRate %>" size=8 maxlength=10 disabled>
					</td>		
					<td class=gray>
						&nbsp;Teachers hourly rate.
					</td>							
				</tr>
				<tr>
					<td class=gray>
						<input type=text name="intMinTeacherCost" value="" size=8 maxlength=10 disabled>
					</td>		
					<td class=gray>
						&nbsp;Minimum total teacher cost per student.
					</td>							
				</tr>	
				<tr>
					<td class=gray>
						<input type=text name="intMaxTeacherCost" value="" size=8 maxlength=10 disabled>
					</td>		
					<td class=gray>
						&nbsp;Maximum total teacher cost per student.
					</td>							
				</tr>
				<tr>
					<td class=gray>
						<input type=text name="intMiscCost" value="$<% = dblGrandTotal %>" size=8 maxlength=10 disabled>
					</td>		
					<td class=gray>
						&nbsp;Total miscellaneous costs per student.
					</td>							
				</tr>		
				<tr>
					<td class=gray>
						<input type=text name="intMinTotalCost" value="" size=8 maxlength=10 disabled>
					</td>		
					<td class=gray>
						&nbsp;<B>Minimum total deduction per student account.</b>
					</td>							
				</tr>	
				<tr>
					<td class=gray>
						<input type=text name="intMaxTotalCost" value="" size=8 maxlength=10 disabled>
					</td>		
					<td class=gray>
						&nbsp;<B>Maximum total deduction per student account.</b>
					</td>							
				</tr>	
			</table>	
<%
end function

function vbfInstructorFields
%>
			<table style="width:100%;">
				<tr>
					<td class=gray colspan=3 nowrap align="center">
						&nbsp;Registration Deadline
					</td>
					<td class=gray nowrap align="center">
						&nbsp;Min # Students
					</td>	
					<td class=gray nowrap align="center">
						&nbsp;Max # Students
					</td>	
					<td class=gray nowrap align="center">
						&nbsp;Grade&nbsp;
					</td>		
					<td class=gray nowrap align="center">
						&nbsp;to Grade&nbsp;
					</td>																
				</tr>
				<tr>
					<% if not bolLock then%>
					<td>						
						<select name="regMonth" onChange="jfChanged();"  <% = strDisabled%> >
							<option value="">Month</option>
							<% 
							dim sqlMonth
							sqlMonth = "select strValue,strText from common_lists where intList_ID = 1 order by intOrder"
							Response.Write oFunc.MakeListSQL(sqlMonth,"","",regMonth)								
							%>
						</select>
					</td>		
					<td>
						<select name="regDay" onChange="jfChanged();" <% = strDisabled%> >
							<option value="">Day</option>
							<% 
							dim sqlDay
							sqlDay = "select strValue,strText from common_lists where intList_ID = 2 order by intOrder"
							Response.Write oFunc.MakeListSQL(sqlDay,"","",regDay)								
							%>
						</select>
					</td>											
					<td>
						<select name="regYear" onChange="jfChanged();" <% = strDisabled%> >
							<option value="">Year</option>
							<% = oFunc.MakeYearList(2,1,regYear) %>
						</select>
					</td>	
					<% else %>
					<td colspan="3" class="svplain10" align="center">
						<% = regMonth & "/" & regDay & "/" & regYear %>
					</td>
					<%end if%>
					<td align=center>
						<input type=text  <% = strDisabled%> name="intMin_Students" value="<% = intMin_Students%>" maxlength=3 size=4 onChange="jfChanged();">
					</td>	
					<td align=center>
						<input type=text <% = strDisabled%>  name="intMax_Students" value="<% = intMax_Students%>" maxlength=3 size=4 onChange="jfChanged();">
					</td>	
					<td align=center  class="svplain10">
						<% if not bolLock then%>
						<select name="sGrade_Level" onChange="jfChanged();" <% = strDisabled%> >
							<option value="">
							<% 
							dim strGradeList
							strGradeList = "K,1,2,3,4,5,6,7,8,9,10,11,12"
							Response.Write oFunc.MakeList(strGradeList,strGradeList,replace(sGrade_Level & ""," ",""))								
							%>
						</select>
						<%else%>
						<% = sGrade_Level %>
						<%end if %>
					</td>			
					<td align=center class="svplain10">
						<% if not bolLock then%>
						<select name="sGrade_Level2" onChange="jfChanged();" <% = strDisabled%> >
							<option value="">
							<% 
							Response.Write oFunc.MakeList(strGradeList,strGradeList,replace(sGrade_Level2 & ""," ",""))								
							%>
						</select>
						<%else%>
						<% = sGrade_Level2 %>
						<%end if %>
					</td>								
				</tr>
			</table>
			<table style="width:100%;">				
				<tr>
					<td class=gray colspan=3 align="center">
						&nbsp;Class Start Date
					</td>
					<td class=gray colspan=3 align="center">
						&nbsp;Class End Date
					</td>	
					<% if bolLock then %>
					<td class=gray align="center">
						&nbsp;Meets Every
					</td>		
					<% end if %>														
				</tr>
				<tr>
					<% if not bolLock then%>
					<td valign=top>
						<select name="monthStart" onChange="jfChanged();" <% = strDisabled%> >
							<option value="">Month</option>
							<% 
							Response.Write oFunc.MakeListSQL(sqlMonth,"","",monthStart)								
							%>
						</select>
					</td>		
					<td valign=top>
						<select name="dayStart" onChange="jfChanged();" <% = strDisabled%> >
							<option value="">Day</option>
							<% 
							Response.Write oFunc.MakeListSQL(sqlDay,"","",dayStart)								
							%>
						</select>
					</td>											
					<td valign=top>
						<select name="yearStart" onChange="jfChanged();" <% = strDisabled%> >
							<option value="">Year</option>
							<% = oFunc.MakeYearList(2,1,yearStart) %>
						</select>		
					</td>		
					<% else %>
					<td colspan="3" class="svplain10" align="center">
						<% = monthStart & "/" & dayStart & "/" & yearStart %>
					</td>
					<%end if%>	
					<% if not bolLock then%>
					<td valign=top>
						<select name="monthEnd" onChange="jfChanged();" <% = strDisabled%> >
							<option value="">Month</option>
							<% 
							Response.Write oFunc.MakeListSQL(sqlMonth,"","",monthEnd)								
							%>
						</select>
					</td>		
					<td valign=top>
						<select name="dayEnd" onChange="jfChanged();" <% = strDisabled%> >
							<option value="">Day</option>
							<% 
							Response.Write oFunc.MakeListSQL(sqlDay,"","",dayEnd)								
							%>
						</select>
					</td>											
					<td valign=top>
						<select name="yearEnd" onChange="jfChanged();" <% = strDisabled%> >
							<option value="">Year</option>
							<% = oFunc.MakeYearList(2,1,yearEnd) %>
						</select>		
					</td>	
					<% else %>
					<td colspan="3" class="svplain10" align="center">
						<% = monthEnd & "/" & dayEnd & "/" & yearEnd %>
					</td>
					<%end if%>								
					
					<% if not bolLock then%>
					</tr>
					<tr>
						<td class="gray" colspan="8">
							&nbsp;Meets Every
						</td>
					</tr>
					<tr>
						<td class="svplain10" align="center" colspan="8">
					<%
							sql = "select strValue,strText from common_lists where intList_ID = 4 order by intOrder"
							response.Write oFunc.MakeCheckList(sql,"strText","strText",szDays_Meet_On,"szDays_Meet_On",6)
					%>
						</td>
					</tr>
					<% else %>
						<td class="svplain10" align="center">		
						<% = szDays_Meet_On %>
						</td>
					</tr>
					<% end if%>		
			</table>		
			<table style="width:100%;">				
				<tr>
					<td class=gray colspan=4 align="center">
							&nbsp;Class Start Time
					</td>
					<td class=gray>
							&nbsp;
					</td>
					<td class=gray colspan=4 align="center">
						&nbsp;Class End Time
					</td>		
					<td class=gray colspan=4 align="center">
						&nbsp;Schedule Comments
					</td>													
				</tr>
				<tr>
					<% if not bolLock then%>
					<td valign=top>
						<select name="hourStart" onChange="jfChanged();" <% = strDisabled%> >
							<% 
							dim strHour
							strHour = "1,2,3,4,5,6,7,8,9,10,11,12"
							Response.Write oFunc.MakeList(strHour,strHour,hourStart)								
							%>
						</select>
					</td>	
					<td valign=top>
						:
					</td>	
					<td valign=top>
						<select name="minuteStart" onChange="jfChanged();" <% = strDisabled%> >
							<% 
							dim strMinute
							dim str0
							strMinute = "00,01"
							for i = 2 to 60
								if i < 10 then str0 = "0"
								strMinute = strMinute & "," &  str0 & i
								str0 = ""
							next 
							Response.Write oFunc.MakeList(strMinute,strMinute,minuteStart)								
							%>
						</select>
					</td>											
					<td valign=top>
						<select name="amPmStart" onChange="jfChanged();" <% = strDisabled%> >
							<% 
							dim strAmPm
							strAmPm = "AM,PM"
							Response.Write oFunc.MakeList(strAmPm,strAmPm,amPmStart)								
							%>
						</select>		
					</td>		
					<% else %>
					<td colspan="4" class="svplain10" align="center">
						<% = hourStart & ":" & minuteStart & " " & amPmStart %>
					</td>
					<%end if%>	
					<td>
							&nbsp;
					</td>	
					<% if not bolLock then%>
					<td valign=top>
						<select name="hourEnd" onChange="jfChanged();" <% = strDisabled%> >
							<% 
							Response.Write oFunc.MakeList(strHour,strHour,hourEnd)								
							%>
						</select>
					</td>		
					<td valign=top>
						:
					</td>	
					<td valign=top>
						<select name="minuteEnd" onChange="jfChanged();" <% = strDisabled%> >
							<% 
							Response.Write oFunc.MakeList(strMinute,strMinute,minuteEnd)								
							%>
						</select>
					</td>											
					<td valign=top>
						<select name="amPmEnd" onChange="jfChanged();" <% = strDisabled%> >
							<% 
							Response.Write oFunc.MakeList(strAmPm,strAmPm,amPmEnd)								
							%>
						</select>		
					</td>	
					<% else %>
					<td colspan="4" class="svplain10" align="center">
						<% = hourEnd & ":" & minuteEnd & " " & amPmEnd %>
					</td>
					<%end if%>	
					<% if not bolLock then%>
					<td align=center>
						<textarea cols=20 rows=2 name="szSchedule_Comments" wrap=virtual onKeyDown="jfMaxSize(128,this);"  <% = strDisabled%> ><% = szSchedule_Comments%></textarea>						
					</td>	
					<% else %>
					<tdclass="svplain10" align="center">
						<% = szSchedule_Comments%>
					</td>
					<%end if%>	
				</tr>
			</table>
<%
end function

function vbfGuardianFields
	if intMin_Students = "" then
		intMin_Students = 1
	end if
	
	if intMax_Students = "" then
		intMax_Students = 1
	end if
%>
			<table style="width:100%;">
				<tr>
					<td class=gray>
						&nbsp;Min # Students &nbsp;
					</td>	
					<td class=gray>
						&nbsp;Max # Students &nbsp;
					</td>	
					<td class=gray>
						&nbsp;Class Duration &nbsp;
					</td>		
						
					<td class=gray>
						&nbsp;Session Length &nbsp;
					</td>																
				</tr>
				<tr>					
					<td align=center>
						<input type=text name="intMin_Students" value="<% = intMin_Students%>" maxlength=3 size=4 onChange="jfChanged();">
					</td>	
					<td align=center>
						<input type=text name="intMax_Students" value="<% = intMax_Students%>" maxlength=3 size=4 onChange="jfChanged();">
					</td>	
					<td>
						<select name="intDuration_ID" onChange="jfChanged();">
							<option>
							<% 
							sql = "select intDuration_ID,szDuration_Name from trefDuration order by szDuration_Name"
							Response.Write oFunc.MakeListSQL(sql,"intDuration_ID","szDuration_Name",intDuration_ID)								
							%>
						</select>
					</td>								
					<td valign=top align=center>
						<select name="intSession_Minutes" onChange="jfChanged();">
							<option>
							<% 
							dim strHour
							strMinutes = "30,60,90,120,150,180,210,240,270,300"
							strHour = ":30,1:00,1:30,2:00,2:30,3:00,3:30,4:00,4:30,5:00"
							Response.Write oFunc.MakeList(strMinutes,strHour,intSession_Minutes)								
							%>
						</select>
					</td>	
				</tr>
			</table>		
			<table style="width:100%;">				
				<tr>
					<td class=gray>
						&nbsp;Meets Every
					</td>
				</tr>
				<tr>
					<td class="svplain10" align="center" colspan="8">
						<%
						sql = "select strValue,strText from common_lists where intList_ID = 4 order by intOrder"
						response.Write oFunc.MakeCheckList(sql,"strText","strText",szDays_Meet_On,"szDays_Meet_On",6)
						%>
					</td>
				</tr>
				<tr>	
					<td class=gray >
						&nbsp;Comments
					</td>													
				</tr>			
					<td align=center>
						<textarea style="width:100%;" rows=2 name="szSchedule_Comments" wrap=virtual onKeyDown="jfMaxSize(128,this);" ><% = szSchedule_Comments%></textarea>						
					</td>	
				</tr>
			</table>			
<%
end function
%>