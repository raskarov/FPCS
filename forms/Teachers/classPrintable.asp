<%@ Language=VBScript %>
<%
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Dimention Variables, Make db Connection, print HTML header.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim intClass_Id
dim intInstructor_ID
dim sqlInstructor
dim curPay_Rate
dim sqlClass
dim sqlMaterials
dim intCount
dim strMaterials
dim strClassTitle
dim strInstructMessage
dim intStudent_id
dim intClassMatStart		'Keeps track of the number of existing Resources during edit mode
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
dim dblTotalMat				'Total of Resources (materials)		
dim strPrintTitle	
dim strFormType	

dim intPOS_Subject_ID
dim szClass_Name
dim szASD_Course_ID
dim szLocation
dim dtReg_Deadline
dim intMin_Students
dim intMax_Students
dim sGrade_Level
dim sGrade_Level2
dim dtClass_Start
dim dtClass_End
dim szStart_Time
dim szEnd_Time
dim szSchedule_Comments
dim decHours_Student
dim decHours_Planning
dim szDays_Meet_On
dim decOriginal_Student_Hrs
dim decOriginal_Planning_hrs
dim dtHrs_Last_Updated			

strPrintTitle = "Class Schedule"
strFormType = "Schedule" 

dim oFunc	'wsc object
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

' Get needed form variables   
if Request.Form.Count > 0 then
	set objRequest = Request.Form
else
	set objRequest = Request.QueryString
end if

For Each Item in objRequest	
	execute("dim " & Item)
	strObjValue = objRequest(Item)
	execute(Item & " = strObjValue")
next

session.Value("simpleTitle") = strFormType
Server.Execute(Application.Value("strWebRoot") & "Includes/simpleHeader.asp")	
	
if intClass_Id <> "" then
	call vbfPrintContract
elseif session.Contents("strClassList") <> "" then
	dim strContractList
	dim intContract_Guardian_ID
	dim intGuardian_ID 
	dim intVendor_ID 
	
	strContractList = session.Contents("strClassList")
	
	if instr(1,strContractList,"|") > 0 then
		arList = split(strContractList,"|")
		for i = 0 to ubound(arList)
			arValues = split(arList(i),",")
			intClass_id = arValues(0)
			intInstructor_ID = arValues(1)
			intInstruct_Type_id = arValues(2)
			intContract_Guardian_ID = arValues(3)
			intGuardian_ID = arValues(4)
			intVendor_ID = arValues(5)
			call vbfPrintContract
			Response.Write "<p>"
		next
	else
		arValues = split(strContractList,",")
		intClass_id = arValues(0)
		intInstructor_ID = arValues(1)
		intInstruct_Type_id = arValues(2)
		intContract_Guardian_ID = arValues(3)
		intGuardian_ID = arValues(4)
		intVendor_ID = arValues(5)
		call vbfPrintContract
	end if 
end if 

function vbfPrintContract 
	' Initualize variables from their different sources.  We do not want the form elements to be enabled
	' when we are dealing with students. Our rule is that classes can only be edited by the instructor or
	' fpcs admin. So a parent can not edit a class that effects other students. 
	if strDisabled <> "" then
		'Coming from viewClasses.asp with a defined student_id. 
		intClass_Id = intClass_Id
		strDisabled = " disabled "
	elseif intClass_Id <> "" then
		'Coming from veiw/delete class (viewClasses.asp from the teachers version)
		intClass_Id = intClass_Id
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

	if intClass_Id = "" or (intClass_Id <> "" and Session.Value("intStudent_ID") <> "" and bolInWindow = "") _
		or (intClass_ID <> "" and strThisIsACopy <> "") then
		session.Value("blnFromClassAdmin") = true
	else
		session.Value("blnFromClassAdmin") = false
	end if

	intCount = 0
	strClassTitle = "Add a Contract "
	strCalcType = "jfAddHRS();"

	' Session.Value("intStudent_ID") may not always be destroyed when coming direclty from root.
	' (This script is executed coming from ilp1.asp and default.asp)
	' If it's coming from default AND from a teacher ADD A CLASS request 
	' request("bolFromTeacher") will be defined and we can not have intStudent_ID populated

	if bolFromTeacher = "" then
		intStudent_id = Session.Value("intStudent_ID")
	else
		Session.Value("intStudent_ID") = ""
		Session.Value("studentFirstName") = ""
	end if 

	if Session.Value("strStudentName") <> "" then
		strStudentName =  Session.Value("strStudentName") 
	end if 

	' ####NOTE#### Session variables for both intInstructor_ID and intVendor_ID are used in ilpMain.asp
	' Define instructor information
	if intInstructor_ID <> "" or intInstructor_ID <> "" then
		Session.Value("intInstructor_ID") = intInstructor_ID
		strPrintTitle = "Parent/ASD Teacher Instructional Contract"
		strFormType = "Contract"
		
		set rsGetInstructor = server.CreateObject("ADODB.RECORDSET")
		rsGetInstructor.CursorLocation = 3
		sqlInstructor = "select szFirst_Name, szLast_Name, curPay_Rate from tblInstructor " & _
						"where intInstructor_ID = " & intInstructor_ID
		rsGetInstructor.Open sqlInstructor, oFunc.FPCScnn
		strTeacherName = rsGetInstructor("szFirst_Name") & " " & rsGetInstructor("szLast_Name")
		curPay_Rate = rsGetInstructor("curPay_Rate")
		rsGetInstructor.Close
		set rsGetInstructor = nothing
	else
		' Ensures Session.Value("intInstructor_ID") is erased if we are not dealing with an instructor
		Session.Value("intInstructor_ID") = ""
	end if 

	Session.Value("strTeacherName") = strTeacherName

	' Define Vendor information
	if intVendor_id <> "" or intVendor_id <> "" then
		dim intVendor_ID
		dim curCharge_Amount
		dim szChargeDesc
		dim sqlVendor
		
		set rsVendor = server.CreateObject("ADODB.RECORDSET")
		rsVendor.CursorLocation = 3
		
		Session.Value("intVendor_id") = intVendor_ID
		sqlVendor = "select v.szVendor_Name,v.curCharge_Amount,c.szDesc,v.intCharge_Type_ID " & _
			  "from tblVendors v, trefCharge_Type c " & _
			  "where v.intCharge_Type_ID = c.intCharge_Type_id " & _
			  "and v.intVendor_id = " & intVendor_ID
		rsVendor.Open sqlVendor,oFunc.FPCScnn
		
		'vbfPrint sqlVendor
		szVendor_Name =rsVendor("szVendor_Name")
		szChargeDesc = rsVendor("szDesc")
		curCharge_Amount = rsVendor("curCharge_Amount")
		intCharge_Type_ID = rsVendor("intCharge_Type_ID")
		
		Session.Value("strTeacherName") = szVendor_Name	
		rsVendor.Close
		set rsVendor = nothing
			
		strInstructMessage = szVendor_Name
		strCalcType = "jfVendorAdd();"
	end if

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'' This select block sets what parts of the html form we show 
	'' depending on the Instruction type.
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	select case intInstruct_Type_ID
		case "4" 'Contract ASD Teacher
			strInstructMessage = strTeacherName
			arRate = oFunc.InstructorCosts(intInstructor_ID)
			if isArray(arRate) then
				curInstructionRate = formatNumber(arRate(9),2)
			end if 
			
			' This next bit of code creates the html neede to allow a user to
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
				
				rsIDS.Open sqlContracts, oFunc.FPCScnn									

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
				strParams = "?intInstruct_Type_ID=" & intInstruct_Type_ID & _
							"&bolFromTeacher=True&intInstructor_ID=" & 	intInstructor_ID & _
							"&strThisIsACopy=true"						
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
		if intInstructor_ID <> "" or intInstructor_ID <> "" then
			strAddSQL = "i.intInstructor_ID,i.curPay_Rate " & _
						  "from tblClasses c, tblInstructor i " & _
						  "where c.intClass_ID = " & intClass_Id & _
						  " and c.intInstructor_ID = i.intInstructor_ID" 	
		elseif intGuardian_id <> "" or intGuardian_id <> "" then
			strAddSQL = "g.intGuardian_ID " & _
						  "from tblClasses c, tblGuardian g " & _
						  "where c.intClass_ID = " & intClass_Id & _
						  " and c.intGuardian_ID = g.intGuardian_ID" 
		elseif intVendor_id <> "" or intVendor_id <> "" then
			strAddSQL = "v.intVendor_id, szService_Desc,curContract_Amount,bolOn_Premises,szPremises," & _
						  "bolDistrict_Equip,szEquip_list,bolPay_On_Completion,szPay_details, " & _
						  "c.intCharge_Type_ID as intClass_Charge_ID, " & _
						  "curCharge_Amount, decNum_Units, curUnit_Cost " & _
						  "from tblClasses c, tblVendors v " & _
						  "where c.intClass_ID = " & intClass_Id & _
						  " and c.intVendor_id = v.intVendor_id" 
		end if 
		
		sqlClass = "select c.intPOS_Subject_ID,c.intInstructor_ID,c.szClass_Name,c.szASD_Course_ID," & _
				   "c.szLocation,c.dtReg_Deadline,c.intMin_Students,c.intMax_Students," & _
				   "c.sGrade_Level,c.sGrade_Level2,c.dtClass_Start,c.dtClass_End,c.szStart_Time,c.szEnd_Time," & _
				   "c.szSchedule_Comments,c.decHours_Student,c.decHours_Planning,c.szDays_Meet_On, " & _
				   "c.decOriginal_Student_Hrs, c.decOriginal_Planning_hrs, dtHrs_Last_Updated, " & _
				   	strAddSQL 
		rsClass.Open sqlClass, oFunc.FPCScnn		
		
		'This for loop dimentions and defines all the columns we selected in sqlClass
		'and we use the variables created here to populate the form.
		for each item in rsClass.Fields
			'execute("dim " & rsClass.Fields(intCount).Name)
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
		
		' See if this class is limited to select familes and if so get them in a comma seperated list
		' so we can auto populate them on the form
		dim sqlRestricted
		sqlRestricted = "select a.intFamily_ID, f.szFamily_Name " & _
					    "from tascClass_Family a, tblFamily f " & _
						"where a.intClass_ID = " & intClass_ID & _
						" and a.intFamily_ID = f.intFamily_ID " & _
						" order by f.szFamily_Name "  
		rsClass.Open sqlRestricted, oFunc.FPCScnn

		strFamilyValues = "no"
		if rsClass.RecordCount > 0 then		
			do while not rsClass.EOF
				strFamilyList = strFamilyList & rsClass("intFamily_ID") & ", "
				strFamilyNames = strFamilyNames & rsClass("szFamily_Name") & "<BR>"
				rsClass.MoveNext
			loop
			strFamilyList = Left(strFamilyList, len(strFamilyList) - 2)
			strFamilyNames = Left(strFamilyNames, len(strFamilyNames) - 4)
			strFamilyValues = "yes"
		end if
		
		rsClass.Close
		set rsClass = nothing
		
		'This next section breaks up date information that is stored in single columns in the 
		'database because they are displayed as individual drop downs in the HTML form.
		'So we need the individual values to populate the drop downs.
		dim month
		dim day
		dim year
		dim monthStart
		dim dayStart
		dim yearStart
		dim dayEnd
		dim monthEnd
		dim yearEnd
		dim hourStart
		dim minuteStart
		dim amPmStart
		dim hourEnd
		dim minuteEnd
		dim amPmEnd
		 
		month = datePart("m",dtReg_Deadline)
		day = datePart("d",dtReg_Deadline)
		year = datePart("yyyy",dtReg_Deadline)
		 
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
		
		'Now we make another query and get all of the materials needed for the class
		set rsMaterials = server.CreateObject("ADODB.RECORDSET")
		rsMaterials.CursorLocation = 3
		sqlMaterials = "select intClass_Materials_ID,szMaterial_Name,szMaterial_Desc," & _
					   "intMaterial_QTY,curUnit_Price " & _
					   "from tblClass_Materials where intClass_id = " & intClass_ID
		rsMaterials.Open sqlMaterials,oFunc.FPCScnn
		strMaterials = strMaterials & "<table><tr><Td class=gray valign=top> Item Name</td>" & chr(13)
		strMaterials = strMaterials & "<Td class=gray valign=top> Description</td>" & chr(13)
		strMaterials = strMaterials & "<Td class=gray valign=top> Qty</td>" & chr(13)
		strMaterials = strMaterials & "<td class=gray> Unit Price</td>" & chr(13)
		strMaterials = strMaterials & "<td class=gray valign=top> Total</td></tr>" & chr(13)	
		intCount = 0
		do while not rsMaterials.EOF	
			dblCost = cdbl(rsMaterials("intMaterial_QTY")) * CDBL(rsMaterials("curUnit_Price"))
			dblTotalMat = dblTotalMat + dblCost
			strMaterials = strMaterials & "<Tr><td valign=top class=svplain10>" & rsMaterials("szMaterial_Name") & "</td>" & chr(13)
			strMaterials = strMaterials & "<td valign=top class=svplain10>" & rsMaterials("szMaterial_Desc") & "</td>" & chr(13)
			strMaterials = strMaterials & "<td valign=top class=svplain10 align=right>" & rsMaterials("intMaterial_QTY") & "</td>" & chr(13)	
			strMaterials = strMaterials & "<td valign=top class=svplain10 align=right>$" & formatNumber(rsMaterials("curUnit_Price"),2) & "</td>" & chr(13)							  						  								  	
			strMaterials = strMaterials & "<td valign=top class=svplain10 align=right>$" & formatNumber(dblCost,2) & "</td></tr>" & chr(13)	
			rsMaterials.MoveNext	
			intCount = intCount + 1								  									  
		loop
		strMaterials = strMaterials & "</tr></table>"
		strClassTitle = "Scheduled Class"	
		intClassMatStart = rsMaterials.RecordCount
		rsMaterials.Close
		set rsMaterials = nothing			
	end if 	
	Session.Value("strTitle") = strFormType
	Session.Value("strLastUpdate") = "22 Feb 2002"

	' This recordset is used many times to get Class Information 
	set rsInfo = server.CreateObject("ADODB.recordset")
	rsInfo.CursorLocation = 3	
	%>
	<link rel="stylesheet" href="<% =  Application.Value("strWebRoot") %>CSS/printStyle.css">
	<table width=100%>
		<tr>
			<td align=left>
				<img src="<% = Application("strImageRoot")%>fpcsLogo.gif">
			</td>
			<td align=right class=svplain10 width=100%>
				<% = Application.Contents("SchoolAddress") %>
			</td>
		</tr>
		<tr class=yellowHeader>	
			<Td colspan=2>
				<table align=right ID="Table1"><tr><td align=right><font face=arial size=2 color=white><% = date()%></font></td></tr></table>
				&nbsp;<b><% = strPrintTitle %></b>											
			</td>					
		</tr>
		<tr>
			<td colspan=2>
			<table>		
					<tr>	
						<Td colspan=32>
							<font class=svplain11>
								<b><i>Parties Involved:</I></B> 
							</font>
						</td>
					</tr>
					<tr>
						<% if intInstructor_ID & "" <> "" then%>
						<td class=gray >
							&nbsp;ASD Teacher
						</td>
						<% end if %>
						<td class=gray>
							&nbsp;Parent
						</td>	
						<td class=gray>
							&nbsp;Student
						</td>																
					</tr>
					<tr>
						<% if intInstructor_ID & "" <> "" then%>
						<td class=svplain10>
								<% = strInstructMessage %>
						</td>	
						<% end if %>
						<td align=center class=svplain10>
							<%				
								if intStudent_ID <> "" then											
									dim sqlGaurdian
									sqlGaurdian = "Select g.intGuardian_ID,g.szLast_Name + ',' + g.szFirst_Name as Name " & _
													 "from tblGuardian g, tblILP i " & _
													 "where g.intGuardian_ID = i.intContract_Guardian_ID " & _
													 " and i.intStudent_ID = " & intStudent_ID &_
													 " and i.intClass_ID = " & intClass_ID & _
													 " order by szLast_Name"	
									rsInfo.Open sqlGaurdian, oFunc.FPCScnn
									if rsInfo.RecordCount > 0 then
										Response.Write rsInfo("name")	
									end if		
									rsInfo.Close
								end if								 
							%>	
						</td>	
						<td align=center class=svplain10>
							<% = strStudentName%>
						</td>									
					</tr>
				</table>
				<table ID="Table3">
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
						<td class=gray>
								&nbsp;Name of Class
						</td>
						<td class=gray>
								&nbsp;ASD Course ID
						</td>
						<td class=gray>
								&nbsp;Course Category
						</td>
						<!--<td class=gray>
							&nbsp;Subject
						</td>-->
						<td class=gray>
							&nbsp;Location
						</td>											
					</tr>
					<tr>
						<td class=svplain10>
							<% = szClass_Name%>
						</td>
						<td class=svplain10>
							<% = szASD_Course_ID%>
						</td>
						<td class=svplain10 align=center>
							
							<%
								sql = "select intPOS_Subject_ID, szSubject_Name from trefPOS_Subjects where intPOS_Subject_ID = " & intPOS_Subject_ID
								rsInfo.Open sql, oFunc.FPCScnn
								if rsInfo.RecordCount > 0 then
									Response.Write rsInfo("szSubject_Name")	
								end if 
								rsInfo.Close
							%>
						</td>
						<!--
						<td>
							<input type=text name="szSubject" value="<% = szSubject%>" maxlength=64 size=20 onChange="jfChanged();">
						</td> -->
						<td class=svplain10>
							<% = szLocation%>
						</td>									
					</tr>
				</table>
				<table ID="Table4">
					<tr>
						<td class=gray >
							&nbsp;Registration Deadline
						</td>
						<td class=gray>
							&nbsp;Min # Students
						</td>	
						<td class=gray>
							&nbsp;Max # Students
						</td>	
						<td class=gray>
							&nbsp;Grade&nbsp;
						</td>		
						<td class=gray>
							&nbsp;to Grade&nbsp;
						</td>																
					</tr>
					<tr>
						<td class=svplain10>
								<% = month & "/" & day & "/" & year %>
						</td>	
						<td align=center class=svplain10>
						<% = intMin_Students%>
						</td>	
						<td align=center class=svplain10>
							<% = intMax_Students%>
						</td>	
						<td align=center class=svplain10>
							<% = sGrade_Level%>
						</td>			
						<td align=center class=svplain10>
							<%= sGrade_Level2 %>
						</td>								
					</tr>
				</table>
				<table ID="Table5">				
					<tr>
						<td class=gray >
							&nbsp;Class Start Date
						</td>
						<td class=gray >
							&nbsp;Class End Date
						</td>	
						<td class=gray align=center>
							&nbsp;Meets Every
						</td>																
					</tr>
					<tr>
						<td valign=top class=svplain10>
							<% = monthStart & "/" & dayStart & "/" & yearStart %>
						</td>				
						<td valign=top class=svplain10>
							<% = monthEnd & "/" & dayEnd & "/" & yearEnd %>	
						</td>				
						<td align=center class=svplain10>
							<% 							
								dim sqlDays
								sqlDays = "select strText from common_lists where intList_ID = 4 and strValue='" & szDays_Meet_On & "'"
								rsInfo.Open sqlDays,oFUnc.FPCScnn
								if rsInfo.RecordCount > 0 then
									do while not rsInfo.EOF
										Response.Write 	rsInfo("strText") 
										rsInfo.MoveNext
										if not rsInfo.EOF then
											response.Write ", "
										end if 
									loop
								end if			
								rsInfo.Close		
							%>
						</td>
					</tr>
				</table>		
				<table ID="Table6">				
					<tr>
						<td class=gray >
								&nbsp;Class Start Time
						</td>
						<td class=gray >
							&nbsp;Class End Time
						</td>		
						<td class=gray >
							&nbsp;Schedule Comments
						</td>													
					</tr>
					<tr>
						<td valign=top class=svplain10 align=center>
							<% = hourStart%>:<% = minuteStart %> <% =amPmStart %>
						</td>	
						<td valign=top class=svplain10 align=center>
							<%= hourEnd %>:<% = minuteEnd %> <% = amPmEnd %>	
						</td>	
						<td align=center class=svplain10>
							<% = szSchedule_Comments%>					
						</td>	
					</tr>
				</table>
				<% if strFamilyNames <> "" then %>
				<table ID="Table7">
					<tr>	
						<Td colspan=2>
							<font class=svplain11>
								<b><i><% = strFormType %> Restrictions</I></B> 
							</font>
						</td>
					</tr>
					<tr>
						
						<td class=gray valign=top>						
							<b>Restricted to the following families:</b><BR>
							<% = strFamilyNames%>						
						</td>
					</tr>				
				</table>	
				<BR>
				<% end if %>	
				
				<% if dblTotalMat <> "" then %>	
				<table ID="Table8">
					<tr>	
						<Td colspan=2>
							<font class=svplain11>
								<b><i>Class Costs</I></B> 
							</font>
							<font class=svplain11>
								 (Resources Required)
							</font>
						</td>
					</tr>
					<tr>
						<td class=gray>
							&nbsp;<i>Itemized Costs for an <u>Individual</u> Student Only.</i>. 		
						</td>										
					</tr>
					<tr>
						<td class=svplain10>
							<% = strMaterials %>
						</td>
					</tr>				
				</table>
				<BR>
				<% end if %>	
							
				<table ID="Table9">				
					<%
						if intVendor_ID <> "" then
							call vbfVendorFields
						elseif intInstructor_ID & "" <> "" then
							call vbfClassDetailsForASD
						end if 
					%>
				</table>
			</td>
		</tr>
	</table>
	<% if intInstructor_ID & "" <>  "" then%>
	<table>
		<tr>
			<td colspan=2 class=svplain10>
				Signatures below indicate acceptance of all applicable<br> sections of the 
				Educational Plan & Annual Budget.
			</td>
		</tr>
		<tr>
			<td>
			<pre>

				
	________________________________________  ____________
	Student Signature                          Date		


	________________________________________  ____________
	Parent Teacher Signature                   Date


	________________________________________  ____________
	Partnering ASD Teacher Signature           Date


	________________________________________  ____________
	FPCS Administrator                         Date
				 </pre>	
			</td>
		</tr>
	</table>

	<% 
	end if
	set rsInfo = nothing			
end function

%>
	<script language=javascript>
		if (window.print){
	      window.print()
	    }
	    else {
	      alert("Mac users: please press Apple-P to print this form.\nWindows users: Please press ctrl-P to print this form.")
		}
	</script>
<%
call oFunc.CloseCN	
Server.Execute(Application.Value("strWebRoot") & "Includes/footer.asp")
 
 function vbfClassDetailsForASD
 %>
				<tr>	
					<Td colspan=2>
						<font class=svplain11>
							<b><i>Class Costs</I></B> 
						</font>
						<font class=svplain11>
							 (Teachers Time)
						</font>
					</td>
				</tr>
				<tr>
					<td class=gray align=right>
						<% = formatNumber(decHours_Student,1) %>
					</td>		
					<td class=gray>
						&nbsp;Number of teacher hours with student.
					</td>							
				</tr>		
				<tr>
					<td class=gray align=right>
						<% = formatNumber(decHours_Planning,1) %>
					</td>		
					<td class=gray>
						&nbsp;Number of hours for teacher planning.
					</td>							
				</tr>	
				<tr>
					<td class=gray align=right>
						<% = formatNumber((CDBL(decHours_Student) + cdbl(decHours_Planning)),1) %>
					</td>		
					<td class=gray>
						&nbsp;<B>Total teacher hours.</b>
					</td>							
				</tr>	
 				<tr>
					<td class=gray align=right>
						<% = formatNumber((CDBL(decHours_Student) + cdbl(decHours_Planning))/intMax_Students,1)%>
					</td>		
					<td class=gray>
						&nbsp;Minimum number of hours to be charged to each student.
					</td>							
				</tr>	
				<tr>
					<td class=gray align=right>
						<% = formatNumber((CDBL(decHours_Student) + cdbl(decHours_Planning))/intMin_Students,1)%>
					</td>		
					<td class=gray>
						&nbsp;Maximum number of hours to be charged to each student.
					</td>							
				</tr>	
				<tr>
					<td class=gray align=right>
						$<% = curInstructionRate %>
					</td>		
					<td class=gray>
						&nbsp;Teachers hourly rate.
					</td>							
				</tr>
				<tr>
					<td class=gray align=right>
						$<% = formatNumber(((CDBL(decHours_Student) + cdbl(decHours_Planning))/intMax_Students)*curInstructionRate,2) %>
					</td>		
					<td class=gray>
						&nbsp;Minimum total teacher cost per student.
					</td>							
				</tr>	
				<tr>
					<td class=gray align=right>
						$<% = formatNumber(((CDBL(decHours_Student) + cdbl(decHours_Planning))/cdbl(intMin_Students))*cdbl(curInstructionRate),2) %>
					</td>		
					<td class=gray>
						&nbsp;Maximum total teacher cost per student.
					</td>							
				</tr>
				<tr>
					<td class=gray align=right>
						$<% = formatNumber(dblTotalMat,2) %>
					</td>		
					<td class=gray>
						&nbsp;Total miscellaneous costs per student.
					</td>							
				</tr>		
				<tr>
					<td class=gray align=right>
						$<% = formatNumber(dblTotalMat + (((CDBL(decHours_Student) + cdbl(decHours_Planning))/intMax_Students)*curInstructionRate),2)%>
					</td>		
					<td class=gray>
						&nbsp;<B>Minimum total deduction per student account.</b>
					</td>							
				</tr>	
				<tr>
					<td class=gray align=right>
						$<% = formatNumber(dblTotalMat + (((CDBL(decHours_Student) + cdbl(decHours_Planning))/intMin_Students)*curInstructionRate),2)%>
					</td>		
					<td class=gray>
						&nbsp;<B>Maximum total deduction per student account.</b>
					</td>							
				</tr>		
 <%
 end function
 
 function vbfVendorFields
	dim sqlChargeID
	dim strChargeList	' Option List of Charges
	
	if intClass_Charge_ID <> "" then		
		' This overrides the value stored in the vendor profile by the value
		' that has been add to the class record.
		intCharge_Type_ID = intClass_Charge_ID
	end if 
	
	if curUnit_Cost <> "" then
		' This overrides the value stored in the vendor profile by the value
		' that has been add to the class record.
		curCharge_Amount = cdbl(curUnit_Cost)
	end if 
	
	sqlChargeID = "select szDesc from trefCharge_Type where intCharge_Type_ID = " & intCharge_Type_ID
	rsInfo.Open sqlChargeID,oFunc.FPCScnn
	if rsInfo.RecordCount > 0 then
		strChargeList = rsInfo("szDesc")
	end if
	rsInfo.Close
	
 %>				
				<tr>	
					<Td colspan=2>
						<font class=svplain11>
							<b><i>Vendor Class Costs</I></B> 
						</font>
					</td>
				</tr>
				<tr>
					<td class=gray align=right>
						<% = decNum_Units %>
					<td class=gray>
						&nbsp;<% = strChargeList %> at
					</td>							
				</tr>		
				<tr>
					<td class=gray align=right>
						$<% = formatNumber(curCharge_Amount,2) %>
					</td>		
					<td class=gray onChange="jfChanged();">
						&nbsp;Cost per unit.
					</td>							
				</tr>					
				<tr>
					<td class=gray align=right>
						$<% = formatNumber(cdbl(decNum_Units) * cdbl(curCharge_Amount),2) %>
					</td>		
					<td class=gray>
						&nbsp;<B>Contract Amount.</b>
					</td>							
				</tr>	
			 </table>
			 <br>
			 <table ID="Table10">
				<tr>	
					<Td colspan=2>
						<font class=svplain11>
							<b><i>Vendor Contract Information</I></B> 
						</font>
					</td>
				</tr>
				<tr>
					<td class=gray colspan=2>
						Description of Service Vendor is Providing				
					</td>
				</tr>
				<tr>
					<td colspan=2 class=svplain10>
						<% = szService_desc %>
					</td>
				</tr>
				<tr>
					<td class=gray>
						&nbsp;Is Service to be provided <BR>
						&nbsp;on District premises? 
						<input disabled type=checkbox <% if bolOn_Premises then Response.Write " checked " %> name="bolOn_Premises" value="TRUE" onChange="jfChanged();" ID=Checkbox <% if bolon_premises then response.write " checked " %>1><b>Yes</b>&nbsp;
					</td>
					<td class=gray>
						&nbsp;Is Service to be provided <BR>
						&nbsp;using District equipment or supplies? 
						<input disabled type=checkbox <% if bolDistrict_Equip then Response.Write " checked " %> name="bolDistrict_Equip" value="TRUE" onChange="jfChanged();" ID=Checkbox <% if boldistrict_equip then response.write " checked " %>1><b>Yes</b>&nbsp;
					</td>										
				</tr>	
				<tr>
					<td class=svplain10>
						If yes indicate where?					
					</td>
					<td class=svplain10>
						If yes indicate which equipment and supplies.				
					</td>
				</tr>
				<tr>
					<td class=svplain10>
						<% = szPremises%>
					</td>
					<td align=center class=svplain10>
						<% = szEquip_List %>
					</td>
				</tr>
				<tr>
					<td class=gray colspan=2>
						&nbsp;Will the District pay Contractor for Service updon satisfactory completion
						<br>&nbsp;and acceptance of ALL work required uder this contract? 
						<input disabled type=checkbox <% if bolPay_On_Completion then Response.Write " checked " %> name="bolPay_On_Completion" value="TRUE" ID=Checkbox <% if bolpay_on_completion then response.write " checked " %>1><b>Yes</b>&nbsp;
					</td>
				</tr>
				<tr>
					<td class=svplain10 colspan=2>
						If NO indicate when (partial) payment(s) to be made. 				
					</td>
				</tr>
				<tr>
					<td colspan=2 class=svplain10>
						<% = szPay_Details %>
					</td>
				</tr>
 <%
 end function
%>