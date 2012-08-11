<%@  language="VBScript" %>
<%
response.Buffer = false
server.ScriptTimeout = 900
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Dimention Variables, Make db Connection, print HTML header.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim intClass_Id
dim intInstructor_ID
dim sqlInstructor
dim curPay_Rate
dim sqlClass
dim intCount
dim strClassTitle
dim strInstructMessage
dim intStudent_id	
dim intClassMatStart		'Keeps track of the number of existing Resources during edit mode
dim strAddSQL				'Dynamic peice of sql defined depending on instructor,guardian or vendor
dim curInstructionRate		'Holds the hourly rate of instruction including taxes and benefits
dim strDisabled				'This string is used in form elements to disable them when we are adding ILP's
dim strFamilyList			'Contains list of families that this class is restricted to 	
dim strFamilyValues			'This is used to keep track of whether the families pulldown is populated.
							'If in edit mode it was populated and the admin decided to make it open
							'to everyone we needed some way of nowing that all family restrictions 
							'for this class dhould be deleted and not replaced with others.	
dim strPrintTitle	
dim strFormType	

dim intPOS_Subject_ID
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
dim rsInfo
dim monthReg

dim dblClassCharge 
dim dblClassBudget 
dim dblTotalCharge 
dim dblTotalBudget 

dim dayReg
dim yearReg
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

dim strStudentName
dim strTeacherName
dim szClass_Name
dim sql
dim strMaterials			'contains all of the materials from classAdmin.asp
dim intLength				'used to take off ending comma in the constructed strMaterials
dim monthEnroll				'dtStudent_Enrollment will be broken into these 3 variables
dim dayEnroll
dim yearEnroll
dim intILP_ID
dim intColor
dim strColor
dim HideSigs 

strFormType = "Class " 

dim oFunc	'wsc object
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()	
intStudent_ID = request("intStudent_ID")
intClass_ID = request("intClass_ID")
intILP_ID = request("intILP_ID")
if intStudent_ID <> "" then
	strStudentName = oFunc.StudentInfo(intStudent_ID,3)
end if
intInstructor_ID = request("intInstructor_ID")


session.Value("simpleTitle") = "Student Packet"
Server.Execute(Application.Value("strWebRoot") & "Includes/simpleHeader.asp")
%>
<link rel="stylesheet" href="<% =  Application.Value("strWebRoot") %>CSS/printStyle.css">
<%

' Get list of all Contracts and ILPS 
if intStudent_ID <> "" then
	sql = "SELECT	i.intClass_ID, i.intILP_ID,  " & _ 
			"CASE ISNULL(c.intInstructor_ID, 1) WHEN 1 THEN 0 ELSE 1 END AS isContract, " & _
			"CASE c.intPOS_SUBJECT_ID WHEN 22 THEN 0 ELSE 1 END AS isSponsor " & _ 
			"FROM	tblILP i INNER JOIN " & _ 
			"	tblClasses c ON i.intClass_ID = c.intClass_ID " & _ 
			"WHERE (i.sintSchool_Year = " & session.Contents("intSchool_Year") & ") " & _
			" AND (i.intStudent_ID = " & intStudent_ID & ") " & _ 
			"ORDER BY isSponsor, c.szClass_Name "
				
elseif intInstructor_ID <> "" then
	sql = "select c.intClass_ID,gi.intILP_ID, '01' as isContract " & _
			  "from tblInstructor i,tblClasses c left outer join tblILP_Generic gi " & _
			  " ON c.intClass_id = gi.intClass_ID " & _
			  "where i.intInstructor_ID = c.intInstructor_ID and " & _
		      "i.intInstructor_ID = " &  intInstructor_ID & _ 
		      " and ( c.intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _
		      " order by c.szClass_Name " 
		
end if

' Get Recordset
if ucase(intClass_ID) = "ALL" or ucase(intILP_ID) = "ALL"  or ucase(request("strAction")) = "A" then
	set rsAll = server.CreateObject("ADODB.RECORDSET")
	rsAll.CursorLocation = 3
	rsAll.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
end if

' Fusebox logic
select case ucase(request("strAction"))
	case "C"	
		call vbsContract
	case "I"
		call vbsILP
	case "IP"
		call Philiosophy(intStudent_ID)
	case "G"
		call vbsGoodsServices
	case "S" 
		call vbfStudentPacket
	case "T"
		call vbForms("Testing")
	case "P"
		call vbForms("")
	case "A"
		call vbfStudentPacket
		response.Write "<p></p>"
		call Philiosophy(intStudent_ID)
		response.Write "<p></p>"
		rsAll.MoveFirst
		
		do while not rsAll.EOF
			'intClass_ID = "ALL"
			'intILP_ID = "ALL"
			if isNumeric(rsAll("intClass_ID")) then
				call vbfPrintContract(rsAll("intClass_ID"))	
				response.Write "<p></p>"
			end if
			
			if isNumeric(rsAll("intILP_ID")) then
				call vbfPrintILP(rsAll("intILP_ID"))
				response.Write "<p></p>"
			end if
			rsAll.MoveNext
		loop		
		dblClassCharge = 0 
		dblClassBudget = 0 
		dblTotalCharge = 0 
		dblTotalBudget = 0 
		intColor = 0 
	case "V"
		call VendorServiceReport
	case "AP"


		strPacketList = request.QueryString("strPacketList")
		strPacketList = Left(strPacketList,len(strPacketList)-1)
		strPacketList = Right(strPacketList,len(strPacketList)-1)
		strPacketList = replace(strPacketList,"s","")


		dim LastStudent
		LastStudent = 0
		
		sql = "SELECT	s.szLast_Name, s.szFirst_Name,i.intStudent_ID,i.intClass_ID, i.intILP_ID,  " & _ 
				"CASE ISNULL(c.intInstructor_ID, 1) WHEN 1 THEN 0 ELSE 1 END AS isContract, " & _
				"CASE c.intPOS_SUBJECT_ID WHEN 22 THEN 0 ELSE 1 END AS isSponsor " & _ 
				"FROM	tblILP i INNER JOIN " & _ 
				"	tblClasses c ON i.intClass_ID = c.intClass_ID INNER JOIN " & _ 
				"	tblStudent s ON i.intStudent_ID = s.intStudent_ID " & _
				"WHERE (i.sintSchool_Year = " & session.Contents("intSchool_Year") & ") " & _
				" AND (i.intStudent_ID in (" & strPacketList & ")) " & _ 
				"ORDER BY s.szLast_Name, s.szFirst_Name,isSponsor, c.szClass_Name "
				
		set rsPacket = server.CreateObject("ADODB.RECORDSET")
		rsPacket.CursorLocation = 3
		rsPacket.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
		
		do while not rsPacket.EOF
			if rsPacket("intStudent_ID") <> LastStudent then
				strStudentName = rsPacket("szFirst_Name") & " " & rsPacket("szLast_Name")
				intStudent_ID = rsPacket("intStudent_ID")
				call vbfStudentPacket
				response.Write "<p></p>"
				call Philiosophy(rsPacket("intStudent_ID"))
				response.Write "<p></p>"
				LastStudent = rsPacket("intStudent_ID")				
			end if
			
			if isNumeric(rsPacket("intClass_ID")) then
				call vbfPrintContract(rsPacket("intClass_ID"))	
				response.Write "<p></p>"
			end if
			
			if isNumeric(rsPacket("intILP_ID")) then
				call vbfPrintILP(rsPacket("intILP_ID"))
				response.Write "<p></p>"
			end if
			
			rsPacket.movenext
		loop
		
		rsPacket.close
		set rsPacket = nothing
		
		update = "update tblEnroll_Info set dtPacket_Printed = CURRENT_TIMESTAMP " & _
				 " WHERE intStudent_ID in (" & strPacketList & ") and " & _
				 "		 sintSchool_Year = " & session.Contents("intSchool_Year")
		oFunc.ExecuteCN(update)
		
	%>
<script language="javascript">
    window.opener.location.reload();
    window.focus();
	</script>
<%
	case else
		call vbsContract
		response.Write "<p></p>"
		call vbsILP		
		call vbsGoodsServices
end select

sub vbsContract
	if isNumeric(intClass_ID) then
		call vbfPrintContract(intClass_ID)	
	elseif ucase(intClass_ID) = "ALL" then
		bolAddBreak = false
		do while not rsAll.EOF			
			if rsAll("isContract") = "1"  then
				if bolAddBreak then response.Write "<p></p>"
				call vbfPrintContract(rsAll("intClass_ID"))	
				if ucase(request("strAction")) = "A" then
					intClass_ID = rsAll("intClass_ID")
					vbsGoodsServices
					response.Write "<p></p>"
				elseif ucase(request("strAction")) <> "A" then
					rsAll.MoveNext
					if not rsAll.EOF then
						if rsAll("isContract") = "1"  then
							response.Write "<p></p>"
							bolAddBreak = false
						else
							bolAddBreak = true
						end if
					end if 
					rsAll.MovePrevious
				end if
			end if
			rsAll.MoveNext
		loop
	end if
end sub

sub vbsILP
	if isNumeric(intILP_ID) then
		call vbfPrintILP(intILP_ID)
	elseif ucase(intILP_ID) = "ALL" then
		do while not rsAll.EOF
			if isNumeric(rsAll("intILP_ID")) then
				call vbfPrintILP(rsAll("intILP_ID"))
			end if
			rsAll.MoveNext
			if not rsAll.EOF then
				response.Write "<p></p>"
			elseif ucase(request("strAction")) = "A" then
				response.Write "<p></p>"
			end if
		loop
	end if
end sub

if isObject(rsAll) then
	rsAll.Close
	set rsAll = nothing
end if

sub vbsGoodsServices
	if isNumeric(intStudent_ID) or isNumeric(intClass_ID) then
		call vbfGoodsServices()
	end if
end sub

function vbfPrintContract(intClass_ID)

	intCount = 0
	strClassTitle = "Class Contract "

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'' This next section will fill the form in with class info 
	'' if we have a valid class id passed to this script.
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'sqlClass gets most of the class information
		set rsClass = server.CreateObject("ADODB.RECORDSET")	
		rsClass.CursorLocation = 3 
				
		sqlClass = "SELECT c.intPOS_Subject_ID, c.intInstructor_ID, c.szClass_Name, c.szASD_Course_ID,   " & _
					" c.szLocation, c.dtReg_Deadline, c.intMin_Students, c.intMax_Students,  " & _
					" c.sGrade_Level, c.sGrade_Level2, c.dtClass_Start, c.dtClass_End, c.szStart_Time,  " & _
					" c.szEnd_Time, c.szSchedule_Comments, c.decHours_Student,  " & _
					" c.decHours_Planning, c.szDays_Meet_On, c.decOriginal_Student_Hrs, c.decOriginal_Planning_Hrs, " & _
					" c.dtHrs_Last_Updated, i.curPay_Rate,  i.szFIRST_NAME + ' ' + i.szLAST_NAME AS iName,  " & _
					"g.szFIRST_NAME + ' ' + g.szLAST_NAME AS gName, d.szDuration_Name, c.intSession_Minutes,  pos.szSubject_Name " & _
					"FROM tblClasses c LEFT OUTER JOIN " & _
					" trefPOS_Subjects pos ON c.intPOS_Subject_ID = pos.intPOS_Subject_ID LEFT OUTER JOIN " & _
					" trefDuration d ON c.intDuration_ID = d.intDuration_ID LEFT OUTER JOIN " & _
					" tblGUARDIAN g ON c.intGuardian_ID = g.intGUARDIAN_ID LEFT OUTER JOIN " & _
					" tblINSTRUCTOR i ON c.intInstructor_ID = i.intINSTRUCTOR_ID " & _ 
					"where c.intClass_ID = " & intClass_Id 
		rsClass.Open sqlClass, Application("cnnFPCS")'oFunc.FPCScnn		
		
		'This for loop dimentions and defines all the columns we selected in sqlClass
		'and we use the variables created here to populate the form.

		for each item in rsClass.Fields
			'execute("dim " & rsClass.Fields(intCount).Name)
			execute(rsClass.Fields(intCount).Name & " = item")		
			intCount = intCount + 1
		next 
	
		rsClass.Close
		
		' Get instructors hourly rate
		' JD: EDIT get the flat rate, instead:
		if iName <> "" then
			if Session.Contents("intSchool_Year") < 2012 then
			    arrate = ofunc.instructorcosts(intinstructor_id)
			    if isarray(arrate) then
			    	curinstructionrate = formatnumber(arrate(9),2)
			    end if
			else
			
			    sql4 ="select intFlat_Inst_Id, flatRate from tblInstructor_Flat_Rate where intSchool_year = " & session.Contents("intSchool_Year")
                set rs4 = server.CreateObject("ADODB.RECORDSET")
                rs4.CursorLocation = 3
                rs4.Open sql4, Application("cnnFPCS")'oFunc.FPCScnn
                curinstructionrate = formatNumber(rs4("flatRate"), 2)
                rs4.Close()
            end if
 
		end if 
			
		' See if this class is limited to select familes and if so get them in a comma seperated list
		' so we can auto populate them on the form
		dim sqlRestricted
		sqlRestricted = "select a.intFamily_ID, f.szFamily_Name " & _
					    "from tascClass_Family a, tblFamily f " & _
						"where a.intClass_ID = " & intClass_ID & _
						" and a.intFamily_ID = f.intFamily_ID " & _
						" order by f.szFamily_Name "  
		rsClass.Open sqlRestricted, Application("cnnFPCS")'oFunc.FPCScnn

		strFamilyValues = "no"
		if rsClass.RecordCount > 0 then		
			do while not rsClass.EOF
				strFamilyList = strFamilyList & rsClass("intFamily_ID") & ", "
				strFamilyNames = strFamilyNames & rsClass("szFamily_Name") & ", "
				rsClass.MoveNext
			loop
			strFamilyList = Left(strFamilyList, len(strFamilyList) - 2)
			strFamilyNames = Left(strFamilyNames, len(strFamilyNames) - 2)
			strFamilyValues = "yes"
		end if
		
		rsClass.Close
		set rsClass = nothing
		
		'This next section breaks up date information that is stored in single columns in the 
		'database because they are displayed as individual drop downs in the HTML form.
		'So we need the individual values to populate the drop downs.
		
		
		if dtReg_Deadline <> "" then
			monthReg = datePart("m",dtReg_Deadline)
			dayReg = datePart("d",dtReg_Deadline)
			yearReg = datePart("yyyy",dtReg_Deadline)
		end if
		
		if dtClass_Start <> "" then
			monthStart = datePart("m",dtClass_Start)
			dayStart = datePart("d",dtClass_Start)
			yearStart = datePart("yyyy",dtClass_Start)
		end if 
		
		if dtClass_End <> "" then
			monthEnd = datePart("m",dtClass_End)
			dayEnd = datePart("d",dtClass_End)
			yearEnd = datePart("yyyy",dtClass_End)
		end if
				 
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
				
	Session.Value("strTitle") = strFormType
	Session.Value("strLastUpdate") = "22 Feb 2002"

	if intInstructor_ID & "" <> "" then
		strPrintTitle = "Parent/Teacher Contract"
	else
		strPrintTitle = "Class Schedule"
	end if 
	
	' This recordset is used many times to get Class Information 
	set rsInfo = server.CreateObject("ADODB.recordset")
	rsInfo.CursorLocation = 3		
	%>
<table width="100%" id="Table1">
    <% = vbfFormHeader(strPrintTitle & " #" & intClass_ID) %></b>
    <tr>
        <td colspan="2">
            <table id="Table3" style="width: 100%;" cellpadding="3">
                <tr>
                    <td colspan="32">
                        <font class="svplain11"><b><i>Parties Involved:</i></b> </font>
                    </td>
                </tr>
                <tr>
                    <% if iName <> "" then %>
                    <td class="gray" nowrap>
                        ASD Teacher
                    </td>
                    <% end if %>
                    <td class="gray" nowrap>
                        Parent
                    </td>
                    <td class="gray" style="width: 100%;">
                        Student
                    </td>
                </tr>
                <tr>
                    <% if iName <> "" then %>
                    <td class="svplain10" nowrap>
                        <% = iName %>
                    </td>
                    <% end if %>
                    <td align="center" class="svplain10" nowrap>
                        <%		
								if gName <> "" then
									response.Write gName
								else		
									if intStudent_ID <> "" then											
										dim sqlGaurdian
										sqlGaurdian = "Select g.intGuardian_ID,g.szLast_Name + ',' + g.szFirst_Name as Name " & _
														"from tblGuardian g, tblILP i " & _
														"where g.intGuardian_ID = i.intContract_Guardian_ID " & _
														" and i.intStudent_ID = " & intStudent_ID &_
														" and i.intClass_ID = " & intClass_ID & _
														" order by szLast_Name"	
										rsInfo.Open sqlGaurdian, Application("cnnFPCS")'oFunc.FPCScnn
										if rsInfo.RecordCount > 0 then
											Response.Write rsInfo("name")	
										end if		
										rsInfo.Close
									end if	
								end if							 
							%>
                    </td>
                    <td class="svplain10" style="width: 100%;">
                        <% = strStudentName%>
                    </td>
                </tr>
            </table>
            <table id="Table4" style="width: 100%;">
                <tr>
                    <td colspan="2" nowrap>
                        <font class="svplain11"><b><i>
                            <% = strFormType %>
                            Information</i></b> </font><font class="svplain"></font>
                    </td>
                </tr>
                <tr>
                    <td class="gray" nowrap>
                        &nbsp;Name of Class
                    </td>
                    <td class="gray" nowrap>
                        &nbsp;ASD Course ID
                    </td>
                    <td class="gray" nowrap>
                        &nbsp;Course Category
                    </td>
                    <!--<td class=gray>
							&nbsp;Subject
						</td>-->
                    <td class="gray">
                        &nbsp;Location
                    </td>
                </tr>
                <tr>
                    <td class="svplain10" nowrap>
                        <% = szClass_Name%>
                    </td>
                    <td class="svplain10">
                        <% = szASD_Course_ID%>
                    </td>
                    <td class="svplain10" align="center">
                        <% = szSubject_Name %>
                    </td>
                    <!--
						<td>
							<input type=text name="szSubject" value="<% = szSubject%>" maxlength=64 size=20 onChange="jfChanged();">
						</td> -->
                    <td class="svplain10" style="width: 100%;">
                        <% = szLocation%>
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
								
				if strFamilyNames <> "" then 
				%>
            <table id="Table7" style="width: 100%;">
                <tr>
                    <td colspan="2">
                        <font class="svplain11"><b><i>
                            <% = strFormType %>
                            Restrictions</i></b> </font>
                    </td>
                </tr>
                <tr>
                    <td class="gray" valign="top" style="width: 100%;">
                        <b>Restricted to the following families:</b><br>
                        <% = strFamilyNames%>
                    </td>
                </tr>
            </table>
            <br>
            <% end if 									
				
				dblGrandTotal = 0 
				if intClass_ID <> "" AND intInstructor_ID <> "" then	
						'sqlItems = "SELECT (CASE ci.intItem_ID WHEN 3 THEN " & _
						'			" (SELECT vs.szVend_Service_Name " & _
						'			" FROM trefVendor_Services vs, tblClass_Attrib ca2 " & _
						'			" WHERE ci.intClass_Item_ID = ca2.intClass_Item_ID AND ca2.intItem_Attrib_Id = 26 AND vs.intVend_Service_ID = ca2.szValue)  " & _
						'			" ELSE ca.szValue END) AS Description, ci.intQty, ci.curUnit_Price, i.szName, ci.intClass_Item_ID AS ExistingItemID, ig.szName AS ItemType,  " & _
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
			   			
					set rsItems = server.CreateObject("ADODB.Recordset")
					rsItems.CursorLocation = 3
					rsItems.Open sqlItems, Application("cnnFPCS")'oFunc.FPCScnn
					
					
					if rsItems.RecordCount < 1 then		
						' do nothing for now
					else
				%>
            <table cellpadding="3" id="Table5">
                <tr>
                    <td class="svplain11" colspan="6">
                        <b><i>Required Goods and Services</i></b>
                    </td>
                </tr>
                <tr>
                    <td class="gray" align="center">
                        <b>Type</b>
                    </td>
                    <td class="gray" align="center">
                        <b>Category</b>
                    </td>
                    <td class="gray" align="center">
                        <b>Name</b>
                    </td>
                    <td class="gray" align="center">
                        <b>Qty</b>
                    </td>
                    <td class="gray" align="center">
                        <b>Unit Price</b>
                    </td>
                    <td class="gray" align="center">
                        <b>Total</b>
                    </td>
                </tr>
                <%		
						dim dblTotal
						
						do while not rsItems.EOF
					%>
                <tr>
                    <td class="gray" align="center">
                        <% = rsItems("ItemType") %>
                    </td>
                    <td class="gray" align="center">
                        <% = rsItems("szName") %>
                    </td>
                    <td class="gray" align="center">
                        <% if rsItems("Description") & "" = "" then
									response.Write rsItems("szName")
							else
									response.Write rsItems("Description")
							end if			
							%>
                    </td>
                    <td class="gray" align="center">
                        <% = rsItems("intQty") %>
                    </td>
                    <td class="gray" align="center">
                        $<% = rsItems("curUnit_Price") %>
                    </td>
                    <td class="gray" align="right">
                        <% 
								dblTotal = cdbl(rsItems("intQty")) * cdbl(rsItems("curUnit_Price"))
								dblGrandTotal = dblGrandTotal + dblTotal
								Response.Write "$" & dblTotal
							%>
                    </td>
                </tr>
                <%
							rsItems.MoveNext
						loop
				%>
                <tr>
                    <td colspan="5" class="gray" align="right">
                        <b>Grand Total:</b>
                    </td>
                    <td class="gray" align="right">
                        $<% = dblGrandTotal %>
                    </td>
                </tr>
            </table>
            <br>
            <%
							end if
							rsItems.Close
							set rsItems = nothing	
						end if 
					%>
            <table id="Table9">
                <%
						if intVendor_ID <> "" then
							call vbfVendorFields
						elseif intInstructor_ID & "" <> "" then
							call vbfClassDetailsForASD(dblGrandTotal)
						end if 
					%>
            </table>
        </td>
    </tr>
</table>
<% if intInstructor_ID & "" <>  "" and false then%>
<table id="Table6">
    <tr>
        <td colspan="2" class="svplain10">
            Signatures below indicate acceptance of all applicable<br>
            sections of the Educational Plan & Annual Budget.
        </td>
    </tr>
    <tr>
        <td>
        </td>
    </tr>
</table>
<% 
	end if
	set rsInfo = nothing			
end function

%>
<script language="javascript">
		<% if request("noprint") = "" then %>
		if (window.print){
	      window.print()
	    }
	    else {
	      alert("Mac users: please press Apple-P to print this form.\nWindows users: Please press ctrl-P to print this form.")
		}
		<% end if %>
	</script>
<%
call oFunc.CloseCN	
Server.Execute(Application.Value("strWebRoot") & "Includes/footer.asp")

 function vbfClassDetailsForASD(pGSCost)
 %>
<tr>
    <td colspan="2">
        <font class="svplain11"><b><i>Class Costs</i></b> </font><font class="svplain11">(Teachers
            Time) </font>
    </td>
</tr>
<tr>
    <td class="gray" align="right">
        <% = formatNumber(decHours_Student,1) %>
    </td>
    <td class="gray">
        &nbsp;Number of teacher hours with student.
    </td>
</tr>
<tr>
    <td class="gray" align="right">
        <% = formatNumber(decHours_Planning,1) %>
    </td>
    <td class="gray">
        &nbsp;Number of hours for teacher planning.
    </td>
</tr>
<tr>
    <td class="gray" align="right">
        <% = formatNumber((CDBL(decHours_Student) + cdbl(decHours_Planning)),1) %>
    </td>
    <td class="gray">
        &nbsp;<b>Total teacher hours.</b>
    </td>
</tr>
<tr>
    <td class="gray" align="right">
        <% = formatNumber((CDBL(decHours_Student) + cdbl(decHours_Planning))/intMax_Students,1)%>
    </td>
    <td class="gray">
        &nbsp;Minimum number of hours to be charged to each student.
    </td>
</tr>
<tr>
    <td class="gray" align="right">
        <% = formatNumber((CDBL(decHours_Student) + cdbl(decHours_Planning))/intMin_Students,1)%>
    </td>
    <td class="gray">
        &nbsp;Maximum number of hours to be charged to each student.
    </td>
</tr>
<tr>
    <td class="gray" align="right">
        $<% = curInstructionRate %>
    </td>
    <td class="gray">
        &nbsp;Teachers hourly rate.
    </td>
</tr>
<tr>
    <td class="gray" align="right">
        $<% = formatNumber(((CDBL(decHours_Student) + cdbl(decHours_Planning))/intMax_Students)*curInstructionRate,2) %>
    </td>
    <td class="gray">
        &nbsp;Minimum total teacher cost per student.
    </td>
</tr>
<tr>
    <td class="gray" align="right">
        $<% = formatNumber(((CDBL(decHours_Student) + cdbl(decHours_Planning))/cdbl(intMin_Students))*cdbl(curInstructionRate),2) %>
    </td>
    <td class="gray">
        &nbsp;Maximum total teacher cost per student.
    </td>
</tr>
<tr>
    <td class="gray" align="right">
        $<% = pGSCost %>
    </td>
    <td class="gray">
        &nbsp;Total miscellaneous costs per student.
    </td>
</tr>
<tr>
    <td class="gray" align="right">
        $<% = formatNumber((((CDBL(decHours_Student) + cdbl(decHours_Planning))/intMax_Students)*curInstructionRate)+pGSCost,2) %>
    </td>
    <td class="gray">
        &nbsp;<b>Minimum total deduction per student account.</b>
    </td>
</tr>
<tr>
    <td class="gray" align="right">
        $<% = formatNumber((((CDBL(decHours_Student) + cdbl(decHours_Planning))/cdbl(intMin_Students))*cdbl(curInstructionRate))+pGSCost,2) %>
    </td>
    <td class="gray">
        &nbsp;<b>Maximum total deduction per student account.</b>
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
	rsInfo.Open sqlChargeID,Application("cnnFPCS")'oFunc.FPCScnn
	if rsInfo.RecordCount > 0 then
		strChargeList = rsInfo("szDesc")
	end if
	rsInfo.Close
	
 %>
<tr>
    <td colspan="2">
        <font class="svplain11"><b><i>Vendor Class Costs</i></b> </font>
    </td>
</tr>
<tr>
    <td class="gray" align="right">
        <% = decNum_Units %>
    <td class="gray">
        &nbsp;<% = strChargeList %>
        at
    </td>
</tr>
<tr>
    <td class="gray" align="right">
        $<% = formatNumber(curCharge_Amount,2) %>
    </td>
    <td class="gray" onchange="jfChanged();">
        &nbsp;Cost per unit.
    </td>
</tr>
<tr>
    <td class="gray" align="right">
        $<% = formatNumber(cdbl(decNum_Units) * cdbl(curCharge_Amount),2) %>
    </td>
    <td class="gray">
        &nbsp;<b>Contract Amount.</b>
    </td>
</tr>
</table>
<br>
<table id="Table10">
    <tr>
        <td colspan="2">
            <font class="svplain11"><b><i>Vendor Contract Information</i></b> </font>
        </td>
    </tr>
    <tr>
        <td class="gray" colspan="2">
            Description of Service Vendor is Providing
        </td>
    </tr>
    <tr>
        <td colspan="2" class="svplain10">
            <% = szService_desc %>
        </td>
    </tr>
    <tr>
        <td class="gray">
            &nbsp;Is Service to be provided
            <br>
            &nbsp;on District premises?
            <input disabled type="checkbox" <% if bolOn_Premises then Response.Write " checked " %>
                name="bolOn_Premises" value="TRUE" onchange="jfChanged();" id="Checkbox" <% if bolon_premises then response.write " checked " %>1><b>Yes</b>&nbsp;
        </td>
        <td class="gray">
            &nbsp;Is Service to be provided
            <br>
            &nbsp;using District equipment or supplies?
            <input disabled type="checkbox" <% if bolDistrict_Equip then Response.Write " checked " %>
                name="bolDistrict_Equip" value="TRUE" onchange="jfChanged();" id="Checkbox" <% if boldistrict_equip then response.write " checked " %>1><b>Yes</b>&nbsp;
        </td>
    </tr>
    <tr>
        <td class="svplain10">
            If yes indicate where?
        </td>
        <td class="svplain10">
            If yes indicate which equipment and supplies.
        </td>
    </tr>
    <tr>
        <td class="svplain10">
            <% = szPremises%>
        </td>
        <td align="center" class="svplain10">
            <% = szEquip_List %>
        </td>
    </tr>
    <tr>
        <td class="gray" colspan="2">
            &nbsp;Will the District pay Contractor for Service updon satisfactory completion
            <br>
            &nbsp;and acceptance of ALL work required uder this contract?
            <input disabled type="checkbox" <% if bolPay_On_Completion then Response.Write " checked " %>
                name="bolPay_On_Completion" value="TRUE" id="Checkbox" <% if bolpay_on_completion then response.write " checked " %>1><b>Yes</b>&nbsp;
        </td>
    </tr>
    <tr>
        <td class="svplain10" colspan="2">
            If NO indicate when (partial) payment(s) to be made.
        </td>
    </tr>
    <tr>
        <td colspan="2" class="svplain10">
            <% = szPay_Details %>
        </td>
    </tr>
    <%
 end function
 
 function vbfInstructorFields
%>
    <table id="Table8" style="width: 100%;">
        <tr>
            <td class="gray" nowrap>
                &nbsp;Registration Deadline
            </td>
            <td class="gray" nowrap>
                &nbsp;Min # Students
            </td>
            <td class="gray" nowrap>
                &nbsp;Max # Students
            </td>
            <td class="gray" nowrap>
                &nbsp;Grade&nbsp;
            </td>
            <td class="gray" style="width: 100%;">
                &nbsp;to Grade&nbsp;
            </td>
        </tr>
        <tr>
            <td class="svplain10" nowrap align="center">
                <% = monthReg & "/" & dayReg & "/" & yearReg %>
            </td>
            <td align="center" class="svplain10" nowrap>
                <% = intMin_Students%>
            </td>
            <td align="center" class="svplain10" nowrap>
                <% = intMax_Students%>
            </td>
            <td align="center" class="svplain10" nowrap>
                <% = sGrade_Level%>
            </td>
            <td class="svplain10">
                &nbsp;&nbsp;&nbsp;&nbsp;<%= sGrade_Level2 %>
            </td>
        </tr>
    </table>
    <table id="Table11" style="width: 100%;">
        <tr>
            <td class="gray" nowrap>
                &nbsp;Class Start Date
            </td>
            <td class="gray" nowrap>
                &nbsp;Class End Date
            </td>
            <td class="gray" align="left" style="width: 100%;">
                &nbsp;Meets Every
            </td>
        </tr>
        <tr>
            <td valign="top" class="svplain10" nowrap>
                <% = monthStart & "/" & dayStart & "/" & yearStart %>
            </td>
            <td valign="top" class="svplain10" nowrap>
                <% = monthEnd & "/" & dayEnd & "/" & yearEnd %>
            </td>
            <td class="svplain10">
                <% 	
								if szDays_Meet_On <> "" then						
									dim sqlDays
									sqlDays = "select strText from common_lists where intList_ID = 4 and strValue='" & szDays_Meet_On & "'"
									rsInfo.Open sqlDays,Application("cnnFPCS")'oFUnc.FPCScnn
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
								end if		
							%>
            </td>
        </tr>
    </table>
    <table id="Table12" style="width: 100%;">
        <tr>
            <td class="gray" nowrap>
                &nbsp;Class Start Time
            </td>
            <td class="gray" nowrap>
                &nbsp;Class End Time
            </td>
            <td class="gray" style="width: 100%;">
                &nbsp;Schedule Comments
            </td>
        </tr>
        <tr>
            <td valign="top" class="svplain10" align="center" nowrap>
                <% = hourStart%>:<% = minuteStart %>
                <% =amPmStart %>
            </td>
            <td valign="top" class="svplain10" align="center" nowrap>
                <%= hourEnd %>:<% = minuteEnd %>
                <% = amPmEnd %>
            </td>
            <td class="svplain10">
                <% = szSchedule_Comments%>
            </td>
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
    <table id="Table13">
        <tr>
            <td class="gray">
                &nbsp;Min # Students &nbsp;
            </td>
            <td class="gray">
                &nbsp;Max # Students &nbsp;
            </td>
            <td class="gray">
                &nbsp;Class Duration &nbsp;
            </td>
            <td class="gray">
                &nbsp;Session Length &nbsp;
            </td>
        </tr>
        <tr>
            <td align="center" class="svplain10">
                <% = intMin_Students%>
            </td>
            <td align="center" class="svplain10">
                <% = intMax_Students%>
            </td>
            <td class="svplain10" align="center">
                <% = szDuration_Name %>
            </td>
            <td valign="top" align="center" class="svplain10">
                <% = intSession_Minutes%>
                minutes
            </td>
        </tr>
    </table>
    <table id="Table14">
        <tr>
            <td class="gray" class="svplain10">
                &nbsp;Meets Every
            </td>
            <td class="gray" class="svplain10">
                &nbsp;Comments
            </td>
        </tr>
        <tr>
            <td class="svplain10" align="center">
                <% = szDays_Meet_On %>
            </td>
            <td align="center" class="svplain10">
                <% = szSchedule_Comments %>
            </td>
        </tr>
    </table>
    <%
end function

function vbfPrintILP(intILP_ID)

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'' This section will get ILP info for both Genric and non-Generic ILP's 
	'' Depending on the incoming request
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
	dim strILPTable
	dim strGradeTable
	dim strILPFields 
	 
	if intStudent_ID <> ""  then			
		' Pull from real ILP's
		strILPFields = "i.dtStudent_Enrolled,"
		strILPTable = " tblILP "
		strGradeTable = " tblGrading_Scale "
		
		sql = "SELECT i.dtStudent_Enrolled, i.dtCreate, c.szClass_Name, i.sintSchool_Year, i.intClass_ID, i.intSemester, i.decCourse_Hours, i.szCurriculum_Desc, i.szGoals,  " & _
			" i.szRequirements, i.szTeacher_Role, i.szStudent_Role, i.szParent_Role, i.szEvaluation, i.bolPass_Fail, i.szOther_Grading, i.bolGradingScale, i.intGrading_Scale_ID, i.szILP_Additions, tblILP_Generic.szILP_Additions AS Teacher_Additions, " & _
			" i.intContract_Guardian_ID, (CASE isNull(c.intGuardian_ID, 1)  " & _
			" WHEN 1 THEN ins.szFirst_Name + ' ' + ins.szLast_Name ELSE g.szFIRST_NAME + ' ' + g.szLAST_NAME END) AS TaughtBy, p.szPhilosophy " & _
			"FROM tblENROLL_INFO e LEFT OUTER JOIN " & _
			" tblPhilosophy p ON e.intPhilosophy_ID = p.intPhilosophy_ID RIGHT OUTER JOIN " & _
			" tblClasses c INNER JOIN " & _
			" tblILP i ON c.intClass_ID = i.intClass_ID ON e.intSTUDENT_ID = i.intStudent_ID LEFT OUTER JOIN " & _
			" tblGUARDIAN g ON c.intGuardian_ID = g.intGUARDIAN_ID LEFT OUTER JOIN " & _
			" tblINSTRUCTOR ins ON c.intInstructor_ID = ins.intINSTRUCTOR_ID LEFT OUTER JOIN " & _
			" tblILP_Generic ON c.intClass_ID = tblILP_Generic.intClass_ID " & _
			"WHERE (i.intILP_ID = " & intILP_ID & ") AND (e.sintSCHOOL_YEAR = " & session.Contents("intSchool_Year") & ") "
			
	elseif ucase(request("ILP_TYPE")) = "I" then
		' Pull for ILP Bank (for guardian)
		strILPFields = "i.dtStudent_Enrolled,"
		strILPTable = " tblILP "
		strGradeTable = " tblGrading_Scale "
		
		sql = "SELECT i.dtStudent_Enrolled, c.szClass_Name, i.sintSchool_Year, i.intClass_ID, i.intSemester, i.decCourse_Hours, i.szCurriculum_Desc, i.szGoals,  " & _
			" i.szRequirements, i.szTeacher_Role, i.szStudent_Role, i.szParent_Role, i.szEvaluation, i.bolPass_Fail, i.szOther_Grading, i.intGrading_Scale_ID " & _			
			"FROM tblClasses c INNER JOIN " & _
			" tblILP i ON c.intClass_ID = i.intClass_ID " & _
			"WHERE (i.intILP_ID = " & intILP_ID & ")"
	elseif ucase(request("ILP_TYPE")) = "C" then	
		' Pull for ILP Bank (for instructor)
		strILPFields = ""
		strILPTable = " tblILP_Generic "
		strGradeTable = " tblGrading_Scale_Generic "	
		sql = "SELECT     c.szClass_Name, i.sintSchool_Year, i.intClass_ID, i.intSemester, i.decCourse_Hours, i.szCurriculum_Desc, i.szGoals, i.szRequirements,  " & _
			" i.szTeacher_Role, i.szStudent_Role, i.szParent_Role, i.szEvaluation, i.bolPass_Fail, i.szOther_Grading, i.intGrading_Scale_ID  " & _
			"FROM tblClasses c INNER JOIN " & _
			" tblILP_Generic i ON c.intClass_ID = i.intClass_ID " & _
			"WHERE (i.intILP_ID = " & intILP_ID & ")"		
	else
		' Pull from Generic (for teachers)
		strILPFields = ""
		strILPTable = " tblILP_Generic "
		strGradeTable = " tblGrading_Scale_Generic "	
		sql = "SELECT     c.szClass_Name, i.sintSchool_Year, i.intClass_ID, i.intSemester, i.decCourse_Hours, i.szCurriculum_Desc, i.szGoals, i.szRequirements,  " & _
			" i.szTeacher_Role, i.szStudent_Role, i.szParent_Role, i.szEvaluation, i.bolPass_Fail, i.szOther_Grading, i.intGrading_Scale_ID,i.szILP_Additions AS Teacher_Additions,  " & _
			" i.intContract_Guardian_id, ins.szFIRST_NAME + ' ' + ins.szLAST_NAME AS TaughtBy " & _
			"FROM tblClasses c INNER JOIN " & _
			" tblILP_Generic i ON c.intClass_ID = i.intClass_ID LEFT OUTER JOIN " & _
			" tblINSTRUCTOR ins ON c.intInstructor_ID = ins.intINSTRUCTOR_ID " & _
			"WHERE (i.intILP_ID = " & intILP_ID & ")"		
	end if 

	'We need to populate the ILP info since we were given an ILP ID
	set rsILP = server.CreateObject("ADODB.RECORDSET")
	rsILP.CursorLocation = 3	
	
	rsILP.Open sql,Application("cnnFPCS")'oFunc.FPCScnn

	intCount = 0
	'This for loop dimentions and defines all the columns we selected in sqlClass
	'and we use the variables created here to populate the form.
	for each item in rsILP.Fields
		execute("dim " & rsILP.Fields(intCount).Name)
		execute(rsILP.Fields(intCount).Name & " = item")		
		intCount = intCount + 1
	next  
	

	rsILP.Close

	set rsILP = nothing
	
	'Seperate student enrollment date for select box populating
	
	intCount = 0
	if intGrading_Scale_ID <> "" then
		set rsGrade = server.CreateObject("ADODB.RECORDSET")
		rsGrade.CursorLocation = 3
		sql = "select intA_Upper,intA_Lower,intB_Upper,intB_Lower," & _
				"intC_Upper,intC_Lower,intD_Upper,intD_Lower,intF_Upper,intF_Lower " & _
				"from " & strGradeTable  & _
				" where intGrading_Scale_Id = " & intGrading_Scale_ID
		rsGrade.Open sql,Application("cnnFPCS")'oFunc.FPCScnn
	
		if rsGrade.RecordCount > 0 then
			for each item in rsGrade.Fields
				execute("dim " & rsGrade.Fields(intCount).Name)
				execute(rsGrade.Fields(intCount).Name & " = item")		
				intCount = intCount + 1
			next  
		end if
		
		rsGrade.Close
		set rsGrade = nothing
	end if 			

if sintSchool_Year = "" then
	sintSchool_Year = session.Value("intSchool_Year")
end if
%>
    <table width="100%" id="Table15" cellpadding="2" cellspacing="2">
        <% = vbfFormHeader("ILP (Individual Learning Plan) #" & intClass_ID) %>
        <tr>
            <td colspan="2">
                <% if request("ILP_TYPE") & "" = "" then %>
                <table id="Table17">
                    <tr>
                        <% if strStudentName <> "" then %>
                        <td class="gray" align="center">
                            &nbsp;Date Enrolled
                        </td>
                        <td class="gray">
                            &nbsp;Student
                        </td>
                        <% end if %>
                        <td class="gray">
                            &nbsp;Instructor
                        </td>
                        <td class="gray" align="center">
                            &nbsp;Class Name&nbsp;
                        </td>
                        <td class="gray" align="center">
                            &nbsp;School Yr&nbsp;
                        </td>
                        <td class="gray" align="center" title="Number of hours this course contributes to core hours.">
                            &nbsp;Course Hrs&nbsp;
                        </td>
                    </tr>
                    <tr>
                        <% if strStudentName <> "" then %>
                        <td class="svplain10" valign="middle" align="center">
                            <% = formatDateTime(dtCreate,2) %>
                        </td>
                        <td class="svplain10">
                            &nbsp;<% = strStudentName %>&nbsp;
                        </td>
                        <% end if %>
                        <td class="svplain10" valign="middle">
                            &nbsp;<% = TaughtBy %>&nbsp;
                        </td>
                        <td class="svplain10">
                            &nbsp;<% = szClass_Name %>&nbsp;
                        </td>
                        <td class="svplain10" valign="middle" align="center">
                            <%= sintSchool_Year %>
                        </td>
                        <td align="center" class="svplain10" valign="middle">
                            <% = decCourse_Hours %>
                        </td>
                    </tr>
                </table>
                <% end if %>
                <table id="Table18">
                    <tr>
                        <td class="gray">
                            &nbsp;&nbsp;<b>1. Description of the course including methods, curriculum and supplies
                                needed ...</b>
                        </td>
                    </tr>
                    <tr>
                        <td class="svplain10">
                            <%=szCurriculum_Desc%>
                        </td>
                    </tr>
                </table>
                <table id="Table19">
                    <tr>
                        <td class="gray">
                            <table cellpadding="3" cellspacing="0" class="gray2" id="Table38">
                                <tr>
                                    <td valign="top">
                                        <b>2.</b>
                                    </td>
                                    <td><b>Standards: Common core or GLE</b>
                                        <!--<b>The student will learn ... </b>
                                        <br>
                                        (a minimum of 2 examples)-->
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <!--<td>
                            &nbsp;
                        </td>-->
                        <td class="gray">
                            <table cellpadding="3" cellspacing="0" class="gray2" id="Table39">
                                <tr>
                                    <td valign="top">
                                        <b>3.</b>
                                    </td>
                                    <td>
                                        <b>Student will be involved in these
                                            <br>
                                            activities/assignments ... </b>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td class="svplain10" valign="top">
                            <%=szTeacher_Role%>
                        <%If False Then %>
                    <%End If %>
                        </td>
                        <!--<td>
                            &nbsp;
                        </td>-->
                        <td class="svplain10" valign="top">
                            <%=szRequirements%>
                        </td>
                    </tr>
                </table>
                <br>
                <table id="Table20">
                    <!--<tr>
                        <td colspan="3">
                            <font class="svplain11"><b><i>Roles</i></b> </font>
                        </td>
                    </tr>-->
                    <tr>
                        <td class="gray">
                            &nbsp;<b>4. Materials, Resources</b>
                        </td>
                        <td class="gray">
                            &nbsp;<b>5. Role of Parent/Teacher/Vendor/any additional responsibilities of the student</b>
                        </td>
                        <!--<td class="gray">
                            <table cellpadding="3" cellspacing="0" class="gray2" border="0">
                                <tr>
                                    <td valign="top">
                                        <b>6.</b>
                                    </td>
                                    <td>
                                        <b>Role of Parent/Teacher/Vendor</b>
                                    </td>
                                </tr>
                            </table>
                        </td>-->
                    </tr>
                    <tr>
                        <td class="svplain10" valign="top">
                            <%=szStudent_Role%>
                        </td>
                        <td class="svplain10" valign="top">
                            <%=szParent_Role%>
                        </td>
                        <!--<td class="svplain10" valign="top">
                            <%=szTeacher_Role%>
                        </td>-->
                    </tr>
                </table>
                <table id="Table21">
                    <tr>
                        <td colspan="3">
                            <font class="svplain11"><b><i>Evaluation and Grading</i></b> </font>
                        </td>
                    </tr>
                    <tr>
                        <td class="gray" valign="top">
                            <nobr>&nbsp;<input type=checkbox name="bolPass_Fail" <% if bolPass_Fail&"" <> "" and bolPass_Fail <> 0 then response.Write(" checked ")%> ID="Checkbox1" disabled>Pass/Fail&nbsp;</nobr>
                        </td>
                        <td class="gray">
                            &nbsp;<input type="checkbox" name="bolGrading_Scale" <% if bolGradingScale &"" <> "" then response.Write(" checked ")%>
                                id="Checkbox2" disabled>Grading Scale
                        </td>
                        <td class="gray">
                            &nbsp;<input type="checkbox" name="bolOther_Grading" <% if szOther_Grading&"" <> "" then response.Write(" checked ")%>
                                id="Checkbox3" disabled>Other
                        </td>
                    </tr>
                    <tr>
                    <td>&nbsp;</td>
                        <td colspan="1">
                    <% if bolGradingScale &"" <> "" then %>
                            <table id="Table22">
                                <tr>
                                    <td class="gray" nowrap>
                                        &nbsp;A =
                                    </td>
                                    <td nowrap class="svplain">
                                        90% to 100%
                                    </td>
                                    <td class="gray" nowrap>
                                        &nbsp;B =
                                    </td>
                                    <td nowrap class="svplain">
                                        80% to 89%
                                    </td>
                                    <td class="gray" nowrap>
                                        &nbsp;C =
                                    </td>
                                    <td nowrap class="svplain">
                                        70% to 79%
                                    </td>
                                </tr>
                                <tr>
                                    <td class="gray" nowrap>
                                        &nbsp;D =
                                    </td>
                                    <td nowrap class="svplain">
                                        60% to 69%
                                    </td>
                                    <td class="gray" nowrap>
                                        &nbsp;F =
                                    </td>
                                    <td nowrap class="svplain">
                                        0% to 59%
                                    </td>
                                    <td nowrap class="svplain" colspan="2">
                                        &nbsp;
                                    </td>
                                </tr>
                            </table>
                    <% end if %>
                        </td>
                        <td class="svplain8" colspan="3">
                    <% if szOther_Grading <> "" then%>
                    <!--<tr>-->
                            <b>6. Explain:</b><br>
                            <%=szOther_Grading%>
                   <!-- </tr>-->
                    <% end if %>
                        </td>
                    </tr>
                    <tr>
                        <td class="gray" colspan="3">
                            <table cellpadding="3" cellspacing="0" class="gray2" id="Table41">
                                <tr>
                                    <td valign="top">
                                        <b>7.</b>
                                    </td>
                                    <td>
                                        <b>What will be evaluated? </b>(worksheets, tests, class participation, daily work,
                                        logs, attendance, etc.)<br>
                                        <b>How will student be evaluated?</b> (# of projects, work correction, logs, informal/formal
                                        evaluations, portfolios, oral presentations, etc.)
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td valign="top" colspan="3" class="svplain10">
                            <% = szEvaluation %>
                        </td>
                    </tr>
                    <tr>
                        <td class="gray" colspan="3">
                            <table cellpadding="3" cellspacing="0" class="gray2" id="Table41">
                                <tr>
                                    <td valign="top">
                                        <b>8.</b>
                                    </td>
                                    <td>
                                        <b>Course Syllabus: work out a timeline (scope and sequence) of all major topics to be covered.
                                        <b>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td valign="top" colspan="3" class="svplain10">
                            <%Dim re
                    Set re=New RegExp
                    re.Pattern="[<]tbody[>]"
                    re.IgnoreCase=True
                    If re.Test(szGoals) then %>
                            <table class="svplain10">
                                <%=szGoals%></table>
                            <%Else %>
                            <pre>
<%=szGoals %>
                            </pre>
                            <%End If %>
                        </td>
                    </tr>

                    <% if szILP_Additions <> "" then %>
                    <tr>
                        <td colspan="3">
                            <br>
                        </td>
                    </tr>
                    <tr>
                        <td class="gray" colspan="3">
                            &nbsp;<b>ILP Additions</b>
                        </td>
                    </tr>
                    <tr>
                        <td class="svplain10" colspan="3">
                            <% = szILP_Additions %>
                        </td>
                    </tr>
                    <% end if %>
                    <% if Teacher_Additions <> "" then %>
                    <tr>
                        <td colspan="3">
                            <br>
                        </td>
                    </tr>
                    <tr>
                        <td class="gray" colspan="3">
                            &nbsp;<b>Instructor Additions</b>
                        </td>
                    </tr>
                    <tr>
                        <td class="svplain10" colspan="3">
                            <% = Teacher_Additions %>
                        </td>
                    </tr>
                    <% end if %>
                </table>
            </td>
        </tr>
    </table>
    <%
end function

function vbfGoodsServices()
	dim strFor
	If intStudent_ID <> ""  THEN
		dim strGSWhere
		if intILP_ID <> "" and intILP_ID <> "ALL" then
			strGSWhere = intILP_ID
		elseif intClass_ID <> "" then
			strGSWhere = "(SELECT top 1 intILP_ID " & _ 
						"FROM tblILP " & _ 
						"WHERE     (intClass_ID = " & intClass_ID & ") " & _
						"AND (intStudent_ID = " & intStudent_ID & ") )"
		end if 
		
		sqlItems = "SELECT oi.intOrdered_Item_ID as ExistingItemID, v.szVendor_Name,oi.bolSponsor_Approved, " & _
				"ig.szName AS grp_name,ig.intItem_Group_ID, i.szName AS item_name, ci.bolRequired, oi.bolApproved," & _
				"oi.szDeny_Reason,oi.intQty, (oi.intQty * oi.curUnit_Price) as Total, " & _
				"           (SELECT     (CASE oa2.szValue " & _
				"						WHEN '0' then 'No'	" & _
				"						WHEN '1' then 'Yes'	" & _
				"						ELSE 'Not Given'		" & _
				"						END) as consum			" & _
				"            FROM          tblOrd_Attrib oa2 " & _
				"            WHERE      oa2.intOrdered_Item_Id = oi.intOrdered_Item_Id AND " & _
				"			oa2.intItem_Attrib_ID = 15) AS Consumable, " & _
				"           (SELECT     oa2.szValue " & _
				"            FROM          tblOrd_Attrib oa2 " & _
				"            WHERE      oa2.intOrdered_Item_Id = oi.intOrdered_Item_Id AND " & _
				"			oa2.intItem_Attrib_ID = 25) AS Semester, " & _
				"           (SELECT     oa2.szValue " & _
				"            FROM          tblOrd_Attrib oa2 " & _
				"            WHERE      oa2.intOrdered_Item_Id = oi.intOrdered_Item_Id AND " & _
				"			oa2.intItem_Attrib_ID = 3) + ' - ' + " & _
				"           (SELECT     oa2.szValue " & _
				"            FROM          tblOrd_Attrib oa2 " & _
				"            WHERE      oa2.intOrdered_Item_Id = oi.intOrdered_Item_Id AND " & _
				"            oa2.intItem_Attrib_ID = 4) AS Dates, " & _
				"          (SELECT     oa2.szValue " & _
				"             FROM          tblOrd_Attrib oa2 " & _
				"             WHERE      oa2.intOrdered_Item_Id = oi.intOrdered_Item_Id AND " & _
				"			 (oa2.intItem_Attrib_ID = 9 OR " & _
				"              oa2.intItem_Attrib_ID = 5 OR " & _
				"              oa2.intItem_Attrib_ID = 22)) AS iName " & _
				" FROM tblOrdered_Items oi INNER JOIN " & _
				"       tblVendors v ON oi.intVendor_ID = v.intVendor_ID INNER JOIN " & _
				"       trefItems i ON oi.intItem_ID = i.intItem_ID INNER JOIN " & _
				"       trefItem_Groups ig ON i.intItem_Group_ID = ig.intItem_Group_ID " & _
				"	   LEFT OUTER JOIN " & _
				"       tblClass_Items ci ON oi.intClass_Item_ID = ci.intClass_Item_ID " & _			   
				" WHERE (oi.intILP_ID = " &  strGSWhere & ")" & _ 
				" ORDER by i.szName "
										
		strFor = "<nobr>Class Name: " & szClass_Name & "</nobr><BR><nobr>Student Name: " & strStudentName & "</nobr>"
				
	elseif intClass_ID <> "" then

		sqlItems = "SELECT ci.intClass_Item_ID as ExistingItemID,v.szVendor_Name, ig.intItem_Group_ID,ig.szName AS grp_name, i.szName AS item_name, " & _
				"ci.intQty, (ci.intQty * ci.curUnit_Price) as Total, " & _
				"           (SELECT     ca2.szValue " & _
				"            FROM          tblClass_Attrib ca2 " & _
				"            WHERE      ca2.intClass_Item_Id = ci.intClass_Item_Id AND " & _
				"			ca2.intItem_Attrib_ID = 15) AS Consumable, " & _
				"           (SELECT     ca2.szValue " & _
				"            FROM          tblClass_Attrib ca2 " & _
				"            WHERE      ca2.intClass_Item_Id = ci.intClass_Item_Id AND " & _
				"			ca2.intItem_Attrib_ID = 25) AS Semester, " & _
				"           (SELECT     ca2.szValue " & _
				"            FROM          tblClass_Attrib ca2 " & _
				"            WHERE      ca2.intClass_Item_Id = ci.intClass_Item_Id AND " & _
				"			ca2.intItem_Attrib_ID = 3) + ' - ' + " & _
				"           (SELECT     ca2.szValue " & _
				"            FROM          tblClass_Attrib ca2 " & _
				"            WHERE      ca2.intClass_Item_Id = ci.intClass_Item_Id AND " & _
				"            ca2.intItem_Attrib_ID = 4) AS Dates, " & _
				"          (SELECT     ca2.szValue " & _
				"             FROM          tblClass_Attrib ca2 " & _
				"             WHERE      ca2.intClass_Item_Id = ci.intClass_Item_Id AND " & _
				"			 (ca2.intItem_Attrib_ID = 9 OR " & _
				"              ca2.intItem_Attrib_ID = 5 OR " & _
				"              ca2.intItem_Attrib_ID = 22)) AS iName, '' as szDeny_Reason " & _
				" FROM tblClass_Items ci INNER JOIN " & _
				"       tblVendors v ON ci.intVendor_ID = v.intVendor_ID INNER JOIN " & _
				"       trefItems i ON ci.intItem_ID = i.intItem_ID INNER JOIN " & _
				"       trefItem_Groups ig ON i.intItem_Group_ID = ig.intItem_Group_ID " & _
				" WHERE (ci.intClass_ID = " & intClass_ID & ")  " & _
				"order by i.szName "
			if szClass_Name = "" then
				szClass_Name = oFunc.ClassInfo(intClass_ID,1)
			end if
			strFor = "Class Name: " & szClass_Name				
	end if	

	set rsItems = server.CreateObject("ADODB.Recordset")
	rsItems.CursorLocation = 3
	rsItems.Open sqlItems, Application("cnnFPCS")'oFunc.FPCScnn
	
	if rsItems.recordcount < 1 then
		rsItems.close
		set rsItems = nothing
		exit function
	end if
	if request("strAction") = "" then response.Write "<p></p>"
%>
    <table width="100%" id="Table25" cellpadding="2" cellspacing="2">
        <tr>
            <td align="left">
                <img src="<% = Application("strImageRoot")%>fpcsLogo.gif">
            </td>
            <td align="right" class="svplain10" width="100%">
                <% = Application.Contents("SchoolAddress") %>
            </td>
        </tr>
        <tr class="yellowHeader">
            <td colspan="2">
                <table align="right" id="Table26">
                    <tr>
                        <td align="right">
                            <font face="arial" size="2" color="white">
                                <% = date()%></font>
                        </td>
                    </tr>
                </table>
                &nbsp;<b>Goods and Services #<%=intClass_ID%></b>
            </td>
        </tr>
        <tr>
            <td class="svplain10">
                <% = strFor %>
            </td>
        </tr>
        <%
if rsItems.RecordCount < 1 then
%>
        <table id="Table23">
            <tr>
                <td class="gray">
                    &nbsp;No Goods or Services have been added to this class.
                </td>
            </tr>
        </table>
        <%
else
%>
        <table cellpadding="3" id="Table24" border="1" cellspacing="0">
            <tr>
                <td class="svplain10" align="center">
                    <b>Category</b>
                </td>
                <td class="svplain10" align="center">
                    <b>Vendor</b>
                </td>
                <td class="svplain10" align="center">
                    <b>Name</b>
                </td>
                <td class="svplain10" align="center" title="Consumable Item">
                    <b>Cons</b>
                </td>
                <td class="svplain10" align="center">
                    <b>Dates</b>
                </td>
                <td class="svplain10" align="center">
                    <b>Total</b>
                </td>
                <td class="svplain10" align="center">
                    <b>Status</b>
                </td>
            </tr>
            <%
	dim dblGrandTotal
	dim dblTotal
		
	do while not rsItems.EOF	
		' Determine what the user can delete/view/edit
		strVeiwEdit = "View/Edit"
		strClass = "svplain10"
		if intStudent_ID <> ""  then
			' added 'not admin' logic to all admins to delete all goods/services.
			if not rsItems("bolApproved") then 
				strClass = "grayStrike"
				strDeleteBt = "Rejected"
			elseif (rsItems("bolApproved") = true or rsItems("bolRequired") = true) then
				strDeleteBt = "Approved"
			else
				strDeleteBt = "Pending"
			end if
		else
			strDeleteBt =  ""
		end if
		
%>
            <tr>
                <td class="<% = strClass %>" align="center" title="<% = rsItems("grp_name") %>">
                    &nbsp;<% = rsItems("item_name") %>
                </td>
                <td class="<% = strClass %>" align="center">
                    &nbsp;<% = rsItems("szVendor_Name") %>
                </td>
                <td class="<% = strClass %>" align="center">
                    &nbsp;<% =  rsItems("iName") %>
                </td>
                <td class="<% = strClass %>" align="center">
                    &nbsp;<% = rsItems("Consumable") %>
                </td>
                <td class="<% = strClass %>" align="center">
                    &nbsp;<% = rsItems("Dates") %>
                </td>
                <td class="<% = strClass %>" align="right">
                    &nbsp;<% 
				' Exclude item in grand total if item has been denied
				if strClass <> "grayStrike" then
					dblGrandTotal = dblGrandTotal + rsItems("Total")
				end if
			    response.Write "$" & rsItems("Total")
			%>
                </td>
                <td class="svplain10" align="center" title="<% = rsItems("szDeny_Reason") %>">
                    &nbsp;<% = strDeleteBt %>
                </td>
            </tr>
            <%
		rsItems.MoveNext
	loop
%>
            <tr>
                <td colspan="5" class="svplain10" align="right">
                    <b>Grand Total:</b>
                </td>
                <td class="svplain10" align="right">
                    $<% = dblGrandTotal %>
                </td>
                <td class="svplain10" colspan="2">
                    &nbsp;
                </td>
            </tr>
        </table>
        <br>
        <table>
            <% if HideSigs = "" then %>
            <tr>
                <td colspan="10">
                </td>
            </tr>
            <% end if %>
        </table>
        <br>
        <%
end if
rsItems.Close
set rsItems = nothing

end function

function vbfFormHeader(pText)
%>
        <tr>
            <td align="left">
                <img src="<% = Application("strImageRoot")%>fpcsLogo.gif">
            </td>
            <td align="right" class="svplain10" width="100%" nowrap>
                <% = Application.Contents("SchoolAddress") %>
            </td>
        </tr>
        <tr class="yellowHeader">
            <td colspan="2">
                <table align="right" id="Table27">
                    <tr>
                        <td align="right">
                            <font face="arial" size="2" color="white">
                                <% = date()%></font>
                        </td>
                    </tr>
                </table>
                &nbsp;<b><% = pText %></b>
            </td>
        </tr>
        <%
end function

function vbfStudentPacket
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'Name:		packet.asp
	'Purpose:	Main information page contaning Course management, budgets,
	'			and student status information
	'Date:		26 oCt 2004
	'Author:	Scott Bacon (ThreeShapes.com LLC)
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	dim intShort_ILP_ID , strWhere
	dim intPreviousID		' Used to determine when a course is changed in rsBudget
	dim dblTargetBalance	' Target start - all budgeted expenses
	dim dblActualBalance	' Actual start - all actual expenses
	dim dblWithdraw			' Amount to reduce budget funding by due to Budget Transfer withdrawal 
	dim dblDeposit			' Amount to reduce budget funding by due to Budget Transfer deposit 
	dim dblBudgetCost		' Calculated cost for a budgeted item
	dim dblUnitCost			' Used to handle teachers cost vs budgeted goods/services
	dim dblShipping			' Used to track shipping costs
	dim dblCharge 
	dim dblAdjBudget 
	dim mDivCount
	dim mLablelCount
	dim strBList
	dim strDateField
	dim bStatus				' budgeted item status
	dim strItemType			' tells user if item is requiestion or reimbursement
	dim oHtml

	mLablelCount = 0
	mDivCount	 = 0	

	oFunc.ResetSelectSessionVariables

	set oBudget = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/StudentBudgetInfo.wsc"))
	set oHtml   = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/htmlFuncs.wsc"))

	'Initialize some key variables

	if intStudent_ID <> "" then	
		
		'oBudget.PopulateStudentFunding oFunc.FPCScnn,intStudent_ID,session.Contents("intSchool_Year")
		oBudget.PopulateStudentFunding Application("cnnFPCS"),intStudent_ID,session.Contents("intSchool_Year")
		
		dblDeposits = oBudget.Deposits
		dblWithdraw = oBudget.Withdrawls
		dblActualBalance = oBudget.ActualFunding
		dblTargetBalance = oBudget.BudgetFunding 	
		intEnroll_Info_ID = oBudget.EnrollInfoId
		myBudgetBalance = oBudget.BudgetBalance
		myActualBalance = oBudget.ActualBalance	
	else
		'terminate page since page was improperly called.
		response.Write "<html><body><H1>Page improperly called.</h1></body></html>"
		response.End
	end if

	if oBudget.FamilyId & "" = "" or oBudget.StudentGrade & "" = ""then		
	%>
        <table cellspacing="0" cellpadding="4" width="85%" id="Table34">
            <tr>
                <td class="svplain10">
                    <% if oBudget.FamilyId & "" = "" then %>
                    <b>This student does not belong to a family in the Student Information System.</b>
                    <br>
                    An Administrator will need to add the student to a family before work on the packet
                    can begin.
                    <% else %>
                    <b>A grade has not been selected for this student.</b><br>
                    Before work can begin on the packet you will need to go to the student profile and
                    enter the students' current grade.
                    <%end if%>
                </td>
            </tr>
        </table>
        <%		
		set oBudget = nothing
		response.End
	elseif isNumeric(oBudget.EnrollmentId) and isNumeric(oBudget.IepId) then
		' Student Profiles have been updated by family. Now we check to see if a sponsor has been selected.
		if not isNumeric(oBudget.SponsorID)  then Sponsor = "No Sponsor Selected": SponsorEmail=""
	elseif false then 
		%>
        <table cellspacing="0" cellpadding="4" width="85%" id="Table35">
            <tr>
                <td class="svplain10">
                    <b>Before you can plan any courses you must update your students information for SY
                        <% = session.Contents("intSchool_Year")%>. To do this click on the 'Family Manager'
                        link on the menu above follow the instructions found on that page.</b> </b>
                </td>
            </tr>
        </table>
        <%
				
		set oBudget = nothing
		response.End
	end if

	'Find out if student is in High School
	if isNumeric(oBudget.StudentGrade) then
		if cint(oBudget.StudentGrade) >= 9 then
			bolHighSchool = true
		else
			bolHighSchool = false
		end if
	end if
%>
        <form name="main" action="<%=Application("strSSLWebRoot")%>forms/packet/packet.asp"
        method="post" id="Form2">
        <input type="hidden" name="intStudent_ID" value="<%=intStudent_ID%>" id="Hidden1">
        <input type="hidden" name="bolHighSchool" value="<%=bolHighSchool%>" id="Hidden7">
        <input type="hidden" name="courseTitleData" value="" id="Hidden9">
        <input type="hidden" name="simpleHeader" value="<% = request("simpleHeader") %>"
            id="Hidden10">
        <input type="hidden" name="lastIndex" value="" id="Hidden11">
        <input type="hidden" name="ClassName" value="" id="Hidden12">
        <table style="width: 640px;" id="Table40">
            <tr>
                <td style="width: 100%;">
                    <table style="width: 100%;" id="Table42">
                        <% = vbfFormHeader("<font face=arial size=2 color=white><b> Student Packet/Budget for " & oBudget.StudentName & " </b> &nbsp;&nbsp;Grade: " &  oBudget.StudentGrade & "</font>") %>
                    </table>
                </td>
            </tr>
            <tr>
                <td style="width: 100%;">
                    <table id="Table44" style="width: 100%;">
                        <tr>
                            <td valign="top" style="width: 50%;">
                                <table id="Table45" style="width: 100%;">
                                    <tr>
                                        <td valign='top' style='height: 100%;'>
                                            <table id="Table46" cellspacing='1' cellpadding='4' style='height: 100%; width: 100%;'>
                                                <tr>
                                                    <td class="TableHeader" align="center">
                                                        <b>Progress<br>
                                                            Chart</b>
                                                    </td>
                                                    <td class="TableHeader" align="center">
                                                        <b>Enrollment</b>
                                                    </td>
                                                    <td class="TableHeader" align="center">
                                                        <b>Core<br>
                                                            Units</b>
                                                    </td>
                                                    <td class="TableHeader" align="center">
                                                        <b>Elective<br>
                                                            Units</b>
                                                    </td>
                                                    <td class="TableHeader" align="center">
                                                        <b>Class<br>
                                                            Time</b>
                                                    </td>
                                                    <td class="TableHeader" align="center">
                                                        <b>Contract<br>
                                                            Hrs</b>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td class="TableHeader" align="center">
                                                        <b>Goal</b>
                                                    </td>
                                                    <td class="TableCell" valign="middle" align="center" nowrap>
                                                        <% if oBudget.PercentEnrolledLocked <> "" then 
																response.Write oBudget.PercentEnrolledLocked
															else
																response.Write oBudget.PlannedEnrollment
															end if%>%
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
                                                    <td class="TableHeader" align="center">
                                                        <b>Achieved</b>
                                                    </td>
                                                    <td class="<% 
												if oBudget.ActualEnrollment < oBudget.PlannedEnrollment then 
													response.Write "ErrorCell" 													
												else 
													response.Write "TableCell" 
												end if
													   %>"
                                                        valign="middle" align="center">
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
													   %>"
                                                        align="center">
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
												%>"
                                                        align="center">
                                                        <b>
                                                            <% = round(oBudget.ElectiveUnits,1) %></b>
                                                    </td>
                                                    <td class="<%
												if oBudget.TotalHours < oBudget.GoalClassTime then 
													response.Write "ErrorCell" 
												else 
													response.Write "TableCell" 
												end if
													  %>"
                                                        align="center">
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
													  %>"
                                                        align="center">
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
                    'oBudget.PopulateFamilyBudgetInfo oFunc.FpcsCnn, oBudget.FamilyId,session.Contents("intSchool_Year") 
                    oBudget.PopulateFamilyBudgetInfoApplication("cnnFPCS"), oBudget.FamilyId,session.Contents("intSchool_Year") 
                    %>
                            <td valign="top" style="width: 50%;">
                                <table cellpadding="4" style="width: 100%;" id="Table47">
                                    <tr>
                                        <td class="TableHeader" align="center">
                                            <b>*Family Elective<br>
                                                Spending Limits </b>
                                        </td>
                                        <td class="TableHeader" align="center">
                                            <b>Family Budget<br>
                                                Limit</b>
                                        </td>
                                        <td class="TableHeader" align="center">
                                            <b>Family Amount<br>
                                                Budgeted</b>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="TableHeader">
                                            <b>Family Budget</b>
                                        </td>
                                        <td class="TableCell" align="right">
                                            $<% = formatNumber(oBudget.FamilyBudgetFunding,2) %>
                                        </td>
                                        <td class="TableCell" align="right">
                                            $<% = formatNumber(oBudget.FamilyElectiveBudget,2) %>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td colspan="3" class="svplain7">
                                            <b>*</b> Each family can not spend more than 50% of their students combined budgets
                                            on Music, Art and/or P.E. classes.
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
                    <table id="Table48" cellspacing="2">
                        <tr>
                            <td bgcolor="white" colspan="11">
                                <span class="svplain10"><b>Sponsor Teacher:</b>&nbsp;<% = oBudget.SponsorName%>&nbsp;</span>
                            </td>
                            <td class="TableHeader" align="center">
                                Budget
                            </td>
                            <td class="TableHeader" align="center">
                                Spent
                            </td>
                        </tr>
                        <tr>
                            <td rowspan="4" colspan="8" class="svplain8" valign="bottom">
                                <%
							if oBudget.TSTestingSigned < 0 then
								packetHelper = packetHelper & "<li>ASD Testing Agreement must be signed.  " & _
											   "</li>"
							end if
							
							if not oBudget.IsProgressSigned then
								packetHelper = packetHelper & "<li>Progress Report Agreement must be signed.  " & _
											   "</li>"
							end if
														
							if not oBudget.IsPhilosophyFilled then
								packetHelper = packetHelper & "<li>Must provide an ILP Philosophy. " & _
											   "</li>"
							end if
								
							if not oBudget.HasSponsorCourse then
								packetHelper = packetHelper & "<li>Packet must include an ASD Sponsor/Oversight class with at least 1 contract hour.</li>"
							end if
													
							If packetHelper <> "" then
						%>
                                <table class="svplain8" id="Table49">
                                    <tr>
                                        <td style='width: 140px;' class="TableHeader">
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
                                <table class="svplain8" id="Table50">
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
                                <table class="svplain8" id="Table51">
                                    <tr>
                                        <td style='width: 140px;' class="TableHeaderBlue">
                                            &nbsp;<b>Packet Helper</b>
                                        </td>
                                        <td>
                                            Almost there. Be sure all parties have signed off on each course. The final step
                                            will be completed after the entire Packet has been approved by the Academic Advisor.
                                        </td>
                                    </tr>
                                </table>
                                <%
							end if
						%>
                                <br>
                            </td>
                            <td align="right" class="svplain8" colspan="2">
                                Beginning Balance:
                            </td>
                            <td bgcolor="white" style="width: 0%;">
                                &nbsp;
                            </td>
                            <td class="TableCell" align="right">
                                $<%=formatNumber(oBudget.BasePlannedFunding,2)%>
                            </td>
                            <td class="TableCell" align="right">
                                $<%=formatNumber(oBudget.BaseActualFunding,2)%>
                            </td>
                        </tr>
                        <tr>
                            <td align="right" class="svplain8" colspan="2">
                                Budget Transfer Deposits:
                            </td>
                            <td bgcolor="white" style="width: 0%;">
                                &nbsp;
                            </td>
                            <td class="TableCell" align="right">
                                <nobr>$<%=formatNumber(oBudget.Deposits,2)%></nobr>
                            </td>
                            <td class="TableCell" align="right">
                                <nobr>$<%=formatNumber(oBudget.Deposits,2)%></nobr>
                            </td>
                        </tr>
                        <tr>
                            <td align="right" class="svplain8" colspan="2">
                                Budget Transfer Withdrawals:
                            </td>
                            <td bgcolor="white" style="width: 0%;">
                                &nbsp;
                            </td>
                            <td class="TableCell" align="right">
                                <nobr>- $<%=formatNumber(oBudget.Withdrawls,2)%></nobr>
                            </td>
                            <td class="TableCell" align="right">
                                <nobr>- $<%=formatNumber(oBudget.Withdrawls,2)%></nobr>
                            </td>
                        </tr>
                        <tr>
                            <td align="right" class="svplain8" nowrap colspan="2">
                                Available Remaining Funds:
                            </td>
                            <td bgcolor="white" style="width: 0%;">
                                &nbsp;
                            </td>
                            <td class="TableCell" align="right">
                                <nobr>$<%=formatNumber(myBudgetBalance,2)%></nobr>
                            </td>
                            <td class="TableCell" align="right">
                                <nobr>$<%=formatNumber(myActualBalance,2)%></nobr>
                            </td>
                        </tr>
                        <%

	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Get student Information
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	sql = "SELECT     ISF.szCourse_Title, POS.txtCourseTitle, ISF.intShort_ILP_ID, I.szName, tblILP.intILP_ID, tblILP.bolApproved AS aStatus,  " & _ 
		"                      tblILP.bolSponsor_Approved AS sStatus, oi.bolApproved, oi.bolSponsor_Approved,  " & _ 
		"                      CASE isNull(tblClasses.intPOS_Subject_ID,1) when 1 then case ISF.intPOS_Subject_ID WHEN 22 THEN 0 ELSE 1 END ELSE case tblClasses.intPOS_Subject_ID WHEN 22 THEN 0 ELSE 1 END END AS isSponsor, oi.intQty, oi.curUnit_Price, oi.curShipping, ISF.intCourse_Hrs,  " & _ 
		"                      tblILP.decCourse_Hours, oi.intQty * oi.curUnit_Price + oi.curShipping AS total, oi.intOrdered_Item_ID, tblClasses.intInstructor_ID, " & _
		"	CASE isNull(tps2.szSubject_Name,'a') when 'a' then tps.szSubject_Name else tps2.szSubject_Name end as szSubject_Name,  " & _ 
		"                      tblClasses.intClass_ID, tblClasses.intInstruct_Type_ID, tblILP.intContract_Guardian_ID, tblClasses.intGuardian_ID, tblClasses.intVendor_ID,  " & _ 
		"                      tblClasses.szClass_Name, CASE WHEN tblClasses.intInstructor_ID IS NOT NULL  " & _ 
		"                      THEN ins.szFirst_Name + ' ' + ins.szLast_Name WHEN tblClasses.intGuardian_ID IS NOT NULL  " & _ 
		"                      THEN g.szFirst_Name + ' ' + g.szLast_Name END AS teacherName, tblILP.szAdmin_Comments, tblILP.szSponsor_Comments,  " & _ 
		"                      tblILP.bolReady_For_Review, tblILP.dtReady_For_Review, " & _ 
		"                          (SELECT     TOP 1 oa2.szValue " & _ 
		"                            FROM          tblOrd_Attrib oa2 " & _ 
		"                            WHERE      oa2.intOrdered_Item_Id = oi.intOrdered_Item_Id AND (oa2.intItem_Attrib_ID = 9 OR " & _ 
		"                                                   oa2.intItem_Attrib_ID = 5 OR " & _ 
		"                                                   oa2.intItem_Attrib_ID = 6 OR " & _ 
		"                                                   oa2.intItem_Attrib_ID = 22 OR " & _ 
		"                                                   oa2.intItem_Attrib_ID = 33) " & _ 
		"                            ORDER BY oa2.intOrd_Attrib_ID) AS oiDesc, oi.bolClosed, oi.bolReimburse, I.intItem_Group_ID, oi.szDeny_Reason, tblVendors.szVendor_Name,  " & _ 
		"                      tblVendors.szVendor_Phone, tblVendors.szVendor_Fax, tblVendors.szVendor_Email, tblVendors.szVendor_Website, oi.dtCREATE AS oiCreate,  " & _ 
		"                      DM_TEACHER_CLASS_COST.TeacherCostPerStudent, DM_TEACHER_RATES.HourlyRateTaxBen,  " & _ 
		"                      DM_TEACHER_CLASS_COST.HoursChargedPerStudent, " & _ 
		"					   tblILP.GuardianStatusId, tblILP.SponsorStatusId,tblILP.InstructorStatusId,tblILP.AdminStatusId," & _
		"					   tblILP.GuardianStatusDate,tblILP.SponsorStatusDate,tblILP.InstructorStatusDate, tblILP.AdminStatusDate, " & _
		"					   tblILP.GuardianComments, tblILP.InstructorComments, " & _
		"					   (select u.szName_First + ' ' + u.szName_Last from tblUsers u where u.szUser_ID = tblILP.GuardianUser) as GuardianUser, " & _
		"					   (select u.szName_First + ' ' + u.szName_Last from tblUsers u where u.szUser_ID = tblILP.SponsorUser) as SponsorUser, " & _
		"					   (select u.szName_First + ' ' + u.szName_Last from tblUsers u where u.szUser_ID = tblILP.InstructorUser) as InstructorUser, " & _
		"					   (select u.szName_First + ' ' + u.szName_Last from tblUsers u where u.szUser_ID = tblILP.AdminUser) as AdminUser," & _
		"					   (select u.szName_First + ' ' + u.szName_Last from tblUsers u where u.szUser_ID = tblClasses.szUser_Approved) as AdminUser2," & _					    
		"						tblClasses.intInstructor_ID,tblClasses.intContract_Status_ID, tblClasses.dtApproved, tblClasses.szUser_Approved, tblILP.bolSponsorAlert, tblILP.bolParentAlert, " & _
		"			  CASE isNull(tblClasses.szClass_Name,'a') WHEN 'a' then CASE isNull(POS.txtCourseTitle,'a') WHEN 'a' then ISF.szCourse_Title  else POS.txtCourseTitle end else tblClasses.szClass_Name end as ClassLabel, " & _
		"			  tblClasses.szASD_COURSE_ID, POS.txtCourseNbr " & _
		"FROM         tblClasses INNER JOIN " & _ 
		"                      tblILP ON tblClasses.intClass_ID = tblILP.intClass_ID LEFT OUTER JOIN " & _ 
		"                      trefItems I INNER JOIN " & _ 
		"                      tblOrdered_Items oi ON I.intItem_ID = oi.intItem_ID ON tblILP.intILP_ID = oi.intILP_ID RIGHT OUTER JOIN " & _ 
		"                      tblILP_SHORT_FORM ISF ON tblILP.intShort_ILP_ID = ISF.intShort_ILP_ID LEFT OUTER JOIN " & _ 
		"                      tblProgramOfStudies POS ON ISF.lngPOS_ID = POS.lngPOS_ID INNER JOIN " & _ 
		"                      trefPOS_Subjects tps ON tps.intPOS_Subject_ID = ISF.intPOS_Subject_ID LEFT OUTER JOIN " & _ 
		"                      trefPOS_Subjects tps2 ON tps2.intPOS_Subject_ID = tblClasses.intPOS_Subject_ID LEFT OUTER JOIN " & _ 
		"                      DM_TEACHER_RATES ON tblClasses.intInstructor_ID = DM_TEACHER_RATES.InstructorId AND  " & _ 
		"                      DM_TEACHER_RATES.StartSchoolYear = " & session.Contents("intSchool_Year") & " LEFT OUTER JOIN " & _ 
		"                      DM_TEACHER_CLASS_COST ON tblClasses.intClass_ID = DM_TEACHER_CLASS_COST.ClassId LEFT OUTER JOIN " & _ 
		"                      tblVendors ON oi.intVendor_ID = tblVendors.intVendor_ID LEFT OUTER JOIN " & _ 
		"                      tblINSTRUCTOR INS ON tblClasses.intInstructor_ID = INS.intINSTRUCTOR_ID LEFT OUTER JOIN " & _ 
		"                      tblGUARDIAN g ON tblClasses.intGuardian_ID = g.intGUARDIAN_ID  " & _
		"WHERE     (ISF.intStudent_ID = " & intStudent_ID & ") AND (ISF.intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _ 
		"ORDER BY isSponsor, ClassLabel, ISF.intShort_ILP_ID "

set rsBudget = server.CreateObject("ADODB.RECORDSET")
rsBudget.CursorLocation = 3
rsBudget.Open sql,Application("cnnFPCS")'oFunc.FPCScnn

intPreviousID = 0

if rsBudget.RecordCount < 1 then
%>
                        <tr>
                            <td colspan="13" align="center" class="svplain10">
                                <br>
                                <b>No courses have been planned yet. To get started click the 'Plan New Course' button
                                    above.</b>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <%
	rsBudget.Close
	set rsBudget = nothing
	set oBudget = nothing
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
				strClassHeader = "TableHeader" '"TableHeaderBlack"
				CourseHelper = " This course has been rejected.  The Guardian or the Sponsor must delete this course. The funds budgeted by this course will not be released until the course is deleted."			
		elseif  rsBudget("AdminStatusId")  = "2" or rsBudget("SponsorStatusId") = "2" then
			' Needs Work
			strClassHeader = "TableHeader" '"TableHeaderRed"
			CourseHelper = " This course needs work before it can be signed off on. Please fix any problems and re-sign the contract after any issues have been resolved."
		elseif rsBudget("intILP_ID") & "" = "" then
			strClassHeader = "TableHeader" '"SubHeader"
			CourseHelper = " This course is in the <b>planned stage</b>. The next step is to implement the plan.  This can be done by selecting 'Implement Plan' under 'Actions' and then click the 'go' button."
		elseif rsBudget("GuardianStatusId") & "" <> "1" or rsBudget("SponsorStatusId") & "" <> "1" or _
			(rsBudget("AdminStatusId") & "" <> "1" and rsBudget("intContract_Status_Id") & "" <> "5") or _
			(rsBudget("intInstructor_ID") & "" <> "" and rsBudget("intInstructor_ID") & "" <> oBudget.SponsorId & "" and rsBudget("InstructorStatusId") & ""  <> "1") then
			 strClassHeader = "TableHeader" '"TableheaderBlue"
			 CourseHelper = " This course has not yet been signed by all parties. In order for this course to be complete all parties must sign."
		else
			strClassHeader = "TableHeader" '"TableHeaderGreen"			
			CourseHelper = "Congratulations! This course has been approved."
		end if 				
		
		if rsBudget("bolSponsorAlert") then
			strClassHeader = "TableHeader" '"TableHeaderGrape"
		end if 
		
		if rsBudget("bolParentAlert") then
			strClassHeader = "TableHeader" '"TableHeaderTeal"
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
		
		if rsBudget("szClass_Name") & "" = "" then 
			if rsBudget("txtCourseTitle") & "" <> "" then
				myClassName = replace(rsBudget("txtCourseTitle"),"'","\'")
			else
				myClassName = replace(rsBudget("szCourse_Title"),"'","\'")
			end if
		else 
			myClassName = replace(rsBudget("szClass_Name"),"'","\'")
		end if
		
		if rsBudget("szClass_Name") & "" <> "" then
			myClassName = replace(replace(rsBudget("szClass_Name"),"'","\'"),"""","")
		end if
		'response.Write szClass_Name & "<<<"
%>
        <tr>
            <td colspan="11" style="width: 100%;">
                <table style="width: 100%;" cellpadding='2' cellspacing='1' id="Table2">
                    <tr class="<% = strClassHeader %>" <% if mDivcount > 1 then response.Write "id=""div" & mDivCount & """"%>>
                        <td align="left" style="width: 50%;">
                            &nbsp;<b>Course Title</b>
                        </td>
                        <td align='center' style="width: 30%;">
                            <b>Subject</b>
                        </td>
                        <td align='center' nowrap style="width: 0%;">
                            &nbsp;<b>Hrs</b>&nbsp;
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
                    </tr>
                    <% 
							mDivCount = mDivCount + 1
							strBList = strBList & mDivCount & ","
							strSmallList = mDivCount & ","
							%>
                    <tr id="div<% = mDivCount%>">
                        <td colspan="3">
                            <% if  rsBudget("intILP_ID") & "" <> "" then  %>
                            <table style="width: 100%;" cellspacing="1" cellpadding="0" id="Table16">
                                <tr class="svplain">
                                    <td valign="middle" rowspan="2" align="center" class="TableCell" valign="middle"
                                        style="width: 130px;">
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
                                    <td align="center">
                                        Admin
                                    </td>
                                </tr>
                                <tr class="svplain">
                                    <td align="center">
                                        <% if rsBudget("GuardianStatusId") & "" = "" then%>
                                        not signed
                                        <% else %>
                                        <span title="signed on: <% = rsBudget("GuardianStatusDate")%>">
                                            <% = rsBudget("GuardianUser")%></span>
                                        <% end if %>
                                    </td>
                                    <td valign="middle" align="center">
                                        <% 
												if rsBudget("SponsorStatusId") & ""  = "1" then %>
                                        <span title="signed on: <% = rsBudget("SponsorStatusDate")%>">
                                            <% = rsBudget("SponsorUser")%></span>
                                        <% else
													response.Write InterpretStatus(rsBudget("SponsorStatusId"))
												end if%>
                                    </td>
                                    <% if rsBudget("intInstructor_ID") & "" <> "" and rsBudget("intInstructor_ID") & "" <> oBudget.SponsorId & "" THEN %>
                                    <td valign="middle" align="center">
                                        <% 												
													if rsBudget("InstructorStatusId") & ""  = "1" then %>
                                        <span title="signed on: <% = rsBudget("InstructorStatusDate")%>">
                                            <% = rsBudget("InstructorUser")%></span>
                                        <% else
														response.Write InterpretStatus(rsBudget("InstructorStatusId"))
													end if%>
                                    </td>
                                    <% end if %>
                                    <td valign="middle" align="center">
                                        <% if rsBudget("intInstructor_ID") & "" <> "" and rsBudget("intContract_Status_ID") & "" = "5" then 
													' This is ASD course is pre-approved via the principal class approval admin
												%>
                                        <span title="signed on: <% = rsBudget("dtApproved")%>">
                                            <% = rsBudget("AdminUser2")%></span>
                                        <% elseif rsBudget("AdminStatusId") & "" = "1" then 
													' Signed Schedule
												%>
                                        <span title="signed on: <% = rsBudget("AdminStatusDate")%>">
                                            <% = rsBudget("AdminUser") %></span>
                                        <% else 
													response.Write InterpretStatus(rsBudget("AdminStatusId"))												
												 end if %>
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
                    <input type="hidden" name="MustAmend<% = rsBudget("intILP_ID")%>" value="1" id="Hidden2">
                    <%
							end if 
							
							%>
                    <tr id="div<% = mDivCount%>">
                        <td colspan="3" style="width: 100%;">
                            <table style="width: 100%;" cellpadding="0" cellspacing="1" id="Table28">
                                <% if rsBudget("intILP_ID") & "" <> "" then %>
                                <tr>
                                    <td valign="top" style="width: 100%;" align="center">
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
																			"<td class='TableCell' style='width:130px;' align='center' valign='top'><b>Admin Comments</b></td>" & _
																			"<td class='TableCell' >" & rsBudget("szAdmin_Comments") & "</td></tr>"
													end if
													
													if rsBudget("szSponsor_Comments") & "" <> "" then
														strCommentTable = strCommentTable & "<tr>" & _
																			"<td class='TableCell' style='width:130px;' align='center' valign='top'><b>Sponsor Comments</b></td>" & _
																			"<td class='TableCell'>" & rsBudget("szSponsor_Comments") & "</td></tr>"
													end if
													
													if rsBudget("InstructorComments") & "" <> "" then
														strCommentTable = strCommentTable & "<tr>" & _
																			"<td class='TableCell' style='width:130px;' align='center' valign='top'><b>Instructor Comments</b></td>" & _
																			"<td class='TableCell'>" & rsBudget("InstructorComments") & "</td></tr>"
													end if
													
													if rsBudget("GuardianComments") & "" <> "" then
														strCommentTable = strCommentTable & "<tr>" & _
																			"<td class='TableCell' style='width:130px;' align='center' valign='top'><b>Guardian Comments</b></td>" & _
																			"<td class='TableCell'>" & rsBudget("GuardianComments") & "</td></tr>"
													end if
													
													strCommentTable = strCommentTable & "<tr >" & _
																	"<td  class='TableCell' style='width:130px;background-color:#F0F0F0;' align='center' valign='middle'>&nbsp;<b>Course Helper</b></td>" & _
																	"<td class='TableCell' >" & CourseHelper & "</td></tr>"
																									
													strCommentTable = "<table cellpadding='2' style='width:100%;'>" & strCommentTable & "</table>"																			
												%>
                                        <% = strCommentTable %>
                                    </td>
                                </tr>
                                <% else
											response.write "<tr >" & _
														   "<td  class='TableCell' style='width:130px;background-color:#F0F0F0;' align='center' valign='middle'>&nbsp;<b>Course Helper</b></td>" & _
														   "<td class='TableCell' >" & CourseHelper & "</td></tr>"
										 end if %>
                            </table>
                        </td>
                    </tr>
                </table>
                <nobr>
            </td>
            <td class="ltGray" colspan="2" style="width: 0%;">
                &nbsp;
            </td>
        </tr>
        <% 
					mDivCount = mDivCount + 1
					strBList = strBList & mDivCount & ","	
					strSmallList = strSmallList & mDivCount & ","
				%>
        <tr id="div<% = mDivCount%>">
            <td class="TableSubHeader" align="center">
                Budget Item
            </td>
            <td class="TableSubHeader" align="center">
                Status
            </td>
            <td class="TableSubHeader" style="width: 100%;">
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
            <td style="width: 0%;" class="TableSubHeader" align="center" title="Adjustments are needed to handle over expendatures and to release unused budgeted funds once the budget is closed.">
                Budget Adjust
            </td>
            <td style="width: 0%;" class="TableSubHeader" align="center" title="(Budget Total - Actual Charges) + Budget Adjust">
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
			dblClassCharge = round(cdbl(rsBudget("TeacherCostPerStudent")),2)
			dblClassBudget = round(cdbl(rsBudget("TeacherCostPerStudent")),2)
				%>
        <tr id="div<%=mDivCount%>">
            <td class="<% = strClass %>">
                Instruction
            </td>
            <td class="<% = strClass %>" align="center">
                n/a
            </td>
            <td class="<% = strClass %>" style="width: 100%;">
                Instruction by:
                <% = rsBudget("teacherName") %>
            </td>
            <td class="<% = strClass %>" align="center" nowrap>
                <%= round(rsBudget("HoursChargedPerStudent"),3)%>
            </td>
            <td class="<% = strClass %>" align="right" title="Teachers Hourly Rate" nowrap>
                $<%= formatNumber(round(rsBudget("HourlyRateTaxBen"),3),3)%>
            </td>
            <td class="<% = strClass %>" align="center">
                n/a
            </td>
            <td class="<% = strClass %>" align="right" nowrap>
                $<%= formatNumber(round(rsBudget("TeacherCostPerStudent"),2),2)%>
            </td>
            <td class="<% = strClass %>" align="right" nowrap>
                $<%= formatNumber(round(rsBudget("TeacherCostPerStudent"),2),2)%>
            </td>
            <td style="width: 0%;" class="<% = strClass %>" align="right" nowrap>
                $0.00
            </td>
            <td style="width: 0%;" class="<% = strClass %>" align="right" nowrap>
                $0.00
            </td>
            <td class="ltGray" style="width: 0%;">
                &nbsp;
            </td>
            <td class="<% = strClass %>" align="right" nowrap>
                -$<%= formatNumber(round(rsBudget("TeacherCostPerStudent"),2),2)%>
            </td>
            <td class="<% = strClass %>" align="right" nowrap>
                -$<%= formatNumber(round(rsBudget("TeacherCostPerStudent"),2),2)%>
            </td>
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
		'liInfo = LineItemInfo(rsBudget("intOrdered_Item_ID"),dblBudgetCost, rsBudget("bolClosed"), oFunc.FPCScnn,strClass)
		liInfo = LineItemInfo(rsBudget("intOrdered_Item_ID"),dblBudgetCost, rsBudget("bolClosed"), Application("cnnFPCS"),strClass)
		bStatus = GetBudgetStatus(rsBudget("intItem_Group_ID"),rsBudget("bolApproved"),liInfo(4),rsBudget("bolReimburse"))		
		
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
            <td class="<% = strClass %>">
                <% = rsBudget("szName") %>
            </td>
            <td class="<% = strClass %>" align="center">
                <% = bStatus %>
            </td>
            <td class="<% = strClass %>" style="width: 100%;">
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
            <td style="width: 0%;" class="<% = strClass %>" align="right" nowrap title="Adjustments are needed to handle over expendatures and to release unused budgeted funds once the budget is closed.">
                $<% = formatNumber(liInfo(2),2) %>
            </td>
            <td style="width: 0%;" class="<% = strClass %>" align="right" nowrap title="(Budget Total - Actual Charges) + Budget Adjust">
                $<% = formatNumber((dblBudgetCost - cdbl(liInfo(1))) + cdbl(liInfo(2)),2)%>
            </td>
            <td bgcolor="white" style="width: 0%;">
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
            <td class="svplain10" colspan="10" align="right">
                Available Remaining Funds:
            </td>
            <td bgcolor="white" style="width: 0%;">
                &nbsp;&nbsp;&nbsp;
            </td>
            <td class="TableHeader" align="right">
                $<%=formatNumber(dblTargetBalance,2)%>
                <input type="hidden" name="budgetBalance" value="<%=formatNumber(dblTargetBalance,2)%>"
                    id="Hidden3">
            </td>
            <td class="TableHeader" align="right">
                $<%=formatNumber(dblActualBalance,2)%>
            </td>
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
    </table>
    </td> </tr>
</table>
</form>
<%
response.Write oHtml.ToolTipDivs
set oHtml = nothing
set oBudget = nothing
end function
''''''''''''''''''''''''''''''
' END PACKET
''''''''''''''''''''''''''''''

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
    <td colspan="10" align="right" class="svplain10">
        <b>Course Totals:</b>
    </td>
    <td bgcolor="white" style="width: 0%;">
        &nbsp;&nbsp;&nbsp;
    </td>
    <td class="TableHeader" align="right">
        <nobr>
						<% if instr(1,dblBudgetCost,"-") > 0 then
								response.Write "+ $" & formatNumber(replace(dblClassBudget,"-",""),2)
						   else
								response.Write "- $" & formatNumber(dblClassBudget,2)
						   end if						
						%></nobr>
    </td>
    <td class="TableHeader" align="right">
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
<tr bgcolor="white" id="div<% = mDivCount%>">
    <td colspan="13">
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

dim strTestTable

function vbForms(pFormName)
	call vbfMakeTestTable
	
	if ucase(request("strAction")) = "A" and request("intStudent_ID") <> "" then
		strField = "intStudent_ID " 
		intFamilyID = request("intStudent_ID")
	elseif session.Contents("intFamily_ID") <> "" then
		intFamilyID = session.Contents("intFamily_ID")
		strField = "intFamily_ID " 
	elseif request("intStudent_ID") <> "" then
		intFamilyID = oFunc.StudentInfo(request("intStudent_ID"),"6")
		strField = "intFamily_ID " 
	end if

	if intFamilyID <> "" then
		sql = "SELECT s.szFIRST_NAME, s.szLAST_NAME " & _ 
				"FROM tblSTUDENT s INNER JOIN " & _ 
				" tblENROLL_INFO ei ON s.intSTUDENT_ID = ei.intSTUDENT_ID " & _ 
				"WHERE (ei.sintSCHOOL_YEAR = " & session.Contents("intSchool_Year") & _
				") AND (s." & strField & " = " & intFamilyID & ") " & _
				" and (ei.bolASD_Testing IS NULL OR ei.bolASD_Testing = 0) "

		dim rs1
		set rs1 = server.CreateObject("ADODB.RECORDSET")
		rs1.CursorLocation = 3	
		rs1.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
		
		if rs1.RecordCount > 0 then		
			do while not rs1.EOF
				if pFormName = "Testing" then
					response.Write vbfTestingForm(ucase(rs1(0) & " " & rs1(1))) 
				else
					response.Write vbfProgressForm(ucase(rs1(0) & " " & rs1(1))) 
				end if							
				rs1.MoveNext
				if not rs1.EOF then response.Write "<p></p>"
			loop
		else
			if pFormName = "Testing" then
				response.Write vbfTestingForm("") 
			else
				response.Write vbfProgressForm("") 
			end if	
		end if
		rs1.Close
		set rs1 = nothing
	else
		if pFormName = "Testing" then
			response.Write vbfTestingForm("") 
		else
			response.Write vbfProgressForm("") 
		end if
	end if
	
	if ucase(request("strAction")) = "A" and pFormName = "Testing" then
		response.Write "<p></p>"
	end if
	
end function

function vbfTestingForm(pStudentName)
	if pStudentName = "" then pStudentName = "________________________"
%>
<table style="width: 100%;">
    <%
		vbfFormHeader("ASD Required Testing Agreement")											
%>
    <tr class="svplain11">
        <td colspan="2">
            All students enrolled in FPCS are required to participate in the applicable Anchorage
            School District mandatory testing listed below.
            <br>
            <br>
        </td>
    </tr>
    <tr>
        <td colspan="2">
            <% = strTestTable %>
            <font class="svplain8">Times and locations to be announced later. Sponsor teachers will
                have access to test results.
                <br>
                <br>
            </font>
        </td>
    </tr>
    <tr>
        <td>
            <pre style="font-family: Veranda,Tahoma; font-size: 10pt; font-weight: bolder;">
By signing below I agree that <% = pStudentName %> will participate
in all Anchorage School District required testing.




________________________________________________  Date:___________
Guardian Signature                             




________________________________________________
Print Guardian Name	
</pre>
        </td>
    </tr>
</table>
<%
end function

function vbfMakeTestTable()

	strTestTable = "<table cellspacing=0 cellpadding=2 border=1>" & _
				   "<tr class=""gray"">" & _
				   "	<td>" & _
					"		Test" & _
					"	</td>" & _
					"	<td>" & _
					"		Dates" & _
					"	</td>" & _
					"	<td>" & _
					"		Grade Level" & _
					"	</td>" & _
					"	<td>" & _
					"		Notes" & _
					"	</td>" & _
					"</tr>" 
									 
			sql = "select strTest_Name, strTesting_Dates, strGrade_Level,strTest_Desc " & _
					"from tblTesting_Info " & _
					"WHERE intSchool_Year = " & session.Contents("intSchool_Year") & _
					" order by 1"
			dim rs
			set rs = server.CreateObject("ADODB.RECORDSET")
			rs.CursorLocation = 3
			rs.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
			
			if rs.RecordCount > 0 then
				do while not rs.EOF
					strTestTable = strTestTable & "<tr class=""svplain"">" & _
								   "<td valign=top>" & _
									rs(0) & _	
									"&nbsp;</td><td valign=top>" & rs(1) & "&nbsp;</td>" & _
									"<td valign=top>" & _
									rs(2) & _	
									"&nbsp;</td>" & _
									"<td valign=top>" & _
									rs(3) & _	
									"&nbsp;</td>" & _
									"</tr>"
					rs.MoveNext
				loop
			end if
			rs.Close
			set rs = nothing
		strTestTable = strTestTable & "</table>"
end function

function vbfProgressForm(pStudentName)
	if pStudentName = "" then pStudentName = "________________________"
%>
<table width="100%" id="Table36">
    <tr>
        <td align="left">
            <img src="<% = Application("strImageRoot")%>fpcsLogo.gif">
        </td>
        <td align="right" class="svplain10" nowrap>
            <% = Application.Contents("SchoolAddress") %>
        </td>
    </tr>
    <tr class="yellowHeader">
        <td colspan="2">
            <table align="right" id="Table37">
                <tr>
                    <td align="right">
                        <font face="arial" size="2" color="white">
                            <% = date()%></font>
                    </td>
                </tr>
            </table>
            &nbsp;<b>Mandatory Student Progress Reports Agreement</b>
        </td>
    </tr>
    <tr class="svplain11">
        <td colspan="2">
            Parent/sponsor teacher communication is a critical component of our charter school.
            Mandatory progress reports of all courses are to be completed and turned in to the
            FPCS office at a minimum of twice a year. Reports are to be submitted between Nov.
            29 to Dec. 10 and March 7 to March 18.<br>
            <br>
            Signatures of parents and sponsor teacher are required.<br>
        </td>
    </tr>
    <tr>
        <td>
            <pre style="font-family: Veranda,Tahoma; font-size: 10pt; font-weight: bolder;">
I am aware that progress report forms for <% = pStudentName %> 
will be available on-line in the printable forms link of the 
Student On-line System.




________________________________________________  Date:___________
Guardian Signature                             




________________________________________________  Date:___________
Sponsor Teacher Signature 
</pre>
        </td>
    </tr>
</table>
<%
end function

function Philiosophy(pStudentID)
	dim sql 
	sql = "SELECT tblPhilosophy.szPhilosophy " & _ 
			"FROM tblENROLL_INFO INNER JOIN " & _ 
			" tblPhilosophy ON tblENROLL_INFO.intPHILOSOPHY_ID = tblPhilosophy.intPhilosophy_ID " & _ 
			"WHERE (tblENROLL_INFO.intSTUDENT_ID = " & pStudentID & ") AND (tblENROLL_INFO.sintSCHOOL_YEAR = " & session.Contents("intSchool_Year") & ") "
	set rsP = server.CreateObject("ADODB.RECORDSET")
	rsP.CursorLocation = 3
	rsP.Open sql,Application("cnnFPCS")'oFunc.FPCScnn
	%>
<table style='width: 100%'>
    <% = vbfFormHeader("ILP Philosophy for " & oFunc.StudentInfo(pStudentID,3)) %>
    <%
	if rsP.RecordCount > 0 then
		response.Write "<tr><td class='svplain8' colspan='2'><BR>" & rsP(0) & "</td></tr>"
	else
		response.Write "<tr><td class='svplain8' colspan='2'>No ILP Philosophy Defined.</td></tr>"
	end if
	response.Write "</table>"
	rsP.Close
	set rsP = nothing
	if ucase(request("strAction")) = "A" then
		response.Write "<p></p>"
	end if
end function 

function VendorServiceReport()	
%>
    <table style="width: 650px; height: 100%;" id="Table3">
        <%

dim strWhere
if request("hdnVendors") <> "" then 
	strVenList = right(request("hdnVendors"),len(request("hdnVendors"))-1)
	strVenList = left(strVenList,len(strVenList)-1)
	strWhere = " AND (tblVendors.intVendor_ID not in (" & strVenList & ")) "
end if

if request("hdnRange") <> "" then 
sRange=Replace(request("hdnRange"),"'',","")
    If sRange<>"" Then
	sRange = left(sRange,len(sRange)-1)
	strWhere = strWhere & " AND SUBSTRING(UPPER(tblVendors.szVendor_Name),1,1) IN (" & sRange & ")"
    End If
end if

sql = "SELECT tblVendors.szVendor_Name, tblSTUDENT.szLAST_NAME, tblSTUDENT.szFIRST_NAME,  " & _ 
		" SUM(tblOrdered_Items.intQty * tblOrdered_Items.curUnit_Price + tblOrdered_Items.curShipping) AS total, tblILP_SHORT_FORM.szCourse_Title,  " & _ 
		" tblProgramOfStudies.txtCourseTitle, tblILP.bolApproved, tblILP.bolSponsor_Approved, tblILP.bolReady_For_Review, " & _
		" tblILP.intILP_ID,tblILP.intClass_ID, tblOrdered_Items.intILP_ID " & _ 
		"FROM tblVendors INNER JOIN " & _ 
		" tblOrdered_Items ON tblVendors.intVendor_ID = tblOrdered_Items.intVendor_ID INNER JOIN " & _ 
		" tblSTUDENT ON tblOrdered_Items.intStudent_ID = tblSTUDENT.intSTUDENT_ID INNER JOIN " & _ 
		" tblILP ON tblOrdered_Items.intILP_ID = tblILP.intILP_ID INNER JOIN " & _ 
		" tblILP_SHORT_FORM ON tblILP.intShort_ILP_ID = tblILP_SHORT_FORM.intShort_ILP_ID LEFT OUTER JOIN " & _ 
		" tblProgramOfStudies ON tblILP_SHORT_FORM.lngPOS_ID = tblProgramOfStudies.lngPOS_ID  " & _
        "WHERE (tblOrdered_Items.intSchool_Year = " & session.Contents("intSchool_Year") & ") AND (tblOrdered_Items.intItem_ID = 3)  " & strWhere & " " & _ 
		"GROUP BY tblVendors.szVendor_Name, tblSTUDENT.szLAST_NAME, tblSTUDENT.szFIRST_NAME, tblILP_SHORT_FORM.szCourse_Title,  " & _ 
		" tblProgramOfStudies.txtCourseTitle, tblILP.bolApproved, tblILP.bolSponsor_Approved, " & _
		" tblILP.bolReady_For_Review, tblILP.intILP_ID,tblILP.intClass_ID, tblOrdered_Items.intILP_ID " & _ 
		"ORDER BY tblVendors.szVendor_Name, tblSTUDENT.szLAST_NAME, tblSTUDENT.szFIRST_NAME, tblILP_SHORT_FORM.szCourse_Title,  " & _ 
		" tblProgramOfStudies.txtCourseTitle "

	dim rs 
	set rs = server.CreateObject("ADODB.RECORDSET")
	dim rs2 
	set rs2 = server.CreateObject("ADODB.RECORDSET")
	dim strVendName
	dim dblGrandTotal
	dim IlpList
	dim StudentList
	dim ClassList
	dim szPO_Number
	
	intStudent_id = 1
	rs2.CursorLocation = 3
	rs.CursorLocation = 3
	rs.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
	strVendName = ""
	
	if rs.RecordCount > 0 then
		do while not rs.EOF
			if strVendName <> rs("szVendor_Name") then							
				if strVendName <> "" then
					response.Write vbfPONumber(szPO_Number) & " <td class='svplain8' align='right'><b>total:</b></td><td class='svplain8' align='right'><b>$" & _
								formatNumber(dblSubTotal,2) & "</b></td></tr></table>"
					dblGrandTotal = dblGrandTotal + dblSubTotal
				end if 					
				
				if IlpList <> "" then
					arList = split(IlpList,",")
					arSName = split(StudentList,",")
					arCName = split(ClassList,"~|")
					for i = 0 to ubound(arList)
						if arList(i) <> ""  then
							response.Write "<p></p>"
							strStudentName = arSName(i)
							szClass_Name = arCName(i)
							vbfPrintILP(arList(i))
							response.Write "<p></p>"
							intILP_ID = arList(i)							
							HideSigs = "TRUE"
							vbfGoodsServices()
							response.Write "<p></p>"
						end if
					next 									
				end if

				dblSubTotal = 0
				strVendName = rs("szVendor_Name")
				response.Write vbsTableHeader(rs("szVendor_Name"))
				IlpList = ""	
				StudentList = ""
				ClassList = ""
				szPO_Number = ""
			end if
					
%>
        <td class="svplain8">
            <% = rs("szCourse_Title") & rs("txtCourseTitle") %>
        </td>
        <td class="svplain8">
            <% = rs("szLAST_NAME") & ", " & rs("szFIRST_NAME")  %>
        </td>
        <td class="svplain8" align="right">
            $<% = formatNumber(rs("total"),2) %>
        </td>
        </tr>
        <%			
				if szPO_Number = "" then
					sql = "select szPO_Number " & _
						   " FROM tblLINE_ITEMS LEFT OUTER JOIN " & _
						   "	tblORDERED_ITEMS on tblORDERED_ITEMS.intORDERED_ITEM_ID =  tblLINE_ITEMS.intORDERED_ITEM_ID " & _
						   " WHERE tblORDERED_ITEMS.intILP_ID = " & rs("intILP_ID") 
					rs2.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
					if rs2.RecordCount > 0 then
						if rs2("szPO_Number") & "" <> "" then
							szPO_Number = rs2("szPO_Number")
						end if
					end if
					rs2.Close
				end if		
				dblSubTotal = dblSubTotal + rs("total")
				IlpList = IlpList & rs("intILP_ID") & ","	
				StudentList = StudentList & rs("szFIRST_NAME") & " " & rs("szLAST_NAME") & ","
				ClassList = ClassList & rs("szCourse_Title") & rs("txtCourseTitle") & "~|"
				intStudent_id = intStudent_id + 1
				'if intStudent_id > 25 then exit function						
			rs.MoveNext
		loop
	end if
	dblGrandTotal = dblGrandTotal + dblSubTotal
	response.Write vbfPONumber(szPO_Number) & " <td colspan='2' class='svplain8' align='right'><b>total:</b></td><td class='svplain8' align='right'><b>$" & _
					formatNumber(dblSubTotal,2) & "</b></td></tr></td></table>"
	if IlpList <> "" then
		arList = split(IlpList,",")
		arSName = split(StudentList,",")
		arCName = split(ClassList,"~|")
		for i = 0 to ubound(arList)
			if arList(i) <> ""  then
				response.Write "<p></p>"
				strStudentName = arSName(i)
				szClass_Name = arCName(i)
				vbfPrintILP(arList(i)) 
				response.Write "<p></p>"
				intILP_ID = arList(i)							
				HideSigs = "TRUE"
				vbfGoodsServices()
			end if
		next 					
	end if
%>
    </table>
    </td> </tr>
</table>
<%
	rs.Close
	set rs = nothing
	set rs2 = nothing
end function

function vbsTableHeader(pName)
%>
<tr>
    <td valign="top">
        <table>
            <% = vbfFormHeader("Vendor Service Status Report for School Year " & oFunc.SchoolYearRange) %>
        </table>
    </td>
</tr>
<tr>
    <td valign="top">
        <table id="Table32" style="width: 650px;">
            <tr>
                <td class="TableCell" style="width: 30%;">
                    Vendor Name
                </td>
                <td class="TableCell" style="width: 30%;">
                    Course Name
                </td>
                <td class="TableCell" style="width: 30%;">
                    Student Name
                </td>
                <td class="TableCell" style="width: 10%;">
                    Budget
                </td>
            </tr>
            <tr>
                <td class="svplain8" rowspan="100" valign="top">
                    <b>
                        <% = pName %></b><br>
                    <br>
                </td>
                <%
end function

function vbfPONumber(pNum)
%>
            <tr>
                <td class="svplain8" align="left">
                    <table border="1" cellspacing="0" id="Table33" bordercolor="#c0c0c0">
                        <tr>
                            <td class="svplain8">
                                PO#:</b>
                                <% if pNum <> "" then 
										response.Write "&nbsp;" & pNum
									   else
									%>
                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                <%
									  end if
									%>
                            </td>
                        </tr>
                    </table>
                </td>
                <%
end function 
%>
        