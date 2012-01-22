<%@ Language=VBScript %>
<%
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Make db Connection, print HTML header.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim oFunc
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()
   
dim insert
dim update
dim dtDeadline
dim dtClassStart
dim dtClassEnd
dim dtClass
dim strStartTime
dim strEndTime
dim intClass_ID
dim strAddField
dim strAddValue

oFunc.BeginTransCN

' Fusebox like logic for script steering. This if statement executes the needed functions
' based on incoming requests from the HTTP header
if Request.Form("edit") = "" and Request.Form("intClass_ID") = "" then
	call vbfInsertClass
else
	call vbfUpdateClass
end if 

function vbfInsertClass
	call vbfFormatDates
	
	' Should not have to reset these variables because they should be erased by default.asp.
	' We are doing it because we have noted some incosistances with the erasing of these variables.
	' Source of problem may be due to user not using the applications navigation and instead 
	' they are using browser back buttons to get to pages (which doesn't always send a request
	' to the server so session.abandon is not being fired.)
	
	session("intCharge_Type_ID") = ""
	session("curUnit_Cost") = ""
	session("decNum_Units") = ""
	
	if Request.Form("intInstructor_ID") <> "" then
		strAddField = "intInstructor_ID,dtReg_Deadline,sGrade_Level,sGrade_Level2,dtClass_Start,dtClass_End"
		strAddValue = "'" & oFunc.EscapeTick(request("intInstructor_ID")) & "'," & _
					  "'" & dtDeadline & "'," & _
					  "'" & oFunc.EscapeTick(request("sGrade_Level")) & "'," & _
					  "'" & oFunc.EscapeTick(request("sGrade_Level2")) & "'," & _
					  "'" & dtClassStart & "'," & _
					  "'" & dtClassEnd & "'," 
					  
		intInstruct_Type_Id = 4
	elseif Request.Form("intGuardian_id") <> "" then
		strAddField = "intGuardian_id,dtClass_Start,dtClass_End"
		strAddValue = "'" & oFunc.EscapeTick(request("intGuardian_id")) & "'," & _
					  "'" & Application.Contents("dtSchool_Year_Start" & session.Contents("intSchool_YEar")) & "'," & _
					  "'" & Application.Contents("dtSchool_Year_End" & session.Contents("intSchool_YEar")) & "',"
		intInstruct_Type_Id = 1	
	end if 
		
		
	insert = "insert into tblClasses(intPOS_Subject_ID," & strAddField & ",szClass_Name," & _
			 "szASD_Course_ID,szLocation,intMin_Students,intMax_Students," & _
			 "szDays_Meet_On,szStart_Time,szEnd_Time," & _
			 "szSchedule_Comments,decHours_Student,decHours_Planning,intInstruct_Type_id," & _
			 "decOriginal_Student_Hrs,decOriginal_Planning_Hrs,intDuration_ID," & _
			 "intSession_Minutes, szUSER_CREATE, intSchool_Year)" & _
			 "VALUES (" & _
			 "'" & oFunc.EscapeTick(request("intPOS_Subject_ID")) & "'," & _
			 strAddValue & _
			 "'" & UCase(oFunc.EscapeTick(request("szClass_Name"))) & "'," & _
			 "'" & UCase(oFunc.EscapeTick(request("szASD_Course_ID"))) & "'," & _
			 "'" & UCase(oFunc.EscapeTick(request("szLocation"))) & "'," & _			 
			 oFunc.CheckDecimal(request("intMin_Students")) & "," & _
			 oFunc.CheckDecimal(request("intMax_Students")) & "," & _			 
			 "'" & oFunc.EscapeTick(request("szDays_Meet_On")) & "'," & _
			 "'" & strStartTime & "'," & _
			 "'" & strEndTime & "'," & _
			 "'" & oFunc.EscapeTick(request("szSchedule_Comments")) & "'," & _
			 oFunc.CheckDecimal(request("decHours_Student")) & "," & _
			 oFunc.CheckDecimal(request("decHours_Planning")) & ","	& _
			 "'" & intInstruct_Type_id & "'," & _
			 oFunc.CheckDecimal(request("decHours_Student")) & "," & _			 
			 oFunc.CheckDecimal(request("decHours_Planning")) & "," & _
			 oFunc.CheckDecimal(request("intDuration_ID")) & "," & _
			 oFunc.CheckDecimal(request("intSession_Minutes")) & "," & _
			 "'" & Session.Value("strUserID")	& "'," & _
			 session.Contents("intSchool_Year") & ")"

			 
	oFunc.ExecuteCN(insert)	

	intClass_ID = oFunc.GetIdentity
	
	' Inserts the records that will limit a class to specific families
	if Request.Form("intFamily_ID") <> "" then
		arFamily = split(Request.Form("intFamily_ID"),",")
		if isArray(arFamily) then
			for iFam = 0 to ubound(arFamily)
				insert = "insert into tascClass_Family (intClass_ID,intFamily_ID,szUser_Create) " & _
						 "values (" & _
						 intClass_ID & "," & _
						 arFamily(iFam) & ",'" & _
						 session.Value("strUserID") & "')"
				oFunc.ExecuteCN(insert)
			next
		else
			insert = "insert into tascClass_Family (intClass_ID,intFamily_ID,szUser_Create) " & _
				     "values (" & _
					 intClass_ID & "," & _
					 Request.Form("intFamily_ID") & ",'" & _
					 session.Value("strUserID") & "')"
			oFunc.ExecuteCN(insert)
		end if 
	end if 
	
	' Save Class name for use in ILPMain.asp when creating a generic ILP
	Session("szClass_Name") = request("szClass_Name")
	Session("intClass_ID") = intClass_ID
	'Used for saving ILP Bank information in ilpMain.asp and ilpInsert.asp
	session.Contents("intPOS_Subject_ID_from_class") = request("intPOS_Subject_ID")
	
	oFunc.CommitTransCN
	call oFunc.CloseCN

	Response.Redirect("../ILP/ILPMain.asp?intContract_Guardian_id=" & request("intGuardian_id"))
end function 

function vbfUpdateClass
	call vbfFormatDates
	intClass_ID = Request.Form("intClass_ID")
	
	if Request.Form("intInstructor_ID") <> "" then
		strAddValue = "intInstructor_ID = '" & oFunc.EscapeTick(request("intInstructor_ID")) & "'," & _
					  "dtReg_Deadline = '" & dtDeadline & "'," & _
					  "dtClass_Start = '" & dtClassStart & "'," & _
					  "dtClass_End = '" & dtClassEnd & "',"
		intInstruct_Type_Id = 4
	elseif Request.Form("intGuardian_id") <> "" then
		strAddValue = "intGuardian_id = '" & oFunc.EscapeTick(request("intGuardian_id")) & "',"
		intInstruct_Type_Id = 1	
	end if 
	
	if request("dtEffective") <> "" then
		' This means class student or planning hours have changed and need to be logged.
		
		' updates tblClasses with new hours
		update = "update tblClasses set " & _
				 "decHours_Student = '" & oFunc.CheckDecimal(request("decHours_Student")) & "'," & _
				 "decHours_Planning = '" & oFunc.CheckDecimal(request("decHours_Planning")) & "',"& _
				 "dtHrs_Last_Updated = '" & request("dtEffective")  & "', " & _
				 "szUSER_MODIFY = '" & Session.Value("strUserID") & "' " & _
				 "where intClass_id = " & intClass_ID
		oFunc.ExecuteCN(update)
		
		' Creates new historical hours log for statements
		insert = "insert into tblClass_Hrs_Change(intClass_id,decHrs_Student_Change, decOLDHrs_Student_Change, " & _
				 "decHrs_Planning_Change, decOLDHrs_Planning_Change, dtEffective,strReason) values (" & _
				 intClass_id & ", " & _
				 "'" & oFunc.CheckDecimal(request("decHours_Student")) & "',"  & _		
				 "'" & oFunc.CheckDecimal(request("startStudentHrs")) & "',"  & _		
				 "'" & oFunc.CheckDecimal(request("decHours_Planning")) & "',"  & _	
				 "'" & oFunc.CheckDecimal(request("startPlanningHrs")) & "',"  & _	
				 "'" & request("dtEffective") & "'," & _
				 "'" & oFunc.EscapeTick(request("strReason")) & "')"
		oFunc.ExecuteCN(insert)
	end if 
	
	update = "update tblClasses set " & _
			 "intPOS_Subject_ID = '" & oFunc.EscapeTick(request("intPOS_Subject_ID")) & "'," & _
			 strAddValue & _
			 "szClass_Name = '" & UCase(oFunc.EscapeTick(request("szClass_Name"))) & "'," & _
			 "szASD_Course_ID = '" & UCase(oFunc.EscapeTick(request("szASD_Course_ID"))) & "'," & _
			 "szLocation = '" & UCase(oFunc.EscapeTick(request("szLocation"))) & "'," & _			 
			 "intMin_Students = " & oFunc.CheckDecimal(request("intMin_Students")) & "," & _
			 "intMax_Students = " & oFunc.CheckDecimal(request("intMax_Students")) & "," & _
			 "sGrade_Level = '" & oFunc.EscapeTick(request("sGrade_Level")) & "'," & _
			 "sGrade_Level2 = '" & oFunc.EscapeTick(request("sGrade_Level2")) & "'," & _			 
			 "szDays_Meet_On = '" & oFunc.EscapeTick(request("szDays_Meet_On")) & "'," & _
			 "szStart_Time = '" & strStartTime & "'," & _
			 "szEnd_Time = '" & strEndTime & "'," & _
			 "szSchedule_Comments = '" & oFunc.EscapeTick(request("szSchedule_Comments")) & "'," & _
			 "intInstruct_Type_Id = '" & intInstruct_Type_Id & "', "& _
			 "intDuration_ID = " & oFunc.CheckDecimal(request("intDuration_ID")) & "," & _
			 "intSession_Minutes = " & oFunc.CheckDecimal(request("intSession_Minutes")) & "," & _
			 "szUSER_MODIFY = '" & Session.Value("strUserID") & "' " & _
			 "where intClass_ID = " & intClass_ID
	oFunc.ExecuteCN(update)		
	
	' deletes and then Inserts the records that will limit a class to specific families
	if Request.Form("intFamily_ID") <> "" then
	
		delete = "delete from tascClass_Family where intClass_ID = " & intClass_ID
		oFunc.ExecuteCN(delete)
		
		if Request.Form("intFamily_ID") <> "" then
			arFamily = split(Request.Form("intFamily_ID"),",")
			if isArray(arFamily) then
				for iFam = 0 to ubound(arFamily)				
					insert = "insert into tascClass_Family (intClass_ID,intFamily_ID,szUser_Modify) " & _
							 "values (" & _
							 intClass_ID & "," & _
							 arFamily(iFam) & ",'" & _
							 session.Value("strUserID") & "')"
					oFunc.ExecuteCN(insert)
				next
			else
				insert = "insert into tascClass_Family (intClass_ID,intFamily_ID,szUser_Modify) " & _
					     "values (" & _
						 intClass_ID & "," & _
						 Request.Form("intFamily_ID") & ",'" & _
						 session.Value("strUserID") & "')"
				oFunc.ExecuteCN(insert)
			end if 
		end if 
	else
		delete = "delete from tascClass_Family where intClass_ID = " & intClass_ID
		oFunc.ExecuteCN(delete)
	end if 
end function

function vbfFormatDates
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'' Combine date variables into one
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	dtDeadline = request("regMonth") & "/" & request("regDay") & "/" & request("regYear")
	dtClassStart = request("monthStart") & "/" & request("dayStart") & "/" & request("yearStart")
	dtClassEnd = request("monthEnd") & "/" & request("dayEnd") & "/" & request("yearEnd")
	strStartTime = request("hourStart") & ":" & request("minuteStart") & " " & request("amPmStart")
	strEndTime = request("hourEnd") & ":" & request("minuteEnd") & " " & request("amPmEnd")
end function

oFunc.CommitTransCN
oFunc.CloseCN
set oFunc = nothing

dim strRefresh 
' This 'if' is used to tell viewClasses.asp to refresh so the 
' class name list will be correct since the user has changed a class name.
if Request.Form("new_name") <> "" then
	strRefresh = "window.opener.location.reload();"
else
	strRefresh = ""
end if 

%>
<HTML>
<HEAD>
<script language=javascript>
	function jfRefresh(){
		// Don't refresh page if coming from reqApprovalAdmin.asp
		var strScript = window.opener.location.href;
		if (strScript.indexOf("reqApprovalAdmin") < 1) {
			<% = strRefresh %>
		}
	}
</script>
</HEAD>
<BODY onLoad="alert('Class Updated');window.opener.focus();jfRefresh();window.close();">
</BODY>
</HTML>
