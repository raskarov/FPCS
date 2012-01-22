<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		StudentInsert.asp
'Purpose:	This script inserts or updates all SIS info for a given student.
'Date:		9 July 2001
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim oFunc
dim strMessage		' This string will be printed at the bottom of default.asp 
dim dtBirth
dim dtLottery
dim dtLottery_Received
'dim strInsert
'dim dtIEP
'dim bolIEP	
'dim strUpdate
dim intStudent_ID
dim intReEnroll_State

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

dtBirth = Request.Form("dtBirth") '& "/" & Request.Form("Day") & "/" & Request.Form("Year")

dtLottery = " NULL "
dtLottery_Received = " NULL "

if request("intReEnroll_State") = "" then
	intReEnroll_State = 1
else
	intReEnroll_State = request("intReEnroll_State")
end if

if Request.Form("lotteryMonth") <> "" then
	dtLottery = "'" & Request.Form("lotteryMonth") & "/" & Request.Form("lotteryDay") & "/" & Request.Form("lotteryYear") & "'"
end if

if Request.Form("lotteryRecvdMonth") <> "" then
	dtLottery_Received = "'" & Request.Form("lotteryRecvdMonth") & "/" & Request.Form("lotteryRecvdDay") & "/" & Request.Form("lotteryRecvdYear") & "'"
end if

' IEP sectino removed 6-9-03.  Functionality moved to iep.asp
'if Request.Form("IEPyear") <> "" then
'	strInsert = ",dtIEP_Renewal"	
'	dtIEP = "'" & Request.Form("IEPMonth") & "/" & Request.Form("IEPDay") & "/" & Request.Form("IEPYear") & "', "
'	strUpdate = "dtIEP_Renewal = " & dtIEP 
'end if 

'bkm - 30-sept-2002
'if UCase(Request.Form("bolIEP")) = "YES" then
'	bolIEP = 1
'elseif UCase(Request.Form("bolIEP")) = "NO" then
'	bolIEP = 0
'else
'	bolIEP = "NULL"
'end if
'response.End
if Request.Form("intStudent_ID") = "" then
	'If we get here we need to create a new record in the database.	
	dim insert	
	dim DOBInsert
	dim DOBValue

	DOBInsert= " dtBirth, "
	DOBValue = "'" & dtBirth & "',"

	if not isdate(dtBirth) then
		DOBInsert= ""
		DOBValue = ""
	end if
	
	oFunc.BeginTransCN		
	
	insert = "insert into tblStudent(szLast_Name,szFirst_Name,sMid_Initial," & _
			 "szSSN, sSex, intRace_ID, intTuition_id, " & DOBInsert & _
			 "intFirst_Lang, intHome_Lang,szGrade, intGrad_Year, " & _
			 "szPrevious_School,szPrev_School_Year,szPrev_School_Addr," & _
		 	 "szPrev_School_City,szPrev_School_State,szPrev_School_Country," & _
			 "szPrev_School_Zip_Code,szPrev_Anch_School,intPrev_Anch_Year," & _
			 "szContact_Last_Name,szContact_First_Name,szContact_Phone," & _
			 "szDR_Last_Name,szDR_First_Name,szDR_Phone,szDaycare_Name," & _
			 "szDaycare_Phone,szMed_Alert_1,szMed_Alert_2,szDisability_1," & _
			 "szDisability_2" & strInsert & ",szUSER_CREATE, dtLottery,dtLottery_Received,szNew_Wait_List_Num) VALUES ('" & _
			 oFunc.EscapeTick(Request.Form("szLast_Name")) & "','" & _
			 oFunc.EscapeTick(Request.Form("szFirst_Name")) & "','" & _ 
			 oFunc.EscapeTick(Request.Form("sMid_Initial")) & "','" & _
			 oFunc.EscapeTick(Request.Form("szSSN")) & "','" & _
			 Request.Form("sSex") & "','" &_
			 Request.Form("intRace_ID") & "','" & _ 
			 request("intTuition_id") & "'," & DOBValue & _
			 oFunc.CheckDecimal(Request.Form("intFirst_Lang")) & "," & _
			 oFunc.CheckDecimal(Request.Form("intHome_Lang")) & ",'" & _
			 oFunc.EscapeTick(Request.Form("szGrade")) & "','" & _
			 oFunc.CheckDecimal(Request.Form("intGrad_Year")) & "', '" & _
			 oFunc.EscapeTick(request("szPrevious_School")) & "','" & _
			 oFunc.EscapeTick(request("szPrev_School_Year")) & "','" & _
			 oFunc.EscapeTick(request("szPrev_School_Addr")) & "','" & _
			 oFunc.EscapeTick(request("szPrev_School_City")) & "','" & _
			 oFunc.EscapeTick(request("szPrev_School_State")) & "','" & _
			 oFunc.EscapeTick(request("szPrev_School_Country")) & "','" & _
			 oFunc.EscapeTick(request("szPrev_School_Zip_Code")) & "','" & _  
			 oFunc.EscapeTick(request("szPrev_Anch_School")) & "','" & _ 
			 oFunc.CheckDecimal(request("intPrev_Anch_Year")) & "','" & _ 
			 oFunc.EscapeTick(request("szContact_Last_Name")) & "','" & _ 
			 oFunc.EscapeTick(request("szContact_First_Name")) & "','" & _ 
			 oFunc.EscapeTick(request("szContact_Phone")) & "','" & _ 
			 oFunc.EscapeTick(request("szDR_Last_Name")) & "','" & _ 
			 oFunc.EscapeTick(request("szDR_First_Name")) & "','" & _ 
			 oFunc.EscapeTick(request("szDR_Phone")) & "','" & _ 
			 oFunc.EscapeTick(request("szDaycare_Name")) & "','" & _ 
			 oFunc.EscapeTick(request("szDaycare_Phone")) & "','" & _ 
			 oFunc.EscapeTick(request("szMed_Alert_1")) & "','" & _ 
			 oFunc.EscapeTick(request("szMed_Alert_2")) & "','" & _ 
			 oFunc.EscapeTick(request("szDisability_1")) & "','" & _ 
			 oFunc.EscapeTick(request("szDisability_2")) & "','" & _ 
			 Session.Value("strUserID")	& "'," & _
			 dtLottery & "," & dtLottery_Received & ",'" & _
			 oFunc.EscapeTick(request("szNew_Wait_List_Num")) & "' )"
			 
	oFunc.ExecuteCN(insert)
	
	intStudent_ID = oFunc.GetIdentity
	' Insert Enrollment Record
	call vbsInsertEnroll 
	'Insert Student States Record
	call vbsInsertStudentState(intStudent_ID)
	
	' Insert Exemptions record if needed
	if session.Contents("strRole") = "ADMIN" and _
		(request.Form("bolASD_Contract_HRS_Exempt") = 1 or _
		request("intCore_Credit_Percent") <> "" or _
		request("intElective_Credit_Percent") <> "") and _
		request.Form("intStudent_Exemption_ID") = "" then		 
		call vbsInsertExemptions
	end if
	
	oFunc.CommitTransCN		 
	strMessage = "Student Information was Added."
	
	if Request.Form("bolNewStudent") <> "" then 
		' We add the new student to the student list on the family manager page and close this window.
%>
<html>
<head>
<script language=javascript>
	function jfAddStudentToList(){
		// Passes info needed to create a new option in the student select list that is
		// contained in familyManager.asp 
		window.opener.jfAddOption('<% = Request.Form("szLast_Name")& "," & Request.Form("szFirst_Name")%>','<%=intStudent_ID%>','selStudent_ID');
		window.close();
	}

</script>
<body onload="jfAddStudentToList();" bgcolor=white>
</body>
</html>
<%	
		Response.End
	end if 
elseif  Request.Form("intStudent_ID") <> ""  then
	' A change has been made in edit mode so we update the data.
	intStudent_ID = Request.Form("intStudent_ID")
	oFunc.BeginTransCN
	dim update
	update = "update tblStudent set " & _
			 "szLast_Name = '" & oFunc.EscapeTick(Request.Form("szLast_Name")) & "'," & _
			 "szFirst_Name = '" & oFunc.EscapeTick(Request.Form("szFirst_Name")) & "'," & _ 
			 "sMid_Initial = '" & oFunc.EscapeTick(Request.Form("sMid_Initial")) & "'," & _
			 "szSSN = '" & oFunc.EscapeTick(Request.Form("szSSN")) & "'," & _
			 "sSex = '" & Request.Form("sSex") & "'," &_
			 "intRace_ID = " & Request.Form("intRace_ID") & "," & _ 
			 "dtBirth = '" & dtBirth & "'," & _
			 "intFirst_Lang = " & oFunc.CheckDecimal(Request.Form("intFirst_Lang")) & "," & _
			 "intHome_Lang = " & oFunc.CheckDecimal(Request.Form("intHome_Lang")) & "," & _
			 "intGrad_Year = " & oFunc.CheckDecimal(Request.Form("intGrad_Year")) & "," & _
			 "szPrevious_School = '" & oFunc.EscapeTick(request("szPrevious_School")) & "'," & _
			 "szPrev_School_Year = '" & oFunc.EscapeTick(request("szPrev_School_Year")) & "'," & _
			 "szPrev_School_Addr = '" & oFunc.EscapeTick(request("szPrev_School_Addr")) & "'," & _
			 "szPrev_School_City = '" & oFunc.EscapeTick(request("szPrev_School_City")) & "'," & _
			 "szPrev_School_State = '" & oFunc.EscapeTick(request("szPrev_School_State")) & "'," & _
			 "szPrev_School_Country = '" & oFunc.EscapeTick(request("szPrev_School_Country")) & "'," & _
			 "szPrev_School_Zip_Code = '" & oFunc.EscapeTick(request("szPrev_School_Zip_Code")) & "'," & _  
			 "szPrev_Anch_School = '" & oFunc.EscapeTick(request("szPrev_Anch_School")) & "'," & _ 
			 "intPrev_Anch_Year = " & oFunc.CheckDecimal(request("intPrev_Anch_Year")) & "," & _ 
			 "szContact_Last_Name = '" & oFunc.EscapeTick(request("szContact_Last_Name")) & "'," & _ 
			 "szContact_First_Name = '" & oFunc.EscapeTick(request("szContact_First_Name")) & "'," & _ 
			 "szContact_Phone = '" & oFunc.EscapeTick(request("szContact_Phone")) & "'," & _ 
			 "szDR_Last_Name = '" & oFunc.EscapeTick(request("szDR_Last_Name")) & "'," & _ 
			 "szDR_First_Name = '" & oFunc.EscapeTick(request("szDR_First_Name")) & "'," & _ 
			 "szDR_Phone = '" & oFunc.EscapeTick(request("szDR_Phone")) & "'," & _ 
			 "szDaycare_Name = '" & oFunc.EscapeTick(request("szDaycare_Name")) & "'," & _ 
			 "szDaycare_Phone = '" & oFunc.EscapeTick(request("szDaycare_Phone")) & "'," & _ 
			 "szMed_Alert_1 = '" & oFunc.EscapeTick(request("szMed_Alert_1")) & "'," & _ 
			 "szMed_Alert_2 = '" & oFunc.EscapeTick(request("szMed_Alert_2")) & "'," & _ 
			 "szDisability_1 = '" & oFunc.EscapeTick(request("szDisability_1")) & "'," & _ 
			 "szDisability_2 = '" & oFunc.EscapeTick(request("szDisability_2")) & "'," & _ 
			 "szUser_Modify = '" & Session.Value("strUserID") & "', " & _
			 "dtLottery = " & dtLottery & ", dtLottery_Received = " & dtLottery_Received & "," & _
			 "szNew_Wait_List_Num = '" & oFunc.EscapeTick(request("szNew_Wait_List_Num")) & "' " & _
			 "where intStudent_ID = " & Request.Form("intStudent_ID")

	oFunc.ExecuteCN(update)
	
	intStudent_ID = Request.Form("intStudent_ID")
	if request("intEnroll_Info_ID") = "" then
		call vbsInsertEnroll				
	else 
		call vbsUpdateEnroll
	end if
	
	' Do we need to insert/update exemptions?
	if request.Form("intStudent_Exemption_ID") <> "" then 
		call vbsUpdateExemptions
	elseif session.Contents("strRole") = "ADMIN" and _
		(request.Form("bolASD_Contract_HRS_Exempt") = 1 or _
		request("intCore_Credit_Percent") <> "" or _
		request("intElective_Credit_Percent") <> "") and _
		request.Form("intStudent_Exemption_ID") = "" then		 
		call vbsInsertExemptions
	end if
	
	' Check to see if we need to insert or update Student States
	if request("intStudent_State_ID") & "" <> "" then
		call vbsUpdateStudentState(request("intStudent_State_ID"))
	else
		call vbsInsertStudentState(Request.Form("intStudent_ID"))
	end if
	
	oFunc.CommitTransCN
	strMessage = "Student Information was Updated."
end if


' handle managing conditionally enrolled locks
if request("oldReEnrollState") & "" <> "129" and request("intReEnroll_State") & "" = "129" then
	' Student is being set to conditionally enrolled so engage the spending lock
	insert = "insert into STUDENT_LOCKED_ACCOUNTS (StudentID, SchoolYear, UserCreated) " & _
			 " values (" & intStudent_ID & "," & session.Contents("intSchool_Year") & _
				",'" & session.Contents("strUserId") & "')"
	oFunc.ExecuteCN(insert)
	
	Application.Contents("LockedStudentAccounts" & session.Contents("intSchool_Year")) = Application.Contents("LockedStudentAccounts" & session.Contents("intSchool_Year")) & intStudent_ID & ","
elseif 	request("oldReEnrollState") & "" = "129" and request("intReEnroll_State") & "" <> "129" then
	'studnet is no longer conditionally enrolled so remove spending lock
	delete = "delete from STUDENT_LOCKED_ACCOUNTS where StudentID = " & intStudent_ID & " and SchoolYear = " &  session.Contents("intSchool_Year")
	oFunc.ExecuteCN(delete)
	Application.Contents("LockedStudentAccounts" & session.Contents("intSchool_Year")) = replace(Application.Contents("LockedStudentAccounts" & session.Contents("intSchool_Year")), intStudent_ID & ",","")
end if

if request("hdnNeedChanged") & "" <> "" then
	' Save the changed Enroll Info Need data
	delete = "delete from STUDENT_ENROLL_INFO_NEEDED where StudentID = " & intStudent_ID & _
				" and SchoolYear = " & session.Contents("intSchool_Year") 
	oFunc.ExecuteCN(delete)
	
	iList = split(request("EnrollNeed"), ",")
	for i = 0 to ubound(iList)
		if iList(i) <> "" then
			insert = "insert into STUDENT_ENROLL_INFO_NEEDED (StudentID, SchoolYear,NeededEnrollInfoCD,UserCreated) " & _
						" values(" & intStudent_ID & "," & session.Contents("intSchool_Year") & ",'" & trim(iList(i)) & "','" &  Session.Contents("strUserID") & "')"
			'response.Write insert
			oFunc.ExecuteCN(insert)
		end if
	next
end if

'response.Write request("hdnNeedChanged") & "<<" & request("EnrollNeed")
'response.End
call oFunc.CloseCN	 

sub vbsInsertEnroll
	'Insert new Enrollment Info Record
	insert = "insert into tblEnroll_Info (intStudent_ID,sintSchool_Year," & _
			  "szPrivate_School_Name,szOther_District_Name,intPercent_Enrolled_D2," & _
			  "intPercent_Enrolled_Fpcs,bolCharter_Grad,szUser_Create) values (" & _
			 intStudent_ID & "," & _
			 session.Value("intSchool_Year") & "," & _
			 "'" & oFunc.EscapeTick(Request.Form("szPrivate_School_Name")) & "'," & _
			 "'" & oFunc.EscapeTick(Request.Form("szOther_District_Name")) & "'," & _				 
			 "'" & oFunc.CheckDecimal(Request.Form("intPercent_Enrolled_D2")) & "'," & _
			 "'" & oFunc.CheckDecimal(Request.Form("intPercent_Enrolled_Fpcs")) & "'," & _
			 "'" & oFunc.EscapeTick(Request.Form("bolCharter_Grad")) & "'," & _
			 "'" & oFunc.EscapeTick(session.Value("strUserID")) & "')" 

	oFunc.ExecuteCN(insert)		 
end sub		

function vbsUpdateEnroll
	'Updates an Enrollment Record
	update = "update tblEnroll_Info set " & _
			 "szPrivate_School_Name = '" & oFunc.EscapeTick(Request.Form("szPrivate_School_Name")) & "'," & _
			 "szOther_District_Name = '" & oFunc.EscapeTick(Request.Form("szOther_District_Name")) & "'," & _
			 "intPercent_Enrolled_D2 = '" & oFunc.CheckDecimal(Request.Form("intPercent_Enrolled_D2")) & "'," & _
			 "intPercent_Enrolled_Fpcs = '" & oFunc.CheckDecimal(Request.Form("intPercent_Enrolled_Fpcs")) & "'," & _
			 "bolCharter_Grad = '" & oFunc.EscapeTick(Request.Form("bolCharter_Grad")) & "'," & _
			 "szUser_Modify = '" & oFunc.EscapeTick(session.Contents("strUserID")) & "'," & _
			 "dtModify = '" & now() & "' " & _
			 "where intEnroll_Info_ID = " & Request.Form("intEnroll_Info_ID")
			  
	oFunc.ExecuteCN(update)
end function

sub vbsInsertExemptions
	' Inserts an Exemption record
	insert = "insert into tblStudent_Exemptions(intStudent_ID,intSchool_Year," &_
				 "bolASD_Contract_HRS_Exempt,szHRS_Exempt_Reason,intCore_Credit_Percent," & _
				 "szCore_Exemption_Reason,intElective_Credit_Percent," & _
				 "szElective_Exemption_Reason,szUser_Create,dtCreate) " & _
				 "values (" & _
				 intStudent_ID & "," & _
				 session.Contents("intSchool_Year") & "," & _
				 request.Form("bolASD_Contract_HRS_Exempt") & ",'" & _
				 oFunc.EscapeTick(request.Form("szHRS_Exempt_Reason")) & _
				 "'," & oFunc.CheckDecimal(request("intCore_Credit_Percent"))  & _
				 ",'" & oFunc.EscapeTick(request("szCore_Exemption_Reason")) & "'" & _
				 "," & oFunc.CheckDecimal(request("intElective_Credit_Percent"))  & _
				 ",'" & oFunc.EscapeTick(request("szElective_Exemption_Reason")) & "'" & _
				 ",'" & session.Contents("strUserID") & "'" & _
				 ",'" & now() & "')"
	oFunc.ExecuteCN(insert)	
end sub

sub vbsUpdateExemptions
	'Update Exemptions
	update = "update tblStudent_Exemptions " & _
				"set bolASD_Contract_HRS_Exempt = '" & request.Form("bolASD_Contract_HRS_Exempt") & "'," & _
				"szHRS_Exempt_Reason = '" & oFunc.EscapeTick(request.Form("szHRS_Exempt_Reason")) & "', " & _
				"intCore_Credit_Percent = " & oFunc.CheckDecimal(request("intCore_Credit_Percent")) & "," & _
				"szCore_Exemption_Reason = '" & oFunc.EscapeTick(request("szCore_Exemption_Reason")) & "'," & _
				"intElective_Credit_Percent = " & oFunc.CheckDecimal(request("intElective_Credit_Percent")) & "," & _
				"szElective_Exemption_Reason = '" & oFunc.EscapeTick(request("szElective_Exemption_Reason")) & "', " & _				 
				"szUser_Modify = '" & oFunc.EscapeTick(session.Contents("strUserID")) & "'," & _
				"dtModify = '" & now() & "' " & _
				"where intStudent_Exemption_ID = " & request.Form("intStudent_Exemption_ID")
	oFunc.ExecuteCN(update)
end sub

sub vbsInsertStudentState(pStudentID)
		
	if UCASE(session.Contents("strRole")) = "ADMIN" then
		eStateField = "intReEnroll_State,"
		eStateValue = intReEnroll_State & "," 
		if request("dtWithdrawn") <> "" then
			if isDate(request("dtWithdrawn")) then
				eStateField = eStateField & "dtWithdrawn,"
				eStateValue = eStateValue & "'" & request("dtWithdrawn") & "'," 
			end if
		end if
	else
		eStateField = ""
		eStateValue = ""
	end if
	
	dim insert
	insert = "INSERT INTO tblStudent_States " & _
			"(intStudent_id, " & eStateField & " intSchool_Year,szGrade, " & _
			" szUSER_CREATE) " & _
			"VALUES (" & _
			pStudentID & "," & _
			eStateValue & _
			session.Contents("intSchool_Year") & "," & _
			"'" & oFunc.EscapeTick(request("szGrade")) & "'," & _
			"'" & oFunc.EscapeTick(session.Contents("strUserID")) & "')"
	oFunc.ExecuteCN(insert)
end sub 

sub vbsUpdateStudentState(pStudentStateID)
	if UCASE(session.Contents("strRole")) = "ADMIN" then
		eStateSet = "intReEnroll_State = " & request("intReEnroll_State") & "," 
		if request("dtWithdrawn") <> "" then
			if isDate(request("dtWithdrawn")) then
				strWithdrawn = ", dtWithdrawn = '" & request("dtWithdrawn") & "' " 
			end if
		end if
	else
		eStateSet = ""
	end if
	
	dim update
	update = "update tblStudent_States set " & _
			eStateSet & _
			"szUser_Modify = '" & oFunc.EscapeTick(session.Contents("strUserID")) & "', " & _
			"szGrade = '" & oFunc.EscapeTick(request("szGrade")) & "' " & _
			strWithdrawn & _
			"WHERE intStudent_State_ID = " & pStudentStateID 
	oFunc.ExecuteCN(update)
end sub 

if request("bolHeader") <> "" then
	response.Redirect("studentProfile.asp?intStudent_ID=" & intStudent_ID)
	response.End
end if

%>
<html>
<head>
<script language=javascript>
	function jfClose(){
		alert("Student info has been updated.");
		<% if Request.Form("intEnroll_Info_ID") = "" then %>
		opener.window.location.reload();
		<% end if %>
		opener.window.focus();
		window.close();
	}

</script>
<body onload="jfClose();" bgcolor=white>
</body>
</html>

