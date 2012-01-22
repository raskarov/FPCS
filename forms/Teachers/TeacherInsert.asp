<%@ Language=VBScript %>
<%
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Make db Connection, print HTML header.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim oFunc
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()
   
dim insert
dim dtBirth
dim strMessage
dim intInstructor_ID
dim sql
dim delete
dim fvCount   'html form variable count 
dim bolActive

if request("bolActive") = "" then 
	bolActive = 0
else
	bolActive = 1
end if

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Combine date variables into one
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
dtBirth = request("dtBirth")

if request("dtCert_Expire") <> "" then
	dtCert_Expire = " '" & request("dtCert_Expire") & "' " 
else
	dtCert_Expire = " NULL " 
end if

oFunc.BeginTransCN

'If we are just verifing a teachers profile make update and procede to next item
fvCount = Request.Form("intCount")
if fvCount <> "" and Request.Form("btAccept") <> "" then
	call vbfVerifyProfile
end if 

intInstructor_ID = request("intInstructor_ID")
 
if Request.form("intInstructor_ID") = "" and Request.QueryString("intInstructor_ID") = "" then
	insert = "insert into tblInstructor(szTitle,szFirst_Name,szLast_Name,sMid_Initial," & _
			 "szSSN,szMailing_ADDR,szCity,sState,szZip_Code,szHome_Phone,szBusiness_Phone," & _
			 "intBusiness_Ext,szCell_Phone,szEmail,szEmail2,dtBirth,intPay_Type_id," & _
			 "curPay_Rate,intDist_Code," & _
			 "bolOn_ASD_Leave,bolSubstitute,bolASD_Employee,bolASD_Eligible_For_Hire," & _
			 "bolASD_Retired,bolGroup_Instruction,bolIndividual_Instruction," & _
			 "intYears_Experience,bolK_8,bolK_12,bolSpecial_Ed,bolSecondary,szSecondary_List," & _
			 "bolMy_Classroom,bolMy_Home,bolStudents_Home,bolFPCS_Classroom,bolOther,szOther_Desc," & _
			 "bolAvail_Weekdays,bolAvail_Wk_Afternoon,bolAvail_Wk_Evening,bolAvail_Wk_Ends,strASD_School,bolAvail_Summers, szUSER_CREATE,dtCert_Expire) " & _
			 "VALUES (" & _
			 "'" & oFunc.EscapeTick(request("szTitle")) & "'," & _
			 "'" & UCase(oFunc.EscapeTick(request("szFirst_Name"))) & "'," & _
			 "'" & UCase(oFunc.EscapeTick(request("szLast_Name"))) & "'," & _
			 "'" & UCase(oFunc.EscapeTick(request("sMid_Initial"))) & "'," & _
			 "'" & oFunc.EscapeTick(request("szSSN")) & "'," & _
			 "'" & UCase(oFunc.EscapeTick(request("szMailing_ADDR"))) & "'," & _
			 "'" & UCase(oFunc.EscapeTick(request("szCity"))) & "'," & _
			 "'" & UCase(oFunc.EscapeTick(request("sState"))) & "'," & _
			 "'" & oFunc.EscapeTick(request("szZip_Code")) & "'," & _
			 "'" & oFunc.EscapeTick(request("szHome_Phone")) & "'," & _
			 "'" & oFunc.EscapeTick(request("szBusiness_Phone")) & "'," & _
			 "'" & oFunc.EscapeTick(request("intBusiness_Ext")) & "'," & _
			 "'" & oFunc.EscapeTick(request("szCell_Phone")) & "'," & _
			 "'" & oFunc.EscapeTick(request("szEmail")) & "'," & _
			 "'" & oFunc.EscapeTick(request("szEmail2")) & "'," & _
			 "'" & dtBirth & "'," & _			 			 
			 "'" & oFunc.EscapeTick(request("intPay_Type_id")) & "'," & _
			 "convert(money,'" & oFunc.EscapeTick(request("curPay_Rate")) & "')," & _
			 "'" & oFunc.EscapeTick(request("intDist_Code")) & "'," & _		
			 "'" & oFunc.ConvertCheckToBit(request("bolOn_ASD_Leave")) & "'," & _	
			 "'" & oFunc.ConvertCheckToBit(request("bolSubstitute")) & "'," & _	
			 "'" & oFunc.ConvertCheckToBit(request("bolASD_Employee")) & "'," & _	
			 "'" & oFunc.ConvertCheckToBit(request("bolASD_Eligible_For_Hire")) & "'," & _	
			 "'" & oFunc.ConvertCheckToBit(request("bolASD_Retired")) & "'," & _	
			 "'" & oFunc.ConvertCheckToBit(request("bolGroup_Instruction")) & "'," & _	
			 "'" & oFunc.ConvertCheckToBit(request("bolIndividual_Instruction")) & "'," & _	
			 "'" & oFunc.CheckDecimal(request("intYears_Experience")) & "'," & _	
			 "'" & oFunc.ConvertCheckToBit(request("bolK_8")) & "'," & _	
			 "'" & oFunc.ConvertCheckToBit(request("bolK_12")) & "'," & _	
			 "'" & oFunc.ConvertCheckToBit(request("bolSpecial_Ed")) & "'," & _	
			 "'" & oFunc.ConvertCheckToBit(request("bolSecondary")) & "'," & _	
			 "'" & oFunc.EscapeTick(request("szSecondary_List")) & "'," & _	
			 "'" & oFunc.ConvertCheckToBit(request("bolMy_Classroom")) & "'," & _	
			 "'" & oFunc.ConvertCheckToBit(request("bolMy_Home")) & "'," & _	
			 "'" & oFunc.ConvertCheckToBit(request("bolStudents_Home")) & "'," & _	
			 "'" & oFunc.ConvertCheckToBit(request("bolFPCS_Classroom")) & "'," & _	
			 "'" & oFunc.ConvertCheckToBit(request("bolOther")) & "'," & _	
			 "'" & oFunc.EscapeTick(request("szOther_Desc")) & "'," & _	
			 "'" & oFunc.ConvertCheckToBit(request("bolAvail_Weekdays")) & "'," & _	
			 "'" & oFunc.ConvertCheckToBit(request("bolAvail_Wk_Afternoon")) & "'," & _	
			 "'" & oFunc.ConvertCheckToBit(request("bolAvail_Wk_Evening")) & "'," & _	
			 "'" & oFunc.ConvertCheckToBit(request("bolAvail_Wk_Ends")) & "'," & _	
			 "'" & oFunc.EscapeTick(request("strASD_School")) & "'," & _	
			 "'" & oFunc.ConvertCheckToBit(request("bolAvail_Summers")) & "', " & _
			 "'" & Session.Value("strUserID")	& "'," & dtCert_Expire & ")"
	oFunc.ExecuteCN(insert)		 
	
	intInstructor_ID = oFunc.GetIdentity
	
	call vbfInsertPayData()
	
	strMessage = "A new teacher has been added."
	
elseif (Request.form("intInstructor_ID") <> "" or Request.QueryString("intInstructor_ID") <> "") _
	   and ucase(request("changed")) = ucase("yes") then
	dim update	
	
	update = "update tblInstructor set " & _
			 "szTitle = '" & oFunc.EscapeTick(request("szTitle")) & "'," & _
			 "szFirst_Name = '" & UCase(oFunc.EscapeTick(request("szFirst_Name"))) & "'," & _
			 "szLast_Name = '" & UCase(oFunc.EscapeTick(request("szLast_Name"))) & "'," & _
			 "sMid_Initial = '" & UCase(oFunc.EscapeTick(request("szVendor_Zip_Code"))) & "'," & _
			 "szSSN = '" & oFunc.EscapeTick(request("szSSN")) & "'," & _
			 "szMailing_ADDR = '" & UCase(oFunc.EscapeTick(request("szMailing_ADDR"))) & "'," & _
			 "szCity = '" & UCase(oFunc.EscapeTick(request("szCity"))) & "'," & _
			 "sState = '" & UCase(oFunc.EscapeTick(request("sState"))) & "', " & _
			 "szZip_Code = '" & oFunc.EscapeTick(request("szZip_Code")) & "', " & _
			 "szHome_Phone = '" & oFunc.EscapeTick(request("szHome_Phone")) & "', " & _
			 "szBusiness_Phone = '" & oFunc.EscapeTick(request("szBusiness_Phone")) & "', " & _
			 "intBusiness_Ext = '" & oFunc.EscapeTick(request("intBusiness_Ext")) & "', " & _
			 "szCell_Phone = '" & oFunc.EscapeTick(request("szCell_Phone")) & "'," & _
			 "szEmail = '" & oFunc.EscapeTick(request("szEmail")) & "', " & _
			 "szEmail2 = '" & oFunc.EscapeTick(request("szEmail2")) & "', " & _
			 "dtBirth = '" & dtBirth & "', " & _
			 "intPay_Type_id = '" & oFunc.EscapeTick(request("intPay_Type_id")) & "', " & _			 
			 "bolOn_ASD_Leave = '" & oFunc.ConvertCheckToBit(request("bolOn_ASD_Leave")) & "'," & _	
			 "bolSubstitute = '" & oFunc.ConvertCheckToBit(request("bolSubstitute")) & "'," & _	
			 "bolASD_Employee = '" & oFunc.ConvertCheckToBit(request("bolASD_Employee")) & "'," & _	
			 "bolASD_Eligible_For_Hire = '" & oFunc.ConvertCheckToBit(request("bolASD_Eligible_For_Hire")) & "'," & _	
			 "bolASD_Retired = '" & oFunc.ConvertCheckToBit(request("bolASD_Retired")) & "'," & _	
			 "bolGroup_Instruction = '" & oFunc.ConvertCheckToBit(request("bolGroup_Instruction")) & "'," & _	
			 "bolIndividual_Instruction = '" & oFunc.ConvertCheckToBit(request("bolIndividual_Instruction")) & "'," & _	
			 "intYears_Experience = '" & oFunc.CheckDecimal(request("intYears_Experience")) & "'," & _	
			 "bolK_8 = '" & oFunc.ConvertCheckToBit(request("bolK_8")) & "'," & _	
			 "bolK_12 = '" & oFunc.ConvertCheckToBit(request("bolK_12")) & "'," & _	
			 "bolSpecial_Ed = '" & oFunc.ConvertCheckToBit(request("bolSpecial_Ed")) & "'," & _	
			 "bolSecondary = '" & oFunc.ConvertCheckToBit(request("bolSecondary")) & "'," & _	
			 "szSecondary_List = '" & oFunc.EscapeTick(request("szSecondary_List")) & "'," & _	
			 "bolMy_Classroom = '" & oFunc.ConvertCheckToBit(request("bolMy_Classroom")) & "'," & _	
			 "bolMy_Home = '" & oFunc.ConvertCheckToBit(request("bolMy_Home")) & "'," & _	
			 "bolStudents_Home = '" & oFunc.ConvertCheckToBit(request("bolStudents_Home")) & "'," & _	
			 "bolFPCS_Classroom = '" & oFunc.ConvertCheckToBit(request("bolFPCS_Classroom")) & "'," & _	
			 "bolOther = '" & oFunc.ConvertCheckToBit(request("bolOther")) & "'," & _	
			 "szOther_Desc = '" & oFunc.EscapeTick(request("szOther_Desc")) & "'," & _	
			 "bolAvail_Weekdays = '" & oFunc.ConvertCheckToBit(request("bolAvail_Weekdays")) & "'," & _	
			 "bolAvail_Wk_Afternoon = '" & oFunc.ConvertCheckToBit(request("bolAvail_Wk_Afternoon")) & "'," & _	
			 "bolAvail_Wk_Evening = '" & oFunc.ConvertCheckToBit(request("bolAvail_Wk_Evening")) & "'," & _	
			 "bolAvail_Wk_Ends = '" & oFunc.ConvertCheckToBit(request("bolAvail_Wk_Ends")) & "'," & _
			 "strASD_School = '" & oFunc.EscapeTick(request("strASD_School")) & "'," & _
			 "bolAvail_Summers = '" & oFunc.ConvertCheckToBit(request("bolAvail_Summers")) & "', " & _
			 "szUSER_MODIFY = '" & Session.Value("strUserID") & "', " & _			 
			 "dtCert_Expire = " & dtCert_Expire & " " & _
			 "where intInstructor_ID = " & intInstructor_ID
			 
	oFunc.ExecuteCN(update)
	
	if Request.Form("bolChangePayData") <> "" then							  
		  ' This code could return an incorrect record due to the use of the MAX function
			'set rsGetEffectiveDT = server.CreateObject("ADODB.RECORDSET")
			'rsGetEffectiveDT.CursorLocation = 3
			
			'sql = "SELECT MAX(dtEffective_Start) AS effective, " & _
			'	  "intInstructor_Pay_Data_ID " & _
			'	  "FROM tblInstructor_Pay_Data " & _
			'	  "WHERE intInstructor_id = " & intInstructor_ID & _
			'	  " AND dtEffective_End IS NULL " & _
			'	  "GROUP BY intInstructor_Pay_Data_ID "
			'rsGetEffectiveDT.Open sql,oFunc.FPCScnn	
			
			'if rsGetEffectiveDT.RecordCount > 0 then
			
		' This code replaced the above on 5-6-05 smb	
		' We need to set a new effective date for an updated teacher pay rate 
		' So we first set the previous pay effective end date to be the day before 
		' the new rates effective date
		
		' PLEASE NOTE: A teacher will only have ONE active pay record for 
		' an ENTIRE school year.  This means if a teachers per deim was $200 at the beginning 
		' of the school year but mid year the teachers per deim went to $250 then the teachers
		' per deim would be $250 for the entire year, start to finish.  
		if Request.Form("intInstructor_Pay_Data_ID") <> "" then
			'dtEffectiveEnd = "6/30/" & (cint(session.Contents("intSchool_Year")))
		
			update = "update tblInstructor_Pay_Data set " & _
					 "dtEffective_End = CURRENT_TIMESTAMP, " & _
					 "szUSER_MODIFY = '" & Session.Value("strUserID") & "' " & _
					 "where intInstructor_Pay_Data_ID = " & Request.Form("intInstructor_Pay_Data_ID")
			oFunc.ExecuteCN(update)
		end if 				
		
		'Add new pay rate record
		call vbfInsertPayData()
		
		' No longer need this section since we will always insert a new pay record
		' each time a change has been made. Commented out 5-9-05
			'elseif Request.Form("intInstructor_Pay_Data_ID") <> "" then
	
				'	update = "update tblInstructor_Pay_Data set " & _
				'			 "curPer_Hour = convert(money,'" & oFunc.EscapeTick(request("curPer_Hour")) & "')," & _
				'			 "curPer_Hour_Benefits = convert(money,'" & oFunc.EscapeTick(request("curPer_Hour_Benefits")) & "')," & _
				'			 "curPay_Rate = convert(money,'" & oFunc.EscapeTick(request("curPay_Rate")) & "')," & _
				'			 "intPay_Type_id = '" & oFunc.EscapeTick(request("intPay_Type_id")) & "'," & _	
				'			 "bolASD_Full_Time = '" & oFunc.ConvertCheckToBit(request("bolASD_Full_Time")) & "'," & _	
				'			 "decASD_Full_Time_Percent = '" & oFunc.EscapeTick(request("decASD_Full_Time_Percent")) & "'," & _	
				'			 "bolASD_Part_Time = '" & oFunc.ConvertCheckToBit(request("bolASD_Part_Time")) & "'," & _
				'			 "decASD_Part_Time_Percent = '" & oFunc.EscapeTick(request("decASD_Part_Time_Percent")) & "'," & _	
				'			 "decFPCS_Hours_Goal = '" & oFunc.CheckDecimal(request("decFPCS_Hours_Goal")) & "', " & _
				'			 "bolActive = " & bolActive & ", " & _	
				'			 "szUSER_MODIFY = '" & Session.Value("strUserID") & "' " & _
				'			 "where intInstructor_Pay_Data_ID = " & Request.Form("intInstructor_Pay_Data_ID")
				'	oFunc.ExecuteCN(update)
	end if 
	strMessage = "Teacher has been updated."
end if 

if fvCount <> "" then
	' This is fired if a forced action (verify profile) is active
	' and the user modified the record. We will still record the results of the 
	' forced event.
	call vbfVerifyProfile
end if 

oFunc.CommitTransCN
oFunc.CloseCN
set oFunc = nothing
function vbfInsertPayData()
	'Creates a teachers pay info record.
	'This is called in the update section as well because if data is changed
	'that effects the amount a student will be charged we can not update the
	'old record since we must maintain a history of pay data. 

	insert = "insert into tblInstructor_Pay_Data (" & _
			 "intInstructor_ID," & _
			 "curPay_Rate,intPay_Type_id,bolASD_Full_Time,fltASD_Full_Time_Percent,bolActive," & _
			 "bolASD_Part_Time,fltASD_Part_Time_Percent,fltFPCS_Hours_Goal,dtEffective_Start, " & _
			 "intSchool_Year_Start, szUSER_CREATE,bolMasters_Degree,szSalary_Placement) " & _
			 " VALUES (" & _
			 intInstructor_ID & "," & _
			 "convert(money,'" & oFunc.EscapeTick(request("curPay_Rate")) & "')," & _
			 "'" & oFunc.EscapeTick(request("intPay_Type_id")) & "'," & _	
			 "'" & oFunc.ConvertCheckToBit(request("bolASD_Full_Time")) & "'," & _	
			 "'" & oFunc.CheckDecimal(request("fltASD_Full_Time_Percent")) & "'," & _
			 bolActive & "," & _	
			 "'" & oFunc.ConvertCheckToBit(request("bolASD_Part_Time")) & "'," & _
			 "'" & oFunc.CheckDecimal(request("fltASD_Part_Time_Percent")) & "'," & _	
			 "'" & oFunc.CheckDecimal(request("fltFPCS_Hours_Goal")) & "'," & _	
			 "convert(datetime,'" & application.contents("dtSchool_Year_Start" & session.contents("intSchool_Year"))  & "')," & _
			 session.Contents("intSchool_Year") & ", '" & Session.Value("strUserID") & "'," & _
			"'" & oFunc.ConvertCheckToBit(request("bolMasters_Degree")) & "'," & _
			"'" & oFunc.EscapeTick(request("szSalary_Placement")) & "')"
	oFunc.ExecuteCN(insert)
end function

function vbfVerifyProfile
	insert = "insert into tblInstr_Verify(intInstructor_ID,szItem_Verified) " & _
			 "values (" & Request.Form("intInstructor_ID") & ",'Verified Profile')"
	oFunc.ExecuteCN(insert)

	' Delete this forced action from users list
	delete = "delete from tascUsers_Action " & _
			 "where intUser_Action_ID = " & session.Value("arActions")(fvCount,2)
	oFunc.ExecuteCN(delete)
	
	oFunc.CommitTransCN
	oFunc.CloseCN
	set oFunc = nothing
	
	'Erase this action from the array so it doesn't get executed again
	arEdit = session.Value("arActions")
	arEdit(fvCount,1) = "" 
	session.Value("arActions") = arEdit
	
	for i = 0 to ubound(arEdit)
		if arEdit(i,1) <> "" then
			Response.Redirect(Application("strSSLWebRoot") & arEdit(i,1))
		end if 
	next 
	
	session.Value("bolActionNeeded") = false
	Response.Redirect(session.Value("strURL"))
	Response.End
end function

if request("bolWin") <> "" then
%>
<html>
<body bgcolor=white onLoad="alert('Teacher Profile Updated');window.opener.location.reload();window.opener.focus();window.close();">
</body>
</html>
<%
else
%>
<html>
<body bgcolor=white onLoad="window.location.href='<%=Application.Value("strSSLWebRoot")%>/forms/teachers/addTeacher.asp?intInstructor_ID=<% = intInstructor_ID %>';">
</body>
</html>
<%
'Doesn't work with Mac IE 5.0
'Response.Redirect("../../default.asp?strMessage=" & strMessage)
end if 
%>

