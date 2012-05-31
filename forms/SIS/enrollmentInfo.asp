<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		enrollmentInfo.asp
'Purpose:	Used to bring to light school conflicts and enrollment
'			percentages.
'Date:		3 June 2002
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim intForcedCount
dim intCount
dim strAlert
dim strEnrollmentHeader
dim strButton
dim strFirstName

dim oFunc		'wsc object
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()
oFunc.ResetSelectSessionVariables
if Request.QueryString("intCount") <> "" then
	intForcedCount = Request.QueryString("intCount")
elseif Request.Form("intCount") <> "" then
	intForcedCount = Request.Form("intCount")
end if 

if Request.Form("intStudent_ID") <> "" then
	' Comes from submited form
	intStudent_ID = Request.Form("intStudent_ID")
	if Request.Form("intEnroll_Info_ID") = "" then
		call vbfAddEnrollInfo
	else
		call vbfUpdateEnrollInfo
	end if 
elseif Request.QueryString("intStudent_ID") <> "" then
	intStudent_ID = Request.QueryString("intStudent_ID")
else
	session.Value("arStudents") = ""
	session.Value("intCount") = ""
end if 

if isArray(session.Value("arStudents")) = false and request("bolForced") <> "" then
	'Student State Matrix Values:
    '1	=	Re-Enrollment Letter was Sent
    '2	=	Re-Enrollment Letter was Received
	'4	=	Responded with ReEnroll = Yes
	'8	=	First Phone call has occurred
	'16=	Second Phone call has occurred
	'32=	Letter of Dismissal was Sent
	'64=	Student is an Exit Candidate
	'
	'possible combinations
	'7	=	Perfect				=ENROLLED
	'15=	1 Call Enroll		=ENROLLED
	'31=	2 Call ReEnroll	=ENROLLED
	'1	=	Hanging
	'9	=	1 Call
	'25=	2 Call
	'57=	Dismiss
	'121=	Exit Dismiss
	'67=	Not ReEnroll
	'75=	1 Call Exit
	'86=    Withdrawn
	'91=	2 Call Exit	
	'123=   Early Graduation

	dim strStudents
	dim sql
	set rsGetStudents = server.CreateObject("ADODB.RECORDSET")
	rsGetStudents.CursorLocation = 3
		  
	'bkm 23-jun-02 - restrict to those who have a Student State of Re-Enrolled - see comments
	'above for matrix values/definitions
	sql = "SELECT s.intSTUDENT_ID, s.szFIRST_NAME, s.szLAST_NAME, s.intGRAD_YEAR, ss.szGRADE " & _
		"FROM tblSTUDENT s INNER JOIN " & _
		"tblStudent_States ss ON s.intSTUDENT_ID = ss.intStudent_id " & _
		"WHERE intFamily_ID = " & Session.Value("intFamily_ID") & " AND (NOT EXISTS " & _
		"	(SELECT e.intstudent_id " & _
		"	FROM tblEnroll_Info e " & _
		"	WHERE e.sintSchool_Year = " & Session.Value("intSchool_Year") & " AND s.intStudent_ID = e.intStudent_ID)) " & _ 
		"AND (ss.intSchool_Year = " & Session.Value("intSchool_Year") & ") AND (ss.intReEnroll_State IN (" & Application.Contents("ActiveEnrollList") & ") )"		  
		  
	rsGetStudents.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
	
	if rsGetStudents.RecordCount > 0 then
		redim arStudents(rsGetStudents.RecordCount,5)
		intCount = 0 
		do while not rsGetStudents.EOF
			arStudents(intCount,0) =  rsGetStudents("intStudent_ID") 
			arStudents(intCount,1) =  rsGetStudents("szFirst_Name") 
			arStudents(intCount,2) =  rsGetStudents("szLast_Name") 
			arStudents(intCount,3) =  rsGetStudents("intGrad_year") 
			arStudents(intCount,4) =  rsGetStudents("szGrade") 
			intCount = intCount + 1
			rsGetStudents.MoveNext
		loop
		rsGetStudents.Close
		set rsGetStudents = nothing
		session.Value("arStudents") = arStudents
		session.Value("intStudent_Count") = 0
		intCount = 0
	else
		' No Students to process so it's sensless to force this action 
		' We erase this forced action and go on to the next step
		if Request.Form("intCount") then
			fvCount = Request.Form("intCount")
		else
			fvCount = Request.QueryString("intCount")
		end if
		call oFunc.ForcedActionHandling(oFunc,fvCount)
	end if 
elseif session.Value("bolActionNeeded") = true then
	arStudents = session.Value("arStudents")	
	intCount = session.Value("intStudent_Count")
else
	' Get existing Enroll Record
	sql = "select e.intEnroll_Info_ID, e.intStudent_ID, s.szFirst_Name, s.szLast_Name, " & _
		  "e.szPrivate_School_Name,e.szOther_District_Name,e.intPercent_Enrolled_D2," & _
		  "e.intPercent_Enrolled_Fpcs,e.bolCharter_Grad, se.bolASD_Contract_HRS_Exempt, " & _
		  "se.szHRS_Exempt_Reason,se.intStudent_Exemption_ID," & _
		  "se.intCore_Credit_Percent,se.szCore_Exemption_Reason," & _
		  "se.szElective_Exemption_Reason,se.intElective_Credit_Percent " & _
		  "from tblEnroll_Info e, tblStudent s LEFT OUTER JOIN " & _
		  " tblStudent_Exemptions se on (s.intStudent_ID = se.intStudent_ID " & _
		  " and se.intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _
		  "where e.intStudent_ID = " & intStudent_ID & _
		  " and e.intStudent_ID = s.intStudent_ID " & _
		  " and e.sintSchool_Year = " & session.Contents("intSchool_Year")
		  
	set rsEnrollInfo = server.CreateObject("ADODB.RECORDSET")
	rsEnrollInfo.CursorLocation = 3
	rsEnrollInfo.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
	
	if rsEnrollInfo.RecordCount > 0 then
		'This for loop dimentions and defines all the columns we selected in sql
		'and we use the variables created here to populate the form.
		for each item in rsEnrollInfo.Fields
			execute("dim " & item.Name)
			execute(item.Name & " = item")		
		next  
	end if 
	strEnrollmentHeader = "<B>Enrollment Information for: " & szFirst_Name & " " & szLast_Name & "</b>"
	strFirstName = szFirst_Name
	strButton = "<input type=submit value=""Home Page"" class=""btSmallGray"" onClick=""window.location.href='" & Application("strSSLWebRoot") & "';"">" & chr(13) & _
				"<input type=submit value=""Update"" class=""NavSave"" onClick=""jfCheckPercent();"">"
	call vbfPrintEnrollForm
end if

if isArray(arStudents) then
	if intCount < ubound(arStudents,1) then 
		strFirstName = arStudents(intCount,1)
		strEnrollmentHeader = " <b>Enrollment Step 2 " & arStudents(intCount,1) & " " &  arStudents(intCount,2) & " </b> (Student " & intCount +1 & " of " & ubound(arStudents,1) & ")"
		intStudent_ID = arStudents(intCount,0)
		strButton = "<input type=submit value=""Next >"" class=""btSmallGray"" onClick=""jfCheckPercent();"">"
		call vbfPrintEnrollForm
	end if
end if 
function vbfPrintEnrollForm()
	
	Session.Value("strTitle") = "Enrollment Step 2"
	Session.Value("strLastUpdate") = "03 June 2002"
	if request("bolForced") <> "" then
		Server.Execute(Application.Value("strWebRoot") & "includes/simpleheader.asp")
	else
		Server.Execute(Application.Value("strWebRoot") & "includes/header.asp")
	end if
%>
<form action=enrollmentInfo.asp name=main method=post onSubmit="return false;">
<input type=hidden name=intCount value="<% = intForcedCount%>">
<input type=hidden name="intStudent_ID" value="<%=intStudent_ID%>">
<input type=hidden name="intEnroll_Info_ID" value="<%=intEnroll_Info_ID%>">
<table width=100%>
	<tr>	
		<Td class=yellowHeader >
				&nbsp;<% = strEnrollmentHeader %>
		</td>
	</tr>
	<tr>
		<Td>
			<table>
				<tr>	
					<Td colspan=2>
						<font class=svplain11>
							<b><i>Enrollment Status for School Year: <% = session.Contents("intSchool_Year")%></I></B>   
						</font>
					</td>
				</tr>
				<tr>	
					<Td class=gray>
						&nbsp;If  <% = strFirstName  %> will be enrolled in a private school for this
						school year what is the name of that school?
					</td>
				</tr>	
				<tr>						
					<td>
						<input type=text name="szPrivate_School_Name" value="<%=szPrivate_School_Name%>" maxlength=128 size=40>
					</td>
				</tr>	
				<tr>	
					<Td class=gray>
						&nbsp;If  <% = strFirstName %> will be enrolled in another school district 
						(such as Galena (IDEA), Alyeska Central School, Nenana (CyberLinks), etc.) what
						percentage will <% = strFirstName %> be enrolled in that school district?
					</td>
				</tr>	
				<tr>						
					<td>
						<select name="intPercent_Enrolled_D2">
							<option>
						<%
							Response.Write oFunc.MakeList("25,50,75,100","25%,50%,75%,100%",intPercent_Enrolled_D2)
						%>
						</select>
					</td>
				</tr>	
				<tr>	
					<Td class=gray>
						&nbsp;What percentage are you planning to enroll <% = strFirstName %> in FPCS
						for this new year?
					</td>	
				</tr>	
				<tr>					
					<td>
						<select name="intPercent_Enrolled_FPCS">
							<option>
						<%
							Response.Write oFunc.MakeList("25,50,75,100","25%,50%,75%,100%",intPercent_Enrolled_FPCS)
						%>
						</select>
					</td>
				</tr>		
				<tr>	
					<Td class=gray>
						&nbsp;Do you plan for <% = strFirstName %> to graduate from FPCS?
					</td>	
				</tr>	
				<tr>					
					<td>
						<select name="bolCharter_Grad">
						<%
							Response.Write oFunc.MakeList("1,0","Yes,No",oFunc.TrueFalse(bolCharter_Grad))
						%>
						</select>
					</td>
				</tr>	
				<% if session.Contents("strRole") = "ADMIN" then %>
				<tr>	
					<Td class=gray>
						&nbsp;Does this student qualify for an ASD Contract Hours Exemption?						
					</td>	
				</tr>			
				<tr>					
					<td class=svplain10>
						<select name=bolASD_Contract_HRS_Exempt ID="Select1">
						<%
							response.Write oFunc.MakeList("0,1","No,Yes",oFunc.TrueFalse(bolASD_Contract_HRS_Exempt))
						%>						
						</select>	
						If yes please give reason: 
						<input type=text name=szHRS_Exempt_Reason value="<%=szHRS_Exempt_Reason%>" maxlength=511 size=25>			
						<input type=hidden name="intStudent_Exemption_ID" value="<%=intStudent_Exemption_ID%>">
					</td>
				</tr>
				<tr>	
					<Td class=gray>
						&nbsp;If this student qualifies for Core or Elective credit exemptions enter the
						percent of exemption the student qualifies for.						
					</td>	
				</tr>			
				<tr>					
					<td class=svplain10>
						Core Credit Exemption%:<input type=text name="intCore_Credit_Percent" size=4 value="<%=intCore_Credit_Percent%>">
						 Exemption Reason: <input type=text name="szCore_Exemption_Reason" value="<%=szCore_Exemption_Reason%>" size=15 maxlength=511>						 
					</td>
				</tr>
				<tr>					
					<td class=svplain10>
						Elective Credit Exempt%:<input type=text name="intElective_Credit_Percent" value="<%=intElective_Credit_Percent%>" size=4 ID="Text1">
						 Exemption Reason: <input type=text name="szElective_Exemption_Reason" value="<%= szElective_Exemption_Reason %>" size=15 ID="Text2" maxlength=511>						 
					</td>
				</tr>
			  <% end if %>		
			</table>
		</td>
	</tr>
</table>
<% = strButton %>
</form>
<script language=javascript>
	<% = strAlert %>
	function jfCheckPercent() {
		if (main.intPercent_Enrolled_FPCS.value == "") {
			alert("You must enter the percentage you plan for your child to enroll in FPCS for the new year.");
			return false;
		}else{
			main.submit();
		}
	}
	
	<%
	if request.querystring("intCount") <> "" then
		response.write "jfInstruct();"
	end if
	%>
	function jfInstruct(){
		var message = "We now are at step 2 of 3 of the Initial Online Enrollment Process. This form asks some simple questions ";
		message += " about school participation for the new school year for each of the students you have enrolled."
		alert(message);
	}
</script>
<%
	Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")
end function

function vbfAddEnrollInfo
	dim insert
	
	' Insert records
	insert = "insert into tblEnroll_Info (intStudent_ID,sintSchool_Year," & _
			  "szPrivate_School_Name,szOther_District_Name,intPercent_Enrolled_D2," & _
			  "intPercent_Enrolled_Fpcs,bolCharter_Grad,szUser_Create) values (" & _
			 Request.Form("intStudent_ID") & "," & _
			 session.Value("intSchool_Year") & "," & _
			 "'" & oFunc.EscapeTick(Request.Form("szPrivate_School_Name")) & "'," & _
			 "'" & oFunc.EscapeTick(Request.Form("szOther_District_Name")) & "'," & _				 
			 "'" & oFunc.CheckDecimal(Request.Form("intPercent_Enrolled_D2")) & "'," & _
			 "'" & oFunc.CheckDecimal(Request.Form("intPercent_Enrolled_Fpcs")) & "'," & _
			 "'" & oFunc.EscapeTick(Request.Form("bolCharter_Grad")) & "'," & _
			 "'" & oFunc.EscapeTick(session.Value("strUserID")) & "')" 
	oFunc.ExecuteCN(insert)		 

	'Check to see if we need to loop to the next student 	
	session.Value("intStudent_Count") = session.Value("intStudent_Count") + 1
	arStudents = session.Value("arStudents")
	
	if session.Contents("strRole") = "ADMIN" and _ 
			(request.Form("bolASD_Contract_HRS_Exempt") = 1 or _
			request("intCore_Credit_Percent") <> "" or _
			request("szElective_Exemption_Reason") <> "") then
		call vbsInsertExemption
	end if
	
	if isArray(arStudents) then
		if session.Value("intStudent_Count") < ubound(arStudents,1) then 
			'countinue
		else
			'Finished going through each student in our forced action loop
			if Request.Form("intCount") <> "" then
				call oFunc.ForcedActionHandling(oFunc,Request.Form("intCount"))
			end if		
			Response.End
		end if 
	end if
end function


sub vbsInsertExemption
	' Inserts an Exemption record
	insert = "insert into tblStudent_Exemptions(intStudent_ID,intSchool_Year," &_
				 "bolASD_Contract_HRS_Exempt,szHRS_Exempt_Reason,intCore_Credit_Percent," & _
				 "szCore_Exemption_Reason,intElective_Credit_Percent," & _
				 "szElective_Exemption_Reason,szUser_Create,dtCreate) " & _
				 "values (" & _
				 request.Form("intStudent_ID") & "," & _
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

function vbfUpdateEnrollInfo
	'Updates an Enrolment Record
	update = "update tblEnroll_Info set " & _
			 "szPrivate_School_Name = '" & oFunc.EscapeTick(Request.Form("szPrivate_School_Name")) & "'," & _
			 "szOther_District_Name = '" & oFunc.EscapeTick(Request.Form("szOther_District_Name")) & "'," & _
			 "intPercent_Enrolled_D2 = '" & oFunc.CheckDecimal(Request.Form("intPercent_Enrolled_D2")) & "'," & _
			 "intPercent_Enrolled_Fpcs = '" & oFunc.CheckDecimal(Request.Form("intPercent_Enrolled_Fpcs")) & "'," & _
			 "bolCharter_Grad = '" & oFunc.EscapeTick(Request.Form("bolCharter_Grad")) & "'," & _
			 "szUser_Modify = '" & oFunc.EscapeTick(Request.Form("strUserID")) & "'," & _
			 "dtModify = '" & now() & "' " & _
			 "where intEnroll_Info_ID = " & Request.Form("intEnroll_Info_ID")			 
	oFunc.ExecuteCN(update)
	
	' Update or Insert Exemptions if needed
	if request.Form("intStudent_Exemption_ID") <> "" then
		update = "update tblStudent_Exemptions " & _
				 "set bolASD_Contract_HRS_Exempt = '" & request.Form("bolASD_Contract_HRS_Exempt") & "'," & _
				 "szHRS_Exempt_Reason = '" & oFunc.EscapeTick(request.Form("szHRS_Exempt_Reason")) & "', " & _
				 "intCore_Credit_Percent = " & oFunc.CheckDecimal(request("intCore_Credit_Percent")) & "," & _
				 "szCore_Exemption_Reason = '" & oFunc.EscapeTick(request("szCore_Exemption_Reason")) & "'," & _
				 "intElective_Credit_Percent = " & oFunc.CheckDecimal(request("intElective_Credit_Percent")) & "," & _
				 "szElective_Exemption_Reason = '" & oFunc.EscapeTick(request("szElective_Exemption_Reason")) & "' " & _				 
				 "where intStudent_Exemption_ID = " & request.Form("intStudent_Exemption_ID")
		oFunc.ExecuteCN(update)
	else
		if session.Contents("strRole") = "ADMIN" and _
			(request.Form("bolASD_Contract_HRS_Exempt") = 1 or _
			request("intCore_Credit_Percent") <> "" or _
			request("szElective_Exemption_Reason") <> "") and _
			request.Form("intStudent_Exemption_ID") = "" then
		 
			call vbsInsertExemption
		end if
	end if
	
	strAlert="alert('Enrollment Information has been Updated');"
end function

oFunc.CloseCN()
set oFunc = nothing

%>