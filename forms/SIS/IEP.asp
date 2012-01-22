<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		IEP.asp  
'Purpose:	This script contains the Special Education Students (IEP)
'			section that was formerly in studentProfile.asp
'Date:		04 Sept 2002
'Author:	Bryan K Mofley (ThreeShapes.com LLC)
'
'rev:		14-May-2003 BKM - removed javascript - assigned to tblIEP instead of tblStudent
'rev:		17-May-2003 BKM 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim oFunc				'windows scripting component generalized functions
dim objRequest			'will contain either the FORM or QUERYSTRING object
'dim intStudent_id
dim mstrMessage			'Success for Fail message for new insert
dim mstrValidationError	'Error message indicating which items failed validation

	Session.Value("strTitle") = "Student IEP"
	Session.Value("strLastUpdate") = "17 May 2003"

	set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
	call oFunc.OpenCN()

	if Request.Form.Count > 0 then
		set objRequest = Request.Form
	else
		set objRequest = Request.QueryString
	end if

	intStudent_id = objRequest("intStudent_id")
	if objRequest("cmdSubmit") <> "" then
		mstrValidationError = vbfValidate(objRequest)
		if mstrValidationError <> "" then
			Response.Write "<font class=svError>" & mstrValidationError & "</font>"
		else
			call vbfInsert(objRequest)
			Response.Write mstrMessage
		end if
	end if

	'**************************************************************
	'Section:	Populate Form
	'Purpose:	populates the IEP form with data from tblIEP or Form object
	'note:		this section is not in a function because variables
	'			are dynamically DIM'd and need to have module
	'			level scope
	'**************************************************************
	if mstrValidationError <> "" then
		'dimention local variables from the form object
		'we'll use these variables in the embedded HTML to populate the form
		for each item in objRequest
			execute("dim " & item)
			execute(item & " = """ & objRequest(item) & """")
		next 
	elseif intStudent_id <> "" then
		'dimention local variables from tblIEP for the given student
		'we'll use these variables in the embedded HTML to populate the form
		dim rsStudent
		dim sqlStudent
		dim intCount
		dim item
		dim Exp_Month
		dim Exp_Day
		dim Exp_Year
		dim Next_Eval_Month
		dim Next_Eval_Day
		dim Next_Eval_Year
			
		set rsStudent = Server.CreateObject("ADODB.RECORDSET")
		rsStudent.CursorLocation = 3
		'grab the most recent entry in tblIEP for this student
		sqlStudent =	"SELECT tblIEP.* " & _
						"FROM   tblIEP " & _
						"WHERE  (intStudent_ID = " & intStudent_id & ") AND " & _
						"	(dtCREATE = " & _
						"		(SELECT MAX(dtCREATE) " & _
						"        FROM  tblIEP " & _
						"        WHERE (intStudent_ID = " & intStudent_id & ")))"
					
		rsStudent.Open sqlStudent,oFunc.FPCScnn
		if not rsStudent.BOF and not rsStudent.EOF then
			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'' This for loop will dimension AND assign our student info variables
			'' for us. We'll use them later to populate the form.
			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
			intCount = 0
			for each item in rsStudent.Fields
				execute("dim " & rsStudent.Fields(intCount).Name)
				execute(rsStudent.Fields(intCount).Name & " = item")
				intCount = intCount + 1
			next							
				
			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'' The ILP dates are stored in the database as a single field, but our form 
			'' displays it as three seperate select lists so we break the single
			'' date up to use the parts in our form populating.
			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''		
			Exp_Month = datePart("m",dtExpiration)
			Exp_Day = datePart("d",dtExpiration)
			Exp_Year = datePart("yyyy",dtExpiration)
			
			Next_Eval_Month = datePart("m",dtNext_Eval)
			Next_Eval_Day = datePart("d",dtNext_Eval)
			Next_Eval_Year = datePart("yyyy",dtNext_Eval)		
		else
			'Response.Write "Student ID " & Session.Value("intStudent_id") & " is not a valid ID"
			'Session.Value("intStudent_id") = ""		
			'Session.Value("SISEditMode") = ""	
		end if
		rsStudent.Close
		
		'now we grab the students name for non-admin users
		sqlStudent = "SELECT szLAST_NAME, szFIRST_NAME, sMID_INITIAL " & _
					 "FROM   tblSTUDENT " & _
					 "WHERE (intSTUDENT_ID = " & intStudent_id & ")"
		rsStudent.Open sqlStudent,oFunc.FPCScnn
		if not rsStudent.BOF and not rsStudent.EOF then	
			intCount = 0
			for each item in rsStudent.Fields
				execute("dim " & rsStudent.Fields(intCount).Name)
				execute(rsStudent.Fields(intCount).Name & " = item")
				intCount = intCount + 1
			next
		end if	
		rsStudent.Close				 
		set rsStudent = nothing	
	end if 
	'**************************************************************
	'End Section:	Populate Form
	'**************************************************************
	Server.Execute(Application.Value("strWebRoot") & "includes/Header.asp")

	if Session.Contents("strRole") = "ADMIN" then
%>
	<script language="javascript">
		function jfChangeStudent(obj){
		//reloads page with newly selected student
			var strURL = "<% = Application.Value("strWebRoot")%>forms/SIS/IEP.asp?intStudent_ID=" + obj.value;
			window.open(strURL, "_self");
		}
	</script>
<%	end if %>
	
	<form name=main action="iep.asp" method=post>
	<input type="hidden" name="intSchool_Year" value="<%=Session.Contents("intSchool_Year")%>">
	<input type=hidden name="isFamManager" value="<%=request("isFamManager")%>">
	<table width=100%>
		<tr>	
			<td colspan=2 class=yellowHeader>
					&nbsp;<b>Special Education Students</b>&nbsp;&nbsp;&nbsp;
			<% if Session.Value("strRole") = "ADMIN" then %>	
					<select name="intStudent_ID" onchange="jfChangeStudent(this);">
						<option value="">
					<%
						dim sqlStudentName
						'sqlStudentName = "SELECT intStudent_ID,szLast_Name + ',' + szFirst_Name AS Name " & _
						'				"FROM tblStudent ORDER BY szLast_Name"
						sqlStudentName = "SELECT s.intSTUDENT_ID, " & _
									"Name = (Case ss.intReEnroll_State WHEN 86 then " & _
									"s.szLAST_NAME + ',' + s.szFIRST_NAME + ': Withdrawn (' + convert(varChar(20),ss.dtModify) + ')'" & _ 
									"WHEN 123 THEN s.szLAST_NAME + ',' + s.szFIRST_NAME + ': Graduated (' + convert(varChar(20),ss.dtModify) + ')'" & _ 
									"ELSE s.szLAST_NAME + ',' + s.szFIRST_NAME END) " & _
									"FROM tblSTUDENT s INNER JOIN " & _ 
									"tblStudent_States ss ON s.intSTUDENT_ID = ss.intStudent_id " & _ 
									"WHERE (ss.intReEnroll_State in (" & application.Contents("strEnrollmentList") & ")) AND (ss.intSchool_Year = " & Session.Value("intSchool_Year") & ") " & _ 
									"ORDER BY Name"										
						Response.Write oFunc.MakeListSQL(sqlStudentName,"intStudent_ID","Name",intStudent_id)												 
					%>
					</select>
			<% else 
				Response.Write szLast_Name & ", " & szFirst_Name & " " & sMID_INITIAL %>
				<input type=hidden name="intStudent_ID" value="<% = intStudent_id%>">
				<input type=hidden name="szLAST_NAME"   value="<% = szLAST_NAME%>">
				<input type=hidden name="szFIRST_NAME"  value="<% = szFIRST_NAME%>">
				<input type=hidden name="sMID_INITIAL"  value="<% = sMID_INITIAL%>">
			<% end if %>
			</td>
		</tr>
		<tr>	
			<td colspan=2 class=gray>					
					Has your child ever been eligible for Special Education Services?<BR>
					<b>Yes</b><input type=radio name="bolIEP"  value="yes" <% if bolIEP = true or bolIEP = "yes" then Response.Write " checked " %> >
					<b>No</b><input type=radio name="bolIEP"  value="no" <% if (bolIEP <> "" AND bolIEP = false) or  bolIEP = "no" then Response.Write " checked " %> >
					
			</td>
		</tr>
		<tr>
			<td class=gray colspan=2>
					If your child has ever been certified in a Special Education area, 
					please check it below:
			</td>										
		</tr>
		<tr>
			<td colspan=2>
				<table>
					<tr>
						<td class=svplain10>
							<nobr>Gifted/Talented</nobr>
						</td>
						<td>
							<input type=checkbox value="true" name="bolGifted" <% if bolGifted then response.Write " checked "%>>
						</td>
						<td class=svplain10>
							<nobr>Autism/Asberger's Syndrome</nobr>
						</td>
						<td>
							<input type=checkbox value="true" name="bolAutism" <% if bolAutism then response.Write " checked "%> ID="Checkbox1">
						</td>
						<td class=svplain10>
							<nobr>Deafness</nobr>
						</td>
						<td>
							<input type=checkbox value="true" name="bolDeafness" <% if bolDeafness then response.Write " checked "%> ID="Checkbox2">
						</td>
					</tr>
					<tr>
						<td class=svplain10>
							<nobr>Deaf-Blindness</nobr>
						</td>
						<td>
							<input type=checkbox value="true" name="bolDeaf_Blindness" <% if bolDeaf_Blindness then response.Write " checked "%> ID="Checkbox3">
						</td>
						<td class=svplain10>
							<nobr>Early Childhood Developmental Delay</nobr>
						</td>
						<td>
							<input type=checkbox value="true" name="bolDev_Delay" <% if bolDev_Delay then response.Write " checked "%> ID="Checkbox4">
						</td>
						<td class=svplain10>
							<nobr>Emotional Disturbance</nobr>
						</td>
						<td>
							<input type=checkbox value="true" name="bolEmotional" <% if bolEmotional then response.Write " checked "%> ID="Checkbox5">
						</td>
					</tr>
					<tr>
						<td class=svplain10>
							<nobr>Hearing Impairment</nobr>
						</td>
						<td>
							<input type=checkbox value="true" name="bolHearing" <% if bolHearing then response.Write " checked "%> ID="Checkbox6">
						</td>
						<td class=svplain10>
							<nobr>Specific Learning Disability</nobr>
						</td>
						<td>
							<input type=checkbox value="true" name="bolLearning_Dis" <% if bolLearning_Dis then response.Write " checked "%> ID="Checkbox7">
						</td>
						<td class=svplain10>
							<nobr>Mental Retardation</nobr>
						</td>
						<td>
							<input type=checkbox value="true" name="bolMental_Retardation" <% if bolMental_Retardation then response.Write " checked "%> ID="Checkbox8">
						</td>
					</tr>
					<tr>
						<td class=svplain10>
							<nobr>Multiple Disability</nobr>
						</td>
						<td>
							<input type=checkbox value="true" name="bolMulti_Dis" <% if bolMulti_Dis then response.Write " checked "%> ID="Checkbox9">
						</td>
						<td class=svplain10>
							<nobr>Orthopedic Impairment</nobr>
						</td>
						<td>
							<input type=checkbox value="true" name="bolOrthopedic" <% if bolOrthopedic then response.Write " checked "%> ID="Checkbox10">
						</td>
						<td class=svplain10>
							<nobr>Other Health Impairment</nobr>
						</td>
						<td>
							<input type=checkbox value="true" name="bolOther_Health" <% if bolOther_Health then response.Write " checked "%> ID="Checkbox11">
						</td>
					</tr>
					<tr>
						<td class=svplain10>
							<nobr>Traumatic Brain Injury</nobr>
						</td>
						<td>
							<input type=checkbox value="true" name="bolBrain_Injury" <% if bolBrain_Injury then response.Write " checked "%> ID="Checkbox12">
						</td>
						<td class=svplain10>
							<nobr>Speech or Language Impairment</nobr>
						</td>
						<td>
							<input type=checkbox value="true" name="bolSpeech_Impairment" <% if bolSpeech_Impairment then response.Write " checked "%> ID="Checkbox13">
						</td>
						<td class=svplain10>
							<nobr>Visual Impairment</nobr>
						</td>
						<td>
							<input type=checkbox value="true" name="bolVisual_Impairment" <% if bolVisual_Impairment then response.Write " checked "%> ID="Checkbox14">
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>	
			<td colspan=2 class=gray>					
					Has your child been receiving special education services?
					<b>Yes</b><input type=radio name="bolCurrent_SES"  value="yes" <% if  bolCurrent_SES = true or bolCurrent_SES = "yes" then Response.Write " checked " %> >
					<b>No</b><input type=radio name="bolCurrent_SES"  value="no" <% if  (bolCurrent_SES <> "" AND bolCurrent_SES = false) or bolCurrent_SES = "no" then Response.Write " checked " %> >
					
			</td>
		</tr>	
		<tr>	
			<td colspan=2 class=gray>					
					Has your child been formally exited from special education?
					<b>Yes</b><input type=radio name="bolFormal_Exit"  value="yes" <% if  bolFormal_Exit = true or bolFormal_Exit = "yes" then Response.Write " checked " %> >
					<b>No</b><input type=radio name="bolFormal_Exit"  value="no" <% if  (bolFormal_Exit <> "" AND bolFormal_Exit = false) or bolFormal_Exit = "no" then Response.Write " checked " %> >
					
			</td>
		</tr>
		<tr>	
			<td class=gray>					
					What is the expiration date of the most current IEP?
			</td>
			<td>
				<table>
					<tr>
						<td>
							<select name="Exp_month">
								<option value="">
								<% 
								sqlMonth = "select strValue,strText from common_lists where intList_ID = 1 order by intOrder"
								Response.Write oFunc.MakeListSQL(sqlMonth,"","",Exp_Month)								
								%>
							</select>
						</td>		
						<td>
							<select name="Exp_day">
								<option value="">
								<% 
								sqlDay = "select strValue,strText from common_lists where intList_ID = 2 order by intOrder"
								Response.Write oFunc.MakeListSQL(sqlDay,"","",Exp_Day)								
								%>
							</select>
						</td>											
						<td>
							<select name="Exp_year">	
								<option value="">
								<% = oFunc.MakeYearList(4,2,Exp_Year) %>
							</select>
						</td>		
					</tr>
				</table>	
			</td>
		</tr>	
		<tr>	
			<td class=gray>					
					When is the next eligibility evaluation due?
			</td>
			<td>
				<table>
					<tr>
						<td>
							<select name="Next_Eval_month">
								<option value="">
								<% 
								sqlMonth = "select strValue,strText from common_lists where intList_ID = 1 order by intOrder"
								Response.Write oFunc.MakeListSQL(sqlMonth,"","",Next_Eval_Month)								
								%>
							</select>
						</td>		
						<td>
							<select name="Next_Eval_day">
								<option value="">
								<% 
								sqlDay = "select strValue,strText from common_lists where intList_ID = 2 order by intOrder"
								Response.Write oFunc.MakeListSQL(sqlDay,"","",Next_Eval_Day)								
								%>
							</select>
						</td>											
						<td>
							<select name="Next_Eval_year">	
								<option value="">
								<% = oFunc.MakeYearList(4,2,Next_Eval_Year) %>
							</select>
						</td>		
					</tr>
				</table>	
			</td>
		</tr>
		<tr>
			<td colspan=2 class=yellowHeader>&nbsp;<b>Bilingual/Multicultural</b></td>
		</tr>
		<tr>	
			<td colspan=2 class=gray>					
				Has your child ever learned another language besides English?
					<b>Yes</b><input type=radio name="bolBilingual"  value="yes" <% if  bolBilingual = true or bolBilingual = "yes" then Response.Write " checked " %> >
					<b>No</b><input type=radio name="bolBilingual"  value="no" <% if  (bolBilingual <> "" AND bolBilingual = false) or bolBilingual = "no" then Response.Write " checked " %> >
					
			</td>
		</tr>	
		<tr>
			<td class=gray>
				If so, what language or culture has your child experienced or learned?
			</td>
			<td>
				<select name="intLanguage_ID">
					<option value="">- - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				<%							
					dim sqlLanguage
					sqlLanguage = "select intLanguage_id,szLanguage_Desc from trefLanguage order by szLanguage_Desc"
					Response.Write oFunc.MakeListSQL(sqlLanguage,"","",intLanguage_ID)
				%>
				</select>				
			</td>
		</tr>	
		<tr>	
			<td colspan=2 class=gray>					
				Is your child eligible for the Bilingual Program?
					<b>Yes</b><input type=radio name="bolBilingual_Elig"  value="yes" <% if  bolBilingual_Elig = true or bolBilingual_Elig = "yes" then Response.Write " checked " %> >
					<b>No</b><input type=radio name="bolBilingual_Elig"  value="no" <% if  (bolBilingual_Elig <> "" AND bolBilingual_Elig = false) or bolBilingual_Elig = "no" then Response.Write " checked " %> >
					
			</td>
		</tr>
		<tr>	
			<td colspan=2 class=gray>					
				Has your child currently been receiving services in the Bilingual Program?
					<b>Yes</b><input type=radio name="bolBilingual_Current"  value="yes" <% if  bolBilingual_Current = true or bolBilingual_Current = "yes" then Response.Write " checked " %> >
					<b>No</b><input type=radio name="bolBilingual_Current"  value="no" <% if  (bolBilingual_Current <> "" AND bolBilingual_Current = false) or bolBilingual_Current = "no" then Response.Write " checked " %> >
					
			</td>
		</tr>	
		<tr>
			<td colspan=2 class=yellowHeader>&nbsp;<b>Migrant Education</b></td>
		</tr>
		<tr>	
			<td colspan=2 class=gray>					
				Has your child been receiving services under the Migrant Education Program?
					<b>Yes</b><input type=radio name="bolMigrant_ED"  value="yes" <% if  bolMigrant_ED = true or bolMigrant_ED = "yes" then Response.Write " checked " %> >
					<b>No</b><input type=radio name="bolMigrant_ED"  value="no" <% if  (bolMigrant_ED <> "" AND bolMigrant_ED = false) or bolMigrant_ED = "no" then Response.Write " checked " %> >
					
			</td>
		</tr>																	
		<tr>
			<td colspan=2>
				&nbsp;	
			</td>					
		</tr>
		<tr>
			<td class=gray colspan=2>
				PLEASE BE REMINDED THAT YOU WILL NEED TO HIRE A SPONSOR TEACHER
				TO ATTEND ANY AND ALL IEP MEETINGS WITH YOU.  IF YOU HAVE ANY 
				QUESTIONS PLEASE CALL THE OFFICE IMMEDIATELY.
			</td>
		</tr>
	</table>
	<% if request("isFamManager") <> "" then%>
	<input type=button value="Return to Family Manager" onClick="window.location.href='<%=Application("strWebRoot")%>admin/familyManager.asp';" class="btSmallGray" NAME="Button1">
	<%else%>
	<input type=button value="Home Page" onClick="window.location.href='<%=Application("strWebRoot")%>';" class="btSmallGray" >
	<%end if%>
	<input type=submit value="Update" class="NavSave" name="cmdSubmit">
	</form>	
		
<%
	call oFunc.CloseCN()
	set oFunc = nothing
	Server.Execute(Application.Value("strWebRoot") & "Includes/footer.asp")
	
'*************************************************************************
'functions/procedures below this line
'*************************************************************************
	
	function vbfValidate(pobjRequest)
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'Name:		vbfValidate 
	'Purpose:	Server side validation of the form prior to allowing inserts
	'Date:		14 May 2003
	'Author:	Bryan K Mofley (ThreeShapes.com LLC)
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''		
	dim strError		'Store any returned errors
	dim oVal			'validation wsc object
	dim dtExpiration	'fully qualified date
	dim dtNext_Eval		'fully qualified date
	
		'dimention all of the form/querystring objects
		for each item in pobjRequest
			execute("dim " & item)
			execute(item & " = """ & pobjRequest(item) & """")
		next 

		set oVal = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/formValidation.wsc"))
		oVal.validateField intStudent_ID,"blank","","Student" 		
	
		if bolIEP = "yes" then
			'at least one type of special education is requred if they indiacted their
			'child was eligible for Special Education Services
			if bolGifted & bolAutism & bolDeafness & bolDeaf_Blindness & bolDev_Delay & _
				bolEmotional & bolHearing & bolLearning_Dis & bolMental_Retardation & bolMulti_Dis & _
				bolOrthopedic & bolOther_Health & bolBrain_Injury & bolSpeech_Impairment & bolVisual_Impairment = "" then
				strError = strError & "You must select at least one Certified Special Education Area<BR>"
			end if
			
			if bolCurrent_SES = "" then
				strError = strError & "You must provide an answer for the question ... <BR>" & _
						   "&nbsp;&nbsp;'Has your child currently been receiving special education services?'<BR>"
			end if
			
		end if
	
		if bolCurrent_SES = "yes" then
			'Expiration and Eval dates are required if they indicated their child
			'has currently been receiving special education services
			if Exp_month = "" and Exp_day = "" and Exp_year = "" then
				strError = strError & "You must provide the expiration date of the current IEP.<BR>"
			end if	
			
			if Next_Eval_month = "" and Next_Eval_Day = "" and Next_Eval_Year = "" then
				strError = strError & "You must provide a date for when the next IEP evaluation is due.<BR>"
			end if 
		end if
		
		'removed the check for bolIEP=yes  BKM 12-July-2003
		if  (Exp_month & Exp_day & Exp_year <> "" OR Next_Eval_month & Next_Eval_day & Next_Eval_year <> "") then 'bolIEP = "yes" and 
			'if they supplied a date (even if it was not required) check to make sure it is valid
			dtExpiration = Exp_month & "/" & Exp_day & "/" & Exp_year
			dtNext_Eval = Next_Eval_month & "/" & Next_Eval_day & "/" & Next_Eval_year	
			oVal.validateField dtExpiration,"date","","Expiration" 
			oVal.validateField dtNext_Eval,"date","","Next Eligibility"		
		end if

		if bolBilingual = "yes" then		
 			oVal.validateField intLanguage_ID,"blank","","Language/Culture" 
		end if
		if oVal.ValidationError & "" <> "" then
			strError = strError & oVal.ValidationError 
		end if
		
		if strError <> "" then
			strError = "<BR><font color=red><b>The following items need to be corrected.</B><BR>" & strError & "</font>"
		end if
		
		vbfValidate = strError	
	end function
	
	function vbfInsert(pobjRequest)
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'Name:		vbfInsert 
	'Purpose:	Inserts a new record into tblIEP if necessary
	'Date:		14 May 2003
	'Author:	Bryan K Mofley (ThreeShapes.com LLC)
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	dim strSQL			'SQL for update statement
	dim strSQLfields	'SQL for INSERT field names
	dim strSQLvalues	'SQL for INSERT field values
	dim i				'counter in for next loop
	dim vntValue		'variant value of form field being passed to SQL statement
	
		' Since some of the Form objects will NOT be used in the SQL statement, there is no need
		' to turn the http header variables into vbs variables (this will actually mess up the
		' SQL statement being created below).  Instead, we only use those that begin
		' with our SQL field name standards (int, bol, or dt)
		
		for each i in pobjRequest
			if Left(i,3) = "int" or Left(i,3) = "bol" or Left(i,2) = "dt"  then
				strSQLfields = strSQLfields & i & ","
				select case pobjRequest(i)
					case "yes", "true"
						vntValue = 1
					case "no", "false"
						vntValue = 0
					case ""
						vntValue = "NULL"
					case else
						vntValue = pobjRequest(i)			
				end select
				strSQLvalues = strSQLvalues & vntValue & ","
			end if
		next
		
		if pobjRequest("Exp_month") <> "" and pobjRequest("Next_Eval_month") <> "" then
			'validation was laready done in vbfValidate so if the dates are blank, exclude them from the INSERT
			strSQLfields = strSQLfields & "dtExpiration,dtNext_Eval,"
			strSQLvalues = strSQLvalues & "'" & pobjRequest("Exp_month") & "/" & pobjRequest("Exp_day") & "/" & pobjRequest("Exp_year") & "',"
			strSQLvalues = strSQLvalues & "'" & pobjRequest("Next_Eval_month") & "/" & pobjRequest("Next_Eval_day") & "/" & pobjRequest("Next_Eval_year") & "',"
		end if
		
		strSQLfields = strSQLfields & "szUser_Create) "
		strSQLvalues = strSQLvalues & "'" & Session.Value("strUserID") & "')"
		
					
		strSQL = "INSERT INTO tblIEP (" 
		strSQL = strSQL & strSQLfields
		strSQL = strSQL & "VALUES (" & strSQLvalues 

		on error resume next
		oFunc.BeginTransCN
		oFunc.ExecuteCN(strSQL)
		oFunc.CommitTransCN
		
		'detect SQL errors and email developers if necessary
		if Err.number <> 0 then
			Session.Contents("ErrorNum") = Err.number
			Session.Contents("ErrorDesc") = Err.Description
			Server.Execute(Application.Value("strWebRoot") & "admin/debugEmailer.asp")		
			mstrMessage = "<font color=red>An error has occured.<br>A detailed error message has been mailed to the web developer.<br>" & _
				Session.Contents("ErrorNum") & "<br>" & Session.Contents("ErrorDesc") & "<br>" & _
				Err.Source & "</font>"
			Session.Contents("ErrorNum") = ""
			Session.Contents("ErrorDesc") = ""
		else
			mstrMessage = "<font class=svError><b>Student IEP Information was Updated.</b></font>"
		end if
		on error goto 0
	end function
	
%>