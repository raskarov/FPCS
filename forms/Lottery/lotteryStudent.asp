<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		StudentProfile.asp  
'Purpose:	This script collects the student information
'			or displays the student information.
'Date:		9 July 2001
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim oFunc			'Main object that exposes many of our custom functions 
dim oVal			'Form Validation Object
dim strUpdate		'Addition update informtion to be added to sql command
dim dtIEP			'IEP date info combined into one string
dim dtBirth			'Birth date info combined into one string

set oVal = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/formValidation.wsc"))
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

intStudent_id = request.QueryString("intStudent_id")
bolDelete = request.QueryString("bolDelete")
bolConfirm = request.QueryString("bolConfirm")

if bolDelete <> "" then
	call vbsConfirmDelete
elseif bolConfirm <> "" then
	call vbsDeleteStudent
end if 

if request.Form.Count > 0 then
	' Transfers all of the post http header variables into vbs variables
	' so we can more readily access them
	for each i in request.Form
		execute("dim " & i)
		execute(i & " = """ & request.Form(i) & """")
	next 
end if 

' Validate Form data for Student
if btSaveStudentInfo <> ""  then
	strError = vbfValidateStudent()
end if

if IEPyear <> "" then	
	dtIEP = "'" & IEPMonth & "/" & IEPDay & "/" & IEPYear & "'"
	strUpdate = "dtIEP_Renewal = " & dtIEP 	& ","
else
	strUpdate = "dtIEP_Renewal = NULL, "
	dtIEP = " NULL, " 
end if 

dtBirth = mMonth & "/" & mDay & "/" & mYear

' Create new student entry
if intStudent_ID = "" and btSaveStudentInfo <> "" and strError = "" then
	call vbsInsertStudent
elseif intStudent_ID <> "" and btSaveStudentInfo <> ""  and strError = "" then
	call vbsUpdateStudent
end if 


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' When the next if is true we get the student info 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if intStudent_id <> "" and btSaveStudentInfo = "" then
	dim rsStudent
	dim sqlStudent
	dim intCount
	dim item
	dim intCount2
		
	set rsStudent = Server.CreateObject("ADODB.RECORDSET")
	rsStudent.CursorLocation = 3
	sqlStudent = "select szFirst_Name,szLast_Name,sMid_Initial,szSSN," & _
				"sSex,dtBirth," & _
				"szGrade,intGrad_Year,intFirst_Lang,intHome_Lang, " & _
				"szPrevious_School,szPrev_School_Year,szPrev_School_Addr," & _
				"szPrev_School_City,szPrev_School_State,szPrev_School_Country," & _
				"szPrev_School_Zip_Code,szPrev_Anch_School,intPrev_Anch_Year," & _
				"szContact_Last_Name,szContact_First_Name,szContact_Phone," & _
				"szDR_Last_Name,szDR_First_Name,szDR_Phone,szDaycare_Name," & _
				"szDaycare_Phone,szMed_Alert_1,szMed_Alert_2,szDisability_1," & _
				"szDisability_2,bolIEP,dtIEP_Renewal,szExceptionality " & _					
				"from tblStudent where intStudent_ID=" & intStudent_id & _
				" and intFamily_ID = " & session.Contents("intFamily_ID")
					
	rsStudent.Open sqlStudent,Application("cnnFPCS")'oFunc.FPCScnn
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
			
		'set format for several fields
		'szSSN = oFunc.Reformat(szSSN, Array("", 3, "-", 2, "-", 4))
		'szPrev_School_Zip_Code = oFunc.Reformat(szPrev_School_Zip_Code, Array("", 5, "-", 4))
		'szDaycare_Phone = oFunc.Reformat(szDaycare_Phone, Array("(", 3, ") ", 3, "-", 4))
		'szDR_Phone = oFunc.Reformat(szDR_Phone, Array("(", 3, ") ", 3, "-", 4))
			
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'' The Birth data is stored in the database as a single field, but our form 
	'' displays it as three seperate select lists so we break the single
	'' date up to use the parts in our form populating.
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''		
		mMonth = datePart("m",dtBirth)
		mDay = datePart("d",dtBirth)
		mYear = datePart("yyyy",dtBirth)
		
		iepMonth = datePart("m",dtIEP_Renewal)
		iepDay = datePart("d",dtIEP_Renewal)
		iepYear = datePart("yyyy",dtIEP_Renewal)
	else
		Response.Write "Student ID " & Session.Value("intStudent_id") & " is not a valid ID"
		Session.Value("intStudent_id") = ""		
		Session.Value("SISEditMode") = ""	
	end if
	rsStudent.Close
	set rsStudent = nothing	
end if 
	
%>	
<html>
	<head>
		<title>Student Enrollment Page</title>
		<link rel="stylesheet" type="text/css" href="../../css/homestyle.css">
	</head>
	<body bgcolor="white">	
	<form action="lotteryStudent.asp" method=POST name=main ID="Form1">
	<input type=hidden name="intStudent_ID" value="<%=intStudent_ID%>">
	<table width=100% ID="Table1">
		<tr>	
			<Td class=navyHeader>
					&nbsp;<b>Student Enrollment Form</b>&nbsp;&nbsp;&nbsp;		
			</td>
		</tr>
		<tr>
			<td class="svplain10">
				<table cellspacing="0" cellpadding=4 bordercolor="e6e6e6" border="1" ID="Table15">
					<tr>
						<td class=svplain10>
							<b>Instructions:</b><br>
							Please enter your student information on this form. You
							must click the 'Save' button at the bottom of this form
							for the student information to be saved.
							 * denotes the information that is required
							in order to add a student to the enrollment database.
							<BR>
							<% = strError %>
						</td>
					</tr>
				</table>
				<br>
			</td>
		</tr>
		<tr>
			<td bgcolor=f7f7f7>
				<table ID="Table2">
					<tr>	
						<Td colspan=6>
							<font class=svplain11>
								<b><i>Students Information</I></B>
							</font>
						</td>
					</tr>
					<tr>
						<td class=gray>
								&nbsp;Legal Name: Last*
						</td>
						<td class=gray>
								&nbsp;First Name*
						</td>
						<td class=gray>
								&nbsp;MI*
						</td>		
						<td class=gray>
								&nbsp;Social Security No.*
						</td>
						<td class=gray>
								&nbsp;Sex*
						</td>			
					</tr>
					<tr>
						<td>
							<input type=text name="szLast_Name" value="<% = szLast_Name%>" maxlength=50 size=17  <% = strDisable %> ID="Text1">							
						</td>
						<td>
							<input type=text name="szFirst_Name" value="<% = szFirst_Name%>" maxlength=50 size=15  <% = strDisable %> ID="Text2">
						</td>
						<td>
							<input type=text name="sMid_Initial" value="<% = sMid_Initial%>" maxlength=1 size=2  <% = strDisable %> ID="Text3">
						</td>						
						<td>
							<input type=text name="szSSN" value="<% = szSSN%>" maxlength=11 size=20 ID="Text4">
						</td>
						<td>
							<select name="sSex"   ID="Select2">
								<% = oFunc.MakeList("M,F","Male,Female",sSex) %>
							</select>
						</td>
					</tr>
				</table>
				
				<table ID="Table3">
					<tr>					
						<td class=gray colspan=3 >
								&nbsp;Date of Birth*&nbsp;
						</td>			
						<td class=gray>
								&nbsp;Grade*&nbsp;
						</td>				
						<td class=gray>
								&nbsp;Grad Yr*&nbsp;
						</td>						
					</tr>
					<tr>						
						<td>
							<select name="mMonth"   ID="Select4">
								<option value="">
								<% 
								dim sqlMonth
								sqlMonth = "select strValue,strText from common_lists where intList_ID = 1 order by intOrder"
								Response.Write oFunc.MakeListSQL(sqlMonth,"","",mMonth)								
								%>
							</select>
						</td>		
						<td>
							<select name="mDay"   ID="Select5">
								<option value="">
								<% 
								dim sqlDay
								sqlDay = "select strValue,strText from common_lists where intList_ID = 2 order by intOrder"
								Response.Write oFunc.MakeListSQL(sqlDay,"","",mDay)								
								%>
							</select>
						</td>											
						<td>
							<select name="mYear"   ID="Select6">	
								<option value="">
								<% = oFunc.MakeYearList(0,20,mYear) %>
							</select>
						</td>	
						<td align=center>
							<select name="szGrade"   ID="Select7">
								<option value="">
								<% 
								dim strGrades
								strGrades = "K,1,2,3,4,5,6,7,8,9,10,11,12"							
								Response.Write oFunc.MakeList(strGrades,strGrades,szGrade)								
								%>
							</select>
						</td>
						<td align=center>
							<select name="intGrad_Year"   ID="Select8">
								<option value="">
								<% 
									Response.Write oFunc.MakeYearList(13,0,intGrad_Year) 							
								%>
							</select>
						</td>														
					</tr>
				</table>				
				
				<table ID="Table4">
					<tr>	
						<Td class=gray>
								&nbsp;First Language Student Learned
						</td>		
						<td class=gray>
								&nbsp;Language Spoken at Home&nbsp;
						</td>									
					</tr>
					<tr>
						<td>
							<select name="intFirst_Lang"   ID="Select9">
								<option value="">- - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
							<%							
								dim sqlLanguage
								sqlLanguage = "select intLanguage_id,szLanguage_Desc from trefLanguage order by szLanguage_Desc"
								Response.Write oFunc.MakeListSQL(sqlLanguage,"","",intFirst_Lang)
							%>
							</select>
						</td>		
						<td>
							<select name="intHome_Lang"   ID="Select10">
								<option value="">- - - - - - - - - - - - - - - - - - - - - - - -
							<%							
								Response.Write oFunc.MakeListSQL(sqlLanguage,"","",intHome_Lang)
							%>
							</select>
						</td>											
					</tr>
				</table>
				<br>	
				<table ID="Table5">
				<tr>	
					<Td colspan=6>
						<font class=svplain11>
							<b><i>Previous School: Out of District</I></B>
						</font>
					</td>
				</tr>
				<tr>
					<td class=gray>
							&nbsp;School Name
					</td>
					<td class=gray>
							&nbsp;Year
					</td>					
				</tr>
				<tr>
					<td>
						<input type=text name="szPrevious_School" value="<%=szPrevious_School%>" maxlength=256 size=30   ID="Text5">
					</td>
					<td>
						<input type=text name="szPrev_School_Year" value="<%=szPrev_School_Year%>" maxlength=4 size=5   ID="Text6">
					</td>			
				</tr>
			</table>
			
			<table ID="Table6">
				<tr>
					<td class=gray>
							&nbsp;Address
					</td>
					<td class=gray>
							&nbsp;City
					</td>
					<td class=gray>
							&nbsp;State
					</td>
					<Td class=gray>
							&nbsp;Country
					</td>				
					<Td class=gray>
							&nbsp;Zip
					</td>									
				</tr>
				<tr>
					<td>
						<input type=text name="szPrev_School_Addr" value="<% = szPrev_School_Addr%>" maxlength=256 size=30   ID="Text7">
					</td>
					<td>
						<input type=text name="szPrev_School_City" value="<%=szPrev_School_City%>" maxlength=50 size=10   ID="Text8">
					</td>
					<td>
						<select name="szPrev_School_State"   ID="Select11">
						<%
							dim sqlState
							sqlState = "select strValue,strText from Common_Lists where intList_Id = 3 order by strValue"
							Response.Write oFunc.MakeListSQL(sqlState,"","",szPrev_School_State)
						%>
						</select>						
					</td>
					<td>
						<input type=text name="szPrev_School_Country" value="<%=szPrev_School_Country%>" maxlength=25 size=7   ID="Text9">
					</td>
					<td>
						<input type=text name="szPrev_School_Zip_Code" value="<%=szPrev_School_Zip_Code%>" maxlength=11 size=5   ID="Text10">
					</td>		
				</tr>
			</table>	
			<BR>
			<table ID="Table7">
				<tr>	
					<Td colspan=6>
						<font class=svplain11>
							<b><i>Previous Anchorage School</I></B>
						</font>
					</td>
				</tr>
				<tr>
					<td class=gray>
							&nbsp;School Name
					</td>
					<td class=gray>
							&nbsp;Year
					</td>					
				</tr>
				<tr>
					<td>
						<input type=text name="szPrev_Anch_School" value="<%=szPrev_Anch_School%>" maxlength=256 size=30   ID="Text11">
					</td>
					<td>
						<input type=text name="intPrev_Anch_Year" value="<%=intPrev_Anch_Year%>" maxlength=4 size=5   ID="Text12">
					</td>			
				</tr>
			</table>	
			<BR>
			<table ID="Table8">
				<tr>	
					<Td colspan=2>
						<font class=svplain11>
							<b><i>Special Education Students</I></B> 
						</font>
					</td>
				</tr>
				<tr>	
					<Td colspan=2 class=gray>					
							Has your student been certified with a type of
							exceptionality through the ASD special education department?<BR>
							
							<b>Yes</b><input type=radio name="bolIEP"  value="1"   <% if  bolIEP = true or bolIEP = "1" then Response.Write " checked " %> ID="Radio1">
							<b>No</b><input type=radio name="bolIEP"  value="0"   <% if  bolIEP = false or bolIEP = "0" then Response.Write " checked " %> ID="Radio2">
							<b>Not Sure</b><input type=radio name="bolIEP"  value="NULL" <% if bolIEP & "" = "" or bolIEP = "NULL" then Response.Write " checked " %> ID="Radio3">
							
					</td>
				</tr>
				<tr>
					<td class=gray colspan=2>
							&nbsp;(if yes) When does your student’s IEP need to be renewed?
					</td>										
				</tr>
				<tr>
					<td colspan=2>
						<table ID="Table9">
							<tr>
								<td>
									<select name="IEPmonth"   ID="Select12">
										<option value="">
										<% 
										sqlMonth = "select strValue,strText from common_lists where intList_ID = 1 order by intOrder"
										Response.Write oFunc.MakeListSQL(sqlMonth,"","",iepMonth)								
										%>
									</select>
								</td>		
								<td>
									<select name="IEPday"   ID="Select13">
										<option value="">
										<% 
										sqlDay = "select strValue,strText from common_lists where intList_ID = 2 order by intOrder"
										Response.Write oFunc.MakeListSQL(sqlDay,"","",iepDay)								
										%>
									</select>
								</td>											
								<td>
									<select name="IEPyear"   ID="Select14">	
										<option value="">
										<% = oFunc.MakeYearList(4,2,iepYear) %>
									</select>
								</td>		
							</tr>
						</table>	
					</TD>					
				</tr>
				<TR>
					<td class=gray>
						&nbsp;(if yes) What is your student’s type of exceptionality?
					</td>
					<td>
						<input type=text size=30 maxlength=128 name="szExceptionality" value="<% = szExceptionality%>" ID="Text13">
					</td>
				</tr>
			</table>	
			<BR>
			<table ID="Table10">
				<tr>	
					<Td colspan=6>
						<font class=svplain11>
							<b><i>Emergency Information</I></B> 
						</font>
					</td>
				</tr>
				<tr>	
					<Td colspan=6>
						<font class=svplain10>
							Contact other than Guardians Entered
						</font>
					</td>
				</tr>
				<tr>
					<td class=gray>
							&nbsp;Last Name
					</td>
					<td class=gray>
							&nbsp;First Name
					</td>
					<td class=gray>
							&nbsp;Phone
					</td>										
				</tr>
				<tr>
					<td>
						<input type=text name="szContact_Last_Name" value="<% = szContact_Last_Name %>" maxlength=50 size=17   ID="Text14">
					</td>
					<td>
						<input type=text name="szContact_First_Name" value="<% = szContact_First_Name %>" maxlength=50 size=15   ID="Text15">
					</td>
					<td>
						<input type=text name="szContact_Phone" value="<% = szContact_Phone %>" maxlength=20 size=15   ID="Text16">
					</td>									
				</tr>
			</table>	
			<table ID="Table11">
				<tr>	
					<Td colspan=6>
						<font class=svplain10>
							Information of Students Doctor
						</font>
					</td>
				</tr>
				<tr>
					<td class=gray>
							&nbsp;Last Name
					</td>
					<td class=gray>
							&nbsp;First Name
					</td>
					<td class=gray>
							&nbsp;Phone
					</td>										
				</tr>
				<tr>
					<td>
						<input type=text name="szDR_Last_Name" value="<% = szDR_Last_Name %>" maxlength=50 size=17   ID="Text17">
					</td>
					<td>
						<input type=text name="szDR_First_Name" value="<% = szDR_First_Name %>" maxlength=50 size=15   ID="Text18">
					</td>
					<td>
						<input type=text name="szDR_Phone" value="<% = szDR_Phone %>" maxlength=20 size=15   ID="Text19">
					</td>									
				</tr>
			</table>	
			<table ID="Table12">
				<tr>
					<td class=gray>
							&nbsp;Daycare
					</td>
					<td class=gray>
							&nbsp;Phone
					</td>										
				</tr>
				<tr>
					<td>
						<input type=text name="szDaycare_Name" value="<% = szDaycare_Name%>" maxlength=128 size=25   ID="Text20">
					</td>
					<td>
						<input type=text name="szDaycare_Phone" value="<% = szDaycare_Phone %>" maxlength=20 size=15   ID="Text21">
					</td>									
				</tr>
			</table>		
			<table ID="Table13">
				<tr>
					<td class=gray colspan=2>
							&nbsp;Medic Alert Information
					</td>						
				</tr>
				<tr>
					<td>
						<input type=text name="szMed_Alert_1" value="<% = szMed_Alert_1 %>" maxlength=128 size=50 ID="Text22">
					</td>												
				</tr>
			</table>		
			<table ID="Table14">
				<tr>
					<td class=gray colspan=2>
							&nbsp;Disabilities
					</td>						
				</tr>
				<tr>
					<td>
						<input type=text name="szDisability_1" value="<% = szDisability_1 %>" maxlength=128 size=50 ID="Text24">
					</td>							
				</tr>
			</table>							
			</td>
		</tr>	
	</table>
		<input type=button value="Cancel" id="Button1" onClick="window.location.href='lotteryMain.asp';" NAME="Button1">	
		<input type=submit value="SAVE" id="Submit2" NAME="btSaveStudentInfo">	
	</form>
</body>
</html>
<%
function vbfValidateStudent()
	'Validates Student form information
	
	dim strError	' Used to store any returned errors
		
	' Now do the validation
	oVal.validateField szLast_Name,"blank","","Last Name" 
	oVal.validateField szFirst_Name,"blank","","First Name"
	oVal.validateField szSSN,"regexp","ssn","szSSN"
	oVal.validateField mMonth,"blank","","Birth 'Month'"
	oVal.validateField mDay,"blank","","Birth 'Day'"
	oVal.validateField mYear,"blank","","Birth 'Year'"
	oVal.validateField szGrade,"blank","","szGrade"
	oVal.validateField intGrad_Year,"blank","","intGrad_Year"	
	oVal.validateField intFirst_Lang,"blank","","First Language Student Learned"	
	oVal.validateField intHome_Lang,"blank","","Language Spoken at Home"	
		
	if bolIEP = "1" then
		oVal.validateField IEPmonth,"blank","","IEP Renewal 'Month'"
		oVal.validateField IEPday,"blank","","IEP Renewal 'Day'"
		oVal.validateField IEPyear,"blank","","IEP Renewal 'Year'"
		oVal.validateField szExceptionality,"blank","","Student’s type of exceptionality"
	end if 
	
	if szPrev_School_Year <> "" then
		if isNumeric(szPrev_School_Year) = false then
			strMoreError = "'Previous School Year' is an invalid number.<BR>"
		end if 
	end if 			
	
	if intPrev_Anch_Year <> "" then
		if isNumeric(intPrev_Anch_Year) = false then
			strMoreError = strMoreError & "'Previous Anchorage School Year' is an invalid number.<BR>"
		end if 
	end if 
	
	if oVal.ValidationError & "" <> ""  then
		strError = "<BR><font color=red><b>The following items need to be corrected before this information can be saved.</B><BR>"
		strError = strError & oVal.ValidationError & strMoreError & "</font>"	
	elseif strMoreError <> "" then
		strError = "<BR><font color=red><b>The following items need to be corrected before this information can be saved.</B><BR>"
		strError = strError & strMoreError & "</font>"
	end if
	
	vbfValidateStudent = strError	
end function

sub vbsInsertStudent()
	dim insert		
	oFunc.BeginTransCN			
			
	insert = "insert into tblStudent(intFamily_ID,szLast_Name,szFirst_Name,sMid_Initial," & _
			 "szSSN, sSex, dtBirth," & _
			 "intFirst_Lang, intHome_Lang,szGrade, intGrad_Year, " & _
			 "szPrevious_School,szPrev_School_Year,szPrev_School_Addr," & _
		 	 "szPrev_School_City,szPrev_School_State,szPrev_School_Country," & _
			 "szPrev_School_Zip_Code,szPrev_Anch_School,intPrev_Anch_Year," & _
			 "szContact_Last_Name,szContact_First_Name,szContact_Phone," & _
			 "szDR_Last_Name,szDR_First_Name,szDR_Phone,szDaycare_Name," & _
			 "szDaycare_Phone,szMed_Alert_1,szMed_Alert_2,szDisability_1," & _
			 "szDisability_2,bolIEP,dtIEP_Renewal,szExceptionality,bolLottery) VALUES (" & _
			 session.Contents("intFamily_ID") & ",'" & _
			 oFunc.EscapeTick(Request.Form("szLast_Name")) & "','" & _
			 oFunc.EscapeTick(Request.Form("szFirst_Name")) & "','" & _ 
			 oFunc.EscapeTick(Request.Form("sMid_Initial")) & "','" & _
			 oFunc.EscapeTick(Request.Form("szSSN")) & "','" & _
			 Request.Form("sSex") & "','" &_
			 dtBirth & "'," & _
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
			 oFunc.EscapeTick(request("szDisability_2")) & "'," & _ 
			 bolIEP & "," & _
			 dtIEP & _
			 "'" & oFunc.EscapeTick(request("szExceptionality")) & "',1)"
			 
			' response.Write insert
			' response.End
	oFunc.ExecuteCN(insert)
	
	intStudent_ID = oFunc.GetIdentity
	
	' Create lottery record
	insert = "insert into tblLottery_Appl (" & _
			 "intStudent_ID, dtLottery, dtCREATE) " & _
			 "values (" & _
			 intStudent_ID & "," & _
			 "'" & now() & "'," & _
			 "'" & now() & "')"
	oFunc.ExecuteCN(insert)
	
	oFunc.CommitTransCN		 
	session.Contents("strInstructions") = "Student has been added. To edit this students " & _
										  "information find the student under the 'Student Requesting Enrollment' " & _
										  "section and click the 'View/Edit Student Info.' button. If you are finished " & _
										  "adding guardian and student information please click the 'Finished' button " & _
										  "to finish the enrollment."
										  
	response.Redirect(Application.Value("strWebRoot") & "/forms/lottery/lotteryMain.asp")
end sub

sub vbsUpdateStudent
	oFunc.BeginTransCN
	dim update		
	
	update = "update tblStudent set " & _
			 "szLast_Name = '" & oFunc.EscapeTick(Request.Form("szLast_Name")) & "'," & _
			 "szFirst_Name = '" & oFunc.EscapeTick(Request.Form("szFirst_Name")) & "'," & _ 
			 "sMid_Initial = '" & oFunc.EscapeTick(Request.Form("sMid_Initial")) & "'," & _
			 "szSSN = '" & oFunc.EscapeTick(Request.Form("szSSN")) & "'," & _
			 "sSex = '" & Request.Form("sSex") & "'," &_
			 "dtBirth = '" & dtBirth & "'," & _
			 "intFirst_Lang = " & oFunc.CheckDecimal(Request.Form("intFirst_Lang")) & "," & _
			 "intHome_Lang = " & oFunc.CheckDecimal(Request.Form("intHome_Lang")) & "," & _
			 "szGrade = '" & oFunc.EscapeTick(Request.Form("szGrade")) & "'," & _
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
			 strUpdate & _
			 "bolIEP = " & bolIEP & "," & _
			 "szExceptionality = '" & oFunc.EscapeTick(request("szExceptionality")) & "' " & _
			 "where intStudent_ID = " & Request.Form("intStudent_ID") & _
			 " AND intFamily_ID = " & session.Contents("intFamily_ID")
			 
	oFunc.ExecuteCN(update)
	
	oFunc.CommitTransCN
	session.Contents("strInstructions") = "Student information has been saved. To edit this students " & _
										  "information find the student under the 'Student Requesting Enrollment' " & _
										  "section and click the 'View/Edit Student Info.' button."
										  
	response.Redirect(Application.Value("strWebRoot") & "/forms/lottery/lotteryMain.asp")		
end sub

sub vbsConfirmDelete 
%>
<html>
<head>
<title>Confirm Student Deletion</title>
<link rel="stylesheet" type="text/css" href="../../css/homestyle.css">
</head>
		
<body bgcolor=white>
<table width=100% height=100%>
	<tr>
		<td class=svplain11 valign=middle align=center>
			<b>Are you sure you want to delete student 
			'<% = request.QueryString("studentName")%>'?</b><br><br>
			<input type=button value="Cancel" onClick="window.location.href='lotteryMain.asp';" NAME="Button1">	
			<input type=button value="Yes, Delete Student." onclick="window.location.href='lotteryStudent.asp?intStudent_ID=<%=intStudent_ID%>&bolConfirm=true';">						
		</td>
	</tr>
</table>

</body>
</html>
<%
end sub

sub vbsDeleteStudent()
	' Deletes student and lottery entry. Lottery record is deleted by
	' cascading delete.
	
	'intFamily_ID is added to where clause as a precaution against
	'malicous header info being submitted by a hacker
	
	delete = "delete from tblStudent " & _
			 "where intStudent_ID = " & intStudent_ID & _
			 "and intFamily_ID = " & session.Contents("intFamily_ID") 
	oFunc.ExecuteCN(delete)
	
	session.Contents("strInstructions") = "Student information has been deleted. "
										  
										  
	response.Redirect(Application.Value("strWebRoot") & "/forms/lottery/lotteryMain.asp")
	
end sub

set oFunc = nothing
set oVal = nothing
%>