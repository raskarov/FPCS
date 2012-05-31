<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		StudentProfile.asp  
'Purpose:	This script collects the student information
'			or displays the student information.
'Date:		9 July 2001
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim oFunc
dim bolExistingStudent
dim intYearsLeft
dim bolHeader

Session.Value("strTitle") = "SIS Student Profile"
Session.Value("simpleTitle") = "SIS Student Profile"
Session.Value("strLastUpdate") = "09 June 2002"

if Request.QueryString("bolNewStudent") <> "" or Request.QueryString("bolUpdate") <> "" then
	Server.Execute(Application.Value("strWebRoot") & "Includes/simpleHeader.asp")
	bolHeader = false
else
	Server.Execute(Application.Value("strWebRoot") & "Includes/Header.asp")
	bolHeader = true
end if

   set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
   call oFunc.OpenCN()

	intStudent_id = Request.QueryString("intStudent_id")
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'' When the next if is true we get the student info 
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	if intStudent_id <> "" then
		dim rsStudent
		dim sqlStudent
		dim intCount
		dim item
		dim mMonth
		dim mDay
		dim mYear
		dim intCount2
		dim lotteryMonth
		dim lotteryDay
		dim lotteryYear
		dim lotteryRecvdMonth
		dim lotteryRecvdDay
		dim lotteryRecvdYear		
		
			
		set rsStudent = Server.CreateObject("ADODB.RECORDSET")
		rsStudent.CursorLocation = 3
				
		sqlStudent = "SELECT S.szFIRST_NAME, S.szLAST_NAME, S.sMID_INITIAL, S.szSSN, S.sSEX, S.intRACE_ID,  " & _
					"S.intTUITION_ID, S.dtBIRTH, SS.szGRADE, S.intGRAD_YEAR,  " & _
					" S.intFIRST_LANG, S.intHOME_LANG, S.szPrevious_School, S.szPrev_School_Year,  " & _
					"S.szPrev_School_Addr, S.szPrev_School_City,  " & _
					"S.szPrev_School_State, S.szPrev_School_Country,  " & _
					"S.szPrev_School_Zip_Code, S.szPrev_Anch_School, S.intPrev_Anch_Year,  " & _
					" S.szContact_Last_Name, S.szContact_First_Name,  " & _
					"S.szContact_Phone, S.szDR_Last_Name, S.szDR_First_Name, S.szDR_Phone,  " & _
					"S.szDaycare_Name, S.szDaycare_Phone, S.szMed_Alert_1,  " & _
					"S.szMed_Alert_2, S.szDisability_1, S.szDisability_2, EI.intENROLL_INFO_ID,  " & _
					"EI.szPrivate_School_Name, EI.szOther_District_Name,  " & _
					"EI.intPercent_Enrolled_D2, EI.intPercent_Enrolled_Fpcs, EI.bolCharter_Grad,  " & _
					"SE.bolASD_Contract_HRS_Exempt, SE.szHRS_Exempt_Reason, " & _
					" SE.intCore_Credit_Percent, SE.szCore_Exemption_Reason,  " & _
					"SE.intElective_Credit_Percent, SE.szElective_Exemption_Reason, " & _
					"SE.intStudent_Exemption_ID, SS.intReEnroll_State, SS.intStudent_State_ID,SS.dtWithdrawn, " & _
					"S.dtLottery, S.dtLottery_Received, S.szNew_Wait_List_Num, EI.intPercent_Enrolled_Locked " & _
					"FROM         tblSTUDENT S LEFT OUTER JOIN " & _
					"tblStudent_States SS ON S.intSTUDENT_ID = SS.intStudent_id  " & _
					"AND SS.intSchool_Year = " & session.Contents("intSchool_Year") & " LEFT OUTER JOIN " & _
					" tblStudent_Exemptions SE ON S.intSTUDENT_ID = SE.intStudent_ID  " & _
					"AND SE.intSchool_Year = " & session.Contents("intSchool_Year") & " LEFT OUTER JOIN " & _
					"tblENROLL_INFO EI ON S.intSTUDENT_ID = EI.intSTUDENT_ID  " & _
					"AND EI.sintSCHOOL_YEAR = " & session.Contents("intSchool_Year") & " " & _
					"WHERE     (S.intSTUDENT_ID = " & intStudent_id & ")"

					
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
			szSSN = oFunc.Reformat(szSSN, Array("", 3, "-", 2, "-", 4))
			szPrev_School_Zip_Code = oFunc.Reformat(szPrev_School_Zip_Code, Array("", 5, "-", 4))
			'szDaycare_Phone = oFunc.Reformat(szDaycare_Phone, Array("(", 3, ") ", 3, "-", 4))
			'szDR_Phone = oFunc.Reformat(szDR_Phone, Array("(", 3, ") ", 3, "-", 4))
				
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'' The Birth data is stored in the database as a single field, but our form 
		'' displays it as three seperate select lists so we break the single
		'' date up to use the parts in our form populating.
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''		
			lotteryMonth = datePart("m",dtLottery)
			lotteryDay = datePart("d",dtLottery)
			lotteryYear = datePart("yyyy",dtLottery)
			
			lotteryRecvdMonth = datePart("m",dtLottery_Received)
			lotteryRecvdDay = datePart("d",dtLottery_Received)
			lotteryRecvdYear = datePart("yyyy",dtLottery_Received)
			' Removed 6-9-2003. We now use IEP.asp
			'iepMonth = datePart("m",dtIEP_Renewal)
			'iepDay = datePart("d",dtIEP_Renewal)
			'iepYear = datePart("yyyy",dtIEP_Renewal)
		else
			Response.Write "Student ID " & Session.Value("intStudent_id") & " is not a valid ID"
			Session.Value("intStudent_id") = ""		
			Session.Value("SISEditMode") = ""	
		end if
		rsStudent.Close
		set rsStudent = nothing	
	end if 
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'' Now we either print a blank student form or a populated one based
	'' on the logic above.
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	%>
	<script language="javascript">
	
		function jfSubmit(objForm) {
			if (jfValidate(objForm) == true) {         
				objForm.submit();
			}
		}
		
		function jfValidate(objForm) {
		//Ensure all approriate fields have been filled out
			var strErrMsg			= '';
			var strLast_Name		= objForm.szLast_Name.value;
			var strFirst_Name		= objForm.szFirst_Name.value;		
			var strSSN				= objForm.szSSN.value;				
			var intRace_id			= objForm.intRace_id.value;
			//var intMonth			= objForm.month.value;
			//var intDay				= objForm.day.value;
			//var intYear				= objForm.year.value;
			var strGrade			= objForm.szGrade.value;
			var intGrad_Year		= objForm.intGrad_Year.value;
			var intFirst_Lang		= objForm.intFirst_Lang.value;
			var intHome_Lang		= objForm.intHome_Lang.value;
			var intPercent_Enrolled_FPCS = objForm.intPercent_Enrolled_FPCS.value;
			var szContact_Last_Name = objForm.szContact_Last_Name.value;
			var szContact_First_Name = objForm.szContact_First_Name.value;
			var szContact_Phone		= objForm.szContact_Phone.value;
			var szDR_Last_Name		= objForm.szDR_Last_Name.value;
			var szDR_First_Name		= objForm.szDR_First_Name.value;
			var szDR_Phone			= objForm.szDR_Phone.value;
			var szPrev_School_Year  = objForm.szPrev_School_Year.value;
			var intPrev_Anch_Year   = objForm.szPrev_School_Year.value;
			
			<% if ucase(session.contents("strRole")) = "ADMIN" then %>
			var lotteryRecvdMonth   = objForm.lotteryRecvdMonth.value;
			var lotteryRecvdDay     = objForm.lotteryRecvdDay.value;
			var lotteryRecvdYear    = objForm.lotteryRecvdYear.value;
			var lotteryMonth		= objForm.lotteryMonth.value;
			var lotteryDay			= objForm.lotteryDay.value;
			var lotteryYear			= objForm.lotteryYear.value;
			var intReEnroll_State  = objForm.intReEnroll_State.value;
			<% end if %>  
			  
			if(strLast_Name.length == 0) {strErrMsg += 'Last Name\n';}
			if(strFirst_Name.length == 0) {strErrMsg += 'First Name\n';}
			if(strSSN.length == 0) {strErrMsg += 'Social Security Number\n';}
			if(intRace_id.length == 0) {strErrMsg += 'Ethnic Origin\n';}
			if(objForm.dtBirth.value == "") {strErrMsg += 'Birth Date\n';}
			if(strGrade.length == 0) {strErrMsg += 'Grade\n';}
			if(intGrad_Year.length == 0) {strErrMsg += 'Graduating Year\n';}
			if(intFirst_Lang.length == 0) {strErrMsg += 'First Language\n';}
			if(intHome_Lang.length == 0) {strErrMsg += 'Home Language\n';}
			if(intPercent_Enrolled_FPCS.length == 0) {strErrMsg += 'Percent Planning to Enroll in FPCS\n';}
			if(isNaN(szPrev_School_Year)) {strErrMsg += 'Previous School Out of District Year Must be a Vaild Number.\n';}
			if(isNaN(intPrev_Anch_Year)) {strErrMsg += 'Previous Anchorage School Year Must be a Vaild Number.\n';}
			if(szContact_Last_Name.length == 0) {strErrMsg += 'Emergency Contact Last Name\n';}
			if(szContact_First_Name.length == 0) {strErrMsg += 'Emergency Contact First Name\n';}
			if(szContact_Phone.length == 0) {strErrMsg += 'Emergency Contact Phone\n';}
			if(szDR_Last_Name.length == 0) {strErrMsg += 'Doctors Last Name\n';}
			if(szDR_First_Name.length == 0) {strErrMsg += 'Doctors First Name\n';}
			if(szDR_Phone.length == 0) {strErrMsg += 'Doctors Phone\n';}
			
			<% if ucase(session.contents("strRole")) = "ADMIN" then %>  
				if (lotteryMonth != "" || lotteryDay != "" || lotteryYear != ""){
					if (lotteryMonth == "" || lotteryDay == "" || lotteryYear == "") {
						strErrMsg += 'Invalid Lottery Date\n';
					}
				}
				
				if (lotteryRecvdMonth != "" || lotteryRecvdDay != "" || lotteryRecvdYear != ""){
					if (lotteryRecvdMonth == "" || lotteryRecvdDay == "" || lotteryRecvdYear == "") {
						strErrMsg += 'Invalid Lottery Received Date\n';
					}
				}
				if (intReEnroll_State == "") {
					strErrMsg += 'Enrollment State can not be Blank\n';
				}
			<% end if %>
			   
			if (strErrMsg.length == 0 ) {
				if(strSSN.length != 0) {
					if (checkSSN(objForm.szSSN) == false) {return false;}
				}
				//if (checkDate(objForm.year, objForm.month, objForm.day, "Birth Date") == false) {return false;}
				objForm.szPrev_School_Zip_Code.value = stripCharsInBag(objForm.szPrev_School_Zip_Code.value, ZIPCodeDelimiters);
				if(strSSN.length != 0) {
					objForm.szSSN.value = stripCharsInBag(objForm.szSSN.value, SSNDelimiters);
				}
				 
				objForm.szDR_Phone.value = stripCharsInBag(objForm.szDR_Phone.value, phoneNumberDelimiters);
				return true;
			} else {    
				strErrMsg = 'The Following Information is Required:\n \n' + strErrMsg;
				alert(strErrMsg);
				return false;
			}
		}
		
		function jfConfirm(){
			var bolContinue = confirm("Are you sure you want to close without saving any changes you may have made? Click 'OK' to close this window without saving. Otherwise click 'Cancel' and then click 'Save Changes and Close'.");
			if (bolContinue == false) {
				return false;
			}
			window.opener.focus();
			window.close();
		}
	</script>
	<form action="studentInsert.asp" method=POST name=main onsubmit="return false;">
	<script language=javascript src="<%= Application.Value("strWebRoot") %>includes/CalendarPopup.js"></script>	
	<script language="javascript">
		function jfChangeStudent(item){
			//reloads page with newly selected student
			var strURL = "<% = Application.Value("strWebRoot")%>forms/SIS/studentProfile.asp?intStudent_ID=" + item.value;
			window.open(strURL, "_self");
		}
		var cal = new CalendarPopup('divCal');
		cal.showNavigationDropdowns();
		cal.setYearSelectStartOffset(10);
		
		var cal2 = new CalendarPopup('divCal');
		cal2.showNavigationDropdowns();
		cal2.setYearSelectStartOffset(70);
		cal2.offsetX = 40;

		
	</script>
	<% if Request("strMessage") <> "" then %>
	&nbsp;<font class=svPlain11 color=red><b><% = Request("strMessage") %></b></font><br>	
	<% end if %>
	<input type=hidden name=changed value="">
	<input type=hidden name="bolNewStudent" value = "<%=Request.QueryString("bolNewStudent")%>">
	<input type=hidden name="intStudent_State_ID" value="<% = intStudent_State_ID %>" ID="Hidden3">
	<input type="hidden" name="hdnNeedChanged" id="hdnNeedChanged" value="">
	<table width=100%>
		<tr>	
			<Td class=yellowHeader>
					&nbsp;<b>SIS Online Enrollment Form</b>&nbsp;&nbsp;&nbsp;
			<% if Request.QueryString("bolNewStudent") = "" and session.Value("strRole") <> "GUARD"then %>	
					<select name="intStudent_ID" onchange="jfChangeStudent(this);">
						<option value="">
						<option value="">New Student
					<%
						'this change was requested by Val
						dim sqlStudentName
						sqlStudentName = "SELECT intStudent_ID,szLast_Name + ',' + szFirst_Name AS Name " & _
										"FROM tblStudent ORDER BY szLast_Name"
						Response.Write oFunc.MakeListSQL(sqlStudentName,"intStudent_ID","Name",Request("intStudent_ID"))												 
					%>
					</select>
			<% else %>
				<input type=hidden name="intStudent_ID" value="<% = Request.QueryString("intStudent_ID")%>">
				<font size=1><b>* = required fields.</b></font>
			<% end if %>
			</td>
		</tr>
		<tr>
			<td bgcolor=f7f7f7>
				<table>
					<tr>
						<td>
							<table>
								<tr>	
									<Td colspan=6>
										<font class=svplain11>
											<b><i>Students Information</I></B>
										</font>
									</td>
								</tr>
								<tr>
									<td class=gray nowrap>
											&nbsp;Legal Name: Last*&nbsp;
									</td>
									<td class=gray nowrap>
											&nbsp;First Name*&nbsp;
									</td>
									<td class=gray nowrap>
											&nbsp;MI
									</td>		
									<td class=gray nowrap>
											&nbsp;Social Security No.*
									</td>
									<td class=gray nowrap style="width:100%;">
											&nbsp;Sex*
									</td>			
								</tr>
								<tr>
									<% if session.Value("strRole") <> "GUARD" then %>	
									<td>
										<input type=text name="szLast_Name" value="<% = szLast_Name%>" maxlength=50 size=17 onChange="jfUpper(this);" <% = strDisable %>>							
									</td>
									<td>
										<input type=text name="szFirst_Name" value="<% = szFirst_Name%>" maxlength=50 size=15 onChange="jfUpper(this);" <% = strDisable %>>
									</td>
									<td>
										<input type=text name="sMid_Initial" value="<% = sMid_Initial%>" maxlength=1 size=2 onChange="jfUpper(this);" <% = strDisable %>>
									</td>
									<% else %>
									<td class=gray nowrap>
										&nbsp;<% = szLast_Name%>
										<input type=hidden name="szLast_Name" value="<% = szLast_Name%>" >							
									</td>
									<td class=gray nowrap>
										&nbsp;<% = szFirst_Name%>
										<input type=hidden name="szFirst_Name" value="<% = szFirst_Name%>" >	
									</td>
									<td class=gray nowrap>
										&nbsp;<% = sMid_Initial%>
										<input type=hidden name="sMid_Initial" value="<% = sMid_Initial%>" >	
									</td>
									<% end if %>
									<td nowrap>
										<input type=text name="szSSN" value="<% = szSSN%>" maxlength=11 size=20 onChange="checkSSN(this);">
									</td>
									<td>
										<select name="sSex"   style="width:100%;">
											<% = oFunc.MakeList("M,F","Male,Female",sSex) %>
										</select>
									</td>
								</tr>
							</table>
							<table ID="Table4">
								<tr>	
									<Td class=gray>
											&nbsp;First Language Student Learned*
									</td>		
									<td class=gray>
											&nbsp;Language Spoken at Home*&nbsp; 
									</td>									
								</tr>
								<tr>
									<td>
										<select name="intFirst_Lang"   ID="Select11">
											<option value="">- - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
										<%							
											dim sqlLanguage
											sqlLanguage = "select intLanguage_id,szLanguage_Desc from trefLanguage order by szLanguage_Desc"
											Response.Write oFunc.MakeListSQL(sqlLanguage,"","",intFirst_Lang)
										%>
										</select>
									</td>		
									<td>
										<select name="intHome_Lang"   ID="Select12">
											<option value="">- - - - - - - - - - - - - - - - - - - - - - - -
										<%							
											Response.Write oFunc.MakeListSQL(sqlLanguage,"","",intHome_Lang)
										%>
										</select>
									</td>											
								</tr>
							</table>
							<table ID="Table3" style="width:100%;">
								<tr>	
									<Td class=gray style="width:100%;">
											&nbsp;Ethnic Origin*
									</td>	<!--	
									<td class=gray>
											&nbsp;Tuition Status&nbsp;
									</td> -->								
									<td class=gray nowrap>
											&nbsp;Grade*&nbsp;
									</td>				
									<td class=gray nowrap>
											&nbsp;Grad Yr*&nbsp;
									</td>		
									<td class=gray  nowrap>
											&nbsp;Date of Birth*&nbsp;
									</td>					
								</tr>
								<tr>
									<td nowrap>
										<select name="intRace_id"   ID="Select5" style="width:100%;">
											<option value="">
										<%							
											dim sqlRace
											sqlRace = "select intRace_id,szRace_Desc from trefRace"
											Response.Write oFunc.MakeListSQL(sqlRace,"","",intRace_id)
										%>
										</select>
									</td>						
									<!--
									<td>
										<select name="intTuition_ID"  >
											<% 
											dim sqlTuition
											if intTuition_ID = "" then intTuition_ID = "1"
											sqlTuition = "select intTuition_id,szTuition_desc from trefTuition order by intTuition_id"
											Response.Write oFunc.MakeListSQL(sqlTuition,"","",intTuition_ID)								
											%>
										</select>
									</td>		
									<td>
										<select name="month"   ID="Select6">
											<option value="">
											<% 
											dim sqlMonth
											sqlMonth = "select strValue,strText from common_lists where intList_ID = 1 order by intOrder"
											Response.Write oFunc.MakeListSQL(sqlMonth,"","",mMonth)								
											%>
										</select>
									</td>		
									<td>
										<select name="day"   ID="Select7">
											<option value="">
											<% 
											dim sqlDay
											sqlDay = "select strValue,strText from common_lists where intList_ID = 2 order by intOrder"
											Response.Write oFunc.MakeListSQL(sqlDay,"","",mDay)								
											%>
										</select>
									</td>											
									<td>
										<select name="year"   ID="Select8">	
											<option value="">
											<% = oFunc.MakeYearList(0,20,mYear) %>
										</select>
									</td>	-->
									<% 
										dim strGrades
											
											' These next couple if statements are used to auto set a grade and
											' graduation year for a student that is returning for another school
											' year.  We only auto calc this info for existing students since the
											' assumption is that a new students SIS form will contain current
											' info or be blank.
											
											' Do we have an re-enrolling student?		
																
											if intENROLL_INFO_ID&"" = "" and intStudent_id <> "" then 
												set rsCheck = server.CreateObject("ADODB.RECORDSET")
												rsCheck.CursorLocation = 3
												sql = "select * from tblEnroll_info where intStudent_id = " & intStudent_id & _
													" and sintSchool_year = " & session.Contents("intSchool_Year") -1
												rsCheck.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
												
												if rsCheck.RecordCount > 0 then
													bolExistingStudent = true
												end if
												rsCheck.Close
												set rsCheck = nothing
											end if
											
											' Ony auto calc for re-enrolling students
											'if false then
												if intENROLL_INFO_ID&"" = "" and ucase(szGrade) <> "K" then
													szGrade = szGrade + 1
													intYearsLeft = 12 - szGrade
													intGrad_Year = session.Contents("intSchool_Year") + intYearsLeft
												elseif intENROLL_INFO_ID&"" = "" and ucase(szGrade) = "K" then
													szGrade = 1		
													intGrad_Year = session.Contents("intSchool_Year") + 11							
												end if 
											'end if
									%>
									<td align=center class="svplain8" align="center" nowrap>
										<% if oFunc.IsAdmin then %>
										<select name="szGrade"   ID="Select9">
											<option value="">
											<% 								
											strGrades = "K,1,2,3,4,5,6,7,8,9,10,11,12"							
											Response.Write oFunc.MakeList(strGrades,strGrades,szGrade)								
											%>
										</select>
										<% else %>
											<% = szGrade %>
											<input type="hidden" name="szGrade" value="<% = szGrade %>">
										<% end if %>
									</td>
									<td align=center nowrap>
										<select name="intGrad_Year"   ID="Select10">
											<option value="">
											<% 
												Response.Write oFunc.MakeYearList(13,0,intGrad_Year) 							
											%>
										</select>
									</td>			
									<td class="svplain8" nowrap>
										<input type=text name="dtBirth" size=10 value="<% = dtBirth %>" maxlength=10  ID="Text14" > 
												<a href="#" onclick="cal2.select(document.forms[0].dtBirth,'adtBirth','M/d/yyyy','0<% = trim(dtBirth) %>');return false;" id="adtBirth" name="adtBirth">calendar</a>
									</td>											
								</tr>
							</table>
						</td>
					</tr>
				</table>												
				<br>	
					
				
<input type=hidden name="intEnroll_Info_ID" value="<%=intEnroll_Info_ID%>" ID="Hidden1">
<table width=100% ID="Table1">
	<tr>	
		<Td colspan=6>
			<font class=svplain11>
				<b><i>Enrollment Information for SY <% = oFunc.SchoolYearRange %></I></B>
			</font>
		</td>
	</tr>
	
	<tr>
		<Td>
			<table>
				<tr>
					<td class=gray >
						Enrollment Status:
					</td>
					<td class="svplain8" align="left">	
						<script language="javascript">
							function jfShowNeed(val){
								var obj = document.getElementById('EnrollNeedID');
								if (val == "129") {									
									obj.style.display = 'block';
								}else{
									obj.style.display = 'none';
								}
							
							}
						</script>
						<%
						sql = "SELECT intReEnroll_State, strCase " & _
								"FROM trefReEnroll_States " & _
								" WHERE bolActive = 1 " & _
								"ORDER BY strCase "
								
						sEnrollList = oFunc.MakeListSQL(sql,"","",intReEnroll_State)
						sEnrollValue = oFunc.SelectedListText
						
						if oFunc.IsAdmin then 
						%>		
						<input type="hidden" name="oldReEnrollState" value="<% = intReEnroll_State %>">					
						<select name="intReEnroll_State" onChange="jfShowNeed(this.value);">
						<option value=""></option>
						<% = sEnrollList %>
						</select>
						<%else
							Response.Write sEnrollValue
							end if 
						%>
					</td>
					<% if oFunc.IsAdmin then  %>
					<td class=gray nowrap>
						Withdrawl Date:
					</td>
					<td class="svplain8">									
						<input type=text name="dtWithdrawn" size=10 value="<% = dtWithdrawn %>" maxlength=10  ID="Text13" > 
						<a href="#" onclick="cal.select(document.forms[0].dtWithdrawn,'aWithdraw','M/d/yyyy','0<% = trim(dtWithdrawn) %>');return false;" id="aWithdraw" name="aWithdraw">calendar</a>
						
					</td>
					<% end if %>
				</tr>
				<tr id="EnrollNeedID" style="display:<% if intReEnroll_State = 129 then response.Write "block" else response.Write "none" %>;">
					<td colspan="100">
						<table>
							
								<%
									sql = "SELECT NeededEnrollInfoCD, Label " & _ 
											"FROM trefNEEDED_ENROLL_INFO " & _ 
											"WHERE (IsActive = 1) " & _ 
											"ORDER BY OrderID "
									
									dim rsNeed 
									set rsNeed = server.CreateObject("ADODB.RECORDSET")
									rsNeed.CursorLocation = 3
									rsNeed.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
									
									if  request("intStudent_ID") & "" <> "" then 
										sql = "select NeededEnrollInfoCD from STUDENT_ENROLL_INFO_NEEDED " & _
											" WHERE STudentID = " & request("intStudent_ID") & " and SchoolYear = " & session.Contents("intSchool_Year")
											'response.Write sql 									  
										set rsNeed2 = server.CreateObject("ADODB.RECORDSET")
										rsNeed2.CursorLocation = 3
										rsNeed2.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
									end if
									
									do while not rsNeed.EOF
										if  request("intStudent_ID") & "" <> "" then 
											if rsNeed2.RecordCount > 0 then
												isChecked = ""
												do while not rsNeed2.EOF and isChecked = ""
													if rsNeed("NeededEnrollInfoCD") & "" = rsNeed2("NeededEnrollInfoCD") & "" then
														isChecked = " checked "
													end if 
													rsNeed2.MoveNext
												loop
												rsNeed2.MoveFirst
											end if
										end if 
								%>
									<tr><td class='tablecell' nowrap><%= rsNeed("Label") %><input type=checkbox name='EnrollNeed' value='<% = rsNeed("NeededEnrollInfoCD")%>' onChange="document.getElementById('hdnNeedChanged').value = 'true';" <% = isChecked %>></td></tr>
								<%
										rsNeed.MoveNext
									loop
									
									rsNeed.Close
									set rsNeed = Nothing
									if  request("intStudent_ID") & "" <> "" then 
										rsNeed2.Close
										set rsNeed2 = Nothing
									end if
									
								%>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<% 
		if szFirst_Name = "" then szFirst_Name = "this student"
	%>		
				<tr>	
					<Td class=gray>
						&nbsp;If  <% = szFirst_Name  %> will be enrolled in a private school for this
						school year what is the name of that school?
					</td>
				</tr>	
				<tr>						
					<td>
						<input type=text name="szPrivate_School_Name" value="<%=szPrivate_School_Name%>" maxlength=128 size=40 ID="Text1">
					</td>
				</tr>	
				<tr>	
					<Td class=gray>
						&nbsp;If  <% = szFirst_Name %> will be enrolled in another school district 
						(such as Galena (IDEA), Alyeska Central School, Nenana (CyberLinks), etc.) what
						percentage will <% = szFirst_Name %> be enrolled in that school district?
					</td>
				</tr>	
				<tr>						
					<td>
						
						<select name="intPercent_Enrolled_D2" ID="Select1">
						<%
							Response.Write oFunc.MakeList("0,25,50,75,100","0%,25%,50%,75%,100%",intPercent_Enrolled_D2)
						%>
						</select>						
					</td>
				</tr>	 
				<tr>	
					<Td class=gray>
						&nbsp;What percentage are you planning to enroll <% = szFirst_Name %> in FPCS
						for this new year?*
					</td>	
				</tr>	
				<tr>					
					<td class="svplain8">
						<% if (ucase(session.Contents("strROLE")) = "GUARD" or ucase(session.Contents("strRole")) = "TEACHER") and _
								isNumeric(intPercent_Enrolled_Locked) then %>
								<% = intPercent_Enrolled_FPCS %>%
								<input type="hidden" name="intPercent_Enrolled_FPCS" value="<%=intPercent_Enrolled_FPCS%>" ID="Hidden4">
						<% else %>
							<select name="intPercent_Enrolled_FPCS" ID="Select2">
								<option>
							<%
								Response.Write oFunc.MakeList("25,50,75,100","25%,50%,75%,100%",intPercent_Enrolled_FPCS)
							%>
							</select>
						<% end if %>
					</td>
				</tr>		
				<tr>	
					<Td class=gray>
						&nbsp;Do you plan for <% = szFirst_Name %> to graduate from FPCS?
					</td>	
				</tr>	
				<tr>					
					<td>
						<select name="bolCharter_Grad" ID="Select3">
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
						<select name=bolASD_Contract_HRS_Exempt ID="Select4">
						<%
							response.Write oFunc.MakeList("0,1","No,Yes",oFunc.TrueFalse(bolASD_Contract_HRS_Exempt))
						%>						
						</select>	
						If yes please give reason: 
						<input type=text name=szHRS_Exempt_Reason value="<%=szHRS_Exempt_Reason%>" maxlength=511 size=25 ID="Text2">			
						<input type=hidden name="intStudent_Exemption_ID" value="<%=intStudent_Exemption_ID%>" ID="Hidden2">
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
						Core Credit Exemption%:<input type=text name="intCore_Credit_Percent" size=4 value="<%=intCore_Credit_Percent%>" ID="Text3">
						 Exemption Reason: <input type=text name="szCore_Exemption_Reason" value="<%=szCore_Exemption_Reason%>" size=15 maxlength=511 ID="Text4">						 
					</td>
				</tr>
				<tr>					
					<td class=svplain10>
						Elective Credit Exempt%:<input type=text name="intElective_Credit_Percent" value="<%=intElective_Credit_Percent%>" size=4 ID="Text5">
						 Exemption Reason: <input type=text name="szElective_Exemption_Reason" value="<%= szElective_Exemption_Reason %>" size=15 ID="Text6" maxlength=511>						 
					</td>
				</tr>
			  <% end if %>		
			</table>
		</td>
	</tr>
</table>	

<br>	
<% if ucase(session.Contents("strRole")) = "ADMIN" then %>			
<table ID="Table7">
	<tr>	
		<Td colspan=6>
			<font class=svplain11>
				<b><i>Lottery Information</I></B>
			</font>
		</td>
	</tr>
	<tr>
		<td class=gray>
				&nbsp;Lottery Date
		</td>
		<td class=gray>
				&nbsp;Application Date
		</td>	
		<td class=gray>
				&nbsp;Wait List Number
		</td>				
	</tr>
	<tr>
		<td>
			<table cellpadding=0 cellspacing=0>
				<tr>
					<td>
						<select name="lotteryMonth"   ID="Select14">
							<option value="">
							<% 
							sqlMonth = "select strValue,strText from common_lists where intList_ID = 1 order by intOrder"
							Response.Write oFunc.MakeListSQL(sqlMonth,"","",lotteryMonth)								
							%>
						</select>
					</td>		
					<td>
						<select name="lotteryDay"   ID="Select15">
							<option value="">
							<% 
							sqlDay = "select strValue,strText from common_lists where intList_ID = 2 order by intOrder"
							Response.Write oFunc.MakeListSQL(sqlDay,"","",lotteryDay)								
							%>
						</select>
					</td>											
					<td>
						<select name="lotteryYear"   ID="Select16">	
							<option value="">
							<% = oFunc.MakeYearList((year(now) - cint(application.Value("dtYearAppStarted"))+2),1,lotteryYear) %>
						</select>
					</td>
				</tr>
			</table>							
		</td>
		<td>
			<table cellpadding=0 cellspacing=0 ID="Table8">
					<tr>
						<td>
							<select name="lotteryRecvdMonth"   ID="Select17">
								<option value="">
									<% 
									sqlMonth = "select strValue,strText from common_lists where intList_ID = 1 order by intOrder"
									Response.Write oFunc.MakeListSQL(sqlMonth,"","",lotteryRecvdMonth)								
									%>
								</select>
						</td>		
						<td>
							<select name="lotteryRecvdDay"   ID="Select18">
								<option value="">
								<% 
								sqlDay = "select strValue,strText from common_lists where intList_ID = 2 order by intOrder"
								Response.Write oFunc.MakeListSQL(sqlDay,"","",lotteryRecvdDay)								
								%>
							</select>
						</td>											
						<td>
							<select name="lotteryRecvdYear"   ID="Select19">	
								<option value="">
								<% = oFunc.MakeYearList((year(now) - cint(application.Value("dtYearAppStarted"))+2),1,lotteryRecvdYear) %>
							</select>
						</td>
					</tr>
				</table>
		</td>	
		<td>
			<input type=text name="szNew_Wait_List_Num" value="<%=szNew_Wait_List_Num%>" maxlength=12 size=25    ID="Text15">
		</td>		
	</tr>
</table>			
				
<br>
<% end if %>				
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
						<input type=text name="szPrevious_School" value="<%=szPrevious_School%>" maxlength=256 size=30   ID="Text7">
					</td>
					<td>
						<input type=text name="szPrev_School_Year" value="<%=szPrev_School_Year%>" maxlength=4 size=5   ID="Text8">
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
						<input type=text name="szPrev_School_Addr" value="<% = szPrev_School_Addr%>" maxlength=256 size=30   ID="Text9">
					</td>
					<td>
						<input type=text name="szPrev_School_City" value="<%=szPrev_School_City%>" maxlength=50 size=10   ID="Text10">
					</td>
					<td>
						<select name="szPrev_School_State"   ID="Select13">
						<%
							dim sqlState
							sqlState = "select strValue,strText from Common_Lists where intList_Id = 3 order by strValue"
							Response.Write oFunc.MakeListSQL(sqlState,"","",szPrev_School_State)
						%>
						</select>						
					</td>
					<td>
						<input type=text name="szPrev_School_Country" value="<%=szPrev_School_Country%>" maxlength=25 size=7   ID="Text11">
					</td>
					<td>
						<input type=text name="szPrev_School_Zip_Code" value="<%=szPrev_School_Zip_Code%>" maxlength=11 size=5   ID="Text12">
					</td>		
				</tr>
			</table>			
			<BR>
			<table>
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
						<input type=text name="szPrev_Anch_School" value="<%=szPrev_Anch_School%>" maxlength=256 size=30  >
					</td>
					<td>
						<input type=text name="intPrev_Anch_Year" value="<%=intPrev_Anch_Year%>" maxlength=4 size=5  >
					</td>			
				</tr>
			</table>	
			<BR>
			<!--<table>
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
							<% if intStudent_id = "" then bolIEP = " " %>
							<b>Yes</b><input type=radio name="bolIEP"  value="yes"   <% if  bolIEP = true then Response.Write " checked " %> >
							<b>No</b><input type=radio name="bolIEP"  value="no"   <% if  bolIEP = false then Response.Write " checked " %> >
							
					</td>
				</tr>
				<tr>
					<td class=gray colspan=2>
							&nbsp;(if yes) When does your student’s IEP need to be renewed?
					</td>										
				</tr>
				<tr>
					<td colspan=2>
						<table>
							<tr>
								<td>
									<select name="IEPmonth"  >
										<option value="">
										<% 
										sqlMonth = "select strValue,strText from common_lists where intList_ID = 1 order by intOrder"
										Response.Write oFunc.MakeListSQL(sqlMonth,"","",iepMonth)								
										%>
									</select>
								</td>		
								<td>
									<select name="IEPday"  >
										<option value="">
										<% 
										sqlDay = "select strValue,strText from common_lists where intList_ID = 2 order by intOrder"
										Response.Write oFunc.MakeListSQL(sqlDay,"","",iepDay)								
										%>
									</select>
								</td>											
								<td>
									<select name="IEPyear"  >	
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
						What is your student’s type of exceptionality?
					</td>
					<td>
						<input type=text size=30 maxlength=128 name="szExceptionality" value="<% = szExceptionality%>">
					</td>
				</tr>
				<TR>
					<td class=gray colspan=2>
						PLEASE BE REMINDED THAT YOU WILL NEED TO HIRE A SPONSOR TEACHER
						TO ATTEND ANY AND ALL IEP MEETINGS WITH YOU.  IF YOU HAVE ANY 
						QUESTIONS PLEASE CALL THE OFFICE IMMEDIATELY.
					</td>
				</tr>
			</table>	
			<BR>-->
			<table>
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
							&nbsp;Last Name*
					</td>
					<td class=gray>
							&nbsp;First Name*
					</td>
					<td class=gray>
							&nbsp;Phone*
					</td>										
				</tr>
				<tr>
					<td>
						<input type=text name="szContact_Last_Name" value="<% = szContact_Last_Name %>" maxlength=50 size=17  >
					</td>
					<td>
						<input type=text name="szContact_First_Name" value="<% = szContact_First_Name %>" maxlength=50 size=15  >
					</td>
					<td>
						<input type=text name="szContact_Phone" value="<% = szContact_Phone %>" maxlength=20 size=15  >
					</td>									
				</tr>
			</table>	
			<table>
				<tr>	
					<Td colspan=6>
						<font class=svplain10>
							Information of Students Doctor
						</font>
					</td>
				</tr>
				<tr>
					<td class=gray>
							&nbsp;Last Name*
					</td>
					<td class=gray>
							&nbsp;First Name*
					</td>
					<td class=gray>
							&nbsp;Phone*
					</td>										
				</tr>
				<tr>
					<td>
						<input type=text name="szDR_Last_Name" value="<% = szDR_Last_Name %>" maxlength=50 size=17   >
					</td>
					<td>
						<input type=text name="szDR_First_Name" value="<% = szDR_First_Name %>" maxlength=50 size=15   >
					</td>
					<td>
						<input type=text name="szDR_Phone" value="<% = szDR_Phone %>" maxlength=20 size=15  >
					</td>									
				</tr>
			</table>	
			<table>
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
						<input type=text name="szDaycare_Name" value="<% = szDaycare_Name%>" maxlength=128 size=25  >
					</td>
					<td>
						<input type=text name="szDaycare_Phone" value="<% = szDaycare_Phone %>" maxlength=20 size=15  >
					</td>									
				</tr>
			</table>		
			<table>
				<tr>
					<td class=gray colspan=2>
							&nbsp;Medic Alert Information
					</td>						
				</tr>
				<tr>
					<td>
						<input type=text name="szMed_Alert_1" value="<% = szMed_Alert_1 %>" maxlength=128 size=25  >
					</td>
					<td>
						<input type=text name="szMed_Alert_2" value="<% = szMed_Alert_2 %>" maxlength=128 size=25  >
					</td>								
				</tr>
			</table>		
			<table>
				<tr>
					<td class=gray colspan=2>
							&nbsp;Disabilities
					</td>						
				</tr>
				<tr>
					<td>
						<input type=text name="szDisability_1" value="<% = szDisability_1 %>" maxlength=128 size=25  >
					</td>
					<td>
						<input type=text name="szDisability_2" value="<% = szDisability_2 %>" maxlength=128 size=25  >
					</td>								
				</tr>
			</table>							
			</td>
		</tr>	
	</table>
	<% if Request.QueryString("bolNewStudent") = "" then %>
		<% if not bolHeader then %>
	&nbsp;&nbsp;<input type=button value="Close Without Saving" class="btSmallGray" onClick="jfConfirm();">	
	<input type=submit value="Save Changes and Close" class="NavSave" onclick="jfSubmit(this.form);">	
		<% else %>
	&nbsp;&nbsp;<input type=submit value="Save Changes" class="NavSave" onclick="jfSubmit(this.form);" NAME="Submit1">	
	<input type=hidden name=bolHeader value="true">
		<% end if %>
	
	<% else %>
		<input type=button value="Cancel" class="btSmallGray" onClick="window.opener.focus();window.close();">	
		<input type=submit value="Add Record" id="NavSave" onclick="jfSubmit(this.form);">	
	<% end if %>	
	</form>
	<DIV ID="divCal" STYLE="position:absolute;visibility:hidden;background-color:white;layer-background-color:white;"></DIV>
	<%
	call oFunc.CloseCN()
	set oFunc = nothing
Server.Execute(Application.Value("strWebRoot") & "Includes/footer.asp")

%>