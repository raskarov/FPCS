<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		Forms\Teachers\addTeacher.asp  
'Purpose:	Add a teacher or Update teacher info
'Date:		9 July 2001
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Make db Connection, print HTML header.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim oFunc
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()
oFunc.ResetSelectSessionVariables()

dim strTitle
strTitle = "Add an ASD Teacher"

if Request.form("intInstructor_ID") <> "" or Request.QueryString("intInstructor_ID") <> "" then
	
	dim sql
	dim month
	dim day
	dim year
	dim intCount
	dim intInstructor_ID
	dim dtMaxDate			'Used to determine the current pay data for a teacher
	dim dblFlatRate			'Holds per diem broken in to hourly rate
	dim dblTaxBen			'Holds hourly plus tax/benefits
	'JD Flat rate
	dim InstructorFlatRate
	
	set rsIns = server.CreateObject("ADODB.RECORDSET")
	sql = "select szTitle,szFirst_Name,szLast_Name,sMid_Initial," & _
		  "szSSN,szMailing_ADDR,szCity,sState,szZip_Code,szHome_Phone,szBusiness_Phone," & _
		  "intBusiness_Ext,szCell_Phone,szEmail,szEmail2,dtBirth,intPay_Type_id," & _
		  "bolMasters_Degree,intDist_Code," & _
		  "bolOn_ASD_Leave,bolSubstitute,bolASD_Employee,bolASD_Eligible_For_Hire,dtCert_Expire," & _
	      "bolASD_Retired,bolGroup_Instruction,bolIndividual_Instruction,strASD_School," & _
		  "intYears_Experience,bolK_8,bolK_12,bolSpecial_Ed,bolSecondary,szSecondary_List," & _
		  "bolMy_Classroom,bolMy_Home,bolStudents_Home,bolFPCS_Classroom,bolOther,szOther_Desc," & _
		  "bolAvail_Weekdays,bolAvail_Wk_Ends,bolAvail_Wk_Afternoon,bolAvail_Wk_Evening,bolAvail_Summers,dtCert_Expire, szSalary_Placement " & _
		  "from tblInstructor " & _
		  "where intInstructor_ID = " & request("intInstructor_ID")
		  'LINE 36 IS FOR TESTING AND WILL NEED TO BE DELETED
	rsIns.Open sql,Application("cnnFPCS")'oFunc.FPCScnn
	

	'This for loop dimentions and defines all the columns we selected in sqlClass
	'and we use the variables created here to populate the form.
	intCount = 0
	for each item in rsIns.Fields
		execute("dim " & rsIns.Fields(intCount).Name)
		execute(rsIns.Fields(intCount).Name & " = item")		
		intCount = intCount + 1
	next

	'added by BKM 26-Apr-202
	szSSN = oFunc.Reformat(szSSN, Array("", 3, "-", 2, "-", 4))
	szZip_Code = oFunc.Reformat(szZip_Code, Array("", 5, "-", 4))
	'szHome_Phone = oFunc.Reformat(szHome_Phone, Array("(", 3, ") ", 3, "-", 4))
	'szBusiness_Phone = oFunc.Reformat(szBusiness_Phone, Array("(", 3, ") ", 3, "-", 4))
	'szCell_Phone = oFunc.Reformat(szCell_Phone, Array("(", 3, ") ", 3, "-", 4))

	
	rsIns.Close
	set rsIns = nothing
		
	strTitle = "View an ASD Teacher Profile"
	if Request.form("intInstructor_ID") <> "" then
		intInstructor_ID = Request.form("intInstructor_ID") <> ""
	elseif Request.QueryString("intInstructor_ID") <> "" then
		intInstructor_ID = Request.QueryString("intInstructor_ID")
	end if	
	
	set rsGetPayData = server.CreateObject("ADODB.RECORDSET")
	rsGetPayData.CursorLocation = 3
	'sql = "select Max(dtEffective_Start) " & _
	'	  "from tblInstructor_Pay_Data " & _
	'	  "where intInstructor_ID = " & request("intInstructor_ID") & _
	'	  " and dtEffective_End is Null " 
	'rsGetPayData.Open sql,cn

	'if rsGetPayData.RecordCount > 0 and rsGetPayData(0) & "" <> "" then
		'dtMaxDate = rsGetPayData(0)
		'rsGetPayData.Close
	
		'sql = "select intInstructor_Pay_Data_ID,intInstructor_ID,curPer_Hour,curPer_Hour_Benefits," & _
		'	  "curPay_Rate,intPay_Type_id,bolASD_Full_Time,decASD_Full_Time_Percent," & _
		'	  "bolASD_Part_Time,decASD_Part_Time_Percent,decFPCS_Hours_Goal,dtEffective_Start " & _
		'	  "from tblInstructor_Pay_Data " & _
		'	  "where intInstructor_ID = " & request("intInstructor_ID") & _
		'	  " and dtEffective_Start <= '6/30/" & (session.Value("intSchool_Year")+1) & "' AND " & _
		'	  "(dtEffective_End IS NULL OR " & _
		'	  " dtEffective_End >= '7/1/" & (session.Value("intSchool_Year")+1) & "')"
		
		sql = "SELECT TOP 1 intInstructor_Pay_Data_ID, intInstructor_ID, " & _
		  " curPay_Rate, intPay_Type_id, bolASD_Full_Time,bolMasters_Degree, szSalary_Placement, " & _
	      " fltASD_Full_Time_Percent, bolASD_Part_Time, fltASD_Part_Time_Percent, " & _
	      " fltFPCS_Hours_Goal,bolActive, dtEffective_Start, dtEffective_End " & _
		  "FROM tblInstructor_Pay_Data " & _
		  "WHERE (intInstructor_ID = " & intInstructor_ID & _
		  ") AND (intSchool_Year_Start <= " & session.contents("intSchool_Year") & ") " & _
		  "ORDER BY dtEffective_Start desc, intInstructor_Pay_Data_ID DESC "
		  'response.Write sql
		  
		rsGetPayData.Open sql,Application("cnnFPCS")'oFunc.FPCScnn
		intCount = 0
		'added the IF NOT..THEN bkm 3-may-2002		
		if not rsGetPayData.BOF and not rsGetPayData.EOF then
			for each item in rsGetPayData.Fields
				execute("dim " & rsGetPayData.Fields(intCount).Name)
				execute(rsGetPayData.Fields(intCount).Name & " = item")		
				intCount = intCount + 1
			next
		end if
	'end if 
	
	rsGetPayData.Close
	set rsGetPayData = nothing
		
	arRate = oFunc.InstructorCosts(request("intInstructor_ID"))
	if isArray(arRate) then
		dblTaxBen = formatNumber(arRate(9),2)
		curPay_Rate = arRate(10)
		dblFlatRate = arRate(0)
	end if
	
	'JD: create the recordset to get the flat instructor rate
	if session.Contents("intSchool_Year") => 2012 then
    sql4 ="select intFlat_Inst_Id, flatRate from tblInstructor_Flat_Rate where intSchool_year = " & session.Contents("intSchool_Year")
    set rs4 = server.CreateObject("ADODB.RECORDSET")
    rs4.CursorLocation = 3
    rs4.Open sql4, Application("cnnFPCS")'oFunc.FPCScnn
    InstructorFlatRate = rs4("flatRate")
    rs4.Close()
    end if

end if 

if request("bolWin") = "" then
	Server.Execute(Application.Value("strWebRoot") & "includes/header.asp")
else
	Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")
end if
%>
<script language=javascript src="<%= Application.Value("strWebRoot") %>includes/CalendarPopup.js"></script>	
<script language=javascript>
	var cal = new CalendarPopup('divCal');
	cal.showNavigationDropdowns();
	cal.setYearSelectStartOffset(10);
	
	var cal2 = new CalendarPopup('divCal');
	cal2.showNavigationDropdowns();
	cal2.setYearSelectStartOffset(70);
	
	<% if request.QueryString("bolForced") <> "" then%>
		alert("You must verify or edit your profile information before proceding.");
	<% end if %>
	
	function jfQuit() {							
		var strPrompt = "You have made changes that will effect the pay rate of this teacher ";
		strPrompt += "which will effect the cost of every contract this teacher has entered into for";
		strPrompt += " the current school year (<% = oFunc.SchoolYearRange() %>) having the potential to put student budgets in the negative. ";		
		strPrompt += "\nAre you sure you want to continue?";		
		if (confirm(strPrompt)) {	
			document.main.bolChangePayData.value = "TRUE";
			return "false";
		}
		else {
			return "true";
		}	
	}
	

	function jfSubmit(objForm) {
		var bolQuit = "false";
		<% if request("intInstructor_ID") <> "" then %>
		if (document.main.bolPrompt.value == "True") {		
			bolQuit = jfQuit();
		}
		<% end if %>
		
		if (bolQuit != "true") {
			if (jfValidate(objForm) == true  ) {
				objForm.submit();
			}
		}
	}
		
	function jfValidate(objForm) {
	//added bkm 26-Apr-2002
	//Ensure all approriate fields have been filled out
		var strErrMsg			= '';
		var strLast_Name		= objForm.szLast_Name.value;
		var strFirst_Name		= objForm.szFirst_Name.value;
		var dtBirth				= objForm.dtBirth.value;
		var strMailing_Addr		= objForm.szMailing_ADDR.value;	
		var strCity				= objForm.szCity.value;	
		var strState			= objForm.sState.value;	
		var strZip_Code			= objForm.szZip_Code.value;	
		var strHome_Phone		= objForm.szHome_Phone.value;
		var strBusiness_Phone	= objForm.szBusiness_Phone.value;
		var strBusiness_Ext		= objForm.intBusiness_Ext.value;
		var strCell_Phone		= objForm.szCell_Phone.value;
		var strSSN				= objForm.szSSN.value;	
		var strEmail			= objForm.szEmail.value;
		var strEmail2			= objForm.szEmail2.value;
		
		var fltASD_Full_Time_Percent = objForm.fltASD_Full_Time_Percent.value;
		var fltASD_Part_Time_Percent = objForm.fltASD_Part_Time_Percent.value;
		var fltFPCS_Hours_Goal = objForm.fltFPCS_Hours_Goal.value;
		var intYears_Experience = objForm.intYears_Experience.value;
		var curPay_Rate = objForm.curPay_Rate.value;

		//these are the required fields - they must be populated
		if(strLast_Name.length == 0) {strErrMsg += 'Last Name\n';}
		if(strFirst_Name.length == 0) {strErrMsg += 'First Name\n';}
		if(dtBirth.length == 0) {strErrMsg += 'Birth Day\n';}
		if(strMailing_Addr.length == 0) {strErrMsg += 'Mailing Address\n';}
		if(strCity.length == 0) {strErrMsg += 'City\n';}
		if(strState.length == 0) {strErrMsg += 'State\n';}
		if(strZip_Code.length == 0) {strErrMsg += 'ZIP Code\n';}
		if(strHome_Phone.length == 0) {strErrMsg += 'Home Phone\n';}
		if(strEmail.length == 0) {
			strErrMsg += 'Email\nIf you do not have an email account you can get one free at www.hotmail.com';
		}else{
			if (isEmail(strEmail) == false) {strErrMsg += 'Email Address is invalid\n';}
		}
		
		
		if(strEmail2.length != 0) {
			if (isEmail(strEmail2) == false) {strErrMsg += '2nd Email Address is invalid\n';}
		}
		if(fltASD_Full_Time_Percent.length != 0) {
			if (isFloat(fltASD_Full_Time_Percent) == false) {strErrMsg += 'Full-time Teacher % must be a number\n';}
		}
		if(fltASD_Part_Time_Percent.length != 0) {
			if (isFloat(fltASD_Part_Time_Percent) == false) {strErrMsg += 'Part-time Teacher % must be a number\n';}
		}
		if(fltFPCS_Hours_Goal.length != 0) {
			if (isFloat(fltFPCS_Hours_Goal) == false) {strErrMsg += 'Goal for FPCS hours % must be a number\n';}
		}
		if(intYears_Experience.length != 0) {
			if (isInteger(intYears_Experience) == false) {strErrMsg += 'Years of Teaching Experience must be a number\n';}
		}
		if(curPay_Rate.length != 0) {
			if (isFloat(curPay_Rate) == false) {strErrMsg += 'Per Deim Rate must be a number\n';}
		}

		if (strErrMsg.length == 0 ) {
			//if all of the required fields are populated then we test the values
			//in some of them.  Additionaly, if some UnRequired fields are populated,
			//we test their values as well
			if (checkZIPCode(objForm.szZip_Code) == false) {return false;}
			if(strSSN.length != 0) {
				if (checkSSN(objForm.szSSN) == false) {return false;}
			}
			if (checkUSPhone(objForm.szHome_Phone) == false) {return false;}
			if(strBusiness_Phone.length != 0){
				if (checkUSPhone(objForm.szBusiness_Phone) == false) {return false;}
			}
			if(strCell_Phone.length != 0){
				if (checkUSPhone(objForm.szCell_Phone) == false) {return false;}
			}
			//if (checkEmail(objForm.szEmail) == false) {return false;}
			if(strEmail2.length != 0) {
				if (checkEmail(objForm.szEmail2) == false) {return false;}
			}

			
			//items below have already passed validation - we strip characters to pass raw data to database
			objForm.szZip_Code.value = stripCharsInBag(objForm.szZip_Code.value, ZIPCodeDelimiters);
			if(strSSN.length != 0) {
				objForm.szSSN.value = stripCharsInBag(objForm.szSSN.value, SSNDelimiters);
			}
			objForm.szHome_Phone.value = stripCharsInBag(objForm.szHome_Phone.value, phoneNumberDelimiters);
			if(strBusiness_Phone.length != 0) {
				objForm.szBusiness_Phone.value = stripCharsInBag(objForm.szBusiness_Phone.value, phoneNumberDelimiters);
			}
			if(strCell_Phone.length != 0) {
				objForm.szCell_Phone.value = stripCharsInBag(objForm.szCell_Phone.value, phoneNumberDelimiters);
			}
			
			return true;
		} else {
			strErrMsg = 'Please Enter the Following:\n \n' + strErrMsg;
			alert(strErrMsg);
			return false;
		}
	}

	function jfSetPrompt(){
		document.main.bolPrompt.value = "True";
	}
	
	function jfViewAnotherInstructor(id){
		var strURL = "<%=Application.Value("strWebRoot")%>forms/Teachers/addTeacher.asp?intInstructor_ID=" + id.value;
		window.location.href = strURL;
	}
</script>

<form action="TeacherInsert.asp" method=Post name=main onSubmit="return false;">
<input type=hidden name="bolChangePayData" value="">
<input type=hidden name=changed value="">
<input type=hidden name="intInstructor_Pay_Data_ID" value="<% = intInstructor_Pay_Data_ID %>">
<% if session.Contents("strRole") <> "ADMIN" or request("intInstructor_ID") = "" then%>
<input type=hidden name="intInstructor_ID" value="<% = intInstructor_ID%>">
<% end if %>
<input type=hidden name="bolPrompt" value="">
<table width=100%>
	<tr>	
		<Td class=yellowHeader>
				&nbsp;<b><% = strTitle %></b>
				<% if Session.Contents("strRole") = "ADMIN"  and request("intInstructor_ID") <> "" then %>
				<select name="intInstructor_ID" onChange="jfViewAnotherInstructor(this);">
					<option value="">
						<%
				dim sqlInstructor
				sqlInstructor = "Select intInstructor_ID,szLast_Name + ',' + szFirst_Name + ': ' + convert(varchar,intInstructor_ID) as Name " & _
									"from tblInstructor order by szLast_Name"
				Response.Write oFunc.MakeListSQL(sqlInstructor,intStudent_ID,Name,intInstructor_ID)												 
						%>
				</select>
				<% end if %>
		</td>
	</tr>
	<tr>
		<td bgcolor=f7f7f7>
			<table>
				<tr>	
					<Td colspan=2>
						<font class=svplain11>
							<b><i>Teacher Information</I></B> 
						</font>
						<font class=svplain>
						</font>
					</td>
				</tr>
				<tr>
					<td class=gray>
							&nbsp;Title
					</td>
					<td class=gray>
							&nbsp;First Name
					</td>		
					<td class=gray>
							&nbsp;M.I.
					</td>	
					<td class=gray>
							&nbsp;Last Name&nbsp;
					</td>
					<td class=gray >
							&nbsp;Date of Birth
					</td>	
					<td class=gray>
						&nbsp;Is Active&nbsp;
					</td>												
				</tr>
				<tr>
					<td>
						<select name="szTitle" onChange="jfChanged();">
							<% 
								dim strTeacherTitle
								strTeacherTitle = "Mr.,Mrs.,Ms.,Dr."
								Response.Write oFunc.MakeList(strTeacherTitle,strTeacherTitle,szTitle)																
							%>
						</select>				
					</td>
					<td>
						<input type=text name="szFirst_Name" value="<% = szFirst_Name%>" maxlength=50 size=15 onChange="jfChanged();">
					</td>
					<td>
						<input type=text name="sMid_Initial" value="<% = sMid_Initial%>" maxlength=1 size=2 onChange="jfChanged();">
					</td>
					<td>
						<input type=text name="szLast_Name" value="<% = szLast_Name%>" maxlength=50 size=17 onChange="jfChanged();">
					</td>		
					<td class="svplain8">
						<input type=text name="dtBirth" size=10 value="<% = dtBirth %>" maxlength=10 onChange="jfChanged();" ID="Text2" > 
						<a href="#" onclick="cal2.select(document.forms[0].dtBirth,'aBirth','M/d/yyyy','0<% = trim(dtBirth) %>');jfChanged();return false;" id="aBirth" name="aBirth">select</a>
					</td>	
					<td class=gray>
						<%  dim strChecked 
							strChecked = ""
						    if bolActive = "True" then
								strChecked = " checked "
							end if
						%>
						<input type=checkbox name="bolActive" value="1" <% = strChecked %>  onChange="jfChanged();document.main.bolChangePayData.value = 'TRUE';"> <b>yes</b>
					</td>							
				</tr>
			</table>
			<table>
				<tr>
					<td class=gray>
							&nbsp;Mailing Address
					</td>
					<td class=gray>
							&nbsp;City
					</td>		
					<td class=gray>
							&nbsp;State
					</td>	
					<td class=gray>
							&nbsp;Zip Code&nbsp;
					</td>												
				</tr>
				<tr>
					<td>
						<input type=text name="szMailing_ADDR" value="<% = szMailing_ADDR%>" maxlength=256 size=50 onChange="jfChanged();">
					</td>
					<td>
						<input type=text name="szCity" value="<% = szCity%>" maxlength=50 size=25 onChange="jfChanged();">
					</td>
					<td>
						<select name="sState" onChange="jfChanged();">
						<%
							dim sqlState
							sqlState = "select strValue,strText from Common_Lists where intList_Id = 3 order by strValue"
							Response.Write oFunc.MakeListSQL(sqlState,"","",session("sState"))
						%>
						</select>
					</td>
					<td>
						<input type=text name="szZip_Code" value="<% = szZip_Code%>" maxlength=10 size=7 onChange="jfChanged(); checkZIPCode(this);">
					</td>									
				</tr>
			</table>
			<table>
				<tr>
					<td class=gray>
							&nbsp;Home Phone
					</td>
					<td class=gray>
							&nbsp;Bus. Phone
					</td>
					<td class=gray>
							&nbsp;Ext
					</td>
					<td class=gray>
							&nbsp;Cell Phone
					</td>
					<Td class=gray>
							&nbsp;Masters Degree
					</td>	
					<td class=gray>
							&nbsp;SSN
					</td>
					
					<!--<Td class=gray>
							&nbsp;Teacher Type
					</td>-->										
				</tr>
				<tr>
					<td>
						<input type=text name="szHome_Phone" value="<% = szHome_Phone %>" maxlength=15 size=15 onChange="jfChanged(); checkUSPhone(this);">
					</td>
					<td>
						<input type=text name="szBusiness_Phone" value="<% = szBusiness_Phone %>" maxlength=15 size=15 onChange="jfChanged(); checkUSPhone(this);">
					</td>
					<td>
						<input type=text name="intBusiness_Ext" value="<% = intBusiness_Ext %>" maxlength=4 size=4 onChange="jfChanged();">
					</td>
					<td>
						<input type=text name="szCell_Phone" value="<% = szCell_Phone %>" maxlength=15 size=15 onChange="jfChanged(); checkUSPhone(this);">
					</td>	
					<td class=svplain10 align=center>
						<b>yes</b> <input type=checkbox name=bolMasters_Degree <% if bolMasters_Degree then Response.Write " checked " %>  onChange="jfChanged();jfSetPrompt();">										
					<td>
						<input type=text name="szSSN" value="<% = szSSN %>" maxlength=15 size=15 onChange="jfChanged(); checkSSN(this);">
					</td>
					<!--<td>
						<select name="bolContract_ASD" onChange="jfChanged();">
						<% 
							dim strContractVal
							dim strContractText
							strContractASD = "-1,1"
							strContractText = "Sponcer Teacher,Contract ASD Teacher"
							'Response.Write vbfMakeList(strContractASD,strContractText,vbfTrueFalse(bolContract_ASD))
						%>
						</select>
					</td>-->
				</tr>
			</table>	
			<table>
				<tr>	
					<Td class=gray>
							&nbsp;Email Address
					</td>	
					<Td class=gray>
							&nbsp;Second Email Address 
					</td>	

					<td class=gray>
							&nbsp;District Code
					</td>
					<!--
					<td class=gray>
							&nbsp;Work Location
					</td> -->											
				</tr>
				<tr>
					<td>
						<input type=text name="szEmail" value="<% = szEmail %>" maxlength=64 size=30 onChange="jfChanged();">
					</td>
					<td>
						<input type=text name="szEmail2" value="<% = szEmail2 %>" maxlength=64 size=30 onChange="jfChanged();">
					</td>	
					
					<td>
						<% if intDist_Code = "" then intDist_Code = "540" %>
						<input type=text name="intDist_Code" value="<% = intDist_Code %>" maxlength=15 size=15 onChange="jfChanged();">
					</td>
					<!--
					<td>
						<input type=text name="intWK_Location" value="<% = intWK_Location %>" maxlength=15 size=15 onChange="jfChanged();">
					</td> -->
				</tr>
			</table>
			<table>
				<Tr>
					<Td>
						<table>
							<tr>	
								<Td class=gray>
										&nbsp;Status with Anchorage School District besides FPCS<BR>
										&nbsp;(Check only 1 box for benefit pay)
								</td>						
							</tr>
							<tr>
								<td class=svplain10>
									<input type=checkbox name="bolASD_Full_Time" <% if bolASD_Full_Time then Response.Write " checked " %> onChange="jfChanged();jfSetPrompt();">
									Benefit paid by other ASD school <input type=text size=3 maxlength=3 name="fltASD_Full_Time_Percent" value="<%=fltASD_Full_Time_Percent%>" onChange="jfChanged();jfSetPrompt();">% (ASD FTE)
								</td>
							</tr>
							<Tr>
								<td class=svplain10>
									<input type=checkbox name="bolASD_Part_Time" <% if bolASD_Part_Time then Response.Write " checked " %> onChange="jfChanged();jfSetPrompt();">
									Check only if less than 50% (Will not include benefits in calculations) <input type=text size=3 maxlength=3 name="fltASD_Part_Time_Percent" value="<%=fltASD_Part_Time_Percent%>" onChange="jfChanged();jfSetPrompt();">% (FTCS FTE)
								</td>	
							</tr>
							<Tr>
								<td class=svplain10>									
									Name of School <input type=text size=15 maxlength=50 name="strASD_School" value="<%=strASD_School%>" onChange="jfChanged();">
								</td>	
							</tr>
							<tr>
								<td class=svplain10>
									<input type=checkbox name="bolOn_ASD_Leave" <% if bolOn_ASD_Leave then Response.Write " checked " %> onChange="jfChanged();">
									On leave from ASD
								</td>	
							</tr>
							<tr>
								<td class=svplain10>
									<input type=checkbox name="bolSubstitute" <% if bolSubstitute then Response.Write " checked " %> onChange="jfChanged();">
									Substitute Teacher
								</td>	
							</tr>
							<tr>
								<td class=svplain10>
									<input type=checkbox name="bolASD_Employee" <% if bolASD_Employee then Response.Write " checked " %> onChange="jfChanged();">
									Other ASD employee (not teaching)
								</td>	
							</tr>
							<tr>
								<td class=svplain10>
									<input type=checkbox name="bolASD_Eligible_For_Hire" <% if bolASD_Eligible_For_Hire then Response.Write " checked " %> onChange="jfChanged();">
									On ASD eligible-to-hire list
								</td>	
							</tr>
							<tr>
								<td class=svplain10>
									<input type=checkbox name="bolASD_Retired" <% if bolASD_Retired  then Response.Write " checked " %> onChange="jfChanged();">
									Retired ASD Teacher
								</td>	
							</tr>
						</table>
					</td>
					<Td valign=top>
						<table>
							<tr>	
								<Td class=gray>
										&nbsp;FPCS FTE Percentage:<BR>
										&nbsp;(Percentage based on 100% = 1365 hrs) 20% = 0.2 FTE
								</td>						
							</tr>
							<tr>
								<td class=svplain10>
									<input type=text name="fltFPCS_Hours_Goal" size=5 value="<% = fltFPCS_Hours_Goal %>" maxlength=5 onChange="jfChanged();jfSetPrompt();">%
								</td>
							</tr>
						</table>
						<table>
							<tr>	
								<Td class=gray>
										&nbsp;I am interested in teaching:
								</td>						
							</tr>
							<tr>
								<td class=svplain10>
									<input type=checkbox name="bolGroup_Instruction" <% if bolGroup_Instruction then Response.Write " checked " %> onChange="jfChanged();">
									Group Instruction
								</td>
							</tr>
							<tr>
								<td class=svplain10>
									<input type=checkbox name="bolIndividual_Instruction" <% if bolIndividual_Instruction then Response.Write " checked " %> onChange="jfChanged();">
									Individual Instruction (one on one)
								</td>
							</tr>
						</table>
						<table>
							<tr>	
								<Td class=gray>
										&nbsp;Years of Teaching Experience:
								</td>						
							</tr>
							<tr>
								<td class=svplain10>
									<input type=text name="intYears_Experience" size=3 value="<% = intYears_Experience %>" maxlength=3 onChange="jfChanged();"> years
								</td>
							</tr>
						</table>
						<table ID="Table1" cellpadding="3">
							<tr>	
								<Td class=gray>
										Teaching Certificate<BR>Expiration Date
								</td>	
								<Td class=gray align="center">
										Teaching Salary Placement
								</td>					
							</tr>
							<tr>
								<td class=svplain8>
									<input type=text name="dtCert_Expire" size=10 value="<% = dtCert_Expire %>" maxlength=10 onChange="jfChanged();" ID="Text1" > 
									<a href="#" onclick="cal.select(document.forms[0].dtCert_Expire,'certExpire','M/d/yyyy','<% = trim(dtCert_Expire) %>'); jfChanged();return false;" id="certExpire" name="certExpire">select</a>
								</td>
								<td class=svplain8>
									<input type="text" name="szSalary_Placement" value="<% = szSalary_Placement %>" maxlength="63" ID="Text3" size=25 onChange="jfChanged();jfSetPrompt();" >
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			
			<table>
				<tr>	
					<Td class=gray>
							&nbsp;Alaska Certification: (Check all that apply)
					</td>						
				</tr>
				<tr>
					<td class=svplain10>
						<input type=checkbox name="bolK_8" <% if bolK_8 then Response.Write " checked " %> onChange="jfChanged();">
						K-8 |
						<input type=checkbox name="bolK_12" <% if bolK_12 then Response.Write " checked " %> onChange="jfChanged();">
						K-12 |
						<input type=checkbox name="bolSpecial_Ed" <% if bolSpecial_Ed then Response.Write " checked " %> onChange="jfChanged();">
						Special Education
					</td>
				</tr>
				<tr>
					<td class=svplain10>
						<input type=checkbox name="bolSecondary" <% if bolSecondary then Response.Write " checked " %> onChange="jfChanged();">
						Secondary (list subject and grades)
						<input type=text size=20 name="szSecondary_List" value="<% = szSecondary_List %>" onChange="jfChanged();">
					</td>
				</tr>
			<table>
			
			<table>
				<tr>
					<td>
						<table>
							<tr>	
								<Td class=gray>
										&nbsp;I am available to teach FPCS students:<BR>
										&nbsp;(Check all that apply)
								</td>						
							</tr>
							<tr>
								<td class=svplain10>
									<input type=checkbox name="bolMy_Classroom" <% if bolMy_Classroom then Response.Write " checked " %> onChange="jfChanged();">
								in my classroom
								</td>
							</tr>
							<tr>
								<td class=svplain10>
									<input type=checkbox name="bolMy_Home" <% if bolMy_Home then Response.Write " checked " %> onChange="jfChanged();">
									at my home
								</td>
							</tr>
							<tr>
								<td class=svplain10>
									<input type=checkbox name="bolStudents_Home" <% if bolStudents_Home then Response.Write " checked " %> onChange="jfChanged();">
									at student's home
								</td>
							</tr>
							<tr>
								<td class=svplain10>
									<input type=checkbox name="bolFPCS_Classroom" <% if bolFPCS_Classroom then Response.Write " checked " %> onChange="jfChanged();">
									at FPCS classroom
								</td>
							</tr>
							<tr>
								<td class=svplain10>
									<input type=checkbox name="bolOther" <% if bolOther then Response.Write " checked " %> onChange="jfChanged();">
									Other <input type=text size=15 name="szOther_Desc" value="<% = szOther_Desc %>" onChange="jfChanged();">
								</td>
							</tr>
						</table>	
					</td>
					<td valign=top>
						<table>
							<tr>	
								<Td class=gray>
										&nbsp;I am available:<BR>
										&nbsp;(Check all that apply)
								</td>						
							</tr>
							<tr>
								<td class=svplain10>
									<input type=checkbox name="bolAvail_Weekdays" <% if bolAvail_Weekdays then Response.Write " checked " %> onChange="jfChanged();">
								weekdays
								</td>
							</tr>
							<tr>
								<td class=svplain10>
									<input type=checkbox name="bolAvail_Wk_Afternoon" <% if bolAvail_Wk_Afternoon then Response.Write " checked " %> onChange="jfChanged();">
									weekday afternoons
								</td>
							</tr>
							<tr>
								<td class=svplain10>
									<input type=checkbox name="bolAvail_Wk_Evening" <% if bolAvail_Wk_Evening then Response.Write " checked " %> onChange="jfChanged();">
									weekday evenings
								</td>
							</tr>
							<tr>
								<td class=svplain10>
									<input type=checkbox name="bolAvail_Wk_Ends" <% if bolAvail_Wk_Ends then Response.Write " checked " %> onChange="jfChanged();">
									weekends
								</td>
							</tr>
							<tr>
								<td class=svplain10>
									<input type=checkbox name="bolAvail_Summers" <% if bolAvail_Summers then Response.Write " checked " %> onChange="jfChanged();">
									summers
								</td>
							</tr>
						</table>		
					</td>
				</tr>
			</table>	
			<table>		
				<tr>
					<td class=gray rowspan="2" align="center">
							This pay data is in effect <BR>as of 
							<% = dtEffective_Start %>
					</td>
					<!--JD flat rate-->
					<% if session.Contents("intSchool_Year") => 2012 then%>
					<td class=gray>
					        &nbsp;Flat Rate per Hour
					</td>
					<% end if%>
					<td class=gray>
							&nbsp;Base Pay per Hour
					</td>
					<td class=gray>
							&nbsp;Pay per Hour w/Benefits
					</td>		
					<td class=gray>
							&nbsp;Per Deim Rate
					</td>		
					<td class=gray>
							&nbsp;Pay Type
					</td>													
				</tr>
				<tr>
					<!--JD flat rate-->
					<td align=center class="svplain10">
						$<% = formatNumber(InstructorFlatRate, 2) %>
					</td>

					<td align=center class="svplain10">
						$<% = dblFlatRate %>
					</td>
					<td align=center class="svplain10">
						$<% = dblTaxBen%>
					</td>
					<td align=center>
						<input type=text name="curPay_Rate" value="<% = curPay_Rate %>" maxlength=8 size=8 onChange="jfChanged();jfSetPrompt();">
					</td>
					<td>
						<select name="intPay_Type_id" onChange="jfChanged();jfSetPrompt();">
							<option value="">
						<% 
							dim strPayType
							strPayType = "select intPay_type_id, szPay_Type_Name from trefPay_Types order by szPay_Type_Name"
							Response.Write oFunc.MakeListSQL(strPayType,"intPay_type_id","szPay_Type_Name",intPay_type_id)
						%>
						</select>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<% 
if ucase(session.Contents("strRole")) = "ADMIN" then
if request("intInstructor_ID") <> "" then 
	if request("bolWin") <> "" then	
%>
<input type=hidden name="bolWin" value="True">
<input type=button value="CLOSE" class="btSmallGray" onClick="window.opener.focus();window.close();">
<%	elseif Request.QueryString("bolForced") = "" then  %>
<input type=button value="Home Page" class="btSmallGray" onClick="window.location.href='<%=Application.Value("strWebRoot")%>';">
<%	end if  
	if Request.QueryString("bolForced") <> "" then  
%>
<input type=hidden name=intCount value="<% = Request.QueryString("intCount")%>">
<input type=submit value="ACCEPT" class="NavSave" name="btAccept" onClick="jfSubmit(this.form);">
<%  end if %>		
<input type=submit value="UPDATE" class="NavSave" onClick="jfSubmit(this.form);">
<% else %>
<input type=submit value="ADD TEACHER" class="NavSave" onClick="jfSubmit(this.form);">
<% end if 
end if%>
<DIV ID="divCal" STYLE="position:absolute;visibility:hidden;background-color:white;layer-background-color:white;"></DIV>
</form>
</BODY>
</HTML>
<% 
oFunc.CloseCN
%>