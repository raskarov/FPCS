<%@ Language=VBScript %>
<%

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		lottryMain.asp
'Purpose:	This script is the main data gathering/launch point
'			for the required lottery information. 
'Date:		4-2-2003
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim oFunc			'Main object that exposes many of our custom functions 
dim oVal			'Form Validation Object
dim strInstructions	'Holds our instruction text for guiding the user
dim strError		'Contains validation error information 
dim strStudentList	'Contains HTML list of students with edit and delete buttons
dim bolShowFinish	'Tells us whether or not to show the finish button. Requires at least one student and guardian.
dim strAddGuardList	'Contains HTML list of additionl guardians with edit and delete buttons

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

set oVal = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/formValidation.wsc"))

' This will tell us if we are using the main form or the additional guardian form
bolAdditionalGuardian = request.QueryString("bolAdditionalGuardian")
intGuardian_ID = request.QueryString("intGuardian_ID")
bolDelete = request.QueryString("bolDelete")
bolConfirm = request.QueryString("bolConfirm")

if bolDelete <> "" then
	call vbsConfirmDelete
elseif bolConfirm <> "" then
	call vbsDeleteGuardian
end if 

if request.Form.Count > 0 then
	' Transfers all of the post http header variables into vbs variables
	' so we can more readily access them
	for each i in request.Form
		execute("dim " & i)
		execute(i & " = """ & request.Form(i) & """")
	next 
end if 

' initialize instructions 
if session.Contents("strInstructions") = "" then
	session.Contents("strInstructions") = "To start off please provide the Guardian information. A '*' " & _
					"signifies that the information is required. When you finish " & _
					"entering the guardian information click 'SAVE' at the bottom " & _
					"of this page. The next step will be adding student information. " 
elseif bolAdditionalGuardian <> "" then
	session.Contents("strInstructions") = "Fill out the form below to add an additional parent/guardian. " & _
									  "To save this information you must click the 'Save Guardian Info' " & _
									  "button at the bottom of the page. * denotes required information."
	
		
end if 				
				  
' Handle Inserts and Updates only if the 'SAVE' button was clicked
if saveGuardian <> "" or finished <> "" then
	if bolAdditionalGuardian <> "" then
		strError = vbfValidateGuardian("")	
	else
		' Validate Form data for Guardians. Primary guardian data can never be blank.	
		strError = vbfValidateGuardian(1)
	end if
		
	if strError = "" and (szLast_Name2 <> "" or szFirst_Name2 <> "") then
		strError = strError & vbfValidateGuardian(2)
	end if 

	if strError = "" and session.Contents("intGuardian_ID1") = "" and szLast_Name1 <> "" then
		call vbsInsertGuardian(1)
	elseif strError = "" and session.Contents("intGuardian_ID1") <> "" and bolAdditionalGuardian = "" then
		call vbsUpdateGuardian(1)
	end if 
	
	if strError = "" and session.Contents("intGuardian_ID2") = "" and szLast_Name2 <> "" then
		call vbsInsertGuardian(2)
	elseif strError = "" and session.Contents("intGuardian_ID2") <> "" and bolAdditionalGuardian = "" then		
		call vbsUpdateGuardian(2)
	end if 

	if strError = "" and intGuardian_ID = "" and szLast_Name <> "" and bolAdditionalGuardian <> "" then				
		call vbsInsertGuardian("")
	elseif strError = "" and intGuardian_ID <> "" and bolAdditionalGuardian <> "" then
		call vbsUpdateGuardian("")		
	end if 
	
	' If finished then lets make our exit 
	if finished <> "" then
		response.Redirect("finished.htm")
		oFunc.CloseCN()
		set oFunc = nothing
		set oVal = nothing
		response.End
	end if
end if 

if (session.Contents("intFamily_ID") <> "" and strError = "" and bolAdditionalGuardian = "") _
	or (intGuardian_ID <> "") then
	' Retrieve and populate guardian information 
	dim sql				' Contains our sql statement
	dim intGuardCount	'used so we can keep track of first 2 guardians
	dim intCount		'used to itterate through recordset field collection 
	
	set rsGuardInfo = server.CreateObject("ADODB.RECORDSET")
	rsGuardInfo.CursorLocation = 3
	
	if intGuardian_ID <> "" then
		'View single guardian only 
		strWhere = " AND g.intGUARDIAN_ID = " & intGuardian_ID
	end if
	
	sql = "SELECT g.intGUARDIAN_ID, g.szFIRST_NAME, g.sMID_INITIAL, g.szLAST_NAME, " & _
		    " g.szEMPLOYER, g.szBUSINESS_PHONE, g.intPHONE_EXT,  " & _
			" g.szCELL_PHONE, g.szPAGER, g.bolACTIVE_MILITARY, g.szRANK, g.szEMAIL, " & _
			" g.szAddress, g.szCity, g.szState, g.szCountry, g.szZip_Code,  " & _
			" g.szHome_Phone, g.bolSnail_Mail " & _
			"FROM tascFAM_GUARD fg INNER JOIN " & _
			" tblGUARDIAN g ON fg.intGUARDIAN_ID = g.intGUARDIAN_ID " & _
			"WHERE (fg.intFamily_ID = " & session.Contents("intFamily_ID") & ") " & _
			strWhere & _
			" order by g.intGUARDIAN_ID "
	rsGuardInfo.Open sql, oFunc.FPCScnn
	
	if rsGuardInfo.RecordCount > 0 then
		intGuardCount = 1
		do while not rsGuardInfo.EOF
			if intGuardCount < 3 then
				' For the first 2 guardians dynamically create and assign
				' variables to contain guardian information
				intCount = 0 
				for each item in rsGuardInfo.Fields
					'Set intGuardCount to "" since when we use lotteryAdditionalGuardian.asp
					'we do not append numbers to the end of our fields
					if bolAdditionalGuardian <> "" then intGuardCount = ""					
					execute("dim " & rsGuardInfo.Fields(intCount).Name & intGuardCount)
					execute(rsGuardInfo.Fields(intCount).Name & intGuardCount & " = item")
					intCount = intCount + 1
				next	
				if bolAdditionalGuardian = "" then
					intGuardCount = intGuardCount + 1
				end if
			else
				strAddGuardList = strAddGuardList & "<tr><td class=svplain10>" & _
							 rsGuardInfo("szFIRST_NAME") & " " & rsGuardInfo("szLAST_NAME") & "</td><td>" & _
							 "<input type=button onclick=""window.location.href=" & _
							 "'lotteryMain.asp?intGuardian_ID=" & rsGuardInfo("intGUARDIAN_ID") & _
							 "&bolAdditionalGuardian=true';"" id=""btSmallGray"" value=""View/Edit Guardian Info."">" & _
							 "<input type=button onclick=""window.location.href=" & _
							 "'lotteryMain.asp?intGUARDIAN_ID=" & rsGuardInfo("intGUARDIAN_ID") & _
							 "&bolDelete=true&guardianName=" & rsGuardInfo("szFIRST_NAME") & " " & rsGuardInfo("szLAST_NAME") & "';"" id=""btSmallGray"" value=""Delete Guardian Entry"">" & _
							 "</td></tr>"
			end if
			rsGuardInfo.MoveNext
		loop
	end if		
	rsGuardInfo.Close
	
	' Retrieve students if they exist
	sql = "SELECT szLAST_NAME + ', ' + szFIRST_NAME as Name, intSTUDENT_ID " & _
			"FROM tblSTUDENT " & _
			"WHERE (intFamily_ID = " & session.Contents("intFamily_ID") & ")"	
	rsGuardInfo.Open sql, oFunc.FPCScnn
	
	if rsGuardInfo.RecordCount > 0 then
		do while not rsGuardInfo.EOF
			strStudentList = strStudentList	& "<tr><td class=svplain10>" & _
							 rsGuardInfo("Name") & "</td><td>" & _
							 "<input type=button onclick=""window.location.href=" & _
							 "'lotteryStudent.asp?intStudent_ID=" & rsGuardInfo("intstudent_ID") & _
							 "';"" id=""btSmallGray"" value=""View/Edit Student Info."">" & _
							 "<input type=button onclick=""window.location.href=" & _
							 "'lotteryStudent.asp?intStudent_ID=" & rsGuardInfo("intstudent_ID") & _
							 "&bolDelete=true&studentName=" & rsGuardInfo("Name") & "';"" id=""btSmallGray"" value=""Delete Student Entry"">" & _
							 "</td></tr>"
			rsGuardInfo.MoveNext
		loop
		session.Contents("strInstructions") = session.Contents("strInstructions") & _
						" When you are finished entering student and guardian informtion " & _
						" click the 'Finished' button. If you make any changes to the guardian information" & _
						" on this page click the 'Save Guardian Info' button at the bottom of this page."
		bolShowFinish = true
	else
		 session.Contents("strInstructions") = session.Contents("strInstructions") & " Next step is to add " & _
					  "student information. To do so simply click the 'Add a Student' " & _
					  "button.  To modify your guardian information make your changes and " & _
					  "then click the 'Save Guardian Info' button."
		bolShowFinish = false
	end if 
	rsGuardInfo.Close
	set rsGuardInfo = nothing
end if

'Print HTML header information
%>
<html>
	<head>
		<title>Enrollment Step 2</title>
		<link rel="stylesheet" type="text/css" href="../../css/homestyle.css">
	</head>
	<body bgcolor="white">
		<table ID="Table1">			
			<form action="lotteryMain.asp" name="main" method="post" ID="Form1">
				<input type=hidden name="bolFulltime" value="<%=bolFulltime%>"> 
				<input type=hidden name="bolComeToMeeting" value="<%=bolComeToMeeting%>">
				<input type=hidden name="bolVolunteer" value="<%=bolVolunteer%>">
			<% if bolAdditionalGuardian = ""  then %>
				<tr>
					<td class="gray">
						&nbsp;&nbsp;&nbsp;<b>Welcome to the family enrollment page.</b> Completion of 
						this information<br>
						&nbsp; initiates the enrollment of your children in Frontier Charter School and<br>
						&nbsp; provides Anchorage School District with the data required for their 
						records.<br>
						&nbsp; After Frontier receives this page, you will be contacted about how to<br>
						&nbsp; complete the enrollment process.
					</td>
				</tr>
			<%	end if %>
				<tr>
					<td class="svplain10">
						<table cellspacing="0" cellpadding="4" bordercolor="e6e6e6" border="1" ID="Table3">
							<tr>
								<td class="svplain10">
									<b>Instructions:</b><br>
									<% = session.Contents("strInstructions") %>
									<BR>
									<% = strError %>
								</td>
							</tr>
						</table>
						<br>
					</td>
				</tr>							
			<%  if bolAdditionalGuardian <> ""  then 
					' Shows form for adding a single guardian			
					call vbsAdditionalGuardianForm
			   else
			%>
				<% if session.Contents("intFamily_ID") <> "" and strError = "" then %>
								
				<tr>
					<td class="NavyHeader">
						&nbsp;<B>Students Requesting Enrollment </B>&nbsp;&nbsp;&nbsp;&nbsp; <input type="button" value="Add a Student" id="btSmallGray" onclick="window.location.href ='lotteryStudent.asp';" NAME="Button3">
					</td>
				</tr>
				<tr>
					<td class="svplain10">
						<table ID="Table2" cellpadding="3" cellspacing="0" bordercolor="e6e6e6" border="1">
							<% = strStudentList %>
						</table>
						<br>
					</td>
				</tr>
					<%
					end if 
					%>
				<tr>
					<td class="NavyHeader">
						&nbsp;<B>Parent/Guardian Information</B>
						<%
						if session.Contents("intGuardian_ID1") <> "" and session.Contents("intGuardian_ID2") <> "" then
						%>
						<input type="button" value="Add Another Parent/Guardian" id="btSmallGray" onclick="window.location.href='lotteryMain.asp?bolAdditionalGuardian=true';">
						<%
						end if 
						%>
					</td>
				</tr>
				<tr>
				<tr>
					<td class="svplain10">
						<table ID="Table13" cellpadding="3" cellspacing="0" bordercolor="e6e6e6" border="1">
							<% = strAddGuardList %>
						</table>
						<br>
					</td>
				</tr>
					<td>
						<table ID="Table4">
							<tr>
								<Td colspan="6">
									<font class="svplain10"><b><i>Primary Parent/Guardian </i></b></font>
								</Td>
							</tr>
							<tr>
								<td class="gray">
									&nbsp;Last Name*
								</td>
								<td class="gray">
									&nbsp;First Name*
								</td>
								<td class="gray">
									&nbsp;MI*
								</td>
								<td class="gray">
									&nbsp;Email Address*&nbsp;
								</td>
							</tr>
							<tr>
								<td>
									<input type="text" name="szLast_Name1" value="<%= szLast_Name1%>" maxlength="50" size="17"   ID="Text1">
								</td>
								<td>
									<input type="text" name="szFirst_Name1" value="<%= szFirst_Name1%>" maxlength="50" size="15"   ID="Text2">
								</td>
								<td>
									<input type="text" name="sMid_Initial1" value="<%= sMid_Initial1%>" maxlength="1" size="2"   ID="Text8">
								</td>
								<td class="svplain10">
									<input type=text name="szEmail1" size=25 value="<%= szEmail1%>" maxlength=128>
								</td>
							</tr>
						</table>
						<table ID="Table5">
							<tr>
								<Td class="gray">
									&nbsp;Employer
								</Td>
								<td class="gray">
									&nbsp;Active Military&nbsp;
								</td>
								<td class="gray">
									&nbsp;Rank&nbsp;
								</td>
								<td class="gray">
									&nbsp;Pager
								</td>
							</tr>
							<tr>
								<td>
									<input type="text" name="szEmployer1" value="<%= szEmployer1%>" maxlength="128" size="30"   ID="Text9">
								</td>
								<td>
									<select name="bolActive_Military1" ID="Select2">
										<option value="">- - - - - - - - - - -
											<%
											Response.Write oFunc.MakeList("TRUE,FALSE","Yes,No", oFunc.TFText(bolActive_Military1))
										%>
									</select>
								</td>
								<td>
									<input type="text" name="szRank1" value="<%= szRank1%>" maxlength="50" size="4"   ID="Text10">
								</td>
								<td>
									<input type="text" name="szPager1" value="<%= szPager1%>" maxlength="15" size="15"   ID="Text11">
								</td>
							</tr>
						</table>
						<table ID="Table6">
							<tr>
								<td class="gray">
									&nbsp;Home Phone*&nbsp;
								</td>
								<td class="gray">
									&nbsp;Business Phone&nbsp;
								</td>
								<td class="gray">
									&nbsp;Ext.
								</td>
								<td class="gray">
									&nbsp;Cell Phone
								</td>
							</tr>
							<tr>
								<td align="center">
									<input type="text" name="szHome_Phone1" value="<%= szHome_Phone1%>" maxlength="15" size="15"   ID="Text3">
								</td>
								<td align="center">
									<input type="text" name="szBusiness_Phone1" value="<%= szBusiness_Phone1%>" maxlength="15" size="15"   ID="Text12">
								</td>
								<td>
									<input type="text" name="intPhone_Ext1" value="<%= intPhone_Ext1%>" maxlength="4" size="4"   ID="Text13">
								</td>
								<td>
									<input type="text" name="szCell_Phone1" value="<%= szCell_Phone1%>" maxlength="15" size="15"   ID="Text14">
								</td>
							</tr>
						</table>
						<table ID="Table7">
							<tr>
								<td class="gray">
									&nbsp;Address*
								</td>
								<td class="gray">
									&nbsp;City*
								</td>
								<td class="gray">
									&nbsp;State*
								</td>
								<Td class="gray">
									&nbsp;Country*
								</Td>
								<Td class="gray">
									&nbsp;Zip*
								</Td>
							</tr>
							<tr>
								<td>
									<input type="text" name="szAddress1" value="<%= szAddress1%>" maxlength="256" size="30"   ID="Text16">
								</td>
								<td>
									<input type="text" name="szCity1" value="<%= szCity1%>" maxlength="128" size="10"   ID="Text17">
								</td>
								<td>
									<select name="szState1" ID="Select3">
										<option value="">
											<%
								'Create State select list									
								sql = "select strValue, strText " & _
										"from Common_Lists " & _
										"where intList_ID=3 order by strValue "
								' Set Alaska as default state
								if szState1 = "" then szState1 = "AK"
								response.Write oFunc.MakeListSQL(sql,"strValue","strText",szState1)
								if szCountry1 = "" then szCountry1 = "USA"
							%>
									</select>
								</td>
								<td>
									<input type="text" name="szCountry1" value="<%= szCountry1%>" maxlength="50" size="7"   ID="Text18">
								</td>
								<td>
									<input type="text" name="szZip_Code1" value="<%= szZip_Code1%>" maxlength="12" size="5"   ID="Text19">
								</td>
							</tr>
						</table>
						<br>
						<table ID="Table8">
							<tr>
								<Td colspan="6">
									<font class="svplain10"><b><i>Parent/Guardian #2</i></b> </font>
								</Td>
							</tr>
							<tr>
								<td class="gray">
									&nbsp; Last Name*
								</td>
								<td class="gray">
									&nbsp;First Name*
								</td>
								<td class="gray">
									&nbsp;MI
								</td>
								<td class="gray">
									&nbsp;Email Address&nbsp;
								</td>
							</tr>
							<tr>
								<td>
									<input type="text" name="szLast_Name2" value="<%= szLast_Name2%>" maxlength="50" size="17"   ID="Text15">
								</td>
								<td>
									<input type="text" name="szFirst_Name2" value="<%= szFirst_Name2%>" maxlength="50" size="15"   ID="Text20">
								</td>
								<td>
									<input type="text" name="sMid_Initial2" value="<%= sMid_Initial2%>" maxlength="1" size="2"   ID="Text21">
								</td>
								<td class="svplain10">
									<input type=text name="szEmail2" size=25 value="<%= szEmail2%>" maxlength=128 ID="Text4">
								</td>
							</tr>
						</table>
						<table ID="Table9">
							<tr>
								<Td class="gray">
									&nbsp;Employer
								</Td>
								<td class="gray">
									&nbsp;Active Military&nbsp;
								</td>
								<td class="gray">
									&nbsp;Rank&nbsp;
								</td>
								<td class="gray">
									&nbsp;Pager
								</td>
							</tr>
							<tr>
								<td>
									<input type="text" name="szEmployer2" value="<%= szEmployer2%>" maxlength="128" size="30"   ID="Text22">
								</td>
								<td>
									<select name="bolActive_Military2" ID="Select4">
										<option value="">- - - - - - - - - - -
											<%
											Response.Write oFunc.MakeList("TRUE,FALSE","Yes,No", oFunc.TFText(bolActive_Military2))
										%>
									</select>
								</td>
								<td>
									<input type="text" name="szRank2" value="<%= szRank2%>" maxlength="20" size="4"   ID="Text23">
								</td>
								<td>
									<input type="text" name="szPager2" value="<%= szPager2%>" maxlength="15" size="15"   ID="Text24">
								</td>
							</tr>
						</table>
						<table ID="Table10">
							<tr>
								<td class="gray">
									&nbsp;Home Phone&nbsp;
								</td>
								<td class="gray">
									&nbsp;Business Phone&nbsp;
								</td>
								<td class="gray">
									&nbsp;Ext.
								</td>
								<td class="gray">
									&nbsp;Cell Phone
								</td>
							</tr>
							<tr>
								<td align="center">
									<input type="text" name="szHome_Phone2" value="<%= szHome_Phone2%>" maxlength="15" size="15"   ID="Text44">
								</td>
								<td align="center">
									<input type="text" name="szBusiness_Phone2" value="<%= szBusiness_Phone2%>" maxlength="15" size="15"   ID="Text25">
								</td>
								<td>
									<input type="text" name="intPhone_Ext2" value="<%= intPhone_Ext2%>" maxlength="4" size="4"   ID="Text26">
								</td>
								<td>
									<input type="text" name="szCell_Phone2" value="<%= szCell_Phone2%>" maxlength="15" size="15"   ID="Text27">
								</td>
							</tr>
						</table>
						<table ID="Table11">
							<tr>
								<td class="gray">
									&nbsp;Address (if different)
								</td>
								<td class="gray">
									&nbsp;City
								</td>
								<td class="gray">
									&nbsp;State
								</td>
								<Td class="gray">
									&nbsp;Country
								</Td>
								<Td class="gray">
									&nbsp;Zip
								</Td>
							</tr>
							<tr>
								<td>
									<input type="text" name="szAddress2" value="<%= szAddress2%>" maxlength="256" size="30"   ID="Text28">
								</td>
								<td>
									<input type="text" name="szCity2" value="<%= szCity2%>" maxlength="50" size="10"   ID="Text29">
								</td>
								<td>
									<select name="szState2" ID="Select5">
										<option value="">
											<%
								'Create State select list									
								sql = "select strValue, strText " & _
										"from Common_Lists " & _
										"where intList_ID=3 order by strValue "
								' Set Alaska as default state
								if szState2 = "" then szState2 = "AK"
								response.Write oFunc.MakeListSQL(sql,"strValue","strText",szState2)
								
								if szCountry2 = "" then szCountry2 = "USA"
							%>
									</select>
								</td>
								<td>
									<input type="text" name="szCountry2" value="<%= szCountry2%>" maxlength="25" size="7"   ID="Text30">
								</td>
								<td>
									<input type="text" name="szZip_Code2" value="<%= szZip_Code2%>" maxlength="10" size="5"   ID="Text31">
								</td>
							</tr>
						</table>
					</td>
				</tr>
		</table>
		<% end if ' ends test for bolAdditionalGuardian %>
		<table ID="Table12">
			<tr>
				<td colspan="2" class="svplain8">
					<!--<input type="button" value="Cancel" onclick="window.location.href='default.htm';">-->
					<input type="submit" value="Save Guardian Info" name="saveGuardian">
					<% if bolShowFinish and bolAdditionalGuardian = "" then %>
					<input type="submit" value="Finished" name="Finished" ID="Submit1">
					<% else %>
					<input type="button" value="Cancel" onclick="window.location.href='lotteryMain.asp';" ID="Button1" NAME="Button1">
					<% end if %>
					<br>
					<br>
					&nbsp; (* Required Fields. Information will NOT be saved if you do not click 
					the 'Save Gaurdian Info.' button.)
					<br>
					<br>
				</td>
			</tr>
			</form>
		</table>
		<% 
call oFunc.CloseCN()
set oFunc = nothing
set oVal = nothing
Server.Execute(Application.Value("strWebRoot") & "Includes/footer.asp")

function vbfValidateGuardian(num)
	'Dynamically assign guardian values based on 'num'
	'which tells us which guardian we are validating.
	'if num is 1 then this is the Primary guardian
	
	dim strError	' Used to store any returned errors
	session.Contents("strError") = "" ' resets error message
	
	if num <> "" then 
		execute("szLast_Name = szLast_Name" & num)
		execute("szFirst_Name = szFirst_Name" & num)
		execute("sMid_Initial = sMid_Initial" & num)
		execute("szEmail = szEmail" & num)
		execute("szHome_Phone = szHome_Phone" & num)
		execute("szAddress = szAddress" & num)
		execute("szCity = szCity" & num)
		execute("szState = szState" & num)
		execute("szCountry = szCountry" & num)
		execute("szZip_Code = szZip_Code" & num)
	end if
	
	' Now do the validation
	oVal.validateField szLast_Name,"blank","","Last Name" 
	oVal.validateField szFirst_Name,"blank","","First Name"
	'oVal.validateField sMid_Initial,"blank","","Middle Initial"
		
	if num = "1" then
		oVal.validateField szHome_Phone,"blank","","Home Phone"
	end if
	
	if num = "1" or szEmail <> "" then
		oVal.validateField szEmail,"email","","Email Address"
	end if 
	
	if num = "1" or szAddress <> "" then
		oVal.validateField szAddress,"blank","","Address" 
		oVal.validateField szCity,"blank","","City" 
		oVal.validateField szState,"blank","","State" 
		oVal.validateField szCountry,"blank","","Country" 
		oVal.validateField szZip_Code,"blank","","Zip Code" 
	end if
	
	if num = "1" then 
		num = " the Primary guardian"
	elseif num = "" then
		num = " the guardian below"
	else
		num = "guardian #" & num
	end if
	if oVal.ValidationError & "" <> "" then
		strError = "<BR><font color=red><b>The following items need to be corrected for " & num & ".</B><BR>"
		strError = strError & oVal.ValidationError & "</font>"
	end if
	
	session.Contents("strError") = strError ' this is so we have access to strError in lotteryAdditionalGuardian.asp
	vbfValidateGuardian = strError	
end function

sub vbsInsertGuardian(num)
	' Creates a family,guardian and associated records . 'num' tells us which set of
	' guardian variables to use
	
	' We dynamically assign variables based on the set of variables
	' for our 'num' guardian
	execute("szLast_Name = szLast_Name" & num)
	execute("szFirst_Name = szFirst_Name" & num)
	execute("sMid_Initial = sMid_Initial" & num)
	execute("szEmail = szEmail" & num)
	execute("szHome_Phone = szHome_Phone" & num)
	execute("szAddress = szAddress" & num)
	execute("szCity = szCity" & num)
	execute("szState = szState" & num)
	execute("szCountry = szCountry" & num)
	execute("szZip_Code = szZip_Code" & num)
	
	execute("szEmployer = szEmployer" & num)
	execute("bolActive_Military = bolActive_Military" & num)
	execute("szRank = szRank" & num)
	execute("szPager = szPager" & num)
	execute("szBusiness_Phone = szBusiness_Phone" & num)
	execute("intPhone_Ext = intPhone_Ext" & num)
	execute("szCell_Phone = szCell_Phone" & num)
	
	oFunc.BeginTransCN
	if session.Contents("intFamily_ID") = "" then				
		'create family record.	
		insert = "insert into tblFamily(" & _
				 " szFamily_Name, szAddress, szCity, szState, szCountry, szZip_Code, " & _
				 "szHome_Phone, szEMAIL, bolLottery, dtCREATE) " & _
				 "values (" & _
				 "'" & oFunc.EscapeTick(szLast_Name) & "'," & _
				 "'" & oFunc.EscapeTick(szAddress) & "'," & _
				 "'" & oFunc.EscapeTick(szCity) & "'," & _
				 "'" & oFunc.EscapeTick(szState) & "'," & _
				 "'" & oFunc.EscapeTick(szCountry) & "'," & _
				 "'" & oFunc.EscapeTick(szZip_Code) & "'," & _
				 "'" & oFunc.EscapeTick(szHome_Phone) & "'," & _
				 "'" & oFunc.EscapeTick(szEmail) & "'," & _
				 "1," & _
				 "'" & now() & "')"
		oFunc.ExecuteCN(insert)
		session.Contents("intFamily_ID") = oFunc.GetIdentity
	end if
	
	' Determine if this is a primary guardian
	if num = "1" then
		intPrimary = 1
	else
		intPrimary = 0
	end if
		
	'create guardian record
	insert = "insert into tblGuardian(" & _
			 "szFIRST_NAME, szLAST_NAME, sMID_INITIAL, szEMPLOYER, szBUSINESS_PHONE, " & _
			 "intPHONE_EXT, szCELL_PHONE, szPAGER,bolACTIVE_MILITARY, szRANK, " & _
			 "szAddress, szCity, szState, szCountry, szZip_Code, szHome_Phone, " & _
			 "bolSnail_Mail,szEmail, bolLottery, bolPrimary, dtCREATE) " & _
			 " values (" & _
			 "'" & oFunc.EscapeTick(szFirst_Name) & "'," & _
			 "'" & oFunc.EscapeTick(szLast_Name) & "'," & _
			 "'" & oFunc.EscapeTick(sMID_INITIAL) & "'," & _
			 "'" & oFunc.EscapeTick(szEMPLOYER) & "'," & _
			 "'" & oFunc.EscapeTick(szBUSINESS_PHONE) & "'," & _
			 "'" & oFunc.CheckDecimal(intPHONE_EXT) & "'," & _
			 "'" & oFunc.EscapeTick(szCELL_PHONE) & "'," & _
			 "'" & oFunc.EscapeTick(szPAGER) & "'," & _
			 "'" & oFunc.TrueFalse(bolACTIVE_MILITARY) & "'," & _
			 "'" & oFunc.EscapeTick(szRANK) & "'," & _
			 "'" & oFunc.EscapeTick(szAddress) & "'," & _
			 "'" & oFunc.EscapeTick(szCity) & "'," & _
			 "'" & oFunc.EscapeTick(szState) & "'," & _
			 "'" & oFunc.EscapeTick(szCountry) & "'," & _
			 "'" & oFunc.EscapeTick(szZip_Code) & "'," & _
			 "'" & oFunc.EscapeTick(szHome_Phone) & "'," & _
			 "'" & oFunc.TrueFalse(bolSnail_Mail) & "'," & _
			 "'" & oFunc.EscapeTick(szEmail) & "'," & _
			 "1," & _
			 intPrimary & "," & _
			 "'" & now() & "')" 
	oFunc.ExecuteCN(insert)
	session.Contents("intGuardian_ID" & num) = oFunc.GetIdentity
	
	'Now create association between the Family and Guardian
	insert = "insert into tascFAM_GUARD (" & _
			 "intFamily_ID, intGUARDIAN_ID, dtCREATE)" & _
			 "values (" & _ 
			 session.Contents("intFamily_ID") & "," & _
			 session.Contents("intGUARDIAN_ID" & num) & "," & _
			 "'" & now() & "')"
	oFunc.ExecuteCN(insert)				 
	oFunc.CommitTransCN	
	session.Contents("strInstructions") = "Your guardian information has been saved."	
	bolAdditionalGuardian = "" 	
	intGuardian_ID = ""
end sub

sub vbsUpdateGuardian(num)
	' Updates a guardian record. 'num' tells us which set of
	' guardian variable to use
	
	' We dynamically assign variables based on the set of variables
	' for our 'num' guardian
	execute("szLast_Name = szLast_Name" & num)
	execute("szFirst_Name = szFirst_Name" & num)
	execute("sMid_Initial = sMid_Initial" & num)
	execute("szEmail = szEmail" & num)
	execute("szHome_Phone = szHome_Phone" & num)
	execute("szAddress = szAddress" & num)
	execute("szCity = szCity" & num)
	execute("szState = szState" & num)
	execute("szCountry = szCountry" & num)
	execute("szZip_Code = szZip_Code" & num)
	
	execute("szEmployer = szEmployer" & num)
	execute("bolActive_Military = bolActive_Military" & num)
	execute("szRank = szRank" & num)
	execute("szPager = szPager" & num)
	execute("szBusiness_Phone = szBusiness_Phone" & num)
	execute("intPhone_Ext = intPhone_Ext" & num)
	execute("szCell_Phone = szCell_Phone" & num)
	
	update = "update tblGuardian set " & _
			 "szLast_Name = '" & oFunc.EscapeTick(szLast_Name) & "'," & _
			 "szFirst_Name = '" & oFunc.EscapeTick(szFirst_Name) & "'," & _
			 "sMid_Initial = '" & oFunc.EscapeTick(sMid_Initial) & "'," & _
			 "szEmail = '" & oFunc.EscapeTick(szEmail) & "'," & _
			 "szHome_Phone = '" & oFunc.EscapeTick(szHome_Phone) & "'," & _
			 "szAddress = '" & oFunc.EscapeTick(szAddress) & "'," & _
			 "szCity = '" & oFunc.EscapeTick(szCity) & "'," & _
			 "szState = '" & oFunc.EscapeTick(szState) & "'," & _
			 "szCountry = '" & oFunc.EscapeTick(szCountry) & "'," & _
			 "szZip_Code = '" & oFunc.EscapeTick(szZip_Code) & "'," & _
			 "szEmployer = '" & oFunc.EscapeTick(szEmployer) & "'," & _
			 "bolActive_Military = " & oFunc.TrueFalse(bolActive_Military) & "," & _
			 "szRank = '" & oFunc.EscapeTick(szRank) & "'," & _
			 "szPager = '" & oFunc.EscapeTick(szPager) & "'," & _
			 "szBusiness_Phone = '" & oFunc.EscapeTick(szBusiness_Phone) & "'," & _
			 "intPhone_Ext = '" & oFunc.CheckDecimal(intPhone_Ext) & "'," & _
			 "szCell_Phone = '" & oFunc.EscapeTick(szCell_Phone) & "' " & _
			 "Where intGuardian_ID = " & session.Contents("intGuardian_ID" & num)
			 
	oFunc.ExecuteCN(update)
	
	session.Contents("strInstructions") = "Your guardian information has been saved."
	bolAdditionalGuardian = "" 	
	intGuardian_ID = ""
end sub

sub vbsAdditionalGuardianForm()
								  
%>
				<tr>
					<input type=hidden name="bolAdditionalGuardian" value="true" ID="Hidden1">
					<input type=hidden name="intGuardian_ID" value="<% = intGuardian_ID%>">
					<td class="NavyHeader">
						&nbsp;<B>Parent/Guardian Information</B> &nbsp;&nbsp;&nbsp;&nbsp;
					</td>
				</tr>
				<tr>
					<td>
						<table ID="Table15">
							<tr>
								<Td colspan="6">
									<font class="svplain10"><b><i>Additional Parent/Guardian</i></b> </font>
								</Td>
							</tr>
							<tr>
								<td class="gray">
									&nbsp; Last Name*
								</td>
								<td class="gray">
									&nbsp;First Name*
								</td>
								<td class="gray">
									&nbsp;MI
								</td>
								<td class="gray">
									&nbsp;Email Address&nbsp;
								</td>
							</tr>
							<tr>
								<td>
									<input type="text" name="szLast_Name" value="<%= szLast_Name%>" maxlength="50" size="17"   ID="Text5">
								</td>
								<td>
									<input type="text" name="szFirst_Name" value="<%= szFirst_Name%>" maxlength="50" size="15"   ID="Text6">
								</td>
								<td>
									<input type="text" name="sMid_Initial" value="<%= sMid_Initial%>" maxlength="1" size="2"   ID="Text7">
								</td>
								<td class="svplain10">
									<input type=text name="szEmail" size=25 value="<%= szEmail%>" maxlength=128 ID="Text32">
								</td>
							</tr>
						</table>
						<table ID="Table16">
							<tr>
								<Td class="gray">
									&nbsp;Employer
								</Td>
								<td class="gray">
									&nbsp;Active Military&nbsp;
								</td>
								<td class="gray">
									&nbsp;Rank&nbsp;
								</td>
								<td class="gray">
									&nbsp;Pager
								</td>
							</tr>
							<tr>
								<td>
									<input type="text" name="szEmployer" value="<%= szEmployer%>" maxlength="128" size="30" ID="Text33">
								</td>
								<td>
									<select name="bolActive_Military" ID="Select1">
										<option value="">- - - - - - - - - - -
											<%
											Response.Write oFunc.MakeList("TRUE,FALSE","Yes,No", oFunc.TFText(bolActive_Military2))
										%>
									</select>
								</td>
								<td>
									<input type="text" name="szRank" value="<%= szRank%>" maxlength="20" size="4"   ID="Text34">
								</td>
								<td>
									<input type="text" name="szPager" value="<%= szPager%>" maxlength="15" size="15"   ID="Text35">
								</td>
							</tr>
						</table>
						<table ID="Table17">
							<tr>
								<td class="gray">
									&nbsp;Home Phone&nbsp;
								</td>
								<td class="gray">
									&nbsp;Business Phone&nbsp;
								</td>
								<td class="gray">
									&nbsp;Ext.
								</td>
								<td class="gray">
									&nbsp;Cell Phone
								</td>
							</tr>
							<tr>
								<td align="center">
									<input type="text" name="szHome_Phone" value="<%= szHome_Phone%>" maxlength="15" size="15"   ID="Text36">
								</td>
								<td align="center">
									<input type="text" name="szBusiness_Phone" value="<%= szBusiness_Phone%>" maxlength="15" size="15"   ID="Text37">
								</td>
								<td>
									<input type="text" name="intPhone_Ext" value="<%= intPhone_Ext%>" maxlength="4" size="4"   ID="Text38">
								</td>
								<td>
									<input type="text" name="szCell_Phone" value="<%= szCell_Phone%>" maxlength="15" size="15"   ID="Text39">
								</td>
							</tr>
						</table>
						<table ID="Table18">
							<tr>
								<td class="gray">
									&nbsp;Address (if different)
								</td>
								<td class="gray">
									&nbsp;City
								</td>
								<td class="gray">
									&nbsp;State
								</td>
								<Td class="gray">
									&nbsp;Country
								</Td>
								<Td class="gray">
									&nbsp;Zip
								</Td>
							</tr>
							<tr>
								<td>
									<input type="text" name="szAddress" value="<%= szAddress%>" maxlength="256" size="30"   ID="Text40">
								</td>
								<td>
									<input type="text" name="szCity" value="<%= szCity%>" maxlength="50" size="10"   ID="Text41">
								</td>
								<td>
									<select name="szState" ID="Select6">
										<option value="">
											<%
								'Create State select list									
								sql = "select strValue, strText " & _
										"from Common_Lists " & _
										"where intList_ID=3 order by strValue "
								' Set Alaska as default state
								if szState = "" then szState = "AK"
								response.Write oFunc.MakeListSQL(sql,"strValue","strText",szState)
								
								if szCountry = "" then szCountry = "USA"
							%>
									</select>
								</td>
								<td>
									<input type="text" name="szCountry" value="<%= szCountry%>" maxlength="25" size="7"   ID="Text42">
								</td>
								<td>
									<input type="text" name="szZip_Code" value="<%= szZip_Code%>" maxlength="10" size="5"   ID="Text43">
								</td>
							</tr>
						</table>
					</td>
				</tr>
<%
	session.Contents("strInstructions") = " "
end sub

sub vbsConfirmDelete 
%>
<html>
<head>
<title>Confirm Guardian Deletion</title>
<link rel="stylesheet" type="text/css" href="../../css/homestyle.css">
</head>
		
<body bgcolor=white>
<table width=100% height=100% ID="Table14">
	<tr>
		<td class=svplain11 valign=middle align=center>
			<b>Are you sure you want to delete guardian 
			'<% = request.QueryString("guardianName")%>'?</b><br><br>
			<input type=button value="Cancel" onClick="window.location.href='lotteryMain.asp';" NAME="Button1" ID="Button2">	
			<input type=button value="Yes, Delete Guardian." onclick="window.location.href='lotteryMain.asp?intGuardian_ID=<%=intGuardian_ID%>&bolConfirm=true';" ID="Button3" NAME="Button3">						
		</td>
	</tr>
</table>

</body>
</html>
<%
	response.End
end sub

sub vbsDeleteGuardian()
	' Deletes guardian and tascFam_Guard entry. tascFam_Guard record is deleted by
	' cascading delete.
	
	'intFamily_ID is added to where clause as a precaution against
	'malicous header info being submitted by a hacker
	set rsCheck = server.CreateObject("ADODB.RECORDSET")
	rsCheck.CursorLocation = 3
	
	sql = "select * from tascFam_Guard where intFamily_ID = " & _
		  session.Contents("intFamily_ID") & _
		  " and intGuardian_ID = " & intGuardian_ID
	
	rsCheck.Open sql, oFunc.FPCScnn
	
	if rsCheck.RecordCount > 0 then	  
		delete = "delete from tblGuardian " & _
				"where intGuardian_ID = " & intGuardian_ID 
		oFunc.ExecuteCN(delete)
	end if
	
	rsCheck.Close
	set rsCheck = nothing
	intGuardian_ID = "" 
	session.Contents("strInstructions") = "Guardian has been deleted."
end sub
%>