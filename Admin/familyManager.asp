<%@ Language=VBScript %>
<%
'*******************************************
'Name:		Admin\familyManager.asp
'Purpose:	Allows FPCS staff create edit family profiles
'
'CalledBy:	
'
'Inputs:	Request.QueryString("szUserID")
'
'Author:	ThreeShapes.com LLC
'Date:		20 May 2002
'*******************************************

'per http://support.microsoft.com/default.aspx?scid=kb;EN-US;q234067
Response.CacheControl = "no-cache" 
Response.Expires = -1

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

oFunc.ResetSelectSessionVariables 

dim intFamily_Id
dim strButton
dim blnIncludeDir
dim strStudentName		'contains list of students with 'View/Edit' button for guardians
dim strStudentOptions
dim strStudentNameAdmin	'contains list of students with links to view edit student info in admin mode
dim strSISClass
dim strIEPClass
dim strURL
dim strGuardianOptions
dim intGuardian_ID
dim strJFAction

strButton = "Add Family"

' Define Family ID 
if Request("intFamily_ID") <> "" and session.Contents("strRole") = "ADMIN" then
	intFamily_Id = Request("intFamily_ID")
elseif session.Value("intFamily_id") <> "" then
	intFamily_Id = session.Value("intFamily_id")   ' Only defined in guardian log on
elseif request("intStudent_ID") <> "" then
	intFamily_Id = oFunc.StudentInfo(request("intStudent_ID"),6)
end if

'response.write Request("intFamily_ID") & " - " & session.Value("intFamily_id") & " - " & intFamily_Id

if session.Value("strRole") = "ADMIN" or session.Value("strRole") = "GUARD" then
	if Request.Form("intFamily_ID") = "new" and session.Value("strRole") = "ADMIN" then
		call vbsInsertFamily		
	elseif Request.Form("changed") <> "" and Request.Form("intFamily_ID") <> "new" then
		call vbsUpdateFamily
	end if
else
	' The user requesting this page does not have rights to view it
	response.Write "<h1>Page Improperly Called.</h1>"
	response.End
end if	

if (intFamily_ID <> "" and ucase(intFamily_ID) <> "NEW") OR (session.Contents("strRole") <> "GUARD" and intFamily_ID <> "" and ucase(intFamily_ID) <> "NEW")then
	
	'Populate the family info.
	dim sql
	dim intCount
	
	set rsFamily = server.CreateObject("ADODB.RECORDSET")
	rsFamily.CursorLocation = 3
	
	sql = "select szFamily_Name,szDesc,szAddress,szCity,szState," & _
			"szCountry,szZip_Code,szHome_Phone,szEMAIL,bolIncludeDir " & _
			"from tblFamily " & _
			"where intFamily_ID = " & intFamily_Id 
			
	rsFamily.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
	
	if rsFamily.RecordCount > 0 then
		intCount = 0
		'This for loop dimentions and defines all the columns we selected in sqlClass
		'and we use the variables created here to populate the form.
		for each item in rsFamily.Fields
			execute("dim " & rsFamily.Fields(intCount).Name)
			execute(rsFamily.Fields(intCount).Name & " = item")		
			intCount = intCount + 1
		next  
		
		if bolIncludeDir then
			bolIncludeDir = " checked"
		else
			bolIncludeDir = " "
		end if
	end if 
	
	strButton = "Update"
	if ucase(session.Contents("strRole")) <> "GUARD" then
		strJFAction = " onClick=""jfMakeList(main.selStudent_ID);jfMakeList(main.selGuardian_ID);jfChanged();"""
	end if
	
	' Get this families students
	dim strStudentList
	set rsStudents = server.CreateObject("ADODB.RECORDSET")
	rsStudents.CursorLocation = 3
	
	if session.Value("strRole") = "ADMIN" then
		strMore = " + ': ' + convert(nchar(5),intStudent_ID) "
	else
		strMore = ""
	end if 		
	
	sqlStudents = "SELECT s.intSTUDENT_ID, s.szFIRST_NAME + ' ' + s.szLAST_NAME AS name, " & _
					"ei.intENROLL_INFO_ID, ss.intStudent_State_id " & _
					", ss.intStudent_State_id, SUM(tblIEP.intIEP_ID) AS bolIEP " & _
					"FROM tblSTUDENT s LEFT OUTER JOIN " & _
					" tblIEP ON s.intSTUDENT_ID = tblIEP.intStudent_ID AND " & _
					" tblIEP.intSchool_Year = " & session.Contents("intSchool_Year") & _
					" LEFT OUTER JOIN " & _
					" tblStudent_States ss ON s.intSTUDENT_ID = ss.intStudent_id " & _
					" AND ss.intSchool_Year = " & session.Contents("intSchool_Year") & _
					" AND ss.intReEnroll_State IN (" & Application.Contents("ActiveEnrollList") & ") LEFT OUTER JOIN " & _
					" tblENROLL_INFO ei ON s.intSTUDENT_ID = ei.intSTUDENT_ID " & _
					" AND ei.sintSCHOOL_YEAR = " & session.Contents("intSchool_Year") & " " & _
					"WHERE (s.intFamily_ID = " & intFamily_Id & ") " & _
					"GROUP BY s.intSTUDENT_ID, s.szFIRST_NAME, s.szLAST_NAME, " & _
					"ei.intENROLL_INFO_ID, ss.intStudent_State_id"
					
	rsStudents.Open sqlStudents,Application("cnnFPCS")'oFunc.FPCScnn
	
	do while not rsStudents.EOF
		strStudentList = strStudentList & rsStudents("intStudent_ID") & ", "
		if session.Value("strRole") = "ADMIN" then
			strStudentName = strStudentName & _
							"<a href='javascript:' onClick=""jfViewProfile('" & rsStudents("intStudent_ID") & _
							"','studentProfile.asp');"" >" & rsStudents("Name") & ": " & rsStudents("intStudent_ID") & "</a><BR>"
			strStudentOptions = strStudentOptions & "<option value=""" & rsStudents("intStudent_ID") & """>" & 	rsStudents("Name") & ": " & rsStudents("intStudent_ID") & "</option>"		
			intStudent_ID = intStudent_ID & rsStudents("intStudent_ID") & ","
		else
			if (rsStudents("intENROLL_INFO_ID") & "" = "" or rsStudents("bolIEP") & "" = "") _ 
			and rsStudents("intStudent_State_id")&"" <> ""  then
				strInstructions = "<B>Instructions:</b><BR>Each school year the SIS and IEP inforamtion " & _
								  "must be updated for each student. Next to the students name please find " & _
								  "buttons labeled 'SIS Info' and 'IEP Info'. If either of these buttons are red " & _
								  "that indicates that the inforomation has not been updated. To update this " & _
								  "information simply click on the red button and complete the form that it displays."
			end if
			
			if rsStudents("intStudent_State_id") & "" <> "" then
				strURL = application.Contents("strSSLWebRoot") & "forms/SIS/iep.asp?isFamManager=true&intStudent_ID=" & rsStudents("intStudent_ID")
				if rsStudents("intENROLL_INFO_ID") & "" = "" then
					strSISClass = "btSmallRed"
				else
					strSISClass = "btSmallGray"
				end if
				
				if rsStudents("bolIEP") & "" = "" then
					strIEPClass = "btSmallRed"
				else
					strIEPClass = "btSmallGray"
				end if
				
				strStudentName = strStudentName & "<tr><td class=gray>" & rsStudents("Name") & "</td>" & _
								"<td><input class='btsmallgray' type=button value='SIS Info' onclick=""jfViewProfile('" & rsStudents("intStudent_ID") & _
								"','studentProfile.asp');"" id=" & strSISClass & ">" & _
								"<input class='btsmallgray' type=button value='IEP Info' onClick=""window.location='" & _
								strURL & "';"" id=" & strIEPClass & "></td></tr>"
			else
				strStudentName = strStudentName & "<tr class=gray><td>" & rsStudents("Name") & "</td>" & _
								"<td align=center>not enrolled</td></tr>"
			end if
		end if
		rsStudents.MoveNext
	loop
	rsStudents.Close
	set rsStudents = nothing
	
	if len(strStudentList) > 0 then
		strStudentList = left(strStudentList,len(strStudentList)-2)
	end if
		
	' Get this families Guardians
	dim strGuardianName
	dim strGuardianList
	set rsGuardian = server.CreateObject("ADODB.RECORDSET")
	rsGuardian.CursorLocation = 3
	
	if session.Value("strRole") = "ADMIN" then
		strMore = " + ': ' + convert(nchar(5),fg.intGuardian_ID) "
	else
		strMore = ""
	end if 
	
	sqlGuardian = "select fg.intGuardian_ID, g.szFirst_Name + ' ' + g.szLast_Name " & strMore & " as Name " & _
					"from tblGuardian g, tascFam_Guard fg " & _ 
					"where fg.intFamily_ID = " & intFamily_Id & _
					" and fg.intGuardian_ID = g.intGuardian_ID "  
	rsGuardian.Open sqlGuardian,Application("cnnFPCS")'oFunc.FPCScnn
	
	do while not rsGuardian.EOF
		strGuardianList = strGuardianList & rsGuardian("intGuardian_ID") & ", "
		if session.Contents("strRole") = "ADMIN" then
			strGuardianName = strGuardianName & _
							"<a href='javascript:' onClick=""jfViewProfile('" & rsGuardian("intGuardian_ID") & _
							"','guardianProfile.asp');"" >" & rsGuardian("Name") & "</a><BR>"
			strGuardianOptions = strGuardianOptions & "<option value=""" & rsGuardian("intGuardian_ID") & """>" & 	rsGuardian("Name") & "</option>"		
			intGuardian_ID = intGuardian_ID & rsGuardian("intGuardian_ID") & ","							
		else	
			strGuardianName = strGuardianName & "<tr><td class=gray>" & rsGuardian("Name") & "</td>" & _
							"<td><input type=button value='View/Edit Info' onclick=""jfViewProfile('" & rsGuardian("intGuardian_ID") & _
							"','guardianProfile.asp');"" class=btsmallgray></td></tr>"
		end if
		rsGuardian.MoveNext
	loop
	rsGuardian.Close
	set rsGuardian = nothing
	if len(strGuardianList) > 0 then
		strGuardianList = left(strGuardianList,len(strGuardianList)-2)
	end if 
end if 
	
Session.Value("strTitle") = "Family Manager"
Server.Execute(Application.Value("strWebRoot") & "Includes/simpleheader.asp")
%>
<script language=javascript>
	function jfGetFamily(item){
		window.location.href = "familyManager.asp?intFamily_ID=" + item.value;
	}
	function jfAddNewStudent() {
		var winStudent;
		winStudent = window.open("../forms/sis/studentProfile.asp?bolNewStudent=yes","winStudent","scrollbars=yes,width=640,height=500,resizable=yes");
		winStudent.moveTo(0,0);
		winStudent.focus();
	}
	
	function jfAddNewGuardian() {
		var winGuardian;
		var intLen = document.main.intStudent_ID.length;
		var strStudents = "";
		for(i=0;i<intLen;i++){
			//if (document.main.intStudent_ID.options[i].selected == true){
				strStudents += document.main.intStudent_ID.value;
			//}		
		}
		
		winGuardian = window.open("../forms/sis/guardianProfile.asp?bolNewGuardian=yes&strStudents="+strStudents,"winGuardian","scrollbars=yes,width=640,height=500,resizable=yes");
		winGuardian.moveTo(0,0);
		winGuardian.focus();
	}
	function jfAddOption(text,val,list) {
		// called by html page created in studentInsert.asp or guardianInsert.asp. 
		// This adds a newly created student or guardian to the bottom of the drop down list.
		var obj = new Option(text,val);
		var objList = document.getElementById(list);
		var intLen =  objList.length;
		objList.options[intLen] = obj;
		jfMakeList(objList)	
		jfChanged();	
		// The next line does select the item in the list BUT does NOT include it in the HTTP header
		//document.all.item(list).options[intLen].selected = true;
	}
	
	<% if strMessage <> "" then response.write "alert('" & strMessage & "');" %>
	
	function jfViewProfile(id,script){
		var winProfile;
		var strParam;
		if (script == "guardianProfile.asp") {
			strParam = "intFamily_ID=<%=intFamily_Id %>&intGuardian_id=" + id + "&strStudents=";
			<% if session.Value("strRole") = "GUARD" then 
				response.write  "strParam += """ & replace(strStudentList," ","") & """;"
			   else %>
			//var intLen = document.main.intStudent_ID.length;
			//for(i=0;i<intLen;i++){
			//	if (document.main.intStudent_ID.options[i].selected == true){
					strParam += document.main.intStudent_ID.value;
			//	}		
			//}
			<% end if %>
		}else{
			strParam = "intStudent_id=" + id;
		}
		winProfile = window.open("../forms/sis/" + script +"?bolUpdate=yes&exempt=true&"+strParam,"winProfile","scrollbars=yes,width=640,height=500,resizable=yes");
		winProfile.moveTo(0,0);
		winProfile.focus();
	}
	function jfCheckForm(){
		if (main.changed.value != "" ) {
			var bolAnswer = confirm("You have made changes to this form. Click 'OK' to save changes or \nclick 'Cancel' to discard changes and return to the home page.");
			if (bolAnswer){
				jfMakeList(main.selStudent_ID);
				jfMakeList(main.selGuardian_ID);
				main.submit();
			}else{
				window.location.href='<% = Application.Value("strWebRoot")%>';
			}			
		}else{
			window.location.href='<% = Application.Value("strWebRoot")%>';
		}
	}
</script>
<form name=main method=post action="familyManager.asp" <% if Request.QueryString("bolForced") <> "" then Response.Write "onSubmit='return false';"%>>
<input type=hidden name="changed" value="">
<input type=hidden name="bolLottery" value="<% request("bolLottery")%>">
<table width="100%">
	<tr>
		<Td class="yellowHeader">
			&nbsp;<b>Family <% if request("bolLottery") <> "" then response.Write" Lottery " %>Manager</b>
		</Td>
	</tr>
	<tr>
		<td bgcolor="f7f7f7">
			<table>
				<tr>
					<td>
						<table>
							<tr>
								<% if session.Value("strRole") = "ADMIN" then %>
								<td class=Gray>								
									<nobr>&nbsp;<b>Select a Family</nobr><br>
								</td>
								<td width=100%>
									<select name="intFamily_ID" onChange="jfGetFamily(this);">
										<option value="new">New Family						
									<%
										dim sqlFamilies
										if request("bolLottery") =  "" then
											sqlFamilies = "select intFamily_ID, Name = " & _
														"CASE " & _
														"WHEN szDesc is null then szFamily_Name + ': ' +  convert(varchar,intFamily_ID) " & _
														"WHEN szDesc is not null then szFamily_Name + ', ' + szDesc + ': ' +  convert(varchar,intFamily_ID) " & _
														"END " & _
														"from tblFamily where bolLottery is null or bolLottery = 0 order by Name"
										else
											sqlFamilies = "SELECT f.intFamily_ID, f.szFamily_Name + ': ' + CONVERT(varchar, f.intFamily_ID) AS Name " & _
															"FROM tblFAMILY f INNER JOIN " & _
															" tblSTUDENT s ON f.intFamily_ID = s.intFamily_ID " & _
															"WHERE (s.bolLottery = 1) " & _
															"ORDER BY Name"
										end if											
										Response.Write oFunc.MakeListSQL(sqlFamilies,"intFamily_ID","Name",intFamily_Id)	
									%>
									</select>
								</td>
								<% elseif session.Value("strRole") = "GUARD" then%>
								<Td class=Gray>
									&nbsp;<B>Family Name:</b> <% = szFamily_Name %>&nbsp;
								</td>
								<Td class=Gray>
									&nbsp;<B>Description:</b> <% = szDesc %>&nbsp;
								</td>
								<% end if %>
							</tr>
						</table>
						<% if session.Value("strRole") = "ADMIN" then%>
						<table>
							<tr>
								<td class=gray>
										&nbsp;Family Name
								</td>
								<td class=gray>
										&nbsp;Description
								</td>				
							</tr>
							<tr>
								<td valign=top>
									<input type=text name="szFamily_Name" value="<% = szFamily_Name %>" maxlength=100 size=40 onChange="jfChanged();">
								</td>
								<td>
									<textarea name="szDesc" cols=30 rows=1 wrap=virtual onChange="jfChanged();"><% = szDesc %></textarea>
								</td>
							</tr>
						</table>
						<% else %>
							<input type=hidden name="szFamily_Name" value="<% = szFamily_Name %>">
							<input type=hidden name="szDesc" value="<% = szDesc %>">
						<% end if %>
						<table>
							<tr>
								<td class=gray>
										&nbsp;Address
								</td>		
								<td class=gray>
										&nbsp;Home Phone
								</td>			
							</tr>
							<tr>
								<td>
									<input type=text name="szAddress" value="<% = szAddress%>" maxlength=256 size=60 onChange="jfChanged();">
								</td>
								<td>
									<input type=text name="szHome_Phone" value="<% = szHome_Phone%>" maxlength=50 size=13 onChange="jfChanged();">
								</td>
							</tr>
						</table>
						<table>
							<tr>
								<td class=gray>
										&nbsp;City 
								</td>			
								<td class=gray>
										&nbsp;State 
								</td>	
								<td class=gray>
										&nbsp;Country 
								</td>
								<td class=gray>
										&nbsp;Zip 
								</td>
								<TD class=gray rowspan=2>&nbsp;Add Family<BR>&nbsp;to Directory&nbsp;</TD>
								<TD align="right" rowspan=2 valign=middle class=svplain10>Yes&nbsp;<INPUT type="checkbox" name="bolIncludeDir" <% = bolIncludeDir %> onChange="jfChanged();" ID="Checkbox1"></TD>
							</tr>
							<tr>
								<td>
									<input type=text name="szCity" value="<% = szCity %>" maxlength=128 size=20 onChange="jfChanged();">
								</td>
								<td>
									<select name="szState" onChange="jfChanged();">
									<%
										dim sqlState
										sqlState = "select strValue,strText from Common_Lists where intList_Id = 3 order by strValue"
										Response.Write oFunc.MakeListSQL(sqlState,"","",szState)
									%>
									</select>						
								</td>
								<td>
									<input type=text name="szCountry" value="<% = szCountry %>" maxlength=50 size=7 onChange="jfChanged();">
								</td>
								<td>
									<input type=text name="szZip_Code" value="<% = szZip_Code %>" maxlength=12 size=5 onChange="jfChanged();">
								</td>									
							</tr>
							<tr>
								<td class=gray colspan=4>
										&nbsp;Family Email Address 
								</td>
							</tr>
							<tr>
								<td colspan=4 class='svplain9'>
									<% if ucase(session.Contents("strRole")) = "ADMIN" then %>
									<input type=text name="szEMAIL" value="<% = szEMAIL %>" maxlength=127 size=40 onChange="jfChanged();">
									<% else %>
									<input type=hidden name="szEMAIL" value="<% = szEMAIL %>" >
									<% = szEMAIL %>
									<% end if %>
								</td>
							</tr>
						</table>
						<br>
						<table>							
						<% if intFamily_ID <> "" then
								if session.Value("strRole") = "ADMIN" then%>
							<tr>	
								<Td colspan=3>
									<font class=svplain11>
										<b><i>Students in Family</I></B> 
									</font>
								</td>
							</tr>
							<tr>
								<td class=gray>
										&nbsp;List of All Students  &nbsp; &nbsp;
										<input type=button value="add new student to list" onClick="jfAddNewStudent();" class="btSmallGray">
								</td>
								<td>
									&nbsp;
								</td>			
								<td class=gray>
										&nbsp;Students in this Family &nbsp;
								</td>
								<td>
									&nbsp;
								</td>
							<tr>
								<td>
									<select name="AllStudentsList" multiple size=5 onChange="jfChanged();">
									<%
										dim sqlStudents
										dim strMore
										if session.Value("strRole") = "ADMIN" then
											strMore = " + ': ' + convert(nchar(5),intStudent_ID) "
										else
											strMore = ""
										end if 
										sqlStudents = "select intStudent_ID, szLast_Name + ',' + szFirst_Name" & strMore & " as Name " & _
													  "from tblStudent  order by name"
										Response.Write oFunc.MakeListSQL(sqlStudents,"intStudent_ID","Name","")
									%>
									</select>
								</td>
								<td valign=middle>
									<input type=button value="Add >" class="btSmallGray" onclick="jfSelectItemFromTo(AllStudentsList, selStudent_ID);jfChanged();">
								</td>
								<td class=gray valign=top>
									<select name="selStudent_ID" id="selStudent_ID" multiple size=5 onChange="jfChanged();">
										<% = strStudentOptions %>
									</select>	
									<input type=hidden name="intStudent_ID" value="<% = intStudent_ID%>">								
								</td>
								<td valign=middle>
									<input type=button value="View Profile of Selected Student" class="btSmallGray" onclick="jfViewProfile(main.selStudent_ID.options[main.selStudent_ID.selectedIndex].value,'studentProfile.asp');"><br>
									<input type=button value="Remove Selected Student" class="btSmallGray" onclick="jfRemoveItems(selStudent_ID);jfChanged();">									
								</td>
								<script language=javascript>
									function jfSelectItemFromTo(selectFrom, selectTo) {
										//based on ideas from excite.com's weather selection - heavily modified
										var blnSelected = false;
										var selected = selectFrom.selectedIndex;
										if (selected != -1){
											for (j=0; j<selectFrom.length; j++) {
												if (selectFrom.options[j].selected){
													var selectedText = selectFrom.options[j].text;
													var selectedValue = selectFrom.options[j].value;
													if (selectedValue != "") {
														var toLength = selectTo.length;
														var i;
														// If item is already added, give it focus
														for (i=0; i<toLength; i++) {
															if (selectTo.options[i].value == selectedValue) {
																blnSelected = true;
															}
														}
														if (!blnSelected){
															// Add new option 
															selectTo.options[selectTo.length] = new Option(selectedText, selectedValue);
														}
													}
												}
												blnSelected = false;
											}
										}	
										jfMakeList(selectTo);							
									}
									
									function jfMakeList(selObj){
										//since we aren't using the multi-select in a proper way, we take all of the
										//options in the selChosenActions dropdown and write them to a hidden field
										var strItems = "";
										for (i=0; i< selObj.length; i++) {
											strItems = strItems + selObj.options[i].value + ",";
										}
										if (selObj.name == "selStudent_ID") {
											document.main.intStudent_ID.value = strItems.substr(0, strItems.length - 1); 
										}else{
											document.main.intGuardian_ID.value = strItems.substr(0, strItems.length - 1); 											
										}									
									}
									
									function jfRemoveItems(pobjSelect){
										//remove items from multiple select list
										//Since setting an option to NULL changes the index
										//value of the item beneath it, we have to make a couple
										//of passes at the object.  We first grab the quantity of
										//items selected, then we use that as a counter to remove
										//the selected items
										var iCnt = 0;
											for (i=0; i<pobjSelect.length; i++) {
												if (pobjSelect.options[i].selected){
													iCnt ++;
												}
											}
											for (j=0; j<iCnt; j++){
												for (i=0; i<pobjSelect.length; i++) {
													if (pobjSelect.options[i].selected){
														pobjSelect.options[i] = null;
													}
												}
											}
											jfMakeList(pobjSelect);
										}
								</script>
							</tr>
							<% else 
									if strInstructions <> "" then
							%>
							<tr>
								<td colspan=3>
									<table cellpadding=4 cellspacing=0 border=e6e6e6>
										<tr>
											<td class=svplain10>
												<% = strInstructions %>
											</td>											
										</tr>									
									</table>
									<br>
								</td>
							</TR>
							<%		end if	%> 
							<tr>	
								<Td class=gray >
									&nbsp;<b>Students in Family</B> &nbsp;
								</td>
								<td>
									&nbsp;
								</td>	
								<Td class=gray >
									&nbsp;<b>Guardians in Family</B>&nbsp;
								</td>
							</tr>
							<tr>
								<td class=gray valign=top>
									<table cellpadding=4 cellspacing=0 border=1 bordercolor=white>
										<% = strStudentName %>
									</table>
								</td>
								<td>
									&nbsp;
								</td>
								<td class=gray valign=top>
									<table cellpadding=4 cellspacing=0 border=1 bordercolor=white ID="Table1">												
										<% = strGuardianName %>
									</table>
								</td>								
							</tr>
							<% end if %>
						</table>
						<BR>
						<table>
							
							<% if session.Value("strRole") = "ADMIN" then%>
							<tr>	
								<Td colspan=2>
									<font class=svplain11>
										<b><i>Guardians in Family</I></B>
									</font>
								</td>
							</tr>
							<tr>
								<td class=gray>
										&nbsp;List of All Guardians &nbsp; &nbsp;
										<input type=button value="add new guardian to list" onClick="jfAddNewGuardian();" class=btSmallGray>
								</td>	
								<td>
									&nbsp;
								</td>										
								<td class=gray>
										&nbsp;Guardians in this Family &nbsp;
								</td>
								<td>
									&nbsp;
								</td>
							</tr>
							<tr>
								<td>
									<select name="selAllGuardians" multiple size=5">
									<%
										dim sqlGuardian										
										if session.Value("strRole") = "ADMIN" then
											strMore = " + ': ' + convert(nchar(5),intGuardian_ID) "
										else
											strMore = ""
										end if 
										
										'JD:Select only 'active' guardians
										sqlGuardian = "select intGuardian_ID as id, szLast_Name + ',' + szFirst_Name" & strMore & " as Name " & _
													  "from tblGuardian where blnDeleted=0 order by name"
										Response.Write oFunc.MakeListSQL(sqlGuardian,"id","Name","")
									%>
									</select>
								</td>
								<td valign=middle>
									<input type=button value="Add >" class="btSmallGray" onclick="jfSelectItemFromTo(selAllGuardians, selGuardian_ID);jfChanged();" NAME="Button1">
								</td>
								<td class=gray valign=top>
									<select name="selGuardian_ID" multiple size=5 onChange="jfChanged();" ID="Select1">
										<% = strGuardianOptions %>
									</select>	
									<input type=hidden name="intGuardian_ID" value="<% = intGuardian_ID %>" ID="Hidden1">								
								</td>
								<td valign=middle>
									<input type=button value="View Profile of Selected Guardian" class="btSmallGray" onclick="jfViewProfile(selGuardian_ID.options[selGuardian_ID.selectedIndex].value,'guardianProfile.asp');" NAME="Button2"><br>
									<input type=button value="Remove Selected Guardian" class="btSmallGray" onclick="jfRemoveItems(selGuardian_ID);jfChanged();" NAME="Button3">									
								</td>
							</tr>
							<% end if
						end if %>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<%
if ucase(session.Contents("strRole")) = "ADMIN" then
	strClick = "jfCheckForm();"
else
	strClick = "window.location.href='" & Application.Value("strWebRoot")& "';"
end if

%>	
<% if Request.QueryString("bolForced") <> "" then%>
<script language=javascript>
	function jfConfirmSIS() {
		var message = "It is critical that the information contained on this page is up to date."
		message += " Please be sure that you have reviewed all student ";
		message += "and guardian profiles and have made the appropriate changes before you continue. "
		message += "\nClicking 'OK' is confirmation that you have reviewed all student and guardian inforamtion ";
		message += "and have made any needed changes. Click 'CANCEL' to continue working on this page.";
		var bolAns = confirm(message);
		if (bolAns == true) {
			main.submit();
		}else{
			return false;
		}
	}
</script>	
<input type=hidden name="intCount" value="<% = Request.QueryString("intCount")%>">
&nbsp;<input type=button value="View Instructions" onClick="jfInstructions();" class="btSmallGray" >
<input type=submit value="Confirm/Update" class="NavSave" onClick="jfConfirmSIS();">
<% else %>
&nbsp;<input type=button value="Home Page" onClick="<% = strClick %>" class="btSmallGray" >&nbsp;
<input type=submit value="<%=strButton%>" class="NavSave" <% = strJFAction %>>
<% end if %>
</form>	
<%

set oFunc = nothing
Server.Execute(Application.Value("strWebRoot") & "Includes/footer.asp")

function vbfUpdateStudent(fam_id)
	oFunc.BeginTransCN
	dim update
	' This update will handle if any students where deleted from a family
	update = "update tblStudent set intFamily_id = null " & _
			 "where intFamily_ID = " & fam_id 
	oFunc.ExecuteCN(update)
	
	if instr(1,request("intStudent_ID"),",") > 0 then
		arStudents = split(replace(request("intStudent_ID")," ",""),",")
		for i = 0 to ubound(arStudents)
			strSQL = strSQL & " intStudent_ID = " & arStudents(i) & " or"
		next 
		strSQL = left(strSQL,len(strSQL) -2)
	else
		strSQL = " intStudent_ID = " & request("intStudent_ID") 
	end if 
	
	if request("intStudent_ID")  <> "" then
		' This update adds all the selected students into a family
		update = "update tblStudent set intFamily_ID = " & fam_id & _
				 " where " & strSQL
		oFunc.ExecuteCN(update)
	end if 
	oFunc.CommitTransCN
end function

function vbfUpdateGuardian(fam_id)
	' Handles all updating, inserting, deleteing of tascFam_Guard records based on users
	' selections.
	oFunc.BeginTransCN
	dim update
	dim sql
	dim delete
	dim bolFound
	
	if instr(1,request("intGuardian_ID"),",") > 0 then
		'Create array from list
		arGuardian = split(replace(Request.Form("intGuardian_ID")," ",""),",")
	else
		' Create array even though we don't have a list so we don't have to write two kinds of 
		' logic to handle our loops
		arGuardian = array(Request.Form("intGuardian_ID"),"")
	end if 

	'Get existing fam_guard records 
	set rsExisting = server.CreateObject("ADODB.RECORDSET")
	rsExisting.CursorLocation = 3
	sql = "select intGuardian_Id, intFamGuard_ID,bolLives_With,dtCreate ,szUser_Create " & _
		  "from tascFam_Guard where intFamily_ID = '" & replace(intFamily_Id,"new","") & "'"
	rsExisting.Open sql, Application("cnnFPCS")'oFunc.FPCScnn			
	
	if rsExisting.RecordCount > 0 then
		' We have existing records so we'll delete them all and recreate based on what the
		' user has now selected 
		delete = "delete from tascFam_Guard where intFamily_ID = " & intFamily_Id
		oFunc.ExecuteCN(delete)
	end if
	
	if isArray(arGuardian) then
		for i = 0 to ubound(arGuardian)
			if arGuardian(i) <> "" then		'saves us from looping on our dummy array element
				bolFound = false
				if rsExisting.RecordCount > 0 then
					do while not rsExisting.EOF										 								
						if arGuardian(i) = rsExisting("intGuardian_Id") then
							' Resaves previous records
							insert = "insert into tascFam_Guard (intFamily_ID,intGuardian_ID,dtCreate,szUser_Create," & _
									 "dtModify,szUser_Modify) values(" & _
									 intFamily_Id & "," & _
									 rsExisting("intGuardian_Id") & ",'" & _
									 rsExisting("dtCreate") & "','" & _
									 rsExisting("szUser_Created") & "','" & _
									 oFunc.DateTimeFormat(now()) & "','" & _
									 session.Value("strUserID") & "')"
							oFunc.ExecuteCN(insert)	
							bolFound = true											 
						end if 										
						rsExisting.MoveNext
					loop
					rsExisting.MoveFirst
				end if 
				if bolFound = false then
					' Inserts new fam_guard records
					insert = "insert into tascFam_Guard (intFamily_ID,intGuardian_ID,dtCreate,szUser_Create" & _
							 ") values(" & _
							 intFamily_Id & "," & _
							 arGuardian(i) & ",'" & _
							 oFunc.DateTimeFormat(now()) & "','" & _
							 session.Value("strUserID") & "')"
					oFunc.ExecuteCN(insert)	
				end if 										
			end if 
		next
	end if 
	rsExisting.close
	set rsExisting = nothing
	oFunc.CommitTransCN
end function

sub vbsInsertFamily
	'Creates Family Record
	dim insert
	if Request.Form("bolIncludeDir") = "on" then
		blnIncludeDir = 1
	else
		blnIncludeDir = 0
	end if
	
	insert = "insert into tblFamily (szFamily_Name,szDesc,szAddress,szCity,szState,szCountry," & _
				"szZip_Code,szHome_Phone,szEMAIL,bolIncludeDir,dtCreate,szUser_Create) " & _
				"values " & _
				"('" & oFunc.EscapeTick(Request.Form("szFamily_Name")) & "'," & _
				"'" & oFunc.EscapeTick(Request.Form("szDesc")) & "'," & _
				"'" & oFunc.EscapeTick(Request.Form("szAddress")) & "'," & _
				"'" & oFunc.EscapeTick(Request.Form("szCity")) & "'," & _
				"'" & oFunc.EscapeTick(Request.Form("szState")) & "'," & _
				"'" & oFunc.EscapeTick(Request.Form("szCountry")) & "'," & _
				"'" & oFunc.EscapeTick(Request.Form("szZip_Code")) & "'," & _
				"'" & oFunc.EscapeTick(Request.Form("szHome_Phone")) & "'," & _
				"'" & oFunc.EscapeTick(Request.Form("szEMAIL")) & "'," & _
				blnIncludeDir & "," & _
				"'" & oFunc.DateTimeFormat(now()) & "'," & _
				"'" & session.Value("strUserID") & "')" 
	oFunc.ExecuteCN(insert)
	strMessage = "Family Added"
	intFamily_Id = oFunc.GetIdentity
	call vbfUpdateStudent(intFamily_Id)
	call vbfUpdateGuardian(intFamily_Id)
end sub

sub vbsUpdateFamily
	' Updates Family Record
	dim update
	
	if Request.Form("bolIncludeDir") = "on" then
		blnIncludeDir = 1
	else
		blnIncludeDir = 0
	end if
	
	update = "update tblFamily set " & _
				"szFamily_Name = '" & oFunc.EscapeTick(Request.Form("szFamily_Name")) & "'," & _
				"szDesc = '" & oFunc.EscapeTick(Request.Form("szDesc")) & "'," & _
				"szAddress = '" & oFunc.EscapeTick(Request.Form("szAddress")) & "'," & _
				"szCity = '" & oFunc.EscapeTick(Request.Form("szCity")) & "'," & _
				"szState = '" & oFunc.EscapeTick(Request.Form("szState")) & "'," & _
				"szCountry = '" & oFunc.EscapeTick(Request.Form("szCountry")) & "'," & _
				"szZip_Code = '" & oFunc.EscapeTick(Request.Form("szZip_Code")) & "'," & _
				"szHome_Phone = '" & oFunc.EscapeTick(Request.Form("szHome_Phone")) & "'," & _
				"szEMAIL = '" & oFunc.EscapeTick(Request.Form("szEMAIL")) & "'," & _
				"bolIncludeDir = " & blnIncludeDir & "," & _
				"dtModify = '" & oFunc.DateTimeFormat(now()) & "'," & _
				"szUser_Modify = '" & session.Value("strUserID")  & "' " & _
				"where intFamily_ID = " & intFamily_Id 
	'response.write "<br><br>" & update 
	oFunc.ExecuteCN(update)
	strMessage = "Family Updated"	

	if session.Value("strRole") = "ADMIN" then
		call vbfUpdateStudent(intFamily_Id)		
		call vbfUpdateGuardian(intFamily_Id)
	end if
end sub
%>
