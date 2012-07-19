<%@ Language=VBScript %>
<%
dim oFunc				'wsc object
dim sqlMessage			'SQL to get messages for guards or teachers to be displayed on page
dim strMessage			'Contains message to display tpo appropriate user
dim strMessageHTML		'HTML table that contains the message in strMessage


Response.Expires = -1000	'Makes the browser not cache this page

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
set oList = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/dbOptionsList.wsc"))



call oFunc.OpenCN()

'Reset Session Variables
Session.Contents("intStudent_ID") = ""
Session.Contents("intInstructor_ID") = ""
'session.Value("intFamily_id")
oFunc.ResetSelectSessionVariables()
 
Session.Value("strTitle") = "FPCS Information Systems"
Session.Value("strLastUpdate") = "12 May 2002"

Server.Execute(Application.Value("strWebRoot") & "Includes/header.asp")

' This section will get messages for users that have either Guard or Teacher 
' roles and will include the message in the page that this script generates.
if session.Contents("strRole") = "TEACHER" then
	sqlMessage = "select szMessage " & _
				  "from tblMessages " & _
				  "where intGroup_ID = 2 " & _
				  "and bolShow_Message = 1"	
elseif session.Contents("strRole") = "GUARD" then
	sqlMessage = "select szMessage " & _
					"from tblMessages " & _
					"where intGroup_ID = 1 " & _
					"and bolShow_Message = 1"
end if

'Only query for message if needed		  
if sqlMessage <> "" then
	set rsMessages = server.CreateObject("ADODB.RECORDSET")
	rsMessages.CursorLocation = 3
	rsMessages.Open sqlMessage, Application("cnnFPCS")'oFunc.FPCScnn
	
	' Create html table and insert the message to display
	if rsMessages.RecordCount > 0 then
		do while not rsMessages.EOF
			strMessage = strMessage & rsMessages(0) & "<BR><BR>"
			rsMessages.MoveNext
		loop
		strMessageHTML = "<br><BR><center>" & chr(13) & _
					"<table cellspacing=0 border=1 bordercolor=#ffcc66 width=400 >" & chr(13) & _
					"	<tr>" & chr(13) & _
					"		<td>" & chr(13) & _
					"			<table>" & chr(13) & _
					"				<tr>" & chr(13) & _
					"					<Td>" & chr(13) & _
					"						<font size=-1 face=tahoma><b>SYSTEM MESSAGE</b></font>" & chr(13) & _
					"					</td>" & chr(13) & _
					"				</tr>" & chr(13) & _
					"				<tr>" & chr(13) & _
					"					<td>" & chr(13) & _
					"						<font size=-1 face=tahoma>" & strMessage & "</font>" & chr(13) & _
					"					</td>" & chr(13) & _
					"				</tr>" & chr(13) & _
					"			</table>" & chr(13) & _
					"		</td>" & chr(13) & _
					"	</tr>" & chr(13) & _
					"</table></center>"
	end if 
	rsMessages.Close
	set rsMessages = nothing
end if 


   
%>
<script language="javascript">
	function jfAction(myForm,myItem,requiredField,fieldTitle)	{	
		var myPath = myItem.value;
		if (myPath.indexOf("**") > 0 ) {
			if (requiredField != "") {
				var strField = requiredField.value;	
				// Verify we have the info we need to take action
				if (strField.length > 0) {
					
					if (strField == "" || typeof strField == "undefined")	{
						alert("You must provide a value for " + fieldTitle + ".");
						//myItem.selectedIndex = 0;
						return false;
					}
				}else{
					alert("You must provide a value for " + fieldTitle + ".");
					return false;
				}
			}
			var sRemove = "**";
			//var re = new RegExp(sRemove,"g");
			myPath = myPath.replace(/\*\*/, '');
		}
		// Now form the URL 
		var rootPath;
		rootPath = "<% = Application.Value("strWebRoot") %>";
		rootPath = rootPath + myPath;
		if (requiredField != "") {
			if (typeof strField != "undefined") {
				if (rootPath.indexOf("?") > 0) {
					rootPath += "&" + requiredField.name + "=" + strField;
				}else{
					rootPath += "?" + requiredField.name + "=" + strField;
				}
			}
		}
		//Execute the action 
		
		window.location.href = rootPath;
	}
	
	function jfAction2(myForm,myLink,requiredField,fieldTitle)	{		
		
		if (requiredField != "") {
			var strField = document.getElementById(requiredField).value;	
			// Verify we have the info we need to take action
			if (strField.length > 0) {
				
				if (strField == "" || typeof strField == "undefined")	{
					alert("You must provide a value for " + fieldTitle + ".");
					//myItem.selectedIndex = 0;
					return false;
				}
			}else{
				alert("You must provide a value for " + fieldTitle + ".");
				return false;
			}
		}
		// Now form the URL 
		var rootPath;
		rootPath = "<% = Application.Value("strWebRoot") %>";
		rootPath = rootPath + myLink;
		
		if (requiredField != "") {
			if (rootPath.indexOf("?") > 0) {
				rootPath += "&" + requiredField + "=" + strField;
			}else{
				rootPath += "?" + requiredField + "=" + strField;
			}
		}
		//Execute the action 
		window.location.href = rootPath;
	}
	
	function jfManageEmail(pEmail){
		var sList = document.main.strEmailList;
		
		if (sList.value.indexOf(";"+pEmail+";") == -1 ) {
			// Email is not in list so add it
			sList.value = sList.value + pEmail + ";";
		}else{
			// Email is in list so remove it
			var re = new RegExp(pEmail + ";",'gi');
			sList.value = sList.value.replace(re,'');
		}
	}
	
	function jfOpenMailClient(){
		var sList = document.main.strEmailList;
		window.location.href = "mailto:" + sList.value;
	}
	
	function jfEmailAll(){
	  var sEmail = document.getElementById("allMailList");
          var sDiv = document.getElementById("divEmail");
	  var sTd = document.getElementById("taEmail");
		
	   sTd.value = sEmail.value;
	   sDiv.style.display = "";
	   scroll(0,0);

	}
	
	function jfReimburse(){
		var reimbWin;
		var strURL = "<%=Application.Value("strWebRoot")%>forms/Requisitions/reimburseForm.asp";
		reimbWin = window.open(strURL,"reimbWin","width=710,height=500,scrollbars=yes,resize=yes,resizable=yes");
		reimbWin.moveTo(0,0);
		reimbWin.focus();
	}
</script>
<% 
'response.write session.Contents("student_list") & "<BR>"
 %>


<% if ucase(session.contents("strUserID")) = "SCOTT" then %>
<a href="./UserAdmin/SessionBridge.asp">Test Link</a><br><br>
<% end if %>
<script type="text/javascript" language="javascript">
	function jfViewProgress(pPage){
		var winProgress;
		var strURL = "<%=Application.Value("strWebRoot")%>useradmin/sessionbridge.asp?page=" + pPage;
		winProgress = window.open(strURL,"winProgress","width=950,height=650,scrollbars=yes,resize=yes,resizable=yes");
		winProgress.moveTo(0,0);
		winProgress.focus();
	}
</script>
<form name="main" action="" ID="Form1">
<input type="hidden" name="strEmailList" value=";">
<table width="100%" ID="Table1">
	<tr>
		<td class="yellowHeader">
			<% if session.Contents("strRole") <> "GUARD" then %>
			&nbsp;<b>FPCS Online System Home Page</b>
			<% elseif Session.Contents("strFamily_Name") <> "" then%>
			&nbsp;<b><% = Session.Contents("strFamily_Name") %> Family Home Page</b>
			<% end if %>
		</td>
	</tr>
	<tr>		
			<% if Session.Value("strRole") = "ADMIN" then %>
		<td bgcolor="f7f7f7">
			<table ID="Table2">
				<tr>
					<td colspan="6">
						<font class="svplain11"><b><i>Student Information Options</i></b> </font>
					</td>
				</tr>
				<tr>
					<td class="TableHeader" colspan="3">
						&nbsp;<b>Add a New Family/Student to SIS: <input type="button" value="Add &gt;" onclick="window.location.href='<% = Application.Value("strWebRoot") %>Admin/familyManager.asp';" class="btSmallGray" ID="Button1" NAME="Button1">
					</td>
				</tr>
					<tr>
						<td class="gray">
							&nbsp;<b>Select a Student</b>
						</td>
						<td class="gray">
							&nbsp;<b>Action to take?</b>
						</td>
					</tr>
					<tr>
						<td align="center">
							<select name="intStudent_ID" style="width:300px;" ID="intStudent_ID">
								<option value>
									<%
							dim sqlStudent														
							
								sqlStudent = "SELECT     s.intSTUDENT_ID, (CASE ss.intReEnroll_State WHEN 86 THEN s.szLAST_NAME + ',' + s.szFIRST_NAME + ': Withdrawn (' + CASE isNull(ss.dtWithdrawn, " & _ 
											" 1) WHEN 1 THEN 'No Date Entered' ELSE CONVERT(varChar(100), ss.dtWithdrawn)  " & _ 
											" END + ')' WHEN 123 THEN s.szLAST_NAME + ',' + s.szFIRST_NAME + ': Graduated (' + CONVERT(varChar(20), ss.dtModify)  " & _ 
											" + ')' ELSE s.szLAST_NAME + ',' + s.szFIRST_NAME END) AS Name, ss.intReEnroll_State, ss.dtWithdrawn " & _ 
											"FROM tblSTUDENT s INNER JOIN " & _ 
											" tblStudent_States ss ON s.intSTUDENT_ID = ss.intStudent_id " & _ 
											"WHERE (ss.intReEnroll_State IN (" & Application.Contents("strEnrollmentList") & ")) AND (ss.intSchool_Year = " & Session.Value("intSchool_Year") & ") " & _ 
											"ORDER BY Name "
							
						
							Response.Write oFunc.MakeListSQL(sqlStudent,intStudent_ID,Name,"")												 
						%>
							</select>
						</td>
						<td>
							<select name="studentPath"  style="width:180px;" ID="Select2">
								<option value>
						<%
							dim sqlScripts
							sqlScripts = "Select strPath + case when bolRequire_Validate = 1 then '**' else '' end as strPath,strScript_Name " & _
											 "from tblFPCS_Scripts where intScript_Group_Id = 1 and bolVersion_2_Off  is  null order by strScript_Name"
							Response.Write oFunc.MakeListSQL(sqlScripts,"strPath","strScript_Name","")												 
						%>
							</select>
							<input type="button" value="go" class="btSmallGray" onClick="jfAction(this.form,this.form.studentPath,this.form.intStudent_ID,'Student');" ID="Button2" NAME="Button2">
						</td>
					</tr>
			</table>
			<br>
			<table ID="Table3">
				<tr>
					<td colspan="6">
						<font class="svplain11"><b><i>Teacher Information Options</i></b> </font>
					</td>
				</tr>
				<tr>
					<td class="TableHeader" colspan="3">
						&nbsp;<b>Add a New Teacher to TIS: <input type="button" onclick="window.location.href='<% = Application.Value("strWebRoot") %>Forms/Teachers/addTeacher.asp?new=true;'" value="Add &gt;" class="btSmallGray" ID="Button3" NAME="Button3">
					</td>
				</tr>
					<!--<input type="hidden" name="intInstruct_Type_ID" value="4"> -->
					<input type="hidden" name="bolFromTeacher" value="True" ID="Hidden1">
					<input type="hidden" name="intInstructor_ID" value="" ID="intInstructor_ID">
					<tr>
						<td class="gray">
							&nbsp;<b>Select a Teacher</b>
						</td>
						<td class="gray">
							&nbsp;<b>Action to take?</b>
						</td>
					</tr>
					<tr>
						<td align="center">
							<table ID="Table6">
								<tr>
									<td>
										<select name="intInstructor_ID1" onChange="this.form.intInstructor_ID.value=this.value;" style="width:300px;" ID="Select3">
											<option value="">ACTIVE TEACHERS
												<%
										response.Write oList.ActiveTeachers(session.Contents("intSchool_Year"),"")								
												%>
										</select>
									</td>
								</tr>
								<tr>
									<td>
										<select name="intInstructor_ID2" onChange="this.form.intInstructor_ID.value=this.value;" style="width:300px;" ID="Select4">
											<option value="">INACTIVE TEACHERS
												<%
										response.Write oList.InactiveTeachers(session.Contents("intSchool_Year"),"")								
												%>
										</select>
									</td>
								</tr>
							</table>
						</td>
						<td valign="top">
							<select name="teacherPath"  style="width:180px;" ID="Select5">
								<option value="">
									<%
							sqlScripts = "Select strPath + case when bolRequire_Validate = 1 then '**' else '' end as strPath,strScript_Name " & _
											 "from tblFPCS_Scripts where intScript_Group_Id = 2  and bolVersion_2_Off  is  null  order by strScript_Name"
							Response.Write oFunc.MakeListSQL(sqlScripts,"strPath","strScript_Name","")												 
						%>
							</select>
							<input type="button" value="go" class="btSmallGray" onClick="jfAction(this.form,this.form.teacherPath,this.form.intInstructor_ID,'Teacher');" NAME="Button1" ID="Button4">
						</td>
					</tr>
			</table>
			<br>
			<table ID="Table4">
				<tr>
					<td colspan="6">
						<font class="svplain11"><b><i>Vendor Information Options</i></b> </font>
					</td>
				</tr>
				<tr>
					<td class="TableHeader" colspan="3">
						&nbsp;<b>Add a New Vendor to VIS: <input type="button" value="Add &gt;" onclick="window.location.href='<% = Application.Value("strWebRoot") %>Forms/VIS/vendorAdmin.asp?new=true';" class="btSmallGray" ID="Button5" NAME="Button5">
					</td>
				</tr>
					<tr>
						<td class="gray">
							&nbsp;<b>Select a Vendor</b>
						</td>
						<td class="gray">
							&nbsp;<b>Action to take?</b>
						</td>
					</tr>
					<tr>
						<td align="center">
							<input type="hidden" name="intVendor_ID" value="" ID="intVendor_ID">
							<table ID="Table5">
								<tr>
									<td>
									
									
							<select name="intVendor_ID1" onChange="this.form.intVendor_ID.value=this.value;" style="width:300px;" ID="Select6">
								<option >APPROVED, PENDING & REMOVED VENDOR LIST
									<%
							dim sqlVendor
							sqlVendor = "SELECT     intVendor_ID, szVendor_Name + ': ' + vStatus AS VendorName " & _ 
										"FROM         (SELECT     v.intVendor_ID, v.szVendor_Name, " & _ 
										"                                                  (SELECT     TOP 1 upper(szVendor_Status_CD) " & _ 
										"                                                    FROM          tblVendor_Status vs " & _ 
										"                                                    WHERE      vs.intVendor_ID = v.intVendor_ID AND vs.intSchool_Year <= " & session.Contents("intSchool_Year") & _ 
										"                                                    ORDER BY intSchool_Year DESC, intVendor_Status_ID DESC) AS vStatus " & _ 
										"                       FROM          tblVendors v " & _ 
										"                       WHERE      (SELECT     TOP 1 upper(szVendor_Status_CD) " & _ 
										"                                               FROM          tblVendor_Status vs " & _ 
										"                                               WHERE      vs.intVendor_ID = v.intVendor_ID AND vs.intSchool_Year <= " & session.Contents("intSchool_Year")  & _ 
										"                                               ORDER BY intSchool_Year DESC, intVendor_Status_ID DESC) IN ('APPR', 'PEND', 'REMV')) DERIVEDTBL " & _ 
										"ORDER BY szVendor_Name "
							Response.Write oFunc.MakeListSQL(sqlVendor,"intVendor_ID","VendorName","")													 
						%>
							</select>
							</td>
						</tr>
						<tr>
							<td>
							<select name="intVendor_ID2" onChange="this.form.intVendor_ID.value=this.value;" style="width:300px;" ID="Select7">
								<option value="">REJECTED VENDOR LIST
									<%
							sqlVendor = "SELECT     intVendor_ID, szVendor_Name + ': ' + " & _ 
										"                          (SELECT     TOP 1 upper(szVendor_Status_CD) " & _ 
										"                            FROM          tblVendor_Status vs " & _ 
										"                            WHERE      vs.intVendor_ID = v.intVendor_ID AND vs.intSchool_Year <= " & session.Contents("intSchool_Year") & _ 
										"                            ORDER BY intSchool_Year DESC, intVendor_Status_ID DESC) AS vendorName " & _ 
										"FROM         tblVendors v " & _ 
										"WHERE     ((SELECT     TOP 1 upper(szVendor_Status_CD) " & _ 
										"                         FROM         tblVendor_Status vs " & _ 
										"                         WHERE     vs.intVendor_ID = v.intVendor_ID AND vs.intSchool_Year <= " & session.Contents("intSchool_Year") & _ 
										"                         ORDER BY intSchool_Year DESC, intVendor_Status_ID DESC) IN ('REJC')) " & _ 
										"ORDER BY szVendor_Name "
							Response.Write oFunc.MakeListSQL(sqlVendor,"intVendor_ID","vendorName","")												 
						%>
							</select>
							</td>
						</tr>							
					</table>
						</td>
						<td valign="top">
							<select name="vendorPath"  style="width:180px;" ID="Select8">
								<option value>
									<%
							sqlScripts = "Select strPath + case when bolRequire_Validate = 1 then '**' else '' end as strPath,strScript_Name " & _
										 "from tblFPCS_Scripts where intScript_Group_Id = 3  and bolVersion_2_Off  is  null  order by strScript_Name"
							Response.Write oFunc.MakeListSQL(sqlScripts,"strPath","strScript_Name","")												 
						%>
							</select>
							<input type="button" value="go" class="btSmallGray" onClick="jfAction(this.form,this.form.vendorPath,this.form.intVendor_ID,'Vendor');" NAME="Button1" ID="Button6">
						</td>
					</tr>
			</table>
			<br>
			<table ID="Table7">
				<tr>
					<td>
						<table ID="Table8">
							<tr>
								<td colspan="6">
									<font class="svplain11"><b><i>FPCS Reports</i></b> </font>
								</td>
							</tr>
								<tr>
									<td class="TableHeader">
										&nbsp;<b>Select a Report</b> 
									</td>
								</tr>
								<tr>
									<td>
										<select name="reportPath" ID="Select9">
											<option value>
												<%
													sqlScripts = "Select strPath,strScript_Name " & _
																	 "from tblFPCS_Scripts where intScript_Group_Id = 4  and bolVersion_2_Off  is  null order by strScript_Name"
													Response.Write oFunc.MakeListSQL(sqlScripts,"strPath","strScript_Name","")												 
												%>
										</select>
										<input type="button" value="go" class="btSmallGray" onClick="jfAction(this.form,this.form.reportPath,this.form.reportPath,'');" NAME="Button1" ID="Button7">
									</td>
								</tr>
						</table>
					</td>
					<td>
						<table ID="Table9">
							<tr>
								<td colspan="6">
									<font class="svplain11"><b><i>FPCS Admin Tools</i></b> </font>
								</td>
							</tr>
								<tr>
									<td class="TableHeader">
										&nbsp;<b>Select an Admin Tool</b> 
									</td>
								</tr>
								<tr>
									<td>
										<select name="toolsPath" ID="Select10">
											<option value>
												<%
										sqlScripts = "Select strPath,strScript_Name " & _
														 "from tblFPCS_Scripts where intScript_Group_Id = 5  and bolVersion_2_Off  is  null  order by strScript_Name"
										Response.Write oFunc.MakeListSQL(sqlScripts,"strPath","strScript_Name","")												 
									%>
										</select>
										<input type="button" value="go" class="btSmallGray" onClick="jfAction(this.form,this.form.toolsPath,this.form.toolsPath,'');" NAME="Button1" ID="Button8">
									</td>
								</tr>
						</table>
					</td>
				</tr>
			</table>
			<%elseif Session.Value("strRole") = "TEACHER" then 
				
				sql = "SELECT     t1.nonSponsor, t2.Sponsor, t3.Total, t4.Classes " & _ 
						"FROM         (SELECT     COUNT(*) AS nonSponsor " & _ 
						"                       FROM          tblClasses c INNER JOIN " & _ 
						"                                              tblILP i ON c.intClass_ID = i.intClass_ID INNER JOIN " & _ 
						"                                              tblENROLL_INFO ON i.intStudent_ID = tblENROLL_INFO.intSTUDENT_ID " & _ 
						"                       WHERE      (c.intInstructor_ID = " & session.Contents("instruct_ID") & ") AND (i.sintSchool_Year = " & session.Contents("intSchool_Year") & ") AND (tblENROLL_INFO.sintSCHOOL_YEAR = " & session.Contents("intSchool_Year") & ") AND  " & _ 
						"                                              (tblENROLL_INFO.intSponsor_Teacher_ID <> " & session.Contents("instruct_ID") & ") AND (i.instructorStatusId <> 1 OR " & _ 
						"                                              i.instructorStatusId IS NULL)) t1 CROSS JOIN " & _ 
						"                          (SELECT     COUNT(*) AS Sponsor " & _ 
						"                            FROM          tblClasses c INNER JOIN " & _ 
						"                                                   tblILP i ON c.intClass_ID = i.intClass_ID INNER JOIN " & _ 
						"                                                   tblENROLL_INFO ON i.intStudent_ID = tblENROLL_INFO.intSTUDENT_ID " & _ 
						"                            WHERE      (c.intInstructor_ID = " & session.Contents("instruct_ID") & ") AND (i.sintSchool_Year = " & session.Contents("intSchool_Year") & ") AND (tblENROLL_INFO.sintSCHOOL_YEAR = " & session.Contents("intSchool_Year") & ") AND  " & _ 
						"                                                   (tblENROLL_INFO.intSponsor_Teacher_ID = " & session.Contents("instruct_ID") & ") AND (i.SponsorStatusId <> 1 OR " & _ 
						"                                                   i.SponsorStatusId IS NULL)) t2 CROSS JOIN " & _ 
						"                          (SELECT     COUNT(*) AS Total " & _ 
						"                            FROM          tblClasses c INNER JOIN " & _ 
						"                                                   tblILP i ON c.intClass_ID = i.intClass_ID " & _ 
						"                            WHERE      (c.intInstructor_ID = " & session.Contents("instruct_ID") & ") " & _ 
						"                            GROUP BY i.sintSchool_Year " & _ 
						"                            HAVING      (i.sintSchool_Year = " & session.Contents("intSchool_Year") & ")) t3 CROSS JOIN " & _
						"				( Select count(*) Classes from tblClasses where intInstructor_ID = " & session.Contents("instruct_ID") & _
						"							and  intSchool_Year = " & session.Contents("intSchool_Year") & ") t4 "
	
			set rsList = server.CreateObject("ADODB.RECORDSET")
			rsList.CursorLocation = 3
			rsList.Open sql, Application("cnnFPCS")'oFunc.FPCSCnn


			
			if rsList.RecordCount > 0 then
				if not isNumeric(rsList("nonSponsor")) then 
					myNonSponsor = 0 
				else
					myNonSponsor = rsList("nonSponsor")
				end if
				
				if not isNumeric(rsList("Sponsor")) then 
					mySponsor = 0 
				else
					mySponsor = rsList("Sponsor")
				end if
				
				if not isNumeric(rsList("Total")) then 
					myTotal = 0 
				else
					myTotal = rsList("Total")
				end if
				
				if not isNumeric(rsList("Classes")) then 
					myClasses = 0 
				else
					myClasses = rsList("Classes")
				end if
			else
				myNonSponsor = 0 
				mySponsor = 0 
				myTotal = 0
				myClasses = 0
			end if
			
			rsList.Close
			
			
			%>
			<td bgcolor="f7f7f7">
			<table ID="Table10">
				<tr>
					<td colspan="6">
						<font class="svplain11"><b><i>Teacher Information Options</i></b> </font>
					</td>
				</tr>
					<input type="hidden" name="intInstruct_Type_ID" value="4" ID="Hidden4"> 
					<input type="hidden" name="bolFromTeacher" value="True" ID="Hidden5">
					<tr>
						<td class="TableHeader">
							&nbsp;<b>Current Teacher</b>
						</td>
						<td class="TableHeader">
							&nbsp;<b>Action to take?</b>
						</td>
					</tr>
					<tr>		
						<td class="gray" align="center">		
							<% = Session.Value("strFullName") %>
						<%					
						if 	session.Contents("instruct_id") <> "" then		
														 
						%>
								<input type="hidden" name="intInstructor_ID" value="<%=session.Contents("instruct_id") %>" ID="intInstructor_ID">						
							</td>
							<td>
								<select name="insPath" ID="Select11">
									<option value>
									<%
							sqlScripts = "Select strPath + case when bolRequire_Validate = 1 then '**' else '' end as strPath,strScript_Name " & _
											 "from tblFPCS_Scripts where intScript_Group_Id = 2  and bolVersion_2_Off  is  null  order by strScript_Name"
							Response.Write oFunc.MakeListSQL(sqlScripts,"strPath","strScript_Name","")												 
						%>
								</select>
								<input type="button" value="go" class="btSmallGray" onClick="jfAction(this.form,this.form.insPath,this.form.intInstructor_ID,'');" NAME="Button1" ID="Button9">
							</td>
						<% else %>
							<td colspan="3" class="gray">
								<b>Your user has a valid role as &quot;Teacher&quot; but does not
								have a valid Instructor Profile. The FPCS staff will need to 
								associate your user with a valid instructor.</b>
							</td>
						<% end if %>
						</tr>
						<tr>
							<td class="svplain8" colspan=3>
								You are instructing <% = myClasses %> courses with <% = myTotal %> contracts of which <% = (myNonSponsor + mySponsor) %> contracts have not been signed by you.
								<BR>Click <a href="<% = Application.Value("strWebRoot") %>forms/teachers/ContractManager.asp">HERE</a> to go to the Instructor Contract Manager.
							</td>
						</tr>
                        <%If False Then %>
						<tr>
							<td colspan="5">
								<span style="font-family:arial;color:red;font-size:15pt;"><b>NEW!</b></span>
<span style="font-family:arial;color:black;font-size:10pt;">
<b>&nbsp;&nbsp;Thanks to the diligence of the FPCS staff and teachers we now have two new tools to help track student progress. They are the 'Online Progress by Student' 
and 'Online Progress by Class'.  Not very catchy names BUT they should be a help. <BR><br> To see what this is all about simply click a link below.  This is new code so there
may be some bugs.  Please report any bugs to <a href="mailto:fpcs_admin@fpcs.net">Bug Report</a>. Be sure to explain your situation well.<br><br>
<% if session.Contents("student_List") <> "" then%>
<a Href="javascript:void(0);" onClick="jfViewProgress('progressForm.aspx');">Online Progress By Student</a><br>
<% end if %>
<a Href="javascript:void(0);" onClick="jfViewProgress('ClassProcessCtrl.aspx');">Online Progress By Class</a></b></span><br>
							
							</td>	
						</tr>
                        <%End If %>
						<% if session.Contents("student_List") <> "" then%>
						<tr>
							<td align="center" colspan='5'>
								&nbsp;<br>
<div style="border: #000000 1px solid;display:none;width:450;"  id='divEmail'>
<table style='width:450;background-color:#ffffff;'>
	<tr>
		<td class='svplain10'>
			Below is a list of all your family emails.  Copy and paste this into the 'To:' or 'BCC:' section of your email program.
		</td>
	</tr>
	<tr>
		<td class='svplain8' id="tdEmail">
			<textarea style='width:99%;' rows='10' id='taEmail'></textarea>
		</td>
	</tr>
</table>	
</div>


							</td>
						</tr>
						<tr>
							<td  colspan="6">							
								<table ID="Table11">
									<tr>
										<td>
											<font class="svplain11"><b><i>Sponsor Teacher Students:</i></b> </font>
											<span class="svplain8">When Student column is <span class="green">&nbsp;green&nbsp;</span> that signifies the Principal has initially approved the Packet.</span>
										</td>
										<td colspan="2" class="svplain8" align="right" valign="bottom">
											<% 
											set oHtml = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/htmlFuncs.wsc"))
						
						strDefinitions = "<table> " & _
								"	<tr> " & _
								"		<td class='TableCell' valign='top'> " & _
								"			<i>GNS</i> " & _
								"		</td> " & _
								"		<td class='TableCell'> " & _
								"			'Guardin Not Signed' Number of Courses that have not been signed by the Guardian. " & _
								"		</td> " & _
								"	</tr> " & _
								"	<tr> " & _
								"		<td class='TableCell' valign='top'> " & _
								"			<i>SNS</i> " & _
								"		</td> " & _
								"		<td class='TableCell'> " & _
								"			'Sponsor Not Signed' Number of Courses that have not been signed by the Sponsor. " & _
								"		</td> " & _
								"	</tr> " & _
								"	<tr> " & _
								"		<td class='TableCell' valign='top'> " & _
								"			<i>INS</i> " & _
								"		</td> " & _
								"		<td class='TableCell'> " & _
								"			'Instructor Not Signed' Number of Courses that have not been signed by the Instructor. " & _
								"		</td> " & _
								"	</tr> " & _
								"	<tr> " & _
								"		<td class='TableCell' valign='top'> " & _
								"			<i>ANS</i> " & _
								"		</td> " & _
								"		<td class='TableCell'> " & _
								"			'Admin Not Signed' Number of Courses that have not been signed by the Admin. " & _
								"		</td> " & _
								"	</tr> " & _
								"	<tr> " & _
								"		<td class='TableCell' valign='top'> " & _
								"			<i>Alerts</i> " & _
								"		</td> " & _
								"		<td class='TableCell'> " & _
								"			Number of Sponsor Alerts a student has within their Packet. " & _
								"		</td> " & _
								"	</tr> " & _
								"	<tr> " & _
								"		<td class='TableCell' valign='top'> " & _
								"			<i>PRS 1</i> " & _
								"		</td> " & _
								"		<td class='TableCell'> " & _
								"			Progress Report for Semester One. Has three states 'N/A' which means guardian has not " & _
								"			completed the progress report, 'View' which means the report has been completed and  " & _
								"			ready for sponsors review and 'Reviewed' which means the sponsor has reviewed the report. " & _
								"		</td> " & _
								"	</tr> " & _
								"	<tr> " & _
								"		<td class='TableCell' valign='top'> " & _
								"			<i>PRS 2</i> " & _
								"		</td> " & _
								"		<td class='TableCell'> " & _
								"			Progress Report for Semester Two. " & _
								"		</td> " & _
								"	</tr> " & _
								"</table>	 " 					
								response.Write oHtml.ToolTip("<a href='#'>Column Definitions</a>&nbsp;",strDefinitions,true,"Column Definitions",false,"ToolTip","400px","",true,true) 
								
								response.Write oHtml.ToolTipDivs
								set oHtml = nothing											
								%>
										</td>
									</tr>
						<tr>
							<td colspan="10">
								<table ID="Table12">									
									<%
									
									
									sql = "SELECT     s.szLAST_NAME + ', ' + s.szFIRST_NAME AS Name, f.szHome_Phone, f.szEMAIL, s.intSTUDENT_ID, ss.intReEnroll_State, ss.szGrade, " & _ 
											"                          (SELECT     COUNT(*) " & _ 
											"                            FROM          tblILP i " & _ 
											"                            WHERE      i.intStudent_ID = s.intStudent_ID AND (i.GuardianStatusID <> 1 OR " & _ 
											"                                                   i.GuardianStatusID IS NULL) AND i.sintSchool_Year = ss.intSchool_Year) AS GuardNotSign, " & _ 
											"                          (SELECT     COUNT(*) " & _ 
											"                            FROM          tblILP i " & _ 
											"                            WHERE      i.intStudent_ID = s.intStudent_ID AND (i.SponsorStatusID <> 1 OR " & _ 
											"                                                   i.SponsorStatusID IS NULL) AND i.sintSchool_Year = ss.intSchool_Year) AS SponsorNotSign, " & _ 
											"                          (SELECT     COUNT(*) " & _ 
											"                            FROM          tblILP i INNER JOIN " & _ 
											"                                                   tblClasses c3 ON i.intClass_ID = c3.intClass_ID " & _ 
											"                            WHERE      i.intStudent_ID = s.intStudent_ID AND (i.InstructorStatusID <> 1 OR " & _ 
											"                                                   i.InstructorStatusID IS NULL) AND i.sintSchool_Year = ss.intSchool_Year AND c3.intInstructor_Id IS NOT NULL) AS InstructorNotSign, " & _ 
											"                          (SELECT     COUNT(*) " & _ 
											"                            FROM          tblILP i INNER JOIN " & _ 
											"                                                   tblClasses c3 ON i.intClass_ID = c3.intClass_ID " & _ 
											"                            WHERE      i.intStudent_ID = s.intStudent_ID AND (i.AdminStatusID <> 1 OR " & _ 
											"                                                   i.AdminStatusID IS NULL) AND i.sintSchool_Year = ss.intSchool_Year AND c3.intInstructor_ID IS NULL) AS AdminNotSign, " & _ 
											"                          (SELECT     COUNT(*) " & _ 
											"                            FROM          tblILP i " & _ 
											"                            WHERE      i.intStudent_ID = s.intStudent_ID AND (i.AdminStatusID = 2) AND i.sintSchool_Year = ss.intSchool_Year) AS AdminMustAdmin, " & _ 
											"			p1.PROGRESS_ABBR as PROGRESS_1, p1.PROGRESS_REPORT_STATUS_TEXT PROGRESS_FULL_1, " & _ 
											"			p2.PROGRESS_ABBR as PROGRESS_2, p2.PROGRESS_REPORT_STATUS_TEXT PROGRESS_FULL_2,	 " & _ 
											"                          (SELECT     dtSponsor_Reviewed " & _ 
											"                            FROM          tblProgress_Reports " & _ 
											"                            WHERE      (intStudent_ID = s.intStudent_ID) AND (intSchool_Year = ss.intSchool_Year) AND (intReporting_Period_ID = 1)) AS Teacher_Viewed_I, " & _ 
											"                          (SELECT     dtSponsor_Reviewed " & _ 
											"                            FROM          tblProgress_Reports " & _ 
											"                            WHERE      (intStudent_ID = s.intStudent_ID) AND (intSchool_Year = ss.intSchool_Year) AND (intReporting_Period_ID = 2))  " & _ 
											"                      AS Teacher_Viewed_II, e.AdminPacketSigned,e.PacketSignDate,  " & _ 
											" (SELECT COUNT(*) " & _ 
											"		FROM          tblILP i " & _ 
											"		WHERE      i.intStudent_ID = s.intStudent_ID AND i.bolSponsorAlert = 1 AND i.sintSchool_Year = ss.intSchool_Year) AS Alerts, " & _
											" (SELECT COUNT(*) " & _ 
											"		FROM          tblILP i " & _ 
											"		WHERE      i.intStudent_ID = s.intStudent_ID AND i.bolParentAlert = 1 AND i.sintSchool_Year = ss.intSchool_Year) AS GuardAlerts " & _
											"FROM         tblENROLL_INFO e INNER JOIN " & _ 
											"                      tblSTUDENT s ON e.intSTUDENT_ID = s.intSTUDENT_ID INNER JOIN " & _ 
											"                      tblFAMILY f ON s.intFamily_ID = f.intFamily_ID INNER JOIN " & _ 
											"                      tblStudent_States ss ON ss.intStudent_id = s.intSTUDENT_ID " & _ 
											" left outer join " & _ 
											" dbo.PROGRESS_REPORTS_STATUS pr1 on pr1.Student_ID = s.intSTUDENT_ID and pr1.SCHOOL_YEAR =  " & _ 
											" ss.intSchool_Year and pr1.SEMISTER_ID = 1 left outer join PROGRESS_REPORT_STATUS p1 " & _ 
											" on pr1.REPORT_STATUS_ID = p1.PROGRESS_REPORT_STATUS_ID LEFT OUTER JOIN " & _ 
											"dbo.PROGRESS_REPORTS_STATUS pr2 on pr2.Student_ID = s.intSTUDENT_ID and pr2.SCHOOL_YEAR =  " & _ 
											" ss.intSchool_Year and pr2.SEMISTER_ID = 2 left outer join PROGRESS_REPORT_STATUS p2 " & _ 
											" on pr2.REPORT_STATUS_ID = p2.PROGRESS_REPORT_STATUS_ID  " & _ 
											"WHERE     (e.intSponsor_Teacher_ID = " & session.Contents("instruct_ID") & ") AND (e.sintSCHOOL_YEAR = " & session.Contents("intSchool_Year") & ") AND (ss.intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _ 
											"ORDER BY s.szLAST_NAME, s.szFIRST_NAME "
										
										'if ucase(session.Contents("strUserID")) = "CHRONIH30" then
										'	response.Write sql
										'end if	 					
									rsList.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
									
									if rsList.RecordCount > 0 then
										do while not rsList.EOF
											if rowCount mod 30 = 0 then
												call vbsTeacherHeader 
											end if
											
											if rsList("AdminPacketSigned") then
												StudentCss = "green"
												myAttrib = " class='lnkWhite' title='Packet Signed: " & FormatDateTime(cdate(rsList("PacketSignDate")),2) & "' "
											else
												StudentCss = "TableCell"
												myAttrib = ""
											end if
									%>
									<tr >
										<td class="<% = StudentCss %>">
											&nbsp;<a href="<% = Application.Value("strWebRoot") %>Forms/packet/packet.asp?intstudent_ID=<%=rsList("intStudent_ID")%>" <% = myAttrib %>>
										<% = rsList("Name") %></a>
										<% if rsList("intReEnroll_State") <> 7 AND rsList("intReEnroll_State") <> 15 _
											AND rsList("intReEnroll_State") <> 31 then
												if rsList("intReEnroll_State") = "129" then
													response.Write " <span style='color:red;'><b>Conditionally Enrolled</b></span> "
												else
													response.Write " <span style='color:red;'><b>Not Active</b></span> "
												end if
											end if
										%>																						 
										</td>
										<td align="center" class="TableCell">
											<% = rsList("szGrade") %>
										</td>
										<td class="TableCell" nowrap>
											<% = oFunc.FormatPhone(rsList("szHome_Phone"))  %>
										</td>
										<td align="center" class="TableCell" title="Number of Courses that have not been signed by the Guardian.">
											<% = rsList("GuardNotSign") %>
										</td>
										<td align="center" class="TableCell" title="Number of Courses that have not been signed by the Sponsor.">
											<% = rsList("SponsorNotSign") %>
										</td>
										<td align="center" class="TableCell" title="Number of Courses that have not been signed by the Instructor.">
											<% = rsList("InstructorNotSign") %>
										</td>
										<td align="center" class="TableCell" title="Number of Courses that have not been signed by the Admin.">
											<% = rsList("AdminNotSign") %>
										</td>	
										<td align="center" class="<%  if rsList("AdminMustAdmin") > 0 then  %>TableheaderRed<% else %>TableCell<%end if %>" title="Number of Courses that have a status of 'Must Amend' set by the Admin.">
											<% = rsList("AdminMustAdmin") %>
										</td>										
										<td align="center" class="<%  if rsList("Alerts") > 0 then  %>TableHeaderGrape<% else %>TableCell<%end if %>" title="Number of Sponsor Alerts a student has within their Packet.">											
											<% = rsList("Alerts") %>
										</td>
										<td align="center" class="<%  if rsList("GuardAlerts") > 0 then  %>TableHeaderTeal<% else %>TableCell<%end if %>" title="Number of Guardian Alerts a student has within their Packet.">											
											<% = rsList("GuardAlerts") %>
										</td>											
										<td class="TableCell">
											<table cellpadding=0 cellspacing=0 style="width:100%;">
												<tr>
													<td class="svplain8">
														<a href="mailto:<% = rsList("szEMAIL") %>"><% = rsList("szEMAIL") %></a>&nbsp;
													</td>
													<td align="right">
													<% 
													if instr(1,sMailList,rsList("szEMAIL") & ";") < 1 then %>	
														<input type="checkbox" value="<% = rsList("szEMAIL") %>" onChange="jfManageEmail('<% = rsList("szEMAIL") %>');" ID="Checkbox1" NAME="Checkbox1">		
													<% 
													end if 
													sMailList = sMailList & rsList("szEMAIL") & ";"	
													%>
													</td>
												</tr>
											</table>																						
										</td>
										<td class="TableCell" align="center">
											<% = ProgressText(rsList("PROGRESS_1"), rsList("PROGRESS_FULL_1"), rsList("intStudent_ID"), session.contents("intSchool_Year")) %>&nbsp;
										</td>
										<td class="TableCell" align="center">
											<% = ProgressText(rsList("PROGRESS_2"), rsList("PROGRESS_FULL_2"), rsList("intStudent_ID"), session.contents("intSchool_Year")) %>&nbsp;
										</td>
									</tr>
									<%
											rsList.MoveNext
											rowCount = rowCount + 1
										loop
									end if
									rsList.Close
									set rsList = nothing
									%>
								</table>
								<input type='hidden' id='allMailList' value="<% = sMailList %>">
							</td>
						</tr>						
						<% 							
						end if	
						%>
				</table>
			
			<%elseif Session.Value("strRole") = "GUARD" then %>
		<td>
	<table ID="Table13">
		<tr>
			<td>						
			<table border="0" cellspacing="2" cellpadding="4" style="width:100%;" ID="Table14">
				<tr>
					<td colspan="6">	
						<table ID="Table15">
							<tr>
								<td>
									<nobr><font class="svplain11"><b><i>THINGS TO DO!</i></b> </font></nobr>
									<br><br>
								</td>
								<td width="100%" class="svplain8" align="center">
									<hr size="1" color="929292">
									<b><i>the tools you need to complete the task ...</i></b>
								</td>
								<td>
									<img src="<% = Application.Value("strWebRoot") %>images/CheckList.gif">
								</td>
							</tr>
						</table>	
					</td>
				</tr>				
					<% if session.Value("intFamily_id") <> "" then 
							sqlStudent = "SELECT s.intSTUDENT_ID, s.szLAST_NAME + ',' + s.szFIRST_NAME AS Name, s.szFIRST_NAME, ss.szGrade " & _ 
										"FROM tblSTUDENT s INNER JOIN " & _ 
										"tblFAMILY f ON s.intFamily_ID = f.intFamily_ID INNER JOIN " & _ 
										"tblStudent_States ss ON s.intSTUDENT_ID = ss.intStudent_id " & _ 
										"WHERE     (f.intFamily_ID = " & Session.Value("intFamily_ID") & ") AND (ss.intSchool_Year = " & Session.Value("intSchool_Year") & ") " & _ 
										"AND ss.intReEnroll_State IN (" & Application.Contents("ActiveEnrollList") & ")  " & _ 										
										"ORDER BY Name" 
'response.write sqlStudent
							dim strStudentList
							strStudentList = oFunc.MakeListSQL(sqlStudent,intStudent_ID,Name,"")
							
							if oFunc.makeListRecordCount = 1 then
								strStudentList = replace(strStudentList,"<option value=""","")
								arStudentInfo = split(strStudentList,""">")
						%>
					<tr>
						<td class="gray">
							&nbsp;<b>Active Student</b>
						</td>
						<td class="gray">
							&nbsp;<b>Clickable Links</b>
						</td>
						<td class="gray">
							&nbsp;<b>Descriptions</b>
						</td>
					</tr>
					<tr>	
						<td align="center" class="svplain10" valign="top" rowspan="100">
							<input type="hidden" name="intStudent_ID" id="intStudent_ID" value="<% = arStudentInfo(0) %>">
							&nbsp;<b><% = arStudentInfo(1) %></b>&nbsp;
						</td>
						<%
							elseif 	oFunc.makeListRecordCount > 1 then										 
						%>
					<tr>
						<td class="TableHeader">
							&nbsp;<b>FIRST... </b><br>
							&nbsp;Select a Student
						</td>
						<td class="TableHeader">
							&nbsp;<b>THEN...</b><br>
							&nbsp;Click a Link
						</td>
						<td class="TableHeader" valign="bottom">
							&nbsp;<b>Descriptions</b>
						</td>
					</tr>
					<tr>
						<td align="center" class="svplain10" rowspan="100" valign="top">
							<select name="intStudent_ID" ID="intStudent_ID" size=5>
								<% = strStudentList %>
							</select>
						</td>
						
						<%  else %>
					<tr>						
						<td align="center" class="svplain10" colspan="2">
							There are no active students in your account for School Year
							<% = oFunc.SchoolYearRange %>.  <BR><BR>
							If you would like to change to another school year 
							click <a href="<% = Application.Value("strWebRoot") %>Admin/ChangeSchoolYear.asp">HERE</a>. <br><br>
							If you have questions  
							please contact the FPCS office at (907)742-3700.
						</td>
					</tr>						
						<%	end if 
												
							if oFunc.makeListRecordCount > 0 then 
								'*****************************
								'BKM 20-July-2003
								'Removed drop down and added a table instead (per Scott's request)
								'*****************************
								dim rsActions
								dim intCntAction
								sqlScripts = "SELECT strPath, strScript_Name, strScript_Desc, bolRequire_Validate " & _
											 "FROM   tblFPCS_Scripts " & _
											 "WHERE  ((intScript_Group_ID = 1) AND bolVersion_2_Off  is  null) and (bolAdmin_Only IS NULL) OR " & _
											 "	(bolAdmin_Only = 0) " & _
											 "ORDER BY strScript_Name"
								set rsActions = Server.CreateObject("ADODB.Recordset")
								with rsActions
									intCntAction = 0
									.CursorLocation = 3
									.Open sqlScripts, Application("cnnFPCS")'oFunc.FPCScnn
									if not .BOF and not .EOF then
										do until .EOF
											intCntAction = intCntAction + 1
											'NOTE:  Why doesn't "this.form.strPath1" work!!!  Had to change to form name..."main1"
								%>	
					<tr>
						<td class="svplain10" onclick="jfAction2(this.form,'<%=rsActions("strPath")%>',<% if rsActions("bolRequire_Validate") then response.Write "'intStudent_ID','Student'" else response.Write "'',''" %>);" style="CURSOR:pointer;" nowrap>
							<u><% = rsActions("strScript_Name")%></u>							
						</td>
						<td class="svplain10">
							<i><% = rsActions("strScript_Desc")%></i>
						</td>
					</tr>
								<%												
											.MoveNext
										loop
									end if
								end with
								
								%>
					<%If False Then %>	
					<tr>									
						<td class="svplain10" onclick="jfViewProgress('progressForm.aspx');" style="CURSOR:pointer;" nowrap>
							<u>Student Progress Report</u>							
						</td>
						<td class="svplain10">
							<i>Access to online progress reports, questionnaires and Year End Eval for Elementary Students.</i>
						</td>
					</tr>	
					<%End If %>			
					
							</table>
						</td>
						<% end if %>
					</tr>		
					<tr>					
						<td>
						<%If False Then %>
                        <input type="button" class="btSmallGray" value="Change School Year" onclick="window.location.href='<%=Application.Value("strWebRoot")%>Admin/ChangeSchoolYear.asp';" ID="Button10" NAME="Button10">
						<%End If %>
                        <input type="button" class="btSmallGray" value="Change Password" onclick="window.location.href='<%=Application.Value("strWebRoot")%>UserAdmin/ChangePassword.asp';" ID="Button11" NAME="Button1">
						<input type="button" class="btSmallGray" value="Reimbursement Form" onclick="jfReimburse();">
						<input type="button" class="btSmallGray" value="Family Manager" onclick="window.location.href='<% = Application.Contents("strSSLWebRoot") & "admin/familyManager.asp?intFamily_ID=" & session.Contents("intFamily_ID") %>';">
						<br>
		</td>
	</tr>
	<tr>
		<td>
		<br><br>
			<table border="0" cellspacing="2" cellpadding="1" ID="Table18" style="width:100%;">
				<tr>
					<td colspan="20">
					<table ID="Table19">
						<tr>
							<td>
								<nobr><font class="svplain11"><b><i>THINGS TO EXPLORE!</i></b> </font>
								<br><BR>
							</td>
							<td width="100%" class="svplain8" align="center">
								<hr size="1" color="929292">
								<b><i>find the information you need ...</i></b>
							</td>
							<td>
								<img src="<% = Application.Value("strWebRoot") %>images/Binoculars.gif">
							</td>
						</tr>
						<tr>
							<td colspan="3" align="center">
								 <table style="width:100%;">
									<tr>
										<td style="width:33%;"  class="green" align="center">
											<b><a class="linkWht2" href="<% = Application.Value("strWebRoot") %>Forms/Teachers/classSearch.asp">Search for Classes</a></b>
										</td>
										<td style="width:33%;" class="TableheaderBlue" align="center">
											<b><a class="linkWht2" href="<% = Application.Value("strWebRoot") %>forms/Teachers/teacherBiosViewer.asp">Search for Teachers</a></b>
										</td>
										<td style="width:33%;"  class="TableheaderRed" align="center">
											<b><a class="linkWht2" href="<% = Application.Value("strWebRoot") %>forms/VIS/VendorSearchEngine.asp">Search for Vendors</a></b>
										</td>
									</tr>
								 </table>
								 <br>
								 <table style="width:80%;" ID="Table20">
									<tr>
										<td style="width:50%;"  class="TableHeaderOrange" align="center">
											<a class="linkWht2" href="<% = Application.Value("strWebRoot") %>Reports/directory.asp"><b>Directories</b></a>
										</td>
										<td style="width:50%;" bgcolor=#ffffcc class="TableHeaderPurple" align="center">
											<a class="linkWht2" href="<% = Application.Value("strWebRoot") %>forms/ilp/ilpBankViewer.asp"><b>Search for ILP's</b></a>
										</td>
									</tr>
								</table>
							</td>
						</tr>
					</table>						
					</td>
				</tr>
			</table>
		</td>
	 </tr>
	</t
	<% if cint(session.Contents("intSchool_Year")) >= 2004 then %>
	<tr>
		<td>
		<br><br>
			<table border="0" cellspacing="2" cellpadding="1" ID="Table16" style="width:100%;">
				<tr>
					<td colspan="20">
					<table ID="Table17">
						<tr>
							<td>
								<nobr><font class="svplain11"><b><i>THINGS TO KNOW!</i></b> </font><br><br>
							</td>
							<td width="100%" class="svplain8" align="center">
								<hr size="1" color="929292">
								<i><b>a wealth of information ...</b></i>
							</td>
							<td>
								<img src="<% = Application.Value("strWebRoot") %>images/LightBulb.gif">
							</td>
						</tr>
					</table>						
					</td>
				</tr>
				<tr>
					<td class="TableHeaderSmall" align="center">
						<b>Student</b>
					</td>
					<td class="TableHeaderSmall" align="center" title="Number of Parent Alerts." >
						<b>Alerts</b>
					</td>
					<td class="TableHeaderSmall" align="center" title="Grade" >
						<b>Grd</b>
					</td>
					<td class="TableHeaderSmall"  align="center" title="Core Units (based on ILP's entered in the system)">
						<b>CU</b>
					</td>
					<td class="TableHeaderSmall"  align="center" title="Elective Units (based on ILP's entered in the system)">
						<b>EU</b>
					</td>
					<td class="TableHeaderSmall"  align="center" title="ASD Teacher Contract Hours">
						<b>CH</b>
					</td>
					<td class="TableHeaderSmall" align="center" title="Total Instruction Hours (based on ILP's entered in the system)">
						<b>IH</b>
					</td>
					<td class="TableHeaderSmall" align="center" title="Target Enrollment Percentage">
						<b>TE%</b>
					</td>
					<td class="TableHeaderSmall" align="center" title="Actual Enrollment Percentage">
						<b>AE%</b>
					</td>
					<td class="TableHeaderSmall" align="center" title="Target Funding (Based on Target Enrollment %)">
						<b>Tgt Fund</b>
					</td>
					<td class="TableHeaderSmall" align="center" title="Actual Funding (Based on Actual Enrollment %)">
						<b>Act Fund</b>
					</td>
					<td class="TableHeaderSmall" align="center" title="Amount of funds that have been budgeted.">
						<b>Budgeted</b>
					</td>
					<td class="TableHeaderSmall" align="center" title="Amount of budgeted funds that have been actually spent.">
						<b>Spent</b>
					</td>
					<td class="TableHeaderSmall" align="center" title="Budget Transfer AND Withdrawl Totals">
						<b>Trans</b>
					</td>
					<td class="TableHeaderSmall" align="center" title="Budget Balance (Tgt Fund + Trans - Budgeted)">
						<b>Bgt Bal</b>
					</td>
					<td class="TableHeaderSmall" align="center" title="Actual Balance (Act Fund + Trans - Spent)">
						<b>Act Bal</b>
					</td>
				</tr>
				<%
				dim rsSdt
				set rsSdt = server.CreateObject("ADODB.RECORDSET")
				rsSdt.Open sqlStudent, Application("cnnFPCS")'oFunc.FPCScnn
				do while not rsSdt.EOF
					set oBudget = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/StudentBudgetInfo.wsc"))
					'oBudget.PopulateStudentFunding oFunc.FPCSCnn, rsSdt(0), session.Contents("intSchool_Year")
					oBudget.PopulateStudentFunding Application("cnnFPCS"), rsSdt(0), session.Contents("intSchool_Year")
																	
					dblTotalTFund = dblTotalTFund + cdbl(oBudget.BudgetFunding)
					dblTotalAFund = dblTotalAFund + cdbl(oBudget.ActualFunding)
					dblTotalBCost = dblTotalBCost + cdbl(oBudget.TotalAmountBudgeted)
					dblTotalACost = dblTotalACost + cdbl(oBudget.TotalAmountSpent)
					dblTotalBBal = dblTotalBBal + cdbl(oBudget.BudgetBalance)
					dblTotalABal = dblTotalABal + cdbl(oBudget.ActualBalance)
					
				%>
				<tr>
					<td class="TableCellSmall" nowrap>
						<% = rsSdt(2) %>				
					</td>
					<td align="center" title="Number of Parent Alerts." class="<% if oBudget.ParentAlert > 0 then response.Write "TableHeaderTeal" else response.Write "TableCellSmall" end if %>">
						<% = oBudget.ParentAlert %>
					</td>
					<td class="TableCellSmall" align="center" title="Grade" nowrap>
						<% = rsSdt(3) %>				
					</td>
					<td class="TableCellSmall" align="center" title="Core Credits (based on ILP's entered in the system)" nowrap>
						<% = round(oBudget.CoreUnits,1) %>				
					</td>
					<td class="TableCellSmall" align="center" title="Elective Credits (based on ILP's entered in the system)" nowrap>
						<% = round(oBudget.ElectiveUnits,1) %>				
					</td>
					<td class="TableCellSmall" align="center" title="ASD Teacher Contract Hours" nowrap>
						<% = oBudget.ContractHours %>				
					</td>
					<td class="TableCellSmall" align="center" title="Total Instruction Hours (based on ILP's entered in the system)" nowrap>
						<% = oBudget.TotalHours %>				
					</td>
					<td class="TableCellSmall" align="center" title="Target Enrollment Percentage" nowrap>
						<% = oBudget.PlannedEnrollment %>%		
					</td>
					<td class="TableCellSmall" align="center" title="Actual Enrollment Percentage" nowrap>
						<% = oBudget.ActualEnrollment %>%			
					</td>
					<td class="TableCellSmall" align="right" title="Target Funding (Based on Target Enrollment %)" nowrap>
						$<% = formatNumber(oBudget.BasePlannedFunding,2) %>				
					</td>
					<td class="TableCellSmall" align="right" title="Actual Funding (Based on Actual Enrollment %)" nowrap>
						$<% = formatNumber(oBudget.BaseActualFunding,2) %>				
					</td>
					<td class="TableCellSmall" align="right" title="Amount of funds that have been budgeted." nowrap>
						$<% = formatNumber(oBudget.TotalAmountBudgeted,2) %>				
					</td>
					<td class="TableCellSmall" align="right" title="Amount of budgeted funds that have been actually spent." nowrap>
						$<% = formatNumber(oBudget.TotalAmountSpent,2)%>				
					</td>
					<td class="TableCellSmall" align="right" title="Budget Transfer Depostits - Withdrawls Totals" nowrap>
						$<% = formatNumber(oBudget.TotalTransfers,2) %>				
					</td>					
					<td class="TableCellSmall" align="right" title="Budget Balance (Tgt Fund + Trans - Budgeted)" nowrap>
						$<% = formatNumber(oBudget.BudgetBalance,2) %>				
					</td>
					<td class="TableCellSmall" align="right" title="Actual Balance (Act Fund + Trans - Spent)" nowrap>
						$<% = formatNumber(oBudget.ActualBalance,2) %>				
					</td>
				</tr>
				<%
						rsSdt.MoveNext
						set oBudget = nothing
					loop
					rsSdt.Close
					set rsSdt = nothing
					%>	
				<tr>
					<td class="gray" colspan="9" align="right">
						<b>Totals:</b>
					</td>
					<td class="svplain" align="right" nowrap>
						$<% = formatNumber(dblTotalTFund,2) %>				
					</td>
					<td class="svplain" align="right" nowrap>
						$<% = formatNumber(dblTotalAFund,2) %>				
					</td>
					<td class="svplain" align="right" nowrap>
						$<% = formatNumber(dblTotalBCost,2) %>				
					</td>
					<td class="svplain" align="right" nowrap>
						$<% = formatNumber(dblTotalACost,2) %>				
					</td>
					<td class="svplain" align="right" nowrap>
						&nbsp;$0.00			
					</td>
					<td class="svplain" align="right" nowrap>
						$<% = formatNumber(dblTotalBBal,2) %>				
					</td>
					<td class="svplain" align="right" nowrap>
						$<% = formatNumber(dblTotalABal,2) %>				
					</td>
				</tr>
			</table>
			<br>
			<span class="svplain8">
				<b>Not sure what all this data means? </b>Simply mouse over a column to get the
				columns description.
			</span>
			</td>
		</tr>	
	
	<% end if %>		
<% else %>
		<tr>
			<td colspan="3" class="gray">
				<b>The user you logged onto the system with 
				has a valid role as &quot;Guardian&quot; but has not been 
				completely set up in the system. Please contact the FPCS staff 
				to have this corrected if you believe your account should be
				active.</b>
			</td>
		</tr>				
<% end if %>	
	<!--		
	<tr>
		<td>
			<BR>
			<table border="0" cellspacing="2" cellpadding="4" ID="Table18" style="width:100%;">
				<tr>
					<td colspan="20">
					<table ID="Table19">
						<tr>
							<td>
								<nobr><font class="svplain11"><b><i>ASD Required Testing</i></b> </font></nobr>								
							</td>
							<td width="100%">
								<hr size="1" color="929292">
							</td>
						</tr>
					</table>						
					</td>
				</tr>
				<tr>
				<script language=javascript>				
					function jfPrintTestForm(){
						var strURL = "<%=Application.Value("strWebRoot")%>forms/PrintableForms/testingRequirement.asp";
						var testWin = window.open(strURL,"testWin","width=710,height=500,scrollbars=yes,resizable=yes");
						testWin.moveTo(0,0);
						testWin.focus();
					}
				</script>
					<td class="svplain11">
						The table below contains testing information for the current
						school year.<br>
						If you have not yet signed the ASD Required Testing Agreement<br>
						click <a href="javascript:" onclick="jfPrintTestForm();">HERE</a> to print the form.<br><br>
						<table cellspacing=1 cellpadding=2 border=0 ID="Table20">
							<tr class="TableHeader">
								<td>
									<B>Test</B>
								</td>
								<td>
									<B>Dates</B>
								</td>
								<td>
									<B>Grade Level</B>
								</td>
								<td>
									<B>Notes</B>
								</td>
							</tr>
					<%
					'sql = "select strTest_Name, strTesting_Dates, strGrade_Level,strTest_Desc " & _
					'	  "from tblTesting_Info " & _
					'	  "WHERE intSchool_Year = " & session.Contents("intSchool_Year") & _
					'	  " order by 1"
					'dim rs
					'set rsTest = server.CreateObject("ADODB.RECORDSET")
					'rsTest.CursorLocation = 3
					'rsTest.Open sql, oFunc.FPCScnn
					
					'if rsTest.RecordCount > 0 then
					'	do while not rsTest.EOF
					%>
							<tr>
								<td valign=top class="TableCell">
									<% ' = rsTest(0) %>&nbsp;
								</td>
								<td valign=top class="TableCell">
									<% ' = rsTest(1) %>&nbsp;
								</td>
								<td valign=top class="TableCell">
									<% ' = rsTest(2) %>&nbsp;
								</td>
								<td valign=top class="TableCell">
									<% ' = rsTest(3) %>&nbsp;
								</td>
							</tr>
					<%
							'rsTest.MoveNext
						'loop
					'end if
					'rsTest.Close
					'set rsTest = nothing
					%>
						</table>
						<font class="svplain8">
						Times and locations to be announced later.  
						Sponsor teachers will have access to test results<br><br>
						</font>
					</td>
				</td>
			</tr>	
		</table>
		</td>
	</tr>-->
</table>
			<% 
				end if 
				response.Write strMessageHTML
			%>
		</td>
	</tr>
</table>
&nbsp;<font class="svPlain11" color="red"><b><% = request("strMessage") %></b></font>
</form>
<%
	oFunc.CloseCN
	set oFunc = nothing
	set oList = nothing
'end if
Server.Execute(Application.Value("strWebRoot") & "Includes/footer.asp")


sub vbsTeacherHeader
%>
	<tr class="TableHeader">
			<td>
				&nbsp;<b>Student</b>
				<br>&nbsp;(click name for packet)
			</td>
			<td>
				&nbsp;<b>Grade</b>&nbsp;
			</td>	
			<td>
				&nbsp;<b>Home Phone</b>&nbsp;
			</td>
			<td title="Number of Courses that have not been signed by the Guardian." align="center">
				<b>GNS</b>
			</td>			
			<td title="Number of Courses that have not been signed by the Sponsor." align="center">
				<b>SNS</b>
			</td>	
			<td title="Number of Courses that have not been signed by the Instructor." align="center">
				<b>INS</b>
			</td>
			<td title="Number of Courses that have not been signed by the Admin." align="center">
				<b>ANS</b>
			</td>
			<td title="Number of Courses that have a status of 'Must Amend' set by the Admin." align="center">
				<b>AMA</b>
			</td>	
			<td title="Number of Sponsor Alerts a student has within their Packet.">
				&nbsp;<b>SA</b>&nbsp;
			</td>
			<td title="Number of Guardian Alerts a student has within their Packet.">
				&nbsp;<b>GA</b>&nbsp;
			</td>										
			<td>
				<table cellpadding=0 cellspacing=0 style="width:100%;">
					<tr>
						<td class="svplain8" style="color:white;">
							&nbsp;<b>Email</b>&nbsp;
						</td>
						<td align="right">
							<input type=button value="email checked" class="btSmallWhite" onclick="jfOpenMailClient();" ID="Button12" NAME="Button12">
							<input type=button value="email all" class="btSmallWhite" onclick="jfEmailAll();" ID="Button12" NAME="Button12">
						</td>
					</tr>
				</table>								
			</td>	
			<td align="center">
				&nbsp;<b>PRS 1</b>&nbsp;
			</td>		
			<td align="center">
				&nbsp;<b>PRS 2</b>&nbsp;
			</td>								
		</tr>
<%
end sub

function ProgressText(byval pShortName, pLongName, pStudentId, pSchoolYear)
	dim sColor 
	
	if pShortName = "NSIGN" then
		sColor = "blue"
	elseif pShortName = "GSIGN" then
		sColor = "orange"
	elseif pShortName = "SSIGN" or pShortName = "BGSSG" then
		sColor = "black"
	elseif pShortName = "NCHNG" then
		sColor = "red"
	elseif pShortName = "ASIGN" or pShortName  = "ASIGP" then
		sColor = "green"
	elseif pShortName&"" = "" then
		pShortName = "NSTRT"
		pLongName = "NOT STARTED"
		sColor = "red"
	end if

	ProgressText = "<span onClick=""jfViewProgress('progressForm.aspx|||STUDENT_ID=" & pStudentId & "~~SCHOOL_YEAR=" & pSchoolYear & "~~');"" style='cursor:pointer;color:" & sColor & "' title='" & pLongName & "' ><u>" & pShortName & "</u></span>"
end function
%>
