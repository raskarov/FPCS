<%@ Language=VBScript %>
<%
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Dimention Variables, make db Connection, print HTML header.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim strStudentName
dim intStudent_ID 
dim strTeacherName
dim szClass_Name
dim intClass_ID
dim sql
dim intCount
dim intContract_Guardian_ID
dim strMaterials			'contains all of the materials from classAdmin.asp
dim intLength				'used to take off ending comma in the constructed strMaterials
dim monthEnroll				'dtStudent_Enrollment will be broken into these 3 variables
dim dayEnroll
dim yearEnroll
dim oFunc		'wsc object
dim strGenericILPList1
dim strSelectGenericILP2
dim bolShowSelectILP
dim intILPGenericID			' Holds an exisiting generic ILP ID
dim bolLock
dim strDisabled

bolLock = false

bolShowSelectILP = false
strSelectGenericILP1 = "<td class=gray>&nbsp;<input type=button value='Select Exisiting ILP' class='btSmallGray' onClick='jfOpenILPBank();'></td>"
strSelectGenericILP2 = "<td></td>"	
' request("bolLateAdd") is defined in veiwClasses.asp and flags when an ILP is to be 
' added to a class that was created previously without an ILP.  It's only functionality is to
' close the browser and refresh the parent window when ilpInsert.asp is finishing.

Session.Contents("strTitle") = "ILP (Individual Learning Plan)"
Session.Contents("simpleTitle") = "ILP (Individual Learning Plan)"
Session.Contents("strLastUpdate") = "12 May 2002"
Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'' Get Student Name 
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
   call oFunc.OpenCN()
   
	if Session.Contents("blnFromClassAdmin") = true then
		'When coming from classInsert.asp via classAdmin.asp we stored most of our values
		'in session variables since they were already defined in classAdmin.asp.  This 
		'made it so we didn't have to redefine them in the redirect of classInsert.asp
		'Also at this point there are no ILP's for the class so we won't be using a sql
		'that will give us these values.
		szClass_Name = Session.Contents("szClass_Name") 
		strTeacherName = Session.Contents("strTeacherName") 
		intClass_ID = Session.Contents("intClass_ID")	
		intContract_Guardian_ID = Request("intContract_Guardian_ID")
		' we use this session variable only with the ILPBank.
		session.Contents("strParams") = "szClass_Name=" & szClass_Name & "&intClass_ID=" & intClass_ID & "&strTeacherName=" & _
				   strTeacherName & "&&bolAddILPtoExistingContract=true" & _
				   "&intContract_Guardian_id=" & request("intContract_Guardian_id") & _
				   "&maxCost=" & request("maxCost") 
	elseif Request.QueryString("szClass_Name") <> "" then
		'When coming from viewClasses.asp
		szClass_Name = Request.QueryString("szClass_Name") 
		intClass_ID = Request.QueryString("intClass_ID")
		'ilpInsert.asp uses Session.Contents("intClass_ID") to insert/update. 
		Session.Contents("intClass_ID") = intClass_ID
		strTeacherName = Request.QueryString("strTeacherName")
		Session.Contents("intVendor_ID") = Request.QueryString("intVendor_ID")
		
		' we use this session variable only with the ILPBank.
		session.Contents("strParams") = "szClass_Name=" & szClass_Name & "&intClass_ID=" & intClass_ID & "&strTeacherName=" & _
				   strTeacherName & "&bolAddILPtoExistingContract=true" & _
				   "&intContract_Guardian_id=" & request("intContract_Guardian_id")
	else
		'Redirect to the home page if this URL is entered direcly
		Server.Transfer Application.Value("strMiniRoot") & "default.asp"
	end if 
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'' Get Vendor Name 
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	if Session.Contents("szVendor_Name") <> "" then
		strTeacherName = Session.Contents("szVendor_Name")
	end if
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'' Get Student Name 
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	intStudent_ID = Session.Contents("intStudent_ID")

	if Session.Contents("strStudentName") <> "" then
		strStudentName = "for " & Session.Contents("strStudentName")
	end if 
        Dim rsSyllabus
		set rsSyllabus = server.CreateObject("ADODB.RECORDSET")
	
	if request("intILP_ID") <> "" or request("intILP_ID_Generic") <> "" then	
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'' This section will get ILP info for both exisiting Genric and  
		'' non-Generic ILP's depending on the incoming request.
		'' This handle's our explict ilp requests
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
		dim strILPTable
		dim sqlILPID
		dim strILPFields
		dim strFROM 
		
		
		if request("intILP_ID") <> "" then			
			strILPFields = "i.intILP_ID as genILP,i.bolILP_Bank,i.szILP_Name,i.szUser_Create," & _
				"i.dtStudent_Enrolled, i.intShort_ILP_ID,  " & _
				" trefPOS_Subjects.szSubject_Name, " & _
				"c.intPOS_Subject_ID,c.intInstructor_ID,i.bolApproved,i.bolGradingScale," & _
				"i.bolSponsor_Approved,i.szILP_Additions, tblILP_Generic.szILP_Additions AS Teacher_Additions,"  & _
				" '1' as Enrolled, c.intContract_Status_ID, i.GuardianStatusId, i.SponsorStatusId, i.InstructorStatusId,i.AdminStatusId, " 
			strILPTable = " tblILP " 
			
			strFROM  =  " FROM tblClasses c INNER JOIN " & _ 
						" tblILP i ON c.intClass_ID = i.intClass_ID LEFT OUTER JOIN " & _ 
						" tblILP_Generic ON c.intClass_ID = tblILP_Generic.intClass_ID LEFT OUTER JOIN " & _ 
						" trefPOS_Subjects ON c.intPOS_Subject_ID = trefPOS_Subjects.intPOS_Subject_ID LEFT OUTER JOIN " & _ 
						" tblINSTRUCTOR ins ON c.intInstructor_ID = ins.intINSTRUCTOR_ID " 
			sqlILPID = request("intILP_ID")
		else
			strILPFields = "i.bolILP_Bank,i.szILP_Name,c.intPOS_Subject_ID as intPOS_Subject_ID," & _
						   "c.intInstructor_ID, i.intILP_ID as genILP, i.szUser_Create, i.szILP_Additions AS Teacher_Additions,i.bolGradingScale," & _
						   "(SELECT COUNT(i2.intILP_ID) AS total " & _ 
						   " FROM tblILP i2 " & _ 
					       " WHERE      i2.intClass_ID = i.intClass_ID) AS Enrolled, c.intContract_Status_ID, " 
			strILPTable = " tblILP_Generic "
			strFROM = 	" FROM   tblINSTRUCTOR ins RIGHT OUTER JOIN " & _
						"        tblClasses c RIGHT OUTER JOIN " & _
						"        tblILP_Generic i ON c.intClass_ID = i.intClass_ID ON ins.intINSTRUCTOR_ID = c.intInstructor_ID " 
			sqlILPID = request("intILP_ID_Generic")	
			
			' Teachers may replace an existing generic ILP that is tied to a class
			' with a generic ILP from the ILP bank. If this happens we want to 
			' keep the same generic ilp id and just overwrite the current data
			' intExisitingGenericILP helps us make sure we do not create
			' a new generic ilp if one already exisits for a class
			if request("intExisitingGenericILP") <> "" then
				intILPGenericID = request("intExisitingGenericILP")
			else
				intILPGenericID = request("intILP_ID_Generic")	
			end if
			
			bolShowSelectILP = true
		end if

		'We need to populate the ILP info since we were given an ILP ID
		set rsILP = server.CreateObject("ADODB.RECORDSET")
		rsILP.CursorLocation = 3
		
		'bkm 2-oct-2002
		'multiple outer joins - allows us to grab from both tblProgramOfStudies and trefPOS_Subjects.  For those ILP's
		'that have a ShortForm that do NOT have an associated tblProgramOfStudies we use the trefPOS_Subjects in the
		'display to the user indicating which ShortForm the ILP was based on
		
		if szClass_Name = "" then
			' Only select class name and id if not already provided.
			strGetClassInfo = " c.szClass_Name, i.intClass_ID,"
		end if 
		sql =	"SELECT     " & strILPFields & strGetClassInfo & " i.intSemester, i.decCourse_Hours, i.szCurriculum_Desc, i.szGoals, i.szRequirements,  " & _
				"i.szTeacher_Role, i.szStudent_Role, i.szParent_Role, i.szEvaluation, i.szEvaluationFrequency, i.bolPass_Fail,i.szOther_Grading, ins.szFIRST_NAME, ins.szLAST_NAME,  " & _
				"i.intContract_Guardian_id " & _
				strFROM & " WHERE     (i.intILP_ID = " & sqlILPID & ")"	

		rsILP.Open sql,Application("cnnFPCS")'oFunc.FPCScnn
		
		rsSyllabus.CursorLocation = 3
        Dim sqlSyllabus
        sqlSyllabus = "SELECT [syllabusId],[intILP_ID],[weekNo],[dtStart],[dtEnd],[szDescription] FROM [dbo].[tblSyllabus] where intILP_ID=" & sqlILPID & " order by weekNo"
		rsSyllabus.Open sqlSyllabus,Application("cnnFPCS")'oFunc.FPCScnn

		intCount = 0
		'This for loop dimentions and defines all the columns we selected in sqlClass
		'and we use the variables created here to populate the form.
		for each item in rsILP.Fields
			execute("dim " & rsILP.Fields(intCount).Name)
			execute(rsILP.Fields(intCount).Name & " = item")		
			intCount = intCount + 1
		next  
		
		rsILP.Close
		set rsILP = nothing
		
		'Seperate student enrollment date for select box populating
		monthEnroll = datePart("m",dtStudent_Enrolled)
		dayEnroll = datePart("d",dtStudent_Enrolled)
		yearEnroll = datePart("yyyy",dtStudent_Enrolled)
		
		intCount = 0		
	else
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'' If no ILP ID has been given we check to see if a generic ILP exists in the
		'' the system based on the class_id
		'' Happens when a guardian is adding a class.
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

		' Check for Template ILP
		dim strTeacherJoin
		set rsTempILP = server.CreateObject("ADODB.RECORDSET")
		rsTempILP.CursorLocation = 3
		
		sql = "SELECT intInstructor_ID, intGuardian_ID, intPOS_Subject_ID " & _
				"FROM tblClasses " & _
				"WHERE (intClass_ID = " & intClass_ID & ")"
		
		rsTempILP.Open sql, Application("cnnFPCS")'oFunc.FPCScnn			
		
		' Do we have an ILP Template?
		if rsTempILP.RecordCount > 0 then
			if rsTempILP(0) <> "" then
				' Found a teachers template
				' Class instructed by a Contract Teacher so get ILP from the generic ilp table
				strILPTable = "tblILP_Generic"
				intInstructor_ID = rsTempILP(0)
				strTeacherJoin = "tblInstructor ins ON (c.intInstructor_Id = ins.intInstructor_Id) " 
			elseif rsTempILP(1) <> "" then
				' found a guardians template
				' Class instructed by a parent so get ILP from the ilp table
				strILPTable = "tblILP"
				strTeacherJoin = "tblGuardian ins ON (c.intGuardian_ID = ins.intGuardian_ID) " 
			end if
			intPOS_Subject_ID = rsTempILP("intPOS_Subject_ID")
		end if
		
		rsTempILP.Close
						
		if strTeacherJoin <> "" then
			' We found an ILP so time to populate variables
			 
			sql = "select top 1 i.intILP_ID as intTempILPID, c.szClass_Name,i.intClass_id,i.intSemester,i.decCourse_Hours,i.szCurriculum_Desc," & _
				  "i.szGoals,i.szRequirements,i.szTeacher_Role,i.szStudent_Role,i.szParent_Role,i.szEvaluation,i.szEvaluationFrequency," & _
				  " ins.szFirst_Name, ins.szLast_Name, i.bolGradingScale,c.intContract_Status_ID " & _
				  "FROM " & strILPTable & " i inner join tblClasses c ON i.intClass_ID = c.intClass_ID LEFT OUTER JOIN " & _
				  strTeacherJoin & _
				  "WHERE c.intClass_ID = " & intClass_ID
				  
			rsTempILP.Open sql,Application("cnnFPCS")'oFunc.FPCScnn
		
			intCount = 0
			
			if rsTempILP.RecordCount > 0 then	
				'This for loop dimentions and defines all the columns we selected in sqlClass
				'and we use the variables created here to populate the form.
				for each item in rsTempILP.Fields
					execute("dim " & rsTempILP.Fields(intCount).Name)
					execute(rsTempILP.Fields(intCount).Name & " = item")		
					intCount = intCount + 1
				next  	
			else
				' A contract or schedule exists without an ILP
				bolShowSelectILP = true
			end if		
			rsTempILP.Close						
		else
			' No ILP's exisit for the class so we generate a list of existing ILP's
			response.Write "ERROR " 
			response.End
			bolShowSelectILP = true
		end if
		set rsTempILP = nothing
	end if 
	
	if request("bolHideAddBank") <> "" then
		szILP_NAME = ""
		bolILP_BANK  = ""
		intPOS_SUBJECT_ID = ""
	end if 
	
	bolLock = false
	
	' Determine if we need to lock the ILP 
	if ((InstructorStatusId & "" = "1" or  SponsorStatusId & "" = "1" or _
		GuardianStatusId & "" = "1" or AdminStatusId & "" = "1" or _
		intContract_Status_ID & "" = "5") and request("bolHideAddBank") = "") or _
		(oFunc.LockYear and ucase(session.Contents("strRole")) <> "ADMIN") then
		
		' don't need to lock if course is taught by a guardian and the
		' sponsor and guradian have not both signed  smb 4-24-2006
		if intInstructor_ID & "" = "" and SponsorStatusId & "" <> "1" or _
		   GuardianStatusId & "" <> "1" and AdminStatusId & "" <> "1" then
			bolLock = false
		else
			bolLock = true
			strDisabled = " disabled "
		end if
	end if
		
	
%>
<form action="ILPInsert.asp" method=Post name=main>
<input type=hidden name="changed" value="">
<input type=hidden name="intContract_Guardian_ID" value="<%=intContract_Guardian_ID%>">
<input type=hidden name="intGuardian_ID" value="<%=request("intGuardian_ID")%>">
<input type=hidden name="intVendor_ID" value="<%=request("intVendor_ID")%>">
<input type=hidden name="bolAddILPtoExistingContract" value="<% = Request.QueryString("bolAddILPtoExistingContract")%>">
<input type=hidden name="bolLateAdd" value="<% = request("bolLateAdd") %>">
<input type=hidden name="maxCost" value="<% = request("maxCost") %>">
<input type=hidden name="hdnHrsChanged" value="">
<table width=100%>
	<tr>	
		<Td class=yellowHeader>
		<%
			dim strILPSFdesc
			'strILPSFdesc = "<nobr>(Class Name:<i>" & szClass_Name & "</i>)</nobr>"
	
		%>
				&nbsp;<B>ILP  <% = strStudentName %>&nbsp;&nbsp;<% = strILPSFdesc%></B>
		</td>
	</tr>
	<tr>
		<td bgcolor=f7f7f7 style="width:100%;">
		<script language=javascript>			
			function jfOpenILPBank(){
				var winILPBank;
				var url = "ilpBankViewer.asp?isPopUp=<%=request("isPopUp")%>&fromMain=true&";
				url += "bolLateAdd=<%=request("bolLateAdd")%>&intExisitingGenericILP=<%=request("intILP_ID_Generic")%>";
				winILPBank = window.open(url,"winILPBank","width=710,height=500,scrollbars=yes,resize=yes,resizable=yes");
				winILPBank.moveTo(0,0);
				winILPBank.focus();
			}
			
			function jfValidateForm(pForm){
				if (!pForm.bolPass_Fail.checked && !pForm.bolGradingScale.checked && !pForm.bolOther_Grading.checked) {
					alert("You must select an Evaluation/Grading method before you can save the ILP.");
					return false;
				}else if(pForm.bolOther_Grading.checked && pForm.bolOther_Grading.value == ""){
					alert("You must provide an explaination in box 7 when 'other' is checked under 'Evaluatin and Grading'.");
					return false;
				}else{
					pForm.submit();
				}
			}
		</script>
			<% if bolFromILPBank <> true then %>
			<table cellpadding="3" style="width:100%;">
				<tr>		
					<% if bolShowSelectILP = true and not bolLock then response.Write strSelectGenericILP1 %>
					<td class=gray>
							Instructor
					</td>
					<td class=gray nowrap>
							Class Name
					</td>
					<td class=gray nowrap align="center">
							School Year
					</td>		
					<!--<td class=gray>
							&nbsp;Semester&nbsp;
					</td>-->
					<td align="center" class=gray title="Number of hours this course contributes to core hours.">
							Course Hrs
					</td>	
					<% if strStudentName <> "" then %>	
					<td class=gray >
						Class Enroll Date
					</td>		
					<% end if %>											
				</tr>
				<tr>			
					<% if bolShowSelectILP = true and not bolLock then response.Write strSelectGenericILP2 %>		
					<td class=svplain8>
							<% = strTeacherName %>
					</td>	
					<td class=svplain8>
							<% = szClass_Name %>
					</td>
					<td align=center class=svplain8>
						<input type=hidden name="sintSchool_Year"  value="<%= session.Contents("intSchool_Year")  %>">
						<%= session.Contents("intSchool_Year") %>
					</td>
					<td align=center class=svplain8>
					<% if (strDisabled & "" <> "" and oFunc.IsAdmin) or strDisabled & "" = "" then %>					
					<input type=text name="decCourse_Hours" value="<% if (decCourse_Hours = "" or decCourse_Hours = 0)and Session.Contents("intShort_ILP_ID") <> "" then response.Write oFunc.CourseInfo(Session.Contents("intShort_ILP_ID"),"2") else response.Write decCourse_Hours %>" size=4 maxlength=4 onChange="jfChanged();this.form.hdnHrsChanged.value='true';">					
					<% else %>
						<% = decCourse_Hours%>
					<% end if %>
					</td>	
					<% if strStudentName <> "" then %>	
					<td class=svplain8  align=center >
						<% if dtStudent_Enrolled <> "" then 
							response.Write dtStudent_Enrolled
						%>
							<input type=hidden name="dtStudent_Enrolled" value="<% = dtStudent_Enrolled%>">
						<%
						   else
						%>
						   <input type=hidden name="dtStudent_Enrolled" value="<% = date()%>">
						   <% = date()%>				
						<%
						   end if
						%>
					</td>										
					<% end if %>					
				</tr>
			</table>		
			<% else %>
			<table ID="Table1">
				<tr>		
					<td class=gray>
							&nbsp;ILP Name
					</td>
					<td class=gray>
							&nbsp;Subject&nbsp;
					</td>											
				</tr>
				<tr>				
					<td class=svplain10>
						<input type=text size=20 maxlength=64 name="szILP_Name" value="<%=szILP_Name%>">														
					</td>
					<td>
						<select name="intPOS_Subject_ID">
								<option value="12">
						<%
							sql = "select intPOS_Subject_ID, lower(szSubject_Name) szSubject_Name from trefPOS_Subjects order by szSubject_Name"									
							response.Write oFunc.MakeListSQL(sql,"intPOS_Subject_ID","szSubject_Name",request("intPOS_Subject_ID"))
						%>
						</select>
					</td>			
				</tr>
			</table>			
			<% end if %>	
			<table style="width:100%;" ID="Table4">
				<tr>
					<td class=gray  style="width:100%;">
							&nbsp;<B>1. Description of the course including methods needed ...</B>
					</td>									
				</tr>
				<tr>
					<td style="width:100%;" class="svplain8">
					<% if not bolLock then %>
						<textarea name="szCurriculum_Desc" onKeyDown="jfMaxSize(7000,this);"  style='width:100%;' rows='4' wrap='virtual' onFocus="this.rows=12;" onBlur="this.rows=4;" ID="Textarea1"><%=szCurriculum_Desc%></textarea>
					<% else 
						response.Write szCurriculum_Desc					
					end if %>
					</td>
				</tr>
			</table>			
			<table width=100% ID="Table5">
				<tr>
					<td class=gray style="width:50%;">
							<table cellpadding=3 cellspacing=0 class=gray ID="Table6">
								<tr>
									<td valign=top><b>2.</b></td>
									<td>
										<b>Scope and sequence</b><br>
											(add entire syllabus plan. ex. wk 1 - 8/2/12 to 8/6/12 Coniferous Tree Study)
									</td>
								</tr>
							</table>
					</td>
					<td class=gray style="width:50%;">
						<table cellpadding=3 cellspacing=0 class=gray ID="Table7">
								<tr>
									<td valign=top><b>3.</b></td>
									<td>
										<b>Activities student will be involved</B>
									</td>
								</tr>
							</table>
					</td>									
				</tr>
				<tr>
					<td valign="top" class="svplain8">
                    <table>
                    <tr>
                    <td class="gray">Week</td>
                    <td class="gray">Start</td>
                    <td class="gray">End</td>
                    <td class="gray">Description</td>
                    </tr>
                    <% If rsSyllabus<>Empty Then
                    do while not rsSyllabus.EOF %>
                    <% if bolLock Then %>
                    <tr>
                    <td>
                    <%=rsSyllabus("WeekNo")%>
                    </td>
                    <td><%=rsSyllabus("dtStart")%></td>
                    <td><%=rsSyllabus("dtEnd")%></td>
                    <td><%=rsSyllabus("szDescription")%></td>
                    </tr>
                    <%Else %>
                    <tr>
                    <td><input type="hidden" name="syllabusId" value='<%=rsSyllabus("syllabusId") %>' />
                    <input type="text" name="WeekNo" value='<%=rsSyllabus("WeekNo")%>' maxlength="2" size="2" />
                    </td>
                    <td><input type="text" class="date" name="dtStart" value='<%=rsSyllabus("dtStart")%>' maxlength="10" size="10" />
                    </td>
                    <td><input type="text" class="date" name="dtEnd" value='<%=rsSyllabus("dtEnd")%>' maxlength="10" size="10" />
                    </td>
                    <td><input type="text" name="szDescription" value='<%=rsSyllabus("szDescription")%>' maxlength="200" size="20" /></td>
                    </tr>

                    <%End If
                    rsSyllabus.MoveNext()
                    Loop
                    rsSyllabus.Close()
                    Set rsSyllabus = Nothing
                    End If
                     %>
                     <%if not bolLock and strILPTable<>" tblILP_Generic " Then %>
                    <tr>
                    <td><input type="hidden" name="syllabusId" value='new' />
                    <input class="syllabus" type="text" name="WeekNo" value='' maxlength="2" size="2" />
                    </td>
                    <td><input type="text" class="date syllabus" name="dtStart" value='' maxlength="10" size="10" />
                    </td>
                    <td><input type="text" class="date syllabus" name="dtEnd" value='' maxlength="10" size="10" />
                    </td>
                    <td><input class="syllabus" type="text" name="szDescription" value='' maxlength="200" size="20" /></td>
                    </tr>


                     <%End If %>
                    </table>
                     <%if not bolLock Then %>
                     <%End If %>
				 	<%' if not bolLock then %>
						<%'<textarea name="szGoals" style="width:100%;" rows=4 wrap=virtual onFocus="this.rows=12;" onBlur="this.rows=4;" onKeyDown="jfMaxSize(7000,this);" ID="Textarea2"><%=szGoals% ></textarea>%>
					<%' else 
						'response.Write szGoals					
					'end if %>
					</td>
					<td valign="top" class="svplain8">
					<% if not bolLock then %>
						<textarea name="szRequirements"  style="width:100%;" rows=4 wrap=virtual onFocus="this.rows=12;" onBlur="this.rows=4;" onKeyDown="jfMaxSize(7000,this);" ID="Textarea3"><%=szRequirements%></textarea>
					<% else 
						response.Write szRequirements					
					end if %>
					</td>
				</tr>
			</table>
			<table width=100% ID="Table8">
				<tr>
					<td class=gray style="width:33%;"> 
							&nbsp;<b>4. Standards: Common core or GLE</B>
					</td>
					<td class=gray style="width:33%;">
							&nbsp;<b>5. Materials, Resources and</B>
					</td>		
					<td class=gray style="width:33%;">
							<table cellpadding=3 cellspacing=0 class=gray ID="Table9">
								<tr>
									<td valign=top><b>6.</b></td>
									<td>
										<b>Role of Parent/Teacher/Vendor/any aditional responcibilities of the student</b>
									</td>
								</tr>
							</table>
					</td>												
				</tr>
				<tr>
					<td style="width=33%;" valign="top" class="svplain8">
					<% if not bolLock then %>
						<textarea name=szTeacher_Role style="width:100%;" onFocus="this.rows=12;" onBlur="this.rows=4;" cols=22 rows=3 wrap=virtual onKeyDown="jfMaxSize(7000,this);" ID="Textarea4"><%=szTeacher_Role%></textarea>
					<% else 
						response.Write szTeacher_Role					
					end if %>
					</td>
					<td style="width=33%;" valign="top" class="svplain8">
					<% if not bolLock then %>
						<textarea name=szStudent_Role style="width:100%;" onFocus="this.rows=12;" onBlur="this.rows=4;" cols=22 rows=3 wrap=virtual onKeyDown="jfMaxSize(7000,this);" ID="Textarea5"><%=szStudent_Role%></textarea>
					<% else 
						response.Write szStudent_Role					
					end if %>
					</td>
					<td style="width=33%;" valign="top" class="svplain8">
					<% if not bolLock then %>
						<textarea name=szParent_Role style="width:100%;" onFocus="this.rows=12;" onBlur="this.rows=4;" cols=22 rows=3 wrap=virtual onKeyDown="jfMaxSize(7000,this);" ID="Textarea6"><%=szParent_Role%></textarea>
					<% else 
						response.Write szParent_Role					
					end if %>
					</td>
				</tr>
			</table>	

			<table width=100% ID="Table10">
				<tr>	
					<Td colspan=3>
						<font class=svplain11>
							<b><i>Evaluation and Grading</I></B> 
						</font>
						<font class=svplain8> (Please check method(s) that apply)</font>
					</td>				
				</tr>
				<tr>
					<%
						gradeDisabled = strDisabled
						
						if gradeDisabled <> "" then 
							if not bolPass_Fail and not bolGradingScale and szOther_Grading&"" = "" then
								gradeDisabled = ""
							end if
						end if										
					%>
					<td class=gray rowspan=2 valign=top>
						<nobr>&nbsp;<input <% = gradeDisabled %> type=checkbox name="bolPass_Fail" <% if bolPass_Fail&"" <> "" and bolPass_Fail <> 0 then response.Write(" checked ")%> ID="Checkbox1" value='1'>Pass/Fail&nbsp;</nobr>
					</td>
					<td class=gray>
						&nbsp;<input <% = gradeDisabled %> type=checkbox name="bolGradingScale" <% if bolGradingScale then response.Write(" checked ")%> ID="Checkbox2" value="1">Grading Scale 
					</td>	
					<td class=gray>
						&nbsp;<input <% = gradeDisabled %> type=checkbox name="bolOther_Grading" <% if szOther_Grading&"" <> "" then response.Write(" checked ")%> ID="Checkbox3" value='1'>Other  (if checked you <b>MUST</b> explain why in box #7.)
					</td>										
				</tr>
				<tr>					
					<td>
						<table ID="Table11">
							<tr>
								<td class=gray nowrap>
									&nbsp;A = 
								</td>
								<td nowrap class="svplain">
									90% to 100%
								</td>							
							</tr>
							<tr>
								<td class=gray nowrap>
									&nbsp;B = 
								</td>
								<td nowrap class="svplain">
									80% to 89%
								</td>									
							</tr>
							<tr>
								<td class=gray nowrap>
									&nbsp;C = 
								</td>
								<td nowrap class="svplain">
									70% to 79%
								</td>								
							</tr>
							<tr>
								<td class=gray nowrap>
									&nbsp;D = 
								</td>
								<td nowrap class="svplain">
									60% to 69%
								</td>								
							</tr>
							<tr>
								<td class=gray nowrap>
									&nbsp;F = 
								</td>
								<td nowrap class="svplain">
									0% to 59%
								</td>								
							</tr>
						</table>
					</td>
					<td class=svplain8 valign="top">
						<b>7. Explain:</b><br>
					<% if gradeDisabled = "" then %>	
						<textarea rows=7 style="width:100%;" onFocus="this.rows=12;" onBlur="this.rows=7;" name="szOther_Grading" ID="Textarea7"><%=szOther_Grading%></textarea>
					<% else 
						response.Write szOther_Grading					
					end if %>
					</td>
				</tr>
				<tr>
					<td class=gray colspan=3>
							<table cellpadding=3 cellspacing=0 class=gray ID="Table12">
								<tr>
									<td valign=top><b>8.</b></td>
									<td>
										<b>What will be evaluated? 
										</b>(worksheets, tests, class participation, daily work, logs, attendance, etc.)<br>
										<b>What will be the measurable outcomes?</b> (Logs are not permissable without measurable goals included, i.e. run 1.5 hours without stopping, run 4 days per week, or 40 minutes without stopping )
									</td>
								</tr>
							</table>
					</td>
				</tr>
				<tr>
					<td valign=top colspan=3 class="svplain8">
					<% if not bolLock then %>
						<textarea name="szEvaluation" style="width:100%;" onFocus="this.rows=12;" onBlur="this.rows=5;" cols=60 rows=5 wrap=virtual onKeyDown="jfMaxSize(7000,this);"><% = szEvaluation %></textarea>
					<% else 
						response.Write szEvaluation					
					end if %>
					</td>
				</tr>
				<tr>
					<td class=gray colspan=3>
							<table cellpadding=3 cellspacing=0 class=gray ID="Table12">
								<tr>
									<td valign=top><b>9.</b></td>
									<td>
										<b>How often will evaluation marks to be teacher of record? 
										</b>(the end of weeks 4, 8, 12, and 18)<br>
										
									</td>
								</tr>
							</table>
					</td>
				</tr>
				<tr>
					<td valign=top colspan=3 class="svplain8">
					<% if not bolLock then %>
						<textarea name="szEvaluationFrequency" style="width:100%;" onFocus="this.rows=12;" onBlur="this.rows=5;" cols=60 rows=5 wrap=virtual onKeyDown="jfMaxSize(500,this);"><% = szEvaluationFrequency %></textarea>
					<% else 
						response.Write szEvaluationFrequency					
					end if %>
					</td>
				</tr>
			</table>	
			<% 
			'response.Write bolFromILPBank & " - '" & bolILP_Bank  &  "' - " & intInstructor_ID & " - " & intStudent_ID
			'if (bolILP_Bank <> true or bolILP_Bank & "" = "")  _
			'	and ((intInstructor_ID <> "" and intStudent_ID = "") or _
			'	 (intInstructor_ID & "" = "" and intStudent_ID <> "")) then 
			'if not bolILP_Bank or bolILP_Bank & "" = "" or intGeneric_ILP_ID = "" then
			' bolFromILPBank <> true is when the user has NOT selected an ILP from the bank
			' bolILP_Bank is from the checkbox on this form. If it has already been checked
			' we do not want to show the Add to Bank table
			' If we have an instructor id and a student id this means a guradian is 
			' adding an existing class created by an ASD instructor and the guardian
			' will not be able to add this ilp to the bank the teacher will have to.
			' If we have a intContract_Guardian_ID and a Student ID then if the other 
			' conditions are true then we DO want to show the add to bank since we 
			' have a parent instructed class that has not been added to the bank
			' Assign intPOS_Subject_ID if not defined
			if intPOS_Subject_ID = "" and session.Contents("intPOS_Subject_ID") <> "" then
				intPOS_Subject_ID = session.Contents("intPOS_Subject_ID")
			elseif intPOS_Subject_ID = "" and session.Contents("intPOS_Subject_ID_from_class") <> "" then
				intPOS_Subject_ID = session.Contents("intPOS_Subject_ID_from_class")
			end if 						
			
			if ((bolLock or szILP_Additions <> "") and strILPTable = " tblILP ") or _
			(intTempILPID <> "" ) then			
			' Show this section when the ILP is locked whether due to the fact that
			' we have a new ilp being copied from a temp ILP or the students ilp is locked
			%>
			<br>
			<table width=100% ID="Table13">
				<tr>
					<td class=gray style="width:33%;"> 
							&nbsp;<b>Guardian ILP Modifications:</B> After an ILP has been signed 
							the signed ILP can not be altered.  Please
							enter additional information in the box below.
					</td>
				</tr>
				<tr>
					<td>
						<textarea name="szILP_Additions" style="width:100%;" onFocus="this.rows=12;" onBlur="this.rows=4;" cols=22 rows=3 wrap=virtual onKeyDown="jfMaxSize(4000,this);" ID="Textarea8"><%=szILP_Additions%></textarea>												
						<% if bolLock then %>
							<input type=hidden name="bolILP_ADD" value="true">
								<% if request("isPopUp") = "" and strILPTable <> " tblILP " then 
									' We need to create a new ILP from an existing ILP
									' Only want to do it IF the existing ILP is from tblILP_Generic
								%>
								<input type=hidden name="intILP_ID" value="<% = sqlILPID & intTempILPID %>" ID="Hidden3">
								<% end if %>					 
						<% end if %>
					</td>
				</tr>
			</table>
			<% end if  
			
			if (Enrolled > 0 or Teacher_Additions <> "") and intInstructor_ID <> "" and request("bolHideAddBank") = "" then
			%>
			<br>
			<table width=100% ID="Table14">
				<tr>
					<td class=gray style="width:33%;"> 
							&nbsp;<b>Instructor ILP Modifications:</B> 
							<% if request("intILP_ID") = "" then %>
							*Please Note* There are currently 
							<% = Enrolled %> student(s) enrolled in this class. Since this ILP has already been
							accepted in its present form please make any modifications in the box below. These
							changes will also appear on the students version of the ILP.				
							<% else %>
							*Please Note* After enrolling your student in this class the 
							instructor made the following modifications to the original ILP ...
							<% end if %>
					</td>
				</tr>
				<tr>
					<td class="svplain8">
						<% if request("intILP_ID") <> "" then %>
							<% = Teacher_Additions %>
							<% elseif request.QueryString("bolHideAddBank") = "" then %>
							<textarea name="Teacher_Additions" style="width:100%;" onFocus="this.rows=12;" onBlur="this.rows=4;" cols=22 rows=3 wrap=virtual onKeyDown="jfMaxSize(4000,this);" ID="Textarea9"><%=Teacher_Additions%></textarea>
							<input type=hidden name="bolILP_TEACHER_ADD" value="true" ID="Hidden4">	
							<% if request("isPopUp") = "" then 
								' We need to create a new ILP from an existing ILP
							%>	
							<input type=hidden name="intILP_ID_Generic" value="<% =  intILPGenericID %>" ID="Hidden5">
							<% end if
						end if %>
					</td>
				</tr>
			</table>
			<% end if %>			
			<table ID="Table2">
				<tr>	
					<Td colspan=3>
						<font class=svplain11>
							<b><i>ILP Bank Information</I></B> 
						</font>
					</td>
				</tr>	
				<tr>
					<td class=gray>
						&nbsp;Add ILP to Bank?&nbsp;
					</td>
					<td class=svplain10>
						<b>Yes</b><input type=checkbox name="bolILP_Bank" <% if bolILP_Bank = true then response.Write " checked " end if %> ID="Checkbox4">
					</td>
				</tr>
				<tr>		
					<td class=gray>
						&nbsp;ILP Name:&nbsp;
					</td>
					<td class=svplain10>
						<input type=text size=64 maxlength=64 name="szILP_Name" value="<%=szILP_Name%>" ID="Text1">											
					</td>
				</tr>
			</table>
			<input type=hidden name="intPOS_Subject_ID" value="<%=intPOS_Subject_ID%>" ID="Hidden1">																
		</td>
	</tr>
</table>
<%	
' first check to see if year is locked
if not oFunc.LockYear then								
		IF request("isPopUp") = "" then 	
			'New ILP is being Created		
	%>
		<input type='<% if bolLock then %>submit<%else%>button<% end if %>' value="SAVE (goods & services page next)" class="NavSave"  <% if not bolLock then %> onclick="jfValidateForm(this.form);" <% end if %>>
	<% 
	ELSE
			'existing ILP is being modified
	%>
		<input type=button value="Close without saving" onClick="window.opener.focus();window.close();" class="NavLink">	
		<input type=button value="SAVE" class="NavSave" onclick="jfValidateForm(this.form);">
		<input type=hidden name="edit" value="yes">
		<input type=hidden name="intILP_ID" value="<% = request("intILP_ID") %>">
		<input type=hidden name="intILP_ID_Generic" value="<% =  intILPGenericID %>">
	<% end if 
end if
%>
</form>
<%If Not bolLock Then %>
<script type="text/javascript">
    $(document).ready(function () {
        $(".date").datepicker({ showAnim: 'fade', numberOfMonths: 1, showOn: "focus", changeMonth: true, maxDate: '+1y' });
        $('.syllabus').change(function () { AddRow(this); });
        function AddRow(elem) {
            var tr = $(elem).parent().parent();
            var newTr = tr.parent().append('<tr> \
                    <td><input type="hidden" name="syllabusId" value="new" /> \
                    <input class="syllabus" type="text" name="WeekNo" value="" maxlength="2" size="2" /> \
                    </td> \
                    <td><input type="text" class="date syllabus" name="dtStart" value="" maxlength="10" size="10" /> \
                    </td> \
                    <td><input type="text" class="date syllabus" name="dtEnd" value="" maxlength="10" size="10" /> \
                    </td> \
                    <td><input class="syllabus" type="text" name="szDescription" value="" maxlength="200" size="20" /></td> \
                    </tr>');
            $('input[type="text"]', tr).removeClass('syllabus').unbind('change');
            $('.syllabus', newTr).change(function () { AddRow(this); });
            $(".date", newTr).datepicker({ showAnim: 'fade', numberOfMonths: 1, showOn: "focus", changeMonth: true, maxDate: '+1y' });
        }
    });
</script>
<%End If %>
<%
call oFunc.CloseCN()
set oFunc = nothing

Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")
%>

