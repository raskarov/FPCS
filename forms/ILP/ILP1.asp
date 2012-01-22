<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		ILP1.asp
'Purpose:	Forms for the initializing of an ILP. Gathers Type of instructor
'				and begining class info.
'Date:		9 July 2001
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Dimension Variables, make db Connection, print HTML header.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim intStudent_ID 
dim strTeacherText
dim intInstructor_ID
dim rsStudent
dim vntGrade
dim bolHighSchool
dim strShort_Ilp		'contains short ilp list or link to ilpShortForm.asp if none exist

dim oFunc		'wsc object
Session.Value("strTitle") = "Add a Class"
Session.Value("strLastUpdate") = "22 Feb 2002"
Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")

if Session.Value("bolUserLoggedIn") = false then
	Response.Expires = -1000	'Makes the browser not cache this page
	Response.Buffer = True		'Buffers the content so our Response.Redirect will work
	Session.Value("strURL") = Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
	Server.Execute(Application.Value("strWebRoot") & "UserAdmin/Login.asp")
else 
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'' Get Student Name 
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
    call oFunc.OpenCN()
   
	if request("intStudent_id") <> "" then
		intStudent_ID = request("intStudent_id")
		Session.Value("intStudent_ID") = intStudent_ID
	elseif Session.Value("intStudent_ID") <> "" then
		intStudent_ID = Session.Value("intStudent_ID")
	else
		'Redirect to the home page if this URL is entered direcly
		Server.Transfer Application.Value("strMiniRoot") & "default.asp"
	end if 

	set rsStudent = server.CreateObject("ADODB.RECORDSET")
	with rsStudent
		.CursorLocation = 3
		sqlStudent = "select s.szFirst_Name,s.szLast_Name, ss.szGrade " & _
					 "from tblStudent s inner join tblStudent_States ss " & _
					 "on s.intStudent_ID = ss.intStudent_ID and " & _
					 "ss.intSchool_year = " & session.Contents("intSchool_Year") & " " & _
					 " where s.intStudent_ID=" & intStudent_ID
		.Open sqlStudent, oFunc.FPCScnn	
		'bkm 18-jun-02
		'added check for valid intStudent_ID - redirects to home page if invalid
		if not .BOF and not .EOF then
			strStudentName = rsStudent("szFirst_Name") & " " & rsStudent("szLast_Name")
			'bkm 24-jun-02 - used by ILPShortForm drop down
			vntGrade =  rsStudent("szGrade")	'bkm 24-jun-02
			
			if isNumeric(vntGrade) then
				if cint(vntGrade) >= 9 then
					bolHighSchool = true
				else
					bolHighSchool = false
				end if
			end if
			Session.Value("strStudentName") = strStudentName
		else
			Response.Redirect Application.Value("strWebRoot") & "?strMessage=Student not found"
		end if
		.Close
	end with
	set rsStudent = nothing

	if Session.Value("intSponsorTeacherID" & intStudent_ID & session.Value("intSchool_Year")) = "" then
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'' Before anyone can start filling out the ILP there must be a Sponsor
	'' Teacher assigned to the student. We check that first and if one hasn't 
	'' been assigned they are taken to addSponsorTeacher.asp which allows 
	'' them to set one up.
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		dim sqlSponsor 
		dim intSchoolYear
		
		intSchoolYear = session.Value("intSchool_Year")  'from log in 
		
		set rsSponsor = server.CreateObject("ADODB.RECORDSET")
		rsSponsor.CursorLocation = 3
		sqlSponsor = "select intSponsor_Teacher_ID from tblEnroll_Info " & _
					 "where intStudent_ID = " & intStudent_ID & _
					 " and sintSchool_YEAR = " & intSchoolYear
		rsSponsor.Open sqlSponsor, oFunc.FPCScnn	
			
		if rsSponsor.RecordCount > 0 then
			if rsSponsor("intSponsor_Teacher_ID") & "" <> "" then
				' Sponsor Teacher has been selected
				Session.Value("intSponsorTeacherID" & intStudent_ID & session.Value("intSchool_Year")) = rsSponsor("intSponsor_Teacher_ID")
			else
				' No sponsor teacher has been selected
				rsSponsor.Close
				set rsSponsor = nothing
				Response.Redirect(Application.Value("strWebRoot") & "forms/packet/addSponsorTeacher.asp?intShort_ILP_ID=" & request("intShort_ILP_ID"))
			end if 			
		else
			' No sponsor teacher has been selected
			rsSponsor.Close
			set rsSponsor = nothing
			Response.Redirect(Application.Value("strWebRoot") & "forms/packet/addSponsorTeacher.asp?intShort_ILP_ID=" & request("intShort_ILP_ID"))
		end if 
		rsSponsor.Close
		set rsSponsor = nothing
	end if 
	
	' Without this if statement the instructor_id would persist when the 
	' instruction type is changed giving us undesirable results.
	if request("intInstruct_Type_ID") <> "" then 
		intInstructor_ID = Request("intInstructor_ID")
	end if
	
	Session.Value("intShort_ILP_ID") = Request("intShort_ILP_ID")
%>
<%
' first check to see if year is locked
if oFunc.LockYear then
%>
<span class="svplain8"><b>This school year has been locked and no modifications can be made.</b></span><br><BR>
<input type=button class="Navlink" value="Cancel" onClick="window.location.href='<%=Application("strSSLWebRoot")%>forms/packet/packet.asp?intStudent_ID=<%=intStudent_ID%>';" id="Button3" NAME="Button1">
</body>
</html>
<%	
	response.End
end if
%>
<form action="" name=main method=get ID="Form1">
<input type=hidden name=viewing value="true" ID="Hidden2">
<input type=hidden name="fromILP" value="true" ID="Hidden3"> <!-- this toggles the contract guardian field in classAdmin.asp -->

<table width=100% ID="Table1">
	<tr>	
		<Td class=yellowHeader >
				&nbsp;<b>Implement Plan for Course: '<% = oFunc.CourseInfo(request("intShort_ILP_ID"),3)%>'</b> <font size=1>(Student: <% = strStudentName %>)</font> 
		</td>
	</tr>
	<tr>
		<td bgcolor=f7f7f7>
			<table ID="Table2">
				<!--<TR>					
						<%
							'dim strSQLshortILP
							'if bolHighSchool then
								'show course description from High School Program Of Studies table
							'	strSQLshortILP =	"SELECT sf.intShort_ILP_ID, sf.intSchool_Year, pos.txtCourseNbr + ':' + pos.txtCourseTitle AS CourseDesc " & _
							'			"FROM tblILP_SHORT_FORM sf INNER JOIN " & _
							'			"tblProgramOfStudies pos ON sf.lngPOS_ID = pos.lngPOS_ID " & _
							'			"WHERE (sf.intStudent_ID = " & intStudent_ID & ") AND (sf.intSchool_Year = " & Session.Contents("intSchool_Year") & ") " & _
							'			" and not exists(select 'x' from tblILP i where i.intShort_ILP_ID = sf.intShort_ILP_ID) " & _
							'			"ORDER BY CourseDesc"
							'else
								'show course description from trefPOS_Subjects (Non-High School)
							'	strSQLshortILP =	"SELECT sf.intShort_ILP_ID, pos.szSubject_Name + ':' + sf.szCourse_Title AS CourseDesc, " & _
							'			"sf.intSchool_Year " & _
							'			"FROM tblILP_SHORT_FORM sf INNER JOIN " & _
							'			"trefPOS_Subjects pos ON sf.intPOS_Subject_ID =pos.intPOS_Subject_ID " & _
							'			"WHERE (sf.intStudent_ID = " & intStudent_ID & ") AND (sf.intSchool_Year = " & Session.Contents("intSchool_Year") & ") " & _
							'			" and not exists(select 'x' from tblILP i where i.intShort_ILP_ID = sf.intShort_ILP_ID) " & _
							'			"ORDER BY CourseDesc"
							'end if
							'strShort_Ilp =  oFunc.MakeListSQL(strSQLshortILP,"intShort_ILP_ID","CourseDesc",Request("intShort_ILP_ID"))												 
							'if oFunc.makeListRecordCount > 0 then
						%>
					<TD>
						<FONT class="svplain11">Select a Course: </FONT><FONT class="svplain"></FONT>
					</TD>
					<TD CLASS=svplain11>						
						<SELECT name="intShort_ILP_ID" id="Select1" onChange="this.form.action = 'ILP1.asp',this.form.submit();" >
							<OPTION value="">Course Outlines</OPTION>
							<% 'strShort_Ilp %>
						</SELECT>
						<%
						    'else
						%>
					<TD CLASS=svplain11>	
							In order to add class you must first create 
							a corrisponding course in the Course Outline. <br>
							To create one now 
							<a href="<%=Application("strSSLWebRoot")%>forms/ilp/ilpShortForm.asp?intStudent_id=<% = intStudent_ID %>">
							click here</a>.<br><br>
							<b>Note:</b> Courses from your Course Outline that have been attached to 
							a class will not show up in the list.  You can use a course from your 
							Course Outline only once.						
						<%
							'end if
						%>    
						
					</TD>
				</TR>-->
				<input type=hidden name="intShort_ILP_ID" value="<% = request("intShort_ILP_ID") %>">
				<% if request("intShort_ILP_ID") <> "" then 
					
					' This gets the intPOS_Subject_ID for use in classAdmin.asp
					' and ilpMain.asp
					set rsGetPOS = server.CreateObject("ADODB.Recordset")
					rsGetPOS.CursorLocation = 3
					sql = "select intPOS_Subject_ID from tblILP_Short_Form " & _
						  " where intShort_ILP_ID = " & request("intShort_ILP_ID")
					rsGetPOS.Open sql, oFunc.FPCScnn
						
					if rsGetPOS.RecordCount > 0 then
						Session.Contents("intPOS_Subject_ID") = rsGetPOS(0)
					end if	
					rsGetPOS.Close
					set rsGetPOS = nothing
					
					' Now print the 'Who will teach course' Section
				%>
				<tr>	
					<Td>
						<font class=svplain11>
							Who will teach the course?  
						</font>
						<font class=svplain>
						</font>
					</td>					
					<td>
						<select name="intInstruct_Type_ID" onChange="this.form.action = 'ILP1.asp',this.form.submit();" ID="Select2">
							<option value="">
						<%
							dim sqlInstruct
							sqlInstruct = "Select intInstruct_Type_ID,szInstruct_Name " & _
											 "from trefInstruct_Type where bolShow = 1 order by szInstruct_Name"
							Response.Write oFunc.MakeListSQL(sqlInstruct,intInstruct_Type_ID,szInstruct_Name,Request("intInstruct_Type_ID"))												 
						%>
						</select>	
					</td>
				</tr>		

<%
				end if 
				
	select case Request("intInstruct_Type_ID")		
		case "1"	'Parent Teacher
			call vbfGuardianList
		case "4"	'Contract ASD Teacher
			if not oFunc.LockSpending then
				strTeacherText = "Select a Contract ASD Teacher: " 				
			else
				strTeacherText = "<B>PLEASE NOTE: Spending has been locked <BR>and no ASD charges can be made."
			end if
			call vbfTeacherList()
		case "5"	'Vendor Teacher
			strTeacherText = "Select a Vendor: " 
			call vbfVendorList	
		case "6"	'Class from ASD School
			call vbfClasses("intInstruct_Type_ID = 6 ")			
	end select

	if intInstructor_ID <> "" and Request("intInstruct_Type_ID") = "4" then				
		call vbfClasses("intInstructor_ID = " & intInstructor_ID)
	elseif request("intGuardian_ID") <> "" and Request("intInstruct_Type_ID") = "1" then
		call vbfClasses( "intGuardian_ID = " & request("intGuardian_ID"))
	elseif request("intVendor_ID") <> "" and Request("intInstruct_Type_ID") = "5" then
		call vbfClasses( "intVendor_ID = " & request("intVendor_ID"))
	end if
			
function vbfTeacherList()
	' Prints list of teachers 
%>
				<tr>	
					<Td class=svplain11>
						<% = strTeacherText %>
					</td>
					<td class=svplain11>
						<% if not oFunc.LockSpending then %>
						<select name="intInstructor_ID" onChange="this.form.action = 'ILP1.asp',this.form.submit();" ID="Select3">
							<option value="">
						<%
							set oList = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/dbOptionsList.wsc"))
							Response.Write oList.ActiveTeachers(session.Contents("intSchool_Year"),intInstructor_ID)		
							set oList = nothing
						%>
						</select>	<b>OR</b> <input type="button" value="Search for a Class" class="NavSave" onClick="jfClassSearch();" ID="Button4" NAME="Button4">
						<script language="javascript">
							function jfClassSearch(){
								var strURL = "<%=Application.Value("strWebRoot")%>forms/Teachers/classSearch.asp?intStudent_ID=<% = intStudent_ID %>";
								strURL += "&intShort_ILP_ID=<% = Request("intShort_ILP_ID") %>&intInstruct_Type_ID=<% = Request("intInstruct_Type_ID")%>&bolWin=true";
								strURL += "&intPOS_Subject_ID=<%= Session.Contents("intPOS_Subject_ID")%>";
								var searchWin = window.open(strURL,"searchWin","width=710,height=500,scrollbars=yes,resizable=yes");
								searchWin.moveTo(0,0);
								searchWin.focus();
							}
						</script>
						<% end if%>
					</td>
				</tr>
<%
end function

function vbfGuardianList
	' Prints list of Guardians based on intStudent_ID
%>
				<tr>	
					<Td>
						<font class=svplain11>
							Select a Parent/Guardian: 
						</font>
						<font class=svplain>
						</font>
					</td>
					<td>
						<select name="intGuardian_ID" onChange="this.form.action = 'ILP1.asp',this.form.submit();" ID="Select4">
							<option value="">
						<%
							dim sqlGaurdian
							' This union will give us all guardians that belong to the
							' student's family and all parent's that are not part of students family
							' but allowed the students family to have acces to their class.
							' The class has to have the same POS_Subject in order for the
							' non-family parent nto show up in the list. 
							' If a non-family parent does show up we will know they are not
							' part of the family because thier guardian id will be negative.				 
							'JD: Allow only active guardians to show in the ddl			 
							sqlGaurdian = "SELECT g.intGUARDIAN_ID as Guard_ID, g.szLAST_NAME + ',' + g.szFIRST_NAME AS Name " & _
											"FROM tblGUARDIAN g INNER JOIN " & _
											"					tascSTUDENT_GUARDIAN sg ON g.intGUARDIAN_ID = sg.intGUARDIAN_ID " & _
											"WHERE (sg.intSTUDENT_ID = " & intStudent_ID & "AND blnDeleted = 0) " & _
											"UNION " & _
											"SELECT DISTINCT  " & _
											"					(CASE WHEN fg.intFamily_ID = s.intFamily_ID THEN g.intGUARDIAN_ID ELSE g.intGUARDIAN_ID * - 1 END) as Guard_ID,  " & _
											"					g.szLAST_NAME + ',' + g.szFIRST_NAME AS Name " & _
											"FROM tblSTUDENT s INNER JOIN " & _
											"					tascClass_Family cf ON s.intFamily_ID = cf.intFamily_ID INNER JOIN " & _
											"					tblClasses ON cf.intClass_ID = tblClasses.intClass_ID INNER JOIN " & _
											"					tblGUARDIAN g ON tblClasses.intGuardian_ID = g.intGUARDIAN_ID INNER JOIN " & _
											"					tascFAM_GUARD fg ON g.intGUARDIAN_ID = fg.intGUARDIAN_ID " & _
											"WHERE (s.intSTUDENT_ID = " & intStudent_ID & ") AND " & _
											"(tblClasses.intGuardian_ID IS NOT NULL) AND " & _
											"(tblClasses.intPOS_Subject_ID = " & Session.Contents("intPOS_Subject_ID") & ") " & _
											" AND (tblClasses.intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _
											"ORDER BY 2"
							Response.Write oFunc.MakeListSQL(sqlGaurdian,"Guard_ID","Name",request("intGuardian_ID"))												 
						%>
						</select>	
					</td>
				</tr>
<%
end function

function vbfVendorList
	' Prints list of Guardians based on intStudent_ID
%>
				<tr>	
					<Td>
						<font class=svplain11>
							Select a Vendor: 
						</font>
						<font class=svplain>
						</font>
					</td>
					<td>
						<select name="intVendor_ID" onChange="this.form.action = 'ILP1.asp',this.form.submit();" ID="Select5">
							<option value="">
						<%
							dim sqlVendor
							sqlVendor = "Select intVendor_ID, szVendor_Name " & _
											 "from tblVendors order by szVendor_Name"										 
							Response.Write oFunc.MakeListSQL(sqlVendor,"intVendor_ID","szVendor_Name",request("intVendor_ID"))												 
						%>
						</select>	
					</td>
				</tr>				
<%
end function

function vbfClasses(strClause)
	' Prints list of classes based on the conditions provided by the 'strClause' parameter.
	dim sqlClasses
	dim strClassList
	dim strSQLADD
	dim bolClassAdd
	dim intFamily_ID
	
	
	bolClassAdd = true
	' If we have a negative number in the strClause it's because we have a
	' parent instructed class and the parent is NOT part of the student's current
	' family.  If this is the case we do NOT want the user to be able to add 
	' a class for a non-family parent so we set the bol flag and use it to hide
	' the create a class button latter. 
	if instr(1,strClause,"-") > 0 then
		bolClassAdd = false
		strClause = replace(strClause,"-","")
	end if
	
	' The following logic was replaced by the simple assignment to strSQLADD
	' since the way it handles the family id is applicable to all roles
	' SMB 10-18-2005
	
	'if session.Value("strROLE") = "ADMIN" then
	'	strSQLADD = "" 		
	'else
	'	if session.Value("strROLE") = "TEACHER" then 
	'		intFamily_ID = oFunc.FamilyInfo(1,intStudent_id,1)
	'	else	
	'		intFamily_ID = session.Value("intFamily_ID")
	'	end if
	'	strSQLADD = " AND " & _
	'				"(EXISTS " & _
	'				"	(SELECT 'x' " & _
	'				"		FROM tascClass_Family a " & _
	'				"		WHERE c.intClass_ID = a.intClass_ID AND " & _
	'				"		a.intFamily_ID = " & intFamily_ID & ") OR " & _
	'				"	NOT EXISTS " & _
	'				"		(SELECT 'x' " & _ 
	'				"		FROM tascClass_Family a " & _
	''				"		WHERE c.intClass_ID = a.intClass_ID)) "
	'end if 		 					 	
	
	
	strSQLADD = " AND " & _
					"(EXISTS " & _
					"	(SELECT 'x' " & _
					"		FROM tascClass_Family a " & _
					"		WHERE c.intClass_ID = a.intClass_ID AND " & _
					"		a.intFamily_ID = s.intFamily_ID) OR " & _
					"	NOT EXISTS " & _
					"		(SELECT 'x' " & _ 
					"		FROM tascClass_Family a " & _
					"		WHERE c.intClass_ID = a.intClass_ID)) "
					
	if Application.Contents("bolUseContractApproval"&session.Contents("intSchool_Year")) and intInstructor_ID & "" <> "" then
		strSQLADD = strSQLADD & " AND c.intContract_Status_ID = 5 " 
	end if
	
	'17-sept-2002 bkm - only show classes that have not met the max # of students
				 
	sqlClasses = "SELECT     c.intClass_ID, c.szClass_Name, c.intMax_Students, COUNT(tblILP.intStudent_ID) AS CountStudents " & _
				"FROM tblClasses c INNER JOIN " & _
				" tblILP_SHORT_FORM ON c.intPOS_Subject_ID = tblILP_SHORT_FORM.intPOS_Subject_ID LEFT OUTER JOIN " & _
				" tblILP ON c.intClass_ID = tblILP.intClass_ID INNER JOIN " & _
                " tblSTUDENT s ON s.intSTUDENT_ID = tblILP_SHORT_FORM.intStudent_ID " & _
				"WHERE " & strClause & " AND (NOT EXISTS " & _
				"  (SELECT  intILP_Id " & _
				"  FROM  tblILP i " & _
				"  WHERE c.intClass_id = i.intClass_Id AND i.intStudent_id = " & intStudent_id & ")) " & _
				" AND (c.intSchool_Year = " & session.Contents("intSchool_Year") & ") AND  " & _
				"  (tblILP_SHORT_FORM.intShort_ILP_ID = " & Request("intShort_ILP_ID") & ") " & _
				strSQLADD & _
				" GROUP BY c.intClass_ID, c.szClass_Name, c.intMax_Students " & _
				"HAVING  (COUNT(tblILP.intStudent_ID) < c.intMax_Students) " & _
				"ORDER BY c.szClass_Name"

	'if Ucase(session.contents("strUserID")) = "SCOTT" then response.write sqlClasses 

	strClassList = oFunc.MakeListSQL(sqlClasses,"intClass_ID","szClass_Name",Request("intClass_ID"))


	if oFunc.makeListRecordCount > 0 then
%>
				<tr>	
					<Td class=svplain11>
					 Select an Existing Class: 
					</td>
					<script language=javascript>
						function jfShowClass(id){
							if (id != "") {
								document.main.action = '../Teachers/classAdmin.asp'
								document.main.submit();
							}
						}
					</script>
					<td class=svplain11>
						<select name="intClass_ID" onChange="jfShowClass(this.value);" ID="Select6">
							<option value="">
						<% = strClassList %>
						</select>
					</td>
					<% if request("bolFromSearch") <> "" then%>
					<script language="javascript">
						jfShowClass('<% = request("intClass_ID") %>');
					</script>
					<% end if %>
				</tr>
<%	else %>
				<tr>	
					<Td class=svplain11 colspan=2>
					<br>
					 The selected instructor currently does not offer <BR>
					 classes that match the subject you selected or the <br>
					 maximum number of students for that class has been met.
					 <br><br>
					</td>
				</tr>
<%	end if %>
				</form>
				<form action="../Teachers/classAdmin.asp" method=get name=create onSubmit="return false;" ID="Form2">
				<input type=hidden name="intInstructor_ID" value="<% = intInstructor_ID %>" ID="Hidden4">	
				<input type=hidden name="intGuardian_ID" value="<% = request("intGuardian_ID") %>" ID="Hidden5">
				<input type=hidden name=intInstruct_Type_ID value="<%=request("intInstruct_Type_ID")%>" ID="Hidden6"> 
				<input type=hidden name=intVendor_ID value="<%=request("intVendor_ID")%>" ID="Hidden7"> 
				<input type=hidden name=intClass_ID value="" ID="Hidden8">	
				<input type=hidden name=bolValidated value=false ID="Hidden9">	<!-- used to let us know if we need to add a confirmation message or not -->				
<%  
' Only can create a class if parent instructed AND parent is part of student's family
if Request("intInstruct_Type_ID") = "1"  and bolClassAdd then %>
				<tr>
					<Td class=svplain11 colspan=2>
					 You may create a new class for the selected course <Br>by clicking 'create':  						
						<input type=submit value="CREATE" class="btSmallGray" name="create" onClick="this.form.action = '../Teachers/classAdmin.asp',this.form.submit();">
					</td>
					</form>
					<TR>
						<td>
							<br>
						</td>		
					</tr>
				</tr>
<%
	end if
end function 

if request("intInstruct_Type_ID") = "" then
%>
				<tr>
					<td colspan=10 class="instruct" style="padding-left: 7px;">
						<b>Note: <br>
						Parent instruction includes parent taught classes, UAA classes, <BR>
						ASD/Charge Back classes and vendor classes (such as little gym, <BR>
						kumon etc.). </b>
					</td>				
				</tr>
<%
end if
%>					
			</table><BR>
			<input type=button class="Navlink" value="Cancel" onClick="window.location.href='<%=Application("strSSLWebRoot")%>forms/packet/packet.asp?intStudent_ID=<%=intStudent_ID%>';" id="Button1" NAME="Button1">
			<input type=button class="Navlink" value="Reset Form" onClick="window.location.href='<%=Application("strSSLWebRoot")%>forms/ilp/ilp1.asp?intStudent_ID=<%=intStudent_ID%>&intShort_ILP_ID=<%=request("intShort_ILP_ID")%>';" id="Button2" NAME="Button1">
		</td>
	</tr>
</table>
</form>
<%
   call oFunc.CloseCN()
   set oFunc = nothing
end if
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")
%>
