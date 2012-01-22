<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		addCourse.asp
'Purpose:	Handles inserting and updating short form records.
'Date:		moved 10/26/2004
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim intStudent_ID		'Unique Student ID 
dim bolHighSchool		'True = High Scool Student: False = Elementry/Jr High Student
dim sql					'generic sql string
dim intPOS_Subject_ID	'defined in vbsGetShortForm. Dim'd here for global access.
dim szSemester          'JD: keep track of the semester value in list
dim lngPOS_ID			' same as previous comment
dim intShort_ILP_ID		' same as previous comment
dim szCourse_Title		' same as previous comment
dim intCourse_Hrs		' same as previous comment
dim strMessage			'Message that will be displayed in alert box of opener window
dim intGrade			' grade level of student
dim intSchoolTypeId	   	' 1 = elementary, 2 = middles school, 3 = high school

'prevent page from cacheing
Response.AddHeader "Cache-Control","No-Cache"
Response.Expires = -1

if request.QueryString("intStudent_ID") <> "" then
	'Define variable if get properly called
	intStudent_ID = request.QueryString("intStudent_ID")
	bolHighSchool = request.QueryString("bolHighSchool")
	intShort_ILP_ID = request.QueryString("intShort_ILP_ID")
	intGrade = request.QueryString("GRADE")
	
elseif request.Form("intStudent_ID") <> "" then
	'Define variable if post properly called
	intStudent_ID = request.Form("intStudent_ID")
	bolHighSchool = request.Form("bolHighSchool")	
	intShort_ILP_ID = request.Form("intShort_ILP_ID")
	intGrade = request.Form("GRADE")
else
	'terminate page since page was improperly called.
	response.Write "<html><body><H1>Page improperly called.</h1></body></html>"
	response.End
end if 

if not isNumeric(intGrade) then
	intSchoolTypeId = 1
	bolHighSchool = false
elseif intGrade < 6 then
	intSchoolTypeId = 1
	bolHighSchool = false
elseif intGrade < 9 then 
	intSchoolTypeId	= 2
	bolHighSchool = true
else
	intSchoolTypeId = 3
	bolHighSchool = true
end if

if bolHighSchool & "" = "" then bolHighSchool = false

' Now we can proceded to define needed function object
dim oFunc		'wsc object
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

' Fuse Box type logic
if request.QueryString("intShort_ILP_ID") <> ""  then
	'We are in Edit Mode which was called from Packet.asp
	call vbsGetShortForm(intShort_ILP_ID)
elseif request.Form("state") = "update" then
	' Update record
	call vbsUpdateShortForm
elseif request.Form("state") = "insert" then
	' Insert Record
	call vbsInsertShortForm
elseif request("intPOS_Subject_ID") <> "" then
	' This is a result of intPOS_Subject_ID being changed in the Select list
	' and the student is in high school. 
	intPOS_Subject_ID = request("intPOS_Subject_ID")
    'JD: Keep track of the semester selected
	if request("szSemester") <> "" then
    szSemester = request("szSemester")
end if

end if 



if strMessage <> "" then
' If we just made an update or insert we send a message and refresh the opener window
' and close this window.
%>
<html>
<body onload="window.opener.location.href='Packet.asp?intStudent_ID=<%=intStudent_ID%>&intShort_ILP_ID=<%=intShort_ILP_ID%><%=session.Contents("strSimpleHeader")%>';window.opener.focus();window.close();">
</body>
</html>
<%
	response.End
end if 
'Prepare and print header
Session.Value("strTitle") = "Plan a Course"
Session.Value("strLastUpdate") = "03 June 2002"
Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")
%>
<script language=javascript>
	function jfValidateCourse(state){
		var strError = "";
				
		<% if ucase(bolHighSchool) = "TRUE" then %>
		if (document.main.lngPOS_ID.value == "") {
			strError += "'Course Title' can not be blank.\n";
		}
		<%else%>
		if (document.main.szCourse_Title.value == "") {
			strError += "'Course Title' can not be blank.\n";
		}
		<% end if %>
				
		if (isNaN(document.main.intCourse_Hrs.value) || document.main.intCourse_Hrs.value == ""){
			strError += "'Course Hours' must be a valid number and can not be blank.\n";
		}
		if (strError != ""){
			alert("The following items need to be corrected before your information can be saved.\n" + strError);
		}else{
			document.main.state.value = state;
			main.submit();
		}
	}
</script>
<form action="addCourse.asp" method=post name=main onsubmit="return false;" ID="Form1">
<input type=hidden name="intStudent_ID" value="<% = intStudent_ID %>" ID="Hidden1">
<input type=hidden name="bolHighSchool" value="<% = bolHighSchool %>" ID="Hidden2">
<input type=hidden name="intShort_ILP_ID" value="<% = intShort_ILP_ID %>" ID="Hidden3">
<input type=hidden name="state" value="" ID="Hidden4">
<input type=hidden name="GRADE" value="<% = intGrade %>" >
<table width=100% class=yellowHeader ID="Table1">
	<tr>
		<td>
			&nbsp;<b>Plan a Course</b>
		</td>
	</tr>
</table>
<table ID="Table2">
	<tr>
		<td class="gray">
			&nbsp;Subject Area
		</td>
		<td> 
			<select name="intPOS_Subject_ID" <% if ucase(bolHighSchool) = "TRUE" then%> onchange="main.submit();" <% end if %> ID="Select1">
				<option value=''>
				<%
					' Create POS Subject HTML Option List
					sql = "select distinct ps.intPOS_Subject_ID, upper(szSubject_Name) as szSubject_Name " & _
					      " from trefPOS_Subjects ps " & _
					      " inner join tblProgramOfStudies po on po.intPOS_Subject_id = ps.intPOS_Subject_id " & _
					      " and po.intSchool_Level_id = " & intSchoolTypeId & _
					      " where bolShow = '1' and IS_ACTIVE = 1 " & _
					      " order by szSubject_Name"
'response.write sql
					response.Write oFunc.MakeListSQL(sql,"intPOS_Subject_ID","szSubject_Name",intPOS_Subject_ID)	
				%>
			</select>
		</td>
	</tr>
	
	<tr>
		<td class="gray">
			&nbsp;Semester
		</td>
		<td>
		<%
		    'JD: Sort by semester list
			if intSchoolTypeId = 2 or intSchoolTypeId = 3 then
		%>
		<%
		    if intPOS_Subject_ID <> "" and intPOS_Subject_ID<>22 then
			    response.Write "<select name='szSemester' onchange='main.submit()'>"
    			response.Write "<option value=''>"
				response.Write oFunc.MakeList("1,2","1st,2nd",szSemester)
    		    response.Write "</select>"
		     end if 
		 %>
		 <% end if %>
        </td>
	</tr>
	
	
	<tr>
		<td class="gray" valign='middle'>
			&nbsp;Course Title
		</td>
		<td class='svplain8'>

		<%
			if intSchoolTypeId = 2 or intSchoolTypeId = 3 then
		%>

<%
                'JD:Dont show courses when there's no semester selected, except for sponsorship/oversight courses
				if intPOS_Subject_ID <> "" and szSemester <> "" and intPOS_Subject_ID <> 22 then
					response.Write "<select name='lngPOS_ID'>"
					'sql = "select lngPOS_ID,txtCourseNbr + ': ' + txtCourseTitle as Course " & _
					'	"from tblProgramOfStudies " & _
					'	"where IS_ACTIVE = 1 " & _
					'	"and intPOS_SUBJECT_ID = " & intPOS_Subject_ID & _
					'	" and intSchool_Level_id = " & intSchoolTypeId & _
					'	" and txtCourseTitle like  '%sem " & szSemester & "'" & _       
					'	" order by txtCourseTitle "
					'JD:Select by semester but also include tutoring for both TO TEST
					sql = "select lngPOS_ID,txtCourseNbr + ': ' + txtCourseTitle as Course, txtCourseNbr " & _
                        "from tblProgramOfStudies " & _
                        "where IS_ACTIVE = 1 " & _
                        "and intPOS_SUBJECT_ID = " & intPOS_Subject_ID & _
                        " and intSchool_Level_id = " & intSchoolTypeId & _
                        " and txtCourseTitle like  '%sem " & szSemester & "'" & _ 
                        "union " & _
                        "select lngPOS_ID,txtCourseNbr + ': ' + txtCourseTitle as Course, txtCourseNbr " & _
                        "from tblProgramOfStudies " & _
                        "where IS_ACTIVE = 1 " & _
						"and intPOS_SUBJECT_ID = " & intPOS_Subject_ID & _
						" and intSchool_Level_id = " & intSchoolTypeId & _
                        "and txtCourseNbr = '00000' " & _
                        "order by txtCourseNbr "
					Response.Write oFunc.MakeListSQL(sql,"lngPOS_ID","Course",lngPOS_ID)	
					response.Write "</select>"
				elseif intPOS_Subject_ID = 22 then
				    response.Write "<select name='lngPOS_ID'>"
					'sql = "select lngPOS_ID,txtCourseNbr + ': ' + txtCourseTitle as Course " & _
					'	"from tblProgramOfStudies " & _
					'	"where IS_ACTIVE = 1 " & _
					'	"and intPOS_SUBJECT_ID = " & intPOS_Subject_ID & _
					'	" and intSchool_Level_id = " & intSchoolTypeId & _
					'	" and txtCourseTitle like  '%sem " & szSemester & "'" & _       
					'	" order by txtCourseTitle "
					'JD:Select by semester but also include tutoring for both TO TEST
					sql = "select lngPOS_ID,txtCourseNbr + ': ' + txtCourseTitle as Course, txtCourseNbr " & _
                        "from tblProgramOfStudies " & _
                        "where IS_ACTIVE = 1 " & _
                        "and intPOS_SUBJECT_ID = " & intPOS_Subject_ID & _
                        " and intSchool_Level_id = " & intSchoolTypeId & _
                        "union " & _
                        "select lngPOS_ID,txtCourseNbr + ': ' + txtCourseTitle as Course, txtCourseNbr " & _
                        "from tblProgramOfStudies " & _
                        "where IS_ACTIVE = 1 " & _
						"and intPOS_SUBJECT_ID = " & intPOS_Subject_ID & _
						" and intSchool_Level_id = " & intSchoolTypeId & _
                        "and txtCourseNbr = '00000' " & _
                        "order by txtCourseNbr "
					Response.Write oFunc.MakeListSQL(sql,"lngPOS_ID","Course",lngPOS_ID)	
					response.Write "</select>"

				end if 
		%>

		<%
			else			
		%>
		<table>
<tr>
<td>
			<input type=text name="szCourse_Title" value="<% = szCourse_Title %>" maxlength=255 size=25 ID="Text1">
<td class="svplain8">
<b>Note:</b> Course titles must be in a <br>
'Curriculum - Level (Sem1 or Sem2 or All Year) format<br>
Example: Saxon Math - Gr. 3-4 Sem 1
</td>
</tr>
</table>
		<%
			end if 
		%>
		</td>
	</tr>
	<!--<tr>
		<td class="gray">
			&nbsp;Which Semester
		</td>
		<td>
			<select name="szSemesters">
				<%
				'Response.Write oFunc.MakeList("1,2,Both","1st,2nd,Both",szSemesters)
				%>
			</select>
        </td>
	</tr>-->
	<tr>
		<td class="gray">
			&nbsp;Course Hours&nbsp;
		</td>
		<td>
			<table ID="Table3">
				<tr>
					<td>
						<input type=text name="intCourse_Hrs" value="<% = intCourse_Hrs %>" size=5 maxlength=3 ID="Text2">
					</td>
					<td class=svplain10>
						Includes all hours student spends on this subject <br>
						(i.e. time with parent,teacher,vendor and homework).
					</td>
				</tr>
			</table>			
        </td>
	</tr>
	<tr>
		<td colspan=2>
		<input type=button value="Close without saving" class="NavLink" onClick="window.opener.focus();window.close();">
		<% if intPOS_Subject_ID <> "" or bolHighSchool <> "TRUE" then 
				if request("intShort_ILP_ID") <> "" then %>
			<input type=submit value="SAVE" onclick="jfValidateCourse('update');" class="NavSave">			
		<%		else %>
			<input type=submit value="SAVE" onclick="jfValidateCourse('insert');" class="NavSave">
		<%		end if 
		   end if	 %>			
		</td>
	</tr>
</table>
</form>
<% if intSchoolTypeId = 3 then %>
<span class="svplain8"><a href="https://docs.google.com/viewer?a=v&pid=explorer&chrome=true&srcid=0B5w5Wuf_btjhYWIyYTc3MmQtNzVkZS00OWVmLWIyYmUtNTUyYmM1ODVhZDBk&hl=en_US&pli=1" target="_blank">ASD High School Program of Studies Link</a></span>
<% elseif intSchoolTypeId = 2 then %>
<span class="svplain8"><a href="https://docs.google.com/viewer?a=v&pid=explorer&chrome=true&srcid=0B5w5Wuf_btjhOTUwNjgzMTctNGUwYi00MTJlLWIxODUtNzcyY2YwYTA5ZDlm&hl=en_US" target="_blank">ASD Middle School Program of Studies Link</a></span>
<% end if %>
<%
sub vbsGetShortForm(id)
	' Set up the recordset
	dim rsShortForm
	set rsShortForm = server.CreateObject("ADODB.Recordset")
	rsShortForm.CursorLocation = 3
	
	' intSchool_Year and intStudent_ID are only included in this sql as 
	' a safe guard against improper requests whether honest or malicous.
	' This sql retrieves our short form information		
	sql = "SELECT intPOS_Subject_ID, lngPOS_ID, szCourse_Title, intCourse_Hrs " & _
		  "FROM tblILP_SHORT_FORM " & _
		  "WHERE (intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _
		  "AND (intStudent_ID = " & intStudent_ID & " ) " & _
		  "AND (intShort_ILP_ID = " & id & ")" 
	rsShortForm.Open sql, oFunc.FPCScnn
	
	'Populate variables
	intPOS_Subject_ID = rsShortForm("intPOS_Subject_ID")
	lngPOS_ID = rsShortForm("lngPOS_ID")
	szCourse_Title = rsShortForm("szCourse_Title")
	intCourse_Hrs = rsShortForm("intCourse_Hrs")
	intShort_ILP_ID = id	
	
	' Clean up rs obj
	rsShortForm.Close
	set rsShortForm = nothing
end sub

sub vbsUpdateShortForm
	dim update
	' Setup and execute update sql
	update = "update tblILP_Short_Form set " & _
			 "intPOS_Subject_ID = " & request.Form("intPOS_Subject_ID") & "," & _
			 "szCourse_Title = '" & oFunc.EscapeTick(request.Form("szCourse_Title")) & "'," & _
			 "lngPOS_ID = " & oFunc.CheckDecimal(request.Form("lngPOS_ID")) & "," & _
			 "intCourse_Hrs = " & oFunc.CheckDecimal(request.Form("intCourse_Hrs")) & "," & _
			 "szUser_Modify = '" & session.Contents("strUserID") & "'," & _
			 "dtModify = '" & now() & "' " & _
			 "WHERE intShort_ILP_ID = " & request.Form("intShort_ILP_ID")
	oFunc.ExecuteCN(update)
	
	'THE FOLLOWING MAY NOT BE NEEDED AND SHOULD BE DELETED IF HERE AFTER 5-15-2003 smb
	' Update the ilp course hrs if needed so our course and ilp hrs are consistant
	'set rsILP = server.CreateObject("ADODB.RECORDSET")
	'rsILP.CursorLocation = 3
	
	'sql = "select intILP_ID from tblILP where intShort_ILP_ID = " & request.Form("intShort_ILP_ID")
	'rsILP.Open sql, oFunc.FPCScnn
	
	'if rsILP.RecordCount > 0 then
	'	update = "update tblILP set " & _
	'			 "decCourse_Hours = " & oFunc.CheckDecimal(request.Form("intCourse_Hrs")) & _
	'			 " where intShort_Form_ID = " & request.Form("intShort_ILP_ID")
	'	oFunc.ExecuteCN(update)
	'end if
	
	strMessage = "Update Complete"
end sub

sub vbsInsertShortForm
	dim insert
	' Setup and execute insert sql
	insert = "insert into tblILP_Short_Form(intStudent_ID,intPOS_Subject_ID," & _
			 "szCourse_Title,lngPOS_ID,intCourse_Hrs,intSchool_Year,szUser_Create," & _
			 "dtCreate) values (" & _
			 request.Form("intStudent_ID") & "," & _
			 request.Form("intPOS_Subject_ID") & "," & _
			 "'" & oFunc.EscapeTick(request.Form("szCourse_Title")) & "'," & _
			 oFunc.CheckDecimal(request.Form("lngPOS_ID")) & "," & _
			 oFunc.CheckDecimal(request.Form("intCourse_Hrs")) & "," & _
			 session.Contents("intSchool_Year") & "," & _
			 "'" & session.Contents("strUserID") & "'," & _
			 "'" & now() & "')"
	oFunc.ExecuteCN(insert)	
	strMessage = "Course Added"		 
	intShort_ILP_ID = oFunc.GetIdentity
end sub

'Closing remarks
call oFunc.CloseCN()
set oFunc = nothing
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")
%>

