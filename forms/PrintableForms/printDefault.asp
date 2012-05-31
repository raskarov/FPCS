<%@ Language=VBScript %>
<%
' TOGGLES SHOWING GOODS/SERVICES 
if session.Contents("strRole") = "ADMIN" then
	bolShow = true
else
	bolShow = false
end if

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Dimention Variables, make db Connection, print HTML header.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim intStudent_ID 
dim intContract_Guardian_ID
dim bolGetGuardian
dim sqlTeacher
dim sqlStudent
dim strURLAddition
dim strList
dim strIdType 
dim strDisabled
dim oFunc		'wsc object
dim bolGeneric
dim intActualPercent
dim strStudentInfo
dim strConSched			'tells us the difference between a contacted or scheduled class
dim strHideGoodService

Session.Contents("strTitle") = "View Classes"
Session.Contents("strLastUpdate") = "22 Feb 2002"
if request("simpleHeader") = "" then
	Server.Execute(Application.Value("strWebRoot") & "includes/header.asp")
else
	Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")
end if

Session.Contents("blnFromClassAdmin") = ""
session.Contents("strILPList") = ""
session.Contents("strClassList") = ""
Session.Contents("intILP_ID") = ""

if Session.Contents("bolUserLoggedIn") = false then
	Response.Expires = -1000	'Makes the browser not cache this page
	Response.Buffer = True		'Buffers the content so our Response.Redirect will work
	Session.Contents("strURL") = Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
	Server.Execute(Application.Value("strWebRoot") & "UserAdmin/Login.asp")
	response.End
end if
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()
ofunc.ResetSelectSessionVariables
bolGetGuardian = false
bolGeneric = "true"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Get Student Name 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This script is used to view classes that a student may have signed up for 
' or to view the classes that an instructor is scheduled to teach.
if Request.QueryString("intStudent_id") <> "" then
	'This section gives the classes for a student
	intStudent_ID = request("intStudent_id")
	strHideGoodService = "&bolHideGoodsServices=true"
	set rsStudent = server.CreateObject("ADODB.RECORDSET")
		rsStudent.CursorLocation = 3
		sqlStudent = "SELECT s.szFIRST_NAME, s.szLAST_NAME " & _
						"FROM tblSTUDENT s " & _
						"WHERE (s.intSTUDENT_ID = " & intStudent_ID & ") " 
	
		rsStudent.Open sqlStudent,Application("cnnFPCS")'oFunc.FPCScnn	

	Session.Contents("strStudentName") = rsStudent("szFirst_Name") & " " & rsStudent("szLast_Name")				 
	rsStudent.Close
	set rsStudent = nothing
		
	sql = "select ISF.szCourse_Title, POS.txtCourseTitle,c.intClass_ID,c.szClass_Name, c.intInstructor_ID,c.intInstruct_Type_id," & _
			"c.intGuardian_id,c.intVendor_id,i.intContract_Guardian_ID,i.intILP_ID as Class_ILP, " & _
			" teacherName = " & _
			" CASE " & _
			"  WHEN c.intInstructor_ID is not null THEN ins.szFirst_Name + ' ' + ins.szLast_Name " & _
			"  WHEN c.intVendor_ID is not null THEN v.szVendor_Name " & _
			"  WHEN c.intGuardian_ID IS NOT NULL THEN g.szFirst_Name + ' ' + g.szLast_Name " & _
			" END, i.decCourse_Hours, CASE c.intPOS_SUBJECT_ID WHEN 22 THEN 0 ELSE 1 END AS isSponsor  " & _
			"from tblILP i cross join tblClasses c " & _
			" LEFT OUTER JOIN tblInstructor ins ON c.intInstructor_ID = ins.intInstructor_ID " & _
			" LEFT OUTER JOIN tblVendors v ON c.intVendor_ID = v.intVendor_ID " & _
			" LEFT OUTER JOIN tblGuardian g ON c.intGuardian_ID = g.intGuardian_ID " & _
			" RIGHT OUTER JOIN tblILP_Short_Form ISF on i.intShort_ILP_ID = ISF.intShort_ILP_ID " & _
			" LEFT OUTER JOIN tblProgramOfStudies POS ON ISF.lngPOS_ID = POS.lngPOS_ID " & _
			"where i.intClass_ID = c.intClass_id and " & _
		    "i.intStudent_ID =" &  intStudent_ID & _ 
		    " and (c.intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _
		    " ORDER BY isSponsor, POS.txtCourseTitle, ISF.szCourse_Title"
		   
	bolGetGuardian = true
	
	dim sqlStudents		

	sqlStudents = "SELECT s.intSTUDENT_ID, " & _
				"Name = (Case ss.intReEnroll_State WHEN 86 then " & _
				"s.szLAST_NAME + ',' + s.szFIRST_NAME + ': Withdrawn (' + convert(varChar(20),ss.dtModify) + ')'" & _ 
				"WHEN 123 THEN s.szLAST_NAME + ',' + s.szFIRST_NAME + ': Graduated (' + convert(varChar(20),ss.dtModify) + ')'" & _ 
				"ELSE s.szLAST_NAME + ',' + s.szFIRST_NAME END) " & _
				"FROM tblSTUDENT s INNER JOIN " & _ 
				"tblStudent_States ss ON s.intSTUDENT_ID = ss.intStudent_id " & _ 
				"WHERE (ss.intReEnroll_State in(" & application.Contents("strEnrollmentList") & ")) AND (ss.intSchool_Year = " & Session.Value("intSchool_Year") & ") " & _ 
				"ORDER BY Name" 
																	
	strList = oFunc.MakeListSQL(sqlStudents,"intStudent_id","Name",intStudent_ID)
	strIdType = "intStudent_ID"
	strDisabled = "&strDisabled=true"
	bolGeneric = "false"
elseif Request.QueryString("intInstructor_ID") <> "" then  
	
	set rsTeacher = server.CreateObject("ADODB.RECORDSET")
		rsTeacher.CursorLocation = 3
		sqlTeacher = "select szFirst_Name,szLast_Name " & _
						"from tblInstructor where intInstructor_ID=" & Request.QueryString("intInstructor_ID")
		rsTeacher.Open sqlTeacher,Application("cnnFPCS")'oFunc.FPCScnn	

	Session.Contents("strTeacherFirstName") = rsTeacher("szFirst_Name")
	Session.Contents("strTeacherLastName") = rsTeacher("szLast_Name")
	Session.Contents("intInstructor_ID") = Request.QueryString("intInstructor_ID")	
	rsTeacher.Close
	set rsTeacher = nothing
	
	'This section gets the classes for a teacher
	sql = "select c.intClass_ID,c.szClass_Name as szCourse_Title, '' as txtCourseTitle, c.intInstructor_ID,c.intInstruct_Type_id, i.szFirst_Name + ' ' + i.szLast_Name as teacherName, " & _
			"gi.intILP_ID as Class_ILP ,'' as strILPList, c.intGuardian_id,c.intVendor_id, gi.decCourse_Hours " & _
			"from tblInstructor i,tblClasses c left outer join tblILP_Generic gi " & _
			" ON c.intClass_id = gi.intClass_ID " & _
			"where i.intInstructor_ID = c.intInstructor_ID and " & _
		    "i.intInstructor_ID =" &  Request.QueryString("intInstructor_ID") & _ 
		    " and (c.intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _
		    " order by c.szClass_Name " 
	' This section creates the select list for teachers
	dim sqlTeachers 
	
	sqlTeachers = "select intInstructor_id, szLast_Name + ', ' + szFirst_Name as Name" & _
					" from tblInstructor order by szLast_Name " 
	strList = oFunc.MakeListSQL(sqlTeachers,"intInstructor_id","Name",Request.QueryString("intInstructor_ID"))
	strURLAddition = "&intInstruct_Type_ID=4&bolFromTeacher=True"
	strIdType = "intInstructor_ID"
	
elseif Request.QueryString("intVendor_ID") <> "" then	   
	'This section gets the classes for a vendor
	Session.Contents("intVendor_ID") = Request.QueryString("intVendor_ID")
	sql = "select c.intClass_ID,c.szClass_Name,v.szVendor_Name as teacherName,c.intInstructor_ID,c.intInstruct_Type_id," & _
			"c.intGuardian_id,c.intVendor_id, gi.intILP_ID as Class_ILP " & _
			"from tblVendors v,tblClasses c left outer join tblILP_Generic gi " & _
			" ON c.intClass_id = gi.intClass_ID " & _
			"where v.intVendor_ID = c.intVendor_ID and " & _
		    "v.intVendor_ID =" &  Request.QueryString("intVendor_ID") & _ 
		    " and (c.intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _
		    " order by c.szClass_Name "  	
	
	' This section creates the select list for Vendors
	dim sqlVendors		
	sqlVendors = "select intVendor_id, szVendor_Name " & _
					" from tblVendors order by szVendor_Name " 
	strList = oFunc.MakeListSQL(sqlVendors,"intVendor_id","szVendor_Name",Request.QueryString("intVendor_ID"))
	strIdType = "intVendor_ID"
end if 

set rsClasses = server.CreateObject("ADODB.RECORDSET")
rsClasses.CursorLocation = 3
rsClasses.Open sql, Application("cnnFPCS")'oFunc.FPCScnn

%>
<script language=javascript>
	function jfGo(class_id,instructor_id,instruct_type,intContract_Guardian_ID,intGuardian_ID,intVendor_ID,script) {
		var classWin;
		var strURL = script + "?bolInWindow=true&plain=yes<%=strDisabled%>&intClass_id="+class_id;
		strURL += "&intInstructor_id="+instructor_id+"&intInstruct_Type_ID="+instruct_type;
		strURL += "&intContract_Guardian_ID="+intContract_Guardian_ID;
		strURL += "<% = strHideGoodService %>&intGuardian_id="+intGuardian_ID;
		strURL += "&intVendor_ID="+intVendor_ID;
		<% IF Request.QueryString("intStudent_id") <> "" then %>
		strURL += "&intStudent_ID=<%=Request.QueryString("intStudent_id")%>";
		<% end if %>
		classWin = window.open(strURL,"classWin","width=640,height=500,scrollbars=yes,resizable=yes");
		classWin.moveTo(0,0);
		classWin.focus();
	}
	
	function jfViewILP(bolGeneric,ilp_id,class_ID,class_name,cg,vendor,teacherName,script) {
		var ilpWin;
		var strURL;
		var strILP;
		
		if (bolGeneric == "true")	{
			strILP = "intILP_ID_Generic";
		}else{
			strILP = "intILP_ID";
		}
		strURL = script + "?plain=yes&"+strILP+"=" + ilp_id + "&intClass_id=" + class_ID;
		strURL += "&szClass_Name=" + class_name;
		strURL += "&intVendor_ID=" + vendor;
		strURL += "&strTeacherName=" + teacherName;
		strURL += "&intContract_Guardian_ID=" + cg;
		<% IF Request.QueryString("intStudent_id") <> "" then %>
		strURL += "&intStudent_ID=<%=Request.QueryString("intStudent_id")%>";
		<% end if %>
		ilpWin = window.open(strURL,"ilpWin","width=710,height=500,scrollbars=yes,resizable=yes");
		ilpWin.moveTo(0,0);
		ilpWin.focus();
	}
	
	function jfViewNext(item){
		var strURL = "<%=Application.Value("strWebRoot")%>forms/teachers/viewClasses.asp?<%=strIdType%>=" + item.value 
		window.location.href = strURL + "<%=strURLAddition%>";
	}
		
	function jfViewCosts(studentID,ilpID,classID){
		var strURL = "<%=Application.Value("strWebRoot")%>forms/Requisitions/req1.asp?intClass_ID="+classID;
		strURL += "&intStudent_ID=" + studentID + "&intILP_ID=" + ilpID;
		var costsWin = window.open(strURL,"costsWin","width=710,height=500,scrollbars=yes,resize=yes,resizable=yes");
		costsWin.moveTo(0,0);
		costsWin.focus();
	}
	
	function jfPrintScreens(strURL){
		var winPrint;
		<% IF Request.QueryString("intStudent_id") <> "" then %>
		strURL += "&intStudent_ID=<%=Request.QueryString("intStudent_id")%>";
		<% end if %>
		var winPrint = window.open(strURL,"winPrint","width=710,height=500,scrollbars=yes,resize=yes,resizable=yes");
		winPrint.moveTo(0,0);
		winPrint.focus();
	}
	
	function jfReimburse(studentID){
		var strURL = "<%=Application.Value("strWebRoot")%>forms/Requisitions/reimburseForm.asp?";
		strURL += "intStudent_ID=" + studentID;
		var reimWin = window.open(strURL,"reimWin","width=710,height=500,scrollbars=yes,resizable=yes");
		reimWin.moveTo(0,0);
		reimWin.focus();
	}
	
	function jfBudget(studentID){
		var strURL = "<%=Application.Value("strWebRoot")%>forms/Budget/BudgetWorkSheet.asp?print=true&";
		strURL += "intStudent_ID=" + studentID;
		var reimWin = window.open(strURL,"reimWin","width=710,height=500,scrollbars=yes,resizable=yes");
		reimWin.moveTo(0,0);
		reimWin.focus();
	}
	
	function jfPrintAll(class_ID,ilp_ID,action){
		var winPrint;
		var strURL = "<%=Application.Value("strWebRoot")%>forms/PrintableForms/allPrintable.asp?intClass_ID="+class_ID;
		strURL += "&intStudent_ID=<%=Request("intStudent_id")%>";
		strURL += "&intILP_ID=" + ilp_ID + "&strAction=" + action;
		winPrint = window.open(strURL,"winPrint","width=710,height=500,scrollbars=yes,resize=yes,resizable=yes");
		winPrint.moveTo(0,0);
		winPrint.focus();
	
	}
</script>
<form  name=main ID="Form1">
<table width=100% ID="Table1" cellpadding='4'>
	<tr>	
		<Td class=yellowHeader >
				&nbsp;<b>Printable Forms
				<% if false then 'strList <> "" and Session.Contents("strRole") = "ADMIN" then %>
				<select name="<%=strFieldName%>" onChange="jfViewNext(this);" ID="Select1">
					<% = strList %>
				</select>
				<%
				   elseif Session.Contents("strRole") = "TEACHER" then 
						Response.Write Session.Contents("strFullName")
				   end if 
				%>
		</td>
	</tr>
	<tr>
		<td bgcolor=f7f7f7>
			<% = strStudentInfo %>
			<table ID="Table2">
				<tr>	
					<Td class="TableHeader" valign=middle align=center>
						&nbsp;<B>Class Name</b>&nbsp;
					</td>
					<Td class="TableHeader" valign=middle align=center>
						&nbsp;<b>Contract&nbsp;
					</td>					
					<Td class="TableHeader" valign=middle align=center>
						&nbsp;<b>ILP</b>&nbsp;
					</td>
					<% if false then %>
					<Td class="TableHeader" valign=middle align=center>
						&nbsp;<b>Goods/<BR>Services</b>&nbsp;
					</td>
					<% end if %>
					<Td class="TableHeader" valign=middle align=center>
						&nbsp;<b>Class Packet</b>&nbsp;
					</td>
				</tr>
<%	
		intColorCount = 0
		if rsClasses.RecordCount > 0 then
				do while not rsClasses.EOF	
					if bolGetGuardian = true then
						 intContract_Guardian_ID = rsClasses("intContract_Guardian_ID")
					end if
					if intColorCount mod 2 = 0 then
						strBgColor = " bgcolor=white " 
					else
						strBgColor = ""
					end if 
					
					if rsClasses("intInstructor_ID") <> "" then
						strConSched ="Contracted Class"
					else
						strConSched = "Scheduled Class"
					end if
%>
				<tr <% = strBgColor %>>
					<Td class="TableCell" title="<%=strConSched%>">
						&nbsp;<%  if rsClasses("szClass_Name") & ""<> "" then response.write rsClasses("szClass_Name") else response.write rsClasses("szCourse_Title") & rsClasses("txtCourseTitle")%>&nbsp;
					</td>					
					<td align=center class="TableCell">
						<input type=button value="Print" class="btSmallGray" onCLick="jfPrintAll('<% =rsClasses("intClass_ID")%>','','C');" NAME="Button1">						
					</td>
					<%
					strClassList = strClassList & rsClasses("intClass_ID") & "," & rsClasses("intInstructor_ID") & "," & rsClasses("intInstruct_Type_id") & "," & intContract_Guardian_ID & "," & rsClasses("intGuardian_ID") & "," & rsClasses("intVendor_ID") & "|"
					if rsClasses("Class_ILP") <> "" then
					%>
					<Td class="TableCell" valign=middle>
						<input type=button value="Print" class="btSmallGray" onCLick="jfPrintAll('','<% =rsClasses("Class_ILP")%>','I');" NAME="Button3">							
					</td>
					<%
						if request("intSTUDENT_ID") <> "" then
							strILPList = strILPList & bolGeneric & "," & rsClasses("Class_ILP") & "," & rsClasses("intClass_ID") & "," & Replace(rsClasses("szClass_Name"), "'", "\'") & "," & intContract_Guardian_ID & "," & rsClasses("intVendor_ID") & "," & Replace(rsClasses("teacherName"), "'", "\'") & "|"
						end if
					else
					%>
					<Td align=center class=svplain10 class="TableCell">
						No ILP Provided
					</td>
					<%
					end if 
					%>
					<% if false then %>
					<Td align=center class="TableCell">
						<input type=button value="View/Edit" class="btSmallGray" onCLick="jfViewCosts('<%=intStudent_ID%>','<% =rsClasses("Class_ILP")%>','<% =rsClasses("intClass_ID")%>');" NAME="Button5">						
					</td>	
					<% end if %>
					<Td align=center class="TableCell">
						<input type=button value="Print" class="btSmallGray" onCLick="jfPrintAll('<% =rsClasses("intClass_ID")%>','<% =rsClasses("Class_ILP") %>','');" NAME="Button3">
					</td>			
				</tr>
<%				rsClasses.MoveNext
				if len(strILPList) > 0 then
					session.Contents("strILPList") = left(strILPList,len(strILPList)-1)
				end if 
				session.Contents("strClassList") = left(strClassList,len(strClassList)-1)
				intColorCount = intColorCount + 1 
			loop	
		else
%>
				<tr>	
					<Td colspan=2 class=gray>
						&nbsp;No Scheduled Classes.
					</td>
				</tr>
<%
		end if 
	rsClasses.Close
	set rsClasses = nothing	
	call oFunc.CloseCN
	set oFunc = nothing
%>			
			</table>
		</td>
	</tr>
</table>

<table>
	<tr>
		<td class=gray>
			Other Print Options: 
		</td>
		<td>
			<select name="selPrintOption">
				<option>Make Selection</option>
				<option value="printPacket">Student Packet/Budget</option>
				<option value="printPhilosophy">ILP Philosophy</option>
				<% if len(strClassList) > 0 then %>
				<option value="printAllContracts">All Contracts (ASD Only)</option>
				<% end if %>
				<% if len(strILPList) > 0 then %>
				<option value="printAllILPs">All ILPs</option>
				<% end if %>								
				<option value="printAll">Entire Packet (all of above)</option>
				<option value="printReimbursements">Reimbursements</option>
			</select>
		</td>
		<td>
			<input type=button value="Print" onclick="jfExecutePrint(this.form.selPrintOption.value);" class="btSmallGray">
		</td>
	</tr>
</table>
</form>
<script language=javascript>
	function jfExecutePrint(pOption){
		if (pOption == 'printAllContracts'){
			jfPrintScreens('allPrintable.asp?intClass_ID=ALL&strAction=C');
		}else if (pOption == 'printAllILPs'){
			jfPrintScreens('allPrintable.asp?intILP_ID=ALL&strAction=I');
		}else if (pOption == 'printPacket'){
			//jfPrintScreens('allPrintable.asp?strAction=S');
			jfPrintScreens('printPacket2.asp?strAction=S');
		}else if (pOption == 'printReimbursements'){
			jfReimburse('<% = intStudent_ID %>');
		}else if (pOption == 'printBudget'){
			jfPrintScreens('allPrintable.asp?strAction=B');
		}else if (pOption == 'printAll'){
			jfPrintScreens('allPrintable.asp?strAction=A');
		}else if (pOption == 'printTesting'){
			jfPrintScreens('allPrintable.asp?strAction=T');
		}else if (pOption == 'printProgress'){
			jfPrintScreens('allPrintable.asp?strAction=P');
		}else if (pOption == 'printPhilosophy'){
			jfPrintScreens('allPrintable.asp?strAction=IP');
		}
	}	
</script>	
<%
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")
%>