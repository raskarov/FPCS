<%@ Language=VBScript %>
<%
' TOGGLES SHOWING GOODS/SERVICES 
'if session.Contents("strRole") = "ADMIN" then'
'	bolShow = true
'else
'	bolShow = false
'end if
bolShow = true
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
dim intActualPercent
dim strStudentInfo
dim strConSched			'tells us the difference between a contacted or scheduled class
dim strHideGoodService

Session.Contents("strTitle") = "View Classes"
Session.Contents("strLastUpdate") = "22 Feb 2002"
Server.Execute(Application.Value("strWebRoot") & "includes/header.asp")
Session.Contents("blnFromClassAdmin") = ""
session.Contents("strILPList") = ""
session.Contents("strClassList") = ""
Session.Contents("intILP_ID") = ""

if Session.Contents("bolUserLoggedIn") = false then
	Response.Expires = -1000	'Makes the browser not cache this page
	Response.Buffer = True		'Buffers the content so our Response.Redirect will work
	Session.Contents("strURL") = Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
	Server.Execute(Application.Value("strWebRoot") & "UserAdmin/Login.asp")
else 
   set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
   call oFunc.OpenCN()
   
   set oHtml = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/htmlFuncs.wsc"))
	'Reset Session Variables
	Session.Contents("intStudent_ID") = ""
	Session.Contents("intInstructor_ID") = ""
	oFunc.ResetSelectSessionVariables() 

	bolGetGuardian = false
	
	if request("intClass_ID") <> "" and request("pValue") <> "" and request("intInstructor_ID") <> "" then
		call vbsUpdateContractStatus(request("intClass_ID"),request("pValue"),request("intInstructor_ID"))
	end if
	
	if request("myType") <> "" and request("intClass_ID") <> "" then
		call vbsUpdateComments(request("intClass_ID"),request("message"),request("myType"))
	end if
	if Request.QueryString("intInstructor_ID") <> "" then  
		
		set rsTeacher = server.CreateObject("ADODB.RECORDSET")
			rsTeacher.CursorLocation = 3
			sqlTeacher = "select szFirst_Name,szLast_Name " & _
						 "from tblInstructor where intInstructor_ID=" & Request.QueryString("intInstructor_ID")
			rsTeacher.Open sqlTeacher,oFunc.FPCScnn	

		Session.Contents("strTeacherFirstName") = rsTeacher("szFirst_Name")
		Session.Contents("strTeacherLastName") = rsTeacher("szLast_Name")
		Session.Contents("intInstructor_ID") = Request.QueryString("intInstructor_ID")	
		rsTeacher.Close
		set rsTeacher = nothing
		
		'This section gets the classes for a teacher
		Session.Contents("intStudent_ID") = ""
		sql = "select c.intClass_ID,c.szClass_Name, c.intInstructor_ID,c.intInstruct_Type_id, i.szFirst_Name + ' ' + i.szLast_Name as teacherName, " & _
			  "gi.intILP_ID as Class_ILP , c.intGuardian_id,c.intVendor_id, gi.decCourse_Hours, " & _
			  "(SELECT COUNT(i2.intILP_ID) AS total " & _ 
						   " FROM tblILP i2 " & _ 
					       " WHERE      i2.intClass_ID = c.intClass_ID) AS Enrolled, " & _
			  "c.intMax_Students,c.intMin_Students, c.decHours_Student + c.decHours_Planning as hourTotal, " & _
			  "c.intContract_Status_ID, c.dtReady_For_Review, c.dtApproved, c.szUser_Approved, " & _
			  "c.szInstructor_Comments, c.szComments, cst.szContract_Status_Name " & _					       
			  "from tblInstructor i INNER JOIN " & _
			  " tblClasses c ON i.intInstructor_ID = c.intInstructor_ID LEFT OUTER JOIN " & _
			  " tblILP_Generic gi  ON c.intClass_id = gi.intClass_ID LEFT OUTER JOIN " & _
			  " tblContract_Status_Types cst ON c.intContract_Status_Id = cst.intContract_Status_Id " & _
			  "where i.intInstructor_ID =" &  Request.QueryString("intInstructor_ID") & _ 
		      " and (c.intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _
		      " order by c.szClass_Name " 
		'if ucase(session.contents("strUserId")) = "SCOTT" then response.write sql      

		' This section creates the select list for teachers
		dim sqlTeachers 
		
		sqlTeachers = "select intInstructor_id, szLast_Name + ', ' + szFirst_Name as Name" & _
					  " from tblInstructor order by szLast_Name " 
		strList = oFunc.MakeListSQL(sqlTeachers,"intInstructor_id","Name",Request.QueryString("intInstructor_ID"))
		strURLAddition = "&intInstruct_Type_ID=4&bolFromTeacher=True"
		strIdType = "intInstructor_ID"
		
	elseif Request.QueryString("intVendor_ID") <> "" then	   
		'This section gets the classes for a teacher
		Session.Contents("intStudent_ID") = ""
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
	rsClasses.Open sql, oFunc.FPCScnn
'response.Write "<B>TESTSING...<BR></b>" & sql
%>
<script language=javascript>
	function jfGo(class_id,instructor_id,instruct_type,intContract_Guardian_ID,intGuardian_ID,intVendor_ID,script) {
		var classWin;
		var strURL = script + "?bolInWindow=true&isPopUp=yes<%=strDisabled%>&intClass_id="+class_id;
		strURL += "&intInstructor_id="+instructor_id+"&intInstruct_Type_ID="+instruct_type;
		strURL += "&intContract_Guardian_ID="+intContract_Guardian_ID;
		strURL += "<% = strHideGoodService %>&intGuardian_id="+intGuardian_ID;
		strURL += "&intVendor_ID="+intVendor_ID;
		classWin = window.open(strURL,"classWin","width=720,height=500,scrollbars=yes,resizable=yes");
		classWin.moveTo(0,0);
		classWin.focus();
	}
	function jfDeleteClass(class_id) {
		var answer;
		answer = confirm("Are you sure you want to delete this class?");
		if (answer) {
			var winDel;
			winDel = window.open("deleteClass.asp?intClass_id="+class_id,"winDel","width=200,height=200,scrollbars=yes,resizable=yes");
			winDel.moveTo(0,0);
			winDel.focus();			
		}
	}
	function jfDropStudents(class_id) {
		var answer;
		answer = confirm("Are you sure you want to drop all students from the class?");
		if (answer) {
			var winDel;
			winDel = window.open("deleteClass.asp?studentdrop="+class_id,"winDel","width=200,height=200,scrollbars=yes,resizable=yes");
			winDel.moveTo(0,0);
			winDel.focus();			
		}
	}
	function jfDeleteILP(ilp_id) {
		var answer;
		answer = confirm("Are you sure you want to delete this class?");
		if (answer) {
			var winDel;
			winDel = window.open("deleteClass.asp?intILP_id="+ilp_id,"winDel","width=200,height=200,scrollbars=yes,resizable=yes");
			winDel.moveTo(0,0);
			winDel.focus();			
		}
	}	
	function jfViewRoll(class_id) {
		var winRoll;
		winRoll = window.open("../../Reports/studentsInClass.asp?intClass_id="+class_id,"winRoll","width=640,height=480,scrollbars=yes,resizable=yes");
		winRoll.moveTo(0,0);
		winRoll.focus();					
	}
	
	function jfViewILP(ilp_id,class_ID,class_name,cg,vendor,teacherName,script) {
		var ilpWin;
		var strURL;
		var strILP;
		
		strURL = "../ilp/" + script + "?isPopUp=yes&intILP_ID_Generic=" + ilp_id + "&intClass_id=" + class_ID;
		strURL += "&szClass_Name=" + class_name;
		strURL += "&intVendor_ID=" + vendor;
		strURL += "&strTeacherName=" + teacherName;
		strURL += "&intContract_Guardian_ID=" + cg;
		ilpWin = window.open(strURL,"ilpWin","width=710,height=500,scrollbars=yes,resizable=yes");
		ilpWin.moveTo(0,0);
		ilpWin.focus();
	}
	
	function jfViewNext(item){
		var strURL = "<%=Application.Value("strWebRoot")%>forms/teachers/viewClasses.asp?<%=strIdType%>=" + item.value 
		window.location.href = strURL + "<%=strURLAddition%>";
	}
	
	function jfContractStatus(pClassID,pValue){
		if (pValue ) {pValue = 2;} else {pValue = 1;}
		if (pValue == '') { pValue = 0;}
		var strURL = "<%=Application.Value("strWebRoot")%>forms/teachers/viewClasses.asp?intClass_ID=" + pClassID;
		strURL += "&pValue=" + pValue + "&intInstructor_ID=<% = request("intInstructor_ID")%>"; 
		window.location.href = strURL + "<%=strURLAddition%>";
	}
	
	function jfSaveComment(txtObj, pClassID, myType){
		var myObj = document.getElementById(txtObj);
		var sMsg = myObj.value.replace(/\'/g," ");
		sMsg = sMsg.replace(/\"/g," ");
		sMsg = sMsg.replace(/\&/g," ");
		var strURL = "<%=Application.Value("strWebRoot")%>forms/teachers/viewClasses.asp?intClass_ID=" + pClassID;
		strURL += "&myType=" + myType + "&intInstructor_ID=<% = request("intInstructor_ID")%>&message="+sMsg; 
		window.location.href = strURL + "<%=strURLAddition%>";
	}
	
	
	function jfAddILP(classID,className,teacherName){
		var strURL = "<%=Application.Value("strWebRoot")%>forms/ILP/ILPMain.asp?intClass_ID="+classID+"&szClass_Name=" +className;
		strURL += "&strTeacherName=" + teacherName + "&bolLateAdd=true";
		addILPWin = window.open(strURL,"addILPWin","width=710,height=500,scrollbars=yes,resize=yes,resizable=yes");
		addILPWin.moveTo(0,0);
		addILPWin.focus();
	}
	
	function jfViewCosts(studentID,ilpID,classID){
		var strURL = "<%=Application.Value("strWebRoot")%>forms/Requisitions/req1.asp?intClass_ID="+classID;
		strURL += "&intStudent_ID=" + studentID + "&intILP_ID=" + ilpID;
		costsWin = window.open(strURL,"costsWin","width=710,height=500,scrollbars=yes,resize=yes,resizable=yes");
		costsWin.moveTo(0,0);
		costsWin.focus();
	}
	
	function jfPrintScreens(strURL){
		var winPrint;
		winPrint = window.open(strURL,"winPrint","width=710,height=500,scrollbars=yes,resize=yes,resizable=yes");
		winPrint.moveTo(0,0);
		winPrint.focus();
	}
</script>
<form  name=main ID="Form1">
<table width=100% ID="Table1" cellpadding="2">
	<tr>	
		<Td class=yellowHeader >
				<b>Manage Classes for: <% =Session.Contents("strStudentName") %></b>
				<% if strList <> "" and Session.Contents("strRole") = "ADMIN" then %>
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
	<% if Application.Contents("bolUseContractApproval"&session.Contents("intSchool_Year")) and request("intInstructor_ID") <> "" then%>
	<tr>
		<td class="svplain8">
			<b>Contract Status Instructions</b>
			<ul>
				<li>Each ASD contract (contract, ILP and associated goods/services) created must be approved by the Principal before students can enroll in the class.</li>
				<li>Once you are ready for the Principal to review the ASD contract, click the 'RFR' (Ready For Review) checkbox for the class.</li>
				<li>After the Pricipal has reviewed the contract you will see a status of 'Needs Work', 'Rejected' or 'Approved'. </li>
				<li>Click on 'View/Make' under the 'Comments' column to read or make comments. If a contract has a comment the 'View/Make' link will be <span class='yellow'>yellow</span>.</li>
			</ul>
		</td>
	</tr>
	<% end if %>
	<tr>
		<td bgcolor=f7f7f7>
			<% = strStudentInfo %>
			<table ID="Table2">
				<% call vbsClassHeader %>
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
						strBgColor = " bgcolor='#ececec' "
					end if 
					
					if intColorCount mod 20 = 0 and intColorCount <> 0 then
						call vbsClassHeader
					end if 
					
					if rsClasses("intInstructor_ID") <> "" then
						strConSched ="Contracted Class"
					else
						strConSched = "Scheduled Class"
					end if
%>
				<tr <% = strBgColor %>>
					<Td class="TableCell" title="<%=strConSched%>">
						<% = rsClasses("szClass_Name")%>
					</td>
					<% if Application.Contents("bolUseContractApproval"&session.Contents("intSchool_Year")) then %>
					<td class="TableCell" align="center">
						<% if rsClasses("intContract_Status_ID") < 4 or rsClasses("intContract_Status_ID") & "" = "" then %>
							<input type="checkbox" value="1" name="ContractStatus<% = rsClasses("intClass_Id")%>" <% if rsClasses("intContract_Status_ID") = 2 then response.Write " checked "%> onClick="jfContractStatus('<% = rsClasses("intClass_ID")%>',this.checked);">RFR 
						<% 
								strBR = "<BR>"
							end if
						   if rsClasses("intContract_Status_ID") > 2 then 
								response.Write strBR
								response.Write rsClasses("szContract_Status_Name") 
							end if
							strBR = ""
						 %>
					</td>
					<% end if %>
					<td class="TableCell" align="center">
						<%
							strComment = "<table style='width:100%;'>"
							
							if oFunc.IsAdmin then
								strComment = strComment & "<tr><td class='TableHeader'>Principal Comments</td></tr><tr><td><textarea onChange=""document.getElementById('Comment" & rsClasses("intClass_ID") & "').value = this.value;"" "  & _
											" rows=5 style='width:100%;'>" & rsClasses("szComments") & "</textarea><input type=hidden name='Comment" & rsClasses("intClass_ID") & "' value='" & rsClasses("szComments") & "'></td></tr>" & _
											"<tr><td class='TableHeader'>Teacher Comments</td></tr>" & _
											"<tr><td class='tablecell'>" & oHtml.IIF(rsClasses("szInstructor_Comments") & "" = "", "No comments.",rsClasses("szInstructor_Comments")) & "</td></tr>"
								strType = "szComments"
							elseif oFunc.IsTeacher then
								strComment = strComment & "<tr><td class='TableHeader'>Teacher Comments</td></tr><tr><td><textarea onChange=""document.getElementById('Comment" & rsClasses("intClass_ID") & "').value = this.value;"" "  & _
											" rows=5 style='width:100%;'>" & rsClasses("szInstructor_Comments") & "</textarea><input type=hidden name='Comment" & rsClasses("intClass_ID") & "' value='" & rsClasses("szInstructor_Comments") & "'></td></tr>" & _
											"<tr><td class='TableHeader'>Principal Comments</td></tr>" & _
											"<tr><td class='tablecell'>" & oHtml.IIF(rsClasses("szComments") & "" = "", "No comments.",rsClasses("szComments")) & "</td></tr>"
								strType = "szInstructor_Comments"
							end if
							strComment = strComment & "<tr><td align=right><input type='button' class='btSmallGray' value='SAVE' onclick=""jfSaveComment('Comment" & rsClasses("intClass_ID") & "', '" & rsClasses("intClass_ID") & "','" & strType & "');""></td></tr></table>"
							response.Write oHtml.ToolTip("<u>" & oHtml.IIF(rsClasses("szInstructor_Comments") & "" <> "" or rsClasses("szComments") & "" <> "", "<span class='yellow'>View/Make</span>","View/Make") & "</u>",strComment,true,"Comments for: " & left(rsClasses("szClass_Name"),50) ,false,"tooltip","450px","",false,true) 
						%>
					</td>
					<td align=center class="TableCell">
						<nobr><input type=button value="(<% = rsClasses("hourTotal") %> hrs) View/Edit" class="btSmallGray" onCLick="jfGo('<% =rsClasses("intClass_ID")%>','<% =rsClasses("intInstructor_ID")%>','<% = rsClasses("intInstruct_Type_id") %>','<%=intContract_Guardian_ID%>','<%=rsClasses("intGuardian_ID")%>','<%=rsClasses("intVendor_ID")%>','classAdmin.asp');" NAME="btSmallGray" style="width:100px;">						
						</nobr>
					</td>
					<%
					strClassList = strClassList & rsClasses("intClass_ID") & "," & rsClasses("intInstructor_ID") & "," & rsClasses("intInstruct_Type_id") & "," & intContract_Guardian_ID & "," & rsClasses("intGuardian_ID") & "," & rsClasses("intVendor_ID") & "|"
					if rsClasses("Class_ILP") <> "" then
					%>
					<Td class="TableCell">
						<input type=button value="(<% = rsClasses("decCourse_Hours")%> hrs) View/Edit" class="btSmallGray" onCLick="jfViewILP('<% =rsClasses("Class_ILP")%>','<% =rsClasses("intClass_ID")%>','<% = server.URLEncode(Replace(rsClasses("szClass_Name"), "'", "\'"))%>','<%=intContract_Guardian_ID%>','<%=rsClasses("intVendor_ID")%>','<%=Replace(rsClasses("teacherName"), "'", "\'")%>','ilpMain.asp');" NAME="Button2" style="width:100px;">
					</td>
					<%
						strILPList = strILPList  & "," & rsClasses("Class_ILP") & "," & rsClasses("intClass_ID") & "," & Replace(rsClasses("szClass_Name"), "'", "\'") & "," & intContract_Guardian_ID & "," & rsClasses("intVendor_ID") & "," & Replace(rsClasses("teacherName"), "'", "\'") & "|"
					else
					%>
					<Td align=center class="TableCell">
						<input type=button value="Add ILP" class="btSmallGray" onClick="jfAddILP('<% =rsClasses("intClass_ID")%>','<% =Replace(rsClasses("szClass_Name"), "'", "\'")%>','<%=Replace(rsClasses("teacherName"), "'", "\'")%>');" NAME="Button4">
					</td>
					<%
					end if 
					%>
					<% if bolShow = true then %>
					<Td align=center class="TableCell">
						<input type=button value="View/Edit" class="btSmallGray" onCLick="jfViewCosts('<%=intStudent_ID%>','<% =rsClasses("Class_ILP")%>','<% =rsClasses("intClass_ID")%>');" NAME="Button5">						
					</td>
					<% end if %>
					<% if Request.QueryString("intStudent_ID") <> "" then 
							if session.Contents("strRole") = "ADMIN" then 
					%>
					<td align=center class="TableCell" valign="middle">		
								<% if not oFunc.LockYear then
								%>
								<input type=button value="DELETE" class="btSmallGray" onCLick="jfDeleteILP('<% =rsClasses("Class_ILP")%>');" NAME="Button6">&nbsp;
								<% end if %>
					</td>
						<%  end if 	
					    else
						%>
					<td align=center class="TableCell" valign="middle">	
							<% if not oFunc.LockYear and rsClasses("Enrolled") < 1 then%>
								<input type=button value="Delete" class="btSmallGray" onCLick="jfDeleteClass('<% =rsClasses("intClass_ID")%>','<% =rsClasses("intInstructor_ID")%>','<% = intStudent_ID %>');" NAME="Button7">						
							<% else %>
								N/A
							<% end if %>
					</td>
					<td align=center class="TableCell" valign="middle">
					<% if not oFunc.LockYear and rsClasses("Enrolled") >0 then%>
					<input type=button value="Drop Students" class="btSmallGray" onCLick="jfDropStudents('<% =rsClasses("intClass_ID")%>');" NAME="Button10">&nbsp;					
					<% else %>
					N/A
					<% end if %>
					</td>
					<% end if %>
					<td align=center class="TableCell">
						<input type=button value="C" class="btSmallGray" onCLick="jfPrintScreens('../PrintableForms/allPrintable.asp?strAction=C&intClass_ID=<%=rsClasses("intClass_ID")%>');" NAME="Button1">
						<input type=button value="I" class="btSmallGray" onCLick="jfPrintScreens('../PrintableForms/allPrintable.asp?strAction=I&intILP_ID=<%=rsClasses("Class_ILP")%>');" NAME="Button3"></nobr>						
						<input type=button value="GS" class="btSmallGray" onCLick="jfPrintScreens('../PrintableForms/allPrintable.asp?strAction=G&intClass_ID=<%=rsClasses("intClass_ID")%>&intILP_ID=<%=rsClasses("Class_ILP")%>&intStudent_ID=<%=intStudent_ID%>');" NAME="Button1">
					</td>
					<td align=center class="TableCell">
						<input type=button value="<% = rsClasses("Enrolled") & "/" & rsClasses("intMax_Students")%> Enrolled" class="<% if rsClasses("Enrolled") < rsClasses("intMin_Students") then response.Write "btSmallRed" else response.Write "btSmallGray" %>" onCLick="jfViewRoll('<% =rsClasses("intClass_ID")%>');" NAME="Button8" style="width:75px;">
					</td>
				</tr>
<%				
				dblTotalContractHrs = dblTotalContractHrs + cdbl(rsClasses("hourTotal"))
				rsClasses.MoveNext
				if len(strILPList) > 0 then
					'session.Contents("strILPList") = left(strILPList,len(strILPList)-1)
				end if 
				'session.Contents("strClassList") = left(strClassList,len(strClassList)-1)
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
				<tr class="svplain8"> 
					<td colspan="2" align="right">
						Total Contract Hours:
					</td>
					<td>
						&nbsp;<% = formatNumber(dblTotalContractHrs,1)%>
					</td>
					<td colspan="5">
						&nbsp;
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<input type=button value="Home Page" onClick="window.location.href='<%=Application.Value("strWebRoot")%>';" class="btSmallGray" NAME="Button9">
<%
if len(strClassList) > 0 then
%>
<input type=button value="Print all Contracts/Schedules" class="btSmallGray" onClick="jfPrintScreens('../PrintableForms/allPrintable.asp?straction=c&intInstructor_ID=<%=request("intInstructor_ID")%>&intClass_ID=all');" NAME="Button10">
<% 
end if 

if len(strILPList) > 0 then %>
		<input type=button value="Print all ILP's"  class="btSmallGray" onClick="jfPrintScreens('../PrintableForms/allPrintable.asp?straction=i&intInstructor_ID=<%=request("intInstructor_ID")%>&intILP_ID=all');" NAME="Button11">
<% end if 
response.Write oHtml.ToolTipDivs
%>
</form>
<%
end if
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")

sub vbsClassHeader()
%>
				<tr>	
					<Td class="TableHeader" valign=middle align=center>
						<B>Class Name</b>
					</td>	
					<% if Application.Contents("bolUseContractApproval"&session.Contents("intSchool_Year")) then %>
					<Td class="TableHeader" valign=middle align=center>
						<B>Contract<BR>Status</b>
					</td>				
					<% end if %>
					<Td class="TableHeader" valign=middle align=center>
						<B>Comments</b>
					</td>	
					<Td class="TableHeader"  valign=middle align=center>
						<b>Contract/<BR>Schedules</b>
					</td>					
					<Td class="TableHeader"  valign=middle align=center>
						<b>ILP</b>
					</td>
					<% if bolShow = true then %>
					<Td class="TableHeader"  valign=middle align=center>
						<b>Goods/<BR>Services</b>
					</td>
					<% end if %>
					<% if session.Contents("strRole") <> "GUARD" then %>
					<Td class="TableHeader"  valign=middle align=center>
						<b>Delete</b>
					</td>
					<Td class="TableHeader" valign=middle align=center>
					<b>Drop students</b>
					</td>
					<% end if %>
					<Td class="TableHeader"  valign=middle align=center>
						<b>Print Forms</b>
					</td>
					<Td class="TableHeader" valign=middle align=center>
						<b>Enrolled List</b>
					</td>
				</tr>
<%				
end sub

sub vbsUpdateContractStatus(pClassID,pValue,pInstructorID)
	dim update,strTime
	
	if pValue = 1 then
		strTime = " NULL "
	else
		strTime = " CURRENT_TIMESTAMP "
	end if 
	
	update = "update tblClasses set intContract_Status_ID = " & pValue & _
			 ", dtReady_For_Review =  " & strTime & " " & _
			 "WHERE intClass_ID = " & pClassID & _
			 " AND intInstructor_ID = " & pInstructorID
	oFunc.ExecuteCN(update)
end sub

sub vbsUpdateComments(pClassID,pMessage,pType)
	dim update

	update = "update tblClasses set " &  pType & " = '" & oFunc.EscapeTick(pMessage) & _
			 "', dtModify = CURRENT_TIMESTAMP, szUser_Modify = '" & oFunc.EscapeTick(session.contents("szUserId")) & "' "  & _
			 "WHERE intClass_ID = " & pClassID 
			 
	oFunc.ExecuteCN(update)
end sub
%>