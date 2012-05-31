<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		classSearch.asp
'Purpose:	Teacher Class Search Engine
'Date:		7 APRIL 2005
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim sql
dim oFunc
dim rs, strSqlFamily

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

'JD: Deny access if VENDOR
if ucase(session.Contents("strRole")) = "VENDOR" then
	'terminate page since page was improperly called.
	response.Write "<html><body><H1>Page improperly called.</h1></body></html>"
	response.End
end if
'JD

if request("bolWin") = "" then
	Server.Execute(Application.Value("strWebRoot") & "includes/header.asp")
else
	Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")
end if

if request("intStudent_ID") <> "" then
	set oBudget = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/StudentBudgetInfo.wsc"))
	'oBudget.PopulateStudentFunding oFunc.FPCScnn, request("intStudent_ID"), session.Contents("intSchool_Year") 
	oBudget.PopulateStudentFunding Application("cnnFPCS"), request("intStudent_ID"), session.Contents("intSchool_Year") 
	strSqlFamily = " AND tascClass_Family.intFamily_ID <> " & oBudget.FamilyID & " "
	set oBudget = nothing
	
	strSqlNoDualEnroll = " AND (NOT EXISTS " & _
				"  (SELECT  intILP_Id " & _
				"  FROM  tblILP i " & _
				"  WHERE tblClasses.intClass_id = i.intClass_Id AND i.intStudent_id = " & request("intStudent_ID") & ")) " 
				
	addClassQueryStr = "intStudent_ID=" & request("intStudent_ID") & _
					   "&intShort_ILP_ID=" & request("intShort_ILP_ID") & _
					   "&intInstruct_Type_ID=" & request("intInstruct_Type_ID") & _
					   "&intPOS_Subject_ID=" & request("intPOS_Subject_ID")
end if
%>
<script language=javascript>

	function jfPrintAll(class_ID,ilp_ID){
		var winContractApproval;
		var strURL = "<%=Application.Value("strWebRoot")%>forms/PrintableForms/allPrintable.asp?intClass_ID="+class_ID;
		strURL += "&noprint=true&intILP_ID=" + ilp_ID ;
		winContractApproval = window.open(strURL,"winContractApproval","width=710,height=500,scrollbars=yes,resize=yes,resizable=yes");
		winContractApproval.moveTo(20,20);
		winContractApproval.focus();	
	}
	
	function jfAddClass(pInstructorID,pClassID){
		var qString = "<%=Application.Value("strWebRoot")%>forms/ilp/ilp1.asp?<% = addClassQueryStr %>";
		qString += "&bolFromSearch=true&intInstructor_Id=" + pInstructorID + "&intClass_ID=" + pClassID
		window.opener.location.href = qString;
		window.opener.focus();
		window.close();
	}
	
	function jfHighLight(row){
		var obj = document.getElementById('ROW'+row);
		var lastRow = document.main.lastRow.value;
		var lastRowColor = document.main.lastRowColor.value;	
		// Reset last row to its normal state
		if (lastRow != ""){	
			var obj2 = document.getElementById('ROW'+lastRow);
			obj2.className = lastRowColor;
		}
		// Highlight current row and retsain original info
		document.main.lastRowColor.value = obj.className;
		document.main.lastRow.value = row;
		//obj.style.backgroundColor = "e6e6e6";
		obj.className = "SubHeader";
	}
	
	function jfViewBio(pId){
		var sUrl = "<%=Application.Value("strWebRoot")%>forms/teachers/teacherBiosViewer.asp?";
		sUrl += "simpleHeader=true&intInstructor_ID=" + pId;
		winBio= window.open(sUrl,"winBio","width=710,height=500,scrollbars=yes,resize=yes,resizable=yes");
		winBio.moveTo(20,20);
		winBio.focus();	
	}
</script>	
<form name="main" method="post" action="classSearch.asp">
<input type="hidden" name="Search" value="true">
<input type=hidden name="lastRow" ID="Hidden1">
<input type=hidden name="LineItemsChanged" value="," ID="Hidden8">
<input type=hidden name="lastRowColor" ID="Hidden3">
<input type="hidden" name="hdnReset" value="" ID="Hidden2">
<!-- Remaining hiiden variables used when using this page to add a class from forms/ilp/ilp1.asp -->
<input type="hidden" name="intStudent_ID" value="<% = request("intStudent_ID") %>">
<input type="hidden" name="intShort_ILP_ID" value="<% = request("intShort_ILP_ID") %>" ID="Hidden4">
<input type="hidden" name="intInstruct_Type_ID" value="<% = request("intInstruct_Type_ID") %>" ID="Hidden5">
<input type="hidden" name="bolWin" value="<% = request("bolWin") %>" ID="Hidden6">
<table style="width:100%;" ID="Table3">
	<tr>
		<td class="yellowHeader">
			&nbsp;<b>ASD Class Search Engine</b>
		</td>
	</tr>
	<tr>
		<td>
			<table>						
				<tr>
					<td style="width:0%;">
						<table style="width:100%;" cellpadding="2">
							<tr>
								<td class="TableHeader">
									ASD Teacher Name
								</td>
								<td class="TableHeader">
									Subject
								</td>
								<!--<td class="TableHeader">
									Grade
								</td>-->
							</tr>
							<tr>
								<td>
									<select name="intInstructor_ID" onchange="this.form.hdnReset.value='true';">
										<option value="">
									<%
										sql = "SELECT DISTINCT tblINSTRUCTOR.intInstructor_ID, tblINSTRUCTOR.szLAST_NAME + ', ' + tblINSTRUCTOR.szFIRST_NAME as Name " & _ 
												"FROM tblINSTRUCTOR INNER JOIN " & _ 
												" tblClasses ON tblINSTRUCTOR.intINSTRUCTOR_ID = tblClasses.intInstructor_ID " & _ 
												"WHERE tblClasses.intSchool_Year = " & session.Contents("intSchool_Year") & " " & _ 
												"ORDER BY Name "
										Response.Write oFunc.MakeListSQL(sql,"intInstructor_ID","Name",request("intInstructor_ID"))	
									%>
									</select>
								</td>
								<td>
									<% 
										if request("bolWin") <> "" then
											strDisable = " disabled "
									%>
										<input type="hidden" value="<% = request("intPOS_Subject_ID") %>" name="intPOS_Subject_ID">
									<%
										end if
									%>
									<select name="intPOS_Subject_ID"  <% = strDisable %> ID="Select1" onchange="this.form.hdnReset.value='true';">
										<option value="">
									<%
										sql = "select intPOS_Subject_ID, upper(szSubject_Name) Name from trefPOS_Subjects where bolShow = '1' order by szSubject_Name"
										Response.Write oFunc.MakeListSQL(sql,"intPOS_Subject_ID","Name",request("intPOS_Subject_ID"))	
									%>
									</select>	
								</td>
								<!--<td>
									<select name="sGrade" onchange="this.form.hdnReset.value='true';">
										<option value="">
										<% 
										dim strGradeList
										strGradeList = "K ,1 ,2 ,3 ,4 ,5 ,6 ,7 ,8 ,9 ,10,11,12"
										Response.Write oFunc.MakeList(strGradeList,strGradeList,replace(request("sGrade") & ""," ",""))								
										%>
									</select>-->
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td  style="width:0%;">
						<table ID="Table1" style="width:100%;" cellpadding="2">
							<tr>
								<td class="TableHeader">
									Key Word(s) to Search
								</td>
								<td class="TableHeader">
									Meets On
								</td>
								<td class="TableHeader" colspan="4">
									Start Time
								</td>
							</tr>
							<tr>
								<td>
									<input type="text" name="KeyWords" size="40" maxlength="128" value="<% = request("KeyWords") %>" onchange="this.form.hdnReset.value='true';">
								</td>
								<td>
									<select name="szDays_Meet_On" onchange="this.form.hdnReset.value='true';">
										<option value="">
										<% 
										dim sqlDays
										sqlDays = "select strValue,strText from common_lists where intList_ID = 4 order by intOrder"
										Response.Write oFunc.MakeListSQL(sqlDays,"","",request("szDays_Meet_On"))								
										%>
									</select>	
								</td>
								<td valign=top>
									<select name="hourStart" ID="Select2" onchange="this.form.hdnReset.value='true';">
										<option value="">
										<% 
										dim strHour
										strHour = "1,2,3,4,5,6,7,8,9,10,11,12"
										Response.Write oFunc.MakeList(strHour,strHour,request("hourStart"))								
										%>
									</select>
								</td>	
								<td valign=top>
									:
								</td>	
								<td valign=top>
									<select name="minuteStart" ID="Select3" onchange="this.form.hdnReset.value='true';">
										<option value="">
										<% 
										dim strMinute
										strMinute = "00,15,30,45"
										Response.Write oFunc.MakeList(strMinute,strMinute,request("minuteStart"))								
										%>
									</select>
								</td>											
								<td valign=top>
									<select name="amPmStart" ID="Select4" onchange="this.form.hdnReset.value='true';">
										<% 
										dim strAmPm
										strAmPm = "AM,PM"
										Response.Write oFunc.MakeList(strAmPm,strAmPm,request("amPmStart"))								
										%>
									</select>		
								</td>	
							</tr>
							<tr>
								<td class="gray">
									Key Words: Match Exact Words <input type="checkbox" name="searchType" value="exact" <% if request("searchType") <> "" then response.Write " checked "%>  value="true" ID="Radio1" onchange="this.form.hdnReset.value='true';">
								</td>
								<td colspan="2">
									<input type="submit" value="Search!" class="NavSave">
									<% if request("bolWin") <> "" then
									%>
									<input type="button" value="Cancel" onclick="window.opener.focus();window.close();" class="NavSave">
									<%
									   end if
									%>
								</td>
							</tr>
							<tr>
								<td colspan="8" class="svplain8">
									Please note: Classes that have not yet been approved by the principal <BR>will not show up when conducting a search.
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<%

if request.Form("Search") <> "" then
	sql = "SELECT tblINSTRUCTOR.intINSTRUCTOR_ID, tblINSTRUCTOR.szLAST_NAME, tblINSTRUCTOR.szFIRST_NAME, trefPOS_Subjects.intPOS_Subject_ID,  " & _ 
			" trefPOS_Subjects.szSubject_Name, tblClasses.szClass_Name, tblClasses.intMin_Students, tblClasses.intMax_Students, tblClasses.sGrade_Level,  " & _ 
			" tblClasses.sGrade_Level2, tblClasses.intClass_ID, tblClasses.dtClass_Start, tblClasses.dtClass_End, tblClasses.szDays_Meet_On,  " & _ 
			" tblClasses.szStart_Time, tblClasses.szEnd_Time, tblClasses.szSchedule_Comments, i.intILP_ID " & _ 
			", tblClasses.intMax_Students, " & _ 
			"                          (SELECT     COUNT(ii.intILP_ID) " & _ 
			"                            FROM          tblILP ii " & _ 
			"                            WHERE      ii.intClass_ID = tblClasses.intClass_ID) AS SlotsTaken " & _ 
			"FROM tblClasses INNER JOIN " & _ 
			" tblILP_Generic as i ON tblClasses.intClass_ID = i.intClass_ID INNER JOIN " & _ 
			" tblINSTRUCTOR ON tblClasses.intInstructor_ID = tblINSTRUCTOR.intINSTRUCTOR_ID INNER JOIN " & _ 
			" trefPOS_Subjects ON tblClasses.intPOS_Subject_ID = trefPOS_Subjects.intPOS_Subject_ID " & _ 
			"WHERE (tblClasses.intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _
			" AND (tblClasses.intInstructor_ID IS NOT NULL) AND (NOT EXISTS " & _ 
			" (SELECT     'x' " & _ 
			"	FROM tascClass_Family " & _ 
			"	WHERE tblClasses.intClass_ID = tascClass_Family.intClass_ID " & strSqlFamily & "))  " & _ 
			strSqlNoDualEnroll & _
			"	AND tblClasses.intContract_Status_ID = 5 " 

	if Request.Form("keywords") <> "" then		
		if Request.Form("searchType") <> ""  then	
			strKeyWords = " like upper('%" & oFunc.EscapeTick(Request.Form("keywords"))& "%') " 
			sql = sql & " and (upper(convert(varChar(8000),substring(szCurriculum_Desc,1,8000)))" & strKeyWords & " or " & _
					"upper(convert(varChar(8000),substring(szGoals,1,8000)))" & strKeyWords & " or " & _
					"upper(convert(varChar(8000),substring(szRequirements,1,8000)))" & strKeyWords & " or " & _
					"upper(convert(varChar(8000),substring(szTeacher_Role,1,8000)))" & strKeyWords & " or " & _
					"upper(convert(varChar(8000),substring(szStudent_Role,1,8000)))" & strKeyWords & " or " & _
					"upper(szILP_NAME)" & strKeyWords & " or " & _
					"upper(szClass_Name)" & strKeyWords & " or " & _
					"upper(convert(varChar(8000),substring(szParent_Role,1,8000)))" & strKeyWords & ") " 
		else
			arWords = split(Request.Form("keywords")," ")
			if isArray(arWords) then
				sql = sql & " and ("
				for i = 0 to ubound(arWords)
					strKeyWords = " like upper('%" & oFunc.EscapeTick(arWords(i))& "%') "
					sql = sql & " upper(convert(varChar(8000),substring(szCurriculum_Desc,1,8000)))" & strKeyWords & " or " & _
					"upper(convert(varChar(8000),substring(szGoals,1,8000)))" & strKeyWords & " or " & _
					"upper(convert(varChar(8000),substring(szRequirements,1,8000)))" & strKeyWords & " or " & _
					"upper(convert(varChar(8000),substring(szTeacher_Role,1,8000)))" & strKeyWords & " or " & _
					"upper(convert(varChar(8000),substring(szStudent_Role,1,8000)))" & strKeyWords & " or " & _
					"upper(szILP_NAME)" & strKeyWords & " or " & _
					"upper(szClass_Name)" & strKeyWords & " or " & _
					"upper(convert(varChar(8000),substring(szParent_Role,1,8000)))" & strKeyWords & " or" 
				next	
				sql = left(sql,len(sql)-2) 	
				sql = sql & ") "	
			end if 
		end if 
	end if 		
	
	if request("szDays_Meet_On") <> "" then
		sql = sql & " AND upper(szDays_Meet_On) like '%" & ucase(request("szDays_Meet_On")) & "%' " 
	end if	
	
	if request("intInstructor_ID") <> "" then
		sql = sql & " AND tblClasses.intInstructor_ID = " & request("intInstructor_ID") & " " 
	end if	
	
	if request("intPOS_Subject_ID") <> "" then
		sql = sql & " AND tblClasses.intPOS_Subject_ID = " & request("intPOS_Subject_ID") & " " 
	end if	
	
	if request("hourStart") <> "" and request("minuteStart") <> "" then
		sql = sql & " AND tblClasses.szStart_Time = '" & request("hourStart") & ":" & request("minuteStart") & " " & request("amPmStart") & "' "
	end if	

	if request("orderby") <> "" then
		sql = sql & " ORDER BY " & request("orderby")
	else
		sql = sql & " ORDER BY szClass_Name" 
	end if
	
	set rs = server.CreateObject("ADODB.RECORDSET")
	rs.CursorLocation = 3
	rs.Open sql, Application("cnnFPCS")'oFunc.FPCScnn

if ucase(session.contents("strUserId")) = "SCOTT" then response.write sql
	
	if request("PageNumber") <> "" and request("hdnReset") = "" then
		intPageNum = cint(request("PageNumber"))	
	else
		intPageNum = 1
	end if
	
	with rs
		if .RecordCount > 0 then
			.PageSize = 25
			.AbsolutePage = intPageNum
			intViewingTo = .AbsolutePosition + .PageSize -1 
			if intViewingTo > .recordcount then intViewingTo = .RecordCount
%>
<br>
<input type="hidden" name="PageNumber" value="<% = intPageNum%>" ID="Hidden7">
<table cellpadding="2">
	<tr>
		<td colspan=10 class="svplain8" nowrap>
			
			Viewing <% = .AbsolutePosition %> - <% = intViewingTo %>  of <% = .RecordCount %> Matches &nbsp;
			
			<table ID="Table4" cellpadding="2"><tr><td>
			<%
				if cint(.RecordCount) > cint(.PageSize) then
					for i = 1 to .PageCount
					
						if intViewingTo/.PageSize = i or (.RecordCount = intViewingTo and i = .PageCount) then 
							strCssClass = "NavSave"
						else
							strCssClass = "btSmallWhite"
						end if
					%>
					<input type="button" class="<% = strCssClass %>" value="<%=i%>" onClick="this.form.PageNumber.value='<%=i%>';this.form.submit();" ID="Button2" NAME="Button2">
					<%
					next 
				end if
			%>
			</td></tr></table>
		</td>					
	</tr>
<%			
			intCount = 0
			intCount2 = 0 
			intMax = (.AbsolutePosition + .PageSize)
			
			do while .AbsolutePosition < intMax and not .EOF
				if intCount mod 2 = 0 then
					strColor = "TableCell"
				else
					strColor = "gray"
				end if
				
				if intCount2 = 0 or intCount2 mod .PageSize = 0 then
					call PrintHeader
				end if
				
%>	
			<tr id="ROW<%=intCount%>" onClick="jfHighLight('<%=intCount%>');" class="<% = strColor %>">				
				<% if request("bolWin") <> "" then %>
				<td class="TableHeader" align="center">
					<% if (cint(rs("intMax_Students") & "") - cint(rs("SlotsTaken"))) < 1 then%>
						<b>full</b>
					<% else %>
						<a href="javascript:"  style="color:white;" onclick="jfAddClass('<% = rs("intInstructor_ID")%>','<% = rs("intClass_ID") %>');">select</a>
					<% end if %>
				</td>
				<% end if %>
				<td >
					<% = rs("szClass_Name") %>
				</td>
				<td >
					<% = rs("szFIRST_NAME") & " " & rs("szLAST_NAME") %>
				</td>
				<td >
					<% = ucase(rs("szSubject_Name")) %>
				</td>
				<td align="center">
					<%	if (cint(rs("intMax_Students") & "") - cint(rs("SlotsTaken"))) < 1 then
							response.Write "<B>class full</b>"
						else						
							response.Write (cint(rs("intMax_Students") & "") - cint(rs("SlotsTaken"))) &  " of " & rs("intMax_Students")
						end if
					%>
				</td>
				<td align="center">
					<% = rs("sGrade_Level") %> - <% = rs("sGrade_Level2") %>
				</td>
				<td>
					<a href="#" onCLick="jfPrintAll('<% =rs("intClass_ID")%>','<% =rs("intILP_ID")%>');">Contract/Ilp</a>
					&nbsp;<a href="#" onCLick="jfViewBio('<% =rs("intInstructor_ID")%>');">Teacher Bio</a>
					
				</td>
			</tr>
<%				
				.MoveNext
				intCount = intCount + 1
				intCount2 = intCount2 + 1
			loop
%>
	<input type=hidden name="intCount" value="<%=intCount%>" ID="Hidden10">
	<input type=hidden name="intCount2" value="<%=intCount2%>" ID="Hidden12">
	<input type="hidden" name="orderby" value="<% = request("orderby") %>">
</table>
<%		
		else
			%>
			<span class="svplain8"><B>0 Matches Found.</B></span>
			<%
		end if
		.close		
	end with
end if

%>
</form>
<%
call oFunc.CloseCN()
set oFunc = nothing
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")


function PrintHeader
%>
	<Tr>
		<% if request("bolWin") <> "" then %>
		<td class="TableHeader" align="center">
			Add Class
		</td>
		<% end if %>
		<td class="TableHeader">
			<a href="#" class="linkWht" onclick="document.forms[0].orderby.value='szClass_Name';document.forms[0].submit();">Class Name</a>
		</td>
		<td class="TableHeader">
			<a href="#" class="linkWht" onclick="document.forms[0].orderby.value=' szLAST_NAME,szFIRST_NAME';document.forms[0].submit();">Instructor</a>
		</td>
		<td class="TableHeader">
			<a href="#" class="linkWht" onclick="document.forms[0].orderby.value=' tblClasses.intPOS_Subject_ID';document.forms[0].submit();">Subject</a>
		</td>
		<td class="TableHeader" align="center">
			Available<BR>Spots
		</td>
		<td class="TableHeader">
			Grade
		</td>
		<td class="TableHeader" align="center">
			Links
		</td>
	</Tr>
<%
end function
%>