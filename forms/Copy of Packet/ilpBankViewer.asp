<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		ilpBankViewer.asp
'Purpose:	This script is the ilp search engine.  With the results that it
'			returns it creates links to view the ilp (via genericILPViewer.asp)
'			or sends the ilp data to dynamically fill in the ilp form in ilpMain.asp
'Date:		10-01-2002
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

' We dimension the following variables here to give them global scope 
dim szFirst_Name
dim szLast_Name
dim szEmail
dim szBio
dim szPhoto_Link
dim intInstructor_Bios_ID
dim strPhoto
dim intInstructor_ID
dim szUser_ID

dim intCount
dim intStart
dim intEnd
dim strReturnData
dim strHTMLClasses		'Contains drop down with list of instructors classes if show classes is on
dim strPOS
dim strtLink
dim strUpdate

Session.Value("strTitle") = "ILP Bank"
Session.Value("strLastUpdate") = "08 Sept 2002"


if request("fromMain") <> "" then
	Server.Execute(Application.Value("strWebRoot") & "includes/simpleheader.asp")
else
	oFunc.ResetSelectSessionVariables
	session.Contents("intStudent_ID") = ""
	Server.Execute(Application.Value("strWebRoot") & "includes/header.asp")
end if 

'  UPDATE ILP BANK STATUS
if len(request("IList")) > 1 then
	arList = split(request("IList"),",")
	for i = 0 to ubound(arList)
		if arList(i) <> "" then
			if request.Form("pp" & arList(i)) & "" = "1" then
				sPubVal = "1"
			else
				sPubVal = "0"
			end if
			
			strUpdate = "Update tblILP set isPublic = " & sPubVal  & _
				", dtModify = CURRENT_TIMESTAMP, szUSER_MODIFY = '" & session.Contents("strUserID") & "' " & _
				"Where intILP_ID = " & arList(i)
				'response.Write strUpdate & "<BR>"
			oFunc.ExecuteCN(strUpdate)
		end if
	next
	strUpdate = ""
end if

if len(request("CList")) > 1 then
	arList = split(request("CList"),",")
	for i = 0 to ubound(arList)
		if arList(i) <> "" then
			if request.Form("pp" & arList(i)) & "" = "1" then
				sPubVal = "1"
			else
				sPubVal = "0"
			end if
			
			strUpdate = "Update tblILP_GENERIC set isPublic = " & sPubVal  & _
				", dtModify = CURRENT_TIMESTAMP, szUSER_MODIFY = '" & session.Contents("strUserID") & "' " & _
				"Where intILP_ID = " & arList(i)
				
			oFunc.ExecuteCN(strUpdate)
		end if
	next
	strUpdate = ""
end if

if len(request("CCheck")) > 1 then
	cCheck = right(request("CCheck"),len(request("CCheck"))-1)
	cCheck = left(cCheck,len(cCheck)-1)
	if len(cCheck) > 1 then 
		update = "update tblILP_Generic set bolILP_Bank = 0 " & _
				", dtModify = CURRENT_TIMESTAMP, szUSER_MODIFY = '" & session.Contents("strUserID") & "' " & _
				" WHERE intILP_ID in (" & cCheck & ") "
		oFunc.ExecuteCN(update)
	end if
end if

if len(request("ICheck")) > 1 then
	iCheck = right(request("ICheck"),len(request("ICheck"))-1)
	iCheck = left(iCheck,len(iCheck)-1)
	if len(iCheck) > 1 then 
		update = "update tblILP set bolILP_Bank = 0 " & _
				", dtModify = CURRENT_TIMESTAMP, szUSER_MODIFY = '" & session.Contents("strUserID") & "' " & _
				" WHERE intILP_ID in (" & iCheck & ") "
				'response.Write update
		oFunc.ExecuteCN(update)
	end if
end if

' USER FILTER CRITERIA 
if Request.Form("szUser_ID") <> "" then
	' Teacher filter
	if ucase(Request.Form("szUser_ID")) = ucase(session.Contents("strUserID")) or oFunc.IsAdmin then
		' Teacher has rights to see all thier ILP's in the bank.  So does the admin
		strSQL = " and upper(i.szUser_Create) = upper('" & Request.Form("szUser_ID") & "') "		
	else
		' Teachers and guards can only see public ILP's for other teachers
		if request.Form("sAccess") & "" <> " and IsPublic = 0 " then 
			strSQL = " and (upper(i.szUser_Create) = upper('" & Request.Form("szUser_ID") & "')  AND isPublic = 1) "
		else
			' show no data because private is turned on 
			strSQL = " and (upper(i.szUser_Create) = 'NO_USER') "
		end if
	end if
elseif request("intGuardian_ID") <> "" then
	' This is only applicable if the user is a guardian
	' Guardian can see all of thier ILP's in the bank
	strSQL = strSQL & " AND (upper(i.szUser_Create) = upper('" & Session.Contents("strUserID") & "')) " 		
elseif Request.Form("intInstructor_ID") & "" = "" and _
   request("intGuardian_ID") & "" = "" and not oFunc.IsAdmin then
   ' Teachers and Guards can only see public ILP's with the exclusion
   ' of the cases presented prior in the script
	strSQL = strSQL & " AND ((upper(i.szUser_Create) = upper('" & Session.Contents("strUserID") & "') " & request.Form("sAccess") & ") "
	if request.Form("sAccess") & "" <> " and IsPublic = 0 " then 
	  strSQL = strSQL & " or isPublic = 1 "
	end if
	
	strSQL = strSQL & ") "
	
	if oFunc.IsTeacher then
		strSQL = strSQL & " or (upper(i.szUser_Create) in (SELECT UPPER(gu.szUser_ID) " & _ 
				"FROM	tblINSTRUCTOR i INNER JOIN " & _ 
				"	tblGUARDIAN g ON i.szFIRST_NAME = g.szFIRST_NAME AND i.szLAST_NAME = g.szLAST_NAME INNER JOIN " & _ 
				"	tascGUARD_USERS gu ON g.intGUARDIAN_ID = gu.intGUARDIAN_ID " & _ 
				"WHERE	(i.intINSTRUCTOR_ID = " & session.Contents("instruct_ID") & ")))  "
	end if
end if

' KEY WORD FILTER CRITERIA
if Request.Form("keywords") <> "" then		
	if Request.Form("searchType") = "exact" then	
		strKeyWords = " like upper('%" & oFunc.EscapeTick(Request.Form("keywords"))& "%') " 
		strSQL = strSQL & " and (upper(convert(varChar(8000),substring(szCurriculum_Desc,1,8000)))" & strKeyWords & " or " & _
				 "upper(convert(varChar(8000),substring(szGoals,1,8000)))" & strKeyWords & " or " & _
				 "upper(convert(varChar(8000),substring(szRequirements,1,8000)))" & strKeyWords & " or " & _
				 "upper(convert(varChar(8000),substring(szTeacher_Role,1,8000)))" & strKeyWords & " or " & _
				 "upper(convert(varChar(8000),substring(szStudent_Role,1,8000)))" & strKeyWords & " or " & _
				 "upper(szILP_NAME)" & strKeyWords & " or " & _
				 "upper(convert(varChar(8000),substring(szParent_Role,1,8000)))" & strKeyWords & ") " 
	else
		arWords = split(Request.Form("keywords")," ")
		if isArray(arWords) then
			strSQL = strSQL & " and ("
			for i = 0 to ubound(arWords)
				strKeyWords = " like upper('%" & oFunc.EscapeTick(arWords(i))& "%') "
				strSQL = strSQL & " upper(convert(varChar(8000),substring(szCurriculum_Desc,1,8000)))" & strKeyWords & " or " & _
				 "upper(convert(varChar(8000),substring(szGoals,1,8000)))" & strKeyWords & " or " & _
				 "upper(convert(varChar(8000),substring(szRequirements,1,8000)))" & strKeyWords & " or " & _
				 "upper(convert(varChar(8000),substring(szTeacher_Role,1,8000)))" & strKeyWords & " or " & _
				 "upper(convert(varChar(8000),substring(szStudent_Role,1,8000)))" & strKeyWords & " or " & _
				 "upper(szILP_NAME)" & strKeyWords & " or " & _
				 "upper(convert(varChar(8000),substring(szParent_Role,1,8000)))" & strKeyWords & " or" 
			next	
			strSQL = left(strSQL,len(strSQL)-2) 	
			strSQL = strSQL & ") "	
		end if 
	end if 
end if 

if request("sAccess") <> "" then
	strSQL = strSQL &  request("sAccess")
end if 

' SUBJECT FILTER CRITERIA
if request.Form("intPOS_Subject_ID") <> "" then
	strPOS = " AND (intPOS_Subject_ID = " & request.Form("intPOS_Subject_ID")	& ") "
end if

' Allows Sponsor Teachers to see all ILP's in the Bank for their Sponsored Guardians
if request.Form("showSponsoredILP") <> "" and oFunc.IsTeacher then
	if len(session.Contents("student_list")) > 0 then
		strStudentList = replace(session.Contents("student_list"),"~~",",")
		strStudentList = replace(strStudentList,"~","")
		strSQL = strSQL & "	OR (UPPER(i.szUSER_CREATE) IN " & _ 
						"	((SELECT DISTINCT UPPER(gu.szUser_ID) " & _ 
						"	FROM	tascGUARD_USERS gu INNER JOIN " & _ 
						"	tascStudent_Guardian sg ON sg.intGuardian_ID = gu.intGuardian_ID " & _ 
						"	WHERE	sg.intStudent_ID IN (" & strStudentList & "))) " & request.Form("sAccess") & " ) " 
	end if
end if 

if request.Form("isSubmitted") <> "" and (strSQL <> "" or strPOS <> "" or (oFunc.IsAdmin and request("IList") <> "")) then
	dim strColumn 
	strColumn = ""
	'if ucase(session.Contents("strRole")) = "ADMIN" then
		strColumn = "<TD class=gray>Public<input type=submit value=""Remove"" class=""btsmallgray""></td>"
	'end if
	
	strResults = "<table cellpadding='2'>" 
	
	if  request("fromMain") = "" then
		strResults2 = strResults2 & "   <tr>" & _
				"		<td colspan='15' align='right'><input type='submit' value='Save Changes' class='NavSave'></tr>" 
	end if 
	
	strResults2 = strResults2 & "	<tr>" & _	
				"		<td>" & _			
				"		</td>" & _
				"		<td class='TableHeader'>" & _
				"			<b>ILP Name</b> (click to view)" & _
				"		</td>" & _
				"		<td  class='TableHeader'>" & _
				"			<b>Subject</b>" & _
				"		</td>" & _
				"		<td  class='TableHeader' align='center'>" & _
				"			<b>Created By</b>" 
				
	if request("fromMain") = "" then 
			strResults2 = strResults2 & "		</td>" & _
							"		<td  class='TableHeader' align='center'>" & _
							"			<b>Access</b>" & _
							"		</td>" & _
							"		<td  class='TableHeader' align='center'>" & _
							"			<b>Delete From Bank</b>" & _
							"		</td>" 
	end if
	strResults2 = strResults2 &	"	</tr>"	 
				
	
	strResults = strResults & strResults2
	
	sql = "SELECT i.intILP_ID, i.szILP_Name, i.szSubject_Name, i.ILP_Type,i.intGuardian_ID, " & _
		  " i.intInstructor_ID, i.isPublic, u.szName_First, u.szName_Last, i.szUser_Create, i.intStudent_ID  " & _
		  "FROM v_ILP_Bank i left outer join " & _
		  " tblUsers u on i.szUser_Create = u.szUser_ID " & strFrom & _
		  "WHERE 1=1 " & strPOS & _
          strSQL & " order by i.szILP_Name, i.szSubject_Name"  
          
          '"WHERE bolILP_Bank = 1 and (intPOS_Subject_ID <> 22 " & strPOS & ") " & _
	set rsSearch = server.CreateObject("ADODB.RECORDSET")
	rsSearch.CursorLocation = 3
	if ucase(session.Contents("strUserId")) = "CHRONIH30" then
		'response.Write "<h1>testing</h1><BR>" & sql
	end if
	'response.End
	rsSearch.Open sql, oFunc.FPCScnn
	
	intCount = 0 
	dim strCheckBox
	
	if rsSearch.RecordCount > 0 then					
		do while not rsSearch.EOF							
			intCount = intCount + 1
			
			if intCount mod 25 = 0 then
				strResults = strResults & strResults2
			end if
			
			if intCount mod 2 = 0 then
				strColor = "TableCell"
			else
				strColor = "gray"
			end if			
				
			if request("fromMain") <> "" then
				strLink = "<a href='javascript:' " & _
						  "onClick=""jfGetGenericILP('" & rsSearch("intILP_ID") & "','" & rsSearch("ILP_Type") & "');"">" & _
						   "Select This ILP:</a>"
			else 
				strLink = intCount
			end if
			
			if rsSearch("isPublic") then
				sPublic = " selected "
				sPText = "public"
			else
				sPublic =  " "
				sPText = "private"
			end if
			
			strCheck = ""
			if (ucase(session.Contents("strUserID")) = ucase(rsSearch("szUser_Create"))) or _
				 (oFunc.IsAdmin) or (trim(rsSearch("szName_First")) & " " & trim(rsSearch("szName_Last")) = session.Contents("strFullName")) _
				 or (instr(1,session.Contents("student_list"), "~" & rsSearch("intStudent_ID") & "~")) then
				'strCheckBox = "<td align=center>" & _
				'				"<input type=checkbox name=""" & lcase(rsSearch("ILP_Type")) & rsSearch("intILP_ID") & """ value='1' " & sPublic & " ></td>"
				'if lcase(rsSearch("ILP_Type")) = "i" then
				'	strGList = strGList & rsSearch("intILP_ID") & "|"
				'elseif lcase(rsSearch("ILP_Type")) = "c" then
				'	strIList = strIList & rsSearch("intILP_ID") & "|"
				'end if
				
				if lcase(rsSearch("ILP_Type")) = "i" then
					selectObj = "IList"
					checkObj = "ICheck"
				else
					selectObj = "CList"
					checkObj = "CCheck"
				end if 
				
				strList = "<select id='pp" & rsSearch("intILP_ID") & "'  name='pp" & rsSearch("intILP_ID") & "' onChange=""jfUpdateList('" & rsSearch("intILP_ID") & "','" & selectObj & "');""> " & _
						  "<option value='0' style='color:red;'>Private</option>" & _
						  "<option value='1'  style='color:green;' " & sPublic & ">Public</option>" & _
						  "</select>"
				
				strCheck = "<input type='checkbox' name='db" &  rsSearch("intILP_ID") & "'  name='db" & rsSearch("intILP_ID") & "' value='1' onChange=""jfUpdateList('" & rsSearch("intILP_ID") & "','" & checkObj & "');""> Delete"
				
			else
				strList = sPText 				
			end if
			
			strResults = strResults & _
						 "<tr id=""ROW"  & intCount & """ onClick=""jfHighLight('" & intCount & "');"" class=""" & strColor & """><td nowrap>&nbsp;<B>" & strLink & "</b></td>" & _
						 "<td> " & _
						 "<a href='javascript:' onClick=""jfViewGenericILP('" & _
						 rsSearch("intILP_ID") & "','" & rsSearch("ILP_Type") & "');"">" & rsSearch("szILP_Name")  & _
						 "</a></td><TD>" & rsSearch("szSubject_Name") & _
						 "</td><td nowrap>" & mid(rsSearch("szName_First"),1,1) & ". " & rsSearch("szName_Last") & " " & "</TD>" 
						 
			if request("fromMain") = "" then 
				strResults = strResults & "<td class='svplain8' align='center'>" & strList & "</td><td align='center' class='svplain8'>" & strCheck & "</td>"		
			end if
			
			strResults = strResults & "</tr>"
			rsSearch.MoveNext				 
		loop
		strReturnData = "<tr><td class=svHeader10 colspan=4><b>Results: </b>" & rsSearch.RecordCount & " Record(s) Found<BR><BR></td></tr>" & _
						strResults & "</table>"
	else
		strReturnData = "<tr><td class=svHeader10 ><b>Results: </b>0 Record(s) Found</td></tr>"
	end if 
	
	rsSearch.Close
	set rsSearch = nothing
	call vbfSearchForm
else
	' Stop script.  We must have the "intInstructor_ID" parameter provided by the user
	call vbfSearchForm
end if

oFunc.CloseCN()
set oFunc = nothing

function vbfSearchForm
%>
<script language="javascript">
	function jfUpdateList(pVal,pObjName) {
		var obj;
		obj = document.getElementById(pObjName);
		
		if (obj.value.indexOf(","+pVal+",") == -1 ) {
			obj.value = obj.value + pVal + ",";
		}
	}	
</script>
<form action=ilpBankViewer.asp method=post name="main" ID="Form1">
<input type=hidden name="lastRow" ID="Hidden2">
<input type=hidden name="LineItemsChanged" value="," ID="Hidden8">
<input type=hidden name="lastRowColor" ID="Hidden3">
<input type=hidden name="fromMain" value="<% = request("fromMain") %>" ID="Hidden1">
<input type=hidden name="bolLateAdd" value="<% = request("bolLateAdd") %>" ID="Hidden4">
<input type=hidden name="isPopUp" value="<%= request("isPopUp") %>" ID="Hidden5">
<input type=hidden name="isSubmitted" value="true" ID="Hidden6">
<input type=hidden name="intExisitingGenericILP" value="<% = request("intExisitingGenericILP") %>" ID="Hidden7">
<input type=hidden name="strSession" value="<% = request("strSession")%>" ID="Hidden9">
<input type="hidden" name="CList" id="CList" value=",">
<input type="hidden" name="IList" id="IList" value=",">
<input type="hidden" name="ICheck" id="ICheck" value=",">
<input type="hidden" name="CCheck" id="CCheck" value=",">

<table width=100% ID="Table1">
	<tr>	
		<Td class=yellowHeader>
				&nbsp;<b>ILP Bank Search Engine</b>
		</td>
	</tr>
	<tr>
		<td bgcolor=f7f7f7>
			<table ID="Table2">
				<tr>
					<td>
					
						<table ID="Table3">
							<tr>	
								<Td colspan=2>
									<font class=svplain11>
										<b><i>Search Criteria</I></B> 
									</font>
									<font class=svplain>
									</font>
								</td>
							</tr>
							<tr>
								<td class=gray>
										&nbsp;<b>Search by Teacher:</B>
								</td>
								<td>
									<select name="szUser_ID" style="width:100%;" ID="Select1">
										<option>
										<%
											sql = "Select iu.szUser_ID, i.szLast_Name + ', ' + i.szFirst_Name as Name " & _
												"from tblInstructor i inner join tascInstr_USER iu on i.intInstructor_ID = iu.intInstructor_ID order by i.szLast_Name "
											Response.Write oFunc.MakeListSQL(sql,"szUser_ID","Name",request.Form("szUser_ID"))
										%>
									</select>
								</td>
							</tr>				
							<tr>
								<td class=gray>
										&nbsp;<b>Search by Subject:</B>
								</td>
								<td>
										<select name="intPOS_Subject_ID"  style="width:100%;" ID="Select2">
											<option value="">ALL
											<%
												sql = "select intPOS_Subject_ID, upper(szSubject_Name) szSubject_Name from trefPOS_Subjects where bolShow = 1 order by szSubject_Name"									
												response.Write oFunc.MakeListSQL(sql,"intPOS_Subject_ID","szSubject_Name",request("intPOS_Subject_ID"))
											%>
										</select>
								</td>
							</tr>
							<tr>
								<td class=gray>
										&nbsp;<b>Search by Access:</B>
								</td>
								<td>
									<select name="sAccess" style="width:100%;" ID="Select3">
										<option value="">ALL</option>
										<option value=" and IsPublic = 1 " <% if request("sAccess") = " and IsPublic = 1 " then response.Write " selected " %>>PUBLIC ONLY</option>
										<option value=" and IsPublic = 0 " <% if request("sAccess") = " and IsPublic = 0 " then response.Write " selected " %>>PRIVATE ONLY</option>
									</select>
								</td>
							</tr>	
							<tr>
								<td class=gray>
										&nbsp;<b>Search by key word(s):</B>
								</td>
								<td>
										<input type=text name=keywords size=40 maxlength=50 value="<% = Request.Form("keywords") %>" ID="Text1">
								</td>
							</tr>
							<% if session.Contents("strRole") = "GUARD" then %>
							<tr>
								<td class=gray colspan=2>
										&nbsp;<b>Show only ILP's deposited by <% = Session.Contents("strFullName")%></B>
										<input type=checkbox name="intGuardian_ID" value="<%=Session.Contents("intGuardian_ID")%>" <% if request("intGuardian_ID") <> "" then response.Write " checked " %> ID="Checkbox1">
								</td>
							</tr>
							<% end if %>	
							<% if oFunc.IsTeacher then %>
							<tr>
								<td class=gray colspan=2>
										&nbsp;<b>Include all saved ILP's for Sponsored Students</B>
										<input type=checkbox name="showSponsoredILP" value="true" <% if request("showSponsoredILP") <> "" then response.Write " checked " %> ID="Checkbox2">
								</td>
							</tr>
							<% end if %>				
							<tr>
								<td class="TableCell" colspan=2>
									&nbsp;Match All Words<input type=radio value="exact" name="searchType" checked ID="Radio1">
									&nbsp;&nbsp;Match Any Word<input type=radio value="any" name="searchType" ID="Radio2">
								</td>
							</tr>				
							<tr>
								<td class=svplain10 colspan=2>
									<% if request("fromMain") = "" then %>
									<input type=button value="Home" onClick="window.location.href='<%=Application.Value("strWebRoot")%>';" class="btSmallGray" NAME="btSmallGray" ID="Button1">
									<!--<input type=button value="Deposit an ILP" onClick="jfDepILP();" class="btSmallGray">-->
									<% else %>
									<input type=button value="Close" onClick="window.opener.focus();window.close();" class="btSmallGray" ID="Button2" NAME="Button2">
									<% end if%>
									<input type=submit value="Submit" class="btSmallGray" NAME="Submit1" ID="Submit1">
								</td>
							</tr>
						</table>
					</td>
					<td valign="top">
						<table ID="Table4">
							<tr>	
								<Td >
									<font class=svplain11>
										<b><i>Search Instructions</I></B> 
									</font>
									<font class=svplain>
									</font>
								</td>
							</tr>
							<tr>
								<td style="width:100%;" class="svplain">
								
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>						
			<BR>
			<table ID="Table5" cellpadding=2>
				<% = strReturnData%>
			</table>						
		</td>
	</tr>	
</table>
</form>
<script language=javascript>
	function jfGetGenericILP(id,type){
			// This function is used when a user is in the process of adding a class
			// and has selected an existing ILP.  This code will load the selected
			// ilp into ilpMain.asp and will allow the user to procede to the next
			// step in the adding a class work flow.
			var ILP_ID_TYPE;
			if (type == "I"){
				ILP_ID_TYPE = "intILP_ID";
			}else{
				ILP_ID_TYPE = "intILP_ID_Generic";
			}
			var url = "<%=Application("strSSLWebRoot")%>forms/ILP/ILPMain.asp?<%= session.contents("strParams") %>";
			url += "&isPopUp=<%=request("plain")%>&" + ILP_ID_TYPE +"=" + id;
			url += "&bolHideAddBank=true&bolLateAdd=<% =request("bolLateAdd") %>";
			url += "&intExisitingGenericILP=<%=request("intExisitingGenericILP") %>&strSession=<% = request("strSession")%>";
			window.opener.location.href = url;
			window.opener.focus();
			window.close();
		}
	function jfViewGenericILP(id,type){
		// Opens new window to view generic ilp based on 'id'
		// which is the intILP_ID number
		var url = "<%=Application("strSSLWebRoot")%>forms/PrintableForms/allPrintable.asp?";
		url += "noprint=true&strAction=I&ILP_TYPE=" + type + "&intILP_ID=" + id;
		var winViewILP;
		winViewILP = window.open(url,"winViewILP","width=640,height=480,scrollbars=yes,resizable=yes");
		winViewILP.moveTo(0,0);
		winViewILP.focus();
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
	
</script>
<%
end function
%>
</BODY>
</HTML>
