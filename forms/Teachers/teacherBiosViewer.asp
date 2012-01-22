<%@ Language=VBScript %>
<%
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

'JD: Deny access if VENDOR
if ucase(session.Contents("strRole")) = "VENDOR" then
	'terminate page since page was improperly called.
	response.Write "<html><body><H1>Page improperly called.</h1></body></html>"
	response.End
end if
'JD


oFunc.ResetSelectSessionVariables()
' We dimension the following variables here to give them global scope and they are defined in vbfGetBio
dim szFirst_Name
dim szLast_Name
dim szEmail
dim szBio
dim szPhoto_Link
dim intInstructor_Bios_ID
dim strPhoto
dim intInstructor_ID

dim intCount
dim intStart
dim intEnd
dim strReturnData
dim strHTMLClasses		'Contains drop down with list of instructors classes if show classes is on

Session.Value("strTitle") = "Teacher Bios"
Session.Value("strLastUpdate") = "26 July 2002"

if request.QueryString("simpleHeader") <> "" then
	Server.Execute(Application.Value("strWebRoot") & "includes/simpleheader.asp")
else
	Server.Execute(Application.Value("strWebRoot") & "includes/header.asp")
end if

if Request.QueryString("intInstructor_ID") <> "" then
	' Show existing record
	call vbfGetBio(Request.QueryString("intInstructor_ID"))
elseif Request.Form("keywords") <> "" then	
	
	if Request.Form("searchType") = "exact" then
		strSQL = "upper(szBio) like upper('%" & oFunc.EscapeTick(Request.Form("keywords")) & "%') " 
	else
		arWords = split(Request.Form("keywords")," ")
		if isArray(arWords) then
			for i = 0 to ubound(arWords)
				strSQL = strSQL & " upper(szBio) like upper('%" & oFunc.EscapeTick(arWords(i)) & "%') or " 
			next	
			strSQL = left(strSQL,len(strSQL)-3)
		end if 
	end if 
	
	sql = "select i.intInstructor_ID, szFirst_Name, szLast_Name,szBio " & _
		  "from tblInstructor i, tblInstructor_Bios b " & _
		  "where i.intInstructor_ID = b.intInstructor_ID AND " & strSQL
	
	set rsSearch = server.CreateObject("ADODB.RECORDSET")
	rsSearch.CursorLocation = 3
	rsSearch.Open sql, oFunc.FPCScnn
	
	intCount = 0 
	if rsSearch.RecordCount > 0 then					
		do while not rsSearch.EOF				
			dim intPadding   ' this number tells us how many characters on each side of the key word 
							 ' we want to display back to the user in a Bio blurb
			
			intStrLocal = instr(1,rsSearch("szBio"),Request.Form("keywords"))
			intPadding = 50 
			if Request.Form("searchType") = "exact" then	
				if intStrLocal > intPadding then
					intStart = intStrLocal - intPadding
				else
					intStart = 1
				end if
				
				if (intStart + len(Request.Form("keywords")) + intPadding) <= len(rsSearch("szBio")) then
					intEnd = len(Request.Form("keywords")) + (intPadding*2)
				else
					intEnd = len(right(rsSearch("szBio"),len(rsSearch("szBio"))-intStart))
				end if 
				strBioPeice = mid(rsSearch("szBio"),intStart,intEnd)
			else
				strBioPeice = left(rsSearch("szBio"),80) 
			end if
			intCount = intCount + 1
			strResults = strResults & _
						 "<tr><td class=gray>&nbsp;<B>" & intCount &": </b>" & _
						 "<a href='teacherBiosViewer.asp?intInstructor_id=" & rsSearch("intInstructor_ID") & "'>" & _
						 "" & rsSearch("szFirst_Name") & " " & rsSearch("szLast_Name") & _
						 "</a></td><tr><TD class=svPlain10>&nbsp;" & strBioPeice & _
						 "</td></tr>"		
			rsSearch.MoveNext				 
		loop
		strReturnData = "<tr><td class=svHeader10><b>Results: </b>" & rsSearch.RecordCount & " Record(s) Found</td></tr>" & _
						strResults
	else
		strReturnData = "<tr><td class=svHeader10><b>Results: </b>0 Record(s) Found</td></tr>"
	end if 
	
	rsSearch.Close
	set rsSearch = nothing
	call vbfSearchForm
else
	' Stop script.  We must have the "intInstructor_ID" parameter provided by the user
	call vbfSearchForm
end if

function vbfGetBio(instructor_ID)
	' This function takes the parameter instructor_ID and pulls the bio info
	' for the given instructor
	
	dim sqlTeacher
	set rsInstInfo = server.CreateObject("ADODB.Recordset")
	rsInstInfo.CursorLocation = 3
	
	sqlTeacher = "select i.szFirst_Name, i.szLast_Name, i.szEmail, b.szBio, b.szPhoto_Link, " & _
				 "b.bolShow_Classes,b.szAdditional_Contact, b.intInstructor_Bios_ID " & _
				 "from tblInstructor i left outer join tblInstructor_Bios b " & _
				 "ON i.intInstructor_ID = b.intInstructor_ID " & _
				 "where i.intInstructor_ID = " & instructor_ID
	
	rsInstInfo.Open sqlTeacher, oFunc.FPCScnn
	
	'This for loop dimentions and defines all the columns we selected in sqlTeacher
	'and we use the variables created here to populate the form.
	for each item in rsInstInfo.Fields
		execute(item.Name & " = item")		
	next
	
    ' Check to see if we have the teachers photo 
    Set fso = CreateObject("Scripting.FileSystemObject")
    if not fso.FileExists(Server.MapPath(Application("strImageRoot") & "teachers/" & instructor_ID & ".jpg")) then
		strPhoto = "Photo not available"
	else
		strPhoto =  "<img src='" & Application("strImageRoot") & "teachers/" & instructor_ID & ".jpg'>" 
	end if 

	rsInstInfo.Close
	set rsInstInfo = nothing
	
	intInstructor_ID = instructor_ID
	
	if ucase(bolShow_Classes)= "TRUE" then
		sqlClasses  = "SELECT convert(varChar,c.intClass_ID) + '|' + convert(varChar,tblILP_Generic.intILP_ID) as IDs, c.szClass_Name " & _
						"FROM tblClasses c LEFT OUTER JOIN " & _
						" tblILP_Generic ON c.intClass_ID = tblILP_Generic.intClass_ID " & _
						"WHERE (NOT EXISTS " & _
						" (SELECT 'x' " & _
						" FROM tascClass_Family f " & _
						" WHERE f.intClass_ID = c.intClass_ID)) AND (c.intInstructor_ID = " & instructor_ID & ")  " & _
						" AND (c.intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _
						"	AND c.intContract_Status_ID = 5 " & _
						"ORDER BY c.szClass_Name"
						
		strClassList = oFunc.MakeListSql(sqlClasses,"IDs","szClass_Name","")
		              
		if oFunc.makeListRecordCount > 0 then
			strHTMLClasses = "<tr>" & chr(13) & _
							"	<td class=gray>" & chr(13) & _
							"			&nbsp;<b>Current Classes:</B>" & chr(13) & _
							"	</td>" & chr(13) & _
							"	<td class=gray>" & chr(13) & _
							"		<select name='intClass_ID' onChange=""jfGo(this.value);"">" & chr(13) & _
							"			<option> " & chr(13) & _
							strClassList & chr(13) & _
							"		</select>" & chr(13) & _
							"	</td>" & chr(13) & _
							"</tr>" 
		end if 
	end if
	
	call vbfBioTable
end function 
oFunc.CloseCN()
set oFunc = nothing

function vbfSearchForm
%>
<form action=teacherBiosViewer.asp method=post>
<table width=100%>
	<tr>	
		<Td class=yellowHeader>
				&nbsp;<b>Teacher Bio Search Engine</b>
		</td>
	</tr>
	<tr>
		<td bgcolor=f7f7f7>
			<table>
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
							&nbsp;<b>Select Teacher:</B>
					</td>
					<td class=gray>
						<select name=intInstructor_ID onChange="window.location.href='teacherBiosViewer.asp?intInstructor_id=' + this.value;">
							<option>
							<%
								set oList = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/dbOptionsList.wsc"))
								response.Write oList.ActiveTeachers(session.Contents("intSchool_Year"),"")
								set oList = nothing
							%>
						</select>
					</td>
				</tr>
				<tr>
					<td class=gray colspan=2>
							&nbsp;<b>OR</B>
					</td>
				</tr>
				<tr>
					<td class=gray>
							&nbsp;<b>Search by key word(s):</B>
					</td>
					<td class=gray>
							<input type=text name=keywords size=40 maxlength=50 value="<% = Request.Form("keywords") %>">
					</td>
				</tr>
				<tr>
					<td class=gray colspan=2>
						&nbsp;Match All Words<input type=radio value="exact" name="searchType" checked>
						&nbsp;&nbsp;Match Any Word<input type=radio value="any" name="searchType">
					</td>
				</tr>
				<tr>
					<td class=gray colspan=2>
						<input type=button value="Home" onClick="window.location.href='<%=Application.Value("strWebRoot")%>';" class="NavLink" >
						<input type=submit value="Submit" class="NavLink">
					</td>
				</tr>
			</table>
			<BR><BR>
			<table>
				<% = strReturnData%>
			</table>						
		</td>
	</tr>
</table>
</form>
<%
end function

function vbfBioTable
%>
<script language=javascript>
	function jfGo(ids) {
		var classWin;
		var arIDS = ids.split("|");
		var class_ID = arIDS[0];
		var ilp_ID = arIDS[1];
		var strURL = "<%=Application.Value("strWebRoot")%>/forms/printableForms/allPrintable.asp?";
		strURL += "noprint=true&intClass_ID="+class_ID+"&intILP_ID="+ilp_ID;	
		classWin = window.open(strURL,"classWin","width=640,height=500,scrollbars=yes,resizable=yes");
		classWin.moveTo(0,0);
		classWin.focus();
	}
</script>
<table width=100%>
	<tr>	
		<Td class=yellowHeader>
				&nbsp;<b>Teacher Bio</b>
		</td>
	</tr>
	<tr>
		<td bgcolor=f7f7f7>
			<table>
				<tr>
					<td>
						<table>
							<tr>
								<td class=gray valign=top>
									 <% = strPhoto %>
								</td>
								<td valign=top>
									<table>   
										<tr>
											<td class=gray>
													&nbsp;<b>Teacher's Name:</B>
											</td>
											<td class=gray>
													&nbsp;<% = szFirst_Name & " " & szLast_Name %>
											</td>
										</tr>
										<tr>
											<td class=gray>
													&nbsp;<b>Teacher's Email:</B>
											</td>
											<td class=gray>
													&nbsp;<a href="mailto:<% = szEmail %>"><% = szEmail %></a>
											</td>
										</tr>
										<tr>
											<td class=gray>
													&nbsp;<b>Additional Contact Info:</B>
											</td>
											<td class=gray>
													<% = szAdditional_Contact%>
											</td>
										</tr>
										<% = strHTMLClasses %>
									</table>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td class=gray colspan=2>
							&nbsp;<b>Biographical Information:</B>
					</td>
				</tr>
				<tr>
					<td colspan=2>
						<%=replace(replace(szBio & "",chr(13),"<BR><BR>")," ", "&nbsp; ")%>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<%	if request.QueryString("simpleHeader") = "" then %>
<input type=button value="Home" onClick="window.location.href='<%=Application.Value("strWebRoot")%>';" class="btSmallGray" >
<input type=button value="Back to List" onClick="history.go(-1);" class="btSmallGray" NAME="Button1">
<input type=button value="New Search" onClick="window.location.href='<%=Application.Value("strWebRoot")%>/forms/teachers/teacherBiosViewer.asp';" class="btSmallGray">
<%	else %>
<input type="button" value="close window" class="btSmallGray" onclick="window.close();">
<%
	end if
end function
%>
</BODY>
</HTML>
