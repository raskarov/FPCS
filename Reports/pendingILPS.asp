<%@ Language=VBScript %>
<%
' TOGGLES SHOWING GOODS/SERVICES 
if session.Contents("strRole") <> "ADMIN" then
	response.write "<h1>Page Improperly Called.</h1>"
	response.end
end if

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Dimention Variables, make db Connection, print HTML header.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim intStudent_ID 

Session.Contents("strTitle") = "Pending ILP's"
Session.Contents("strLastUpdate") = "22 Feb 2002"
if request("simpleHeader") = "" then
	Server.Execute(Application.Value("strWebRoot") & "includes/header.asp")
else
	Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")
end if


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

%>
<script language=javascript>
	function jfPacket(id){
		var winILPPend;
		var strURL = "<%=Application.Value("strWebRoot")%>forms/Packet/packet.asp?simpleHeader=true&intStudent_ID="+id;
		winILPPend = window.open(strURL,"winILPPend","width=800,height=550,scrollbars=yes,resize=yes,resizable=yes");
		winILPPend.moveTo(0,0);
		winILPPend.focus();	
	}
</script>
<table width=100% ID="Table1">
	<tr>	
		<Td class=yellowHeader >
				&nbsp;<b>Pending ILP's List for Active Students</b>
		</td>
	</tr>
	<tr>
		<td class="svplain8">
		<%
		set oHtml = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/htmlFuncs.wsc"))
						
						strDefinitions = "<table> " & _
								"	<tr> " & _
								"		<td class='TableCell' valign='top'> " & _
								"			<i>Working</i> " & _
								"		</td> " & _
								"		<td class='TableCell'> " & _
								"			This column reports the number of ILP's that are at the 'implemented' or 'must amend' stage. " & _
								"		</td> " & _
								"	</tr> " & _
								"	<tr> " & _
								"		<td class='TableCell' valign='top'> " & _
								"			<i>Ready for Sponsor</i> " & _
								"		</td> " & _
								"		<td class='TableCell'> " & _
								"			This column reports the number of ILP's that are ready for the sponsor teachers review. " & _
								"		</td> " & _
								"	</tr> " & _
								"	<tr> " & _
								"		<td class='TableCell' valign='top'> " & _
								"			<i>Ready for Admin</i> " & _
								"		</td> " & _
								"		<td class='TableCell'> " & _
								"			This column reports the number of ILP's that are ready for the administrators review. " & _
								"		</td> " & _
								"	</tr> " & _
								"</table>	 " 					
								response.Write oHtml.ToolTip("<a href='#'>Click here for column definitions</a>&nbsp;",strDefinitions,true,"Column Definitions",false,"ToolTip","400px","",true,true) 
								
								response.Write oHtml.ToolTipDivs
								set oHtml = nothing		
		%>
		</td>
	</tr>
	<tr>
		<td bgcolor=f7f7f7>
			<table ID="Table2">
				<tr>	
					<Td class="TableHeader" valign=middle align=center>
						&nbsp;<B>Student Name</b> (click to view packet)&nbsp;
					</td>
					<Td class="TableHeader" valign=middle align=center>
						&nbsp;<B>Sponsor Name</b> (click to email)&nbsp;
					</td>
					<Td class="TableHeader" valign=middle align=center>
						&nbsp;<b>Working&nbsp;
					</td>	
					<Td class="TableHeader" valign=middle align=center>
						&nbsp;<b>Ready for Sponsor&nbsp;
					</td>	
					<Td class="TableHeader" valign=middle align=center>
						&nbsp;<b>Ready for Admin&nbsp;
					</td>										
				</tr>
<%	
	'This section gives the classes for a student
set rsStudent = server.CreateObject("ADODB.RECORDSET")
rsStudent.CursorLocation = 3
sqlStudent = "SELECT     s.szLAST_NAME + ', ' + s.szFIRST_NAME AS Name, f.szHome_Phone, f.szEMAIL, s.intSTUDENT_ID, ss.intReEnroll_State, " & _ 
"                          (SELECT     COUNT(*) " & _ 
"                            FROM          tblILP i " & _ 
"                            WHERE      i.intStudent_ID = s.intStudent_ID AND (bolApproved = 0 OR " & _ 
"                                                   bolApproved IS NULL) AND (bolSponsor_Approved = 0 OR " & _ 
"                                                   bolSponsor_Approved IS NULL) AND i.sintSchool_Year = " & session.Contents("intSchool_Year") & ") AS IMPLEMENTED, " & _ 
"                          (SELECT     COUNT(*) " & _ 
"                            FROM          tblILP i " & _ 
"                            WHERE      i.intStudent_ID = s.intStudent_ID AND (bolSponsor_Approved = 0 OR " & _ 
"                                                   bolSponsor_Approved IS NULL) AND (bolReady_For_Review = 1) AND (bolApproved = 0 OR " & _ 
"                                                   bolApproved IS NULL) AND i.sintSchool_Year = " & session.Contents("intSchool_Year") & ") AS NA_SPONSOR, ss.szGRADE, " & _ 
"                          (SELECT     COUNT(*) " & _ 
"                            FROM          tblILP i " & _ 
"                            WHERE      i.intStudent_ID = s.intStudent_ID AND (bolApproved = 0 OR " & _ 
"                                                   bolApproved IS NULL) AND (bolSponsor_Approved = 1) AND i.sintSchool_Year = " & session.Contents("intSchool_Year") & ") AS NA_ADMIN,  " & _ 
"                      tblINSTRUCTOR.szFIRST_NAME + ' ' + tblINSTRUCTOR.szLAST_NAME AS teacher_name, tblINSTRUCTOR.szEmail AS teacher_email " & _ 
"FROM         tblENROLL_INFO e INNER JOIN " & _ 
"                      tblSTUDENT s ON e.intSTUDENT_ID = s.intSTUDENT_ID INNER JOIN " & _ 
"                      tblFAMILY f ON s.intFamily_ID = f.intFamily_ID INNER JOIN " & _ 
"                      tblStudent_States ss ON ss.intStudent_id = s.intSTUDENT_ID INNER JOIN " & _ 
"                      tblINSTRUCTOR ON e.intSponsor_Teacher_ID = tblINSTRUCTOR.intINSTRUCTOR_ID AND  " & _ 
"                      e.intSponsor_Teacher_ID = tblINSTRUCTOR.intINSTRUCTOR_ID " & _ 
"WHERE     (e.sintSCHOOL_YEAR = " & session.contents("intSchool_year") & " ) AND (ss.intSchool_Year = " & session.contents("intSchool_year") & " ) AND ss.intReEnroll_State  IN (" & Application.Contents("ActiveEnrollList") & ") " & _ 
"ORDER BY s.szLAST_NAME "	

rsStudent.Open sqlStudent,oFunc.FPCScnn	

		intColorCount = 0
		if rsStudent.RecordCount > 0 then
				do while not rsStudent.EOF						
					if intColorCount mod 2 = 0 then
						strBgColor = " bgcolor=white " 
					else
						strBgColor = ""
					end if 
					
%>
				<tr <% = strBgColor %>>
					<Td class="TableCell" title="<%=strConSched%>">
						&nbsp;<a href="javascript:" onclick="jfPacket('<% = rsStudent("intStudent_ID") %>');"><% = rsStudent("Name")%></a>&nbsp;
					</td>
					<Td class="TableCell">
						&nbsp;<a href="mailto:<% = rsStudent("teacher_email") %>"><% = rsStudent("teacher_name")%></a>&nbsp;
					</td>
					<td align=center class="TableCell">
						<% = rsStudent("IMPLEMENTED")%>		
					</td>
					<td align=center class="TableCell">
						<% = rsStudent("NA_SPONSOR")%>		
					</td>					
					<td align=center class="TableCell">
						<% = rsStudent("NA_ADMIN")%>		
					</td>
				</tr>
<%				rsStudent.MoveNext
				intColorCount = intColorCount + 1 
			loop	
		else
%>
				<tr>	
					<Td colspan=2 class=gray>
						&nbsp;No Pending ILP's are in the system.
					</td>
				</tr>
<%
		end if 
	rsStudent.Close
	set rsStudent = nothing	
	call oFunc.CloseCN
	set oFunc = nothing
%>			
			</table>
		</td>
	</tr>
</table>
<%
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")
%>