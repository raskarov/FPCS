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
dim     intStudent_ID 


Session.Contents("strTitle") = "Student/Sponsor Teacher List"
Session.Contents("strLastUpdate") = "05 May 2004"
if request("simpleHeader") = "" then
	Server.Execute(Application.Value("strWebRoot") & "includes/header.asp")
else
	Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")
end if

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()
ofunc.ResetSelectSessionVariables

%>
<script language=javascript>
	function jfAuthList(id){
		var winAuthAct;
		var strURL = "<%=Application.Value("strWebRoot")%>reports/vendorAuthList.asp?intVendor_ID="+id;
		winAuthAct = window.open(strURL,"winAuthAct","width=640,height=550,scrollbars=yes,resize=yes,resizable=yes");
		winAuthAct.moveTo(0,0);
		winAuthAct.focus();	
	}
</script>
<%Response.Write Session.Contents("ActiveEnrollList") %>
<form action="<%=Application("strSSLWebRoot")%>Reports/studentSponsors.asp" name="main" method="post">


<table width=100% ID="Table1">
	<tr>	
		<td class=yellowHeader >
				&nbsp;<b>Active Student/Sponsor Teacher List</b>
		</td>
	</tr>
	<tr>
	<td>
	<%'JD 052711 add order by %>
	    <select name="EntityOrder" onChange="this.form.submit();" ID="Select5">
			<option value="">Sort the list By
			<%
			    dim strBolValues
			    dim strBolText
			    strBolValues = "0,1"
			    strBolText = "Student,Teacher"									 
			    Response.Write oFunc.MakeList(strBolValues,strBolText,Request("EntityOrder"))												 
			%>
		</select>
		</td>
    </tr>
	<tr>
		<td bgcolor=f7f7f7>
			<table ID="Table2">
				<tr>	
					<Td class="TableHeader" valign=middle align=center>
						<B>Student Name (click for email)</b>
					</td>
					<Td class="TableHeader" valign=middle align=center>
						<b>Sponsor Teacher (click for email)
					</td>						
				</tr>
<%	
	'This section gives the classes for a student
set rs = server.CreateObject("ADODB.RECORDSET")
rs.CursorLocation = 3

'JD 052711 add order by student or sponsor
strOrderBy = ""

		if Request("EntityOrder") ="1" then
	        strOrderBy = " ORDER BY i.szLAST_NAME, i.szFIRST_NAME "
		else
		    strOrderBy = " ORDER BY s.szLAST_NAME, s.szFIRST_NAME "
		end if		


		
sql = "SELECT s.szLAST_NAME + ', ' + s.szFIRST_NAME AS studentName, " & _ 
		" i.szFIRST_NAME + ' ' + i.szLAST_NAME AS teacherName, i.szEmail, i.szHOME_PHONE,  " & _ 
		" f.szEMAIL AS smail, f.szHome_Phone AS sphone " & _ 
		"FROM tblSTUDENT s INNER JOIN " & _ 
		" tblENROLL_INFO ei ON s.intSTUDENT_ID = ei.intSTUDENT_ID INNER JOIN " & _ 
		" tblFAMILY f ON s.intFamily_ID = f.intFamily_ID INNER JOIN " & _ 
		" tblStudent_States ON s.intSTUDENT_ID = tblStudent_States.intStudent_id LEFT OUTER JOIN " & _ 
		" tblINSTRUCTOR i ON ei.intSponsor_Teacher_ID = i.intINSTRUCTOR_ID AND ei.intSponsor_Teacher_ID = i.intINSTRUCTOR_ID " & _ 
		"WHERE (ei.sintSCHOOL_YEAR = " & session.Contents("intSchool_Year") & ") AND (tblStudent_States.intReEnroll_State IN (" & Application.Contents("ActiveEnrollList") & ")) AND (tblStudent_States.intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _ 
		strOrderBy
		
		response.Write sql
rs.Open sql,oFunc.FPCScnn	

intColorCount = 0
if rs.RecordCount > 0 then
		do while not rs.EOF						
			if intColorCount mod 2 = 0 then
				strBgColor = " bgcolor=white " 
			else
				strBgColor = ""
			end if 
					
%>
		<tr <% = strBgColor %>>
			<Td class="TableCell" valign=top title="Phone: <% = rs("sPhone") %>"> 
				<a href="mailto:<% = rs("smail")%>"><% = rs("studentName") %></a>
			</td>					
			<td align=left class="TableCell" title="Phone: <% = rs("szHOME_PHONE") %>"> 
				<a href="mailto:<% = rs("szEmail")%>"><% = rs("teacherName") %></a>				
			</td>		
		</tr>
<%				rs.MoveNext
			intColorCount = intColorCount + 1 
		loop	
	else
%>
				<tr>	
					<Td colspan=2 class=gray>
						&nbsp;No Active Students for the School Year <% = session.contents("intSchool_Year") %>.
					</td>
				</tr>
<%
		end if 
	rs.Close
	set rs = nothing	
	call oFunc.CloseCN
	set oFunc = nothing
%>			
			</table>
		</td>
	</tr>
</table>
</form>
<%
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")
%>