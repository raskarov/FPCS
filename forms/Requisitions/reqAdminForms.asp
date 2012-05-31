<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		reqAdminForm.asp
'Purpose:	Gives user with admin rights the ability to print
'			requisitions based on family,vendor and date.
'Date:		04 JAN 2004
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Quick Security Check
If ucase(session.Contents("strRole")) <> "ADMIN" then
	response.Write "<h1>Page illegally called.</h1>"
	response.End
end if

dim oFunc			' Main functions object
dim sql				' string to contain sql query commands
dim objRequest		' Contains the incoming form info via the request object
dim strObjValue		' Contains the value of an item in the request collection
dim strWhere		' Refines req search sequal
dim intCount		' Used to keep track of html table rows

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

Session.Value("strTitle") = "Requisition Forma"
Session.Value("strLastUpdate") = "04 Jan 2004"
Server.Execute(Application.Value("strWebRoot") & "includes/Header.asp")

' Get needed form variables   
if Request.Form.Count > 0 then
	set objRequest = Request.Form
else
	set objRequest = Request.QueryString
end if

For Each Item in objRequest	
	execute("dim " & Item)
	strObjValue = objRequest(Item)
	execute(Item & " = strObjValue")
next

%>
<table width=100% bgcolor=f7f7f7 ID="Table1">
	<tr>	
		<Td class=yellowHeader>
			&nbsp;<b>Requisition Forms</b> &nbsp;
		</td>
	</tr>
	<tr>
		<td>
			<form action=reqAdminForms.asp method=post name=main>
			<input type=hidden name="lastRowColor">
			<input type=hidden name="lastRow">
			<table class=svplain10>
				<tr>
					<td bgcolor=e6e6e6 colspan=2>
						&nbsp;Filter Requisitions by
					</td>
				</tr>
				<tr>
					<td>
						&nbsp;Family&nbsp;
					</td>
					<td>
						<select name=intFamily_ID ID="Select1">
							<option value=""></option>
						<%
							sql = "select intFamily_ID, Name = " & _
									"CASE " & _
									"WHEN szDesc is null then szFamily_Name + ': ' +  convert(varchar,intFamily_ID) " & _
									"WHEN szDesc is not null then szFamily_Name + ', ' + szDesc + ': ' +  convert(varchar,intFamily_ID) " & _
									"END " & _
									"FROM tblFamily f " & _
									"WHERE exists(" & _
										"SELECT s.intSTUDENT_ID, " & _
										"Name = (Case ss.intReEnroll_State WHEN 86 then " & _
										"s.szLAST_NAME + ',' + s.szFIRST_NAME + ': Withdrawn (' + convert(varChar(20),ss.dtModify) + ')'" & _ 
										"WHEN 123 THEN s.szLAST_NAME + ',' + s.szFIRST_NAME + ': Graduated (' + convert(varChar(20),ss.dtModify) + ')'" & _ 
										"ELSE s.szLAST_NAME + ',' + s.szFIRST_NAME END) " & _
										"FROM tblSTUDENT s INNER JOIN " & _ 
										"tblStudent_States ss ON s.intSTUDENT_ID = ss.intStudent_id " & _ 
										"WHERE f.intFamily_ID = s.intFamily_ID " & _
										"and (ss.intReEnroll_State in (" & application.Contents("strEnrollmentList") & ")) " & _
										"AND (ss.intSchool_Year = " & Session.Value("intSchool_Year") & ") " & _ 
										")"  & _
									"ORDER BY 2"
							response.Write oFunc.MakeListSQL(sql,"intFamily_id","Name",intFamily_ID)
						%>
						</select>
					</td>
				</tr>
				<tr>
					<td>
						&nbsp;Vendor&nbsp;
					</td>
					<td>
						<select name=intVendor_ID ID="Select2">
							<option value=""></option>
						<%
							sql = "SELECT intVendor_ID, szVendor_Name + ': ' + CONVERT(varchar, intVendor_ID) AS Name " & _
									"FROM dbo.tblVendors " & _
									"WHERE (bolApproved = 1) " & _
									"ORDER BY szVendor_Name"
							response.Write oFunc.MakeListSQL(sql,"intVendor_ID","Name",intVendor_ID)
						%>
						</select>
					</td>
				</tr>			
			</table>
			<input type=button value="Submit" id=btSmallGray onclick="this.form.submit();">
			</form>
		</td>
	</tr>
</table>
<%
' Handle Requisition Search Request
sql = ""
if intFamily_ID <> "" and intVendor_ID <> "" then
	strWhere = " (intFamily_ID = " & intFamily_ID & ") " & _
			   "AND (intVendor_ID = " & intVendor_ID & ") "
elseif intFamily_ID <> "" then
	strWhere = " (intFamily_ID = " & intFamily_ID & ") "
elseif intVendor_ID <> "" then
	strWhere = "(intVendor_ID = " & intVendor_ID & ") "
end if
	
if strWhere <> "" then	
	set rs = server.CreateObject("ADODB.RECORDSET")
	rs.CursorLocation = 3 'adUseClient
	sql = "SELECT CONVERT(varChar, dtApproval_Changed, 101) AS date," & _
			"szVendor_Name, szFamily_Name, intFamily_ID, " & _
			" intVendor_ID, COUNT(*) AS item_count " & _
			"FROM dbo.v_Requisitions " & _
			"WHERE " & strWhere & _
			"GROUP BY CONVERT(varChar, dtApproval_Changed, 101), szVendor_Name, szFamily_Name,  " & _
			"intFamily_ID, intVendor_ID, intSchool_Year " & _
			"HAVING (intSchool_Year = '" & session.Contents("intSchool_Year") &"') " & _
			"ORDER BY szFamily_Name, szVendor_Name,  " & _
			"CONVERT(varChar, dtApproval_Changed, 101)"
	rs.Open sql,Application("cnnFPCS")'oFunc.FPCScnn
	if rs.RecordCount > 0 then
		call vbfPrintResults(rs)
	else
		response.Write "<span class=svplain8><b>No records found.</b></span>"
	end if
	rs.Close
	set rs = nothing
end if

' Close open items
response.Write "</body></html>"
oFunc.CloseCN()
set oFunc = nothing

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Subs and functions below here
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
function vbfPrintResults(pRS)
	' Prints out search results
%>
<script language=javascript>
	function jfHighLight(row){
		var obj = document.getElementById(row);
		var lastRow = document.main.lastRow.value;
		var lastRowColor = document.main.lastRowColor.value;	
		// Reset last row to its normal state
		if (lastRow != ""){	
			var obj2 = document.getElementById(lastRow);
			obj2.style.backgroundColor = lastRowColor;
		}
		// Highlight current row and retsain original info
		document.main.lastRowColor.value = obj.style.backgroundColor;
		document.main.lastRow.value = row;
		obj.style.backgroundColor = "e6e6e6";
	}
	function jfShowReq(date,famID,vendID){
		var reqWin;
		var URL = "<%=Application("strSSLWebRoot")%>forms/requisitions/reqForms.asp";
		URL += "?date="+date+"&intFamily_ID="+famID+"&intVendor_ID="+vendID
		reqWin = window.open(URL,"reqWin","width=790,height=550,scrollbars=yes,resizable=yes");
		reqWin.moveTo(0,0);
		reqWin.focus();
	}
</script>
<table class=svplain10 bordercolor=e6e6e6 cellspacing=0 cellpadding=4 border=1>
	<tr>
		<td colspan=4 class=NavyWhite10>			
			<table align=right cellpadding=0 cellspacing=0  class=NavyWhite8>
				<tr>
					<td>
						click on row to view reqs
					</td>
				</tr>
			</table>
			<B>Search Results</B>
		</td>
	</tr>
	<tr>
		<td align=center>
			Date
		</td>
		<td align=center>
			Vendor
		</td>
		<td align=center>
			Family
		</td>
		<td align=center>
			# of Items
		</td>
	</tr>
<%
	intCount = 0
	do while not pRS.eof
%>
	<tr id="tr<%=intCount%>" onClick="jfHighLight('tr<%=intCount%>');jfShowReq('<%=pRS(0)%>','<%=pRS(3)%>','<%=pRS(4)%>');" style="cursor:hand">
		<td align=center>
			&nbsp;<% = pRS(0) %>
		</td>
		<td>
			&nbsp;<% = pRS(1) %>&nbsp;
		</td>
		<td >
			&nbsp;<% = pRS(2) %>&nbsp;
		</td>
		<td align=center >
			&nbsp;<% = pRS(5) %>
		</td>
	</tr>
<%
		pRS.MoveNext
		intCount = intCount + 1
	loop
	
%>
</table>
<%
end function
%>