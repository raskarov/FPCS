<%@ Language=VBScript %>
<%
' TOGGLES SHOWING GOODS/SERVICES 
if session.Contents("strRole") = "GUARD" then
	response.write "<h1>Page Improperly Called.</h1>"
	response.end
end if

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Dimention Variables, make db Connection, print HTML header.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim intStudent_ID 

Session.Contents("strTitle") = "Vendor List"
Session.Contents("strLastUpdate") = "05 May 2004"

Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()
ofunc.ResetSelectSessionVariables

%>
<script language=javascript>
	function jfPacket(id){
		var winILPPend;
		var strURL = "<%=Application.Value("strWebRoot")%>forms/ILP/ILPShortForm.asp?simpleHeader=true&intStudent_ID="+id;
		winILPPend = window.open(strURL,"winILPPend","width=800,height=550,scrollbars=yes,resize=yes,resizable=yes");
		winILPPend.moveTo(0,0);
		winILPPend.focus();	
	}
</script>
<table width=100% ID="Table1">
	<tr>	
		<Td class=yellowHeader >
				&nbsp;<b>Vendor Authorized Services</b>
		</td>
	</tr>
	<tr>
		<td bgcolor=f7f7f7>
			<table ID="Table2">
				<tr>	
					<Td class=gray valign=middle align=center>
						&nbsp;<B>Vendor Name</b>&nbsp;
					</td>
					<Td class=gray valign=middle align=center>
						&nbsp;<b>Address&nbsp;
					</td>	
					<Td class=gray valign=middle align=center>
						&nbsp;<b>Phone&nbsp;
					</td>		
					<Td class=gray valign=middle align=center>
						&nbsp;<b>Email&nbsp;
					</td>				
					<td rowspan=2 valign=top>
						<input type=button value="Close Window" id=btSmallGray onclick="window.close();window.opener.focus();">
					</td>			
				</tr>
<%	
	'This section gives the classes for a student
set rsVendor = server.CreateObject("ADODB.RECORDSET")
rsVendor.CursorLocation = 3
sqlVendor = "SELECT intVendor_ID, szVendor_Name, VendorAddress, VendorCity," & _
			" VendorState, VendorZip,szVendor_Phone,szVendor_Email " & _
			"FROM v_Active_Vendors " & _
			"WHERE (intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _
			" AND intVendor_ID = " & request("intVendor_ID") & _
			" ORDER BY szVendor_Name"

rsVendor.Open sqlVendor,Application("cnnFPCS")'oFunc.FPCScnn	

intColorCount = 0
if rsVendor.RecordCount > 0 then
		do while not rsVendor.EOF						
			if intColorCount mod 2 = 0 then
				strBgColor = " bgcolor=white " 
			else
				strBgColor = ""
			end if 
					
%>
		<tr <% = strBgColor %>>
			<Td class = svplain11 valign=top> 
				&nbsp;<% = rsVendor("szVendor_Name") %>&nbsp;
			</td>					
			<td align=center class = svplain11>
				<% = rsVendor("VendorAddress")%><br>
				<% = rsVendor("VendorCity")%>, <% = rsVendor("VendorState")%> <% = rsVendor("VendorZip")%> 
			</td>
			<Td class = svplain11 valign=top> 
				&nbsp;<% = rsVendor("szVendor_Phone") %>&nbsp;
			</td>
			<Td class = svplain11 valign=top> 
				&nbsp;<a href="mailto:<% = rsVendor("szVendor_Email") %>"><% = rsVendor("szVendor_Email") %></a>&nbsp;
			</td>
		</tr>
<%				rsVendor.MoveNext
			intColorCount = intColorCount + 1 
		loop	
	else
%>
				<tr>	
					<Td colspan=2 class=gray>
						&nbsp;No Active Vendors for the School Year <% = session.contents("intSchool_Year") %>.
					</td>
				</tr>
<%
	end if 
rsVendor.Close
%>							
			</table>
			<br>
			<table border=1 bordercolor=c0c0c0 cellpadding=2 cellspacing=0>
				<tr>
					<td class=gray>
						<nobr>&nbsp;<B>Category</B>&nbsp;</nobr>
					</td>
					<td class=gray>
						&nbsp;<B>Authorized Subjects</B>
					</td>
				</tr>
<%
sqlVendor = "SELECT tblVendor_Auth.intVendor_ID, trefItems.szName, trefPOS_Subjects.szSubject_Name " & _
			"FROM tblVendor_Auth INNER JOIN " & _
			" trefItems ON tblVendor_Auth.intItem_ID = trefItems.intItem_ID INNER JOIN " & _
			" trefPOS_Subjects ON tblVendor_Auth.intPOS_SUbject_ID = trefPOS_Subjects.intPOS_Subject_ID " & _
			"WHERE (tblVendor_Auth.intVendor_ID = " & request("intVendor_ID") & ") " & _
			"ORDER BY trefItems.szName, trefPOS_Subjects.szSubject_Name"
rsVendor.Open sqlVendor,Application("cnnFPCS")'oFunc.FPCScnn	

dim strLastItem
if rsVendor.RecordCount > 0 then	
	do while not rsVendor.EOF						
		if strLastItem & "" <> rsVendor("szName") then		
			if strList <> "" then
%>
					<td class=svplain11 valign=top>
						<% = ucase(strList) %>
					</td>
				</tr>
<%
			end if
%>				
				<tr>
					<td class=svplain11 valign=top>
						<% = rsVendor("szName") %>
					</td>			
<%
			strList = rsVendor("szSubject_Name")
		else
			strList = strList & ", " & rsVendor("szSubject_Name")
		end if
		strLastItem = rsVendor("szName")
		rsVendor.MoveNext
	loop					
end if	
%>
					<td class=svplain11 valign=top>
						<% = ucase(strList) %>
					</td>
				</tr>
<%				
set rsVendor = nothing	
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