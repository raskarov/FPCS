<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		vendorAdmin.asp
'Purpose:	Admin tool for adding/viewing/modifying Vendor information
'Date:		9-04-01
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Dimention Variables, make db Connection, print HTML header.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()
Session.Value("strTitle") = "Vendor Authorizations"
Session.Value("strLastUpdate") = "15 Jan 2003"
dim blnFullDisplay	'if True will display the full page to the user, otherwise the scalled down version will display

Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")

' This section shows what actions the vendor is allowed to take.
dim sqlAuth
dim rsAuth
dim sqlItems
dim rsItems
dim bolAuth
dim update
dim delete
dim insert
dim strAlert
dim strClose

if request.Form("intCount") <> "" then
	'Form has been submitted
	oFunc.BeginTransCN
	for i = 0 to request.Form("intCount")
		if request.Form("intVendor_Auth_ID"&i) <> "" then
			if request.Form("bolRights"&i) = 1 then
				' update an Auth record
				update = "update tblVendor_Auth set " & _
						"intPOS_Subject_ID = " & request.Form("intPOS_Subject_ID"&i) & _
						", bolReimburse_Only = " & request.Form("bolReimburse_Only"&i) & _
						", bolRequisition_Only = " & request.Form("bolRequisition_Only"&i) & _
						", bolRequisition_Only_OverRide = " & request.Form("bolRequisition_Only_OverRide"&i) & _
						",dtModify = convert(dateTime,'" & now() & "')" & _
						",szUser_Modify = '" & session.Contents("strUserID") & "' " & _
						"where intVendor_Auth_ID = " & request.Form("intVendor_Auth_ID"&i)
				oFunc.ExecuteCN(update)
			else
				'Delete an Auth record
				delete = "delete from tblVendor_Auth where intVendor_Auth_ID = " & request.Form("intVendor_Auth_ID"&i)
				oFunc.ExecuteCN(delete)
			end if
		elseif request.Form("bolRights"&i) = 1 then
			' insert an Auth record
			insert = "insert into tblVendor_Auth(intVendor_ID,intItem_ID,intPOS_Subject_ID," & _
					 "bolReimburse_Only,bolRequisition_Only,bolRequisition_Only_OverRide," & _
					 "dtCreate,szUser_Create) values (" & _
					 request.Form("intVendor_ID") & "," & _
					 request.Form("intItem_ID") & "," & _
					 request.Form("intPOS_Subject_ID"&i) & "," & _
					 request.Form("bolReimburse_Only"&i) & "," & _
					 request.Form("bolRequisition_Only"&i) & "," & _
					 request.Form("bolRequisition_Only_OverRide"&i) & "," & _
					 "convert(DateTime,'" & now() & "')," & _
					 "'" & session.Contents("strUserID") & "')"
			oFunc.ExecuteCN(insert)
		end if
	next
	oFunc.CommitTransCN
	strAlert = "alert('Update Made');"
	strClose = "jfClose();"
end if 
%>
<script language=javascript>
	<% = strClose %>
	function jfClose(){		
		<% = strAlert %>
		window.opener.location.reload();window.opener.focus();window.close();
	}
</script>
<table ID="Table1">
	<form action=vendorAuthAdmin.asp method=post name=main>
	<input type=hidden name="intVendor_ID" value="<% = request("intVendor_ID")%>" ID="Hidden1">
	<input type=hidden name="intItem_ID" value="<% = request("intItem_ID") %>" ID="Hidden2">
	<input type=hidden name="szVendor_Name" value="<% = request("szVendor_Name") %>">
	<tr>	
		<Td colspan=2 class=yellowHeader>
			&nbsp;<b><i>Authorized Vendor Actions for <% = request("szVendor_Name") %></I></B> &nbsp;
			<input type=button value="Close without saving" onclick="jfClose();" class="btSmallGray" >
			<input type=submit value="SAVE" class="NavSave">
		</td>
	</tr>	
	<tr>
<%
set rsAuth = server.CreateObject("ADODB.RECORDSET")
rsAuth.CursorLocation = 3

sqlAuth = "SELECT trefItems.szName AS item, trefItem_Groups.szName AS iGroup, " & _
		  "tblVendor_Auth.intVendor_Auth_ID, tblVendor_Auth.intPOS_SUbject_ID, " & _
		  "tblVendor_Auth.bolReimburse_Only, tblVendor_Auth.bolRequisition_Only, " & _
		  "tblVendor_Auth.bolRequisition_Only_OverRide, tblVendors.szVendor_Name," & _
		  "ps.szSubject_Name " & _
		  "FROM tblVendor_Auth INNER JOIN " & _
		  "trefItems ON tblVendor_Auth.intItem_ID = trefItems.intItem_ID INNER JOIN " & _
		  "trefItem_Groups ON trefItems.intItem_Group_ID = trefItem_Groups.intItem_Group_ID INNER JOIN " & _
          "tblVendors ON tblVendor_Auth.intVendor_ID = tblVendors.intVendor_ID INNER JOIN " & _
          "trefPOS_Subjects ps ON tblVendor_Auth.intPOS_Subject_ID = ps.intPOS_Subject_ID " & _
		  "WHERE (tblVendor_Auth.intVendor_ID = " & request("intVendor_ID") & ")" & _
		  "AND (tblVendor_Auth.intItem_ID = " & request("intItem_ID") & ") " & _
		  "UNION " & _
		  "SELECT i.szName AS item, ig.szName AS iGroup, NULL AS Expr1, ps.intPOS_Subject_ID, " & _
		  "CONVERT(bit, 0) AS Expr2, CONVERT(bit, 0) AS Expr3, CONVERT(bit, 0) " & _
		  "AS Expr4, v.szVendor_Name,ps.szSubject_Name " & _
		  "FROM trefItems i INNER JOIN " & _
		  "trefItem_Groups ig ON i.intItem_Group_ID = ig.intItem_Group_ID CROSS JOIN " & _
		  "tblVendors v CROSS JOIN " & _
		  "trefPOS_Subjects ps " & _
		  "WHERE (v.intVendor_ID = " & request("intVendor_ID") & ") " & _
		  "AND (i.intItem_ID = " & request("intItem_ID") & ") " & _
		  "AND (NOT EXISTS " & _
          "(SELECT     'x' " & _
		  "   FROM tblVendor_AUth va " & _
          "   WHERE ps.intPOS_SUbject_ID = va.intPOS_SUbject_ID AND " & _
          "   (va.intItem_ID = " & request("intItem_ID") & ") " & _
          "    AND (va.intVendor_ID = " & request("intVendor_ID") & ")))" & _
		  "ORDER BY 9"
rsAuth.open sqlAuth, oFunc.FPCScnn

if rsAuth.RecordCount > 0 then
%>
		<td>
			<table ID="Table2">
				<tr>	
					<td class=gray>
						&nbsp;<b>POS Subjects</b>
					</td>
					<td class=gray align=center>
						<b>Rights</b>
					</td>
					<td class=gray title="Reimbursement Only">
						<b>RMB Only</b>
					</td>
					<td class=gray title="Requisition Only">
						<b>RQ Only</b>
					</td>
					<td class=gray title="Requisition Only Over Ride">
						<b>RQ Only Ovrd</b>
					</td>
				</tr>
<%
	dim intCount
	intCount = 0
	do while not rsAuth.EOF						
		%>
				<tr>	
					<input type=hidden name="intVendor_Auth_ID<% = intCount %>" value="<% = rsAuth("intVendor_Auth_ID") %>">
					<td class=gray>
						<input type=hidden name="intPOS_Subject_ID<% = intCount %>" value="<%=rsAuth("intPOS_Subject_ID")%>">						
						&nbsp;<% = rsAuth("szSubject_Name") %>
					</td>		
					<td class=gray align=center>
						<select name=bolRights<% = intCount %> ID="Select3">
						<%
							response.Write oFunc.MakeList("0,1","Denied,Granted",oFunc.TrueFalse(IsNumeric(rsAuth("intVendor_Auth_ID"))))
						%>
						</select>	
					</td>												
					<td class=gray align=center>
						<select name=bolReimburse_Only<% = intCount %>>
						<%
							response.Write oFunc.MakeList("0,1","False,True",oFunc.TrueFalse(rsAuth("bolReimburse_Only")))
						%>
						</select>	
					</td>
					<td class=gray align=center>
						<select name=bolRequisition_Only<% = intCount %> ID="Select1">
						<%
							response.Write oFunc.MakeList("0,1","False,True",oFunc.TrueFalse(rsAuth("bolRequisition_Only")))
						%>
						</select>		
					</td>
					<td class=gray align=center>
						<select name=bolRequisition_Only_OverRide<% = intCount %> ID="Select2">
						<%
							response.Write oFunc.MakeList("0,1","False,True",oFunc.TrueFalse(rsAuth("bolRequisition_Only_OverRide")))
						%>
						</select>		
					</td>
				</tr>	
			
		<%
		intCount = intCount + 1
		rsAuth.MoveNext
	loop
end if

rsAuth.Close
set rsAuth = nothing
	
%>						
								
			</table>
		</td>
	</tr>
	<tr>	
		<td colspan=10>
			<input type=hidden name=intCount value="<% = intCount %>">
			<input type=button value="Close without saving" onclick="window.opener.location.reload();window.opener.focus();window.close();" class="btSmallGray">
			<input type=submit value="SAVE" class="NavSave">
		</td>
	</tr>
	</form>
</table>
<%
call oFunc.CloseCN()
set oFunc = nothing
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")
%>