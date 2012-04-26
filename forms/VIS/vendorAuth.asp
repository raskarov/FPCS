<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		vendorAuth.asp
'Purpose:	Displays what items a vendor is authorized to
'			provide good/services for and if there are any reimbursement
'			or requisition.
'Date:		23-JAN-2003
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Dimention Variables, make db Connection, print HTML header.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

Session.Value("strTitle") = "Authorized Vendor Actions"
Session.Value("strLastUpdate") = "23 JAN 2003"

dim sql
dim intVendor_ID
intVendor_ID = request("intVendor_ID")

if request("bolSimple") <> "" then
	Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")
else
	Server.Execute(Application.Value("strWebRoot") & "includes/header.asp")
end if 
%>	
<script language=javascript>
	function jfAuthAdmin(){
		var winAuth;
		<% if request("bolSimple") <> "" then %>
		var ID = document.main.intItem_ID.value;
		var index = window.opener.document.main.intVendor_ID.selectedIndex;
		var szVendor_Name = window.opener.document.main.intVendor_ID.options(index).text;
		<% else %>		
		var ID = document.main.intItem_ID.value;
		var index = document.main.intVendor_ID.selectedIndex;
		var szVendor_Name = document.main.intVendor_ID.options(index).text;
		<% end if %>
		var URL = "vendorAuthAdmin.asp?intVendor_ID=<%=intVendor_ID%>&intItem_ID=" + ID;
		URL += "&szVendor_Name=" + szVendor_Name;		
		winAuth = window.open(URL,"winAuth","width=800,height=500,scrollbars=yes,resizable=on");
		winAuth.moveTo(0,0);
		winAuth.focus();		
	}
</script>	
<table width=100%>		
	<form name=main ID="Form1">	
	<tr>	
		<Td class=yellowHeader >
				&nbsp;<b>Authorized Vendor Actions: </b> 
				<% if request("bolSimple") = "" then %>
				<select name="intVendor_ID" onchange="window.location.href='<% = Application.Value("strWebRoot") %>forms/VIS/vendorAdmin.asp?intVendor_ID='+this.value+'&szVendor_Name='+this.text;" ID="Select2">
				<%
					dim sqlVendor
					sqlVendor = "Select intVendor_ID,szVendor_Name " & _
										"from tblVendors order by szVendor_Name"
					Response.Write oFunc.MakeListSQL(sqlVendor,"intVendor_ID","szVendor_Name", request("intVendor_ID"))	
				%>
				</select>
				<% else %>
				<input type=hidden name=intVendror_ID value="<% = request("intVendor_ID") %>">
				<% end if %>
		</td>
	</tr>
</table>
<table ID="Table1">	
	<tr>
		<td class=gray>
			<b>Add/Edit Authorized Item:</b>
		</td>
		<td>
			<select name="intItem_ID">
				<option>
			<%
				sql = "select intItem_ID,szName " &_
						"from trefItems " & _
						"order by szName "
				response.Write oFunc.MakeListSQL(sql,"intItem_ID","szName","")
			%>
			</select>
			<input type=button value="Submit" class="btSmallGray" onclick="jfAuthAdmin();" NAME="btSmallGray">
			<% if request("bolSimple") <> "" then %>
			<input type=button value="Close Window" class="btSmallGray" onclick="window.opener.focus();window.close();">
			<% end if %>
		</td>
	</tr>
	</form>
	<tr>
		<td colspan=2>
			<table ID="Table2">
<%
' This section shows what actions the vendor is allowed to take.
dim sqlAuth
dim rsAuth
set rsAuth = server.CreateObject("ADODB.RECORDSET")
rsAuth.CursorLocation = 3

sqlAuth = "SELECT ig.szName, i.szName AS item_name, pos.szSubject_Name, " & _
		  "va.bolReimburse_Only,va.bolRequisition_Only,va.bolRequisition_Only_OverRide " & _
		  "FROM trefItems i INNER JOIN " & _
          "trefItem_Groups ig ON i.intItem_Group_ID = ig.intItem_Group_ID " & _
          "INNER JOIN " & _
          "tblVendor_Auth va ON i.intItem_ID = va.intItem_ID INNER JOIN " & _
          "trefPOS_Subjects pos ON va.intPOS_SUbject_ID = pos.intPOS_Subject_ID " & _
		  "WHERE (va.intVendor_ID = " & request("intVendor_ID") & ")" & _
		  "order by ig.szName, i.szName,pos.szSubject_Name "
rsAuth.Open sqlAuth,Application("cnnFPCS")'oFunc.FPCScnn

if rsAuth.RecordCount > 0 then
%>
				<tr>
					<td class=gray>
						<b>Good/Service</b>
					</td>
					<td class=gray>
						<b>Item Type</b>
					</td>
					<td class=gray>
						<b>POS Subject</b>
					</td>
					<td class=gray align=center>
						<b>RMB Only</b>
					</td>
					<td class=gray align=center>
						<b>REQ Only</b>
					</td>
					<td class=gray align=center>
						<b>REQ Only Ovrd</b>
					</td>
				</tr>
<%
	do while not rsAuth.EOF
		%>
				<tr>
					<td class=gray>
						<% = rsAuth("szName") %>
					</td>
					<td class=gray>
						<% = rsAuth("item_name") %>
					</td>
					<td class=gray>
						<% = rsAuth("szSubject_Name") %>
					</td>
					<td class=gray align=center>
						<% = rsAuth("bolReimburse_Only") %>
					</td>
					<td class=gray align=center>
						<% = rsAuth("bolRequisition_Only") %>
					</td>
					<td class=gray align=center>
						<% = rsAuth("bolRequisition_Only_OverRide") %>
					</td>
				</tr>	
			
		<%
		rsAuth.MoveNext
	loop
else
%>
				<TR>
					<td class=gray>
						No Actions Authorized.
					</td>
				</tr>
				
				
<%	
end if
rsAuth.Close
set rsAuth = nothing
%>						
								
			</table>
		</td>
	</tr>
</table>
<%
call oFunc.CloseCN
Server.Execute(Application.Value("strWebRoot") & "Includes/footer.asp")
%>