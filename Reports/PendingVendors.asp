<%@ Language=VBScript %>
<%

Session.Value("strSimpleTitle") = "Pending Vendors Report"
Session.Value("strLastUpdate") = "09 Dec 2002"
Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

dim sql 
dim rs
set rs = server.CreateObject("ADODB.RECORDSET")
rs.CursorLocation = 3

sql = "SELECT intVendor_ID, szVendor_Name, szVendor_Phone, szVendor_Email, dtCREATE " & _
		"FROM tblVendors " & _
		"WHERE (bolApproved IS NULL) ORDER BY dtCREATE"
rs.Open sql, oFunc.FPCScnn

if rs.RecordCount > 0  then 
%>
<table>
	<tr>
		<td class=svplain10>
			Vendor Name
		</td>
		<td class=svplain10>
			Phone
		</td>
		<td class=svplain10>
			Email
		</td>
	</tr>
<%
	do while not rs.EOF
%>
	<tr>
		<td class=svplain10>
			<a href="javascript:" onclick="jfViewVendor('<%=rs("intVendor_ID")%>');>
			<% = rs("szVendor_Name")%>
			</a>
		</td>	
		<td class=svplain10>
			<% = rs("szVendor_Phone"
		</td>
	</tr>
<%
		rs.MoveNext
	loop
%>
</table>

SELECT     tblVendors.intVendor_ID, tblVendors.szVendor_Name, tblVendors.szVendor_Phone, tblVendors.szVendor_Email, tblVendors.dtCREATE, 
                      tblClass_Items.intClass_ID, tblClasses.szClass_Name, tblOrdered_Items.intILP_ID, tblOrdered_Items.intOrdered_Item_ID, 
                      tblClass_Items.intClass_Item_ID
FROM         tblClass_Items INNER JOIN
                      tblClasses tblClasses_1 ON tblClass_Items.intClass_ID = tblClasses_1.intClass_ID RIGHT OUTER JOIN
                      tblClasses INNER JOIN
                      tblILP ON tblClasses.intClass_ID = tblILP.intClass_ID INNER JOIN
                      tblOrdered_Items ON tblILP.intILP_ID = tblOrdered_Items.intILP_ID RIGHT OUTER JOIN
                      tblVendors ON tblOrdered_Items.intVendor_ID = tblVendors.intVendor_ID ON tblClass_Items.intVendor_ID = tblVendors.intVendor_ID
WHERE     (tblVendors.bolApproved IS NULL)
ORDER BY tblVendors.dtCREATE
