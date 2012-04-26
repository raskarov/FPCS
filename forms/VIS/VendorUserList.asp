<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		VendorUserList.asp
'Purpose:	Creates Vendor User list Showing Vendor name and user name
'Date:		July 7 2005
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim sql
dim oFunc
dim rs

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

' must be an admin to access this page
if not oFunc.IsAdmin then
	response.Write "<h1>Page Improperly Called</h1>"
	response.End
end if

if request("bolWin") = "" then
	Server.Execute(Application.Value("strWebRoot") & "includes/header.asp")
else
	Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")
end if

sql = "SELECT v.intVendor_ID, v.szVendor_Name, vu.szUser_ID, v.szVendor_Email, v.szVendor_Phone " & _ 
		"FROM tascVendor_User vu INNER JOIN " & _ 
		" tblVendors v ON vu.intVendor_ID = v.intVendor_ID " & _ 
		" WHERE bolService_Vendor = 1 " & _
		" AND (select top 1 upper(szVendor_Status_CD) from tblVendor_Status vs where vs.intVendor_ID = v.intVendor_ID and " & _
		" vs.intSchool_Year <= " & session.Contents("intSchool_Year") & _ 
		" order by intSchool_Year desc, intVendor_Status_ID desc) in ('APPR','PEND') " & _
		"ORDER BY v.szVendor_Name "
		
set rs = server.CreateObject("ADODB.RECORDSET")
rs.CursorLocation = 3
rs.Open sql, Application("cnnFPCS")'oFunc.FpcsCnn
rsCount = rs.RecordCount
%>
<table cellpadding="3">
	<tr>
		<td class="yellowHeader" colspan="10">
			<b>Service Vendor User List</b> (Approved and Pending Only)
		</td>
	</tr>
	<tr>
		<td class="TableHeader">
			&nbsp;<b>Vendor name</b>
		</td>
		<td class="TableHeader" align="center">
			<b>User Name</b>
		</td>
		<td class="TableHeader" align="center">
			<b>Default<BR>Password</b>
		</td>
		<td class="TableHeader" align="center">
			<b>Email</b>
		</td>
		<td class="TableHeader" align="center">
			<b>Phone #</b>
		</td>			
	</tr>
	<%
do while not rs.EOF
	%>
	<tr>
		<td class="TableCell">
			<% = rs("szVendor_Name") %>&nbsp;
		</td>
		<td class="TableCell">
			<% = rs("szUser_ID") %>&nbsp;
		</td>
		<td class="TableCell"  align="center">
			<% = rs("intVendor_ID") %>&nbsp;
		</td>
		<td class="TableCell">
			<a href="mailto:<%=rs("szVendor_Email")%>"><% = rs("szVendor_Email") %></a>&nbsp;
		</td>
		<td class="TableCell">
			<% = rs("szVendor_Phone") %>&nbsp;
		</td>		
	</tr>
	<%
	rs.MoveNext
loop
	
	%>
</table>
&nbsp;<span class="svplain8"><% = rsCount %> Service Vendors Returned</span>
<%
rs.Close
set rs = nothing
set oFunc = nothing
Server.Execute(Application.Value("strWebRoot") & "Includes/footer.asp")
%>