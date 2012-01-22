<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		CreateVendorAccts.asp
'Purpose:	Creates Vendor login accounts for Approved vendors that 
'			do not have an existing user account
'Date:		June 17 2005
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

if request("btCreate") = "" then
%>
<form action="CreateVendorAccts.asp" method="post">
	<span class="svplain8">
		This page will create user accounts for <b>vendors</b> that are <b>approved, pending and removed</b> but do not
		currently have a user account.  <br><br>
		To create vendor accounts click the button below.<br><br>
		<input type="submit" value="Create Accounts" class="NavSave" name="btCreate">
	</span>
</form>
<%
else
	sql = "SELECT     intVendor_ID, szVendor_Name,szContact_First_Name, szContact_Last_Name, szVendor_Email " & _ 
			"FROM         tblVendors v " & _ 
			"WHERE     ((SELECT     TOP 1 upper(szVendor_Status_CD) " & _ 
			"                         FROM         tblVendor_Status vs " & _ 
			"                         WHERE     vs.intVendor_ID = v.intVendor_ID AND vs.intSchool_Year <= " & session.Contents("intSchool_Year") & _ 
			"                         ORDER BY intSchool_Year DESC, intVendor_Status_ID DESC) in ('APPR','PEND','REMV')) AND (NOT EXISTS " & _ 
			"                          (SELECT     'x' " & _ 
			"                            FROM          tascVendor_User vu " & _ 
			"                            WHERE      vu.intVendor_ID = v.intVendor_ID)) " & _ 
			"ORDER BY szVendor_Name "
	set rs = server.CreateObject("ADODB.RECORDSET")
	rs.CursorLocation = 3
	rs.Open sql, oFunc.FpcsCnn
	
	if rs.RecordCount > 0 then
		set oCrypto = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/Crypto.wsc"))
		do while not rs.EOF
			oFunc.BeginTransCN
			strUserName = ucase(replace(rs("szVendor_Name")," ",""))
			strUserName = replace(left(strUserName,5) & rs("intVendor_ID"),"'","")
			oCrypto.Text = rs("intVendor_ID")		
			Call oCrypto.Encypttext
			strEncPwd = oCrypto.EncryptedText
			insert = "insert into tblUsers (szUser_Id,szPassword,blnActive,blnForcePWDchange,dtCreate,szUser_Create)" & _
					 " values('" & _
					 oFunc.EscapeTick(strUserName) & "','" & strEncPwd & "',1,0,CURRENT_TIMESTAMP,'" & session.Contents("strUserID") & "')"
			oFunc.ExecuteCN(insert)
			
			insert = "insert into tascUserRoles(szUser_ID, szRole_CD,dtCreate,szUser_Create) " & _
					 " values ('" & oFunc.EscapeTick(strUserName) & "', 'VENDOR',CURRENT_TIMESTAMP,'" & session.Contents("strUserID") & "')"
			oFunc.ExecuteCN(insert)
			
			insert = "insert into tascVendor_User(intVendor_ID, szUser_Id,dtCreate,szUser_Create)" & _
					 " values (" & rs("intVendor_ID") & ",'" & _
					 oFunc.EscapeTick(strUserName) & "',CURRENT_TIMESTAMP,'" & session.Contents("strUserID") & "')"
			oFunc.ExecuteCN(insert)
			oFunc.CommitTransCN
			rs.MoveNext
		loop
		set oCrypto = nothing
	end if
	%>
	<span class="svplain8"><% = rs.RecordCount %> Vendor Accounts were created.</span>
	<%
	rs.Close
	set rs = nothing
end if  ' ends request("btCreate") = "" 
set oFunc = nothing
Server.Execute(Application.Value("strWebRoot") & "Includes/footer.asp")
%>