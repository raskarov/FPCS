<%@ Language=VBScript %>
<%
'*******************************************
'Name:		Admin\GeneratePW.asp
'Purpose:	Allows FPCS staff to generate encrypted passwords
'
'CalledBy:	
'
'Author:	ThreeShapes.com LLC
'Date:   23 April 2002
'*******************************************
'option explicit
dim wscLists
dim oFunc
dim oCrypto		'wsc object
dim strEncPwd
dim rs
dim strSQL
dim iCounter

iCounter = 0
Session.Value("strTitle") = "Generate Encrypted Passwords"
Server.Execute(Application.Value("strWebRoot") & "Includes/header.asp")

set rs = Server.CreateObject("ADODB.Recordset")
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()
strSQL = Request("txtSQL")

'"SELECT tblUsers.szUser_ID, tblUsers.szPassword " & _
'			"FROM tascUserRoles INNER JOIN " & _
'			"tblUsers ON tascUserRoles.szUser_ID = tblUsers.szUser_ID " & _
'			"WHERE(tascUserRoles.szRole_CD = 'TEACHER') " & _
'			"ORDER BY tblUsers.szUser_ID"
set oCrypto = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/Crypto.wsc"))
if Trim(strSQL) <> "" then
	with rs
		.CursorLocation = 3
		.LockType = 4 'adLockBatchOptimistic
		.Open strSQL, oFunc.FPCScnn
		if not .BOF and not .EOF then
			do until .EOF
				'decrypt password for database compare
				'oCrypto.Key = "something"	'actual key is not shown here
				oCrypto.Text = LCase(rs("PWD"))
				Call oCrypto.Encypttext
				strEncPwd = oCrypto.EncryptedText
				rs("szPassword") = strEncPwd
				iCounter = iCounter + 1
				.MoveNext
			loop
			.UpdateBatch
		end if
		.Close
		response.Write "Completed " & iCounter & " updates.<BR><BR>"
	end with
end if
set rs = nothing
set oCrypto = nothing	
%>				
<FORM name="frmMain" method="post" id="Form1">
<TEXTAREA name="txtSQL" cols="70" rows="10" id="Textarea1"><% = strSQL %></TEXTAREA>
<BR>
<INPUT type="submit" name="Submit" id="Submit1">
<%
	set wscLists = nothing
	set oFunc = nothing
'end if
Server.Execute(Application.Value("strWebRoot") & "Includes/footer.asp")

%>
