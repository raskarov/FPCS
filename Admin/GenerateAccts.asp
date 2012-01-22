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
dim strSQL1
dim strSQL2
dim strSQL3
dim iCounter

iCounter = 0
Session.Value("strTitle") = "Generate User Accounts"
Server.Execute(Application.Value("strWebRoot") & "Includes/header.asp")

set rs = Server.CreateObject("ADODB.Recordset")
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()
strSQL = Request("txtSQL")

strSQL = "SELECT tblUsers.szUser_ID, tblUsers.szPassword " & _
			"FROM tascUserRoles INNER JOIN " & _
			"tblUsers ON tascUserRoles.szUser_ID = tblUsers.szUser_ID " & _
			"WHERE(tascUserRoles.szRole_CD = 'TEACHER') " & _
			"ORDER BY tblUsers.szUser_ID"


strSQL1 = "INSERT INTO tblUsers " & _
				"(szUser_ID, szName_Last, szName_First, szEmail, blnActive, blnForcePWDchange) " & _
				"SELECT     UPPER(LEFT(szLAST_NAME, 6) + LEFT(szFIRST_NAME, 1) + CAST(intINSTRUCTOR_ID AS varchar(5))), UPPER(szLAST_NAME),  " & _
				"UPPER(szFIRST_NAME), szEmail, 1, 1 " & _
				"FROM         tblINSTRUCTOR"

strSQL2 = "INSERT INTO tascUserRoles " & _
				"(szUser_ID, szRole_CD) " & _
				"SELECT     tblUsers.szUser_ID, 'TEACHER' AS szRl " & _
				"FROM         tblINSTRUCTOR INNER JOIN " & _
				"tblUsers ON UPPER(LEFT(tblINSTRUCTOR.szLAST_NAME, 6) + LEFT(tblINSTRUCTOR.szFIRST_NAME, 1)  " & _
				"+ CAST(tblINSTRUCTOR.intINSTRUCTOR_ID AS varchar(5))) = tblUsers.szUser_ID"

strSQL3 = "INSERT INTO tascINSTR_USER " & _
				"(szUser_ID, intINSTRUCTOR_ID) " & _
				"SELECT     tblUsers.szUser_ID, tblINSTRUCTOR.intINSTRUCTOR_ID " & _
				"FROM         tblINSTRUCTOR INNER JOIN " & _
				"tblUsers ON UPPER(LEFT(tblINSTRUCTOR.szLAST_NAME, 6) + LEFT(tblINSTRUCTOR.szFIRST_NAME, 1)  " & _
				"+ CAST(tblINSTRUCTOR.intINSTRUCTOR_ID AS varchar(5))) = tblUsers.szUser_ID"



set oCrypto = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/Crypto.wsc"))
if Request("chkTeachers") = "on" then
'if Trim(strSQL) <> "" then
	oFunc.ExecuteCN(strSQL1)
	oFunc.ExecuteCN(strSQL2)
	oFunc.ExecuteCN(strSQL3)		
	with rs
		.CursorLocation = 3
		.LockType = 4 'adLockBatchOptimistic
		.Open strSQL, oFunc.FPCScnn
		if not .BOF and not .EOF then
			do until .EOF
				'decrypt password for database compare
				'oCrypto.Key = "something"	'actual key is not shown here
				oCrypto.Text = LCase(rs("szUser_ID"))
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
<FORM name="frmMain" method="post">
<TABLE>
	<TR>
		<TD><INPUT type="checkbox" name="chkTeachers"></TD>
		<TD style="FONT-FAMILY:Verdana">Teachers</TD>
	</TR>
	<TR>
		<TD><INPUT type="checkbox" name="chkGuardians"></TD>
		<TD style="FONT-FAMILY:Verdana">Guardians</TD>
	</TR>
</TABLE>
<BR>
<BR>
<TEXTAREA name="txtSQL" cols="70" rows="10"><% = strSQL %></TEXTAREA>
<BR>
<INPUT type="submit" name="Submit">
<%
	set wscLists = nothing
	set oFunc = nothing
'end if
Server.Execute(Application.Value("strWebRoot") & "Includes/footer.asp")

%>
