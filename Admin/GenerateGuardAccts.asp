<%@ Language=VBScript %>
<%
'*******************************************
'Name:		Admin\GenerateGuardAccts.asp
'Purpose:	Allows FPCS staff to generate User 
'			accounts for guardians that
'			1) previously had no user account AND
'			2) belong to a family that has a student that
'				recently was re-enrolled for the given school year
'
'NOTE:		This page does NOT create Guardian accounts.  It simply creates
'			user accounst for existing guardians that meet the criteria above.
'			When a Student wins the Lottery, the guardians associated with that
'			student still have to be created manually, especially in light of the
'			fact that the guardian might already exist for a previous student
'
'CalledBy:	
'
'Author:	ThreeShapes.com LLC
'Date:		12 June 2003
'*******************************************
'option explicit
dim oFunc
dim oCrypto		'wsc object
dim strEncPwd	'Encrypted Password
dim rs			'used to encrypt password
dim strSQL		'used by rs to encrypt password
dim strSQL1		'Creates the User account
dim strSQL2		'Creates the association tascUserRoles as a GUARD
dim strSQL3		'Creates the assocation in tascGUARD_USERS
dim strSQL4		'Creates the assocation in tascUsers_Action to force password changes
dim iCounter	'counts the records where the password was encrypted
dim lngNow		'date stamp converted to a long (then into a strig)

	if Session.Value("strRole") <> "ADMIN" then 
		response.Write "We're sorry, but only authorized staff members may access this page."
		response.End
	end if
	iCounter = 0
	Session.Value("strTitle") = "Generate Guardian User Accounts"
	Server.Execute(Application.Value("strWebRoot") & "Includes/header.asp")

	set rs = Server.CreateObject("ADODB.Recordset")
	set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
	call oFunc.OpenCN()

	lngNow = clng(now()): lngNow = cstr(lngNow)

	'strSQL =	"SELECT tblUsers.szUser_ID, tblUsers.szPassword " & _
	'			"FROM   tblUsers " & _
	'			"WHERE  tblUsers.szPassword = '" & lngNow & "CreateGuard' "

	strSQL =	"SELECT     tblUsers.szUser_ID, tblUsers.szName_Last, tblUsers.szName_First, tblUsers.szPassword, tascGUARD_USERS.intGUARDIAN_ID,  " & _
				"					LOWER(LEFT(REPLACE(tblUsers.szName_Last, ' ', ''), 3) + CAST(tascGUARD_USERS.intGUARDIAN_ID AS varchar(5))  " & _
				"					+ LEFT(REPLACE(tblUsers.szName_First, ' ', ''), 3)) AS PWD " & _
				"FROM         tblUsers INNER JOIN " & _
				"					tascGUARD_USERS ON tblUsers.szUser_ID = tascGUARD_USERS.szUser_ID INNER JOIN " & _
				"					tascUsers_Action ON tblUsers.szUser_ID = tascUsers_Action.szUser_ID " & _
				"WHERE     (tascUsers_Action.intAction_ID = 2) AND (tblUsers.blnActive = 1) and tblUsers.szPassword = '" & lngNow & "CreateGuard'"


'	strSQL1 = "INSERT INTO tblUsers " & _
'					"(szUser_ID, szName_Last, szName_First, szEmail, blnActive, blnForcePWDchange, szPassword, szUser_Create) " & _
'					"SELECT     UPPER(LEFT(szLAST_NAME, 6) + LEFT(szFIRST_NAME, 1) + CAST(intGUARDIAN_ID AS varchar(5))), UPPER(szLAST_NAME),  " & _
'					"UPPER(szFIRST_NAME), szEmail, 1, 1, '" & lngNow & "CreateGuard', '" & Session.Value("strUserID") & "' " & _
'					"FROM  (SELECT DISTINCT " & _
'					"      tblGUARDIAN.intGUARDIAN_ID, tblGUARDIAN.szLAST_NAME, tblGUARDIAN.szFIRST_NAME, tblGUARDIAN.szEMAIL, tblGUARDIAN.sMID_INITIAL " & _
'					"	FROM         tblUsers RIGHT OUTER JOIN " & _
 '                   "		tascFAM_GUARD INNER JOIN " & _
  '                  "		tblGUARDIAN ON tascFAM_GUARD.intGUARDIAN_ID = tblGUARDIAN.intGUARDIAN_ID INNER JOIN " & _
'                    "		tblSTUDENT INNER JOIN " & _
'                    "		tblStudent_States ON tblSTUDENT.intSTUDENT_ID = tblStudent_States.intStudent_id ON tascFAM_GUARD.intFamily_ID = tblSTUDENT.intFamily_ID ON  " & _
'                    "		tblUsers.szUser_ID = UPPER(LEFT(tblGUARDIAN.szLAST_NAME, 6) + LEFT(tblGUARDIAN.szFIRST_NAME, 1)  " & _
'                    "		+ CAST(tblGUARDIAN.intGUARDIAN_ID AS varchar(5))) LEFT OUTER JOIN " & _
'                    "		tascGUARD_USERS ON tblGUARDIAN.intGUARDIAN_ID = tascGUARD_USERS.intGUARDIAN_ID" & _
					'"	WHERE (tascGUARD_USERS.intGuard_User_ID IS NULL) AND tblStudent_States.intReEnroll_State IN (" & Application.Contents("ActiveEnrollList") & ")  AND (tblStudent_States.intSchool_Year = " & Request("selSchoolYear") & ") AND (tblUsers.szUser_ID IS NULL)) X"

STRsql1 = "INSERT INTO tblUsers " & _
	  "(szUser_ID, szName_Last, szName_First, szEmail, blnActive, blnForcePWDchange, szPassword, szUser_Create, szUser_Modify) " & _
"SELECT COALESCE " & _
"( " & _
"	( " & _
"		SELECT UPPER(LEFT(szLAST_NAME, 6) + LEFT(szFIRST_NAME, 1) + CAST(intGUARDIAN_ID AS varchar(5))) " & _
"		FROM TBLUSERS " & _
"		WHERE SZUSER_ID = SZlAST_NAME + '_' + SZFIRST_NAME " & _
"	), " & _
"	UPPER(SZlAST_NAME + '_' + SZFIRST_NAME) " & _
"), " & _
"UPPER(szLAST_NAME), UPPER(szFIRST_NAME), szEmail, 1, 1, '40313CreateGuard', 'scott', intGUARDIAN_ID " & _
"FROM  " & _
"(SELECT DISTINCT tblGUARDIAN.intGUARDIAN_ID, tblGUARDIAN.szLAST_NAME, tblGUARDIAN.szFIRST_NAME, tblGUARDIAN.szEMAIL,  " & _
"tblGUARDIAN.sMID_INITIAL " & _
"FROM  " & _
"tascFAM_GUARD  " & _
"INNER JOIN tblGUARDIAN ON tascFAM_GUARD.intGUARDIAN_ID = tblGUARDIAN.intGUARDIAN_ID  " & _
"INNER JOIN tblSTUDENT  " & _
"INNER JOIN tblStudent_States ON tblSTUDENT.intSTUDENT_ID = tblStudent_States.intStudent_id  " & _
"ON tascFAM_GUARD.intFamily_ID = tblSTUDENT.intFamily_ID  " & _
"LEFT OUTER JOIN tascGUARD_USERS ON tblGUARDIAN.intGUARDIAN_ID = tascGUARD_USERS.intGUARDIAN_ID  " & _
"WHERE (tascGUARD_USERS.intGuard_User_ID IS NULL) AND tblStudent_States.intReEnroll_State IN (" & Application.Contents("ActiveEnrollList") & ")  " & _
"AND (tblStudent_States.intSchool_Year = " & Request("selSchoolYear") & ")  " & _
") X"



	strSQL2 = "INSERT INTO tascUserRoles " & _
					"(szUser_ID, szRole_CD, szUser_Create) " & _
					"SELECT tblUsers.szUser_ID, 'GUARD' AS szRl, '" & Session.Value("strUserID") & "' " & _
					"FROM   tblUsers " & _
					"WHERE  tblUsers.szPassword = '" & lngNow & "CreateGuard' "

	strSQL3 = "INSERT INTO tascGUARD_USERS " & _
					"(szUser_ID, intGUARDIAN_ID, szUser_Create) " & _
					"SELECT     tblUsers.szUser_ID, tblGUARDIAN.intGUARDIAN_ID, '" & Session.Value("strUserID") & "' " & _
					"FROM         tblGUARDIAN INNER JOIN " & _
					"tblUsers ON convert(varchar,tblguardian.intguardian_id) = tblusers.szuser_modify " & _
					"WHERE tblUsers.szPassword = '" & lngNow & "CreateGuard' "
	
	strSQL4 = "INSERT INTO tascUsers_Action " & _
					"(szUser_ID, intAction_ID, intOrder_ID) " & _
					"SELECT tblUsers.szUser_ID, 2, 1 " & _
					"FROM   tblUsers " & _
					"WHERE tblUsers.szPassword = '" & lngNow & "CreateGuard' "

	set oCrypto = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/Crypto.wsc"))
	if Request("cmdOK") = "OK" then
		'response.Write strSQL1 & "<BR><BR>"
		'response.Write strSQL2 & "<BR><BR>"
		'response.Write strSQL3 & "<BR><BR>"
		'response.Write strSQL & "<BR>"

		oFunc.ExecuteCN(strSQL1)
		oFunc.ExecuteCN(strSQL2)
		oFunc.ExecuteCN(strSQL3)	
		oFunc.ExecuteCN(strSQL4)	
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
			response.Write "<div class='svplain10'><font color=red>Created/Updated " & _
				iCounter & " User Accounts with an Encrypted Password.</font></div><BR><BR>"
		end with
	end if
	set rs = nothing
	set oCrypto = nothing	
%>	
<script language="javascript">
	function jfPreview(){
		var winPrevGuard;
		var URL = "PreviewGuardAccts.asp?selSchoolYear=" + frmMain.selSchoolYear.value;
		winPrevGuard = window.open(URL,"winPrevGuard","width=800,height=500,scrollbars=yes,resizable=on");
		winPrevGuard.moveTo(0,0);
		winPrevGuard.focus();		
	}
</script>			
<FORM name="frmMain" method="post">
<div class='svplain10'>
This page allows FPCS admin staff to generate User Accounts for Guardians that
<ul>
	<li>previously had no user account AND</li>
	<li>belong to a family that has a student that recently was re-enrolled for the given school year</li>
</ul> 

Please note that this page does NOT create Guardian accounts.  It simply creates
user accounst for existing guardians that meet the criteria above.  Furthermore, it resets the
password for ALL user accounts that have NEVER logged in.<br><br>
When a Student wins the Lottery, the guardians associated with that
student still have to be created manually, especially in light of the
fact that the guardian might already exist for a previous student<br><br>
If you would like to preview the list of guardians for which accounts 
will be created, supply the School Year and click Preview.  If you are ready to create the accounts,
click OK.

<BR><br>
School Year <select name="selSchoolYear">
	<% = oFunc.MakeYearList((year(now) - cint(application.Value("dtYearAppStarted"))+3),1,lotteryYear) %>
</select>
<INPUT type="submit" class="btSmallGray" name="cmdOK" value="OK">&nbsp;
<input type=button class="btSmallGray" value="Preview" onclick="jfPreview();"> 
</div>
</form>
<%
	set oFunc = nothing
'end if
Server.Execute(Application.Value("strWebRoot") & "Includes/footer.asp")

%>