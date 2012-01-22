<%@ Language=VBScript %>
<%
'*******************************************
'Name:		UserAdmin\EmailPassword.asp
'Purpose:	Emails the users password
'
'CalledBy:	UserAdmin\login.asp	
'
'Inputs:	Request.QueryString("szUserID")
'
'Author:	3Shapes (Scott Bacon, Bryan Mofley, Guy Mofley)
'Date:      19-May-2000
'*******************************************

'option explicit
dim strTitle
dim strMsg
dim strEmail
dim strUserID
dim strPassword
dim strMDBEnc
dim rsEmailValid
dim strSQLvalidate
dim objMail
dim strBody
dim strMsgUID
dim oCrypto
dim oFunc

response.end

	strTitle = "FPCS.org - Email My Password"

	
		set rsEmailValid = Server.CreateObject("ADODB.RecordSet")
		rsEmailValid.CursorLocation = 3 'adUseClient 

		strSQLvalidate =	" SELECT replace(szUser_ID,' ', '') as szUser_ID, replace(szEmail,' ', '') as szEmail, szPassword, szName_First " & _
					" FROM tblUsers U" & _
					" INNER JOIN TEMP_GUARD_USER T ON T.NEW_USER = SZUSER_ID " & _
					" where szUser_Id > 'Bredar_William'"

		
		set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
		call oFunc.OpenCN()
		rsEmailValid.Open strSQLvalidate, oFunc.FPCScnn

		if rsEmailValid.BOF and rsEmailValid.EOF then
			strMsg = "<TR><TD colspan=3>Unable to find the User ID you specified.<BR> " & _
						"Make sure you are using the User ID you <BR> " & _
						"used when you signed up.<BR><BR> " & _
						"If you think this is an error, please contact<BR>" & _
						"the FPCS Admin Staff.</TD></TR> " & strMsgUID
		else 
			'response.write rsEmailValid.RecordCount & "<BR><BR>"
			do while NOT rsEmailValid.EOF

			if not instr(rsEmailValid("szUser_id"),".") > 0 and not instr(rsEmailValid("szUser_id"),"'") > 0 then
			set oCrypto = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/Crypto.wsc"))
				'encrypt password for database compare
				'oCrypto.Key = "something"	'actual key is not shown here
				oCrypto.Text = rsEmailValid("szPassword")
				Call oCrypto.Encypttext
				strPassword = oCrypto.EncryptedText
			set oCrypto = nothing
		
			'*******************************************
			'Name:		Send Mail
			'Purpose:	Send an email to the registered user
			'http://msdn.microsoft.com/library/en-us/cdosys/html/_cdosys_messaging_examples_creating_and_sending_a_message.asp?frame=true
			'*******************************************

strBody = "FAMILIES ENROLLED IN FPCS IN THE CURRENT 2009/2010 SCHOOL YEAR:" & vbCrLf & vbCrLf & _
"The APC and FPCS have decided to temporarily return to our former online system for the 2010/2011 school year. " & _
"This system is now available for you to begin creating your 2010/ 2011 school year ILP’s and budgeting for curriculum orders. " & _
"Service vendors will not be available until July 1, 2010. " & vbCrLf & vbCrLf & _
"FAMILIES NEW TO FPCS FOR THE 2010/2011 SCHOOL YEAR:" & vbCrLf & vbCrLf & _
"Please contact the office 742-3700 to set up training for the FPCS online system or contact your sponsor teacher for assistance. " & vbCrLf & vbCrLf & _
"TO LOG IN TO THE SYSTEM PLEASE GO TO ... " & vbCrLf & _
"http://www.fpcs.net " & vbCrLf & vbCrLf & _
"Near the bottom of the screen you will see the FPCS log in section. Your log in information is as follows ..." & vbCrLf & _
"User Name: " & rsEmailValid("szUser_id") & vbCrLf & vbCrLf & _
"Password: " & strPassword  & vbCrLf & vbCrLf & _
"Passwords are case sensitive, so be sure to enter exactly as you see here." & vbCrLf & vbCrLf & _
"Thank you," & vbCrLf & _
"FPCS Staff"

			RESPONSE.WRITE strBody & "<br><br>" & 	rsEmailValid("szEmail") 
			Set cdoMessage = Server.CreateObject("CDO.Message") 
			set cdoConfig = Server.CreateObject("CDO.Configuration")
			cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 ' cdoSendUsingPort
			cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
			cdoConfig.Fields.Update
			set cdoMessage.Configuration = cdoConfig
			cdoMessage.From = "FPCS_ADMIN@FPCS.NET" 
			cdoMessage.To = rsEmailValid("szEmail") 
			cdoMessage.Subject = "FPCS Account Log in for 2010-2011 School Year" 
			cdoMessage.TextBody = strBody 
			cdoMessage.Send 
			Set cdoMessage = Nothing 
			
		
			
			'***********************************************************
			'end EMAIL section
			'***********************************************************
			
			strMsg = "<td>Your password has been sent to your registered email address.<BR><BR>" & vbCrLf & _
					"Click the 'Sign In' button to log in once you receive your password.<BR>" & vbCrLf & _
					"</td>"
	end if
			rsEmailValid.MoveNext
		   loop
		end if
		call oFunc.CloseCN()
		set oFunc = nothing
	Session.Value("strTitle") = "FPCSonline - Email Password"
	Session.Value("strLastUpdate") = "22 Feb 2002"
	'Server.Execute(Application.Value("strWebRoot") & "Includes/header.asp")
%>

	<center>
		<font color="black" face="Verdana,Arial,Helvetica,Sans-serif" size="-1">
			<form action=EmailPassword.asp method=get name='frmEmailPW' ID="Form1">
				<table  ID="Table1">
					<tr>
						<%= strMsg %>	
					</tr>
				</table>
			</form>
		</font>
	</center>
<%
	'Server.Execute(Application.Value("strWebRoot") & "Includes/footer.asp")
%>