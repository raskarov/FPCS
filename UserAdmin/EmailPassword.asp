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

	strTitle = "FPCS.org - Email My Password"
	strUserID = Request.QueryString("szUserID")
	
	strMsgUID = "<tr><td><table><tr>" & vbCrLf & _
					"<td>" & vbCrLf & _
					"	<font size='-1' face=tahoma><b>User ID:</b></font>" & vbCrLf & _
					"</td>" & vbCrLf & _
					"<td>" & vbCrLf & _
					"	<input type='text' size='20' name='szUserID' value='" & strUserID & "'>" & vbCrLf & _
					"	<input type='hidden' name='hThisForm' value='TRUE'>" & vbCrLf & _
					"</td>" & vbCrLf & _
					"<td>" & vbCrLf & _
					"	<input type='submit' name='submit' value='Email Password'>" & vbCrLf & _
					"</td>" & vbCrLf & _
				"</tr></table></td></tr>"
	
	
	if Request.QueryString("hThisForm") = "TRUE" then
	
		set rsEmailValid = Server.CreateObject("ADODB.RecordSet")
		rsEmailValid.CursorLocation = 3 'adUseClient 

		strSQLvalidate =	"SELECT szUser_ID, szEmail, szPassword, szName_First FROM tblUsers " & _
								"WHERE blnActive = 1 AND szUser_ID = '" & trim(strUserID) &  "'"

		
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
			strBody =	"This is an automatic email notification from Family Partnership Charter School." & vbCrLf & _
						vbCrLf & "You recently requested your password be sent to you via email." & _
						vbCrLf & vbCrLf	& "Your User ID is: " & strUserID & vbCrLf & _
						"Your password is: " & strPassword & vbCrLf & vbCrLf & _
						"If you need additional help please check our web site at http://" & _
						Request.ServerVariables("SERVER_NAME") & Application.Value("strWebRoot")
						
			Set cdoMessage = Server.CreateObject("CDO.Message") 
			set cdoConfig = Server.CreateObject("CDO.Configuration")
			cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 ' cdoSendUsingPort
			cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
			cdoConfig.Fields.Update
			set cdoMessage.Configuration = cdoConfig
			cdoMessage.From = "OFFICE@FPCS.NET" 
			cdoMessage.To = rsEmailValid("szEmail") 
			cdoMessage.Subject = "FPCS Account Request" 
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
		call oFunc.CloseCN()
		set oFunc = nothing
	else
		strMsg = "<TR><TD>Please enter your registered User ID. <BR>" & _
					"We will email your password to the address on file. </TD></TR> " & strMsgUID
	end if
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