<%@ Language=VBScript %>
<%
'*******************************************
'Name:		Admin\debugEmailer.asp
'Purpose:	Emails 3Shape developers debug/error messages
'
'Author:	Bryan K Mofley (ThreeShapes.com LLC)
'Date:		17 May 2003
'*******************************************
on error resume next
dim strPathFile
dim strTextBody

	strPathFile = Server.MapPath(Request.ServerVariables("PATH_INFO"))

	strTextBody = "QueryString Collection:" & vbCrLf
	For Each Item in Request.QueryString
		For intLoop = 1 to Request.QueryString(Item).Count 
			strTextBody = strTextBody & Item & " = " & Request.QueryString(Item)(intLoop) & vbCrLf
		Next 
	Next
	
	strTextBody = strTextBody & vbCrLf & vbCrLf & "Form Collection:" & vbCrLf
	For Each Item in Request.Form
		For intLoop = 1 to Request.Form(Item).Count 
			strTextBody = strTextBody & Item & " = " & Request.Form(Item)(intLoop) & vbcrlf
		Next 
	Next 
	
	strTextBody = strTextBody & vbCrLf & vbCrLf & "Cookies Collection:" & vbCrLf
	For Each Item in Request.Cookies
		If Request.Cookies(Item).HasKeys Then
			'use another For...Each to iterate all keys of dictionary
			For Each ItemKey in Request.Cookies(Item) 
				strTextBody = strTextBody & "Sub Item: " & Item & "(" & ItemKey  & ")" & Request.Cookies(Item)(ItemKey) & vbCrLf
			Next 
		Else
			'Print out the cookie string as normal
			strTextBody = strTextBody & Item & " = " & Request.Cookies(Item) & vbCrLf
		End If
	Next 
	
	strTextBody = strTextBody & vbCrLf & vbCrLf & "ClientCertificate Collection:" & vbCrLf
	For Each Item in Request.ClientCertificate
		For intLoop = 1 to Request.ClientCertificate(Item).Count 
			strTextBody = strTextBody &  Item & " = " & Request.ClientCertificate(Item)(intLoop) & vbCrLf
		Next 
	Next 
	
	strTextBody = strTextBody & vbCrLf & vbCrLf & "ServerVariables Collection:" & vbCrLf
	For Each Item in Request.ServerVariables
		For intLoop = 1 to Request.ServerVariables(Item).Count 
			strTextBody = strTextBody &  Item & " = " & Request.ServerVariables(Item)(intLoop) & vbCrLf
		Next 
	Next 

	' Set up CDO object and set properties
	'http://msdn.microsoft.com/library/en-us/cdosys/html/_cdosys_messaging_examples_creating_and_sending_a_message.asp?frame=true
	Set cdoMessage = Server.CreateObject("CDO.Message")
	set cdoConfig = Server.CreateObject("CDO.Configuration")
	cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 ' cdoSendUsingPort
	cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
	cdoConfig.Fields.Update
	set cdoMessage.Configuration = cdoConfig

	cdoMessage.From = "3Shapes Autobot <dev@3Shapes.com>" 
	cdoMessage.Subject = Request.ServerVariables("SERVER_NAME") & " Application Error"
	cdoMessage.TextBody = "An error occured while processing the following file:" & vbCrLf & _
		vbTab &  strPathFile & vbCrLf & vbCrLf & _
		"Error:" & Session.Contents("ErrorNum") & vbCrLf & "Description:" & Session.Contents("ErrorDesc") & _
		"----------------------------------------" & vbCrLf & vbCrLf & strTextBody
	cdoMessage.To = "3Shapes Debugger <debug@3Shapes.com>"

	cdoMessage.Send 

	Set cdoMessage = Nothing 
%>