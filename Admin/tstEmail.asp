<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D" NAME="CDO for Windows 2000 Type Library" -->
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library" -->
<html>
<body>
<%
Dim oMessage, oConfig
Dim strTo, strFrom, strSubject, strBody

Set oMessage = CreateObject("CDO.Message")
Set oConfig= CreateObject("CDO.Configuration")

oConfig.Fields(cdoSendUsingMethod) = cdoSendUsingPort
oConfig.Fields(cdoSMTPServer) = "127.0.0.1"
oConfig.Fields(cdoSMTPServerPort) = 25
oConfig.Fields(cdoSMTPConnectionTimeout) = 10
oConfig.Fields(cdoSMTPAuthenticate) = cdoAnonymous
oConfig.Fields.Update

strTo = "bryan@3shapes.com,scott@3shapes.com"
strFrom = "bryan@threeshapes.com"
strSubject = "Subject: Test from ASP"
strBody = "Message text  from ASP"

Set oMessage.Configuration = oConfig 

oMessage.To = strTo ' recipient
oMessage.From = strFrom ' sender
oMessage.Subject = strSubject ' Subject
oMessage.TextBody = strBody ' Message

On Error Resume Next
oMessage.Send

If Err.Number = 0 Then
    Response.Write("Message sent successfully!")
Else
    Response.Write("ERROR! " & Err.description)
    Err.Clear
End If
On Error Goto 0
%>
</body>
</html>
