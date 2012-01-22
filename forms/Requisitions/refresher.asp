<%@ Language=VBScript %>
<% 
	if session.Contents(session.contents("strUserId") & "StartTime") = "" then
		session.Contents(session.contents("strUserId") & "StartTime") = dateadd("h",-3,now())
	end if 
%>
<html>  
  <head>
    <title>refresher</title>
    <link rel="stylesheet" type="text/css" href="<% = Application.Value("strWebRoot") %>css/homestyle.css">
    <META HTTP-EQUIV=Refresh CONTENT="59; URL=<% = Application("strSSLWebRoot") & "forms/requisitions/refresher.asp" %>"> 
  </head>
  <body class="svplain8">
		<span style="font-size=10pt;"><b>Session Helper</b></span><br><br>
		Keep this page open while working on the Goods and Service Approval Admin. This
		page will help keep your session alive which will keep you from automatically 
		being logged off due to long periods of inactivity.<br><br>
		You are logged in as: <% = ucase(session.contents("strUserId")) %><br>
		Session Helper Start: <% = session.Contents(session.contents("strUserId") & "StartTime") %>
		Last Refreshed: <% = dateadd("h",-3,Now()) %>
  </body>
</html>
