<%@ Language=VBScript %>
<%
'*******************************************
'Name:		UserAdmin\logout.asp
'Purpose:	Simply deletes the users cookie by setting the 
'			date to the past, the abondons the session so
'			as to reinitialize all session variables, and
'			finally redirects to the home page
'
'Author:	3Shapes (Scott Bacon, Bryan Mofley, Guy Mofley and others)
'Date:      23-May-2000
'*******************************************
	'Response.Cookies("remember").expires = now()-365
	Session.Abandon
	Response.Redirect(Application.Value("strSSLWebRoot"))
%>
