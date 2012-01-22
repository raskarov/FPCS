<%@ Language=VBScript %>
<%
'*******************************************
'Name:		admin\maint.asp
'Purpose:	Toggles the bolMaintenance Session
'				variable.  This session variable is 
'				evaluted by Includes\header.asp.  If
'				set to TRUE, then a "Site unavailable"
'				message is displayed
'
'CalledBy:	Guy's FTP SandX program	
'Calls:		
'
'Inputs:		Request.QueryString("maint")
'
'Author:		Bryan K Mofley
'				3Shapes is Scott Bacon, Bryan Mofley, Guy Mofley
'Date:      07-Jun-2000
'*******************************************

option explicit
dim strMaint		'from Request.QueryString

strMaint = Request.QueryString("maint")

if lcase(strMaint) = "true" then
	Application("bolMaintenance") = True
else
	Application("bolMaintenance") = False
end if
%>