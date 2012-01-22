<%@ Language=VBScript %>
<%
'**********************************************************************
'Name:      simpleHeader.asp
'Purpose:   HTML header. This include provides the opening tags. 
'
'Note**:	This page is  called indirectly by way of Server.Execute.  Thus,
'           all variables used here are not available to other pages.  Use
'           Session and Applicaiton variables instead.
'
'Author:    Scott Bacon, ThreeShapes LLC
'Date:      04/24/2002
'**********************************************************************
' No seeing webpages unless your are in SSL mode (should be port 443)
'if Request.ServerVariables("SERVER_PORT_SECURE") <> 1 then
'	Response.Redirect "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
'end if

' No seeing webpages until you log in
if Session("bolUserLoggedIn") = false and inStr(1,Request.ServerVariables("URL"),"EmailPassword.asp") < 1 _
	and inStr(1,Request.ServerVariables("URL"),"login.asp") < 1 then
	'session.Value("strURL") = "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
	'Response.Redirect(Application("strSSLWebRoot") & "UserAdmin/login.asp")
	Session.Value("strURL") = Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
	Response.Redirect(Application("strWebRoot") & "UserAdmin/login.asp")
end if 

' No seeing students info that doesn't belong to you
if ucase(session.Contents("strRole")) <> "ADMIN" then
	if request("intStudent_ID") <> "" then
		if instr(1,session.Contents("student_list"),"~" & request("intStudent_ID") & "~") < 1 then
			
			dim cnCrack
			dim insertCrack
			set cnCrack = server.CreateObject("ADODB.CONNECTION")
			cnCrack.Open(Application("cnnFPCS"))
			insertCrack = "insert into tblCrack_Attempts(szUser_ID,szIP_Address,szURL,szALL,szUser_Create) " & _
						" values (" & _
						"'" & session.Contents("strUserID") & "'," & _
						"'" & request.ServerVariables("REMOTE_ADDR") & "'," & _
						"'" & replace(Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING"),"'","''") & "'," & _
						"'" & replace(request.ServerVariables("ALL_HTTP"),"'","''") & "'," & _
						"'" & session.Contents("strUserID") & "')" 
			cnCrack.Execute(insertCrack)
			cnCrack.Close
			set cnCrack = nothing
			response.Write "<H1>Security Breach Attempt. Your user information has been captured and sent to the Admin.</h1>" & _
						   "<br>This action can cause your account to be terminated and legal steps may also be taken."
			response.End
		end if
	end if
end if

dim myGuid

myGuid = createGuid

if len(myGuid) > 1 then
	if Session.Contents("instruct_id") <> "" then
		tId = Session.Contents("instruct_id")
	else
		tId = " NULL "
	end if

	if Session.Contents("intGuardian_ID") <> "" then
		gId = Session.Contents("intGuardian_ID") 
	else
		gId = " NULL "
	end if

	set cnGuid = server.CreateObject("ADODB.CONNECTION")
	cnGuid.Open(Application("cnnFPCS"))
	insertSession = "INSERT INTO SESSION_BRIDGE " & _
			   " (SESSION_ID, USER_ID, ROLE, STUDENT_LIST, SCHOOL_YEAR, TEACHER_ID, GUARDIAN_ID, FULL_NAME, DATE_CREATED) " & _
			   " VALUES     ('" & myGuid & "','" & session.Contents("strUserID") & "','" & session.Contents("strRole") & "','" & session.Contents("student_list") & "'," & _
			   session.Contents("intSchool_Year") & "," & _
			   tId  & "," & gId & ",'',current_timestamp)" 
		'response.write 	   insertSession

	cnGuid.Execute(insertSession)
	cnGuid.Close
	set cnGuid = nothing
	myGuid = left(myGuid, len(myGuid)-1)
	myGuid = right(myGuid, len(myGuid)-1)

	sPage = replace(request.querystring("page"),"|||","?")
	sPage = replace(sPage,"~~","&")
	IF instr(1,sPage,"?") < 1 then
		sPage = sPage & "?"	
	End IF


	response.redirect("http://empowernet/" & sPage & "&GUID=" & lcase(myGuid) )
	
end if

Function createGuid()
	Set TypeLib = Server.CreateObject("Scriptlet.TypeLib")
	tg = TypeLib.Guid
	createGuid = left(tg, len(tg)-2)
	Set TypeLib = Nothing
End Function

%>