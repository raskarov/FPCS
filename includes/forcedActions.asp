<%
function vbfActionHandling(oFPCS,fvCount)
	oFPCS.BeginTransCN	
	dim i
	dim insert
	dim delete
	
	insert = "insert into tblForce_Action_Results(intAction_ID,szUser_ID,dtCompleted,intSchool_Year) " & _
			 "values (" & session.Contents("arActions")(fvCount,0) & _
			 ",'" & session.Contents("strUserID") & "','" & now() & "'," & _
			 session.Contents("intSchool_Year") & ")"
	oFPCS.ExecuteCN(insert)

	' Delete this forced action from users list
	delete = "delete from tascUsers_Action " & _
			 "where intUser_Action_ID = " & session.Contents("arActions")(fvCount,2)

	oFPCS.ExecuteCN(delete)
	
	oFPCS.CommitTransCN
	oFPCS.CloseCN
	
	'Erase this action from the array so it doesn't get executed again
	arEdit = session.Contents("arActions")
	arEdit(fvCount,1) = "" 
	session.Contents("arActions") = arEdit
	
	for i = 0 to ubound(arEdit)
		if arEdit(i,1) <> "" then
			Response.Redirect(Application.Contents("strSSLWebRoot") & arEdit(i,1) & "&exempt=true")
			response.end
		end if 
	next 

	'set oFunc = nothing  'THIS SHOULD PROBABLY BE UNCOMMENTED SINCE oFunc wil not be destroyed otherwise
	session.Contents("bolActionNeeded") = false
	if session.Contents("strURL") <> "" then
		Response.Redirect(session.Contents("strURL"))
	else
		Response.Redirect(Application.Contents("strSSLWebRoot"))
	end if 
	Response.End
end function

%>

