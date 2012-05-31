<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		guardianInsert.asp
'Purpose:	This script inserts or updates all guardian inforamation
'			recieved from guadianProfile.asp
'Date:		9 July 2001
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim oFunc
dim strMessage		' This string will be printed at the bottom of default.asp 
dim dtBirth

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

if Request.Form("intGuardian_ID") = "" and Request.form("changed") <> "" then
	'If we get here we need to create a new record in the database.	
	dim insert
	dim intGuardian_ID
	oFunc.BeginTransCN		
	
	' Create a new Gaurdian Record
	insert = "insert into tblGuardian (" & _
			 "szLast_Name,szFirst_Name,sMid_Initial, szEmployer, szBusiness_Phone, " & _
			 "intPhone_Ext, szCell_Phone, szPager,bolActive_Military,szRank, " & _
			 "szEmail,szAddress,szCity,szState,szCountry,szZip_Code,szHome_Phone,szUSER_CREATE) VALUES ('" & _
			 oFunc.EscapeTick(Request.Form("szLast_Name")) & "','" & _
			 oFunc.EscapeTick(Request.Form("szFirst_Name")) & "','" & _ 
			 oFunc.EscapeTick(Request.Form("sMid_Initial")) & "','" & _
			 oFunc.EscapeTick(Request.Form("szEmployer")) & "','" & _ 
			 oFunc.EscapeTick(Request.Form("szBusiness_Phone")) & "'," & _
			 oFunc.CheckDecimal(Request.Form("intPhone_Ext")) & ",'" & _
			 oFunc.EscapeTick(Request.Form("szCell_Phone")) & "','" & _
			 oFunc.EscapeTick(Request.Form("szPager")) & "'," & _
			 oFunc.TrueFalse(Request.Form("bolActive_Military")) & ",'" & _
			 oFunc.EscapeTick(Request.Form("szRank")) & "','" & _
			 oFunc.EscapeTick(Request.Form("szEmail")) & "','" & _
			 oFunc.EscapeTick(Request.Form("szAddress")) & "','" & _
			 oFunc.EscapeTick(Request.Form("szCity")) & "','" & _
			 oFunc.EscapeTick(Request.Form("szState")) & "','" & _
			 oFunc.EscapeTick(Request.Form("szCountry")) & "','" & _
			 oFunc.EscapeTick(Request.Form("szZip_Code")) & "','" & _
			 oFunc.EscapeTick(Request.Form("szHome_Phone")) & "','" & _
			 Session.Value("strUserID")	& "')"

	oFunc.ExecuteCN(insert)
	
	'Create a new tascStudent_Guardian record
	intGuardian_ID = oFunc.GetIdentity
	call vbfInsertGaurdStudent(intGuardian_ID)
	
	oFunc.CommitTransCN		 
	strMessage = "Guardian Information was Added."
	
	if Request.Form("bolNewGuardian") <> "" then 
		' We add the new Guardian to the Guardian list on the family manager page and close this window.
%>
<html>
<head>
<script language=javascript>
	function jfAddGuardianToList(){
		// Passes info needed to create a new option in the Guardian select list that is
		// contained in familyManager.asp 
		window.opener.jfAddOption('<% = Request.Form("szLast_Name")& "," & Request.Form("szFirst_Name")%>','<%=intGuardian_ID%>','selGuardian_ID');
		window.opener.focus();
		window.close();
	}

</script>
<body onload="jfAddGuardianToList();" bgcolor=white>
</body>
</html>
<%	
		Response.End
	end if 
elseif  Request.Form("changed") <> ""  and Request.Form("intGuardian_ID") <> ""  then
	' A change has been made in edit mode so we update the data.
	oFunc.BeginTransCN
	dim update
	dim sql
	
	update = "update tblGuardian set " & _
			 "szLast_Name = '" & oFunc.EscapeTick(Request.Form("szLast_Name")) & "'," & _
			 "szFirst_Name = '" &oFunc.EscapeTick(Request.Form("szFirst_Name")) & "'," & _ 
			 "sMid_Initial = '" & oFunc.EscapeTick(Request.Form("sMid_Initial")) & "'," & _
			 "szEmployer = '" & oFunc.EscapeTick(Request.Form("szEmployer")) & "'," & _ 
			 "szBusiness_Phone = '" & oFunc.EscapeTick(Request.Form("szBusiness_Phone")) & "'," & _
			 "intPhone_Ext = " & oFunc.CheckDecimal(Request.Form("intPhone_Ext")) & "," & _
			 "szCell_Phone = '" & oFunc.EscapeTick(Request.Form("szCell_Phone")) & "'," & _
			 "szPager = '" & oFunc.EscapeTick(Request.Form("szPager")) & "'," & _
			 "bolActive_Military = " & oFunc.TrueFalse(Request.Form("bolActive_Military")) & "," & _
			 "szRank = '" & oFunc.EscapeTick(Request.Form("szRank")) & "'," & _
			 "szEmail = '" & oFunc.EscapeTick(Request.Form("szEmail")) & "'," & _
			 "szAddress = '" & oFunc.EscapeTick(Request.Form("szAddress")) & "'," & _
			 "szCity = '" & oFunc.EscapeTick(Request.Form("szCity")) & "'," & _
			 "szState = '" & oFunc.EscapeTick(Request.Form("szState")) & "'," & _
			 "szCountry = '" & oFunc.EscapeTick(Request.Form("szCountry")) & "'," & _
			 "szZip_Code = '" & oFunc.EscapeTick(Request.Form("szZip_Code")) & "'," & _
			 "szHome_Phone = '" & oFunc.EscapeTick(Request.Form("szHome_Phone")) & "'," & _
			 "szUser_Modify = '" & Session.Value("strUserID")& "' " & _
			 "where intGuardian_ID = " & Request.Form("intGuardian_ID")
	oFunc.ExecuteCN(update)
	
	' Delete existing student-guardian records for this family
	set rsAssoc = server.CreateObject("ADODB.RECORDSET")
	rsAssoc.CursorLocation = 3
	sql = "select intAssoc_ID from tascStudent_Guardian sg, tblStudent s " & _
		  "where sg.intStudent_ID = s.intStudent_ID " & _
		  "and s.intFamily_id = " & Request.Form("intFamily_id") & _
		  " and intGuardian_id = " & Request.Form("intGuardian_ID")
	rsAssoc.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
	
	do while not rsAssoc.EOF
		delete = "delete from tascStudent_Guardian where intAssoc_ID = " & rsAssoc("intAssoc_ID")
		oFunc.ExecuteCN(delete)
		rsAssoc.MoveNext
	loop
	
	rsAssoc.Close
	set rsAssoc = nothing
	
	'create new student-guardian records for this family
	call vbfInsertGaurdStudent(Request.Form("intGuardian_ID"))
	
	oFunc.CommitTransCN
	strMessage = "Guardian Information was Updated."
end if

function vbfInsertGaurdStudent(id)
	arStudents = split(Request.Form("strStudents"),",")
	if isArray(arStudents) then
		for i = 0 to ubound(arStudents)
			if arStudents(i) <> "" and Request.Form("intGuardian_Type_ID"&arStudents(i)) <> "" then
				insert ="insert into tascStudent_Guardian(intStudent_Id,intGuardian_ID,intGuardian_Type_ID," & _
						"dtCreate,szUser_Create) values(" & _
						arStudents(i) & "," & _
						id & "," & _
						Request.Form("intGuardian_Type_ID"&arStudents(i)) & "," & _
						"'" & now() & "','" & _
						Session.Value("strUserID")	& "')"
				oFunc.ExecuteCN(insert)
			end if 
		next
	end if 
end function
call oFunc.CloseCN
%>
<html>
<head>
<script language=javascript>
	function jfClose(){
		alert("Guardian info has been updated.");
		window.close();
	}

</script>
<body onload="jfClose();" bgcolor=white>
</body>
</html>

