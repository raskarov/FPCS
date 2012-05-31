<%@ Language=VBScript %>
<%
'*******************************************
'Name:		Admin\PreviewGuardAccts.asp
'Purpose:	Allows FPCS staff to preview the User 
'			accounts for guardians that
'			1) previously had no user account AND
'			2) belong to a family that has a student that
'				recently was re-enrolled for the given school year
'
'CalledBy:	Admin\GenerateGuardAccts.asp
'
'Author:	ThreeShapes.com LLC
'Date:		12 June 2003
'*******************************************
'option explicit
dim oFunc
dim intRow
dim rs
dim strSQLpreview
dim strBGcolor

if Session.Value("strRole") <> "ADMIN" then 
	response.Write "We're sorry, but only authorized staff members may access this page."
	response.End
end if
iCounter = 0
Session.Value("strTitle") = "Preview Guardian User Accounts"
Server.Execute(Application.Value("strWebRoot") & "Includes/simpleHeader.asp")


set rs = Server.CreateObject("ADODB.Recordset")
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

%>
<div class='svplain10'>
<%
strSQLpreview = "SELECT DISTINCT " & _
                "      tblGUARDIAN.intGUARDIAN_ID, tblGUARDIAN.szLAST_NAME, tblGUARDIAN.szFIRST_NAME, tblGUARDIAN.sMID_INITIAL, tblGUARDIAN.szEMAIL, " & _
                "      UPPER(LEFT(tblGUARDIAN.szLAST_NAME, 6) + LEFT(tblGUARDIAN.szFIRST_NAME, 1) + CAST(tblGUARDIAN.intGUARDIAN_ID AS varchar(5))) AS UserNameWillBe " & _
				"FROM  tascGUARD_USERS RIGHT OUTER JOIN " & _
                "      tascFAM_GUARD INNER JOIN " & _
                "      tblGUARDIAN ON tascFAM_GUARD.intGUARDIAN_ID = tblGUARDIAN.intGUARDIAN_ID INNER JOIN " & _
                "      tblSTUDENT INNER JOIN " & _
                "      tblStudent_States ON tblSTUDENT.intSTUDENT_ID = tblStudent_States.intStudent_id ON tascFAM_GUARD.intFamily_ID = tblSTUDENT.intFamily_ID ON  " & _
                "      tascGUARD_USERS.intGUARDIAN_ID = tblGUARDIAN.intGUARDIAN_ID " & _
				"WHERE (tascGUARD_USERS.intGuard_User_ID IS NULL) AND tblStudent_States.intReEnroll_State IN (" & Application.Contents("ActiveEnrollList") & ")  AND (tblStudent_States.intSchool_Year = " & Request("selSchoolYear") & ") " & _
                "ORDER BY tblGUARDIAN.szLAST_NAME"


	rs.CursorLocation = 3
	rs.Open strSQLpreview, Application("cnnFPCS")'oFunc.FPCScnn
	if not rs.BOF and not rs.EOF then
%>
<table border=1 cellspacing=0 cellpadding=1 ID="Table1">
<thead>
	<tr bgcolor='#639976' style='COLOR: WHITE;'>
<%
		for each fld in rs.fields
			Response.Write "<td style='FONT-SIZE:8pt'>" & fld.name & "</td>" & vbCrLf
		next
%>		
	</tr>
</thead>
<%
		intRow = 0
		do until rs.EOF
			if intRow mod 2 = 0 then
				strBGcolor = " bgcolor=white"
			else
				strBGcolor = " bgcolor=silver"
			end if
			Response.Write "<tr" & strBGcolor & ">" & vbCrLf
			for each fld in rs.fields
				dim vntValue
				if trim(fld.value) = "" then
					vntValue = "&nbsp;"	'pad as space so TD cell will show borders correclty
				else
					vntValue = fld.value
				end if
				Response.Write "<td style='FONT-SIZE:7pt'>" & vntValue & "</td>" & vbCrLf
			next
			Response.Write "</tr>" & vbCrLf
			intRow = intRow + 1
			rs.movenext
		loop
%>	
</table>
<%
	else
		response.Write "There are no Guardians for School Year " & Request("selSchoolYear") & " that require User Accounts.<BR>"	
	end if
%>
<input type=button value="Close" onclick="window.close();" id="btSmallGray" NAME="btSmallGray">
<%	
set rs = nothing
set oFunc = nothing
%>
</div>
</body>
</html>
