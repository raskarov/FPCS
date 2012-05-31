<%@ Language=VBScript %>
<%
'*******************************************
'Name:		Admin\fixBudget.asp
'Purpose:	Finds all goods/services that exist in Order Items but not
'			in the budget table and creates budget records to balance
'			actual expense to budgeted.
'
'Author:	ThreeShapes.com LLC
'Date:		02 Oct 2003
'*******************************************
DO NOT USE
'per http://support.microsoft.com/default.aspx?scid=kb;EN-US;q234067
Response.CacheControl = "no-cache" 
Response.Expires = -1

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

if Ucase(session.Contents("strRole")) <> "ADMIN" then
	response.Write "<h1>You are not authourized to view this page.</h1>"
	response.End
end if

sql = "SELECT oi.intItem_ID, oi.intOrdered_Item_ID, oi.intQty," & _
		" oi.curUnit_Price, oi.intStudent_ID, " & _
		"i.intShort_ILP_ID, ti.intItem_Group_ID, tblOrd_Attrib.szValue " & _
		"FROM tblOrdered_Items oi INNER JOIN " & _
		" tblILP i ON oi.intILP_ID = i.intILP_ID AND oi.intILP_ID = i.intILP_ID INNER JOIN " & _
		" trefItems ti ON oi.intItem_ID = ti.intItem_ID INNER JOIN " & _
		" tblOrd_Attrib ON oi.intOrdered_Item_ID = tblOrd_Attrib.intOrdered_Item_ID " & _
		"WHERE (oi.intSchool_Year = 2004) AND (NOT EXISTS " & _
		" (SELECT 'x' " & _
		" FROM tblILP_SHORT_FORM sf INNER JOIN " & _
		" tblBudget b ON sf.intShort_ILP_ID = b.intShort_ILP_ID " & _
		" WHERE b.intOrdered_Item_ID = oi.intOrdered_Item_ID  " & _
		"AND (sf.intSchool_Year = 2004))) AND (tblOrd_Attrib.intOrder = 1) " & _
		"ORDER BY oi.intStudent_ID"
set rs = server.CreateObject("ADODB.RECORDSET")
rs.CursorLocation = 3
rs.Open sql, Application("cnnFPCS")'oFunc.FPCScnn

if rs.RecordCount > 0 then
	do while not rs.EOF
		if rs("intShort_ILP_ID") & "" <> "" then
			insert = "insert into tblBudget(" & _
					"intBudget_Item_ID,intItem_ID,intShort_ILP_ID,szDesc,intQTY,curUnit_Price," & _
					"intOrdered_Item_ID,dtCreate,szUser_Create) values (" & _
					rs("intItem_Group_ID") & "," & rs("intItem_ID") & "," & _
					rs("intShort_ILP_ID") & ",'" & replace(rs("szValue"),"'","''") & "'," & _
					rs("intQTY") & "," & replace(rs("curUnit_Price"),",","") & "," & rs("intOrdered_Item_ID") & ",'" & _
					now() & "','fixBudget')"
			response.Write insert & "<BR><BR>"
			'oFunc.ExecuteCN(insert)		
		end if		 		
		rs.MoveNext
	loop
end if

response.Write "Number of records corrected: " & rs.RecordCount
rs.Close
set rs = nothing
oFunc.CloseCN
set oFunc = nothing
%>