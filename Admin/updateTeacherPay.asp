<%@ Language=VBScript %>
<%
'*******************************************
'Name:		Admin\updateTeacherPay.asp
'Purpose:	Transforms legacy active/inactive tracking to 
'			new historical teacher payroll tracking of active status
'
'Author:	ThreeShapes.com LLC
'Date:		10 May 2005
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


' get all inactive teachers
sql = "SELECT tblINSTRUCTOR.dtInactive, tblINSTRUCTOR.szLAST_NAME, tblINSTRUCTOR.szFIRST_NAME, tblInstructor_Pay_Data.curPer_Hour,  " & _ 
		" tblInstructor_Pay_Data.intPay_Type_id, tblInstructor_Pay_Data.dtEffective_Start, tblInstructor_Pay_Data.dtEffective_End,  " & _ 
		" tblInstructor_Pay_Data.bolActive, tblInstructor_Pay_Data.intInstructor_Pay_Data_ID, tblInstructor_Pay_Data.intInstructor_ID,  " & _ 
		" tblInstructor_Pay_Data.curPer_Hour_Benefits, tblInstructor_Pay_Data.curPay_Rate, tblInstructor_Pay_Data.bolASD_Full_Time,  " & _ 
		" tblInstructor_Pay_Data.decASD_Full_Time_Percent, tblInstructor_Pay_Data.bolASD_Part_Time,  " & _ 
		"tblInstructor_Pay_Data.decASD_Part_Time_Percent,  tblInstructor_Pay_Data.decFPCS_Hours_Goal " & _ 
		"FROM tblINSTRUCTOR INNER JOIN " & _ 
		" tblInstructor_Pay_Data ON tblINSTRUCTOR.intINSTRUCTOR_ID = tblInstructor_Pay_Data.intInstructor_ID " & _ 
		"WHERE (tblINSTRUCTOR.bolIs_Active = 0) AND (tblINSTRUCTOR.dtInactive > '1/1/2004') AND (tblINSTRUCTOR.dtInactive < '3/3/2005') " & _ 
		"ORDER BY tblINSTRUCTOR.szLAST_NAME, tblINSTRUCTOR.szFIRST_NAME, tblInstructor_Pay_Data.dtEffective_Start DESC, tblInstructor_Pay_Data.intInstructor_Pay_Data_ID DESC "
		
set rs = server.CreateObject("ADODB.RECORDSET")
rs.CursorLocation = 3
rs.Open sql, Application("cnnFPCS")'oFunc.FPCScnn

if rs.RecordCount > 0 then
	id = 0
	oFunc.BeginTransCN
	do while not rs.EOF
		if id <> rs("intINSTRUCTOR_ID") then
			' as the teacher id changes we know we need to insert a new 
			' pay record to show that the teacher is now inactive
			insert = "INSERT INTO tblInstructor_Pay_Data " & _ 
					" (intInstructor_ID, curPer_Hour, curPer_Hour_Benefits, curPay_Rate, intPay_Type_id,  " & _ 
					"bolASD_Full_Time, decASD_Full_Time_Percent,  " & _ 
					" bolASD_Part_Time, decASD_Part_Time_Percent, decFPCS_Hours_Goal,  " & _ 
					"dtEffective_Start, bolActive, dtCREATE, szUSER_CREATE) " & _ 
					"VALUES(" & _
					rs("intInstructor_ID") & "," & _
					rs("curPer_Hour") & "," & _
					rs("curPer_Hour_Benefits") & "," & _
					rs("curPay_Rate") & "," & _
					rs("intPay_Type_id") & "," & _
					oFunc.TrueFalse(rs("bolASD_Full_Time")) & "," & _
					rs("decASD_Full_Time_Percent") & "," & _
					oFunc.TrueFalse(rs("bolASD_Part_Time")) & "," & _
					rs("decASD_Part_Time_Percent") & "," & _
					rs("decFPCS_Hours_Goal") & "," & _
					"'7/1/2004'," & _
					"0," & _
					"CURRENT_TIMESTAMP,'AUTO_SCRIPT') "
					'response.Write oFunc.TrueFalse(rs("bolASD_Full_Time")) & "<BR><BR>"
					'response.Write insert
			oFunc.ExecuteCN(insert)
			id = rs("intINSTRUCTOR_ID")
		end if
		
		' Add an end date to any existing record that has not been closed
		if rs("dtEffective_End") & "" = "" then
			strAdd = " , dtEffective_End = '7/1/2004' "
		end if
		
		' make teachers active in the database for prior records
		update = "update tblInstructor_Pay_Data set bolActive = 1 " & _
				 strAdd & " WHERE intInstructor_Pay_Data_ID = " & rs("intInstructor_Pay_Data_ID") 
		oFunc.ExecuteCN(update)				 
		rs.MoveNext
	loop
	oFunc.CommitTransCN
end if

response.Write "Number of records corrected: " & rs.RecordCount
rs.Close
set rs = nothing
oFunc.CloseCN
set oFunc = nothing
%>