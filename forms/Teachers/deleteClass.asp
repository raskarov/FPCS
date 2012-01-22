<%@ Language=VBScript %>
<%
dim delete
dim sql
dim strMessage
dim oFunc
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

session.Value("simpleTitle") = "Delete Class"
Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")   

if Request.QueryString("intClass_id") <> "" then
	' This section is for delete a teachers class. 
	set rsILP = server.CreateObject("ADODB.RECORDSET")
	rsILP.CursorLocation = 3
	sql = "select * from tblILP where intClass_ID = " & Request.QueryString("intClass_id")
	rsILP.Open sql,oFunc.FPCScnn
	
	if rsILP.RecordCount > 0 then
		strMessage = "You can not delete this class because it has enrolled students. To " & _
					 "delete this class all students must be removed. " 
	else		
		rsILP.Close
		'Check to see if we need to delete any Generic ILP/Grading Scales 
		'linked to Class as long as they are not suppose to be part of the ILP Bank
		sql = "select intILP_ID,intGrading_Scale_ID " & _
			  "from tblILP_Generic " & _
			  "where intClass_ID = " & Request.QueryString("intClass_id") & _
			  " and bolILP_Bank <> 1 "
		rsILP.Open sql,oFunc.FPCScnn
		
		oFunc.BeginTransCN
		if rsILP.RecordCount > 0 then	
			delete = "delete from tblILP_Generic where intILP_ID = " & rsILP("intILP_ID")	
			oFunc.executeCN(delete)	
			if rsILP("intGrading_Scale_ID") & "" <> "" then
				delete = "delete from tblGrading_Scale_Generic where intGrading_Scale_ID = " & rsILP("intGrading_Scale_ID")
				oFunc.executeCN(delete)	
			end if								
		end if 
		
		'Delete all the records that limited the class to specific families.
		delete = "delete from tascClass_Family where intClass_id = " & Request.QueryString("intClass_id")
		oFunc.executeCN(delete)
		'Now delete the class Items and the class
		delete = "delete from tblClass_Items where intClass_id = " & Request.QueryString("intClass_id")
		oFunc.executeCN(delete)
		
		delete = "delete from tblClasses where intClass_id = " & Request.QueryString("intClass_id")
		oFunc.executeCN(delete)
		oFunc.CommitTransCN
		strMessage = "Class and materials have been deleted from the database."
	end if 
	
	rsILP.Close
	set rsILP = nothing
	
end if
if Request.QueryString("studentdrop") <> ""  then
	set rsStudents = server.CreateObject("ADODB.RECORDSET")
	rsStudents.CursorLocation = 3
	sql = "select intILP_Id from tblILP where intClass_Id = " & Request.QueryString("studentdrop")
	rsStudents.Open sql,oFunc.FPCScnn
	Dim studentCount 
	studentCount = rsStudents.RecordCount
Do While NOT rsStudents.EOF
	dim currentILP
    currentILP = rsStudents.Fields("intILP_Id").Value	
	if currentILP <> "" then
	'This section is for deleting a students ILP/Materials	
	dim deleteClassStudent
	dim deleteGradingStudent
	dim deleteGenILPStudent
	dim deleteAsocStudent
	
	set rsMaterials = server.CreateObject("ADODB.RECORDSET")
	rsMaterials.CursorLocation = 3
	sql = "select count(*) " & _
		  "from tblOrdered_Items " & _
		  "where bolApproved = 1 and intILP_ID = " & currentILP
	rsMaterials.Open sql , oFunc.FPCScnn

	
	' if rsMaterials(0) > 0 then
		' strMessage = "You can not delete this class because it has Goods or Services " & _
					 ' "that have been ordered. Contact FPCS for more help."
	' else
		' Check to see if we can delete the class record
		' Class must be taught by guardian in order to do so
		sql = "SELECT i.intClass_ID " & _
				"FROM tblILP i INNER JOIN  " & _
				"tblClasses c ON i.intClass_ID = c.intClass_ID " & _
				"WHERE     (c.intInstructor_ID IS NULL) " & _
				"AND (c.intGuardian_ID IS NOT NULL) " & _
				"AND (c.intInstruct_Type_ID = 1) " & _
				"AND (i.intILP_ID =" & currentILP & ")"
				
		' we'll reuse rsMaterials quite a bit
		rsMaterials.Close
		rsMaterials.Open sql, oFunc.FPCScnn
		
		if rsMaterials.RecordCount > 0 then
			' Class is taught by guardian.
			' Now let's see if anyone else is in the class
			' If so we can not delete it
			intClass_ID = rsMaterials("intClass_ID")
			rsMaterials.Close
			sql = "select * from tblILP where intClass_ID = " & intClass_ID 
			
			rsMaterials.Open sql, oFunc.FPCScnn
			
			' if rsMaterials.RecordCount = 1 then
				' ' only one student enrolled in the class so we can delete it
				' deleteClassStudent = "delete from tblClasses where intClass_ID = " & intClass_ID
				' deleteAsocStudent = "delete from tascClass_Family where intClass_id = " & intClass_ID
				' ' now let's check to see if we can delete the generic ILP
				' 'sql = "select intILP_ID,intGrading_Scale_ID " & _
				' '		"from tblILP_Generic " & _
				' '		"where intClass_ID = " & intClass_ID & _
				' '		" and bolILP_Bank <> 1 "
				' 'rsMaterials.Close
				' 'rsMaterials.Open sql, oFunc.FPCScnn
				
				' 'if rsMaterials.RecordCount > 0 then
					' ' Generic ILP is not saved in ILP Bank so we can delete it
				' '	deleteGenILPStudent = "delete from tblILP_Generic where intILP_ID = " & rsMaterials("intILP_ID")
				' '	if rsMaterials("intGrading_Scale_ID") & "" <> "" then
				' '		deleteGradingStudent = "delete from tblGrading_Scale_Generic where intGrading_Scale_ID = " & rsMaterials("intGrading_Scale_ID")
				' '	end if				
				' 'end if 			
			' end if
		end if		
		
		if strMessage = "" then
			rsMaterials.Close
			sql = "select intShort_ILP_ID from tblILP where intILP_ID = " & currentILP
			rsMaterials.Open sql, oFunc.FPCScnn
			
			oFunc.BeginTransCN
			'Delete the Ordered Goods/Services (ordered Item attributes are auto deleted using
			' a cascade delete
			delete = "delete from tblOrdered_Items where intILP_ID = " & currentILP
			oFunc.ExecuteCN(delete)
			' Delete Progress Ratings for this ILP
			delete = "delete from tblCourse_Progress_Ratings where intILP_ID = " & currentILP
			oFunc.ExecuteCN(delete)
			' delete prgress report comments for this ILP
			delete = "delete from PROGRESS_REPORT_COMMENTS where progressReportID in (select ProgressReportID from PROGRESS_REPORTS where IlpId = " &  currentILP & ")"
			oFunc.ExecuteCN(delete)
			' delete the progress report record
			delete = "delete from PROGRESS_REPORTS where IlpId = " & currentILP
			oFunc.ExecuteCN(delete)
			'Delete the ILP
			delete = "delete from tblILP where intILP_id = " & currentILP
response.write delete
			oFunc.executeCN(delete)		
			' Reset budget records so we can reuse them if this course is implemented again
			update = "update tblBudget set intOrdered_Item_ID = NULL where intShort_ILP_ID = " & rsMaterials("intShort_ILP_ID")
			oFunc.ExecuteCN(update)
			
			if deleteGenILPStudent <> "" then
				oFunc.ExecuteCN(deleteGenILPStudent)
			end if 
			
			if deleteGradingStudent <> "" then
				oFunc.ExecuteCN(deleteGradingStudent)
			end if 								
			
			if deleteClassStudent <> "" then
				oFunc.ExecuteCN(deleteAsocStudent)
				oFunc.ExecuteCN(deleteClassStudent)
			end if 

			oFunc.CommitTransCN
		end if
	'end if 
	end if
	rsStudents.MoveNext
	Loop
	strMessage = "All students were dropped"
end if


if Request.QueryString("intILP_ID") <> "" then
	'This section is for deleting a students ILP/Materials	
	dim deleteClass
	dim deleteGrading
	dim deleteGenILP
	dim deleteAsoc
	
	set rsMaterials = server.CreateObject("ADODB.RECORDSET")
	rsMaterials.CursorLocation = 3
	sql = "select count(*) " & _
		  "from tblOrdered_Items " & _
		  "where bolApproved = 1 and intILP_ID = " & Request.QueryString("intILP_ID")
	rsMaterials.Open sql , oFunc.FPCScnn

	
	if rsMaterials(0) > 0 then
		strMessage = "You can not delete this class because it has Goods or Services " & _
					 "that have been ordered. Contact FPCS for more help."
	else
		' Check to see if we can delete the class record
		' Class must be taught by guardian in order to do so
		sql = "SELECT i.intClass_ID " & _
				"FROM tblILP i INNER JOIN  " & _
				"tblClasses c ON i.intClass_ID = c.intClass_ID " & _
				"WHERE     (c.intInstructor_ID IS NULL) " & _
				"AND (c.intGuardian_ID IS NOT NULL) " & _
				"AND (c.intInstruct_Type_ID = 1) " & _
				"AND (i.intILP_ID =" & Request.QueryString("intILP_ID") & ")"
				
		' we'll reuse rsMaterials quite a bit
		rsMaterials.Close
		rsMaterials.Open sql, oFunc.FPCScnn
		
		if rsMaterials.RecordCount > 0 then
			' Class is taught by guardian.
			' Now let's see if anyone else is in the class
			' If so we can not delete it
			intClass_ID = rsMaterials("intClass_ID")
			rsMaterials.Close
			sql = "select * from tblILP where intClass_ID = " & intClass_ID 
			
			rsMaterials.Open sql, oFunc.FPCScnn
			
			if rsMaterials.RecordCount = 1 then
				' only one student enrolled in the class so we can delete it
				deleteClass = "delete from tblClasses where intClass_ID = " & intClass_ID
				deleteAsoc = "delete from tascClass_Family where intClass_id = " & intClass_ID
				' now let's check to see if we can delete the generic ILP
				'sql = "select intILP_ID,intGrading_Scale_ID " & _
				'		"from tblILP_Generic " & _
				'		"where intClass_ID = " & intClass_ID & _
				'		" and bolILP_Bank <> 1 "
				'rsMaterials.Close
				'rsMaterials.Open sql, oFunc.FPCScnn
				
				'if rsMaterials.RecordCount > 0 then
					' Generic ILP is not saved in ILP Bank so we can delete it
				'	deleteGenILP = "delete from tblILP_Generic where intILP_ID = " & rsMaterials("intILP_ID")
				'	if rsMaterials("intGrading_Scale_ID") & "" <> "" then
				'		deleteGrading = "delete from tblGrading_Scale_Generic where intGrading_Scale_ID = " & rsMaterials("intGrading_Scale_ID")
				'	end if				
				'end if 			
			end if
		end if		
		
		if strMessage = "" then
			rsMaterials.Close
			sql = "select intShort_ILP_ID from tblILP where intILP_ID = " & request("intILP_ID")
			rsMaterials.Open sql, oFunc.FPCScnn
			
			oFunc.BeginTransCN
			'Delete the Ordered Goods/Services (ordered Item attributes are auto deleted using
			' a cascade delete
			delete = "delete from tblOrdered_Items where intILP_ID = " & Request.QueryString("intILP_ID")
			oFunc.ExecuteCN(delete)
			' Delete Progress Ratings for this ILP
			delete = "delete from tblCourse_Progress_Ratings where intILP_ID = " & Request.QueryString("intILP_ID")
			oFunc.ExecuteCN(delete)
			' delete prgress report comments for this ILP
			delete = "delete from PROGRESS_REPORT_COMMENTS where progressReportID in (select ProgressReportID from PROGRESS_REPORTS where IlpId = " &  Request.QueryString("intILP_ID") & ")"
			oFunc.ExecuteCN(delete)
			' delete the progress report record
			delete = "delete from PROGRESS_REPORTS where IlpId = " & Request.QueryString("intILP_ID")
			oFunc.ExecuteCN(delete)
			'Delete the ILP
			delete = "delete from tblILP where intILP_id = " & Request.QueryString("intILP_id")
response.write delete
			oFunc.executeCN(delete)		
			' Reset budget records so we can reuse them if this course is implemented again
			update = "update tblBudget set intOrdered_Item_ID = NULL where intShort_ILP_ID = " & rsMaterials("intShort_ILP_ID")
			oFunc.ExecuteCN(update)
			
			if deleteGenILP <> "" then
				oFunc.ExecuteCN(deleteGenILP)
			end if 
			
			if deleteGrading <> "" then
				oFunc.ExecuteCN(deleteGrading)
			end if 								
			
			if deleteClass <> "" then
				oFunc.ExecuteCN(deleteAsoc)
				oFunc.ExecuteCN(deleteClass)
			end if 
			
			strMessage = "ILP, Goods and Services have been deleted."
			
			oFunc.CommitTransCN
		end if
	end if 
	
	rsMaterials.Close
	set rsMaterials = nothing		
	
	
end if 

%>
<html>
<head>
<title>Delete Class</title>
<link rel="stylesheet" href="<% = strPath %>/CSS/homestyle.css">
</head>
<body background=c0c0c0>
<form id=form1 name=form1>
<table width=100% height=100%>
	<tr>
		<Td align=center valign=middle>
			<table>
				<tr>
					<Td class=svplain10>
					
						<% = strMessage %><br><BR>
						<center>
						<input type=button value="Close" onCLick="window.opener.location.reload();window.opener.focus();window.close();" id=button1 name=button1>
						</center>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</body>
</html>
<%
oFunc.closeCN
set oFunc = nothing	
%>