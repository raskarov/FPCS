<%@ Language=VBScript %>
<%
dim insert
dim intILP_ID
dim intGrading_Scale_ID
dim update
dim delete
dim arMaterials
dim sql
dim rsGetMat
dim intOrdered_Materials_id
dim strField
dim strValues
dim strILPTable
dim strUpdate
dim szOther_Grading
dim bolPass_Fail
dim oFunc						'wsc object
dim bolILP_Bank
dim bolGradingScale


if Session.Contents("bolUserLoggedIn") = false then
	Response.Expires = -1000	'Makes the browser not cache this page
	Response.Buffer = True		'Buffers the content so our Response.Redirect will work
	Session.Contents("strURL") = Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
	Server.Execute(Application.Value("strWebRoot") & "UserAdmin/Login.asp")
else 
   set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
   call oFunc.OpenCN()
   	
   	' Set up ILP Bank value
   	if request("bolILP_Bank") <> "" or request("bolFromILPBank") <> "" then
   		bolILP_Bank = 1
   	else
   		bolILP_Bank = 0 
   	end if
   	
   	szOther_Grading = request("szOther_Grading")
   
   	if Request("bolPass_Fail") <> "" then
   		bolPass_Fail = 1
   	else
   		bolPass_Fail = 0
   	end if
   	
   	if request("bolGradingScale") <> "" then
   		bolGradingScale = 1
   	else
   		bolGradingScale = 0 
   	end if
   	
   	' This will remove an ILP from the ILP Bank
   	if request("bolRemove") <> "" then
   		update = "update tblILP_Generic set bolILP_Bank = 0 where intILP_ID = " & request("bolRemove")
		oFunc.ExecuteCN(update)
	end if   		
	' Fusebox like logic for script steering. This if statement executes the needed function
	' based on incoming requests from the HTTP header
	if request("bolFromILPBank") <> "" then
		'Reset session only if coming from ILP Bank. Otherwise we could lose some records.
		Session.Contents("intILP_ID") = ""		
	end if 
	
	if Session.Contents("intILP_ID") = "" and Request.Form("edit") = "" and request("bolILP_ADD") = "" then
		'This means we need to insert an ILP
		'Session.Contents("intILP_ID") makes it so we can only add this ILP once a session 
		'to cut down on corrupting data by using the browser back button and resubmitting
		call vbfInsertILP
	elseif request.Form("edit") <> "" and request("bolILP_ADD") <> "" then
		call vbfILPAdds(request("intILP_ID"), "tblILP", request("szIlp_Additions"))
	elseif request.Form("edit") <> "" and request("bolILP_TEACHER_ADD") <> "" then
		call vbfILPAdds(request("intILP_ID_Generic"), "tblILP_GENERIC", request("Teacher_Additions"))
	elseif  request.Form("edit") = "" and request("bolILP_ADD") <> "" then
		' Transfer a generic ASD ILP to a Student ILP
		'response.End
		vbsCopyGenericToReal(request("intILP_ID"))
	elseif Request.Form("edit") <> "" then
		'We are now in edit mode and need to update the ILP	
		call vbfUpdateILP
	end if 	
	
	call oFunc.CloseCN
	set oFunc = nothing

	if Request.Form("edit") <> "" or request("bolFromILPBank") <> "" or request("bolLateAdd") <> "" then
		dim strRefreshPage
		if request("bolLateAdd") <> "" or (Request.Form("edit") <> "" and request("hdnHrsChanged") <> "" ) then
			strRefreshPage = "window.opener.location.reload();"
		end if
%>
	<HTML>
	<HEAD>
	<script language=javascript>
		function jfRefresh(){
			// Don't refresh page if coming from reqApprovalAdmin.asp
			var strScript = window.opener.location.href;
			if (strScript.indexOf("reqApprovalAdmin") < 1) {
				<% = strRefreshPage %>
			}
		}
	</script>
	</HEAD>
	<BODY onLoad="window.opener.focus();jfRefresh();window.close();">
	</BODY>
	</HTML>
	<%
	else
		dim strURL
		strURL = "forms/requisitions/req1.asp?intClass_ID=" & Session.Contents("intClass_Id")
		strURL = strURL & "&intILP_ID=" & Session.Contents("intILP_ID")
		strURL = strURL & "&intStudent_ID=" & Session.Contents("intStudent_ID")
		strURL = strURL & "&bolFromILPInsert=true"
	%>
	<html>
	<body bgcolor=white onLoad="window.location.href='<%= Application.Value("strWebRoot") & strURL%>';">
	</body>
	</html>
	<%
	end if
end if

function vbfInsertILP
	oFunc.BeginTransCN
	if request("intContract_Guardian_ID") <> "" then
		strField = strField & ",intContract_Guardian_ID"
		strValues = strValues & "'" & oFunc.EscapeTick(request("intContract_Guardian_ID")) & "'," 
	end if
	
	if Session.Contents("intStudent_ID") <> "" then
		' First check to see if this course has already been implemented
		set rsI = server.CreateObject("ADODB.RECORDSET")
		rsI.CursorLocation = 3
		sql = "select * from tblILP where intShort_ILP_ID = " & Session.Contents("intShort_ILP_ID")
		rsI.Open sql,Application("cnnFPCS")'oFunc.FPCScnn
		
		if rsI.RecordCount > 0 then
			%>
			<html>
				<header>
					<script language="javascript">
						window.location.href="<% = Application.Value("strWebRoot") %>forms/packet/packet.asp?intStudent_ID=<% = Session.Contents("intStudent_ID")%>&strMessage=Could not implement class.Class has already been implemented.";
					</script>
				</header>
				<body>
				</body>
			</html>
			<%
			rsI.Close
			set rsI = nothing
			response.End
		end if
		' Adding an ILP for a student from a Generic ILP that already was tied
		' to a class
		strField = strField & ",intStudent_ID"
		strValues = strValues & "'" & Session.Contents("intStudent_ID") & "'," 
		
		strILPTable = "tblILP"
		call vbfInsertRecords()				
		
	else			
		' Handles teacher created ILP's (ASD teacher ILP's are always generic)				
		strILPTable = "tblILP_Generic"
		call vbfInsertRecords()	
	end if
	Session.Contents("intShort_ILP_ID") = ""
	oFunc.CommitTransCN
end function

function vbfInsertRecords()	
	if strILPTable = "tblILP_Generic" then					
			strEnrolledField = ",intPOS_Subject_ID,intInstructor_ID"
			strEnrolledValue = "'" & request("intPOS_Subject_ID") & "'," & _								   
								oFunc.CheckDecimal(session.Contents("intInstructor_ID")) & ","		
	else
		strEnrolledField = ",dtStudent_Enrolled"
		strEnrolledValue = "'" & request("dtStudent_Enrolled") & "'," 
		'added strShortILPField - bkm 21-jun-02
		strShortILPField = ",intShort_ILP_ID"
		if Session.Contents("intShort_ILP_ID") & "" <> "" then
			strShortILPValue = Session.Contents("intShort_ILP_ID") & "," 
		else
			response.Write "ERROR- A value is required for intShort_ILP_ID"
			response.End 
		end if
	end if 		
	insert = "insert into " & strILPTable & "(intClass_Id,sintSchool_year," & _
				"intSemester,decCourse_Hours" & strEnrolledField & strShortILPField & ",szCurriculum_Desc," & _
			    "szGoals,szRequirements,szTeacher_Role,szStudent_Role,bolILP_Bank,szILP_Name," & _
			    "szParent_Role" & strField & ",szEvaluation,szEvaluationFrequency,bolPass_Fail,szOther_Grading,szUSER_CREATE,bolGradingScale,szILP_Additions) Values (" & _		     
				"'" & Session.Contents("intClass_Id") & "'," & _
				"'" & session.Contents("intSchool_Year") & "'," & _
			    "'" & oFunc.EscapeTick(request("intSemester")) & "'," & _
				"'" & oFunc.EscapeTick(request("decCourse_Hours")) & "'," & _
				strEnrolledValue & strShortILPValue & _
				"'" & oFunc.EscapeTick(request("szCurriculum_Desc")) & "'," & _
				"'" & oFunc.EscapeTick(request("szGoals")) & "'," & _
				"'" & oFunc.EscapeTick(request("szRequirements")) & "'," & _
				"'" & oFunc.EscapeTick(request("szTeacher_Role")) & "'," & _
				"'" & oFunc.EscapeTick(request("szStudent_Role")) & "'," & _
				bolILP_Bank & "," & _
				"'" & oFunc.EscapeTick(request("szILP_Name")) & "'," & _
				"'" & oFunc.EscapeTick(request("szParent_Role")) & "'," & _
				strValues & _
				"'" & oFunc.EscapeTick(request("szEvaluation")) & "'," & _
				"'" & oFunc.EscapeTick(request("szEvaluationFrequency")) & "'," & _
				bolPass_Fail & "," & _
				"'" & oFunc.EscapeTick(szOther_Grading) & "'," & _
				"'" & Session.Contents("strUserID")	& "', " & bolGradingScale & "," & _
				"'" & oFunc.EscapeTick(request("szILP_Additions")) & "')"
	oFunc.ExecuteCN(insert)
	
	intILP_ID = oFunc.GetIdentity
	'If page is submited twice we only want to save the record once.
	' Since this session variable is defined it will prevent us from inserting twice in the first if statement
	Session.Contents("intILP_ID") = intILP_ID
	
	
	if Session.Contents("intShort_ILP_ID") & "" <> "" then
		call vbsUpdateHrs(Session.Contents("intShort_ILP_ID"),request("decCourse_Hours"))
	end if	
end function 
Function vbfSaveSyllabus
Dim intILP_ID
		intILP_ID = oFunc.EscapeTick(Request.Form("intILP_ID"))
        'If intILP_ID = Empty Then
		'intILP_ID = oFunc.EscapeTick(Request.Form("intILP_ID_Generic"))
        'End If
Dim insert
Dim update
Dim cmd
insert = "insert dbo.tblSyllabus(intILP_ID, weekNo, dtStart, dtEnd, szDescription) values(@intILP_ID, @weekNo, '@dtStart', '@dtEnd', '@szDescription')"
update = "update dbo.tblSyllabus set weekNo=@weekNo, dtStart='@dtStart', dtEnd='@dtEnd', szDescription='@szDescription' where syllabusId=@syllabusId"
For i = 1 to Request.Form("syllabusId").Count
If Request.Form("syllabusId")(i)="new" Then
    If Request.Form("weekNo")(i)>"" Then
        cmd = Replace(insert,"@intILP_ID",intILP_ID)
        cmd = Replace(cmd,"@weekNo",oFunc.EscapeTick(Request.Form("weekNo")(i)))
        cmd = Replace(cmd,"@dtStart",oFunc.EscapeTick(Request.Form("dtStart")(i)))
        cmd = Replace(cmd,"@dtEnd",oFunc.EscapeTick(Request.Form("dtEnd")(i)))
        cmd = Replace(cmd,"@szDescription",oFunc.EscapeTick(Request.Form("szDescription")(i)))
        ok=true
    Else
    ok=false
    'GoTo NextLoop:
    End If
Else
    If Request.Form("weekNo")(i)>"" Then
        cmd = Replace(update,"@syllabusId",oFunc.EscapeTick(Request.Form("syllabusId")(i)))
        cmd = Replace(cmd,"@weekNo",oFunc.EscapeTick(Request.Form("weekNo")(i)))
        cmd = Replace(cmd,"@dtStart",oFunc.EscapeTick(Request.Form("dtStart")(i)))
        cmd = Replace(cmd,"@dtEnd",oFunc.EscapeTick(Request.Form("dtEnd")(i)))
        cmd = Replace(cmd,"@szDescription",oFunc.EscapeTick(Request.Form("szDescription")(i)))
        ok=true
    Else
        ok=False
    'GoTo NextLoop:
    End If

End If
        If intILP_ID = Empty Then
        ok=false
        End If
If ok Then
	oFunc.ExecuteCN(cmd)
End If
'NextLoop:
Next
End Function

function vbfUpdateILP
	' It's time to Update the ILP. We start with the main table ... 
	if Request.Form("intILP_ID_Generic") <> "" then
		' Updates Teacher ILP's
		intILP_ID = Request.Form("intILP_ID_Generic")
		strField = "szILP_Additions='" & oFunc.EscapeTick(request("Teacher_Additions")) & "', " 
		strValues = ""
		strILPTable = "tblILP_Generic"
	else		
		' Updates Student ILPs	
		intILP_ID = Request.Form("intILP_ID")
		strField =  "dtStudent_Enrolled = '" & request("dtStudent_Enrolled") & "'," & _
					" szILP_Additions = '" &  oFunc.escapeTick(request("szILP_Additions")) & "'," 
		strILPTable = "tblILP"			
	end if 
	'oFunc.BeginTransCN
		
	update = "update " & strILPTable & " set " & _
				"intSemester = '" & oFunc.EscapeTick(Request("intSemester")) & "'," & _
				"decCourse_Hours = '" & oFunc.EscapeTick(Request("decCourse_Hours")) & "'," & _
			    strField & _
				"szCurriculum_Desc = '" & oFunc.EscapeTick(Request("szCurriculum_Desc")) & "'," & _
				"szGoals = '" & oFunc.EscapeTick(Request("szGoals")) & "'," & _
				"szRequirements = '" & oFunc.EscapeTick(Request("szRequirements")) & "'," & _
				"szTeacher_Role = '" & oFunc.EscapeTick(Request("szTeacher_Role")) & "'," & _
				"szStudent_Role = '" & oFunc.EscapeTick(Request("szStudent_Role")) & "'," & _
				"szParent_Role = '" &oFunc.EscapeTick( Request("szParent_Role")) & "'," & _
				"szEvaluation = '" & oFunc.EscapeTick(Request("szEvaluation")) & "'," & _
				"szEvaluationFrequency = '" & oFunc.EscapeTick(Request("szEvaluationFrequency")) & "'," & _
				"bolILP_Bank= " & bolILP_Bank & "," & _
				"bolGradingScale= " & bolGradingScale & "," & _
				"szILP_Name='" & oFunc.EscapeTick(request("szILP_Name")) & "', " & _				
				"bolPass_Fail = " & bolPass_Fail & "," & _
				"szOther_Grading = '" & oFunc.EscapeTick(szOther_Grading) & "', " & _
				"szUSER_MODIFY = '" & Session.Contents("strUserID") & "' " & _
				"where intILP_ID = " & intILP_ID
				
	oFunc.ExecuteCN(update)
	'oFunc.CommitTransCN
	
	' Update hours in short ilp
	if intILP_ID <> "" and request("hdnHrsChanged") <> "" then 'and session.Contents("strRole") <> "TEACHER" then
		sql = "select intShort_ILP_ID from tblILP where intILP_ID = " & intILP_ID
		set rs = server.CreateObject("ADODB.RECORDSET")
		rs.Open sql,Application("cnnFPCS")'oFunc.FPCScnn
		
		if rs.RecordCount > 0 then
			if rs("intShort_ILP_ID") & "" <> ""  then
				call vbsUpdateHrs(rs("intShort_ILP_ID"),Request("decCourse_Hours"))
			end if 
		end if
		rs.Close
		set rs = nothing
	end if
    Call vbfSaveSyllabus
end function 

sub vbfILPAdds(pID,pTable,pText)

	dim update, strAddUpdate
	
	if oFunc.IsAdmin and isNumeric(request.Form("decCourse_Hours")) then
		strAddUpdate = ", decCourse_Hours = " & request.Form("decCourse_Hours") & " " 
	end if
	
	if bolGradingScale = 1 or bolPass_Fail = 1 or szOther_Grading <> "" then
		strAddUpdate = strAddUpdate & ", bolGradingScale= " & bolGradingScale & "," & _				
				"bolPass_Fail = " & bolPass_Fail & "," & _
				"szOther_Grading = '" & oFunc.EscapeTick(szOther_Grading) & "' "
	end if
	
	update = "update " & pTable & " set " & _
			 "szILP_Additions = '" & oFunc.EscapeTick(pText) & "'," & _
			 "dtModify = '" & oFunc.DateTimeFormat(now()) & "',"& _
			 "szUser_Modify = '" & session.Contents("strUserID") & "',"  & _
			 "bolILP_Bank= " & bolILP_Bank & "," & _
			 "szILP_Name='" & oFunc.EscapeTick(request("szILP_Name")) & "' " & _
			 strAddUpdate & _
			 " where intILP_ID = " & pID
	oFunc.ExecuteCN(update)
	
end sub

sub vbsUpdateHrs(shortForm,hrs)
	' Update the course hrs in our short form to maintain consitancy between
	' planned hrs and actual hrs
	update = "update tblILP_Short_Form set " & _
			 "intCourse_Hrs = " & oFunc.CheckDecimal(hrs) & _
			 " Where intShort_ILP_ID = " & shortForm
	oFunc.ExecuteCN(update)
end sub

sub vbsCopyGenericToReal(pID)
	dim insert
	
	if Session.Contents("intShort_ILP_ID") & "" <> "" then
		strShortILPValue = Session.Contents("intShort_ILP_ID")
	else
		response.Write "ERROR- Need a value for intShort_ILP_ID"
		response.End
	end if
	
	insert = "INSERT INTO tblILP " & _ 
			" (intStudent_ID, intClass_ID, sintSchool_Year, intSemester, decCourse_Hours, intContract_Guardian_id, szCurriculum_Desc, szGoals, szRequirements,  " & _ 
			" szTeacher_Role, szStudent_Role, szParent_Role, szEvaluation, szEvaluationFrequency, bolPass_Fail, bolGradingScale,szOther_Grading, intEnroll_Info_ID,  " & _ 
			"  bolILP_Bank, szILP_Name, szILP_Additions,dtStudent_Enrolled,intShort_ILP_ID,szUSER_CREATE) " & _ 
			"SELECT " & Session.Contents("intStudent_ID") & " as intStudent_ID, " & Session.Contents("intClass_ID")	 & " as intClass_ID, " & session.Contents("intSchool_Year") & " as sintSchool_Year, intSemester, decCourse_Hours, " & _
			request("intContract_Guardian_ID") & " as intContract_Guardian_ID, szCurriculum_Desc, szGoals, szRequirements,  " & _ 
			" szTeacher_Role, szStudent_Role, szParent_Role, szEvaluation, szEvaluationFrequency, bolPass_Fail,bolGradingScale, szOther_Grading, intEnroll_Info_ID,  " & _ 
			 bolILP_Bank & " as bolILP_Bank,'" & oFunc.EscapeTick(request("szILP_Name")) & "' as szILP_Name, '" & oFunc.escapeTick(request("szIlp_Additions")) & "' as  szILP_Additions, " & _
			" convert(datetime,'" & request("dtStudent_Enrolled") & "') as  dtStudent_Enrolled, " & strShortILPValue & " as intShort_ILP_ID, '" & Session.Contents("strUserID") & "' as szUSER_CREATE" & _
			" FROM tblILP_Generic " & _ 
			"WHERE (intILP_ID = " & pID & ") "
	oFunc.ExecuteCN(insert)

	Session.Contents("intILP_ID") = oFunc.GetIdentity
end sub
%>
