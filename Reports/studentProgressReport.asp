<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		studentProgressReport.asp
'Purpose:	Facilitates the storing/reporting of course progress
'Date:		07 Dec 2004
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim intStudent_ID 
dim sql
dim mError		'conitains our error messages after validation is complete
dim strDiasbled 
dim strStudentName
dim arInfo
dim arFamInfo
dim bolPrint
dim  printCount
printCount = 0
intReporting_Period_ID = request("intReporting_Period_ID")

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

'Initialize some key variables
if request("intStudent_ID") = "" then
	'terminate page since page was improperly called.
	response.Write "<html><body><H1>Page improperly called.</h1></body></html>"
	response.End
end if

if request.Form.Count > 0 then
	' Transfers all of the post http header variables into vbs variables
	' so we can more readily access them
	for each i in request.Form
		execute("dim " & i)
		execute(i & " = """ & replace(replace(replace(request.Form(i),"""","'"),chr(13),""),chr(10),"") & """")
	next 
end if 

if request("print") <> "" then bolPrint = true



'Validate Budget Transfer form if needed
if hdnSave <> "" then
	vbsInsertProgress()
end if

'prevent page from cacheing
Response.AddHeader "Cache-Control","No-Cache"
Response.Expires = -1

'Print the header
Session.Value("strTitle") = "Student Progress Report"
Session.Value("strLastUpdate") = "08 Dec 2004"

if request("SimpleHeader") <> "" or bolPrint then
	Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")
	if bolPrint then
	%>
	<link rel="stylesheet" href="<% =  Application.Value("strWebRoot") %>CSS/printStyle.css">
	<%
	end if
else
	Server.Execute(Application.Value("strWebRoot") & "includes/Header.asp")
end if

%>
<form name=main method=post action="studentProgressReport.asp" ID="Form1">
<input type="hidden" name="SimpleHeader" value="<% = request("SimpleHeader") %>">
<%

dim arStudents
arStudents = split(request("intStudent_ID"),",")
for i = 0 to ubound(arStudents)
	if arStudents(i) <> "" then 
		intStudent_ID = arStudents(i)
		arInfo = oFunc.StudentInfo(intStudent_ID,8)
		studentGrade = arInfo(4)
		strStudentName = arInfo(2)
		if bolPrint and printCount > 1 then 
			response.Write "<p></p>"
			strUniqueName = printCount
		else
			strUniqueName = ""
		end if
%>
<input type=hidden name=intStudent_ID value="<% = intStudent_ID %>" ID="Hidden1">
<table style="width:100%;" ID="Table3">
	<tr>
		<td class="yellowHeader">
			&nbsp;<b>Progress Report for <% = strStudentName %>, SY <% = oFunc.SchoolYearRange %></b>&nbsp;&nbsp;Grade: <%=studentGrade%>
		</td>
	</tr>
	<tr>
		<td>
			<table ID="Table1">
				<tr>
					<td class="svplain8" nowrap>
						<b><% if bolPrint then response.Write "Reporting Period: " else response.Write "Select Reporting Period:" end if %></b>
					</td>
					<td class="svplain8">
						<%
							sql = "SELECT intReporting_Period_ID, szReporting_Period_Name " & _ 
										"FROM trefReporting_Periods " & _ 
										"ORDER BY szReporting_Period_Name "
							strList =  oFunc.MakeListSQL(sql,"intReporting_Period_ID","szReporting_Period_Name",intReporting_Period_ID)
								
							if bolPrint then
								response.Write oFunc.SelectedListText
							else
						%>
						<select name="intReporting_Period_ID" onchange="this.form.submit();" ID="Select1">
							<option value=""></option>
							<% = strList %>							
						</select> <b>Semester I due <% = Application.Contents("dtSem_One_Progress_Deadline" &  session.Contents("intSchool_Year")) %> : Semester II due <% = Application.Contents("dtSem_Two_Progress_Deadline" &  session.Contents("intSchool_Year")) %></b>
						<%  end if %>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>
		
		
<%
if intReporting_Period_ID <> "" then
	'sql = "SELECT     ISF.szCourse_Title, POS.txtCourseTitle, ISF.intShort_ILP_ID, tblILP.intILP_ID, tblILP.bolApproved AS aStatus, tblILP.bolSponsor_Approved AS sStatus,  " & _ 
	'	"                      CASE ISF.intPOS_Subject_ID WHEN 22 THEN 0 ELSE 1 END AS isSponsor, ISF.intCourse_Hrs, tblILP.decCourse_Hours, tblClasses.intInstructor_ID,  " & _ 
	'	"                      tps.szSubject_Name, tblClasses.intClass_ID, tblClasses.intInstruct_Type_ID, tblILP.intContract_Guardian_ID, tblClasses.intGuardian_ID,  " & _ 
	'	"                      tblClasses.intVendor_ID, tblClasses.szClass_Name, CASE WHEN tblClasses.intInstructor_ID IS NOT NULL  " & _ 
	''	"                      THEN ins.szFirst_Name + ' ' + ins.szLast_Name WHEN tblClasses.intGuardian_ID IS NOT NULL  " & _ 
	'	"                      THEN g.szFirst_Name + ' ' + g.szLast_Name END AS teacherName, tblILP.szAdmin_Comments, tblILP.szSponsor_Comments,  " & _ 
	'	"                      tblILP.bolReady_For_Review,  " & _ 
	'	"                      CASE WHEN tblILP.bolApproved = 1 THEN 'a-appr' WHEN tblILP.bolApproved = 0 THEN 'a-must amend' ELSE CASE WHEN tblILP.bolSponsor_Approved " & _ 
	'	"                       = 1 THEN 's-appr' WHEN tblILP.bolSponsor_Approved = 0 THEN 's-must amend' WHEN tblILP.bolReady_For_Review = 1 THEN 'ready for review' ELSE " & _ 
	''	"                       'implemented' END END AS ILPStatus, trefProgress_Ratings.intProgress_Rating_ID, trefProgress_Ratings.szProgress_Rating_Name,  " & _ 
	'	"                      tblProgress_Reports.intProgress_Report_ID, tblProgress_Reports.bolReEnroll," & _ 
	'	"					   tblProgress_Reports.bolHave_Materials, tblProgress_Reports.szMaterials_Not_Received,  " & _
     '   "					   tblProgress_Reports.bolVendors_Paid, tblProgress_Reports.bolTrain_PSC, tblProgress_Reports.bolTrain_GS, " & _ 
      '  "					   tblProgress_Reports.bolTrain_Reimburse, tblProgress_Reports.bolTrain_Grad, tblProgress_Reports.szParent_Comments,  " & _
       ' "					   tblProgress_Reports.szSponsor_Comments AS SponsorComments, tblProgress_Reports.szRole_Create,tblProgress_Reports.dtCREATE, " & _
      '  "					   tblProgress_Reports.szUSER_CREATE,tblProgress_Reports.dtSponsor_Reviewed " & _
	'	"FROM         tblProgress_Reports INNER JOIN " & _ 
	'	"                      tblCourse_Progress_Ratings INNER JOIN " & _ 
	'	"                      trefProgress_Ratings ON tblCourse_Progress_Ratings.intProgress_Rating_ID = trefProgress_Ratings.intProgress_Rating_ID ON  " & _ 
	''	"                      tblProgress_Reports.intProgress_Report_ID = tblCourse_Progress_Ratings.intProgress_Report_ID " & _
	'	"					   and tblProgress_Reports.intReporting_Period_ID = " & intReporting_Period_ID & " RIGHT OUTER JOIN " & _ 
	'	"                      trefPOS_Subjects tps INNER JOIN " & _ 
	'	"                      tblILP_SHORT_FORM ISF ON tps.intPOS_Subject_ID = ISF.intPOS_Subject_ID INNER JOIN " & _ 
	'	"                      tblClasses INNER JOIN " & _ 
	'	"                      tblILP ON tblClasses.intClass_ID = tblILP.intClass_ID ON ISF.intShort_ILP_ID = tblILP.intShort_ILP_ID ON  " & _ 
	'	"                      tblCourse_Progress_Ratings.intILP_ID = tblILP.intILP_ID LEFT OUTER JOIN " & _ 
	'	"                      tblProgramOfStudies POS ON ISF.lngPOS_ID = POS.lngPOS_ID LEFT OUTER JOIN " & _ 
	'	"                      tblINSTRUCTOR INS ON tblClasses.intInstructor_ID = INS.intINSTRUCTOR_ID LEFT OUTER JOIN " & _ 
	'	"                      tblGUARDIAN g ON tblClasses.intGuardian_ID = g.intGUARDIAN_ID " & _ 
	'	"WHERE     (ISF.intStudent_ID = " & intStudent_ID & ") AND (ISF.intSchool_Year = " & session.Contents("intSchool_Year") & _
	'	") ORDER BY isSponsor, POS.txtCourseTitle, ISF.szCourse_Title, ISF.intShort_ILP_ID "
			
	
	' BAD SQL  Did not include the Progress reporting period and over wrote I semester data when entering Second Semester
	'sql = "SELECT     tblILP.intILP_ID, CASE isNull(tblClasses.intPOS_Subject_ID, 1)  " & _ 
	'	"                      WHEN 1 THEN CASE tblClasses.intPOS_Subject_ID WHEN 22 THEN 0 ELSE 1 END ELSE CASE tblClasses.intPOS_Subject_ID WHEN 22 THEN 0 ELSE 1 " & _ 
	'	"                       END END AS isSponsor, tblILP.decCourse_Hours, tblClasses.intInstructor_ID, tblClasses.intClass_ID, tblClasses.intInstruct_Type_ID,  " & _ 
	'	"                      tblILP.intContract_Guardian_ID, tblClasses.intGuardian_ID, tblClasses.intVendor_ID, tblClasses.szClass_Name,  " & _ 
	'	"                      CASE WHEN tblClasses.intInstructor_ID IS NOT NULL THEN ins.szFirst_Name + ' ' + ins.szLast_Name WHEN tblClasses.intGuardian_ID IS NOT NULL  " & _ 
	''	"                      THEN g.szFirst_Name + ' ' + g.szLast_Name END AS teacherName, tblILP.szAdmin_Comments, tblILP.szSponsor_Comments,  " & _ 
	'	"                      tblILP.bolReady_For_Review, tblILP.dtReady_For_Review, tblILP.GuardianStatusId, tblILP.SponsorStatusId, tblILP.InstructorStatusId,  " & _ 
	'	"                      tblILP.AdminStatusId, tblILP.GuardianStatusDate, tblILP.SponsorStatusDate, tblILP.InstructorStatusDate, tblILP.AdminStatusDate,  " & _ 
	'	"                      tblILP.GuardianComments, tblILP.InstructorComments, tblClasses.intContract_Status_ID, tblClasses.dtApproved, tblClasses.szUser_Approved,  " & _ 
	'	"                      tblILP.bolSponsorAlert, tblILP.bolParentAlert, tblClasses.szASD_Course_ID, tblProgress_Reports.intProgress_Report_ID,  " & _ 
	'	"                      tblProgress_Reports.intReporting_Period_ID, tblProgress_Reports.bolHave_Materials, tblProgress_Reports.szMaterials_Not_Received,  " & _ 
	'	"                      tblProgress_Reports.bolVendors_Paid, tblProgress_Reports.bolTrain_PSC, tblProgress_Reports.bolTrain_GS,  " & _ 
	'	"                      tblProgress_Reports.bolTrain_Reimburse, tblProgress_Reports.bolTrain_Grad, tblProgress_Reports.bolReEnroll,  " & _ 
	'	"                      tblProgress_Reports.szSponsor_Comments AS SponsorComments, tblProgress_Reports.szParent_Comments, tblProgress_Reports.szRole_Create,  " & _ 
	'	"                      tblProgress_Reports.dtSponsor_Reviewed, tblCourse_Progress_Ratings.intCourse_Progress_Rating_ID,tblProgress_Reports.dtCreate,  " & _ 
	'	"                      tblCourse_Progress_Ratings.intProgress_Rating_ID, tblENROLL_INFO.intSponsor_Teacher_ID AS Sponsor_ID, tblProgress_Reports.szUSER_CREATE,  " & _ 
	'	"                      tblILP.intStudent_ID, tblILP.sintSchool_Year, tps2.szSubject_Name " & _ 
	'	"FROM         tblProgress_Reports FULL OUTER JOIN " & _ 
	'	"                      tblINSTRUCTOR INS RIGHT OUTER JOIN " & _ 
	'	"                      tblCourse_Progress_Ratings RIGHT OUTER JOIN " & _ 
	'	"                      tblClasses INNER JOIN " & _ 
	'	"                      tblILP ON tblClasses.intClass_ID = tblILP.intClass_ID INNER JOIN " & _ 
	'	"                      tblENROLL_INFO ON tblILP.intStudent_ID = tblENROLL_INFO.intSTUDENT_ID AND tblILP.sintSchool_Year = tblENROLL_INFO.sintSCHOOL_YEAR ON  " & _ 
	'	"                      tblCourse_Progress_Ratings.intILP_ID = tblILP.intILP_ID RIGHT OUTER JOIN " & _ 
	'	"                      trefPOS_Subjects tps2 ON tblClasses.intPOS_Subject_ID = tps2.intPOS_Subject_ID ON  " & _ 
	'	"                      INS.intINSTRUCTOR_ID = tblClasses.intInstructor_ID LEFT OUTER JOIN " & _ 
	'	"                      tblGUARDIAN g ON tblClasses.intGuardian_ID = g.intGUARDIAN_ID ON  " & _ 
	'	"                      tblProgress_Reports.intProgress_Report_ID = tblCourse_Progress_Ratings.intProgress_Report_ID AND  " & _ 
	'	"                      tblProgress_Reports.intSchool_Year = tblILP.sintSchool_Year AND tblProgress_Reports.intStudent_ID = tblILP.intStudent_ID " & _ 
	'	" WHERE     (tblILP.intStudent_ID = " & intStudent_ID & ") AND (tblILP.sintSchool_Year = " & session.Contents("intSchool_Year") & ") " & _ 
	'	" ORDER BY isSponsor, tblClasses.szClass_Name "
	
	sql = "SELECT     i.intILP_ID, CASE isNull(c.intPOS_Subject_ID, 1)  " & _ 
			"                      WHEN 1 THEN CASE c.intPOS_Subject_ID WHEN 22 THEN 0 ELSE 1 END ELSE CASE c.intPOS_Subject_ID WHEN 22 THEN 0 ELSE 1 END END AS isSponsor, " & _ 
			"                       i.decCourse_Hours, c.intInstructor_ID, c.intClass_ID, c.intInstruct_Type_ID, i.intContract_Guardian_ID, c.intGuardian_ID, c.intVendor_ID,  " & _ 
			"                      c.szClass_Name, CASE WHEN c.intInstructor_ID IS NOT NULL THEN ins.szFirst_Name + ' ' + ins.szLast_Name WHEN c.intGuardian_ID IS NOT NULL  " & _ 
			"                      THEN g.szFirst_Name + ' ' + g.szLast_Name END AS teacherName, i.szAdmin_Comments, i.szSponsor_Comments, i.bolReady_For_Review,  " & _ 
			"                      i.dtReady_For_Review, i.GuardianStatusId, i.SponsorStatusId, i.InstructorStatusId, i.AdminStatusId, i.GuardianStatusDate, i.SponsorStatusDate,  " & _ 
			"                      i.InstructorStatusDate, i.AdminStatusDate, i.GuardianComments, i.InstructorComments, c.intContract_Status_ID, c.dtApproved, c.szUser_Approved,  " & _ 
			"                      i.bolSponsorAlert, i.bolParentAlert, c.szASD_Course_ID, pr.intProgress_Report_ID, pr.intReporting_Period_ID, pr.bolHave_Materials,  " & _ 
			"                      pr.szMaterials_Not_Received, pr.bolVendors_Paid, pr.bolTrain_PSC, pr.bolTrain_GS, pr.bolTrain_Reimburse, pr.bolTrain_Grad, pr.bolReEnroll,  " & _ 
			"                      pr.szSponsor_Comments AS SponsorComments, pr.szParent_Comments, pr.szRole_Create, pr.dtSponsor_Reviewed,  " & _ 
			"                      cpr.intCourse_Progress_Rating_ID, pr.dtCREATE, cpr.intProgress_Rating_ID, ei.intSponsor_Teacher_ID AS Sponsor_ID, pr.szUSER_CREATE,  " & _ 
			"                      i.intStudent_ID, i.sintSchool_Year, tps2.szSubject_Name " & _ 
			"FROM         tblGUARDIAN g RIGHT OUTER JOIN " & _ 
			"                      trefPOS_Subjects tps2 LEFT OUTER JOIN " & _ 
			"                      tblCourse_Progress_Ratings cpr INNER JOIN " & _ 
			"                      tblProgress_Reports pr ON cpr.intProgress_Report_ID = pr.intProgress_Report_ID AND pr.intReporting_Period_ID = " & intReporting_Period_ID & " RIGHT OUTER JOIN " & _ 
			"                      tblClasses c INNER JOIN " & _ 
			"                      tblILP i ON c.intClass_ID = i.intClass_ID INNER JOIN " & _ 
			"                      tblENROLL_INFO ei ON i.intStudent_ID = ei.intSTUDENT_ID AND i.sintSchool_Year = ei.sintSCHOOL_YEAR ON  " & _ 
			"                      pr.intSchool_Year = i.sintSchool_Year AND pr.intStudent_ID = i.intStudent_ID AND cpr.intILP_ID = i.intILP_ID ON  " & _ 
			"                      tps2.intPOS_Subject_ID = c.intPOS_Subject_ID LEFT OUTER JOIN " & _ 
			"                      tblINSTRUCTOR INS ON c.intInstructor_ID = INS.intINSTRUCTOR_ID ON g.intGUARDIAN_ID = c.intGuardian_ID " & _ 
			"WHERE     (i.intStudent_ID = " & intStudent_ID & ") AND (i.sintSchool_Year =  " & session.Contents("intSchool_Year") & ") " & _ 
			"ORDER BY isSponsor, c.szClass_Name "

	dim rs 
	set rs = server.CreateObject("ADODB.RECORDSET")
	rs.CursorLocation = 3
	rs.Open sql, Application("cnnFPCS")'oFunc.FPCScnn

	if rs.RecordCount > 0 then
		dim bolChecked
		dim strILPList
		dim rs2 
		set rs2 = server.CreateObject("ADODB.RECORDSET")		
		rs2.CursorLocation = 3
		sql = "SELECT     intProgress_Rating_ID, szProgress_Rating_Name, szProgress_Rating_Short_Name " & _ 
				"FROM         trefProgress_Ratings " & _ 
				"ORDER BY intProgress_Rating_ID "
		rs2.Open sql,Application("cnnFPCS")'oFunc.FPCScnn
		
		if (rs("szUser_Create") & "" <> "" and ucase(rs("szUser_Create")) <> ucase(session.Contents("strUserID"))) or bolPrint then
			strDiasbled = " disabled "
		end if 
		
		do while not rs.EOF
%>
		<table style="width:100%;" ID="Table2">
			<tr>
				<td class="TableHeader" style="width:45%;">
					Course Title
				</td>
				<td class="TableHeader" style="width:30%;">
					Taught By
				</td>
				<td class="TableHeader" style="width:25%;">
					ILP Status 
					<%if not bolPrint then%>
					&nbsp;&nbsp;<input type="button" value="view ilp" class="btSmallGray" onclick="jfViewILP('<%=rs("intILP_ID")%>','<%=rs("intClass_ID")%>','<%=replace(rs("szClass_Name"),"'","\'")%>','<%=rs("intContract_Guardian_ID")%>','<%=rs("intVendor_ID")%>','<%=rs("teacherName")%>');" ID="Button1" NAME="Button1">
					<%end if %>
				</td>	
				<%if bolPrint then%>
				<td class="TableHeader" style="width:25%;" nowrap>
					Progress Rating
				</td>	
				<%end if %>	
			</tr>
			<tr>
				<td class="TableCell">
					<% = rs("szClass_Name") %>
				</td>
				<td class="TableCell">
					<% = rs("teacherName")%>
				</td>
				<td class="TableCell">
					<% 
						if rs("AdminStatusId") = "3" or rs("SponsorStatusId") = "3" or _
							rs("InstructorStatusId") = "3" then
								'Rejected 
								response.Write "rejected"
						elseif  rs("AdminStatusId")  = "2" or rs("SponsorStatusId") = "2" then
							' Needs Work
							response.Write "must ammend"
						elseif rs("GuardianStatusId") & "" = "1" and rs("SponsorStatusId") & "" = "1" and _
							(rs("AdminStatusId") & "" = "1" or rs("intContract_Status_Id") & "" = "5") and _
							(rs("InstructorStatusId") & ""  = "1" or _
							rs("intInstructor_ID") & "" = "" or  _
							(rs("intInstructor_ID") & "" <> "" and _
							rs("intInstructor_ID") & "" = rs("Sponsor_ID") & "")) then
							
							response.Write "signed"
						else
							' Not Signed
							response.Write "not signed"
						end if  
					
					%>
				</td>
			<% if bolPrint then 
				rs2.MoveFirst
				strProgress = ""
				do while not rs2.EOF
					if rs2("intProgress_Rating_ID") = rs("intProgress_Rating_ID") then
						if rs2("intProgress_Rating_ID") = 3 then
							strProgress = "<b>" & rs2("szProgress_Rating_Short_Name") & "</b>"
						else
							strProgress = rs2("szProgress_Rating_Short_Name")
						end if
						exit do
					end if
					rs2.MoveNext
				loop
			%>
				<td class="TableCell" nowrap>
					<% = strProgress%>
				</td>
			</tr>
			<% else %>
			</tr>
			<tr>
				<td colspan=10 align="center">
					<table ID="Table4" cellspacing="4" cellpadding="2">
						<tr>
							<%
							rs2.MoveFirst
							do while not rs2.EOF
								if rs2("intProgress_Rating_ID") = rs("intProgress_Rating_ID") then
									bolChecked = " checked "
								else
									bolChecked = ""
								end if
								
								%>
								<td class="TableCell" align="center" >
									<% = rs2("szProgress_Rating_Name") %><br>
									<% if strDiasbled <> "" then 
											if bolChecked <> "" then response.Write "X" else response.Write "&nbsp;"
									   else %>
									<input type=radio name="progress<% =rs("intILP_ID")%>" id="progress<% =rs("intILP_ID")%>" <% = bolChecked %> value="<% = rs2("intProgress_Rating_ID")%>" >
									<% end if %>
								</td>
								<%								
								rs2.MoveNext
							loop								
							%>
						</tr>
					</table>
					<br>
				</td>
			</tr>
			<% end if%>
		</table>			
<%			
			strILPList = strILPList & "," & rs("intILP_ID")
			rs.MoveNext
		loop
		
		if strILPList <> "" then
			strILPList = right(strILPList,len(strILPList)-1)
		end if
		rs.MoveFirst
		intProgress_Report_ID = rs("intProgress_Report_ID")
%>
		<input type="hidden" name="intProgress_Report_ID" value="<% = rs("intProgress_Report_ID") %>" ID="Hidden2">
		<table cellpadding="2" ID="Table5">
			<tr>
				<td class="SubHeader" colspan=2 <% if not bolPrint then %> style="font-size:10pt;" <% end if %>nowrap>
					<b>Please answer all of the following questions.</b>
				</td>
			</tr>
			<tr>
				<td class="TableCell" nowrap>
					<b>Do you have all of the <br>curriculum/materials that you ordered?</b>
				</td>
				<td class="TableCell">
					Yes<input type=radio  <% = strDiasbled %> id="bolHave_Materials" name="bolHave_Materials<%= strUniqueName%>" <% if rs("bolHave_Materials") then response.write " checked "%> value="1">
					| No<input type=radio  <% = strDiasbled %> name="bolHave_Materials<%= strUniqueName%>" <% if rs("bolHave_Materials") = 0 then response.write " checked "%> value="0">
					| N/A<input type=radio  <% = strDiasbled %> name="bolHave_Materials<%= strUniqueName%>" <% if rs("bolHave_Materials") & "" = "" then response.write " checked "%> value="NULL">
				</td>
			</tr>				
			<tr>
				<td class="TableCell" nowrap>
					<b>Are your vendors getting paid?</b>
				</td>
				<td class="TableCell">
					Yes<input type=radio  <% = strDiasbled %> id="bolVendors_Paid" name="bolVendors_Paid<%= strUniqueName%>" <% if rs("bolVendors_Paid") then response.write " checked "%> value="1" ID="radio1">
					| No<input type=radio  <% = strDiasbled %> id="Radio  <% = strdiasbled %>2" name="bolVendors_Paid<%= strUniqueName%>" <% if rs("bolVendors_Paid") = 0 then response.write " checked "%> value="0" ID="radio2">
					| N/A<input type=radio  <% = strDiasbled %> id="Radio  <% = strdiasbled %>1" name="bolVendors_Paid<%= strUniqueName%>" <% if rs("bolVendors_Paid") & "" = "" then response.write " checked "%> value="NULL" ID="radio2">
				</td>
			</tr>
			<% if intReporting_Period_ID = 2 then %>
			<tr>				
				<td class="TableCell" nowrap>
					<b>Do you intend to enroll your child<br> for the next school year (SY <% = Right(oFunc.SchoolYear,2) & "-" & right(oFunc.SchoolYear +1,2)%>)?</b>
				</td>
				<td class="TableCell">
					Yes<input type=radio  <% = strDiasbled %> id="Radio  <% = strdiasbled %>8" name="bolReEnroll<%= strUniqueName%>" <% if rs("bolReEnroll") then response.write " checked "%> value="1" ID="radio1">
					| No<input type=radio  <% = strDiasbled %> id="Radio  <% = strdiasbled %>9" name="bolReEnroll<%= strUniqueName%>" <% if rs("bolReEnroll") = 0 then response.write " checked "%> value="0" ID="radio2">
					| N/A<input type=radio  <% = strDiasbled %> id="Radio  <% = strdiasbled %>10" name="bolReEnroll<%= strUniqueName%>" <% if rs("bolReEnroll") & "" = "" then response.write " checked "%> value="NULL" ID="radio10">
				</td>
			</tr>
			<% end if %>
			<tr>
				<td class="TableCell" valign="top" >
					<b>If you answered 'No' to any of the above questions please explain. </b>
				</td>
				<td class="svplain10">
					<% if (ucase(session.Contents("strRole")) = "GUARD" or rs("szMaterials_Not_Received") & "" = "") and not bolPrint then %>
					<textarea style="width:100%;"  onKeyDown="jfMaxSize(1000,this);" rows="4" name="szMaterials_Not_Received" wrap=hard ID="Textarea1"><%=rs("szMaterials_Not_Received") %></textarea>
					<% else %>
					<%=rs("szMaterials_Not_Received") %>
					<input type="hidden" name="szMaterials_Not_Received" value="<% = replace(rs("szMaterials_Not_Received")&"", """","")%>" ID="Hidden6">
					<% end if%>
					
				</td>
			</tr>		
			<tr>
				<td class="SubHeader" colspan=2 <% if not bolPrint then %> style="font-size:10pt;" <% end if %> nowrap>
					<b>Do you need training on any of the following ...</b>
				</td>
			</tr>
			<tr>
				<td class="TableCell" nowrap>
					<b>Vendor Personal Service Contracts?</b>
				</td>
				<td class="TableCell">
					Yes<input type=radio  <% = strDiasbled %> id="bolTrain_PSC" name="bolTrain_PSC<%= strUniqueName%>" <% if rs("bolTrain_PSC") then response.write " checked "%> value="1" ID="radio3">
					| No<input type=radio  <% = strDiasbled %> id="Radio  <% = strdiasbled %>3" name="bolTrain_PSC<%= strUniqueName%>" <% if rs("bolTrain_PSC") = 0 then response.write " checked "%> value="0" ID="radio4">
				</td>
			</tr>	
			<tr>
				<td class="TableCell" nowrap>
					<b>Goods and Services (in ILP packet)?</b>
				</td>
				<td class="TableCell">
					Yes<input type=radio  <% = strDiasbled %> id="bolTrain_GS" name="bolTrain_GS<%= strUniqueName%>" <% if rs("bolTrain_GS")  then response.write " checked "%> value="1" ID="radio5">
					| No<input type=radio  <% = strDiasbled %> id="Radio  <% = strdiasbled %>4" name="bolTrain_GS<%= strUniqueName%>" <% if rs("bolTrain_GS") = 0 then response.write " checked "%> value="0" ID="radio6">
				</td>
			</tr>	
			<tr>
				<td class="TableCell" nowrap>
					<b>Reimbursement process?</b>
				</td>
				<td class="TableCell">
					Yes<input type=radio  <% = strDiasbled %> id="bolTrain_Reimburse" name="bolTrain_Reimburse<%= strUniqueName%>" <% if rs("bolTrain_Reimburse")  then response.write " checked "%> value="1" ID="radio7">
					| No<input type=radio  <% = strDiasbled %> id="Radio  <% = strdiasbled %>5" name="bolTrain_Reimburse<%= strUniqueName%>" <% if rs("bolTrain_Reimburse") = 0 then response.write " checked "%> value="0" ID="radio8">
				</td>
			</tr>	
			<tr>
				<td class="TableCell" nowrap>
					<b>Junior/Senior Credit Check; <BR>do you know what classes you need?</b>
				</td>
				<td class="TableCell">
					Yes<input type=radio  <% = strDiasbled %> id="bolTrain_Grad" name="bolTrain_Grad<%= strUniqueName%>" <% if rs("bolTrain_Grad") then response.write " checked "%> value="1" ID="radio9">
					| No<input type=radio  <% = strDiasbled %> id="Radio  <% = strdiasbled %>7" name="bolTrain_Grad<%= strUniqueName%>" <% if rs("bolTrain_Grad") = 0 then response.write " checked "%> value="0" ID="radio10">
					| N/A<input type=radio  <% = strDiasbled %> id="Radio  <% = strdiasbled %>6" name="bolTrain_Grad<%= strUniqueName%>" <% if rs("bolTrain_Grad") & "" = "" then response.write " checked "%> value="NULL" ID="radio10">
				</td>
			</tr>	
			<tr>
				<td class="SubHeader" colspan=2 <% if not bolPrint then %> style="font-size:10pt;" <% end if %>>
					<b>Comments: <% = strStudentName %></b>
				</td>
			</tr>
			<tr>
				<td class="svplain10" valign="top" nowrap>
					<b>Parent Comments</b><br>
					<% if ucase(rs("szRole_Create")) = "GUARD" then 
							dim rsG 
							set rsG = server.CreateObject("ADODB.RECORDSET")
							rsG.CursorLocation = 3
							
							sql = "SELECT     tblGUARDIAN.szFIRST_NAME + ' ' + tblGUARDIAN.szLAST_NAME AS Name " & _ 
									"FROM         tblGUARDIAN INNER JOIN " & _ 
									"                      tascGUARD_USERS ON tblGUARDIAN.intGUARDIAN_ID = tascGUARD_USERS.intGUARDIAN_ID INNER JOIN " & _ 
									"                      tblUsers ON tascGUARD_USERS.szUser_ID = tblUsers.szUser_ID " & _ 
									"WHERE     (tblUsers.szUser_ID = '" & rs("szUser_Create") & "') "
							
							rsG.Open sql,Application("cnnFPCS")' oFunc.FPCScnn
							
							if rsG.RecordCount > 0 then response.Write rsG(0)
							rsG.Close
							set rsG = nothing
							
							response.Write "<BR> " & rs("dtCreate") 
										
					   end if
					%>
					
				</td>
				<td class="svplain10" style="width:100%;" align="left">
					<% if ucase(session.Contents("strRole")) = "GUARD" then %>
					<textarea  style="width:100%;"   onKeyDown="jfMaxSize(2000,this);"  rows="4" name="szParent_Comments" wrap=hard ID="Textarea2"><%=rs("szParent_Comments") %></textarea>
					<% else %>
					<%=rs("szParent_Comments") %>		
					<input type="hidden" name="Parent_Comments" value="<% = replace(rs("szParent_Comments")&"", """","")%>" ID="Hidden7">			
					<% end if%>
				</td>
			</tr>	
			<tr>
				<td class="svplain10" valign="top" nowrap>
					<b>Sponsor Teacher Comments</b><br>
					<% = arInfo(8) & "<BR>" & rs("dtSponsor_Reviewed") %>
				</td>
				<td class="svplain10">
					<% if ucase(session.Contents("strRole")) = "TEACHER" then %>
					<textarea onKeyDown="jfMaxSize(2000,this);"  style="width:100%;"  rows="4" name="SponsorComments" wrap=hard ID="Textarea3"><%=rs("SponsorComments") %></textarea>
					<% else %>
					<%=rs("SponsorComments") %>
					<% end if%>
				</td>
			</tr>	
		</table>
		<% if not bolPrint then%>
		<br><br>
		<% end if %>
<%		
		rs2.Close
		set rs2 = nothing
		if not oFunc.LockYear and not bolPrint then
			 if ucase(session.Contents("strRole")) = "TEACHER" and rs("szUSER_CREATE") & "" <> "" then
			%>
			<input type="button" value="Click Here to Mark as Read and Save Comments" class="NavSave" onclick="jfValidate(this.form);" ID="Button3" NAME="Button2">
			<%else%>
			<input type="button" value="Submit Progress Report" class="btSave" onclick="jfValidate(this.form);" ID="Button2" NAME="Button2">
			<%end if
			if request("SimpleHeader") = "" then
			%>
			<input type="button" value="Cancel (does not save)" class="btSave" onclick="window.location.href='<%=Application.Value("strWebRoot")%>';">
			<%
			end if
		end if
	else
		response.Write "<span class='svplain10'>Currently there are no implemented classes for this student. Go to the student packet to plan and implement classes.</span>"
	end if
	rs.Close
	set rs = nothing
end if
%>
		</td>
	</tr>
</table>
<% 
	end if
	printCount = printCount + 1
next 
if not bolPrint then %>
<input type=hidden name="ilpList" value = "<% = strILPList %>" ID="Hidden3">
<input type="hidden" name="hdnSave" value="" ID="Hidden4">
<input type="hidden" name="hdnNotSameUser" value="<%=strDiasbled%>" ID="Hidden5">
<% end if%>
</form>
<script language="javascript">
	function jfValidate(myForm){
		var ilpList = myForm.ilpList.value;
		var obj;
		var strError = "";
		var bolMissing;
		
		<% if intProgress_Report_ID & "" = "" then %>
		// Check to be sure every course has been given a progress status
		ilpList = ilpList.split(",");
		for (var i = 0;i<ilpList.length;i++){
			eval("obj = document."+myForm.name+".progress" + ilpList[i] + ";");
			bolMissing = true;
			
			for (var j=0;j<obj.length;j++){
				if (obj[j].checked == true) {
					bolMissing = false
				}
			}
			
			if (bolMissing == true) {
				strError = "One or more courses are missing the current progress status.  You must provide the current status for all courses.\n";			
				break;
			}
		}
		
		// verify that the required fields below are populated
		var arFields = new Array("bolHave_Materials","bolVendors_Paid","bolTrain_PSC","bolTrain_GS","bolTrain_Reimburse","bolTrain_Grad");
		var arLabels = new Array("Do you have all materials?","Are your vendors getting paid?","Need training Personal Service Contracts?","Need training Goods/Services?","Need training Reimbursement process?","Need training Jr/Sr Credit check?");
		
		for (i=0;i<arFields.length;i++){
			eval("obj = document."+myForm.name+"." + arFields[i]);
			if (arFields[i] != "bolHave_Materials" && arFields[i] != "bolVendors_Paid" && arFields[i] != "bolTrain_Grad") {
				if (obj[0].checked == false && obj[1].checked == false){
					strError += "You must provide a value for '" + arLabels[i] + "'.\n";				
				}
			}else{
				if (obj[0].checked == false && obj[1].checked == false && obj[2].checked == false){
					strError += "You must provide a value for '" + arLabels[i] + "'.\n";				
				}
			}
		}
		<% end if %>
			
		if (strError != "") {
			alert("They following errors must be fixed before this form can be saved ...\n" + strError);
		}else{
			myForm.hdnSave.value = "insert";
			myForm.submit();
		}
	}
	
	function jfViewILP(ilp_id,class_ID,class_name,cg,vendor,teacherName) {
		var ilpWin;
		var strURL;
		var strILP;
				
		strURL = "<%=Application.Value("strWebRoot")%>forms/ILP/ilpMain.asp?plain=yes&intILP_ID=" + ilp_id + "&intClass_id=" + class_ID;
		strURL += "&szClass_Name=" + class_name;
		strURL += "&intVendor_ID=" + vendor;
		strURL += "&strTeacherName=" + teacherName;
		strURL += "&intContract_Guardian_ID=" + cg;
		ilpWin = window.open(strURL,"ilpWin","width=710,height=500,scrollbars=yes,resizable=yes");
		ilpWin.moveTo(0,0);
		ilpWin.focus();
	}
	
	<% if request("print") <> "" then %>
		if (window.print){
	      window.print()
	    }
	    else {
	      alert("Mac users: please press Apple-P to print this form.\nWindows users: Please press ctrl-P to print this form.")
		}
		<% end if %>
</script>
<%
call oFunc.CloseCN()
set oFunc = nothing
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")

sub vbsInsertProgress()	
	dim bolNoTrans
	bolNoTrans = false
	if intProgress_Report_ID = "" then
		if session.Contents(intStudent_ID & "|" & intReporting_Period_ID & "|" & session.Contents("intSchool_Year"))  then exit sub
		oFunc.BeginTransCN
		dim insert
		insert = "insert into tblProgress_Reports (intStudent_ID, intSchool_Year, " & _
					"intReporting_Period_ID, bolHave_Materials, szMaterials_Not_Received, " & _
					"bolVendors_Paid, bolTrain_PSC, bolTrain_GS, " & _
                    "bolTrain_Reimburse, bolTrain_Grad, "
        
        if intReporting_Period_ID = 2 then
			insert = insert & "bolReEnroll,"
        end if
                   
        if ucase(session.Contents("strRole")) = "GUARD" then
			insert = insert & " szParent_Comments," 
		elseif ucase(session.Contents("strRole")) = "TEACHER" then
			insert = insert & " szSponsor_Comments,dtSponsor_Reviewed," 
		end if
                    
        insert = insert & "dtCREATE, szUSER_CREATE, szRole_Create) " & _
                    " values (" & _
                    intStudent_ID & "," & session.Contents("intSchool_Year") & ", " & _
					intReporting_Period_ID & "," & bolHave_Materials & "," & _
					"'" & oFunc.EscapeTick(szMaterials_Not_Received) & "'," & _
					bolVendors_Paid & "," & bolTrain_PSC & "," & bolTrain_GS & "," & _
                    bolTrain_Reimburse & "," & bolTrain_Grad & "," 
         
        if intReporting_Period_ID = 2 then
			insert = insert & bolReEnroll & ","
        end if
                   
		if ucase(session.Contents("strRole")) = "GUARD" then
			insert = insert & "'" & oFunc.EscapeTick(szParent_Comments) & "'," 
		elseif ucase(session.Contents("strRole")) = "TEACHER" then
			insert = insert & "'" & oFunc.EscapeTick(SponsorComments) & "',convert(datetime,'" & now() & "')," 
		end if
        
        insert = insert & "convert(datetime,'" & now() & "'),'" &  oFunc.EscapeTick(session.Contents("strUserID")) & "', '" & ucase(session.Contents("strRole")) & "')"
                    'response.Write insert
                    'response.End
        oFunc.ExecuteCn(insert)
        intProgress_Report_ID = oFunc.GetIdentity
        session.Contents(intStudent_ID & "|" & intReporting_Period_ID & "|" & session.Contents("intSchool_Year")) = "lock"
        dim arList
        arList = split(ilpList,",")
        for i =0 to ubound(arList)
			if arList(i) <> "" then
				execute("intProgress_Rating_ID = progress" & arList(i))
				insert = "insert INTO tblCourse_Progress_Ratings " & _
						 " (intProgress_Report_ID, intProgress_Rating_ID, " & _
						 "intILP_ID, dtCREATE, szUSER_CREATE) " & _
						 " values (" & _
						 intProgress_Report_ID & "," & intProgress_Rating_ID & "," &  _
						 arList(i) & "," & _
						 "convert(datetime,'" & now() & "'),'" &  oFunc.EscapeTick(session.Contents("strUserID")) & "')"
				oFunc.ExecuteCN(insert)
			end if
        next        
	else
		oFunc.BeginTransCN
		dim update
		if hdnNotSameUser = "" then
			' perform update on report data only if the user is the same as the one who created the records
			update = " update tblProgress_Reports set " & _
					 "bolHave_Materials = " & bolHave_Materials & "," & _
					 "szMaterials_Not_Received = '" & oFunc.EscapeTick(szMaterials_Not_Received) & "'," & _
					 "bolVendors_Paid = " & bolVendors_Paid & "," & _
					 "bolTrain_PSC = " & bolTrain_PSC & "," & _
					 "bolTrain_Reimburse = " & bolTrain_Reimburse & "," & _
					 "bolTrain_GS = " & bolTrain_GS & "," & _					 
					 "bolTrain_Grad = " & bolTrain_Grad &  "," 
			
			if intReporting_Period_ID = 2 then
				update = update & "bolReEnroll = " & bolReEnroll & "," 
			end if	 
			
			if ucase(session.Contents("strRole")) = "GUARD" then
					update = update & " szParent_Comments = '" & oFunc.EscapeTick(szParent_Comments) & "'," 
			elseif ucase(session.Contents("strRole")) = "TEACHER" then
					update = update & " szSponsor_Comments = '" & oFunc.EscapeTick(SponsorComments) & "',"  & _
							" dtSponsor_Reviewed = convert(datetime,'" & now() & "')," 
			end if
		
			update = update & "szUser_Modify = '" &  oFunc.EscapeTick(session.Contents("strUserID")) & "'," & _
					 " dtModify = convert(datetime,'" & now() & "') " & _
					 " WHERE intProgress_Report_ID = " & intProgress_Report_ID
				'response.Write update		
			oFunc.ExecuteCN(update)
			
			arList = split(ilpList,",")
			for i =0 to ubound(arList)
				if arList(i) <> "" then
					execute("intProgress_Rating_ID = progress" & arList(i))

					update = "update tblCourse_Progress_Ratings set " & _
							 " intProgress_Rating_ID = " & intProgress_Rating_ID & _
							 " , szUser_Modify = '" &  oFunc.EscapeTick(session.Contents("strUserID")) & "'," & _
							 " dtModify = convert(datetime,'" & now() & "') " & _
							 " WHERE intILP_ID = " & arList(i) & _
							 " AND intProgress_Report_ID = " & intProgress_Report_ID 
					oFunc.ExecuteCN(update)
				end if
			next
		else
			if ucase(session.Contents("strRole")) = "GUARD" then
					update = "update tblProgress_Reports set " & _
							 " szParent_Comments = '" & oFunc.EscapeTick(szParent_Comments) & "' " & _
							 " , szUser_Modify = '" &  oFunc.EscapeTick(session.Contents("strUserID")) & "'," & _
							 " dtModify = convert(datetime,'" & now() & "') " & _
							 " WHERE intProgress_Report_ID = " & intProgress_Report_ID
				oFunc.ExecuteCN(update)
			elseif ucase(session.Contents("strRole")) = "TEACHER" then
					update = "update tblProgress_Reports set " & _
							 " szSponsor_Comments = '" & oFunc.EscapeTick(SponsorComments) & "', " & _
							 " dtSponsor_Reviewed = CURRENT_TIMESTAMP, " & _
							 " szUser_Modify = '" &  oFunc.EscapeTick(session.Contents("strUserID")) & "'," & _
							 " dtModify = convert(datetime,'" & now() & "') " & _
							 " WHERE intProgress_Report_ID = " & intProgress_Report_ID	
					oFunc.ExecuteCN(update)		
			else
				bolNoTrans = true		
			end if			
		end if		
	end if
	
	if not bolNoTrans then
		oFunc.CommitTransCN
	end if
		
	intStudent_ID = request("intStudent_ID")
	arInfo = oFunc.StudentInfo(intStudent_ID,8)
	
	if (update <> "") and ucase(session.Contents("strRole")) = "TEACHER" then
		' Alerts Guardian
		arFamInfo = oFunc.FamilyInfo("1",request("intStudent_ID"),"6")
		if arFamInfo(4) <> "" then
			call SendMail(arFamInfo(3),true,arInfo(9)) 
		end if
	elseif (insert <> "" or update <> "") and ucase(session.Contents("strRole")) = "GUARD" then
		' Alerts Teacher
		if arInfo(9) & "" <> "" then
			call SendMail(arInfo(9),false,"info@3shapes.com")
		end if
	end if 
end sub

sub SendMail(pTo,pSendToGuard,pFrom)	
	strStudentName = arInfo(2)
	Set cdoMessage = Server.CreateObject("CDO.Message")
	set cdoConfig = Server.CreateObject("CDO.Configuration")
	cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 ' cdoSendUsingPort
	cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "127.0.0.1"
	cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	cdoConfig.Fields.Update
	set cdoMessage.Configuration = cdoConfig
	
	cdoMessage.From = pFrom
	cdoMessage.Subject = "Progress Report for " & strStudentName
	if pSendToGuard then
		if SponsorComments <> "" then
			SponsorComments = chr(13) & chr(10) & chr(13) & chr(10) & "Sponsor Teacher Comments: " & chr(10) & chr(13) & SponsorComments
		end if 
		if Parent_Comments <> "" then
			Parent_Comments = chr(13) & chr(10) & chr(13) & chr(10) & "Guardian Comments: " & chr(10) & chr(13) & Parent_Comments
		end if		
		
		cdoMessage.TextBody = "This is an automatic email to inform you that your sponsor teacher " & _
							  " has reviewed the Progress Report for " & strStudentName & Parent_Comments & SponsorComments
	else
		cdoMessage.TextBody = "To view the progress report click the link below." & chr(10) & chr(13) & _
								Application("strSSLWebRoot") & "reports/studentProgressReport.asp?intStudent_ID=" & _
								intStudent_ID & "&intReporting_Period_ID=" & intReporting_Period_ID & _
								"&intSchool_Year=" & session.Contents("intSchool_Year") & "&directPath=true"
	end if

	cdoMessage.To =  pTo '"scott@3shapes.com" ' pTo
	cdoMessage.Send
	
	'Clean up Objects
	Set cdoMessage = Nothing 
end sub
%>