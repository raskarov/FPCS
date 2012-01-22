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

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

'Initialize some key variables
if ucase(session.Contents("strRole")) = "ADMIN" then
	intReporting_Period_ID = request("intReporting_Period_ID")
else
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
Server.Execute(Application.Value("strWebRoot") & "includes/Header.asp")
%>
<form name=main method=post action="adminstudentProgressReport.asp" ID="Form1">
<input type=hidden name=intStudent_ID value="<% = intStudent_ID %>" ID="Hidden1">
<input type=hidden name="studentList" value=",">
<table style="width:100%;" ID="Table3">
	<tr>
		<td class="yellowHeader">
			&nbsp;<b>Student Progress Report Viewer</b>
		</td>
	</tr>
	<tr>
		<td>
			<table ID="Table1">
				<tr>
					<td class="TableHeader" nowrap>
						&nbsp;Select Reporting Period:&nbsp;
					</td>
					<td class="svplain10">
						<select name="intReporting_Period_ID" onchange="this.form.submit();" ID="Select1">
							<option value=""></option>
							<%
								sql = "SELECT intReporting_Period_ID, szReporting_Period_Name " & _ 
										"FROM trefReporting_Periods " & _ 
										"ORDER BY szReporting_Period_Name "
								response.Write oFunc.MakeListSQL(sql,"intReporting_Period_ID","szReporting_Period_Name",intReporting_Period_ID)
							%>							
						</select>
					</td>
					<%if intReporting_Period_ID <> "" then%>
					<td class="TableHeader" nowrap>
						&nbsp;Order By:&nbsp;
					</td>
					<td class="svplain10">
						<select name="strFilter" onchange="this.form.submit();">
							<%
								strFilterList1 = ",tblProgress_Reports.dtCreate,tblProgress_Reports.dtSponsor_Reviewed,bolHave_Materials,bolVendors_Paid"
								strFilterList2 = "Student Name,Date Submitted,Date Sponsor Submitted,Has Materials,Vendors Paid"
								response.Write oFunc.MakeList(strFilterList1,strFilterList2,strFilter)
							%>	
						</select>
					</td>
					<td>
						<input type=button value="Print Selected" onclick="jfPrint();" class="btSmallGray">
					</td>
					<% end if %>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>
		
		
<%
if intReporting_Period_ID <> "" then
	if strFilter <> "" then strFilter =  strFilter & " desc, "
	sql = "SELECT     tblSTUDENT.szLAST_NAME, tblSTUDENT.szFIRST_NAME, tblSTUDENT.intSTUDENT_ID, tblProgress_Reports.intProgress_Report_ID,  " & _ 
			"                      tblProgress_Reports.bolHave_Materials, tblProgress_Reports.szMaterials_Not_Received, tblProgress_Reports.bolVendors_Paid,  " & _ 
			"                      tblProgress_Reports.bolTrain_PSC, tblProgress_Reports.bolTrain_GS, tblProgress_Reports.bolTrain_Reimburse, tblProgress_Reports.bolTrain_Grad,  " & _ 
			"                      tblProgress_Reports.szParent_Comments, tblProgress_Reports.szRole_Create, tblProgress_Reports.szSponsor_Comments,  " & _ 
			"                      tblProgress_Reports.dtSponsor_Reviewed, tblProgress_Reports.dtCREATE " & _ 
			"FROM         tblStudent_States INNER JOIN " & _ 
			"                      tblSTUDENT ON tblStudent_States.intStudent_id = tblSTUDENT.intSTUDENT_ID LEFT OUTER JOIN " & _ 
			"                      tblProgress_Reports ON tblSTUDENT.intSTUDENT_ID = tblProgress_Reports.intStudent_ID AND tblProgress_Reports.intSchool_Year = " & session.Contents("intSchool_Year") & " AND  " & _ 
			"                      tblProgress_Reports.intReporting_Period_ID = " & intReporting_Period_ID & " " & _ 
			"WHERE     (tblStudent_States.intSchool_Year = " & session.Contents("intSchool_Year") & ") AND tblStudent_States.intReEnroll_State  IN (" & Application.Contents("ActiveEnrollList") & ")  " & _ 
			"ORDER BY " & strFilter & " tblSTUDENT.szLAST_NAME, tblSTUDENT.szFIRST_NAME " 
	dim rs 
	set rs = server.CreateObject("ADODB.RECORDSET")
	rs.CursorLocation = 3
	'response.Write "TESTING <BR>" & sql 
	rs.Open sql, oFunc.FPCScnn

	if rs.RecordCount > 0 then	
		
	
%>	
		<table  ID="Table2" cellpadding="3">
			<tr>
				<td class="TableHeader" align="center">
					<b>Student Name/Report Link</b>
				</td>
				<td class="TableHeader"  align="center">
					<b>Date Submitted</b>
				</td>
				<td class="TableHeader"  align="center">
					<b>Date Sponsor Submitted</b>
				</td>	
				<td class="TableHeader"  align="center">
					<b>Has Materials</b>
				</td>			
				<td class="TableHeader"  align="center">
					<b>Vendors Paid</b>
				</td>
				<td class="TableHeader" align="center" >
					<b>Training Requested</b>
				</td>
				<td class="TableHeader" align="center" >
					<b>Print</b>
				</td>
			</tr>
<%		
		do while not rs.EOF
			if rs("intProgress_Report_ID") & "" = "" then
				strNeedTaining = ""
			elseif rs("bolTrain_PSC") or _ 
				rs("bolTrain_GS") or _ 
				rs("bolTrain_Reimburse") or _ 
				rs("bolTrain_Grad") then
				strNeedTaining = "<span style='color:red;'>True</span>"
			else
				strNeedTaining = "False"
			end if
			
			if rs("intProgress_Report_ID") & "" = "" then
				strMaterials = ""
			elseif rs("bolHave_Materials") then 
				strMaterials =  "True" 
			elseif rs("bolHave_Materials") & "" = "" then
				strMaterials = "N/A"
			else 
				strMaterials = "<span style='color:red;'>False</span>"
			end if
			
			if rs("intProgress_Report_ID") & "" = "" then
				strVend = ""
			elseif rs("bolVendors_Paid") then 
				strVend =  "True" 
			elseif rs("bolVendors_Paid") & "" = "" then 
				strVend =  "N/A" 
			else 
				strVend = "<span style='color:red;'>False</span>"
			end if
			
%>
			<tr>
				<td class="TableCell" >
					<a href="javascript:" onclick="jfViewReport('<% = rs("intStudent_ID")%>','<% = intReporting_Period_ID%>');"><% = rs("szLAST_NAME") & ", " & rs("szFIRST_NAME") %></a>&nbsp;
				</td>
				<td class="TableCell"  align="center">
					<% if isDate(rs("dtCREATE")) then response.Write formatdatetime(rs("dtCREATE"),2) %>&nbsp;
				</td>
				<td class="TableCell"  align="center">
					<%  if isDate(rs("dtSponsor_Reviewed")) then response.Write formatdatetime(rs("dtSponsor_Reviewed"),2) %> &nbsp;
				</td>	
				<td class="TableCell"  align="center">
					<% = strMaterials%>&nbsp;
				</td>			
				<td class="TableCell"  align="center">
					<% = strVend %>&nbsp;
				</td>
				<td class="TableCell" align="center">
					<% = strNeedTaining %>&nbsp;
				</td>
				<td class="TableCell" align="center">
					<input type="checkbox" onClick="jfPrintList('<% = rs("intStudent_ID") %>',this);" name="chk<% = rs("intStudent_ID") %>">
				</td>
			</tr>		
				
<%			
			rs.MoveNext
		loop
		response.Write "</table>"
	end if ' end if recordcount > 0
	rs.Close
	set rs = nothing
end if
%>
		</td>
	</tr>
</table>
<input type=hidden name="ilpList" value = "<% = strILPList %>" ID="Hidden3">
<input type="hidden" name="hdnSave" value="" ID="Hidden4">
<input type="hidden" name="hdnNotSameUser" value="<%=strDiasbled%>" ID="Hidden5">
</form>
<script language="javascript">
	function jfViewReport(pStudentID,pPeriodID) {
		var winSPR;
				
		strURL = "<%=Application.Value("strWebRoot")%>Reports/StudentProgressReport.asp?SimpleHeader=true&intStudent_id=" + pStudentID + "&intReporting_Period_ID=" + pPeriodID;
		winSPR = window.open(strURL,"winSPR","width=710,height=500,scrollbars=yes,resizable=yes");
		winSPR.moveTo(0,0);
		winSPR.focus();
	}
	
	function jfPrintList(pID,pObj){
		var sIDs = document.getElementById('studentList');
		sIDs.value = sIDs.value.replace(","+pID+",",",");
		if (pObj.checked){ sIDs.value = sIDs.value + pID + ","; }
		document.main.studentList.value = sIDs.value;
	}
	
	function jfPrint(){
		var winPrint;
		var studentList = document.main.studentList.value;	
		strURL = "<%=Application.Value("strWebRoot")%>Reports/StudentProgressReport.asp?print=true&intStudent_id=" + studentList + "&intReporting_Period_ID=<%=intReporting_Period_ID%>";
		winPrint = window.open(strURL,"winPrint","width=710,height=500,scrollbars=yes,resizable=yes");
		winPrint.moveTo(0,0);
		winPrint.focus();
	}
</script>
<%
call oFunc.CloseCN()
set oFunc = nothing
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")


%>