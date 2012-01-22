<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		budgetWorkSheet.asp
'Purpose:	Tool that allows guardians to add/view/edit/delete costs
'			for planned courses
'Date:		25 March 2003
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim intStudent_ID, intShort_ILP_ID 
dim dblFunds		' contains the remaining funds in a students budget
dim sql
dim mError		'conitains our error messages after validation is complete

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

'Initialize some key variables
if request("intStudent_ID") <> "" then
	intStudent_ID = request("intStudent_ID") 
	intShort_ILP_ID = request("intShort_ILP_ID")	
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
		execute(i & " = """ & request.Form(i) & """")
	next 
end if 


'Validate Budget Transfer form if needed
if btSubmit <> "" then
	mError = vbfValidate()
	if mError = "" then
		call vbsInsertTransfer
		strMessage = "alert('Transfer has been made.');"
		session.Value("simpleOnLoad") = strMessage
	end if
end if

'prevent page from cacheing
Response.AddHeader "Cache-Control","No-Cache"
Response.Expires = -1

'Print the header
Session.Value("strTitle") = "Budget Transfers"
Session.Value("strLastUpdate") = "18 JAN 2003"
Server.Execute(Application.Value("strWebRoot") & "includes/Header.asp")
session.Value("simpleOnLoad") = ""
%>
<form name=main method=post action="budgetTransfer.asp">
<input type=hidden name=intStudent_ID value="<% = intStudent_ID %>">
<input type=hidden name=intShort_ILP_ID value="<% = intShort_ILP_ID %>">
<table width="100%" ID="Table3">
	<tr>
		<td class="yellowHeader">
			&nbsp;<b> Budget Transfers</b> (for the <% = oFunc.FamilyInfo(1,intStudent_ID,2)%> family)
		</td>
	</tr>
<%
if mError <> "" then
%>
	<tr>
		<td>
			<table cellpadding=4 cellspacing=0 border=1 bordercolor=c0c0c0>
				<tr>
					<td class=error10>
						<B>The following errors must be fixed in order to submit your request ...</B><br>
						<% = mError%>
					</td>
				</tr>
			</table>
		</td>		
	</tr>
<%
end if 
%>
	<tr>
		<td>
			<table cellpadding='4'>
<%
' Get all students in the family
intFamily_ID = oFunc.FamilyInfo(1,intStudent_ID,1)		

set rsStudents = server.CreateObject("ADODB.RECORDSET")
rsStudents.CursorLocation = 3

if ucase(session.Contents("strRole")) = "ADMIN" then
	sql = "SELECT s.intSTUDENT_ID, s.szFIRST_NAME + ' ' + s.szLAST_NAME as Name " & _
		"FROM tblSTUDENT s INNER JOIN " & _
		" tblStudent_States ss ON s.intSTUDENT_ID = ss.intStudent_id " & _
		"WHERE (s.intFamily_ID = " & intFamily_ID & ") " & _
		" AND (ss.intSchool_Year = " & session.Contents("intSchool_Year") & _
		") " & _
		"ORDER BY s.szFIRST_NAME "
else
	sql = "SELECT s.intSTUDENT_ID, s.szFIRST_NAME + ' ' + s.szLAST_NAME as Name " & _
			"FROM tblSTUDENT s INNER JOIN " & _
			" tblStudent_States ss ON s.intSTUDENT_ID = ss.intStudent_id " & _
			"WHERE (s.intFamily_ID = " & intFamily_ID & ") " & _
			" AND (ss.intSchool_Year = " & session.Contents("intSchool_Year") & _
			") AND (ss.intReEnroll_State = 7 OR " & _
			" ss.intReEnroll_State = 15 OR " & _
			" ss.intReEnroll_State = 31) " & _
			"ORDER BY s.szFIRST_NAME "
end if
rsStudents.Open sql, oFunc.FPCScnn

if rsStudents.RecordCount = 1 then
%>			
				<tr>
					<td class=gray>
						This family only has one student
						actively enrolled in FPCS for school year <% = oFunc.SchoolYearRange %>.
						A family must have at least two actively enrolled students
						in order to make a budget transfer.<br><br>
						<input type=button value="Return to Budget Worksheet" class="btSmallGray" onclick="window.location.href='budgetWorkSheet.asp?intStudent_ID=<%=intStudent_ID%>&intShort_ILP_ID=<%=intShort_ILP_ID%>';">
					</td>
				</tr>
<%
elseif rsStudents.RecordCount < 1 then
%>			
				<tr>
					<td class=gray>
						No active students have been found.  The student you are working with may have withdrawn from FPCS. Contact the FPCS office for more information.<br><br>
						<input type=button value="Return to Budget Worksheet" class="btSmallGray" onclick="window.location.href='budgetWorkSheet.asp?intStudent_ID=<%=intStudent_ID%>&intShort_ILP_ID=<%=intShort_ILP_ID%>';" ID="Button1" NAME="Button1">
					</td>
				</tr>
<%
else
	' Create table with student names and remaining budget funds
%>
				<tr>
					<td class="TableHeader">
						&nbsp;Students in Family&nbsp;
					</td>
					<td class="TableHeader">
						&nbsp;Available Budget&nbsp;
					</td>
				</tr>
<%
	do while not rsStudents.EOF
		arFund = oFunc.GetStudentBalances(rsStudents("intStudent_ID"))
%>
				<tr>
					<td class="TableCell">
						&nbsp;<% = rsStudents("Name") %>&nbsp;
					</td>
					<td class="TableCell" align=right>
						$<% = formatNumber(arFund(0),2) %>
					</td>
				</tr>

<%	
		rsStudents.MoveNext
	loop
		if strMessage <> "" then
			' reset form
			intFrom_Student_ID = ""
			intTo_Student_ID = ""
			curAmount = ""
		end if
		' Now create the budget transfer form
%>
			
				<tr>
					<td class=gray>
						&nbsp;Transfer Funds <b>From</b> Student:&nbsp; 
					</td>
					<td class="TableCell">
						<select name="intFrom_Student_ID">
							<option></option>
							<%
								rsStudents.MoveFirst
								response.Write oFunc.MakeListRS(rsStudents,"intSTUDENT_ID","Name",intFrom_Student_ID)
							%>
						</select>
					</td>					
				</tr>
				<tr>
					<td class=gray>
						&nbsp;Transfer Funds <b>To</b> Student:&nbsp;
					</td>
					<td class="TableCell">
						<select name="intTo_Student_ID" ID="Select1">
							<option></option>
							<%
								rsStudents.MoveFirst
								response.Write oFunc.MakeListRS(rsStudents,"intSTUDENT_ID","Name",intTo_Student_ID)
							%>
						</select>
					</td>					
				</tr>
				<tr>
					<td class=gray align=right>
						Transfer Amount: $ 
					</td>
					<td class="TableCell">
						<input type=text name="curAmount" value="<% = curAmount %>" size=7 maxlength=7>
					</td>					
				</tr>
				<tr>
					<td></td>
					<td>
						<input type=submit value="Transfer Funds" name="btSubmit" class="NavSave">
					</td>
				</tr>
<%
end if 

rsStudents.Close
set rsStudents = nothing
%>
			</table>
		</td>
	</tr>
</table>
</form>
<%
call oFunc.CloseCN()
set oFunc = nothing
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")

function vbfValidate()
	dim strError
	' Validate the transfer budget form
	if intFrom_Student_ID = "" or intTo_Student_ID = "" then
		strError = "Neither 'Transfer Funds From Student' or 'Transfer Funds To Student' fields can be blank.<br>"
	elseif intFrom_Student_ID = intTo_Student_ID then
		strError = strError & "The 'Transfer Funds From Student' field can not be the same as the 'Transfer Funds To Student' field.<br>"		
	end if
	
	
	
	
	
	if curAmount = "" then 
		strError = strError & "Transfer Amount can not be blank. <BR>"
	else
		if intFrom_Student_ID <> "" then
			' This next line will give us the remaining budget of the student who the user
			' wants to transfer finds FROM 	
			arFund2 = oFunc.GetStudentBalances(intFrom_Student_ID)
			fromBudget = arFund2(0)
		end if
		
		if not isNumeric(curAmount) then
			strError = strError & "Transfer Amount must be a valid number.<BR>"
		elseif cdbl(fromBudget) - cdbl(curAmount) < 0 then
			strError = strError & "You have requested to transfer more than the available funds in the student account.<BR>"
		end if
	end if
	
	vbfValidate = strError
end function

sub vbsInsertTransfer
	dim insert
	
	insert = "insert into tblBudget_Transfers (" & _
			 "intFrom_Student_ID, intTo_Student_ID, " & _
			 "curAmount, intSchool_Year, dtCREATE, szUSER_CREATE) " & _
			 " values (" & _
			 intFrom_Student_ID & "," & _
			 intTo_Student_ID & "," & _
			 curAmount & "," & _
			 session.Contents("intSchool_Year") & "," & _
			 "'" & now() & "'," & _
			 "'" & session.Contents("strUserID") & "')"
	oFunc.ExecuteCN(insert)
	
end sub
%>