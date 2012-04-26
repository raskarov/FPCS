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
dim SBA 

SBA = Application.Contents("SchoolBudgetAccount")
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

	if not oFunc.IsAdmin then
		sWhere = " AND NOT EXISTS (Select 'x' from STUDENT_LOCKED_ACCOUNTS sl " & _
								 " where ss.intStudent_ID = sl.StudentID and SchoolYear = " & session.Contents("intSchool_Year") & ") "
	end if
	
	sql = "SELECT s.intSTUDENT_ID, s.szFIRST_NAME + ' ' + s.szLAST_NAME as Name " & _
			"FROM tblSTUDENT s INNER JOIN " & _
			" tblStudent_States ss ON s.intSTUDENT_ID = ss.intStudent_id " & _
			"WHERE (s.intFamily_ID = " & intFamily_ID & ") " & _
			" AND (ss.intSchool_Year = " & session.Contents("intSchool_Year") & _
			") AND (ss.intReEnroll_State = 7 OR " & _
			" ss.intReEnroll_State = 15 OR " & _
			" ss.intReEnroll_State = 31) " & sWhere & _
			"ORDER BY s.szFIRST_NAME "
			
rsStudents.Open sql, Application("cnnFPCS")'oFunc.FPCScnn

if rsStudents.RecordCount = 1 and not oFUnc.IsAdmin then
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
	set oBudget = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/StudentBudgetInfo.wsc"))
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
		'arFund = oFunc.GetStudentBalances(rsStudents("intStudent_ID"))
		oBudget.PopulateStudentFunding Application("cnnFPCS"),rsStudents("intStudent_ID"),session.Contents("intSchool_Year")
		'oBudget.PopulateStudentFunding oFunc.FPCScnn,rsStudents("intStudent_ID"),session.Contents("intSchool_Year")
%>
				<tr>
					<td class="TableCell">
						&nbsp;<% = rsStudents("Name") %>&nbsp;
					</td>
					<td class="TableCell" align=right>
						$<% = formatNumber(oBudget.BudgetBalance,2) %>
					</td>
				</tr>

<%	
		rsStudents.MoveNext
	loop
	
	set oBudget = nothing
	
		if strMessage <> "" then
			' reset form
			intFrom_Student_ID = ""
			intTo_Student_ID = ""
			curAmount = ""
		end if
		' Now create the budget transfer form

		' first check to see if year is locked
		if not oFunc.LockYear and not oFunc.LockSpending then

%>
			
				<tr>
					<td class=gray>
						&nbsp;Transfer Funds <b>From</b> Student:&nbsp; 
					</td>
					<td class="TableCell" align="left">
						<select name="intFrom_Student_ID">
							<option></option>
							<% if oFunc.IsAdmin then %>
							<option value="<%=SBA%>" <% if intFrom_Student_ID = SBA then response.Write " selected " %>>FPCS Account</option>
							<% end if %>
							<%
								rsStudents.MoveFirst
								response.Write oFunc.MakeListRS(rsStudents,"intSTUDENT_ID","Name",intFrom_Student_ID)
							%>
						</select>
					</td>					
				</tr>				
				<tr>
					<td class=gray align="right">
						&nbsp;Transfer Funds <b>To</b> Student:&nbsp;
					</td>
					<td class="TableCell" align="left">
						<select name="intTo_Student_ID" ID="Select1">
							<option></option>
							<% if oFunc.IsAdmin then %>
							<option value="<% = SBA %>" <% if intFrom_Student_ID = SBA then response.Write " selected " %>>FPCS Account</option>
							<% end if %>
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
					<td class="TableCell"  align="left">
						<input type=text name="curAmount" value="<% = curAmount %>" size=7 maxlength=7>
					</td>					
				</tr>
				<tr>
					<td colspan="2" class=gray>
						Comment
					</td>					
				</tr>
				<tr>
					<td colspan="2" class=gray>
						<input type="text" name="szComment" maxlength="512" style="width:100%;" ID="Text1">
					</td>					
				</tr>
				<tr>
					<td></td>
					<td>
					<%
					' first check to see if year is locked
					if not oFunc.LockYear and not oFunc.LockSpending then
					%>
						<input type=submit value="Transfer Funds" name="btSubmit" class="NavSave">
					<% end if %>
					</td>
				</tr>				
<%
			end if
end if 

rsStudents.Close
set rsStudents = nothing
%>
			</table>
		</td>
	</tr>
</table>
<br><br>
<% call vbsShowStatement %>
</form>
<%
call oFunc.CloseCN()
set oFunc = nothing
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")

function vbfValidate()
	dim strError
	dim oBudget
	set oBudget = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/StudentBudgetInfo.wsc"))
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
			'arFund2 = oFunc.GetStudentBalances(intFrom_Student_ID)
			if oFunc.IsAdmin and intFrom_Student_ID = SBA then
				' Don't validate student since funds are not coming from a studnet account but
				' the school account
			else
				oBudget.PopulateStudentFunding Application("cnnFPCS"),intFrom_Student_ID,session.Contents("intSchool_Year")
				'oBudget.PopulateStudentFunding oFunc.FPCScnn,intFrom_Student_ID,session.Contents("intSchool_Year")
				fromBudget = oBudget.BudgetBalance
				'fromBudget = arFund2(0)
			end if
		end if
		
		if not isNumeric(curAmount) then
			strError = strError & "Transfer Amount must be a valid number.<BR>"
		elseif (cdbl(fromBudget) - cdbl(curAmount) < 0) and not (oFunc.IsAdmin and intFrom_Student_ID = SBA) then
			strError = strError & "You have requested to transfer more than the available funds in the student account.<BR>"
		end if
		
		if szComment & "" = "" then
			strError = strError & "You must provide a comment.<BR>"
		end if
	end if
	
	vbfValidate = strError
end function

sub vbsInsertTransfer
	dim insert
	
	insert = "insert into tblBudget_Transfers (" & _
			 "intFrom_Student_ID, intTo_Student_ID, " & _
			 "curAmount, intSchool_Year, dtCREATE, szUSER_CREATE, szComment) " & _
			 " values (" & _
			 intFrom_Student_ID & "," & _
			 intTo_Student_ID & "," & _
			 curAmount & "," & _
			 session.Contents("intSchool_Year") & "," & _
			 "'" & now() & "'," & _
			 "'" & session.Contents("strUserID") & "'," & _
			 "'" & oFunc.EscapeTick(szComment) & "')"
	oFunc.ExecuteCN(insert)

end sub

sub vbsShowStatement
	dim sql, rs2
	
	sql = "SELECT     s2.szLAST_NAME + ', ' + s2.szFIRST_NAME AS From_Student,  " & _ 
			"s1.szLAST_NAME + ', ' + s1.szFIRST_NAME AS To_Student, b.curAmount,  " & _ 
			"	b.dtCREATE, b.szComment, b.szUser_Create, b.intFrom_Student_ID, b.intTo_Student_ID " & _ 
			"FROM tblSTUDENT s1 INNER JOIN " & _ 
			"	tblBudget_Transfers b ON s1.intSTUDENT_ID = b.intTo_Student_ID INNER JOIN " & _ 
			"	tblSTUDENT s2 ON b.intFrom_Student_ID = s2.intSTUDENT_ID " & _ 
			"WHERE (b.intSchool_Year = " & session.Contents("intSchool_Year")& ") AND (s2.intFamily_ID = " & intFamily_ID & " OR " & _ 
			"	s1.intFamily_ID = " & intFamily_ID & ") " & _ 
			"ORDER BY b.dtCREATE "
	
	set rs2 = server.CreateObject("ADODB.RECORDSET")
	rs2.CursorLocation = 3
	rs2.Open sql, Application("cnnFPCS")'oFunc.FpcsCnn
	
	if rs2.RecordCount > 0 then
%>
		<span class="svplain10">&nbsp;<b>Budget Transfer History</b></span>
		<table cellpadding='3'>
			<tr>
				<td class="TableHeader">Date</td>
				<td class="TableHeader">Comments</td>
				<td class="TableHeader">Created By</td>
				<td class="TableHeader">From Student</td>
				<td class="TableHeader">To Student</td>
				<td class="TableHeader">Amount</td>
			</tr>
<%		
		dim FpcsTo, FpcsFrom
		FpcsTo = 0
		FpcsFrom = 0
		do while not rs2.EOF
%>
			<tr>
				<td class="TableCell">
					<% = formatDateTime(rs2("dtCREATE"),2) %>
				</td>
				<td class="TableCell">
					&nbsp;<% = rs2("szComment") %>
				</td>	
				<td class="TableCell">
					<% = rs2("szUser_Create") %>
				</td>
				<td class="TableCell">
					<% = rs2("From_Student") %>
				</td>
				<td class="TableCell">
					<% = rs2("To_Student") %>
				</td>
				<td class="TableCell" align="right">
					$<% = formatNUmber(rs2("curAmount"),2) %>
				</td>
			</tr>
<%			
			if SBA & ""  = cstr(rs2("intFrom_Student_ID")) & "" then
				FpcsFrom = FpcsFrom + cdbl(rs2("curAmount"))
			end if
			
			if SBA & "" = cstr(rs2("intTo_Student_ID")) & "" then
				FpcsTo = FpcsTo + cdbl(rs2("curAmount"))
			end if
			
			rs2.MoveNext						
		loop
		
		if FpcsFrom > 0 and FpcsTo > 0  and oFunc.IsAdmin then
%>		
		<tr>
			<td colspan=3 class="TableHeader" align="right">
				FPCS Account Totals:
			</td>
			<td class="TableHeader">
				Withdrawn
			</td>
			<td class="TableHeader">
				Deposited
			</td>
			<td class="TableHeader">
				Net
			</td>
		</tr>
		<tr>
			<td colspan=3 class="TableCell" align="right">
				&nbsp;
			</td>
			<td class="TableCell" align="right">
				$<% = formatNumber(FpcsFrom,2) %>
			</td>
			<td class="TableCell" align="right">
				$<% = formatNumber(FpcsTo,2) %>
			</td>
			<td class="TableCell" align="right">
				$<% = formatNumber(FpcsTo - FpcsFrom,2) %>
			</td>
		</tr>
<%		
		end if
%>
		</table>
<%
	end if
	
	rs2.Close
	
	if false then
		sql = "SELECT a.Withdraw, b.Deposit, b.Deposit - a.Withdraw AS total " & _ 
				"FROM (SELECT     SUM(curAmount) AS Withdraw " & _ 
				"	FROM	tblBudget_Transfers " & _ 
				"	WHERE	(intFrom_Student_ID = " & SBA & ") " & _ 
				"	GROUP BY intSchool_Year " & _ 
				"	HAVING      (intSchool_Year = " & session.Contents("intSchool_Year")& ")) a CROSS JOIN " & _ 
				"	(SELECT     SUM(curAmount) AS Deposit " & _ 
				"	FROM	tblBudget_Transfers " & _ 
				"	WHERE (intTo_Student_ID = " & SBA & ") " & _ 
				"	GROUP BY intSchool_Year " & _ 
				"	HAVING (intSchool_Year = " & session.Contents("intSchool_Year")& ")) b "

		rs2.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
		
		if rs2.RecordCount > 0 then
%>
<BR><BR>
<span class="svplain10">&nbsp;<b>FPCS Account Summary</b> <BR>(includes all transactions for all students for <% = oFunc.SchoolYearRange %> school year.)</span>
<table cellpadding="3">
	<tr>
		<td class="TableHeader">Withdrawls</td>
		<td class="TableHeader">Deposits</td>
		<td class="TableHeader">Net</td>
	</tr>
	<tr>
		<td class="TableCell" align="right">
			$-<% = formatNUmber(rs2("Withdraw"),2) %>
		</td>
		<td class="TableCell" align="right">
			$<% = formatNUmber(rs2("Deposit"),2) %>
		</td>
		<td class="TableCell" align="right">
			$<% = formatNUmber(rs2("total"),2) %>
		</td>
	</tr>
</table>
<%			
		end if
		rs2.Close		
	end if	
	set rs2 = nothing
end sub
%>