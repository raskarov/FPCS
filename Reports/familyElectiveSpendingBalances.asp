<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		familyElectiveSpendingBalances.asp
'Purpose:	Report showing the Available funds LEFT in the families
'			elective spending budget for a given year.
'Date:		03/06/2006
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Server.Execute(Application.Value("strWebRoot") & "includes/Header.asp")
%>
<table style="width:100%;" ID="Table3">
	<tr>
		<td class="yellowHeader" colspan="3">
			&nbsp;<b>Family Elective Spending Balance Report</b>
		</td>
	</tr>	
	<tr>
		<td>
			<form name="main" method="post" action="./familyElectiveSpendingBalances.asp" ID="Form1">						
			<table ID="Table1" cellpadding="3">
				<tr>
					<td class="svplain8"><B>Show all Families</B></td>
					<td>
						<input type="checkbox" name="showAll" value="1" onclick="this.form.submit();"  <% if request("showAll") & "" <> "" then response.Write " checked " %>>						
					</td>
				</tr>
			</table>
			</form>
		</td>
	</tr>
<%

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()
oFunc.ResetSelectSessionVariables
session.Contents("intStudent_ID") = ""
oFunc.OpenCN

dim sql

sql = "SELECT DISTINCT f.szFamily_Name, f.intFamily_ID, f.szDesc, f.szHome_Phone, f.szEMAIL " & _ 
	"FROM	tblFAMILY f INNER JOIN " & _ 
	"	tblSTUDENT s ON f.intFamily_ID = s.intFamily_ID INNER JOIN " & _ 
	"	tblStudent_States ss ON s.intSTUDENT_ID = ss.intStudent_id " & _ 
	"WHERE	(ss.intSchool_Year = " & session.Contents("intSchool_Year") & ") AND (ss.intReEnroll_State IN (" & application.Contents("strEnrollmentList") & ")) " & _ 
	"ORDER BY f.szFamily_Name, f.szDesc "
	
set rs = server.CreateObject("ADODB.RECORDSET")
rs.CursorLocation = 3
rs.Open sql,Application("cnnFPCS")'oFunc.FPCScnn

if rs.RecordCount > 0 then
%>
		<tr>
			<td class="TableHeader">
				<B>Family Name</B>
			</td>
			<td class="TableHeader" align="center">
				<b>Remaining Budget</B>
			</td>
		</tr>
<%	
	do while not rs.EOF
		set oBudget = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/StudentBudgetInfo.wsc"))
		'oBudget.PopulateFamilyBudgetInfo oFunc.FPCScnn, rs("intFamily_ID"), session.Contents("intSchool_Year")
		oBudget.PopulateFamilyBudgetInfo Application("cnnFPCS"), rs("intFamily_ID"), session.Contents("intSchool_Year")
		
		if request("showAll") & "" <> ""  or (oBudget.AvailableElectiveBudget < 0 and _
				request("showAll") & "" = "") then
%>
		<tr>
			<td class="TableCell">
				<% = rs("szFamily_Name") & ", " & rs("szDesc") %>
			</td>
			<td class="TableCell" align="right">
			<% if oBudget.AvailableElectiveBudget >= 0 then 
					response.Write "$" & formatNumber(oBudget.AvailableElectiveBudget,2)
				else
					response.Write "<span class='sverror'>$" & formatNumber(oBudget.AvailableElectiveBudget,2) & "</span>"	
				end if
			%>
			</td>
		</tr>
<%		
		end if
		set oBudget = nothing
		rs.MoveNext
	loop

else ' recordcount = 0
%>
		<tr>
			<td class="svplain10" colspan="3">
				<b>No Students are active for the select school year.</b>
			</td>
		</tr>
<%
end if

rs.Close
set rs = nothing

oFunc.CloseCn
set oFunc = nothing
%>
	</table>
<%
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")
%>