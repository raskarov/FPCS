<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		studentEnrollmentPercentages.asp
'Purpose:	Creates a list of students and displays there enrollment
'			percentages.
'Date:		7 Jan 2002
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'prevent page from cacheing
Response.AddHeader "Cache-Control","No-Cache"
Response.Expires = -1
server.ScriptTimeout = 10000

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Dimension Variables, make db Connection.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim sqlStudents
dim intPercentage
dim intBlankProjected
dim intProjected0,intProjected25,intProjected50,intProjected75,intProjected100
dim elementry0Act,elementry25Act,elementry50Act,elementry75Act,elementry100Act
dim middle0Act,middle25Act,middle50Act,middle75Act,middle100Act
dim high0Act,high25Act,high50Act,high75Act,high100Act

dim elementry0Proj,elementry25Proj,elementry50Proj,elementry75Proj,elementry100Proj
dim middle0Proj,middle25Proj,middle50Proj,middle75Proj,middle100Proj
dim high0Proj,high25Proj,high50Proj,high75Proj,high100Proj
dim int0,int25,int50,int75,int100
dim intCount

dim oFunc		'wsc object
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

' Get all needed student info
sqlStudents = "SELECT s.intSTUDENT_ID, s.szLAST_NAME + ', ' + s.szFIRST_NAME AS Name, " & _ 
			" tblENROLL_INFO.intPercent_Enrolled_Fpcs, ss.szGrade "& _
			"FROM tblSTUDENT s INNER JOIN " & _ 
			"tblStudent_States ss ON s.intSTUDENT_ID = ss.intStudent_id " & _ 
			"LEFT OUTER JOIN " & _
            " tblENROLL_INFO ON s.intSTUDENT_ID = tblENROLL_INFO.intSTUDENT_ID " & _
            " AND (tblENROLL_INFO.sintSCHOOL_YEAR = " & Session.Value("intSchool_Year") & ") " & _
			"WHERE ss.intReEnroll_State IN (" & Application.Contents("ActiveEnrollList") & ") AND (ss.intSchool_Year = " & Session.Value("intSchool_Year") & ") " & _ 			
			"ORDER BY s.szLast_Name" 
 
set rsStudents = server.CreateObject("ADODB.RECORDSET")
rsStudents.CursorLocation = 3
rsStudents.Open sqlStudents, oFunc.FPCScnn

' Print the HTML header
Session.Value("strTitle") = "Student Enrollment Percentages"
Session.Value("strLastUpdate") = "03 June 2002"
Server.Execute(Application.Value("strWebRoot") & "includes/Header.asp")

' Print the Report Header
%>
<table width=100%>
	<tr>	
		<Td class=yellowHeader>
				&nbsp;<b>Student Enrollment Percentages</b> (<%=rsStudents.RecordCount%> active students)
		</td>
	</tr>
	<tr>
</table>
<table cellpadding=2 cellspacing=1 >
<%

' Start printing the stundent enrollment information
do while not rsStudents.EOF
	'Find out what the actual enrollment % is for a specific student
	intPercentage = cint(oFunc.StudentPercentage(rsStudents("intStudent_ID")))
	arStudentEnroll = oFunc.arStudentEnroll
	'reprint table header after every 40 rows 	
	if intCount mod 40 = 0 then		
		if intCount > 0 then
			response.Write "<tr><td colspan=11><p></td></tr>"
		end if 
		call vbfPrintHeaders
	end if 
	
	response.Write "<tr><td class='TableCell'>&nbsp;" & rsStudents("Name") & "</td>"
	response.Write "<td class='TableCell' align=center>&nbsp;" & rsStudents("szGrade") & "</td>"	
	response.Write "<td class='TableCell' align=center>&nbsp;" & rsStudents("intPercent_Enrolled_Fpcs") & "%</td>"
	response.Write "<td class='TableCell' align=center>&nbsp;" & formatNumber(arStudentEnroll(0)/90,1) & "</td>"
	response.Write "<td class='TableCell' align=center>&nbsp;" & formatNumber(arStudentEnroll(1)/90,1) & "</td>"
	response.Write "<td class='TableCell' align=center>&nbsp;" & arStudentEnroll(2) & "</td>"
	select case intPercentage
		case 0 
			response.Write "<td class='TableCell' align=center>X</td><td class='TableCell'>&nbsp;</td><td class='TableCell'>&nbsp;</td><td class='TableCell'>&nbsp;</td><td class='TableCell'>&nbsp;</td>"
			int0 = int0 + 1
			call vbsFundingCategoryCalc(rsStudents("szGrade"),0,"Act")
		case 25
			response.Write "<td class='TableCell'>&nbsp;</td><td class='TableCell' align=center>X</td><td class='TableCell'>&nbsp;</td><td class='TableCell'>&nbsp;</td><td class='TableCell'>&nbsp;</td>"			
			int25 = int25 + 1
			call vbsFundingCategoryCalc(rsStudents("szGrade"),25,"Act")
		case 50
			response.Write "<td class='TableCell'>&nbsp;</td><td class='TableCell'>&nbsp;</td><td class='TableCell' align=center>X</td><td class='TableCell'>&nbsp;</td><td class='TableCell'>&nbsp;</td>"
			int50 = int50 + 1
			call vbsFundingCategoryCalc(rsStudents("szGrade"),50,"Act")
		case 75
			response.Write "<td class='TableCell'>&nbsp;</td><td class='TableCell'>&nbsp;</td><td class='TableCell'>&nbsp;</td><td class='TableCell' align=center>X</td><td class='TableCell'>&nbsp;</td>"
			int75 = int75 + 1
			call vbsFundingCategoryCalc(rsStudents("szGrade"),75,"Act")
		case 100
			response.Write "<td class='TableCell'>&nbsp;</td><td class='TableCell'>&nbsp;</td><td class='TableCell'>&nbsp;</td><td class='TableCell'>&nbsp;</td><td class='TableCell' align=center>X</td>"
			int100 = int100 + 1
			call vbsFundingCategoryCalc(rsStudents("szGrade"),100,"Act")
	end select
	
	select case rsStudents("intPercent_Enrolled_Fpcs")
		case 0 
			intProjected0 = intProjected0 + 1
			call vbsFundingCategoryCalc(rsStudents("szGrade"),0,"Proj")
		case 25
			intProjected25 = intProjected25 + 1
			call vbsFundingCategoryCalc(rsStudents("szGrade"),25,"Proj")
		case 50
			intProjected50 = intProjected50 + 1
			call vbsFundingCategoryCalc(rsStudents("szGrade"),50,"Proj")
		case 75
			intProjected75 = intProjected75 + 1
			call vbsFundingCategoryCalc(rsStudents("szGrade"),75,"Proj")
		case 100
			intProjected100 = intProjected100 + 1
			call vbsFundingCategoryCalc(rsStudents("szGrade"),100,"Proj")
		case else
			intBlankProjected = intBlankProjected + 1
	end select
	rsStudents.MoveNext
	response.Write "</tr>"
	intCount = intCount + 1
loop

'Print report footer info
%>
</table>
<p>
<table>
	<tr>
		<td class='TableCell'>
			<b>&nbsp;Actual Enrollment&nbsp;</b>
		</td>
		<td class='TableCell'>
			<b>&nbsp;0%</b>
		</td>
		<td class='TableCell'>
			<b>&nbsp;25%</b>
		</td>
		<td class='TableCell'>
			<b>&nbsp;50%</b>
		</td>
		<td class='TableCell'>
			<b>&nbsp;75%</b>
		</td>
		<td class='TableCell'>
			<b>&nbsp;100%</b>
		</td>
		<td class='TableCell'>
			<b>&nbsp;Totals&nbsp;</b>
		</td>
	</tr>
	<tr>
		<td class='TableCell' align=right>
			K to 5th Grade Totals</b>&nbsp;
		</td>
		<td class='TableCell' align=center>
			<% = elementry0Act %>
		</td>
		<td class='TableCell' align=center>
			<% = elementry25Act %>
		</td>
		<td class='TableCell' align=center>
			<% = elementry50Act %>
		</td>
		<td class='TableCell' align=center>
			<% = elementry75Act %>
		</td>
		<td class='TableCell' align=center>
			<% = elementry100Act %>
		</td>
		<td class='TableCell' align=center>
			<b><% = elementry0Act + elementry25Act + elementry50Act + elementry75Act + elementry100Act%></b>
		</td>
	</tr>
	<tr>
		<td class='TableCell' align=right>
			6th to 8th Grade Totals</b>&nbsp;
		</td>
		<td class='TableCell' align=center>
			<% = middle0Act %>
		</td>
		<td class='TableCell' align=center>
			<% = middle25Act %>
		</td>
		<td class='TableCell' align=center>
			<% = middle50Act %>
		</td>
		<td class='TableCell' align=center>
			<% = middle75Act %>
		</td>
		<td class='TableCell' align=center>
			<% = middle100Act %>
		</td>
		<td class='TableCell' align=center>
			<b><% = middle0Act + middle25Act + middle50Act + middle75Act + middle100Act%></b>
		</td>
	</tr>
	<tr>
		<td class='TableCell' align=right>
			9th to 12th Grade Totals</b>&nbsp;
		</td>
		<td class='TableCell' align=center>
			<% = high0Act %>
		</td>
		<td class='TableCell' align=center>
			<% = high25Act %>
		</td>
		<td class='TableCell' align=center>
			<% = high50Act %>
		</td>
		<td class='TableCell' align=center>
			<% = high75Act %>
		</td>
		<td class='TableCell' align=center>
			<% = high100Act %>
		</td>
		<td class='TableCell' align=center>
			<b><% = high0Act + high25Act + high50Act + high75Act + high100Act%></b>
		</td>
	</tr> 
	<tr>
		<td class='TableCell' align=right >
			<b>Actual Totals</b>&nbsp;
		</td>
		<td class='TableCell' align=center>
			<b><% = int0 %></b>
		</td>
		<td class='TableCell' align=center>
			<b><% = int25 %></b>
		</td>
		<td class='TableCell' align=center>
			<b><% = int50 %></b>
		</td>
		<td class='TableCell' align=center>
			<b><% = int75 %></b>
		</td>
		<td class='TableCell' align=center>
			<b><% = int100 %></b>
		</td>
		<td class='TableCell' align=center>
			<b><% = int0 + int25 + int50 + int75 + int100%></b>
		</td>
	</tr>
	<tr>
		<td colspan=6>
			&nbsp;
		</td>
	</tr>
	<tr>
		<td class='TableCell'>
			<b>&nbsp;Projected Enrollment&nbsp;</b>
		</td>
		<td class='TableCell'>
			<b>&nbsp;0%</b>
		</td>
		<td class='TableCell'>
			<b>&nbsp;25%</b>
		</td>
		<td class='TableCell'>
			<b>&nbsp;50%</b>
		</td>
		<td class='TableCell'>
			<b>&nbsp;75%</b>
		</td>
		<td class='TableCell'>
			<b>&nbsp;100%</b>
		</td>
		<td class='TableCell'>
			<b>&nbsp;Totals&nbsp;</b>
		</td>
	</tr>
	<tr>
		<td class='TableCell' align=right>
			K to 5th Grade Totals</b>&nbsp;
		</td>
		<td class='TableCell' align=center>
			<% = elementry0Proj %>
		</td>
		<td class='TableCell' align=center>
			<% = elementry25Proj %>
		</td>
		<td class='TableCell' align=center>
			<% = elementry50Proj %>
		</td>
		<td class='TableCell' align=center>
			<% = elementry75Proj %>
		</td>
		<td class='TableCell' align=center>
			<% = elementry100Proj %>
		</td>
		<td class='TableCell' align=center>
			<b><% = elementry0Proj + elementry25Proj + elementry50Proj + elementry75Proj + elementry100Proj%></b>
		</td>
	</tr>
	<tr>
		<td class='TableCell' align=right>
			6th to 8th Grade Totals</b>&nbsp;
		</td>
		<td class='TableCell' align=center>
			<% = middle0Proj %>
		</td>
		<td class='TableCell' align=center>
			<% = middle25Proj %>
		</td>
		<td class='TableCell' align=center>
			<% = middle50Proj %>
		</td>
		<td class='TableCell' align=center>
			<% = middle75Proj %>
		</td>
		<td class='TableCell' align=center>
			<% = middle100Proj %>
		</td>
		<td class='TableCell' align=center>
			<b><% = middle0Proj + middle25Proj + middle50Proj + middle75Proj + middle100Proj%></b>
		</td>
	</tr>
	<tr>
		<td class='TableCell' align=right>
			9th to 12th Grade Totals</b>&nbsp;
		</td>
		<td class='TableCell' align=center>
			<% = high0Proj %>
		</td>
		<td class='TableCell' align=center>
			<% = high25Proj %>
		</td>
		<td class='TableCell' align=center>
			<% = high50Proj %>
		</td>
		<td class='TableCell' align=center>
			<% = high75Proj %>
		</td>
		<td class='TableCell' align=center>
			<% = high100Proj %>
		</td>
		<td class='TableCell' align=center>
			<b><% = high0Proj + high25Proj + high50Proj + high75Proj + high100Proj%></b>
		</td>
	</tr> 
	<tr>
		<td class='TableCell' align=right>
			(<% = intBlankProjected %> blank projections) <b>Projected Totals</b>&nbsp;
		</td>
		<td class='TableCell' align=center>
			<b><% = intProjected0 %></b>
		</td>
		<td class='TableCell' align=center>
			<b><% = intProjected25 %></b>
		</td>
		<td class='TableCell' align=center>
			<b><% = intProjected50 %></b>
		</td>
		<td class='TableCell' align=center>
			<b><% = intProjected75 %></b>
		</td>
		<td class='TableCell' align=center>
			<b><% = intProjected100 %></b>
		</td>
		<td class='TableCell' align=center>
			<b><% = intProjected0 + intProjected25 + intProjected50 + intProjected75 + intProjected100 + intBlankProjected%></b>
		</td>
	</tr>	
<%
' Close objects
rsStudents.Close
set rsStudents = nothing
call oFunc.CloseCN()
set oFunc = nothing
' Conclude html 
response.Write "</table>"
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Functions and Subroutines
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
function vbfPrintHeaders
	'We put the headers in a fucntion so we could reuse them easily
%>
	<tr>
		<td class="TableHeader">
			<b>&nbsp;Student Name</b>
		</td>
		<td class="TableHeader">
			<b>&nbsp;Grd&nbsp;</b>
		</td>
		<td class="TableHeader">
			<b>&nbsp;Proj&nbsp;</b>
		</td>
		<td class="TableHeader">
			<b>&nbsp;Core&nbsp;</b>
		</td>
		<td class="TableHeader">
			<b>&nbsp;Elec&nbsp;</b>
		</td>
		<td class="TableHeader">
			<b>&nbsp;ASD Hrs&nbsp;</b>
		</td>
		<td class="TableHeader">
			<b>&nbsp;0%</b>
		</td>
		<td class="TableHeader">
			<b>&nbsp;25%</b>
		</td>
		<td class="TableHeader">
			<b>&nbsp;50%</b>
		</td>
		<td class="TableHeader">
			<b>&nbsp;75%</b>
		</td>
		<td class="TableHeader">
			<b>&nbsp;100%</b>
		</td>
	</tr> 
<%
end function

sub vbsFundingCategoryCalc(grade,percentage,uniqueID)
	' Organizes enrollment %'s into grade groupings
	' So for grades k - 5 we will know how many students fall under 0%,25% etc enrollment
	' Requires parameters 'grade' (grade level of student) and percentage (integer 
	
	'Change K into integer
	if ucase(grade) = "K" then grade = 0
	'no grade provided
	if grade & "" = "" then grade = -1
	grade = cint(grade)
	
	select Case grade
		case 0,1,2,3,4,5
			execute("elementry" & percentage & uniqueID & "=" & "elementry" & percentage & uniqueID & " + 1")
		case 6,7,8
			execute("middle" & percentage & uniqueID & "=" & "middle" & percentage & uniqueID & " + 1")
		case 9,10,11,12
			execute("high" & percentage & uniqueID & "=" & "high" & percentage & uniqueID & " + 1")
	end select
end sub
%>
