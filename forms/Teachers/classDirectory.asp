<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		classDirectory.asp
'Purpose:	Displays a list of classes, each student in a class, and student info
'			for each student
'Date:		5-5-2003
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim intInstructor_ID		'Unique Instructor ID 
dim sql						'generic sql string

'prevent page from cacheing
Response.AddHeader "Cache-Control","No-Cache"
Response.Expires = -1

if request("intInstructor_ID") <> "" then
	'Define variable if get properly called
	intInstructor_ID = request("intInstructor_ID")
else
	'terminate page since page was improperly called.
	response.Write "<html><body><H1>Page improperly called.</h1></body></html>"
	response.End
end if 

' Now we can proceded to define needed function object
dim oFunc		'wsc object
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

'Prepare and print header
Session.Value("strTitle") = "Add/Edit Budget Item"
Session.Value("strLastUpdate") = "25 March 2003"
Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")
%>

<form action="classDirectory.asp" method=post name=main onSubmit="return false;" ID="Form1">
<input type=hidden name="intInstructor_ID" value="<% = intInstructor_ID %>" ID="Hidden1">
<input type=button value="Back to FPCS OS Home" onclick="window.location.href='../../default.asp';">
<table ID="Table1">
	<tr>
		<td>
				
<table ID="Table2" border=1 cellpadding=2 cellspacing=0 bordercolor=black>
	<tr>
		<td class="YellowHeader11" colspan=10>
			&nbsp;<b>Class Directory for </b> <% = oFunc.InstructorInfo(intInstructor_ID,3) %>
		</td>
	</tr>
<%
	sql = "SELECT s.szLAST_NAME, s.szFIRST_NAME, c.szClass_Name, pos.szSubject_Name," & _
			" f.szHome_Phone, f.szEMAIL,s.intFamily_ID " & _
			"FROM tblSTUDENT s INNER JOIN " & _
			" tblILP i ON s.intSTUDENT_ID = i.intStudent_ID INNER JOIN " & _
			" tblFAMILY f ON s.intFamily_ID = f.intFamily_ID RIGHT OUTER JOIN " & _
			" tblINSTRUCTOR ins INNER JOIN " & _
			" tblClasses c ON ins.intINSTRUCTOR_ID = c.intInstructor_ID INNER JOIN " & _
			" trefPOS_Subjects pos ON c.intPOS_Subject_ID = pos.intPOS_Subject_ID " & _
			" ON i.intClass_ID = c.intClass_ID " & _
			"WHERE (ins.intINSTRUCTOR_ID = " & intINSTRUCTOR_ID & ") AND " & _
			"(c.dtClass_Start > CONVERT(DATETIME, '" & (session.Contents("intSchool_Year") -1) & "-06-30 00:00:00', 102))  " & _
			" AND (c.dtClass_End < CONVERT(DATETIME,  '" & session.Contents("intSchool_Year") & "-07-01 00:00:00', 102)) " & _
			" ORDER BY s.szLAST_NAME"
			
	set rs = server.CreateObject("ADODB.RECORDSET")
	rs.CursorLocation = 3
	rs.Open sql, oFunc.FPCScnn
	
	' Set up rs for guardian list
	set rsGuard = server.CreateObject("ADODB.RECORDSET")
	rsGuard.CursorLocation =3
	
	if rs.RecordCount > 0 then
%>
	<tr>
		<td class=gray>
			&nbsp;<b>Student's Name</b>&nbsp;
		</td>
		<td class=gray>
			&nbsp;<b>Class Name</b>&nbsp;
		</td>
		<td class=gray>
			&nbsp;<b>Subject</b>&nbsp;
		</td>
		<td class=gray>
			&nbsp;<b>Home #</b>&nbsp;
		</td>
		<td class=gray>
			&nbsp;<b>Email</b>&nbsp;
		</td>
		<td class=gray>
			&nbsp;<b>Guardians</b>&nbsp;
		</td>
	</tr>
<%
		do while not rs.EOF
%>
	<tr bgcolor=f7f7f7>
		<td class=svplain10>
			<nobr>&nbsp;<% if  rs(0) & "" = "" then 
					response.Write "(No students enrolled)"
			   else
					response.Write rs(0) & ", " & rs(1) 
			   end if
			%>&nbsp;</nobr>
		</td>
		<td class=svplain10>
			<% = rs(2) %>
		</td>
		<td class=svplain10>
			&nbsp;<% = rs(3) %>&nbsp;
		</td>
		<td class=svplain10>
			&nbsp;<% = rs(4) %>&nbsp;
		</td>
		<td class=svplain10>
			&nbsp;<a href="mailto:<% = rs(5) %>"><% = rs(5) %></a>&nbsp;
		</td>
		<td class=svplain10>
			
			<%					
				if rs(6) & "" <> "" then			
					sql = "SELECT g.szFIRST_NAME + ' ' + g.szLAST_NAME AS Name " & _
							"FROM tascFAM_GUARD fg INNER JOIN " & _
							" tblGUARDIAN g ON fg.intGUARDIAN_ID = g.intGUARDIAN_ID " & _
							"WHERE (fg.intFamily_ID = " & rs(6) & ")"
					rsGuard.Open sql,oFunc.FPCScnn
					
					strGuardList = ""
					
					if rsGuard.RecordCount > 0 then
						do while not rsGuard.EOF
							strGuardList = strGuardList & rsGuard("Name") & ", "
							rsGuard.MoveNext
						loop					
						strGuardList = left(strGuardList,len(strGuardList)-1)					
					end if 
					
					rsGuard.Close
					response.Write strGuardList	
				else 
					response.Write "&nbsp;"				
				end if
				strGuardList = ""
			%>
		</td>
	</tr>
<%
			rs.MoveNext
		loop		
	end if
%>
</table>
<%

set rsGuard = nothing
rs.Close
set rs = nothing

'Closing remarks
call oFunc.CloseCN()
set oFunc = nothing
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")
%>

