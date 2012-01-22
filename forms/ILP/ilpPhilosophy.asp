<%@ Language="VBScript" %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		ilpPhilosophy.asp
'Purpose:	Form that allows the user to enter an education philosophy
'			that is automatically included on ILP's for students
'			that are selected to use the philosophy.
'Date:		25 Jan 2004
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim oFunc				'wsc object
dim intStudent_ID
dim intFamily_ID
dim intPhilosophy_ID
dim strPhilosophy
dim sql
dim rs

intStudent_ID = request("intStudent_ID")

if intStudent_ID = "" then
	response.Write "<h1>Page improperly called.</h1>"
	response.End
end if

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

if Session.Value("intFamily_ID") <> "" then
	intFamily_ID = Session.Value("intFamily_ID")
else
	intFamily_ID = oFunc.FamilyInfo(1,intStudent_ID,1)
end if

if request("btSubmit") <> "" then
	if request("intPhilosophy_ID") <> "" then
		intPhilosophy_ID = request("intPhilosophy_ID")
		call vbsUpdate()
	elseif session.Contents(intStudent_ID & "PhilosophySaved") = "" then
		' seesion varable ensures that we don't try to insert
		' the philosophy multiple times (via page refreshes/multiple submits)
		call vbsInsert()
	end if
end if
Server.Execute(Application.Value("strWebRoot") & "includes/Header.asp")
%>
<table width="100%" ID="Table1">
	<tr>
		<td class="yellowHeader">
			&nbsp;<b>ILP Philosophy</b>
		</td>
	</tr>
</table>
<%

sql = "SELECT p.szPhilosophy, p.intPhilosophy_ID " & _
		"FROM tblENROLL_INFO e INNER JOIN " & _
		" tblPhilosophy p ON e.intPhilosophy_ID = p.intPhilosophy_ID " & _
		"WHERE     (e.sintSCHOOL_YEAR = " & session.Contents("intSchool_year") & ") " & _
		" AND (e.intSTUDENT_ID = " & intStudent_ID & ")"
set rs = server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
rs.Open sql, oFunc.FPCScnn

if rs.RecordCount > 0 then
	strPhilosophy = rs(0)
	intPhilosophy_ID = rs(1)
else
	rs.Close
	sql = "SELECT top 1 p.szPhilosophy, p.intPhilosophy_ID " & _
		"FROM tblENROLL_INFO e INNER JOIN " & _
		" tblPhilosophy p ON e.intPhilosophy_ID = p.intPhilosophy_ID " & _
		"WHERE (e.intSTUDENT_ID = " & intStudent_ID & ") " & _
		" ORDER BY e.sintSchool_Year desc "
	rs.CursorLocation = 3 
	rs.Open sql, oFunc.FPCScnn
	if rs.RecordCount > 0 then
		strPhilosophy = rs(0)
	end if
end if
rs.Close

sql  ="SELECT s.intSTUDENT_ID, s.szFIRST_NAME, e.intPhilosophy_ID " & _
			"FROM tblSTUDENT s INNER JOIN " & _
			" tblFAMILY f ON s.intFamily_ID = f.intFamily_ID INNER JOIN " & _
			" tblStudent_States ss ON s.intSTUDENT_ID = ss.intStudent_id LEFT OUTER JOIN " & _
			" tblENROLL_INFO e ON s.intSTUDENT_ID = e.intSTUDENT_ID " & _
			"WHERE (f.intFamily_ID = " & intFamily_ID & ") AND " & _
			"(ss.intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _
			" AND ss.intReEnroll_State IN (" & Application.Contents("ActiveEnrollList") & ") AND (e.sintSCHOOL_YEAR = " & session.Contents("intSchool_Year") & ") " & _
			"ORDER BY s.szFIRST_NAME"
			
rs.Open sql, oFunc.FPCScnn

%>
<form action=ilpPhilosophy.asp method=post>
<input type=hidden name="intPhilosophy_ID" value="<% = intPhilosophy_ID%>">
<input type=hidden name="intStudent_ID"	value="<% = intStudent_ID %>">
<table cellpadding=3>
	<tr>
		<td class=gray>
			An ILP educational philosophy statement is required in your Student Packet.  
			This philosophical statement will be used to represent your educational 
			beliefs as reflected in the development of your child’s Student Packet.  
			Below enter the educational philosophy that you would like included in the 
			Student Packet for this child.
		</td>
	</tr>
	<tr>
		<td>
			<textarea style="width:100%;" rows=8 name=szPhilosophy wrap=virtual onKeyDown="jfMaxSize(4000,this);"><%=strPhilosophy%></textarea>
		</td>
	</tr>
	<tr>
		<td align=center>
			<table cellpadding=4>
				<tr>
					<td class=gray align=center>
						Use the above philosophy for the following students... <b>(check all that apply THEN click 'Save')</b>
					</td>
				</tr>
				<tr>
					<td align=center>
						<table>
							<tr class=svplain8>														
<%
	do while not rs.EOF
		strChecked = ""
		if isNumeric(rs("intPhilosophy_ID")) and isNumeric(intPhilosophy_ID) then			
			if cint(intPhilosophy_ID) = cint(rs("intPhilosophy_ID")) then
				strChecked = " checked "
			end if
		end if
%>							
								<td>
									<input type=checkbox name="UsePhilosophy" value="<% = rs("intStudent_ID") %>" <% = strChecked %>><% = rs("szFIRST_NAME") %>
								</td>
<%		
		rs.MoveNext
	loop
%>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>
			<hr width=100% size=1>
			<% ' first check to see if year is locked
				if not oFunc.LockYear then
			%>
			<input type=submit value="SAVE" class="navSave" NAME="btSubmit">
			<% end if %>
			<input type=button value="return to packet (does not save)" class="btSmallgray" onClick="window.location.href='<%=Application.Value("strWebRoot")%>forms/packet/packet.asp?intStudent_ID=<%=intStudent_ID%>'">
		</td>
	</tr>
</table>
</form>
<%
rs.Close
set rs = nothing
call oFunc.CloseCN()
set oFunc = nothing

Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")

sub vbsInsert()
	' First insert the philiosophy
	dim insert
	oFunc.BeginTransCN
	insert = "insert into tblPhilosophy(szPhilosophy,intFamily_ID,szUser_Create) " & _
			 "values ('" & oFunc.EscapeTick(request("szPhilosophy")) & "'," & _
			 intFamily_ID & "," & _
			 "'" & oFunc.EscapeTick(session.Contents("strUser_ID")) & "')"
	oFunc.ExecuteCN(insert)	 
	intPhilosophy_ID = oFunc.GetIdentity
	call vbsSetPhilosophy
	oFunc.CommitTransCN
	session.Contents(intStudent_ID & "PhilosophySaved") = "true"
end sub 

sub vbsUpdate()
	dim update
	oFunc.BeginTransCN
	update = "update tblPhilosophy " & _
		     "set szPhilosophy = '" & oFunc.EscapeTick(request("szPhilosophy")) & "', " & _
		     " szUser_Modify = '" & oFunc.EscapeTick(session.Contents("strUser_ID")) & "' " & _
		     " where intPhilosophy_ID = " & intPhilosophy_ID
	oFunc.ExecuteCN(update)
	' Reset all records where intPhilosophy_ID is used
	update = "update tblEnroll_Info " & _
			 " set intPhilosophy_id = null " & _
			 ", szUser_Modify = '" & oFunc.EscapeTick(session.Contents("strUser_ID")) & "' " & _
			 " where intPhilosophy_ID = " & intPhilosophy_ID & _			 
			 " and sintSchool_Year = " & session.Contents("intSchool_Year") 
	oFunc.ExecuteCN(update)
	' Add intPhilosophy_ID to selected students enroll info
	call vbsSetPhilosophy
	oFunc.CommitTransCN
end sub

sub vbsSetPhilosophy()
	dim update
	arStudents = split(request("UsePhilosophy"),",")
	for i = 0 to ubound(arStudents)
		if arStudents(i) <> "" then 
			update = "Update tblEnroll_Info " & _
					 " set intPhilosophy_ID = " & intPhilosophy_ID & _
					 ", szUser_Modify = '" & oFunc.EscapeTick(session.Contents("strUser_ID")) & "' " & _
					 " where intStudent_ID = " & arStudents(i) & _
					 " and sintSchool_Year = " & session.Contents("intSchool_Year") 
			oFunc.ExecuteCN(update)
		end if
	next
end sub
%>