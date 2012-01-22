<%@ Language=VBScript %>
<%
   set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
   call oFunc.OpenCN()
   oFunc.ResetSelectSessionVariables()
' We dimension the following variables here to give them global scope and they are defined in vbfGetBio
dim szFirst_Name
dim szLast_Name
dim szEmail
dim szBio
dim szPhoto_Link
dim intInstructor_Bios_ID
dim strPhoto
dim intInstructor_ID

Session.Value("strTitle") = "Teacher Bios"
Session.Value("strLastUpdate") = "26 July 2002"
Server.Execute(Application.Value("strWebRoot") & "includes/header.asp")

' Set up bolShow_Classes for both insert and update if needed
if request.Form("bolShow_Classes") = "" then
	bolShow_Classes = 0 
else 
	bolShow_Classes = 1
end if 

if Request.QueryString("intInstructor_ID") <> "" then
	' Show existing record
	call vbfGetBio(Request.QueryString("intInstructor_ID"))
elseif Request.Form("intInstructor_Bios_ID") <> "" then
	' Update and show existing record
	dim update
	update = "update tblinstructor_Bios set " & _ 
			 "szAdditional_Contact = '" & oFunc.EscapeTick(Request.Form("szAdditional_Contact")) & "', " & _
			 "szBio = '" & oFunc.EscapeTick(Request.Form("szBio")) & "', " & _
			 "bolShow_Classes = '" & bolShow_Classes & "' " & _
			 "where intInstructor_Bios_ID = " & Request.Form("intInstructor_Bios_ID")
	oFunc.ExecuteCN(update)
	strMessage = "alert('Bio has been updated.');"
	call vbfGetBio(Request.Form("intInstructor_ID"))
elseif Request.Form("intInstructor_Bios_ID") = "" and Request.Form("szBio") <> "" then
	' Create and show a new record
	dim insert
	insert = "insert into tblinstructor_Bios(intInstructor_ID,szAdditional_Contact," & _
	         "szBio,bolShow_Classes,szUser_Create) values(" & _
			 Request.Form("intInstructor_ID") & "," & _
			 "'" & oFunc.EscapeTick(Request.Form("szAdditional_Contact")) & "', " & _
			 "'" & oFunc.EscapeTick(Request.Form("szBio")) & "', " & _
			 "'" & bolShow_Classes & "', " & _
			 "'" & session.Contents("strUserID") & "')"
	oFunc.ExecuteCN(insert)
	strMessage = "alert('New bio has been creaated.');"
	call vbfGetBio(Request.Form("intInstructor_ID"))
else
	' Stop script.  We must have the "intInstructor_ID" parameter provided by the user
	Response.Write " You did not provide a valid instructor id. Process ended prematurely."
	Response.End
end if

function vbfGetBio(instructor_ID)
	' This function takes the parameter instructor_ID and pulls the bio info
	' for the given instructor
	
	dim sqlTeacher
	set rsInstInfo = server.CreateObject("ADODB.Recordset")
	rsInstInfo.CursorLocation = 3
	
	sqlTeacher = "select i.szFirst_Name, i.szLast_Name, i.szEmail, b.szBio, b.szPhoto_Link, " & _
				 "b.bolShow_Classes, b.szAdditional_Contact, b.intInstructor_Bios_ID " & _
				 "from tblInstructor i left outer join tblInstructor_Bios b " & _
				 "ON i.intInstructor_ID = b.intInstructor_ID " & _
				 "where i.intInstructor_ID = " & instructor_ID
	
	rsInstInfo.Open sqlTeacher, oFunc.FPCScnn
	
	'This for loop dimentions and defines all the columns we selected in sqlTeacher
	'and we use the variables created here to populate the form.
	for each item in rsInstInfo.Fields
		execute(item.Name & " = item")		
	next
	
    ' Check to see if we have the teachers photo 
    Set fso = CreateObject("Scripting.FileSystemObject")
    if not fso.FileExists(Server.MapPath(Application("strImageRoot") & "teachers/" & instructor_ID & ".jpg")) then
		strPhoto = "Photo not available"
	else
		strPhoto =  "<a href='" & Application("strImageRoot") & "teachers/" & instructor_ID & ".jpg' target='_new'>" &_
					"Click here to view picture</a>"
	end if 

	rsInstInfo.Close
	set rsInstInfo = nothing
	
	intInstructor_ID = instructor_ID

end function 
oFunc.CloseCN()
set oFunc = nothing
%>
<script language=javascript>
 <% = strMessage %>
</script>
<form action=teacherBiosAdmin.asp method=post>
<input type=hidden name=intInstructor_Bios_ID value="<%=intInstructor_Bios_ID%>">
<input type=hidden name=intInstructor_ID value="<%=intInstructor_ID%>">
<table width=100%>
	<tr>	
		<Td class=yellowHeader>
				&nbsp;<b>Teacher Bio Editor</b>
		</td>
	</tr>
	<tr>
		<td bgcolor=f7f7f7>
			<table>
				<tr>	
					<Td colspan=2>
						<font class=svplain11>
							<b><i>Teacher's Information</I></B> 
						</font>
						<font class=svplain>
						</font>
					</td>
				</tr>
				<tr>
					<td class=gray>
							&nbsp;<b>Teacher's Name:</B>
					</td>
					<td class=gray>
							&nbsp;<% = szFirst_Name & " " & szLast_Name %>
					</td>
				</tr>
				<tr>
					<td class=gray>
							&nbsp;<b>Teacher's Email:</B>
					</td>
					<td class=gray>
							&nbsp;<a href="mailto:<% = szEmail %>"><% = szEmail %></a>
					</td>
				</tr>
				<tr>
					<td class=gray>
							&nbsp;<b>Additional Contact Info:</B>
					</td>
					<td class=gray>
							<input type=text name="szAdditional_Contact" size=40 maxlength=128 value="<% = szAdditional_Contact%>">
					</td>
				</tr>
				<tr>
					<td class=gray>
							&nbsp;<b>Photo:</B>
					</td>
					<td class=gray>
						 <% = strPhoto %>
					</td>
				</tr>
				<tr>
					<td class=gray>
							&nbsp;<b>Show Classes:</B>
					</td>
					<td class=gray>
						<% 
						if ucase(bolShow_Classes) = "TRUE" then
								strChecked = " checked "
						   else 
								strChecked = ""
						   end if
						%>
						 </b><input type=checkbox name="bolShow_Classes" value="1" <% = strChecked %>>
						 (check for yes)
					</td>
				</tr>
				<tr>
					<td class=gray colspan=2>
							&nbsp;<b>Biographical Information:</B>
					</td>
				</tr>
				<tr>
					<td colspan=2>
						<textarea cols=54 rows=10 name=szBio wrap=virtual onKeyDown="jfMaxSize(7000,this);"><%=szBio%></textarea>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<input type=button value="Home" class="NavLink" onClick="window.location.href='<%=Application.Value("strWebRoot")%>';"  >
<input type=submit value="Submit" class="NavSave">
</form>
</BODY>
</HTML>
