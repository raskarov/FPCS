<%@ Language=VBScript %>
<%
'*******************************************
'Name:		Admin\familyEmailer.asp
'Purpose:	Allows FPCS staff to email all familes, groups of families,
'			or individually selected teachers from the FPCS database.
'
'CalledBy:	
'
'Inputs:	Request.QueryString("szUserID")
'
'Author:	ThreeShapes.com LLC
'Date:		22 April 2002
'*******************************************

Session.Value("strTitle") = "Family Email Interface"
Server.Execute(Application.Value("strWebRoot") & "Includes/header.asp")
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()
dim strHtmlList 

if Request.Form("strBody") <> "" or request("listOnly") <> "" then
	'This block handles getting email addresses and sending the emails
	dim strFrom
	dim strWhere
	dim strCase
	dim sqlGetEmail	
	dim strFamilyList
	strFamilyList = Request.Form("strFamilies")
	'Response.Write inStr(1,strFamilyList,"*") & "<<"
	if inStr(1,strFamilyList,",") > 0 and inStr(1,strFamilyList,"'") < 1 then
		' Specific families where hand selected
		arFamilyList = split(strFamilyList,",")
		strWhere = " AND (f.intFamily_ID = '" & arFamilyList(0) & "' "		
		for w = 1 to ubound(arFamilyList)
			strWhere = strWhere & " or f.intFamily_ID = '" & arFamilyList(w) & "' "
		next 	
		strWhere = strWhere & ")"	
	elseif strFamilyList = "all" then
		' This code put here only to emphasis that with 'all' we don't limit the where clause
		strWhere = ""
	elseif instr(1,strFamilyList,"'") > 0 then
		strWhere = " and ss.szGrade in (" & strFamilyList & ") " 
	else
		' Only a single selection was made
		strWhere = " and (f.intFamily_ID = '" & strFamilyList & "') "		
	end if

	
	sqlGetEmail = "SELECT DISTINCT f.szEMAIL, szDesc + ' '  + szFamily_Name as Fam_Name  " & _
					"FROM tblFAMILY f INNER JOIN " & _
					" tblSTUDENT s ON f.intFamily_ID = s.intFamily_ID INNER JOIN " & _
					" tblStudent_States ss ON s.intSTUDENT_ID = ss.intStudent_id " & _
					"WHERE (ss.intSchool_Year = " & session.Contents("intSchool_Year") & _
					") AND ss.intReEnroll_State IN (" & Application.Contents("ActiveEnrollList") & ") AND (f.szEMAIL IS NOT NULL) AND (f.szEMAIL <> '') " & _
					strWhere
		
'response.write sqlGetEmail & " ORDER BY FAM_NAME "		
	set rsGetEmail = server.CreateObject("ADODB.RECORDSET")
	rsGetEmail.Open sqlGetEmail, oFunc.FPCScnn
	
	if request("listOnly") = "" then 
		' Set up CDO object and set properties
		'http://msdn.microsoft.com/library/en-us/cdosys/html/_cdosys_messaging_examples_creating_and_sending_a_message.asp?frame=true
		Set cdoMessage = Server.CreateObject("CDO.Message")
		set cdoConfig = Server.CreateObject("CDO.Configuration")
		cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 ' cdoSendUsingPort
		cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "127.0.0.1"
		cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25


		cdoConfig.Fields.Update
		set cdoMessage.Configuration = cdoConfig
		
		cdoMessage.From = Request.Form("strFrom")
		cdoMessage.Subject = Request.Form("strSubject")
		cdoMessage.TextBody = Request.Form("strBody") 
		on error resume next
		do while not rsGetEmail.EOF
			if len(rsGetEmail("szEmail")) > 0 then
				if instr(1,rsGetEmail("szEmail"),"@") > 0 and _
					instr(1,rsGetEmail("szEmail"),".") > 0 then
					cdoMessage.To = rsGetEmail("szEmail")	
					cdoMessage.Send 
				end if
			end if 
			rsGetEmail.MoveNext
		loop
		
		if request("strCC") <> "" then
			cdoMessage.To = request("strCC")	
			cdoMessage.Send 
		end if
		
		'Clean up Objects
		Set cdoMessage = Nothing 
		%>
		<script language=javascript>
			alert('Email messages have been sent.');
		</script>
		<%	
	else
		do while not rsGetEmail.EOF
			if len(rsGetEmail("szEmail")) > 0 then
				if instr(1,rsGetEmail("szEmail"),"@") > 0 and _
					instr(1,rsGetEmail("szEmail"),".") > 0 then
					strList = strList & rsGetEmail("szEmail") & ";"
					strHtmlList = strHtmlList & "<tr><td class='TableCell'>" & _
						     rsGetEmail("Fam_Name") & "</td>" & _
						     "<td class='TableCell'>" & _
						     "<a href='mailto:" & rsGetEmail("szEmail") & "'>" & rsGetEmail("szEmail")  & _
							"</a></td></tr>"
				end if
			end if 
			rsGetEmail.MoveNext
		loop
	end if
	rsGetEmail.Close
	set rsGetEmail = nothing
end if 

%>
<form name=main method=post action="familyEmailer.asp" ID="Form1" onsubmit="return false;">
<table width="100%" ID="Table1">
	<tr>
		<Td class="yellowHeader">
			&nbsp;<b>FPCS Family Mass Emailer</b>
		</Td>
	</tr>
	<tr>
		<td bgcolor="f7f7f7">
			<table ID="Table2">
				<tr>
					<td class=Gray>
						<nobr>&nbsp;<b>Select Family(s)</nobr><br>
						<nobr>&nbsp;to Email</b></nobr><br>
						(hold ctrl key down to make multiple selections)
					</td>
					<td width=100%>
						<select name="strFamilies" multiple size=5 ID="Select1">
							<option value="all">ALL Families	
							<option value="'K'">Families w/ Kindergartners	
							<option value="'1'">Families w/ 1st Graders					
							<option value="'2'">Families w/ 2nd Graders
							<option value="'3'">Families w/ 3rd Graders
							<option value="'4'">Families w/ 4th Graders
							<option value="'5'">Families w/ 5th Graders
							<option value="'6'">Families w/ 6th Graders
							<option value="'7'">Families w/ 7th Graders
							<option value="'8'">Families w/ 8th Graders
							<option value="'9'">Families w/ 9th Graders
							<option value="'10'">Families w/ 10th Graders
							<option value="'11'">Families w/ 11th Graders
							<option value="'12'">Families w/ 12th Graders
						<%							
							dim sqlFamilies
							sqlFamilies = "SELECT DISTINCT f.szFamily_Name + ' ' + f.szDesc as Name, f.intFamily_ID " & _
											"FROM tblFAMILY f INNER JOIN " & _
											" tblSTUDENT s ON f.intFamily_ID = s.intFamily_ID INNER JOIN " & _
											" tblStudent_States ss ON s.intSTUDENT_ID = ss.intStudent_id " & _
											"WHERE (ss.intSchool_Year = " & session.Contents("intSchool_Year") & ") AND ss.intReEnroll_State IN (" & Application.Contents("ActiveEnrollList") & ") AND f.szEmail is not null " & _
											" and f.szEmail <> '' " & _
											"ORDER BY Name"
							Response.Write oFunc.MakeListSQL(sqlFamilies,"intFamily_ID","Name","")	
						%>
						</select>
					</td>
				</tr>
				<tr>
					<td class="gray" nowrap>
						&nbsp;<b>Create Email List Only</b>
					</td>
					<td>
						<input type="checkbox" name="listOnly" value="1" onclick="this.form.submit();" ID="Checkbox1">
					</td>
				</tr>
				<tr>
					<td class=Gray>
						<nobr>&nbsp;<b>From Email Address:</b></nobr>						
					</td>
					<td>
						<input type=text name=strFrom size=43 ID="Text2">
					</td>
				</tr>
				<tr>
					<td class=Gray>
						<nobr>&nbsp;<b>CC Email Address:</b></nobr>						
					</td>
					<td>
						<input type=text name=strCC size=43 ID="Text3">
					</td>
				</tr>
				<tr>
					
					<td class=Gray>
						<nobr>&nbsp;<b>Subject:</b></nobr><br>
					</td>
					<td>
						<input type=text name=strSubject size=43 ID="Text1">
					</td>
				</tr>
				<tr>
					<td class=Gray valign=top>
						<nobr>&nbsp;<b>Message:</b></nobr><br>
					</td>
					<td>
						<textarea cols=55 rows=12 name=strBody wrap=virtual ID="Textarea1"></textarea>
					</td>
				</tr>
				<% if strList <> "" then %>
				<tr>
					<td class=Gray valign=top>
						<nobr>&nbsp;<b>Email List</b></nobr><br>
					</td>
					<td>
						<textarea cols=55 rows=12 wrap=virtual ID="Textarea2" NAME="Textarea2"><% = strList %></textarea>
					</td>
				</tr>
				<% end if %>
			</table>
		</td>
	</tr>
</table>	
&nbsp;<input type=submit value="Send" onClick="jfValidate(this.form);">
<script language="javascript">
	function jfValidate(pForm){
		var eMsg = "";
		if (pForm.strFamilies.value == "") { eMsg += "You must select at least one Family.\n";}
		if (pForm.strFrom.value == "") { eMsg += "You must provide a value for 'From Email Address'\n";}
		if (pForm.strBody.value == "") { eMsg += "You must provide a value for 'Message'\n";}
		if (eMsg != "") { 
			alert(eMsg);
		}else{
			pForm.submit();
		}
	}
</script>
<% if strHtmlList <> "" then %>
<BR><BR>
<span style='font-family:Arial;font-size:11pt;'><B>List of selected family emails and names.</b></span>
<table>
	<tr>
		<td class="TableHeader">
			&nbsp;<b>Family Name</b>
		</td>
		<td class="TableHeader">
			&nbsp;<b>Family Email</b>
		</td>
	</tr>
	<% = strHtmlList %>
</table>
<% end if %>

</form>	
<%

set oFunc = nothing
Server.Execute(Application.Value("strWebRoot") & "Includes/footer.asp")

%>