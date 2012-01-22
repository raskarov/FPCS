<%@ Language=VBScript %>
<%
'*******************************************
'Name:		Admin\teacherEmailer.asp
'Purpose:	Allows FPCS staff to email all teachers, groups of teachers,
'			or individually selected teachers from the FPCS database.
'
'CalledBy:	
'
'Inputs:	Request.QueryString("szUserID")
'
'Author:	ThreeShapes.com LLC
'Date:		22 April 2002
'*******************************************

Session.Value("strTitle") = "Teachers Email Interface"
Server.Execute(Application.Value("strWebRoot") & "Includes/header.asp")
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()


if Request.Form("strBody") <> "" or request("listOnly") <> "" then
	'This block handles getting email addresses and sending the emails
	dim strFrom
	dim strWhere
	dim strCase
	dim sqlGetEmail	
	dim strTeacherList
	strTeacherList = Request.Form("strTeachers")
	
	if inStr(1,strTeacherList,"*") > 0 then 
	
		' A certain group of teachers where selected
		strCase = replace(strTeacherList,"*","")
		strFrom = " , trefPay_Types pt " 
		strWhere = " AND i.intPay_Type_ID = pt.intPay_Type_ID and "
		if instr(1,strCase,",") > 0 then
			arTypeList = split(strCase,",")
			strWhere = strWhere & "("
			for i = 0 to ubound(arTypeList)
				strWhere = strWhere & " pt.intPay_Type_ID = " & arTypeList(i) & " or "  
			next
			strWhere = left(strWhere,len(strWhere)- 3)
			strWHere = strWhere & ")"
		else
			strWHere = strWhere & " pt.intPay_Type_ID = " & strCase
		end if
	elseif inStr(1,strTeacherList,",") > 0 then
		' Specific teachers where hand selected
		arInstrcutList = split(strTeacherList,",")
		strWhere = " AND intInstructor_ID = '" & arInstrcutList(0) & "' "		
		for w = 1 to ubound(arInstrcutList)
			strWhere = strWhere & " or intInstructor_ID = '" & arInstrcutList(w) & "' "
		next 		
	elseif strTeacherList = "all" then
		' This code put here only to emphasis that with 'all' we don't limit the where clause
		strWhere = ""
		strFrom = "" 
	elseif instr(1,strTeacherList, "CST") < 1 then
		' Only a single selection was made
		strWhere = " AND intInstructor_ID = '" & strTeacherList & "' "		
	end if

	if instr(1,strTeacherList, "CST") > 0 then
		strWhere = strWhere & " AND EXISTS " & _
                          "(SELECT     ' x' AS Expr1 " & _
                          " FROM          tblENROLL_INFO AS e2 " & _
                          " WHERE      (sintSCHOOL_YEAR = " & session.contents("intSchool_Year") & ")" & _
			  " AND (intSponsor_Teacher_ID = i.intINSTRUCTOR_ID)) "
	end if
	
	sqlGetEmail = "select szEmail from tblInstructor i " & strFrom & _
	" WHERE ((SELECT     TOP 1 bolActive " & _ 
	"			FROM tblInstructor_Pay_Data ip " & _ 
	"			WHERE (ip.intInstructor_ID = i.intInstructor_ID) AND " & _
	"			(ip.intSchool_Year_Start <= " & session.contents("intSchool_Year") & ") " & _
	"			ORDER BY ip.intSchool_Year_Start DESC, intInstructor_Pay_Data_ID DESC) = 1) " & _
	strWhere
	
	set rsGetEmail = server.CreateObject("ADODB.RECORDSET")
	rsGetEmail.Open sqlGetEmail, oFunc.FPCScnn

'response.write sqlGetEmail

	if request("strFrom") <> "" then
		strFrom = request("strFrom")
	else
		strFrom = "FPCS <Tatum_Alex@asdk12.org>" 
	end if
	'response.Write request("listOnly") & "<<"
	' Set up CDO object and set properties
	'http://msdn.microsoft.com/library/en-us/cdosys/html/_cdosys_messaging_examples_creating_and_sending_a_message.asp?frame=true
	if request("listOnly") = "" then 
		Set cdoMessage = Server.CreateObject("CDO.Message")
		set cdoConfig = Server.CreateObject("CDO.Configuration")
		cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 ' cdoSendUsingPort
		cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "127.0.0.1"
		cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		cdoConfig.Fields.Update
		set cdoMessage.Configuration = cdoConfig
		
		cdoMessage.From = strFrom
		cdoMessage.Subject = Request.Form("strSubject")
		cdoMessage.TextBody = Request.Form("strBody") 
		'on error resume next
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
				end if
			end if 
			rsGetEmail.MoveNext
		loop
	end if
	
	rsGetEmail.Close
	set rsGetEmail = nothing
end if 

%>
<form name=main method=post action="teacherEmailer.asp" onsubmit="return false;">
<table width="100%">
	<tr>
		<Td class="yellowHeader">
			&nbsp;<b>FPCS Teacher Mass Emailer</b>
		</Td>
	</tr>
	<tr>
		<td bgcolor="f7f7f7">
			<table>
				<tr>
					<td class=Gray>
						<nobr>&nbsp;<b>Select Teacher(s)</nobr><br>
						<nobr>&nbsp;&nbsp;to Email</b></nobr>
					</td>
					<td width=100%>
						<select name="strTeachers" multiple size=5>
							<option value="all">ALL Teachers
							<option value="CST">Current Sponsor Teachers							
						<%
							dim sqlTypes
							set rsPayTypes = server.CreateObject("ADODB.RECORDSET")
							rsPayTypes.CursorLocation = 3
							sqlTypes = "select szPay_Type_Name, intPay_Type_ID " & _
									   "from trefPay_Types " & _
									   "order by szPay_Type_Name"
							rsPayTypes.Open sqlTypes, oFunc.FPCScnn
							
							do while not rsPayTypes.EOF
								Response.Write "<option value='*" & rsPayTypes(1) & "'>ALL " & rsPayTypes(0) & " Teachers" & chr(10)
								rsPayTypes.MoveNext
							loop
							rsPayTypes.Close
							set rsPayTypes = nothing							
							
							dim sqlInstructor
							sqlInstructor = "SELECT intINSTRUCTOR_ID, szLAST_NAME + ', ' + szFIRST_NAME AS Teacher_Name " & _ 
											" FROM tblINSTRUCTOR i " & _ 
											" WHERE ((SELECT     TOP 1 bolActive " & _ 
											"			FROM tblInstructor_Pay_Data ip " & _ 
											"			WHERE (ip.intInstructor_ID = i.intInstructor_ID) AND " & _
											"			(ip.intSchool_Year_Start <= " & session.contents("intSchool_Year") & ") " & _
											"			ORDER BY ip.intSchool_Year_Start DESC, intInstructor_Pay_Data_ID DESC) = 1) " & _ 
											" ORDER BY Teacher_Name "
							Response.Write oFunc.MakeListSQL(sqlInstructor,intInstructor_ID,Name,"")	
						%>
						</select>
					</td>
				</tr>
				<tr>
					<td class="gray" nowrap>
						&nbsp;<b>Create Email List Only</b>
					</td>
					<td>
						<input type="checkbox" name="listOnly" value="1" onclick="this.form.submit();">
					</td>
				</tr>
				<tr>
					<td class=Gray>
						<nobr>&nbsp;<b>From</b></nobr><br>
					</td>
					<td>
						<input type=text name=strFrom size=43 ID="Text1">
					</td>
				</tr>
				<tr>
					<td class=Gray>
						<nobr>&nbsp;<b>CC</b></nobr><br>
					</td>
					<td>
						<input type=text name=strCC size=43 ID="Text2">
					</td>
				</tr>
				<tr>
					<td class=Gray>
						<nobr>&nbsp;<b>Subject</b></nobr><br>
					</td>
					<td>
						<input type=text name=strSubject size=43>
					</td>
				</tr>
				<tr>
					<td class=Gray valign=top>
						<nobr>&nbsp;<b>Message</b></nobr><br>
					</td>
					<td>
						<textarea cols=55 rows=12 name=strBody wrap=virtual></textarea>
					</td>
				</tr>
				<% if strList <> "" then %>
				<tr>
					<td class=Gray valign=top>
						<nobr>&nbsp;<b>Email List</b></nobr><br>
					</td>
					<td>
						<textarea cols=55 rows=12 wrap=virtual ID="Textarea1"><% = strList %></textarea>
					</td>
				</tr>
				<% end if %>
			</table>
		</td>
	</tr>			
</table>	
&nbsp;<input type=submit value="Send" onClick="jfValidate(this.form);" ID="Submit1" NAME="Submit1">
<script language="javascript">
	function jfValidate(pForm){
		var eMsg = "";
		if (pForm.strTeachers.value == "") { eMsg += "You must select a Teacher(s).\n";}
		if (pForm.strFrom.value == "") { eMsg += "You must provide a value for 'From Email Address'\n";}
		if (pForm.strBody.value == "") { eMsg += "You must provide a value for 'Message'\n";}
		if (eMsg != "") { 
			alert(eMsg);
		}else{
			pForm.submit();
		}
	}
</script>
</form>	
<%

set oFunc = nothing
Server.Execute(Application.Value("strWebRoot") & "Includes/footer.asp")

%>