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

Session.Value("strTitle") = "Vendor Email Interface"
Server.Execute(Application.Value("strWebRoot") & "Includes/header.asp")
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()


if Request.Form("strBody") <> "" or request("listOnly") <> "" then
	'This block handles getting email addresses and sending the emails
	dim strFrom
	dim strWhere
	dim strCase
	dim sqlGetEmail	
	dim strVendorList
	dim ApprovalType

	strVendorList = Request.Form("strVendorList")
	ApprovalType = "'APPR','PEND'"	

	if strVendorList = "all" then
		' This code put here only to emphasis that with 'all' we don't limit the where clause
		strWhere = ""
	elseif strVendorList = "AONLY" then
		ApprovalType = "'APPR'"
	elseif strVendorList & "" <> ""  then
		strWhere = " and (v.intVendor_ID in (" & strVendorList & ")) "						
	end if
	
	sqlGetEmail = "SELECT     intVendor_ID, szVendor_Name, szVendor_Email, szVendor_Phone " & _
					" FROM    tblVendors v " & _
					" WHERE     ((SELECT     TOP 1 upper(vs.intVendor_ID) " & _
					"				FROM         tblVendor_Status vs " & _
					"				WHERE     vs.intVendor_ID = v.intVendor_ID AND vs.intSchool_Year = " & session.Contents("intSchool_Year") & " AND vs.szVendor_Status_CD IN (" & ApprovalType  & ") AND  " & _
					"					vs.dtContract_Start IS NOT NULL " & _
					"				ORDER BY intSchool_Year DESC, intVendor_Status_ID DESC) = intVendor_ID) AND (bolService_Vendor = 1) " & strWhere & _
					"ORDER BY szVendor_Name "
	

set rsGetEmail = server.CreateObject("ADODB.RECORDSET")
rsGetEmail.CursorLocation = 3
	rsGetEmail.Open sqlGetEmail, Application("cnnFPCS")'oFunc.FPCScnn
	

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
			if len(rsGetEmail("szVendor_Email")) > 0 then
				if instr(1,rsGetEmail("szVendor_Email"),"@") > 0 and _
					instr(1,rsGetEmail("szVendor_Email"),".") > 0 then
					strList = strList & rsGetEmail("szVendor_Email") & ";"
				end if
			end if 
			rsGetEmail.MoveNext
		loop
	end if

	rsGetEmail.moveFirst

	dim sHtml 

	sHtml = "<span class='svplain10'>Selected Vendor Email List (" & rsGetEmail.recordcount & " vendors)</span><br><table><tr><td class='TableHeader' align='center'>Vendor Name</td>" & _
		"<td class='TableHeader' align='center'>Email</td>" & _
		"<td class='TableHeader' align='center'>Phone</td></tr>"

	do while not rsGetEmail.eof

		sHtml = sHtml & "<tr><td class='TableCell' align='center'>" & rsGetEmail("szVendor_Name") & "</td>" & _
				"    <td class='TableCell' align='center'><a href='mailto:" & rsGetEmail("szVendor_Email") & "'>" & rsGetEmail("szVendor_Email") & "</a></td>" & _
				"    <td class='TableCell' align='center'>" & rsGetEmail("szVendor_Phone") & "</td></tr>" & _
		rsGetEmail.moveNext
	loop

	sHtml = sHtml & "</table>"

	rsGetEmail.Close
	set rsGetEmail = nothing
end if 

%>
<form name=main method=post action="vendorEmailer.asp" ID="Form1" onsubmit="return false;">
<table width="100%" ID="Table1">
	<tr>
		<Td class="yellowHeader">
			&nbsp;<b>FPCS Service Vendor Mass Emailer</b>
		</Td>
	</tr>
	<tr>
		<td class="svplain8" colspan="2">
			<b>Please Note:</b><BR>The list of vendors below only includes <u>service</u>
			vendors that have a status of <u>Approved</u> or <u>Pending</u> and have
			a <u>Contract Start Date</u> for the school year you are working with.
		</td>
	</tr>
	<tr>
		<td bgcolor="f7f7f7">
			<table ID="Table2">
				<tr>
					<td class=Gray>
						<nobr>&nbsp;<b>Select Vendor(s)</nobr><br>
						<nobr>&nbsp;to Email</b></nobr>
					</td>
					<td width=100%>
						<select name="strVendorList" multiple size=5 ID="Select1">
							<option value="all">ALL APPR AND PEND Vendors	
							<option value="AONLY">Only APPR Vendors						
						<%							
							dim sqlFamilies
							sqlVendors = "SELECT     intVendor_ID, szVendor_Name " & _
										" FROM         tblVendors v " & _
										" WHERE     ((SELECT     TOP 1 upper(vs.intVendor_ID) " & _
										"				FROM         tblVendor_Status vs " & _
										"				WHERE     vs.intVendor_ID = v.intVendor_ID AND vs.intSchool_Year = " & session.Contents("intSchool_Year") & " AND vs.szVendor_Status_CD IN ('APPR', 'PEND') AND  " & _
										"					vs.dtContract_Start IS NOT NULL " & _
										"				ORDER BY intSchool_Year DESC, intVendor_Status_ID DESC) = intVendor_ID) AND (bolService_Vendor = 1) " & _
										"ORDER BY szVendor_Name "
							Response.Write oFunc.MakeListSQL(sqlVendors,"intVendor_ID","szVendor_Name","")	
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
&nbsp;<input type=submit value="Send" onClick="jfValidate(this.form);" ID="Submit1" NAME="Submit1">
<br><br>
<% = sHtml %>
<script language="javascript">
	function jfValidate(pForm){
		var eMsg = "";
		if (pForm.strVednors.value == "") { eMsg += "You must select at least one Family.\n";}
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