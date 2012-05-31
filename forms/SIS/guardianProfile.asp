<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		guardianProfile.asp
'Purpose:	This script collects the 
'				gaurdian information or displays the gaurdian information.
'Date:		9 July 2001
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim insert
dim update
dim dtBirth 
dim sql
dim oFunc	'wsc object
Session.Value("strTitle") = "SIS Guardian Profile"
Session.Value("simpleTitle") = "SIS Guardian Profile"
Session.Value("strLastUpdate") = "09 June 2002"

if Request.QueryString("bolNewGuardian") <> "" or Request.QueryString("bolUpdate") <> "" then
	Server.Execute(Application.Value("strWebRoot") & "Includes/simpleHeader.asp")
else
	Server.Execute(Application.Value("strWebRoot") & "Includes/header.asp")
end if


   set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
   call oFunc.OpenCN()

	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'' The Birth data is stored in the database as a single field, but our 
	'' form displays it as three seperate select lists so we collect the
	'' date parts into a single variable.
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''		
	dtBirth = request("Month") & "/" & request("Day") & "/" & request("Year")

	if Request.QueryString("intGuardian_id") <> "" then
		set rsGuardian = server.CreateObject("ADODB.RECORDSET")
		rsGuardian.CursorLocation = 3

		sqlGuardian = "select intGuardian_id,szFirst_Name, szLast_Name,sMid_Initial, szEmployer, szBusiness_Phone, " & _
					  "intPhone_Ext, szCell_Phone, szPager,bolActive_Military,szRank, " & _
					  "szEmail,szAddress,szCity,szState,szCountry,szZip_Code,szHome_Phone " & _
					  "from tblGuardian where intGuardian_id = " & Request.QueryString("intGuardian_id")
		rsGuardian.Open sqlGuardian, Application("cnnFPCS")'oFunc.FPCScnn
			
		intCount = 0
			
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'' This for loop will dimension AND assign our student info variables
		'' for us. We'll use them later to populate the form.
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''			
		if rsGuardian.RecordCount > 0 then
			for each item in rsGuardian.Fields
				execute("dim " & rsGuardian.Fields(intCount).Name)
				if item & "" <> "" then
					item = replace(item,"""","'")
				end if 
				execute(rsGuardian.Fields(intCount).Name  & " = """ & item & """")
				intCount = intCount + 1
			next
		end if 

		rsGuardian.Close
		set rsGuardian = nothing

	end if 
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'' Now we either print a blank guardian form or a populated one based
	'' on the logic above.
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
%>
<script language=javascript>
	function jfDeleteGuardian(id){
		var result;
		result = confirm("Are you sure you want to remove this guardian?");
		if (result == true)	{
			location.href="guardianInfo.asp?delete=yes&id=" + id;
		}		
	}
	
	function jfConfirm(){
			var bolContinue = confirm("Are you sure you want to close without saving any changes you may have made?");
			if (bolContinue == false) {
				return false;
			}
			window.opener.focus();
			window.close();
		}
</script>
<form action="gaurdianInsert.asp" method=Post name=main>
<input type=hidden name=changed value="">
<input type=hidden name="intGuardian_id" value="<%=Request.QueryString("intGuardian_id")%>">
<input type=hidden name="strStudents" value="<% =Request.QueryString("strStudents")%>">
<input type=hidden name="bolNewGuardian" value="<%=Request.QueryString("bolNewGuardian")%>">
<input type=hidden name="intFamily_ID" value="<% = Request.QueryString("intFamily_ID")%>">
<table width=100%>
	<tr>	
		<Td class=yellowHeader>
				&nbsp;<b>SIS Online Enrollment Form</b>
		</td>
	</tr>
	<tr>
		<td bgcolor=f7f7f7>
			<table>
				<tr>	
					<Td colspan=6>
						<font class=svplain11>
							<b><i>Guardian Information</I></B> 
						</font>
					</td>
				</tr>
				<tr>
					<td class=gray>
							&nbsp; Last
					</td>
					<td class=gray>
							&nbsp;First
					</td>
					<td class=gray>
							&nbsp;MI
					</td>
					<!--<td class=gray>
							&nbsp;Lives w/ Students
					</td>-->
				</tr>
				<tr>
					<% if session.Value("strRole") <> "GUARD" then %>
					<td>
						<input type=text name="szLast_Name" value="<% = szLast_Name %>" maxlength=50 size=17 onChange="jfChanged();">
					</td>
					<td>
						<input type=text name="szFirst_Name" value="<% = szFirst_Name %>" maxlength=50 size=15 onChange="jfChanged();">
					</td>
					<td>
						<input type=text name="sMid_Initial" value="<% = sMid_Initial %>" maxlength=1 size=2 onChange="jfChanged();">
					</td>	
					<%else%>
					<td class=gray>
						&nbsp;<% = szLast_Name%>
						<input type=hidden name="szLast_Name" value="<% = szLast_Name%>" >							
					</td>
					<td class=gray>
						&nbsp;<% = szFirst_Name%>
						<input type=hidden name="szFirst_Name" value="<% = szFirst_Name%>" >	
					</td>
					<td class=gray>
						&nbsp;<% = sMid_Initial%>
						<input type=hidden name="sMid_Initial" value="<% = sMid_Initial%>" >	
					</td>
					<%end if %>
					<!--<td	 class=gray align=center>
						<b>yes</b><input type=checkbox name="bolLives_With" <% if bolLives_With <> "" then Response.Write "checked" %>>
					</td>-->
				</tr>
			</table>
			<table>
				<tr>	
					<Td class=gray>
							&nbsp;Employer
					</td>
					<td class=gray>
							&nbsp;Active Military&nbsp;
					</td>		
					<td class=gray>
							&nbsp;Rank&nbsp;
					</td>	
					<td class=gray>
							&nbsp;Pager
					</td>	
				</tr>
				<tr>
					<td>
						<input type=text name="szEmployer" value="<% = szEmployer %>" maxlength=128 size=30 onChange="jfChanged();">
					</td>				
					<td>
						<select name="bolActive_Military" onChange="jfChanged();">
							<option value="">- - - - - - - - - - - 
						<%
							Response.Write oFunc.MakeList("TRUE,FALSE","Yes,No", oFunc.TFText(bolActive_Military))
						%>
						</select>
					</td>
					<td>
						<input type=text name="szRank" value="<% =szRank %>" maxlength=20 size=4 onChange="jfChanged();">
					</td>
					<td>
						<input type=text name="szPager" value="<% = szPager %>" maxlength=15 size=15 onChange="jfChanged();">
					</td>
				</tr>
			</table>
			<table>
				<tr>							
					<td class=gray>
							&nbsp;Business Phone&nbsp;
					</td>
					<td class=gray>
							&nbsp;Ext.
					</td>
					<td class=gray>
							&nbsp;Cell Phone
					</td>		
					<td class=gray>
							&nbsp;Email Address
					</td>								
				</tr>
				<tr>					
					<td align=center>
						<input type=text name="szBusiness_Phone" value="<% = szBusiness_Phone %>" maxlength=15 size=15 onChange="jfChanged();">
					</td>
					<td>
						<input type=text name="intPhone_Ext" value="<% = intPhone_Ext %>" maxlength=4 size=4 onChange="jfChanged();">
					</td>
					<td>
						<input type=text name="szCell_Phone" value="<% = szCell_Phone %>" maxlength=15 size=15 onChange="jfChanged();">
					</td>	
					<td>
						<input type=text name="szEmail" value="<% = szEmail %>" maxlength=128 size=30 onChange="jfChanged();">
					</td>				
				</tr>
			</table>
			<table>
				<tr>
					<td class=gray>
							&nbsp;Address (if different)
					</td>
					<td class=gray>
							&nbsp;City
					</td>
					<td class=gray>
							&nbsp;State
					</td>
					<Td class=gray>
							&nbsp;Country
					</td>				
					<Td class=gray>
							&nbsp;Zip
					</td>									
				</tr>
				<tr>
					<td>
						<input type=text name="szAddress" value="<% = szAddress %>" maxlength=256 size=30 onChange="jfChanged();">
					</td>
					<td>
						<input type=text name="szCity" value="<% = szCity %>" maxlength=50 size=10 onChange="jfChanged();">
					</td>
					<td>
						<select name="szState" onChange="jfChanged();">
						<%
							dim sqlState
							sqlState = "select strValue,strText from Common_Lists where intList_Id = 3 order by strValue"
							Response.Write oFunc.MakeListSQL(sqlState,"","",szState3)
						%>
						</select>						
					</td>
					<td>
						<input type=text name="szCountry" value="<% = szCountry %>" maxlength=25 size=7 onChange="jfChanged();">
					</td>
					<td>
						<input type=text name="szZip_Code" value="<% = szZip_Code %>" maxlength=10 size=5 onChange="jfChanged();">
					</td>		
				</tr>
			</table>		
			<%
			arStudents = split(Request.QueryString("strStudents"),",")
			dim strWhere
			dim strGuard
			
			for i = 0 to ubound(arStudents)
				if arStudents(i) <> "" then
					strWhere = strWhere & " or s.intStudent_ID = " & arStudents(i)					
				end if 
			next
			
			
			
			if Request.QueryString("intGuardian_id") <> ""  then
				strGuard = "and sg.intGuardian_ID = " & Request.QueryString("intGuardian_id")
			else
				strGuard = "and sg.intGuardian_ID is null " 
			end if 
			
			if strWhere <> "" then
				strWhere = right(strWhere,len(strWhere)-3)
				set rsRelations = server.CreateObject("ADODB.RECORDSET")
				rsRelations.CursorLocation = 3
				sql = "select distinct s.intStudent_ID, s.szFirst_Name + ' ' + s.szLast_Name as Name, sg.intGuardian_Type_ID " & _
					  "From tblStudent s " & _
					  "LEFT outer join tascStudent_Guardian sg on s.intStudent_ID = sg.intStudent_ID " & _
					  strGuard & _
					  "where " & strWhere
				rsRelations.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
			else
				strMessage = "No students selected."
			end if 
			
			'response.Write sql
			%>
			<table>
				<tr>	
					<Td colspan=6>
						<font class=svplain11>
							<b><i>Guardian/Student Relationships</I></B> 
						</font>
					</td>
				</tr>
				<tr>
					<td class=gray>
							&nbsp; Student Name
					</td>
					<td class=gray>
							&nbsp;Relationship
					</td>
			<%
				if strMessage <> "" then
			%>
				<tr>
					<td colspan=2 class=gray>
						No Students associated with this guardian.  <BR>
						This must be done in the Family Manger after <br>
						students have been added to the family.
					</td>
				</tr>
			<%
				elseif 	rsRelations.RecordCount > 0 then
					do while not rsRelations.EOF
			%>
				<tr>
					<td class=gray>
						<% = rsRelations("Name") %>
					</td>
					<td class=gray>	
						<select name="intGuardian_Type_ID<%=rsRelations("intStudent_ID")%>" onChange="jfChanged();">
							<option>
						<%
							dim sqlGuardian_Type
							sqlGuardian_Type = "select intGuardian_Type_ID, szGuardian_Type_Desc " & _ 
											   " From trefGuardian_Type order by szGuardian_Type_Desc "
							Response.Write oFunc.MakeListSQL(sqlGuardian_Type,"","",rsRelations("intGuardian_Type_ID"))
						
						%>
						<select>
					</td>
				</tr>
			<%
						rsRelations.MoveNext
					loop	
					rsRelations.Close
					set rsRelations = nothing
				end if
			%>
			</table>
		</td>
	</tr>	
</table>
<input type=button value="Close without saving" class="btSmallGray" onClick="jfConfirm();">
<input type=submit value="Save"  class="NavSave">		
</form>

<%
call oFunc.CloseCN
set oFunc = nothing

Server.Execute(Application.Value("strWebRoot") & "Includes/footer.asp")

%>