<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		TeacherContractAdmin.asp
'Purpose:	Admin page for accedemic review and approval of certified 
'			teacher contracts.  If contract approval is required 
'			students will not be able to enroll in a class until
'			approval has been given.
'Date:		2 MAY 2005
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim sql
dim oFunc 
dim rs

if ucase(session.Contents("strRole")) <> "ADMIN" then
	response.Write "<h1>Page Improperly Called</h1>"
	response.End
end if

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

if request("bolWin") = "" then
	Server.Execute(Application.Value("strWebRoot") & "includes/header.asp")
else
	Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")
end if

if request("updatelist") <> "" then call vbsUpdateStatus

%>
<script language=javascript>
	
	function jfPrintAll(class_ID,ilp_ID){
		var winContractApproval;
		var strURL = "<%=Application.Value("strWebRoot")%>forms/PrintableForms/allPrintable.asp?intClass_ID="+class_ID;
		strURL += "&noprint=ture&intILP_ID=" + ilp_ID ;
		winContractApproval = window.open(strURL,"winContractApproval","width=710,height=500,scrollbars=yes,resize=yes,resizable=yes");
		winContractApproval.moveTo(0,0);
		winContractApproval.focus();	
	}
	
	function jfHighLight(row){
		var obj = document.getElementById('ROW'+row);
		var lastRow = document.main.lastRow.value;
		var lastRowColor = document.main.lastRowColor.value;	
		// Reset last row to its normal state
		if (lastRow != ""){	
			var obj2 = document.getElementById('ROW'+lastRow);
			obj2.className = lastRowColor;
		}
		// Highlight current row and retsain original info
		document.main.lastRowColor.value = obj.className;
		document.main.lastRow.value = row;
		//obj.style.backgroundColor = "e6e6e6";
		obj.className = "SubHeader";
	}
	
	function jfUpdateList(id) {
		// if an item as been changed log it only once.  We will use this list
		// to determine which Contract Status' should be modified

		if (document.main.updatelist.value.indexOf(","+id+",") == -1 ) {
			document.main.updatelist.value = document.main.updatelist.value + id + ",";
		}
	}	
</script>	
<form name="main" method="post" action="TeacherContractAdmin.asp">
<input type=hidden name="lastRow" ID="Hidden1">
<input type=hidden name="LineItemsChanged" value="," ID="Hidden8">
<input type=hidden name="lastRowColor" ID="Hidden3">
<input type="hidden" name="updatelist" value="">
<input type="hidden" name="hdnReset" value="">
<input type="hidden" name="Search" value="true">
<table style="width:100%;" ID="Table3">
	<tr>
		<td class="yellowHeader">
			&nbsp;<b>Principal's ASD Contract Approval</b>
		</td>
	</tr>
	<tr>
		<td>
			<table>						
				<tr>
					<td style="width:0%;">
						<table style="width:100%;" cellpadding="2">
							<tr>
								<td class="TableHeader">
									Teachers
								</td>
								<td class="TableHeader">
									Status
								</td>
								<td rowspan="2" valign="middle">
									<input type="submit" value="Search/Save" class="NavSave" ID="Submit1" NAME="Search">
								</td>
							</tr>
							<tr>
								<td>
									<select name="intInstructor_ID" onchange="this.form.hdnReset.value='true';">
										<option value="">All Teachers
									<%
										sql = "SELECT DISTINCT tblINSTRUCTOR.intInstructor_ID, tblINSTRUCTOR.szLAST_NAME + ', ' + tblINSTRUCTOR.szFIRST_NAME as Name " & _ 
												" FROM tblINSTRUCTOR INNER JOIN " & _ 
												" tblClasses ON tblINSTRUCTOR.intINSTRUCTOR_ID = tblClasses.intInstructor_ID " & _ 
												" WHERE (tblClasses.intSchool_Year = " & session.Contents("intSchooL_Year") & ") " & _ 
												" ORDER BY Name "
										Response.Write oFunc.MakeListSQL(sql,"intInstructor_ID","Name",request("intInstructor_ID"))
									%>
									</select>
								</td>
								<td>
									<select name="intContract_Status_ID"  ID="Select1"  onchange="this.form.hdnReset.value='true';">
										<option value="">All 
									<%
										sql = "SELECT intContract_Status_ID, szContract_Status_Name " & _ 
												"FROM tblContract_Status_Types " & _ 
												"WHERE (intYear_Active_Start >= " & session.Contents("intSchool_Year") & ") " & _
												" AND (intYear_Active_End <= " & session.Contents("intSchool_Year") & ") OR " & _ 
												" (intYear_Active_End IS NULL) order by intContract_Status_ID "
										Response.Write oFunc.MakeListSQL(sql,"intContract_Status_ID","szContract_Status_Name",request("intContract_Status_ID"))	
									%>
									</select>	
								</td>			
							</tr>
						</table>
					</td>
				</tr>				
			</table>
		</td>
	</tr>
</table>
<%

if request.Form("Search") <> "" then
	sql = "SELECT c.intClass_ID, c.szClass_Name, i.intINSTRUCTOR_ID, i.szLAST_NAME, i.szFIRST_NAME, CASE isNull(c.intContract_Status_ID,1) when 1 then '1' else c.intContract_Status_ID end as intContract_Status_ID, c.dtReady_For_Review,  " & _ 
			" c.dtApproved, c.szUser_Approved, c.szComments, ilp.intILP_ID, c.szInstructor_Comments  " & _ 
			"FROM tblClasses c INNER JOIN " & _ 
			" tblINSTRUCTOR i ON c.intInstructor_ID = i.intINSTRUCTOR_ID INNER JOIN " & _ 
			" tblILP_Generic ilp ON ilp.intClass_ID = c.intClass_ID " & _
			"WHERE (c.intSchool_Year = " & session.Contents("intSchool_Year") & ") " 

	if request("intInstructor_ID") <> "" then
		sql = sql & " AND c.intINSTRUCTOR_ID = " & request("intInstructor_ID") & " " 
	end if
	
	if request("intContract_Status_ID") <> "" then
		if request("intContract_Status_ID") = 1 then
			sql = sql & " AND (c.intContract_Status_ID = " & request("intContract_Status_ID") & " or c.intContract_Status_ID is null) " 
		else
			sql = sql & " AND c.intContract_Status_ID = " & request("intContract_Status_ID") & " " 
		end if
	end if

	if request("orderby") <> "" then
		sql = sql & " ORDER BY " & request("orderby")
	else
		sql = sql & " ORDER BY i.szLast_Name, i.szFirst_Name, c.szClass_Name " 
	end if
	
'response.write sql
	set rs = server.CreateObject("ADODB.RECORDSET")
	rs.CursorLocation = 3
	rs.Open sql, oFunc.FPCScnn

	
	if request("PageNumber") <> "" and request("hdnReset") = "" then
		intPageNum = cint(request("PageNumber"))	
	else
		intPageNum = 1
	end if
	
	with rs
		if .RecordCount > 0 then
			.PageSize = 25
			.AbsolutePage = intPageNum
			intViewingTo = .AbsolutePosition + .PageSize -1 
			if intViewingTo > .recordcount then intViewingTo = .RecordCount
%>
<br>
<input type="hidden" name="PageNumber" value="<% = intPageNum%>" ID="Hidden7">
<table cellpadding="2">
	<tr>
		<td colspan=10 class="svplain8" nowrap>
			
			Viewing <% = .AbsolutePosition %> - <% = intViewingTo %>  of <% = .RecordCount %> Matches &nbsp;
			
			<table ID="Table4" cellpadding="2"><tr><td>
			<%
				if cint(.RecordCount) > cint(.PageSize) then
					for i = 1 to .PageCount
						if intViewingTo/.PageSize = i or (.RecordCount = intViewingTo and i = .PageCount) then 
							strClass = "NavSave"
						else
							strClass = "btSmallWhite"
						end if
					%>
					<input type="button" class="<% = strClass %>" value="<%=i%>" onClick="this.form.PageNumber.value='<%=i%>';this.form.submit();" ID="Button2" NAME="Button2">
					<%
					next 
				end if
			%>
			</td></tr></table>
		</td>					
	</tr>
<%			
			intCount = 0
			intCount2 = 0 
			intMax = (.AbsolutePosition + .PageSize)
			
			do while .AbsolutePosition < intMax and not .EOF
				if intCount mod 2 = 0 then
					strColor = "TableCell"
				else
					strColor = "gray"
				end if
				
				if intCount2 = 0 or intCount2 mod .PageSize = 0 then
					call PrintHeader
				end if
				
%>	
			<tr id="ROW<%=intCount%>" onClick="jfHighLight('<%=intCount%>');" class="<% = strColor %>">
				<td >
					<% = rs("szFIRST_NAME") & " " & rs("szLAST_NAME") %>
				</td>
				<td >
					<% = rs("szClass_Name") %>
				</td>
				<td >
					<% = rs("dtReady_For_Review") %>
				</td>	
				<td align="center">
					<input type="button" value="View" onCLick="jfPrintAll('<% =rs("intClass_ID")%>','<% = rs("intILP_ID") %>');" class="btSmallGray">
				</td>
				<td align="center">
					<select name="intContract_Status_ID<% = rs("intClass_Id") %>" onchange="jfUpdateList('<% = rs("intClass_ID") %>');">
						<option value="">
						<%
							sql = "SELECT intContract_Status_ID, szContract_Status_Name " & _ 
									"FROM tblContract_Status_Types " & _ 
									"WHERE (intYear_Active_Start >= " & session.Contents("intSchool_Year") & ") " & _
									" AND (intYear_Active_End <= " & session.Contents("intSchool_Year") & ") OR " & _ 
									" (intYear_Active_End IS NULL)  order by intContract_Status_ID "
							Response.Write oFunc.MakeListSQL(sql,"intContract_Status_ID","szContract_Status_Name",rs("intContract_Status_ID"))	
						%>
					</select>
				</td>				
				<td>
					<textarea name="szComments<%= rs("intClass_ID")%>" cols="30" rows="1"  onchange="jfUpdateList('<% = rs("intClass_ID") %>');" onfocus="this.rows=4;" onblur="this.rows=1;" onKeyDown="jfMaxSize(511,this);" ><% = rs("szComments") %></textarea>
					<% 
						if rs("szInstructor_Comments") & "" <> "" then
							response.Write "<br><b>Teacher Comments:</b> " &  rs("szInstructor_Comments")
						end if
					
					%>
				</td>
			</tr>
<%				
				.MoveNext
				intCount = intCount + 1
				intCount2 = intCount2 + 1
			loop
	
%>
	<input type=hidden name="intCount" value="<%=intCount%>" ID="Hidden10">
	<input type=hidden name="intCount2" value="<%=intCount2%>" ID="Hidden12">
	<input type="hidden" name="orderby" value="<% = request("orderby") %>">
</table> 
<%		
		else
			%>
			<span class="svplain8"><B>0 Matches Found.</B></span>
			<%
		end if
		.close		
	end with
end if

%>
</form>
<%
call oFunc.CloseCN()
set oFunc = nothing
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")


function PrintHeader
%>
	<Tr>
		<td class="TableHeader">
			<a href="#" class="linkWht" onclick="document.forms[0].orderby.value=' i.szLAST_NAME,i.szFIRST_NAME';document.forms[0].submit();">Instructor</a>
		</td>
		<td class="TableHeader">
			<a href="#" class="linkWht" onclick="document.forms[0].orderby.value='c.szClass_Name';document.forms[0].submit();">Class Name</a>
		</td>		
		<td class="TableHeader" align="center">
			<a href="#" class="linkWht" onclick="document.forms[0].orderby.value=' c.dtReady_For_Review';document.forms[0].submit();">Ready Date</a>
		</td>		
		<td class="TableHeader">
			Contract
		</td>
		<td class="TableHeader">
			Status
		</td>
		<td class="TableHeader">
			Comments
		</td>
	</Tr>
<%
end function

sub vbsUpdateStatus
	dim update, updateAdd
	dim list, i


	list = split(request("updatelist"),",")

'response.write request("updatelist") & "<BR>"
	for i = 0 to ubound(list)
		if list(i) <> "" then
	'response.write list(i) & " - " & i & "<BR>"

			if request("intContract_Status_ID" & list(i)) = 1 or _
				request("intContract_Status_ID" & list(i)) = 3 then
				updateAdd = " , dtReady_For_Review = NULL "
			elseif request("intContract_Status_ID" & list(i)) = 4 or request("intContract_Status_ID" & list(i)) = 5  then
				 updateAdd = " , dtApproved = CURRENT_TIMESTAMP, szUser_Approved = '" & oFunc.EscapeTick(session.Contents("strUserID")) & "' "
			end if
			update = "Update tblClasses set intContract_Status_ID = " & request("intContract_Status_ID" & list(i)) & _
					 ", szComments = '" & oFunc.EscapeTick(request("szComments" & list(i))) & "' " & updateAdd & _
					 " WHERE intClass_ID = " & list(i)
			oFunc.ExecuteCN(update)
		end if
	next
end sub
%>