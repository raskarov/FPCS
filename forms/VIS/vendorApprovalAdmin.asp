<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		vendorApprovalAdmin.asp
'Purpose:	Gives admin the abilty to view and approve/deny Vendors
'Date:		21 JAN 2003
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'on error resume next
' Dimension Variables
dim intCount			'used as a tracking mechanism for our update subroutine 
dim sql					'sql that helps us populate our form		
dim rsVendors			'recordset that helps us populate our form
dim strOrderBy			'order by for sql statement
dim intNumToShow		'Number of rows to return in sql statement
dim strMessage			'used to displays javascript messages
dim strColor			'used to set alternating bgcolor for rows in table

' Only Admins may view this page. Stop all others
if session.Contents("strRole") <> "ADMIN" then
%>
<html>
<body>
<h1>Page Improperly Called.</h1>
</body>
</html>
<%
	response.End
end if

Session.Value("strTitle") = "Vendor Approval Page"
Session.Value("strLastUpdate") = "21 JAN 2003"
Server.Execute(Application.Value("strWebRoot") & "includes/Header.asp")

'Create object containing all of our FPCS functions
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

'Check to see if we need to save any changes
if request.Form("intCount") <> "" then
	call vbsSaveChanges
end if

set rsVendors = server.CreateObject("ADODB.Recordset")
rsVendors.CursorLocation = 3 'adUseClient

' Next if statements help filter/order our list
if request.Form("numberToShow") <> "" then
	intNumToShow = request.Form("numberToShow")
else
	intNumToShow = 25
end if 

if request.Form("orderBy") <> "" then
	strOrderBy = "Order by " & request.Form("orderBy")
else
	strOrderBy = "ORDER BY v.szVendor_Name "
end if

' Lets get the list of Goods/Sevices that need approval
sql = "SELECT TOP " & intNumToShow & " intVendor_ID, szVendor_Name, szVendor_Phone, bolApproved,  " & _
		"bolNonProfit, bolCertASDTeacher, bolCurrentASDEmp, bolEligibleHire, bolConflictIntWaiver, " & _
		"bolConflictIntWaiverRecvd, bolIncludeDir, dtCREATE, dtMODIFY, szUSER_CREATE, szUSER_MODIFY  " & _
		"FROM tblVendors v  " & _
		"WHERE (bolApproved IS NULL)  " & _
		strOrderBy
	  
'Response.Write sql
'Response.End	  
rsVendors.Open sql, Application("cnnFPCS")'oFunc.FPCScnn

'Start by printing title
%>
<table width=100% ID="Table1">
	<form action="vendorApprovalAdmin.asp" method=post ID="Form1">
	<tr>	
		<Td class=yellowHeader>
			&nbsp;<b>Vendor Approval Admin</b>
		</td>
	</tr>
<%

if rsVendors.RecordCount > 0 then
	' We've got some records so let's make the form
%>
<script language=javascript>
<!-- hide from browsers
	function jfShowVendor(id){
		var winVendor;
		var strURL;		
		strURL = "vendorAdmin.asp?blnFromItems=yes&intVendor_ID=" + id;
		winVendor = window.open(strURL,"winVendor","width=710,height=500,scrollbars=yes,resizable=yes");
		winVendor.moveTo(0,0);
		winVendor.focus();
	}
	
	function jfViewClass(class_id,instructor_id,instruct_type,intContract_Guardian_ID,intGuardian_ID) {
		var winClass;
		var strURL = "../Teachers/classAdmin.asp?bolInWindow=true&plain=yes<%=strDisabled%>&intClass_id="+class_id;
		strURL += "&intInstructor_id="+instructor_id+"&intInstruct_Type_ID="+instruct_type;
		strURL += "&intContract_Guardian_ID="+intContract_Guardian_ID;
		strURL += "&intGuardian_id="+intGuardian_ID;
		winClass = window.open(strURL,"winClass","width=640,height=500,scrollbars=yes,resizable=yes");
		winClass.moveTo(0,0);
		winClass.focus();
	}
	
	function jfViewItem(ilp,student,ExistingItemID,itemGroup,className,pos_id){
			var winItem;
			var url;
			url = "reqGoods.asp?intILP_ID=" + ilp + "&intStudent_ID=" + student;
			url += "&ExistingItemID=" + ExistingItemID;
			url += "&intItem_Group_ID="+itemGroup;
			url += "&strClassName=" + className + "&intPOS_Subject_ID=" + pos_id;
			winItem = window.open(url,"winItem","width=750,height=500,scrollbars=yes,resizable=yes");
			winItem.moveTo(0,0);
			winItem.focus();
	}
	<% = strMessage %>
-->
</script>
	<tr>
		<td>
			<table>
				<tr>
					<td	class=gray>
						<b>Show First:</b>
					</td>
					<td>
						<select name="numberToShow">
							<option value="">
							<%
								response.Write oFunc.makeList("25,50,100,150,200,300,500","25,50,100,150,200,300,500",request.Form("numberToShow"))							
							%>
						</select>
					</td>
					<td	class=gray>
						<b>Order By:</b>
					</td>
					<td>
						<select name="orderby" ID="Select1">
							<option value="">
							<%
								response.Write oFunc.makeList("v.dtCREATE ,v.dtCREATE DESC, v.szVendor_Name","Oldest Date,Newest Date, Vendor Name",request.Form("orderby"))							
							%>
						</select>
					</td>
					<td>
					<input type=submit value="Refresh List And Submit Approval/Denials" class="NavSave">
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<table bgcolor=f7f7f7>	
	<tr>
		<td class=gray align=center>
			&nbsp;Vendor Name&nbsp;
		</td>
		<td class=gray align=center>
			&nbsp;Non Profit&nbsp;
		</td>
		<td class=gray align=center>
			&nbsp;Retired<br>Certificated<br>ASD Teacher&nbsp;
		</td>		
		<td class=gray align=center>
			&nbsp;Current ASD<br>Employee&nbsp;
		</td>
		<td class=gray align=center>
			&nbsp;Eligible for Hire&nbsp;
		</td>
		<td class=gray align=center title="Unit Price">
			&nbsp;Conflict<br>Of Interest<br>Waiver&nbsp;
		</td>
		<td class=gray align=center>
			&nbsp;Conflict Waiver<br>Recvd&nbsp;
		</td>
		<td class=gray align=center title="Approve Vendor">
			&nbsp;Aprv&nbsp;
		</td>
		<td class=gray align=center>
			&nbsp;Include In<br>Directory&nbsp;
		</td>
		<td class=gray align=center>
			&nbsp;Deny&nbsp;
		</td>
	</tr>
<%
	intCount = 0
	do while not rsVendors.EOF
		' Set row color
		if intCount mod 2 = 0 then
			strColor = " bgcolor='white' "
		else
			strColor = ""
		end if
%>
	<tr <% = strColor %>>
		<input type=hidden name="intVendor_ID<%=intCount%>" value="<%=rsVendors("intVendor_ID")%>">
		<td class="TableCell" title="Phone: <% = rsVendors("szVendor_Phone") %>  CreatedBy:  <% = rsVendors("szUSER_CREATE") %> on <% = rsVendors("dtCREATE") %> ">
			<a title="View Vendor Profile" 
				href="javascript:" onclick="jfShowVendor(<%=rsVendors("intVendor_ID")%>);">
			<% = rsVendors("szVendor_Name") %></a>
		</td>
		<td class="TableCell" align="center">
			<%	if IsNull(rsVendors("bolNonProfit")) then 
					Response.Write "&nbsp;"
				elseif cbool(rsVendors("bolNonProfit")) then 
					Response.Write "Yes"
				else 
					Response.Write "No"
				end if 
			%>
		</td>
		<td class="TableCell" align="center">			
			<%	if IsNull(rsVendors("bolCertASDTeacher")) then 
					Response.Write "&nbsp;"
				elseif cbool(rsVendors("bolCertASDTeacher")) then 
					Response.Write "Yes"
				else 
					Response.Write "No"
				end if 
			%>
		</td>
		<td class="TableCell" align="center">
			<%	if IsNull(rsVendors("bolCurrentASDEmp")) then 
					Response.Write "&nbsp;"
				elseif cbool(rsVendors("bolCurrentASDEmp")) then 
					Response.Write "Yes"
				else 
					Response.Write "No"
				end if 
			%>
		</td>
		<td class="TableCell" align="center">
			<%	if IsNull(rsVendors("bolEligibleHire")) then 
					Response.Write "&nbsp;"
				elseif cbool(rsVendors("bolEligibleHire")) then 
					Response.Write "Yes"
				else 
					Response.Write "No"
				end if 
			%>
		</td>
		<td class="TableCell" align="center">
			<%	if IsNull(rsVendors("bolConflictIntWaiver")) then 
					Response.Write "&nbsp;"
				elseif cbool(rsVendors("bolConflictIntWaiver")) then 
					Response.Write "Yes"
				else 
					Response.Write "No"
				end if 
			%>
		</td>
		<td class="TableCell" align="center">
			<%	if IsNull(rsVendors("bolConflictIntWaiverRecvd")) then 
					Response.Write "&nbsp;"
				elseif cbool(rsVendors("bolConflictIntWaiverRecvd")) then 
					Response.Write "Yes"
				else 
					Response.Write "No"
				end if 
			%>
		</td>
		<td align=center class="TableCell">
			<input type=checkbox name="approved<% = intCount%>" value="<% = rsVendors("intVendor_ID")%>" onclick="bolIncludeDir<% = intCount%>.disabled = false;">
		</td>
		<td align=center class="TableCell">
			<select name="bolIncludeDir<% = intCount%>" disabled>
				<option value="">
				<%
					strValues = "TRUE,FALSE"
					strText = "Yes,No"								
					Response.Write oFunc.MakeList(strValues,strText,oFunc.TFText(rsVendors("bolIncludeDir")))
				%>			
			</select>				
		</td>
		<td class="TableCell">
			<textarea cols=20 rows=1 wrap=virtual name="denied<% = intCount%>" onfocus="this.rows=4;" onblur="this.rows=1;" onKeyDown="jfMaxSize(511,this);"></textarea>
		</td>
	</tr>
<%
		rsVendors.MoveNext
		intCount = intCount + 1
	loop
%>
	<tr>
		<td colspan=9>
		<input type=hidden name="intCount" value="<%=intCount%>">
		<input type=submit value="Submit Approval/Denials" class="NavSave">
		</td>
	</tr>
</form>
</table>
		</td>
	</tr>
</table>
<%
else 
' No items returned in recordset so display message
%>
	<tr>
		<td>
			<center><font face=arial size=2><b>No Vendors need to be approved.</b></font></center>
		</td>
	</tr>
</table>
<%
end if

rsVendors.Close
set rsVendors = nothing

call oFunc.CloseCN()
set oFunc = nothing
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")

sub vbsSaveChanges
	' This sub cycles through all of the rows from the form and picks out
	' the approved vendors, sets them to approved in tblVendors and 
	' takes the denied items and sets them to denied and saves the reason.
	' If neither approved or denied there is no action.
	dim update
	dim intLocalCount
	dim bolIncludeDir
	
	intLocalCount = 0 
	oFunc.BeginTransCN
	for i = 0 to request.Form("intCount")
		bolIncludeDir = oFunc.ConvertCheckToBit(Request("bolIncludeDir"&i))	
		if request.Form("Approved"&i) <> "" then
			update = "update tblVendors set bolApproved = 1, " & _
			"bolIncludeDir = " & bolIncludeDir & ", " & _
			"szUser_Modify = '" & session.Value("strUserID") & "' " & _
			"where intVendor_ID = " & request.Form("intVendor_ID"&i) 
			oFunc.ExecuteCN(update)
			intLocalCount = intLocalCount + 1
		elseif request.Form("denied"&i) <> "" then
			update = "update tblVendors set bolApproved = 0, " & _
					 "szDeny_Reason = '" & oFunc.escapeTick(request.Form("denied"&i)) & "', " & _
					 "bolIncludeDir = " & bolIncludeDir & ", " & _
					 "szUser_Modify = '" & session.Value("strUserID") & "' " & _
					 "where intVendor_ID = " & request.Form("intVendor_ID"&i) 
			oFunc.ExecuteCN(update)
			intLocalCount = intLocalCount + 1
		end if 
	next
	oFunc.CommitTransCN
	
	'Send message only if we made updates
	if intLocalCount > 0 then
		strMessage = "alert('" &intLocalCount & " Approval/Denials have been recorded');"	 
	end if 
end sub
%>
