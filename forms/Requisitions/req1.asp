<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		req1.asp
'Purpose:	This script displays currently ordered good/services
'			and contains buttons that link to adding additional
'			goods/services or editing/deleting existing
'Date:		23 AUG 2002
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim sqlGetILPs
dim sql
dim intCount
dim strStudent_Name
dim oFunc		'wsc object
dim strClassName
dim intPOS_Subject_ID		
dim strDeleteBt 
dim strVeiwEdit 
dim bolRejected
dim intInstructor_ID
dim strItemDesc
dim intItemCount
dim strFontColor
dim oHtml
dim bolLock

strMsg = "This page allows you to create new goods and services " & _
		 " by clicking on the 'Add a Good' or 'Add a Service' button or " & _
		 " to view or edit existing goods or services."
		 
Session.Value("strTitle") = "Requisitions"
Session.Value("strLastUpdate") = "19 Aug 2002"
Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")

If Request.Form("intStudent_ID") <> "" then
	intILP_ID = Request.Form("intILP_ID")
	intStudent_ID = Request.Form("intStudent_ID")	
elseif Request.QueryString("intStudent_ID") <> "" then
	intILP_ID = Request.QueryString("intILP_ID")
	intStudent_ID = Request.QueryString("intStudent_ID")
	intClass_id = request.QueryString("intClass_ID")
elseif Request.QueryString("intClass_ID") <> "" then
	intClass_ID = Request.QueryString("intClass_ID")
	intILP_ID = Request.QueryString("intILP_ID")
elseif Request.Form("intClass_ID") <> "" then
	intClass_ID = Request.Form("intClass_ID")
else
%>
	<font class=svplain10><B>The request to view this page is invalid.
	</b></font><br>
	<input type=button value="Home Page" onClick="window.location.href='<%=Application.Value("strWebRoot")%>';" class="btSmallGray">
</body>
</html>
<%
	Response.End
end if 

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()
'set oHtml = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/htmlFuncs.wsc"))

set rsGetName = server.CreateObject("ADODB.Recordset")
rsGetName.CursorLocation = 3
	
if intStudent_ID <> "" then	
	'Get Name of Student and Class	
	sql = "select szFirst_Name + ' ' + szLast_name as name " & _
		  "from tblStudent " & _
		  "where intStudent_ID = " & intStudent_ID
	rsGetName.Open sql,Application("cnnFPCS")'oFunc.FPCScnn
	strStudent_Name = rsGetName(0)
	
	rsGetName.Close
	
	sql = "SELECT tblClasses.szClass_Name, tblClasses.intPOS_Subject_ID " & _
			"FROM tblILP INNER JOIN " & _
			"    tblClasses ON " & _
			"    tblILP.intClass_ID = tblClasses.intClass_ID " & _
			"WHERE (tblILP.intILP_ID = " & intILP_ID & ")" 
	rsGetName.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
	
	strClassName = rsGetName(0)
	intPOS_Subject_ID = rsGetName(1)
	
	' Very important step.  This is run only if we are adding
	' a class for a student and it is only run once. 
	' (IE Run at CREATE, not on EDIT)
	' This step calls a function that will copy all class 
	' goods/services to the students account (stored
	' in tblOrdered_items/tblOrd_Attrib)
	if request.QueryString("bolFromILPInsert") <> "" then
		call vbsCopyToOrdered()
	elseif request.QueryString("intClass_Item_ID") <> "" then
		call vbsCopyToOrdered()
	end if
elseif intClass_ID <> "" then
	' Get Class Name
	sql = "Select szClass_Name,intPOS_Subject_ID,intInstructor_ID, intContract_Status_ID from tblClasses where intClass_ID = " & intClass_ID 
	rsGetName.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
	strClassName = rsGetName(0)
	intPOS_Subject_ID = rsGetName(1)
	intInstructor_ID = rsGetName(2)
	if rsGetName("intContract_Status_ID") & "" = "5" then
		bolLock = true
	else
		bolLock = false
	end if
end if 

rsGetName.Close
set rsGetName = nothing

%>
<script language="javascript">
	function jfChangeStudent(form){
	//reloads page with newly selected student
		var strURL = "<% = Application.Value("strWebRoot")%>forms/Requisitions/req1.asp?intStudent_ID=" + form.selintStudent_ID.value;
		window.open(strURL, "_self");
	}
</script>
<table width=100% bgcolor=f7f7f7 ID="Table1" cellspacing="4">
	<script language=javascript>
		function jfAddEditItem(ilp,student,intClass_ID,ExistingItemID,itemGroup){
			var winGood;
			var url;
			url = "<% = Application.Value("strWebRoot")%>forms/Requisitions/reqGoods.asp?intILP_ID=" + ilp + "&intStudent_ID=" + student;
			url += "&intClass_ID=" + intClass_ID + "&ExistingItemID=" + ExistingItemID;
			url += "&strClassName=<%= replace(strClassName,"&","AND")%>&intPOS_Subject_ID=<%=intPOS_Subject_ID%>";
			url += "&intItem_Group_ID="+itemGroup;			
			winGood = window.open(url,"winGood","width=750,height=500,scrollbars=yes,resizable=yes");
			winGood.moveTo(0,0);
			winGood.focus();
		}
		
		function jfAddGS(ilp,studentID,intClass_ID,ClassItemID){
			var winGood;
			var url;
			url = "<% = Application.Value("strWebRoot")%>forms/Requisitions/req1.asp?intILP_ID=" + ilp + "&intStudent_ID=" + studentID;
			url += "&intClass_ID=" + intClass_ID + "&intClass_Item_ID=" + ClassItemID;
			url += "&bolFromILPInsert=<% = request.QueryString("bolFromILPInsert") %>";		
			window.location.href = url;
		}
		function jfAddEditBudgetItem(ilp,student,intClass_ID,ItemID,itemGroup,budgetID){
			var winGood;
			var url;
			url = "reqGoods.asp?intILP_ID=" + ilp + "&intStudent_ID=" + student;
			url += "&intClass_ID=" + intClass_ID + "&intItem_ID=" + ItemID;
			url += "&strClassName=<%= replace(strClassName,"&","AND")%>&intPOS_Subject_ID=<%=intPOS_Subject_ID%>";
			url += "&intItem_Group_ID="+itemGroup+"&intBudget_ID="+budgetID;
			winGood = window.open(url,"winGood","width=750,height=500,scrollbars=yes,resizable=yes,status=yes");
			winGood.moveTo(0,0);
			winGood.focus();
		}
		
		<% if not oFunc.LockSpending and not oFunc.LockYear then %>
		function jfDelItem(id,name,bolOrdered_Item){
			var bolAnswer = confirm("Are you sure you want to delete '" + name + "'?")
			if (bolAnswer == true) {
				var winDelItem;
				var strURL = "reqDelete.asp?bolOrdered_Item=" + bolOrdered_Item;
				strURL += "&id=" + id;
				winDelItem = window.open(strURL,"winDelItem","width=200,height=200,resizable=no,scrollbars=no");
				winDelItem.moveTo(0,0);
				winDelItem.focus();
			}
		}
		<% end if %>
		
		function jfDeleteBudgetItem(id){
		// Opens up delete window
		var winBudget;
		var bolContinue = confirm("Are you sure you want to delete this budgeted item?");
		if (bolContinue) {
			var URL = "../budget/budgetItemTool.asp?intStudent_ID=<%=intStudent_ID%>&intBudget_ID="+id+"&delete=true";		
			winBudget = window.open(URL,"winBudget","width=1,height=1,scrollbars=yes,resizable=yes");
			winBudget.moveTo(-1000,0);
			winBudget.focus();
		}
	}
	</script>
	<tr>	
		<Td class=yellowHeader>
			&nbsp;<b>BUDGETED GOODS AND SERVICES</b> &nbsp;
		</td>
	</tr>
	<tr>
		<td class="SubHeader" bgcolor=e6e6e6>
			&nbsp;<B>Class:</b> <% = strClassName %></b>
			<% if strStudent_Name <> "" then %>
			&nbsp;&nbsp;&nbsp; <B>Student:</b> <% = strStudent_Name %>
			<% end if %>
		</td>
	</tr>
	<tr>
		<td>
			<% '= oHtml.InstructMsg(strMsg,"")%>
		</td>
	</tr>
	<tr>
		<td class=svplain10>
						<%
						' first check to see if year is locked
						if not oFunc.LockSpending and not oFunc.LockYear and not bolLock then

						%>
						<input type=button value="Add a Good" class="NavLink" onClick="jfAddEditItem('<% = intILP_ID %>','<% = intStudent_ID %>','<% = intClass_ID %>','','2');" NAME="Button1">
						<input type=button value="Add a Service" class="NavLink" onClick="jfAddEditItem('<% = intILP_ID %>','<% = intStudent_ID %>','<% = intClass_ID %>','','1');" NAME="Button2">
						<% end if %>
						<%if request.QueryString("bolFromILPInsert") <> "" then
							if intStudent_ID <> "" then'remove /
						%>
						<input type=button value="Finished"  class="NavLink" onclick="window.location.href='<%=Application.Value("strWebRoot")%>forms/packet/packet.asp?intStudent_ID=<%=intstudent_ID%><% = session.Contents("strSimpleHeader") %>';" NAME="Button4">
						<%
							elseif intInstructor_ID <> "" then
						%>
						<input type=button value="Finished"  class="NavLink" onclick="window.location.href='<%=Application.Value("strWebRoot")%>forms/teachers/viewclasses.asp?intInstructor_ID=<%=intInstructor_ID%>';" NAME="Button4">
						<%
							else
						%>
						<input type=button value="Finished"  class="NavLink" onclick="window.location.href='<%=Application.Value("strWebRoot")%>default.asp?strMessage=Class Added';" NAME="Button4">
						<%  end if 
						else
						%>
						<input type=button class="NavLink" value="Close Window" onclick="window.opener.location.reload();window.opener.focus();window.close();" id="Button4" NAME="Button5">
						<%
						end if 
						%>	
		</td>
	</tr>
</table>
<table ID="Table2">
<form name=main  method=post ID="Form2">
<input type=hidden name=intStudent_id value="<%=intStudent_id %>" ID="Hidden1">
<input type=hidden name=sintSchool_Year value="<%=oFunc.SchoolYear%>" ID="Hidden2">
<%

If intStudent_ID <> "" THEN

	sqlItems = "SELECT oi.intOrdered_Item_ID as ExistingItemID, v.szVendor_Name,oi.bolSponsor_Approved, " & _
			   "ig.szName AS grp_name,ig.intItem_Group_ID, i.szName AS item_name, ci.bolRequired, oi.bolApproved," & _
			   "oi.szDeny_Reason,oi.intQty, ((oi.intQty * oi.curUnit_Price)+oi.curShipping) as Total, " & _
               "           (SELECT     (CASE oa2.szValue " & _
               "						WHEN '0' then 'No'	" & _
               "						WHEN '1' then 'Yes'	" & _
               "						ELSE 'Not Given'		" & _
               "						END) as consum			" & _
               "            FROM          tblOrd_Attrib oa2 " & _
               "            WHERE      oa2.intOrdered_Item_Id = oi.intOrdered_Item_Id AND " & _
               "			oa2.intItem_Attrib_ID = 15) AS Consumable, " & _
               "           (SELECT     oa2.szValue " & _
               "            FROM          tblOrd_Attrib oa2 " & _
               "            WHERE      oa2.intOrdered_Item_Id = oi.intOrdered_Item_Id AND " & _
               "			oa2.intItem_Attrib_ID = 3) + ' - ' + " & _
               "           (SELECT     oa2.szValue " & _
               "            FROM          tblOrd_Attrib oa2 " & _
               "            WHERE      oa2.intOrdered_Item_Id = oi.intOrdered_Item_Id AND " & _
               "            oa2.intItem_Attrib_ID = 4) AS Dates, " & _
               "          (SELECT top 1 oa2.szValue " & _
               "             FROM          tblOrd_Attrib oa2 " & _
               "             WHERE      oa2.intOrdered_Item_Id = oi.intOrdered_Item_Id AND " & _
               "			 (oa2.intItem_Attrib_ID in (9,5,6,22,33,18)) order by oa2.intOrd_Attrib_ID) AS iName " & _
		       " FROM tblOrdered_Items oi INNER JOIN " & _
               "       tblVendors v ON oi.intVendor_ID = v.intVendor_ID INNER JOIN " & _
               "       trefItems i ON oi.intItem_ID = i.intItem_ID INNER JOIN " & _
               "       trefItem_Groups ig ON i.intItem_Group_ID = ig.intItem_Group_ID " & _
               "	   LEFT OUTER JOIN " & _
               "       tblClass_Items ci ON oi.intClass_Item_ID = ci.intClass_Item_ID " & _			   
			   " WHERE (oi.intILP_ID = " & intILP_ID & ") " & _
			   " ORDER by i.szName "
	'oa2.intItem_Attrib_ID = 9 OR " & _
     '          "              oa2.intItem_Attrib_ID = 5 OR " & _
      '         "              oa2.intItem_Attrib_ID = 6 OR " & _
		'	   "              oa2.intItem_Attrib_ID = 22 or oa2.intItem_Attrib_ID = 33
			   '" WHERE (oi.intStudent_ID = " & intStudent_ID & ") " & _
			   '" WHERE (oi.intILP_ID = " & intILP_ID & ") and (oi.bolApproved = 1 or oi.bolApproved is null) " & _
elseif intClass_ID <> "" then

	sqlItems = "SELECT ci.intClass_Item_ID as ExistingItemID,v.szVendor_Name, ig.intItem_Group_ID,ig.szName AS grp_name, i.szName AS item_name, " & _
			   "ci.intQty, ((ci.intQty * ci.curUnit_Price)+ci.curShipping) as Total, " & _
               "           (SELECT     ca2.szValue " & _
               "            FROM          tblClass_Attrib ca2 " & _
               "            WHERE      ca2.intClass_Item_Id = ci.intClass_Item_Id AND " & _
               "			ca2.intItem_Attrib_ID = 15) AS Consumable, " & _
               "           (SELECT     ca2.szValue " & _
               "            FROM          tblClass_Attrib ca2 " & _
               "            WHERE      ca2.intClass_Item_Id = ci.intClass_Item_Id AND " & _
               "			ca2.intItem_Attrib_ID = 3) + ' - ' + " & _
               "           (SELECT     ca2.szValue " & _
               "            FROM          tblClass_Attrib ca2 " & _
               "            WHERE      ca2.intClass_Item_Id = ci.intClass_Item_Id AND " & _
               "            ca2.intItem_Attrib_ID = 4) AS Dates, " & _
               "          (SELECT  top 1   ca2.szValue " & _
               "             FROM          tblClass_Attrib ca2 " & _
               "             WHERE      ca2.intClass_Item_Id = ci.intClass_Item_Id AND " & _
               "			 (ca2.intItem_Attrib_ID in (9,5,6,22,33,18)) order by ca2.intItem_Attrib_ID ) AS iName, '' as szDeny_Reason, 1 as bolApproved, intContract_Status_ID " & _
		       " FROM tblClass_Items ci INNER JOIN " & _
               "       tblVendors v ON ci.intVendor_ID = v.intVendor_ID INNER JOIN " & _
               "       trefItems i ON ci.intItem_ID = i.intItem_ID INNER JOIN " & _
               "       trefItem_Groups ig ON i.intItem_Group_ID = ig.intItem_Group_ID inner join " & _
               "	   tblClasses c ON c.intClass_ID = ci.intClass_ID " & _
			   " WHERE (ci.intClass_ID = " & intClass_ID & ")  " & _
			   "order by i.szName "
               'ca2.intItem_Attrib_ID = 9 OR " & _
              ' "              ca2.intItem_Attrib_ID = 5 OR " & _
               '"              ca2.intItem_Attrib_ID = 6 OR " & _
			   '"              ca2.intItem_Attrib_ID = 22 or ca2.intItem_Attrib_ID = 33
end if

set rsItems = server.CreateObject("ADODB.Recordset")
rsItems.CursorLocation = 3

rsItems.Open sqlItems, Application("cnnFPCS")'oFunc.FPCScnn

if rsItems.RecordCount < 1 then
%>
<table ID="Table3">
	<tr>
		<Td class=svplain8>
			<b>No Goods or Services have been added to this class.</b>
		</td>
	</tr>
</table>
<%
else
%>
<table cellpadding=3 ID="Table4">
	<tr>
		<td class=svplain8 colspan=6>
			<B>Budgeted Goods and Services.</b>
		</td>
	</tr>
	<tr>
		<td class=TableHeader align=center>
			<b>Category</b>
		</td>
		<Td class=TableHeader align=center>
			<b>Vendor</b>
		</td>		
		<Td class=TableHeader align=center>
			<b>Description</b>
		</td>
		<td class=TableHeader align=center title="Consumable Item">
			<b>Cons</b>
		</td>
		<Td class=TableHeader align=center>
			<b>Dates</b>
		</td>
		<td class=TableHeader align=center>
			<b>Total</b>
		</td>
		<td class=TableHeader align=center>
			<b>View/Edit</b>
		</td>
		<td class=TableHeader align=center>
			<b></b>
		</td>
	</tr>	
<%
	dim dblGrandTotal
	dim dblTotal
		
	do while not rsItems.EOF	
		' Determine what the user can delete/view/edit
		strVeiwEdit = "View/Edit"
		strClass = "gray"
		if rsItems("bolApproved") = false then strClass = "grayStrike"
		'response.Write rsItems("ExistingItemID") & " - " & rsItems("bolRequired") & " - " & rsItems("bolApproved") & " - " & rsItems("iName")
		if intStudent_ID <> "" then
			' added 'not admin' logic to all admins to delete all goods/services.
			if (rsItems("bolApproved") = true or rsItems("bolRequired") = true) and session.Contents("strRole") <> "ADMIN" then
				if rsItems("bolRequired") = true then
					strDeleteBt = "required"
				else
					strDeleteBt = "can't delete"
				end if
				
				strVeiwEdit = "View Only"
			elseif not oFunc.LockSpending and not oFunc.LockYear then
				strDeleteBt = "<input type=button class=btSmallGray value=""delete"" " & _
								" onClick=""jfDelItem('" & rsItems("ExistingItemID") & "','" & replace(replace(rsItems("item_name") & "","'","\'"),"""","") & ":" & replace(replace(rsItems("iName") & "","'","\'"),"""","") & "','true');"">"
			end if
		elseif rsItems("intContract_Status_ID") & "" <> "5" then
			strDeleteBt =  "<input type=button class=btSmallGray value=""delete"" " & _
							" onClick=""jfDelItem('" & rsItems("ExistingItemID") & "','" & replace(replace(rsItems("item_name") & "","'","\'"),"""","") & ":" & replace(replace(rsItems("iName") & "","'","\'"),"""","") & "','false');"">"
		else
			strVeiwEdit = "View Only"
			strDeleteBt = "can't delete"
		end if
		
%>
	<tr>
		<td class=<% = strClass %> align=center  title="<% = rsItems("grp_name") %>">
			<% = rsItems("item_name") %>
		</td>
		<Td class=<% = strClass %> align=center>
			<% = rsItems("szVendor_Name") %>
		</td>		
		<Td class=<% = strClass %> align=center>
			<% =  rsItems("iName") %>
		</td>
		<td class=<% = strClass %> align=center>
			<% = rsItems("Consumable") %>
		</td>
		<Td class=<% = strClass %> align=center>
			<% = rsItems("Dates") %>
		</td>
		<td class=<% = strClass %> align=right>
			<% 
				' Exclude item in grand total if item has been denied
				if strClass <> "grayStrike" then
					dblGrandTotal = dblGrandTotal + rsItems("Total")
				end if
			    response.Write "$" & formatNumber(rsItems("Total"),2)
			%>
		</td>		
		<td class=<% = strClass %> align=center>
			<input type=button value="<% = strVeiwEdit %>" class="btSmallGray" onClick="jfAddEditItem('<% = intILP_ID %>','<% = intStudent_ID %>','<% = intClass_ID %>','<% = rsItems("ExistingItemID") %>','<%=rsItems("intItem_Group_ID")%>');" NAME="Button3">
		</td>
		<td class=gray align=center title="<% = rsItems("szDeny_Reason") %>">
			<% = strDeleteBt %>
		</td>	
	</tr>
<%
		rsItems.MoveNext
	loop
%>
	<tr>
		<td colspan=5 class=gray align=right>
			<B>Grand Total:</b>
		</td>
		<td class=gray align=right>
			$<% = formatNumber(dblGrandTotal,2) %>
		</td>
		<td class=gray colspan=2>
			&nbsp;
		</td>		
	</tr>
</table>	
<br>
<%
end if
rsItems.Close



if intStudent_ID <> "" and intClass_ID <> "" then
	sql = "SELECT ci.intClass_Item_ID,v.szVendor_Name, ig.intItem_Group_ID,ig.szName AS grp_name, i.szName AS item_name, " & _
			   "ci.intQty, ((ci.intQty * ci.curUnit_Price)+ci.curShipping) as Total, " & _
               "           (SELECT     ca2.szValue " & _
               "            FROM          tblClass_Attrib ca2 " & _
               "            WHERE      ca2.intClass_Item_Id = ci.intClass_Item_Id AND " & _
               "			ca2.intItem_Attrib_ID = 15) AS Consumable, " & _
               "           (SELECT     ca2.szValue " & _
               "            FROM          tblClass_Attrib ca2 " & _
               "            WHERE      ca2.intClass_Item_Id = ci.intClass_Item_Id AND " & _
               "			ca2.intItem_Attrib_ID = 3) + ' - ' + " & _
               "           (SELECT     ca2.szValue " & _
               "            FROM          tblClass_Attrib ca2 " & _
               "            WHERE      ca2.intClass_Item_Id = ci.intClass_Item_Id AND " & _
               "            ca2.intItem_Attrib_ID = 4) AS Dates, " & _
               "          (SELECT  top 1   ca2.szValue " & _
               "             FROM          tblClass_Attrib ca2 " & _
               "             WHERE      ca2.intClass_Item_Id = ci.intClass_Item_Id AND " & _
               "			 (ca2.intItem_Attrib_ID = 9 OR " & _
               "              ca2.intItem_Attrib_ID = 5 OR " & _
               "              ca2.intItem_Attrib_ID = 6 OR " & _
			   "              ca2.intItem_Attrib_ID = 22 or ca2.intItem_Attrib_ID = 33) order by ca2.intItem_Attrib_ID ) AS iName, '' as szDeny_Reason, 1 as bolApproved, intContract_Status_ID " & _
		       " FROM tblClass_Items ci INNER JOIN " & _
               "       tblVendors v ON ci.intVendor_ID = v.intVendor_ID INNER JOIN " & _
               "       trefItems i ON ci.intItem_ID = i.intItem_ID INNER JOIN " & _
               "       trefItem_Groups ig ON i.intItem_Group_ID = ig.intItem_Group_ID inner join " & _
               "	   tblClasses c ON c.intClass_ID = ci.intClass_ID " & _
			   " WHERE (ci.intClass_ID = " & intClass_ID & ") AND (ci.bolRequired = 0 OR " & _ 
			   "	ci.bolRequired IS NULL) AND (NOT EXISTS " & _ 
			   "	(SELECT	'x' " & _ 
			   "	FROM	tblOrdered_Items oi " & _ 
			   "	WHERE	oi.intClass_Item_ID = ci.intClass_Item_ID AND oi.intStudent_ID = " & intStudent_ID & ")) " & _
			   "order by i.szName "
		
	rsItems.Open sql, Application("cnnFPCS")'oFunc.Fpcscnn
	
	if rsItems.RecordCount > 0 then
		dblGrandTotal = 0
%>
<table cellpadding=3 ID="Table5">
	<tr>
		<td class=svplain8 colspan=6>
			<B>Suggested but not Required Goods/Services.</b><br>
			The Goods/Services below have been suggested by the instructor of this
			course but they are not required. Currently you have <b>NOT</b> purchased the
			items below. If you would like to purchase any item below click the corresponding <b>ADD</b>
			button.
		</td>
	</tr>
	<tr>
		<td class=TableHeader align=center>
			<b>Category</b>
		</td>
		<Td class=TableHeader align=center>
			<b>Vendor</b>
		</td>		
		<Td class=TableHeader align=center>
			<b>Description</b>
		</td>
		<td class=TableHeader align=center title="Consumable Item">
			<b>Cons</b>
		</td>
		<Td class=TableHeader align=center>
			<b>Dates</b>
		</td>
		<td class=TableHeader align=center>
			<b>Total</b>
		</td>
		<td class=TableHeader align=center>
			&nbsp;
		</td>
	</tr>	
<%		
	do while not rsItems.EOF
%>
	<tr>
		<td class=<% = strClass %> align=center  title="<% = rsItems("grp_name") %>">
			<% = rsItems("item_name") %>
		</td>
		<Td class=<% = strClass %> align=center>
			<% = rsItems("szVendor_Name") %>
		</td>		
		<Td class=<% = strClass %> align=center>
			<% =  rsItems("iName") %>
		</td>
		<td class=<% = strClass %> align=center>
			<% = rsItems("Consumable") %>
		</td>
		<Td class=<% = strClass %> align=center>
			<% = rsItems("Dates") %>
		</td>
		<td class=<% = strClass %> align=right>
			<% 
				' Exclude item in grand total if item has been denied
				if strClass <> "grayStrike" then
					dblGrandTotal = dblGrandTotal + rsItems("Total")
				end if
			    response.Write "$" & formatNumber(rsItems("Total"),2)
			%>
		</td>		
		<td class=<% = strClass %> align=center>
			<input type=button value="ADD" class="btSmallGray" onClick="jfAddGS('<% = intILP_ID %>','<% = intStudent_ID %>','<% = intClass_ID %>','<% = rsItems("intClass_Item_ID") %>');" NAME="Button3" ID="Button1">
		</td>
	</tr>
<%
			rsItems.MoveNext
		loop
%>
	<tr>
		<td colspan=5 class=gray align=right>
			<B>Grand Total:</b>
		</td>
		<td class=gray align=right>
			$<% = formatNumber(dblGrandTotal,2) %>
		</td>
		<td class=gray colspan=2>
			&nbsp;
		</td>		
	</tr>
</table>	
<br>
<%
	end if 
	rsItems.Close
end if 

set rsItems = nothing
call oFunc.CloseCN
set oFunc = nothing

Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")

sub vbsCopyToOrdered()
	' This funtion shold only be called when a class is being added for a student.
	' This function takes all the goods/services that are associated with the class_id
	' and copies the records from tblClass_Items and tblClass_Attrib and stores them
	' in tblOrdered_Items and tblOrdered_Attrib creating records for goods and services 
	' that are now ties to the student/class/ilp.
	dim intOrdered_Item_ID_s
	dim insert_s
	dim update_s
	dim sql_s
	dim intRC_s
	
	
	if request.QueryString("intClass_Item_ID") & "" = "" then
		' Verify that no other items have been added for this ilp
		' If so then the copying of items has already took place
		sql_s = "select * from tblOrdered_Items where intILP_ID = " & intILP_ID
		set rsVerify = server.CreateObject("ADODB.RECORDSET")
		rsVerify.CursorLocation = 3
		rsVerify.Open sql_s, Application("cnnFPCS")'oFunc.FPCScnn
		intRC_s = rsVerify.recordcount
		rsVerify.close	
		set rsVerify = nothing
		if intRC_s > 0 then
			exit sub
		end if 
	end if
		
	' Create ADO RS Objects
	set rsGetGS = server.CreateObject("ADODB.RECORDSET")
	rsGetGS.CursorLocation = 3
	
	set rsGetAttr = server.CreateObject("ADODB.RECORDSET")
	rsGetAttr.CursorLocation = 3	
	
	' Get all item records for class
	sql_s = "SELECT ci.intClass_Item_ID, ci.intItem_ID, ci.intQty, ci.curUnit_Price,  " & _
			"ci.curShipping, ci.intSchool_Year, ci.intVendor_ID, i.intItem_Group_ID " & _
			"FROM tblClass_Items ci INNER JOIN " & _
			" trefItems i ON ci.intItem_ID = i.intItem_ID " 
	
	if request.QueryString("intClass_Item_ID") & "" = "" then
		sql_s = sql_s & "WHERE (ci.intClass_ID =" & intClass_ID & ") and ci.bolRequired = 1 "
	else
		sql_s = sql_s & "WHERE (ci.intClass_Item_ID = " & 	request.QueryString("intClass_Item_ID") & ")"
	end if 
			
	rsGetGS.Open sql_s, Application("cnnFPCS")'oFunc.FPCScnn
	
	'copy class items to ordered items
	do while not rsGetGS.EOF
		insert_s = "insert into tblOrdered_Items (" & _
				 "intVendor_ID,intILP_ID,intItem_ID,intStudent_ID,intQTY," & _
				 "curUnit_Price,curShipping,bolReimburse,intSchool_Year,intClass_Item_ID) " & _
				 "values (" & _
				 rsGetGS("intVendor_ID") & "," & _
				 intILP_ID & "," & _
				 rsGetGS("intItem_ID") & "," & _
				 intStudent_ID & "," & _
				 rsGetGS("intQTY") & "," & _
				 rsGetGS("curUnit_Price") & "," & _
				 rsGetGS("curShipping") & "," & _
				 "'0'," & _
				 session.Contents("intSchool_Year") & "," & _
				 rsGetGS("intClass_Item_ID") & ")" 
		oFunc.ExecuteCN(insert_s)
		
		intOrdered_Item_ID_s = oFunc.GetIdentity
		
		sql_s = "select intItem_Attrib_ID, szValue, intOrder " & _
				"from tblClass_Attrib " & _
				"where intClass_Item_ID = " & rsGetGS("intClass_Item_ID")
		rsGetAttr.Open sql_S,Application("cnnFPCS")'oFunc.FPCScnn
		
		' Now we copy all attributes for a given item
		intItemCount = 0
		do while not rsGetAttr.EOF
			insert_s = "insert into tblOrd_Attrib(" & _
					   "intOrdered_Item_ID,intItem_Attrib_ID,szValue,intOrder) " & _
					   "values (" & _
					   intOrdered_Item_ID_s & "," & _
					   rsGetAttr("intItem_Attrib_ID") & ",'" & _
					   oFunc.EscapeTick(rsGetAttr("szValue")) & "'," & _
					   rsGetAttr("intOrder") & ")" 					   
			oFunc.ExecuteCN(insert_s)
			if intItemCount = 0 then
				strItemDesc = oFunc.EscapeTick(rsGetAttr("szValue"))
			end if
			intItemCount = intItemCount + 1
			rsGetAttr.MoveNext
		loop
		rsGetAttr.Close
		rsGetGS.MoveNext	 
	loop
	
	' Clean up rs objects	
	set rsGetAttr = nothing
	rsGetGS.Close
	set rsGetGS = nothing
end sub
%>


