<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		reimburseForm.asp
'Purpose:	Gives users the ability to print
'			reimbursements based on family id.
'Date:		04 JAN 2004
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

dim oFunc			 ' Main functions object
dim sql				 ' string to contain sql query commands
dim objRequest		 ' Contains the incoming form info via the request object
dim strObjValue		 ' Contains the value of an item in the request collection
dim curShippingTotal ' Total for all shipping costs
dim intFamily_Id	 ' Selects the family we are working with
dim strOC4040		 ' Contains object code totals
dim strOC4020		 ' same
dim strOC3210		 ' same
dim strOC3030		 ' same
dim intHdrCount		 ' Used to track when we need to start a new page
dim bolPrint

server.ScriptTimeout = 2200
curShippingTotal = formatNumber(0,2)

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

Session.Value("strTitle") = "Reimbursements Form"
Session.Value("strLastUpdate") = "04 Jan 2004"
Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")

' Get needed form variables   
if Request.Form.Count > 0 then
	set objRequest = Request.Form
else
	set objRequest = Request.QueryString
end if

For Each Item in objRequest	
	execute("dim " & Item)
	strObjValue = objRequest(Item)
	execute(Item & " = strObjValue")
next

set rs = server.CreateObject("ADODB.RECORDSET")
rs.CursorLocation = 3 'adUseClient

if session.Contents("intFamily_ID") <> "" then
	intFamily_Id = session.Contents("intFamily_ID")
	arFamily = oFunc.FamilyInfo(2,intFamily_Id,6)
elseif intStudent_ID <> "" then
	arFamily = oFunc.FamilyInfo(1,intStudent_ID,6)
	intFamily_ID = arFamily(0)
else
	response.Write "<h1>Page improperly called.</h1>"
	response.End
end if
sql = "SELECT     oi.intStudent_ID, tblClasses.szClass_Name, i.szName AS item_name, " & _
"CASE WHEN oi.intItem_ID = 3 THEN " & _
"                          (SELECT     vs.szVend_Service_Name " & _
"                            FROM          trefVendor_Services vs, tblOrd_Attrib oa2 " & _
"                            WHERE      oi.intOrdered_Item_ID = oa2.intOrdered_Item_ID AND oa2.intItem_Attrib_Id = 26 AND vs.intVend_Service_ID = oa2.szValue)  " & _
"                      WHEN oi.intItem_ID > 99 THEN CASE WHEN isNumeric(oa.szValue) = 1 THEN " & _
"                          ((SELECT     Name " & _
"                              FROM         INVENTORY_DETAIL_TYPES " & _
"                              WHERE     InventoryDetailTypeID = CONVERT(int, oa.szValue)) + ': ' + " & _
"                          (SELECT     ioa.szValue " & _
"                            FROM          tblOrd_Attrib ioa " & _
"                            WHERE      ioa.intOrdered_Item_ID = oi.intOrdered_Item_ID AND ioa.intItem_Attrib_Id IN (6, 9))) ELSE " & _
"                          (SELECT     ioa.szValue " & _
"                            FROM          tblOrd_Attrib ioa " & _
"                            WHERE      ioa.intOrdered_Item_ID = oi.intOrdered_Item_ID AND ioa.intItem_Attrib_Id IN (6, 9)) END " & _
"			WHEN oi.intItem_ID = 6 THEN " & _
"                          (SELECT     ioa.szValue  " & _
"                            FROM          tblOrd_Attrib ioa " & _
"                            WHERE      ioa.intOrdered_Item_ID = oi.intOrdered_Item_ID AND ioa.intItem_Attrib_Id = 5)" & _
" ELSE oa.szValue END + '(#'+cast(oi.intOrdered_Item_ID as varchar(max))+')' AS description " & _
", oi.intQty, oi.curUnit_Price, oi.curShipping, oi.bolReimburse, oi.intSchool_Year,  " & _ 
"                      oi.bolSponsor_Approved,  " & _ 
"                      oi.szDeny_Reason, (CASE WHEN oi.intItem_ID >= 5 AND oi.intItem_ID <= 8 THEN " & _ 
"                          (SELECT     szValue " & _ 
"                            FROM          tblOrd_Attrib oa2 " & _ 
"                            WHERE      oi.intOrdered_Item_ID = oa2.intOrdered_Item_ID AND oa2.intItem_Attrib_Id = 32) ELSE '' END) AS bar_code, oi.intVendor_ID,  " & _ 
"                      tblClasses.intClass_ID, (CASE WHEN oi.intItem_ID = 5 OR " & _ 
"                      oi.intItem_ID = 6 THEN " & _ 
"                          (SELECT     szValue " & _ 
"                            FROM          tblOrd_Attrib oa2 " & _ 
"                            WHERE      oi.intOrdered_Item_ID = oa2.intOrdered_Item_ID AND oa2.intItem_Attrib_Id = 15) ELSE '' END) AS consumable, i.szObject_CD,  " & _ 
"                      oi.szUSER_CREATE, " & _ 
"                          (SELECT     oa2.szValue " & _ 
"                            FROM          tblOrd_Attrib oa2 " & _ 
"                            WHERE      oi.intOrdered_Item_ID = oa2.intOrdered_Item_ID AND oa2.intItem_Attrib_Id = 28) AS Check_Number, " & _ 
"                          (SELECT     oa2.szValue " & _ 
"                            FROM          tblOrd_Attrib oa2 " & _ 
"                            WHERE      oi.intOrdered_Item_ID = oa2.intOrdered_Item_ID AND oa2.intItem_Attrib_Id = 29) AS Check_Date, " & _ 
"                          (SELECT     oa2.szValue " & _ 
"                            FROM          tblOrd_Attrib oa2 " & _ 
"                            WHERE      oi.intOrdered_Item_ID = oa2.intOrdered_Item_ID AND oa2.intItem_Attrib_Id = 30) AS Payee, " & _ 
"                          (SELECT     oa2.szValue " & _ 
"                            FROM          tblOrd_Attrib oa2 " & _ 
"                            WHERE      oi.intOrdered_Item_ID = oa2.intOrdered_Item_ID AND oa2.intItem_Attrib_Id = 31) AS Reciept_Date, oi.intILP_ID, oi.intOrdered_Item_ID,  " & _ 
"                      tblFAMILY.szHome_Phone, s.intSTUDENT_ID AS Expr1, s.intFamily_ID, s.szFIRST_NAME + ' ' +  s.szLAST_NAME AS student_name,  " & _ 
"                      trefPOS_Subjects.szSubject_Name, tblVendors.intVendor_ID AS Expr2, tblVendors.szVendor_Name, tblVendors.szVendor_Phone,  " & _ 
"                      tblVendors.szVendor_Contact, u.szName_First + ' ' + u.szName_Last AS entered_by, " & _ 
"						oi.intQty * oi.curUnit_Price + oi.curShipping AS cost,  " & _
"						ilp.GuardianStatusID,ilp.SponsorStatusID,ilp.InstructorStatusID,ilp.AdminStatusID, tblClasses.intContract_Status_ID, tblClasses.intInstructor_ID, ei.intSponsor_Teacher_ID AS Sponsor_ID " & _
"FROM         tblOrdered_Items oi INNER JOIN " & _ 
"                      tblOrd_Attrib oa ON oi.intOrdered_Item_ID = oa.intOrdered_Item_ID INNER JOIN " & _ 
"                      trefItems i ON oi.intItem_ID = i.intItem_ID INNER JOIN " & _ 
"                      trefItem_Groups ig ON i.intItem_Group_ID = ig.intItem_Group_ID INNER JOIN " & _ 
"                      tblILP ilp ON oi.intILP_ID = ilp.intILP_ID INNER JOIN " & _ 
"                      tblClasses ON ilp.intClass_ID = tblClasses.intClass_ID INNER JOIN " & _ 
"                      tblSTUDENT s ON oi.intStudent_ID = s.intSTUDENT_ID INNER JOIN " & _ 
"                      tblFAMILY ON s.intFamily_ID = tblFAMILY.intFamily_ID INNER JOIN " & _ 
"                      trefPOS_Subjects ON tblClasses.intPOS_Subject_ID = trefPOS_Subjects.intPOS_Subject_ID INNER JOIN " & _ 
"                      tblVendors ON oi.intVendor_ID = tblVendors.intVendor_ID LEFT OUTER JOIN " & _ 
"                      tblUsers u ON oi.szUSER_CREATE = u.szUser_ID INNER JOIN " & _ 
"					   tblEnroll_Info ei ON ei.intStudent_ID = s.intStudent_ID AND ei.sintSchool_Year = " & session.Contents("intSchool_Year") & " " & _
"WHERE     (oa.intOrder = 1) AND (oi.bolReimburse = 1) AND (oi.bolClosed IS NULL OR " & _ 
"                      oi.bolClosed = 0) AND (oi.intSchool_Year = " & session.Contents("intSchool_Year") & ") " 


if request("selectedIDs") <> "" then
	sql = sql & " AND oi.intOrdered_Item_ID in (" & request("selectedIDs") & ") " 
	bolPrint = true
else
	sql = sql & " and (tblFAMILY.intFamily_ID = " & intFamily_ID & ")"
end if

sql = sql &	" ORDER BY student_name, tblVendors.szVendor_Name"

'response.write sql
		
rs.Open sql,Application("cnnFPCS")'oFunc.FPCScnn


if rs.RecordCount > 0 then
	strPayee = rs("payee")
end if

'if ucase(session.contents("strUserId")) = "SCOTT" then response.write sql 


%>
<script language=javascript>
		function jfPrint(){
			if (window.print){
			window.print()
			}
			else {
			alert("Mac users: please press Apple-P to print this form.\nWindows users: Please press ctrl-P to print this form.")
			}
		}
		<% if bolPrint then response.write "jfPrint();" %>
</script>
<link rel="stylesheet" href="<% =  Application.Value("strWebRoot") %>CSS/<% if bolPrint then response.Write "printStyle.css" else response.Write "homestyle.css" %>">
<form action="reimburseForm.asp" method="post">
<input type="hidden" name="intFamily_ID" value="<% = intFamily_ID %>">
<input type="hidden" name="intStudent_ID" value="<% = intStudent_ID %>" ID="Hidden1">
<%
if rs.RecordCount > 0 then
	dim sStatus
	do while not rs.EOF
		
		
		if intHdrCount = 0 and bolPrint then 
			call vbsReimburseHeader()
			call vbsInstructions()
			call vbsTableHeader
		elseif intHdrCount = 0 then 
%>
<table width=100% ID="Table5">
	<tr class="yellowHeader" valign=middle>	
		<Td colspan=2 title="Phone #:<% = arFamily(4) %>" valign="middle">
			<table cellpadding=2 cellspacing=0 align=right ID="Table6" class="yellowHeader"><tr><td align=right valign=middle><b><% = arFamily(1) & ": " & arFamily(2)%>&nbsp;&nbsp;</b></td></tr></table>
			<b>&nbsp;Request for Reimbursement</b>
		</td>					
	</tr>
	<tr>
		<td class="svplain8">
			<b>This page will print the Reimbursement forms based on your selections below.</b>
			 <ul>
				<li>First, please select all the items below that you would like included on the Reimbursement form. </li>
			 <li>Then click the "Print Reimbursement Form' button.</b></li>
			 </ul>
			 <input name="btPrint" type=submit value="Print Reimbursement Form" class="NavSave" ID="Submit1">
		</td>
	</tr>
</table>
<br><BR>
<center>
<table ID="Table7">
<%					
			call vbsTableHeader
		end if
%>
	<tr>
		<% if not bolPrint then %>
		<td align=center valign=top title="Contact: <% = rs("szVendor_Contact")%>  Phone #:<%=rs("szVendor_Phone")%>">
			<input type="checkbox" name="selectedIDs" value="<% = rs("intOrdered_Item_ID") %>">
		</td>
		<% end if %>
		<td align=center valign=middle title="Contact: <% = rs("szVendor_Contact")%>  Phone #:<%=rs("szVendor_Phone")%>">
			<% = rs("szVendor_Name") %>
		</td>
		<td align=center valign=middle title="Entered by: <% = rs("entered_by")%>">
			<% = rs("description") %>
		</td>
		<td align=center valign=middle title="Barcode: <% = rs("bar_code") %>">
			&nbsp;<% if cstr(rs("consumable")) & "" = "1" then response.Write "Y" else response.Write "N" %>&nbsp;
		</td>
		<td align=right valign="middle"	
			title="Unit Cost:$<%= formatnumber(rs("curUnit_Price"),2)%> QTY: <%=rs("intQTY")%> S/H:$<% = formatNumber(rs("curShipping"),2)%>">
			$<% = formatNumber(rs("cost"),2) %>
		</td>
		<td align=center valign=top>
			&nbsp;&nbsp;
		</td>
		<td align=center valign=middle title="Subject: <% = rs("szSubject_Name")%>">
			<% = rs("szClass_Name") & "<BR>" %>
			<%
				if rs("AdminStatusId") = "3" or rs("SponsorStatusId") = "3" or _
					rs("InstructorStatusId") = "3" then
						'Rejected 
						response.Write "rejected"
				elseif  rs("AdminStatusId")  = "2" or rs("SponsorStatusId") = "2" then
					' Needs Work
					response.Write "must ammend"
				elseif rs("GuardianStatusId") & "" = "1" and rs("SponsorStatusId") & "" = "1" and _
					(rs("AdminStatusId") & "" = "1" or rs("intContract_Status_Id") & "" = "5") and _
					(rs("InstructorStatusId") & ""  = "1" or _
					rs("intInstructor_ID") & "" = "" or  _
					(rs("intInstructor_ID") & "" <> "" and _
					rs("intInstructor_ID") & "" = rs("Sponsor_ID") & "")) then
					
					response.Write "signed"
				else
					' Not Signed
					response.Write "not signed"
				end if  
			%>
		</td>
		<td align=center valign=middle>
			<% = rs("student_name") %>
		</td>
		<td align=center valign=middle>
			&nbsp;<% = rs("szObject_CD") %>&nbsp;
		</td>
	</tr>
	<%
		execute("strOC" & rs("szObject_CD") & "= strOC" & rs("szObject_CD") & " + " & rs("cost"))
		curTotal = formatNumber(curTotal + rs("cost"),2)
		rs.MoveNext
		intHdrCount = intHdrCount + 1
		if intHdrCount >= 6 and bolPrint then
			if not rs.EOF then
				call vbsReimburseFooter(curTotal,strOC4040,strOC4020,strOC3210,strOC3030)
				
				response.Write "<P>"
				curTotal = 0 
				strOC4040 = 0
				strOC4020 = 0
				strOC3210 = 0
				strOC3030 = 0
				intHdrCount = 0
			end if
		end if
	loop
	
	if intHdrCount > 0 and bolPrint then call vbsReimburseFooter(curTotal,strOC4040,strOC4020,strOC3210,strOC3030)
else
%>
<BR><BR>
	<center>
	<hr size=1 width=95%>
	<span class=svplain10><b>Either no reimbursements have been entered into the system<br>
		or none of the entered reimbursements have been approved by the sponsor teacher.</b>
		<br><br>Before reimbursements can show up on this form they have to be approved
		by the sponsor teacher.
	</span>
	<hr size=1 width=95%>
	</center>
<%
end if 
rs.Close
oFunc.CloseCN()
set rs = nothing
set oFunc = nothing
%>
</form>
</body>
</html>
<%
''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Procedures below here
''''''''''''''''''''''''''''''''''''''''''''''''''''''
sub vbsReimburseHeader()
%>
<table width=100% ID="Table1">
	<tr>
		<td align=left>
			<img src="<% = Application("strImageRoot")%>fpcsLogo.gif">
		</td>
		<td align=right class=svplain10 width=100%>
			<% = Application.Contents("SchoolAddress") %>
		</td>
	</tr>
	<tr class=yellowHeader valign=middle>	
		<Td colspan=2 title="Phone #:<% = arFamily(4) %>">
			<table cellpadding=2 cellspacing=0 align=right ID="Table2"><tr><td align=right><font face=veranda,arial color=white size=1><b><% = arFamily(1) & ": " & arFamily(2)%>&nbsp;&nbsp;</b></font></td></tr></table>
			<div class="svplain8" style="font-color:white"><b>&nbsp;Request for Reimbursement</b></div>
		</td>					
	</tr>
</table>
<br><BR>
<center>
<table ID="Table1">
	<tr class=svplain10>
		<td>
			<nobr><b>Check Payable to:</b></nobr>
		</td>
		<td>
			<b><u>_________________________________</u></b>
		</td>
		<td width=0%>
			&nbsp;&nbsp;
		</td>
		<td>
			<b>Date:</b>
		</td>
		<td>
			__________________
		</td>
	</tr>
	<tr class=svplain6>
		<td>
			&nbsp;
		</td>
		<td valign=top>
			<b>(person/organization to whom the check is to be written)</b>
		</td>
		<td colspan=3>
			&nbsp;
		</td>
	</tr>
</table>
<br>
</center>
<%
end sub

sub vbsInstructions()
%>
<center>
<table width=90% ID="Table3">
	<tr>
		<td class=svplain10 colspan=2>
			<b>Procedures:</b>
		</td>
	</tr>
	<tr class=svplain10>
		<td valign=top width=0%>
			1.
		</td>
		<td width=100%>
			Staple <b>Original</b> receipts or place them in an envelope 
			and attach to the back of this form.<br> <i>(remember ONLY original
			receipts with stamped company name will be accepted)</i>
		</td>
	</tr>
	<tr class=svplain10>
		<td valign=top width=0%>
			2.
		</td>
		<td>
			<b>Circle</b> items on receipt that are to be reimbursed. 
			(Please do not use a fluorescent marker. It erases the ink.)
		</td>
	</tr>
	<tr class=svplain10>
		<td valign=top width=0%>
			3.
		</td>
		<td>
			In the Requested Amount column enter, in ink, the actual amount 			
			requested for reimbursement for each item.
		</td>
	</tr>
	<tr class=svplain10>
		<td valign=top width=0%>
			4.
		</td>
		<td>
			Bring in all non-consumable items when reimbursement is requested and
			have them barcoded with FPCS barcodes. <u>No reimbursements can be made before 
			items are barcoded.</u>
		</td>
	</tr>
	<tr class=svplain10>
		<td valign=top width=0%>
			5.
		</td>
		<td>
			Make sure all items are referenced in either Box 1 or the 'Guardian Modification' 			
			Box on the student's ILP.			
		</td>
	</tr>
	<tr class=svplain10>
		<td valign=top width=0%>
			6.
		</td>
		<td>
			No single receipt may be over $150.00 with the following two exceptions: <br>			
			(a) receipts for core curriculum purchases, and <br>
			(b) receipts for services from non-profit vendors. (Reciepts cannot exceed $200)			
			<br><br> <b>Cons = Consumable</b>
		</td>
	</tr>
</table>
</center>  
<%
end sub 

sub vbsTableHeader()
%>

<table width=100% cellpadding=2 cellspacing=0 border=1 bordercolor=c0c0c0  class=svplain8 ID="Table4">
	<tr bgcolor=e6e6e6 style="font-color:white">
		<% if not bolPrint then %>
		<td align=center valign=top>
			<b>&nbsp;Select&nbsp;</b>
		</td>
		<% end if %>
		<td align=center valign=top>
			<b>Vendor Name</b>
		</td>	
		<td align=center valign=top>
			<b>Item</b>
		</td>	
		<td align=center valign=top title="Consumable">
			<b>Cons</b>
		</td>	
		<td align=center valign=top>
			<b>Budgeted Amount</b>
		</td>	
		<td align=center valign=top>
			<b>Requested Amount</b>
		</td>	
		<td align=center valign=top>
			<b>Course</b>
		</td>
		<td align=center valign=top>
			<b>Student Name</b>
		</td>
		<td align=center valign=top>
			<b>Obj CD</b>
		</td>
	</tr>
<%
end sub

sub vbsReimburseFooter(pTotal,pOC4040,pOC4020,pOC3210,pOC3030)
%>
	<tr>
		<td colspan=3 align=right>
			<b>Total</b>
		</td>
		<td align=right>
			$<%=pTotal%>&nbsp;
		</td>
		<td colspan=6>
		</td>
	</tr>
</table>
<br><BR>
<pre>
___________________________________    ________________________________
Parent Name (print)                    Parent Signature           


___________________________________    ________________________________
Sponsor Name (print)                   Sponsor Signature           


___________________________________		
FPCS Signature                             

</pre>
<span class=svplain8>
<b>ASD Object Codes:</b>
<b>4040:</b> $<% = formatNumber(pOC4040,2)%>&nbsp;
<b>4020:</b> $<% = formatNumber(pOC4020,2)%>&nbsp;
<b>3210:</b> $<% = formatNumber(pOC3210,2)%>&nbsp;
<b>3030:</b> $<% = formatNumber(pOC3030,2)%>&nbsp;
</span>
<%
end sub
%>