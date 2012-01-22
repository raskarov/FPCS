<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		GoodsServiceDetail.asp
'Purpose:	Prints Goods and Services Detail
'Date:		11 March 2005
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Dimension Variables, make db Connection, print HTML header.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim rs, sql, strLast, strCurrent, curBudgetBalance

Session.Value("strTitle") = "Print Goods and Services Detail"
Session.Value("strLastUpdate") = "19 Aug 2002"
Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")
%>
<link rel="stylesheet" href="<% =  Application.Value("strWebRoot") %>CSS/printStyle.css">
<%

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()
gsList = request("GSList")
gsTest = replace(gsList, ",","")

if gsTest = "" or ucase(session.Contents("strRole")) <> "ADMIN" then
%>
	<font class=svplain10><B>Invalid Page Request.
	</b></font><br><BR>
	<input type=button value="Close Window" onClick="window.opener.focus();window.close();" class="btSmallGray" ID="Button1" NAME="Button1">
</body>
</html>
<%
	set oFunc = nothing
	Response.End
else
	' We have a request to populate an existing item from Ordered tables
%>
<script language="javascript">
	if (window.print){
	    window.print()
	}
	else {
	    alert("Mac users: please press Apple-P to print this form.\nWindows users: Please press ctrl-P to print this form.")
	}
</script>
<table width=100% ID="Table1" cellpadding="2">
<%	
	gsList = right(gsList,len(gsList)-1)
	gsList = left(gsList,len(gsList)-1)
	
	sql = "SELECT tblOrdered_Items.intOrdered_Item_ID, tblOrdered_Items.intVendor_ID, tblOrdered_Items.intItem_ID, tblOrdered_Items.intStudent_ID, tblOrdered_Items.intQty, tblOrdered_Items.curUnit_Price,  " & _ 
			" tblOrdered_Items.bolReimburse, tblOrdered_Items.intSchool_Year, tblOrdered_Items.bolApproved, tblOrdered_Items.curShipping,  " & _ 
			" tblOrd_Attrib.szValue, tblOrd_Attrib.intItem_Attrib_ID, tblSTUDENT.szFIRST_NAME + ' ' + tblSTUDENT.szLAST_NAME AS StudentName,  " & _ 
			"  tblVendors.szVendor_Name, trefItems.szName,   " & _ 
			" trefItem_Groups.szName AS Expr1, trefItem_Attrib.szName AS Label,  " & _
			" tblILP.GuardianStatusID,tblILP.SponsorStatusID,tblILP.InstructorStatusID,tblILP.AdminStatusID, tblClasses.intContract_Status_ID,tblFAMILY.szHome_Phone, tblFAMILY.szEmail, " & _ 
			" (SELECT top 1 oa2.szValue " & _
               "             FROM          tblOrd_Attrib oa2 " & _
               "             WHERE      oa2.intOrdered_Item_Id = tblOrdered_Items.intOrdered_Item_Id AND " & _
               "			 (oa2.intItem_Attrib_ID = 9 OR " & _
               "              oa2.intItem_Attrib_ID = 5 OR " & _
               "              oa2.intItem_Attrib_ID = 6 OR " & _
			   "              oa2.intItem_Attrib_ID = 22 or oa2.intItem_Attrib_ID = 33)) as iName, " & _
			   "			  CASE isNull(tblClasses.szClass_Name,'a') WHEN 'a' then CASE isNull(tblProgramOfStudies.txtCourseTitle,'a') WHEN 'a' then tblILP_SHORT_FORM.szCourse_Title  else tblProgramOfStudies.txtCourseTitle end else tblClasses.szClass_Name end as ClassLabel, " & _
			   " ps.szSubject_Name, ei.intSponsor_Teacher_ID AS Sponsor_ID, tblClasses.intInstructor_ID, tblClasses.dtApproved " & _
			"FROM tblOrd_Attrib INNER JOIN " & _ 
			" tblOrdered_Items ON tblOrd_Attrib.intOrdered_Item_ID = tblOrdered_Items.intOrdered_Item_ID INNER JOIN " & _ 
			" tblSTUDENT ON tblOrdered_Items.intStudent_ID = tblSTUDENT.intSTUDENT_ID INNER JOIN " & _ 
			" tblILP ON tblOrdered_Items.intILP_ID = tblILP.intILP_ID INNER JOIN " & _ 
			" tblVendors ON tblOrdered_Items.intVendor_ID = tblVendors.intVendor_ID INNER JOIN " & _ 
			" trefItems ON tblOrdered_Items.intItem_ID = trefItems.intItem_ID INNER JOIN " & _ 
			" tblILP_SHORT_FORM ON tblILP.intShort_ILP_ID = tblILP_SHORT_FORM.intShort_ILP_ID INNER JOIN " & _ 
			" trefItem_Groups ON trefItems.intItem_Group_ID = trefItem_Groups.intItem_Group_ID INNER JOIN " & _ 
			" trefItem_Attrib ON tblOrd_Attrib.intItem_Attrib_ID = trefItem_Attrib.intItem_Attrib_ID INNER JOIN " & _
			" tblFAMILY ON tblSTUDENT.intFamily_ID = tblFAMILY.intFamily_ID LEFT OUTER JOIN " & _ 
			" tblProgramOfStudies ON tblILP_SHORT_FORM.lngPOS_ID = tblProgramOfStudies.lngPOS_ID INNER JOIN " & _ 
			" tblClasses ON tblClasses.intClass_ID = tblILP.intClass_ID INNER JOIN " & _
			" trefPOS_Subjects ps ON ps.intPOS_SUBJECT_ID = tblClasses.intPOS_SUBJECT_ID INNER JOIN " & _
			" tblEnroll_Info ei ON ei.sintSchool_Year = tblOrdered_Items.intSchool_Year AND ei.intStudent_ID = tblSTUDENT.intStudent_ID " & _
			"WHERE (tblOrdered_Items.intOrdered_Item_ID in (" & gsList & ")) " & _ 
			"ORDER BY tblOrdered_Items.intOrdered_Item_ID, tblOrd_Attrib.intOrder "

	set rs = server.CreateObject("ADODB.RECORDSET")
	rs.CursorLocation = 3
	rs.Open sql, oFunc.FPCScnn
	
	set oBudget = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/StudentBudgetInfo.wsc"))
	do while not rs.EOF
		
		if strOld <> rs("intOrdered_Item_ID") then
			if strOld <> "" then 
				rs.MovePrevious
				vbfPriceInfo 
				rs.MoveNext
				response.Write "<tr><td colspan='2'><p></p></td></tr>"				
			end if
			oBudget.PopulateStudentFunding oFunc.FPCScnn, rs("intStudent_ID"), session.Contents("intSchool_Year") 
			curBudgetBalance = formatNumber(oBudget.BudgetBalance,2)
			call vbfHeader
		end if
		
		if rs("intItem_Attrib_ID") = "26" then
			strValue = rs("iName")
		elseif rs("intItem_Attrib_ID") = "15" then
			if rs("szValue") & "" = "1" then
				strValue = "YES"
			else
				strValue = "NO"
			end if
		else
			strValue = rs("szValue")
		end if 
		
		response.Write "<tr><td class='TableHeader'><b>" & rs("Label") & "</b></td>" & _
					   "<td class='TableCell'>" & strValue & "&nbsp;</td></tr>"
		strOld = rs("intOrdered_Item_ID")
		rs.MoveNext
	loop
	rs.MoveLast
	response.Write vbfPriceInfo
	rs.Close
	set rs = nothing
	set oBudget = nothing
end if			

call oFunc.CloseCN()
set oFunc = nothing
response.Write "</table>" 
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")

function vbfHeader
%>
	<tr>
		<td align=left>
			<img src="<% = Application("strImageRoot")%>fpcsLogo.gif">
		</td>
		<td align=right class=svplain10 width=100%>
			<% = Application.Contents("SchoolAddress") %>
		</td>
	</tr>
	<tr class=yellowHeader>	
		<Td colspan="2">
			<table align=right ID="Table26"><tr><td align=right><font face=arial size=2 color=white><% = date()%></font></td></tr></table>
			&nbsp;<b>Goods and Services  #<%=rs("intOrdered_Item_ID")%></b>									
		</td>				
	</tr>
	<tr>
		<td colspan="2">
			<table style="width:100%;" ID="Table2" cellpadding="2">
				<tr>
					<td class="TableHeader">
						<b>Student Name</b>
					</td>					
					<td class="TableHeader">
						<b>Course Title</b>
					</td>
					<td class="TableHeader">
						<b>Vendor Name</b>
					</td>					
					<td class="TableHeader">
						<b>Subject</b>
					</td>					
				</tr>				
				<tr>
					<td class="TableCell">
						<% = rs("StudentName") %>
					</td>						
					<td class="TableCell">
						<% = rs("ClassLabel")%>
					</td>
					<td class="TableCell">
						<% = rs("szVendor_Name") %>
					</td>										
					<td class="TableCell">
						<% = ucase(rs("szSubject_Name")) %>
					</td>					
				</tr>
			</table>
			<table style="width:100%;" ID="Table2" cellpadding="2">
				<tr>					
					<td class="TableHeader">
						<b>Type</b>
					</td>
					<td class="TableHeader">
						<b>Category</b>
					</td>
					<td class="TableHeader">
						<b>Course Status</b>
					</td>
					<td class="TableHeader">
						<b>Family #</b>
					</td>
					<td class="TableHeader">
						<b>Family Email</b>
					</td>
				</tr>
				<tr>					
					<td class="TableCell">
						<% if rs("bolReimburse") then response.Write "Reimbursement" else response.Write "Requisition" %>
					</td>
					<td class="TableCell">
						<% = rs("szName") %>&nbsp;
					</td>
					<td class="TableCell">
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
					<td style="width:0%" nowrap class="TableCell">
						<nobr><% = oFunc.FormatPhone(rs("szHome_Phone")) %></nobr>
					</td>	
					<td class="TableCell">
						<a href="mailto:<% = rs("szEmail") %>"><% = rs("szEmail") %></a>&nbsp;
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td colspan="2">
			<table ID="Table3" cellpadding="3">
<%
end function

function vbfPriceInfo
%>
				<tr>
					<td class="TableHeader">
						<b>Number of Units</b>
					</td>
					<td class="TableCell">
						<% = rs("intQty") %>
					</td>
				</tr>
				<tr>
					<td class="TableHeader">
						<b>Unit Price</b>
					</td>
					<td class="TableCell">
						$<% = formatNumber(rs("curUnit_Price"),2) %>
					</td>
				</tr>
				<tr>
					<td class="TableHeader">
						<b>Shipping</b>
					</td>
					<td class="TableCell">
						$<% = formatNumber(rs("curShipping"),2) %>
					</td>
				</tr>
				<tr>
					<td class="TableHeader">
						<b>Total</b>
					</td>
					<td class="TableCell">
						$<% = formatNumber(((rs("intQty")*rs("curUnit_Price")) + rs("curShipping")),2) %>
					</td>
				</tr>				
			</table>
			<table align='right' cellpadding="2">
				<tr>
					<td class="TableHeader">
						<b>Student Budget Balance</b>
					</td>
					<td class="TableCell">
						<b>$<% = curBudgetBalance %></b>
					</td>
				</tr>
			</table>
		</td>
	</tr>
<%
end function
%>
