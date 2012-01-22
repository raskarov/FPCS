<%@ Language=VBScript %>
<%
Server.Execute(Application.Value("strWebRoot") & "includes/Header.asp")

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

dim oHtml
dim dblTotalCharge,dblTotalBudget, dblClassBudget, dblClassCharge
set oHtml = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/htmlFuncs.wsc"))

dim sql, rs, ForceVerify

set rs = server.CreateObject("ADODB.RECORDSET")
rs.CursorLocation = 3

if session.Contents("HasUpdated" & session.Contents("intVendor_ID")) & "" = "" then
	' We need to check to see if the vendor has verified their profile
	' If not we will need to force them to do so.
	sql = "SELECT     TOP 1 bolProfile_Verified, szVendor_Status_CD " & _ 
					"	FROM          tblVendor_Status vs " & _ 
					"	WHERE      vs.intVendor_ID = " & session.Contents("intVendor_ID") & " AND vs.intSchool_Year <= " & session.Contents("intSchool_Year") &  " " & _ 
					"	ORDER BY intSchool_Year DESC,intVendor_Status_ID DESC"
					
	rs.Open sql, oFunc.FpcsCnn
	
	if rs.RecordCount > 0 then
		if rs("szVendor_Status_CD") = "REJC" then
			session.Contents("Rejected" & session.Contents("intVendor_ID")) = true
		end if
		
		if rs("bolProfile_Verified") then
			ForceVerify = false
			session.Contents("HasUpdated" & session.Contents("intVendor_ID")) = true
		else
			ForceVerify = true
			session.Contents("HasUpdated" & session.Contents("intVendor_ID")) = false
		end if
	else
		ForceVerify = true
		session.Contents("HasUpdated" & session.Contents("intVendor_ID")) = false
	end if
	
	rs.Close
else
	ForceVerify = not session.Contents("HasUpdated" & session.Contents("intVendor_ID"))
end if

session.Contents("HasUpdated" & session.Contents("intVendor_ID")) = true
ForceVerify = false


sql = "SELECT	tblOrdered_Items.intOrdered_Item_ID, tblVendors.szVendor_Name, tblSTUDENT.szLAST_NAME, tblSTUDENT.szFIRST_NAME,  " & _ 
		"	SUM(tblOrdered_Items.intQty * tblOrdered_Items.curUnit_Price + tblOrdered_Items.curShipping) AS total, tblVendors.intVendor_ID,  " & _ 
		"	tblOrdered_Items.intILP_ID, tblVendors.szDeny_Reason AS vendorDeny, f.szFamily_Name, f.szDesc, f.szHome_Phone, f.szEMAIL,  " & _ 
		"	tblSTUDENT.intSTUDENT_ID, tblILP.intClass_ID, tblOrdered_Items.curShipping, tblOrdered_Items.bolClosed, tblOrdered_Items.bolApproved,  " & _ 
		"	tblOrdered_Items.bolReimburse, tblOrdered_Items.curUnit_Price, tblOrdered_Items.intQty, " & _ 
		"		(SELECT	TOP 1 oa2.szValue " & _ 
		"			FROM	tblOrd_Attrib oa2 " & _ 
		"			WHERE	oa2.intOrdered_Item_Id = tblOrdered_Items.intOrdered_Item_Id AND (oa2.intItem_Attrib_ID = 9 OR " & _ 
		"				oa2.intItem_Attrib_ID = 5 OR " & _ 
		"				oa2.intItem_Attrib_ID = 6 OR " & _ 
		"				oa2.intItem_Attrib_ID = 22 OR " & _ 
		"				oa2.intItem_Attrib_ID = 33) " & _ 
		"			ORDER BY oa2.intOrd_Attrib_ID) AS oiDesc, CASE isNull(tblClasses.szClass_Name, 'a')  " & _ 
		"		WHEN 'a' THEN CASE isNull(tblProgramOfStudies.txtCourseTitle, 'a')  " & _ 
		"		WHEN 'a' THEN tblILP_SHORT_FORM.szCourse_Title ELSE tblProgramOfStudies.txtCourseTitle END ELSE tblClasses.szClass_Name END AS ClassLabel, " & _ 
		"		i.szName, i.intItem_Group_ID, tblOrdered_Items.szDeny_Reason AS oiComment, " & _ 
		"		tblILP.GuardianStatusId, tblILP.SponsorStatusId, tblILP.AdminStatusId, " & _
		"		tblClasses.intInstructor_ID, tblILP.InstructorStatusId, tblClasses.intContract_Status_Id, " & _
		"		ei.intSponsor_Teacher_ID as Sponsor_ID " & _
		"FROM	tblVendors INNER JOIN " & _ 
		"	tblOrdered_Items ON tblVendors.intVendor_ID = tblOrdered_Items.intVendor_ID INNER JOIN " & _ 
		"	tblSTUDENT ON tblOrdered_Items.intStudent_ID = tblSTUDENT.intSTUDENT_ID INNER JOIN " & _ 
		"	tblILP ON tblOrdered_Items.intILP_ID = tblILP.intILP_ID INNER JOIN " & _ 
		"	tblILP_SHORT_FORM ON tblILP.intShort_ILP_ID = tblILP_SHORT_FORM.intShort_ILP_ID LEFT OUTER JOIN " & _ 
		"	tblProgramOfStudies ON tblILP_SHORT_FORM.lngPOS_ID = tblProgramOfStudies.lngPOS_ID LEFT OUTER JOIN " & _ 
		"	tblFAMILY f ON tblSTUDENT.intFamily_ID = f.intFamily_ID INNER JOIN " & _ 
		"	tblClasses ON tblClasses.intClass_ID = tblILP.intClass_ID INNER JOIN " & _ 
		"	trefItems i ON tblOrdered_Items.intItem_ID = i.intItem_ID inner JOIN " & _ 
		"	tblEnroll_Info ei on tblStudent.intStudent_ID = ei.intStudent_ID and ei.sintSchool_Year = tblOrdered_Items.intSchool_Year " & _
		"WHERE	(tblOrdered_Items.intSchool_Year = " & session.Contents("intSchool_Year") & ") AND " & _
		"	(tblOrdered_Items.intVendor_ID = " & session.Contents("intVendor_ID") & ")  AND (tblILP.SponsorStatusId <> 3 OR " & _ 
		"	tblILP.SponsorStatusId IS NULL) AND (tblILP.AdminStatusId <> 3 OR " & _ 
		"	tblILP.AdminStatusId IS NULL) AND (tblILP.InstructorStatusId <> 3 OR " & _ 
		"	tblILP.InstructorStatusId IS NULL) " & _ 
		" and tblILP.GuardianStatusId = 1 and tblILP.SponsorStatusId = 1 and " & _
		" (tblILP.AdminStatusId = 1 or tblClasses.intContract_Status_Id = 5) and " & _
		" (tblILP.InstructorStatusId = 1 or tblClasses.intInstructor_ID is null or " & _
		" (tblClasses.intInstructor_ID is not null and tblClasses.intInstructor_ID = ei.intSponsor_Teacher_ID)) " & _
		"GROUP BY tblOrdered_Items.intOrdered_Item_ID, tblVendors.szVendor_Name, tblSTUDENT.szLAST_NAME, tblSTUDENT.szFIRST_NAME,  " & _ 
		"	tblVendors.intVendor_ID, tblOrdered_Items.intILP_ID, tblVendors.szDeny_Reason, f.szFamily_Name, f.szDesc, f.szHome_Phone, f.szEMAIL,  " & _ 
		"	tblSTUDENT.intSTUDENT_ID, tblILP.intClass_ID, tblOrdered_Items.curShipping, tblOrdered_Items.bolClosed, tblOrdered_Items.bolApproved,  " & _ 
		"	tblOrdered_Items.bolReimburse, tblOrdered_Items.curUnit_Price, tblOrdered_Items.intQty, CASE isNull(tblClasses.szClass_Name, 'a')  " & _ 
		"	WHEN 'a' THEN CASE isNull(tblProgramOfStudies.txtCourseTitle, 'a')  " & _ 
		"	WHEN 'a' THEN tblILP_SHORT_FORM.szCourse_Title ELSE tblProgramOfStudies.txtCourseTitle END ELSE tblClasses.szClass_Name END, i.szName,  " & _ 
		"	i.intItem_Group_ID, tblOrdered_Items.szDeny_Reason, " & _ 
		"		tblILP.GuardianStatusId, tblILP.SponsorStatusId, tblILP.AdminStatusId, " & _
		"		tblClasses.intInstructor_ID, tblILP.InstructorStatusId, tblClasses.intContract_Status_Id, " & _
		"		ei.intSponsor_Teacher_ID " & _
		"ORDER BY tblVendors.szVendor_Name, tblSTUDENT.szLAST_NAME, tblSTUDENT.szFIRST_NAME "

'response.Write sql
rs.Open sql, oFunc.FpcsCnn

%>
<script language=javascript>
	function jfToggle(pList,pID){
		// toggles display of objs in pList on and off
		var arList = pList.split(",");
		var i;
		var obj;
		var sText;
		for(i=0;i< arList.length;i++){
			if (arList[i] != '') {
				obj = document.getElementById('div'+arList[i]);
				if (obj.style.display == 'none'){
					obj.style.display = '';
					sText = 'hide';
				}else{
					obj.style.display = 'none';
					sText = 'show';
				}
			}
		}
		if (pID != ''){
			obj = document.getElementById(pID);
			obj.innerHTML = sText;
		}
	}
</script>	
<table style="width:100%;" ID="Table3" cellpadding="2">
	<tr>
		<td class="yellowHeader" colspan="10">
			&nbsp;<B>Vendor Home Page</B>
		</td>
	</tr>
<% if session.Contents("Rejected" & session.Contents("intVendor_ID")) then %>
	<tr>
		<td class="svplain10">
			<B>Your account currently has a status of 'Rejected' 
			<% if rs("szDeny_Reason") & "" <> "" then %>
			due to the following reason ... <br><br>
			<span class="svError"><% = rs("szDeny_Reason") %></span>
			<% else %>
			.
			<% end if %>
			<br><br>
			You will not be able to access the Vendor Online System until your status
			is changed.		</B>
		<br><br>
		</td>
	</tr>
<% else %>
	<tr>
		<td class="svplain10">
			<B>Welcome to your FPCS Vendor Management page!</B><BR><br>
			From here you can manage your profile and 
			track activity.  Your profile contains your information which acts as your store front in the FPCS online system.
			Teachers and Guardians can search the vendor database to find matches based on data you enter in your profile
			so be sure to keep it <b><i>updated</i></b>. Click below to view your profile.<br><br>
			<b><a href="<% = Application.Value("strWebRoot") %>/forms/vis/vendorAdmin.asp?intVendor_ID=<%=session.contents("intVendor_ID")%>">View Profile</a>
			&nbsp;
			<a href="<%=Application.Value("strWebRoot")%>UserAdmin/ChangePassword.asp">Change Password</a>
			&nbsp;<a href="<%=Application.Value("strWebRoot")%>forms/VIS/VendorSearchEngine.asp">Vendor Search Engine</a> &nbsp;
			<a href="./forms/misc/PersonalServiceContract.pdf" target="_blank">Personal Service Contract</a> </b>
		<br><br>
		</td>
	</tr>
<% if not ForceVerify then %>
	<tr>
		<td class="svplain11">
			<b>Vendor Activity Report</b><br>
			<span class="svplain8">
			This report provides a detailed account of each transaction you have 
			with our students.  After a requisition has been made for your service
			against a students' account a budget is created.  Below you will see
			the budget detail as well as monies that have been paid against the
			budget.  Each time a budget is paid against a 'line item' is created. 
			A line item details when and how much was paid against the budget.  When
			a line item is created a <b>'show'</b> link will be provided under the
			'Line Items' column. Click this link to see the line item detail and then 
			click <b>'hide'</b> to hide the line item information.</span><br><br>
			<table cellpadding="2">
				
			
	<%
	if rs.RecordCount > 0 then		
		dblSubTotal = 0
		strVendName = rs("szVendor_Name")
		intVendor_ID = rs("intVendor_ID")	
		szPO_Number = ""
		set rs2 = server.CreateObject("ADODB.RECORDSET")
		rs2.CursorLocation = 3
		lastStudentId = 0 
		oldClassID = 0 
		mDivCount = 0
		do while not rs.EOF		
			
			if rs("intStudent_ID") <> lastStudentId then
				if lastStudentId > 0 then
					call vbsShowTotals
				end if		
				call PrintHeader
				lastStudentId = rs("intStudent_ID")
				%>
					<tr id="div<% = mDivCount%>">
						<td class="TableSubHeader" align="center" title="If line items have been charged against your budget they can be viewed by clicking on 'show'.">
							Line Items
						</td>	
						<td class="TableSubHeader" align="center">
							Class Name
						</td>
						<!--<td class="TableSubHeader" align="center">
							Status
						</td>-->
						<td class="TableSubHeader">
							Description
						</td>
						<td class="TableSubHeader" align="center">
							QTY
						</td>
						<td class="TableSubHeader" align="center">
							Unit Cost
						</td>
						<td class="TableSubHeader" align="center" title="Shipping and Handling">
							S/H
						</td>
						<td class="TableSubHeader" align="center" title="(QTY * Unit Cost) + S/H">
							Budget Total
						</td>
						<td class="TableSubHeader" align="center" title="Sum of all line items (charged expeneses) entered by the office for a specific budget.">
							Actual Charges
						</td>
						<td class="TableSubHeader" align="center" title="Adjustments are needed to handle over expendatures and to release unused budgeted funds once the budget is closed.">
							Budget Adjust
						</td>
						<!--<td class="TableSubHeader" align="center" title="(Budget Total - Actual Charges) + Budget Adjust">
							Budget Balance
						</td>-->
						<td  class="ltGray"  style="width:0%;">
							&nbsp;
						</td>
						<td  class="ltGray">
							&nbsp;
						</td>
						<td  class="ltGray">
							&nbsp;
						</td>
					</tr>			
					<%
			end if
			
			
			myClassName = replace(rs("ClassLabel"),"'","\'")
			
			if (rs("intOrdered_Item_ID") & "" <> "") then
		
				' Set the budgeted cost for this item
				dblShipping = 0
				if rs("curShipping") & "" <> "" then
					if isNumeric(rs("curShipping")) then
						dblShipping = formatNumber(rs("curShipping"),2)
					end if
				end if
					
				dblBudgetCost = formatNumber(rs("Total"),2)
				'Get Line Item info
				liInfo = LineItemInfo(rs("intOrdered_Item_ID"),dblBudgetCost, rs("bolClosed"), oFunc.FPCScnn,strClass)
				bStatus = GetBudgetStatus(rs("intItem_Group_ID"),rs("bolApproved"),liInfo(4),rs("bolReimburse"))		
				
				dblCharge = formatNumber(liInfo(1),2)
				dblAdjBudget = formatNumber(dblBudgetCost + cdbl(liInfo(2)),2)		
				mDivCount = mDivCount + 1
				strBList = strBList & mDivCount & ","
				strSmallList = strSmallList & mDivCount & ","
				
				if bStatus = "rejc" then
					strClass = "TableCellStrike"
				else
					strClass = "TableCell"
					dblClassCharge = dblClassCharge + cdbl(dblCharge)
					dblClassBudget = dblClassBudget + cdbl(dblAdjBudget)
				end if
				
				if rs("oiComment") <> "" then
					strReason = "<BR><b>Comment:</b> " & rs("oiComment")
				else
					strReason = ""
				end if
				
				if rs("bolReimburse")  then
					strItemType = "Reimburse #" & rs("intOrdered_Item_ID") & ": "
				else
					strItemType = "Requisition #" & rs("intOrdered_Item_ID") & ": "
				end if
				
				' Print row with budget info		
		%>																		
						<tr id="div<% = mDivCount %>">
							<td class="<% = strClass %>" align="center" nowrap>
								<% if liInfo(3) <> "" then%>
								<a href="javascript:" onclick="jfToggle('<%=liInfo(3)%>','a<%=mLablelCount%>');" id="a<%=mLablelCount%>">show</a>
								<% 
									mLablelCount = mLablelCount + 1
								else
									response.Write "&nbsp;"
								end if 
								%>
							</td>
							<td class=<% = strClass %>>
								<% = myClassName %>
							</td>
							<!--<td class="<% = strClass %>" align="center">
								<% = bStatus %>
							</td>-->
							<td class=<% = strClass %>>
								<% response.Write strItemType & rs("oiDesc") & strReason %>
									&nbsp;
							</td>
							<td class=<% = strClass %> align="center" nowrap>
								<% = rs("intQTY") %>
							</td>
							<td class=<% = strClass %> align=right nowrap>
								$<% = formatNumber(rs("curUnit_Price"),2) %>
							</td>
							<td class=<% = strClass %> align=right nowrap title="Shipping and Handling">
								&nbsp;$<% = dblShipping %>
							</td>
							<td class=<% = strClass %> align=right nowrap title="(QTY * Unit Cost) + S/H">
								$<% = dblBudgetCost %>
							</td>
							<td class=<% = strClass %> align=right nowrap title="Sum of all line items (charged expeneses) entered by the office for a specific budget.">
								$<% = formatNumber(liInfo(1),2)%>
							</td>
							<td class=<% = strClass %> align=right nowrap title="Adjustments are needed to handle over expendatures and to release unused budgeted funds once the budget is closed.">
								$<% = formatNumber(liInfo(2),2) %>
							</td>
							<!--<td class=<% = strClass %> align=right nowrap title="(Budget Total - Actual Charges) + Budget Adjust">
								$<% = formatNumber((dblBudgetCost - cdbl(liInfo(1))) + cdbl(liInfo(2)),2)%>
							</td>-->
							<td bgcolor=white style="width:0%;">
								&nbsp;
							</td>
							<td class="<% = strClass %>" align="right" nowrap title="Budget Total - Budget Adjust">
								$<% = dblAdjBudget %>
							</td>
							<td class="<% = strClass %>" align="right" nowrap title="Actual Charges">
								$<% = dblCharge %>
							</td>
							<td class=<% = strClass %> align=right nowrap title="(Budget Total - Actual Charges) + Budget Adjust">
								$<% = formatNumber((dblBudgetCost - cdbl(liInfo(1))) + cdbl(liInfo(2)),2)%>
							</td>
						</tr>
						<% = liInfo(0) %>
						<%	
			end if						
			rs.MoveNext
		loop
		call vbsShowTotals
			%>

						<tr bgcolor="<% = strColor%>">
								<td class=svplain10 colspan="9" align=right>
									<b>All Student Totals:	</b>
								</td>
								<td bgcolor=white  style="width:0%;">
									&nbsp;&nbsp;&nbsp;
								</td>
								<td class="Gray" align=right>
									$<%=formatNumber(dblTotalBudget,2)%>
								</td>
								<td class="Gray" align=right>
									$<%=formatNumber(dblTotalCharge,2)%>
								</td>
								<td class="Gray" align=right>
									$<%=formatNumber(dblTotalBudget - dblTotalCharge,2)%>
								</td>
							</tr>				
			<%
		set rs2 = nothing		
	end if
	
	rs.Close
	set rs = nothing
	%>
			</table>
		</td>
	</tr>
<% else %>
	<tr>
		<td class="svError">
			<b>Before you can access your account activity and other features of the Vendor online system 
			you must first update your Vendor Profile. Click 
			<a href="<% = Application.Value("strWebRoot") %>/forms/vis/vendorAdmin.asp?intVendor_ID=<%=session.contents("intVendor_ID")%>">HERE</a>
			to review and update your profile now.</b>
		</td>
	</tr>
<% end if %>
<% end if %>
</table>
<%
response.Write oHtml.ToolTipDivs
set oHtml = nothing
call oFunc.CloseCN()
set oFunc = nothing
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")

function PrintHeader		
%>		
				<tr>
					<td colspan=9 style="width:100%;">
						<table style="width:100%;" cellpadding=3>
							<tr>
								<td class="tableHeader" style="width:25%;">
									<B>STUDENT NAME</B>
								</td>	
								<td class="tableHeader" style="width:35%;">
									<b>Family Name</b>
								</td>
								<td class="tableHeader" style="width:20%;">
									<b>Phone Number</B>
								</td>									
								<td class="tableHeader" style="width:20%;">
									<b>Email</B>
								</td>
							</tr>
							<tr>
								<td class="TableCell">
									<% = rs("szLAST_NAME") & ", " & rs("szFIRST_NAME")  %>
								</td>
								<td class="TableCell" align="center">
									<% = rs("szDesc") & " " & rs("szFamily_Name") %>
								</td>
								<td class="TableCell">
									<% = rs("szHome_Phone")  %>
								</td>						
								<td class="TableCell" align="center">
									<a href="mailto:<% = rs("szEMAIL") %>"><% = rs("szEMAIL") %></a>
								</td>
							</tr>	
						</table>
					</td>
				<% if mDivCount = 0 then %>
					<td >
						&nbsp;
					</td>
					<td class="TableSubHeader">
						<b>Budgeted</b>
					</td>
					<td class="TableSubHeader" align=center>
						<b>Paid</b>
					</td>
					<td class="TableSubHeader" align=center>
						<b>Balance</b>
					</td>
				<% else %>
					<td colspan=3>
						&nbsp;
					</td>
				<% end if %>
				</tr>
<%							
		
end function


function LineItemInfo(pOrderedID,pBudget,pClosed,pCn,pCellClass)
	' Checks for line item entries and returns the following array if they exist...
	' ar(0) = html table of all line items
	' ar(1) = Total amount Charged (sum of all line items)
	' ar(2) = Budget Adjustment (deifined if budget is closed or is negative)
	' ar(3) = Div List"  Table row id's used to hide or show line item html row
	' ar(4) = If true Line Items do exist else no line items exist
	dim sql
	dim tCharged
	dim tBudget
	dim sHtml
	dim rs
	dim dAdjust
	dim strDivList
	dim strClosed
	dim bolLineItem
	
	tCharged = 0
	dAdjust = 0
	bolLineItem = false
	
	if pClosed then 
		strClosed = "Budget is Closed"
	else
		strClosed = "Budget is Open"
	end if
	
	sql = "SELECT intLine_Item_ID, dtLine_Item, szLine_Item_desc, curUnit_Price, intQuantity, curShipping, " & _ 
			" (curUnit_Price * intQuantity) + curShipping as Total, dtCREATE, szCheck_Number " & _
			"FROM tblLine_Items " & _ 
			"WHERE (intOrdered_Item_ID = " & pOrderedID & ") " & _
			" Order by intLine_Item_ID "
				
	set rs = server.CreateObject("ADODB.RECORDSET")
	rs.CursorLocation = 3
	rs.Open sql, pCn
	
	do while not rs.EOF
		bolLineItem = true
		mDivCount = mDivCount + 1
		strDivList = strDivList & mDivCount & ","
		tCharged = tCharged + formatNumber(rs("Total"),2)	
		if rs("szCheck_Number") & "" <> "" then
			szCheck_Number = "Check #: " & rs("szCheck_Number")
			if rs("szLine_Item_desc") & "" <> "" then
				szCheck_Number = "<BR>" & szCheck_Number
			end if
		else
			szCheck_Number = ""
		end if
			
		sHtml = sHtml & "<tr id='div" & mDivCount & "' style='display:none;'>" & _
				"<td>&nbsp;</td><td   class='TableCellContrast'>Entered: " & formatDateTime(rs("dtCREATE"),2) & "</td>" & _
				"<td class='TableCellContrast' >" & rs("szLine_Item_desc") & szCheck_Number & "</td>" & _
				"<td class='TableCellContrast' align='center' valign='middle'>" & rs("intQuantity") & "</td>" & _
				"<td class='TableCellContrast' align='right' valign='middle'>$" & formatNumber(rs("curUnit_Price"),2) & "&nbsp;</td>" & _
				"<td class='TableCellContrast' align='right' valign='middle'>$" & formatNumber(rs("curShipping"),2) & "</td><td class='TableCellRed' >&nbsp;</td>" & _
				"<td class='TableCellContrast' align='right' valign='middle'>$" & formatNumber(rs("Total"),2) & "</td>" & _
				"<td colspan='2'>&nbsp;</td><td colspan='3' class='TableCellContrast' align='center'>" & strClosed & "</td></tr>" 
		rs.MoveNext		
	loop	
	rs.Close
	set rs = nothing
	
	tBudget = pBudget - tCharged
	if tBudget < 0 or pClosed then
		dAdjust = tBudget * -1
	end if
	
	dim ar(4)
	ar(0) = sHtml
	ar(1) = formatNumber(tCharged,2)
	ar(2) = formatNumber(dAdjust,2)
	ar(3) = strDivList
	ar(4) = bolLineItem
	LineItemInfo = ar
end function

function GetBudgetStatus(pItemGroup,pBappr,pBolLineItems,pIsReimburse)	

	if pBappr & "" = "" and pBolLineItems = false then
		GetBudgetStatus = "pend"
	elseif pBappr = false then
		GetBudgetStatus = "rejc"
	elseif pBolLineItems = true and pIsReimburse = false and pItemGroup = 2 then
		GetBudgetStatus = "<font color='green'><b>pick up</b></font>"
	elseif pBolLineItems = true and pIsReimburse = false and pItemGroup = 1 then
		GetBudgetStatus = "pymt made"
	elseif pBappr = true and pBolLineItems = false and pIsReimburse = false and pItemGroup = 2 then
		GetBudgetStatus = "ordered"
	elseif pBappr = true and pBolLineItems = false and pIsReimburse = false and pItemGroup = 1 then
		GetBudgetStatus = "vend appr"
	elseif pBappr = true and pBolLineItems = true and pIsReimburse = true then
		GetBudgetStatus = "check cut"
	else
		GetBudgetStatus = "pend"
	end if
end function

function PrintBudgetHeader
%>
				<tr id="div<% = mDivCount%>">
					<td class="TableSubHeader" align="center" title="If line items have been charged against your budget they can be viewed by clicking on 'show'.">
						Line Items
					</td>	
					<td class="TableSubHeader" align="center">
						Budget Item
					</td>
					<td class="TableSubHeader" align="center">
						Status
					</td>
					<td class="TableSubHeader">
						Description
					</td>
					<td class="TableSubHeader" align="center">
						QTY
					</td>
					<td class="TableSubHeader" align="center">
						Unit Cost
					</td>
					<td class="TableSubHeader" align="center" title="Shipping and Handling">
						S/H
					</td>
					<td class="TableSubHeader" align="center" title="(QTY * Unit Cost) + S/H">
						Budget Total
					</td>
					<td class="TableSubHeader" align="center" title="Sum of all line items (charged expeneses) entered by the office for a specific budget.">
						Actual Charges
					</td>
					<td class="TableSubHeader" align="center" title="Adjustments are needed to handle over expendatures and to release unused budgeted funds once the budget is closed.">
						Budget Adjust
					</td>
					<td class="TableSubHeader" align="center" title="(Budget Total - Actual Charges) + Budget Adjust">
						Budget Balance
					</td>
					<td  class="ltGray"  style="width:0%;">
						&nbsp;
					</td>
					<td  class="ltGray">
						&nbsp;
					</td>
					<td  class="ltGray">
						&nbsp;
					</td>
				</tr>					
<%
end function

sub vbsShowTotals()
	
	if ilp_ID & "" = "" then		
		'dblClassCharge = "0.00" 
	end if
	
	if ilpShortID & "" = "" then
		'dblClassBudget = "0.00"
	end if 
	' ADD THESE LINES TO MAKE COURSE TOTALS HIDDEN
	'mDivCount = mDivCount + 1
	'strBList = strBList & mDivCount & ","
	'id="div<%=mDivCount" (NEEDS TO BE ADDED TO <TR> TAG IN HTML BELOW)
	'mDivCount = mDivCount + 1
	strDivList = strDivList & mDivCount & ","
	strSmallList = strSmallList & mDivCount & ","
%>
				<tr class=svplain10 bgcolor="<% = strColor%>" >
					<td colspan="9" align=right>		
						<table style="width:100%;" ID="Table1">
							<tr>
								<td align="right" class="svplain10">
									<b>Student Totals:</b>
								</td>
							</tr>
						</table>									
					</td>
					<td bgcolor=white  style="width:0%;">
						&nbsp;&nbsp;&nbsp;
					</td>
					<td class=gray align=right>
						<nobr>
						<% if instr(1,dblBudgetCost,"-") > 0 then
								response.Write "- $" & formatNumber(replace(dblClassBudget,"-",""),2)
						   else
								response.Write "+ $" & formatNumber(dblClassBudget,2)
						   end if						
						%></nobr>
					</td>
					<td class=gray align=right>
						<nobr>
						<% if instr(1,dblActualCost,"-") > 0 then
								response.Write "+ $" & formatNumber(replace(dblClassCharge,"-",""),2)
						   else
								response.Write "- $" & formatNumber(dblClassCharge,2)
						   end if						
						%></nobr>
					</td>
					<td class=gray align=right>
						$<% = formatNumber(dblClassBudget - dblClassCharge ,2)%>
					</td>
				</tr>	
<%
		
%>				
				<tr bgcolor=white>
					<td colspan="20">
						&nbsp;
					</td>
				</tr>	
<%
	'response.Write dblTotalCharge & " - " &  dblClassCharge
	dblTotalCharge = cdbl(dblTotalCharge) + cdbl(dblClassCharge)
	dblTotalBudget = cdbl(dblTotalBudget) + cdbl(dblClassBudget)
	dblClassBudget = 0 
	dblClassCharge = 0 
end sub
%>