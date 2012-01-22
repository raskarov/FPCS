<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		reqGoods.asp
'Purpose:	Dynamically creates goods form based on Item selection.
'			This form is used to gather and display requisition and
'			reimbursement data.
'Date:		23 AUG 2002
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Dimension Variables, make db Connection, print HTML header.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim intStudent_ID 
dim intClass_ID
dim intILP_ID
dim oFunc			'wsc object
dim strQTYText 
dim strPriceText
dim strAddorView	'text that is printed on web form to tell us if we are adding or viewing
dim strBudgetDesc
dim intQty
dim curUnit_Price
dim strStudentName
dim bolTeacherLock

bolTeacherLock = false
intQty = 0
curUnit_Price = 0

Session.Value("strTitle") = "Add a Good or Service."
Session.Value("strLastUpdate") = "19 Aug 2002"
Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

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

if bolResetVend <> "" then
	' Sets vendor to empty string. Needed in case if a user selects a vendor on the form
	' and then changes the selection in the 'Item' drop down. Since Item is higher
	' in the work flow all items beneath must be reset.
	intVendor_ID = ""
end if 

if intILP_ID = "" and intClass_ID = "" then
%>
<font class="svplain10"><B>The request to view this page is invalid. </B></font>
<br>
<input type="button" value="Close Window" onClick="window.opener.focus();window.close();"
	class="btSmallGray"> </body> </html>
<%
	set oFunc = nothing
	Response.End
end if 

strAddorView = "ADD "
if ExistingItemID <> "" and intStudent_Id <> "" then
	' We have a request to populate an existing item from Ordered tables
	sql = "SELECT tblOrdered_Items.intVendor_ID, " & _
			"    tblOrdered_Items.intItem_ID, tblOrdered_Items.intStudent_ID, " & _
			"    tblOrdered_Items.intQty, tblOrdered_Items.curUnit_Price, " & _
			"    tblOrdered_Items.bolReimburse, " & _
			"    tblOrdered_Items.intSchool_Year, " & _
			"    tblOrdered_Items.bolApproved, tblOrdered_Items.curShipping, " & _
			"    tblOrd_Attrib.szValue, tblOrd_Attrib.intItem_Attrib_ID, ci.bolRequired " & _
			"FROM tblOrd_Attrib INNER JOIN" & _
			"    tblOrdered_Items ON " & _
			"    tblOrd_Attrib.intOrdered_Item_ID = tblOrdered_Items.intOrdered_Item_ID LEFT OUTER JOIN " & _
			"	 tblClass_Items ci ON tblOrdered_Items.intClass_Item_ID = ci.intClass_Item_Id " & _
			"WHERE (tblOrdered_Items.intILP_ID = " & intILP_ID & ") AND " & _
			"    (tblOrdered_Items.intOrdered_Item_ID = " & ExistingItemID & ") " & _
			"ORDER BY tblOrd_Attrib.intOrder"
			
	set rsRead = server.CreateObject("ADODB.RECORDSET")
	rsRead.CursorLocation = 3
	rsRead.Open sql, oFunc.FPCScnn
	
	if rsRead.RecordCount > 0 then
		intVendor_ID = rsRead("intVendor_ID")
		intItem_ID = rsRead("intItem_ID")
		intQty = rsRead("intQty")
		curUnit_Price = rsRead("curUnit_Price")
		curShipping = rsRead("curShipping")
		bolReimburse = oFunc.TrueFalse(rsRead("bolReimburse"))
		intSchool_Year = rsRead("intSchool_Year")
		bolApproved = oFunc.TrueFalse(rsRead("bolApproved"))
		intOrd_Item_ID = ExistingItemID
		bolRequired = oFunc.TrueFalse(rsRead("bolRequired"))
		do while not rsRead.EOF
			execute("attrib" & rsRead("intItem_Attrib_ID") & " = rsRead(""szValue"")" )
			rsRead.MoveNext
		loop
	end if 
	
	rsRead.Close
	set rsRead = nothing
	strAddorView = "View/Edit "
elseif ExistingItemID <> "" and intClass_ID <> "" then
	sql ="SELECT tblClass_Items.intVendor_ID, tblClass_Items.intItem_ID, " & _
			"    tblClass_Items.intQty, tblClass_Items.curUnit_Price,  " & _
			"    tblClass_Items.intSchool_Year, tblClass_Items.curShipping, " & _
			"    tblClass_Attrib.intClass_Item_ID, tblClass_Attrib.szValue,  " & _
			"    tblClass_Attrib.intItem_Attrib_ID,tblClass_Items.bolRequired, c.intContract_Status_ID  " & _
			"FROM tblClass_Attrib INNER JOIN " & _
			"    tblClass_Items ON  " & _
			"    tblClass_Attrib.intClass_Item_ID = tblClass_Items.intClass_Item_ID INNER JOIN " & _
			"	 tblClasses c ON c.intClass_ID = tblClass_Items.intClass_ID " & _
			"WHERE (tblClass_Items.intClass_ID = " & intClass_ID & ") AND  " & _
			"    (tblClass_Items.intClass_Item_ID = " & ExistingItemID & ") " & _
			"ORDER BY tblClass_Attrib.intOrder"
	set rsRead = server.CreateObject("ADODB.RECORDSET")
	rsRead.CursorLocation = 3
	rsRead.Open sql, oFunc.FPCScnn
		
	if rsRead.RecordCount > 0 then
		intVendor_ID = rsRead("intVendor_ID")
		intItem_ID = rsRead("intItem_ID")
		intQty = rsRead("intQty")
		curUnit_Price = rsRead("curUnit_Price")
		curShipping = rsRead("curShipping")
		if ExistingItemID <> "" and intStudent_ID <> "" then
			bolReimburse = oFunc.TrueFalse(rsRead("bolReimburse"))
			bolApproved = oFunc.TrueFalse(rsRead("bolApproved"))
		else
			bolReimburse = 0
		end if 
		bolRequired = oFunc.TrueFalse(rsRead("bolRequired"))
		intSchool_Year = rsRead("intSchool_Year")
		
		if rsRead("intContract_Status_ID") & "" = "5" then
			bolTeacherLock = true
		end if
		intClass_Item_ID = ExistingItemID
		
		do while not rsRead.EOF
			execute("attrib" & rsRead("intItem_Attrib_ID") & " = rsRead(""szValue"")" )
			rsRead.MoveNext
		loop
	end if 
	
	rsRead.Close
	set rsRead = nothing
	strAddorView = "View/Edit "
end if
if intItem_Group_ID = "1" then
	strItemTitle = "Service"
elseif intItem_Group_ID = "2" then
	strItemTitle = "Good"
end if 

if intStudent_ID <> "" then
	strStudentName = oFunc.StudentInfo(intStudent_ID,"3")
	strFamilyPhone = oFunc.FamilyInfo("1",intStudent_ID,"5")
end if
%>
<form action="<%=Application("strSSLWebRoot")%>forms/requisitions/reqGoods.asp" name="main" method="post">
	<input type=hidden name=intOrd_Item_ID value="<%=intOrd_Item_ID%>" ID="Hidden1">
	<input type=hidden name=intClass_Item_ID value="<%=intClass_Item_ID%>" ID="Hidden2">
	<input type="hidden" name="viewing" value="true" ID="Hidden3"> <input type=hidden name=intIlp_ID value="<% = intIlp_ID %>" ID="Hidden4">
	<input type=hidden name=intStudent_ID value="<% =intStudent_ID %>" ID="Hidden5">
	<input type=hidden name=intClass_ID value="<% = intClass_ID %>" ID="Hidden6"> <input type=hidden name=strClassName value="<% = strClassName %>" ID="Hidden7">
	<input type=hidden name=bolComplies value="<% =bolComplies %>" ID="Hidden8"> <input type=hidden name=intItem_Group_ID value="<%=intItem_Group_ID%>" ID="Hidden9">
	<input type=hidden name=intPOS_Subject_ID value="<%=intPOS_Subject_ID%>" ID="Hidden10">
	<input type=hidden name=intBudget_ID value="<%=intBudget_ID%>" ID="Hidden13"> <input type="hidden" name="bolResetVend" value="" ID="Hidden11">
	
	<table width="100%" ID="Table1" cellpadding="2">
		<tr>
			<Td class="yellowHeader" nowrap>
				<table style='width:100%;' cellpadding="0">
					<tr>
						<td class="yellowHeader">
							<b>
								<% = strAddorView %>
								A
								<% = ucase(strItemTitle) %>
							</b>
						</td>
						<td class="yellowHeader" align="right">
							<% if strFamilyPhone <> "" and ucase(session.Contents("strRole")) = "ADMIN" then%>
							Family Phone#:
							<%=oFunc.FormatPhone(strFamilyPhone)%>
							<% end if %>
						</td>
					</tr>
				</table>
			</Td>
		</tr>
		<tr>
			<td class="SubHeader">
				<b>Class:</b>
				<% = strClassName%>
				<% if strStudentName <> "" then%>
				&nbsp;&nbsp;<b>Student:</b>
				<%=strStudentName%>
				<% end if %>
				<% if intOrd_Item_ID <> "" then%>
				&nbsp;&nbsp;<b>Item #:</b>
				<%=intOrd_Item_ID%>
				<% end if %>
			</td>
		</tr>
		<tr>
			<td bgcolor="f7f7f7">
				<table ID="Table2">
					<TR>
					<tr>
						<Td class="HeadLine">
							<B>Will this be a Requisition or a Reimbursement?</B>
						</Td>
					</tr>
					<tr>
						<td class="svplain8">
							<select name="bolReimburse" onChange="this.form.submit();" ID="Select5">
								<option value="">Select
									<%
							dim strBolValues
							dim strBolText
							strBolValues = "0,1"
							strBolText = "Requisition,Reimbursement"									 
							Response.Write oFunc.MakeList(strBolValues,strBolText,bolReimburse)												 
						%>
							</select>
							<br>
							<br>
							<% if intItem_Group_ID = 1 and bolReimburse = "" then%>
							<b><u>Requisition:</u></b> Vendor submits invoices directly to FPCS. Business Office
							will remit payment directly to the vendor after services are rendered to your child.
							<br>
							<br>
							<b><u>Reimbursement:</u></b> You pay the vendor for services rendered your 
							child and you will be reimbursed by FPCS based on receipts you submit to FPCS 
							(Reciepts cannot be for more than $200).<br>
							<br>
							<% elseif intItem_Group_ID = 2 and bolReimburse = "" then%>
							<b><u>Requisition:</u></b> FPCS will order the item for you. No cash 
							transaction takes place on your part.<br>
							<br>
							<b><u>Reimbursement:</u></b> You will pay for the item and then request 
							reimbursement from FPCS at a later date (Reimbursement deadlines need to be followed).<br>
							<br>
							<% end if%>
						</td>
					</tr>
					<%				
' This is our 'Progression Logic'. As we get more data from the user we display more to the form.
if bolReimburse <> "" then
	if bolReimburse = "1" and bolComplies = "" and ExistingItemID = "" and intItem_ID = "" then
%>
					<tr>
						<td class="HeadLine">
							Please click <a href="Reimbursement_Instructions.htm" target="_blank">HERE</a> to 
							review the Reimbursement Policy to be sure your request complies with current 
							guidelines. If you believe this request complies, you need to print out the <u>Request 
								for Reimbursement Form</u> and bring it into the FPCS office with the 
							original receipts for the item(s) referenced on the reimbursement form.
							<br>
							<br>
							<b>Yes my request does comply with current guidelines. <a href="javascript:document.main.bolComplies.value='true';document.main.submit();">
									Continue</a>
								<br>
								No my request does not comply with current guidelines. <a href="javascript:window.opener.focus();window.close();">
									Cancel Request</a> </b>
							<br>
							<br>
						</td>
					</tr>
					<%
	else				
		call vbfItemsList 	
	end if
end if 

if intItem_ID <> "" then
	if intItem_ID = 1 and intVendor_ID & "" = "" then
		intVendor_ID = 157
	end if
	call vbfVendorList
end if 
	
if intVendor_ID <> "" then
	call vbfAttribForm
end if 	

function vbfVendorList
	' Prints list of Guardians based on intStudent_ID	
%>
					<tr>
						<Td class="HeadLine">
							<b>
								<% if bolReimbursement = "1" then%>
								Which vendor did you use?
								<% else %>
								Which vendor would you like to use? (Note: if you can't find the vendor, please contact the office)
								<% end if%>
							</b>
						</Td>
					</tr>
					<tr>
						<td class="svplain8">
							<% if intVendor_ID & "" = "157" and intItem_ID & "" = "1" then %>
							<input type="hidden" name="intVendor_ID" value="<% = intVendor_ID %>"> A S D 
							Administration BLDG.<br>
							<br>
							<% else %>
							<select name="intVendor_ID" onChange="this.form.submit();" ID="Select1">
								<option value="">
									<%
							dim sqlAdd
							dim sqlVendor
							dim strVendCriteria
							
							'if bolReimburse = "0" then
								' Some Vendors can only provide goods/services for reimbursements.
								' If the user is adding a requisition filter out the following
							'	sqlAdd = " and (va.bolReimburse_Only <> 1) "
							'end if 
							
							'if intVendor_ID <> "" then
							'	strVendCriteria = " OR va.intVendor_ID = " & intVendor_ID & " " 
							'end if 

							'sql2 = "SELECT     TOP 1 upper(szVendor_Status_CD) " & _ 
							'			"	FROM          tblVendor_Status vs " & _ 
							'			"	WHERE      vs.intVendor_ID = v.intVendor_ID AND vs.intSchool_Year <= " & session.Contents("intSchool_Year") & _ 
							'			"	ORDER BY intSchool_Year DESC, intVendor_Status_ID DESC" 
	
							select case intItem_Group_ID
								case 1
									'sql =" v.bolService_Vendor = 1 " 
									sql =" and vt.bolService_Vendor = 1 " 
									if bolReimburse = "1" then
										' for reimbursements we require the for profit vendors to have a contract date on file
										' but not the non-profit
										'sql2 = " SELECT     TOP 1 upper(szVendor_Status_CD) " & _ 
										'	"	FROM          tblVendor_Status vs " & _ 
										'	"	WHERE      (vs.intVendor_ID = v.intVendor_ID AND vs.intSchool_Year = " & session.Contents("intSchool_Year") & _ 
										'	"       and vs.dtContract_Start is not null) or (vs.intVendor_ID = v.intVendor_ID AND vs.intSchool_Year <= " & session.Contents("intSchool_Year") & " and  v.bolNonProfit = 1) " & _
										'	"	ORDER BY intSchool_Year DESC, intVendor_Status_ID DESC "
										'jd:reimbursement for non-profit service vendors only
										
										sql = sql & " and vt.bolNonProfit = 1 "
									    
									else
										' for requistions we require both for and non profit vendors to have a
										' contract date on file
										'sql2 = " SELECT     TOP 1 upper(szVendor_Status_CD) " & _ 
										'	"	FROM          tblVendor_Status vs " & _ 
										'	"	WHERE      (vs.intVendor_ID = v.intVendor_ID AND vs.intSchool_Year = " & session.Contents("intSchool_Year") & _ 
										'	"       and vs.dtContract_Start is not null) " & _
										'	"	ORDER BY intSchool_Year DESC, intVendor_Status_ID DESC "
										sql = sql & " and vt.dtContract_start is not NULL "

									end if
								case 2
									'sql = " v.bolGoods_Vendor = 1 " 
									sql = " and vt.bolGoods_Vendor = 1 " 
								case 3	
									'sql = "  v.bolGoods_Vendor = 1 " 
									'sql = "  v.bolService_Vendor = 1 " 
									sql = " and vt.bolGoods_Vendor = 1 " 
									sql = " and vt.bolService_Vendor = 1 " 
							end select		
							
							'sqlVendor = "SELECT intVendor_ID,  " & _ 
							'			" szVendor_Name AS Vend_Name " & _ 
							'			"FROM tblVendors v WHERE " & sql & _ 
							'			" AND (" & sql2 & ") IN ('APPR','PEND') " & _
							'			" ORDER BY Vend_Name "
							sqlVendor = "SELECT intVendor_ID " & _
                                        ", szVendor_Name as Vend_Name " & _
                                        "FROM (SELECT intVendor_ID " & _
	                                    "    , szVendor_Name " & _
	                                    "    , v.bolNonProfit" & _
	                                    "    , (SELECT TOP 1 szVendor_Status_CD " & _ 
		                                "        FROM tblVendor_Status vs  " & _
		                                "        WHERE (intSchool_Year <= " & session.Contents("intSchool_Year") & ") "& _
		                                "        AND (vs.intVendor_ID = v.intVendor_ID) " & _
		                                "        ORDER BY intSchool_Year DESC " & _
		                                "        , intVendor_Status_ID DESC) AS szVendor_Status_CD " & _
	                                    "    , v.bolGoods_Vendor " & _
	                                    "    , v.bolService_Vendor " & _
	                                    "    , (SELECT TOP 1 dtContract_Start " & _ 
		                                "        FROM tblVendor_Status vs " & _
		                                "        WHERE (intSchool_Year <= " & session.Contents("intSchool_Year") & ") " & _
		                                "        AND (vs.intVendor_ID = v.intVendor_ID) " & _
		                                "        ORDER BY intSchool_Year DESC " & _
		                                "        , intVendor_Status_ID DESC) as dtContract_Start " & _
	                                    "    , (select top 1 intSchool_year " & _
		                                "        from tblVendor_status vs " & _
		                                "        WHERE (intSchool_Year <= " & session.Contents("intSchool_Year") &  ") " & _
		                                "        AND (vs.intVendor_ID = v.intVendor_ID) " & _
		                                "        ORDER BY intSchool_Year DESC) as intSchool_year " & _
	                                    "    FROM tblVendors v) vt " & _
                                        "WHERE 1 = 1 " & _
                                        "and vt.szVendor_Status_CD = 'APPR' " & _
                                        sql & _
                                        "ORDER BY vt.szVendor_Name  " 
										
										
 
							Response.Write oFunc.MakeListSQL(sqlVendor,"intVendor_ID","Vend_Name",intVendor_ID)												 
						%>
							</select>
							<% if intVendor_ID = "" then %>
							or <input type="button" value="Search for a Vendor" class="NavSave" onclick="jfVendorSearch();">
							<script language="javascript">
							function jfVendorSearch(){
								// we use bolReim to let vendorAdmin.asp know that we are coming from 
								// the goods page.
								var winVendS;
								var sLink = "<%=Application.Value("strWebRoot")%>forms/VIS/VendorSearchEngine.asp?";
								sLink += "intPOS_SUBJECT_ID=<%=intPOS_SUBJECT_ID%>&intClass_Item_ID=<%=intClass_Item_ID%>";
								sLink += "&viewing=true&intItem_ID=<% = intItem_ID %>&intClass_ID=<%=intClass_ID%>";
								sLink += "&bolWin=true&intItem_Group_ID=<%=intItem_Group_ID%>&intILP_ID=<% = intILP_ID%>";
								sLink += "&intStudent_ID=<% = intStudent_ID%>&bolReimburse=<% = bolReimburse%>";
								sLink += "&strClassName=<% = strClassName%>&bolComplies=<%=bolComplies%>&intOrd_Item_ID=<%=intOrd_Item_ID%>";
								winVendS = window.open(sLink,"winVendS","width=750,height=500,scrollbars=yes,resizable=yes");
								winVendS.moveTo(0,0);
								winVendS.focus();
							}
							</script>
							<% end if 
			'if ucase(session.contents("strUserId")) = "CHRONIH166" then  response.write sqlVendor 
			'response.write sqlVendor 
%>
							<br>
							<br>
						</td>
					</tr>
					<% if intItem_Group_ID = 1 then %>
					<tr>
						<td class="svplain8" colspan="10">
						    
						    
						    <%'jd show a different message if non-profit service reimbursement
						        if bolReimburse = "1" then
						     %>
						        <b>
						            If the Non-Profit Service Vendor you are looking for is not on the pulldown list, please contact the FPCS office at 907-742-3700.
						        </b>
						     <%else %>
							    <b>For all 'For-Profit' Service Vendors to be included in the above vendor pull-down 
								    list, the vendor must have completed the following each school year (July 1 - 
								    June 30):<br>
								    1. Updated their vendor profile.<br>
								    2. Completed a contract with FPCS (and ASD).<br>
								    3. Have been issued a contract starting date. </b>
							    <br>
							    <br>
						    <%end if %>

						</td>
					</tr>
					<% end if 
				if intVendor_ID = "" and intItem_Group_ID = 2 then %>
					<tr>
						<Td class="svplain">
							GOODS VENDORS are added by families and teachers as needed.
							<br>
							Goods will be ordered and families WILL ONLY be reimbursed if they follow
							<BR>
							the ILP guidelines and acceptable district purchasing requirements.
							<BR>
							Goods vendors are not “automatically” approved by being on a goods vendor list.
							<br>
							<br>
							If your vendor is not in the vendor list please contact the office. Click <input type=button value="HERE" onClick="jfAddVendor('<% = bolReimburse%>','<% = intILP_ID%>','<% = intStudent_ID%>');" class="btSmallGray" NAME="Button3">
							to do so.
							<br>
							<br>
						</Td>
					</tr>
					<script language="javascript">
					function jfAddVendor(bolReim,ilp,student){
						// we use bolReim to let vendorAdmin.asp know that we are coming from 
						// the goods page.
						var winVend;
						var sLink;
						sLink = "<%=Application.Value("strWebRoot")%>forms/VIS/vendorAdmin.asp?";
						sLink += "intPOS_SUBJECT_ID=<%=intPOS_SUBJECT_ID%>&intClass_Item_ID=<%=intClass_Item_ID%>";
						sLink += "&viewing=true&intItem_ID=<% = intItem_ID %>&intClass_ID=<%=intClass_ID%>";
						sLink += "&bolWin=true&intItem_Group_ID=<%=intItem_Group_ID%>&intILP_ID=<% = intILP_ID%>";
						sLink += "&intStudent_ID=<% = intStudent_ID%>&bolReimburse=<% = bolReimburse%>";
						sLink += "&strClassName=<% = strClassName%>&bolComplies=<%=bolComplies%>&intOrd_Item_ID=<%=intOrd_Item_ID%>";	
						winVend = window.open(sLink,"winVend","width=805,height=500,scrollbars=yes,resizable=yes");
						winVend.moveTo(0,0);
						winVend.focus();
					}
					</script>
					<%
				 end if
			end if ' end if the vendor_id = 157
end function

function vbfItemsList
	'Print select list of Items
%>
					<tr>
						<Td class="HeadLine"><b>
								<% if bolReimburse <> "1" then %>
								What would you like to Requisition?
								<%else%>
								What would you like to be Reimbursed on?
								<%end if%>
							</b>
						</Td>
					</tr>
					<tr>
						<td>
							<select name="intItem_ID" onChange="this.form.bolResetVend.value='true';this.form.submit();"
								ID="Select2">
								<option value="">
									<%
							dim sqlItem
							if not oFunc.IsAdmin then 
								siWhere = " and bolAdmin_Only = 0 "
							end if 
							
							'JD 052711 do not include reimburesements only to requisition
							if bolReimburse = 0 then
		                        ' This will make sure we DO NOT include any attributes that are for
		                        ' reimbursements only since we are not dealing with a reimbursement
		                        siWhere = " AND (bolReimbursement_Only IS NULL OR bolReimbursement_Only = 0)"
		                    else
		                        siWhere = " AND (bolRequisition_Only IS NULL OR bolRequisition_Only = 0)"
	                        end if 
	                        
							sqlItem = "Select intItem_ID, szName " & _
											 "from trefItems " & _
											 "where intItem_Group_ID = " & intItem_Group_ID & _
											 " and (SchoolYearActive <= " & session.Contents("intSchool_Year")  & _
											 " and (SchoolYearInactive > " & session.Contents("intSchool_Year") & " or SchoolYearInactive is null)) " & _
											 siWhere & _
											 " order by szName"										 
							Response.Write oFunc.MakeListSQL(sqlItem,"intItem_ID","szName",intItem_ID)												 
						%>
							</select>
							<br>
							<br>
						</td>
					</tr>
					<%
end function 

function vbfAttribForm
	' This function dynaimicaly creates the Items form based on attributes in tascItems_Attrib
	dim sqlAttrib
	dim strSQL
	dim strJF
	dim strAttrib		' This string contains a comman seperated list of Attrib id's that will
						' be used in reGoodsInsert.asp to insert and update fields.
	dim intTotal		'Total of Price * Quantity
	
	'JD 052711 filter for either reim or req
	'dim strReimburseFilter
	dim strReimburseOrRequistionFilter
	dim strNoAdmin		
	
	intTotal = 0 
	
	'JD 052711 filter for either reim or req
	'if bolReimburse = 0 then
	'	' This will make sure we DO NOT include any attributes that are for
	'	' reimbursements only since we are not dealing with a reimbursement
	'	strReimburseFilter = " AND (trefItem_Attrib.bolReimbursement_Only IS NULL OR trefItem_Attrib.bolReimbursement_Only = 0)"
	'end if 
	
	if bolReimburse = 0 then
		' This will make sure we DO NOT include any attributes that are for
		' reimbursements only since we are not dealing with a reimbursement
		strReimburseOrRequistionFilter = " AND (trefItem_Attrib.bolReimbursement_Only IS NULL OR trefItem_Attrib.bolReimbursement_Only = 0)"
		else
		strReimburseOrRequistionFilter = " AND (trefItem_Attrib.bolRequisition_Only IS NULL Or trefItem_Attrib.bolRequisition_Only = 0)"
	end if 
	
	'get active only
	strReimburseOrRequistionFilter = strReimburseOrRequistionFilter + " AND (trefItem_Attrib.isActive  = 1 or trefItem_Attrib.isActive is null)"

	
	if session.Contents("strRole") <> "ADMIN" then
		' This will filter out all ADMIN only attributes 
		strNoAdmin = " AND (trefItem_Attrib.bolAdmin = 0 or trefItem_Attrib.bolAdmin is NULL) "
	end if 
	
	if intBudget_ID <> "" then
		' This section auto populates specific fields based on info we have
		' in tblBudget
		set rsBudget = server.CreateObject("ADODB.RECORDSET")
		rsBudget.CursorLocation = 3
		sql = "SELECT szDesc, intQTY, curUnit_Price, curShipping " & _
				"FROM tblBudget b " & _
				"WHERE (intBudget_id = " & intBudget_ID & ")"
		rsBudget.Open sql,oFunc.FPCScnn
		
		if rsBudget.RecordCount > 0 then
			intQTY = rsBudget("intQTY")
			curUnit_Price = rsBudget("curUnit_Price")
			curShipping = rsBudget("curShipping")
			strBudgetDesc = rsBudget("szDesc")
		end if
		rsBudget.Close
		set rsBudget = nothing
	end if
	
	' This sql will return all the attributes for a given item id
	sqlAttrib = "SELECT tascItems_Attrib.intItem_Attrib_ID, " & _
				"    tascItems_Attrib.intItem_ID, trefItem_Attrib.szName, " & _
				"    trefItem_Attrib.intForm_Element_ID, " & _
				"    trefItem_Attrib.szForm_Attrib, trefItem_Attrib.bolAdmin, " & _
				"    trefItems.szName AS Item_Name, trefValidation.szJS_Function, " & _
				"    tascItems_Attrib.intOrder_ID, " & _
				"	 trefItem_Attrib.szOption_SQL,trefItems.szALT_QTY_Text, " & _
				"	 trefItems.szALT_Price_Text, trefItem_Attrib.szReplace_Text " & _
				"FROM tascItems_Attrib INNER JOIN " & _
				"    trefItem_Attrib ON " & _
				"    tascItems_Attrib.intItem_Attrib_ID = trefItem_Attrib.intItem_Attrib_ID " & _
				"     INNER JOIN " & _
				"    trefItems ON " & _
				"    tascItems_Attrib.intItem_ID = trefItems.intItem_ID LEFT OUTER JOIN " & _
				"    trefValidation ON " & _
				"    tascItems_Attrib.intValidation_ID = trefValidation.intValidation_ID " & _
				"WHERE (tascItems_Attrib.intItem_ID = " & intItem_ID & ") " & _
				strReimburseOrRequistionFilter & strNoAdmin & _
				" and (tascItems_Attrib.bolVersion_2_Off is null or tascItems_Attrib.bolVersion_2_Off = 0)  " & _
				"ORDER BY tascItems_Attrib.intOrder_ID"
				
			    'JD 052711 filter for either reim or req
				'strReimburseFilter & strNoAdmin & _


	'response.Write sqlAttrib		
	set rsAttrib = server.CreateObject("ADODB.recordset")
	rsAttrib.CursorLocation = 3
	rsAttrib.Open sqlAttrib, oFunc.FPCScnn

	if rsAttrib("szALT_QTY_Text") & "" <> "" then
		strQTYText = rsAttrib("szALT_QTY_Text")
	else
		strQTYText = "Number of Units"
	end if
	
	if rsAttrib("szALT_Price_Text") & "" <> "" then
		strPriceText = rsAttrib("szALT_Price_Text")
	else
		strPriceText = "Unit Price"
	end if
	%>
					<tr>
						<td colspan="2" class="Headline">
							Please fill in the <b>
								<% if bolReimbursement = "1" then%>
								Reimbursement
								<% else %>
								Requisition
								<% end if%>
								form below.</b>
						</td>
					</tr>
					<tr>
						<td>
							<table ID="Table3">
								<%
	
	do while not rsAttrib.EOF
		
		' Get values for Attributes 
		execute("strAttrValue = attrib" &  rsAttrib("intItem_Attrib_ID"))
		
		' I don't like having this here, but the client requested it. This auto populates
		' the Item Attribute for student grade if it is not already defined.		
		if rsAttrib("intItem_Attrib_ID") = "10" and strAttrValue = "" and intStudent_Id <> "" then
			strAttrValue = oFunc.StudentInfo(intStudent_Id,"5")
		end if
		
		' logic to create the form elements
		if 	rsAttrib("intForm_Element_ID") = "1" then
			' Text Box					
			call vbfBeginRow(rsAttrib("szName"))
			
			' We'll substitute the budget description for the first item if applicable
			if strBudgetDesc <> "" and rsAttrib("intOrder_ID") = 1 then
				 strAttrValue = strBudgetDesc
			end if 
			
			Response.Write "<td><input type=text name=attrib" & rsAttrib("intItem_Attrib_ID") & _
							" value=""" & strAttrValue & """ " & _
							rsAttrib("szForm_Attrib") & " ></td></tr>"
		elseif rsAttrib("intForm_Element_ID") = "2" then
			' Select List	
			call vbfBeginRow(rsAttrib("szName"))		
			strSQL =  rsAttrib("szOption_SQL") 
			Response.Write "<td><select name=attrib" & rsAttrib("intItem_Attrib_ID") & " " & rsAttrib("szForm_Attrib") & " >"
			Response.Write "<option value=''>"
			if instr(1,ucase(strSQL),"SELECT") > 0 then 
				' Create a select list from a sql statement
				if rsAttrib("szReplace_Text") & "" <> "" then
						' This section provides dynamic replacement of a value that 
						' a sql statement may need in order to filter a query based on 
						' specific criteria that is not able to be hard coded. 
						' So let's say we have a query that needs to be filtered on
						' intStudent_id but we to dynamically provide the student id 
						' based on a vbs variable. In the trefItem_Attrib
						' we write a sql that looks like 
						'	Select * from table where intStudent_ID = replace_intStudent_ID
						' the 'replace_' string precedes the variable name that we want replaced
						' in the sql statement.  So in this case we are saying we want replace_intStudent_ID
						' to be replace with the vbs variable intStudent_ID.  We know this because 
						' the field 'szReplace_Text' in trefItem_Attrib also has the value 'replace_intStudent_ID'
						' The code below does the work.
						' NOTE:  You can only replace variables that are accessible
						'		 to this script. It doesn't matter if they are
						'		 session,request, vbs, etc variables. 
						
						' Strips the 'replace_' off so we have just the variable name we are looking for
						strValueToGet = replace(rsAttrib("szReplace_Text"),"replace_","")
						' Assigns the desired variable value to strValueToGet
						execute("strValueToGet=" & strValueToGet)
						' replaces our 'replace_variable' holder in the sql with the
						' actual value we are looking for.							
						strSQL = replace(strSQL,rsAttrib("szReplace_Text"),strValueToGet)
					end if 
				Response.Write oFunc.MakeListSQL(strSQL,"","",strAttrValue)
			else	
				if inStr(1,strSQL,"|") > 1 then
					' Creates a select list where the value and text values are
					' different. Consists of a pipe deliminated list
					' which is also a comma seperated list. The pipe seperates
					' the 'value' list from the 'text' list and the comma's
					' seperate each item in that list.
					arLists = split(strSQL,"|")	
					strVal = arLists(0)
					strText = arLists(1)
					Response.Write oFunc.MakeList(strVal,strText,strAttrValue)
				else	
					' Creates a select list from a comma seperated list
					' where the value and text are the same			
					Response.Write oFunc.MakeList(strSQL,strSQL,strAttrValue)
				end if 
			end if
			Response.Write "</select></td></tr>"
		elseif rsAttrib("intForm_Element_ID") = "3" then
			response.Write "<td class='svplain8' colspan=2>" & rsAttrib("szForm_Attrib") & "</td></tr>"
		end if		

		strAttrib = strAttrib & rsAttrib("intItem_Attrib_ID") & ","
		' Adds any validation for the form element if specified in tascItems_Attrib
		if rsAttrib("szJS_Function") <> "" then
			' szJS_Function may have string segments in it that need to be parsed.
			' Currrently string segments that can be parsed are the FIELD name
			' so the java script knows which field to validate against and
			' FRIENDLYNAME which is the name of the field that the user can 
			' recognize and will be displayed in an alert box to give the
			' user further instructions.
			
			strJF = rsAttrib("szJS_Function")
			strJF = replace(strJF,"~*FIELD*~","attrib" & rsAttrib("intItem_Attrib_ID"))
			strJF = replace(strJF,"~*FRIENDLYNAME*~", rsAttrib("szName"))
			strValidate = strValidate & strJF & chr(13)
		end if
	%>
						</td>
					</tr>
					<%
		rsAttrib.MoveNext
	loop
	' Cut off the trailing comma 
	if len(strAttrib) > 0 then
		strAttrib = left(strAttrib,len(strAttrib)-1)
	end if	
	rsAttrib.Close
	set rsAttrib = nothing
	
	if isNumeric(curUnit_Price) and isNumeric(intQty) then
		if curShipping = "" then
			curShipping = 0
		end if
		intTotal = (curUnit_Price*intQty) + curShipping
	end if 	
	
	if intStudent_Id = "" and (session.Contents("strRole") = "ADMIN" _
		or session.Contents("strRole") = "TEACHER") then
	%>
					<tr>
						<td class="gray">
							&nbsp;Required Item:
						</td>
						<td>
							<select name="bolRequired" ID="Select3">
								<%
										response.Write oFunc.MakeList("0,1","No,Yes",bolRequired)
									%>
							</select>
						</td>
					</tr>
					<%
	end if
	%>
					<tr>
						<Td class="TableHeader">
							&nbsp;<% = strQTYText %>:
						</Td>
						<td>
							<input type=text name=intQty size=4 onChange="jfTotal(this.form);" value="<%=intQty%>" ID="Text1">
						</td>
					</tr>
					<tr>
						<td class="TableHeader">
							&nbsp;<% = strPriceText %>:
						</td>
						<td>
							<input type=text name=curUnit_Price size=8 onChange="jfTotal(this.form);" value="$<%=formatNumber(curUnit_Price,2)%>" ID="Text2">
						</td>
					</tr>
					<% if intItem_Group_ID = 2 then %>
					<tr>
						<td class="TableHeader">
							&nbsp;Shipping/Handling/Fees:
						</td>
						<td>
							<input type=text name=curShipping size=8 onChange="jfTotal(this.form);" value="$<%=formatNumber(curShipping,2)%>" ID="Text4">
						</td>
					</tr>
					<% else %>
					<input type="hidden" name="curShipping" value="0">
					<% end if %>
					<tr>
						<td class="TableHeader">
							&nbsp;Total:
						</td>
						<td>
							<input type=text name=intTotal size=8 disabled value="$<% = intTotal %>" ID="Text3">
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<input type=hidden name=strAttribList value="<% = strAttrib %>" ID="Hidden12">
		<script language="javascript">
		var mBolClean = true;		// Toggle to let us know if the form is clean or not
									// To false if validation fails.
		function jfAutoValidate(pFrm){
			// Validates all specified fields
			<% = strValidate %>		
			jfRequiredField('intQty','Number of Units');
			jfRequiredField('curUnit_Price','Unit Price');
			if (mBolClean != false){
				var intQty = pFrm.intQty.value;
				var intPrice = pFrm.curUnit_Price.value;
				var intTotal = pFrm.intTotal.value;
				var sErr = "";
				intPrice = intPrice.replace("$","");
				intPrice = parseFloat(intPrice);
				intTotal = intTotal.replace("$","");
				intTotal = parseFloat(intTotal);
					
				if (isNaN(intQty) == true ){
					sErr += "'Number of Units' must be a valid number.\n";
					mBolClean = false;
				}else if(intQty < .01){
					sErr += "'Number of Units' must be greater than 0.\n";
					mBolClean = false;
				}
				
				if(isNaN(intPrice) == true){
					sErr += "'Unit Price' must be a valid number.\n";
					mBolClean = false;
				}else if(intPrice < .01){
					sErr += "'Unit Price' must be greater than 0.\n";
					mBolClean = false;
				}
				
				<%if ucase(session.contents("strRole")) <> "ADMIN" then%> 
				if(intTotal < 0) {
					sErr += "Total cost can not be less than 0.";
					mBolClean = false;
				}
				<%end if: if intStudent_id <> "" then %>else if(mBolClean){
					jfCheckBudget();
				}
				 <% end if %>
			}
			if (mBolClean == false) {
				// Do nothing. The user should have been alerted to existing
				// problems with their submission from previously called
				// javascript functions.
				if (sErr != "") { alert(sErr);}
				mBolClean = true;
				return false;
			}else{
				// Submit the form.  All validation has been passed.
				pFrm.action = "<%=Application("strSSLWebRoot")%>forms/requisitions/reqGoodsInsert.asp";
				pFrm.submit();			
			}	
		}
		function jfCheckDate(val){
			var arDate = val.split('/');
			if (!isDate(arDate[2],arDate[0],arDate[1])){
				alert("'" + val + "' is an invalid date. Please fix.");
				mBolClean = false;
			}					
		}
		
		function jfRequiredField(fieldName,friendlyName){
			// Requires that 'fieldName' is not blank
			if (eval("document.main." + fieldName + ".value == \"\"")) {
				alert("'" + friendlyName + "' can not be blank.");
				mBolClean = false;
			}			
		}
		
		function jfTotal(pFrm)	{
			// Calculates totals for each resource. Executes when onChange is fired for
			// itemQty or itemPrice text box elements.
			if (pFrm.intQty.value != "" && document.main.curUnit_Price.value != "") {
				var intShipping = parseFloat("0");
				var intQty = pFrm.intQty.value;
				var intPrice = pFrm.curUnit_Price.value ;
				intPrice = intPrice.replace("$","");
				intPrice = parseFloat(intPrice);
				<% if false then%>	
				if (intPrice < .01) {
					pFrm.curUnit_Price.value = "";
					alert("The Unit Price must be greater than 0.");
					return false;				
				}	
				if (intQty < 1) {
					pFrm.intQty.value = "";
					alert("Number of Units must be greater than 0.");
					return false;				
				}		
				<% end if%>
				<% if intItem_Group_ID = 2 then %>
				if(document.main.curShipping.value != ""){
					intShipping = pFrm.curShipping.value ;
					intShipping = intShipping.replace("$","");
					intShipping = parseFloat(intShipping);
					<% if ucase(session.Contents("strRole")) <> "ADMIN" then%>	
					if(intShipping < 0) {
					pFrm.curShipping.value = "";
					alert("Shipping amount must be 0 or greater.");
					pFrm.intTotal.value = "$" + (intPrice * intQty);
					return false;	
					}
					<% end if %>
				}
				<% end if%>				
				pFrm.intTotal.value = "$" + ((intPrice * intQty)+intShipping);
			}
		}
<% if intStudent_ID <> "" then %>
</script>
			<% 
				set oBudget = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/StudentBudgetInfo.wsc"))
				oBudget.PopulateStudentFunding oFunc.FPCSCnn,intStudent_ID, session.contents("intSchool_Year") 
				myBudget = oBudget.BudgetBalance
				bolLimit = false
				if oFunc.IsSpendingLimitSubject(intPOS_SUBJECT_ID) then
					oBudget.PopulateFamilyBudgetInfo oFunc.FpcsCnn, oBudget.FamilyId, session.contents("intSchool_Year") 
					if oBudget.BudgetBalance > oBudget.AvailableElectiveBudget then
						myBudget = oBudget.AvailableElectiveBudget
						bolLimit = true
					end if						
				end if
				set oBudget = nothing
			%>
<script language="javascript">
		function jfCheckBudget(){
			var total = document.main.intTotal.value;
			total = total.replace("$","");
			total = parseFloat(total);
			var remainingBudget = parseFloat('<% = round(myBudget,2) %>');
			var startTotal = parseFloat('<% = intTotal %>');
			var amountAfterChanges = remainingBudget - (total - startTotal);
			if (amountAfterChanges < 0){	
				<% if ucase(session.contents("strRole")) <> "ADMIN" then %>			
				var strError = "<% if bolLimit then %>This budgeted item is subject to the Famliy Elective Spending Limit.\nThis family has $<% = round(myBudget,2) %> left for elective spending.\n<% end if %>Please check the price and quantity you entered. Adding this item will put you over budget by -$";
				strError += round(amountAfterChanges*-1,2) + ". Adjusting your students budget ";
				strError += "or making a budget transfer may create the funds needed to purchase this item.";
				strError += " Click the 'Close without saving' button to exit.";			
				alert(strError);
				mBolClean = false;
				<% else %>
				var strError = "If you add this item you will put this student over budget by -$";
				strError += round(amountAfterChanges*-1,2) + ". Are you sure you want to continue?";
				var bolContinue = confirm(strError);
				if (bolContinue) {
					mBolClean = true;
				}else{
					mBolClean = false;
				}
				
				<% end if %>
			}else{
				mBolClean = true;
			}
		}
		<% end if %>
		function round(number,X) {
			// rounds number to X decimal places, defaults to 2
			X = (!X ? 2 : X);
			return Math.floor(number*Math.pow(10,X))/Math.pow(10,X);
		}
		</script>
		<%
end function 

function vbfBeginRow(name)
	' This function simply prints the begining html for a new row in our
	' attribute form 
%>
		<tr>
			<Td class="TableHeader">
				&nbsp;<% = name %>:
			</Td>
			<%
end function
%>
	</table>
	<input type="button" value="Close without saving" onClick="window.opener.location.reload();window.opener.focus();window.close();"
		class="navLink" NAME="Button1">
	<%
			if (intItem_ID <> "" and (bolApproved & "" = "" or trim(ucase(bolApproved)) = "NULL") and intVendor_ID & "" <> "") or (session.Contents("strRole") = "ADMIN" and bolApproved & "" <> "") then
				' with bolApproved only show save if null
				'  check to see if year is locked
				if (not oFunc.LockSpending and not oFunc.LockYear and not (ucase(session.Contents("strRole")) = "GUARD" and bolRequired & "" = "1") and not bolTeacherLock) _ 
					or (session.Contents("strRole") = "ADMIN") then

			%>
	<input type="hidden" name="hdnContinue" value="" ID="Hidden14"> <input type="button" value="Save & Close" onClick="jfAutoValidate(this.form);" class="NavSave"
		NAME="Button5"> <input type="button" value="Save & Continue" onClick="this.form.hdnContinue.value='true';jfAutoValidate(this.form);"
		class="NavSave" NAME="Button5">
	<% 
				end if
			end if%>
	<br>
	<br>
	<span class="svplain8" style="color:red;">
			<B>Note:</B> Any fields left blank can slow down the processing of your order.
			</span>
	</td> </tr> </table>
</form>
<%
   call oFunc.CloseCN()
   set oFunc = nothing
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")
%>
