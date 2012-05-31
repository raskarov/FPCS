<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		reqApprovalAdmin.asp
'Purpose:	Gives admin the abilty to view and approve/deny goods and
'			services.
'Date:		21 JAN 2003
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Dimension Variables
dim intCount			'used as a tracking mechanism for our update subroutine 
dim sql					'sql that helps us populate our form		
dim rsGetItems			'recordset that helps us populate our form
dim strOrderBy			'order by for sql statement
dim intNumToShow		'Number of rows to return in sql statement
dim strMessage			'used to displays javascript messages
dim strColor			'used to set alternating bgcolor for rows in table
dim strFrom				'used tp return approval list for admin or teacher
dim strApprovedField	'used to define the approved field name 
dim strWhere			'Defines additional search criteria based on user selection
dim strWhere2			'Defines additional search criteria based on role 
dim strApprovedStatus	'Defines where good/service status is approved/denied/pending
dim strLabels
dim intPageCount
dim oHtml
dim oFunc
dim intCount2

' Only Admins may view this page. Stop all others
if session.Contents("strRole") <> "ADMIN" and session.Contents("strRole") <> "TEACHER" then
%>
<html>
<body>
<h1>Page Improperly Called.</h1>
</body>
</html>
<%
	response.End
end if

Session.Value("strTitle") = "Goods/Services Approval Page"
Session.Value("strLastUpdate") = "21 JAN 2003"
Server.Execute(Application.Value("strWebRoot") & "includes/Header.asp")

'Create object containing all of our FPCS functions
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
set oHtml = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/htmlFuncs.wsc"))
call oFunc.OpenCN()

'Define sql and set up which bolApproved to use
if ucase(session.Contents("strRole")) = "ADMIN" then
	strFrom = "FROM tblOrdered_Items oi INNER JOIN " & _
				"tblVendors v ON oi.intVendor_ID = v.intVendor_ID INNER JOIN " & _
				"tblSTUDENT s ON oi.intStudent_ID = s.intSTUDENT_ID INNER JOIN " & _
				"trefItems i ON oi.intItem_ID = i.intItem_ID LEFT OUTER JOIN " & _
				"tblFAMILY f ON s.intFamily_ID = f.intFamily_ID  INNER JOIN " & _
				" tblILP ilp ON oi.intILP_ID = ilp.intILP_ID INNER JOIN " & _
				" tblClasses c ON ilp.intClass_ID = c.intClass_ID LEFT OUTER JOIN " & _
				" tblINSTRUCTOR ins ON c.intInstructor_ID = ins.intINSTRUCTOR_ID INNER JOIN " & _
				" tblSTUDENT_STATES ss ON ss.intStudent_ID = s.intStudent_ID INNER JOIN "   & _
				" tblEnroll_Info ei ON ei.sintSchool_Year = oi.intSchool_Year AND ei.intStudent_ID = s.intStudent_ID " 
	strApprovedField = "bolApproved"
else 'if session.Contents("strRole") = "TEACHER" then
	response.Write "<h1>Page Improperly Called</h1>"
	response.End
	'strFrom = "FROM tblENROLL_INFO INNER JOIN " & _
	'			" tblOrdered_Items oi INNER JOIN " & _
	'			" tblVendors v ON oi.intVendor_ID = v.intVendor_ID INNER JOIN " & _
	'			" tblSTUDENT s ON oi.intStudent_ID = s.intSTUDENT_ID INNER JOIN " & _
	'			" trefItems i ON oi.intItem_ID = i.intItem_ID INNER JOIN " & _
	'			" tblFAMILY f ON s.intFamily_ID = f.intFamily_ID INNER JOIN " & _
	'			" tblILP ilp ON oi.intILP_ID = ilp.intILP_ID INNER JOIN " & _
	'			" tblClasses c ON ilp.intClass_ID = c.intClass_ID ON " & _
	''			" tblENROLL_INFO.intSTUDENT_ID = s.intSTUDENT_ID LEFT OUTER JOIN " & _
	'			" tblINSTRUCTOR ins ON c.intInstructor_ID = ins.intINSTRUCTOR_ID "
	'strWhere2 = " and (tblENROLL_INFO.intSponsor_Teacher_ID = " & session.Contents("instruct_ID") & ") "
	'strApprovedField = "bolSponsor_Approved"
end if

if (request.Form("updateList") <> "" and request.Form("updateList") <> ",") or _
	(request.Form("LineItemsChanged") <> "," and request.Form("LineItemsChanged") <> "") then
	oFunc.BeginTransCN
end if

'Check to see if we need to save any changes
if request.Form("updateList") <> "" and request.Form("updateList") <> "," then
	call vbsSaveChanges
end if
if request.Form("LineItemsChanged") <> "," and request.Form("LineItemsChanged") <> "" then
	call vbsSaveLineItems
end if

if (request.Form("updateList") <> "" and request.Form("updateList") <> ",") or _
	(request.Form("LineItemsChanged") <> "," and request.Form("LineItemsChanged") <> "") then
	oFunc.CommitTransCN
end if

set rsGetItems = server.CreateObject("ADODB.Recordset")
rsGetItems.CursorLocation = 3 'adUseClient

if request.Form("orderBy") <> "" then
	strOrderBy = "Order by " & replace(request.Form("orderBy"),"~",",") & ", s.szLAST_NAME,s.szFIRST_NAME " 
else
	strOrderBy = "ORDER BY s.szLAST_NAME,s.szFIRST_NAME "
end if

if request("byItems") <> "" then
	select case request("byItems")
		case "1"
			sqlAdd = " and (oi.intItem_ID = 1 or oi.intItem_ID = 5 or oi.intItem_ID = 6) and (bolReimburse = 0 OR  bolReimburse is null) "
		case "2"
			sqlAdd = " and (oi.intItem_ID = 2 or oi.intItem_ID = 3) and (bolReimburse = 0 OR  bolReimburse is null) "
		case "3"
			sqlAdd = " and (oi.intItem_ID = 4 or oi.intItem_ID = 7 or oi.intItem_ID = 8 or oi.intItem_ID = 9) and (bolReimburse = 0 OR  bolReimburse is null) "
		case "4" 'Goods Only
			sqlAdd = " and (i.intItem_Group_ID = 2) "
		case "5" ' Services Only
			sqlAdd = " and (i.intItem_Group_ID = 1) "
	end select
end if

if request("bStatus") & "" <> "" then
	strWhere = strWhere & " AND oi." & strApprovedField & " " & request("bStatus")
end if

if request("aStatus") & "" <> "" then
	select case request("aStatus")
		case 1
			'signed
			' Whenever a sponsor teacher is also the ASD teacher of a class there is a 
			' business rule that is enforced by the code in packet.asp that will
			' set both the values for the sponser teacher as instructor fields
			
			strWhere = strWhere & " AND (ilp.GuardianStatusId = 1 and ilp.SponsorStatusID = 1 and ((c.intInstructor_ID is not null and ilp.InstructorStatusID = 1) or c.intInstructor_ID is null) and (ilp.AdminStatusID = 1 or c.intContract_Status_ID = 5)) "
	    case 2
			'not signed
			strWhere = strWhere & " AND (ilp.GuardianStatusId is null or ilp.SponsorStatusID is null or (c.intInstructor_ID is not null and ilp.InstructorStatusID is null) " & _
					   " or (ilp.AdminStatusID is null and c.intContract_Status_ID <> 5)) and ((ilp.InstructorStatusID <> 3 or ilp.InstructorStatusID is null) " & _
					   " and (ilp.SponsorStatusID not in (2,3) or ilp.SponsorStatusID is null) and (ilp.AdminStatusID not in(2,3) or ilp.AdminStatusID is null)) " 			   
		case 3
			' must Amend
			strWhere = strWhere & " AND (ilp.SponsorStatusID = 2 or ilp.AdminStatusID = 2) "
		case 4
			' rejected
			strWhere = strWhere & " AND (ilp.SponsorStatusID = 3 or ilp.AdminStatusID = 3 or ilp.InstructorStatusID = 3) "
	end select
end if

if request("searchField") <> "" then
	strWhere = strWhere & " and upper(" & request("searchField") & ") like '%" & ucase(ofunc.escapeTick(request("strKeyWord"))) & "%' "
end if

if request("type") <> "" then
	if request("type") = 0 then
		strWhere = strWhere & " and (oi.bolReimburse = 0 or oi.bolReimburse IS NULL) "
	else
		strWhere = strWhere & " and oi.bolReimburse = 1 "
	end if
end if 

if request("closedStatus") & "" = "" then
	myClosedStatus = " AND (oi.bolClosed = 0 or oi.bolClosed is null) "
else
	myClosedStatus = request("closedStatus")
end if

' Lets get the list of Goods/Sevices that need approval
sql = "SELECT v.intVendor_ID, oi.intOrdered_Item_ID, v.szVendor_Name, s.szLAST_NAME, s.szFIRST_NAME, " & _
	  "i.szName, ilp.GuardianStatusID,ilp.SponsorStatusID,ilp.InstructorStatusID,ilp.AdminStatusID, c.intContract_Status_ID," & _
	  "ilp.GuardianStatusDate, ilp.SponsorStatusDate, ilp.InstructorStatusDate, ilp.AdminStatusDate, c.dtApproved, " & _
	  "oi.intQty, oi.curUnit_Price,oi.curShipping, oi.intILP_ID, oi.dtCREATE, " & _
	  " f.szDesc, f.szHome_Phone, f.szEMAIL, v.szVendor_Phone, " & _
	  "c.szClass_Name, ilp.intContract_Guardian_ID, ins.szLAST_NAME + ', ' + ins.szFIRST_NAME as teacher " & _
	  ",c.intInstruct_Type_Id, c.intGuardian_Id , c.intClass_ID, c.intInstructor_ID, " & _
	  "s.intStudent_ID, i.intItem_Group_ID, c.intPOS_SUBJECT_ID,oi.dtCreate,oi.szUser_Create, oi." & strApprovedField & ", " & _
	  " oi.szDeny_Reason, oi.bolClosed, v.bolService_Vendor " & _
	  ", (oi.curUnit_Price*oi.intQty)+oi.curShipping as total, ss.intReEnroll_State, " & _
	  " ei.intSponsor_Teacher_ID AS Sponsor_ID, oi.intItem_ID, oi.InventoryDetailID,  i.InventoryCategoryID  " & _
	  strFrom & _
	  "WHERE (oi.intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _
	  " AND (ss.intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _
	  myClosedStatus & _
	  strWhere & strWhere2 & " " & _
	  sqlAdd & strOrderBy

	  'response.End
if request("PageNumber") <> "" then
	intPageNum = cint(request("PageNumber"))	
else
	intPageNum = 1
end if

if request.Form("numberToShow") <> "" then
	intNumToShow = request.Form("numberToShow")
else
	intNumToShow = 25
end if 
if ucase(session.Contents("strUserID")) = "SCOTT" then
	'response.Write sql
end if 

rsGetItems.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
if rsGetItems.RecordCount > 0 then
	rsGetItems.PageSize = intNumToShow
	rsGetItems.AbsolutePage = intPageNum
end if
dim intViewingTo
intViewingTo = rsGetItems.AbsolutePosition + rsGetItems.PageSize -1 
if intViewingTo > rsGetItems.recordcount then intViewingTo = rsGetItems.RecordCount
'Start by printing title
%>
<script language=javascript>
<!-- hide from browsers
	<% if ucase(session.Contents("strRole")) = "ADMIN" then %>
		var winRefresh;
		var strURL;		
		strURL = "./Refresher.asp";
		winRefresh = window.open(strURL,"winRefresh","width=300,height=300,scrollbars=yes,resizable=yes");
		winRefresh.moveTo(0,0);
		winRefresh.focus();
	<% end if %>
	function jfViewILP(id,class_name,teacherName,cg){
		var winILP;
		var strURL;		
		strURL = "../ilp/ilpMain.asp?plain=yes&intILP_ID=" + id;
		strURL += "&szClass_Name=" + class_name;
		strURL += "&strTeacherName=" + teacherName;
		strURL += "&intContract_Guardian_ID=" + cg;
		winILP = window.open(strURL,"winILP","width=710,height=500,scrollbars=yes,resizable=yes");
		winILP.moveTo(0,0);
		winILP.focus();
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
	
	function jfViewPacket(student){
			var winPacket;
			var url;
			url = "<%=Application.Value("strWebRoot")%>forms/packet/packet.asp?simpleHeader=true&intStudent_ID=" + student;
			winPacket = window.open(url,"winPacket","width=800,height=500,scrollbars=yes,resizable=yes");
			winPacket.moveTo(0,0);
			winPacket.focus();
	}
	function jfViewPrint(student){
			var winPrint1;
			var url;
			url = "<%=Application.Value("strWebRoot")%>forms/printableForms/printDefault.asp?simpleHeader=true&intStudent_ID=" + student;
			winPrint1 = window.open(url,"winPrint1","width=750,height=500,scrollbars=yes,resizable=yes");
			winPrint1.moveTo(0,0);
			winPrint1.focus();
	}
	
	function jfViewVendorReport(vendorId){
			var winPrint2;
			var url;
			url = "<%=Application.Value("strWebRoot")%>Reports/VendorServiceReport.asp?simpleHeader=true&detail=true&intVendor_ID=" + vendorId;
			winPrint1 = window.open(url,"winPrint2","width=750,height=500,scrollbars=yes,resizable=yes");
			winPrint1.moveTo(0,0);
			winPrint1.focus();
	}	
	
	function jfAddInventory(oID){
		var winInventory;
		var url;
		url = "<%=Application.Value("strWebRoot")%>Inventory/InventoryAdmin.asp?simpleHeader=true&intOrdered_Item_ID=" + oID;
		winInventory = window.open(url,"winInventory","width=900,height=500,scrollbars=yes,resizable=yes");
		winInventory.moveTo(0,0);
		winInventory.focus();
	}
	
	function jfViewInventory(oID){
		var winInventory;
		var url;
		url = "<%=Application.Value("strWebRoot")%>Inventory/InventoryAdmin.asp?simpleHeader=true&InventoryDetailID=" + oID;
		winInventory = window.open(url,"winInventory","width=900,height=500,scrollbars=yes,resizable=yes");
		winInventory.moveTo(0,0);
		winInventory.focus();
	}
	
	function jfViewBudget(student){
			var winBudget;
			var url;
			url = "<%=Application.Value("strWebRoot")%>forms/budget/budgetWorkSheet.asp?simpleHeader=true&intStudent_ID=" + student;
			winBudget = window.open(url,"winBudget","width=750,height=500,scrollbars=yes,resizable=yes");
			winBudget.moveTo(0,0);
			winBudget.focus();
	}
	<% = strMessage %>
	function jfActiveRow(id){
		var objCol = document.getElementsByTagName("TR");
		var i;
		for (i=0;i<objCol.length;i++){
			objCol[i].style.background = "ffffff";
		}
		var objRow = document.getElementById(id);
		objRow.style.background = "e6e6e6";
		
	}
	
	function jfCalcLineItem(pBudgetAmount,pBudgetID,pLineItemID){
		// handle updating specifi line item	
		var dUnits = document.getElementById("intQuantity"+pLineItemID);
		dUnits.value = (dUnits.value != "")?dUnits.value:1;
		var dPrice= document.getElementById("curUnit_Price"+pLineItemID);
		dPrice.value = (dPrice.value != "")?dPrice.value:0;
		var dShip = document.getElementById("curShipping"+pLineItemID);
		dShip.value = (dShip.value != "")?dShip.value:0;
		var dTotal = document.getElementById("Total"+pLineItemID);
		dTotal.value = (parseFloat(dUnits.value)*parseFloat(dPrice.value))+parseFloat(dShip.value);
		
		// Now based on all line items for a given budget calculate the balance
		var sGList = document.getElementById("liGroupList"+	pBudgetID);
		var aList = sGList.value.split(",");
		var dBalance = 0;
		var i;
		for (i=0;i<aList.length;i++){
			if (aList[i] != "") {
				dBalance += jfGetLineItemTotal(aList[i]);
			}
		}
		
		var tBalance = document.getElementById("Balance"+pBudgetID);
		pBudgetAmount = parseFloat(pBudgetAmount);
		tBalance.innerHTML = "$" + round((pBudgetAmount - dBalance),2);			
	}
	
	function jfGetLineItemTotal(pID){
		var dUnits = document.getElementById("intQuantity"+pID);
		var dPrice= document.getElementById("curUnit_Price"+pID);
		var dShip = document.getElementById("curShipping"+pID);
		return (parseFloat((dUnits.value != "")?dUnits.value:1)*parseFloat((dPrice.value != "")?dPrice.value:0))+parseFloat((dShip.value != "")?dShip.value:0)
	}
	
	function round(number,X) {
		// rounds number to X decimal places, defaults to 2
		X = (!X ? 2 : X);
		var val = parseFloat(val);
		val = number*Math.pow(10,X)/Math.pow(10,X);
		
		val2 = val;
		if (parseInt(val2+.5) > parseInt(val2)) {			
			return Math.ceil(number*Math.pow(10,X))/Math.pow(10,X);
		}else{
			return Math.floor(number*Math.pow(10,X))/Math.pow(10,X);
		} 
	}
	
	function jfPrintList(pID,pObj){
		var sIDs = document.getElementById('GSList');
		sIDs.value = sIDs.value.replace(","+pID+",",",");
		if (pObj.checked){ sIDs.value = sIDs.value + pID + ","; }
		document.getElementById('GSList').value = sIDs.value;
	}
	
	function jfPrintGS(){
		var winPrint;
		var GSList = document.main.GSList.value;	
		strURL = "<%=Application.Value("strWebRoot")%>forms/PrintableForms/GoodsServiceDetail.asp?GSList=" + GSList;
		winPrint = window.open(strURL,"winPrint","width=710,height=500,scrollbars=yes,resizable=yes");
		winPrint.moveTo(0,0);
		winPrint.focus();
	}
-->
</script>
<table width=100% ID="Table1">
	<form action="reqApprovalAdmin.asp" method=post name=main ID="Form1">
	<input type="hidden" name="PageNumber" value="<% = intPageNum%>" ID="Hidden7">
	<tr>	
		<Td class=yellowHeader>
			&nbsp;<b>Goods/Services Approval Admin</b>
		</td>
	</tr>
	<tr>
		<td>
			<table ID="Table2" cellspacing=0>
				<tr>					
					<td>
						<table ID="Table3">		
							<td colspan="10" class=navywhite8 valign=middle>
								<B>&nbsp;Filter Criteria&nbsp;</B>
							</td>
							<tr>
								<td	class=gray>
									<b><nobr>&nbsp;Course Status&nbsp;</nobr></b>
								</td>
								<td	class=gray>
									<b><nobr>&nbsp;B Status&nbsp;</nobr></b>
								</td>
								<td	class=gray>
									&nbsp;<b># to Show</b>&nbsp;
								</td>
								<td	class=gray>
									&nbsp;<b>Type</b>&nbsp;
								</td>
								<td	class=gray>
									&nbsp;<b>Order By</b>&nbsp;
								</td>
								<td	class=gray>
									&nbsp;<b>Items</b>&nbsp;&nbsp;&nbsp;
								</td>	
								<td	class=gray>
									&nbsp;<b>Closed Status</b>&nbsp;
								</td>	
							</tr>
							<tr>
								<td>
									<select name="aStatus" ID="Select7" onchange="this.form.PageNumber.value='1';">
									<%
										response.Write oFunc.makeList(",1,2,3,4",",Signed Date,Not Signed,Must Amend,Rejected",request("aStatus"))
									%>						
									</select>
								</td>
								<td>
									<select name="bStatus" ID="Select4" onchange="this.form.PageNumber.value='1';">
									<%
										response.Write oFunc.makeList(",=1,=0, IS NULL","All,b-appr,b-rejc,b-pend",request("bStatus"))
									%>						
									</select>
								</td>
								<td>
									<select name="numberToShow" ID="Select8" align="center" onchange="this.form.PageNumber.value='1';">
										<%
											response.Write oFunc.makeList("25,50,100,150,200,300","25,50,100,150,200,300",request.Form("numberToShow"))							
										%>
									</select>
								</td>	
								<td>
									<select name="type" ID="Select9">
										<option value="" onchange="this.form.PageNumber.value='1';">All
										<%
											response.Write oFunc.makeList("0,1","Req Only,Reim Only",request.Form("type"))							
										%>
									</select>
								</td>	
								<td>
									<select name="orderby" ID="Select2" onchange="this.form.PageNumber.value='1';">
										<option value="">Students Name
										<%
											response.Write oFunc.makeList("i.szName,oi.dtCreate DESC,oi.dtModify~oi.dtCREATE,v.szVendor_Name","Item Type,Newest Date,Oldest Date,Vendor Name",request.Form("orderby"))							
										%>
									</select>
								</td>					
								<td>
									<select name="byItems" ID="Select3" onchange="this.form.PageNumber.value='1';">
										<option value="">All Items
										<%
											response.Write oFunc.makeList("1,4,2,3","ASD-Cirriculum-Supplies: Req,Goods Only,Services Only,Rental EQ: Comp-Building-Other: Req",request("byItems"))
										%>
									</select>
								</td>		
								<td>
									<select name="closedStatus" ID="Select6" onchange="this.form.PageNumber.value='1';">
										<%
											if request("closedStatus") & "" = "" then
												myClosedStatus = " AND (oi.bolClosed = 0 or oi.bolClosed is null) "
											else
												myClosedStatus = request("closedStatus")
											end if
											response.Write oFunc.makeList(" AND (oi.bolClosed = 0 or oi.bolClosed is null) ,AND 1 = 1 , AND oi.bolClosed = 1 ", "Open Only, Both Open and Closed, Closed Only" ,myClosedStatus)
										%>
									</select>									
								</td>																												
							</tr>
						</table>	
					</td>
				</tr>
				<tr>											
					<td>
						<table ID="Table5">
							<tr>
								<td	class=NavyWhite8 colspan="2">
									&nbsp;<b>Search Criteria</b>&nbsp;
								</td>
							</tr>
							<tr>
								<td	class=gray	>
									&nbsp;<b>Search Field</b>&nbsp;
								</td>	
								<td	class=gray>
									&nbsp;<b>Search Key Word</b>&nbsp;
								</td>	
								<td rowspan="2" valign="top" align="center" >
									&nbsp;&nbsp;<input type=button value="Save and Requery Data" class="NavSave" onClick="this.form.submit();" NAME="Button1" ID="Button1">
								</td>
							</tr>
							<tr>
								<td>
									<select name="searchField" ID="Select5" onchange="this.form.PageNumber.value='1';">
										<option value="">None
										<%
											response.Write oFunc.makeList("s.szLast_Name,v.szVendor_Name,oi.szUser_Create","Student Last Name,Vendor,User",request("searchField"))
										%>
									</select>
								</td>			
								<td>
									<input type=text name="strKeyWord" value="<% = request("strKeyWord")%>" size=22 maxlength=50 ID="Text1" onchange="this.form.PageNumber.value='1';">									
								</td>
							</tr>
						</table>
					</td>						
				</tr>
				<tr>
					<td colspan=10 class="svplain8" nowrap>
						
						&nbsp;viewing <% = rsGetItems.AbsolutePosition %> - <% = intViewingTo %>  of <% = rsGetItems.RecordCount %> records &nbsp;
						
						<table ID="Table4"><tr><td>
						<%
							if cint(rsGetItems.RecordCount) > cint(intNumToShow) then
							for i = 1 to rsGetItems.PageCount
						
								if intViewingTo/rsGetItems.PageSize = i or (rsGetItems.RecordCount = intViewingTo and i = rsGetItems.PageCount) then 
									strClass = "NavSave"
								else
									strClass = "btSmallWhite"
								end if
						%>
							<input type="button" class="<% = strClass %>" value="<%=i%>" onClick="SetPageScroll('<% = i %>GSA');this.form.PageNumber.value='<%=i%>';this.form.submit();" ID="Button2" NAME="Button2">
						<%
							next 
							end if
						%>
						</td></tr></table>
					</td>					
				</tr>
			</table>
		</td>
	</tr>
</table>
<%

if rsGetItems.RecordCount > 0 then
	' We've got some records so let's make the form
%>
<script language=javascript>
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
	
	function jfToggle(pList,pID){
		// toggles display of objs in pList on and off
		var arList = pList.split(",");
		var i;
		var obj;
		var sText;
		for(i=0;i< arList.length;i++){
			if (arList[i] != '') {
				obj = document.getElementById(arList[i]);
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
	
	function jfUpdateList(id) {
		// if an item as been changed log it on;y once.  We will use this list
		// to determine which OI's should be modified
		if (document.main.updatelist.value.indexOf(","+id+",") == -1 ) {
			document.main.updatelist.value = document.main.updatelist.value + id + ",";
		}
	}	
	function jfChangedLI(id){	
		if (document.main.LineItemsChanged.value.indexOf(","+id+",") == -1) {
			document.main.LineItemsChanged.value = document.main.LineItemsChanged.value + id + ",";
		}
	}	
	
	function jfVendProfile(pVendorID){
		var winVendProfile;
		var strURL = "<%=Application.Value("strWebRoot")%>forms/VIS/VendorAdmin.asp?intVendor_ID="+pVendorID;
		strURL += "&bolPrint=true&bolWin=True&intItem_Group_ID=<% = request("intItem_Group_ID")%>";
		winVendProfile = window.open(strURL,"winVendProfile","width=850,height=500,scrollbars=yes,resize=yes,resizable=yes");
		winVendProfile.moveTo(20,20);
		winVendProfile.focus();	
	}
</script>
<input type=hidden name="lastRow" ID="Hidden1">
<input type=hidden name="updatelist" value="," ID="Hidden2">
<input type=hidden name="GSList" value="," ID="Hidden13">
<input type=hidden name="LineItemsChanged" value="," ID="Hidden8">
<input type=hidden name="lastRowColor" ID="Hidden3">
<table cellpadding="2" ID="Table6">	
<%
	intCount = 0
	intCount2 = 0 
	intMax = (rsGetItems.AbsolutePosition + rsGetItems.PageSize)
	dim rsLI 
	set rsLI = server.CreateObject("ADODB.RECORDSET")
	rsLI.CursorLocation = 3
	
	do while rsGetItems.AbsolutePosition < intMax and not rsGetItems.EOF
	
		if intCount = 0 or intCount mod 25 = 0 then
			call PrintHeader
		end if
		
		' Set row color
		if intCount mod 2 = 0 then
			strColor = "plainCell"
		else
			strColor = "gray"
		end if
%>
	<tr class='<%=strColor%>'  id="ROW<%=intCount%>" onClick="jfHighLight('<%=intCount%>');">
		<input type=hidden name="intOrdered_Item_ID<%=intCount%>" value="<%=rsGetItems("intOrdered_Item_ID")%>" ID="Hidden4">
		<input type=hidden name="familyEmail<%=intCount%>" value="<%=rsGetItems("szEMAIL")%>" ID="Hidden5">
		<td title="Family Desc: <% = rsGetItems("szDesc") & chr(13) & "Phone: " & rsGetItems("szHome_Phone")%>">
			<% if rsGetItems("intReEnroll_State") <> 7 and rsGetItems("intReEnroll_State") <> 15 _
				and rsGetItems("intReEnroll_State") <> 31 and rsGetItems("intReEnroll_State") <> 129 then %>
			<span style="color:red;"><b><% = rsGetItems("szLast_Name") & ", " & rsGetItems("szFirst_Name") %></b></span> Not Active
			<%else%>
			<% = rsGetItems("szLast_Name") & ", " & rsGetItems("szFirst_Name") %>
			<%end if%>
			
		</td>
		<td title="Phone: <% = rsGetItems("szVendor_Phone") %>">
			<a href="#" onclick="jfVendProfile('<% = rsGetItems("intVendor_ID") %>');"><% = rsGetItems("szVendor_Name") %></a>
		</td>
		<td align="center">
			<input type="checkbox" onClick="jfPrintList('<% = rsGetItems("intOrdered_Item_ID") %>',this);" name="chk<% = rsGetItems("intOrdered_Item_ID") %>" ID="Checkbox2">
		</td>
		<td>			
			<% = rsGetItems("intOrdered_Item_ID") %></a>
		</td>
		<td  align=center>
			<a href="javascript:" title="View Class/Schedule for '<% = replace(rsGetItems("szClass_Name") & "","'","\'") %>'" onclick="jfViewClass('<%=rsGetItems("intClass_ID")%>','<%=rsGetItems("intInstructor_ID")%>','<%=rsGetItems("intInstruct_Type_Id")%>','<%=rsGetItems("intContract_Guardian_ID")%>','<%=rsGetItems("intGuardian_Id")%>');">
			 C</a> 
			<a href="javascript:" title="View ILP" onclick="jfViewILP('<%=rsGetItems("intILP_ID")%>','<%=replace(rsGetItems("szClass_Name") & "","'","\'")%>','<%=replace(rsGetItems("teacher") & "" ,"'","\'")%>','<% =rsGetItems("intContract_Guardian_ID")%>');">
			I</a>						 
			<a href="javascript:" title="View Goods/Services" onclick="jfViewItem('<%=rsGetItems("intILP_ID")%>','<%=rsGetItems("intStudent_ID")%>','<%=rsGetItems("intOrdered_Item_ID")%>','<%=rsGetItems("intItem_Group_ID")%>','<%=replace(replace(rsGetItems("szClass_Name") & "","'","\'"),"&"," and ")%>','<%=rsGetItems("intPOS_SUBJECT_ID")%>');"> 
			GS</a>
			<a href="javascript:" title="View Packet" onclick="jfViewPacket('<%=rsGetItems("intStudent_ID")%>');">
			 P</a>
			 <a href="javascript:" title="Printable Forms" onclick="jfViewPrint('<%=rsGetItems("intStudent_ID")%>');">
			 Prt</a>
			 <% if rsGetItems("bolService_Vendor") and rsGetItems("intItem_Group_ID") & "" = "1" and rsGetItems("intItem_ID") & "" = "3" then %>
			 <a href="javascript:" title="View Service Vendor Report" onclick="jfViewVendorReport('<%=rsGetItems("intVendor_ID")%>');">
			 VR</a>
			 <% end if %>
			 <% if (rsGetItems("InventoryCategoryID") & "" <> "") then %>
			 <%		if rsGetItems("InventoryDetailID") & "" = ""  then
						InvCaption = "Add to Inventory"
						InvLink = "+Inv"
						jfText = "jfAddInventory('" & rsGetItems("intOrdered_Item_ID") & "');"
					else
						InvCaption = "View Inventory Record"
						InvLink = "Inv"
						jfText = "jfViewInventory('" & rsGetItems("InventoryDetailID") & "');"
					end if
			 %>
			 <a href="javascript:" title="<% = InvCaption %>" onclick="<% = jfText %>">
			 <% = InvLink %></a>
			 <% end if %>
		</td>
		<td   align=right title="Date Created: <%=rsGetItems("dtCreate")%>">
			<% = rsGetItems("szUser_Create") %>
		</td>
		<td  align=right title="Number of Units = <% = rsGetItems("intQty") %>: Unit Price = <%= rsGetItems("curUnit_Price")%>">
			<%
				' get all accounting data that pertains to this budget
				sql = "SELECT intLine_Item_ID, dtLine_Item, szLine_Item_desc, curUnit_Price, intQuantity,  " & _ 
					  "curShipping, szPO_Number, szInvoice_Number, szCheck_Number,  " & _ 
					  "szPayee, dtCheck, dtReciept, dtCREATE, dtMODIFY, szUSER_CREATE, szUSER_MODIFY " & _ 
					  ", (curUnit_Price*intQuantity)+curShipping as total " & _
					  "FROM tblLine_Items " & _ 
					  "WHERE (intOrdered_Item_ID = "  & rsGetItems("intOrdered_Item_ID") & ") " & _
					  " ORDER BY intLine_Item_ID"
				rsLI.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
				dblLineItemCosts = 0
				if rsLI.RecordCount > 0 then
					do while not rsLI.EOF
						dblLineItemCosts = dblLineItemCosts + rsLI("total")
						rsLi.MoveNext
					loop
					rsLI.MoveFirst
				else
					dblLineItemCosts = 0 
				end if
				dblBudgetAmount = rsGetItems("total")
			 %>
			$<% = formatNumber(dblBudgetAmount,2) %>
		</td>
		<td align="right" title="Click here to view Line Item entries." id="Balance<%=intCount%>" onClick="jfToggle('LineItem<%=intCount%>,','');">
			<a href="javascript:" onclick="return false">$<% = formatNumber(dblBudgetAmount-dblLineItemCosts,2) %></a>
		</td>
		<td align="center">
			<input type="checkbox" name="bolClosed<%=intCount%>" <% = oHtml.IIF(rsGetItems("bolClosed"),"checked","") %> value="1" onChange="jfUpdateList('<%=intcount%>');" ID="Checkbox1">
		</td>
		<td   align=center >			
		   <%  
				if rsGetItems("AdminStatusId") = "3" or rsGetItems("SponsorStatusId") = "3" or _
					rsGetItems("InstructorStatusId") = "3" then
						'Rejected 
						response.Write "rejected"
				elseif  rsGetItems("AdminStatusId")  = "2" or rsGetItems("SponsorStatusId") = "2" then
					' Needs Work
					response.Write "must ammend"
				elseif rsGetItems("GuardianStatusId") & "" = "1" and rsGetItems("SponsorStatusId") & "" = "1" and _
					(rsGetItems("AdminStatusId") & "" = "1" or rsGetItems("intContract_Status_Id") & "" = "5") and _
					(rsGetItems("InstructorStatusId") & ""  = "1" or _
					 rsGetItems("intInstructor_ID") & "" = "" or  _
					  (rsGetItems("intInstructor_ID") & "" <> "" and _
					   rsGetItems("intInstructor_ID") & "" = rsGetItems("Sponsor_ID") & "")) then
					' Signed
					myDate = oFunc.GreatestDate(array(rsGetItems("dtApproved"),rsGetItems("SponsorStatusDate"),rsGetItems("InstructorStatusDate"),rsGetItems("AdminStatusDate"),rsGetItems("GuardianStatusDate")))
					if isDate(myDate) then
						myDate = formatDateTime(myDate,2)
					end if
					response.Write "<span title='Date Course Approved: " & mydate & "'>" & mydate & "</span>"
				else
					' Not Signed
					response.Write "not signed"
				end if  	
		   %>
		</td>
		<td align=center >
			<!--<input type=checkbox name="approved<% = intCount%>" value="<% = rsGetItems("intOrdered_Item_ID")%>" <% if ucase(rsGetItems(strApprovedField)) & "" = "TRUE" then response.Write " checked " %>>-->
			<select name="approved<% = intCount%>" onChange="jfUpdateList('<%=intcount%>');" ID="Select1">
				<%
					if rsGetItems(strApprovedField) then
						strApprovedStatus = "1"
					elseif rsGetItems(strApprovedField) = false then
						strApprovedStatus = "0"
					else
						strApprovedStatus = "2"
					end if
					
					if ucase(session.Contents("strRole")) = "ADMIN" then
						strLabels = "b-pend,b-appr,b-rejc"
					else
						strLabels = "s-pend,s-appr,s-rejc"
					end if 
					response.Write oFunc.makeList("2,1,0",strLabels,strApprovedStatus)
				%>
			</select>
		</td>
		<td>
			<textarea cols=20 rows=1 wrap=virtual name="denied<% = intCount%>" onChange="jfUpdateList('<%=intcount%>');" onfocus="this.rows=4;" onblur="this.rows=1;" onKeyDown="jfMaxSize(511,this);" ID="Textarea1"><% = rsGetItems("szDeny_Reason") %></textarea>
		</td>
	</tr>
	<tr id="LineItem<%=intCount%>" style="display:none;" onClick="jfHighLight('<%=intCount%>');">
		<td colspan="12">
			<table style="width:100%;" class="DarkerBorder" ID="Table7">
				<tr class="TableHeader">					
					<td>
						PO#
					</td>
					<td>
						Invoice #
					</td>
					<td>
						Check #
					</td>
					<td>
						Check Date
					</td>
					<td>
						Payee
					</td>
					<td>
						Receipt Date
					</td>
					<td>
						Description
					</td>
					<td>
						Unit Price
					</td>
					<td>
						Qty
					</td>
					<td>
						Shipping
					</td>
					<td>
						Total
					</td>
				</tr>
				<%
				
				if rsLI.RecordCount > 0 then
					intLIIDLast = 0  
					do while not rsLI.EOF
						if intLIIDLast <> clng(rsLI("intLine_Item_ID")) then 
							strLIList = strLIList & intCount2 & ","
							intLIIDLast = rsLI("intLine_Item_ID")
						end if
						%>
				<tr>
					<td>
						<INPUT type="hidden" name="intLine_Item_ID<%=intCount2%>" value="<% = rsLI("intLine_Item_ID")%>" ID="Hidden6">
						<input size="6" type="text" name="szPO_Number<%=intCount2%>" value="<% = rsLI("szPO_Number")%>" maxlength="16" onchange="jfChangedLI('<%=intCount2%>');" ID="Text2">
					</td>
					<td>
						<input size="6" type="text" name="szInvoice_Number<%=intCount2%>" value="<% = rsLI("szInvoice_Number")%>" maxlength="16" onchange="jfChangedLI('<%=intCount2%>');" ID="Text3">
					</td>
					<td>
						<input size="6" type="text" name="szCheck_Number<%=intCount2%>" value="<% = rsLI("szCheck_Number")%>" maxlength="16" onchange="jfChangedLI('<%=intCount2%>');" ID="Text17">
					</td>
					<td>
						<input size="8" type="text" name="dtCheck<%=intCount2%>" value="<% = rsLI("dtCheck")%>" maxlength="16" onchange="jfChangedLI('<%=intCount2%>');" ID="Text4">
					</td>
					<td>
						<input size="20" type="text" name="szPayee<%=intCount2%>" value="<% = rsLI("szPayee")%>" maxlength="128" onchange="jfChangedLI('<%=intCount2%>');" ID="Text5">
					</td>
					<td>
						<input size="8" type="text" name="dtReciept<%=intCount2%>" value="<% = rsLI("dtReciept")%>" maxlength="16" onchange="jfChangedLI('<%=intCount2%>');" ID="Text6">
					</td>
					<td>
						<textarea cols=20 rows=1 wrap=virtual name="szLine_Item_desc<% = intCount2%>" onChange="jfChangedLI('<%=intCount2%>');" onfocus="this.rows=4;" onblur="this.rows=1;" onKeyDown="jfMaxSize(511,this);" ID="Textarea2"><% = rsLI("szLine_Item_desc") %></textarea>
					</td>
					<td>
						<input size="6" type="text" name="curUnit_Price<%=intCount2%>" value="<% = rsLI("curUnit_Price")%>" maxlength="10" onchange="jfChangedLI('<%=intCount2%>');jfCalcLineItem('<%=dblBudgetAmount%>','<%=intCount%>','<%=intCount2%>');" ID="Text7">
					</td>
					<td>
						<input size="2" type="text" name="intQuantity<%=intCount2%>" value="<% = rsLI("intQuantity")%>" maxlength="3" onchange="jfChangedLI('<%=intCount2%>');jfCalcLineItem('<%=dblBudgetAmount%>','<%=intCount%>','<%=intCount2%>');" ID="Text8">
					</td>
					<td>
						<input size="6" type="text" name="curShipping<%=intCount2%>" value="<% = rsLI("curShipping")%>" maxlength="10" onchange="jfChangedLI('<%=intCount2%>');jfCalcLineItem('<%=dblBudgetAmount%>','<%=intCount%>','<%=intCount2%>');" ID="Text9">
					</td>
					<td>
						<input type="text" disabled size="8" value="<% = oHtml.IIF(isNumeric(rsLI("total")),formatnumber(rsLI("total"),2),"")%>" name="Total<%=intCount2%>" ID="Text10">
					</td>
				</tr>						
						<%
						rsLI.MoveNext
						intCount2 = intCount2 + 1
					loop
				else
					intCount2 = intCount2 + 1
				end if
				rsLi.Close
				strLIList = strLIList & intCount2
				%>
				<tr>
					<input type="hidden" name="liGroupList<%=intCount%>" value="<% = strLIList %>" ID="Hidden9">
					<td>
						<input size="6" type="text" name="szPO_Number<%=intCount2%>" value="" maxlength="16" onchange="jfChangedLI('<%=intCount2%>');" ID="Text11">
					</td>
					<td>
						<input size="6" type="text" name="szInvoice_Number<%=intCount2%>" value="" maxlength="16" onchange="jfChangedLI('<%=intCount2%>');" ID="Text12">
					</td>
					<td>
						<input size="6" type="text" name="szCheck_Number<%=intCount2%>" value="" maxlength="16" onchange="jfChangedLI('<%=intCount2%>');" ID="Text18">
					</td>
					<td>
						<input size="8" type="text" name="dtCheck<%=intCount2%>" value="" maxlength="16" onchange="jfChangedLI('<%=intCount2%>');" ID="Text13">
					</td>
					<td>
						<input size="20" type="text" name="szPayee<%=intCount2%>" value="" maxlength="128" onchange="jfChangedLI('<%=intCount2%>');" ID="Text14">
					</td>
					<td>
						<input size="8" type="text" name="dtReciept<%=intCount2%>" value="" maxlength="16" onchange="jfChangedLI('<%=intCount2%>');" ID="Text15">
					</td>
					<td>
						<textarea cols=20 rows=1 wrap=virtual name="szLine_Item_desc<% = intCount2%>" onChange="jfChangedLI('<%=intCount2%>');" onfocus="this.rows=4;" onblur="this.rows=1;" onKeyDown="jfMaxSize(511,this);" ID="Textarea3"></textarea>
					</td>
					<td>
						<input size="6" type="text" name="curUnit_Price<%=intCount2%>" value="" maxlength="10" onchange="jfChangedLI('<%=intCount2%>');jfCalcLineItem('<%=dblBudgetAmount%>','<%=intCount%>','<%=intCount2%>');" ID="Text16">
					</td>
					<td>
						<input size="2" type="text" name="intQuantity<%=intCount2%>" value="" maxlength="3" onchange="jfChangedLI('<%=intCount2%>');jfCalcLineItem('<%=dblBudgetAmount%>','<%=intCount%>','<%=intCount2%>');" ID="Text19">
					</td>
					<td>
						<input size="6" type="text" name="curShipping<%=intCount2%>" value="" maxlength="10" onchange="jfChangedLI('<%=intCount2%>');jfCalcLineItem('<%=dblBudgetAmount%>','<%=intCount%>','<%=intCount2%>');" ID="Text20">
					</td>
					<td>
						<input type="text" size="8" disabled name="Total<%=intCount2%>" ID="Text21">
					</td>
				</tr>
				<INPUT type="hidden" name="intOrdered_Item_IDLI<%=intCount2%>" value="<%=rsGetItems("intOrdered_Item_ID")%>" ID="Hidden11">
			</table>
		</td>
	</tr>
<%
		rsGetItems.MoveNext
		intCount = intCount + 1
		intCount2 = intCount2 + 1
		strLIList =""
	loop
%>
	<tr>
		<td colspan=9>
		<input type=hidden name="intCount" value="<%=intCount%>" ID="Hidden10">
		<input type=hidden name="intCount2" value="<%=intCount2%>" ID="Hidden12">
		<input type=button value="Save and Requery Data" class="NavSave" onClick="SetPageScroll('<% = intPageNum %>GSA');this.form.submit();" NAME="Button1" ID="Button4">
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
<BR><BR>
<table ID="Table8">
	<tr>
		<td>
			<center><font face=arial size=3><b>No Goods/Services were found under the search criteria.</b></font></center>
		</td>
	</tr>
</form>
</table>
<%
end if

rsGetItems.Close
set rsGetItems = nothing
set rsLI = nothing

call oFunc.CloseCN()
set oFunc = nothing
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")

sub vbsSaveLineItems
	arList = split(request.Form("LineItemsChanged"),",")
	dim cmd
	dim dtCheck
	dim dtReciept
	dim dblUnitPrice
	dim intQty
	dim dblShipping
	for i = 0 to ubound(arList)
		if arList(i) <> "" then
			dtCheck = oHtml.IIF(isDate(request.Form("dtCheck"&arList(i))),request.Form("dtCheck"&arList(i)),"NULL")
			dtReciept = oHtml.IIF(isDate(request.Form("dtReciept"&arList(i))),request.Form("dtReciept"&arList(i)),"NULL")
			
			IF isDate(dtCheck) then dtCheck = "'"& formatdateTime(dtCheck,2) & "'"
			IF isDate(dtReciept) then dtReciept = "'"& formatdateTime(dtReciept,2) & "'"
			
			if isnumeric(oFunc.EscapeTick(request.Form("curUnit_Price"&arList(i)))) then
				dblUnitPrice = oFunc.EscapeTick(request.Form("curUnit_Price"&arList(i)))
			else
				dblUnitPrice = "0"
			end if
			
			if isnumeric(oFunc.EscapeTick(request.Form("intQuantity"&arList(i)))) then
				intQty = oFunc.EscapeTick(request.Form("intQuantity"&arList(i)))
			else
				intQty = "0"
			end if
			
			if isnumeric(oFunc.EscapeTick(request.Form("curShipping"&arList(i)))) then
				dblShipping = oFunc.EscapeTick(request.Form("curShipping"&arList(i)))
			else
				dblShipping = "0"
			end if
			
			if request.Form("intLine_Item_ID" & arList(i)) <> "" then
				cmd = "UPDATE    tblLine_Items " & _ 
						"SET szLine_Item_desc ='" & oFunc.EscapeTick(request.Form("szLine_Item_desc"&arList(i))) & "', " & _ 
						"intQuantity =" & intQty & ", " & _ 
						"szPO_Number ='" & oFunc.EscapeTick(request.Form("szPO_Number"&arList(i))) & "',  " & _ 
						"szInvoice_Number ='" & oFunc.EscapeTick(request.Form("szInvoice_Number"&arList(i))) & "',  " & _ 
						"szCheck_Number ='" & oFunc.EscapeTick(request.Form("szCheck_Number"&arList(i))) & "',  " & _ 
						"szPayee ='" & oFunc.EscapeTick(request.Form("szPayee"&arList(i))) & "',  " & _ 
						"dtCheck =" & dtCheck & ",  " & _ 
						"dtReciept =" & dtReciept & ",  " & _ 
						"szUSER_MODIFY ='" & oFunc.EscapeTick(session.Contents("strUserID")) & "'," & _ 
						"curShipping =" & dblShipping & ",  " & _ 
						"curUnit_Price =" & dblUnitPrice & ",  " & _ 
						"dtMODIFY = CURRENT_TIMESTAMP " & _ 
						"WHERE     intLine_Item_ID = " & request.Form("intLine_Item_ID" & arList(i))
			else
				cmd = "INSERT INTO tblLine_Items " & _ 
						" (intOrdered_Item_ID, dtLine_Item, szLine_Item_desc, curUnit_Price, " & _ 
						" intQuantity, curShipping, szPO_Number, szInvoice_Number, szCheck_Number,  " & _ 
						" szPayee, dtCheck, dtReciept, dtCREATE, szUSER_CREATE) " & _ 
						"VALUES     (" & request.Form("intOrdered_Item_IDLI" & arList(i)) & "," & _
						"CURRENT_TIMESTAMP," & _
						"'" & oFunc.EscapeTick(request.Form("szLine_Item_desc"&arList(i))) & "'," & _
						dblUnitPrice & "," & _
						intQty & "," & _
						dblShipping & "," & _
						"'" & oFunc.EscapeTick(request.Form("szPO_Number"&arList(i))) & "'," & _
						"'" & oFunc.EscapeTick(request.Form("szInvoice_Number"&arList(i))) & "'," & _
						"'" & oFunc.EscapeTick(request.Form("szCheck_Number"&arList(i))) & "'," & _
						"'" & oFunc.EscapeTick(request.Form("szPayee"&arList(i))) & "'," & _
						dtCheck & "," & _
						dtReciept & "," & _
						"CURRENT_TIMESTAMP," & _
						"'" & oFunc.EscapeTick(session.Contents("strUserID")) & "') "
			end if
			
			oFunc.ExecuteCN(cmd)
		end if
	next
end sub

sub vbsSaveChanges
	' This sub cycles through all of the rows from the form and picks out
	' the approved items sets them to approved in tblOrdered_Items and 
	' takes the denied items and sets them to denied and saves the reason.
	' If niether approved or denied there is no action.
	dim update
	dim bolClosed
	arList = split(request("updatelist"),",")
	for i = 0 to ubound(arList)
		if arList(i) <> "" then
			bolClosed = oHtml.IIF(request.Form("bolClosed"&arList(i)) <> "","1","NULL")
			if request.Form("Approved"&arList(i)) = "1" then
				' approved budget
				update = "update tblOrdered_Items set " & strApprovedField & " = 1, " & _
						 "szDeny_Reason = '" & oFunc.escapeTick(request.Form("denied"&arList(i))) & "', " & _
						 "dtApproval_Changed = CURRENT_TIMESTAMP " & _
						 ",dtModify = CURRENT_TIMESTAMP, " & _
						 "bolClosed = " & bolClosed & ", " & _
						 "szUSER_MODIFY ='" & oFunc.EscapeTick(session.Contents("strUserID")) & "' " & _ 
						 "where intOrdered_Item_ID = " & request.Form("intOrdered_Item_ID"&arList(i))							
			elseif request.Form("Approved"&arList(i)) = "0" then
				' rejected budget
				update = "update tblOrdered_Items set " & strApprovedField & " = 0, " & _
						"szDeny_Reason = '" & oFunc.escapeTick(request.Form("denied"&arList(i))) & "', " & _
						"dtApproval_Changed = CURRENT_TIMESTAMP," & _
						"bolClosed = 1, " & _
						"szUSER_MODIFY ='" & oFunc.EscapeTick(session.Contents("strUserID")) & "'," & _ 
						"dtModify = CURRENT_TIMESTAMP " & _
						"where intOrdered_Item_ID = " & request.Form("intOrdered_Item_ID"&arList(i)) 
				' no need to have line items since the budget is rejected so do some clean
				' up to insure this is true				
				delete = "delete from tblLine_Items where intOrdered_Item_ID = " & request.Form("intOrdered_Item_ID"&arList(i)) 

				oFunc.ExecuteCN(delete)
				
			elseif request.Form("Approved"&arList(i)) = "2" then
				update = "update tblOrdered_Items set " & strApprovedField & " = NULL, " & _
						 "szDeny_Reason = '" & oFunc.escapeTick(request.Form("denied"&arList(i))) & "', " & _
						 "dtApproval_Changed = CURRENT_TIMESTAMP, " & _
						 "bolClosed = " & bolClosed & ", " & _
						 "szUSER_MODIFY ='" & oFunc.EscapeTick(session.Contents("strUserID")) & "'" & _ 
						 ",dtModify = CURRENT_TIMESTAMP " & _
						 "where intOrdered_Item_ID = " & request.Form("intOrdered_Item_ID"&arList(i)) 
			end if 
			oFunc.ExecuteCN(update)
		end if		
	next
end sub

sub PrintHeader
%>
	<tr class="NavyWhite8">
		<td align=center>
			Students Name
		</td>
		<td align=center>
			Vendor Name
		</td>
		<td align=center>
			<input type="button" value="Print" class="btSmallWhite" onclick="jfPrintGS();" ID="Button5" NAME="Button3">
		</td>
		<td align=center>
			Item #
		</td>		
		<td align=center>
			Links
		</td>
		<td align=center title="Unit Price">
			User
		</td>
		<td align=center>
			Budget
		</td>
		<td align=center>
			Balance
		</td>
		<td align=center>
			Closed
		</td>
		<td align=center title="Shows if the ILP was Approved">
			Course<BR>Status
		</td>
		<td align=center title="Approve Item">
			Aprv
		</td>
		<td align=center>
			Comments/Reject Reason
			<input type=button value="Save" class="NavSave" onClick="SetPageScroll('<% = intPageNum %>GSA');this.form.submit();" NAME="Button1" ID="Button6">
		</td>
	</tr>
<%
end sub
%>
<script language="javascript">
	RestoreScroll('<% = intPageNum %>GSA');
</script>