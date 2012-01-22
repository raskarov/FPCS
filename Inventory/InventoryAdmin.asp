<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		InventoryAdmin.asp
'Purpose:	Management screens for inputing and editing Inventory Items
'Date:		22 March 2006
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Page Header Setup
Session.Value("strTitle") = "Inventory Control Panel"
Session.Value("strLastUpdate") = "22 March 2006"

if request("simpleHeader") <> "" then 
	Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")
else
	Server.Execute(Application.Value("strWebRoot") & "includes/Header.asp")
end if

' Format incoming request variables into VBS variables for ease of use
if Request.Form.Count > 0 then
	set objRequest = Request.Form
else
	set objRequest = Request.QueryString
end if

for each item in objRequest
	if item & "" <> "" then
		execute("dim " & item)
		sValue = objRequest(item)
		if sValue & "" <> "" then sValue = replace(sValue,"""",""):sValue = replace(sValue,"'","''")
		
		execute(item & " = """ & sValue & """")
	end if
next 

' helps us ensure that a record is only inserted once to prevent muti-submit misshaps
if Session("SessionTracker") & "" = "" or  request("SaveOnce") <> "" then
	Session("SessionTracker") = Session("SessionTracker") + 1
end if

' Script Setup
dim cssNew, cssSearch, scriptToCall, oHtml, TitleText, HeaderText, RequiredCss, IsSearch, oFunc, StudentReissueCost
dim ReIssueComments, PrintMode, IsFundsAvailable, totalPaidBack, IsHeld

set oHtml = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/htmlFuncs.wsc"))

StudentReissueCost = "$0.00"
'Create object containing all of our FPCS functions
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

if panel & ""  = "" and not oFunc.IsAdmin then panel = "SEARCH"

PrintMode = false
IsFundsAvailable = true
IsHeld = false

if not oFunc.IsAdmin then PrintMode = true

if oFunc.IsAdmin and IsEdit & "" = "" and InventoryDetailID & "" <> "" then
	PrintMode = true
end if

' Determine Panel to Show
select case ucase(panel)
	case "","NEW"
		cssNew = "inventoryOptionSelected"
		cssSearch = "inventoryOption"
		
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		''' Handle Insert and updates
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		if request("SaveOnce") <> "" and InventoryDetailID & "" = "" _
			and  session.Contents("SessionTracker") - SessionTracker < 2  and oFunc.IsAdmin  then 
			call InsertDetailRecord
			if intStudent_ID <> "" and IlpID <> "" and intOrdered_Item_ID <> "" then
				HeldForStudentID = request("intStudent_ID")
				HeldWithOrdID = true
				oiCost = TotalCostNew
				'HeldForIlpID = request("IlpId")
				'response.Write "INSERT - " & HeldForStudentID & " - " & HeldForIlpID
			end if 
		elseif request("SaveOnce") <> "" and InventoryDetailID & "" <> ""  and oFunc.IsAdmin then 
			call UpdateDetailRecord
		elseif InventoryDetailID & "" <> "" then
					
			sql = "SELECT     id.InventoryCategoryID, id.Title, id.Description, id.DistrictControlNum, id.SchoolControlNum, id.TotalCostNew, id.DatePurchased, id.PONumber,  " & _
                  "    id.InventoryDetailTypeID, id.VendorID, id.SerialNumber, id.OtherRefNum, id.Edition, id.Author, id.ISBN, id.Publisher, id.PublishDate, id.WarrantyInfo, " & _ 
                  "    id.GeneralComments, id.InventoryStatusCD, id.Manufacturer, id.ModelNumber, id.Location, id.DateCreated, id.DateModified, id.UserCreated, " & _ 
                  "    id.UserModified, id.HeldForStudentID, id.HeldForIlpID, id.DateHoldEnd, id.myCount, id.HeldWithOrdID,  " & _
                  "    oi.intQty * oi.curUnit_Price + oi.curShipping AS oiCost " & _
				  "	FROM         INVENTORY_DETAILS AS id LEFT OUTER JOIN " & _
				  "						tblOrdered_Items AS oi ON id.CreatedUsingOrdItemId = oi.intOrdered_Item_ID " & _
				  "	WHERE     (id.InventoryDetailID = " & oFunc.EscapeTick(InventoryDetailID) & ")"
				  
			dim rsSQL
			set rsSQL = server.CreateObject("ADODB.RECORDSET")
			'response.Write sql 
			rsSQL.CursorLocation = 3
			rsSQL.Open sql,oFunc.FPCScnn
			
			for each field in rsSQL.Fields
				'response.Write field.name & "<BR>"
				if field & "" <> "" then
					execute("dim " & field.name)
					execute(field.name &  " = """ & field & """")
				end if 
			next 
			
			rsSQL.Close
			set rsSQL = nothing
			
			if IsTemplate & "" <> "" and oFunc.IsAdmin then
				InventoryDetailID = ""
				PrintMode = false
			end if
		end if
	
		if ReissueCost& "" <> ""  and isDate(ReissueDate&"") & InventoryDetailID & "" <> "" _
			and session.Contents("SessionTracker") - SessionTracker < 2  and oFunc.IsAdmin then
			call InsertReissueRecord
		end if
		
		if len(ChangedIssueCosts) > 1 and oFunc.IsAdmin  then
			call UpdateReissueRecords
		end if
		
		If InventoryDetailID & "" <> "" and HoldStudentID & "" <> "" & HeldForIlpID & "" <> "" and SaveHold & "" <> "" then
			if isNumeric(HoldStudentID&"") then
				call PlaceHold(HoldStudentID, HeldForIlpID, InventoryDetailID)
			end if
		end if
		
		' Checkout Item for a Student
		if InventoryDetailID & "" <> "" and SelectedStudent_ID & "" <> "" _
			and isDate(DateCheckedOut) and oFunc.IsAdmin and CheckOut & "" <> "" then
				call CheckOutItem(InventoryDetailID,SelectedStudent_ID)
		end if
		
		'Check in item for a student
		if CheckedOutInventoryID & "" <> "" and InventoryDetailId & "" <> "" and oFunc.IsAdmin then
			call CheckInItem(CheckedOutInventoryID,InventoryDetailId,Comments)
		end if
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'''	Handle populating inventory item form using Ordered Item
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		if request.QueryString("intOrdered_Item_ID") & "" <> "" and InventoryCategoryID & "" = "" then
			dim rsOI 
			set rsOI = server.CreateObject("ADODB.Recordset")
			rsOI.CursorLocation = 3
			' Check to see if item has already been added to the inventory
			sql = "SELECT     InventoryDetailID " & _ 
					"FROM         INVENTORY_DETAILS " & _ 
					"WHERE     (CreatedUsingOrdItemId = " & intOrdered_Item_ID & ") "
			rsOI.Open sql, oFunc.FPCScnn
		
			if rsOI.RecordCount > 0 then
				' Item is already in inventory so redirect to show inventory record
				InventoryDetailID = rsOI("InventoryDetailID")
				rsOI.Close
				set rsOI = nothing
				response.Redirect("./InventoryAdmin.asp?InventoryDetailID=" & InventoryDetailID & "&simpleHeader=" & simpleHeader)
			else
				' Item is not in inventory so get data to fill in inventory form
				rsOI.close
				sql = "SELECT	oi.intVendor_ID, oi.intStudent_ID, oi.intQty * oi.curUnit_Price + oi.curShipping AS total,  " & _ 
						"	oa.szValue, ia.InventoryFieldName, i.InventoryCategoryID, oi.intILP_ID " & _ 
						"FROM	tblOrdered_Items AS oi INNER JOIN " & _ 
						"	tblOrd_Attrib AS oa ON oi.intOrdered_Item_ID = oa.intOrdered_Item_ID INNER JOIN " & _ 
						"	trefItem_Attrib AS ia ON oa.intItem_Attrib_ID = ia.intItem_Attrib_ID INNER JOIN " & _ 
						"	trefItems AS i ON oi.intItem_ID = i.intItem_ID " & _ 
						"WHERE	(oi.intOrdered_Item_ID = " & intOrdered_Item_ID & ") "
				rsOI.Open sql, oFunc.FPCScnn
						
				if rsOI.RecordCount > 0 then
					do while not rsOI.EOF
						if rsOI("InventoryFieldName") & "" <> "" then
							execute("dim " & rsOI("InventoryFieldName"))
							if rsOI("szValue") & "" <> "" then
								myValue = replace(rsOI("szValue"),"""","'")
							else
								myValue = ""
							end if
							
							execute(rsOI("InventoryFieldName") & " = """ & myValue & """")							
						end if
						rsOI.MoveNext
					loop
					rsOI.MoveFirst
					dim VendorID, intStudentID, TotalCostNew, InventoryCategoryID
					VendorID = rsOI("intVendor_ID")
					intStudent_ID = rsOI("intStudent_ID")
					TotalCostNew = rsOI("total")
					InventoryCategoryID = rsOI("InventoryCategoryID")
					InventoryStatusCD = "AV"
					intILP_ID = rsOI("intILP_ID")
					rsOI.Close
					set rsOI = nothing
					
					InventoryStatusCD = "OH"	
					isPlaceOnHold = "TRUE"				
				else 
					response.Write "<h1>Error:  Invalid Budget ID</h1>"
					rsOI.Close
					set rsOI = nothing
					response.End				
				end if
			end if
			
		end if		
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		
		If IsNumeric(InventoryDetailID&"") Then
			If not PrintMode then
				TitleText = TitleText & "EDIT Existing Inventory Item"
			else
				TitleText = TitleText & "View Existing Inventory Item"
			end if
		Else
			TitleText = "Add a New Inventory Item"
		End If
		
		if InventoryStatusCD & "" = "" then
			InventoryStatusCD = "AV"
		end if
		
		HeaderText = "Required Fields"		
		RequiredCss = "InventoryRequired"
		IsSearch = false
		
		call PrintJsScripts
		PanelToCall = "SingleItemPanel"
	case "SEARCH"
		cssSearch = "inventoryOptionSelected"
		cssNew = "inventoryOption"
				
		TitleText = "Inventory Search"
		HeaderText = "Searchable Fields"
		RequiredCss = "InventoryNotRequired"
		IsSearch = true				
		
		call PrintJsScripts
		PanelToCall = "SingleItemPanel"
end select

' Initial HTML for Form
%>

<form name="main" action="InventoryAdmin.asp" method="post">
<input type="hidden" name="Locks" value="">
<input type="hidden" name="SaveOnce" id="SaveOnce">
<input type="hidden" name="SearchOnce" id="SearchOnce">
<input type="hidden" name="simpleHeader" value="<% = simpleHeader %>">
<input type="hidden" name="intOrdered_Item_ID" value="<% = intOrdered_Item_ID %>">
<input type="hidden" name="panel" value="<% = request("panel") %>">
<input type="hidden" name="SessionTracker" value="<% = session.Contents("SessionTracker") %>">
<input type="hidden" name="IsEdit" value="<% = IsEdit %>">
<input type="hidden" name="isPlaceOnHold" value="<% =  isPlaceOnHold %>">
<input type="hidden" name="refreshParent" value="<% = refreshParent %>">
<% if intOrdered_Item_ID & "" <> "" and intILP_ID <> "" then %>
<input type="hidden" name="IlpID" id="IlpId" value="<% = intILP_ID %>">
<input type="hidden" name="intStudent_ID" id="intStudent_ID" value="<%=intStudent_ID%>">
<% end if %>
<table style="width:100%;" ID="Table3">
    <tr>
        <td style="width:100%;">
            <table style="width:100%;" cellpadding=0 cellspacing=0 style="border-bottom: black 1px solid;" ID="Table4">
				<tr>
					<td class="inventoryMain" nowrap>
						Inventory Control Panel	
					</td>
					<td style="width:5px;">
						&nbsp;
					</td>
					<% if oFunc.IsAdmin then %>
					<td class="<% = cssNew %>" nowrap>
						<a href="InventoryAdmin.asp?panel=new" class="White8Verd">Add New Item</a>
					</td>
					<% end if %>
					<td>
						&nbsp;
					</td>
					<% if request("simpleHeader") = "" then %>
					<td class="<% = cssSearch %>" nowrap>
						<a href="InventoryAdmin.asp?panel=search" class="White8Verd">Search For Item</a>
					</td>	
					<td>
						&nbsp;
					</td>				
					<td class="inventoryOption" nowrap>
						<a href="InventoryLists.asp?" class="White8Verd">Inventory Lists</a>
					</td>
					<% end if %>
					<td style="width:100%;">
						&nbsp;
					</td>
				</tr>
			</table>
		</td>
    </tr>
    <tr>
        <td style="width:100%; background-color:#eaeaea;">
<%
			' Load the correct panel
			execute("call " & PanelToCall)
%>
		</td>
	</tr>
</table>
<% = oHtml.ToolTipDivs %>
</form>
<DIV ID="divCal" STYLE="position:absolute;visibility:hidden;background-color:white;layer-background-color:white;"></DIV>
</body>
</html>
<%
set oHtml = nothing
call oFunc.CloseCN()
set oFunc = nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Procedures below here												 '''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
function PrintJsScripts
%>
<script language=javascript src="<%= Application.Value("strWebRoot") %>includes/CalendarPopup.js"></script>	
<script language="javascript">
	function jfGetForm(obj){
		if (obj.value != "") {
			document.main.submit();
		}else{
			alert("You must select a Category");
		}
	}
	
	var cal = new CalendarPopup('divCal');
	cal.showNavigationDropdowns();
	
	var cal2 = new CalendarPopup('divCal');
	cal2.showNavigationDropdowns();		
	
	var cal3 = new CalendarPopup('divCal');
	cal3.showNavigationDropdowns();	
	
	var cal4 = new CalendarPopup('divCal');
	cal4.showNavigationDropdowns();	
	
	function jfInventoryValidate(){
		// Add static required field validaors
		var oForm = document.forms[0];		
		var sRequired = document.getElementById('sRequiredFields');
		var aRequired = sRequired.value.split(",");
		var aRLabels = document.getElementById('sRequiredFieldsLabels').value.split(",");
		var i;
		var obj;
		var sError = "";
		
		for (i=0;i <= aRequired.length-1;i++) {	
			if (aRequired[i] != "" && aRequired[i] != "undefined") {
				obj = document.getElementById(aRequired[i]);
				if (obj.value == "undefined" || obj.value == ""){
					sError += aRLabels[i] + " cannot be blank.\n";
				}
			}
		}
		
		sError += isValidCost(document.forms[0].TotalCostNew,"Total Cost New");
		
		<% if InventoryDetailID & "" = "" and intOrdered_Item_ID & "" <> "" then%>
			
		<% else %>
		if (((document.forms[0].ReissueCost.value == "" || document.forms[0].ReissueCost.value == "undefined") &&
			(document.forms[0].ReissueDate.value != "" && document.forms[0].ReissueDate.value != "undefined")) ||
			((document.forms[0].ReissueCost.value != "" && document.forms[0].ReissueCost.value != "undefined") &&
			(document.forms[0].ReissueDate.value == "" || document.forms[0].ReissueDate.value == "undefined"))){
			sError += "Reissue Date and Cost must be provided.\n";						
		}
		
		if (document.forms[0].ReissueCost != "" && document.forms[0].ReissueCost != "undefined"){
			sError += isValidCost(document.forms[0].ReissueCost,"Reissue Cost");
		}
		<% end if %>
			
		if (sError.length > 0){
			alert("Cannot save until the following problems are fixed ...\n" + sError);			
		}else{
			var oCheck = document.getElementById('SaveOnce');
			if (oCheck.value == "" || oCheck.value == "undefined"){
				oCheck.value = "saved";
				document.main.submit();
			}
		}			
	}
	
	function isValidCost(pObj, pLabel){
		var myError = "";
		if (pObj.value != "" && pObj.value != 'undefined'){
			var iCost = pObj.value;
			iCost = iCost.replace("$","");
			iCost = parseFloat(iCost);
			if(isNaN(iCost) == true){
				myError = "'" + pLabel + "' must be a valid number.\n";
			}else if(iCost < .01){
				myError = "'" + pLabel + "' must be greater than 0.\n";
			}	
			if (myError == "") { pObj.value = iCost;}	
		}
		return myError;
	}
	
	function jfValidateHold() {
		var oForm = document.forms[0];
		var sError = "";
		
		if (oForm.HoldStudentID.value == "" || oForm.HoldStudentID.value == "undefined") {
			sError += "You must select a Student in order to place a hold.\n";
		}
		
		if (oForm.HeldForIlpID.value == "" || oForm.HeldForIlpID.value == "undefined") {
			sError += "You must select a course in order to place a hold.\n";
		}
		
		if (sError.length > 0){
			alert("Cannot save until the following problems are fixed ...\n" + sError);			
		}else{
			var oCheck = document.getElementById('SaveHold');
			if (oCheck.value == "" || oCheck.value == "undefined"){
				oCheck.value = "saved";
				document.main.submit();
			}
		}			
	}
	
	function jfValidateCheckOut() {
		var oForm = document.forms[0];
		var sError = "";
		
		if (oForm.SelectedStudent_ID.value == "" || oForm.SelectedStudent_ID.value == "undefined") {
			sError += "You must select a Student in order to check an item out.\n";
		}
		
		if (oForm.IlpID.value == "" || oForm.IlpID.value == "undefined") {
			sError += "You must select a course in order to check an item out.\n";
		}
		
		if (oForm.DateCheckedOut.value == "" || oForm.DateCheckedOut.value == "undefined") {
			sError += "You must provide a vaild Checkout Date.\n";
		}
		
		if (sError.length > 0){
			alert("Cannot check an item out until the following problems are fixed ...\n" + sError);			
		}else{
			var oCheck = document.getElementById('CheckOut');
			if (oCheck.value == "" || oCheck.value == "undefined"){
				oCheck.value = "saved";
				document.main.submit();
			}
		}			
	}
	
	function jfViewItem(pDetailID){
		var winInventory;
		var url;
		url = "<%=Application.Value("strSSLWebRoot")%>Inventory/InventoryAdmin.asp?refreshParent=true&simpleHeader=true&InventoryDetailID=" + pDetailID + "&panel=new";
		winInventory = window.open(url,"winInventory","width=950,height=500,scrollbars=yes,resizable=yes");
		winInventory.moveTo(0,0);
		winInventory.focus();
	}

	function jfTryAutoFill(){
        	var url = "<% = Application("strAutoFill") %>?URL=HTTPS://<% = request.servervariables("server_name") & request.servervariables("URL") & "?panel=new||InventoryCategoryID=" & request("InventoryCategoryID") & "&" & request.servervariables("Query_String") %>";
        	var winFill = window.open(url,"winFill","width=900,height=500,scrollbars=yes,resizable=yes");
        	//winFill.moveTo(0,0);
        	//winFill.focus();
    	}
</script>
<%
end function

function SingleItemPanel
	dim  rs, sqlVendor
	
	
   
%>
<input type="hidden" name="InventoryDetailID" value="<% = InventoryDetailID %>" ID="InventoryDetailID">
<table style="width:100%;" ID="Table1">   
    <tr>
        <td class="svplain9" colspan="100">
			<table cellpadding="0" cellspacing="0" style="width:100%;">
				<tr  class="svplain9">
					<td valign="top" >
						<B><% = TitleText %> </B>
					</td>
					<td align="right">
						<% if oFunc.IsAdmin then 
								if PrintMode and InventoryDetailID <> "" then 
						%>
							<a href="#"  onclick="window.location.href='./InventoryAdmin.asp?simpleHeader=<%=simpleHeader%>&InventoryDetailID=<%=InventoryDetailID%>&IsTemplate=true&IsEdit=true';">Use as Template</a> | 
							<a href="#" onclick="window.location.href='./InventoryAdmin.asp?simpleHeader=<%=simpleHeader%>&InventoryDetailID=<%=InventoryDetailID%>&IsEdit=true';">Edit</a>
						<%
								end if
							end if
						%>
					</td>
				</tr>
			</table>
            
        </td>
    </tr>
    <tr>
        <td class="<% if IsSearch then response.Write "TableHeaderBlue" else response.Write "TableHeaderRed"%>" colspan="100"  style="width:100%;">
            &nbsp;<b><% = HeaderText %></b>
        </td>
    </tr>
    <tr>
        <td style="width:95%;">
            <table ID="Table2" cellpadding=3>
                <tr>
                    <td class="<% =RequiredCss %>" style="width:0%;">
                        Category:
                    </td>
                    <td style="<% if InventoryCategoryID & "" <> "" then response.Write "width:33%;" else response.Write "width:150px;"%>" class="TableCell">
						<% 
							sql = "SELECT	InventoryCategoryID, Name " & _ 
									"FROM	INVENTORY_CATEGORIES " & _ 
									"WHERE	(IsActive = 1) " & _ 
									"ORDER BY Name "	
							 CatList =  oFunc.MakeListSQL(sql,"InventoryCategoryID","Name",InventoryCategoryID)	
						
						 if not PrintMode or IsSearch then %>
						<select name="InventoryCategoryID" class="InventorySelect" ID="InventoryCategoryID" onchange="jfGetForm(this);" style="width:100%;">
							<option value=""></option>
							<%
								response.Write CatList									
							%>
						</select>
						<% else %>
							<input type="hidden" value="<% = InventoryCategoryID %>" name="InventoryCategoryID" id="InventoryCategoryID">
							<% = oFunc.SelectedListText %>
						<% end if %>
					</td>
					<% if InventoryCategoryID <> "" then %>           
                    <td class="<% =RequiredCss %>" style="width:0%;">
                       <nobr>Vendor:</nobr>
                    </td>
                    <td style="width:33%;" class="TableCell">
						<% 
							if not IsSearch then
								sqlVendor = "SELECT     intVendor_ID, szVendor_Name + ': ' + vStatus AS VendorName " & _ 
											"FROM         (SELECT     v.intVendor_ID, v.szVendor_Name, " & _ 
											"                                                  (SELECT     TOP 1 upper(szVendor_Status_CD) " & _ 
											"                                                    FROM          tblVendor_Status vs " & _ 
											"                                                    WHERE      vs.intVendor_ID = v.intVendor_ID AND vs.intSchool_Year <= " & session.Contents("intSchool_Year") & _ 
											"                                                    ORDER BY intSchool_Year DESC, intVendor_Status_ID DESC) AS vStatus " & _ 
											"                       FROM          tblVendors v " & _ 
											"                       WHERE      (SELECT     TOP 1 upper(szVendor_Status_CD) " & _ 
											"                                               FROM          tblVendor_Status vs " & _ 
											"                                               WHERE      vs.intVendor_ID = v.intVendor_ID AND v.bolGoods_Vendor = 1 and vs.intSchool_Year <= " & session.Contents("intSchool_Year")  & _ 
											"                                               ORDER BY intSchool_Year DESC, intVendor_Status_ID DESC) IN ('APPR', 'PEND')) DERIVEDTBL " & _ 
											"ORDER BY szVendor_Name "
							else	
								if not oFunc.IsAdmin then addWhere = " AND (INVENTORY_DETAILS.InventoryStatusCD = 'AV') "
								
								sqlVendor = "SELECT DISTINCT tblVendors.intVendor_ID, tblVendors.szVendor_Name as VendorName " & _ 
											"FROM	tblVendors INNER JOIN " & _ 
											"	INVENTORY_DETAILS ON tblVendors.intVendor_ID = INVENTORY_DETAILS.VendorID " & _ 
											"WHERE	(INVENTORY_DETAILS.InventoryCategoryID = " & InventoryCategoryID & ")  " & addWhere & _ 
											"ORDER BY tblVendors.szVendor_Name "
							end if 
										
							VendList = oFunc.MakeListSQL(sqlVendor,"intVendor_ID","VendorName",VendorID)
						
						if not PrintMode or IsSearch then
						%>
						<select name="VendorID" id="VendorID" class="InventorySelect" style="width:100%;">	
							<option value=""></option>
						<%
							response.Write VendList	
						%>
						</select>
						<% else %>	
							<% = oFunc.SelectedListText %>
						<% end if %>
					</td>	 
					<% if not IsSearch then %>
						<td class="<% =RequiredCss %>" style="width:0%;">
							<nobr>Total Cost New:</nobr>
						</td>
						<td style="width:33%;" class="TableCell">
							<% if not PrintMode then %>
							<nobr>$<input type=text value="<% = formatNumber(TotalCostNew,2) %>" class="InventorySelect" name="TotalCostNew" id="TotalCostNew" style="width:98%;"></nobr>
							<% else
									response.Write "$" & formatNumber(TotalCostNew,2)
							   end if
							%>
						</td>
                    <% else %>
						<td class="<% =RequiredCss %>" style="width:0%;">
							<nobr>Cost Between:</nobr>
						</td>
						<td style="width:33%;" class="svplain8">
							<nobr><input type=text value="<% = TotalCostNew %>" size="20" class="InventorySelect" name="TotalCostNew" id="TotalCostNew" style="width:40%;"> and
							<input type=text value="<% = TotalCostNew2 %>" size="20" class="InventorySelect" name="TotalCostNew2" id="TotalCostNew2" style="width:40%;"></nobr>
						</td>
                    <% end if %>
				</tr>				
				<tr>
					<td class="<% =RequiredCss %>">
						<nobr>FPCS Control #:</nobr>
					</td>
					<td class="TableCell">
						<% if not PrintMode or IsSearch then %>
						<input type="text" name="SchoolControlNum" id="SchoolControlNum" class="InventorySelect"  maxlength="50" style="width:100%;" value="<% if InventoryDetailID <> "" then response.Write SchoolControlNum else response.Write "" %>">
						<% else
								response.Write SchoolControlNum 
						   end if
						%>
					</td>
					<% if not IsSearch then %>	
						<td class="<% =RequiredCss %>" >
							<nobr>Date Purchased:</nobr>
						</td>
						<td class="TableCell">
							<% if not PrintMode then %>
							<input type="text" style="width:100%;" name="DatePurchased" value="<%if isDate(DatePurchased) then response.Write DatePurchased else response.write FormatDatetime(Now(),2)%>" ID="DatePurchased" class="InventorySelect" onclick="cal.select(document.forms[0].DatePurchased,'DatePurchased','M/dd/yyyy','<%if isDate(DatePurchased) then response.Write DatePurchased%>');return false;">
							<% else 
									response.Write DatePurchased
							   end if 
							%>
						</td>
					<% else %>
						<td class="<% =RequiredCss %>" >
							<nobr>Purchased Between:</nobr>
						</td>
						<td class="svplain8">
							<nobr><input type="text" size="10" name="DatePurchased" value="<%if isDate(DatePurchased) then response.Write DatePurchased%>" ID="DatePurchased" class="InventorySelect" onclick="cal3.select(document.forms[0].DatePurchased,'DatePurchased','M/dd/yyyy','<%if isDate(DatePurchased) then response.Write DatePurchased%>');return false;">
							and
							<input type="text" size="10" name="DatePurchased2" value="<%if isDate(DatePurchased2) then response.Write DatePurchased2%>" ID="DatePurchased2" class="InventorySelect"  onclick="cal4.select(document.forms[0].DatePurchased2,'DatePurchased2','M/dd/yyyy','<%if isDate(DatePurchased2) then response.Write DatePurchased2%>');return false;"></nobr>
						</td>
					<%
					end if
					
					call StatusList()
					response.Write "</tr>"
					
					set rs = server.CreateObject("ADODB.RECORDSET")
					rs.CursorLocation = 3
					
					' Get fields that make up the form
					sql = "SELECT Fields, FieldLabels, RequiredFields, RequiredFieldsLabels," & _
							" SearchResultFields, SearchResultFieldsLabels " & _ 
							"FROM INVENTORY_CATEGORIES " & _ 
							"WHERE (InventoryCategoryID = " & InventoryCategoryID & ") "
					rs.Open sql, oFunc.FPCScnn
					if rs.RecordCount > 0 then
						sFieldLabels = rs("FieldLabels")
						sFields = rs("Fields")			
						arFields = split(rs("Fields")& "",",")
						arFieldLabels = split(rs("FieldLabels")& "",",")
						
						arRequiredFields = split(rs("RequiredFields")& "",",")
						arRequiredFieldsLabels = split(rs("RequiredFieldsLabels")& "",",")
						sRequiredFields =  rs("RequiredFields") 
						sRequiredFieldsLabels = rs("RequiredFieldsLabels")
						
						if IsSearch then
							sSearchResultFields = rs("SearchResultFields")
							sSearchResultFieldsLabels = rs("SearchResultFieldsLabels")
						end if
					end if
					rs.Close
									
					if isArray(arRequiredFields) then
						ii = 3		
							for i = 0 to ubound(arRequiredFields)	
								if ucase(arRequiredFields(i)) <> "WARRANTYINFO" and _
									ucase(arRequiredFields(i)) <> "GENERALCOMMENTS" then
									execute("sValue = " & arRequiredFields(i))		
									if arRequiredFields(i) <> "InventoryDetailTypeID" then
						%>
							<td class="<% =RequiredCss %>">
								<nobr><% = arRequiredFieldsLabels(i) %>:</nobr>
							</td>		
							<td class="TableCell">
								<%if not PrintMode or IsSearch then 
									if (ucase(arRequiredFields(i)) = "PONUMBER" or ucase(arRequiredFields(i)) = "DISTRICTCONTROLNUM") _
										and InventoryDetailID & "" = "" then
											sValue = ""
									end if
									
								%>
								<input type="text" name="<% = arRequiredFields(i) %>" id="<% = arRequiredFields(i) %>" value="<% = sValue%>" maxlength="50" style="width:100%;"  class="InventorySelect">
								<% else
										response.Write sValue
								   end if 
								%>
							</td>
						<%		
									else
										call TypeList(InventoryCategoryID,InventoryDetailTypeID)															
									end if
									ii = ii + 1
									if ii mod 3 = 0 then response.Write "</tr><tr>"		
									sValue = ""
								else
									if instr(1,ucase(sRequiredFields),"WARRANTYINFO") > 0 then
										call AddStaticTextArea("WarrantyInfo",WarrantyInfo,"Warranty Info",False)
									end if
									
									if instr(1,ucase(sRequiredFields),"GENERALCOMMENTS") > 0 then
										call AddStaticTextArea("GeneralComments",GeneralComments,"General Comments",False)										
									end if
								end if
							next
						end if	
						
						if not IsSearch then
							sRequiredFields =  "InventoryCategoryID,VendorID,TotalCostNew,SchoolControlNum,DatePurchased,InventoryStatusCD," & sRequiredFields 
						else
							sRequiredFields =  "InventoryCategoryID,VendorID,TotalCostNew,TotalCostNew2,SchoolControlNum,DatePurchased,DatePurchased2,InventoryStatusCD," & sRequiredFields 
						end if
						
						sRequiredFieldsLabels = "Category,Vendor,Total Cost New,SI Control #,Date Purchased,Status," & sRequiredFieldsLabels
						
						if mid(sRequiredFields,len(sRequiredFields),1) = "," then
							sRequiredFields = left(sRequiredFields,len(sRequiredFields)-1)
						end if
						
						if mid(sRequiredFieldsLabels,len(sRequiredFieldsLabels),1) = "," then
							sRequiredFieldsLabels = left(sRequiredFieldsLabels,len(sRequiredFieldsLabels)-1)
						end if
						
					%>
					<input type="hidden" name="sRequiredFields" id="sRequiredFields" value="<% = sRequiredFields%>">
					<input type="hidden" name="sRequiredFieldsLabels" id="sRequiredFieldsLabels" value="<% = sRequiredFieldsLabels%>" >
				</tr>
			</table>
		</td>
	</tr>
	<% if not IsSearch then %>
	<tr>
        <td class="TableHeaderBlue" colspan="100"  style="width:100%;">
            &nbsp;<b>Additional Fields</b>
        </td>
    </tr>
    <% end if %>
    <tr>
        <td style="width:95%;">
            <table ID="Table5" cellpadding=3 style="width:100%;">
				<tr>
					<td class="InventoryNotRequired" valign="top" colspan=10>
						Description: 
						<% if not PrintMode or IsSearch then %>
						<textarea name="Description" id="Description"  class="InventorySelect" onKeyDown="jfMaxSize(2000,this);" style="width:100%;" rows="2"><% = Description %></textarea>
						<% else
								response.Write Description
						   end if
						%>
					</td>	
				</tr>
				<tr>
				<%
				if isArray(arFields) then
					ii = 0
						for i = 0 to ubound(arFields)	
							if ucase(arFields(i)) <> "WARRANTYINFO" and _
							   ucase(arFields(i)) <> "GENERALCOMMENTS" then
								execute("sValue = " & arFields(i))	
								if arFields(i) <> "InventoryDetailTypeID" and arFields(i) <> "Description" then								
					%>
						<td class="InventoryNotRequired" style="width:0%;">
							<nobr><% = arFieldLabels(i) %>:</nobr>
						</td>		
						<td style="width:33%;" class="TableCell">
							<% if not PrintMode or IsSearch then 
								if (ucase(arFields(i)) = "PONUMBER" or ucase(arFields(i)) = "DISTRICTCONTROLNUM" or ucase(arFields(i)) = "SERIALNUMBER") _
										and InventoryDetailID & "" = "" then
											sValue = ""
									end if
							%>
							<input type="text" name="<% = arFields(i) %>" id="<% = arFields(i) %>" value="<% = sValue%>" maxlength="50" style="width:100%;"  class="InventorySelect">															
							<% else
									response.Write sValue
								end if 
							%>
						</td>
					<%			
								elseif arFields(i) = "InventoryDetailTypeID" then
									call TypeList(InventoryCategoryID,InventoryDetailTypeID)															
								end if	
								ii = ii + 1
								if ii mod 3 = 0 then response.Write "</tr><tr>"		
								sValue = ""
							else
								if instr(1,ucase(sFields),"WARRANTYINFO") > 0 then 
									call AddStaticTextArea("WarrantyInfo",WarrantyInfo,"Warranty Info",False)									
								end if
								
								if instr(1,ucase(sFields),"GENERALCOMMENTS") > 0 then
									call AddStaticTextArea("GeneralComments",GeneralComments,"General Comments",False)								
								end if
							end if
						next				
				end if	
			end if				
				%>
					<input type="hidden" name="sFields" value="<% = sFields %>" ID="Hidden1">
					</tr>
            </table>
        </td>
    </tr>
    <% if not IsSearch and InventoryCategoryID & "" <> "" then %>
    <tr>
		<td >
			<% if InventoryDetailID & "" = "" and intOrdered_Item_ID & "" <> "" then%>
			
			<% else %>
			<table style="width:100%;">
				<tr>
					<td style="width:50%;" valign="top">
						<% call ReIssueCosts() %>
					</td>
				</tr>
			</table>
			<% end if %>
			<table style="width:100%;" ID="Table7">
				<tr>
					<td style="width:50%;" valign="top">
						<%  if oFunc.IsAdmin then 
								if InventoryDetailID & "" <> "" and PrintMode then
									call InventoryLog() 
								end if
							else
								if ucase(InventoryStatusCD) <> "CO" then
									call PlaceOnHold()
								end if
							end if
						%>
					</td>
				</tr>
			</table>
			<% if TotalCostNew & "" <> "" and totalPaidBack <> "" and oFunc.IsAdmin then %>
			<table style="width:100%;" ID="Table9">
				<tr>
					<td style="width:100%;" valign="top">
				<% = InvestmentRecapture() %>
					</td>
				</tr>
			</table>
			<% end if %>
		</td>
    </tr>
    <% end if %>
</table>
</td>
</tr>
</table>
<% 		
		if InventoryCategoryID & "" <> "" then
			if not IsSearch then 
				if oFunc.IsAdmin and PrintMode and (SelectedStudent_ID & "" <> "" or HeldForStudentID & "" <> "" )then
			%>
				&nbsp;&nbsp;&nbsp;<input type="button" value="Check Out" class="TableHeaderRed" onclick="jfValidateCheckOut();" ID="Button1" NAME="Button1">
				&nbsp;<button  class="TableHeaderBlue" onclick="window.location.href='./InventoryAdmin.asp?simpleHeader=<%=simpleHeader%>&InventoryDetailID=<%=InventoryDetailID%>&IsEdit=';" ID="Button4">Cancel</button>
			<%
				elseif oFunc.IsAdmin and not PrintMode then
			%>
				&nbsp;&nbsp;&nbsp;<input type="button" value="Save Record" class="TableHeaderRed" onclick="jfInventoryValidate();" ID="Button3" NAME="Button1">				
				&nbsp;<button  class="TableHeaderBlue" onclick="window.location.href='./InventoryAdmin.asp?simpleHeader=<%=simpleHeader%>&InventoryDetailID=<%=InventoryDetailID%>&IsEdit=';">Cancel</button>
			
				<% if oFunc.isAdmin and InventoryDetailID& "" = "" then %>
                		 &nbsp;<button class="TableHeaderBlueGray" onClick="jfTryAutoFill();" >Try Auto-Fill</button>
               			<% end if %>
			<%
				elseif HoldStudentID <> "" and not IsHeld then
			%>
				&nbsp;&nbsp;&nbsp;<input type="button" value="Place on Hold" class="TableHeaderRed" onClick="jfValidateHold();">
			<%
				end if%>
			
			<%
			else
			%>
			<input type="button" value="Search" class="TableHeaderBlue" onclick="this.form.SearchOnce.value='true';this.form.submit();" ID="Button2" NAME="Button1">
			<input type="hidden" name="SearchResultFieldsLabels" value="<% = sSearchResultFieldsLabels %>">
			<input type="hidden" name="SearchResultFields" value="<% = sSearchResultFields %>" ID="Hidden3">
			<%
			end if
		end if
		
		if request("simpleHeader") <> "" then 
		%>
			<input type="button" class="TableHeaderBlue" value="Close Window" onclick="<% if refreshParent & "" <> "" then response.Write "window.opener.location.reload();" end if%>window.opener.focus();window.close();">
		<%
		end if
		if SearchOnce & ""<> "" then 
			SearchItems() 
		end if	
end function

function TypeList(pInventoryCategoryID,pInventoryDetailTypeID)
	sql  = "SELECT     InventoryDetailTypeID, Name " & _ 
			"FROM         INVENTORY_DETAIL_TYPES " & _ 
			"WHERE     (InventoryCategoryID = " & pInventoryCategoryID & ") AND (IsActive = 1) " & _ 
			"ORDER BY Name "
	strDetailList = oFunc.MakeListSQL(sql,"InventoryDetailTypeID","Name",pInventoryDetailTypeID)
	
	' If our count is greater than 0 we do have a list
	if oFunc.makeListRecordCount > 0 then	
		' show detail list			
	%>
		<td class="<% = RequiredCss %>">
			Type:
		</td>		
		<td class="TableCell">
			<% if not PrintMode or IsSearch then %>
			<select name="InventoryDetailTypeID" id="InventoryDetailTypeID" class="InventorySelect">
				<option value=""></option>
				<% = strDetailList %>
			</select>
			<% else	
					response.Write oFunc.SelectedListText
			   end if
			%>
		</td>	
	<%
	end if 	
end function

function StatusList()
	sql  = "SELECT     InventoryStatusCD, Name " & _ 
			"FROM         INVENTORY_STATUS_CODES " & _ 
			"WHERE     (IsActive = 1) " & _ 
			"ORDER BY Name "
	' Non Admins only allowed to see available inventory		
	if not oFunc.IsAdmin and IsSearch then InventoryStatusCD = "AV"
	
	strStatusList = oFunc.MakeListSQL(sql,"InventoryStatusCD","Name",InventoryStatusCD)
	
	' If our count is greater than 0 we do have a list
	if oFunc.makeListRecordCount > 0 then	
		' show detail list			
	%>
		<td class="<% = RequiredCss %>">
			Status:
		</td>		
		<td class="TableCell">
			<% if not PrintMode then %>
			<nobr><select name="InventoryStatusCD" id="InventoryStatusCD" class="InventorySelect">
				<option value=""></option>
				<% = strStatusList %>
			</select>
				<% if DateHoldEnd <> "" then %>
				<span style="font-family:arial;font-size=7pt;color=red;">Held until <%= DateHoldEnd%></span></nobr>
				<% end if%>
			<% else
					response.Write oFunc.SelectedListText
			%>
				<input type="hidden" name="InventoryStatusCD" id="InventoryStatusCD" value="AV">
			<%
			   end if 
			%>
		</td>	
	<%
	end if 	
end function

function AddStaticTextArea(pID,pValue,pLabel,pRequired)
	if pRequired then
		pRequired = "InventoryRequired"
	else
		pRequired = "InventoryNotRequired"
	end if
%>
				<tr>
					<td class="<% = pRequired %>" valign="top" colspan=10>
						<% = pLabel %>: 
						<% if not PrintMode or IsSearch then %>
						<textarea name="<%= pID %>" id="<%= pID %>"  class="InventorySelect" onKeyDown="jfMaxSize(2000,this);" style="width:100%;" rows="2"><% = pValue %></textarea>
						<% else
								response.Write pValue
						   end if
						%>
					</td>	
				</tr>
<%	
end function

function InsertDetailRecord()
	dim insert, arFields1, arFields2, reqFields, nonReqFields
	arFields1 = split(sRequiredFields,",")
	
	for i = 0 to ubound(arFields1)
		if arFields1(i) <> "" then
			if i > 2 then
				execute("reqFields = reqFields & ""'"" & oFunc.EscapeTick(" &  arFields1(i) & ") & ""',""")	
			else
				execute("reqFields = reqFields & oFunc.EscapeTick(" &  arFields1(i) & ") & "",""")
			end if		
		end if		
	next
	
	reqFields = left(reqFields,len(reqFields)-1)
	arFields2 = split(sFields,",")
	for i = 0 to ubound(arFields2)
		execute("nonReqFields = nonReqFields & ""'"" &  oFunc.EscapeTick(" & arFields2(i) & ") & ""',""")
	next
	nonReqFields = left(nonReqFields,len(nonReqFields)-1)
	
	if intOrdered_Item_ID & "" <> "" then
		' add reference to original ordered item record
		addColumn = ", CreatedUsingOrdItemId "
		addValue = "," & intOrdered_Item_ID
	end if 
	
	insert = "insert into INVENTORY_DETAILS(" & sRequiredFields & "," & sFields & ", DateCreated,UserCreated" & addColumn & ") " & _
			 " values(" & reqFields & "," & nonReqFields & ",CURRENT_TIMESTAMP,'" & session.Contents("strUserID") & "'" & addValue & ")"

	oFunc.ExecuteCN(insert)	
	InventoryDetailID = oFunc.GetIdentity
	
	if intOrdered_Item_ID & "" <> "" then
		' creates reference to inventory record 
		update = "update tblOrdered_Items set InventoryDetailId = " & InventoryDetailID & _
				 " Where intOrdered_Item_ID = " & intOrdered_Item_ID 
		oFunc.ExecuteCN(update)
		'CheckOutItem InventoryDetailID,intStudent_ID
		PlaceHold intStudent_ID, IlpId, InventoryDetailID		
	end if
	
	TitleText = "<span class='error10'><b>Recorded has been created.</b></span><BR>"
	
	PrintMode = true
	
end function


function UpdateDetailRecord()
	dim update, arFields1, arFields2, reqFields, nonReqFields
	arFields1 = split(sRequiredFields,",")
	
	update = "update INVENTORY_DETAILS set " 
	
	for i = 0 to ubound(arFields1)
		if arFields1(i) <> "" then
			if i > 2 then
				execute("update = update &  arFields1(i) & "" = '"" & oFunc.EscapeTick(" & arFields1(i) & ") & ""',""")			
			else
				execute("update = update &  arFields1(i) & "" = "" & oFunc.EscapeTick(" & arFields1(i) & ") & "",""")			
			end if
		end if		
	next
	
	update = left(update,len(update)-1)
	
	arFields2 = split(sFields,",")
	for i = 0 to ubound(arFields2)
		execute("update = update &  "","" & arFields2(i) & "" = '"" & " & arFields2(i) & " & ""'""")
	next
	
	update = update & ", UserModified = '" & session.Contents("strUserID") & "', DateModified = CURRENT_TIMESTAMP "
	update = update & " WHERE InventoryDetailID = " & InventoryDetailID
	
	
	'response.Write update
	oFunc.ExecuteCN(update)
	
	TitleText = "<span class='error10'><b>Recorded has been updated.</b></span><BR>"	
	
	PrintMode = true
		
end function

function SearchItems()
	dim sql, arFields1, arFields2, reqFields, nonReqFields, sWhere, sCheckText, sCheckText2, rsSelect
	dim arFields, arLabels, k
	arFields1 = split(sRequiredFields & "," & sFields,",")
	
	for i = 0 to ubound(arFields1)
		if arFields1(i) <> "" then
			execute("sCheckText = " & arFields1(i) )
			if sCheckText & "" <> "" then
				if ucase(arFields1(i)) = "TOTALCOSTNEW" then
					execute("sCheckText2 = " & arFields1(i+1))
					if sCheckText2 & "" <> "" then
						execute("sWhere = sWhere & "" and TotalCostNew between convert(money,'"" & sCheckText & ""') and convert(money,'"" & sCheckText2 & ""') """)
					end if
					i = i + 1
				elseif ucase(arFields1(i)) = "DATEPURCHASED" then
					execute("sCheckText2 = " & arFields1(i+1))
					if isDate(sCheckText) and  isDate(sCheckText2) then
						execute("sWhere = sWhere & "" and DATEPURCHASED between convert(datetime,'"" & sCheckText & ""') and convert(datetime,'"" & sCheckText2 & ""') """)
					elseif isDate(sCheckText) then
						execute("sWhere = sWhere & "" and DATEPURCHASED >= convert(datetime,'"" & sCheckText & ""') """)
					end if
					i = i + 1
				elseif isNumeric(sCheckText& "") then
					execute("sWhere = sWhere & "" and "" & 	arFields1(i) & "" = '"" & sCheckText & ""' """)						
				else
					execute("sWhere = sWhere & "" and "" & 	arFields1(i) & "" like '%"" & sCheckText & ""%' """)						
				end if
			end if			
		end if									  		
	next
	
	if instr(1,ucase(SearchResultFields),"INVENTORYDETAILTYPEID") > 0 then
		sAdd = ", (select name from INVENTORY_DETAIL_TYPES where InventoryDetailTypeID = id.InventoryDetailTypeID) as TypeName "
	end if
	
	if SearchResultFields <> "" then mySearchFields = ", " & SearchResultFields
	
	sql = "select InventoryDetailID " & mySearchFields & sAdd & _
			" from INVENTORY_DETAILS id INNER JOIN " &_
            "	tblVendors v ON id.VendorID = v.intVendor_ID " & _
			" WHERE 1=1 " & sWhere
	set rsSelect = server.CreateObject("ADODB.RECORDSET")
	rsSelect.CursorLocation = 3

'response.Write sql
'response.End
	rsSelect.Open sql, oFunc.FPCScnn
	
	response.Write "<BR><BR><span class='svplain8'><b>" & rsSelect.RecordCount & " results found.</span>"
	if rsSelect.RecordCount > 0 then
		arLabels = split(SearchResultFieldsLabels,",")
		arFields = split(SearchResultFields,",")
		response.Write "<table style='width:100%;'><tr>"
		k = 1
		for i = 0 to ubound(arLabels)
			if arLabels(i) & "" <> "" then
		%>
			<td class="TableHeaderRed">
				<% = arLabels(i) %>&nbsp;
			</td>
		<%
			end if
		next
		response.Write "</tr>"
		do while not rsSelect.EOF
			response.Write "<TR onClick=""jfViewItem('" & rsSelect("InventoryDetailID") & "');"" style='cursor:pointer'>"
			for i = 0 to ubound(arFields)
				if arFields(i) & "" <> "" then
					if ucase(arFields(i)) = "INVENTORYDETAILTYPEID" then
						myValue = rsSelect("TypeName")
					else	
						if ucase(arFields(i)) = "TOTALCOSTNEW" then
							myValue = "$" & formatnumber(rsSelect(arFields(i)),2)
						else
							myValue = rsSelect(arFields(i))
						end if
					end if
			%>
			<td class="<% if k mod 2 = 0 then response.Write "svplain8" else response.Write "ltGray8" %>">
				<% = myValue %>&nbsp;
			</td>	
			<%
				end if
			next
			response.Write "</TR>"
			rsSelect.MoveNext
			k = k + 1
		loop
	end if
	response.Write "</table>"
end function

function ReIssueCosts()
	dim sql, rsCost, k, rCount
	if InventoryDetailID & "" <> "" then
			
			sql = "SELECT ReIssueCostID, ReIssueCost, ReIssueCostDate, Comments, DateModified, UserModified " & _ 
				"FROM INVENTORY_REISSUE_COSTS " & _ 
				"WHERE (InventoryDetailID = " & InventoryDetailID & ") " & _ 
				"ORDER BY ReIssueCostDate "
				
			set rsCost = server.CreateObject("ADODB.RECORDSET")
			rsCost.CursorLocation = 3
			rsCost.Open sql, oFunc.FPCScnn
			rCount = rsCost.RecordCount
	end if
		
	If oFunc.IsAdmin or rCount > 0 then	
		if oFunc.IsAdmin then
%>	
	
		<script language="javascript">
			function jsChangedIssueCost(pID){	
				if (document.main.ChangedIssueCosts.value.indexOf(","+pID+",") == -1 ) {
					document.main.ChangedIssueCosts.value = document.main.ChangedIssueCosts.value + pID + ",";
				}
			}	
		</script> 
		<input type="hidden" name="ChangedIssueCosts" id="ChangedIssueCosts" value=",">
		<table  style="width:100%;">
			<tr>
				<td class="TableHeaderBlueGray" colspan="4">
					&nbsp;<B>Reissue Costs</B>
				</td>
			</tr>
			<tr>
				<td class="InventoryNotRequired">
					<nobr>Reissue Cost</nobr>
				</td>	
				<td class="InventoryNotRequired">
					Comments
				</td>		
				<td class="InventoryNotRequired">
					<nobr>Reissue Date</nobr>
				</td>		
				<td class="InventoryNotRequired">
					<nobr>Issued By</nobr>
				</td>				
			</tr>
<%
		end if
		
		if InventoryDetailID & "" <> "" then
			k = 10
			if rsCost.RecordCount > 0 then
				if not oFunc.IsAdmin then rsCost.MoveLast
				do while not rsCost.EOF 
					if oFunc.IsAdmin then
						rsCost.MoveNext
						if rsCost.EOF then
							IsLast = true
						else
							IsLast = false
						end if
						rsCost.MovePrevious
					end if
					if oFunc.IsAdmin then
					%>
			<script language="javascript">
				var cal<%=k%> = new CalendarPopup('divCal');
				cal<%=k%>.showNavigationDropdowns();
			</script>
			<tr>
				<td valign="top" class="TableCell" align="center">
					<% if not PrintMode and IsLast then %>
					<nobr>$<input class="InventorySelect" type="text" name="Cost<% = rsCost("ReIssueCostID") %>" value="<% = formatNumber(rsCost("ReIssueCost"),2) %>" size=10 onchange="jsChangedIssueCost('<% = rsCost("ReIssueCostID") %>');"></nobr>
					<% else %>
					$<% = formatNumber(rsCost("ReIssueCost"),2) %>
					<% end if %>
				</td>
				<td class="TableCell" valign="top" style="width:100%;">
					<% if not PrintMode and IsLast then %>
					<textarea class="InventorySelect"  style="width:100%;"  rows="1" name="Comments<% = rsCost("ReIssueCostID") %>" onKeyDown="jfMaxSize(512,this);" onfocus="this.rows=4;" onblur="this.rows=1;" onchange="jsChangedIssueCost('<% = rsCost("ReIssueCostID") %>');"><% = rsCost("Comments") %></textarea>
					<% else %>
					 <% = rsCost("Comments") %> &nbsp;
					<% end if %>
				</td>
				<td valign="top" class="TableCell" align="center">
					<% if not PrintMode and IsLast  then %>
					<input class="InventorySelect" type="text" name="Date<% = rsCost("ReIssueCostID") %>" value="<%if isDate(rsCost("ReIssueCostDate")) then response.Write rsCost("ReIssueCostDate")%>" ID="<% = rsCost("ReIssueCostID") %>Date" onclick="cal<%= k%>.select(document.forms[0].Date<% = rsCost("ReIssueCostID") %>,'Date<% = rsCost("ReIssueCostID") %>','M/dd/yyyy','<%if isDate(rsCost("ReIssueCostDate")) then response.Write rsCost("ReIssueCostDate")%>');return false;" size=10 onchange="jsChangedIssueCost('<% = rsCost("ReIssueCostID") %>');">
					<% else %>
					 <% = rsCost("ReIssueCostDate") %>
					<% end if %>
				</td>		
				<td class="TableCell" valign="top" align="center">
					<% = rsCost("UserModified") %>
				</td>		
			</tr>
					<%
					end if
					StudentReissueCost = rsCost("ReIssueCost")
					ReIssueComments = rsCost("Comments")
					rsCost.MoveNext
					k = k + 1
				loop
			end if
			
			rsCost.Close
			set rsCost = nothing
		end if
	end if 	
	
	if not PrintMode then
%>			
			<script language="javascript">
				var cal<%=k%> = new CalendarPopup('divCal');
				cal<%=k%>.showNavigationDropdowns();
			</script>
			<tr>
				<td valign="top" class="TableCell">
					<nobr>$<input type="text" name="ReissueCost" value="" size=10 ID="Text1" class="InventorySelect"></nobr>
				</td>	
				<td valign="top" style="width:100%;">
					<textarea class="InventorySelect" style="width:100%;" rows="1" name="ReissueComments" onKeyDown="jfMaxSize(512,this);" ID="Comments" onfocus="this.rows=4;" onblur="this.rows=1;"></textarea>
				</td>			
				<td valign="top">
					<input class="InventorySelect" type="text" name="ReissueDate" value="" ID="ReissueDate" onclick="cal<%= k%>.select(document.forms[0].ReissueDate,'ReissueDate','M/dd/yyyy','<%if isDate(ReissueDate) then response.Write ReissueDate%>');return false;" size=10 >
				</td>	
				<td class="TableCell" valign="top">
					<% = session.Contents("strUserID") %>
				</td>			
			</tr>
<%
	end if
%>
		</table>
		
	</td>
<%
end function

function InsertReissueRecord	
	insert = "insert into INVENTORY_REISSUE_COSTS (ReIssueCost, ReIssueCostDate, Comments, InventoryDetailID, DateModified, UserModified, DateCreated, UserCreated) " & _
			 " values(" & ReIssueCost & ",'" & ReissueDate & "','" & ReissueComments & "'," & InventoryDetailID & ",CURRENT_TIMESTAMP,'" & session("strUserID") & "',CURRENT_TIMESTAMP,'" & session("strUserID") & "')"
	oFunc.ExecuteCN(insert)
	
	PrintMode = true
end function

function UpdateReissueRecords
	dim update, sIds, i
	arIds = split(ChangedIssueCosts,",")
	
	for i = 0 to ubound(arIds)
		if arIds(i) & "" <> "" then
			update = "update INVENTORY_REISSUE_COSTS set ReIssueCost = " & Request("Cost" & arIds(i))  & ", " & _
					 " ReIssueCostDate = '" & Request("Date" & arIds(i))  & "', " & _
					 " Comments = '" & oFunc.EscapeTick(Request("Comments" & arIds(i))) & "', " & _
					 " DateModified = CURRENT_TIMESTAMP, " & _
					 " UserModified = '" & session("strUserID") & "' " & _
					 " where ReIssueCostID = " & arIds(i)
			oFunc.ExecuteCN(update)			
		end if
	next	
	
	PrintMode = true
end function

function InvestmentRecapture()
%>	
		<table ID="Table8" style="width:100%;">
			<tr>
				<td class="TableHeaderBlueGray" colspan="7">
					&nbsp;<B>Investment Recapture</B>
				</td>
			</tr>
			<tr>
				<td class="InventoryNotRequired" align="center">
					&nbsp;Total Cost New
				</td>						
				<td class="InventoryNotRequired" align="center">
					<nobr>Total Cost Recaptured</nobr>
				</td>
				<td class="InventoryNotRequired" align="center">
					<nobr>Balance</nobr>
				</td>
			</tr>
			<tr>
				<td class='TableCell' align=center>
					$<% = formatNumber(TotalCostNew,2) %>
				</td>
				<td class='TableCell' align=center>
					$<% = formatNumber(totalPaidBack,2) %>
				</td>
				<td class='TableCell' align=center>
					$<% = formatNumber(TotalCostNew - totalPaidBack,2) %>
				</td>
			</tr>
		</table>
<%
end function

function InventoryLog()
%>	
	
		<input type="hidden" name="CheckOut" id="CheckOut" value="">
		<table ID="Table6" style="width:100%;">
			<tr>
				<td class="TableHeaderBlueGray" colspan="7">
					&nbsp;<B>Library Card</B>
				</td>
			</tr>
			<tr>
				<td class="InventoryNotRequired">
					&nbsp;Student
				</td>						
				<td class="InventoryNotRequired" align="center">
					<nobr>Course</nobr>
				</td>
				<td class="InventoryNotRequired" align="center">
					<nobr>Checked Out</nobr>
				</td>
				<td class="InventoryNotRequired" align="center">
					<nobr>Due Date</nobr>
				</td>
				<td class="InventoryNotRequired" align="center">
					<nobr>Checked In</nobr>
				</td>
				<td class="InventoryNotRequired" align="center">
					Cost
				</td>
				<td class="InventoryNotRequired" style="width:100%;">
					&nbsp;Comments
				</td>
			</tr>
<%
		myCheckIn = "true"
		if InventoryDetailID & "" <> "" then
			dim sql, rsLog, k
			sql = "SELECT    coi.CheckedOutInventoryID, coi.StudentID, coi.Comments, coi.DateCheckedIn, coi.DateCheckedOut,  " & _ 
					"	coi.CheckedOutInventoryID, coi.OrderedItemID, coi.IlpID, s.szLAST_NAME,  " & _ 
					"	s.szFIRST_NAME, coi.DateDue, c.szClass_Name,  " & _ 
					"	oi.intQty * oi.curUnit_Price + oi.curShipping AS cost " & _ 
					"FROM	INVENTORY_CHECKED_OUT coi INNER JOIN " & _ 
					"	tblSTUDENT s ON coi.StudentID = s.intSTUDENT_ID INNER JOIN " & _ 
					"	tblILP i ON coi.IlpID = i.intILP_ID INNER JOIN " & _ 
					"	tblClasses c ON i.intClass_ID = c.intClass_ID LEFT OUTER JOIN " & _ 
					"	tblOrdered_Items oi ON coi.OrderedItemID = oi.intOrdered_Item_ID " & _ 
					"WHERE	(coi.InventoryDetailID = " & oFunc.EscapeTick(InventoryDetailID) & ") " & _ 
					"ORDER BY coi.CheckedOutInventoryID,coi.DateCheckedOut, coi.DateCheckedIn "
					
					'response.Write sql 

			set rsLog = server.CreateObject("ADODB.RECORDSET")
			rsLog.CursorLocation = 3
			rsLog.Open sql, oFunc.FPCScnn
			
			k = 40			
			if rsLog.RecordCount > 0 then
				do while not rsLog.EOF	
					if isNumeric(rsLog("cost")) then
						myCost = formatNumber(rsLog("cost"),2)
						totalPaidBack = totalPaidBack + rsLog("cost")
					else
						myCost = "0.00"
					end if
					%>
			<tr>
				<td class="TableCell" valign="top">
					<nobr><% = rsLog("szFirst_Name") & " " & rsLog("szLast_Name") %></nobr>
				</td>
				<td class="TableCell" valign="top">
					<nobr><% = rsLog("szClass_Name") %></nobr>
				</td>
				<td class="TableCell" valign="top" align="center">
					<% = rsLog("DateCheckedOut") %>&nbsp;
				</td>
				<td class="TableCell" valign="top" align="center">
					<% if isDate(rsLog("DateDue")) then %>
					<% = formatdatetime(rsLog("DateDue"),2) %>
					<% end if %>
					&nbsp;
				</td>
				<td class="TableCell" valign="top" align="center">
					<%if isDate(rsLog("DateCheckedIn")) then %>
					<% = formatdatetime(rsLog("DateCheckedIn"),2) %> &nbsp;
					<% else %>
					<input type="hidden" name="CheckedOutInventoryID" value="<% =  rsLog("CheckedOutInventoryID")%>">
					<script language="javascript">
						function jfCheckIn(){
							var ans = confirm("Are you sure you want to check this item in?");
							if (ans){
								document.forms[0].submit();						
							}
						}
					</script>
					<input type="checkbox" onClick="jfCheckIn();">
					<%end if%>
				</td>	
				<td class="TableCell" valign="top">
					<nobr>$<% = myCost%></nobr>
				</td>	
				<td class="TableCell" valign="top">
					<%if rsLog("DateCheckedIn") & "" <> "" then %>
						<% = rsLog("Comments") %>&nbsp;
					<% else %>
					<textarea  class="InventorySelect"  id="comments" name="comments" style="width:100%;" rows="1"  onfocus="this.rows=4;" onblur="this.rows=1;" onKeyDown="jfMaxSize(2000,this);"><% = rsLog("Comments") %></textarea>
					<% end if %>
				</td>				
			</tr>
					<%
					myCheckIn = rsLog("DateCheckedIn")
					rsLog.MoveNext
					k = k + 1
				loop
			end if
			
			rsLog.Close
			set rsLog = nothing
		end if
		
		if myCheckIn <> "" then
%>			
			<script language="javascript">
				var cal<%=k%> = new CalendarPopup('divCal');
				cal<%=k%>.showNavigationDropdowns();
				var cal<%=k+1%> = new CalendarPopup('divCal');
				cal<%=k+1%>.showNavigationDropdowns();
				var cal<%=k+2%> = new CalendarPopup('divCal');
				cal<%=k+2%>.showNavigationDropdowns();
			</script>
			<tr>
				<td align="center" valign="top">					
					<select name="SelectedStudent_ID" style="width:225px;" ID="SelectedStudent_ID" onchange="window.location.href='./InventoryAdmin.asp?simpleHeader=<%=simpleHeader%>&InventoryDetailID=<%=InventoryDetailID%>&SelectedStudent_ID=' + this.value;">
						<option value="">Select a Student
							<%
								dim sqlStudent, sqlWhere														
									if oFunc.IsGuardian then
										sqlWhere = " AND s.intFamily_ID = " & session.Contents("intFamily_id") & " "
									end if
									
									if oFunc.IsTeacher then
										sqlWhere = " AND s.intStudent_ID = "
									end if
									
									sqlStudent = "SELECT     s.intSTUDENT_ID, (CASE ss.intReEnroll_State WHEN 86 THEN s.szLAST_NAME + ',' + s.szFIRST_NAME + ': Withdrawn (' + CASE isNull(ss.dtWithdrawn, " & _ 
												" 1) WHEN 1 THEN 'No Date Entered' ELSE CONVERT(varChar(100), ss.dtWithdrawn)  " & _ 
												" END + ')' WHEN 123 THEN s.szLAST_NAME + ',' + s.szFIRST_NAME + ': Graduated (' + CONVERT(varChar(20), ss.dtModify)  " & _ 
												" + ')' ELSE s.szLAST_NAME + ',' + s.szFIRST_NAME END) AS Name, ss.intReEnroll_State, ss.dtWithdrawn " & _ 
												"FROM tblSTUDENT s INNER JOIN " & _ 
												" tblStudent_States ss ON s.intSTUDENT_ID = ss.intStudent_id " & _ 
												"WHERE (ss.intReEnroll_State IN (" & Application.Contents("strEnrollmentList") & ")) AND (ss.intSchool_Year = " & Session.Value("intSchool_Year") & ") " & _ 
												sqlWhere & _
												"ORDER BY Name "																													
								Response.Write oFunc.MakeListSQL(sqlStudent,"intStudent_ID","Name",SelectedStudent_ID & HeldForStudentID)												 
							%>
					</select>
				</td>
				<% 
				if SelectedStudent_ID & "" <> "" or HeldForStudentID & "" <> "" then %>
				<td valign="top">
					<select id="IlpID" name="IlpID"  style="width:225px;" >
						<option value=""></option>
						<% 
							sql = "SELECT	tblILP.intILP_ID, tblClasses.szClass_Name " & _ 
								"FROM	tblILP INNER JOIN " & _ 
								"	tblClasses ON tblILP.intClass_ID = tblClasses.intClass_ID " & _ 
								"WHERE	(tblILP.intStudent_ID = " & SelectedStudent_ID & HeldForStudentID & ") AND (tblILP.sintSchool_Year = " & session.Value("intSchool_Year") & ") " & _ 
								"ORDER BY tblClasses.szClass_Name "
							Response.Write oFunc.MakeListSQL(sql,"intILP_ID","szClass_Name",IlpID & HeldForIlpID)												 
						%>
					</select>
				</td>
				<td valign="top">
					<input class="InventorySelect" type="text" name="DateCheckedOut" ID="DateCheckedOut" value="<% = formatDateTime(now(),2)%>" onclick="cal<%= k%>.select(document.forms[0].DateCheckedOut,'DateCheckedOut','M/dd/yyyy','<% = formatDateTime(now(),2)%>');return false;" size=10 >
				</td>	
				<td valign="top">
					<input class="InventorySelect" type="text" name="DateDue" value="" ID="DateDue" onclick="cal<%= k+1%>.select(document.forms[0].DateDue,'DateDue','M/dd/yyyy','<%if isDate(DateDue) then response.Write DateDue%>');return false;" size=10 >
				</td>
				<td valign="top">
					<input class="InventorySelect" type="text" name="DateCheckedIn" value="" ID="DateCheckedIn" onclick="cal<%= k+2%>.select(document.forms[0].DateCheckedIn,'DateCheckedIn','M/dd/yyyy','<%if isDate(DateCheckedIn) then response.Write DateCheckedIn%>');return false;" size=10 >
				</td>
				<td class="svplain8" valign="top">
					<% if HeldWithOrdID and oiCost & "" <> "" then %>
						<nobr>$<% = formatNumber(oiCost,2) %></nobr>
					<% else %>
						<nobr>$<input  class="InventorySelect"  type="text" name="CheckOutCost" id="CheckOutCost" value="<% = formatNumber(StudentReissueCost,2) %>" size="10"></nobr>
					<% end if %>
				</td>		
				<td valign="top">
					<textarea  class="InventorySelect"  id="LogComments" name="LogComments" style="width:100%;" rows="1"  onfocus="this.rows=4;" onblur="this.rows=1;" onKeyDown="jfMaxSize(2000,this);"></textarea>
				</td>	
				<% 
					if StudentReissueCost > 0 then
						dBudget = AvailableBudget(SelectedStudent_ID & HeldForStudentID,StudentReissueCost)						
				%>
			</tr>
			<tr>
				<td class="sverror" colspan="10">
					<% if not IsFundsAvailable then %>
					Checking this item out will put this students budget in the negative by
					-$<% = FormatNumber((cdbl(dBudget) - cdbl(StudentReissueCost))*-1,2)%>.<br>
					<% end if %>
					Students current Budget Balance: $<% = FormatNumber(cdbl(dBudget),2) %>.
				</td>
						
				<%	end if
				end if%>				
			</tr>
<%
				if HeldForStudentID & "" <> "" then
%>
			<tr>
				<td class="svplain8" colspan="10">
					This is an ordered item from the students packet and is currently 'On Hold'.
				</td>
			</tr>
<%
				end if
			end if
%>
		</table>		
	</td>
<%
end function

function CheckOutItem(pInventoryDetailID,pStudentID)
	dim sql, rsCOIC, update
	sql = "SELECT	S.szLAST_NAME, S.szFIRST_NAME, ICO.DateCheckedOut, ICO.DateDue " & _ 
			"FROM	INVENTORY_CHECKED_OUT ICO INNER JOIN " & _ 
			"	tblSTUDENT S ON ICO.StudentID = S.intSTUDENT_ID " & _ 
			"WHERE	(ICO.InventoryDetailID = " & pInventoryDetailID & ") AND (ICO.DateCheckedIn IS NULL) "
	set rsCOIC = server.CreateObject("ADODB.RECORDSET")
	rsCOIC.CursorLocation = 3
	rsCOIC.Open sql, oFunc.FPCScnn
	
	if rsCOIC.RecordCount > 0 then
		TitleText = "<span class='error10'><b>This item can not be checked out because it is already checked out to " & _
					rsCOIC("szFirst_Name") & " " & rsCOIC("szLast_Name") & ".<BR> The due date is " & _
					rsCOIC("DateDue") & " and was checked out on " & rsCOIC("DateCheckedOut") & ".</b></span><BR>"		
	else
		if DateDue & "" = "" then 
			DateDue = " null "
		else
			DateDue = "'" & DateDue & "' "
		end if 
		
		' We need to determine if we have a legitimate cost
		IsWithCost = false
		if isNumeric(CheckOutCost & "") then
			CheckOutCost = cdbl(CheckOutCost)
			if CheckOutCost > 0 then
				IsWithCost = true
			end if
		end if
		
		if not isDate(DateCheckedOut) then DateCheckedOut = FormatDateTime(DateCheckedOut,3)
		
		if IsWithCost then
			' Run Sp that handles adding Budget and Line Item entries
			runSp = "Exec ts_InventoryCheckoutWithCost " & _
							pInventoryDetailID & "," & _
							pStudentID & "," & _
							IlpID & "," & _
							session.Contents("intSchool_Year") & "," & _
							application.Contents("SchoolVendorID") & "," & _
							CheckOutCost & "," & _
							"'" & session.Contents("strUserId") & "'," & _
							"'" & DateCheckedOut & "'," & _
							DateDue & "," & _
							"'" & oFunc.EscapeTick(LogComments) & "'" 
							
			oFunc.ExecuteCN(runSP)
		else
			' Run Sp that manages simple checkout 
			
			' We may have an intOrdered_Item_ID if this Inventory item
			' was created from an ordered item
			if isNumeric(intOrdered_Item_ID & "") then 
				oID = intOrdered_Item_ID			
			else
				oID = " null "
			end if
				
			runSp = "Exec ts_InventoryCheckout " & _
							pInventoryDetailID & "," & _
							pStudentID & "," & _
							IlpID & "," & _
							session.Contents("intSchool_Year") & "," & _
							"'" & session.Contents("strUserId") & "'," & _
							"'" & DateCheckedOut & "'," & _
							DateDue & "," & _
							"'" & oFunc.EscapeTick(LogComments) & "'," & _
							oID
							
			oFunc.ExecuteCN(runSP)
		end if
		
		TitleText = "<span class='error10'><b>Item has been checked out.</b></span><BR>"
		InventoryStatusCD = "CO"
		DateHoldEnd = ""
		HeldForStudentID = ""
	end if
	
	rsCOIC.Close
	set rsCOIC = nothing
end function

sub CheckInItem(pCheckedOutInventoryID,pInventoryDetailId,pComments)
	dim spText
	spText = "Exec ts_InventoryCheckIn " & _
			 pCheckedOutInventoryID & "," & _
			 pInventoryDetailId & "," & _
			 "'" & oFunc.EscapeTick(pComments) & "'," & _
			 "'" & session.Contents("strUserId") & "'"
	oFunc.ExecuteCN(spText)
	
	TitleText = "<span class='error10'><b>Item has been checked in.</b></span><BR>"
		InventoryStatusCD = "AV"
		DateHoldEnd = ""
		HeldForStudentID = ""
end sub

function PlaceOnHold
%>
	<input type="hidden" name="SaveHold" id="SaveHold" value="">
	<table ID="Table6" style="width:100%;">
		<tr>
			<td class="TableHeaderBlueGray" colspan="6">
				&nbsp;<B>Hold Request</B>
			</td>
		</tr>
		<tr>
			<td class="InventoryNotRequired" style="width:0%;">
				&nbsp;Student to hold item for ... 
			</td>	
			<% if HoldStudentID <> "" or HeldForStudentID & "" <> "" then %>
			<td class="InventoryNotRequired">
				<nobr>Course where item will be used ...</nobr>
			</td>	
			<td class="InventoryNotRequired" align="center">
				<nobr>Held Until</nobr>
			</td>			
			<td class="InventoryNotRequired" align="center">
				<nobr>Cost</nobr>
			</td>
			<td class="InventoryNotRequired" align="center">
				<nobr>Comments</nobr>
			</td>
			<% else %>
			<td class="InventoryNotRequired" style="width:100%;" colspan="4">
				&nbsp;
			</td>
			<% end if %>
		</tr>
		<tr>
			<td align="center" valign="top" style="width:0%;">
				<select name="HoldStudentID" style="width:225px;" ID="Select2" onchange="window.location.href='./InventoryAdmin.asp?simpleHeader=<% = simpleHeader %>&panel=New&InventoryDetailID=<%=InventoryDetailID%>&HoldStudentID='+this.value;">
					<option value="">
						<%
							dim sqlStudent, sqlWhere														
								if oFunc.IsGuardian then
									sqlWhere = " AND s.intFamily_ID = " & session.Contents("intFamily_id") & " "
								end if
								
								if oFunc.IsTeacher then
									sqlFrom = " inner join tblEnroll_Info ei ON ei.intStudent_ID = s.intStudent_ID " & _
												" AND (ei.sintSchool_Year = " & Session.Value("intSchool_Year") & ") "   
									sqlWhere = " AND ei.intSponsor_Teacher_ID = " & session.Contents("instruct_ID") & " "
								end if
								
								sqlStudent = "SELECT s.intSTUDENT_ID, s.szLAST_NAME + ',' + s.szFIRST_NAME AS Name, ss.intReEnroll_State, ss.dtWithdrawn " & _ 
											"FROM tblSTUDENT s INNER JOIN " & _ 
											" tblStudent_States ss ON s.intSTUDENT_ID = ss.intStudent_id " & _ 
											" AND (ss.intSchool_Year = " & Session.Value("intSchool_Year") & ") " & sqlFrom & _ 
											"WHERE (ss.intReEnroll_State IN (" & Application.Contents("ActiveEnrollList") & ")) " & _
											sqlWhere & _
											"ORDER BY Name "																													
							Response.Write oFunc.MakeListSQL(sqlStudent,"intStudent_ID","Name",HeldForStudentID & HoldStudentID)												 
						%>
				</select>
			</td>
			<% 			
			if HoldStudentID <> "" or HeldForStudentID & "" <> "" then 
			%>
			<td class="svplain8"  style="width:0%;" >
				<select name="HeldForIlpID" id="IlpID"  style="width:225px;">
					<option value="">Select a Course</option>
					<%
						sql = "SELECT	tblILP.intILP_ID, tblClasses.szClass_Name " & _ 
							"FROM	tblILP INNER JOIN " & _ 
							"	tblClasses ON tblILP.intClass_ID = tblClasses.intClass_ID " & _ 
							"WHERE	(tblILP.sintSchool_Year = " & session.Contents("intSchool_Year") & ") AND (tblILP.intStudent_ID = " & HoldStudentID & HeldForStudentID & ") " & _
							" ORDER BY tblClasses.szClass_Name "
						Response.Write oFunc.MakeListSQL(sql,"intILP_ID","szClass_Name",HeldForIlpID)												 
					%>
				</select>
			</td>
			<td class="svplain8" align="center">
				<% = FormatDateTime(Dateadd("d",7,now),2)%>
			</td>
			<td class="svplain8" align="center">
				$<% = FormatNumber(StudentReissueCost)%>
			</td>
			<td class="svplain8" align="center">
				<% = ReIssueComments%>
			</td>
		<% 
			if StudentReissueCost > 0 then
				dBudget = AvailableBudget(HoldStudentID & HeldForStudentID,StudentReissueCost)						
		%>
		</tr>
		<tr>
			<td class="sverror" colspan="10">
				<% if not IsFundsAvailable then %>
				Checking this item out will put this students budget in the negative by
				-$<% = FormatNumber((cdbl(dBudget) - cdbl(StudentReissueCost))*-1,2)%>.<br>
				<% end if %>
				Students current Budget Balance: $<% = FormatNumber(cdbl(dBudget),2) %>.
			</td>
						
		<%		end if
		end if %>
		</tr>
	</table>
<%
end function

sub PlaceHold(pStudentID, pIlpID, pInventoryDetailID)
	dim update, sql, sMoreSql
	set rsChk = server.CreateObject("ADODB.RECORDSET")
	rsChk.CursorLocation = 3
	
	sql = "select InventoryStatusCD from INVENTORY_DETAILS where InventoryDetailID = " & pInventoryDetailID
	rsChk.Open sql, oFunc.FPCScnn
	
	TitleText = "<span class='error10'><b>Hold attempt failed. Item is being held for another student.</b></span><BR>"
	
	if rsChk.RecordCount > 0 then
		if ucase(rsChk("InventoryStatusCD")) & "" = "AV" or isPlaceOnHold & "" <> "" then
			if intOrdered_Item_ID & "" <> "" then
				sMoreSql = ", HeldWithOrdId = 1 " 
			end if
			
			update = "update INVENTORY_DETAILS set InventoryStatusCD = 'OH' , " & _
					 " HeldForStudentID = " & pStudentID & ", " & _ 
					 " HeldForIlpID = " & pIlpID & ", " & _
					 " DateHoldEnd = '" & FormatDateTime(Dateadd("d",7,now),2) & "' " & _
					 sMoreSql & _
					 " Where InventoryDetailID = " & pInventoryDetailID
			oFunc.ExecuteCN(update)
			TitleText = "<span class='error10'><b>Item has been placed on hold.</b></span><BR>"
			IsHeld = true
			InventoryStatusCD = "OH"
		end if
	end if
end sub

function AvailableBudget(pStudentID,pCost)
	set oBudget = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/StudentBudgetInfo.wsc"))
	oBudget.PopulateStudentFunding oFunc.FPCScnn, pStudentId, session.Contents("intSchool_Year")
	if cdbl(oBudget.BudgetBalance) - cdbl(pCost) >= 0 then
		IsFundsAvailable = true
	else
		IsFundsAvailable = false
	end if
	dBalance = oBudget.BudgetBalance
	set oBudget = nothing
	
	AvailableBudget = dBalance
end function
%>