<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		reqClassItemAdmin.asp
'Purpose:	Gives admin the abilty to view and approve/deny goods and
'			services.
'Date:		20 OCT 2005
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Dimension Variables
dim intCount			'used as a tracking mechanism for our update subroutine 
dim sql					'sql that helps us populate our form		
dim rs					'recordset that helps us populate our form
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

Session.Value("strTitle") = "Class Item Approval Admin"
Session.Value("strLastUpdate") = "27 OCT 2005"
Server.Execute(Application.Value("strWebRoot") & "includes/Header.asp")

'Create object containing all of our FPCS functions
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
set oHtml = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/htmlFuncs.wsc"))
call oFunc.OpenCN()

if (request.Form("updateList") <> "" and request.Form("updateList") <> ",") or _
	(request.Form("LineItemsChanged") <> "," and request.Form("LineItemsChanged") <> "") then
	'oFunc.BeginTransCN
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
	'oFunc.CommitTransCN
end if

%>
<script language="javascript">
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
	
	function jfChangedLI(id){	
		if (document.main.LineItemsChanged.value.indexOf(","+id+",") == -1) {
			document.main.LineItemsChanged.value = document.main.LineItemsChanged.value + id + ",";
		}
	}
	
	function jfChangedObj(objName,id){	
		var myObj = document.getElementById(objName);
		if (myObj.value.indexOf("," + id + ",") == -1) {
			// add id
			myObj.value = myObj.value + id + ",";
		}else{
			// remove id
			var re = new RegExp("," + id + ",",'g');
			myObj.value = myObj.value.replace(re,',');
		}
		//alert(myObj.value);
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
	
	function jfUpdateList(id) {
		// if an item as been changed log it on;y once.  We will use this list
		// to determine which OI's should be modified
		if (document.main.updatelist.value.indexOf(","+id+",") == -1 ) {
			document.main.updatelist.value = document.main.updatelist.value + id + ",";
		}
	}	
	
	function jfGetLineItemTotal(pID){
		var dUnits = document.getElementById("intQuantity"+pID);
		var dPrice= document.getElementById("curUnit_Price"+pID);
		var dShip = document.getElementById("curShipping"+pID);
		return (parseFloat((dUnits.value != "")?dUnits.value:1)*parseFloat((dPrice.value != "")?dPrice.value:0))+parseFloat((dShip.value != "")?dShip.value:0)
	}
	
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
</script>
<form action="reqClassItemAdmin.asp" method=post name=main ID="Form1">
<input type="hidden" name="PageNumber" value="<% = intPageNum%>" ID="Hidden7">
<input type=hidden name="lastRow" ID="Hidden1">
<input type=hidden name="updatelist" value="," ID="Hidden2">
<input type=hidden name="GSList" value="," ID="Hidden13">
<input type=hidden name="LineItemsChanged" value="," ID="Hidden8">
<input type=hidden name="lastRowColor" ID="Hidden3">
<table width=100% ID="Table1">	
	<tr>	
		<Td class=yellowHeader>
			&nbsp;<b>ASD Class Item Approval Admin</b>
		</td>
	</tr>
	<tr>
		<td>
			<table ID="Table2" cellspacing=0>
				<tr>					
					<td class="svplain10">
						<% if request("selTeacherID") & "" = "" and request("selClassId") & "" = "" then%>
						<BR>Select an Instructor or a Course Name and then click 'Save/Requery'.<br><br>
						<% elseif request("selClassId") & "" = "" then %>
						<BR>Select a Course Name and then click 'Save/Requery'.<br><br>
						<% end if %>
						<table ID="Table3">		
							<tr>
								<td colspan="10" class=navywhite8 valign=middle>
									<B>&nbsp;Filter Criteria&nbsp;</B>
								</td>
							</tr>							
							<tr>
								<td	class=gray>
									<b><nobr>&nbsp;Instructor&nbsp;</nobr></b>
								</td>
								<td	class=gray>
									<b><nobr>&nbsp;Course Name&nbsp;</nobr></b>
								</td>																
							</tr>
							<tr>
								<td>
									<select name="selTeacherID" ID="Select4" onchange="this.form.PageNumber.value='1';this.form.selClassId.selectedIndex=0;this.form.submit();">
									<option value=""></option>
									<%
										' create list of classes that have Class Items and also have students enrolled
										sql = "SELECT DISTINCT i.intINSTRUCTOR_ID, i.szLAST_NAME + ', ' + i.szFIRST_NAME AS TeacherName " & _ 
												"FROM	tblClass_Items INNER JOIN " & _ 
												"	tblClasses ON tblClass_Items.intClass_ID = tblClasses.intClass_ID INNER JOIN " & _ 
												"	tblINSTRUCTOR i ON tblClasses.intInstructor_ID = i.intINSTRUCTOR_ID " & _ 
												"WHERE	(tblClass_Items.intClass_Item_ID IS NOT NULL) AND (tblClasses.intSchool_Year = " & session.Contents("intSchool_Year") & ") AND EXISTS " & _ 
												"	(SELECT	'x' " & _ 
												"		FROM	tblIlp ilp " & _ 
												"		WHERE	ilp.intClass_ID = tblClasses.intClass_ID) " & _ 
												"ORDER BY TeacherName "
										response.Write oFunc.MakeListSQL(sql,"intINSTRUCTOR_ID","TeacherName",request("selTeacherID"))
									%>						
									</select>
								</td>
								<td>
									<select name="selClassId" ID="Select7" onchange="this.form.PageNumber.value='1';">
									<option value=""></option>
									<%
									
										if isnumeric(request("selTeacherID")) and request("selTeacherID") & ""  <> "" then
											sqlWhere = " AND tblClasses.intInstructor_id = " & request("selTeacherID") & " "
										end if 
										' create list of classes that have Class Items and also have students enrolled
										sql = "SELECT DISTINCT tblClasses.intClass_ID, SUBSTRING(tblClasses.szClass_Name, 1, 50) AS ClassName " & _ 
												"FROM	tblClass_Items INNER JOIN " & _ 
												"	tblClasses ON tblClass_Items.intClass_ID = tblClasses.intClass_ID " & _ 
												"WHERE	(tblClass_Items.intClass_Item_ID IS NOT NULL) AND (tblClasses.intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _ 
												"	AND EXISTS (SELECT	'x' " & _ 
												"		FROM	tblIlp ilp " & _ 
												"		WHERE	ilp.intClass_ID = tblClasses.intClass_ID) " & _ 
												sqlWhere & _
												" ORDER BY ClassName "
										response.Write oFunc.MakeListSQL(sql,"intClass_ID","ClassName",request("selClassId"))
									%>						
									</select>
								</td>									
								<td rowspan="2" valign="middle">
									<input type="submit" value="Save/Requery" class="NavSave" ID="Submit1" NAME="Submit1">
								</td>																																	
							</tr>							
						</table>
								
					</td>
				</tr>
			</table>
		</td>
	</tr>
<%

if isnumeric(request("selClassId")) and request("selClassId") & ""  <> "" then
			
	sql = "SELECT     c.intInstructor_ID, c.szClass_Name, ci.intClass_Item_ID, ci.intQty, ci.intClass_ID, it.szName, it.intItem_ID, ci.curUnit_Price, ci.curShipping, ci.intSchool_Year,  " & _ 
			"	v.szVendor_Name, v.szVendor_Phone, v.szVendor_Contact, v.szVendor_Email, ci.bolRequired, ci.bolClosed,  " & _ 
			"	ci.intQty * ci.curUnit_Price + ci.curShipping AS Total, " & _ 
			"	(SELECT     ca2.szValue " & _ 
			"		FROM          tblClass_Attrib ca2 " & _ 
			"		WHERE      ca2.intClass_Item_Id = ci.intClass_Item_Id AND ca2.intItem_Attrib_ID = 15) AS Consumable, " & _ 
			"	(SELECT     TOP 1 ca2.szValue " & _ 
			"		FROM          tblClass_Attrib ca2 " & _ 
			"		WHERE      ca2.intClass_Item_Id = ci.intClass_Item_Id AND (ca2.intItem_Attrib_ID = 9 OR " & _ 
			"		ca2.intItem_Attrib_ID = 5 OR " & _ 
			"		ca2.intItem_Attrib_ID = 6 OR " & _ 
			"		ca2.intItem_Attrib_ID = 22 OR " & _ 
			"		ca2.intItem_Attrib_ID = 33) " & _ 
			"		ORDER BY ca2.intItem_Attrib_ID) AS iName, i.szFIRST_NAME + ' ' + i.szLAST_NAME AS TeacherName, ig.szName AS GroupName,  " & _ 
			"	ig.intItem_Group_ID, ci.bolApproved, ci.szComments, c.intInstruct_Type_ID, c.intGuardian_ID, c.intPOS_Subject_ID " & _ 
			"FROM	tblClass_Items ci INNER JOIN " & _ 
			"	tblClasses c ON ci.intClass_ID = c.intClass_ID INNER JOIN " & _ 
			"	tblINSTRUCTOR i ON c.intInstructor_ID = i.intINSTRUCTOR_ID INNER JOIN " & _ 
			"	trefItems it ON ci.intItem_ID = it.intItem_ID INNER JOIN " & _ 
			"	tblVendors v ON ci.intVendor_ID = v.intVendor_ID INNER JOIN " & _ 
			"	trefItem_Groups ig ON it.intItem_Group_ID = ig.intItem_Group_ID " & _ 
			"WHERE     (c.intClass_ID = " & request("selClassId") & ") " & _ 
			"ORDER BY ci.bolRequired DESC "

	set rs = server.CreateObject("ADODB.RECORDSET")
	rs.CursorLocation = 3
	rs.Open sql, oFunc.FPCScnn
	
	if rs.RecordCount > 0 then 
		set rsLI = server.CreateObject("ADODB.RECORDSET")
		rsLI.CursorLocation = 3
	%>
	<tr>
		<td class="svplain10">
			<br>
			<b>Class Name:</b> <% = rs("szClass_Name") %> <b>Instructor: </b> <% = rs("TeacherName") %> 
		</td>
	</tr>
	<tr>
		<td>
			<table cellpadding="4">
				<%
				call PrintHeader
				
				intCount = 0
				intCount2 = 0
				
				set myRs = server.CreateObject("ADODB.RECORDSET")
				myRs.CursorLocation = 3
				
				do while not rs.EOF	
					' Set row color
					if intCount mod 2 = 0 then
						strColor = "plainCell"
					else
						strColor = "gray"
					end if
									
					%>
					<input type=hidden name="intClass_Item_ID<%=intCount%>" value="<%=rs("intClass_Item_ID")%>" ID="Hidden4">
				<tr class='<%=strColor%>'  id="ROW<%=intCount%>" onClick="jfHighLight('<%=intCount%>');">
					<td >
						<% = rs("szVendor_Name") %>
					</td>
					<td >
						<% = rs("GroupName") %>
					</td>
					<td>
						<% = rs("intClass_Item_ID") %>
					</td>
					<td >
						<% = rs("szName") %>
					</td>
					<td>
						<% = rs("iName") %>
					</td>
					<td align="right" title="Number of Units = <% = rs("intQty") %>: Unit Price = <%= rs("curUnit_Price")%>: Shipping = <% = rs("curShipping") %>">
						$<% = FormatNumber(rs("Total"),2) %>
					</td>	
					<%
					' get all accounting data that pertains to this budget
					sql = "SELECT intClass_Line_Item_ID, dtLine_Item, szLine_Item_desc, curUnit_Price, intQuantity,  " & _ 
						"curShipping, szPO_Number, szInvoice_Number, szCheck_Number,  " & _ 
						"szPayee, dtCheck, dtReciept, dtCREATE, dtMODIFY, szUSER_CREATE, szUSER_MODIFY " & _ 
						", (curUnit_Price*intQuantity)+curShipping as total " & _
						"FROM tblClass_Line_Items " & _ 
						"WHERE (intClass_Item_ID = "  & rs("intClass_Item_ID") & ") " & _
						" ORDER BY intClass_Line_Item_ID"					
						
					rsLI.Open sql, oFunc.FPCScnn
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
					dblBudgetAmount = rs("total")
					%>
					<td align="right" title="Click here to view Line Item entries." id="Balance<%=intCount%>" onClick="jfToggle('LineItem<%=intCount%>,','');">
						<a href="javascript:" onclick="return false">$<% = formatNumber(dblBudgetAmount-dblLineItemCosts,2) %></a>
					</td>			
					<td align="center">
						<% = oFunc.YNText(rs("bolRequired")) %>
					</td>
					<td align="center">
						<input type="checkbox" name="bolClosed<%=intCount%>" <% = oHtml.IIF(rs("bolClosed"),"checked","") %> value="1" onChange="jfUpdateList('<%=intcount%>');" ID="Checkbox1">
					</td>
					<td align="center"> 
					<%
					
					mySql = "SELECT	s.szFIRST_NAME, s.szLAST_NAME, oi.intQty * oi.curUnit_Price + oi.curShipping AS budgetTotal, oi.intOrdered_Item_ID, " & _ 
							"	(SELECT	SUM(li.intQuantity * li.curUnit_Price + li.curShipping) " & _ 
							"	FROM	tblLine_Items li " & _ 
							"	WHERE	li.intOrdered_Item_ID = oi.intOrdered_Item_ID) AS LiTotal, oi.intILP_ID, s.intSTUDENT_ID, ilp.intContract_Guardian_ID " & _ 
							"FROM	tblOrdered_Items oi INNER JOIN " & _ 
							"	tblSTUDENT s ON oi.intStudent_ID = s.intSTUDENT_ID INNER JOIN " & _ 
							"	tblILP ilp ON ilp.intILP_ID = oi.intILP_ID " & _
							"WHERE	(oi.intClass_Item_ID = " & rs("intClass_Item_ID") & ") " & _
							" ORDER BY s.szLAST_NAME, s.szFIRST_NAME "
							
					myRs.Open mySql, oFunc.FPCScnn
					
					myTable = "<table cellpadding='2'><tr class='navywhite8' style='font-weight:bold;'>" & _
							  "<td>Student Name</td>" & _
							  "<td>Budget</td>" & _
							  "<td>Spent</td>" & _
							  "<td>Balance</td>" & _
							  "<td>Links</td>" & _
							  "<td>Include</td></tr>"
					
					sVariance = ""	 
					
					' Used to gather checked students. A checked student will participate in
					' data updates.  Unchecked student records are not effected.
					myCheckList = myCheckList & "<input type='hidden' id='StudentsList" & intCount & "' name='StudentsList" & intCount & "' value=',"
					 
					do while not myRs.EOF
						if myRs("LiTotal") & "" = "" then 
							 LiTotal = 0
						else
							LiTotal = cdbl(myRs("LiTotal"))
						end if
						
						budgetTotal = cdbl(myRs("budgetTotal"))
						sChecked = " checked "
						
						if dblBudgetAmount <> budgetTotal then
							cssBudgetTotal = "TableHeaderOrange"
							sVariance = "<BR><span class='svplain8' title='One or more student budgets are different than the class budget.'>BUDGET VARIANCE!</span>"						
							sChecked = ""
						else
							cssBudgetTotal = "tablecell"
						end if
					
						if dblLineItemCosts <> LiTotal then
							cssLiTotal = "TableHeaderOrange"
							sVariance = "<BR><span class='svplain8' title='One or more student budgets are different than the class budget.'>BUDGET VARIANCE!</span>"													
							sChecked = ""
						else
							cssLiTotal = "tablecell"							
						end if
						
						myTable = myTable & "<tr><td class='tablecell'>" & myRs("szLAST_NAME") & ", " & myRs("szFIRST_NAME") & _
											"</td><td align='right' class='" & cssBudgetTotal & "'>$" &  formatNumber(budgetTotal,2) & "</td>" & _
											"<td align='right' class='" & cssLiTotal & "'>$" &  formatNumber(LiTotal,2) & "</td>" & _
											"<td align='right' class='tablecell'>$" & formatNumber((budgetTotal - LiTotal),2) & "</td>" & _
											"<td  align=center class='tablecell'> " & _ 
											"			<a href=""javascript:"" title=""View Class/Schedule for '" & replace(rs("szClass_Name") & "","'","\'") & "'"" onclick=""jfViewClass('" & rs("intClass_ID") & "','" & rs("intInstructor_ID") & "','" & rs("intInstruct_Type_Id") & "','" & myRs("intContract_Guardian_ID") & "','" & rs("intGuardian_Id") & "');""> " & _ 
											"			 C</a>  " & _ 
											"			<a href=""javascript:"" title=""View ILP"" onclick=""jfViewILP('" & myRs("intILP_ID") & "','" & replace(rs("szClass_Name") & "","'","\'") & "','" & replace(rs("TeacherName") & "" ,"'","\'") & "','" & myRs("intContract_Guardian_ID") & "');""> " & _ 
											"			I</a>						  " & _ 
											"			<a href=""javascript:"" title=""View Goods/Services"" onclick=""jfViewItem('" & myRs("intILP_ID") & "','" & myRs("intStudent_ID") & "','" & myRs("intOrdered_Item_ID") & "','" & rs("intItem_Group_ID") & "','" & replace(replace(rs("szClass_Name") & "","'","\'"),"&"," and ") & "','" & rs("intPOS_SUBJECT_ID") & "');"">  " & _ 
											"			GS</a> " & _ 
											"			<a href=""javascript:"" title=""View Packet"" onclick=""jfViewPacket('" & myRs("intStudent_ID") & "');""> " & _ 
											"			 P</a> " & _ 
											"			 <a href=""javascript:"" title=""Printable Forms"" onclick=""jfViewPrint('" & myRs("intStudent_ID") & "');""> " & _ 
											"			 Prt</a> " & _ 
											"</td> " & _
											"<td align='center' class='tablecell'><input type='checkbox' name='Students" & intCount & "' value='" & myRs("intStudent_ID") & "' " & sChecked & " onChange=""jfChangedObj('StudentsList" & intCount & "','" & myRs("intStudent_ID") & "');"">" & _
											"</td></tr>"
						
						' do not include student id if the students budget is different 
						' than the class budget					
						if sChecked <> "" then
							myCheckList = myCheckList & myRs("intStudent_ID") & ","
						end if
						
						myRs.MoveNext
					loop
					
					myCheckList = myCheckList & "'>"
					
					myRs.Close
					myTable = myTable & "</table>"
					response.Write oHtml.ToolTip("<a href='#'>View Students</a>" & sVariance, myTable, True, "Student List for Class Item ID: " & rs("intClass_Item_ID"),false,"ToolTipWhite","","",true,true)
					'response.Write myTable
					%>
					</td>
					<td align=center >
						<select name="approved<% = intCount%>" onChange="jfUpdateList('<%=intcount%>');" ID="Select1">
							<%
								if rs("bolApproved") then
									strApprovedStatus = "1"
								elseif rs("bolApproved") = false then
									strApprovedStatus = "0"
								else
									strApprovedStatus = "2"
								end if
								 
								response.Write oFunc.makeList("2,1,0","b-pend,b-appr,b-rejc",strApprovedStatus)
							%>
						</select>
					</td>
					<td>
						<textarea cols=20 rows=1 wrap=virtual name="denied<% = intCount%>" onChange="jfUpdateList('<%=intcount%>');" onfocus="this.rows=4;" onblur="this.rows=1;" onKeyDown="jfMaxSize(511,this);" ID="Textarea1"><% = rs("szComments") %></textarea>
					</td>
				</tr>
				<tr id="LineItem<%=intCount%>" style="display:none;" >
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
						if intLIIDLast <> clng(rsLI("intClass_Line_Item_ID")) then 
							strLIList = strLIList & intCount2 & ","
							intLIIDLast = rsLI("intClass_Line_Item_ID")
						end if
						%>
				<tr>
					<td>
						<INPUT type="hidden" name="intClass_Line_Item_ID<%=intCount2%>" value="<% = rsLI("intClass_Line_Item_ID")%>" ID="Hidden6">
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
				<input type="hidden" name="CountTwoToOne<% = intCount2 %>" value="<% = intCount %>" ID="Hidden5">				
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
					<input type="hidden" id="liGroupList<%=intCount%>" name="liGroupList<%=intCount%>" value="<% = strLIList %>" ID="Hidden9">
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
				<INPUT type="hidden" name="intClass_Item_IDLI<%=intCount2%>" value="<%=rs("intClass_Item_ID")%>" ID="Hidden11">
			</table>
		</td>
	</tr>
	<input type="hidden" name="CountTwoToOne<% = intCount2 %>" value="<% = intCount %>">
<%
					rs.MoveNext
					intCount = intCount + 1
					intCount2 = intCount2 + 1
					strLIList =""
				loop
				response.Write myCheckList
%>
			</table>
		</td>
	</tr>
	<%
		set rsLI = nothing
		set myRs = nothing
	else ' no records returned
	
	end if	
	
	rs.Close
	set rs = nothing
end if 

%>
<input type=hidden name="intCount" value="<%=intCount%>" ID="Hidden10">
<input type=hidden name="intCount2" value="<%=intCount2%>" ID="Hidden12">
</table>
<% = oHtml.ToolTipDivs %>
</form>
</body>
</html>
<%

call oFunc.CloseCN()
set oFunc = nothing
set oHtml = nothing

function PrintHeader
%>
				<tr class="NavyWhite8" style="font-weight: bold;">
					<td>
						Vendor
					</td>
					<td>
						Good/Service
					</td>
					<td>
						Class Item ID
					</td>
					<td>
						Item Type
					</td>
					<td>
						Desc
					</td>					
					<td>
						Total
					</td>	
					<td>
						Balance
					</td>					
					<td>
						Required
					</td>
					<td align="center">
						Closed
					</td>
					<td align="center">
						Students
					</td>
					<td align="center">
						Aprv
					</td>
					<td>
						Comments
					</td>
				</tr>
<%
end function 

sub vbsSaveChanges
	' If a class budget is closed, has a approval status change or has a comment change
	' then the class item is updated as well as all associated ordered items.
	dim update
	dim bolClosed, rsChildren
	arList = split(request("updatelist"),",")
	
	oFunc.BeginTransCN
	for i = 0 to ubound(arList)
		if arList(i) <> "" then
			bolClosed = oHtml.IIF(request.Form("bolClosed"&arList(i)) <> "","1","NULL")
			
			StudentsList = request("StudentsList" & arList(i))
			
			' Remove leading comma
			if left(StudentsList,1) = "," then
				StudentsList = right(StudentsList,len(StudentsList)-1)
			end if
			
			' Remove trailing comma
			if right(StudentsList,1) = "," then
				StudentsList = left(StudentsList,len(StudentsList)-1)
			end if
			
			if request.Form("Approved"&arList(i)) = "1" then
				' approved budget
				update = "update tblClass_Items set bolApproved = 1, " & _
						 "szComments = '" & oFunc.escapeTick(request.Form("denied"&arList(i))) & "', " & _
						 "dtApproval_Changed = CURRENT_TIMESTAMP " & _
						 ",dtModify = CURRENT_TIMESTAMP, " & _
						 "bolClosed = " & bolClosed & ", " & _
						 "szUSER_MODIFY ='" & oFunc.EscapeTick(session.Contents("strUserID")) & "' " & _ 
						 "where intClass_Item_ID = " & request.Form("intClass_Item_ID"&arList(i))							
						 
				update2 = "update tblOrdered_Items set bolApproved = 1, " & _
							"szDeny_Reason = '" & oFunc.escapeTick(request.Form("denied"&arList(i))) & "', " & _
							"dtApproval_Changed = CURRENT_TIMESTAMP " & _
							",dtModify = CURRENT_TIMESTAMP, " & _
							"bolClosed = " & bolClosed & ", " & _
							"szUSER_MODIFY ='" & oFunc.EscapeTick(session.Contents("strUserID")) & "' " & _ 
							"where intClass_Item_ID = " & request.Form("intClass_Item_ID"&arList(i)) & _
							" AND intStudent_ID in (" & StudentsList & ")"	
			elseif request.Form("Approved"&arList(i)) = "0" then
				' rejected budget
				update = "update tblClass_Items set bolApproved = 0, " & _
						"szComments = '" & oFunc.escapeTick(request.Form("denied"&arList(i))) & "', " & _
						"dtApproval_Changed = CURRENT_TIMESTAMP," & _
						"bolClosed = 1, " & _
						"szUSER_MODIFY ='" & oFunc.EscapeTick(session.Contents("strUserID")) & "'," & _ 
						"dtModify = CURRENT_TIMESTAMP " & _
						"where intClass_Item_ID = " & request.Form("intClass_Item_ID"&arList(i))
						
				update2 = "update tblOrdered_Items set bolApproved = 0, " & _
							"szDeny_Reason = '" & oFunc.escapeTick(request.Form("denied"&arList(i))) & "', " & _
							"dtApproval_Changed = CURRENT_TIMESTAMP," & _
							"bolClosed = 1, " & _
							"szUSER_MODIFY ='" & oFunc.EscapeTick(session.Contents("strUserID")) & "'," & _ 
							"dtModify = CURRENT_TIMESTAMP " & _
							"where intClass_Item_ID = " & request.Form("intClass_Item_ID"&arList(i)) & _
							" AND intStudent_ID in (" & StudentsList & ")"
				' no need to have line items since the budget is rejected so do some clean
				' up to insure this is true				
				'delete = "delete from tblLine_Items where intOrdered_Item_ID = " & request.Form("intOrdered_Item_ID"&arList(i)) 
				'oFunc.ExecuteCN(delete)
				
			elseif request.Form("Approved"&arList(i)) = "2" then
				update = "update tblClass_Items set bolApproved = NULL, " & _
						 "szComments = '" & oFunc.escapeTick(request.Form("denied"&arList(i))) & "', " & _
						 "dtApproval_Changed = CURRENT_TIMESTAMP, " & _
						 "bolClosed = " & bolClosed & ", " & _
						 "szUSER_MODIFY ='" & oFunc.EscapeTick(session.Contents("strUserID")) & "'" & _ 
						 ",dtModify = CURRENT_TIMESTAMP " & _
						 "where intClass_Item_ID = " & request.Form("intClass_Item_ID"&arList(i)) 
						
				update2 = "update tblOrdered_Items set bolApproved = NULL, " & _
							"szDeny_Reason = '" & oFunc.escapeTick(request.Form("denied"&arList(i))) & "', " & _
							"dtApproval_Changed = CURRENT_TIMESTAMP, " & _
							"bolClosed = " & bolClosed & ", " & _
							"szUSER_MODIFY ='" & oFunc.EscapeTick(session.Contents("strUserID")) & "'" & _ 
							",dtModify = CURRENT_TIMESTAMP " & _
							"where intClass_Item_ID = " & request.Form("intClass_Item_ID"&arList(i)) & _
							" AND intStudent_ID in (" & StudentsList & ")" 
						 
			end if 
			
			oFunc.ExecuteCN(update)
			
			if StudentsList & "" <> "" then
				oFunc.ExecuteCN(update2)	
			end if					
		end if		' ends if arList(i) <> "" then
	next
	oFunc.CommitTransCN
end sub

sub vbsSaveLineItems
	arList = split(request.Form("LineItemsChanged"),",")
	dim cmd
	dim dtCheck
	dim dtReciept
	dim dblUnitPrice
	dim intQty
	dim dblShipping
	
	set rsChild = server.CreateObject("ADODB.RECORDSET")
	rsChild.CursorLocation = 3
	
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
			
			StudentsList = request.Form("StudentsList" & request.Form("CountTwoToOne" & arList(i)))
			
		response.Write "<BR>" & StudentsList
			
			' Remove leading comma
			if left(StudentsList,1) = "," then
				StudentsList = right(StudentsList,len(StudentsList)-1)
			end if
			response.Write "<BR>" & StudentsList
			' Remove trailing comma
			if right(StudentsList,1) = "," then
				StudentsList = left(StudentsList,len(StudentsList)-1)
			end if
			response.Write "<BR>" & StudentsList
			if request.Form("intClass_Line_Item_ID" & arList(i)) <> "" then
				cmd = "UPDATE    tblClass_Line_Items " & _ 
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
						"WHERE     intClass_Line_Item_ID = " & request.Form("intClass_Line_Item_ID" & arList(i))
				oFunc.ExecuteCN(cmd)
				
				cmd = "UPDATE   li  " & _ 
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
						" from tblLine_Items li INNER JOIN tblOrdered_Items oi ON li.intOrdered_Item_ID = oi.intOrdered_Item_ID " & _
						"WHERE     li.intClass_Line_Item_ID = " & request.Form("intClass_Line_Item_ID" & arList(i)) & _
						" AND oi.intStudent_ID in (" & StudentsList & ")" 
					response.Write cmd
				if StudentsList & "" <> "" then
					oFunc.ExecuteCN(cmd)
				end if			
			else
				dim myClassLineItemID
				
				cmd = "INSERT INTO tblClass_Line_Items " & _ 
						" (intClass_Item_ID, dtLine_Item, szLine_Item_desc, curUnit_Price, " & _ 
						" intQuantity, curShipping, szPO_Number, szInvoice_Number, szCheck_Number,  " & _ 
						" szPayee, dtCheck, dtReciept, dtCREATE, szUSER_CREATE) " & _ 
						"VALUES     (" & request.Form("intClass_Item_IDLI" & arList(i)) & "," & _
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
				oFunc.ExecuteCN(cmd)
				myClassLineItemID = oFunc.GetIdentity								
				
				sql = "select intOrdered_Item_ID from tblOrdered_Items where intClass_Item_ID = " & request.Form("intClass_Item_IDLI"&arList(i)) & _
					  " AND intStudent_ID in (" & StudentsList & ")" 
					 response.Write sql 
				rsChild.Open sql, oFunc.FPCScnn
				
				if rsChild.RecordCount > 0 then
					do while not rsChild.eof	
						cmd = "INSERT INTO tblLine_Items " & _ 
								" (intOrdered_Item_ID, dtLine_Item, szLine_Item_desc, curUnit_Price, " & _ 
								" intQuantity, curShipping, szPO_Number, szInvoice_Number, szCheck_Number,  " & _ 
								" szPayee, dtCheck, dtReciept, dtCREATE, szUSER_CREATE, intClass_Line_Item_ID) " & _ 
								"VALUES     (" & rsChild("intOrdered_Item_ID") & "," & _
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
								"'" & oFunc.EscapeTick(session.Contents("strUserID")) & "', " & myClassLineItemID & ") "
						rsChild.MoveNext
						oFunc.ExecuteCN(cmd)
					loop
				end if  ' ends if rsChild.RecordCount > 0 then	
				rsChild.Close
			end if		' if request.Form("intClass_Line_Item_ID" & arList(i)) <> "" then	
		end if			'if arList(i) <> "" then
	next	
	set rsChild = nothing
end sub
%>
