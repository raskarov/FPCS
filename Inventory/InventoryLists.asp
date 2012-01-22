<%@ Language=VBScript %>
<%
dim oFunc				'wsc object
'Create object containing all of our FPCS functions
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

' Page Header Setup
Session.Value("strTitle") = "Inventory Lists"
Session.Value("strLastUpdate") = "19 July 2006"

if request("simpleHeader") <> "" then 
	Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")
else
	Server.Execute(Application.Value("strWebRoot") & "includes/Header.asp")
end if


dim rsHold, rsCO, sql, intFamily_ID

intFamily_ID = request("myFamilyId") 


%>
<script language="javascript">
	function jfAddInventory(pDetailID,pCatId){
		var winInventory;
		var url;
		url = "<%=Application.Value("strSSLWebRoot")%>Inventory/InventoryAdmin.asp?refreshParent=true&simpleHeader=true&InventoryDetailID=" + pDetailID + "&panel=new";
		winInventory = window.open(url,"winInventory","width=950,height=500,scrollbars=yes,resizable=yes");
		winInventory.moveTo(0,0);
		winInventory.focus();
	}
</script>
<form action="./InventoryLists.asp" method="get">
<input type="hidden" name="simpleHeader" value="<% = request("simpleHeader") %>">
<input type="hidden" name="myFamilyId" id="myFamilyId" value="<% = intFamily_ID %>">
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
					<td class="inventoryOption" nowrap>
						<a href="InventoryAdmin.asp?panel=new" class="White8Verd">Add New Item</a>
					</td>
					<% end if %>
					<td>
						&nbsp;
					</td>
					<td class="inventoryOption" nowrap>
						<a href="InventoryAdmin.asp?panel=search" class="White8Verd">Search For Item</a>
					</td>
					<td>
						&nbsp;
					</td>
					<td class="inventoryOptionSelected" nowrap>
						<a href="InventoryLists.asp?" class="White8Verd">Inventory Lists</a>
					</td>
					<td style="width:100%;">
						&nbsp;
					</td>
				</tr>
			</table>
		</td>
    </tr>      
<%
if oFunc.IsAdmin or oFunc.IsTeacher then
%>
	<tr>
		<td style="width:100%;">
			<br>
			<table cellpadding=2 cellspacing=0 ID="Table5">
				<tr>
					<td class="svplain9" nowrap>
						Select a Family	
					</td>
					<td>
						<select name="intFamily_ID" id="intFamily_ID" class="InventorySelect" onchange="document.getElementById('myFamilyId').value=this.value;this.form.submit();">
							<option value=""></option>
							<%
							
							if oFunc.IsTeacher then
								sJoin = "INNER JOIN  tblENROLL_INFO ON s.intSTUDENT_ID = tblENROLL_INFO.intSTUDENT_ID "
								sWhere = " AND (tblENROLL_INFO.sintSCHOOL_YEAR = " & session.Contents("intSchool_Year") & ") AND  " & _ 
										"	(tblENROLL_INFO.intSponsor_Teacher_ID = " & session.Contents("instruct_id") & ") "
							end if
							
							sql = "SELECT DISTINCT f.intFamily_ID, f.szFamily_Name + ', ' +f.szDesc as name " & _ 
									"FROM	tblFAMILY f INNER JOIN " & _ 
									"	tblSTUDENT s ON f.intFamily_ID = s.intFamily_ID " & sJoin & _ 
									"WHERE	EXISTS " & _ 
									"		(SELECT     'x' " & _ 
									"			FROM	INVENTORY_CHECKED_OUT ico " & _ 
									"			WHERE	ico.StudentID = s.intStudent_ID AND ico.DateCheckedIn IS NULL) " & sWhere & " OR " & _ 
									"		EXISTS " & _ 
									"		(SELECT     'x' " & _ 
									"			FROM	INVENTORY_DETAILS id " & _ 
									"			WHERE	id.HeldForStudentID = s.intStudent_ID) " & sWhere & _ 
									"ORDER BY name "
							response.Write oFunc.MakeListSQL(sql,"intFamily_ID","name",intFamily_ID)
							%>
						</select>
					</td>
					<td class="svplain8">
						<nobr>&nbsp;<b>OR</b>&nbsp;</nobr>
					</td>
					<td class="svplain9" nowrap>
						Select a Student	
					</td>
					<td>
						<select name="intFamily_ID2" id="Select1" class="InventorySelect" onchange="document.getElementById('myFamilyId').value=this.value;this.form.submit();">
							<option value=""></option>
							<%
							
							if oFunc.IsTeacher then
								sJoin = "INNER JOIN  tblENROLL_INFO ON s.intSTUDENT_ID = tblENROLL_INFO.intSTUDENT_ID "
								sWhere = " AND (tblENROLL_INFO.sintSCHOOL_YEAR = " & session.Contents("intSchool_Year") & ") AND  " & _ 
										"	(tblENROLL_INFO.intSponsor_Teacher_ID = " & session.Contents("instruct_id") & ") "
							end if
							
							sql = "SELECT DISTINCT f.intFamily_ID, s.szLAST_NAME + ', ' +s.szFIRST_NAME as name " & _ 
									"FROM	tblFAMILY f INNER JOIN " & _ 
									"	tblSTUDENT s ON f.intFamily_ID = s.intFamily_ID " & sJoin & _ 
									"WHERE	EXISTS " & _ 
									"		(SELECT     'x' " & _ 
									"			FROM	INVENTORY_CHECKED_OUT ico " & _ 
									"			WHERE	ico.StudentID = s.intStudent_ID AND ico.DateCheckedIn IS NULL) " & sWhere & " OR " & _ 
									"		EXISTS " & _ 
									"		(SELECT     'x' " & _ 
									"			FROM	INVENTORY_DETAILS id " & _ 
									"			WHERE	id.HeldForStudentID = s.intStudent_ID) " & sWhere & _ 
									"ORDER BY name "
							response.Write oFunc.MakeListSQL(sql,"intFamily_ID","name",intFamily_ID)
							%>
						</select>
					</td>
				</tr>
			</table>
		</td>
	</tr>
<%
end if


if intFamily_ID = "" then intFamily_ID = session.Contents("intFamily_ID")
set rsLists = server.CreateObject("ADODB.RecordSet")
rsLists.CursorLocation = 3

if intFamily_ID & "" <> "" then
	
	rsLists.Open "ts_InventoryLists " & intFamily_ID, oFunc.FPCScnn
	
	if rsLists.RecordCount > 0 then
		call HoldList(rsLists)		
	end if
	
	set rsLists = rsLists.NextRecordset()
	
	if rsLists.RecordCount > 0 then
		call CoList(rsLists)
	end if
	
	rsLists.Close
	set rsLists = nothing			
end if

call oFunc.CloseCN()

%>
	</table>
	</form>
</body>
</html>
<%

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Supporting Procedures below here
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
function HoldList(pRs)
%>
	<tr>
		<td style="width:100%;"	>
			<table ID="Table1" style="width:100%;" cellpadding="3">
				<tr>
					<td class="TableHeaderRed" colspan="10">
						<b>Items On Hold</b>
					</td>
				</tr>
				<tr>
					<td class="ltGray8">
						<b><nobr>Student On Hold For</nobr></b>
					</td>	
					<td class="ltGray8">
						<b><nobr>Item Category</nobr></b>
					</td>	
					<td class="ltGray8">
						<b><nobr>FPCS #</nobr></b>
					</td>
					<td class="ltGray8" style="width:100%;">
						<b>Description</b>
					</td>	
					<td class="ltGray8" align="center">
						<b>Cost</b>
					</td>
					<td class="ltGray8">
						<b>Location</b>
					</td>
					<td class="ltGray8">
						<b><nobr>Held Until</nobr></b>
					</td>				
				</tr>
<%
	do while not pRs.eof
		 if k mod 2 = 0 then myClass = "svplain8" else myClass = "ltGray8" 
%>	
				<tr onclick="jfAddInventory('<% = pRs("InventoryDetailID") %>','<% = pRs("InventoryCategoryID") %>');" style='cursor:pointer'>
					<td valign="top" class="<% = myClass %>">
						<nobr><% = pRs("StudentName") %></nobr>
					</td>
					<td valign="top" class="<% = myClass %>">
						<nobr><% = pRs("Name") %></nobr>
					</td>
					<td valign="top" class="<% = myClass %>">
						<nobr><% = pRs("SchoolControlNum") %></nobr>
					</td>
					<td valign="top" class="<% = myClass %>" style="width:100%;">
						<% = pRs("Description") %>
					</td>
					<td valign="top" class="<% = myClass %>" align="right">
						<nobr>$<% if pRs("Cost")&"" <> "" then response.Write FormatNumber(pRs("Cost"),2) else response.Write "0.00" %></nobr>
					</td>
					<td valign="top" class="<% = myClass %>">
						<nobr><% = pRs("Location") %></nobr>
					</td>
					<td valign="top" class="<% = myClass %>" align="right">
						<nobr><% = pRs("DateHoldEnd") %></nobr>
					</td>					
				</tr>
<%		
		k = k + 1
		pRs.MoveNext
	loop

%>				
			</table>
			<br>
		</td>
	</tr>
<%
end function

function CoList(pRs)
%>
	<tr>
		<td style="width:100%;"	>
			<table ID="Table2" style="width:100%;" cellpadding="3">
				<tr>
					<td class="TableHeaderBlue" colspan="10">
						<b>Checked Out Items</b>
					</td>
				</tr>
				<tr>
					<td class="ltGray8">
						<b><nobr>Checked Out By </nobr></b>
					</td>	
					<td class="ltGray8">
						<b><nobr>Item Category</nobr></b>
					</td>	
					<td class="ltGray8">
						<b><nobr>FPCS #</nobr></b>
					</td>
					<td class="ltGray8" style="width:100%;">
						<b>Description</b>
					</td>	
					<td class="ltGray8" align="center">
						<b>Cost</b>
					</td>
					<td class="ltGray8">
						<nobr><b>Checked Out</b></nobr>
					</td>
					<td class="ltGray8">
						<b><nobr>Due Date</nobr></b>
					</td>				
				</tr>
<%
	do while not pRs.eof
		  if k mod 2 = 0 then myClass = "svplain8" else myClass = "ltGray8" 
%>	
				<tr onclick="jfAddInventory('<% = pRs("InventoryDetailID") %>','<% = pRs("InventoryCategoryID") %>');"  style='cursor:pointer'>
					<td valign="top" class="<% = myClass %>">
						<nobr><% = pRs("StudentName") %></nobr>
					</td>
					<td valign="top" class="<% = myClass %>">
						<nobr><% = pRs("Name") %></nobr>
					</td>
					<td valign="top" class="<% = myClass %>">
						<nobr><% = pRs("SchoolControlNum") %></nobr>
					</td>
					<td valign="top" class="<% = myClass %>" style="width:100%;">
						<% = pRs("Description") %>
					</td>
					<td valign="top" class="<% = myClass %>" align="right">
						<nobr>$<% if pRs("Cost")&"" <> "" then response.Write FormatNumber(pRs("Cost"),2) else response.Write "0.00" %></nobr>
					</td>
					<td valign="top" class="<% = myClass %>" align="right">
						<nobr><% = pRs("DateCheckedOut") %></nobr>
					</td>
					<td valign="top" class="<% = myClass %>" align="right">
						<nobr><% = pRs("DateDue") %></nobr>
					</td>					
				</tr>
<%			
		k = k + 1
		pRs.MoveNext
	loop

%>				
			</table>
		</td>
	</tr>
<%
end function
%>    