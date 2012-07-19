<%@ Language=VBScript %>
<%
' TOGGLES SHOWING GOODS/SERVICES 
if session.Contents("strRole") <> "ADMIN" then
	response.write "<h1>Page Improperly Called.</h1>"
	response.end
end if

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Dimention Variables, make db Connection, print HTML header.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim intStudent_ID 

Session.Contents("strTitle") = "Approved Vendor List"
Session.Contents("strLastUpdate") = "05 May 2004"

if request("simpleHeader") = "" then
	Server.Execute(Application.Value("strWebRoot") & "includes/header.asp")
else
	Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")
end if

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()
ofunc.ResetSelectSessionVariables

if request.Form("updatelist") <> "" then
	call UpdateStatus(request.Form("updatelist"))
end if
%>
<script language=javascript>
	function jfAuthList(id){
		var winAuthAct;
		var strURL = "<%=Application.Value("strWebRoot")%>Forms/VIS/VendorAdmin.asp?bolWin=true&intVendor_ID="+id;
		winAuthAct = window.open(strURL,"winAuthAct","width=840,height=550,scrollbars=yes,resize=yes,resizable=yes");
		winAuthAct.moveTo(0,0);
		winAuthAct.focus();	
	}
	
	function jfUpdateList(id) {
		// if an item as been changed log it on;y once.  We will use this list
		// to determine which OI's should be modified
		if (document.main.updatelist.value.indexOf(","+id+",") == -1 ) {
			document.main.updatelist.value = document.main.updatelist.value + id + ",";
		}
	}	
</script>
<form name="main" action="vendorList.asp" method="post">
<input type="hidden" name="updatelist" value="">
<table width=100% ID="Table1">
	<tr>	
		<Td class=yellowHeader >
				&nbsp;<b>Vendor Status Manager</b>				
		</td>
	</tr>
	<tr>
		<td class="svplain8">
			<b>&nbsp; Select Vendor Status: 
				<select name="szVendor_Status_CD" ID="Select2">
					<option value="ALL">All</option>
					<option value="AP" <% if request("szVendor_Status_CD") = "AP" or  request("szVendor_Status_CD") = "" then response.Write " selected " %> >APPR & PEND</option>
					<option value="APR" <% if request("szVendor_Status_CD") = "APR" then response.Write " selected " %>>APPR & PEND & REMV </option>
					<%
						
						sql = "select szVendor_Status_CD from tblVendor_Status_Codes order by szVendor_Status_CD"
						response.Write oFunc.MakeListSQL(sql,"szVendor_Status_CD","szVendor_Status_CD", request("szVendor_Status_CD"))
					%>
				</select>
				<b>&nbsp;Vendor Type:</b>
				<select name="VendorType" ID="Select3">
					<%
						Response.Write oFunc.MakeList("1,2,3,4","Service Vendor,Goods Vendor,Both,Non-Profit",request("VendorType"))								
					%>
				</select>
				<input type="submit" value="Save/Query" class="NavSave" ID="Submit1" NAME="Submit1"></b>
		</td>
	</tr>
	<tr>
		<td bgcolor=f7f7f7>
			<table ID="Table2">
				<tr>	
					<Td class=gray valign=middle align=center>
						&nbsp;<B>Vendor Name</b>&nbsp;<BR>
						&nbsp;(click on name to view vendor profile)
					</td>
					<Td class=gray valign=middle align=center>
						&nbsp;<b>Status&nbsp;
					</td>	
					<Td class=gray valign=middle align=center>
						<b>Profile Verified<br>by Vendor
					</td>	
					<Td class=gray valign=middle align=center>
						&nbsp;<b>Address&nbsp;
					</td>	
					<Td class=gray valign=middle align=center>
						&nbsp;<b>Phone&nbsp;
					</td>		
					<Td class=gray valign=middle align=center>
						&nbsp;<b>Email&nbsp;
					</td>							
				</tr>
<%	
	'This section gives the classes for a student
	
if request("szVendor_Status_CD") = "AP" or  request("szVendor_Status_CD") = "" then
	sqlWhere = " and (vt.szVendor_Status_CD = 'APPR' or vt.szVendor_Status_CD = 'PEND') "
elseif request("szVendor_Status_CD") = "APR" then
	sqlWhere = " and (vt.szVendor_Status_CD = 'APPR' or vt.szVendor_Status_CD = 'PEND' or vt.szVendor_Status_CD = 'REMV') "
elseif request("szVendor_Status_CD") = "ALL" then 
	
elseif request("szVendor_Status_CD") <> "" then
	sqlWhere = " and vt.szVendor_Status_CD = '" & request("szVendor_Status_CD") & "' " 
end if

if request("VendorType") = "3" then
	' all vendors
	'sqlWhere = sqlWhere & " and (bolGoods_Vendor = 1 and bolService_Vendor = 1) "
elseif request("VendorType") = 2 then 
	' goods vendors
	sqlWhere = sqlWhere & " and (bolGoods_Vendor = 1) "
elseif request("VendorType") = 4 then
	sqlWhere = sqlWhere & " and (bolNonProfit = 1) "
else
	' service vendors
	sqlWhere = sqlWhere & " and (bolService_Vendor = 1) "
end if

set rsVendor = server.CreateObject("ADODB.RECORDSET")
rsVendor.CursorLocation = 3
sqlVendor = "SELECT     intVendor_ID, szVendor_Name, bolNonProfit, VendorAddress, VendorCity, VendorState, VendorZip,  " & _ 
			"szVendor_Phone, szVendor_Email, szVendor_Status_CD, bolGoods_Vendor, bolService_Vendor, bolProfile_Verified, statDate,dtContract_Start " & _ 
			"FROM         (SELECT     intVendor_ID, szVendor_Name,  v.bolNonProfit, " & _ 
			"(CASE isNull(v.szMail_City, 'A') WHEN 'A' THEN v.szVendor_Addr ELSE v.szMail_Addr END)  AS VendorAddress,  " & _ 
			"(CASE isNull(v.szMail_City, 'A') WHEN 'A' THEN v.szVendor_City ELSE v.szMail_City END) AS VendorCity,  " & _ 
			"(CASE isNull(v.szMail_City, 'A') WHEN 'A' THEN v.sVendor_State ELSE v.sMail_State END) AS VendorState, " & _ 
			"(CASE isNull(v.szMail_City, 'A') WHEN 'A' THEN v.szVendor_Zip_Code ELSE v.szMail_Zip_Code END) AS VendorZip,  " & _ 
			"szVendor_Phone, szVendor_Email, " & _ 
			"		(SELECT     TOP 1 szVendor_Status_CD " & _ 
			"		FROM          tblVendor_Status vs " & _ 
			"		WHERE      (intSchool_Year <= " & session.Contents("intSchool_Year") & ") AND (vs.intVendor_ID = v.intVendor_ID) " & _ 
			"		ORDER BY intSchool_Year DESC, intVendor_Status_ID DESC) AS szVendor_Status_CD, v.bolGoods_Vendor, v.bolService_Vendor, " & _ 
			"		(SELECT     TOP 1 bolProfile_Verified " & _ 
			"		FROM          tblVendor_Status vs " & _ 
			"		WHERE      (intSchool_Year <= " & session.Contents("intSchool_Year") & ") AND (vs.intVendor_ID = v.intVendor_ID) " & _ 
			"		ORDER BY intSchool_Year DESC, intVendor_Status_ID DESC) as bolProfile_Verified, " & _
			"		(SELECT     TOP 1 dtCreate " & _ 
			"		FROM          tblVendor_Status vs " & _ 
			"		WHERE      (intSchool_Year <= " & session.Contents("intSchool_Year") & ") AND (vs.intVendor_ID = v.intVendor_ID) " & _ 
			"		ORDER BY intSchool_Year DESC, intVendor_Status_ID DESC) as statDate, " & _
			"		(SELECT     TOP 1 dtContract_Start " & _ 
			"		FROM          tblVendor_Status vs " & _ 
			"		WHERE      (intSchool_Year <= " & session.Contents("intSchool_Year") & ") AND (vs.intVendor_ID = v.intVendor_ID) " & _ 
			"		ORDER BY intSchool_Year DESC, intVendor_Status_ID DESC) as dtContract_Start " & _
			" FROM	tblVendors v) vt " & _ 
			" WHERE 1 = 1 " & sqlWhere & _
			"ORDER BY szVendor_Name "
			
rsVendor.Open sqlVendor,Application("cnnFPCS")'oFunc.FPCScnn	

intColorCount = 0
if rsVendor.RecordCount > 0 then
		do while not rsVendor.EOF						
			if intColorCount mod 2 = 0 then
				strBgColor = " bgcolor=white " 
			else
				strBgColor = ""
			end if 
					
			if rsVendor("bolService_Vendor") and (rsVendor("dtContract_Start") <> "" or rsVendor("bolNonProfit")) _
				and (rsVendor("szVendor_Status_CD") = "APPR" or rsVendor("szVendor_Status_CD") = "PEND") then
				strCellColor = "TableHeaderGreen"
				strALink = " class='linkWht' "
			else
				strCellColor = "tableCell"
				strALink = ""
			end if
%>
		<tr <% = strBgColor %>>
			<Td class ="<% = strCellColor %>" valign=top  title="Status Reference Date: <% = rsVendor("statDate") %>"> 
				&nbsp;<a href="javascript:" <% = strALink %> onclick="jfAuthList('<% = rsVendor("intVendor_ID")%>');"><% = rsVendor("szVendor_Name") %></a>&nbsp;<br>
				<% if rsVendor("dtContract_Start") <> "" then %>
				&nbsp;&nbsp;Contract Start Date: <% = rsVendor("dtContract_Start") %>
				<% end if %>
				<!--&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="javascript:" onclick="jfAuthList('<% = rsVendor("intVendor_ID")%>');">View Authorized Actions</a>-->
			</td>	
			<Td class = "tableCell" valign=top> 
				<select name="Status<% = rsVendor("intVendor_ID") %>" ID="Select1" onChange="jfUpdateList('<%=rsVendor("intVendor_ID")%>');">
					<option value=""></option>
					<%
						sql = "select szVendor_Status_CD from tblVendor_Status_Codes order by szVendor_Status_CD"
						response.Write oFunc.MakeListSQL(sql,"szVendor_Status_CD","szVendor_Status_CD",rsVendor("szVendor_Status_CD"))
					%>
				</select>
			</td>	
			<td class="tableCell" align="center">
				<% = rsVendor("bolProfile_Verified") %>
			</td>		
			<td align=center class = "tableCell">
				<% = rsVendor("VendorAddress")%><br>
				<% = rsVendor("VendorCity")%>, <% = rsVendor("VendorState")%> <% = rsVendor("VendorZip")%> 
			</td>
			<Td class = "tableCell" valign=top> 
				&nbsp;<% = rsVendor("szVendor_Phone") %>&nbsp;
			</td>
			<Td class = "tableCell" valign=top> 
				&nbsp;<a href="mailto:<% = rsVendor("szVendor_Email") %>"><% = rsVendor("szVendor_Email") %></a>&nbsp;
			</td>
		</tr>
<%				rsVendor.MoveNext
			intColorCount = intColorCount + 1 
		loop	
	else
%>
				<tr>	
					<Td colspan=2 class=svplain8>
						&nbsp;No Vendors with the selected Status for the School Year <% = session.contents("intSchool_Year") %>.
					</td>
				</tr>
<%
		end if 
	rsVendor.Close
	set rsVendor = nothing	
	call oFunc.CloseCN
	set oFunc = nothing
%>			
			</table>
		</td>
	</tr>
</table>
</form>
<%
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")

sub UpdateStatus(pList)
	dim insert
	dim arList, rsc
	
	arList = split(pList,",")
	
	set rsc = server.CreateObject("ADODB.RECORDSET")
	rsc.CursorLocation = 3
		
	for i = 0 to ubound(arList)
		if arList(i) <> "" then									
			if request.Form("Status" & arList(i)) <> "" then
				sql = "select intVendor_Status_ID, szVendor_Status_CD, bolProfile_Verified from tblVendor_Status " & _
					" WHERE intVendor_ID = " & arList(i) & _
					" AND intSchool_Year = " & session.Contents("intSchool_Year")  & _
				    " ORDER BY intSchool_Year DESC,intVendor_Status_ID DESC "
				rsc.Open sql,Application("cnnFPCS")' oFunc.FpcsCnn
				
				if rsc.RecordCount > 0 then
					update = "update tblVendor_Status set szVendor_Status_CD = '" & request.Form("Status" & arList(i)) & "', " & _
							" dtModify = CURRENT_TIMESTAMP, szUser_Modify = '" & Session.Value("strUserID") & "' " & _
							" where intVendor_Status_ID = " & rsc("intVendor_Status_ID")
					oFunc.ExecuteCn(update)
				else
					insert = "insert into tblVendor_Status(intVendor_ID, intSchool_Year,szVendor_Status_CD, dtCreate, szUser_Create) " & _
							" values (" & _
							arList(i) & "," & _
							session.Contents("intSchool_Year") & "," & _
							"'" & request.Form("Status" & arList(i)) & "'," & _
							" CURRENT_TIMESTAMP, " & _
							"'" & session.Contents("strUserID") & "')"
					oFunc.ExecuteCN(insert)
				end if
				rsc.Close
			end if 
		end if
	next
	set rsc = nothing
end sub
%>