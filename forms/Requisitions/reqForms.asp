<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		reqForms.asp
'Purpose:	Gives user with admin rights the ability to print
'			requisitions based on family,vendor and date.
'Date:		04 JAN 2004
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Quick Security Check
If ucase(session.Contents("strRole")) <> "ADMIN" then
	response.Write "<h1>Page illegally called.</h1>"
	response.End
end if

dim oFunc			 ' Main functions object
dim sql				 ' string to contain sql query commands
dim objRequest		 ' Contains the incoming form info via the request object
dim strObjValue		 ' Contains the value of an item in the request collection
dim strWhere		 ' Refines req search sequal
dim intCount		 ' Used to keep track of html table rows
dim curShippingTotal ' Total for all shipping costs
dim strSN 
dim strISBN
dim strBR
dim strIN
dim strCons

curShippingTotal = formatNumber(0,2)

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

Session.Value("strTitle") = "Requisition Form"
Session.Value("strLastUpdate") = "04 Jan 2004"

' Add "copy" background to all non-admin access
if ucase(session.Contents("strRole")) <> "ADMIN" then
	session.contents("strBGImagePath") = "../../images/copybg.gif"
else
	session.contents("strBGImagePath") = ""
end if

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

%>
<link rel="stylesheet" href="<% =  Application.Value("strWebRoot") %>CSS/printStyle.css">
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
		<Td colspan=2>
			<table cellpadding=0 cellspacing=0 align=right ID="Table2"><tr><td align=right><font face=arial size=2 color=white><% = formatDateTime(now(),2)%></font></td></tr></table>
			<div class="svplain8" style="font-color:white"><b>&nbsp;Requisition for Services/Equipment/Supplies</b></div>
		</td>					
	</tr>
</table>
<%

sql = "SELECT r.dtApproval_Changed, r.szFamily_Name, f.szDesc, f.szHome_Phone,  " & _
		"f.szEMAIL, u.szName_First + ' ' + u.szName_Last AS entered_by,  " & _
		" r.szVendor_Name, v.szVendor_Phone, v.szVendor_Fax,  " & _
		"v.szVendor_Email, v.szVendor_Contact, v.szVendor_Addr,  " & _
		"v.szVendor_City, v.sVendor_State,  " & _
		"v.szVendor_Zip_Code, r.intQty, r.Description, r.curUnit_Price, r.Shipping, " & _
		"r.szFIRST_NAME + ' ' + r.szLAST_NAME AS student, r.bar_code, r.item_number, r.page_number  " & _
		",r.Stock_Num, r.ISBN, r.Publisher, r.Copy_Year, r.Consumable " & _
		"FROM dbo.v_Requisitions r INNER JOIN " & _
		" dbo.tblFAMILY f ON r.intFamily_ID = f.intFamily_ID INNER JOIN " & _
		" dbo.tblVendors v ON r.intVendor_ID = v.intVendor_ID LEFT OUTER JOIN " & _
		" dbo.tblUsers u ON r.entered_by = u.szUser_ID  " & _
		"WHERE (r.intVendor_ID = " & request("intVendor_ID") & ") " & _
		"AND (r.intFamily_ID = " & request("intFamily_ID") & ")   " & _
		"AND (CONVERT(varchar, r.dtApproval_Changed, 101) = '" & request("date") & "')"
set rs = server.CreateObject("ADODB.RECORDSET")
rs.CursorLocation = 3 'adUseClient
rs.Open sql,oFunc.FPCScnn

if rs.RecordCount > 0 then
%>
<table width=100% cellspacing=0 cellpadding=2 bordercolor=c0c0c0 border=1>
	<tr class=svplain8>
		<td width=50% valign=top>
			<span class="svplain6"><b>Name of Family</b></span><br>
			<% = rs("szFamily_Name") & ": " & rs("szDesc") %>
		</td>
		<td width=50% colspan=2 valign=top title="<% = rs("szVendor_Addr") & " " & rs("szVendor_City") & ", " & rs("sVendor_State") & " " & rs("szVendor_Zip_Code")%>">
			<span class="svplain6"><b>Vendor/Provider</b></span><br>
			<% = rs("szVendor_Name") %>
		</td>
	</tr>
	<tr class=svplain8>
		<td width=50% valign=top>
			<span class="svplain6"><b>Phone Number</b></span><br>
			<% = rs("szHome_Phone") %>
		</td>
		<td width=25% valign=top title="Contact: <% = rs("szVendor_Contact")%>">
			<span class="svplain6"><b>Phone Number</b></span><br>
			<% = rs("szVendor_Phone") %>
		</td>
		<td width=25% valign=top>
			<span class="svplain6"><b>Fax Number</b></span><br>
			<% = rs("szVendor_Fax") %>
		</td>
	</tr>
	<tr class=svplain8>
		<td width=50% valign=top> 
			<span class="svplain6"><b>Email Address</b></span><br>
			<a href="mailto:<% = rs("szEMAIL") %>"><% = rs("szEMAIL") %></a>
		</td>
		<td width=50% colspan=2 valign=top>
			<span class="svplain6"><b>Email Address</b></span><br>
			<a href="mailto:<% = rs("szVendor_Email") %>"><% = rs("szVendor_Email") %></a>
		</td>
	</tr>
</table>
<BR>
<table width=100% cellpadding=4 cellspacing=0 border=1 bordercolor=c0c0c0  class=svplain8>
	<tr bgcolor=e6e6e6 style="font-color:white">
		<td align=center valign=top>
			<b>Page</b>
		</td>	
		<td align=center valign=top>
			<b>QTY</b>
		</td>	
		<td align=center valign=top>
			<b>Item #</b>
		</td>	
		<td align=center valign=top>
			<b>Description</b>
		</td>	
		<td align=center valign=top>
			<b>Unit<br>Price</b>
		</td>	
		<td align=center valign=top>
			<b>Total<br>Price</b>
		</td>
		<td align=center valign=top>
			<b>Student Name</b>
		</td>
		<td align=center valign=top title="Consumable?">
			<b>C</b>
		</td>
		<td align=center valign=top>
			<b>Bar Code</b>
		</td>
	</tr>
<%
	do while not rs.EOF
		strSN = ""
		strISBN = ""
		strBR = ""
		strIN = ""
		strCons = ""
		
		if rs("Stock_Num") <> "" then
			strSN = "S:" & rs("Stock_Num")
		end if
			
		if rs("ISBN") <> "" then
			strISBN = "I:" & rs("ISBN")
		end if
		
		if strSN <> "" and strISBN <> "" then
			strBR = "<BR>"
		end if
		
		if rs("item_number") <> "" then
			strIN = rs("item_number") & " " 
		end if
		
		if rs("Consumable") = "0" then
			strCons = "N "
		elseif rs("Consumable") = "1" then
			strCons = "Y "
		end if
		
		%>
	<tr title="Entered By: <%= rs("entered_by")%>">
		<td align=center valign=top>
			<% = rs("page_number") %>&nbsp;
		</td>
		<td align=center valign=top>
			<% = rs("intQty") %>
		</td>
		<td align=center valign=top>
			<% =  strIN & strSN & strBR & strISBN %>
		</td>
		<td align=center valign=top>
			<% = rs("Description") & " " & rs("Publisher") & " " & rs("Copy_Year")%>
		</td>
		<td align=center valign=top>
			$<% = formatNumber(rs("curUnit_Price"),2) %>
		</td>
		<td align=right valign=top>
			$<% = formatNumber(rs("intQty") * rs("curUnit_Price"),2)%>
		</td>
		<td align=center valign=top>
			<% = rs("student") %>
		</td>
		<td align=center valign=top>
			<% = strCons %>
		</td>
		<td align=center valign=top>
			<% = rs("bar_code") %>&nbsp;
		</td>
	</tr>
		<%
		curShippingTotal = formatNumber(curShippingTotal + rs("Shipping"),2)
		curTotal = formatNumber(curTotal + (rs("intQty") * rs("curUnit_Price")) + rs("Shipping"),2)
		rs.MoveNext
	loop
%>
	<tr>
		<td colspan=5 align=right>
			<b>Estimated Shipping Cost</b>
		</td>
		<td align=right>
			$<%=curShippingTotal%>
		</td>
		<td colspan=2>
		</td>
	</tr>
	<tr>
		<td colspan=5 align=right>
			<b>Total Amount of Requisition</b>
		</td>
		<td align=right>
			$<%=curTotal%>
		</td>
		<td colspan=2>
		</td>
	</tr>
</table>
<span class=svplain7>
<center>
All non consumable goods are the property of FPCS and must be 
returned when student no longer requires them for educational
purposes or student withdraws from FPCS (which ever comes first).
</center>
</span>
<br><BR>
<pre>
___________________________________    ________________________________
Parent Signature               Date    Teacher Signature           Date

</pre>
<% 
end if 
rs.Close
oFunc.CloseCN()
set rs = nothing
set oFunc = nothing
%>
</body>
</html>