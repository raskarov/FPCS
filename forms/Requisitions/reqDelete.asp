<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		reqDelete.asp
'Purpose:	This script deletes a good or a service for tblOrdered_Items
'			and tblClass_Items
'Date:		05 JAN 2003
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim delAttrib
dim delItem
dim updateBudget

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

if ucase(request.QueryString("bolOrdered_Item")) = "TRUE" then
	' Deleting an item from a student	
	sql = "select 'x' from tblLine_Items " & _
		  " where tblLine_Items.intOrdered_Item_ID = " & request.QueryString("id")
	
	set rs = server.CreateObject("ADODB.RECORDSET")
	rs.CursorLocation = 3		    
	rs.Open sql,Application("cnnFPCS")'oFunc.FPCScnn
	if rs.RecordCount < 1 then
			delAttrib = "delete from tblOrd_Attrib " & _
						" where intOrdered_Item_ID = " & request.QueryString("id") & _
						" AND NOT EXISTS (select 'x' from tblLine_Items " & _
						"				  where tblLine_Items.intOrdered_Item_ID = tblOrd_Attrib.intOrdered_Item_ID) "
					    
			delItem = "delete from tblOrdered_Items " & _
					" where intOrdered_Item_ID = " & request.QueryString("id") & _
					" AND NOT EXISTS (select 'x' from tblLine_Items " & _
					"				  where tblLine_Items.intOrdered_Item_ID = tblOrdered_Items.intOrdered_Item_ID) "				    
		' Unlinks budget to ordered item
		updateBudget = "update tblBudget set intOrdered_Item_ID = NULL where intOrdered_Item_ID = " & request.QueryString("id")
		oFunc.ExecuteCN(updateBudget)
	end if
	rs.Close
	set rs = nothing
elseif ucase(request.QueryString("bolOrdered_Item")) = "FALSE" then
	' deleting an item from a class 
	delAttrib = "delete from tblClass_Attrib where intClass_Item_ID = " & request.QueryString("id")
	delItem = "delete from tblClass_Items where intClass_Item_ID = " & request.QueryString("id")
else
	response.Write "<h1>PAGE IMPROPERLY CALLED</h1>"
	response.End
end if

if delAttrib <> "" then
	oFunc.ExecuteCN(delAttrib)
	oFunc.ExecuteCN(delItem)
end if

Session.Value("strTitle") = "Delete an Item"
Session.Value("strLastUpdate") = "05 JAN 2003"
session.Value("simpleOnLoad") = "window.opener.location.reload();window.opener.focus();window.close();"
Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")
session.Value("simpleOnLoad") = ""
%>

<table width=100% height=100%>
	<tr>
		<td align=center valign=middle class=svPlain10>
			<b>
			<% if delAttrib <> "" then %>
			Item has been deleted.
			<% else %>
			Could not delete Item due to existing Line Items.
			<% end if %>
			</b><br><br>
			<center>
			<input type=button  value="Close Window" onclick="window.opener.focus();window.close();">
			</center>
		</td>
	</tr>	
</table>
</body>
</html>

<%
call oFunc.CloseCN
set oFunc = nothing
%>