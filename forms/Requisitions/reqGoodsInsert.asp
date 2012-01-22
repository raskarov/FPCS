<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		reqGoodsInsert.asp
'Purpose:	Insert/updates goods records for both parents and instructors.
'		
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
dim insert
dim update

Session.Value("strTitle") = "Insert/Updating Goods."
Session.Value("strLastUpdate") = "19 Aug 2002"
Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

If Request.Form("intStudent_ID") <> "" then
	intILP_ID = Request.Form("intILP_ID")
	intStudent_ID = Request.Form("intStudent_ID")
elseif Request.Form("intClass_ID") <> "" then
	intClass_ID = Request.Form("intClass_ID")
	intILP_ID = Request.Form("intILP_ID")
else
%>
	<font class=svplain10><B>The request to view this page is invalid.
	</b></font><br>
	<input type=button value="Home Page" onClick="window.location.href='<%=Application.Value("strWebRoot")%>';" id="btSmallGray" >
</body>
</html>
<%
	set oFunc = nothing
	Response.End
end if

intPrice = replace(request("curUnit_Price"),"$","")
intShipping = replace(request("curShipping"),"$","")
intQty = request("intQty")
if not isNumeric(intShipping) then
	intShipping = 0
end if

if not isNumeric(intPrice) then
	intPrice = 0
end if

if not isNumeric(intQty) then
	intQty = 0
end if

if (intPrice*intQty)+intShipping < 0 and ucase(session.Contents("strRole")) <> "ADMIN" then
	response.Write "<h4>Total cost must be greater than 0.<BR><input type=button value='Click Here' onClick='history.go(-1);'> to correct.</h4>"
	response.End		
end if
		
if intClass_ID <> "" then		
	if request("intClass_Item_ID") <> "" then		
		'Update records
		update = "update tblClass_Items set " & _
				 "intVendor_ID = " & request("intVendor_ID") & "," & _
				 "intItem_ID = " & request("intItem_ID") & "," & _
				 "curUnit_Price = convert(money,'" & request("curUnit_Price") & "')," & _
				 "intQTY = " & request("intQty") & "," & _
				 "szUser_Modify = '" & session.Contents("strUserID") & "', " & _
				 "bolRequired = " & request("bolRequired") & ", " & _
				 "curShipping = convert(money,'" & request("curShipping") & "') " & _
				 "where intClass_Item_ID = " & request("intClass_Item_ID")
		oFunc.ExecuteCN(update)
		
		' This recordset is needed because it is possible to add an item attrib in trefItem_Attrib
		' after an tblClass_Attrib record has been created.  If that is so then the 
		' new attribute will not be in tblClass_Attrib.  The corrisponding sql
		' to this recordset will check to see the existance of an attrib in tblClass_Attrib
		' and if it does not exist we will insert a new record
		set rsCheckAttrib = server.CreateObject("ADODB.RECORDSET")
		rsCheckAttrib.CursorLocation = 3
		
		arAttribList = split(request("strAttribList"),",")
		for i = 0 to ubound(arAttribList)
			sql = "select * from tblClass_Attrib " & _
				  " where intClass_Item_Id = " & request("intClass_Item_ID") & _
				  " and intItem_Attrib_Id = " &  arAttribList(i)
			rsCheckAttrib.Open sql, oFunc.FPCScnn
			
			if rsCheckAttrib.RecordCount > 0 then
				update = "update tblClass_Attrib set " & _
						"szValue = '" & oFunc.EscapeTick(request("attrib" & arAttribList(i))) & "'," & _
						"szUser_Modify = '" & session.Contents("strUserID") & "' " & _
						"where intClass_Item_ID = " & request("intClass_Item_ID") & _
						" and intItem_Attrib_ID = " & arAttribList(i)
				oFunc.ExecuteCN(update)		
			else
				insert = "insert into tblClass_Attrib (" & _
					 "intClass_Item_ID,intItem_Attrib_ID,szValue,intOrder,szUser_Create)" & _
					 "values (" & _
					 request("intClass_Item_ID") & "," & _
					 arAttribList(i) & "," & _
					 "'" & oFunc.EscapeTick(request("attrib" & arAttribList(i))) & "'," & _
					 "'" & (i+1) & "'," & _
					 "'" & session.Contents("strUserID") & "')"
				oFunc.ExecuteCN(insert)		
			end if	
			rsCheckAttrib.Close	
		next 
		set rsCheckAttrib = nothing
	else
		' Create New records
		insert = "insert into tblClass_Items ( " & _
				  "intClass_ID,intVendor_ID,intILP_Generic_ID,intItem_ID,intQty,bolRequired,curUnit_Price," & _
				  "curShipping,intSchool_Year,szUser_Create)" & _
				  " values (" & _
				  intClass_ID & "," & _
				  request("intVendor_ID") & ",'" & _
				  request("intILP_ID") & "'," & _
				  request("intItem_ID") & "," & _
				  request("intQty") & "," & _
				  request("bolRequired") & "," & _
				  "convert(money,'" & request("curUnit_Price") & "')," & _
				  "convert(money,'" & request("curShipping") & "')," & _
				  "'" & session.Contents("intSchool_Year") & "'," & _
				  "'" & session.Contents("strUserID") & "')"
				  
		oFunc.ExecuteCN(insert)
	
		intClass_Item_ID = oFunc.GetIdentity
	
		arAttribList = split(request("strAttribList"),",")
		for i = 0 to ubound(arAttribList)
			insert = "insert into tblClass_Attrib (" & _
					 "intClass_Item_ID,intItem_Attrib_ID,szValue,intOrder,szUser_Create)" & _
					 "values (" & _
					 intClass_Item_ID & "," & _
					 arAttribList(i) & "," & _
					 "'" & oFunc.EscapeTick(request("attrib" & arAttribList(i))) & "'," & _
					 "'" & (i+1) & "'," & _
					 "'" & session.Contents("strUserID") & "')"
			oFunc.ExecuteCN(insert)				 				 
		next 	
	end if	
elseif intStudent_ID <> "" then	
	if request("intOrd_Item_ID") <> "" then	
		' Update records
		update = "update tblOrdered_Items set " & _
				 "intVendor_ID = " & request("intVendor_ID") & "," & _
				 "intItem_ID = " & request("intItem_ID") & "," & _
				 "curUnit_Price = convert(money,'" & request("curUnit_Price") & "')," & _
				 "bolReimburse = '" & request("bolReimburse") & "'," & _
				 "intQTY = " & request("intQty") & "," & _
				 "szUser_Modify = '" & session.Contents("strUserID") & "', " & _
				 "curShipping = convert(money,'" & request("curShipping") & "') " & _
				 "where intOrdered_Item_ID = " & request("intOrd_Item_ID")
		oFunc.ExecuteCN(update)
		
		' This recordset is needed because it is possible to add an item attrib in trefItem_Attrib
		' after an tblOrd_Attrib record has been created.  If that is so then the 
		' new attribute will not be in tblOrdAttrib.  The corrisponding sql
		' to this recordset will check to see the existance of an attrib in tblOrd_Atrrib
		' and if it does not exist we will insert a new record
		set rsCheckAttrib = server.CreateObject("ADODB.RECORDSET")
		rsCheckAttrib.CursorLocation = 3
				
		arAttribList = split(request("strAttribList"),",")
		
		for i = 0 to ubound(arAttribList)
			sql = "select * from tblOrd_Attrib " & _
				  " where intOrdered_Item_Id = " & request("intOrd_Item_ID") & _
				  " and intItem_Attrib_Id = " &  arAttribList(i)
			rsCheckAttrib.Open sql, oFunc.FPCScnn
			
			if rsCheckAttrib.RecordCount > 0 then
				update = "update tblOrd_Attrib set " & _
						"szValue = '" & oFunc.EscapeTick(request("attrib" & arAttribList(i))) & "'," & _
						"szUser_Modify = '" & session.Contents("strUserID") & "' " & _
						"where intOrdered_Item_ID = " & request("intOrd_Item_ID") & _
						" and intItem_Attrib_ID = " & arAttribList(i)
				oFunc.ExecuteCN(update)	
				
			else			
				insert = "insert into tblOrd_Attrib (" & _
					 "intOrdered_Item_ID,intItem_Attrib_ID,szValue,intOrder,szUser_Create)" & _
					 "values (" & _
					 request("intOrd_Item_ID") & "," & _
					 arAttribList(i) & "," & _
					 "'" & oFunc.EscapeTick(request("attrib" & arAttribList(i))) & "'," & _
					 "'" & (i+1) & "'," & _
					 "'" & session.Contents("strUserID") & "')"
				oFunc.ExecuteCN(insert)	
			end if 	
			
			rsCheckAttrib.Close	
		next 
		
		' We need to make sure our budget reflects our actual ordered item
		sql = "select intBudget_ID from tblBudget where intOrdered_Item_ID = " & request("intOrd_Item_ID")
		rsCheckAttrib.open sql, oFunc.FPCScnn
		
		if rsCheckAttrib.recordcount > 0 then
			call vbsUpdateBudget(rsCheckAttrib("intBudget_ID"),request("intOrd_Item_ID"))
		else
			call vbsInsertBudget(request("intOrd_Item_ID"))		
		end if 
		
		rsCheckAttrib.Close	
		set rsCheckAttrib = nothing
	else
		'Create New Records
		oFunc.BeginTransCN
		insert = "insert into tblOrdered_Items ( " & _
				  "intVendor_ID,intILP_ID,intItem_ID,intStudent_ID,intQty,curUnit_Price," & _
				  "curShipping,bolReimburse,bolApproved,intSchool_Year,szUser_Create)" & _
				  " values (" & _
				  request("intVendor_ID") & "," & _
				  request("intILP_ID") & "," & _
				  request("intItem_ID") & "," & _
				  request("intStudent_ID") & "," & _
				  request("intQty") & "," & _
				  "convert(money,'" & request("curUnit_Price") & "')," & _
				  "convert(money,'" & request("curShipping") & "')," & _
				  "'" & request("bolReimburse") & "'," & _
				  "null," & _
				  "'" & session.Contents("intSchool_Year") & "'," & _
				  "'" & session.Contents("strUserID") & "')"
		oFunc.ExecuteCN(insert)
	
		intOrderd_Item_ID = oFunc.GetIdentity
	
		arAttribList = split(request("strAttribList"),",")
		for i = 0 to ubound(arAttribList)
			insert = "insert into tblOrd_Attrib (" & _
					 "intOrdered_Item_ID,intItem_Attrib_ID,szValue,intOrder,szUser_Create)" & _
					 "values (" & _
					 intOrderd_Item_ID & "," & _
					 arAttribList(i) & "," & _
					 "'" & oFunc.EscapeTick(request("attrib" & arAttribList(i))) & "'," & _
					 "'" & (i+1) & "'," & _
					 "'" & session.Contents("strUserID") & "')"
			oFunc.ExecuteCN(insert)				 				 
		next 	
		
		if request("intBudget_ID") <> "" then
		'	call vbsUpdateBudget(request("intBudget_ID"),intOrderd_Item_ID)
		else 
		'	call vbsInsertBudget(intOrderd_Item_ID)
		end if 
		
		oFunc.CommitTransCN
	end if
end if

sub vbsUpdateBudget(budgetID,OrderedItemID)
	dim update
	update = "update tblBudget set " & _
			 "intQTY = " & request("intQty") & "," & _
			 "curUnit_Price = " & replace(replace(request("curUnit_Price"),",",""),"$","") & "," & _
			 "curShipping = convert(money,'" & request("curShipping") & "'), " & _
			 "intOrdered_Item_ID = " & OrderedItemID & ", " & _
			 "szDesc = '" & oFunc.EscapeTick(request("attrib" & arAttribList(0))) & "' " & _
			 "WHERE intBudget_ID = " & budgetID
	oFunc.ExecuteCN(update)
end sub

sub vbsInsertBudget(OrderedItemID)
	set rsGetSILP = server.CreateObject("ADODB.RECORDSET")
	rsGetSILP.CursorLocation = 3
	sql = "select isf.intShort_ILP_ID from tblILP_Short_Form isf, tblILP i "  & _
		  "Where i.intILP_ID = " & intILP_ID & _
		  " and i.intShort_ILP_ID = isf.intShort_ILP_ID "
		  
	rsGetSILP.Open sql,oFunc.FPCScnn
	
	if rsGetSILP.RecordCount > 0 then
		dim insert
		dim intShort_ILP_ID
		intShort_ILP_ID = rsGetSILP("intShort_ILP_ID")
		insert = "insert into tblBudget(intBudget_Item_ID,intItem_ID,intShort_ILP_ID," & _
				 "szDesc,intQTY,curUnit_Price,curShipping,intOrdered_Item_ID,dtCreate,szUser_Create)" & _
				 "values (" & _
				 request("intItem_Group_ID") & "," & _
				 request("intItem_ID") & "," & _
				 intShort_ILP_ID & "," & _
				 "'" & oFunc.EscapeTick(request("attrib" & arAttribList(0))) & "'," & _
				 request("intQty") & "," & _
				 "convert(money,'" & request("curUnit_Price") & "')," & _
				 "convert(money,'" & request("curShipping") & "')," & _
				 OrderedItemID & ",'" & _
				 now() & "','" & _
				 session.Contents("strUserID") & "')"
		oFunc.ExecuteCN(insert)
	end if 
	rsGetSILP.Close
	set rsGetSILP = nothing
end sub
if request("hdnContinue") <> "" then
	response.Redirect(Application.Value("strWebRoot") & "forms/requisitions/reqGoods.asp?intStudent_ID=" & request("intStudent_ID") & _
					  "&intClass_ID=" & request("intClass_ID") & "&bolReimburse=" & request("bolReimburse") & _
					  "&intItem_ID=" & request("intItem_ID") & "&intVendor_ID=" & request("intVendor_ID")) & _
					  "&intItem_Group_ID=" & request("intItem_Group_ID") & "&intPOS_Subject_ID=" & request("intPOS_Subject_ID") & _
					  "&intILP_ID=" & request("intILP_ID")
	response.End					  
else

%>
<HTML>
<HEAD>
<script language=javascript>
	var strScript = window.opener.location.href;
	// Don't refresh if opener is reqApprovalAdmin.asp because it will mess up the
	// data the admins have not saved.
	if (strScript.indexOf("reqApprovalAdmin") < 1) {
		window.opener.location.reload();
	}
	window.opener.focus();
	window.close();
</script>
</HEAD>
<BODY>
</BODY>
</HTML>
<% end if %>
