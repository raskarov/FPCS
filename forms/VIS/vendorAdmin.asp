<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		vendorAdmin.asp
'Purpose:	Admin tool for adding/viewing/modifying Vendor information
'Date:		9-04-01
'Author:	Scott Bacon (ThreeShapes.com LLC)
'
'rev:		7-June-2003 BKM - removed javascript - added Unique Vendor Constraint
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim oFunc				'windows scripting component generalized functions
dim objRequest			'will contain either the FORM or QUERYSTRING object
dim mstrMessage			'Success for Fail message for new insert
dim mstrValidationError	'Error message indicating which items failed validation
dim bolPrint
dim FullRights

	Session.Value("strTitle")		= "Vendor Profile"
	Session.Value("strLastUpdate")	= "7 June 2003"
	
	strVendQString = "&intStudent_ID=" & request("intStudent_ID") & _
				 "&intItem_Group_ID=" & request("intItem_Group_ID") & _
				 "&intILP_ID=" & request("intILP_ID") & _
				 "&bolReimburse=" & request("bolReimburse")	& _
				 "&intItem_ID=" & request("intItem_ID")	& _
				 "&intOrd_Item_ID=" &  request("intOrd_Item_ID") & _
				 "&intClass_Item_ID=" & request("intClass_Item_ID") & _
				 "&viewing=" & request("viewing") & _
				 "&intClass_ID=" & request("intClass_ID") & _
				 "&strClassName=" & request("strClassName") & _
				 "&bolComplies=" & request("bolComplies") & _
				 "&intPOS_Subject_ID=" & request("intPOS_Subject_ID")
				 
	if request("bolPrint") <> "" then
		bolPrint = true
	else
		bolPrint = false
	end if	
	
	if request("xsuggestVendor") <> "" and request("intVendor_ID") <> "" then
		response.Write "<h1>Page Improperly Called</h1>"
		response.End 
	end if
	
	set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
	call oFunc.OpenCN()

	if oFunc.IsAdmin or (session.Contents("intVendor_ID") <> "" and session.Contents("intVendor_ID") & "" = request("intVendor_ID") & "") or _
		((oFunc.IsGuardian or oFunc.IsTeacher) and request("intVendor_ID") = "") or _
		(request("xsuggestVendor") <> "" and request("intVendor_ID") = "") then
		FullRights = true
	else
		FullRights = false
	end if

	if Request.Form.Count > 0 then
		set objRequest = Request.Form
	else
		set objRequest = Request.QueryString
	end if
	
	'**************************************
	'Create a Dictionary Collection based on
	'the contents of the Request Collection.
	'This enables us to change the values 
	'(Request.X is ReadOnly).
	dim marRequest
	Set marRequest = Server.CreateObject("Scripting.Dictionary")

	for each item in objRequest
		marRequest.Add item, objRequest(item)
	next 
	'**************************************

	' bolWin is used in this script to let us know this request was made
	' in the midst of the adding an item process.
	if Request("bolWin") <> "" or bolPrint then
		Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")
	elseif not oFunc.IsAdmin and not oFunc.IsGuardian and not oFunc.IsTeacher and request("xsuggestVendor") <> "" then 
		' Someone from the general public has accessed this page without logging in 
		' in order to suggest a vendor so we need to assign some variables
		' and use the header that doesn't force security
		Server.Execute(Application.Value("strWebRoot") & "includes/NonSecureHeader.asp")
		session.Contents("strRole") = "VENDOR"		
	else
		Server.Execute(Application.Value("strWebRoot") & "includes/header.asp")
	end if 
	
	intVendor_ID = objRequest("intVendor_ID")
	if objRequest("cmdSubmit") <> "" then
		mstrValidationError = vbfValidate(marRequest)

		if instr(1,mstrValidationError,"DID NOT SAVE!")  then
			Response.Write mstrValidationError
		else
			if mstrValidationError <> "" and oFunc.IsAdmin then
				Response.Write mstrValidationError
			end if
			
			call vbfInsert(marRequest)
			if request("bolWin") <> "" then
				' redirect back to goods and services page
				%>
				<script language=javascript>
					var qString = "<%=Application.Value("strWebRoot")%>forms/requisitions/reqGoods.asp?intVendor_ID=<% = intVendor_ID %>";
						qString += "<% = strVendQString%>";
						window.opener.location.href = qString;
						window.opener.focus();
						window.close();
				</script>
				<%			
				Server.Execute(Application.Value("strWebRoot") & "Includes/footer.asp")
				set oFunc = nothing
				response.End
			else
				Response.Write mstrMessage
			end if
		end if
	end if

	dim strVendorTitle	'Add or Modify a Vendor
	dim intBack			' Number of pages to go back in browser.  Varies if in edit mode or not.
	
	'default settings
	intBack = -1		' Non Edit mode setting
	strVendorTitle = "Add a New Vendor"	

	if mstrValidationError <> "" then
		'dimention local variables from the form object
		'we'll use these variables in the embedded HTML to populate the form
		for each item in objRequest
			execute("dim " & item)
			execute(item & " = """ & objRequest(item) & """")
		next 
	elseif intVendor_ID <> "" then
		'dimention local variables from tblVendors for the given vendor
		'we'll use these variables in the embedded HTML to populate the form
		dim rsVendor
		dim sqlVendor
		dim intCount
		dim item

		set rsVendor = Server.CreateObject("ADODB.RECORDSET")
		rsVendor.CursorLocation = 3
		'grab the Vendor info
		sqlVendor =	"SELECT     v.*, ct.szDesc as ChargeLabel, " & _ 
					"	(SELECT     TOP 1 upper(szVendor_Status_CD) " & _ 
					"	FROM          tblVendor_Status vs " & _ 
					"	WHERE      vs.intVendor_ID = v.intVendor_ID AND vs.intSchool_Year <= " & session.Contents("intSchool_Year") &  " " & _ 
					"	ORDER BY intSchool_Year DESC,intVendor_Status_ID DESC) AS szVendor_Status_Cd, " & _ 
					"	(SELECT     TOP 1 upper(dtContract_Start) " & _ 
					"	FROM          tblVendor_Status vs " & _ 
					"	WHERE      vs.intVendor_ID = v.intVendor_ID AND vs.intSchool_Year <= " & session.Contents("intSchool_Year") &  " " & _ 
					"	ORDER BY intSchool_Year DESC,intVendor_Status_ID DESC) AS dtContract_Start " & _
					"FROM         tblVendors v LEFT OUTER JOIN trefCharge_Type ct ON v.intCharge_Type_ID = ct.intCharge_Type_ID " & _ 
					"			  LEFT OUTER JOIN trefLocations l on v.intLocation_ID = l.intLocation_ID " & _
					"WHERE     intVendor_ID = " & intVendor_ID
					
		rsVendor.Open sqlVendor,oFunc.FPCScnn
		if not rsVendor.BOF and not rsVendor.EOF then
			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'' This for loop will dimension AND assign our student info variables
			'' for us. We'll use them later to populate the form.
			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
			intCount = 0
			for each item in rsVendor.Fields
				execute("dim " & rsVendor.Fields(intCount).Name)
				execute(rsVendor.Fields(intCount).Name & " = item")
				intCount = intCount + 1
			next	
			intBack = -2		'Edit mode setting	
			strVendorTitle = "View a Vendor"
			
			if request("xGoodService") <> "" then
				xGoodService = request("xGoodService")
			elseif bolService_Vendor and bolGoods_Vendor then
				xGoodService = 3
			elseif bolService_Vendor then
				xGoodService = 1				
			elseif bolGoods_Vendor then
				xGoodService = 2			
			else
				xGoodService = ""
			end if
		else
			'potentially display some error message	
		end if
		rsVendor.Close
		set rsVendor = nothing	
	else
		xGoodService = request("xGoodService") & request("intItem_Group_ID")
	end if 
	
	
	if (oFunc.IsGuardian or oFunc.IsTeacher) and intVendor_ID = "" then
		' guardians and teachers can only add a Goods vendor
		xGoodService = 2
	elseif request("xsuggestVendor") <> "" then
		xGoodService = 1
	elseif oFunc.IsGuardian or oFunc.IsTeacher and intVendor_ID <> "" then
		' these roles can not update profiles
		bolPrint = true
	end if 
	
	'set format for several fields
	'Unformat removes any frivolous characters then Reformat interjects the correct formating characters
	if not IsNull(szSSN) then szSSN	= oFunc.Reformat(oFunc.Unformat(szSSN, Array("-", " ")), Array("", 3, "-", 2, "-", 4))
	if not IsNull(szVendor_Zip_Code) then szVendor_Zip_Code	= oFunc.Reformat(oFunc.Unformat(szVendor_Zip_Code, Array("-", " ")), Array("", 5, "-", 4))
	if not IsNull(szMail_Zip_Code) then szMail_Zip_Code	= oFunc.Reformat(oFunc.Unformat(szMail_Zip_Code, Array("-", " ")), Array("", 5, "-", 4))
	'if not IsNull(szVendor_Phone) then szVendor_Phone = oFunc.Reformat(oFunc.Unformat(szVendor_Phone, Array("(", ")", "-", " ")), Array("(", 3, ") ", 3, "-", 4))
	'if not IsNull(szVendor_Fax) then szVendor_Fax = oFunc.Reformat(oFunc.Unformat(szVendor_Fax, Array("(", ")", "-", " ")), Array("(", 3, ") ", 3, "-", 4))
	'if not IsNull(szVendor_Phone_2) then szVendor_Phone_2 = oFunc.Reformat(oFunc.Unformat(szVendor_Phone_2, Array("(", ")", "-", " ")), Array("(", 3, ") ", 3, "-", 4))
	'**************************************************************
	'End Section:	Populate Form
	'**************************************************************
%>
<script language="javascript">
	function jfViewAuth(){
		var winAuth1;
		var URL = "vendorAuth.asp?intVendor_ID=<%=intVendor_ID%>&bolSimple=true";
		winAuth1 = window.open(URL,"winAuth1","width=800,height=500,scrollbars=yes,resizable=on");
		winAuth1.moveTo(0,0);
		winAuth1.focus();		
	}
</script>
<form action="vendorAdmin.asp" method=post name="main">
<input type=hidden name=count value="<% = objRequest("count") %>">

<input type="hidden" name="intStudent_ID" value="<% = request("intStudent_ID") %>" ID="Hidden5">
<input type="hidden" name="intItem_Group_ID" value="<% = request("intItem_Group_ID") %>" ID="Hidden6">
<input type="hidden" name="intILP_ID" value="<% = request("intILP_ID") %>" ID="Hidden7">
<input type="hidden" name="bolReimburse" value="<% = request("bolReimburse") %>" ID="Hidden15">
<input type="hidden" name="intItem_ID" value="<% = request("intItem_ID") %>" ID="Hidden10">
<input type="hidden" name="bolWin" value="<% = request("bolWin") %>" ID="Hidden9">
<input type=hidden name=intOrd_Item_ID value="<%=request("intOrd_Item_ID")%>" ID="Hidden16">
<input type=hidden name=intClass_Item_ID value="<%=request("intClass_Item_ID")%>" ID="Hidden17">
<input type=hidden name=viewing value="<%= request("viewing")%>" ID="Hidden18">
<input type=hidden name=intClass_ID value="<% = request("intClass_ID") %>" ID="Hidden21">
<input type=hidden name=strClassName value="<% = request("strClassName") %>" ID="Hidden22">
<input type=hidden name=bolComplies value="<% =request("bolComplies") %>" ID="Hidden23">
<input type=hidden name=intPOS_Subject_ID value="<%=request("intPOS_Subject_ID")%>" ID="Hidden25">
<input type=hidden name="xsuggestVendor" value="<% = request("xsuggestVendor") %>">
<table width=100%>
	<tr>	
		<Td class=yellowHeader >
				&nbsp;<b><% = strVendorTitle %></b> 
				<% if objRequest("intVendor_ID") <> "" AND session.Contents("strRole") = "ADMIN" then %>
				<select name="intVendor_ID" onchange="window.location.href='<% = Application.Value("strWebRoot") %>forms/VIS/vendorAdmin.asp?intVendor_ID='+this.value;">
				<option value="">Add a New Vendor</option>
				<%
					'dim sqlVendor
					sqlVendor = "SELECT intVendor_ID,  " & _ 
										" szVendor_Name + ' - ' + (SELECT     TOP 1 upper(szVendor_Status_CD) " & _ 
										"	FROM          tblVendor_Status vs " & _ 
										"	WHERE      vs.intVendor_ID = v.intVendor_ID AND vs.intSchool_Year <= " & session.Contents("intSchool_Year") &  " " & _ 
										"	ORDER BY intSchool_Year DESC,intVendor_Status_ID DESC) AS szVendor_Name " & _ 
										"FROM tblVendors v " & _ 
										"ORDER BY szVendor_Name "
					
					Response.Write oFunc.MakeListSQL(sqlVendor,"intVendor_ID","szVendor_Name", objRequest("intVendor_ID"))	
				%>
				</select>
				<% elseif oFunc.IsVendor then %>
				<input type="hidden" name="intVendor_ID" value="<% = session.Contents("intVendor_ID") %>">
			    <% end if %>
		</td>
	</tr>
	<%
	
	if (xGoodService & "" = "1" or xGoodService & "" = "3") and intVendor_ID & "" <> "" and dtContract_Start & "" = "" and (not bolNonProfit or bolNonProfit & "" = "") and not oFunc.IsAdmin then 
		call vbfStillToDo
	end if 
	%>
	<tr>
		<td bgcolor=f7f7f7>
			<table  style='width:800px;'>
				<tr>	
					<Td  class=svplain11>
							<b><i>Vendor Information</I></b> 
					</td>					
				</tr>
				<%	FRCss = "TableHeader"
					if (not bolPrint or FullRights) and request("intItem_Group_ID") = ""  then %>
					<% if request("xsuggestVendor") = "" then %>
				<tr>
					<td class="svplain8">
						<b>Vendor Type:</b>
						<select name="xGoodService" onchange="this.form.submit();">
							<%
								Response.Write oFunc.MakeList("0,1,2,3",",Service Vendor,Goods Vendor,Both",xGoodService)								
							%>
						</select>
						
					</td>									
				</tr>	
				<% end if %>							
				<tr>
					<td colspan="2" class="svplain8">
						<b><font color="orange" size="3">*</font> = required fields.&nbsp;&nbsp;
						Fields in <span class="required">&nbsp;&nbsp;&nbsp;</span> are seen by all users.</b>
					</td>
				</tr>
				<% 
						FRCss = "required"
					end if %>
				<input type=hidden name="bolGoods_Vendor" value="<% if xGoodService = "2" or xGoodService = "3" then response.Write "1" else response.Write "0"%>" ID="Hidden3">
				<input type=hidden name="bolService_Vendor" value="<% if xGoodService = "1" or xGoodService = "3" then response.Write "1" else response.Write "0"%>" ID="Hidden4">	
			</table>
		</td>
	</tr>
	<% 
	if (request("xGoodService") = "0" or request("xGoodService") = "") and (xGoodService = "0" or xGoodService = "") and request("intItem_Group_ID") = "" then
		' can't go any further until we know what kind of vendor we are working with
	%>
</table>
</form>
</BODY>
</HTML>
	<%
		set oFunc = nothing
		response.End
	end if
	%>
	<tr>
		<td>
			<table  style='width:800px;' ID="Table3">
				<tr>
					<td class="<% = FRCss %>" nowrap>
							&nbsp;Business Name <font color="orange" size="3">*</font><br>
							&nbsp;(Name as it appears on contract with FPCS/ASD)
					</td>
					<td class="<% = FRCss %>" nowrap>
							&nbsp;Vendor First Name <font color="orange" size="3">*</font>
					</td>
					<td class="<% = FRCss %>" nowrap>
							&nbsp;Vendor Last Name <font color="orange" size="3">*</font>
					</td>
					<% if oFunc.IsAdmin or IsVendor then%>		
					<td class="TableCell" align="center">
							&nbsp;Status&nbsp;
					</td>	
					<% end if
					
					   if oFunc.IsAdmin or IsVendor then%>											
					<td class="TableCell" style="width:100%;">
							&nbsp;Status Comments
					</td>		
					<% end if %>									
				</tr>		
				<tr>
					<td class="TableCell" nowrap>
						<% if bolPrint then %>
						<% = szVendor_Name %>&nbsp;
						<% else %>
						<input type=text name="szVendor_Name" value="<% = szVendor_Name%>" maxlength=64 size=35 >
						<% end if %>
					</td>
					<td class="TableCell">
						<% if bolPrint then %>
						<% = szContact_First_Name %>&nbsp;
						<% else %>
						<input type=text name="szContact_First_Name" value="<% = szContact_First_Name%>" maxlength=20 size=15 >
						<% end if %>
					</td>	
					<td class="TableCell">
						<% if bolPrint then %>
						<% = szContact_Last_Name %>&nbsp;
						<% else %>
						<input type=text name="szContact_Last_Name" value="<% = szContact_Last_Name%>" maxlength=20 size=15 >
						<% end if %>
					</td>	
					<% if oFunc.IsAdmin  then%>		
					<td class="TableCell">
						<select name=szVendor_Status_CD ID="Select1">
							<option value=""></option>
							<%
								sql = "SELECT szVendor_Status_CD, szVendor_Status_Name " & _ 
									"FROM tblVendor_Status_Codes " & _ 
									"ORDER BY szVendor_Status_Name "
								Response.Write oFunc.MakeListSQL(sql,"szVendor_Status_CD","szVendor_Status_Name",szVendor_Status_CD)								
							%>
						</select>
					</td>
					<% elseif IsVendor then %>
					<td class="TableCell" align="center">&nbsp;
					<%
						select case oFunc.TrueFalse(bolApproved)
							case "1"
								response.Write "Approved"
							case "0"
								response.Write "Rejected"
							case else
								response.Write "Pending"
						end select
					%>
					</td>
					<% end if %>
					<% if oFunc.IsAdmin then%>		
					<td>
						<textarea  style="width:100%;" rows=1 wrap=virtual name="szDeny_Reason" onfocus="this.rows=4;" onblur="this.rows=1;" onKeyDown="jfMaxSize(511,this);" ID="Textarea1"><% = szDeny_Reason %></textarea>
					</td>
					<% elseif IsVendor then %>
					<td class="TableCell">
						<% = szDeny_Reason %>&nbsp;
					</td>
					<% end if %>											
				</tr>
			</table>
		</td>
	</tr>
	<% 
	if FullRights then %>
	<tr>
		<td>
			<table style='width:800px;'>
				<tr>
					<td class="TableHeader"  style='width:50%;'>
							&nbsp;Street Address <font color="orange" size="3">*</font> 
							(or 'Business Website' required)
					</td>
					<td class="TableHeader">
							&nbsp;City <font color="orange" size="3">*</font>
					</td>		
					<td class="TableHeader" align="center">
							&nbsp;State <font color="orange" size="3">*</font>
					</td>	
					<td class="TableHeader">
							&nbsp;Zip Code <font color="orange" size="3">*</font>&nbsp;
					</td>											
				</tr>		
				<tr>
					<td class="TableCell">
						<% if bolPrint then %>
						<% = szVendor_Addr %>&nbsp;
						<% else %>
						<input type=text style='width:100%;' name="szVendor_Addr" value="<% = szVendor_Addr%>" maxlength=128 size=30 >
						<% end if %>
					</td>
					<td class="TableCell">
						<% if bolPrint then %>
						<% = szVendor_City %>&nbsp;
						<% else %>
						<input type=text name="szVendor_City" value="<% = szVendor_City%>" maxlength=64 size=20 >
						<% end if %>
					</td>	
					<td align="center" class="TableCell">
						<% if bolPrint then %>
						<% = sVendor_State %>&nbsp;
						<% else %>
						<select name="sVendor_State" >
						<%
							dim sqlState
							sqlState = "select sState_CD from trefState order by sState_CD"
							Response.Write oFunc.MakeListSQL(sqlState,"sState_CD","sState_CD",sVendor_State)
						%>
						</select>
						<% end if %>
					</td>	
					<td class="TableCell">
						<% if bolPrint then %>
						<% = szVendor_Zip_Code %>&nbsp;
						<% else %>
						<input type=text name="szVendor_Zip_Code" value="<% = szVendor_Zip_Code%>" maxlength=15 size=15 >
						<% end if %>
					</td>	
				</tr>
			</table>
		</td>
	</tr>	
	<tr>
		<td>			
			<table style='width:800px;' ID="Table1">
				<tr>
					<td class="TableHeader"  style='width:50%;'>
							&nbsp;Mailing Address (if different)
					</td>
					<td class="TableHeader">
							&nbsp;City 
					</td>		
					<td class="TableHeader" align="center">
							&nbsp;State
					</td>	
					<td class="TableHeader">
							&nbsp;Zip Code&nbsp;
					</td>											
				</tr>		
				<tr>
					<td class="TableCell">
						<% if bolPrint then %>
						<% = szMail_Addr %>&nbsp;
						<% else %>
						<input type=text style='width:100%;'  name="szMail_Addr" value="<% = szMail_Addr%>" maxlength=128 size=30 >
						<% end if %>
					</td>
					<td class="TableCell">
						<% if bolPrint then %>
						<% = szMail_City %>&nbsp;
						<% else %>
						<input type=text name="szMail_City" value="<% = szMail_City%>" maxlength=64 size=20 >
						<% end if %>
					</td>	
					<td align="center" class="TableCell">
						<% if bolPrint then %>
						<% = sMail_State %>&nbsp;
						<% else %>
						<select name="sMail_State" >
						<%
							sqlState = "select sState_CD from trefState order by sState_CD"
							Response.Write oFunc.MakeListSQL(sqlState,"sState_CD","sState_CD",sMail_State)
						%>
						</select>
						<% end if %>
					</td>	
					<td class="TableCell">
						<% if bolPrint then %>
						<% = szMail_Zip_Code %>&nbsp;
						<% else %>
						<input type=text name="szMail_Zip_Code" value="<% = szMail_Zip_Code%>" maxlength=15 size=15 >
						<% end if %>
					</td>	
				</tr>
			</table>	
		</td>		
	</tr>
	<% end if %>
	<tr>
		<td>
			<table style='width:800px;' >
				<tr>
					<td class="<% = FRCss %>" nowrap>
							&nbsp;Phone Number <font color="orange" size="3">*</font>&nbsp;
					</td>		
					<td class="<% = FRCss %>" nowrap>
							&nbsp;Fax Number&nbsp;
					</td>	
					<td class="<% = FRCss %>" nowrap>
							&nbsp;Email Address <font color="orange" size="3">*</font>&nbsp;
					</td>
					<td class="<% = FRCss %>" nowrap>
						&nbsp;Location
					</td>
					<td class="<% = FRCss %>"  style='width:100%;' nowrap>
							&nbsp;Business Website&nbsp;
					</td>									
				</tr>		
				<tr>
					<td class="TableCell" nowrap>
						<% if bolPrint then %>
						<% = szVendor_Phone %>&nbsp;
						<% else %>
						<input type=text name="szVendor_Phone" value="<% = szVendor_Phone%>" maxlength=20 size=15 ID="Text5">
						<% end if %>
					</td>	
					<td class="TableCell" nowrap>
						<% if bolPrint then %>
						<% = szVendor_Fax %>&nbsp;
						<% else %>
						<input type=text name="szVendor_Fax" value="<% = szVendor_Fax%>" maxlength=20 size=15 ID="Text6">
						<% end if %>
					</td>
					<td class="TableCell" valign="middle" nowrap>
						<% if bolPrint then %>
						<% = szVendor_Email %>&nbsp;
						<% else %>
						<input type=text name="szVendor_Email" value="<% = szVendor_Email%>" maxlength=64 size=35 ID="Text9">
						<% end if %>
					</td>	
					<td class="TableCell" valign="middle" nowrap>
						<% if bolPrint then %>
						<% = szLocation_Name %>&nbsp;
						<% else %>
						<select name="intLocation_Id">
							<option value=""></option>
							<%
								sql = "select intLocation_Id, szLocation_Name from trefLocations order by szLocation_Name"
								Response.Write oFunc.MakeListSQL(sql,"intLocation_Id","szLocation_Name",intLocation_Id)							
							%>
						</select>
						<% end if %>
					</td>	
					<td class="TableCell">
						<% if bolPrint then %>
						<% = szVendor_Website %>&nbsp;
						<% else %>
						<input type=text name="szVendor_Website" value="<% = szVendor_Website%>"  style='width:100%;'  size=30 ID="Text10">
						<% end if %>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
	</tr>
			<% if xGoodService & "" = "1" or xGoodService & "" = "3" then %>
			<% if FullRights and (xGoodService & "" <> "2" or request("intItem_Group_ID") <> "2") then %>
	<tr>	
		<td>
			<table style='width:800px;'>
				<tr>
					<td class="TableHeader">
							&nbsp;Employer Identification or SSN # <font color="orange" size="3">*</font>
					</td>
					<td class="TableHeader">
							&nbsp;AK Business License #<br>&nbsp;(if contract total exceeds $2,500)
					</td>	
					<td class="TableHeader" align="center">
							&nbsp;License Expiration<br>
							mm/dd/yyyy
					</td>
					<td class="TableHeader" align="center">
							&nbsp;Insurance Expiration mm/dd/yyyy<br>
							&nbsp;(if considered a “high-risk” vendor)
					</td>					
				<tr>
					<td class="TableCell">
						<% if bolPrint then %>
						<% = szVendor_Tax_ID %>&nbsp;
						<% else %>
						<input type=text  style='width:100%;' name="szVendor_Tax_ID" value="<% = szVendor_Tax_ID%>" maxlength=128 size=25 ID="Text1">
						<% end if %>
					</td>
					<td class="TableCell">
						<% if bolPrint then %>
						<% = szBusiness_License %>&nbsp;
						<% else %>
						<input type=text name="szBusiness_License"  style='width:100%;' value="<% = szBusiness_License%>" maxlength=64 size=25  ID="Text2">
						<% end if %>
					</td>
					<td class="TableCell">
						<% if bolPrint then %>
						<% = dtLicense_Expires %>&nbsp;
						<% else %>
						<input type=text name="dtLicense_Expires"  style='width:100%;' value="<% = dtLicense_Expires%>" maxlength=10 size=15 ID="Text7">
						<% end if %>
					</td>
					<td class="TableCell">
						<% if bolPrint then %>
						<% = dtInsurance_Expires %>&nbsp;
						<% else %>
						<input type=text name="dtInsurance_Expires"  style='width:100%;' value="<% = dtInsurance_Expires%>" maxlength=10 size=15 ID="Text4">
						<% end if %>
					</td>					
				</tr>	
			</table>	
		</td>
	</tr>
		<% end if %>
	<tr>
		<td>
			<table ID="Table2" style="width:800px;">
				<tr>
					<td colspan="3" class="<% = FRCss %>">
						&nbsp;Services Provied <font color="orange" size="3">*</font>
					</td>
				</tr>
				<% 
					sql2 = "SELECT     vs.intVend_Service_ID, UPPER(ps.szSubject_Name + ': ' + vs.szVend_Service_Name) AS ServiceName " & _ 
									"FROM         trefVendor_Services vs INNER JOIN " & _ 
									"                      trefPOS_Subjects ps ON vs.intPOS_Subject_ID = ps.intPOS_Subject_ID INNER JOIN " & _ 
									"                      tascVendor_Service_Types ON vs.intVend_Service_ID = tascVendor_Service_Types.intVend_Service_ID " & _ 
									"WHERE     (tascVendor_Service_Types.intVendor_Id = " & Request("intVendor_ID") & ") " & _ 
									"ORDER BY ps.szSubject_Name + ': ' + vs.szVend_Service_Name "
														
					if bolPrint then %>
				<tr>
					<td class="TableCell">&nbsp;
						<%
							set rsProvide = server.CreateObject("ADODB.RECORDSET")
							rsProvide.CursorLocation = 3
							rsProvide.Open sql2, oFunc.FpcsCnn
							if rsProvide.RecordCount > 0 then
								do while not rsProvide.EOF
									strProvide =  strProvide & rsProvide("ServiceName") & ", "
									rsProvide.MoveNext
								loop
								strProvide = left(TRIM(strProvide),len(trim(strProvide))-1)
								response.Write strProvide
							end if
							rsProvide.Close
							set rsProvide = nothing
						%>
					</td>
				</tr>
					<% else %>
				<tr>
					<td class="TableCell" colspan="3">
						To add services select a service from the 'Service List' and then click the '>' button. <BR>
						To delete a service under the vendor provided services list select a service and then click the 'x' button.
					</td>
				</tr>
				<tr>
					<td class="gray">
						&nbsp;Service List
					</td>
					<td>
					
					</td>
					<td class="gray">
						&nbsp;Services Provided by this Vendor
					</td>
				</tr>
				<TR>			
					<TD valign="top"  style="width:50%;">
						<SELECT name="xServiceList"  multiple size="6" style="width:100%;FONT-SIZE:xx-small; ID="Select2">
							<option>----------						
							<%
							sql = "SELECT vs.intVend_Service_ID, UPPER(ps.szSubject_Name + ': ' + vs.szVend_Service_Name) AS ServiceName " & _ 
									"FROM trefVendor_Services vs INNER JOIN " & _ 
									" trefPOS_Subjects ps ON vs.intPOS_Subject_ID = ps.intPOS_Subject_ID " & _ 
									" WHERE vs.Is_Active = 1 " & _
									"ORDER BY ps.szSubject_Name + ': ' + vs.szVend_Service_Name "	
							Response.Write oFunc.MakeListSQL(sql,"intVend_Service_ID","ServiceName","")
							%>
						</SELECT>
					</td>
					<td align=right style="width:0%;">
						<input type=button value=">" title="Add selected Services" class="btSmallGray"
						onclick="jfSelectItemFromTo('xServiceList', 'xServices');" align=right NAME="Button2" ID="Button1"><br>
						<input type=button value="x" title="Remove selected Services" class="btSmallGray"
						onclick="jfRemoveItems('xServices');" align=right NAME="Button2" ID="Button2">
					</td>
					<TD valign="top"  style="width:50%;">
						<SELECT name="xServices"  multiple size="6" style="FONT-SIZE:xx-small;width:100%;" ID="Select3">
							<%
							if Request("intVendor_ID") <> "" and request("xServiceHash") = "" then								
									Response.Write oFunc.MakeListSQL(sql2,"intVend_Service_ID","ServiceName","")								
							elseif request("xServiceHash") <> "" then
								dim hash 
								hash = split(request("xServiceHash"),"|")
								for k = 0 to ubound(hash)
									if hash(k) <> "" then
										myHash = split(hash(k),"~")
										if myHash(0) <> "" then
											response.Write "<option value=""" & myHash(0) & """>" & myHash(1) & "</option>" & chr(10) & chr(13)
										end if
									end if
								next
							end if		 
							%>
						</SELECT>
					</TD>
				</tr>			
				<% end if %>
			</table>		
		</td>
	</tr>
	<tr>
		<td>
			<table style="width:800px;">
				<tr>
					<td class="<% = FRCss %>">
							&nbsp;Please list any training/education/experience in this or a related field.
					</td>										
				</tr>		
				<tr>
					<td align=center class="TableCell">
						<% if bolPrint then %>
						<% = szPrev_Experience %>&nbsp;
						<% else %>
						<textarea style="width:100%;" rows=2 name="szPrev_Experience"  onKeyDown="jfMaxSize(511,this);"><% = szPrev_Experience %></textarea>
						<% end if %>
					</td>
				</tr>
			</table>	
		</td>
	</tr>	
	<tr>
		<td>
			<table style="width:800px;" ID="Table5">
				<tr>
					<td class="<% = FRCss %>">
							&nbsp;Other comments about services provided.
					</td>										
				</tr>		
				<tr>
					<td align=center class="TableCell">
						<% if bolPrint then %>
						<% = szVendor_Comments %>&nbsp;
						<% else %>
						<textarea style="width:100%;" rows=2 name="szVendor_Comments" ID="Textarea2"  onKeyDown="jfMaxSize(2000,this);"><% = szVendor_Comments %></textarea>
						<% end if %>
					</td>
				</tr>
			</table>	
		</td>
	</tr>
	<tr>
		<td>		
			<table style="width:800px;">
				<tr>
					<td class="<% = FRCss %>" nowrap>
						&nbsp;How do you charge? <font color="orange" size="3">*</font>&nbsp;
					</td>
					<td class="<% = FRCss %>" style="width:100%;">
						&nbsp;Other Charge Method&nbsp;
					</td>	
					<td class="<% = FRCss %>" nowrap>
						&nbsp;How much do you charge?&nbsp;				
					</td>	
					<td class="<% = FRCss %>" align="center" nowrap>
							&nbsp;Contract Starting Date
					</td>												
				</tr>	
				<tr>
					<td class="TableCell">
						<% if bolPrint then %>
						<% = ChargeLabel %>&nbsp;
						<% else %>
						<select name="intCharge_Type_ID"  >
							<option value="">
							<%
								dim sqlChargeType
								sqlChargeType = "SELECT intCharge_Type_ID, szDesc FROM trefCharge_Type"
								Response.Write oFunc.MakeListSQL(sqlChargeType,"","",intCharge_Type_ID)
															
							%>			
						</select>
						<% end if %>
					</td>
					<td align=center rowspan=2 class="TableCell">
						<% if bolPrint then %>
						<% = szOther_Charge_Method %>&nbsp; (to change call FPCS)
						<% else %>
						<input type=text  style="width:100%;" name=szOther_Charge_Method  maxlength="1500" value="<% = szOther_Charge_Method %>" >
						<% end if %>
					</td>
					<td align=center rowspan=2 valign="middle" class="TableCell">
						<% if bolPrint or (oFunc.IsVendor and intVendor_ID & "" <> "") then %>
						$<% if  isNumeric(curCharge_Amount) then response.Write formatnumber(curCharge_Amount,2)  %>&nbsp;
						<% else %>
						$<input type=text name=curCharge_Amount value="<% = curCharge_Amount %>" size=10 >
						<% end if %>
					</td>
					<td class="TableCell" nowrap>
						<% if bolPrint or not oFunc.IsAdmin then %>
							<% if dtContract_Start & ""  = ""  and (not bolNonProfit or bolNonProfit & "" = "") then %>
							<span class="sverror"> You can not start services for <BR>payment until a date has been<br>
							entered into this space by FPCS.</span>
							<%else %>
							<% = dtContract_Start %>&nbsp;
							<% end if %> 
						<% else %>
						<input type=text name="dtContract_Start"  value="<% = dtContract_Start%>" maxlength=10 size=20 ID="Text3">
						<% end if %>
					</td>	
				</tr>
			</table>	
		</td>
	</tr>
				<%	end if %>
	<tr>
		<td>					
			<table ID="Table4">
				<% if xGoodService & "" = "1" or xGoodService & "" = "3" then %>

				<tr>
					<td class="<% = FRCss%>">
						&nbsp;Are you Non-Profit? <font color="orange" size="3">*</font>&nbsp;
					</td>
					<td class="TableCell">
						<% if bolPrint then %>
						<% = oFunc.YNText(bolNonProfit) %>&nbsp;
						<% else %>
						<select name="bolNonProfit">
							<option value="">
							<%
								strValues = "TRUE,FALSE"
								strText = "Yes,No"								
								Response.Write oFunc.MakeList(strValues,strText,oFunc.TFText(bolNonProfit))
							%>			
						</select>		
						<% end if %>		
					</td>									
				</tr>	
				<% if FullRights and xGoodService & "" <> "2" then %>
				<tr>
					<td class="TableHeader">
						&nbsp;Have you ever been convicted of a misdemeanor or felony? <font color="orange" size="3">*</font>&nbsp;
					</td>
					<td class="TableCell">
						<% if bolPrint then %>
						<% = oFunc.YNText(bolCrime) %>&nbsp;
						<% else %>	
						<select name="bolCrime">
							<option value="">
							<%
								strValues = "TRUE,FALSE"
								strText = "Yes,No"								
								Response.Write oFunc.MakeList(strValues,strText,oFunc.TFText(bolCrime))
							%>			
						</select>	
						<% end if %>			
					</td>									
				</tr>
				<tr>
					<td class="TableHeader">
						&nbsp;Would you consent to a background check and submit to fingerprinting? <font color="orange" size="3">*</font>&nbsp;
					</td>
					<td class="TableCell">
						<% if bolPrint then %>
						<% = oFunc.YNText(bolConsent) %>&nbsp;
						<% else %>		
						<select name="bolConsent">
							<option value="">
							<%	
								strValues = "TRUE,FALSE"
								strText = "Yes,No"								
								Response.Write oFunc.MakeList(strValues,strText,oFunc.TFText(bolConsent))
							%>			
						</select>	
						<% end if %>			
					</td>									
				</tr>	
				<tr>
					<td class="TableHeader">
						&nbsp;Do you have children who are currently enrolled in FPCS? <font color="orange" size="3">*</font>&nbsp;
					</td>
					<td class="TableCell">
						<% 
						dim trChildEnrolledDisplay
						if IsNull(bolKids_FPCS) or bolKids_FPCS = "" then
							trChildEnrolledDisplay = "none"
						elseif CBool(bolKids_FPCS) then 
							trChildEnrolledDisplay = "block"
						else
							trChildEnrolledDisplay = "none"
						end if
						if bolPrint then %>
						<% = oFunc.YNText(bolKids_FPCS) %>&nbsp;
						<% else %>		
						<select onChange="if (this.value == 'TRUE'){trChildEnrolled.style.display='block'}else{trChildEnrolled.style.display='none'};" name="bolKids_FPCS">
							<option value="">
							<%
								strValues = "TRUE,FALSE"
								strText = "Yes,No"								
								Response.Write oFunc.MakeList(strValues,strText,oFunc.TFText(bolKids_FPCS))								
							%>			
						</select>
						<% end if %>
					</td>									
				</tr>	
				<tr id="trChildEnrolled" style="display:<% = trChildEnrolledDisplay%>;">
					<td class=svplain10 style="padding-left:0.5cm;" colspan="2">
						You have disclosed a potential conflict of interest for which a waiver<br>						
						from the FPCS administration is required before you	can be added to the vendor list. <br>
						You will be contacted within two weeks regarding your waiver status.
					</td>
				</tr>
				<tr>
					<td class="TableHeader">
						&nbsp;Are you a Retired Certificated ASD Teacher? <font color="orange" size="3">*</font>&nbsp;
					</td>
					<td class="TableCell">
						<% if bolPrint then %>
						<% = oFunc.YNText(bolCertASDTeacher) %>&nbsp;
						<% else %>
						<select name="bolCertASDTeacher">
							<option value="">
							<%	
								strValues = "TRUE,FALSE"
								strText = "Yes,No"								
								Response.Write oFunc.MakeList(strValues,strText,oFunc.TFText(bolCertASDTeacher))
							%>				
						</select>
						<% end if %>				
					</td>									
				</tr>	
				<tr>
					<td class="TableHeader">
						&nbsp;Are you, or a member of your immediate family, an ASD employee or on the ASD eligible for hire list? <font color="orange" size="3">*</font>&nbsp;
					</td>
					<td class="TableCell">
						<% 
						  dim trCurrentASDDisplay
							if IsNull(bolCurrentASDEmp) or bolCurrentASDEmp = "" then
								trCurrentASDDisplay = "none"
							elseif CBool(bolCurrentASDEmp) then 
								trCurrentASDDisplay = "block"
							else
								trCurrentASDDisplay = "none"
							end if
							
						  if bolPrint then %>
						<% = oFunc.YNText(bolCurrentASDEmp) %>&nbsp;
						<% else %>
						<select onChange="if (this.value == 'TRUE'){trCurrentASD.style.display='block'}else{trCurrentASD.style.display='none'};" name="bolCurrentASDEmp" ID="Select6">
							<option value="">
							<%	
								strValues = "TRUE,FALSE"
								strText = "Yes,No"								
								Response.Write oFunc.MakeList(strValues,strText,oFunc.TFText(bolCurrentASDEmp))								
							%>			
						</select>	
						<% end if %>			
					</td>									
				</tr>
				<tr id="trCurrentASD" style="display:<% = trCurrentASDDisplay%>">
					<td class="TableHeader" style="padding-left:0.5cm" colspan=2>
						<% if bolPrint then %>
						<% = szASDEmpType & "<BR>Position: " & szPosition & "<BR>Work Location: " & szWorkLocation%>&nbsp;
						<% else %>
						What type of ASD Employee are you?&nbsp;<select name="szASDEmpType" ID="Select7">
							<option value="">
							<%	
								strValues = "Classified,Exempt,Certificated"
								strText = "Classified,Exempt,Certificated"								
								Response.Write oFunc.MakeList(strValues,strText,szASDEmpType)
							%>			
						</select><br>
						Position:&nbsp;<input type=text name=szPosition value="<% = szPosition %>" size=40 ID="Text8"><br>
						Work Location:&nbsp;<input type=text name=szWorkLocation value="<% = szWorkLocation %>" size=40 ID="Text11">
						Employee Name:&nbsp;<input type=text name=szWorkLocation value="<% = szConfilt_Name %>" size=40 ID="Text12">
					 <% end if %>
					</td>									
				</tr>
				<% end if %>
				<tr>
					<td class="<% = FRCss %>">
						&nbsp;Are you currently available to provide services? &nbsp;
					</td>
					<td class="TableCell">
						<% if bolPrint then %>
						<% = oFunc.YNText(bolIsActive) %>&nbsp;
						<% else %>
						<select name="bolIsActive" ID="Select5">
							<%
								strValues = "TRUE,FALSE"
								strText = "Yes,No"								
								Response.Write oFunc.MakeList(strValues,strText,oFunc.TFText(bolIsActive))
							%>			
						</select>		
						<% end if %>		
					</td>									
				</tr>	
<%
			end if 'ends ==> if 
%>				
			<%
				if (xGoodService = 2 or xGoodService = 3) and request("intItem_Group_ID") = "" then
			%>
				<tr>
					<td class="<% = FRCss %>">
						&nbsp;Do you sell non-sectarian materials?&nbsp;
					</td>
					<td class="TableCell">
						<% if bolPrint then %>
						<% = oFunc.YNText(bolNonSectarian) %>&nbsp;
						<% else %>
						<select  name="bolNonSectarian" ID="Select2">
							<option value="">
							<%	
								strValues = "TRUE,FALSE"
								strText = "Yes,No"								
								Response.Write oFunc.MakeList(strValues,strText,oFunc.TFText(bolNonSectarian))
							%>			
						</select>	
						<% end if %>	
					</td>									
				</tr>	
			<% end if %>
				<tr>
					<td class="<% = FRCss %>">
						&nbsp;Do you provide online Curriculum services?&nbsp;
					</td>
					<td class="TableCell">
						<% if bolPrint then %>
						<% = oFunc.YNText(bolOnline_Services) %> &nbsp;
						<% else %>
						<select  name="bolOnline_Services" ID="Select4">
							<option value="">
							<%	
								strValues = "TRUE,FALSE"
								strText = "Yes,No"								
								Response.Write oFunc.MakeList(strValues,strText,oFunc.TFText(bolOnline_Services))
							%>			
						</select>	
						<% end if %>			
					</td>									
				</tr>
			</table>						
		</td>	
	</tr>	
</table>
<% if Request("bolWin") <> "" then %>
<input type=button value="Close Window" class="btSmallGray" onclick="window.opener.focus();window.close();">
<% end if 

  if not bolPrint or ((oFunc.IsGuardian or oFunc.IsTeacher) and intVendor_ID = "") or (oFunc.IsAdmin and not bolPrint) then %>
<input type=button value="Save" class="NavSave" onClick="GetServices(this.form);">
<input type="hidden" name="xServiceIds" value="" ID="Hidden1">
<input type="hidden" name="xServiceHash" value="">
<input type="hidden" name="cmdSubmit" value="" ID="Hidden2">
<script language="javascript">
	function GetServices(pForm){
		<% if xGoodService = 1 or xGoodService = 3 then %>
		var strItems = "";
		var strHash = "";
		for (i=0; i< pForm.xServices.length; i++) {
			strItems = strItems + pForm.xServices.options[i].value + ",";
			strHash += pForm.xServices.options[i].value + "~" + pForm.xServices.options[i].text + "|";
		}
		pForm.xServiceIds.value = strItems.substr(0, strItems.length - 1); 
		pForm.xServiceHash.value = strHash.substr(0, strHash.length - 1);
		<% end if %>
		pForm.cmdSubmit.value = "true";
		pForm.submit();
	}
</script>
<% end if %>
</form>		
<%
Server.Execute(Application.Value("strWebRoot") & "Includes/footer.asp")
set oFunc = nothing

'*************************************************************************
'functions/procedures below this line
'*************************************************************************
	
	function vbfValidate(ByRef pobjRequest)
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'Name:		vbfValidate 
	'Purpose:	Server side validation of the form prior to allowing inserts
	'
	'Inputs:	pobjRequest - Dictionary Object - By Reference to allow changes
	'Date:		14 May 2003
	'Author:	Bryan K Mofley (ThreeShapes.com LLC)
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''		
	dim strError		'Store any returned errors
	dim oVal			'validation wsc object
	dim dtExpiration	'fully qualified date
	dim dtNext_Eval		'fully qualified date
	dim strSQL			'SQL for validating - ensure Vendor name/phone number are unique
	dim rsDuplicate		'Recordset for validating - ensure Vendor name/phone number are unique

	
		'dimention all of the form/querystring objects
		for each item in pobjRequest
			execute("dim " & item)
			execute(item & " = """ & replace(replace(replace(pobjRequest(item),chr(13),""),chr(10),""),"""","'") & """")
		next 
		
		set rsDuplicate = Server.CreateObject("ADODB.RECORDSET")
		set oVal = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/formValidation.wsc"))
		
		'*************************************
		'check for duplicate vendor based on name and phone number
		if pobjRequest("intVendor_ID") = "" then
			strSQL = "SELECT intVendor_ID, szVendor_Name, szVendor_Phone " & _
					"FROM   tblVendors " & _
					"WHERE (UPPER(REPLACE(szVendor_Name, ' ', '')) = UPPER(REPLACE('" & oFunc.EscapeTick(szVendor_Name) & "', ' ', ''))) " & _
					"OR    (REPLACE(szVendor_Phone, '-', '') = REPLACE('" & szVendor_Phone & "', '-', ''))"
			rsDuplicate.Open strSQL, oFunc.FPCScnn
			with rsDuplicate
				if not .BOF and not .EOF then
					strError = "A Vendor with that Name or Phone Number already exists.<BR>"
				end if
			end with
		end if		
		'*************************************
		
		'*************************************
		'Required for all Users entering Vendor Info
		oVal.validateField szVendor_Name,"blank","","Business Name" 
			pobjRequest("szVendor_Name") = UCase(szVendor_Name)
		oVal.validateField szContact_First_Name,"blank","","Vendor First Name" 
			pobjRequest("szContact_First_Name") = UCase(szContact_First_Name)
		oVal.validateField szContact_Last_Name,"blank","","Vendor Last Name" 
			pobjRequest("szContact_Last_Name") = UCase(szContact_Last_Name)
		
		oVal.validateField szContact_Last_Name,"blank","","Phone" 
			pobjRequest("szVendor_Phone") = UCase(szVendor_Phone)
		'oVal.validateField szVendor_Phone,"regexp","phone","Phone" 
		'	pobjRequest("szVendor_Phone") = oFunc.Unformat(szVendor_Phone, Array("(", ")", "-", " "))		
		
		if szVendor_Website = "" then 
			oVal.validateField szVendor_Addr,"blank","","Street Address or Business Website" 
				pobjRequest("szVendor_Addr") = UCase(szVendor_Addr)		
			oVal.validateField szVendor_City,"blank","","City" 
				pobjRequest("szVendor_City") = UCase(szVendor_City)
			oVal.validateField sVendor_State,"blank","","State" 
			
			'oVal.validateField szVendor_Zip_Code,"regexp","zip","Zip Code"
			'pobjRequest("szVendor_Zip_Code") = oFunc.Unformat(szVendor_Zip_Code, Array("-", " "))
		end if
		
		'*************************************
		
		'*************************************
		if pobjRequest("xGoodService") = 1 or pobjRequest("xGoodService") = 3 or pobjRequest("xsuggestVendor") <> "" then
			oVal.validateField intCharge_Type_ID,"blank","","How Do You Charge"
			if intCharge_Type_ID <> "" then
				if intCharge_Type_ID = 5 then
					oVal.validateField szOther_Charge_Method,"blank","","Other Charge Type"
				end if
			end if
			'oVal.validateField intWork_Type_ID,"blank","","Independent Contractor?"
			oVal.validateField bolCrime,"blank","","Convicted of a Crime?"
			oVal.validateField bolConsent,"blank","","Consent to Background Check?" 
			oVal.validateField bolKids_FPCS,"blank","","Children Enrolled in FPCS?"
			oVal.validateField szVendor_Tax_ID,"blank","","Employer Identification or SSN #"
			oVal.validateField xServiceIds,"blank","","Services Provided"												
			'oVal.validateField curCharge_Amount,"numeric","","How Much do you Charge?"
			oVal.validateField bolNonProfit,"blank","","Are you Non-Profit?"
			oVal.validateField bolCurrentASDEmp,"blank","","Are you an ASD employee or on the ASD eligible for hire list?"	
			oVal.validateField bolConsent,"blank","","Would you consent to a background check and submit to fingerprinting?"	
			
			if Len(dtLicense_Expires) > 0 then
				oVal.validateField dtLicense_Expires,"date", "", "License Expiration"
			end if
		
			if Len(dtInsurance_Expires) > 0 then
				oVal.validateField dtInsurance_Expires,"date", "", "Insurance Expiration"
			end if
			
			if Len(dtContract_Start) > 0 then
				oVal.validateField dtContract_Start,"date", "", "Contract Starting Date"
			end if
			
		end if				
			
		'*************************************
		
		'*************************************
		'check the following only if user supplied a value
		if Len(szMail_Zip_Code) <> 0 then
			pobjRequest("szMail_Zip_Code") = szMail_Zip_Code
			'oVal.validateField szMail_Zip_Code,"regexp", "zip", "Mail Zip Code"
			'pobjRequest("szMail_Zip_Code") = oFunc.Unformat(szMail_Zip_Code, Array("-", " "))
		end if
		if Len(szVendor_Fax) <> 0 then
			pobjRequest("szVendor_Fax") = szVendor_Fax
			'oVal.validateField szVendor_Fax,"regexp", "phone", "Fax"
			'pobjRequest("szVendor_Fax") = oFunc.Unformat(szVendor_Fax, Array("(", ")", "-", " "))
		end if
		if Len(szVendor_Phone_2) <> 0 then
			pobjRequest("szVendor_Phone_2") = szVendor_Phone_2
			'oVal.validateField szVendor_Phone_2,"regexp","phone", "2nd Phone"
			'pobjRequest("szVendor_Phone_2") = oFunc.Unformat(szVendor_Phone_2, Array("(", ")", "-", " "))
		end if							
		if Len(szVendor_Email) <> 0 then
			oVal.validateField szVendor_Email,"email","","Email"
		else
			oVal.validateField szVendor_Email,"blank","","Email"
		end if						
		
		
		'force to upper case - no need to check contents	
		if Len(szMail_Addr) <> 0 then
			pobjRequest("szMail_Addr") = UCase(szMail_Addr)
		end if			
		if Len(szMail_City) <> 0 then
			pobjRequest("szMail_City") = UCase(szMail_City)
		end if	
		if Len(szVendor_Contact) <> 0 then
			pobjRequest("szVendor_Contact") = UCase(szVendor_Contact)
		end if	
		if Len(szBusiness_License) <> 0 then
			pobjRequest("szBusiness_License") = UCase(szBusiness_License)
		end if					
		if Len(szCert_Insurance) <> 0 then
			pobjRequest("szCert_Insurance") = UCase(szCert_Insurance)
		end if
		if Len(szVendor_Service) <> 0 then
			pobjRequest("szVendor_Service") = UCase(szVendor_Service)
		end if
		'*************************************
		
		if oVal.ValidationError & "" <> "" then
			strError = strError & oVal.ValidationError 
		end if
		
		if left(UCASE(trim(objRequest("szVendor_Name"))),4) = "THE " then
			strError = strError & "Business name can not start with 'The'.<BR>"
		end if
		
		if strError <> "" and not oFunc.IsAdmin then
			strError = "<BR><div class='svplain10'><font color=red><b><font size='4'>DID NOT SAVE!</font><BR> To save your changes you must correct the following items:</B><BR>" & strError & "</font></div>"
		elseif oVal.CriticalError & "" <> "" and oFunc.IsAdmin then
			strError = "<BR><div class='svplain10'><font color=red><b><font size='4'>DID NOT SAVE!</font><BR> To save your changes you must correct the following items:</B><BR>" & oVal.CriticalError & "</font></div>"
		elseif strError <> "" and oFunc.IsAdmin then
			strError = "<BR><div class='svplain10'><font color=red><b><font size='4'>Saved but required fields where not filled in!</font><BR>" & strError & "</font></div>"
		end if	
			
		vbfValidate = strError		
	end function
	
	function vbfInsert(pobjRequest)
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'Name:		vbfInsert 
	'Purpose:	Inserts a new record into tblIEP if necessary
	'Date:		14 May 2003
	'Author:	Bryan K Mofley (ThreeShapes.com LLC)
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	dim strSQL			'SQL for update statement
	dim strSQLfields	'SQL for INSERT field names
	dim strSQLvalues	'SQL for INSERT field values
	dim item			'counter in for next loop
	dim vntValue		'variant value of form field being passed to SQL statement
	
		' Since some of the Form objects will NOT be used in the SQL statement, there is no need
		' to turn the http header variables into vbs variables (this will actually mess up the
		' SQL statement being created below).  Instead, we only use those that begin
		' with our SQL field name standard conventions (int, bol, s, sz or dt)
		
		for each item in pobjRequest
			if Left(item,3) = "cur" or Left(item,3) = "int" or Left(item,3) = "bol" or Left(item,2) = "dt" or Left(item,1) = "s" then
				if item <> "intVendor_ID" and item <> "szVendor_Status_CD" and _
				item <> "intStudent_ID" and item <> "intItem_Group_ID" and _
				item <> "intILP_ID" and item <> "bolReimburse" and _
				item <> "intItem_ID" and item <> "bolWin" and item <> "intOrd_Item_ID" and _
				item <> "intClass_Item_ID" and item <> "viewing" and _
				item <> "intClass_ID" and item <> "strClassName" and _
				item <> "bolComplies" and item <> "intPOS_Subject_ID" _
				and item <> "dtContract_Start" then '<-- don't want to supply a NULL to the primary key!
					strSQLfields = strSQLfields & item & ","
					select case LCase(pobjRequest(item))
						case "yes", "true", "no", "false"
							vntValue = oFunc.ConvertCheckToBit(pobjRequest(item))
						case ""
							vntValue = "NULL"
						case else
							if Left(item,1) = "s" then
								vntValue = "'" & oFunc.EscapeTick(pobjRequest(item)) & "'"
							elseif Left(item,3) = "cur" then
								vntValue = "convert(money,'" & pobjRequest(item) & "') "
							elseif Left(item,2) = "dt" then
								myDate = cdate(pobjRequest(item))
								myDate = Month(pobjRequest(item)) & "/" & Day(pobjRequest(item)) & "/" & Year(pobjRequest(item))
								vntValue = "convert(datetime,'" & oFunc.EscapeTick(myDate) & "') "
							else 
								vntValue = pobjRequest(item)
							end if
					end select
					strSQLvalues = strSQLvalues & vntValue & ","
					strSQLset = strSQLset & item & "=" & vntValue & ","
				end if
			end if
		next
		
		strSQLfields = strSQLfields & "szUser_Create) "
		strSQLvalues = strSQLvalues & "'" & Session.Value("strUserID") & "')"
		strSQLset = strSQLset & "szUser_Modify='" & Session.Value("strUserID") & "'"

		if pobjRequest("intVendor_ID") = "" then
			strSQL = "INSERT INTO tblVendors (" 
			strSQL = strSQL & strSQLfields
			strSQL = strSQL & "VALUES (" & strSQLvalues 
		else
			strSQL = "UPDATE tblVendors SET " 
			strSQL = strSQL & strSQLset
			strSQL = strSQL & " WHERE intVendor_ID = " & pobjRequest("intVendor_ID")
		end if
		
		'Response.Write strSQL
		'Response.End
		
		'on error resume next
		oFunc.BeginTransCN

'response.write strSQL
		oFunc.ExecuteCN(strSQL)
		
		if pobjRequest("intVendor_ID") = "" then
			intVendor_ID = oFunc.GetIdentity
		else
			intVendor_ID = pobjRequest("intVendor_ID")
		end if
		
		' Now update all services if service vendor
		if pobjRequest("intVendor_ID") <> "" and (pobjRequest("xGoodService") = "1" or objRequest("xGoodService") = "3") then
			delete = "delete from tascVendor_Service_Types where intVendor_ID = " & pobjRequest("intVendor_ID")
			oFunc.ExecuteCN(delete)
		end if
		
		if pobjRequest("xServiceIds") <> "" then
			arIds = split(objRequest("xServiceIds"),",")
			for xp = 0 to ubound(arIds)
				if isNumeric(arIds(xp))then
					insert = "insert into tascVendor_Service_Types(intVendor_ID,intVend_Service_Id, dtCreate, szUser_Create) " & _
							 "values (" & intVendor_ID & "," & arIds(xp) & ",CURRENT_TIMESTAMP,'" & Session.Value("strUserID") & "')"
					oFunc.ExecuteCN(insert)
				end if
			next
		end if
		oFunc.CommitTransCN
		
		' now update the Status
		set rsc = server.CreateObject("ADODB.RECORDSET")
		rsc.CursorLocation = 3
		
		if session.Contents("intSchool_Year") = "" then
			SchoolYear = oFunc.SchoolYear
		else
			SchoolYear = session.Contents("intSchool_Year")
		end if
		
		sql = "select top 1 intVendor_Status_ID, szVendor_Status_CD, bolProfile_Verified from tblVendor_Status " & _
			" WHERE intVendor_ID = " & intVendor_ID & _
			" AND intSchool_Year = " & SchoolYear & _
			" ORDER BY intSchool_Year DESC,intVendor_Status_ID DESC "
		rsc.Open sql, oFunc.FpcsCnn
	
		if objRequest("dtContract_Start") & "" = "" then
			dtContract = " NULL "
		else
			dtContract = "'" & objRequest("dtContract_Start") & "' "
		end if 
			
		if rsc.RecordCount > 0 and oFunc.IsAdmin then		
			' update status if changed							
			update = "update tblVendor_Status set szVendor_Status_CD = '" & objRequest("szVendor_Status_CD") & "', " & _
					" dtContract_Start = " & dtContract & ", " & _
					" dtModify = CURRENT_TIMESTAMP, szUser_Modify = '" & Session.Value("strUserID") & "' " & _
					" where intVendor_Status_ID = " & rsc("intVendor_Status_ID")
			oFunc.ExecuteCn(update)
		elseif rsc.RecordCount > 0 and session.Contents("intVendor_ID") <> ""  then
			' Vendor has now verified their Profile
			
			if rsc("szVendor_Status_CD") = "REMV" then
				myStat = ", szVendor_Status_CD = 'PEND' "
			end if 
			if not rsc("bolProfile_Verified") then
				update = "update tblVendor_Status set bolProfile_Verified = 1, " & _
				" dtModify = CURRENT_TIMESTAMP, szUser_Modify = '" & Session.Value("strUserID") & "' " & _
				myStat & " where intVendor_Status_ID = " & rsc("intVendor_Status_ID")
				oFunc.ExecuteCn(update)
				session.Contents("HasUpdated" & session.Contents("intVendor_ID")) = true
			end if
		else
			if objRequest("szVendor_Status_CD") = "" or not oFunc.IsAdmin then
				' default status
				strStateCd = "PEND"
			else
				strStateCd = objRequest("szVendor_Status_CD")
				dtField = ", dtContract_Start "
				dtContract2 = ", " & dtContract
			end if
			
			insert = "insert into tblVendor_Status(intVendor_ID, szVendor_Status_CD, intSchool_Year, szUser_Create, dtCreate " & dtField & ") " & _
					" values (" & intVendor_ID & ",'" & strStateCd & "'," & SchoolYear & ",'" & _
					Session.Value("strUserID") & "', CURRENT_TIMESTAMP" & dtContract2 & ")"
			oFunc.ExecuteCN(insert)
		end if
		
		rsc.Close
		set rsc = nothing

		'detect SQL errors and email developers if necessary
		if Err.number <> 0 then
			Session.Contents("ErrorNum") = Err.number
			Session.Contents("ErrorDesc") = Err.Description
			Server.Execute(Application.Value("strWebRoot") & "admin/debugEmailer.asp")		
			mstrMessage = "<div class='svplain10'><font color=red>An error has occured.<br>A detailed error message has been mailed to the web developer.<br>" & _
				Session.Contents("ErrorNum") & "<br>" & Session.Contents("ErrorDesc") & "<br></font></div>"
			Session.Contents("ErrorNum") = ""
			Session.Contents("ErrorDesc") = ""
		else
			mstrMessage = "<div class='svplain10'><font color=red><b>Vendor Information was Updated and Saved.</b></font></div>"
		end if
		'on error goto 0
		
		if request("xsuggestVendor") <> "" then
			set oCrypto = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/Crypto.wsc"))
			oFunc.BeginTransCN
			strUserName = ucase(replace(pobjRequest("szVendor_Name")," ",""))
			strUserName = replace(left(strUserName,5) & intVendor_ID,"'","")
			oCrypto.Text = intVendor_ID		
			Call oCrypto.Encypttext
			strEncPwd = oCrypto.EncryptedText
			insert = "insert into tblUsers (szUser_Id,szPassword,blnActive,blnForcePWDchange,dtCreate,szUser_Create)" & _
					 " values('" & _
					 oFunc.EscapeTick(strUserName) & "','" & strEncPwd & "',1,0,CURRENT_TIMESTAMP,'NonLogedInUser')"
			oFunc.ExecuteCN(insert)
			
			insert = "insert into tascUserRoles(szUser_ID, szRole_CD,dtCreate,szUser_Create) " & _
					 " values ('" & oFunc.EscapeTick(strUserName) & "', 'VENDOR',CURRENT_TIMESTAMP,'NonLogedInUser')"
			oFunc.ExecuteCN(insert)
			
			insert = "insert into tascVendor_User(intVendor_ID, szUser_Id,dtCreate,szUser_Create)" & _
					 " values (" & intVendor_ID & ",'" & _
					 oFunc.EscapeTick(strUserName) & "',CURRENT_TIMESTAMP,'NonLogedInUser')"
			oFunc.ExecuteCN(insert)
			oFunc.CommitTransCN
			set oCrypto = nothing
						
			Set cdoMessage = Server.CreateObject("CDO.Message")
			set cdoConfig = Server.CreateObject("CDO.Configuration")
			cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 ' cdoSendUsingPort
			cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
			cdoConfig.Fields.Update
			set cdoMessage.Configuration = cdoConfig
	
			cdoMessage.From = "OFFICE@FPCS.NET"
			cdoMessage.Subject = "New Vendor Account"
			cdoMessage.TextBody = "Thank you for filling out our online vendor profile. " & chr(10) & chr(13) & _ 
								"We will review your application and get back with you if there are any problems. For  " & chr(10) & chr(13)  & _ 
								"now you can log into the Vendor Online System where you can review/edit your vendor profile, " & chr(10) & chr(13)  & _ 
								"search our vendor database and review your business activity within our school. At this point teachers and guardians can find and use your service within the FPCS Online System.  " & chr(10) & chr(13)  & _ 
								" " & chr(10) & chr(13)  & _ 
								"Your Log in information is as follows ... " & chr(10) & chr(13)  & _ 
								"User Name: " & strUserName & chr(10) & chr(13)  & _ 
								"Password: "  & intVendor_ID & chr(10) & chr(13) & _ 
								" "  & chr(10) & chr(13) & _ 
								"Please change your password after logging in to the system. " & chr(10) & chr(13)  & _ 
								" " & chr(10) & chr(13)  & _ 
								"Welcome Aboard! " & chr(10) & chr(13)  & _ 
								"The FPCS Staff " & chr(10) & chr(13)  & _ 
								"http://www.fpcs.net"
			cdoMessage.To = objRequest("szVendor_Email")	
			cdoMessage.Send 
			
			'Clean up Objects
			Set cdoMessage = Nothing 
	
			' Some general public person just added a new vendor without logging in so we
			' boot them back to the front page
			session.Contents("strRole") = ""			
		%>
		<table>
		<% if request("bolNonProfit") <> "TRUE" then  %>
		<% = vbfStillToDo %>
		<% end if %>
		</table>		
		<BR><BR>		
		Your Vendor Profile has been received. You will be given a user name and password to log into the system via email.<BR><BR>
		To close this window click <a href="#" onclick="window.opener.focus();this.window.close();">HERE</a>.		
		<%
			set oFunc = nothing
			Server.Execute(Application.Value("strWebRoot") & "Includes/footer.asp")
			response.End
		end if
	end function
	
function vbfStillToDo
%>
	<tr>
		<td  colspan="10">
			<br>
			<b><font face="Arial" color="red" size="4">STILL TO DO ...</font></b>
		</td>		
	</tr>
	<tr>
		<td class="svplain10" colspan="10"> 
			<b>You must complete and submit a <u>Personal Service Contract</u> or <u>Service Agreement</u>
			with FPCS (ASD). If you need a Personal Service Contract or Service Agreement click here >>
<a href="https://www.fpcs.net/FpcsWeb2/BusinessOffice/Vendors/tabid/76/Default.aspx">Vendor Forms Link</a> <<.</b> <br><br>
			After a contract has been completed, submitted and approved by ASD, a date will be entered into the 
			'Contract Start Date' field of your Vendor Profile.  That is the date you can begin providing
			services for payment.  <br><br>
		</td>
	</tr>
<%
end function
%>
