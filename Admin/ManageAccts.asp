<%@ Language=VBScript %>
<%

'*******************************************
'Name:		Admin\ManageAccts.asp
'Purpose:	Allows FPCS staff to manage web user accounts
'
'CalledBy:	
'
'Inputs:	Request.QueryString("szUserID")
'
'Author:	ThreeShapes.com LLC
'Date:   18 April 2002
'*******************************************
'option explicit
dim oFunc
dim strSQL
dim rs
dim intCount
dim strRoles
dim strFamilies
dim strTeachers
dim strGuardians
dim strForcedActions
dim strAction
dim wscCrypto		'wsc object
dim strEncPwd
dim strMsg
dim blnActive
dim blnForcePWDchange
dim szUser_ID
dim item
dim strUIDdisplay		'INPUT Tag for UserID

'prevent page from cacheing
Response.CacheControl = "no-cache"
Response.Expires = -1

Session.Value("strTitle") = "User Account Management"
Server.Execute(Application.Value("strWebRoot") & "Includes/header.asp")

if Session.Value("strRole") = "ADMIN" then 

	strUIDdisplay = "<INPUT maxlength='50' size='25' name='szUser_ID' value=''>"

	'set some default properties
	strAction = "Submit"
	blnActive = 0
	blnForcePWDchange = 0

	set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
	set wscCrypto = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/Crypto.wsc"))
	'wscCrypto.Key = "something"	'actual key is not shown here

	call oFunc.OpenCN()
	if Request("hblnDirty") <> "" then
		if Request("chkActive") = "on" then blnActive = 1
		if Request("chkPWDchange") = "on" then blnForcePWDchange = 1
		
		if Request("myAction") = "Submit" then
			'we are in insert mode
			'encrypt password for database write
			wscCrypto.Text = Request("szPassword")
			Call wscCrypto.Encypttext
			strEncPwd = wscCrypto.EncryptedText
			strSQL = "INSERT INTO tblUsers " & _
						"(szUser_ID, szName_First, szName_Last, szEmail, szPassword, blnActive, blnForcePWDchange, " & _
						"szUSER_CREATE, szUSER_MODIFY) " & _
						"VALUES ('" & Trim(UCase(Request("szUser_ID"))) & "', " & _
						"'" & Replace(UCase(Request("szName_First")), "'", "''") & "', " & _
						"'" & Replace(UCase(Request("szName_Last")), "'", "''") & "', " & _
						"'" & Request("szEmail") & "', " & _
						"'" & strEncPwd & "', " & _
						blnActive & ", " & blnForcePWDchange & ", " & _
						"'" & Session.Value("strUserID") & "', " & _
						"'" & Session.Value("strUserID") & "')"
			on error resume next
			oFunc.ExecuteCN(strSQL)
			if Err.number <> 0 then
				if InStr(Err.description, "duplicate key") > 0 then
					strMsg = "<FONT face='tahoma'>User ID " & Request("szUser_ID") & " already exists.  Please choose another</FONT>"
				else
					Response.Write Err.number & ":" & Err.description
					Response.End
				end if
			end if
			on error goto 0
		else


			'we are in update mode
			strSQL = "UPDATE tblUsers SET " & _
						"szName_First = '" & Replace(UCase(Request("szName_First")), "'", "''") & "', " & _
						"szName_Last = '" & Replace(UCase(Request("szName_Last")), "'", "''") & "', " & _
						"szEmail = '" & replace(Request("szEmail"),"'","") & "', " & _
						"blnActive = " & blnActive & ", " & _
						"blnForcePWDchange = " & blnForcePWDchange & ",  " & _
						"szUSER_MODIFY = '" & Session.Value("strUserID") & "' " 
			if Request("hblnPWDirty") <> "" then 
				wscCrypto.Text = Request("szPassword")
				Call wscCrypto.Encypttext
				strEncPwd = wscCrypto.EncryptedText
				strSQL = strSQL & ", szPassword = '" & strEncPwd & "' " 
			end if
			strSQL = strSQL & "WHERE (szUser_ID = '" & Trim(UCase(replace(Request("szUser_ID"),"'","''"))) & "')"
			on error resume next
			oFunc.ExecuteCN(strSQL)
			if Err.number <> 0 then
				Response.Write Err.number & ":" & Err.description & "<BR>"
				Response.End
			end if
			on error goto 0			
		end if


		'update associative entities
		strSQL = "DELETE FROM tascUserRoles " & _
					"WHERE szUser_ID = '" & Trim(UCase(replace(Request("szUser_ID"),"'","''"))) & "'"
		oFunc.ExecuteCN(strSQL)
		
		strSQL = "DELETE FROM tascGUARD_USERS " & _
					"WHERE szUser_ID = '" & Trim(UCase(replace(Request("szUser_ID"),"'","''"))) & "'"
		oFunc.ExecuteCN(strSQL)	
		strSQL = "DELETE FROM tascINSTR_USER " & _
					"WHERE szUser_ID = '" & Trim(UCase(replace(Request("szUser_ID"),"'","''"))) & "'"
		oFunc.ExecuteCN(strSQL)
		strSQL = "DELETE FROM tascUsers_Action " & _
					"WHERE szUser_ID = '" & Trim(UCase(replace(Request("szUser_ID"),"'","''"))) & "'"
		oFunc.ExecuteCN(strSQL)
		
		strSQL = "DELETE FROM tascVENDOR_USER " & _
					"WHERE szUser_ID = '" & Trim(UCase(replace(Request("szUser_ID"),"'","''"))) & "'"
		oFunc.ExecuteCN(strSQL)


		dim strCodes
		dim arCodes
		dim i
		if Request("selRoles").Count > 0 then
			strSQL = "INSERT INTO tascUserRoles " & _
						"(szUser_ID, szRole_CD, szUSER_CREATE, szUSER_MODIFY) " & _
						"VALUES     ('" & Trim(UCase(replace(Request("szUser_ID"),"'","''"))) & "', '" & Request("selRoles") & "', " & _
						"'" & Session.Value("strUserID") & "', " & _
						"'" & Session.Value("strUserID") & "')"

			oFunc.ExecuteCN(strSQL)
			
			if Request("selRoles") = "GUARD" and request("selID") <> "" then
				strSQL = "INSERT INTO tascGUARD_USERS " & _
							"(szUser_ID, intGuardian_ID, szUSER_CREATE, szUSER_MODIFY) " & _
							"VALUES     ('" & Trim(UCase(Request("szUser_ID"))) & "', '" & request("selID") & "', " & _
							"'" & Session.Value("strUserID") & "', " & _
							"'" & Session.Value("strUserID") & "')"
				oFunc.ExecuteCN(strSQL)	
			elseif Request("selRoles") = "TEACHER" and request("selID") <> "" then
				strSQL = "INSERT INTO tascINSTR_USER " & _
							"(szUser_ID, intINSTRUCTOR_ID, szUSER_CREATE, szUSER_MODIFY) " & _
							"VALUES     ('" & Trim(UCase(Request("szUser_ID"))) & "', '" & request("selID") & "', " & _
							"'" & Session.Value("strUserID") & "', " & _
							"'" & Session.Value("strUserID") & "')"
				oFunc.ExecuteCN(strSQL)
			elseif Request("selRoles") = "VENDOR" and request("selID") <> "" then
				strSQL = "INSERT INTO tascVENDOR_USER " & _
							"(szUser_ID, intVENDOR_ID, szUSER_CREATE, szUSER_MODIFY) " & _
							"VALUES     ('" & Trim(UCase(Request("szUser_ID"))) & "', '" & request("selID") & "', " & _
							"'" & Session.Value("strUserID") & "', " & _
							"'" & Session.Value("strUserID") & "')"
				oFunc.ExecuteCN(strSQL)
			end if
		end if	
					
		if Request.Form("intAction_ID") <> "" then
			arActions = split(Request.Form("intAction_ID"),",")
			if isArray(arActions) then
				for i = 0 to ubound(arActions)
					strSQL = "INSERT INTO tascUsers_Action " & _
								"(szUser_ID, intAction_ID, intOrder_ID, szUSER_CREATE, szUSER_MODIFY) " & _
								"VALUES     ('" & Trim(UCase(Request("szUser_ID"))) & "'," & arActions(i) & ", " & _
								i + 1 & ", '" & Session.Value("strUserID") & "', " & _
								"'" & Session.Value("strUserID") & "')"
					oFunc.ExecuteCN(strSQL)
				next
			end if 
		end if
	end if
		
	if Request("selUserID") <> "" then
		szUser_ID = replace(Trim(Request("selUserID")),"'","''")
	elseif Request("szUser_ID") <> "" then
		szUser_ID = replace(Trim(Request("szUser_ID")),"'","''")
	end if

	if szUser_ID <> "" then
		strUIDdisplay = "<FONT face='Tahoma'>" & szUser_ID & "</FONT>" & _
							 "<INPUT type='hidden' name='szUser_ID' value='" & szUser_ID & "'>"
		
		if ucase(Request("selRoles")) = "VENDOR" then
			strSQL = "SELECT     u.szPassword, u.blnActive, ur.szRole_CD AS strRoles, v.szVendor_Name as szName_First, " & _ 
					" v.szContact_First_Name + ' ' + v.szContact_Last_Name AS szName_Last,  " & _ 
					"                      v.szVendor_Email as szEmail " & _ 
					"FROM         tascUserRoles ur INNER JOIN " & _ 
					"                      tblUsers u ON ur.szUser_ID = u.szUser_ID INNER JOIN " & _ 
					"                      tascVendor_User vu ON u.szUser_ID = vu.szUser_ID INNER JOIN " & _ 
					"                      tblVendors v ON vu.intVendor_ID = v.intVendor_ID " & _ 
					"WHERE     (u.szUser_ID = '" & szUser_ID & "') "
		else
			strSQL = "SELECT u.szName_First, u.szName_Last, u.szEmail, u.szPassword, u.blnActive, ur.szRole_CD as strRoles " & _ 
						"FROM tascUserRoles ur INNER JOIN " & _ 
						" tblUsers u ON ur.szUser_ID = u.szUser_ID " & _ 
						"WHERE (u.szUser_ID = '" & szUser_ID & "') "
		end if
		
		set rs = Server.CreateObject("ADODB.Recordset")
		with rs
			.CursorLocation = 3

			.Open strSQL, Application("cnnFPCS")'oFunc.FPCScnn
			if not .BOF and not .EOF then
				strAction = "Update"
				'**********************************************************************
				' This for loop will dimension AND assign our User info variables
				' for us. We'll use them later to populate the form.
				'**********************************************************************
				intCount = 0
				for each item in .Fields
					execute("dim " & .Fields(intCount).Name)
					execute(.Fields(intCount).Name & " = item")
					intCount = intCount + 1
				next
				
				if blnActive then 					
					blnActive = " checked "
				else
					blnActive = ""
				end if
				
				if blnForcePWDchange then blnForcePWDchange = " checked "
				'decrypt password for database compare
				'wscCrypto.Text = szPassword
				'Call wscCrypto.Encypttext
				'strEncPwd = wscCrypto.EncryptedText
			end if
			.close		
			
			'***************************
			'BKM 21-Aug-2002 - show/hide Guardian and Teacher
			'dropdowns based on Role(s)
			dim strVisibility
			strVisibility = "hidden"
			
			if strRoles = "GUARD" then
				strVisibility = "visibile"
				'***************GUARDIAN INFO****************
				strSQL = "SELECT intGuardian_ID FROM tascGUARD_USERS WHERE (szUser_ID = '" & szUser_ID & "')"
				.Open strSQL, Application("cnnFPCS")'oFunc.FPCScnn
				if not .BOF and not .EOF then
					strGuardians = rs("intGuardian_ID")					
				end if
				.Close									
			elseif strRoles = "TEACHER" then
				strVisibility = "visibile"
				strSQL = "SELECT intINSTRUCTOR_ID FROM tascINSTR_USER WHERE  (szUser_ID = '" & szUser_ID & "')"
				.Open strSQL, Application("cnnFPCS")'oFunc.FPCScnn
				if not .BOF and not .EOF then
					strTeachers =  rs("intINSTRUCTOR_ID")
				end if
				.Close				
			elseif strRoles = "VENDOR" then
				strVisibility = "visibile"
				strSQL = "SELECT intVendor_ID FROM tascVendor_USER WHERE  (szUser_ID = '" & szUser_ID & "')"
				.Open strSQL, Application("cnnFPCS")'oFunc.FPCScnn
				if not .BOF and not .EOF then
					strVendor = rs("intVendor_ID")
				end if				
				.Close	
			end if
			'end new code
			'***************************

			'***************FORCED ACTIONS****************	
			'if false then
			'	strSQL = "SELECT intAction_ID FROM tascUsers_Action WHERE (szUser_ID = '" & szUser_ID & "')"
			'	.Open strSQL, oFunc.FPCScnn
			'	if not .BOF and not .EOF then
			'		do until .EOF
			'			strForcedActions = strForcedActions & rs("intAction_ID") & ", "
			'			.MoveNext
			'		loop
			'		strForcedActions = Left(strForcedActions, len(strForcedActions) - 2)
			'	end if
			'	.Close	
			'end if
		end with
	else
		' might neeed to clear user name here
		blnActive = " checked "
	end if	
	
	if request("selRoles") <> "" then
		select case ucase(request("selRoles"))
			case "GUARD"
				sql = "Select intGuardian_ID, szLast_Name + ', ' + szFirst_Name as Name from tblGuardian order by 	Name "	
				strIdList = oFunc.MakeListSQL(sql,"intGuardian_ID","Name",strGuardians)
			case "TEACHER"
				sql = "Select intInstructor_ID, szLast_Name + ', ' + szFirst_Name as Name from tblInstructor order by Name "	
				strIdList = oFunc.MakeListSQL(sql,"intInstructor_ID","Name",strTeachers)
			case "VENDOR"
				sql = "Select intVendor_ID, szVendor_Name from tblVendors order by szVendor_Name "	
			strIdList = oFunc.MakeListSQL(sql,"intVendor_ID","szVendor_Name",strVendor)
		end select
	end if
%>
<SCRIPT language="javascript">
	function jfChangeUser(pObj){
	//reloads page with newly selected student
		var mId  = pObj.selUserID.value.replace("'","''");
		//JD:output the account status to the request variable
		var strURL = "<% = Application.Value("strWebRoot")%>Admin/ManageAccts.asp?selUserID=" + pObj.selUserID.value + '&selRoles=' + pObj.selRoles.value+ '&selStatus=' + pObj.selStatus.value;
		window.open(strURL, "_self");
	}

	function jfDirty() {
		frmMain.hblnDirty.value  = true;
	}

	function jfValidate(objForm){
		var strItems = "";
		var strMsg = "";
		//since we aren't using the multi-select in a proper way, we take all of the
		//options in the selChosenActions dropdown and write them to a hidden field
		<% if ucase(strAction) = "SUBMIT" then %>
		if (objForm.szUser_ID.value.length < 1){
			alert("You must enter a User Id.");
			return false;
		}
		<%end if %>
		//for (i=0; i< objForm.selChosenActions.length; i++) {
		//	strItems = strItems + objForm.selChosenActions.options[i].value + ",";
		//}
		//objForm.intAction_ID.value = strItems.substr(0, strItems.length - 1); 
		//if (strItems != ""){jfDirty();}
	
		//bkm - added 21-aug-2002
		//test for a chosen role
		if (objForm.selRoles.value == ""){
			alert("Please select at least one role for this user");
			return false;
		}else{
			<% if ucase(request("selRoles")) <> "ADMIN" then %>
			if (objForm.selID.value == ""){
				alert("You must select a <% = request("selRoles") %> to associate this new user to.");
				return false;
			}
			<% end if %>
		}
		frmMain.submit();
	}
	
	function jfSelectItemFromTo(selectFrom, selectTo) {
		//based on ideas from excite.com's weather selection - heavily modified
		var blnSelected = false;
		var selected = selectFrom.selectedIndex;
		if (selected != -1){
			for (j=0; j<selectFrom.length; j++) {
				if (selectFrom.options[j].selected){
					var selectedText = selectFrom.options[j].text;
					var selectedValue = selectFrom.options[j].value;
					if (selectedValue != "") {
						var toLength = selectTo.length;
						var i;
						// If item is already added, give it focus
						for (i=0; i<toLength; i++) {
							if (selectTo.options[i].value == selectedValue) {
								blnSelected = true;
							}
						}
						if (!blnSelected){
							// Add new option 
							selectTo.options[selectTo.length] = new Option(selectedText, selectedValue);
						}
					}
				}
				blnSelected = false;
			}
		}
	}	
	
	function jfRemoveItems(pobjSelect){
	//remove items from multiple select list
	//Since setting an option to NULL changes the index
	//value of the item beneath it, we have to make a couple
	//of passes at the object.  We first grab the quantity of
	//items selected, then we use that as a counter to remove
	//the selected items
	var iCnt = 0;
		for (i=0; i<pobjSelect.length; i++) {
			if (pobjSelect.options[i].selected){
				iCnt ++;
			}
		}
		for (j=0; j<iCnt; j++){
			for (i=0; i<pobjSelect.length; i++) {
				if (pobjSelect.options[i].selected){
					pobjSelect.options[i] = null;
				}
			}
		}
	}


	function SwapElement(pselObj, nDirection){
	<% 'from http://groups.google.com/groups?q=move+items+option+up+down+javascript&hl=en&lr=&ie=UTF-8&scoring=r&selm=%23GOaRzPnAHA.2164%40tkmsftngp05&rnum=7 %>
	// Take the currently selected element in SelectColumns
	// and swap it with the element that is nDirection from
	// the current element
		with(pselObj)  {
			// If there is more than one item selected, alert the user
			var nCount = 0;
			for(var x = 0; x < length; x++) {
			    if(options[x].selected) {nCount++;}
			}

			if(nCount > 1) {
			    alert("Please select a single column to move up or down");
			    return;
			}

			var nIndex = selectedIndex;
			if(nIndex == -1){
			    alert("Please select a column to move up or down");
			    return;
			}

			// Make sure we are not the top element
			// or bottom element trying to move too far
			var nSwapIndex = nIndex + nDirection;
			if(nSwapIndex < 0 || nSwapIndex >= length){return;}

			var nValue = options[nIndex].value;
			var strText = options[nIndex].text;

			var nSwapValue = options[nSwapIndex].value;
			var strSwapText = options[nSwapIndex].text;

			options[nIndex] = new Option(strSwapText, nSwapValue);
			options[nSwapIndex] = new Option(strText, nValue);

			selectedIndex = nSwapIndex;
		}
	}

</SCRIPT>

<FORM name="frmMain" method="post">
<INPUT type="hidden" name="hblnDirty" value="" >
<INPUT type="hidden" name="hblnPWDirty" value="">
<% = strMsg %>
<TABLE cellspacing="2" cellpadding="4" align="center" border="0" style="width:800px;">
	<TR>
		<TD valign="center" bgcolor="#666699">
			<TABLE cellspacing="0" cellpadding="1" width="100%" border="0">
				<TR>
					<TD><FONT face="tahoma" color="white"><B>User Manager</B></FONT>
					</TD>
					<TD align="right">
						<SELECT name="selRoles" style="FONT-SIZE:xx-small; WIDTH: 150px;" onchange="window.location.href='<% = Application.Value("strWebRoot") %>Admin/ManageAccts.asp?selRoles=' + this.value;">
							<option value="">Select a Role
							<%
							sql = "SELECT     szRole_CD, szRole_Desc " & _ 
								  "FROM       tblRoles " & _ 
								  "ORDER BY szRole_Desc "
							Response.Write oFunc.MakeListSQL(sql,"szRole_CD","szRole_Desc",request("selRoles"))											 
							%>
						</SELECT>
					</TD>
				</TR>
				<%'JD: select account status %>
				<% if request("selRoles") <> "" then %>
				<tr>
					<TD>
					<FONT face="tahoma" color="white" size="-1">Account Status</FONT>
					</TD>
				    <td align="right">
				        <select name="selStatus" onchange="document.body.style.cursor='wait';window.location.href='<% = Application.Value("strWebRoot") %>Admin/ManageAccts.asp?selRoles=<%=request("selRoles") %>&selStatus=' + this.value;">
				        <option value=''>Select Account Status</option>
				        <% 
				            response.write oFunc.MakeList("1,0", "Active,Inactive", request("selStatus"))
				        %>
				        </select>
				    </td>
				</tr>
                <%if request("selStatus")<>"" then %>
				<TR>
					<TD><FONT face="tahoma" color="white" size="-1"><% if request("selUserID") = "" then response.write "<b>Create New User.</b>&nbsp;" end if %> All fields required.</FONT>
					</TD>									
					<TD align="right">
						<SELECT name="selUserID" onchange="document.body.style.cursor='wait'; jfChangeUser(this.form);">
							<option value="">Select Existing User
							<%
								if ucase(request("selRoles")) = "VENDOR" then
									sql = "SELECT ur.szUser_ID, v.szVendor_Name + CASE WHEN u.blnActive = 1 THEN ' : ACTIVE' ELSE ' : INACTIVE' END AS name " & _ 
											"FROM         tascUserRoles ur INNER JOIN " & _ 
											"                      tblUsers u ON ur.szUser_ID = u.szUser_ID INNER JOIN " & _ 
											"                      tascVendor_User ON u.szUser_ID = tascVendor_User.szUser_ID INNER JOIN " & _ 
											"                      tblVendors v ON tascVendor_User.intVendor_ID = v.intVendor_ID " & _ 
											"WHERE     (ur.szRole_CD = 'VENDOR' " & _
											"           and u.blnActive =" & request("selStatus") & ")" & _ 
											"ORDER BY Name "
								else
									sql = "SELECT ur.szUser_ID, u.szName_Last + ', ' + u.szName_First + CASE WHEN u.blnActive = 1 THEN ' : ACTIVE' ELSE ' : INACTIVE' END AS name " & _ 
											"FROM tascUserRoles ur INNER JOIN " & _ 
											" tblUsers u ON ur.szUser_ID = u.szUser_ID " & _ 
											"WHERE (ur.szRole_CD = '" & request("selRoles") & "' " & _ 
											"           and u.blnActive =" & request("selStatus") & ")" & _ 
											"ORDER BY Name "
								end if
							Response.Write oFunc.MakeListSQL(sql,"szUser_ID","Name",szUser_ID)									 
							%>
						</SELECT>
					</TD>					
				</TR>
				<% end if %>
				<% end if %>
				<%'JD %>

			</TABLE>
		</TD>
	</TR>
	<%'JD Add logic to filter by account status %>
	<% if request("selRoles") <> ""  and request("selStatus")<> "" then %>
	<TR>
		<TD valign="middle" align="center">
			<TABLE cellspacing="1" cellpadding="2" border="0" >
				<TR valign="center" align="left">
					<TD class="gray" align="right" nowrap>
						<B>User ID:</B>
					</TD>
					<TD style="width:250px;">
						<% = strUIDdisplay%>
					</TD>
					<TD rowspan="7">
						&nbsp;
					</td>
					<TD class="gray" align="right">
						<B>Active:</B>
					</TD>
					<TD>
						<INPUT type="checkbox" name="chkActive" <% = blnActive %> onchange="jfDirty();">
					</TD>
				</TR>
				<TR valign="center" align="left">
					<TD class="gray" align="right" nowrap>
						<B>
						<% if ucase(request("selRoles")) = "VENDOR" then %>
						Vendor Name:
						<% else %>
						First Name:
						<% end if %>
						</B>
					</TD>
					<TD class="svplain8">
						<% if ucase(request("selRoles")) = "VENDOR" then %>
						<% = szName_First%>
						<% else %>
						<INPUT maxlength="50" size="25" name="szName_First" value="<% = szName_First%>" onchange="jfDirty();">
						<% end if %>
					</TD>
					<TD class="gray" align="right" nowrap>
						<B>New Password:</B>
					</TD>
					<TD>
						<INPUT type="password" maxlength="50" size="25" value="<% '= strEncPwd%>" name="szPassword" onchange="jfDirty(); this.form.hblnPWDirty.value=true;">
					</TD>
				</TR>
				<TR valign="center" align="left">
					<TD class="gray" align="right" nowrap>
						<B>
						<% if ucase(request("selRoles")) = "VENDOR" then %>
						Contact Name:
						<% else %>
						Last Name:
						<% end if %>
						</B>
					</TD>
					<TD class="svplain8">
						<% if ucase(request("selRoles")) = "VENDOR" then %>
						<% = szName_Last%>
						<% else %>
						<INPUT maxlength="50" size="25" name="szName_Last" value="<% = szName_Last%>" onchange="jfDirty();">
						<% end if %>
					</TD>
					<TD class="gray" align="right" nowrap>
						<B>Confirm Password:</B>
					</TD>
					<TD>
						<INPUT type="password" maxlength="50" size="25" value="<% '= strEncPwd%>" name="szPassword2" onchange="jfDirty(); this.form.hblnPWDirty.value=true;">
					</TD>
				</TR>
				<TR align="left">
					<TD class="gray" align="right" nowrap>
						<B>Email address:</B>
					</TD>
					<TD class="svplain8">
						<% if ucase(request("selRoles")) = "VENDOR" then %>
						<% = szEmail%>
						<% else %>
						<INPUT maxlength="50" size="25" name="szEmail" value="<% = szEmail%>" onchange="jfDirty();">
						<% end if %>
					</TD>
					<TD class="gray" id="tdTitleGuardians" style="visibility:<% = strVisibility%>" nowrap>
						<B>Associated <% = request("selRoles") %>: </B>
					</TD>			
					<TD style="visibility:<% = strVisibility%>">
						<SELECT name="selID" style="FONT-SIZE:xx-small;" onchange="jfDirty();">
							<option value="">
							<%
							Response.Write strIdList 									 
							%>
						</SELECT>
					</TD>
					<!--<TD class="gray" align="right" title="Force the user to supply a new password the next time they log in">
						<B>Force Pwd Change:</B>
					</TD> 
					<TD>
						<INPUT type="checkbox" name="chkPWDchange" <% = blnForcePWDchange %> onchange="jfDirty();">
					</TD>-->
				</tr>
				<TR>
					<TD style="FONT-FAMILY:Verdana; COLOR:Gray; FONT-SIZE: 9px;" colspan="5">
						Record created on <% = dtCREATE %> by <% = szUSER_CREATE %>
					</TD>
				</TR>
				<TR>
					<TD style="FONT-FAMILY:Verdana; COLOR:Gray; FONT-SIZE: 9px;" colspan="5">
						Record last modified on <% = dtMODIFY %> by <% = szUSER_MODIFY %>
					</TD>
				</TR>
			</TABLE>
			<BR>
			<TABLE id="Table5">
				
				<!--
				<TR>
					<TD colspan=4 class="gray">
						<B>Force Action</B>
					</TD>
				</TR>
				<TR>
					<TD valign="top">
						<SELECT name="selEnforcableActions"  multiple size="6" style="FONT-SIZE:xx-small; WIDTH: 150px;" onchange="jfDirty();">
							<option value="">Enforcable Actions
							<option>----------
							
							<%
							'dim sqlForceActions 
							'sqlForceActions = "Select intAction_id,szAction_Name from tblForce_Action order by szAction_name "
							'Response.Write oFunc.MakeListSQL(sqlForceActions,"intAction_id","szAction_Name","")'removed strForcedActions
							%>
						</SELECT>
					</TD>
					<TD valign=middle align=center>
						<IMG SRC="<% = Application.Value("strImageRoot")%>add_arrow.gif" 
							title="Add selected 'Forced Action'"
							onclick="jfSelectItemFromTo(selEnforcableActions, selChosenActions);">
					</TD>
					<TD valign="top">
						<SELECT name="selChosenActions"  multiple size="6" style="FONT-SIZE:xx-small; WIDTH: 150px;" onchange="jfDirty();">
							<%
							'dim strSQLSelectedActions 
							'strSQLSelectedActions = "SELECT tascUsers_Action.intAction_ID, tblForce_Action.szAction_Name " & _
							'						"FROM tascUsers_Action INNER JOIN " & _
							'						"tblForce_Action ON tascUsers_Action.intAction_ID = tblForce_Action.intAction_ID " & _
							'						"WHERE (tascUsers_Action.szUser_ID = '" & szUser_ID & "') " & _
							'						"ORDER BY tascUsers_Action.intOrder_ID"
							'Response.Write oFunc.MakeListSQL(strSQLSelectedActions,"intAction_id","szAction_Name","")						 
							%>
						</SELECT>
						<input type="hidden" name="intAction_ID">
					</TD>
					<TD valign=middle align=center>
						<IMG SRC="<% = Application.Value("strImageRoot")%>up.gif" title="Move selected 'Forced Action' Up" onclick="SwapElement(selChosenActions, -1);"><BR>
						<IMG SRC="<% = Application.Value("strImageRoot")%>remove.gif" title="Remove selected 'Forced Action'" onclick="jfRemoveItems(selChosenActions);"><BR>
						<IMG SRC="<% = Application.Value("strImageRoot")%>down.gif" title="Move selected 'Forced Action' Down" onclick="SwapElement(selChosenActions, 1);">
					</TD>
				</TR>-->
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD valign="bottom" align="right">
			<input type="button" onClick="window.location.href='<% = Application.Value("strWebRoot") %>Admin/ManageAccts.asp';" value="Cancel" class="btSmallGray">
			&nbsp;<BUTTON name="cmdAction" onclick="jfValidate(this.form);" class="NavSave"><% = strAction%></BUTTON>
			<input type="hidden" name="myAction" value="<% = strAction%>">
		</TD>
	</TR>
	<% end if %>
</TABLE>
</FORM>
<%
	set oFunc = nothing
	set wscCrypto = nothing		
else
%>
	<h1>Invalid User</h1>
<%
end if 			
	Server.Execute(Application.Value("strWebRoot") & "Includes/footer.asp")
%>
