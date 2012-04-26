<%@ Language=VBScript %>
<%

'Response.Write Request.ServerVariables("URL") & "<br/>"
'for each x in Request.QueryString
'        Response.Write("<br>" & x & " = " & Request.QueryString(x)) 
'    next 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		VensdorSearchEngine.asp
'Purpose:	Vendor Search Engine
'Date:		June 17 2005
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim sql
dim oFunc
dim rs, strSqlFamily

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

if request("bolWin") = "" then
	Server.Execute(Application.Value("strWebRoot") & "includes/header.asp")
else
	Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")
end if

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
				 "&intPOS_Subject_ID=" & request("intPOS_Subject_ID") & _
				 "&bolWin=" & request("bolWin")
%>
<script language=javascript>

	function jfVendProfile(pVendorID){
		var winVendProfile;
		var strURL = "<%=Application.Value("strWebRoot")%>forms/VIS/VendorAdmin.asp?intVendor_ID="+pVendorID;
		strURL += "&bolPrint=true&bolWin=True&intItem_Group_ID=<% = request("intItem_Group_ID")%>";
		winVendProfile = window.open(strURL,"winVendProfile","width=850,height=500,scrollbars=yes,resize=yes,resizable=yes");
		winVendProfile.moveTo(20,20);
		winVendProfile.focus();	
	}
	
	function jfSelectVendor(pVendId){
		var qString = "<%=Application.Value("strWebRoot")%>forms/requisitions/reqGoods.asp?intVendor_ID=" + pVendId;
		qString += "<% = strVendQString%>";
		window.opener.location.href = qString;
		window.opener.focus();
		window.close();
	}
	
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
</script>	
<form name="main" method="post" action="VendorSearchEngine.asp" ID="Form1">
<input type="hidden" name="Search" value="true" ID="Hidden1">
<input type=hidden name="lastRow" ID="Hidden2">
<input type=hidden name="LineItemsChanged" value="," ID="Hidden8">
<input type=hidden name="lastRowColor" ID="Hidden3">
<input type="hidden" name="hdnReset" value="" ID="Hidden4">
<!-- Remaining hidden variables used when using this page to add a class from forms/ilp/ilp1.asp -->
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
<table style="width:100%;" ID="Table3">
	<tr>
		<td class="yellowHeader">
			&nbsp;<b>Vendor Search Engine</b>
		</td>
	</tr>
	<tr>
		<td>
			<table  style="width:100%;">
				<tr>
					<td>										
						<table cellpadding="2" ID="Table2">
							<tr>
								<td class="TableHeader" colspan="2">
									Vendor Name
								</td>
							</tr>									
							<tr>
								<td colspan="2">
									<select name="intVendor_ID" onchange="this.form.hdnReset.value='true';" ID="Select1"  style="width:100%;">
										<option value="">
									<%
										sql2 = "SELECT     TOP 1 upper(szVendor_Status_CD) " & _ 
										"	FROM          tblVendor_Status vs " & _ 
										"	WHERE      vs.intVendor_ID = v.intVendor_ID AND vs.intSchool_Year <= " & session.Contents("intSchool_Year") & _ 
										"	ORDER BY intSchool_Year DESC, intVendor_Status_ID DESC" 

										if request("intItem_Group_ID") <> "" then

											select case request("intItem_Group_ID") 
												case 1
													sql =" v.bolService_Vendor = 1 AND " 
													if request("bolReimburse") <> "" then
				                                        sql = sql & " v.bolNonProfit = 1 AND "
				                                    end if
													sql2 = " SELECT     TOP 1 upper(szVendor_Status_CD) " & _ 
														"	FROM          tblVendor_Status vs " & _ 
														"	WHERE      (vs.intVendor_ID = v.intVendor_ID AND vs.intSchool_Year = " & session.Contents("intSchool_Year") & _ 
														"       and vs.dtContract_Start is not null) or (vs.intVendor_ID = v.intVendor_ID AND vs.intSchool_Year <= " & session.Contents("intSchool_Year") & " and  v.bolNonProfit = 1) " & _
														"	ORDER BY intSchool_Year DESC, intVendor_Status_ID DESC "
												case 2
													sql = " v.bolGoods_Vendor = 1 AND " 
												case 3	
													sql = "  v.bolGoods_Vendor = 1 AND " 
													sql = "  v.bolService_Vendor = 1 AND " 
											end select		
							
										end if 

							sqlVendor = "SELECT intVendor_ID,  " & _ 
									" szVendor_Name AS Vend_Name " & _ 
									"FROM tblVendors v WHERE " & sql & _ 
									" (" & sql2 & ") IN ('APPR') " & _
									" ORDER BY Vend_Name "
										
							Response.Write oFunc.MakeListSQL(sqlVendor,"intVendor_ID","Vend_Name",request.form("intVendor_ID"))					
									%>
									</select>
								</td>
							</tr>
							<% if request("intItem_Group_ID") = "1" or request("intItem_Group_ID") = "" then %>					
							<tr>
								<td class="TableHeader" colspan="2">
									Services 
								</td>
							</tr>
							<tr>
								<td colspan="2">
									<select name="intVend_Service_ID" ID="Select2" onchange="this.form.hdnReset.value='true';" >
										<option value="">
									<%
										sql = "SELECT vs.intVend_Service_ID, UPPER(ps.szSubject_Name + ': ' + vs.szVend_Service_Name) AS ServiceName " & _ 
												"FROM trefVendor_Services vs INNER JOIN " & _ 
												" trefPOS_Subjects ps ON vs.intPOS_Subject_ID = ps.intPOS_Subject_ID " & _ 
												" WHERE vs.Is_Active = 1 " & _
												"ORDER BY ps.szSubject_Name + ': ' + vs.szVend_Service_Name "
										Response.Write oFunc.MakeListSQL(sql,"intVend_Service_ID","ServiceName",request("intVend_Service_ID"))		
									%>
									</select>	
								</td>
							</tr>
							<% end if %>
							<tr>
								<td class="TableHeader" >
									Key Word(s) to Search
								</td>		
								<td class="TableHeader" colspan="2" nowrap style='width:0%;'>
									Vendor Type
								</td>						
							</tr>
							<tr>
								<td>
									<input type="text" name="KeyWords" style="width:100%;"  maxlength="128" value="<% = request("KeyWords") %>" onchange="this.form.hdnReset.value='true';" ID="Text1">
								</td>	
								<td  style='width:0%;'>									
									<select name="xGoodService" ID="Select3"  <% if request("intItem_Group_ID") <> "" then response.Write " disabled " %> onchange="this.form.hdnReset.value='true';" >
										<%
											Response.Write oFunc.MakeList("0,1,2,3",",Service Vendor,Goods Vendor,Both",request("xGoodService") & request("intItem_Group_ID"))								
										%>
									</select>	
									<% if request("intItem_Group_ID") <> "" then %>
									<input type=hidden name="xGoodService" value="<% = request("intItem_Group_ID")%>">
									<% end if %>
								</td>							
							</tr>
							<tr>
								<td class="gray">
									Key Words: Match Exact Words <input type="checkbox" name="searchType" value="exact" <% if request("searchType") <> "" then response.Write " checked "%>  value="true" ID="Radio1" onchange="this.form.hdnReset.value='true';">
								</td>
								<td nowrap>
									<input type="submit" value="Search!" class="NavSave" ID="Submit1" NAME="Submit1">
									<% if request("bolWin") <> "" then
									%>
									<input type="button" value="Close Window" onclick="window.opener.focus();window.close();" class="btSmallGray" ID="Button1" NAME="Button1">
									<%
										end if
									%>
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>	
		</td>
	</tr>
</table>	
<%

if request.Form("Search") <> "" then
	sql = "SELECT DISTINCT v.intVendor_ID, v.szVendor_Name, v.szVendor_Phone, v.szContact_First_Name, v.szContact_Last_Name, v.szVendor_Email, v.szVendor_Website " & _
		   "FROM         tblVendors v left outer join " & _
		   "	tascVendor_Service_Types st on v.intVendor_ID = st.intVendor_ID WHERE 1=1 "

	if Request.Form("keywords") <> "" then		
		if Request.Form("searchType") <> ""  then	
			strKeyWords = " like upper('%" & oFunc.EscapeTick(Request.Form("keywords"))& "%') " 
			sql = sql & " and (upper(convert(varChar(1000),substring(v.szPrev_Experience,1,1000)))" & strKeyWords & " or " & _
					"upper(convert(varChar(1000),substring(v.szTraining,1,1000)))" & strKeyWords & " or " & _
					"upper(convert(varChar(2000),substring(v.szVendor_Comments,1,2000)))" & strKeyWords & " or " & _
					"upper(v.szVendor_Name)" & strKeyWords & " or upper(szContact_First_Name) " & strKeyWords & _
					" or upper(szContact_Last_Name) " & strKeyWords & ") " 
		else
			arWords = split(Request.Form("keywords")," ")
			if isArray(arWords) then
				sql = sql & " and ("
				for i = 0 to ubound(arWords)
					strKeyWords = " like upper('%" & oFunc.EscapeTick(arWords(i))& "%') "
					sql = sql & " upper(convert(varChar(1000),substring(v.szPrev_Experience,1,1000)))" & strKeyWords & " or " & _
					"upper(convert(varChar(1000),substring(v.szTraining,1,1000)))" & strKeyWords & " or " & _
					"upper(convert(varChar(2000),substring(v.szVendor_Comments,1,2000)))" & strKeyWords & " or " & _
					"upper(szVendor_Name)" & strKeyWords & " or " & _
					" upper(szContact_First_Name) " & strKeyWords & _
					" or upper(szContact_Last_Name) " & strKeyWords & " or " 
				next	
				sql = left(sql,len(sql)-3) 	
				sql = sql & ") "	
			end if 
		end if 
	end if 		
	
	if request("intVend_Service_ID") <> "" then
		sql = sql & " AND st.intVend_Service_ID = " & request("intVend_Service_ID") & " " 
	end if	
	
	if request("intVendor_ID") <> "" then
		sql = sql & " AND v.intVendor_ID = " & request("intVendor_ID") & " " 
	end if	
		
	sql = sql & " AND (select top 1 upper(szVendor_Status_CD) from tblVendor_Status vs where " & _
				" vs.intVendor_ID = v.intVendor_ID and ((vs.dtContract_Start is not null " & _
				" and (v.bolNonProfit = 0 or v.bolNonProfit is null) and v.bolService_Vendor = 1) " & _
				" or (v.bolService_Vendor is null or v.bolService_Vendor <> 1) or " & _
				" (v.bolService_Vendor = 1 and  v.bolNonProfit = 1)) and " & _
				" vs.intSchool_Year <= " & session.Contents("intSchool_Year") & _ 
			    "  order by intSchool_Year desc, intVendor_Status_ID desc) in ('APPR') "

	if request("xGoodService") <> "" then
		select case request("xGoodService")
			case 1
				sql = sql & " AND v.bolService_Vendor = 1 "
				if request("bolReimburse") <> "" then
				    sql = sql & " AND v.bolNonProfit = 1 "
				end if
			case 2
				sql = sql & " AND v.bolGoods_Vendor = 1 " 
			case 3	
				sql = sql & " AND v.bolGoods_Vendor = 1 " 
				sql = sql & " AND v.bolService_Vendor = 1 " 
		end select		
	end if 
	if request("orderby") <> "" then
		sql = sql & " ORDER BY " & request("orderby")
	else
		sql = sql & " ORDER BY szVendor_Name" 
	end if
	
	'response.Write sql 
	set rs = server.CreateObject("ADODB.RECORDSET")
	rs.CursorLocation = 3
	rs.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
	
	if request("PageNumber") <> "" and request("hdnReset") = "" then
		intPageNum = cint(request("PageNumber"))	
	else
		intPageNum = 1
	end if
	
	with rs
		if .RecordCount > 0 then
			.PageSize = 25
			.AbsolutePage = intPageNum
			intViewingTo = .AbsolutePosition + .PageSize -1 
			if intViewingTo > .recordcount then intViewingTo = .RecordCount
%>
<br>
<input type="hidden" name="PageNumber" value="<% = intPageNum%>" ID="Hidden11">
<table cellpadding="2" ID="Table5">
	<tr>
		<td colspan=10 class="svplain8" nowrap>
			
			Viewing <% = .AbsolutePosition %> - <% = intViewingTo %>  of <% = .RecordCount %> Matches &nbsp;
			
			<table ID="Table6" cellpadding="2"><tr><td>
			<%
				if cint(.RecordCount) > cint(.PageSize) then
					for i = 1 to .PageCount
					%>
					<input type="button" class="btSmallWhite" value="<%=i%>" onClick="this.form.PageNumber.value='<%=i%>';this.form.submit();" ID="Button2" NAME="Button2">
					<%
					next 
				end if
			%>
			</td></tr></table>
		</td>					
	</tr>
<%			
			intCount = 0
			intCount2 = 0 
			intMax = (.AbsolutePosition + .PageSize)
			
			do while .AbsolutePosition < intMax and not .EOF
				if intCount mod 2 = 0 then
					strColor = "TableCell"
				else
					strColor = "gray"
				end if
				
				if intCount2 = 0 or intCount2 mod .PageSize = 0 then
					call PrintHeader
				end if
				
%>	
			<tr id="ROW<%=intCount%>" onClick="jfHighLight('<%=intCount%>');" class="<% = strColor %>">				
				<% if request("bolWin") <> "" then %>
				<td class="TableHeader" align="center">
					<a href="javascript:"  style="color:white;" onclick="jfSelectVendor('<% = rs("intVendor_ID")%>');">select</a>
				</td>
				<% end if %>
				<td >
					<a href="#" onclick="jfVendProfile('<% = rs("intVendor_ID") %>');"><% = rs("szVendor_Name") %></a>
				</td>
				<td >
					<% = rs("szContact_FIRST_NAME") & " " & rs("szContact_LAST_NAME") %>
				</td>
				<td >
					<% = ucase(rs("szVendor_Phone")) %>
				</td>
				<td align="center">
					<% if rs("szVendor_Email") & "" <> ""  then %>
					<a href="mailto:<% = rs("szVendor_Email") %>"><% = rs("szVendor_Email") %></a>
					<% else %>
					&nbsp;
					<% end if %>
				</td>
				<td align="center">
					<% if rs("szVendor_Website") & "" <> ""  then %>
					<a href="<% if instr(1,ucase(rs("szVendor_Website")),"HTTP") < 1 then  response.Write "http://" & rs("szVendor_Website") else response.Write rs("szVendor_Website")%>" target="_blank">Website</a>
					<% else %>
					&nbsp;
					<% end if %>
				</td>
			</tr>
<%				
				.MoveNext
				intCount = intCount + 1
				intCount2 = intCount2 + 1
			loop
%>
	<input type=hidden name="intCount" value="<%=intCount%>" ID="Hidden12">
	<input type=hidden name="intCount2" value="<%=intCount2%>" ID="Hidden13">
	<input type="hidden" name="orderby" value="<% = request("orderby") %>" ID="Hidden14">
</table>
<%		
		else
			%>
			<span class="svplain8"><B>0 Matches Found.</B></span>
			<%
		end if
		.close		
	end with
end if

%>
</form>
<%
call oFunc.CloseCN()
set oFunc = nothing
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")


function PrintHeader
%>
	<Tr>
		<% if request("bolWin") <> "" then %>
		<td class="TableHeader" align="center">
			Select Vendor
		</td>
		<% end if %>
		<td class="TableHeader">
			<a href="#" class="linkWht" onclick="document.forms[0].orderby.value='v.szVendor_Name';document.forms[0].submit();">Vendor Name</a>
		</td>
		<td class="TableHeader">
			<a href="#" class="linkWht" onclick="document.forms[0].orderby.value=' v.szCONTACT_LAST_NAME,v.szCONTACT_FIRST_NAME';document.forms[0].submit();">Contact Name</a>
		</td>
		<td class="TableHeader">
			Phone
		</td>
		<td class="TableHeader" align="center">
			Email
		</td>
		<td class="TableHeader">
			Website
		</td>
	</Tr>
<%
end function
%>