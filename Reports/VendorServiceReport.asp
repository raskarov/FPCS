<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		vendorServiceReport.asp
'Purpose:	Lists all students vendor services organized by vendor
'Date:		14 Jan 2005
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim intStudent_ID 
dim sql
dim mError		'conitains our error messages after validation is complete
dim strDiasbled 
dim strStudentName
dim arInfo
dim arFamInfo
dim bolPrint
dim  printCount
printCount = 0
intReporting_Period_ID = request("intReporting_Period_ID")

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

'Initialize some key variables
if ucase(session.Contents("strRole")) <> "ADMIN" then
	'terminate page since page was improperly called.
	response.Write "<html><body><H1>Page improperly called.</h1></body></html>"
	response.End
end if

if request("print") <> "" then bolPrint = true


'prevent page from cacheing
Response.AddHeader "Cache-Control","No-Cache"
Response.Expires = -1

'Print the header
Session.Value("strTitle") = "Vendor Service Report"
Session.Value("strLastUpdate") = "13 Jan 2005"

if request("SimpleHeader") <> "" or bolPrint then
	Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")
	if bolPrint then
	%>
	<link rel="stylesheet" href="<% =  Application.Value("strWebRoot") %>CSS/printStyle.css">
	<%
	end if
else
	Server.Execute(Application.Value("strWebRoot") & "includes/Header.asp")
end if

%>
<script language="javascript">
	function jfPrint(strURL){
		var winPrint;
		var sHide = document.main.hdnHide.value;
		var sRange = document.getElementById("sRange");
		
		if (sHide != ",") { strURL += "&hdnVendors=" + sHide;}
		if (sRange.value != "") { 
			var i;
			var sSelected = "";
			
			for (i=0;i < sRange.options.length;i++){
				if (sRange.options[i].selected == true) {
					if (sRange.options[i].value != "undefined") {
						sSelected += sRange.options[i].value + ",";
					}
				}
			}
			strURL += "&hdnRange=" + sSelected;
		}
		var winPrint = window.open(strURL,"winPrint","width=660,height=500,scrollbars=yes,resize=yes,resizable=yes");
		winPrint.moveTo(0,0);
		winPrint.focus();
	}		
</script>
<form name=main method=post action="vendorServiceReport.asp" ID="Form1">
<input type="hidden" name="SimpleHeader" value="<% = request("SimpleHeader") %>" ID="Hidden1">
<input type="hidden" name="hdnHide" value="," ID="Hidden2">
<table style="width:100%;" ID="Table3" cellpadding="2">
	<tr>
		<td class="yellowHeader" colspan="10">
			<table  class="yellowHeader" style="width:100%;" ID="Table1">
				<tr>
					<td class="yellowHeader" valign="top">
						<b>Vendor Service Report School Year <% = oFunc.SchoolYearRange %> <br>
						<% if request("intVendor_ID") & "" = "" then %>
						<div id="dCheck">Show Student Data: <input type="checkbox" name="detail" value="true" ID="Checkbox1" onclick="this.form.submit();" <% if request("detail") <> "" then response.write " checked " %>>
						<input type="button" value="print this page" class="btSmallGray" onclick="jfHideButtons();this.style.display='none';" ID="Button1" NAME="Button1"></div>						
						<% end if %>
					</td>
					<% if request("intVendor_ID") & "" = "" then %>
					<td align="right" class="yellowHeader" valign="top" style="width:5%;" nowrap id="t1">
						Detailed Report Print Options:
					</td>
					<td class="yellowHeader" valign="top" style="width:0%;"  id="t2">
						<select name="pRange" style="font-family:arial;size:7pt;" ID="sRange" multiple size=2>
							<option>Print All</option>
							<option value="'A'">A</option>
							<option value="'B'">B</option>
							<option value="'C'">C</option>
							<option value="'D'">D</option>
							<option value="'E'">E</option>
							<option value="'F'">F</option>
							<option value="'G'">G</option>
							<option value="'H'">H</option>
							<option value="'I'">I</option>
							<option value="'J'">J</option>
							<option value="'K'">K</option>
							<option value="'L'">L</option>
							<option value="'M'">M</option>
							<option value="'N'">N</option>
							<option value="'O'">O</option>
							<option value="'P'">P</option>
							<option value="'Q'">Q</option>
							<option value="'R'">R</option>
							<option value="'S'">S</option>
							<option value="'T'">T</option>
							<option value="'U'">U</option>
							<option value="'V'">V</option>
							<option value="'W'">W</option>
							<option value="'X'">X</option>
							<option value="'Y'">Y</option>
							<option value="'Z'">Z</option>
						</select>
					</td>
					<td align="left" class="yellowHeader" valign="top" style="width:0%;"  id="t3">
						<input type="button" value="print detail" class="btSmallGray" onclick="jfPrint('../Forms/PrintableForms/allPrintable.asp?strAction=V');" ID="Button2" NAME="Button1">
					</td>
					<% end if %>
				</tr>
			</table>
			
		</td>
	</tr>	
	<tr>
		<td class="svplain" colspan="10">
			<b>&nbsp;PLEASE NOTE: </b>
			<ul>
				<li>This report does NOT deal with all types of 
			'Services' just Vendor Services. This does NOT include 'Building Rental',
			'Equipment Rental', 'UAA', 'ASD' or 'Administrative Services'</li>
			<li>Expenses for courses OR budgets that have been
			rejected are NOT included in the data reported here.</li>
			</ul>
		</td>
	</tr>
						
<%

if isNumeric(request("intVendor_ID")&"") then
	strWhere = " AND tblVendors.intVENDOR_ID = " & request("intVendor_ID") & " " 
end if

if request("detail") <> "" then
	sql = "SELECT tblVendors.szVendor_Name, tblSTUDENT.szLAST_NAME, tblSTUDENT.szFIRST_NAME,  " & _ 
		" SUM(tblOrdered_Items.intQty * tblOrdered_Items.curUnit_Price + tblOrdered_Items.curShipping) AS total, tblILP_SHORT_FORM.szCourse_Title,  " & _ 
		" tblProgramOfStudies.txtCourseTitle, tblILP.bolApproved, tblILP.bolSponsor_Approved, tblILP.bolReady_For_Review, tblILP.intILP_ID,COUNT(1)AS myCOUNT," & _
		" tblVendors.intVendor_ID, tblOrdered_Items.intILP_ID, (select sum((li.intQuantity * li.curUnit_Price) + li.curShipping) from tblLine_Items li where li.intOrdered_Item_ID = tblOrdered_Items.intOrdered_Item_ID) as SpentTotal, " & _
		" tblOrdered_Items.bolClosed, tblStudent_States.intReEnroll_State " & _ 
		"FROM tblVendors INNER JOIN " & _ 
		" tblOrdered_Items ON tblVendors.intVendor_ID = tblOrdered_Items.intVendor_ID INNER JOIN " & _ 
		" tblSTUDENT ON tblOrdered_Items.intStudent_ID = tblSTUDENT.intSTUDENT_ID INNER JOIN " & _ 
		" tblILP ON tblOrdered_Items.intILP_ID = tblILP.intILP_ID INNER JOIN " & _ 
		" tblILP_SHORT_FORM ON tblILP.intShort_ILP_ID = tblILP_SHORT_FORM.intShort_ILP_ID LEFT OUTER JOIN " & _ 
		" tblProgramOfStudies ON tblILP_SHORT_FORM.lngPOS_ID = tblProgramOfStudies.lngPOS_ID INNER JOIN " & _ 
		"  tblStudent_States ON tblSTUDENT.intSTUDENT_ID = tblStudent_States.intStudent_id " & _ 
		"WHERE (tblOrdered_Items.intSchool_Year = " & session.Contents("intSchool_Year") & ") AND (tblOrdered_Items.intItem_ID = 3)  " & _ 
		" and (tblOrdered_Items.bolApproved = 1 or tblOrdered_Items.bolApproved is null) " & _
		" and ((tblILP.GuardianStatusID <> 3 or tblILP.GuardianStatusID is null) " & _
		" and (tblILP.SponsorStatusId  <> 3 or tblILP.SponsorStatusId is null) " & _
		" and (tblILP.InstructorStatusID <> 3 or tblILP.InstructorStatusID is null) " & _
		" and (tblILP.adminStatusID <> 3 or tblILP.adminStatusID is null)) " & _
		" AND (tblStudent_States.intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _
		strWhere & _
		"GROUP BY tblVendors.szVendor_Name, tblSTUDENT.szLAST_NAME, tblSTUDENT.szFIRST_NAME, tblILP_SHORT_FORM.szCourse_Title,  " & _ 
		" tblProgramOfStudies.txtCourseTitle, tblILP.bolApproved, tblILP.bolSponsor_Approved, tblILP.bolReady_For_Review, " & _
		" tblILP.intILP_ID,tblVendors.intVendor_ID, tblOrdered_Items.intILP_ID, tblOrdered_Items.intOrdered_Item_ID,tblOrdered_Items.bolClosed, tblStudent_States.intReEnroll_State " & _ 
		"ORDER BY tblVendors.szVendor_Name, tblSTUDENT.szLAST_NAME, tblSTUDENT.szFIRST_NAME, tblILP_SHORT_FORM.szCourse_Title,  " & _ 
		" tblProgramOfStudies.txtCourseTitle "
else	
	sql = "SELECT intVendor_ID,szVendor_Name, SUM(total) AS total,Mail_Addr, Mail_City, Mail_State, Mail_Zip,  " & _ 
			" szVendor_Email, szVendor_Phone, szPO_Number " & _ 
			"FROM (SELECT v.intVendor_ID,v.szVendor_Name, SUM(tblOrdered_Items.intQty * tblOrdered_Items.curUnit_Price + tblOrdered_Items.curShipping) AS total,  " & _ 
			" tblILP.intILP_ID, v.Mail_Addr, v.Mail_City, v.Mail_State,  " & _ 
			" v.Mail_Zip, v.szVendor_Email, v.szVendor_Phone, " & _ 
			" (select top 1 szPO_Number " & _
					" FROM tblLINE_ITEMS li LEFT OUTER JOIN " & _
					"	tblORDERED_ITEMS oi on oi.intORDERED_ITEM_ID =  li.intORDERED_ITEM_ID " & _
					" WHERE oi.intVendor_ID = v.intVendor_ID " & _
					" AND oi.intSCHOOL_YEAR =  " & session.Contents("intSchool_Year") & _
					" AND (oi.intItem_ID = 3) " & _
					" AND li.szPO_Number IS NOT NULL and li.szPO_Number <> '') as szPO_Number " & _
			" FROM v_VendorLabels v INNER JOIN " & _ 
			" tblOrdered_Items ON v.intVendor_ID = tblOrdered_Items.intVendor_ID INNER JOIN " & _ 
			" tblSTUDENT ON tblOrdered_Items.intStudent_ID = tblSTUDENT.intSTUDENT_ID INNER JOIN " & _ 
			" tblILP ON tblOrdered_Items.intILP_ID = tblILP.intILP_ID INNER JOIN " & _ 
			" tblILP_SHORT_FORM ON tblILP.intShort_ILP_ID = tblILP_SHORT_FORM.intShort_ILP_ID LEFT OUTER JOIN " & _ 
			" tblProgramOfStudies ON tblILP_SHORT_FORM.lngPOS_ID = tblProgramOfStudies.lngPOS_ID " & _			
			" WHERE (tblOrdered_Items.intSchool_Year = " & session.Contents("intSchool_Year") & ") AND (tblOrdered_Items.intItem_ID = 3)  " & _ 
			" and (tblOrdered_Items.bolApproved = 1 or tblOrdered_Items.bolApproved is null) " & _
			" and ((tblILP.GuardianStatusID <> 3 or tblILP.GuardianStatusID is null) " & _
			" and (tblILP.SponsorStatusId  <> 3 or tblILP.SponsorStatusId is null) " & _
			" and (tblILP.InstructorStatusID <> 3 or tblILP.InstructorStatusID is null) " & _
			" and (tblILP.adminStatusID <> 3 or tblILP.adminStatusID is null)) " & _
			strWhere & _
			" GROUP BY v.intVendor_ID, v.szVendor_Name, tblILP.bolApproved, tblILP.bolSponsor_Approved, tblILP.bolReady_For_Review, tblILP.intILP_ID, v.Mail_Addr,  " & _ 
			" v.Mail_City, v.Mail_State, v.Mail_Zip, v.szVendor_Email, v.szVendor_Phone) v " & _ 
			" GROUP BY  intVendor_ID,szVendor_Name, Mail_Addr, Mail_City, Mail_State, Mail_Zip, szVendor_Email,  " & _ 
			" szVendor_Phone, szPO_Number " & _ 
			"ORDER BY szVendor_Name "
end if			
	
		'response.Write sql 
	dim rs 
	set rs = server.CreateObject("ADODB.RECORDSET")
	dim rs2 
	set rs2 = server.CreateObject("ADODB.RECORDSET")
	dim strVendName
	dim dblGrandTotal
	dim count
	
	rs2.CursorLocation = 3
	rs.CursorLocation = 3
	rs.Open sql, oFunc.FPCScnn
	strVendName = ""
	intVendor_ID = ""
	if rs.RecordCount > 0 then
		do while not rs.EOF
			if request("detail") <> "" then
				if strVendName <> rs("szVendor_Name") then												
					
					if strVendName  <> "" then
						dblGrandTotal = dblGrandTotal + dblSubTotal
						dblGrandSpentTotal = dblGrandSpentTotal + dblSubSpentTotal
						dblGrandAdjust = dblGrandAdjust + dblSubAdjust 
						response.Write vbfPONumber(szPO_Number) & " <td class='svplain8' align='right'><b>totals:</b></td><td class='svplain8' align='right'><B>$" & _
										formatNumber(dblSubTotal,2) & "</b></td>"  & _	
										"<td class='svplain8' align='right'><b>$" & _
										formatNumber(dblSubSpentTotal,2) & "</b></td>" & _
										"<td class='svplain8' align='right'><b>$" & _
										formatNumber(dblSubAdjust,2) & "</b></td>" & _	
										"<td class='svplain8' align='right'><b>$" & _
										formatNumber((dblSubTotal-dblSubSpentTotal) + dblSubAdjust,2) & "</b></td>" & _			
										"</tr></table>" 
						if request("intVendor_ID") & "" = "" then 
							response.Write "<input type='button' value='hide' id='bt" & intVendor_ID & "' onclick=""jfToggle('" & intVendor_ID & "','" & dblSubTotal & "');"" class=""btSmallGray"" title=""" & strVendName  & """>" 
						end if
						
						response.Write "</td></tr>"																		
						strBtList = strBtList & "bt" &  intVendor_ID
					end if 
					
					dblSubTotal = 0
					dblSubSpentTotal = 0
					dblSubAdjust = 0
					strVendName = rs("szVendor_Name")
					intVendor_ID = rs("intVendor_ID")	
					response.Write vbsTableHeader(intVendor_ID)
					szPO_Number = ""
				end if
				
				if rs("SpentTotal") & "" = "" then 
					mSpentTotal = 0
				else
					mSpentTotal	= rs("SpentTotal")
				end if
				
				if rs("bolClosed") or (rs("total")-mSpentTotal < 0) then
					iAdjust =  (rs("total")-mSpentTotal) * -1
				else
					iAdjust = 0
				end if
				
				if not instr(1,"," & Application.Contents("ActiveEnrollList") & "," ,"," & rs("intReEnroll_State") & ",") > 0 then				
					sName = "<span style='color:red;' title='Student is not actively enrolled.'>" & rs("szLAST_NAME") & ", " & rs("szFIRST_NAME") & "</span>"
				else
					sName = rs("szLAST_NAME") & ", " & rs("szFIRST_NAME")
				end if 
	%>		
						<td class="TableCell">
							<% = rs("szCourse_Title") & rs("txtCourseTitle") %>
						</td>
						<td class="TableCell">
							<% = sName  %>
						</td>
						<td class="TableCell" align="right">
							$<% = formatNumber(rs("total"),2) %>
						</td>
						<td class="TableCell" align="right">
							$<% = formatNumber(mSpentTotal,2) %>
						</td>
						<td class="TableCell" align="right">
							$<% = formatNumber(iAdjust,2) %>
						</td>
						<td class="TableCell" align="right">
							$<% = formatNumber((rs("total")-mSpentTotal) + iAdjust,2) %>
						</td>
					</tr>
	<%						
				dblSubTotal = dblSubTotal + rs("total")		
				dblSubSpentTotal = dblSubSpentTotal + mSpentTotal	
				dblSubAdjust = dblSubAdjust + iAdjust
				if szPO_Number = "" then
					sql = "select szPO_Number " & _
						" FROM tblLINE_ITEMS LEFT OUTER JOIN " & _
						"	tblORDERED_ITEMS on tblORDERED_ITEMS.intORDERED_ITEM_ID =  tblLINE_ITEMS.intORDERED_ITEM_ID " & _
						" WHERE tblORDERED_ITEMS.intILP_ID = " & rs("intILP_ID")  & _
						" AND tblORDERED_ITEMS.intVendor_ID = " & intVendor_ID
					rs2.Open sql, oFunc.FPCScnn
					if rs2.RecordCount > 0 then
						if rs2("szPO_Number") & "" <> "" then
							szPO_Number = rs2("szPO_Number")
						end if
					end if
					rs2.Close
				end if	
			else
				if iCount mod 25 = 0 then
					call vbsTableHeader2	
				end if
				
	%>		
					<tr ID="Vendor<%= rs("intVendor_ID") %>">
						<td class="TableCell">
							<% = rs("szVendor_Name") %>
						</td>
						<td class="TableCell">
							<% = rs("Mail_Addr") & "<BR>" & rs("Mail_City") & ", " & rs("Mail_State") & " " & rs("Mail_Zip") %>&nbsp;
						</td>
						<td class="TableCell" align=center>
							<% = oFunc.FormatPhone(rs("szVendor_Phone")) %>&nbsp;
						</td>
						<td class="TableCell" align=center>
							<a href="mailto:<% = rs("szVendor_Email") %>"><% = rs("szVendor_Email") %></a>&nbsp;
						</td>
						<td class="TableCell" align=center>
							<% = rs("szPO_NUMBER")  %>&nbsp;
						</td>
						<td class="TableCell" align="right">
							$<% = formatNumber(rs("total"),2) %>&nbsp;
						</td>					
						<td>
							<input type='button' value='hide' id='bt<%= rs("intVendor_ID") %>' onclick="jfToggle('<% = rs("intVendor_ID") %>','<% = rs("total") %>');" class="btSmallGray" title="<%=strVendName%>" NAME=bt<%= rs("intVendor_ID") %>">
						</td>
					</tr>
	<%				
				dblSubTotal = dblSubTotal + rs("total")		
				iCount = iCount + 1
				strBtList = strBtList & "bt" &  rs("intVendor_ID")	& ","
			end if					
			rs.MoveNext
		loop
	end if

	if request("detail") <> "" then
		dblGrandTotal = dblGrandTotal + dblSubTotal
		dblGrandSpentTotal = dblGrandSpentTotal + dblSubSpentTotal
		dblGrandAdjust = dblGrandAdjust + dblSubAdjust 
		response.Write vbfPONumber(szPO_Number) & " <td class='svplain8' align='right'><b>totals:</b></td><td class='svplain8' align='right'><B>$" & _
						formatNumber(dblSubTotal,2) & "</b></td>"  & _	
						"<td class='svplain8' align='right'><b>$" & _
						formatNumber(dblSubSpentTotal,2) & "</b></td>" & _
						"<td class='svplain8' align='right'><b>$" & _
						formatNumber(dblSubAdjust,2) & "</b></td>" & _	
						"<td class='svplain8' align='right'><b>$" & _
						formatNumber((dblSubTotal-dblSubSpentTotal) + dblSubAdjust,2) & "</b></td>" & _			
						"</tr></table>" 
		if request("intVendor_ID") & "" = "" then 
			response.Write "<input type='button' value='hide' id='bt" & intVendor_ID & "' onclick=""jfToggle('" & intVendor_ID & "','" & dblSubTotal & "');"" class=""btSmallGray"" title=""" & strVendName  & """>" 
		end if
		
		response.Write "</td></tr>"				
						
		if request("intVendor_ID") & "" = "" then 
			response.Write "<tr><td><table style='width:650px;'>" & _					
					"<tr><td style='width:100%;' class='svplain10' align='right'><b>Grand Budgeted Total:</b></td><td class='svplain10' align='right' nowrap><b>" & _
					"<div id='gTotal'>$" & _
					formatNumber(dblGrandTotal,2) & "</div></b></td></tr></table></td></tr>"
		end if
		strBtList = strBtList & "bt" &  intVendor_ID	
%>
				</table>
			</td>
		</tr>
<% else 
		dblGrandTotal = dblSubTotal
		if strBtList <> "" then
			strBtList = left(strBtList,len(strBtList)-1)
		end if
%>
		<tr>
			<td colspan="5" class="svplain8" align="right" valign="top">
			<b>Total:</b>
			</td>
			<td align="right" class="svplain8"  valign="top">
			<b><div id='gTotal'>$<% = formatNumber(dblGrandTotal,2) %></div>&nbsp;</b>
			</td>
		</tr>
<% end if %>
	<tr>
		<td></td>
	</tr>
</table>
<script language="javascript">
	var gTotal = parseFloat('<% = dblGrandTotal %>');
	
	function jfToggle(pID,pSubTotal){
		var obj = document.getElementById('Vendor' + pID);
		var obj2 = document.getElementById('gTotal');
		var oBtn = document.getElementById('bt' + pID);
		var amount = parseFloat(pSubTotal);
		var sHide = document.main.hdnHide.value;
		
		if (obj.style.display == 'none') {
			obj.style.display = 'block';
			gTotal = gTotal + amount;
			obj2.innerHTML = "$" + formatnum(gTotal);
			sHide = sHide.replace(","+pID+",",",");
			oBtn.value = "hide";
			
		}else{
			obj.style.display = 'none';
			gTotal = gTotal - amount;
			obj2.innerHTML = "$" + formatnum(gTotal);
			sHide += pID + ",";
			oBtn.value = "show";
		}
		document.main.hdnHide.value = sHide;
	}
	
	function formatnum(valuein){ 
		valuein = "" + Math.round( parseFloat(valuein) * 100 ) / 100 
		if (valuein=="NaN") valuein = "0"
		decimalpos = valuein.indexOf(".")
		if (decimalpos==-1)
		{ decimalpos = valuein.length
		valuein = valuein + "." }
		valuein = valuein + "00"
		valueout = valuein.substring(decimalpos,decimalpos+3)
		valuein = valuein.substring(0, decimalpos)
		while (valuein.length>3)
		{
		valueout = "," + valuein.substring(valuein.length-3,valuein.length) + valueout
		valuein = valuein.substring(0, valuein.length-3)


		}


		valueout = valuein + valueout
		return valueout
	}

	function jfHideButtons(){
		var sList = "<% = strBtList %>";
		var aList = sList.split(",");
		var i;
		var obj;
		
		for(i=0;i<aList.length;i++){
			obj = document.getElementById(aList[i]);
			obj.style.display = "none";	
		}
		obj = document.getElementById("t1");
		obj.style.display = "none";	
		
		obj = document.getElementById("t2");
		obj.style.display = "none";	
		
		obj = document.getElementById("t3");
		obj.style.display = "none";	
		
		obj = document.getElementById("dCheck");
		obj.style.display = "none";	
		
		if (window.print){
	      window.print()
	    }
	    else {
	      alert("Mac users: please press Apple-P to print this form.\nWindows users: Please press ctrl-P to print this form.")
		}
	}
</script>
</form>
<%
rs.Close
set rs = nothing
set rs2 = nothing
call oFunc.CloseCN()
set oFunc = nothing
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")

function vbsTableHeader(pID)
%>
	<tr>
		<td align="right">			
			<table ID="Vendor<% = pID%>" style="width:650px;" align="left">	
				<tr>
					<td class="TableCell" style="width:30%;">
						Vendor Name
					</td>
					<td class="TableCell" style="width:30%;">
						Course Name
					</td>
					<td class="TableCell" style="width:30%;">
						Student Name
					</td>
					<td class="TableCell" style="width:10%;" align="center">
						Budget
					</td>	
					<td class="TableCell" style="width:10%;" align="center">
						Spent
					</td>	
					<td class="TableCell" style="width:10%;" align="center">
						Budget Adjustment
					</td>
					<td class="TableCell" style="width:10%;" align="center">
						Balance
					</td>	
				</tr>
				<tr>					
					<td class="svplain8" rowspan="100" valign="top">
						<b><% = rs("szVendor_Name") %></b>
					</td>				
<%
end function

function vbsTableHeader2()
	if iCount = 0 then
%>
				<tr>
					<td colspan="6"><br></td>
				</tr>
<%		
	else
%>
				<tr>
					<td colspan="6"><p></p></td>
				</tr>
<%
	end if
%>
				<tr>
					<td class="TableHeaderBlue" style="width:30%;">
						<b>Vendor Name</b>
					</td>
					<td class="TableHeaderBlue" style="width:30%;">
						<b>Address</b>
					</td>
					<td class="TableHeaderBlue" style="width:30%;">
						<b>Phone Number</b>
					</td>
					<td class="TableHeaderBlue" style="width:10%;">
						<b>Email</b>
					</td>	
					<td class="TableHeaderBlue" style="width:10%;" align="center">
						<b>PO #</b>
					</td>
					<td class="TableHeaderBlue" style="width:10%;" align="center">
						<b>Budget Total</b>
					</td>
				</tr>			
<%
end function

function vbfPONumber(pNum)
%>
				<tr>
					<td class="svplain8" align="left">
						<table border=1 cellspacing=0 ID="Table33" bordercolor="#e6e6e6">
							<tr>
								<td class="svplain8">
									PO#:
									<% if pNum <> "" then 
										response.Write "&nbsp;" & pNum
									   else
									%>									
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									<%
									  end if
									%>
								</td>
							</tr>
						</table>
					</td>	
<%
end function 
%>