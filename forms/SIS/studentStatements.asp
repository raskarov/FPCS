<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		studentStatements.asp  
'Purpose:	This script creates a form that an admin can use to
'			filter student statements and creates the statements 
'			based on form selections.
'Date:		29 JAN 2003
'Author:	Scott M. Bacon(ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim dblTrack		'Keeps track of funding totals
dim strStudent_Name

Session.Value("strTitle") = "Student Statements"
Session.Value("strLastUpdate") = "31 JAN 2003"
server.ScriptTimeout = 3600

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

' Either shows our form or creates the statements
if request.QueryString("intStudent_ID") = "" and request.QueryString("intFamily_ID") = "" _
	and session.Contents("strRole") = "ADMIN" then
	call vbsPrintForm
else
	start = now()
	dim objInstruct_Dict
	set objInstruct_Dict = server.CreateObject("SCRIPTING.DICTIONARY")
	call vbsShowStatement
	endIt = now()
	if session.Contents("strRole") = "ADMIN" then
		response.Write "This took " & formatNumber((datediff("s",start,endIt)/60),2) & " minutes<BR>" 
	end if	
end if

' Close things up 
call oFunc.CloseCN()
set oFunc = nothing
Server.Execute(Application.Value("strWebRoot") & "Includes/footer.asp")


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Subs and Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sub vbsPrintForm
	Server.Execute(Application.Value("strWebRoot") & "includes/Header.asp")
%>
	<form name=main onsubmit="return false;" ID="Form1">
	<input type=hidden name=changed value="" ID="Hidden1">
	<table width=100% ID="Table1">
		<tr>	
			<td colspan=2 class=yellowHeader>
					&nbsp;<b>Student Statements</b>&nbsp;&nbsp;&nbsp;
			</td>
		</tr>		
	</table>
	<table>
		<!--<tr>
			<td class="gray">
				Select Family/Families
			</td>
			<td>
				<select name="intFamily_ID" multiple size=5 onchange="main.intStudent_ID.disabled=true;">
						<option value="">
						<option value="all">All Families
					<%
						'dim sqlFamily 
						'sqlFamily =     "SELECT DISTINCT f.intFamily_ID, f.szFamily_Name + ': ' + f.szDesc AS family " & _
						'				"FROM tblFAMILY f INNER JOIN " & _
						'				" tblSTUDENT s ON f.intFamily_ID = s.intFamily_ID INNER JOIN " & _
						'				" tblStudent_States ss ON s.intSTUDENT_ID = ss.intStudent_id " & _
						'				"WHERE (ss.intSchool_Year = 2003) " & _
						'				"ORDER BY family"
						'Response.Write oFunc.MakeListSQL(sqlFamily,"intFamily_ID","family",intFamily_ID)												 
					%>
				</select>
			</td>
		</tr> -->
		<input type=hidden name="intFamily_ID" value="">
		<tr>
			<td class="gray">
				Select Student(s)
			</td>
			<td>
				<select name="intStudent_ID" multiple size=5> <!--onchange="main.intFamily_ID.disabled=true;"-->
						<option value="">
						<option value="all">All Students</option>
						<option value="active">Active Students Only</option>
						<option value="withdrawn">Withdrawn Students</option>
					<%
						dim sqlStudents
						sqlStudents =   "SELECT s.intSTUDENT_ID, s.szLAST_NAME + ', ' + s.szFIRST_NAME AS Name " & _
										"FROM tblSTUDENT s INNER JOIN " & _
										" tblENROLL_INFO ei ON s.intSTUDENT_ID = ei.intSTUDENT_ID " & _
										"WHERE (ei.sintSCHOOL_YEAR = " & SESSION.Contents("intSchool_Year") & ") " & _
										"ORDER BY s.szLAST_NAME + ', ' + s.szFIRST_NAME"
											
						Response.Write oFunc.MakeListSQL(sqlStudents,"intSTUDENT_ID","Name","")												 
					%>
				</select>
			</td>
		</tr>
	</table>
	<script language=javascript>
		function jfSubmit(){
			// determines parameters to send and opens a new window to display statements			
			var obj;
			
			if (main.intStudent_ID.value != "") {
				obj = main.intStudent_ID;
			}else if (main.intFamily_ID.value != ""){
				obj = main.intFamily_ID;
			}else{
				return false;
			}
			
			var i;
			var strItems = "";
			
			if (obj.options[1].selected == true) {
				for (i=0; i < obj.length; i++) {				
					strItems = strItems + obj.options[i].value + ",";
				}
			}else{	
				for (i=0; i < obj.length; i++) {				
					if (obj.options[i].selected == true){
						strItems = strItems + obj.options[i].value + ",";
					}
				}
			}
			var winStatements;			
			
			strItems = strItems.substr(0, strItems.length - 1); 
			winStatements = window.open("<%=Application("strWebRoot")%>forms/sis/studentStatements.asp?bolSimple=true&"+obj.name+"="+strItems,"winStatements","width=800,height=600,scrollbars=yes,resizable=yes");
			winStatements.moveTo(0,0)
			winStatements.focus();
		}
	</script>
	<input type=button value="Home Page" onClick="window.location.href='<%=Application("strWebRoot")%>';" id="btSmallGray" NAME="btSmallGray">
	<input type=reset value="Reset Form" id="btSmallGray" >
	<!--<input type=reset value="Reset Form" id="Reset1" onclick="main.intFamily_ID.disabled=false;main.intStudent_ID.disabled=false;" NAME="Reset1">-->
	<input type=button value="Get Statements" id="btSmallGray" onclick="jfSubmit();" NAME="Submit1">
	</form>	
		
<%
end sub

sub vbsShowStatement
	' This sub does much of the work needed to create our student statements
	' Start by printing HTML header
	if request("bolSimple") <> "" then
		Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")
	else
		Server.Execute(Application.Value("strWebRoot") & "includes/Header.asp")
	end if
	
	' Dimention variables
	dim sql
	dim pintStudent_ID
	dim intFamily_ID
	dim intPercent	
	dim strAddClause
	dim intCount
	
	' Are we handling students or families?
	if request.QueryString("intStudent_ID") <> "" then
		pintStudent_ID = request.QueryString("intStudent_ID")
	else
		intFamily_ID = request.QueryString("intFamily_ID")
	end if
	
	' Get specific students only
	if pintStudent_ID = "active" then
		' Only return Active Students
		set rsGetStudents = server.CreateObject("ADODB.RECORDSET")
		rsGetStudents.CursorLocation = 3
		
		sql = "SELECT tblSTUDENT.intSTUDENT_ID " & _
				"FROM tblSTUDENT INNER JOIN " & _
				" tblStudent_States ON tblSTUDENT.intSTUDENT_ID = tblStudent_States.intStudent_id " & _
				"WHERE tblStudent_States.intReEnroll_State IN (" & Application.Contents("ActiveEnrollList") & ")  " & _
				" AND (tblStudent_States.intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _
				"ORDER BY tblSTUDENT.szLAST_NAME, tblSTUDENT.szFIRST_NAME"
		rsGetStudents.Open sql, Application("cnnFPCS")'oFunc.FPCScnn

		reDim arStudents((rsGetStudents.RecordCount-1))
		intCount = 0
		do while not rsGetStudents.EOF
			arStudents(intCount) = rsGetStudents("intSTUDENT_ID")
			rsGetStudents.MoveNext
			intCount = intCount + 1
		loop
		rsGetStudents.Close
		set rsGetStudents = nothing
	elseif pintStudent_ID = "withdrawn" then		
		' Only return Inactive Students
		set rsGetStudents = server.CreateObject("ADODB.RECORDSET")
		rsGetStudents.CursorLocation = 3
		
		sql = "SELECT tblSTUDENT.intSTUDENT_ID " & _
				"FROM tblSTUDENT INNER JOIN " & _
				" tblStudent_States ON tblSTUDENT.intSTUDENT_ID = tblStudent_States.intStudent_id " & _
				"WHERE (tblStudent_States.intReEnroll_State = 86) " & _
				" AND (tblStudent_States.intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _
				"ORDER BY tblSTUDENT.szLAST_NAME, tblSTUDENT.szFIRST_NAME"
				
		rsGetStudents.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
		
		reDim arStudents((rsGetStudents.RecordCount-1))
		intCount = 0
		do while not rsGetStudents.EOF
			arStudents(intCount) = rsGetStudents("intSTUDENT_ID")
			rsGetStudents.MoveNext
			intCount = intCount + 1
		loop
		rsGetStudents.Close
		set rsGetStudents = nothing
	else
		' Only Return Selected Students
		arStudents = split(pintStudent_ID,",")
	end if		
	
	set rsStudent = server.CreateObject("ADODB.RECORDSET")
	rsStudent.CursorLocation = 3
	
	for i = 0 to ubound(arStudents)
		if isNumeric(arStudents(i)) then		
			sql = "SELECT 'The ' + f.szFamily_Name + ' Family' AS Family_title, f.szAddress, f.szCity + ', ' + f.szState + ' ' + f.szZip_Code AS addr2,  " & _
					" s.szFIRST_NAME + ' ' + s.szLAST_NAME AS Student_Name, f.szHome_Phone, fd.curFund_Amount, fd.szGRADE " & _
					"FROM tblSTUDENT s INNER JOIN " & _
					" tblFAMILY f ON s.intFamily_ID = f.intFamily_ID INNER JOIN " & _
					" tblFunding FD ON UPPER(s.szGRADE) = UPPER(FD.szGrade) INNER JOIN " & _
					" tblStudent_States ss ON s.intSTUDENT_ID = ss.intStudent_id " & _
					"WHERE (s.intSTUDENT_ID = " & arStudents(i) & ") " & _
					" AND (FD.intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _
					" AND (ss.intSchool_Year = " & session.Contents("intSchool_Year") & ") " 
					 
			rsStudent.Open sql, Application("cnnFPCS")'oFunc.FPCScnn			
			
			if rsStudent.RecordCount < 1  then
				response.Write "ERROR: Student #" & arStudents(i) & _
							   " is not assigned to a family.<p>" 
			else
				intPercent = oFunc.StudentPercentage(arStudents(i))
				strStudent_Name = rsStudent("Student_Name")
%>
<link rel="stylesheet" href="<% =  Application.Value("strWebRoot") %>CSS/printStyle.css">
<script language=javascript>
	function jfViewClass(classID,studentID) {
		var classWin;
		var strURL = "<%=Application.Value("strWebRoot")%>forms/";
		strURL += "printableForms/allPrintable.asp?noprint=true";
		strURL +="&strAction=C&intClass_ID=" +classID + "&intStudent_ID=" + studentID;
		classWin = window.open(strURL,"classWin","width=640,height=500,scrollbars=yes,resizable=yes");
		classWin.moveTo(0,0);
		classWin.focus();
	}
	
	function jfViewCosts(studentID,ilpID,classID){
		var strURL = "<%=Application.Value("strWebRoot")%>forms/Requisitions/req1.asp?intClass_ID="+classID;
		strURL += "&intStudent_ID=" + studentID + "&intILP_ID=" + ilpID;
		costsWin = window.open(strURL,"costsWin","width=710,height=500,scrollbars=yes,resize=yes,resizable=yes");
		costsWin.moveTo(0,0);
		costsWin.focus();
	}
</script>
<table width=100% ID="Table2" cellpadding=2 cellspacing=2>
	<tr>
		<td align=left>
			<img src="<% = Application("strImageRoot")%>fpcsLogo.gif">
		</td>
		<td align=right class=svplain10 width=100%>
			3339 Fairbanks St.<br>
			Anchorage, AK 99503<br>
			Ph: 907-742-3700<br>
			Fax: 907-742-3710
		</td>
	</tr>
	<tr class=yellowHeader>	
		<Td colspan=2>
			<table align=right ID="Table8"><tr><td align=right><font face=arial size=2 color=white><% = date()%></font></td></tr></table>
			&nbsp;<b>Student Statement</b>											
		</td>					
	</tr>
	<tr>
		<td align=left>
			<table>
				<tr>
					<td rowspan=3>
						&nbsp;&nbsp;
					</td>
					<td class=svplain10 >
						<b><% = rsStudent("Family_title") %></b>
					</td>
				</tr>
				<tr>
					<td class=svplain10 >
						<% = rsStudent("szAddress") %>
					</td>
				</tr>
				<tr>
					<td class=svplain10 >
						<% = rsStudent("addr2") %>
					</td>
				</tr>
			</table>
		</td>
		<td align=right class=svplain10 width=100% valign=top>
			<table ID="Table3">
				<tr>
					<td class=svplain10>
						Statement For: <b><% = rsStudent("Student_Name") %></b>
					</td>
					<td rowspan=2>
						&nbsp;&nbsp;
					</td>
				</tr>
				<tr>
					<td class=svplain10>
						<% = rsStudent("szHome_Phone") %>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<table width=100% border=1 cellpadding=0 cellspacing=0 bodercolorlight=cococo bordercolordark=cococo>
	<tr>
		<td class=gray colspan=2>
			&nbsp;<b>FUNDING </b>
		</td>
	</tr>
	<tr>
		<td class=svplain10>
			&nbsp;Available Funding for Grade <% = rsStudent("szGRADE") %>: $<% = formatNumber(rsStudent("curFund_Amount"),2) %>  @ <% = intPercent %>%
			Enrollment
		</td>
		<td align=right class=svplain10>
			
			<nobr>$<% dblTrack = formatNumber(cdbl(rsStudent("curFund_Amount") * (.01 * intPercent)),2)
			   response.Write dblTrack
			%></nobr>
		</td>
	</tr>
	<tr>
		<td class=gray colspan=2>
			&nbsp;<b>Transfers </b> <font size=1>(click a transfer below to go to the budget transefer page)</font>
		</td>
	</tr>
	<% = vbfTransfers(arStudents(i)) %>
	<tr>
		<td class=gray colspan=2>
			&nbsp;<b>Charges for ASD Teachers</b> <font size=1>(click a class below to view contract)</font>
		</td>
	</tr>
	<% = vbfASDTeacherCosts(arStudents(i)) %>
	<tr>
		<td class=gray colspan=2>
			&nbsp;<b>Charges for Goods </b> <font size=1>(click a row below to view goods/services)</font>
		</td>
	</tr>
	<% = vbfGoods(arStudents(i)) %>
	<tr>
		<td class=gray colspan=2>
			&nbsp;<b>Charges for Vendor Services </b> <font size=1>(click a row below to view goods/services)</font>
		</td>
	</tr>
	<% = vbfServices(arStudents(i)) %>
	<tr>
		<td class=gray colspan=2>
			&nbsp;<b>Charges for Reimbursements </b> <font size=1>(click a row below to view goods/services)</font>
		</td>
	</tr>
	<% = vbfReimburse(arStudents(i)) %>
	<tr>
		<td class=gray align=right>
			&nbsp;<b>Remaining Balance</b>
		</td>
		<td class=gray align=right>
			&nbsp;<nobr><b>$<% = formatNumber(dblTrack,2) %></b></nobr>
		</td>
	</tr>
</table>
<p></p>
<%
			end if ' ends recordcount
			rsStudent.close
		end if ' ends isNumeric
		dblTrack = 0
	next
	set rsStudent = nothing
end sub	
	
function vbfASDTeacherCosts(student_ID)
	' Calculates all Teacher costs in relation to instruction time
	dim strHTML
	dim dblCharge
	dim dblSubTotal
	
	set rsASDInfo = server.CreateObject("ADODB.RECORDSET")
	rsASDInfo.CursorLocation = 3
	
	' SQL to return all ASD instructed classes and info 
	sql = "SELECT intInstructor_ID, Teachers_Name, szClass_Name, " & _
		  " dtClass_Start, dtClass_End, hours_to_charge, intClass_ID " & _
		  "FROM v_Hours_Instructor_To_Charge " & _
		  "WHERE (sintSchool_Year = " & session.Contents("intSchool_Year") & ") " & _
		  " AND (intStudent_ID = " & student_ID & " )" & _
		  "ORDER BY Teachers_Name "
	rsASDInfo.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
		
	if rsASDInfo.RecordCount > 0 then
		strHTML = strHTML & "<tr><td><table border=1 bordercolor=e6e6e6 cellpadding=0 cellspacing=4><td class=svplain10><b>&nbsp;" & _
				  "Teacher</B>&nbsp;</td>" & _
				  "<td class=svplain10>" & _
				  "<b>&nbsp;Class Name</B>&nbsp</td>" & _
				  "<td class=svplain10 align=center><b>&nbsp;Class Dates</B>&nbsp</td><td class=svplain10><b>&nbsp;Hrs</B>&nbsp</td>" & _
				  "<td class=svplain10><b>&nbsp;Rate</B>&nbsp</td><td class=svplain10><b>&nbsp;Total</B>&nbsp</td></tr>"
		dblSubTotal = 0 
		do while not rsASDInfo.EOF
			' We use the dictionary obj to insure that we calculate teacher
			' hourly costs only once			
			if not objInstruct_Dict.Exists(rsASDInfo("intInstructor_ID") & "") then
				arTemp = oFunc.InstructorCosts(rsASDInfo("intInstructor_ID"))
				objInstruct_Dict.Add rsASDInfo("intInstructor_ID") & "", arTemp(9)
				erase arTemp
			end if 
			
			dblCharge = formatNumber((CDBL(rsASDInfo("hours_to_charge")) * cdbl(objInstruct_Dict.Item(rsASDInfo("intInstructor_ID")&""))))
			' Subtract spent funds for tracking total
			dblTrack = dblTrack - dblCharge	
			dblSubTotal = dblSubTotal + dblCharge
			
			strHTML = strHTML & "<tr valign=top  style='cursor:hand;' onClick=""jfViewClass('" & rsASDInfo("intClass_ID") & "','" & student_ID & "');"" title='Click to view contract.'><td class=svplain10>" & _
					  rsASDInfo("Teachers_Name") & "</td>" & _
					  "<td class=svplain10>" & rsASDInfo("szClass_Name") & "</td>" & _
					  "<td class=svplain10 align=center >" & rsASDInfo("dtClass_Start") & " - " & _
					  rsASDInfo("dtClass_End") & "</td>" & _
					  "<td class=svplain10 align=right>" & formatNumber(rsASDInfo("hours_to_charge"),2) & "</td>" & _
					  "<td class=svplain10 align=right>$" & formatNumber(objInstruct_Dict.Item(rsASDInfo("intInstructor_ID")& ""),2) & _
					  "</td><td class=svplain10 align=right>$" & dblCharge & _
					  "</td></tr>"
			rsASDInfo.MoveNext		  
		loop
		strHTML = strHTML & "</table></td><td class=svplain10 align=right valign=bottom><nobr>-$" & formatNumber(dblSubTotal,2) & _
				  "</nobr></td></tr>"
	else 
		strHTML = "<tr><td class=svplain10>&nbsp; NONE</td><td class=svplain10 align=right>$0.00</td></tr>"
	end if
	rsASDInfo.Close
	set rsASDInfo = nothing
	'Return HTML
	vbfASDTeacherCosts = strHTML
end function

function vbfGoods(student_ID)	
	' Calculates all charges for goods and print html
	dim sql
	dim strHTML
	dim dblCharge
	dim dblSubTotalMod
	
	set rsGoods = server.CreateObject("ADODB.RECORDSET")
	rsGoods.CursorLocation = 3
	
	sql  = "SELECT szClass_Name, item_name, intQty, curUnit_Price, curShipping, status, " & _
			"bolApproved,bolSponsor_Approved, szDeny_Reason, intClass_ID, intILP_ID " & _
			"FROM v_Ordered_Goods " & _
			"WHERE (intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _
		    " AND (intStudent_ID = " & student_ID & " )" & _
			" AND (bolReimburse = 0 OR bolReimburse IS NULL) " & _
			"ORDER BY szClass_Name"
			
	rsGoods.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
		
	if rsGoods.RecordCount > 0 then		
		strHTML = strHTML & "<tr><td><table border=1 bordercolor=e6e6e6 cellpadding=0 cellspacing=4><tr><td class=svplain10><b>&nbsp;" & _
				  "Status</B>&nbsp;</td><td class=svplain10><b>&nbsp;Class Name</B>&nbsp</td>" & _
				  "<td class=svplain10><b>&nbsp;Item</B>&nbsp</td><td class=svplain10><b>&nbsp;QTY</B>&nbsp</td>" & _
				  "<td class=svplain10><b>&nbsp;Cost</B>&nbsp</td><td class=svplain10><b>&nbsp;S/H</B>&nbsp</td><td class=svplain10><b>&nbsp;Total</B>&nbsp</td></tr>"
		dblSubTotal = 0  
		do while not rsGoods.EOF
			dblCharge = formatNumber(((cdbl(rsGoods("intQty")) * cdbl(rsGoods("curUnit_Price")))+cdbl(rsGoods("curShipping"))),2)			
			if (not rsGoods("bolApproved")) _
					OR ((not rsGoods("bolApproved") or rsGoods("bolApproved") & "" = "") _
					AND not rsGoods("bolSponsor_Approved")) then 
				strCSVClass = "svStrike10"		
			else
				strCSVClass = "svplain10"
				dblSubTotal = dblSubTotal + dblCharge
				' Subtract spent funds from tracking total
				dblTrack = dblTrack - dblCharge	
			end if						
			
			strHTML = strHTML & "<tr style='cursor:hand;' onClick=""jfViewCosts('" & student_ID & "','" & rsGoods("intILP_ID") & "','" & rsGoods("intClass_ID") & "');"" title=""Click to view goods. " & rsGoods("szDeny_Reason") & """ class='" & strCSVClass & "'>" & _
					"<td valign=top ><nobr>&nbsp;" & rsGoods("status") & "&nbsp;</nobr></td>" & _
					"<td valign=top>" & rsGoods("szClass_Name") & _
					"</td><td valign=top>" & rsGoods("item_name") & _
					"</td><td align=center  valign=top>" & rsGoods("intQty") & _
					"</td><td valign=top><nobr>$" & formatNumber(rsGoods("curUnit_Price"),2) & _
					"</nobr></td><td valign=top>$" & formatNumber(rsGoods("curShipping"),2) & _
					"</td><td align=right valign=top><nobr>$" & dblCharge & _
					"</nobr></td></tr>"
			rsGoods.MoveNext
		loop
		if instr(1,dblSubTotal,"-") > 0 then
			dblSubTotalMod = "+$" & formatNumber(replace(dblSubTotal,"-",""),2)
		else
			dblSubTotalMod = "-$" & formatNumber(dblSubTotal,2)
		end if
		strHTML = strHTML & "</table></td><td class=svplain10 align=right valign=bottom><nobr>" & dblSubTotalMod & _
				  "</nobr></td></tr>"
	else
		strHTML = "<tr><td class=svplain10>&nbsp; NONE</td><td class=svplain10 align=right>$0.00</td></tr>"
	end if
	rsGoods.Close
	set rsGoods = nothing
	vbfGoods = strHTML
end function

function vbfServices(student_ID)	
	' Calculates all charges for services and print html
	dim sql
	dim strHTML
	dim dblCharge
	dim dblSubTotal
	dim strCSVClass
	dim dblSubTotalMod
	
	set rsServices = server.CreateObject("ADODB.RECORDSET")
	rsServices.CursorLocation = 3
	
	sql  = "SELECT szVendor_Name, szClass_Name, dtClass_Start, dtClass_End, " & _
			"intQty, UnitType, curUnit_Price,Description,status, " & _
			"bolApproved,bolSponsor_Approved, szDeny_Reason, intClass_ID, intILP_ID " & _
			"FROM v_Ordered_Services " & _
			"WHERE (bolReimburse = 0 OR  bolReimburse IS NULL)" & _
			" AND (intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _
		    " AND (intStudent_ID = " & student_ID & " )" & _
			"ORDER BY szVendor_Name"
	rsServices.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
	
	if rsServices.RecordCount > 0 then
		strHTML = strHTML & "<tr><td><table border=1 bordercolor=e6e6e6 cellpadding=0 cellspacing=4><TR>" & _
				  "<td class=svplain10>&nbsp;<b>Status</B>&nbsp;</td>" & _
				  "<td class=svplain10>&nbsp;<b>Vendor</B>&nbsp;</td>" & _
				  "<td class=svplain10 align=center><b>Class Name</B></td>" & _
				  "<td class=svplain10 align=center><b>Class Dates</B></td>" & _
				  "<td class=svplain10><b>&nbsp;Desc</B>&nbsp</td><td class=svplain10 align=center><b>&nbsp;QTY</B>&nbsp</td>" & _
				  "<td class=svplain10><b>&nbsp;Rate</B>&nbsp</td><td class=svplain10><b>&nbsp;Total</B>&nbsp</td></tr>"
		dblSubTotal = 0
		do while not rsServices.EOF
			dblCharge = formatNumber((cdbl(rsServices("intQty")) * cdbl(rsServices("curUnit_Price"))),2)
			if (not rsServices("bolApproved")) _
					OR ((not rsServices("bolApproved") or rsServices("bolApproved") & "" = "") _
					AND not rsServices("bolSponsor_Approved")) then 
				strCSVClass = "svStrike10"		
			else
				strCSVClass = "svplain10"
				dblSubTotal = dblSubTotal + dblCharge
				' Subtract spent funds from tracking total
				dblTrack = dblTrack - dblCharge
			end if	
				
			strHTML = strHTML & "<tr style='cursor:hand;' onClick=""jfViewCosts('" & student_ID & "','" & rsServices("intILP_ID") & "','" & rsServices("intClass_ID") & "');"" class=""" & strCSVClass & """ title=""Click to view services. " & rsServices("szDeny_Reason") & """>" & _
					"<td valign=top><nobr>&nbsp;" & rsServices("status") & "&nbsp;</nobr></td>" & _
					"<td valign=top>" & rsServices("szVendor_Name") & "</td>" & _
					"<td valign=top>" & rsServices("szClass_Name") & "</td>" & _
					"<td align=center>" & rsServices("dtClass_Start") & _
					" - " & rsServices("dtClass_End") & "</td>" & _
					"<td>" & rsServices("Description") & "</td>" & _
					"<td><nobr>" & rsServices("intQty") & " " & replace(rsServices("UnitType")&""," ","") & "</nobr></td>" & _
					"<td><nobr>$" & formatNumber(rsServices("curUnit_Price"),2) & _
					"</nobr></td><td align=right><nobr>$" & dblCharge & _
					"</nobr></td></tr>"
			rsServices.MoveNext
		loop
		
		if instr(1,dblSubTotal,"-") > 0 then
			dblSubTotalMod = "+$" & formatNumber(replace(dblSubTotal,"-",""),2)
		else
			dblSubTotalMod = "-$" & formatNumber(dblSubTotal,2)
		end if
		
		strHTML = strHTML & "</table></td><td class=svplain10 align=right valign=bottom><nobr>" & dblSubTotalMod & _
				  "</nobr></td></tr>"
	else
		strHTML = "<tr><td class=svplain10>&nbsp; NONE</td><td class=svplain10 align=right>$0.00</td></tr>"
	end if
	rsServices.Close
	set rsServices = nothing
	vbfServices = strHTML
end function

function vbfReimburse(student_ID)	
	' Calculates all charges for reimbursements and print html
	dim sql
	dim strHTML
	dim dblCharge
	dim dblSubTotal
	dim strCSVClass
	dim dblSubTotalMod
	
	set rsReimburse = server.CreateObject("ADODB.RECORDSET")
	rsReimburse.CursorLocation = 3
	
	sql  = "SELECT szClass_Name,Check_Date, Check_Number, Payee, " & _
		   "description, ((intQty * curUnit_Price) + curShipping) AS cost, " & _
		   "status, bolApproved,bolSponsor_Approved, szDeny_Reason, intClass_ID, intILP_ID " & _
		   "FROM v_Reimbursements " & _
		   "WHERE (intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _
		   " AND (intStudent_ID = " & student_ID & " ) " & _
		   "ORDER BY Payee"
	rsReimburse.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
	
	if rsReimburse.RecordCount > 0 then
		strHTML = strHTML & "<tr><td><table border=1 bordercolor=e6e6e6 cellpadding=0 cellspacing=4>" & _
				  "<tr class=svplain10> " & _
				  "<td><b>&nbsp;Status</B>&nbsp;</td>" & _
				  "<td><b>&nbsp;Class Name</B>&nbsp;</td><td><b>&nbsp;Description</B>&nbsp</td>" & _
				  "<td align=center><b>&nbsp;Check Date</B>&nbsp</td>" & _
				  "<td><b>&nbsp;Check#</B>&nbsp</td><td align=center><b>&nbsp;Payee</B>&nbsp</td>" & _
				  "<td><b>&nbsp;Amount</B>&nbsp</td></tr>"
		dblSubTotal = 0
		
		do while not rsReimburse.EOF
			dblCharge = formatNumber(cdbl(rsReimburse("cost")),2)
			if (not rsReimburse("bolApproved")) _
					OR ((not rsReimburse("bolApproved") or rsReimburse("bolApproved") & "" = "") _
					AND not rsReimburse("bolSponsor_Approved")) then 
				strCSVClass = "svStrike10"		
			else
				strCSVClass = "svplain10"
				dblSubTotal = dblSubTotal + dblCharge
				' Subtract spent funds from tracking total
				dblTrack = dblTrack - dblCharge
			end if	
			strHTML = strHTML & "<tr  style='cursor:hand;' onClick=""jfViewCosts('" & student_ID & "','" & rsReimburse("intILP_ID") & "','" & rsReimburse("intClass_ID") & "');"" class=""" & strCSVClass & """ title=""Click to view reimbursements. " & rsReimburse("szDeny_Reason") & """>" & _
					"<td valign=middle>" & rsReimburse("status") & "</td>" & _
					"<td valign=middle>" & rsReimburse("szClass_Name") & "</td>" & _
					"<td>" & rsReimburse("description") & "</td>" & _
					"<td align=center>" & rsReimburse("Check_Date") & "</td>" & _
					"<td align=center>" & rsReimburse("Check_Number") & "</td>" & _
					"<td>" & rsReimburse("Payee") & "</td>" & _
					"</td><td align=right><nobr>$" & dblCharge & _
					"</nobr></td></tr>"
			rsReimburse.MoveNext
		loop
		
		if instr(1,dblSubTotal,"-") > 0 then
			dblSubTotalMod = "+$" & formatNumber(replace(dblSubTotal,"-",""),2)
		else
			dblSubTotalMod = "-$" & formatNumber(dblSubTotal,2)
		end if
		
		strHTML = strHTML & "</table></td><td class=svplain10 align=right valign=bottom><nobr>" & dblSubTotalMod & _
				  "</nobr></td></tr>"
	else
		strHTML = "<tr><td class=svplain10>&nbsp; NONE</td><td class=svplain10 align=right>$0.00</td></tr>"
	end if
	rsReimburse.Close
	set rsReimburse = nothing
	vbfReimburse = strHTML
end function

function vbfTransfers(student_ID)	
	' Calculates all charges for goods and print html
	dim sql
	dim strHTML
	dim strTableHeader
	dim dblCharge
	
	
	strTableHeader = "<tr><td><table onClick=""window.location.href='../budget/budgetTransfer.asp?intStudent_ID=" & student_ID & "';"" border=1 bordercolor=e6e6e6 cellpadding=0 cellspacing=4 title='Click to go to budget transfers.' style='cursor:hand;'>" & _
				  "<tr>" & _
				  "<td class=svplain10><b>&nbsp;Type</B>&nbsp;</td>" & _
				  "<td class=svplain10><b>&nbsp;Date of Trans</B>&nbsp</td>" & _
				  "<td class=svplain10><b>&nbsp;From Acct.</B>&nbsp</td>" & _
				  "<td class=svplain10><b>&nbsp;To Acct.</B>&nbsp</td>" & _
				  "<td class=svplain10><b>&nbsp;Amount</B>&nbsp</td>" & _
				  "</tr>"
				  
	set rsDeposit = server.CreateObject("ADODB.RECORDSET")
	rsDeposit.CursorLocation = 3
	
	sql  = "SELECT s.szFIRST_NAME, s.szLAST_NAME, bt.curAmount, bt.dtCREATE " & _
			"FROM tblBudget_Transfers bt INNER JOIN " & _
			" tblSTUDENT s ON bt.intFrom_Student_ID = s.intSTUDENT_ID " & _
			"WHERE (bt.intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _
			"AND (bt.intTo_Student_ID = " & student_ID & ")"
			
	rsDeposit.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
		
	if rsDeposit.RecordCount > 0 then				
		dblSubTotal = 0  
		do while not rsDeposit.EOF
			dblCharge = formatNumber(cdbl(rsDeposit("curAmount")),2)			
			dblSubTotal = dblSubTotal + dblCharge								
			
			strHTML = strHTML & "<tr class='svplain10'>" & _
					"<td valign=top><nobr>&nbsp;Deposit&nbsp;</nobr></td>" & _
					"<td align=center  valign=top>" & rsDeposit("dtCREATE") & "</td>" & _
					"<td valign=top>&nbsp;" & rsDeposit("szFirst_Name") & " " & rsDeposit("szLast_Name") & "&nbsp;</td>" & _
					"<td valign=top>&nbsp;" & strStudent_Name & "&nbsp;</td>" & _
					"<td align=center  valign=top>$" & formatNumber(rsDeposit("curAmount"),2) & "</td>" & _
					"</tr>"
			rsDeposit.MoveNext
		loop
	end if
	rsDeposit.Close
	
	sql  = "SELECT s.szFIRST_NAME, s.szLAST_NAME, bt.curAmount, bt.dtCREATE " & _
			"FROM tblBudget_Transfers bt INNER JOIN " & _
			" tblSTUDENT s ON bt.intTo_Student_ID= s.intSTUDENT_ID " & _
			"WHERE (bt.intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _
			"AND (bt.intFrom_Student_ID = " & student_ID & ")"
			
	rsDeposit.Open sql, Application("cnnFPCS")'oFunc.FPCScnn
		
	if rsDeposit.RecordCount > 0 then				
		do while not rsDeposit.EOF
			dblCharge = formatNumber(cdbl(rsDeposit("curAmount")),2)			
			dblSubTotal = dblSubTotal - dblCharge								
			
			strHTML = strHTML & "<tr class='svplain10'>" & _
					"<td valign=top><nobr>&nbsp;Withdrawl&nbsp;</nobr></td>" & _
					"<td align=center  valign=top>" & rsDeposit("dtCREATE") & "</td>" & _
					"<td valign=top>&nbsp;" & strStudent_Name & "&nbsp;</td>" & _
					"<td valign=top>&nbsp;" & rsDeposit("szFirst_Name") & " " & rsDeposit("szLast_Name") & "&nbsp;</td>" & _					
					"<td align=center  valign=top>-$" & formatNumber(rsDeposit("curAmount"),2) & "</td>" & _
					"</tr>"
			rsDeposit.MoveNext
		loop		
	end if
	
	if strHTML <> "" then
		strHTML = strTableHeader & strHTML & "</table></td><td class=svplain10 align=right valign=bottom><nobr>$" & formatNumber(dblSubTotal,2) & _
				  "</nobr></td></tr>"
		dblTrack = dblTrack + dblSubTotal
	else
		strHTML = "<tr><td class=svplain10>&nbsp; NONE</td><td class=svplain10 align=right>$0.00</td></tr>"
	end if
	
	set rsDeposit = nothing
	vbfTransfers = strHTML
end function

%>