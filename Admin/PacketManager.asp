<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		PacketManager.asp
'Purpose:	Admin page for accedemic review and approval of certified 
'			teacher contracts.  If contract approval is required 
'			students will not be able to enroll in a class until
'			approval has been given.
'Date:		2 MAY 2005
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim sql
dim oFunc 
dim rs

if ucase(session.Contents("strRole")) <> "ADMIN" then
	response.Write "<h1>Page Improperly Called</h1>"
	response.End
end if

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

set oHtml = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/htmlFuncs.wsc"))

if request("bolWin") = "" then
	Server.Execute(Application.Value("strWebRoot") & "includes/header.asp")
else
	Server.Execute(Application.Value("strWebRoot") & "includes/simpleHeader.asp")
end if

if request("updatelist") <> "" then call vbsUpdateStatus

%>
<script language=javascript>
	
	function jfPrintAll(class_ID,ilp_ID){
		var winContractApproval;
		var strURL = "<%=Application.Value("strWebRoot")%>forms/PrintableForms/allPrintable.asp?intClass_ID="+class_ID;
		strURL += "&noprint=ture&intILP_ID=" + ilp_ID ;
		winContractApproval = window.open(strURL,"winContractApproval","width=710,height=500,scrollbars=yes,resize=yes,resizable=yes");
		winContractApproval.moveTo(0,0);
		winContractApproval.focus();	
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
	
	function jfUpdateList(id) {
		// if an item as been changed log it only once.  We will use this list
		// to determine which Contract Status' should be modified
		if (document.main.updatelist.value.indexOf(","+id+",") == -1 ) {
			document.main.updatelist.value = document.main.updatelist.value + id + ",";
		}
	}	
	
	function jfPacket(id){
		var winILPPend;
		var strURL = "<%=Application.Value("strWebRoot")%>forms/Packet/packet.asp?simpleHeader=true&intStudent_ID="+id;
		winILPPend = window.open(strURL,"winILPPend","width=830,height=550,scrollbars=yes,resize=yes,resizable=yes");
		winILPPend.moveTo(0,0);
		winILPPend.focus();	
	}
	
	function jfPrintPacketList(pID,pObj){
		var sList = document.main.strPacketList;
		
		if (pObj.checked == true) {
			if (sList.value.indexOf(","+pID+",") == -1 ) {
				sList.value = sList.value + pID + ",";
			}
		}else {
			var re = new RegExp(pID + ",",'gi');
			sList.value = sList.value.replace(re,'');
		}
	}
	
	function jfPrintPacketAll(){
		var winPrintPacket;
		var sList = document.main.strPacketList;
		var strURL = "<%=Application.Value("strWebRoot")%>forms/PrintableForms/allPrintable.asp?strAction=AP";
		strURL += "&strPacketList=" + sList.value;
		winPrintPacket = window.open(strURL,"winPrintPacket","width=710,height=500,scrollbars=yes,resize=yes,resizable=yes");
		winPrintPacket.moveTo(0,0);
		winPrintPacket.focus();	
	}
</script>	
<form name="main" method="post" action="PacketManager.asp" ID="Form1">
<input type=hidden name="lastRow" ID="Hidden1">
<input type=hidden name="LineItemsChanged" value="," ID="Hidden8">
<input type=hidden name="lastRowColor" ID="Hidden3">
<input type="hidden" name="updatelist" value="" ID="Hidden2">
<input type="hidden" name="hdnReset" value="" ID="Hidden4">
<input type="hidden" name="Search" value="true" ID="Hidden5">
<input type="hidden" name="strPacketList" value="," ID="Hidden9">
<table style="width:100%;" ID="Table3">
	<tr>
		<td class="yellowHeader">
			&nbsp;<b>Packet Manager</b>
		</td>
	</tr>
	<tr>
		<td>
		<!--
			<table ID="Table1">						
				<tr>
					<td style="width:0%;">
						<table style="width:100%;" cellpadding="2" ID="Table2">
							<tr>
								<td class="TableHeader">
									Teachers
								</td>
								<td class="TableHeader">
									Status
								</td>
								<td rowspan="2" valign="middle">
									<input type="submit" value="Search/Save" class="NavSave" ID="Submit1" NAME="Search">
								</td>
							</tr>
							<tr>
								<td>
									<select name="intInstructor_ID" onchange="this.form.hdnReset.value='true';" ID="Select1">
										<option value="">All Teachers
									<%
										sql = "SELECT DISTINCT tblINSTRUCTOR.intInstructor_ID, tblINSTRUCTOR.szLAST_NAME + ', ' + tblINSTRUCTOR.szFIRST_NAME as Name " & _ 
												" FROM tblINSTRUCTOR INNER JOIN " & _ 
												" tblClasses ON tblINSTRUCTOR.intINSTRUCTOR_ID = tblClasses.intInstructor_ID " & _ 
												" WHERE (tblClasses.intSchool_Year = " & session.Contents("intSchooL_Year") & ") " & _ 
												" ORDER BY Name "
										Response.Write oFunc.MakeListSQL(sql,"intInstructor_ID","Name",request("intInstructor_ID"))
									%>
									</select>
								</td>
								<td>
									<select name="intContract_Status_ID"  ID="Select2"  onchange="this.form.hdnReset.value='true';">
										<option value="">All 
									<%
										sql = "SELECT intContract_Status_ID, szContract_Status_Name " & _ 
												"FROM tblContract_Status_Types " & _ 
												"WHERE (intYear_Active_Start >= " & session.Contents("intSchool_Year") & ") " & _
												" AND (intYear_Active_End <= " & session.Contents("intSchool_Year") & ") OR " & _ 
												" (intYear_Active_End IS NULL) order by intContract_Status_ID "
										Response.Write oFunc.MakeListSQL(sql,"intContract_Status_ID","szContract_Status_Name",request("intContract_Status_ID"))	
									%>
									</select>	
								</td>			
							</tr>
						</table>
					</td>
				</tr>				
			</table>-->
		</td>
	</tr>
</table>
<%


	sql = "SELECT     s.szLAST_NAME + ', ' + s.szFIRST_NAME AS Name, f.szHome_Phone,f.szDesc + ' ' + f.szFamily_Name FamilyName, f.szEMAIL, s.intSTUDENT_ID, ss.intReEnroll_State, ss.szGrade, " & _ 
			"                          (SELECT     COUNT(*) " & _ 
			"                            FROM          tblILP i " & _ 
			"                            WHERE      i.intStudent_ID = s.intStudent_ID AND (i.GuardianStatusID <> 1 OR " & _ 
			"                                                   i.GuardianStatusID IS NULL) AND i.sintSchool_Year = ss.intSchool_Year) AS GuardNotSign, " & _ 
			"                          (SELECT     COUNT(*) " & _ 
			"                            FROM          tblILP i " & _ 
			"                            WHERE      i.intStudent_ID = s.intStudent_ID AND (i.SponsorStatusID <> 1 OR " & _ 
			"                                                   i.SponsorStatusID IS NULL) AND i.sintSchool_Year = ss.intSchool_Year) AS SponsorNotSign, " & _ 
			"                          (SELECT     COUNT(*) " & _ 
			"                            FROM          tblILP i INNER JOIN " & _ 
			"                                                   tblClasses c3 ON i.intClass_ID = c3.intClass_ID " & _ 
			"                            WHERE      i.intStudent_ID = s.intStudent_ID AND (i.InstructorStatusID <> 1 OR " & _ 
			"                                                   i.InstructorStatusID IS NULL) AND i.sintSchool_Year = ss.intSchool_Year AND c3.intInstructor_Id IS NOT NULL) AS InstructorNotSign, " & _ 
			"                          (SELECT     COUNT(*) " & _ 
			"                            FROM          tblILP i INNER JOIN " & _ 
			"                                                   tblClasses c3 ON i.intClass_ID = c3.intClass_ID " & _ 
			"                            WHERE      i.intStudent_ID = s.intStudent_ID AND (i.AdminStatusID <> 1 OR " & _ 
			"                                                   i.AdminStatusID IS NULL) AND i.sintSchool_Year = ss.intSchool_Year AND c3.intInstructor_ID IS NULL) AS AdminNotSign,  " & _ 
			"                      DM_PERCENT_ENROLLED.TotalCoreHours, DM_PERCENT_ENROLLED.TotalElectiveHours,  " & _ 
			"                      DM_PERCENT_ENROLLED.TotalHours, DM_PERCENT_ENROLLED.CoreCredits,  " & _ 
			"                      DM_PERCENT_ENROLLED.ElectiveCredits, DM_PERCENT_ENROLLED.ActualEnrolledPercent,  " & _ 
			"                      DM_PERCENT_ENROLLED.GoalCoreCredits, DM_PERCENT_ENROLLED.GoalElectiveCredits,  " & _ 
			"                      DM_PERCENT_ENROLLED.GoalContractHours, DM_PERCENT_ENROLLED.GoalClassTime, e.intPercent_Enrolled_Fpcs, " & _ 
			"                          (SELECT     MAX(GuardianStatusDate) " & _ 
			"                            FROM          tblILP ti " & _ 
			"                            WHERE      ti.sintSchool_Year = ss.intSchool_Year AND ti.intStudent_ID = ss.intStudent_Id) AS maxGuardDate, e.AdminPacketSigned, e.PacketSignDate, e.intEnroll_Info_ID,  " & _ 
			"                          (SELECT     MAX(SponsorStatusDate) " & _ 
			"                            FROM          tblILP ti " & _ 
			"                            WHERE      ti.sintSchool_Year = ss.intSchool_Year AND ti.intStudent_ID = ss.intStudent_Id) AS maxSponsorDate, " & _
			"                          (SELECT     MAX(InstructorStatusDate) " & _ 
			"                            FROM          tblILP ti " & _ 
			"                            WHERE      ti.sintSchool_Year = ss.intSchool_Year AND ti.intStudent_ID = ss.intStudent_Id) AS maxInstructorDate, " & _
			"                          (SELECT     MAX(AdminStatusDate) " & _ 
			"                            FROM          tblILP ti " & _ 
			"                            WHERE      ti.sintSchool_Year = ss.intSchool_Year AND ti.intStudent_ID = ss.intStudent_Id) AS maxAdminDate, " & _
			"					   e.dtASD_Signed, e.dtProgress_Signed, " & _ 
			"                      tblINSTRUCTOR.intINSTRUCTOR_ID,  " & _ 
			"                      tblINSTRUCTOR.szFIRST_NAME + ' ' + tblINSTRUCTOR.szLAST_NAME AS TeacherName,  " & _ 
			"                      tblINSTRUCTOR.szHOME_PHONE AS TeacherPhone, tblINSTRUCTOR.szEmail AS TeacherEmail,se.TotalTeacherHours, e.bolASD_Testing, e.bolProgress_Agreement, e.intPhilosophy_ID, " & _ 
			"		(SELECT top 1 i2.intILP_ID " & _ 
			"			FROM         tblILP i2 INNER JOIN " & _ 
			"                      tblClasses c ON i2.intClass_ID = c.intClass_ID " & _ 
			"			WHERE     (c.intPOS_Subject_ID = 22) AND (i2.sintSchool_Year = " & session.Contents("intSchool_Year") & ") AND (i2.intStudent_ID = s.intSTUDENT_ID) AND  " & _ 
			"                      (c.decOriginal_Student_Hrs + c.decOriginal_Planning_Hrs > 0)) as HasSponsorCourse, e.dtPacket_Printed  " & _
			" FROM         tblENROLL_INFO e INNER JOIN " & _ 
			"                      tblSTUDENT s ON e.intSTUDENT_ID = s.intSTUDENT_ID INNER JOIN " & _ 
			"                      tblFAMILY f ON s.intFamily_ID = f.intFamily_ID INNER JOIN " & _ 
			"                      tblStudent_States ss ON ss.intStudent_id = s.intSTUDENT_ID LEFT OUTER JOIN " & _ 
			"                      tblINSTRUCTOR ON e.intSponsor_Teacher_ID = tblINSTRUCTOR.intINSTRUCTOR_ID AND  " & _ 
			"                      e.intSponsor_Teacher_ID = tblINSTRUCTOR.intINSTRUCTOR_ID LEFT OUTER JOIN " & _ 
			"                      DM_PERCENT_ENROLLED ON ss.intSchool_Year = DM_PERCENT_ENROLLED.SchoolYear AND  " & _ 
			"                      s.intSTUDENT_ID = DM_PERCENT_ENROLLED.StudentID LEFT OUTER JOIN " & _ 
			"                      DM_STUDENT_EXPENSES se ON ss.intSchool_Year = se.SchoolYear AND  " & _ 
			"                      s.intSTUDENT_ID = se.StudentID " & _
			"WHERE     (e.sintSCHOOL_YEAR = " & session.Contents("intSchool_Year") & ") AND (ss.intSchool_Year = " & session.Contents("intSchool_Year") & ") AND ss.intReEnroll_State IN (" & Application.Contents("ActiveEnrollList") & ") " 

	

	if request("orderby") <> "" then
		sql = sql & " ORDER BY " & request("orderby")
	else
		sql = sql & " ORDER BY s.szLast_Name, s.szFirst_Name" 
	end if
	
	set rs = server.CreateObject("ADODB.RECORDSET")
	rs.CursorLocation = 3
	
	rs.Open sql, oFunc.FPCScnn
	
	if request("PageNumber") <> "" and request("hdnReset") = "" then
		intPageNum = cint(request("PageNumber"))	
	else
		intPageNum = 1
	end if
	
	with rs
		if .RecordCount > 0 then
			.PageSize = 1000
			.AbsolutePage = intPageNum
			intViewingTo = .AbsolutePosition + .PageSize -1 
			if intViewingTo > .recordcount then intViewingTo = .RecordCount
%>
<br>
<input type="hidden" name="PageNumber" value="<% = intPageNum%>" ID="Hidden7">
<table cellpadding="2" ID="Table4">
	<tr>
		<td colspan=10 class="svplain8" nowrap>
			
			Viewing <% = .AbsolutePosition %> - <% = intViewingTo %>  of <% = .RecordCount %> Matches &nbsp;
			
			<table ID="Table5" cellpadding="2"><tr><td>
			<%
				if cint(.RecordCount) > cint(.PageSize) then
					for i = 1 to .PageCount
						if intViewingTo/.PageSize = i or (.RecordCount = intViewingTo and i = .PageCount) then 
							strClass = "NavSave"
						else
							strClass = "btSmallWhite"
						end if
					%>
					<input type="button" class="<% = strClass %>" value="<%=i%>" onClick="this.form.PageNumber.value='<%=i%>';this.form.submit();" ID="Button2" NAME="Button2">
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
			
				strStudentData = "<table><tr class='svplain8'><td><b>Family Name:</b></td>" & _
								 "<td nowrap>" & rs("FamilyName") & "</td></tr>" & _
								 "<tr class='svplain8'><td nowrap><b>Family Phone:</b></td>" & _
								 "<td nowrap>" & rs("szHome_Phone") & "</td></tr>" & _
								 "<tr class='svplain8'><td nowrap><b>Family Email:</b></td>" & _
								 "<td nowrap><a href='mailto:" & rs("szEMAIL") & "'>" & rs("szEMAIL") & "</a></td></tr>" & _
								 "<tr class='svplain8'><td nowrap><b>Sponsor Teacher:</b></td>" & _
								 "<td nowrap>" & rs("TeacherName") & "</td></tr>" & _
								 "<tr class='svplain8'><td nowrap><b>Sponsor Phone:</b></td>" & _
								 "<td nowrap>" & rs("TeacherPhone") & "</td></tr>" & _
								 "<tr class='svplain8'><td nowrap><b>Sponsor Email:</b></td>" & _
								 "<td nowrap><a href='mailto:" & rs("TeacherEmail") & "'>" & rs("TeacherEmail") & "</a></td></tr>" & _
								 "<tr class='svplain8'><td colspan='2'>" & _
								 "<table border='1' style='width:100%;' cellspacing=0 class='svplain8'>" & _
								 "<tr><td></td><td align='center'><b>Goal</b><td align='center'><b>Actual</b></td></tr>" & _
								 "<tr><td>Enroll %</td>" & _
								 "<td align='center'>" & rs("intPercent_Enrolled_Fpcs") & "</td>" & _
								 "<td align='center'>" & rs("ActualEnrolledPercent") & "</td></tr>" & _								 
								 "<tr><td>Goal Units</td>" & _
								 "<td align='center'>" & rs("GoalCoreCredits") & " Core/" & rs("GoalCoreCredits")+ rs("GoalElectiveCredits") & " total</td>" & _
								 "<td align='center'>" & round(CheckNumber(rs("CoreCredits")),1) & " Core/" & round(CheckNumber(rs("CoreCredits")),1)+ round(CheckNumber(rs("ElectiveCredits")),1) & " total</td></tr>" & _								 													 
								 "<tr><td>Contract HRS</td>" & _
								 "<td align='center'>" & rs("GoalContractHours") & "</td>" & _
								 "<td align='center'>" & rs("TotalTeacherHours") & "</td></tr></table></td></tr></table>" 
								 
				if intCount mod 2 = 0 then
					strColor = "TableCell"
				else
					strColor = "gray"
				end if
				
				if intCount2 = 0 or intCount2 mod 25 = 0 then
					call PrintHeader
				end if
				
				' Has the packet met all conditions to be ready for admin review
				if rs("GuardNotSign") = 0 and rs("SponsorNotSign") = 0 and _
				   rs("InstructorNotSign") = 0  and rs("ActualEnrolledPercent") >= rs("intPercent_Enrolled_Fpcs") and _
				   rs("bolASD_Testing") & "" <> "" and rs("bolProgress_Agreement") & "" <> "" and rs("intPhilosophy_ID") & "" <> "" AND _
				   rs("HasSponsorCourse") & "" <> "" then	
					if rs("AdminNotSign") = 0 and rs("AdminPacketSigned") & "" <> "" then
						PacketStatus = "Finished"
						PacketCss = "green"
					else
						PacketStatus = "Ready"
						PacketCss = "yellow"
					end if
				else
					PacketStatus = "Not Ready"	
					PacketCss = "red"			
				end if
				
%>	
			<tr id="ROW<%=intCount%>" onClick="jfHighLight('<%=intCount%>');" class="<% = strColor %>">
				<td >
					<a href="javascript:" onclick="jfPacket('<% = rs("intStudent_ID") %>');"><% response.Write oHtml.ToolTip(rs("NAME"),strStudentData,true,"Student Information",true,"tooltip","","",false,false) %></a>
				</td>
				<td align="center" class="<% = PacketCss%>">
					<% = PacketStatus %>
				</td>
				<td >
					<% if PacketStatus <> "Not Ready" then 
							myDate = oFunc.GreatestDate(array(rs("maxGuardDate"),rs("maxSponsorDate"),rs("maxInstructorDate"),rs("dtASD_Signed"),rs("dtProgress_Signed"),rs("maxAdminDate")))
							if isDate(myDate) then
								myDate = formatDateTime(myDate,2)
							else
								myDate = ""
							end if
							response.Write myDate
					   end if
				
					%>
					
					
				</td>
				<%'JD 052611 display achieved ilp %>
				<td>
				    <% if rs("TotalHours") > rs("GoalClassTime") then %>
				       <span class="sverror"><%=rs("TotalHours") %></span>
				    <% else %>
				        <%=rs("TotalHours") %>
				    <%end if %>
				</td>	
				<td align="center" >
				<% if rs("bolASD_Testing") & "" <> "" then %>
				Yes
				<% else %>
				<span class="sverror">No</span>
				<% end if %>
				</td>
				<td align="center" >
				<% if rs("bolProgress_Agreement") then %>
				Yes
				<% else %>
				<span class="sverror">No</span>
				<% end if %>
				</td>
				<td align="center" >
				<% if rs("intPhilosophy_ID") & "" <> "" then %>
				Yes
				<% else %>
				<span class="sverror">No</span>
				<% end if %>
				</td>
				<td align="center" >
				<% if rs("HasSponsorCourse") & "" <> "" then %>
				Yes
				<% else %>
				<span class="sverror">No</span>
				<% end if %>
				</td>
				<td title="Number of Courses that have not been signed by the Guardian." align="center">
					<b><% = rs("GuardNotSign") %></b>
				</td>			
				<td title="Number of Courses that have not been signed by the Sponsor." align="center">
					<b><% = rs("SponsorNotSign") %></b>
				</td>	
				<td title="Number of Courses that have not been signed by the Instructor." align="center">
					<b><% = rs("InstructorNotSign") %></b>
				</td>
				<td title="Number of Courses that have not been signed by the Admin." align="center">
					<b><% = rs("AdminNotSign") %></b>
				</td>	
				<td <% if rs("AdminPacketSigned") then%>title="Sign on <% = rs("PacketSignDate") %>" <% end if %> align="center">
					<input type="checkbox" name="IsSigned<% = rs("intEnroll_Info_ID") %>" <% if rs("AdminPacketSigned") then response.Write " checked " %> onclick="jfUpdateList('<% = rs("intEnroll_Info_ID") %>');">
				</td>
				<td align="center" <% if isDate(rs("dtPacket_Printed")) then response.Write "class='green' title='Printed: " & formatDateTime(rs("dtPacket_Printed"),2) & "'" else response.Write "class='red'" %>>
					<input type="checkbox" name="PrintPacket" value="<%=rs("intStudent_ID")%>" onChange="jfPrintPacketList('s<%=rs("intStudent_ID")%>',this);">
				</td>
				<td bgcolor="white">
				</td>
			</tr>
<%				
				.MoveNext
				intCount = intCount + 1
				intCount2 = intCount2 + 1
			loop
	
%>
	<input type=hidden name="intCount" value="<%=intCount%>" ID="Hidden10">
	<input type=hidden name="intCount2" value="<%=intCount2%>" ID="Hidden12">
	<input type="hidden" name="orderby" value="<% = request("orderby") %>" ID="Hidden6">
</table> 
<%		
		else
			%>
			<span class="svplain8"><B>0 Matches Found.</B></span>
			<%
		end if
		.close		
	end with


%>
</form>
<%
response.Write oHtml.ToolTipDivs
call oFunc.CloseCN()
set oFunc = nothing
set oHtml = nothing
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")


function PrintHeader
%>
	<Tr>
		<td class="TableHeader" >
			<a href="#" style="color:white;" class="linkWht" onclick="document.forms[0].orderby.value=' s.szLast_Name, s.szFirst_Name';document.forms[0].submit();">Student</a>
		</td>
		<td class="TableHeader" align="center">
			<b>Packet<br>Status</b>
		</td>
		<td class="TableHeader" align="center">
			<b>Ready Date</b>
		</td>	
		<%'JD 052611 Display Achieved ILP %>
		<td class="TableHeader" align="center">
		    <b>Achieved <br /> ILP</b>
		</td>		
		<td class="TableHeader" title="Yes if guardian has signed the Testing Agreement." align="center">
			<b>Test<br>Sign</b>
		</td>
		<td class="TableHeader" title="Yes if guardian has signed the Progress Report Agreement." align="center">
			<b>Progress<br>Sign</b>
		</td>
		<td class="TableHeader" title="Yes if ILP Philosophy has been filled in." align="center">
			<b>ILP<br>Philosophy</b>
		</td>	
		<td class="TableHeader" title="Yes if students packet contains a Sponsor Course." align="center">
			<b>Sponsor<br>Course</b>
		</td>	
		<td class="TableHeader" title="Number of Courses that have not been signed by the Guardian." align="center">
			<b>GNS</b>
		</td>			
		<td class="TableHeader" title="Number of Courses that have not been signed by the Sponsor." align="center">
			<b>SNS</b>
		</td>	
		<td class="TableHeader" title="Number of Courses that have not been signed by the Instructor." align="center">
			<b>INS</b>
		</td>
		<td class="TableHeader" title="Number of Courses that have not been signed by the Admin." align="center">
			<b>ANS</b>
		</td>			
		<td class="TableHeader" align="center">
			<b>Packet<br>Signed?</b>
		</td>
		<td class="TableHeader" align="center">
			<b>Packet<br>Printed?</b>
		</td>
		<td>
			<input type="submit" value="save signatures" class="NavSave" style="width:100px;">
			<input type="button" value="print packets" onclick="jfPrintPacketAll();" class="NavSave" style="width:100px;">
			
		</td>
	</Tr>
<%
end function

sub vbsUpdateStatus
	dim update, updateAdd
	dim list, i
	list = split(request("updatelist"),",")
	for i = 0 to ubound(list)
		if list(i) <> "" then
			if request("IsSigned" & list(i)) & "" <> "" then
				IsSigned = 1
			else 
				IsSigned = 0 
			end if
			
			update = "Update tblEnroll_Info set AdminPacketSigned = " & IsSigned & _
					 ", PacketSignDate = CURRENT_TIMESTAMP, szUSER_Modify =' " & session.Contents("strUserId") & "' " & _
					 " WHERE intEnroll_Info_ID = " & list(i)
			oFunc.ExecuteCN(update)
		end if
	next
end sub

function CheckNumber(pNum)
	if not isNumeric(pNum) then
		pNum = 0
	else
		pNum = pNum
	end if
	
	CheckNumber =  pNum
end function
%>