<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		ContractManager.asp
'Purpose:	Allows Instructors the ability to Sign or reject contracts
'Date:		June 7 2005
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim sql
dim oFunc
dim rs

' Security Check. Must be an Admin
if ucase(session.Contents("strRole")) <> "TEACHER" then
	response.Write "<H1>PAGE ILLEGALLY CALLED</H1>"
	response.End
end if

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

'Print the header
Session.Value("strTitle") = "Contract Manager"
Session.Value("strLastUpdate") = "June 7, 2005"
Server.Execute(Application.Value("strWebRoot") & "includes/Header.asp")	

if request("strILPList") <> "" then	
	call vbsUpdateILPStatus(request("strILPList"))
end if

sql = "SELECT     tblClasses.intClass_ID, tblClasses.szClass_Name, trefPOS_Subjects.szSubject_Name, tblSTUDENT.intSTUDENT_ID, tblSTUDENT.szFIRST_NAME,  " & _ 
		"                      tblSTUDENT.szLAST_NAME, tblFAMILY.szFamily_Name, tblFAMILY.szDesc, tblFAMILY.szHome_Phone, tblFAMILY.szEMAIL,  " & _ 
		"                      tblILP.szAdmin_Comments, tblILP.szSponsor_Comments, tblILP.GuardianStatusId, tblILP.SponsorStatusId, tblILP.InstructorStatusId,  " & _ 
		"                      tblILP.AdminStatusId, tblILP.GuardianStatusDate, tblILP.SponsorStatusDate, tblILP.InstructorStatusDate, tblILP.AdminStatusDate,  " & _ 
		"                      tblILP.GuardianComments, tblILP.InstructorComments, tblILP.GuardianUser, tblILP.SponsorUser, tblILP.InstructorUser, tblILP.AdminUser,  " & _ 
		"                      tblClasses.intContract_Status_ID, tblClasses.dtApproved, tblClasses.szUser_Approved, tblClasses.szComments, tblILP.intILP_ID,tblClasses.szInstructor_Comments, " & _ 
		"CASE WHEN " & _ 
		"                          (SELECT     ei.intSponsor_Teacher_ID " & _ 
		"                            FROM          tblENROLL_INFO ei " & _ 
		"                            WHERE      (ei.sintSCHOOL_YEAR = tblClasses.intSchool_Year) AND (ei.intSTUDENT_ID = tblStudent.intStudent_ID))  " & _ 
		"                      = " & session.Contents("instruct_id") & " THEN 1 ELSE 0 END AS IsSponsor " & _ 
		"FROM         tblClasses INNER JOIN " & _ 
		"                      tblILP ON tblClasses.intClass_ID = tblILP.intClass_ID INNER JOIN " & _ 
		"                      tblSTUDENT ON tblILP.intStudent_ID = tblSTUDENT.intSTUDENT_ID INNER JOIN " & _ 
		"                      trefPOS_Subjects ON tblClasses.intPOS_Subject_ID = trefPOS_Subjects.intPOS_Subject_ID LEFT OUTER JOIN " & _ 
		"                      tblFAMILY ON tblSTUDENT.intFamily_ID = tblFAMILY.intFamily_ID " & _ 
		"WHERE     (tblClasses.intInstructor_ID = " & session.Contents("instruct_id") & ") AND (tblClasses.intSchool_Year = " & session.Contents("intSchool_Year") & ") " & _ 
		"ORDER BY tblClasses.szClass_Name, tblSTUDENT.szLAST_NAME, tblSTUDENT.szFIRST_NAME "
	'response.Write sql 	
set rs = server.CreateObject("ADODB.RECORDSET")
rs.CursorLocation = 3
rs.Open sql, Application("cnnFPCS")'oFunc.FpcsCnn
set oHtml = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/htmlFuncs.wsc")) 
%>
<script language="javascript">
	function jfILPStatus(pID){
		var sList = document.main.strILPList;
		
		if (sList.value.indexOf(","+pID+",") == -1 ) {
			sList.value = sList.value + pID + ",";
		}
	}		
	
	function jfPrintAll(class_ID,ilp_ID,studentId){
		var winContractApproval;
		var strURL = "<%=Application.Value("strWebRoot")%>forms/PrintableForms/allPrintable.asp?intClass_ID="+class_ID;
		strURL += "&noprint=true&intILP_ID=" + ilp_ID + "&intStudent_ID=" + studentId;
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
		obj.className = "yellowHeader";
	}
</script>
<form name=main method=post action="ContractManager.asp">
<input type=hidden name="strILPList" value=",">
<input type=hidden name="lastRow" ID="Hidden1">
<input type=hidden name="lastRowColor" ID="Hidden3">
<table style="width:100%;" ID="Table3">
	<tr>
		<td class="yellowHeader">
			&nbsp;<b>Instructors Contract Manager</b>
		</td>
	</tr>	
	<tr>
		<td style="width:100%;">
			<table style="width:100%;">
			<%
			if rs.RecordCount < 1 then
			%>
				<tr>
					<td class="svplain10" colspan="10">
						Currently there are no contracts that need to be signed.
					</td>
				</tr>
			<%
			else								
				count = 0
				do while not rs.EOF
					if count mod 25 = 0 then
						response.Write PrintHeader					
					end if
					
					if count mod 2 = 0 then
						strColor = "TableCell"
					else
						strColor = "gray"
					end if
			%>
			<tr  id="ROW<%=count%>" onClick="jfHighLight('<%=count%>');" class="<% = strColor %>">
				<td class="TableCell">
					<a href="#" onCLick="jfPrintAll('<% =rs("intClass_ID")%>','<% = rs("intILP_ID") %>','<% = rs("intStudent_ID") %>');"><% = ucase(rs("szClass_Name")) %></a>
				</td>
				<td class="TableCell">
					<% response.Write oHtml.ToolTip(rs("szLAST_NAME") & ", " & rs("szFIRST_NAME"), _
										"<table><tr><td class='svplain8' nowrap>" & _
										"<b>Guardians:</b> " & rs("szDesc") & "<BR>" & _
										"<b>Email:</b> <a href=""mailto:" & rs("szEMAIL") & """>" & rs("szEMAIL") & "</a><BR>" & _										
										"<b>Family Phone:</b> " & rs("szHome_Phone") & "</td></tr></table>", _
										false,"",true,"ToolTip","","",false,false)  %>
				</td>
				
				<td align="center"  class="TableCell">
				<%
				if rs("GuardianStatusId") & "" = "" then
				%>
					not signed
				<% else %>
					<span title="signed on: <% = rs("GuardianStatusDate")%>"><% = rs("GuardianUser")%></span>
				<% end if %>
				</td>
				
				<td class="TableCell" align="center">
				<%
				if rs("SponsorStatusId") & ""  = "1" then %>
				<span title="signed on: <% = rs("SponsorStatusDate")%>"><% = rs("SponsorUser")%></span>
				<% else
					response.Write InterpretStatus(rs("SponsorStatusId"))
				end if
				%>
				</td>				
				<td class="TableCell" align="center">
					<%
					if rs("IsSponsor") & "" = "1" then
					%>
					<input type="hidden" name="IsSponsor<% = rs("intILP_ID")%>" value="1">
					<%
					end if
					
					if rs("InstructorStatusId") & "" <> "1" then												
					%>
					<select ID="Select5" NAME="status<% = rs("intILP_ID")%>" style="font-size:7pt;"  onchange="jfILPStatus('<% = rs("intILP_ID")%>');">
						<option></option>
						<option <% if rs("InstructorStatusId") & ""  = "3" then response.Write " selected " %> value="3">Rejected</option>
						<option value="1">Sign</option>													
					</select> 
					<% else 
							if rs("InstructorStatusId") & ""  = "1" then %>
							<span title="signed on: <% = rs("InstructorStatusDate")%>"><% = rs("InstructorUser")%></span>
							<% else
								response.Write InterpretStatus(rs("InstructorStatusId"))
							end if
					  end if %>
				</td>			
				<td class="TableCell" align="center">
				<%
				if rs("intContract_Status_ID") & ""  = "5" then %>
				<span title="signed on: <% = rs("dtApproved")%>"><% = rs("szUser_Approved")%></span>
				<% else
					response.Write "not signed"
				end if
				%>
				</td>
			<%			
				strCommentTable = ""		
				if rs("szAdmin_Comments") & "" <> "" then
					strCommentTable = strCommentTable & "<tr>" & _
										"<td class='TableCell' style='width:140px;' align=left valign='top'><b>Admin Comments</b></td>" & _
										"<td class='TableCell'>" & rs("szAdmin_Comments") & "</td></tr>"
				end if
				
				if rs("szSponsor_Comments") & "" <> "" then
					strCommentTable = strCommentTable & "<tr>" & _
										"<td class='TableCell' style='width:140px;' align=left valign='top'><b>Sponsor Comments</b></td>" & _
										"<td class='TableCell'>" & rs("szSponsor_Comments") & "</td></tr>"
				end if
				
				if rs("InstructorComments") & "" <> "" then
					strCommentTable = strCommentTable & "<tr>" & _
										"<td class='TableCell' style='width:140px;' align=left valign='top'><b>Instructor Comments</b></td>" & _
										"<td class='TableCell'>" & rs("InstructorComments") & "</td></tr>"
				end if
				
				if rs("GuardianComments") & "" <> "" then
					strCommentTable = strCommentTable & "<tr>" & _
										"<td class='TableCell' style='width:140px;' align=left valign='top'><b>Guardian Comments</b></td>" & _
										"<td class='TableCell'>" & rs("GuardianComments") & "</td></tr>"
				end if
										
				strCommentTable = "<table cellpadding='2' style='width:100%;'>" & strCommentTable & "</table>"	
			%>
				<td class="svplain8">
					<textarea style='width:99%;' rows='1' wrap='virtual' name="szComments<% = rs("intILP_ID") %>" onfocus='this.rows=4;' onblur='this.rows=1;' onKeyDown='jfMaxSize(1999,this);' 
					onChange="jfILPStatus('<% = rs("intILP_ID") %>');" ID="Textarea1"><% = rs("InstructorComments") %></textarea>
					<% if strCommentTable <> "" then response.Write strCommentTable 
						'if rs("szInstructor_Comments") & "" <> "" then
					'		response.Write "<br><b>Teacher Comments:</b> " &  rs("szInstructor_Comments")
					'	end if
					
					%>	
				</td>
			</tr>
			<%							
					count = count + 1								
					rs.MoveNext
				loop							
			end if
			%>
			</table>
		</td>
	</tr>	
</table>
</form>

<%
response.Write oHtml.ToolTipDivs	
set oHtml = nothing
rs.Close
set rs = nothing
call oFunc.CloseCN
set oFunc = nothing
Server.Execute(Application.Value("strWebRoot") & "includes/footer.asp")	
function PrintHeader()
%>
	<tr>
		<td colspan="20">
			<input type=submit value="Save Status & Comments" class="NavSave" style="width:165px;" >		
		</td>
	</tr>
	<tr>
		<td class="TableHeader">
			&nbsp;<B>Class Name</B><br>
			&nbsp;(View contract/Ilp)
		</td>
		<td class="TableHeader">
			&nbsp;<B>Student Name</B><br>
			&nbsp;(Mouse Over)
		</td>
		<td class="TableHeader" align="center">
			<B>Guard<br>Sign</B>
		</td>
		<td class="TableHeader" align="center">
			<B>Sponsor<br>Sign</B>
		</td>
		<td class="TableHeader" align="center">
			<B>Instruct<br>Sign</B>
		</td>
		<td class="TableHeader" align="center">
			<B>Admin<br>Sign</B>
		</td>
		<td class="TableHeader" align="center">
			<B>Comments</B>
		</td>
	</tr>
<%
End function

function InterpretStatus(pStatusId)
	' simply takes the statusId and gives us the corresponding label so 
	' we don't have to make 4 more sub queries to get the label for each role
	' This of course stinks if the labels need to be changed. 
	select case pStatusId
		case "1"
			InterpretStatus = "Signed"
		case "2"
			InterpretStatus = "Must Amend"
		case "3"
			InterpretStatus = "Rejected"
		case else
			InterpretStatus = "not signed"
	end select
end function

sub vbsUpdateILPStatus(pstrILPList)
	arList = split(pstrILPList,",")	
	if isArray(arList) then
		for i = 0 to ubound(arList)		
			if arList(i) <> "" then
				call vbsIlpStatus(arList(i))
			end if
		next
	end if
end sub

sub vbsIlpStatus(pIlpId)
	' update ILP Status and comments based on user Role
	dim update, myStatus
	
	update = ""
	
	if request("status" & pIlpId) & "" = "" then
		myStatus = " NULL "
	else
		myStatus = request("status" & pIlpId)
	end if
	
	update = "update tblILP set "
	if request("IsSponsor" & pIlpId) & "" = "1" then
		' User is both the Instructor and the Sponsor Teacher
		update = update & " SponsorStatusId = " & myStatus & ", " & _
						  " SponsorStatusDate = CURRENT_TIMESTAMP, " & _
						  " SponsorUser = '" & session.Contents("strUserId") & "' " 
		
		update = update & ", InstructorStatusId = " & myStatus  & ", " & _
						  " InstructorStatusDate = CURRENT_TIMESTAMP, " & _
						  " InstructorUser = '" & session.Contents("strUserId") & "', " & _
						  " InstructorComments = '" & oFunc.EscapeTick(request("szComments" & pIlpId)) & "' " 		
						  				  
		if request("status" & pIlpId) & "" = "2" then
			update = update & " , GuardianStatusId = null " 
		end if
	elseif ucase(session.Contents("strRole")) = "TEACHER" then
		' User is only the Instructor
		update = update & " InstructorStatusId = " & myStatus  & ", " & _
						  " InstructorStatusDate = CURRENT_TIMESTAMP, " & _
						  " InstructorUser = '" & session.Contents("strUserId") & "', " & _
						  " InstructorComments = '" & oFunc.EscapeTick(request("szComments" & pIlpId)) & "' " 		
	else						  						 
		exit sub
	end if 
	
	update = update & " where intILP_ID = " &  pIlpId
	oFunc.ExecuteCN(update)
end sub

%>
	