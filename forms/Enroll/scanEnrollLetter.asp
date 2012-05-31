<%@ Language=VBScript %>
<%
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()
Server.Execute(Application.Value("strWebRoot") & "includes/header.asp")

if Request.Form("scanData") <> "" then	
	dim rsState
	dim i
	dim intStateValue
	dim update
	dim insert
	dim strAlert
	dim strIdType
	dim strList
	
	strList = replace(replace(replace(Request.Form("scanData")," ",""),chr(13),""),chr(10),"")

	arData = split(strList,",")
	set rsState = server.CreateObject("ADODB.RECORDSET")
	rsState.CursorLocation = 3
	
	
	'Sets id to student or lottery fields
	'if Request.Form("scanType") = "currentStudents" then
	'	strIdType = "intStudent_ID"
	'else
'		strIdType = "intLottery_ID"
	'end if 
	
	for i = 0 to ubound(arData)
		if arData(i) <> "" then 
			arState = split(arData(i),"!")
			'If we have invalid data then arState will have a ubound of 0
			if ubound(arState) > 0  then
				' Give the state it's matrix value
				' (values viewable in the database in trefReEnroll_States)
				if arState(1) = "1" then			
					intStateValue = 7 ' Yes to reenroll
				else 
					intStateValue = 67 ' No to reenroll
				end if 
		
				' Check for existing state record
				sqlCheck = "select intReEnroll_State from tblStudent_States " & _
						   "where intStudent_ID = " & arState(0) & _
						   " and intSchool_Year = " & Request.Form("intSchool_Year")
		
				rsState.Open sqlCheck, Application("cnnFPCS")'oFunc.FPCScnn
					
				if rsState.RecordCount > 0 then
					'The following is a case where the staterecord needs to be updated
					if cint(rsState("intReEnroll_State")) < 7 then
						' Any value greater than 7 is a state that we can not over write
						' because that recorded state is further into the matrix than
						' this script is designed to handle. 
						update = "update tblStudent_States set intReEnroll_State = " & intStateValue & _
								 " , dtModify = '" & now() & "' " & _
							     " where intStudent_ID = " & arState(0) & _
								 " and intSchool_Year = " & Request.Form("intSchool_Year")
						oFunc.ExecuteCN(update)
					end if
				else
					'Insert new record 				

					insert = "insert into tblStudent_States(intStudent_ID,intReEnroll_State,intSchool_Year, szGrade, dtCreate,szUser_CREATE) " & _
						 " SELECT " & arState(0) & "," & intStateValue & "," & Request.Form("intSchool_Year") & ", " & _
						 "	coalesce((select CASE ss.szGrade WHEN 'K' THEN '1' WHEN '12' THEN '12' ELSE CONVERT(varchar, CONVERT(int, ss.szGrade) + 1) END" & _ 
						 "	FROM tblStudent_States ss " & _	
						 "	WHERE (ss.intSchool_Year = " & cint(Request.Form("intSchool_Year")) - 1 & " AND (ss.intStudent_id = " & arState(0) & "))),'K') , CURRENT_TIMESTAMP,'" & session.Contents("strUserId")  & "' " & _
						 "	from tblStudent where intStudent_ID = " & arState(0)

response.write insert
					oFunc.ExecuteCN(insert)
				end if		
				'Reset rs and array
				rsState.Close	
				erase arState		
			end if
		end if 
	next
	set rsState = nothing
	
	strAlert =  "alert('Updates Finished.');"
end if
%>
<script language=javascript>
	<% = strAlert %>
	function jfConfirm(){
		if(document.main.scanData.value == "") {
			alert("The text box where the scanned barcodes are supose to go is blank.\nPlease click in the text box and then start scanning.");
			document.main.scanData.focus();
			return;
		}
		var year = document.all.item("intSchool_Year").value
		var strMessage = "IMPORTANT! Please confirm that these scanned forms are for the ";
		strMessage += year + " school year.\nIf not click 'Cancel' and change to the correct school year."
		var bolConfirm = confirm(strMessage);
		if (bolConfirm == true) {
			main.submit();
		}			
	}
</script>
<form name=main method=post action="scanEnrollLetter.asp" onSubmit="return false">
<table width="100%">
	<tr>
		<Td class="yellowHeader">
			&nbsp;<b>FPCS Re-Enrollment Scan Interface</b>
		</Td>
	</tr>
	<tr>
		<td bgcolor="f7f7f7">
			<table>
				<tr>
					<td class="gray">
							&nbsp;<b>Select the School Year for Re-Enrollment:</b>
					</td>
					<Td>						
						<select name="intSchool_Year">
							<%
								= oFunc.MakeYearList(2,1,(datePart("yyyy",now())+1))
							%>
						</select>
					</td>
				</tr>
				<!--<tr>
					<td class="gray">
							&nbsp;<b>Are you scanning Current Students or Lottery Winners?</b>
					</td>
					<Td>						
						<select name="scanType">
							<option value="currentStudents">Current Students
							<option value="lottery">Lottery Winners
						</select>
					</td>
				</tr>-->
				<tr>
					<td colspan=2  class="gray">
						&nbsp;<b>Make sure you click in the text box below 
						and then <BR> 
						&nbsp;scan the selected barcodes off the enrollment forms.</b><BR>
						<textarea cols=60 rows=25 name=scanData></textarea>											
					</td>
				</tr>
				<tr>
					<Td colspan=2>
						<input type=submit value="submit" onClick="jfConfirm();">	
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>	
</form>
<%
	Server.Execute(Application.Value("strWebRoot") & "Includes/footer.asp") 
%>
