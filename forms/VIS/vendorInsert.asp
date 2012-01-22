<%@ Language=VBScript %>
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		vendorInsert.asp
'Purpose:	Inserts/Updates info coming from vendorAdmin.asp
'Date:		9-5-2001
'Author:	Scott Bacon (ThreeShapes.com LLC)
'
'mod:		bkm 02-sept-02 - removed ticks from bolApproved in SQL
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Dimention Variables, make db Connection
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Response.Write "intCharge_Type_ID:" & request("intCharge_Type_ID")
on error resume next	'turned back on by bkm 28 Jan 2003
dim strMessage	'passed back to the calling page as a QueryString
dim strValue	'modified value to be used in the INSERT or UPDATE statement
dim objRequest	'contains the Form or QueryString object
dim Item			'an Element within objRequest

set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
call oFunc.OpenCN()

if Request.Form.Count > 0 then
	set objRequest = Request.Form
else
	set objRequest = Request.QueryString
end if

For Each Item in objRequest
	execute("dim " & Item)
	if objRequest(Item) = "" then
		strValue ="NULL"
	else
		if left(Item,1) = "s" then	'"s" and "sz" variables we know could have Ticks (')
			'wrap strings with a single quote
			strValue = chr(39) & oFunc.EscapeTick(objRequest(Item)) & chr(39)
		elseif left(Item,3) = "bol" then
			'"bol" variables we know could should be set to 1 or 0
				strValue = oFunc.ConvertCheckToBit(objRequest(Item))
		else
			strValue = objRequest(Item)
		end if
	end if
	execute(Item & " = strValue")
next


'************************************************************
'removed all of the below assignments.  ConvertCheckToBit to handled in the above For..Next loop
'
'bolFor_Pay		= oFunc.ConvertCheckToBit(Request("bolFor_Pay"))
'bolCrime		= oFunc.ConvertCheckToBit(Request("bolCrime"))
'bolConsent		= oFunc.ConvertCheckToBit(Request("bolConsent"))
'bolKids_FPCS	= oFunc.ConvertCheckToBit(Request("bolKids_FPCS"))

'added bkm 24-jan-03
'bolNonProfit	= oFunc.ConvertCheckToBit(Request("bolNonProfit"))
'bolCertASDTeacher= oFunc.ConvertCheckToBit(Request("bolCertASDTeacher"))
'bolCurrentASDEmp = oFunc.ConvertCheckToBit(Request("bolCurrentASDEmp"))
'bolEligibleHire  = oFunc.ConvertCheckToBit(Request("bolEligibleHire"))
'bolConflictIntWaiver= oFunc.ConvertCheckToBit(Request("bolConflictIntWaiver"))
'************************************************************

if request("intVendor_ID") = "" then
	dim insert
	if session.Contents("strRole") = "ADMIN" then
		insert = "insert into tblVendors (szVendor_Name,szVendor_Phone,szVendor_Fax," & _
				 "szVendor_Contact,szVendor_Addr,szVendor_City,sVendor_State,szVendor_Zip_Code," & _
				 "szMail_Addr, szMail_City, sMail_State, szMail_Zip_Code, szVendor_Tax_ID, " & _
				 "szSSN, szBusiness_License, szCert_Insurance, szVendor_Phone_2, szVendor_Email, " & _
				 "szVendor_Website, szVendor_Service, szPrev_Experience, bolFor_Pay, szTraining, " & _
				 "szPrev_For_Pay, intCharge_Type_ID, szOther_Charge_Method, curCharge_Amount, " & _
				 "intWork_Type_ID, bolCrime, bolConsent, bolKids_FPCS,bolApproved, " & _
				 "bolNonProfit, bolCertASDTeacher, bolCurrentASDEmp, bolEligibleHire, bolConflictIntWaiver, " & _
				 "szASDEmpType, szPosition, szWorkLocation, szReasonNoWaiver, szUSER_MODIFY, szUSER_CREATE) values (" & _
				 szVendor_Name & "," & szVendor_Phone & "," & szVendor_Fax & "," & szVendor_Contact & "," & _
				 szVendor_Addr & "," & szVendor_City & "," & sVendor_State & "," & szVendor_Zip_Code & "," & _
				 szMail_Addr & "," &  szMail_City & "," & sMail_State & "," & szMail_Zip_Code & "," & _	
				 szVendor_Tax_ID & "," & szSSN & "," & szBusiness_License & "," & szCert_Insurance & "," & _
				 szVendor_Phone_2 & "," & szVendor_Email & "," & szVendor_Website & "," & szVendor_Service & "," & _	
				 szPrev_Experience & "," & bolFor_Pay & "," & szTraining & "," & szPrev_For_Pay & "," & _
 				 intCharge_Type_ID & "," & szOther_Charge_Method & "," & _
 				 curCharge_Amount & "," & intWork_Type_ID & "," & _
				 bolCrime & "," & bolConsent & "," & bolKids_FPCS & "," & bolApproved & "," & _
				 bolNonProfit & "," & bolCertASDTeacher & "," & bolCurrentASDEmp & "," & bolEligibleHire & "," & _
				 bolConflictIntWaiver & "," & szASDEmpType & "," & szPosition & "," & _
				 szWorkLocation & "," & szReasonNoWaiver & "," & _
				 "'" & session.Value("strUserID") & "', '" & session.Value("strUserID") & "' )"	
	else		
		'lines below are now being handled by a CASE statement in reqGoods.asp
		'szVendor_Name = left(szVendor_Name,len(szVendor_Name)-1)
		'szVendor_Name = szVendor_Name & " - REQUESTED'" 
		
		'bolApproved is required as a hidden field with no value in the vendorAdmin.asp form.
		'another possible solution would be to removed bolApproved from the SQL statement which forces it to NULL.
		
		insert = "insert into tblVendors (szVendor_Name,szVendor_Phone,szVendor_Fax," & _
				 "szVendor_Contact,szVendor_Addr,szVendor_City,sVendor_State,szVendor_Zip_Code," & _
				 "szMail_Addr, szMail_City, sMail_State, szMail_Zip_Code," & _
				 "szVendor_Phone_2, szVendor_Email,bolApproved, " & _
				 "bolNonProfit, bolCertASDTeacher, bolCurrentASDEmp, bolEligibleHire, bolConflictIntWaiver, " & _
				 "szASDEmpType, szPosition, szWorkLocation, szReasonNoWaiver, szUSER_MODIFY, szUSER_CREATE) values (" & _
				 szVendor_Name & "," & szVendor_Phone & "," & szVendor_Fax & "," & szVendor_Contact & "," & _
				 szVendor_Addr & "," & szVendor_City & "," & sVendor_State & "," & szVendor_Zip_Code & "," & _
				 szMail_Addr & "," &  szMail_City & "," & sMail_State & "," & szMail_Zip_Code & "," & _	
				 szVendor_Phone_2 & "," & szVendor_Email & "," & bolApproved & "," & _
				 bolNonProfit & "," & bolCertASDTeacher & "," & bolCurrentASDEmp & "," & bolEligibleHire & "," & _
				 bolConflictIntWaiver & "," & szASDEmpType & "," & szPosition & "," & _
				 szWorkLocation & "," & szReasonNoWaiver & "," & _
				 "'" & session.Value("strUserID") & "', '" & session.Value("strUserID") & "' )"	 	
	end if
	oFunc.ExecuteCN(insert)
	strMessage = "A new vendor has been added."
elseif request("intVendor_ID") <> "" and ucase(request("changed")) = ucase("yes") then
	dim update
	update = "update tblVendors set " & _
			 "szVendor_Name = " & szVendor_Name & "," & _
			 "szVendor_Phone = " & szVendor_Phone & "," & _
			 "szVendor_Fax = " & szVendor_Fax & "," & _
			 "szVendor_Contact = " & szVendor_Contact & ", " & _
			 "szVendor_Addr = " & szVendor_Addr & "," & _
			 "szVendor_City = " & szVendor_City & "," & _
			 "sVendor_State = " & sVendor_State & "," & _
			 "szVendor_Zip_Code = " & szVendor_Zip_Code & "," & _
			 "szMail_Addr = " & szMail_Addr & "," & _
			 "szMail_City = " & szMail_City & "," & _
			 "sMail_State = " & sMail_State & "," & _
			 "szMail_Zip_Code = " & szMail_Zip_Code & ", " & _
			 "szVendor_Tax_ID = " & szVendor_Tax_ID & "," & _
			 "szSSN = " & szSSN & "," & _
			 "szBusiness_License = " & szBusiness_License & "," & _
			 "szCert_Insurance = " & szCert_Insurance & "," & _
			 "szVendor_Phone_2 = " & szVendor_Phone_2 & "," & _
			 "szVendor_Email = " & szVendor_Email & "," & _
			 "szVendor_Website = " & szVendor_Website & "," & _
			 "szVendor_Service = " & szVendor_Service & "," & _
			 "szPrev_Experience = " & szPrev_Experience & "," & _			 
			 "bolFor_Pay = " & bolFor_Pay & ", " & _
			 "szTraining = " & szTraining & "," & _
			 "szPrev_For_Pay = " & szPrev_For_Pay & "," & _
			 "intCharge_Type_ID = " & intCharge_Type_ID & "," & _
			 "szOther_Charge_Method = " & szOther_Charge_Method & "," & _
			 "curCharge_Amount = " & curCharge_Amount & "," & _			
			 "intWork_Type_ID = " & intWork_Type_ID & "," & _
			 "bolCrime = " & bolCrime & "," & _
			 "bolConsent = " & bolConsent & "," & _
			 "bolKids_FPCS = " & bolKids_FPCS & "," & _
			 "bolApproved = " & bolApproved & "," & _
			 "bolNonProfit = " & bolNonProfit & "," & _
			 "bolCertASDTeacher = " & bolCertASDTeacher & "," & _
			 "bolCurrentASDEmp = " & bolCurrentASDEmp & "," & _
			 "bolEligibleHire = " & bolEligibleHire & "," & _
			 "bolConflictIntWaiver = " & bolConflictIntWaiver & "," & _
			 "szASDEmpType = " & szASDEmpType & "," & _
			 "szPosition = " & szPosition & "," & _
			 "szWorkLocation = " & szWorkLocation & "," & _
			 "szReasonNoWaiver = " & szReasonNoWaiver & "," & _
			 "szUser_Modify = '" & session.Value("strUserID") & "' " & _ 
			 " where intVendor_ID = " & intVendor_ID
	oFunc.ExecuteCN(update)
	strMessage = "Vendor has been updated."
else
	'if we get here, it's because we haven't tested for onChange on a field
	strMessage = "No changes were made."
end if

if Err.number <> 0 then
	if Err.number = -2147217873 then
		strMessage = szVendor_Name & " Already Exists.  Please use a different Vendor Name"
	else
		strMessage =  "Error: " & Err.number & ":" & Err.Description & "<br>" & update
	end if
end if
Err.Clear
on error goto 0
dim intIdent
intIdent = oFunc.GetIdentity
set oFunc = nothing
if blnFromItems <> "" then
	' We are in the process of adding an item and the user has just suggested a vendor
	' so we will return them to adding an item. We refresh the page so the new vendor now shows up
	' in the list.
%>
<html>
<head>
<script language=javascript>
	alert("Your vendor has been added to the vendor list as a request. Select the requested vendor to continue.");
	var strURL = window.opener.location.href;
	if (strURL.indexOf("intVendor_ID") == -1){
		strURL+= "&intVendor_ID=<% = intIdent%>"
	}
	//alert(strURL.indexOf("intVendor_ID") + "\n" + strURL);
	window.opener.location.replace(strURL);
	window.opener.focus();
	window.close();
</script>
</head>
<body>
</body>
</html>

<%
elseif request("count") = "" then
	Response.Redirect("../../default.asp?strMessage=" & strMessage)
else
%>
<html>
<head>
<script language=javascript>
	window.opener.jfAddOption('<%=request("count")%>','','<% = request("szVendor_Name") %>');
	window.close();
</script>
</head>
<body>
</body>
</html>

<%
end if 
%>