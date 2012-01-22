<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Global Variables
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim cn
dim strPath
dim intRecordCount
strPath = "http://24.237.12.58/FPCS"

function vbfOpenCN	
	' Creates and Opens main Connection Object for FPCS
	set cn = server.CreateObject("ADODB.CONNECTION")
	cn.Open Application("FPCS_ConnectionString")
end function

function vbfCloseCN
	' Closes and Destroys main Connection Object for FPCS
	cn.Close
	set cn = nothing
end function

function vbfHeader(title,onLoad)
	' Prints Html Header.
	'	Param Desc:
	'		title		Title for the HTML page
	'		onLoad	Pass the name of the function you would like to run when page
	'					is loaded.
%>	
<html>
<head>
<title><% = title %></title>
<link rel="stylesheet" href="<% = strPath %>/CSS/homestyle.css">
<script language=javascript src="<%= strPath %>/includes/formCheck.js">
</script>
</head>
<body bgcolor="#ffffff" onLoad="<% = onLoad %>">
<%
end function
 
function vbfMakeList(valueList,textList,varFind)
	' Creates and returns an HTML option list based on a set of comma seperated strings.
	' Param Desc:
	'	valueList	comma seperated list for the VALUE portion of the option tag
	'	textList	comma seperated list for the TEXT portion of the option tag
	'	varFind		specific option VALUE to find that you would like auto selected in 
	'				the option list if found as this function iterates through the value list.
	dim strOptionList
	dim strSelected
	dim arValues
	dim arText
	dim i
	
	arValues = split(valueList,",")
	
	if textList = "" then
		arText = split(valueList,",")
	else
		arText = split(textList,",")
	end if
	
	for i = 0 to ubound(arValues)
		if cstr(arValues(i)) = cstr(varFind & "") then
			strSelected = " selected"
		else
			strSelected = ""
		end if 
		strOptionList = strOptionList & "<option value=""" & arValues(i) & _
					    """" & strSelected & ">" & arText(i) & chr(13)	
	next
	
	vbfMakeList = strOptionList
end function 
 
function vbfMakeListRS(recordSet,varValue,varText,varFind)
	' Creates and returns an HTML option list from a recordset.
	' Param Desc:
	'	recordSet	ADO recordset that holds the data that you want to create a list from
	'	varValue	field name in recordSet that holds the data that will be placed in
	'				the VALUE portion of the option tag.
	'	varText		field name in recordSet that holds the data that will be placed in
	'				the TEXT portion of the option tag.
	'	varFind		specific option VALUE to find that you would like auto selected in 
	'				the option list if found as this function iterates through the recordset.
	dim strOptionList
	dim strSelected	
	if recordset.recordCount > 0 then
		recordset.MoveFirst

		if varValue = "" then varValue = 0
		if varText = ""	then varText = 1 	
	
		do while not recordSet.EOF
			strSelected = ""			
			if instr(1,varFind,", ") > 0 then
				arFind = split(varFind,", ")
				for x= 0 to ubound(arFind)
					if cstr(recordSet(varValue) & "") = cstr(arFind(x) & "") then
						strSelected = " selected"
					end if 
				next
			else
				if cstr(recordSet(varValue) & "") = cstr(varFind & "") then
					strSelected = " selected"
				end if 
			end if 
			
			strOptionList = strOptionList & "<option value=""" & recordSet(varValue) & _
						    """" & strSelected & ">" & recordSet(varText) & chr(13)
			recordSet.MoveNext		
		loop 
		recordSet.MoveFirst
	end if 	
	vbfMakeListRS = strOptionList
end function

function vbfMakeListSQL(sql,varValue,varText,varFind)
	' Creates and returns an HTML option list from a sql statement.
	' Param Desc:
	'	sql			sql statement to return option list data
	'	varValue	field name in recordSet that holds the data that will be placed in
	'				the VALUE portion of the option tag.
	'	varText		field name in recordSet that holds the data that will be placed in
	'				the TEXT portion of the option tag.
	'	varFind		specific option VALUE to find that you would like auto selected in 
	'				the option list if found as this function iterates through the recordset.
	dim strOptionList
	dim rs
	
	set rs= server.CreateObject("ADODB.RECORDSET")
	rs.CursorLocation = 3	
	'vbfPrint sql
	rs.Open sql,cn
	intRecordCount = rs.RecordCount
	strOptionList = vbfMakeListRS(rs,varValue,varText,varFind)
	rs.Close
	set rs = nothing
	vbfMakeListSQL = strOptionList
end function

function vbfMakeYearList(plusNow,minusNow,varFind)
	' Creates a list of years spanning Now() + plusNow to Now() - minusNow
	' Param Desc:
	'	plusNow		number of years to add to current year.  Used to set end year in list range.
	'   minusNow	number of years to subtract from current year.  Used to set start year in list range.
	'	varFind		specific option VALUE to find that you would like auto selected in 
	'				the option list if found as this function iterates through the year loop.
	dim intYear
	dim intStart
	dim intNow
	dim intEnd
	dim i
	dim strOptionList
	dim strSelected
	
	intNow = datePart("YYYY",now)
	intEnd = intNow + plusNow
	intStart = intNow - minusNow
	
	for i = intStart to intEnd
		if cstr(i) = cstr(varFind & "") then
			strSelected = " selected"
		else
			strSelected = ""
		end if 
		strOptionList = strOptionList & "<option value=""" & i & """" & strSelected & ">" & i & chr(13)
	next
	vbfMakeYearList = strOptionList
end function

function vbfEscapeTick(text)
	'Escapes single ticks (') from passed in text to prevent crashing an
	'insert or update statement in SQL Server.
	text = replace(text,"'","''")
	vbfEscapeTick = text
end function

function vbfGetIdentity
	'Get Identity for last changed id in a session
	dim intID
	dim rsGetID
	
	set rsGetID = server.CreateObject("ADODB.RECORDSET")
	rsGETID.CursorLocation = 3
	sql = "select @@IDENTITY as New_Num"    
	rsGetID.Open sql,cn

	intID = rsGetId("New_Num")

	rsGetID.Close
	set rsGetID = nothing
	
	vbfGetIdentity = intID
end function

function vbfTrueFalse(strTF)
	' Used to set True false text into their numerical values
	if ucase(strTF) = "TRUE" then
		vbfTrueFalse = 1
	else
		vbfTrueFalse = -1
	end if 
end function

function vbfTFText(strTF)
	' Used to set True false numerical values into text
	if ucase(strTF) = "1" then
		vbfTFText = "TRUE"
	else
		vbfTFText = "FALSE"
	end if 
end function

function vbfPrint(text)
	'Since W2K Server displays absolutly nothing but an error message if
	'a script contains an error (does not display the good code in the
	'browser that came before the error) I often am writting stuff to the
	'screen for testing purposes and to see it I need to do a response.end
	'so I made this little function.
	Response.Write text
	Response.End
end function

function vbfSchoolYear()
	' This function returns the current school year based on July 1 
	' beginning a new school year.
	dim intSchoolYear
	dim intCurrentMonth
	
	intCurrentMonth = datePart("m",now())
	
	if intCurrentMonth < 6 then
		intSchoolYear = datePart("yyyy",now())
	else
		intSchoolYear = datePart("yyyy",dateAdd("yyyy",1,now()))
	end if 
	vbfSchoolYear = intSchoolYear
end function 

Function vbfReformat(ByVal pstrData, ByVal parArgs)
'********************************************************
'Name:		vbfReformat (function)
'Purpose:	Formats a string based on optional parameters.
'
'Note:		This is largely a reversed engineered adaptation
'				of the javascript function 'reformat' found in
'				formCheck.js
'
'Usage:		* To reformat a 10-digit U.S. phone number from "1234567890"
'				to "(123) 456-7890" make this function call:
'				vfReformat("1234567890", Array("(", 3, ") ", 3, "-", 4))
'
'				* To reformat a 9-digit U.S. Social Security number from
'				"123456789" to "123-45-6789" make this function call:
'				vfReformat("123456789", Array("", 3, "-", 2, "-", 4))
'
'Date:		10 October 2001
'Author:		Bryan K Mofley (ThreeShapes.com LLC)
'********************************************************
Dim vntArg		'Element in array - variant
Dim intPos		'Integer position in pstrData
dim intLngth	'Length of String
Dim strResult	'Reformated results
Dim i				'Counter in For..Next loop

	intPos = 1
	strResult = ""
	intLngth = Len(pstrData)
	For i = 0 To UBound(parArgs)
		vntArg = parArgs(i)
		If i Mod 2 = 0 Then
			if intpos < intLngth then
				strResult = strResult & vntArg
			else
				exit for
			end if
		Else
			strResult = strResult & Mid(pstrData, intPos, vntArg)
			intPos = intPos + vntArg
		End If
	Next
	vbfReformat = strResult
End Function

function vbfConvertCheckToBit(checkVal)
	' This function converts a 'on' / null value returned by an HTML 
	' checkbox form element to a 1 or 0 respectively
	if ucase(checkVal) = "ON" then
		vbfConvertCheckToBit = 1
	else
		vbfConvertCheckToBit = 0
	end if		
end function

function vbfCheckDecimal(number)
	number = vbfEscapeTick(number)
	if isNumeric(number) then
		vbfCheckDecimal = number
	else
		vbfCheckDecimal = 0
	end if 
end function

function vbfInstructorRate(intInstructor_ID)
	' This function figures out the cost of taxes and benefits 
	' and returns an adjusted hourly rate for the given instructor
	dim sql
	dim dtMax
	dim intCount
	
	' Get Benefit and  Tax Data
	set rsRates = server.CreateObject("ADODB.RECORDSET")
	sql = "select max(dtEffective_Start) from tblBenefit_Tax_Rates"
	rsRates.Open sql,cn
	dtMax = rsRates(0)
	rsRates.Close
	
	sql = "select decTRS,decMedicare,decWorkmans_Comp,decPERS,curHealth_Cost," & _
		  "decFICA,decUnemployment,curLife_Insurance,curFICA_Cap " & _
		  "from tblBenefit_Tax_Rates " & _
		  "where dtEffective_Start = '" & dtMax & "'"
	rsRates.Open sql,cn
	
	intCount = 0
	for each item in rsRates.Fields
		execute("dim " & rsRates.Fields(intCount).Name)
		execute(rsRates.Fields(intCount).Name & " = item")				
		if instr(1,rsRates.Fields(intCount).Name,"dec") > 0 then
			execute(rsRates.Fields(intCount).Name & " = " & "cdbl(" & rsRates.Fields(intCount).Name & ")")
		end if 
		intCount = intCount + 1
	next	
	
	rsRates.Close
	set rsRates = nothing
	
	'Get Teacher Pay Data
	set rsGetPayData = server.CreateObject("ADODB.RECORDSET")
	rsGetPayData.CursorLocation = 3
	sql = "select Max(dtEffective_Start) " & _
		  "from tblInstructor_Pay_Data " & _
		  "where intInstructor_ID = " & request("intInstructor_ID")
	rsGetPayData.Open sql,cn
	
	dtMax = rsGetPayData(0)
	rsGetPayData.Close
	
	sql = "select intInstructor_Pay_Data_ID,intInstructor_ID,curPer_Hour,curPer_Hour_Benefits," & _
		  "curPay_Rate,intPay_Type_id,bolASD_Full_Time,decASD_Full_Time_Percent," & _
		  "bolASD_Part_Time,decASD_Part_Time_Percent,decFPCS_Hours_Goal,dtEffective_Start " & _
		  "from tblInstructor_Pay_Data " & _
		  "where intInstructor_ID = " & request("intInstructor_ID") & _
		  " and dtEffective_Start = '" & dtMax & "'"
	rsGetPayData.Open sql,cn
	intCount = 0
	if rsGetPayData.RecordCount > 0 then
		for each item in rsGetPayData.Fields
			execute("dim " & rsGetPayData.Fields(intCount).Name)
			execute(rsGetPayData.Fields(intCount).Name & " = item")			
			if instr(1,rsGetPayData.Fields(intCount).Name,"dec") > 0 then
				execute(rsGetPayData.Fields(intCount).Name & " = " & "cdbl(" & rsGetPayData.Fields(intCount).Name & ")")
			end if 
			intCount = intCount + 1
		next
	end if 
	
	rsGetPayData.Close
	set rsGetPayData = nothing
	
	if curPer_Hour <> 0 or decFPCS_Hours_Goal <> 0 then 
		if curPer_Hour <> "" or decFPCS_Hours_Goal <> "" then
			decTRS = (decTRS * .01) * curPer_Hour
			decMedicare = (decMedicare * .01) * curPer_Hour
			decWorkmans_Comp = (decWorkmans_Comp * .01) * curPer_Hour
			decPERS = (decPERS * .01) * curPer_Hour
			curHealth_Cost = curHealth_Cost / (1410 * (decFPCS_Hours_Goal * .01))
			decFICA = (decFICA * .01) * curPer_Hour
			decUnemployment = (decUnemployment * .01) * curPer_Hour
			curLife_Insurance = curLife_Insurance / (1410 * (decFPCS_Hours_Goal * .01))
	
			if decASD_Part_Time_Percent >= 50 or decASD_Full_Time_Percent >= 50 then
				'Addendum Teacher
				vbfInstructorRate = curPer_Hour + decTRS + decMedicare + decWorkmans_Comp
			elseif (decASD_Part_Time_Percent + decFPCS_Hours_Goal) >= 50 and (decFPCS_Hours_Goal >= 20) _
				   and (decASD_Part_Time_Percent < 50) then
				vbfInstructorRate = curPer_Hour + decTRS + decMedicare + curHealth_Cost + decWorkmans_Comp
			elseif (decASD_Full_Time_Percent + decFPCS_Hours_Goal) >= 50 and (decFPCS_Hours_Goal >= 20) _
				   and (decASD_Full_Time_Percent < 50) then
				'Same as above except this teacher is paid by FPCS on not ASD
				vbfInstructorRate = curPer_Hour + decTRS + decMedicare + curHealth_Cost + decWorkmans_Comp
			elseif decFPCS_Hours_Goal < 37.5 then
				'Special Activity Agreement
				vbfInstructorRate = curPer_Hour + decFICA + decMedicare + decWorkmans_Comp
			elseif decFPCS_Hours_Goal >=37.5 and decFPCS_Hours_Goal < 50 then
				vbfInstructorRate = curPer_Hour + decPERS + decFICA + decMedicare + decWorkmans_Comp
			end if 
			
		end if
	end if 
end function
%>
