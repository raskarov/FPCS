<%@ Language=VBScript %>
<%
'***************************************************************************
'Name:		debugAppVar.asp
'Purpose:	Iterates through the Session and Application collections and spits
'				them back as HTML
'				A majority of the code is to handled arrays.  We test for the number
'				of dimensions, number of records and then calculate the aprox size
'				of the array in bytes
'
'Author:		Bryan K Mofley
'Date:		7-Apr-2001
'****************************************************************************
option explicit
dim item			'name of the item in the collection
dim strValue	'value of the item in the collection
dim mintDim		'number of dimensions in an array
dim mintRecs	'number of records in the highest dimension of an array
dim mlngBytes	'number of bytes in an array
dim mlngKBytes	'number of Kilo Bytes in an array

function fcnCalcArray(pobjArray)
'***************************************************************************
'Name:		fcnCalcArray (function)
'Purpose:	Calculates the approximate size of an array.  We test for the 
'				number of dimensions, number of records and then compute the 
'				aprox sizeof the array in bytes
'
'Explain:	mlngBytes = (20 + 12) + (4 * mintDim) + ((50/2) * mintRecs)
'				(20 + 12) = 20 is vbSzie of array, add 12 btyes more for a variant array
'				(4 * mintDim) = each dimension of the array requires 4 btyes
'				((50/2) * mintRecs) = we typically have 5 bytes in dim1 and 45 bytes in dim2
'					we add them together, divide by two then multiply by the total # of recs
'see http://msdn.microsoft.com/library/devprods/vs6/vbasic/vbenlr98/vagrpdatatype.htm
'
'Author:		Bryan K Mofley
'Date:		30 May 2001
'****************************************************************************
dim iCnt			'For...Next loop counter
	on error resume next
	mintDim = 0: mintRecs = 0:	mlngBytes = 0: mlngKBytes = 0
	For iCnt = 1 to 100
		mintRecs = UBound(Application(item), iCnt)
		if Err.number = 0 then
			mintDim = mintDim + 1
		else
			exit for
		end if
	next
	on error goto 0
	mintRecs = UBound(Application(item), mintDim) +1
	mlngBytes = (20 + 12) + (4 * mintDim) + ((50/2) * mintRecs)
	mlngKBytes = Round(mlngBytes / 1024, 1)
end function

%>
<HTML>
<HEAD>
	<LINK rel="stylesheet" type="text/css" href="/FPCSdev/css/homestyle.css">
	<TITLE>FPCS: Debug - Application Variables</TITLE>
	<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
</HEAD>
<BODY>
<TABLE>
<%
	'iterate through the Application Collection
	Response.Write "<TR><TD colspan='3'><B>Application Variables</B></TD></TR>" & vbCrLf
	Response.Write "<TR><TD><B>Type&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</B></TD>" & _
						"<TD><B>Name</B></TD><TD><B>Value</B></TD></TR>" 	& vbCrLf
	for each item in Application.Contents
		if IsObject(Application(item)) then
			Response.Write "<TR><TD><FONT color=red><B>Object</B></FONT></TD><TD><FONT color=red>" & _
			 item & "</FONT></TD><TD><FONT color=red>Possible ActiveX component</FONT></TD></TR>" & vbCrLf
		ElseIf IsArray(Application(item)) then
			call fcnCalcArray(Application(item))
			Response.Write "<TR><TD><FONT color=blue><B>Array</B></FONT></TD><TD><FONT color=blue>" & _
			 item & "</FONT></TD><TD><FONT color=blue>" & mintRecs & " Records in a " & _
			 mintDim & " Dimensional Array: approx size " & mlngKBytes & "KB</B></FONT></TD></TR>" & vbCrLf
		Else
			strValue = Application(item)
			if left(item,3) = "cnn" or left(item,8) = "fpcsDev_" or left(item,9) = "FPCSProd_" or left(item,5) = "fpcs_" then
				'if viewer is not an Admin user, do not show the Oracle connetion strings
				if not cbool(Session("bolAdmin")) then
					strValue = "BLOCKED - Must be ADMIN to view"
				end if
			end if
			Response.Write "<TR><TD>Variable</TD><TD>" & _
			 item & "</TD><TD>" & strValue & "</TD></TR>" & vbCrLf
		end if
	next
	
	Response.Write "<TR><TD><BR><BR></TD></TR>" & vbCrLf
	Response.Write "<TR><TD colspan='3'><B>Session Variables</B></TD></TR>" & vbCrLf
	
	'iterate through the Session Collection
	Response.Write "<TR><TD><B>Type</B></TD><TD><B>Name</B></TD><TD><B>Value</B></TD></TR>" & vbCrLf	
	for each item in Session.Contents
		if IsObject(Session(item)) then
			Response.Write "<TR><TD>Object</TD><TD>" & _
			 item & "</TD><TD>Possible ActiveX component</TD></TR>" & vbCrLf
		ElseIf IsArray(Session(item)) then
			call fcnCalcArray(Application(item))
			Response.Write "<TR><TD><FONT color=red><B>Array</B></FONT></TD><TD><FONT color=red><B>" & _
			 item & "</B></FONT></TD><TD><FONT color=red><B>" & mintRecs & " Records in a " & mintDim & " Dimensional Array:" & _
			 " approx size " & mlngKBytes & "KB</B></FONT></TD></TR>" & vbCrLf
		Else
			strValue = Session(item)
			Response.Write "<TR><TD>Variable</TD><TD>" & _
			 item & "</TD><TD>" & strValue & "</TD></TR>" & vbCrLf
		end if
	next
	'Response.Write "Time:" & Server.ScriptTimeout
	'Response.Write 1 / 0 'used to cause an error for debug purposes
%>
</TABLE>
</BODY>
</HTML>
