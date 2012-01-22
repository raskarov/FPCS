<%@ Language=VBScript %>
<%

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name:		lotteryStart.asp
'Purpose:	This script gathers some prequalifing info from the user
'			and then procedes to lotteryMain.asp
'Date:		4-2-2003
'Author:	Scott Bacon (ThreeShapes.com LLC)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim oFunc			'Main object that exposes many of our custom functions 


'set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
'call oFunc.OpenCN()
%>
<html>
	<head>
		<title>Enrollment Step 1</title>
		<link rel="stylesheet" type="text/css" href="../../css/homestyle.css">
	</head>
	<body bgcolor="white">
		<table ID="Table1">
			<tr>
				<td class="NavyHeader">
					&nbsp;<B>Frontier Enrollment Start</B>
				</td>
			</tr>
			<tr>
				<td>
					<table ID="Table2">
						<form action="lotteryMain.asp" name="main" method="post" ID="Form1">
							<input type="hidden" name="fromStart" value="true">
							<tr>
								<td class="gray">
									&nbsp;Do you homeschool full-time in the Anchorage area? <input type="checkbox" name="bolFulltime" ID="Checkbox1" value="true">
									<b>Yes</b>
								</td>
							</tr>
							<tr>
								<td class="gray">
									&nbsp;Are you willing to attend an introductory meeting? <input type="checkbox" name="bolComeToMeeting" ID="Checkbox2" value="true">
									<b>Yes</b>
								</td>
							</tr>
							<tr>
								<td class="gray">
									&nbsp;Are you willing to volunteer five hours over the course of the school 
									year? <input type="checkbox" name="bolVolunteer" ID="Checkbox3" value="true"> <b>Yes</b>
								</td>
							</tr>
							<tr>
								<td class="gray">
									<input type="button" value="Cancel" onclick="window.location.href='default.htm'" ID="Button1" NAME="Button1">
									<input type="submit" value="Continue">
								</td>
							</tr>
					</table>
				</td>
			</tr>
		</table>
	</body>
</html>
