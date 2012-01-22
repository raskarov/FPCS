<%@ Language=VBScript %>
<html>
	<body>


	<%
	
'Create object containing all of our FPCS functions
set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
oFunc.OpenCnn

arTest = oFunc.GetStudentBalances(50)
 response.Write arTest(0)
	%>
	</body>
</html>
