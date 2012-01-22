<%@ Language=VBScript %>
<%
	dim sURL

	'if Session.Value("strURL") & "" = "" then
		'sURL = Application("strWebRoot") & "default.asp"
		sURL = "http://localhost/empower/" & "default.asp"
	'else
		sURL = Session.Value("strURL")
	'end if
%>
<HTML>
<HEAD>
	<title>FPCS - Application Starting...</title>
	<script language="javascript">
		function jfLaunchApp(){
			var strURL = "<%= sURL %>";
			strFeatures = "toolbar=no,scrollbars=yes,resizable=yes,status=yes";
			var app = window.open(strURL, '<% = Session.Value("strAppWindow")%>', strFeatures);
			//window.location.replace('http://<% = Request.ServerVariables("SERVER_NAME")%>');
			//window.resizeTo(screen.availWidth / 2, screen.availHeight / 2);
			app.moveTo(0, 0);
			app.resizeTo(screen.availWidth , screen.availHeight);
			app.focus();
		}
	</script>
</HEAD>
<BODY onload="jfLaunchApp();" style="font-family:arial;font-size:9pt;">
<font face="Verdana">You have logged in successfully.  If the 
FPCS application pop up window did not appear please
click <a href="javascript:" onclick="jfLaunchApp();"><b>HERE</b></a> to open it.</font>
<br><br>
If you have further problems please email <a href="help@3shapes.com">help@3shapes.com</a>.
</BODY>
</HTML>