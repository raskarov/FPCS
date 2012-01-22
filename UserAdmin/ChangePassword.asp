<%@ Language=VBScript %>
<%
'*******************************************
'Name:		UserAdmin\ChangePassword.asp
'Purpose:	Contains Form and sql code needed to change a password for an existig user
'
'
'Author:	Scott Bacon
'			3Shapes is Scott Bacon, Bryan Mofley, Guy Mofley
'Date:   15-May-2000
'*******************************************
if Request.Form("szPassword1") <> "" then
	dim update
	private strEncPwd
	dim fvCount
	
	set oCrypto = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/Crypto.wsc"))
		'encrypt password for database compare
		'oCrypto.Key = "something"	'actual key is not shown here
		oCrypto.Text = Request.Form("szPassword1")
		Call oCrypto.Encypttext
		strEncPwd = oCrypto.EncryptedText
	set oCrypto = nothing
	
	set oFunc = GetObject ("script:" & Server.MapPath(Application.Value("strWebRoot") & "wsc/FPCSfunctions.wsc"))
	call oFunc.OpenCN()
	
	fvCount = Request.Form("intCount")
	
	update = "update tblUsers set szPassword = '" & strEncPwd & "' " & _
			 "where szUser_ID = '" & Session("strUserID") & "'"
	oFunc.ExecuteCN(update)		
	
	if fvCount <> "" then
		oFunc.ForcedActionHandling oFunc,fvCount 
	else
		oFunc.CloseCN
		set oFunc = nothing
		if ucase(Session.Contents("strRole")) = "VENDOR" then
			Response.Redirect(Application.Contents("strSSLWebRoot") & "/VendorHome.asp")
		else
			Response.Redirect(Application.Contents("strSSLWebRoot"))
		end if
	end if 
end if 


Server.Execute(Application.Value("strWebRoot") & "Includes/simpleheader.asp")
%>
<SCRIPT language=javascript>	
function jfCheckValue(form) {
	if (form.szPassword1.value == "" || form.szPassword2.value == "") {
		alert("You must provide a value for both 'New Password' and 'Confirm Password'");
		form.szPassword1.focus();
		return;
	}
	if (form.szPassword1.value != form.szPassword2.value) {
		alert("The passwords you entered did not match.\nPlease enter a value in 'New Password' and confirm it in 'Confirm Password'");
		form.szPassword1.focus();
		return;
	}
	form.submit();
}	
</script>

<center>
<table cellpadding="3" border="0" cellspacing="0">
	<tr>
		<td align="left">&nbsp;</td>
	</tr>
	<tr>		
		<td>
			<img src="../images/fpcsLogo.gif" >
		</td>
	</tr>
	<tr>
		<td align="left"  class="yellowHeader">
			&nbsp;<b>Change Password</b>			
		</td>
	</tr>
	<tr align="left">
		<td>
			<table cellpadding="1" cellspacing="1">
				<form action="<% = Application.Value("strWebRoot") %>UserAdmin/ChangePassword.asp" method=post name=frmLogin onSubmit="return false;">
				<input type=hidden name="intCount" value="<%=Request.QueryString("intCount")%>">
				<tr valign="middle">
					<td>
						<font size="-1" face=tahoma><b>New Password</b></font>
					</td>
					<td>
						<input type="password" size="20" name="szPassword1" value="<%=strUserID%>">
					</td>
				</tr>
				<tr valign="middle">
					<td>
						<font size="-1" face=tahoma><b>Confirm Password:</b></font>
					</td>
					<td>
						<input type="password" size="20" name="szPassword2">
					</td>
				</tr>
				<tr>
					<td colspan=2 bgcolor=f0f0f0 align=right>
						<input type=hidden name=hLogin value=True>
						<input type="submit" value="Submit" id=cmdLogin name=cmdLogin onClick="jfCheckValue(this.form);">
					</td>
				</tr>
				</form>
			</table>
		</td>
	</tr>
</table>
</center>
</body>
</html>