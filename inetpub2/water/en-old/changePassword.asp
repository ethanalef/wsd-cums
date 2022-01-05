<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
if request.form("action") <> "" then
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next
    set rs = server.createobject("ADODB.Recordset")
	conn.begintrans
	sql = "select * from loginUser where userLevel=5"
	rs.open sql, conn, 2, 2
	if password<>"" then rs("password") = password end if
	rs.update
	conn.committrans
	msg = "Record Updated"
end if
%>
<html>
<head>
<title>更改密碼</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
<!--
function validating(){
	if (document.form1.password.value!=document.form1.password1.value){
		alert("請填入相同的密碼及重入密碼");
        return false;
    }else{
        return true;
    }
}
//  -->
</script>
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.password.focus()">
<!-- #include file="menu.asp" -->
<center>
<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>
<br>
<form name="form1" method="post" action="changePassword.asp">
<table border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td class="b8" align="right">密碼</td>
		<td width=10></td>
		<td><input type="password" name="password" value="<%=password%>" size="50" maxlength="50"></td>
	</tr>
	<tr>
		<td class="b8" align="right">重入密碼</td>
		<td width=10></td>
		<td><input type="password" name="password1" value="<%=password1%>" size="50" maxlength="50"></td>
	</tr>
	<tr>
		<td colspan="3" align="right">
			<input type="submit" value="儲存" onclick="return validating()&&confirm('確定儲存?')" name="action" class="sbttn">
		</td>
	</tr>
</table>
</center>
</form>
</body>
</html>
