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
<title>���K�X</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
<!--
function validating(){
	if (document.form1.password.value!=document.form1.password1.value){
		alert("�ж�J�ۦP���K�X�έ��J�K�X");
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
		<td class="b8" align="right">�K�X</td>
		<td width=10></td>
		<td><input type="password" name="password" value="<%=password%>" size="50" maxlength="50"></td>
	</tr>
	<tr>
		<td class="b8" align="right">���J�K�X</td>
		<td width=10></td>
		<td><input type="password" name="password1" value="<%=password1%>" size="50" maxlength="50"></td>
	</tr>
	<tr>
		<td colspan="3" align="right">
			<input type="submit" value="�x�s" onclick="return validating()&&confirm('�T�w�x�s?')" name="action" class="sbttn">
		</td>
	</tr>
</table>
</center>
</form>
</body>
</html>
