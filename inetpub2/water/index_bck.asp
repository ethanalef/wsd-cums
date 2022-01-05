<!-- #include file="conn.asp" -->

<%
if request("From") = Request.ServerVariables("script_name") and request.form("username") <> "" then
	set reg = new regexp
	reg.pattern="[^a-zA-Z0-9]"
	reg.Global = True
	username=reg.replace(request("username"),"")
	password=reg.replace(request("password"),"")

    SQL = "select * from loginUser where username ='" & username & "' and password ='" & password & "'"
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.open sql, conn,1 ,3
    if rs.eof then
        msg = "Login Failed"
    else
	session.timeout = 1200
        session("userLevel") = rs("userLevel")
        session("username") = rs("username")
        session("workstart")=now
        rs.update
        rs.close
        set rs=nothing
        conn.close
        set conn=nothing
        Response.redirect "en/main.asp"
    end if
	rs.close
	set rs=nothing
end if
conn.close
set conn=nothing
%>
<html>
<head>
<title>水務署員工儲蓄互助社系統</title>
<meta http-equiv="content-type" content="text/html; charset=big5">
<link href="main.css" rel="stylesheet" type="text/css">
<script language=JavaScript>
<!--
function validating()
{
    if(document.login.username.value=="" || document.login.password.value=="")
    {
        alert("Please fill in both Username and Password");
        return false;
    }else{
        return true;
    }
}
//-->
</script>
</head>
<body bgcolor="#ffffff" leftMargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="login.username.focus()">
<div align="center">
<center>
<form method="POST" name="login" action="<% =Request.servervariables("script_name") %>" onSubmit="return validating()">
<input type=hidden name="From" value="<% =Request.servervariables("script_name") %>">
<br><br>
<font face="arial, helvetica, sans-serif" size="5" color="#336699"><b>水務署員工儲蓄互助社系統</b></font><br>
<font face="arial, helvetica, sans-serif" size="4" color="#336699"><b>Water Supplies Department Staff Credit Union<br>Membership, Accounting, Savings and Loans Software</b></font>
<br><br>
<img src="images/image002.gif" broder="0">
<br><br>
<font face="arial, helvetica, sans-serif" size="3" color="#000000">請輸入名稱及密碼登入系統<br>Please Login With Your Username and Password</font>
<br><br>
<table border="0" cellpadding="0" cellspacing="0">
    <tr>
        <td width="130"><b>名稱 Username</b></td>
        <td width="170"><input type="text" name="username" size="20"></td>
        <td width="100">&nbsp;</td>
    </tr>
    <tr>
        <td><b>密碼 Password</b></td>
        <td><input type="password" name="password" size="20"></td>
        <td><input type="submit" value="登入 Login"></td>
    </tr>
    <tr>
        <td colspan=2 height=60>
<%  if msg <> "" then %>
            <center><font color="#0000ff"><b><%= msg%></b></font></center>
<% end if %>
        </td>
    </tr>
</table>
<br><br>
<font size="2">Best Viewed With Microsoft Internet Explorer 5.0 or Higher</font>
<font size="2">Credur Union Ver 1.01.200604</font>
</form>
</center>
</div>
</body>
</html>
