<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="navigator.asp" -->
<%
arrLevel = Array("Inactive","Member","Operator","Supervisor","Administrator","Auditor","Preview")

searchkey = request("searchkey")
if searchkey = "" then
	sql = "select * from loginUser where userlevel<>5 order by userLevel,username"
else
	sql = "select * from loginUser where userlevel<>5 and username like '"&searchkey&"%' order by userLevel,username"
end if

set rs = server.createobject("ADODB.Recordset")
rs.open sql, conn, 3

if searchkey<>"" and rs.recordcount=1 then
	response.redirect "userDetail.asp?id="&rs("username")
end if

if Not rs.eof then
	if request("page") <> "" then
		pageno = cint(request("page"))
	else
		pageno = 1
	end if
	rs.pagesize = 20
	pagesize=rs.pagesize
	rs.absolutepage = pageno
	recordcount=rs.recordcount
	pagecount = rs.pagecount
end if
%>
<html>
<head>
<title>用戶管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0">
<!-- #include file="menu.asp" -->
<br>
<center>
<form name="form1" method="post" action="user.asp">
用戶名稱 : <input type="text" name="searchkey" value="<%=searchkey%>" size="20"> <input type="submit" value="搜尋">
</form>
<%if recordcount>pagesize then navigator("user.asp") end if%>
<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="336699">
	<tr bgcolor="#330000" align="center">
		<td><font size="2" color="#FFFFFF">用戶級別</font></td>
		<td><font size="2" color="#FFFFFF">用戶名稱</font></td>
		<td bgcolor="#FFFFFF"><a href="userDetail.asp"><font size="2">新增</font></a></td>
	</tr>
<%
do while not rs.eof and rowcount < pagesize
	rowcount = rowcount + 1
%>
	<tr bgcolor="#FFFFFF">
		<td><font size="2"><%=arrLevel(rs("userLevel"))%></font></td>
		<td><a href="userDetail.asp?id=<%=rs("uid")%>"><font size="2"><%=rs("username")%></font></a></td>
		<td></td>
	</tr>
<%
	rs.movenext
loop
%>
</table>
<%if recordcount>pagesize then navigator("user.asp") end if%>
</center>
</body>
</html>