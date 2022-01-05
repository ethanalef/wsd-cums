<%requiredLevel=3%>
<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
searchkey = request("searchkey")
if searchkey = "" then
	sql = "select * from glMaster order by glId"
else
	sql = "select * from glMaster where glId like '"&searchkey&"%' order by glId"
end if

set rs = server.createobject("ADODB.Recordset")
rs.open sql, conn, 3

if searchkey<>"" and rs.recordcount=1 then
	response.redirect "glListDetail.asp?id="&rs("glId")
end if
%>
<html>
<head>
<title>G/L</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" alink="#003399" link="#003399" vlink="#003399" bgcolor="#eeeef0">
<!-- #include file="menu.asp" -->
<br>
<center>
<form name="form1" method="post" action="glList.asp">
G/L Account Number : <input type="text" name="searchkey" value="<%=searchkey%>" size="20"> <input type="submit" value="Search">
</form>
<table border="0" cellspacing="1" cellpadding="4" bgcolor="336699">
  <tr bgcolor="#330000">
	<td align="center"><font size="2" color="#FFFFFF">Account Code</font></td>
	<td align="center"><font size="2" color="#FFFFFF">Description</font></td>
  </tr>
<%
do while not rs.eof %>
  <tr bgcolor="#FFFFFF">
	<td><a href="glListDetail.asp?id=<%=rs("glId")%>"><font size="2"><%=rs("glId")%></font></a></td>
	<td><font size="2"><%=rs("glName")%></font></td>
  </tr>
<%
	rs.movenext
loop %>
</table>
</center>
</body>
</html>