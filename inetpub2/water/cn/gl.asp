<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
dim mType(6)
mType(1) = "Fixed Assets"
mType(2) = "Loans"
mType(3) = "Current Assets"
mType(4) = "Liabilities"
mType(5) = "Income"
mType(6) = "Expenses"

if request("del")<>"" then
	glId = request("del")
	set rs = server.createobject("ADODB.Recordset")
	sql = "select count(*) from glTx where glId='"&glId&"'"
	rs.open sql, conn
	if rs(0) > 0 then
		msg = "不能刪除"&glId&", 因為此賬戶曾經有來往紀錄"
	else
		conn.execute("update glMaster set deleted=-1 where glId='"&glId&"'")
		msg = glId&" deleted"
	end if
	rs.close
end if

searchkey = request("searchkey")
if searchkey = "" then
	sql = "select glId,glName,glType from glMaster where deleted=0 order by glId"
else
	sql = "select glId,glName,glType from glMaster where deleted=0 and glId like '"&searchkey&"%' order by glId"
end if
set rs = server.createobject("ADODB.Recordset")
rs.open sql, conn, 3

if searchkey<>"" and rs.recordcount=1 then
	response.redirect "glDetail.asp?glId="&rs("glId")
end if
%>
<html>
<head>
<title>總帳</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0">
<!-- #include file="menu.asp" -->
<br>
<center>
<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>
<form name="form1" method="post" action="gl.asp">
總帳編號 : <input type="text" name="searchkey" value="<%=searchkey%>" size="20"> <input type="submit" value="搜尋">
</form>
<table border="0" cellspacing="1" cellpadding="4" bgcolor="336699">
  <tr bgcolor="#330000" align="center">
	<td><font size="2" color="#FFFFFF">編號</font></td>
	<td><font size="2" color="#FFFFFF">內容</font></td>
	<td><font size="2" color="#FFFFFF">分類</font></td>
<%if session("userLevel")<>5 then%>
	<td bgcolor="#FFFFFF"><a href="glDetail.asp"><font size="2">新增</font></a></td>
<%end if%>
  </tr>
<%
do while not rs.eof %>
  <tr bgcolor="#FFFFFF">
	<td><a href="glDetail.asp?glId=<%=rs("glId")%>"><font size="2"><%=rs("glId")%></font></a></td>
	<td><font size="2"><%=rs("glName")%></font></td>
	<td><font size="2"><%=mType(rs("glType"))%></font></td>
<%if session("userLevel")<>5 then%>
	<td><a href="gl.asp?del=<%=rs("glId")%>" onclick="return confirm('刪除此賬戶?')"><font size="2">刪除</font></a></td>
<%end if%>
  </tr>
<%
	rs.movenext
loop %>
</table>
<br>
</center>
</body>
</html>
