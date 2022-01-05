<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
sql = "select * from cheque where chequeClear=0 order by chequeNum"
set rs = server.createobject("ADODB.Recordset")
rs.open sql, conn, 3
%>
<html>
<head>
<title>支票對數</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0">
<!-- #include file="menu.asp" -->
<form method="Post" action="chequeClear.asp" name="form1">
<br>
<center>
<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="336699">
  <tr bgcolor="#330000" align="center">
	<td><font color="#FFFFFF">支票號碼</font></td>
	<td><font color="#FFFFFF">支票日期</font></td>
	<td><font color="#FFFFFF">收款人</font></td>
	<td><font color="#FFFFFF">金額</font></td>
<%if session("userLevel")<>5 then%>
	<td bgcolor="#FFFFFF"><a href="chequeDetail.asp"><font size="2">新增</font></a></td>
<%end if%>
  </tr>
<%
if not rs.eof then
	do while not rs.eof
%>
  <tr bgcolor="#FFFFFF">
	<td><a href="chequeDetail.asp?id=<%=rs("uid")%>"><%=rs("chequeNum")%></a></td>
	<td><%=right("0"&day(rs("chequeDate")),2)&"/"&right("0"&month(rs("chequeDate")),2)&"/"&year(rs("chequeDate"))%></td>
	<td><%=rs("payee")%></td>
	<td><%=formatnumber(rs("amount"),2)%></td>
<%if session("userLevel")<>5 then%>
	<td><input type="checkbox" name="TS" value="<% =rs("uid") %>"></td>
<%end if%>
  </tr>
<%
		rs.movenext
	loop
%>
<%if session("userLevel")<>5 then%>
  <tr bgcolor="#FFFFFF">
  	<td colspan="5" align="right"><input type="submit" name="clearBttn" value="對數" class="sbttn"></td>
  </tr>
<%end if%>
<%
end if
%>
</table>
</form>
</body>
</html>
