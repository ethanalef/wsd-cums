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
		msg = "����R��"&glId&", �]������ᴿ�g���ө�����"
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
<title>�`�b</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0">
<!-- #include file="menu.asp" -->
<br>
<center>
<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>
<form name="form1" method="post" action="gl.asp">
�`�b�s�� : <input type="text" name="searchkey" value="<%=searchkey%>" size="20"> <input type="submit" value="�j�M">
</form>
<table border="0" cellspacing="1" cellpadding="4" bgcolor="336699">
  <tr bgcolor="#330000" align="center">
	<td><font size="2" color="#FFFFFF">�s��</font></td>
	<td><font size="2" color="#FFFFFF">���e</font></td>
	<td><font size="2" color="#FFFFFF">����</font></td>
<%if session("userLevel")<>5 then%>
	<td bgcolor="#FFFFFF"><a href="glDetail.asp"><font size="2">�s�W</font></a></td>
<%end if%>
  </tr>
<%
do while not rs.eof %>
  <tr bgcolor="#FFFFFF">
	<td><a href="glDetail.asp?glId=<%=rs("glId")%>"><font size="2"><%=rs("glId")%></font></a></td>
	<td><font size="2"><%=rs("glName")%></font></td>
	<td><font size="2"><%=mType(rs("glType"))%></font></td>
<%if session("userLevel")<>5 then%>
	<td><a href="gl.asp?del=<%=rs("glId")%>" onclick="return confirm('�R�������?')"><font size="2">�R��</font></a></td>
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
