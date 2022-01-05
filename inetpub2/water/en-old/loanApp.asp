<%requiredLevel=4%>
<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
if request("del")<>"" then
	glId = request("del")
	set rs = server.createobject("ADODB.Recordset")
	sql = "select count(*) from glTx where glId='"&glId&"'"
	rs.open sql, conn
	if rs(0) > 0 then
		msg = "Can't delete "&glId&" because it get transaction record"
	else
		conn.execute("delete from glMaster where glId='"&glId&"'")
		msg = glId&" deleted"
	end if
	rs.close
end if

searchkey = request("searchkey")
if searchkey = "" then
	sql = "select * from loanApp order by appDate desc"
else
	sql = "select * from loanApp where memNo like '"&searchkey&"%' order by appDate desc"
end if
set rs = server.createobject("ADODB.Recordset")
rs.open sql, conn, 3

if searchkey<>"" and rs.recordcount=1 then
	response.redirect "loanAppDetail.asp?uid="&rs("uid")
end if
%>
<html>
<head>
<title>Loan Application</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" alink="#003399" link="#003399" vlink="#003399" bgcolor="#eeeef0">
<!-- #include file="menu.asp" -->
<br>
<center>
<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>
<form name="form1" method="post" action="gl.asp">
G/L Account Number : <input type="text" name="searchkey" value="<%=searchkey%>" size="20"> <input type="submit" value="Search">
</form>
<table border="0" cellspacing="1" cellpadding="4" bgcolor="336699">
  <tr bgcolor="#330000" align="center">
	<td><font size="2" color="#FFFFFF">Account Code</font></td>
	<td><font size="2" color="#FFFFFF">Description</font></td>
	<td bgcolor="#FFFFFF"><a href="glDetail.asp"><font size="2">Add</font></a></td>
  </tr>
<%
do while not rs.eof %>
  <tr bgcolor="#FFFFFF">
	<td><a href="glDetail.asp?uid=<%=rs("uid")%>"><font size="2"><%=rs("glId")%></font></a></td>
	<td><font size="2"><%=rs("glName")%></font></td>
	<td><a href="gl.asp?del=<%=rs("glId")%>" onclick="return confirm('Delete this record?')"><font size="2">Delete</font></a></td>
  </tr>
<%
	rs.movenext
loop %>
</table>
<br>
</center>
</body>
</html>
