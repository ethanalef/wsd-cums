<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="navigator.asp" -->
<%
arrLevel = Array("Inactive","Member","Operator","Supervisor","Administrator","Auditor","Preview")

if session("userLevel")=5 then
	sql = "select * from userLog order by uid desc"
else
	sql = "select * from userLog where userLevel<4 order by uid desc"
end if
set rs = server.createobject("ADODB.Recordset")
rs.open sql, conn, 3

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
<title>用戶使用紀錄</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0">
<!-- #include file="menu.asp" -->
<br>
<center>
<%if recordcount>pagesize then navigator("userLog.asp") end if%>
<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="336699">
  <tr bgcolor="#330000" align="center">
	<td><font size="2" color="#FFFFFF">時間</font></td>
	<td><font size="2" color="#FFFFFF">用戶</font></td>
	<td><font size="2" color="#FFFFFF">動作</font></td>
	<td><font size="2" color="#FFFFFF">用戶級別</font></td>
  </tr>
<%
do while not rs.eof and rowcount < pagesize
	rowcount = rowcount + 1
%>
  <tr bgcolor="#FFFFFF">
	<td><%=rs("actionTime")%></td>
	<td><%=rs("username")%></td>
	<td><%=rs("actionDes")%></td>
	<td><%=arrLevel(rs("userLevel"))%></td>
  </tr>
<%
	rs.movenext
loop
%>
</table>
<%if recordcount>pagesize then navigator("userLog.asp") end if%>
</center>
</body>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>