<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="navigator.asp" -->
<%
if request("del")<>"" then
	del = request("del")
	conn.execute("update monthlyReport set deleted=-1 where uid="&del)
	msg = "報告書 "&del&" 已刪除"
end if

sql = "select * from monthlyReport where deleted=0 order by uid desc"
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
<title>董事會報告書</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0">
<!-- #include file="menu.asp" -->
<br>
<center>
<%if recordcount>pagesize then navigator("monthlyReport.asp") end if%>
<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="336699">
	<tr bgcolor="#330000" align="center">
		<td><font size="2" color="#FFFFFF">日期</font></td>
		<td><font size="2" color="#FFFFFF">報告書</font></td>
<%if session("userLevel")<>5 then%>
		<td bgcolor="#FFFFFF"><a href="monthlyReportDetail.asp"><font size="2">新增</font></a></td>
<%end if%>
	</tr>
<%
do while not rs.eof and rowcount < pagesize
	rowcount = rowcount + 1
%>
  <tr bgcolor="#FFFFFF">
	<td><a href="monthlyReportDetail.asp?id=<%=rs("uid")%>"><font size="2"><%=right("0"&day(rs("rpDate")),2)&"/"&right("0"&month(rs("rpDate")),2)&"/"&year(rs("rpDate"))%></font></a></td>
	<td align="right"><font size="2"><%=rs("uid")%></font></td>
<%if session("userLevel")<>5 then%>
	<td><a href="monthlyReport.asp?del=<%=rs("uid")%>" onclick="return confirm('刪除此紀錄?')"><font size="2">刪除</font></a></td>
<%end if%>
  </tr>
<%
	rs.movenext
loop
%>
</table>
<%if recordcount>pagesize then navigator("monthlyReport.asp") end if%>
</center>
</body>
</html>
