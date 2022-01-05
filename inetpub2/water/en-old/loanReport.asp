<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="navigator.asp" -->
<%
searchkey = request("searchkey")
if searchkey = "" then
	sql = "select memNo,memName from memMaster order by memNo"
else
	sql = "select memNo,memName from memMaster where memNo like '"&searchkey&"%' order by memNo"
end if

set rs = server.createobject("ADODB.Recordset")
rs.open sql, conn, 3

if searchkey<>"" and rs.recordcount=1 then
	response.redirect "loanReportDetail.asp?id="&rs("memNo")
end if

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
%>
<html>
<head>
<title>貸款申請列表</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.searchkey.focus()">
<!-- #include file="menu.asp" -->
<br>
<center>
<form name="form1" method="post" action="loanReport.asp">
社員號碼 : <input type="text" name="searchkey" value="<%=searchkey%>" size="20"> <input type="submit" value="搜尋">
</form>
<%if recordcount>pagesize then navigator("loanReport.asp?searchkey="&searchkey) end if%>
<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="336699">
	<tr bgcolor="#330000" align="center">
		<td><font size="2" color="#FFFFFF">社員號碼</font></td>
		<td><font size="2" color="#FFFFFF">姓名</font></td>
	</tr>
<%
do while not rs.eof and rowcount < pagesize
	rowcount = rowcount + 1
%>
	<tr bgcolor="#FFFFFF">
		<td><a href="loanReportDetail.asp?id=<%=rs("memNo")%>"><font size="2"><%=rs("memNo")%></font></a></td>
		<td><font size="2"><%=rs("memName")%></font></td>
	</tr>
<%
	rs.movenext
loop
%>
</table>
<%if recordcount>pagesize then navigator("loanReport.asp?searchkey="&searchkey) end if%>
</center>
</body>
</html>
