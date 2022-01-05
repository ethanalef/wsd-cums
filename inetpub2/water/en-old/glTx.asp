<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="../addUserLog.asp" -->
<!-- #include file="navigator.asp" -->
<%
if request("del")<>"" then
	glTxNo = request("del")

	set rs=server.createobject("ADODB.Recordset")
	conn.begintrans
	sql = "select * from glControl"
	rs.open sql, conn
	thisBal="thisBal"&rs("acPeriod")
	rs.close
	sql = "select txAmt,glId,txType from glTx where glTxNo="&glTxNo
	rs.open sql, conn
	txAmt=rs("txAmt")
	glId=rs("glId")
	txType=rs("txType")
	rs.close
	sql = "select "&thisBal&" as thisBal from glMaster where glId='"&glId&"'"
	rs.open sql, conn, 2, 2
	if txType="D" then
		rs(0) = rs(0) - txAmt
	else
		rs(0) = rs(0) + txAmt
	end if
	rs.update
	rs.close
	conn.execute("update glTx set deleted=-1 where glTxNo="&glTxNo)
	addUserLog "Delete G/L Transaction"
	conn.committrans

	msg = glTxNo&"已刪除"
end if

set rs = server.createobject("ADODB.Recordset")
sql = "select * from glTx where deleted=0 order by glTxNo desc"
rs.open sql, conn, 3

if rs.eof then
	response.redirect "glTxDetail.asp"
end if

if not rs.eof then
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
<title>總賬入數</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0">
<!-- #include file="menu.asp" -->
<br>
<center>
<%if msg<>"" then response.write "<font color=red>"&msg&"</font><br>" end if%>
<%if recordcount>pagesize then navigator("glTx.asp") end if%>
<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="336699">
	<tr bgcolor="#330000" align="center">
		<td><font color="#FFFFFF">日期</font></td>
		<td><font color="#FFFFFF">紀錄號碼</font></td>
		<td><font color="#FFFFFF">賬戶號碼</font></td>
		<td><font color="#FFFFFF">內容</font></td>
		<td><font color="#FFFFFF">金額</font></td>
		<td><font color="#FFFFFF">D/C</font></td>
<%if session("userLevel")<>5 then%>
		<td bgcolor="#FFFFFF"><a href="glTxDetail.asp">新增</a></td>
<%end if%>
	</tr>
<%
do while not rs.eof and rowcount < pagesize
	rowcount = rowcount + 1
%>
	<tr bgcolor="#FFFFFF">
		<td><%=right("0"&day(rs("txDate")),2)&"/"&right("0"&month(rs("txDate")),2)&"/"&year(rs("txDate"))%></td>
		<td><%=rs("glTxNo")%></td>
		<td><%=rs("glId")%></td>
		<td><%=rs("txItem")%></td>
		<td align="right"><%if rs("txAmt")<>0 then response.write formatNumber(rs("txAmt"),2) end if%></td>
		<td align="center"><%=rs("txType")%></font></td>
<%if session("userLevel")<>5 then%>
		<td><a href="glTx.asp?del=<%=rs("glTxNo")%>" onclick="return confirm('Delete this record?')">刪除</a></td>
<%end if%>
	</tr>
<%
	rs.movenext
loop
%>
</table>
<%if recordcount>pagesize then navigator("glTx.asp") end if%>
</center>
</body>
</html>
