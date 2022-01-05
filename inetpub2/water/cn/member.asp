<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="navigator.asp" -->
<%
if request("del")<>"" then
	sql = "select * from glControl"
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open sql, conn
	acPeriod=rs("acPeriod")
	acYear=rs("acYear")
	rs.close
	memNo = request("del")
	sql = "select thisShrBal"&acPeriod&",thisLoanBal"&acPeriod&" from memMaster where memNo="&memNo
	rs.open sql, conn
	if rs(0)<>0 or rs(1)<>0 then
		msg = "Can't delete "&memNo&" because it get outstanding balance"
	else
		conn.execute("update memMaster set deleted=-1 where memNo="&memNo)
		msg = memNo&" deleted"
	end if
	rs.close
end if

For Each Field in Request.Form
	TheString = Field & "= Request.Form(""" & Field & """)"
	Execute(TheString)
Next
For Each Field in Request.querystring
	TheString = Field & "= Request.querystring(""" & Field & """)"
	Execute(TheString)
Next
if memNo <> "" then
	sql_filter = sql_filter & " and memNo like '"&memNo&"%'"
end if
if memName <> "" then
	sql_filter = sql_filter & " and memName like '"&memName&"%'"
end if
if memHKID <> "" then
	sql_filter = sql_filter & " and memHKID like '"&memHKID&"%'"
end if
if memSection <> "" then
	sql_filter = sql_filter & " and memSection like '"&memSection&"%'"
end if

sql = "select * from memMaster where deleted=0 " & sql_filter & " order by memNo"
set rs = server.createobject("ADODB.Recordset")
rs.open sql, conn, 3

if searchkey<>"" and rs.recordcount=1 then
	response.redirect "memberDetail.asp?id="&rs("memNo")
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
<title>社員資料修正</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0">
<!-- #include file="menu.asp" -->
<br>
<center>
<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>
<form name="form1" method="post" action="member.asp">
社員編號 : <input type="text" name="memNo" value="<%=memNo%>" size="6">
名稱 : <input type="text" name="memName" value="<%=memName%>" size="10">
身分證號碼 : <input type="text" name="memHKID" value="<%=memHKID%>" size="10">
部門 : <input type="text" name="memSection" value="<%=memSection%>" size="10">
<input type="submit" name="memSearch" value="搜尋">
</form>
<% if request.form("memSearch")<>"" Then %>
<%if recordcount>pagesize then navigator("member.asp?memNo="&memNo&"&memName="&memName&"&memHKID="&memHKID&"&memSection="&memSection) end if%>
<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="336699">
	<tr bgcolor="#330000" align="center">
		<td><font size="2" color="#FFFFFF">社員號碼</font></td>
		<td><font size="2" color="#FFFFFF">名稱</font></td>
		<td><font size="2" color="#FFFFFF">身分證號碼</font></td>
		<td><font size="2" color="#FFFFFF">部門</font></td>
<%if session("userLevel")<>5 and session("userLevel")<>1 and session("userLevel")<>6 then%>
		<td bgcolor="#FFFFFF"><a href="memberDetail.asp"><font size="2">新增</font></a></td>
<%end if%>
	</tr>
<%
do while not rs.eof and rowcount < pagesize
	rowcount = rowcount + 1
%>
	<tr bgcolor="#FFFFFF">
		<td><a href="memberDetail.asp?id=<%=rs("memNo")%>"><font size="2"><%=rs("memNo")%></font></a></td>
		<td><font size="2"><%=rs("memName")%></font></td>
		<td><font size="2"><%=rs("memHKID")%></font></td>
		<td><font size="2"><%=rs("memSection")%></font></td>
<%if session("userLevel")<>5 and session("userLevel")<>1 and session("userLevel")<>6 then%>
		<td><a href="member.asp?del=<%=rs("memNo")%>" onclick="return confirm('刪除此紀錄?')"><font size="2">刪除</font></a></td>
<%end if%>
	</tr>
<%
	rs.movenext
loop
%>
</table>
<%if recordcount>pagesize then navigator("member.asp?memNo="&memNo&"&memName="&memName&"&memHKID="&memHKID&"&memSection="&memSection) end if%>
<%end if%>
</center>
</body>
</html>
