<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="navigator.asp" -->
<%
arrLevel = Array("Inactive","Active")

searchkey = request("searchkey")
if searchkey = "" then
	sql = "select * from handleParty order by status desc,handleName"
else
	sql = "select * from handleParty where handleName like '"&searchkey&"%' order by status desc,handleName"
end if

set rs = server.createobject("ADODB.Recordset")
rs.open sql, conn, 3

if searchkey<>"" and rs.recordcount=1 then
	response.redirect "handlePartyDetail.asp?id="&rs("handleName")
end if

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
<title>委員資料修正</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0">
<!-- #include file="menu.asp" -->
<br>
<center>
<form name="form1" method="post" action="handleParty.asp">
委員名稱 : <input type="text" name="searchkey" value="<%=searchkey%>" size="20"> <input type="submit" value="搜尋">
</form>
<%if recordcount>pagesize then navigator("handleParty.asp") end if%>
<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="336699">
	<tr bgcolor="#330000" align="center">
		<td><font size="2" color="#FFFFFF">委員名稱</font></td>
		<td><font size="2" color="#FFFFFF">狀況</font></td>
<%if session("userLevel")<>5 then%>
		<td bgcolor="#FFFFFF"><a href="handlePartyDetail.asp"><font size="2">新增</font></a></td>
<%end if%>
	</tr>
<%
do while not rs.eof and rowcount < pagesize
	rowcount = rowcount + 1
%>
	<tr bgcolor="#FFFFFF">
		<td><a href="handlePartyDetail.asp?id=<%=rs("uid")%>"><font size="2"><%=rs("handleName")%></font></a></td>
		<td><font size="2"><%if rs("status")=0 then%>Inactive<%else%>Active<%end if%></font></td>
<%if session("userLevel")<>5 then%>
		<td></td>
<%end if%>
	</tr>
<%
	rs.movenext
loop
%>
</table>
<%if recordcount>pagesize then navigator("handleParty.asp") end if%>
</center>
</body>
</html>
