<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="navigator.asp" -->
<%
if request("del")<>"" then
	del = request("del")
	conn.execute("update meetingNotes set deleted=-1 where uid="&del)
	msg = "�|ĳ���� "&del&" �w�R��"
end if

sql = "select * from meetingNotes where deleted=0 order by uid desc"
set rs = server.createobject("ADODB.Recordset")
rs.open sql, conn, 3

'if rs.recordcount=1 then
'	response.redirect "meetingNotesDetail.asp?id="&rs("uid")
'end if

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
<title>�|ĳ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0">
<!-- #include file="menu.asp" -->
<br>
<center>
<%if recordcount>pagesize then navigator("meetingNotes.asp") end if%>
<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="336699">
	<tr bgcolor="#330000" align="center">
		<td><font size="2" color="#FFFFFF">���</font></td>
		<td><font size="2" color="#FFFFFF">�|ĳ</font></td>
<%if session("userLevel")<>5 then%>
		<td bgcolor="#FFFFFF"><a href="meetingNotesDetail.asp"><font size="2">�s�W</font></a></td>
<%end if%>
	</tr>
<%
do while not rs.eof and rowcount < pagesize
	rowcount = rowcount + 1
%>
  <tr bgcolor="#FFFFFF">
	<td><a href="meetingNotesDetail.asp?id=<%=rs("uid")%>"><font size="2"><%=right("0"&day(rs("rpDate")),2)&"/"&right("0"&month(rs("rpDate")),2)&"/"&year(rs("rpDate"))%></font></a></td>
	<td align="right"><font size="2"><%=rs("rpNo")%></font></td>
<%if session("userLevel")<>5 then%>
	<td><a href="meetingNotes.asp?del=<%=rs("uid")%>" onclick="return confirm('�R��������?')"><font size="2">�R��</font></a></td>
<%end if%>
  </tr>
<%
	rs.movenext
loop
%>
</table>
<%if recordcount>pagesize then navigator("meetingNotes.asp") end if%>
</center>
</body>
</html>
