<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
mMonth = request("mMonth")



if IsNumeric(mMonth) then
	if int(mMonth)<1 or int(mMonth)>12 then
		response.redirect "birthdayList.asp"
	end if
else
	response.redirect "birthdayList.asp"
end if

SQl = "select memNo,memName,memcname,memBday from memMaster where wdate is null  and month(memBday)="&mMonth&" order by memBday"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn

if request.form("output")="word" then
	Response.ContentType = "application/msword"
elseif request.form("output")="excel" then
	Response.ContentType = "application/vnd.ms-excel"
elseif request.form("output")="text" then
	spaces=""
	for idx = 1 to 50
		spaces=spaces&" "
	next
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.CreateTextFile(Server.MapPath("..\txt")&"\"&session("username")&".txt", True)
	objFile.Write "水務署員工儲蓄互助社"
	objFile.WriteLine ""
	objFile.Write monthname(mMonth)&"份 社員生日名單列表"
	objFile.WriteLine ""
	objFile.Write left("社員編號"&spaces,10)
	objFile.Write left("社員名稱"&spaces,50)
	objFile.Write left("出生日期"&spaces,20)
	objFile.WriteLine ""
	for idx = 1 to 80
		objFile.Write "-"
	next
	objFile.WriteLine ""
	do while not rs.eof
		objFile.Write left(rs("memNo")&spaces,10)
		objFile.Write left(rs("memName")&rs("memcname")&spaces,50)
		objFile.Write left(right("0"&day(rs("memBday")),2)&"/"&right("0"&month(rs("memBday")),2)&"/"&year(rs("memBday"))&spaces,20)
		objFile.WriteLine ""
		rs.movenext
	loop
	objFile.Close
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.redirect "../txt/"&session("username")&".txt"
end if
%>
<html>
<head>
<title>Birthday List</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0">
	<tr height="30" valign="top" align="center">
		<td colspan="3"><font size="4">水務署員工儲蓄互助社</font></td>
	</tr>
	<tr height="35" valign="top" align="center">
		<td colspan="3"><font size="4"><%=monthname(mMonth)%>份 社員生日名單列表</font></td>
	</tr>
	<tr height="15" valign="bottom">
		<td width=70><b>社員編號</b></td>
		<td width=270><b>社員名稱</b></td>
		<td><b>出生日期</b></td>
	</tr>
	<tr><td colspan=3><hr></td></tr>
<%
do while not rs.eof
%>
	<tr>
		<td><%=rs("memNo")%></td>
		<td><%=rs("memName")%><%=rs("memcname")%></td>
		<td><%=right("0"&day(rs("memBday")),2)&"/"&right("0"&month(rs("memBday")),2)&"/"&year(rs("memBday"))%></td>
	</tr>
<%
	rs.movenext
loop
%>
</table>
</center>
</body>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
