<!-- #include file="../conn.asp" -->
<!-- #include file="navigator.asp" -->
<%
If session("username") = "" then
	response.write "<script>window.close()</script>"
	response.end
end if

editNum = request("editNum")
key = request("key")
if key<>"" then
	stylefield = "and glId like '" & key & "%'"
end if

SQL = "select glId,glName from glMaster where deleted=0 " & stylefield & " order by glId"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn, 3
if rs.eof then
	response.write "<script>window.close()</script>"
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
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<title>A/c Popup</title>
<link href="../main.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
<!--
function PostCode(glId,glName)
{
	window.opener.document.form1.bd3c<%=editNum%>.value = glId;
	window.opener.document.form1.bd4c<%=editNum%>.value = glName;
	window.opener.document.form1.bd5c<%=editNum%>.focus();
	window.close();
}
//-->
</script>
</head>
<body leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<div align="center">
<center>
<table border="0" cellpadding="0" cellspacing="0" width="100%" bgcolor="#87CEEB">
	<form method="POST" action="<%=request.servervariables("script_name")%>" name="form">
    <tr>
        <td height="25">
			&nbsp;&nbsp;<b>搜尋</b>
			<input type="text" name="key" size="20" maxlength="50">
			<input type="submit" value="確定" class="sbttn" name="send">
        </td>
        <td align="right">
            <font face="arial" style="color: black; font: bold" size="3">搜尋總賬</font>&nbsp;&nbsp;
        </td>
    </tr>
	</form>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<%
if key <> "" then
%>
	<tr>
		<td align="center">"<% =key %>" 的搜尋結果</td>
	</tr>
<%
end if
%>
	<tr>
		<td>
<%
if not rs.eof then
	navigator(request.servervariables("script_name")&"?key="&key)
%>
		<table border="0" cellpadding="3" cellspacing="1" width="100%" bgcolor="#336699">
<%
	do while not rs.eof and rowcount < rs.pagesize
		rowcount = rowcount + 1
%>
			<tr>
				<td bgcolor="#ffffff"><a href="JavaScript:PostCode('<%=rs(0)%>','<%=rs(1)%>')"><% =rs("glId") %></a></td>
				<td bgcolor="#ffffff"><a href="JavaScript:PostCode('<%=rs(0)%>','<%=rs(1)%>')"><% =rs("glName") %></a></td>
			</tr>
<%
		rs.movenext
	loop
%>
		</table>
<%
	navigator(request.servervariables("script_name")&"?key="&key)
	response.write "<p>"
else
%>
		<br><p align="center"><font size="4">無符合紀錄</font></p>
<%
end if
%>
		</td>
	</tr>
</table>
</center>
</div>
</body>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>