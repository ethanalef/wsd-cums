<!-- #include file="../conn.asp" -->
<!-- #include file="navigator.asp" -->
<%
If session("username") = "" then
	response.write "<script>window.close()</script>"
	response.end
end if

key = request("key")
if key<>"" then
	stylefield = "and a.memNo like '" & key & "%'"
end if

SQL = "select a.memNo,a.memname,b.code,b.date,b.amount from memmaster a,share b where a.memNo=b.memno " & stylefield &" order by date,code DESC " 
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn, 3

if rs.recordcount=1 then
%>
<script language="JavaScript">
<!--
	window.opener.document.form1.memName.value = '<% =rs("memName") %>';
	window.opener.document.all.tags( "td" )['memName'].innerHTML='<% =rs("memName") %>';
	window.opener.document.form1.sdate.value = '<% =rs("date") %>';
	window.opener.document.all.tags( "td" )['sdate'].innerHTML='<% =rs("date") %>';
	window.opener.document.form1.code.value = '<% =rs("code") %>';
	window.opener.document.all.tags( "td" )['code'].innerHTML='<% =rs("code") %>';
	window.opener.document.form1.amount.value = '<% =rs("amount") %>';
	window.opener.document.all.tags( "td" )['amount'].innerHTML='<% =rs("amount") %>';

	window.opener.document.form1.amount.focus()
	window.close();
//-->
</script>
<%
	response.end
end if

if not rs.eof then
	if request("page") <> "" then
		pageno = cint(request("page"))
	else
		pageno = 1
	end if
	rs.pagesize = 15
	pagesize=rs.pagesize
	rs.absolutepage = pageno
	recordcount=rs.recordcount
	pagecount = rs.pagecount
end if
%>
<html>
<head>
<meta1 HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<title>Member Popup</title>
<link href="../main.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
<!--
function PostCode(memNo,memName,code,date,amount)
{
	window.opener.document.form1.memNo.value=memNo;
	window.opener.document.form1.memName.value=memName;
	window.opener.document.all.tags( "td" )['memName'].innerHTML=memName;
	window.opener.document.form1.code.value=code;
	window.opener.document.all.tags( "td" )['code'].innerHTML=code;
	window.opener.document.form1.sdate.value=date;
	window.opener.document.all.tags( "td" )['sdate'].innerHTML=date;
	window.opener.document.form1.amount.value=amount;
	window.opener.document.all.tags( "td" )['amount'].innerHTML=amount;
	window.opener.document.form1.amount.focus()
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
            <font face="arial" style="color: black; font: bold" size="3">搜尋社員</font>&nbsp;&nbsp;
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
				<td bgcolor="#ffffff"><a href="JavaScript:PostCode('<%=rs(0)%>','<%=rs(1)%>','<%=rs(2)%>','<%=right("0"&day(rs(3)),2)&"/"&right("0"&month(rs(3)),2)&"/"&year(rs(3))%>','<%=rs(4)%>')"><% =rs(0) %></a></td>
				<td bgcolor="#ffffff"><a href="JavaScript:PostCode('<%=rs(0)%>','<%=rs(1)%>','<%=rs(2)%>','<%=right("0"&day(rs(3)),2)&"/"&right("0"&month(rs(3)),2)&"/"&year(rs(3))%>','<%=rs(4)%>')"><% =rs(1) %></a></td>

				<td bgcolor="#ffffff"><a href="JavaScript:PostCode('<%=rs(0)%>','<%=rs(1)%>','<%=rs(2)%>','<%=right("0"&day(rs(3)),2)&"/"&right("0"&month(rs(3)),2)&"/"&year(rs(3))%>','<%=rs(4)%>')"><% =rs(2) %></a></td>
				<td bgcolor="#ffffff"><a href="JavaScript:PostCode('<%=rs(0)%>','<%=rs(1)%>','<%=rs(2)%>','<%=right("0"&day(rs(3)),2)&"/"&right("0"&month(rs(3)),2)&"/"&year(rs(3))%>','<%=rs(4)%>')"><% =rs(3) %></a></td>

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