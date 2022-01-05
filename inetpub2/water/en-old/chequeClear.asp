<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
if ucase(request("From")) = ucase(Request.ServerVariables("script_name")) then
    if lcase(request("send")) <> "" then
        A = split(request("TS"),",",-1,1)
        conn.begintrans
        if isarray(A) then
            if (ubound(A) >= 0) then
                for i = 0 to ubound(A)
                    conn.execute("update cheque set chequeClear=-1 where uid=" & A(i))
                next
            end if
        end if
        conn.committrans
    end if
    conn.close
    set conn = nothing
    response.redirect "cheque.asp"
end if

A = split(request("TS"),",",-1,1)
if isarray(A) then
    if (ubound(A) < 0) then
        response.redirect "cheque.asp"
    end if
else
    response.redirect "cheque.asp"
end if
%>
<html>
<head>
<title>支票對數</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0">
<!-- #include file="menu.asp" -->
<form method="POST" action="chequeClear.asp" name="form1">
<INPUT type="hidden" name="TS" value="<% =request("TS") %>">
<INPUT type="hidden" name="From" value="<% =Request.ServerVariables("script_name") %>">
<center>
<font color=#FF0000>確定為以下支票對數 ?</font>
<br>
<br>
<%
if isarray(A) then
    if ubound(A) >= 0 then
%>
<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="336699">
  <tr bgcolor="#330000" align="center">
	<td><font color="#FFFFFF">#</font></td>
	<td><font color="#FFFFFF">支票號碼</font></td>
	<td><font color="#FFFFFF">支票日期</font></td>
	<td><font color="#FFFFFF">收款人</font></td>
	<td><font color="#FFFFFF">金額</font></td>
  </tr>
<%
        for i = 0 to ubound(a)
            SQL = "select * from cheque where uid=" & A(i)
            Set q = Server.CreateObject("ADODB.Recordset")
            q.open sql, conn
            if not q.eof then
%>
    <tr align="center" bgcolor="#ffffff">
        <td><% =i + 1 %></td>
        <td><%=q("chequeNum")%></td>
		<td><%=right("0"&day(q("chequeDate")),2)&"/"&right("0"&month(q("chequeDate")),2)&"/"&year(q("chequeDate"))%></td>
		<td><%=q("payee")%></td>
		<td><%=formatnumber(q("amount"),2)%></td>
    </tr>
<%
            end if
            Q.close
            set q = nothing
        next
%>
    <tr>
        <td colspan="5" bgcolor="#ffffff" align="right">
            <input type="submit" name="send" value="是" class="sbttn"> <input type="submit" name="send" value="否" class="sbttn">
        </td>
    </tr>
</table>
<%
        conn.close
        set conn = nothing
    end if
end if
%>
</center></div>
</form>
</body>
</html>
