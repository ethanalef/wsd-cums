<!-- #include file="../conn.asp" -->
<!-- #include file="navigator.asp" -->
<%
If session("username") = "" then
	response.write "<script>window.close()</script>"
	response.end
end if

key = request("key")
if key<>"" then
	stylefield = "and memNo like '" & key & "%'"
end if

SQL = "select memNo,memName,memGrade,employCond,firstAppointDate,memBday  from memMaster where deleted=0 " & stylefield & " order by memNo"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn, 3

   ttlbal = 0
   set rs1=conn.execute("select * from share where memno='"&rs(0)&"' ")
   do while  not rs1.eof
      select case left(rs1("code"),1)
             case "A","0","C"
                  ttlbal = ttlbal + rs1("amount")
             case "G","H","B"
                  ttlbal = ttlbal - rs1("amount")
      end select
      rs1.movenext
      loop
   rs1.close
   set rs1=nothing     

if rs.recordcount=1 then
    
         
%>
<script language="JavaScript">
<!--
	window.opener.document.form1.memName.value = '<% =rs("memName") %>';
	window.opener.document.all.tags( "td" )['memName'].innerHTML='<% =rs("memName") %>';
        window.opener.document.all.tags( "td" )['ttlbal'].innerHTML='<% =ttlbal %>';


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
	rs.pagesize = 20
	pagesize=rs.pagesize
	rs.absolutepage = pageno
	recordcount=rs.recordcount
	pagecount = rs.pagecount
end if
%>
<html>
<head>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<title>Member Popup</title>
<link href="../main.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
<!--
function PostCode(memNo,memName,memGrade,employCond,firstAppointDate,memBday,ttlbal)
{
	window.opener.document.form1.memNo.value=memNo;
	window.opener.document.form1.memName.value=memName;
	window.opener.document.all.tags( "td" )['memName'].innerHTML=memName;
	window.opener.document.all.tags( "td" )['ttlbal'].innerHTML=ttlbal;



	window.opener.document.form1.sdate.focus()
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
			&nbsp;&nbsp;<b>�j�M</b>
			<input type="text" name="key" size="20" maxlength="50">
			<input type="submit" value="�T�w" class="sbttn" name="send">
        </td>
        <td align="right">
            <font face="arial" style="color: black; font: bold" size="3">�j�M����</font>&nbsp;&nbsp;
        </td>
    </tr>
	</form>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<%
if key <> "" then
%>
	<tr>
		<td align="center">"<% =key %>" ���j�M���G</td>
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
				<td bgcolor="#ffffff"><a href="JavaScript:PostCode('<%=rs(0)%>','<%=rs(1)%>','<%=rs(2)%>','<%=rs(3)%>','<%=right("0"&day(rs(4)),2)&"/"&right("0"&month(rs(4)),2)&"/"&year(rs(4))%>','<%=right("0"&day(rs(5)),2)&"/"&right("0"&month(rs(5)),2)&"/"&year(rs(5))%>','<%=ttlbal%>')"><% =rs("memNo") %></a></td>
				<td bgcolor="#ffffff"><a href="JavaScript:PostCode('<%=rs(0)%>','<%=rs(1)%>','<%=rs(2)%>','<%=rs(3)%>','<%=right("0"&day(rs(4)),2)&"/"&right("0"&month(rs(4)),2)&"/"&year(rs(4))%>','<%=right("0"&day(rs(5)),2)&"/"&right("0"&month(rs(5)),2)&"/"&year(rs(5))%>','<%=ttlbal%>')"><% =rs("memName") %></a></td>
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
		<br><p align="center"><font size="4">�L�ŦX����</font></p>
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