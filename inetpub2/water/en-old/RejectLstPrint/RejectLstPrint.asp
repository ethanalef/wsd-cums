<!-- #include file="../conn.asp" -->

<%

server.scripttimeout = 1800
mndate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
SQl = "SELECT  a.memno,a.adate,sum(a.bankin) as unpaid ,b.memname,b.memcname,b.memaddr1,b.memaddr2,b.memaddr3,b.memcontacttel,b.accode  FROM  autopay a ,memmaster b where a.memno=b.memno and a.flag='F' and right(a.code,1)='1' and a.pflag=1 group by a.memno,a.adate,b.memname,b.memcname,b.memcname,b.memaddr1,b.memaddr2,b.memaddr3,b.memcontacttel,b.accode  order by a.memno,a.adate,b.memname,b.memcname,b.memaddr1,b.memaddr2,b.memaddr3,b.memcontacttel,b.accode  "
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn
if rs.eof then
   response.redirect "rejectlst.asp"
end if
dim guarantor(3)
dim gender(3)
if request.form("output")="word" then
	Response.ContentType = "application/msword"
elseif request.form("output")="excel" then
	Response.ContentType = "application/vnd.ms-excel"
end if
%>
<html>
<head>
<title>銀行轉賬失效通知書列表</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="0" topmargin="0" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0" >
	<tr height="30" valign="top" align="center"><td colspan="15"><font size="4"  face="標楷體" >水務署員工儲蓄互助社<br>銀行轉賬失效通知書列表<br><font size="2"  face="標楷體" >日期 : <%=mndate%><br></font></font></td></tr>
        

</table>
<table border="0" cellpadding="0" cellspacing="0">
	<tr height="15" valign="bottom">
                <td width=70 align="center"><font size="2"  face="標楷體" >社員編號</font></td>
		<td width=140 align="center"><font size="2"  face="標楷體" >英文名稱</font></td>
		<td width=80 align="center"><font size="2"  face="標楷體" >中文名稱</font></td>		
                <td width="100" align="center"><font size="2"  face="標楷體" >銀行轉賬金額</font></td>
             	<td width="100" align="center"><font size="2"  face="標楷體" >聯絡電話</font></td>
		<td width="500" align="center"><font size="2"  face="標楷體" >地址</font></td>
               
	</tr>
	<tr><td colspan=6><hr></td></tr>
<%


        do while not rs.eof

           
            
%>
	<tr>
		<td width=70 align="center"><font size="2"  face="標楷體" ><%=rs("memNo")%></font></td>
                <td width=140 align="center"><font size="2"  face="標楷體" ><%=rs("memname")%></font></td>
		<td width=80 align="center"><font size="2"  face="標楷體" ><%=rs("memcname")%></font></td>
                <td width="100" align="right"><%=formatnumber(rs("unpaid"),2)%></td>
		<td width="80" align="right"><%=rs("memcontacttel")%></td>
		<td width="500" align="left"><%=rs("memaddr1")%><br><%=rs("memaddr2")%><br><%=rs("memaddr3")%></td>
               
	</tr>



<%
  RS.MOVENEXT
  LOOP
%> 
</font>
</table>
</body>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
