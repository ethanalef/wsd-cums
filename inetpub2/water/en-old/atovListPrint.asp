<!-- #include file="../conn.asp" -->

<%
SQl = "SELECT  memno, SUM(bankin) AS Expr1  FROM  autopay where right(code,1)='1' and flag<>'F' group by memno order by memno "
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn,1,1

ttlcnt=0 


if request.form("output")="word" then
	Response.ContentType = "application/msword"
elseif request.form("output")="excel" then
	Response.ContentType = "application/vnd.ms-excel"

end if
%>
<html>
<head>
<title>�Ȧ���b�W�B�ө���</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0">
	<tr height="30" valign="top" align="center">
	<td colspan="15"><font size="4">���ȸp���u�x�W���U��</font></td>
	</tr>
	<tr height="35" valign="top" align="center">
	<td colspan="15"><font size="4">�Ȧ���b�W�B�ө���</font></td>
	</tr>
	<tr height="15" valign="bottom">
	<td width="80"><b>�����s��</b></td>
	<td width="200"><b>�����W��</b></td>
	<td width="130" align="right"><b>(��b���B)</b></td>
	<td width="130" align="right"><b>(��b�W��)</b></td>
	</tr>
	<tr><td colspan=4><hr></td></tr>
<%
do while not rs.eof
   memno= rs(0) 
   set rs1=conn.execute("select memname  ,memcname,bnklmt from memmaster where memno='"&memno&"' and (mstatus='A' or mstatus='0' or mstatus='1' or mstatus='2' or mstatus='3' ) ") 
           if not rs1.eof then 
              if (rs(1)-rs1(2))> rs1(2) and rs1(2)> 0 then
%>
	<tr>
		<td><%=rs("memNo")%></td>
		<td><%=rs1("memName")%><%=rs1("memcname")%></td>
		<td align="right"><%=formatnumber(rs(1),2)%></td>
		<td align="right"><%=formatNumber(rs1(2),2)%></td>
	</tr>
<%
      
	ttlcnt=ttlcnt+1
        end if
        end if  
	rs.movenext
loop
%>
	<tr><td colspan=4><hr></td></tr>

	
	<tr>
		
		<td>Total Count :</td>
		<td align="right"><%=formatNumber(ttlcnt,2)%></td>
		<td colspan="2"></td>
		
	</tr>

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
