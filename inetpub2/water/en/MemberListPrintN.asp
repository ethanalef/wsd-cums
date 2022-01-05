<!-- #include file="../conn.asp" -->

<%
mMonth = request("mMonth")
 mndate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())


if IsNumeric(mMonth) then
	if int(mMonth)<1 or int(mMonth)>12 then
		response.redirect "birthdayList.asp"
	end if
else
	response.redirect "birthdayList.asp"
end if


SQl =   "select m.memno, m.memname, m.memcname, m.memhkid , " & _
        "	ISNULL(m.memaddr1, '') AS 'memaddr1', ISNULL(m.memaddr2, '') AS 'memaddr2', " & _
        "	ISNULL(m.memaddr3, '') AS 'memaddr3' " & _
        "	 ,b.shttl " & _
        "from memmaster m "&_
        " right join ( select memno , sum(amount) as shttl from share group by memno ) b on m.memno=b.memno "&_
        "WHERE b.shttl > 0 " &_
        "AND m.mstatus NOT IN ('D', 'V','C','B','P','8','9') " &_
        "AND m.wdate is null " &_
        "order by m.memno "



Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn

if request.form("output")="word" then
	Response.ContentType = "application/msword"
elseif request.form("output")="excel" then
	Response.ContentType = "application/vnd.ms-excel"

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
<font size="4"  face="�з���" >
<table border="0" cellpadding="0" cellspacing="0" >
	<tr height="30" valign="top" align="center"><td colspan="15"><font size="4"  face="�з���" >���ȸp���u�x�W���U��<br>
<%=monthname(mMonth)%>�� �����ͤ�W��C��<br><font size="2"  face="�з���" >��� : <%=mndate%></font></font></td></tr>
        

</table>
<br>
<br>

              
<table border="0" cellpadding="0" cellspacing="0"  width="1150">

	<tr height="15" valign="bottom">
		<td width=70 align="center"><font size="3"  face="�з���" >�����s��</font></td>
                <td width=150 align="center"><font size="3"  face="�з���" >�^��W��</font></td> 
		
		<td width=70 align="center"><font size="3"  face="�з���" >����m�W</font></td>
                <td width=70 align="center"><font size="3"  face="�з���" >�W�����Ҹ��X</font></td>
		<td width=100 align="center"><font size="3"  face="�з���" >�Ѫ����l</font></td>
                <td width=230 align="center"></td>
		<td width=230 align="center"><font size="3"  face="�з���" >�a�}</font></center></td>						
                <td width=230 align="center"></td>
	</tr>
	<tr><td colspan=10><hr></td></tr>
<%
do while not rs.eof
               
                            
%>
	<tr>
		<td width=70 align="center"><%=rs("memNo")%></td>
                <td width=150 align="center"><%=rs("memName")%></td>
		
		<td width=70 align="center"><font size="3"  face="�з���" ><%=rs("memcname")%></font></td>		
		<td width=70 align="center"><%=rs("memhkid")%></td>
		<td width=100 align="center"><font size="2"  face="�з���" ><%=rs("shttl")%></font></td>
             
                <td width=230 align="left"><font size="2"  face="�з���" ><%=rs("memaddr1")%></font></td>
                <td width=230 align="left"><font size="2"  face="�з���" ><%=rs("memaddr2")%></td>		
                <td width=230 align="left"><font size="2"  face="�з���" ><%=rs("memaddr3")%></font></td>

	</tr>
<%
	rs.movenext
loop
%>
</table>
</font>
</center>
</body>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
