<!-- #include file="../conn.asp" -->

<%

server.scripttimeout = 1800

SQl = "select A.*,b.memname,b.memcname  from autopay a,memmaster b where a.memno=b.memno and right(a.code,1)='2'  order by a.memno   "
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn
memno=rs("memno")
memname=rs("memname")
memcname=rs("memcname")
yy=year(rs("adate"))
mm=month(rs("adate"))
period =yy&"/"&right("0"&mm,2)
sttlamt = 0   
ttlamt = 0
ttlsamt = 0
TTLASAMT = 0
ttlpamt = 0
ttlpint = 0
ttlisamt = 0
ttlipamt = 0
ttlipint = 0
memcnt = 0
mndate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())




pint = 0 : pamt=0 : samt = 0  
 ipamt = 0 : ipint =0 : isamt = 0
if request.form("output")="Word" then
	Response.ContentType = "application/msword"
elseif request.form("output")="excel" then
	Response.ContentType = "application/vnd.ms-excel"

end if


%>
<html>
<head>
<title>�w�����ө���</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table width="800" border="0">

 <tr>
    <td width="250">&nbsp;</td>
    <td width="400">&nbsp;</td>
    <td width="150">&nbsp;</td>
  </tr>
	<tr>
         <td align="left"><font size="2"  face="�з���" >�@��� : <%=mndate%></font></td>
       
        <td align="center"><b><font size="4"  face="�з���" >���ȸp���u�x�W���U��<br>�w�����ө���</font></b?</td>
       
         <td>&nbsp</td> 
        </tr>
        <tr>
         <td><b><font   face="�з���" >�@�w��������G<%=period%></font></td>
         <td>&nbsp</td>
         <td>&nbsp</td>
       </tr>                                   

</table>
<br>
<table border="0" cellpadding="0" cellspacing="0">

	<tr height="15" valign="bottom">
	<td width="80"><b><font   face="�з���" >�����s��</font></b></td>
	<td width="60"><b><font   face="�з���" >�����W��</font></b></td>
	<td width="100" align="center"><b><font   face="�з���" >(�Q��)</font></b></td>
	<td width="100" align="center"><b><font   face="�з���" >(����)</font></b></td>
	<td width="100" align="center"><b><font   face="�з���" >(�Ѫ�)</font></b></td>
	<td width="100" align="center"><b><font   face="�з���" >(����Q��)</font></b></td>
	<td width="100" align="center"><b><font   face="�з���" >(�������)</font></b></td>
	<td width="100" align="center"><b><font   face="�з���" >(����Ѫ�)</font></b></td>
	<td width="100 " align="center"><b><font   face="�з���" >(�`���B)</font></b></td>
	</tr>
	<tr><td colspan=9><hr></td></tr>
<%
do while not rs.eof
           if memno<>rs("memno") then       
               if sttlamt > 0 then
                   memcnt = memcnt + 1
%>
	<tr>
		<td width=40><%=memNo%></td>
		<td width=40><%=memcname%></td>              		
		<td width=100 align="right"><%=formatNumber(pint,2)%></td>
		<td width=100 align="right"><%=formatNumber(pamt,2)%></td>
		<td width=100 align="right"><%=formatNumber(samt,2)%></td>
		<td width=100 align="right"><%=formatNumber(ipint,2)%></td>
		<td width=100 align="right"><%=formatNumber(ipamt,2)%></td>
		<td width=100 align="right"><%=formatNumber(isamt,2)%></td>
		<td width=100 align="right"><%=formatNumber(sttlamt,2)%></td>

	</tr>
<%
                 ipint=0
                 ipamt=0
                 isamt=0
                 pint=0
                 pamt=0
                 samt=0
                 sttlamt = 0
                end if 
             memno=rs("memno")
             memname=rs("memname")
             memcname=rs("memcname")
         end if
              select case rs("code")
                     case "E2"
                          if rs("flag")<>"F"  then
                           pamt = rs("bankin")
                           ttlpamt = ttlpamt + pamt
                            ttlASAMT=ttlASamt  + rs("bankin")
                         else
                           ipamt = rs("bankin")
                            ttlipamt = ttlipamt + ipamt 
                         end if 
                     case "F2"
                          if rs("flag")<>"F" then
                           pint = rs("bankin")
                           ttlpint = ttlpint + pint
                            ttlASAMT=ttlASamt  + rs("bankin")
                          else
                           ipint = rs("bankin")
			    ttlipint = ttlipint + ipint
                          end if  
                     case "A2"
                          if rs("flag")<>"F" then
                           samt = rs("bankin")
                           ttlsamt = ttlsamt + samt
                           ttlASAMT=ttlASamt  + rs("bankin")
                          else
                          
                           isamt = rs("bankin")                   
                           ttlisamt = ttlisamt + isamt
                          end if 

                 
                          
               end select 
               sttlamt = sttlamt + rs("bankin")
               ttlTemp=ttlTemp+rs("bankin")

	      
	rs.movenext
loop

               if sttlamt > 0 then
                   memcnt = memcnt + 1
%>
	<tr>
		<td><%=memNo%></td>
		<td width=40><%=memcname%></td>              		
		<td width=100 align="right"><%=formatNumber(pint,2)%></td>
		<td width=100 align="right"><%=formatNumber(pamt,2)%></td>
		<td width=100 align="right"><%=formatNumber(samt,2)%></td>
		<td width=100 align="right"><%=formatNumber(ipint,2)%></td>
		<td width=100 align="right"><%=formatNumber(ipamt,2)%></td>
		<td width=100 align="right"><%=formatNumber(isamt,2)%></td>
		<td width=100 align="right"><%=formatNumber(sttlamt,2)%></td>

	</tr>
<%
 
                end if 
 
%>
	<tr><td colspan=9><hr></td></tr>
	<tr>
		<td>�X�@</td>
		<td><%=memcnt%>�H</td>
	        <td align="right"><%=formatNumber(ttlpint,2)%></td>
                <td align="right"><%=formatNumber(ttlpamt,2)%></td>
                <td align="right"><%=formatNumber(ttlsamt,2)%></td>
                <td align="right"><%=formatNumber(ttlipint,2)%></td>
                <td align="right"><%=formatNumber(ttlipamt,2)%></td>
                <td align="right"><%=formatNumber(ttlisamt,2)%></td>
		<td align="right"><%=formatNumber(ttlTemp,2)%></td>
	</tr>
	<tr>
		<td></td>
		<td>=====</td>
	        <td align="right">==========</td>
                <td align="right">==========</td>
                <td align="right">==========</td>
                <td align="right">==========</td>
                <td align="right">==========</td>
                <td align="right">==========</td>
		<td align="right">==========</td>
	</tr>
</table>
</center>
<table border="0" cellpadding="0" cellspacing="0">
	<br>
        <tr>
           <td width="30"></td>
           <td width="100" align="right"><b>�Ȧ���b�X�@ :</b></td> 
           <td><%=formatNumber(ttlasamt,2)%></td>
          
       </tr>
       
</table>

</body>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>