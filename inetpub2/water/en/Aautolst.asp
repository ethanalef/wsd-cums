<!-- #include file="../conn.asp" -->

<%

server.scripttimeout = 1800

SQl = "select * from memmaster where mstatus='A' "
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn,1,1

ttlamt = 0
ttlsamt = 0
ttlpamt = 0
ttlpint = 0
ttlisamt = 0
ttlipamt = 0
ttlipint = 0


%>
<html>
<head>
<title>銀行轉賬細明表</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0">
	<tr height="30" valign="top" align="center">
	<td colspan="15"><font size="4">水務署員工儲蓄互助社</font></td>
	</tr>
	<tr height="35" valign="top" align="center">
	<td colspan="15"><font size="4">銀行轉賬細明表</font></td>
	</tr>
	<tr height="15" valign="bottom">
	<td width="80"><b>社員編號</b></td>
	<td width="200"><b>社員名稱</b></td>
	<td width="60" align="right"><b>(利息) </b></td>
	<td width="60" align="right"><b>(本金)</b></td>
	<td width="60" align="right"><b>(股金)</b></td>
	<td width="80" align="right"><b>(脫期利息) </b></td>
	<td width="80" align="right"><b>(脫期本金)</b></td>
	<td width="80" align="right"><b>(脫期股金)</b></td>
	<td width="80" align="right"><b>(總金額)</b></td>
	</tr>
	<tr><td colspan=9><hr></td></tr>
<%
do while not rs.eof

           memno=rs("memno") 
           ipamt = 0 : ipint =0 : isamt = 0
           pamt = 0 : pint = 0 : samt = 0
            
           sttlamt = 0   
           set rs1 = Server.CreateObject("ADODB.Recordset")
           sql2 = "select * from autopay where memno='"&memno&"' "
           rs1.open sql2, conn,2,2
           pint = 0 : pamt=0 : samt = 0  
           do while not rs1.eof
              select case rs1("code")
                     case "E1"
                          if rs1("flag")<>"F"  then
                           pamt = rs1("bankin")
                           ttlpamt = ttlpamt + pamt
                         else
                           ipamt = rs1("bankin")
                            ttlipamt = ttlipamt + ipamt 
                         end if 
                     case "F1"
                          if rs1("flag")<>"F" then
                           pint = rs1("bankin")
                           ttlpint = ttlpint + pint
                          else
                           ipint = rs1("bankin")
			    ttlipint = ttlipint + ipint
                          end if  
                     case "A1"
                          if rs1("flag")<>"F" then
                           samt = rs1("bankin")
                           ttlsamt = ttlsamt + samt
                          else

                           isamt = rs1("bankin")                   
                           ttlisamt = ttlisamt + isamt
                          end if 

                 
                          
               end select 
               sttlamt = sttlamt + rs1("bankin")
               rs1.movenext
               loop
               rs1.close
               if sttlamt > 0 then 
%>
	<tr>
		<td><%=rs("memNo")%></td>
		<td><%=rs("memName")%><%=rs("memcname")%></td>              		
		<td width=80 align="right"><%=formatNumber(pint,2)%></td>
		<td width=80 align="right"><%=formatNumber(pamt,2)%></td>
		<td width=80 align="right"><%=formatNumber(samt,2)%></td>
		<td width=80 align="right"><%=formatNumber(ipint,2)%></td>
		<td width=80 align="right"><%=formatNumber(ipamt,2)%></td>
		<td width=80 align="right"><%=formatNumber(isamt,2)%></td>
		<td width=80 align="right"><%=formatNumber(sttlamt,2)%></td>

	</tr>
<%
        end if    
	ttlTemp=ttlTemp+sttlamt
	rs.movenext
loop
%>
	<tr><td colspan=9><hr></td></tr>
	<tr>
		<td colspan="8"></td>
		
	
		<td align="right"><%=formatNumber(ttlTemp,2)%></td>
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
