<!-- #include file="../conn.asp" -->

<%

server.scripttimeout = 1800

SQl = "select A.*,b.memname,b.memcname  from autopay a,memmaster b where a.memno=b.memno and a.status='F'  order by a.memno   "
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn
if not rs.eof then
memno=rs("memno")
memnname=rs("memname")
memcname=rs("memcname")
end if

sttlamt = 0   
ttlamt = 0
ttlsamt = 0
TTLASAMT = 0
ttlpamt = 0
ttlpint = 0
ttlisamt = 0
ttlipamt = 0
ttlipint = 0

mndate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())




pint = 0 : pamt=0 : samt = 0  
 ipamt = 0 : ipint =0 : isamt = 0
if request.form("output")="Word" then
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
	objFile.Write "特別個案轉帳細明表"        
	objFile.WriteLine ""	
        objFile.Write "日期 : "&mndate
	objFile.WriteLine ""	
	objFile.Write " 社員編號 "
	objFile.Write "         社員名稱                 "
	objFile.Write "  (利息)   "
	objFile.Write "    (本金)   "
	objFile.Write "   (股金)   "
	objFile.Write " ( 脫期利息 ) "
	objFile.Write " ( 脫期本金 ) "
	objFile.Write " ( 脫期股金 ) "
	objFile.Write " (總金額)  "
	objFile.WriteLine ""
	for idx = 1 to 130
		objFile.Write "-"
	next
	objFile.WriteLine ""
        memno=rs("memno")
        memname=rs("memname")
        memcname=rs("memcname")
	do while not rs.eof
          
           if memno<>rs("memno")  then  
               sttlamt = pamt + pint +samt + ipamt + ipint +isamt
               if sttlamt > 0 then
		objFile.Write left(" "&memNo&spaces,10)
		objFile.Write left(memName&memcname&spaces,25)
		objFile.Write right(spaces&formatnumber(pint,2),13)
		objFile.Write right(spaces&formatnumber(pamt,2),13)
		objFile.Write right(spaces&formatnumber(samt,2),13)
		
		objFile.Write right(spaces&formatnumber(ipint,2),13)
		objFile.Write right(spaces&formatnumber(ipamt,2),13)
		objFile.Write right(spaces&formatnumber(isamt,2),13)
                objFile.Write right(spaces&formatnumber(sttlamt,2),15)
		objFile.WriteLine ""
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
                     case "E2","E1"
                          if rs("flag")<>"F"  then
                           pamt = rs("bankin")
                           ttlpamt = ttlpamt + pamt
                            ttlASAMT=ttlASamt  + rs("bankin")
                         else
                           ipamt = rs("bankin")
                            ttlipamt = ttlipamt + ipamt 
                         end if 
                     case "F2","F1"
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
             
               ttlTemp=ttlTemp+rs("bankin")

 
            
                                                          
		rs.movenext
	loop
               if sttlamt > 0 then
		objFile.Write left(" "&memNo&spaces,10)
		objFile.Write left(memName&memcname&spaces,25)
		objFile.Write right(spaces&formatnumber(pint,2),13)
		objFile.Write right(spaces&formatnumber(pamt,2),13)
		objFile.Write right(spaces&formatnumber(samt,2),13)
		
		objFile.Write right(spaces&formatnumber(ipint,2),13)
		objFile.Write right(spaces&formatnumber(ipamt,2),13)
		objFile.Write right(spaces&formatnumber(isamt,2),13)
                objFile.Write right(spaces&formatnumber(sttlamt,2),15)
		objFile.WriteLine ""
                 ipint=0
                 ipamt=0
                 isamt=0
                 pint=0
                 pamt=0
                 samt=0
                 sttlamt = 0
                end if 
	for idx = 1 to 130
		objFile.Write "-"
	next
	objFile.WriteLine ""
	objFile.Write space(38)
        objFile.Write right(spaces&formatnumber(ttlpint,2),13)
        objFile.Write right(spaces&formatnumber(ttlpamt,2),13) 
        objFile.Write right(spaces&formatnumber(ttlsamt,2),13)
        objFile.Write right(spaces&formatnumber(ttlipint,2),13) 
        objFile.Write right(spaces&formatnumber(ttlipamt,2),13)
        objFile.Write right(spaces&formatnumber(ttlisamt,2),13)
	objFile.Write right(spaces&formatnumber(ttlTemp,2),15)
 	objFile.WriteLine ""
        objFile.WriteLine ""
        objFile.Write "銀行轉帳合共 : "
         objFile.Write  right(spaces&formatnumber(ttlASamt,2),15)
        objFile.WriteLine ""

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
<title>庫房轉賬細明表</title>
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
	<td colspan="15"><font size="4">特別個案轉帳細明表 </font></td>
	</tr>
	<tr height="35" valign="top" align="center">
	<td colspan="15"><font size="3">日期 : <%=mndate%></font></td>
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
 memno=rs("memno")
 memname=rs("memname")
 memcname=rs("memcname")
do while not rs.eof
           if memno<>rs("memno") then       
               sttlamt = pamt+pint+samt+ipamt+ipint+isamt 
               if sttlamt > 0 then
%>
	<tr>
		<td><%=memNo%></td>
		<td><%=memName%><%=memcname%></td>              		
		<td width=80 align="right"><%=formatNumber(pint,2)%></td>
		<td width=80 align="right"><%=formatNumber(pamt,2)%></td>
		<td width=80 align="right"><%=formatNumber(samt,2)%></td>
		<td width=80 align="right"><%=formatNumber(ipint,2)%></td>
		<td width=80 align="right"><%=formatNumber(ipamt,2)%></td>
		<td width=80 align="right"><%=formatNumber(isamt,2)%></td>
		<td width=80 align="right"><%=formatNumber(sttlamt,2)%></td>

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
                     case "E2","E1"
                          if rs("flag")<>"F"  then
                           pamt = rs("bankin")
                           ttlpamt = ttlpamt + pamt
                            ttlASAMT=ttlASamt  + rs("bankin")
                         else
                           ipamt = rs("bankin")
                            ttlipamt = ttlipamt + ipamt 
                         end if 
                     case "F2","F1"
                          if rs("flag")<>"F" then
                           pint = rs("bankin")
                           ttlpint = ttlpint + pint
                            ttlASAMT=ttlASamt  + rs("bankin")
                          else
                           ipint = rs("bankin")
			    ttlipint = ttlipint + ipint
                          end if  
                     case "A2","A1"
                          if rs("flag")<>"F" then
                           samt = rs("bankin")
                           ttlsamt = ttlsamt + samt
                           ttlASAMT=ttlASamt  + rs("bankin")
                          else
                          
                           isamt = rs("bankin")                   
                           ttlisamt = ttlisamt + isamt
                          end if 

                 
                          
               end select 
              
               ttlTemp=ttlTemp+rs("bankin")

	
	rs.movenext
loop
              sttlamt = pamt+pint+samt+ipamt+ipint+isamt 
               if sttlamt > 0 then
%>
	<tr>
		<td><%=memNo%></td>
		<td width=220><%=memName%><%=memcname%></td>              		
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
		<td>合共</td>
		<td></td>
	        <td align="right"><%=formatNumber(ttlpint,2)%></td>
                <td align="right"><%=formatNumber(ttlpamt,2)%></td>
                <td align="right"><%=formatNumber(ttlsamt,2)%></td>
                <td align="right"><%=formatNumber(ttlipint,2)%></td>
                <td align="right"><%=formatNumber(ttlipamt,2)%></td>
                <td align="right"><%=formatNumber(ttlisamt,2)%></td>
		<td align="right"><%=formatNumber(ttlTemp,2)%></td>
	</tr>
</table>
</center>
<table border="0" cellpadding="0" cellspacing="0">
	<br>
        <tr>
           <td width="30"></td>
           <td width="100" align="right"><b>銀行轉帳合共 :</b></td> 
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
