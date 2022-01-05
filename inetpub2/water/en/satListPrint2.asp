<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
server.scripttimeout = 1800

SQl = "select a.memno,a.lnnum,a.trefno,a.amount,b.memname,b.memcname  from sadttran a,memmaster b where a.memno=b.memno   "
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn

ttlamt = 0
ttlsamt = 0
ttlpamt = 0
ttlpint = 0
ttlisamt = 0
ttlipamt = 0
ttlipint = 0

if request.form("output")="word" then
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
	objFile.Write "庫房扣薪細明表"
	objFile.WriteLine ""	
	objFile.Write "日期"&":"&right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
	objFile.WriteLine ""	
	objFile.WriteLine ""	
	objFile.Write " 社員編號 "
	objFile.Write "         社員名稱                 "
	objFile.Write "  (利息)   "
	objFile.Write "    (本金)   "
	objFile.Write "   (股金)   "

	objFile.Write " (總金額)  "
	objFile.WriteLine ""
	for idx = 1 to 130
		objFile.Write "-"
	next
	objFile.WriteLine ""
        xmemno =rs("memno") 
  
	do while not rs.eof
           if rs("memno") <> xmemno then
 
                 if sttlamt > 0 then
  		objFile.Write left(" "&xmemNo&spaces,10)
		objFile.Write left(rs("memname")&" "&rs("memcname")&spaces,25)
		objFile.Write right(spaces&formatnumber(pint,2),13)
		objFile.Write right(spaces&formatnumber(pamt,2),13)
		objFile.Write right(spaces&formatnumber(samt,2),13)
		
		
                objFile.Write right(spaces&formatnumber(sttlamt,2),15)
		objFile.WriteLine ""
		ttlTemp=ttlTemp+sttlamt 

                xmemno = rs("memno")
                sttlamt = 0
                pamt = 0
                pint = 0
                samt = 0
                ipint = 0
                ipamt = 0
                isamt = 0                 
              end if   
           end if  
              select case rs("trefno")
                     case "E2"
                           pamt = rs("amount")
                           ttlpamt = ttlpamt + pamt
                     case "F2"
                           pint = rs("amount")
                           ttlpint = ttlpint + pint
                     case "A2"
                           samt = rs("amount")
                           ttlsamt = ttlsamt + samt
 
               end select 
               sttlamt = sttlamt + rs("amount")
	rs.movenext
	loop
	for idx = 1 to 130
		objFile.Write "-"
	next
	objFile.WriteLine ""
	objFile.Write space(70)
	objFile.Write right(spaces&formatnumber(ttlTemp,2),113)
	objFile.WriteLine ""
 	objFile.Write left("本金合共"&spaces,20)
	objFile.Write right(spaces&formatnumber(ttlpamt,2),20)
	objFile.WriteLine ""       
 	objFile.Write left("利息合共"&spaces,20)
	objFile.Write right(spaces&formatnumber(ttlpint,2),20)
	objFile.WriteLine ""   
 	objFile.Write left("股金合共"&spaces,20)
	objFile.Write right(spaces&formatnumber(ttlsamt,2),20)
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
	<td colspan="15"><font size="4">
server.scripttimeout = 1800

SQl = "s細明表</font></td>
	</tr>
	<tr height="15" valign="bottom">
	<td width="80"><b>社員編號</b></td>
	<td width="200"><b>社員名稱</b></td>
	<td width="60" align="right"><b>(利息) </b></td>
	<td width="60" align="right"><b>(本金)</b></td>
	<td width="60" align="right"><b>(股金)</b></td>

	<td width="80" align="right"><b>(總金額)</b></td>
	</tr>
	<tr><td colspan=9><hr></td></tr>
<%
xmemno = rs("memno")
do while not rs.eof
            if rs("memno") <> xmemno then
 
                 if sttlamt > 0 then
%>
	<tr>
		<td><%=rs("memNo")%></td>
		<td><%=rs("memName")%><%=rs("memcname")%></td>              		
		<td width=80 align="right"><%=formatNumber(pint,2)%></td>
               
		<td width=100 align="right"><%=formatNumber(pamt,2)%></td>

		<td width=100 align="right"><%=formatNumber(samt,2)%></td>
	
		<td width=100 align="right"><%=formatNumber(ipint,2)%></td>

		<td width=100 align="right"><%=formatNumber(ipamt,2)%></td>

		<td width=100 align="right"><%=formatNumber(isamt,2)%></td>
	
		<td width=100 align="right"><%=formatNumber(sttlamt,2)%></td>

	</tr>
<%
		ttlTemp=ttlTemp+sttlamt 

                xmemno = rs("memno")
                sttlamt = 0
                pamt = 0
                pint = 0
                samt = 0
                ipint = 0
                ipamt = 0
                isamt = 0                 
              end if   
           end if  
              select case rs("trefno")
                     case "E2"
                           pamt = rs("amount")
                           ttlpamt = ttlpamt + pamt
                     case "F2"
                           pint = rs("amount")
                           ttlpint = ttlpint + pint
                     case "A2"
                           samt = rs("amount")
                           ttlsamt = ttlsamt + samt

               end select 
               sttlamt = sttlamt + rs("amount")      
rs.movenext
loop
%>
	<tr>
		<td></td>
		<td></td>              		
		<td width=100 align="right"><%=formatNumber(ttlpint,2)%></td>
		
		<td width=100 align="right"><%=formatNumber(ttlpamt,2)%></td>
		<td width=100 align="right"><%=formatNumber(ttlsamt,2)%></td>
		

		<td width=100 align="right"><%=formatNumber(ttlTemp,2)%></td>

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
