<!-- #include file="../conn.asp" -->



<%
server.scripttimeout = 1800

xmon = request.form("nmon")



yy = year(date())
mm = month(date())

if (xmon-mm) >= 0 then
   xyy = yy - 1
   xmon = 12 - (xmon - mm)




else
   xmon = mm - xmon
   xyy = yy
end if

todate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
chkdate = xyy&"."&right("0"&xmon,2)



ttlamt = 0

sql  = "SELECT a.memno,convert(char(10),max(a.sdate),102),b.memname,b.memcname,b.mstatus  FrOM share a,memmaster b where a.memno=b.memno and a.code<>'AI'  group by a.memno,b.memname,b.memcname,b.mstatus having convert(char(7),max(a.sdate),102)<='"&chkdate&"' ORDER  BY a.memno "
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn


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
	objFile.Write "冷戶細明表"
	objFile.WriteLine ""	
	objFile.Write "日期"&":"&right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
	objFile.WriteLine ""	
	objFile.WriteLine ""	
	objFile.Write " 社員編號 "
	objFile.Write "         社員名稱                 "
	objFile.Write "            交易日期   "
	objFile.WriteLine ""
	for idx = 1 to 101
		objFile.Write "-"
	next
	objFile.WriteLine ""
        xmemno =rs("memno") 
        ttlcnt = 0 
	do while not rs.eof

           select case rs("mstatus")
                  case "B","V","C"
                  CASE ELSE  
             xdate =right(rs(1),2)&"/"&mid(rs(1),6,2)&"/"&left(rs(1),4)
             ttlcnt = ttlcnt + 1
                   
                name1 = left(rs("memname")&spaces,24)
                name2 = left(rs("memcname")&spaces,10)   
                ttlname = name1+name2
  		objFile.Write left(" "&xmemNo&spaces,10) 
		objFile.Write left(ttlname&spaces,36)
                objFile.Write left("     "&xdate&spaces,22)
		objFile.WriteLine    


                xmemno = rs("memno")

   
         end select            
        
 
             
	rs.movenext
	loop
	for idx = 1 to 101
		objFile.Write "-"
	next
	objFile.WriteLine ""
	objFile.Write space(72)
	objFile.Write  left("總數 ："&spaces,10)
	objFile.Write right(spaces&formatnumber(ttlcnt,2),13)
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
<title>冷戶列印</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0" >
	<tr height="30" valign="top" align="center">
	<td colspan="15"><font size="4">水務署員工儲蓄互助社</font></td>
	</tr>
	<tr height="35" valign="top" align="center">
	<td colspan="15"><font size="4">冷戶細明表</font></td>
	</tr>
	<tr height="35" valign="top" align="center">
	<td colspan="15"><font size="4">日期 : <%=todate%></font></td>
        </tr>
</center>
</table>
<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="336699">
	<tr bgcolor="#330000" align="center">
		
	<td><font size="2" color="#FFFFFF">社員編號</font></td>
	<td><font size="2" color="#FFFFFF">社員名稱</font></td>
	<td><font size="2" color="#FFFFFF"> 日期 </font></td>

	</tr>
	
<%
   if not rs.eof then
        xmemno =rs("memno") 
	ttlcnt = 0
	do while not rs.eof
           memno = rs("memno") 

          select case rs("mstatus")
                  case "B","V","C"
                  CASE ELSE 
                   
             xdate =right(rs(1),2)&"/"&mid(rs(1),6,2)&"/"&left(rs(1),4)
             ttlcnt = ttlcnt + 1 

%>
   <tr bgcolor="#FFFFFF">
	
  	<td><font size="2"><%=rs("memno")%></font></td>
	<td><font size="2"><%=rs("memname")%><%=rs("memcname")%></font></td>
	<td ><font size="2"><%=xdate%></font></td>

   </tr> 
<%	
                xmemno = rs("memno")
                samt = 0
          end select                                                         
rs.movenext
loop
end if
%>
	<tr>
     		
                <td></td>
		 <td>總數 ：</td>
                
	
		<td width=100 align="right"><%=formatNumber(ttlcnt,2)%></td>
	

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
