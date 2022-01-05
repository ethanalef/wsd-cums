<!-- #include file="../conn.asp" -->

<%


   mndate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
 
   server.scripttimeout = 1800
   ddate = request.form("stdate1")

   yy = right(ddate,4)
   mm = mid(ddate,4,2)
   dd = left(ddate,2)

     
     
   
   ydate = dateserial( yy-2 , mm,dd)
   mdate = dateserial( yy   , 7,dd)
SQl = "SELECT  a.memno,MAX(a.ldate) AS xdate,b.memname,b.memcname,b.mstatus "&_
      "FROM  share a ,memmaster b WHERE   a.memno=b.memno and "&_
      " a.code <>'MF'  "&_
      "GROUP BY  a.memno ,b.memname,b.memcname,b.mstatus "

Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn, 1,1

if request.form("output")="word" then
	Response.ContentType = "application/msword"
elseif request.form("output")="excel" then
	Response.ContentType = "application/vnd.ms-excel"
end if

%>
<html>
<head>
<title>Delinquent Report</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<LINK href="../main.css" rel=STYLESHEET type=text/css>
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0" >
	<tr height="30" valign="top" align="center"><td colspan="15"><font size="3"  face="標楷體" >水務署員工儲蓄互助社<br>冷戶列表<br><font size="2"  face="標楷體" >日期 : <%=mndate%><br><font size="3"  face="標楷體" >巳兩年</font></font></font></td></tr>
        

</table>
<table border="0" cellpadding="0" cellspacing="0">
	<tr height="15" valign="bottom">
		<td width=70 align="center"><font size="2"  face="標楷體" >社員名稱</font></td>
               <td width=180 align="center"><font size="3"  face="標楷體" >英文姓名</font></td>
		<td width=70 align="center"><font size="3"  face="標楷體" >中文姓名</font></td>	
		<td width="130" align="right"><font size="2"  face="標楷體" >最後來往帳日期</font></td>
		<td width="130" align="right"><font size="2"  face="標楷體" >股金結餘</font></td>
	</tr>
	<tr><td colspan=5><hr></td></tr>
<%

        memno=rs("memno")
        do while not rs.eof
 

         xdate=right("0"&day(rs("xdate")),2)&"/"&right("0"&month(rs("xdate")),2)&"/"&year(rs("xdate"))
           xamt = 0

           if (round((mdate - rs("xdate"))/365,2) - 2) > 0 or (rs("mstatus")="D"  )  then
               select case rs("mstatus")
                      case "A","M","N","T","D","H","F","0","1","2"
                          set ms=conn.execute("select * from share where memno='"&rs("memno")&"' and ldate <'"&mdate&"' ")
                          do while not ms.eof
                             select case ms("code")
                                    case "A1","A2","A3","C0","C1","C3" ,"A0","A7" ,"A4" ,"C5","0A"
                                         xamt = xamt + ms("amount")
                                    case "G0" ,"H0","B0","B1","B3","BE","BF","G3","H3","MF" 
                                         xamt = xamt - ms("amount") 
                             end select
                             Lastdate = ms("ldate")
                             ms.movenext
                           loop
                           ms.close 
           
  
            
%>
	<tr>
		<td width=70 align="center"><%=rs("memNo")%></td>
               <td width=180 align="left"><%=rs("memname")%></td> 
		<td width=70 align="center"><font size="3"  face="標楷體" ><%=rs("memcname")%></font> </td>
		<td align="right"><%=Lastdate%></td>
		<td align="right"><%=formatNumber(xamt,2)%></td>
	</tr>
<%
 
      end select
 
  end if
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
