<!-- #include file="../conn.asp" -->

<%


   mndate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())

   server.scripttimeout = 1800

   yy = year(date())
   mm = month(date())
   dd = day(date())
   xyy = yy - 2 
   chkdate = dateserial( xyy , mm,dd)
   idx = request.form("idx")
SQl = "SELECT  memno,memname, memCName,memcontacttel,memMobile,ISNULL(memaddr1, '') AS 'memaddr1', "&_
      "ISNULL(memaddr2, '') AS 'memaddr2', "&_
      "ISNULL(memaddr3, '') AS 'memaddr3' " & _
      "FROM   memmaster  "&_
      "where  accode='9999' and "&_
      " mstatus not in ('C','B','P' ) "&_
      "order by memno "

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
	<tr height="30" valign="top" align="center"><td colspan="15"><font size="4"  face="標楷體" >水務署員工儲蓄互助社<br>退休社員細明表<br><font size="2"  face="標楷體" >日期 : <%=mndate%><br></font></font></td></tr>
        

</table>
<table border="0" cellpadding="0" cellspacing="0">
	<tr height="15" valign="bottom">
                <td width=70 align="center"><font size="2"  face="標楷體" >社員編號</font></td>
		<td width=140 align="center"><font size="2"  face="標楷體" >英文名稱</font></td>
		<td width=80 align="center"><font size="2"  face="標楷體" >中文名稱</font></td>		
                <td width="100" align="center"><font size="2"  face="標楷體" >股金結餘</font></td>
             	<td width="100" align="center"><font size="2"  face="標楷體" >聯絡電話</font></td>
             	<td width="100" align="center"><font size="2"  face="標楷體" >手提電話</font></td>
		<td width="500" align="center"><font size="2"  face="標楷體" >地址</font></td>
               
	</tr>
	<tr><td colspan=7><hr></td></tr>
<%


        do while not rs.eof
           yy = year(date())
           mm = month(date())
           dd = day(date())
            xdate=right("0"&dd,2)&"/"&right("0"&mm,2)&"/"&yy
           ydate = dateserial(yy - 2,mm,dd)
           xamt = 0
         
              set ms=conn.execute("select *  from share where memno='"&rs("memno")&"' and code not in ('CH','AI')  ")
                  do while not ms.eof
                     select case left(ms("code"),1)
                            case "0","A","C"
                                  
                                      xamt = xamt + ms("amount")
                                  
                            case "B","B","H","M"
                              
                                    xamt = xamt - ms("amount")
                              
                     end select
                     ms.movenext
                  loop             
                  ms.close
            select case idx 
                   case "1"
                        if xamt > 0 then
                           pass = 0
                        else 
                           pass = 1
                        end if            
                   case "2"
                       if xamt = 0 then  
                          pass = 0
                       else
                           pass = 2
                       end if
                  case "3"
                       pass = 0
             end select
            if pass = 0 then
           
            
%>
	<tr>
		<td width=70 align="center"><font size="2"  face="標楷體" ><%=rs("memNo")%></font></td>
                <td width=140 align="center"><font size="2"  face="標楷體" ><%=rs("memname")%></font></td>
		<td width=80 align="center"><font size="2"  face="標楷體" ><%=rs("memcname")%></font></td>
                <td width="100" align="right"><%=formatnumber(xamt,2)%></td>
		<td width="80" align="center"><%=rs("memcontacttel")%></td>
		<td width="80" align="center"><%=rs("memMobile")%></td>
		<td width="500" align="left"><%=rs("memaddr1")%><%=rs("memaddr2")%><%=rs("memaddr3")%></td>
               
	</tr>
<%

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
