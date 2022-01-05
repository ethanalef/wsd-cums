<!-- #include file="../conn.asp" -->

<%


   mndate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
 
   server.scripttimeout = 1800
   ddate = request.form("stdate1")

   yy = right(ddate,4)
   mm = mid(ddate,4,2)
   dd = left(ddate,2)
   ddate=dateserial(yy,mm,dd)
  idx = request.form("idx")  
  sidx = request.form("sidx") 

SQl = "SELECT  memno,memname,memcname,mstatus FROM  memmaster  WHERE  MSTATUS='V'   ORDER BY  memno,memname,memcname,mstatus "

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
	<tr height="30" valign="top" align="center"><td colspan="15"><font size="3"  face="標楷體" >水務署員工儲蓄互助社<br>IVA列表<br><font size="2"  face="標楷體" >日期 : <%=mndate%></td></tr>
        

</table>
<table border="0" cellpadding="0" cellspacing="0" >
<tr height="30" valign="top" align="center"><td colspan="15"><font size="3"  face="標楷體" >截至日期 : <%=dd%>日<%=mm%>月<%=yy%>年</font></td></tr>
</table>
<table border="0" cellpadding="0" cellspacing="0">
	<tr height="15" valign="bottom">
		<td width=70 align="center"><font size="2"  face="標楷體" >社員名稱</font></td>
               <td width=180 align="center"><font size="3"  face="標楷體" >英文姓名</font></td>
		<td width=70 align="center"><font size="3"  face="標楷體" >中文姓名</font></td>	
		<td width="130" align="right"><font size="2"  face="標楷體" >貸款結餘</font></td>
		<td width="130" align="right"><font size="2"  face="標楷體" >股金結餘</font></td>
	</tr>
	<tr><td colspan=5><hr></td></tr>
<%

        memno=rs("memno")
        do while not rs.eof
            XAMT = 0
            lnamt = 0
            set ms=conn.execute("select *  from share where memno='"&rs("memno")&"' and ldate<='"&ddate&"'  ")
                  do while not ms.eof
                     select case left(ms("code"),1)
                            case "0","A","C"
                                    if ms("code") <>"AI" then
                                         xamt = xamt + ms("amount")
                                   end if
                            case "B","G","H","M"
                                 xamt = xamt - ms("amount")
                     end select
                     ms.movenext
                  loop             
                  ms.close
           set ms=conn.execute("select *   from loan where memno='"&rs("memno")&"' and ldate<='"&ddate&"' order by ldate desc  " )
            if not ms.eof then
               xx = 0 
               do while not ms.eof  and xx = 0
                 
                  select case  left(ms("code"),1)
                         case "D","0"
                              if ms("code") = "D9" or ms("code")="0D" then
                                 lnamt = lnamt + ms("amount")
                                 xx = 1 
                               end if
                         case "E"
                                lnamt = lnamt - ms("amount")
                 end select
 
                 ms.movenext                    
                loop
            end if
            ms.close
 
            pass  = 0
            select case idx
                   case 1
                          if xamt > 0 then
                             pass = 1
                          end if
                   case 2
                          if xamt = 0 then
                             pass = 1
                          end if
                  case 3
                          pass = 1
           
            end select  
            spass = 0
            select case sidx
                   case 1
                          if lnamt > 0 then
                             spass = 1
                          end if
                   case 2
                          if lnamt = 0 then
                             spass = 1
                          end if
                  case 3
                          spass = 1
                    
           end select    
           if pass = 1 and spass = 1 then               
%>
	<tr>
		<td width=70 align="center"><%=rs("memNo")%></td>
                <td width=180 align="left"><%=rs("memname")%></td> 
		<td width=70 align="center"><font size="3"  face="標楷體" ><%=rs("memcname")%></font> </td>
		<td align="right"><%=formatNumber(lnamt,2)%></td>
		<td align="right"><%=formatNumber(xamt,2)%></td>
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
