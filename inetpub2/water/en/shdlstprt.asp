<!-- #include file="../conn.asp" -->

<%
server.scripttimeout = 1800


todate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())

shddate = request.form("divdate")

ddate1 = dateserial(right(shddate,4),mid(shddate,4,2),left(shddate,2))
ddate2 = dateserial(right(shddate,4),mid(shddate,4,2),left(shddate,2)+1)




SQl = "select a.memno,a.ldate,a.code,a.amount*-1 , b.memname,b.memcname   from share a ,memmaster b where a.memno=b.memno and (left(a.code,1)='C' and a.code<>'C0')  and (a.ldate= '"&ddate1&"'  or  a.ldate = '"&ddate2&"' ) order by  a.code,a.memno  "
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
<title>股息分配列表</title>
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
	<td colspan="15"><font size="4">股息分配細明表</font></td>
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
	
	<td><font size="2" color="#FFFFFF">類別</font></td>
	<td><font size="2" color="#FFFFFF">金額</font></td>
        <td><font size="2" color="#FFFFFF"> 部分總金額</font></td>
       
	
	</tr>
	
<%
        sttlcnt1 = 0
        sttlcnt2 = 0
        sttlcnt3 = 0
        sttlamt1 = 0
        sttlamt2 = 0
        sttlamt3 = 0
        
        ttlcnt = 0
        ttlamt = 0
     
        xcode = rs("code")
	do while not rs.eof
               
    
                select case rs("code")
                       case "CH"
                           status ="暫停股息"  
                           sttlcnt1 = sttlcnt1  + 1
                           sttlamt1 = sttlamt1 + rs(3)*-1
                            amount    = rs(3)*-1 
                       case "C1"
  	 	 	   status ="銀行"  
                           sttlcnt2 = sttlcnt2  + 1
                           sttlamt2 = sttlamt2 +  rs(3)
                            amount    = rs(3) 
                       case "C3" 
                           sttlcnt3 = sttlcnt3  + 1
                           sttlamt3 = sttlamt3 + rs(3)
                           status ="現金/支票"
                              amount    = rs(3) 
                      
               end select
                
               memno     = rs("memno")
               ttlamt = ttlamt +  amount
               ttlcnt = ttlcnt + 1
               memname   = rs("memname")
               memcname  = rs("memcname")
                
               code       = rs("code")
               rs.movenext
               if not rs.eof then
                  if xcode <> rs("code") then
                     xcode =  rs("code")
                  end if 
               else
                   xcode = ""
               end if
               if code <>  xcode   then                      
                  select case code
                         case "CH"
                               xamt = sttlamt1
                         case "C1"
                               xamt = sttlamt2
                         case "C3"
                              xamt = sttlamt3
                  end select
                  opt =1
                  
              end if        
            
%>
   <tr bgcolor="#FFFFFF">
	
  	<td><font size="2"><%=memno%></font></td>
        <td><font size="2"><%=memname%>(<%=memcname%>)</font></td> 
	<td><font size="2"><%=status%></font></td>
        <td><font size="2"><%=formatnumber(amount,2)%></font></td>
        <%if opt = 1 then %>
        <td align="right"><font size="2"><%=formatnumber(xamt ,2)%></font></td>
        <%else%>
        <td></td>
        <% end if
           opt = 0
          %>

   </tr> 
<%	
              

loop
%>
	<tr>
		<td></td>
		<td>總數:<%=ttlcnt%></td>              		                 
		 <td>總金額 ：</td>
             
		<td width=100 align="right"><font size="2"><b><%=formatNumber(ttlamt,2)%></b></font></td>

				
              


	</tr>
</table>
<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="336699">
       <tr>
             <td>暫停股息</td>
             <td wiidth="10">
              <td align="right"><font size="2"><%=formatnumber(sttlcnt1  ,2)%></font></td>
       </tr>
       <tr>
             <td>銀行</td>
             <td wiidth="10">
              <td align="right"><font size="2"><%=formatnumber(sttlcnt2  ,2)%></font></td>
       </tr>
       <tr>
             <td>現金/支票</td>
             <td wiidth="10">
              <td align="right"><font size="2"><%=formatnumber(sttlcnt3  ,2)%></font></td>
       </tr>
</center>
</body>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
