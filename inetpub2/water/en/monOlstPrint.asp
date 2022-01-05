<!-- #include file="../conn.asp" -->

<%
server.scripttimeout = 1800
server.scripttimeout = 2400
yy = request.form("xyr")
mm = request.form("xmon")
mndate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
xxx = yy+mm
randomize
xx = ROUND(rnd(xxx)*2000,0)
idx=round(rnd(XX)*36+1,0)

xidx = "#temp"&idx

conn.begintrans
              conn.execute( "create table "&xidx&"  ( memno int ,lnnum char(10), ldate smalldatetime, code char(2) , amount money,lnflag char(1)) ")
              conn.execute( "insert into "&xidx&"  (memno,ldate,code,amount ,lnflag ) select memno,ldate,code,amount,lnflag from share where (code='A0' or code='B0' or code='B1'  or code='0A' or code='A7' or code='A4' or code='MF' ) and   year(ldate) ='"&yy&"' and month(ldate) ='"&mm&"' order by memno,ldate,code "  )
              conn.execute( "insert into "&xidx&"  (memno,lnnum,ldate,code,amount ) select memno,lnnum,ldate,code,amount from loan  where (code='E0' or code='F0' or code='E6' or code='F6' or code='E7' or code='F7' or code='E9') and  year(ldate)='"&yy&"' and month(ldate)='"&mm&"'  order by memno,ldate,code  " )
conn.committrans

sql = "select * from "&xidx&"   order by memno,ldate,code"   
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn,1,1
if rs.eof then 
   response.redirect "monOtlst.asp"
end if

ttlamt   = 0
ttlsamt  = 0
ttlpamt  = 0
ttlpint  = 0
ttlisamt = 0
ttlipamt = 0
ttlipint = 0
ttlcnt   = 0
ttlcamt  = 0

if request.form("output")="word" then
	Response.ContentType = "application/msword"
elseif request.form("output")="excel" then
	Response.ContentType = "application/vnd.ms-excel"

end if
%>
<html>
<head>

<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>

<table border="0" cellpadding="0" cellspacing="0" >
	<tr height="30" valign="top" align="center"><td colspan="15"><font size="4"  face="標楷體" >水務署員工儲蓄互助社<br>其他帳細項明鈿表<br><font size="2"  face="標楷體" >日期 : <%=mndate%></font></font></td></tr>
    <tr height="30" ><td colspan=9></td></tr>

	<tr height="15" valign="bottom">
        <font size="2"  face="標楷體" >
		<td width="80">社員編號</td>
		<td width="80">社員名稱</td>
		<td width="90" align="center">貸款編號</td>
		<td width="60" align="center">過帳日期</td>
		<td width="100" align="center">退股</td> 
		<td width="80" align="center">股金</td>
		<td width="80" align="center"> 百佳券 </td> 
		<td width="80" align="center"> 冷戶費 </td> 
		<td width="100" align="center"> 利息</td>
		<td width="100" align="center">本金</td>
	</tr>
	<tr><td colspan=10><hr></td></tr>

<%
memno = rs("memno") 
xdate = rs("ldate")
saveamt = 0
do while not rs.eof
	if rs("memno") <> memno or xdate<> rs("ldate")  then
        sql1 = "select memname,memcname,memhkid  from memmaster where memno='"&memno&"'  order by memno"  
        Set rs1 = Server.CreateObject("ADODB.Recordset")
        rs1.open sql1, conn,2,2
        if not rs1.eof then
                 memname = rs1("memname")
                 memcname=rs1("memcname")   
                 memhkid = rs1("memhkid")     
        end if
        rs1.close 
		sxdate =right("0"&day(xdate),2)&"/"&right("0"&month(xdate),2)&"/"&year(xdate) 
        sttlamt = pkamt  + pint + pamt+psamt+samt
        gttlamt = gttlamt + sttlamt
		ttlcnt = ttlcnt + 1

		%>
		<tr bgcolor="#FFFFFF">
				<td width="80"><%=memno%></td>
				<td width="80"><%=memcname%></td>
				<td width="90" align="center"><%=lnnum%></td>	
				<td width="60" align="center"><%=sxdate%></font></td>
				<td width="100" align="right"><%=formatnumber(psamt,2)%></td>
				<td width="80" align="right"><%=formatnumber(samt,2)%></td>
				<td width="80" align="right"><%=formatnumber(pkamt,2)%></td>	
				<td width="80" align="right"><%=formatnumber(camt,2)%></td>	 
				<td width="100" align="right"><%=formatnumber(pint,2)%></td>
				<td width="100" align="right"><%=formatnumber(pamt,2)%></td> 
		</tr> 
		<%	
        samt  = 0
        pkamt = 0          
        pint = 0 
        pamt =  0          
        psamt = 0
        lnnum = ""    
        camt = 0
        memno =rs("memno") 
    end if
    xdate = rs("ldate")  
    select case rs("code")
        case "B0","B1"
            if rs("lnflag")="Y" then
                pkamt = pkamt + rs("amount")
                ttlpkamt = ttlpkamt + rs("amount") 
            else 
                ttlpsamt = ttlpsamt + rs("amount")  
            end if   
            psamt = psamt+rs("amount")  
			
        case "A0","0A","A4","A7"
				samt =samt +  rs("amount")                       
                ttlsamt = ttlsamt + samt
				
        case "MF"
                camt =    rs("amount") 
                ttlcamt = ttlcamt + camt
				
        case "E7"
                lnnum = rs("lnnum")
                pamt =rs("amount")
                ttlpamt = ttlpamt + pamt
				
        case "F7"     
                lnnum = rs("lnnum")  
                pint = rs("amount")
                ttlpint  = ttlpint + pint
				
        case "E0","E6"
                lnnum = rs("lnnum")
                pamt =rs("amount")
                ttlpamt = ttlpamt + pamt
				
        case "F0","F6"     
                lnnum = rs("lnnum")  
                pint = rs("amount")
                ttlpint  = ttlpint + pint
    end select             
    rs.movenext
loop

sttlamt = pkamt  + pint + pamt+psamt+samt +camt
gttlamt = gttlamt + sttlamt

if sttlamt > 0 then
    sql1 = "select memname,memcname,memhkid  from memmaster where memno='"&memno&"'  order by memno"  
    Set rs1 = Server.CreateObject("ADODB.Recordset")
    rs1.open sql1, conn,2,2
    if not rs1.eof then
        memname = rs1("memname")
        memcname=rs1("memcname")   
        memhkid = rs1("memhkid")     
	end if
	rs1.close 

	%>
   <tr bgcolor="#FFFFFF">
		<td width="80"><%=memno%></td>
		<td width="80"><%=memcname%></td>
		<td width="90" align="center"><%=lnnum%></td>	
		<td width="60" align="center"><%=xdate%></font></td>
		<td width="100" align="right"><%=formatnumber(psamt,2)%></td>
		<td width="80" align="right"><%=formatnumber(samt,2)%></td>
		<td width="80" align="right"><%=formatnumber(pkamt,2)%></td>	
		<td width="100" align="right"><%=formatnumber(camt,2)%></td>	 
		<td width="100" align="right"><%=formatnumber(pint,2)%></td>
		<td width="100" align="right"><%=formatnumber(pamt,2)%></td> 
   </tr> 
	<%
end if
%>
	<tr><td colspan=10><hr></td></tr>
	<tr>
		
		<td >總人數 ：</td>             		
        <td  width="80" align="right"><%=formatNumber(ttlcnt+1,0)%></td>
        <td width="90"> 
		<td>總金額 ：</td>
		<td width="100" align="right"><%=formatnumber(ttlpsamt,2)%></td>
		<td width="80" align="right"><%=formatnumber(ttlsamt,2)%></td>
		<td width="80" align="right"><%=formatnumber(ttlpkamt,2)%></td>	
		<td width="100" align="right"><%=formatnumber(ttlcamt,2)%></td>	 
		<td width="100" align="right"><%=formatnumber(ttlpint,2)%></td>
		<td width="100" align="right"><%=formatnumber(ttlpamt,2)%></td> 
	</tr>
	<tr>
	<tr><td colspan=10><hr></td></tr>
		

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
