<!-- #include file="../conn.asp" -->

<%
server.scripttimeout = 2400

yy = request.form("xyr")
mm = request.form("xmon")

xxx = yy+mm
randomize
xx = ROUND(rnd(xxx)*2000,0)
idx=round(rnd(XX)*16+1,0)
mndate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
xidx = "#temp"&idx

  conn.begintrans

              conn.execute( "create table "&xidx&"  ( memno int ,lnnum char(10), ldate smalldatetime, code char(2) , amount money ) ")
              conn.execute( "insert into "&xidx&"  (memno,ldate,code,amount ) select memno,ldate,code,amount from share where (right(code,1)='3' or code='A8') and  year(ldate) ='"&yy&"' and month(ldate) ='"&mm&"' order by memno,ldate,code "  )
              conn.execute( "insert into "&xidx&"  (memno,lnnum,ldate,code,amount ) select memno,lnnum,ldate,code,amount from loan  where (right(code,1)='3' ) and year(ldate)='"&yy&"' and month(ldate)='"&mm&"'  order by memno,ldate,code  " )
    

  conn.committrans

sql = "select * from  "&xidx&" order by memno,ldate,code"   
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn,1,1
if rs.eof then 
   
   response.redirect "monCtlst.asp"
end if
 
 
ttlamt = 0
ttlsamt = 0
ttlpamt = 0
ttlpint = 0
ttlisamt = 0
ttlipamt = 0
ttlipint = 0
ttlcnt = 0
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
	objFile.Write "現金帳細項明細表"
	objFile.WriteLine ""	

	objFile.WriteLine ""
      
	objFile.Write "社員編號   "
	objFile.Write "      社員名稱            "
	objFile.Write "貸款編號   "
	objFile.Write "交易日期   "
	objFile.Write "       (股金)   "
	objFile.Write "     (利息)   "
	objFile.Write "   (本金)   "
	objFile.Write "  (總金額)  "
	objFile.WriteLine ""
	for idx = 1 to 117
		objFile.Write "-"
	next
	objFile.WriteLine ""
  
          memno =rs("memno")             
          xdate = rs("ldate") 
       
        saveamt = 0
	do while not rs.eof
           if rs("memno") <> memno or (rs("memno")=memno and  xdate<>rs("ldate") )  then
              sql1 = "select memname,memcname,memhkid  from memmaster where memno='"&memno&"'  order by memno"  
              Set rs1 = Server.CreateObject("ADODB.Recordset")
              rs1.open sql1, conn,2,2
              if not rs1.eof then
                 memname = rs1("memname")
                 memcname=rs1("memcname")   
                 memhkid = rs1("memhkid")     
              end if
               rs1.close 

              sttlamt = samt  + pint + pamt
              gttlamt = gttlamt + sttlamt
              ttlcnt = ttlcnt + 1
              objFile.Write left("　　"&memno&spaces,8)
              objFile.Write left("　　　　"&memcname&spaces,20)
              objFile.Write left(lnnum&spaces,12)
              objFile.Write left(xdate&spaces,12)
              objFile.Write right(spaces&formatnumber(samt,2),13)
              objFile.Write right(spaces&formatnumber(pint,2),13)
              objFile.Write right(spaces&formatnumber(pamt,2),13)
              objFile.Write right(spaces&formatnumber(sttlamt,2),13) 
              objFile.WriteLine ""                         
 
              samt = 0          
              pint = 0 
              pamt =  0          
              lnnum = ""  
              if memno <> rs("memno") then  
                 memno =rs("memno") 
              end if
             xdate=rs("ldate")              
           end if
           select case rs("code")
                  case "A3", "A8"
                        xdate = rs("ldate")  
                       samt = samt + rs("amount")                       
                       ttlsamt = ttlsamt +  rs("amount")
                  case "E3"
                       lnnum = rs("lnnum")
                       pamt =pamt + rs("amount")
                        xdate = rs("ldate")
                       ttlpamt = ttlpamt +  rs("amount")
                  case "F3"     
                        lnnum = rs("lnnum")  
                        pint =pint +  rs("amount")
                         xdate = rs("ldate")  
                        ttlpint  = ttlpint +  rs("amount")
          end select             
       	rs.movenext
	loop
         sttlamt = samt  + pint + pamt
         
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

             
              gttlamt = gttlamt + sttlamt
              ttlcnt = ttlcnt + 1
            
             objFile.Write left("　　"&memno&spaces,8)
              objFile.Write left("　　　　"&memcname&spaces,20)
              objFile.Write left(lnnum&spaces,12)
              objFile.Write left(xdate&spaces,12)
              objFile.Write right(spaces&formatnumber(samt,2),13)
              objFile.Write right(spaces&formatnumber(pint,2),13)
              objFile.Write right(spaces&formatnumber(pamt,2),13)
              objFile.Write right(spaces&formatnumber(sttlamt,2),13) 
              objFile.WriteLine ""             
                  
              samt = 0          
              pint = 0 
              pamt =  0          
              lnnum = ""    
             
        end if      
	for idx = 1 to 117
		objFile.Write "-"
	next
         objFile.WriteLine ""  
	objFile.Write "總人數 ： "
        objFile.Write  right(spaces&formatnumber(ttlcnt,0),13)
        objFile.Write "　　　　　　　　　　　　　　"
    	objFile.Write "總金額 ： "
	objFile.Write right(spaces&formatnumber(ttlpint,2),13)
	objFile.Write right(spaces&formatnumber(ttlpamt,2),13)
	objFile.Write right(spaces&formatnumber(ttlsamt,2),13)
	objFile.Write right(spaces&formatnumber(gttlamt,2),13)
	objFile.WriteLine ""
	
	objFile.Write "　　　　"
        objFile.Write  "         ======"
        objFile.Write "　　　　　　　　　　　　　　"
    	objFile.Write "　　　　  "
	objFile.Write " ============"
	objFile.Write " ============"
	objFile.Write " ============"
	objFile.Write " ============"
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
<title現金帳細項列表</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0" >
	<tr height="30" valign="top" align="center"><td colspan="15"><font size="4"  face="標楷體" >水務署員工儲蓄互助社<br>現金帳細項明細表<br><font size="2"  face="標楷體" >日期 : <%=mndate%></font></font></td></tr>
        <tr height="30" ><td colspan=9></td></tr>

        
	<tr height="15" valign="bottom">
        <font size="2"  face="標楷體" >
	<td width="80">社員編號</td>
	<td width="80">社員名稱</td>
	<td width="90" align="center">貸款編號</td>
	<td width="60" align="center"> 日期 </td>
	<td width="60" align="center">股金</td>
	<td width="80" align="center"> 利息</td>
	<td width="80" align="center">本金</td>
	
	<td width="80" align="center">總金額</td>
	</tr>
	<tr><td colspan=9><hr></td></tr>


	
<%
       
        memno=rs("memno")
           xdate = rs("ldate") 
        saveamt = 0
	do while not rs.eof
           if rs("memno") <> memno or ( rs("memno")= memno and xdate<>rs("ldate") )   then
              sql1 = "select memname,memcname,memhkid  from memmaster where memno='"&memno&"'  order by memno"  
              Set rs1 = Server.CreateObject("ADODB.Recordset")
              rs1.open sql1, conn,2,2
              if not rs1.eof then
                 memname = rs1("memname")
                 memcname=rs1("memcname")   
                 memhkid = rs1("memhkid")     
              end if
               rs1.close 

              sttlamt = samt  + pint + pamt
              gttlamt = gttlamt + sttlamt
               ttlcnt = ttlcnt + 1 
%>
    <tr bgcolor="#FFFFFF">
	
  	<td width="80"><%=memno%></td>
	<td width="80"><%=memcname%></td>
	<td width="90" align="center"><%=lnnum%></td>	
	<td width="60" align="center"><%=xdate%></font></td>
	<td width="100" align="right"><%=formatnumber(samt,2)%></font></td>
	<td width="100" align="right"><%=formatnumber(pint,2)%></font></td>
	<td width="100" align="right"><%=formatnumber(pamt,2)%></font></td>
	<td width="100" align="right"><%=formatnumber(sttlamt,2)%></font></td>
   </tr> 
<%	
             samt = 0          
              pint = 0 
              pamt =  0          
              lnnum = ""    
              memno =rs("memno") 
             
           end if
           select case rs("code")
                  case "A3", "A8"
                        xdate = rs("ldate")  
                       samt = samt + rs("amount")                       
                       ttlsamt = ttlsamt +  rs("amount")
                  case "E3"
                       lnnum = rs("lnnum")
                       pamt =pamt + rs("amount")
                        xdate = rs("ldate")
                       ttlpamt = ttlpamt +  rs("amount")
                  case "F3"     
                        lnnum = rs("lnnum")  
                        pint =pint +  rs("amount")
                         xdate = rs("ldate")  
                        ttlpint  = ttlpint +  rs("amount")
          end select             
rs.movenext
loop
          sttlamt = pint+pamt + samt 
          gttlamt = gttlamt + sttlamt

         if sttlamt > 0 then
               ttlcnt = ttlcnt + 1
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
	<td width="100" align="right"><%=formatnumber(samt,2)%></font></td>
	<td width="100" align="right"><%=formatnumber(pint,2)%></font></td>
	<td width="100" align="right"><%=formatnumber(pamt,2)%></font></td>
	<td width="100" align="right"><%=formatnumber(sttlamt,2)%></font></td>
   </tr> 
<%end if %>
	<tr><td colspan=9><hr></td></tr>
	<tr>
		
		<td >總人數 ：</td>             		
                <td  width="80" align="right"><%=formatNumber(ttlcnt,0)%></td>
                <td width="90"> 
		 <td>總金額 ：</td>
                 <td width=100 align="right"><%=formatNumber(ttlsamt,2)%></td>	
		 <td width=100 align="right"><%=formatNumber(ttlpint,2)%></td>
	 	<td width=100 align="right"><%=formatNumber(ttlpamt,2)%></td>	
	
		<td width=100 align="right"><%=formatNumber(gttlamt,2)%></td>
	

	</tr>
	<tr>
		
		<td width="80">             		
                <td width="80" align="right">======</td>	
                <td width="90"> 
		<td width="60"> 
                <td width=100 align="right">==========</td>	
		<td width=100 align="right">==========</td>
	 	<td width=100 align="right">==========</td>		
		<td width=100 align="right">==========</td>
	

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
