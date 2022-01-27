<!-- #include file="../conn.asp" -->

<%
sqlstr=" and "
mndate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
server.scripttimeout = 1800

sxdate=request.form("cutdate")
yy = right(sxdate,4)
mm = mid(sxdate,4,2)
dd = left(sxdate,2)


mdate=dateserial(yy,mm,dd)


set rs = server.createobject("ADODB.Recordset")

xxx = yy+mm
randomize
xx = ROUND(rnd(xxx)*2000,0)
idx=round(rnd(XX)*26+1,0)


xidx = "#temp"&idx

  conn.begintrans

              conn.execute( "create table "&xidx&"  ( memno int , ldate smalldatetime, code char(2) , amount money ) ")
              conn.execute( "insert into "&xidx&"  (memno,code,amount ) select memno,code,sum(amount) from share where ldate<='"&mdate&"'  group by memno,code order by memno,code "  )
              conn.execute( "insert into "&xidx&"  (memno,code,amount ) select memno,code,sum(amount) from loan  where ldate<='"&mdate&"'  group by memno,code order by memno,code  " )
    

  conn.committrans

sql = "select * from  "&xidx&" order by memno,ldate,code"   
rs.open sql, conn, 1, 1

if rs.eof then
   response.redirect "memIlst.asp"
end if

ttlamt = 0
    
if request.form("output")="word" then
	Response.ContentType = "application/msword"
elseif request.form("output")="excel" then
	Response.ContentType = "application/vnd.ms-excel"

end if
%>
<html>
<head>
<title>>社員報告(保險)</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>

<table width="1012" border="0">

 <tr>
    <td width="99">&nbsp;</td>
    <td width="780">&nbsp;</td>
    <td width="142">&nbsp;</td>
  </tr>
	<tr>
        <td>&nbsp</td>
        <td align="center"><b><font size="4"  face="標楷體" >水務署員工儲蓄互助社<br>社員報告(保險)</font></b?</td>
        <td align="center"><font size="2"  face="標楷體" >日期 : <%=mndate%></font></td>
        </tr>
       

</table>
<table border="0" cellpadding="0" cellspacing="0">
	<tr height="15" valign="bottom">
	<td width="80" align="center"><font size="3" face="標楷體"  >社員編號</font></td>
	<td width=180 align="center"><font size="3"  face="標楷體" >英文姓名</font></td>
	<td width=70 align="center"><font size="3"  face="標楷體" >中文姓名</font></td>	
	<td width="80"  align="center"><font size="3" face="標楷體"  > 身分證</font></td>
	<td width="80"  align="center"><font size="3" face="標楷體"  > 出生日期</font></td>
	<td width="50"  align="center"><font size="3" face="標楷體"  > 性別</font></td>
	<td width="80"  align="center"><font size="3" face="標楷體"  > 年齡</font></td>
	<td width="130" align="center"><font size="3" face="標楷體"  >股金結餘</font></td>
        <td width="130" align="center"><font size="3" face="標楷體"  >貸款結餘</font></td>
        <td width="130" align="center"><font size="3" face="標楷體"  >分類</font></td> 
	</tr>
	<tr><td colspan=10><hr></td></tr>
<% 
   tlnamt = 0
 
   ttlamt = 0 
   ttl1   = 0
   ttl2   =0
   ttl3   =0
   ttl4   =0
   ttl5   = 0
   ttl6   =0
   ttl7   =0
   ttl8   =0
   ttl9   = 0
   ttl10   =0
   ttl11   =0
   ttl12   =0
   ttl13   = 0
   ttl14   =0
   ttl15   =0
   ttl16   =0
   ttl17   = 0
   ttl18   =0
   ttl19   =0
   cnt1   = 0
   cnt2   =0
   cnt3   =0
   cnt4   =0
   cnt5   = 0
   cnt6   =0
   cnt7   =0
   cnt8   =0
   cnt9   = 0
   cnt10   =0
   cnt11   =0
   cnt12   =0
   cnt13   = 0
   cnt14   =0
   cnt15   =0
   cnt16   =0
   cnt17   = 0
   cnt18   =0
   cnt19   =0
   xamt = 0
   lnamt = 0
   clsbal = 0
   ttlamt = 0
   ttllncnt = 0
   xmemno=rs("memno")
   ymemno = xmemno 
   xamt = 0
   lnamt = 0          

   XX = 0
   do while not rs.eof 


    
      if xmemno<> rs("memno")   then
         set ms=conn.execute("select memname,memcname,membday,memhkid,memGender,membday,mstatus  from memmaster where memno='"&xmemno&"' ")
         if not ms.eof then
             memname = ms("memname")    
             memcname=ms("memcname")
             xmstatus=  ms("mstatus")
             xmembday = ms("membday")
             xmemhkid = ms("memhkid")
             xmemgender =ms("memgender")
             if xmembday<>"" then
                 membday = right("0"&day(xmembday),2)&"/"&right("0"&month(xmembday),2)&"/"&year(xmembday)
                 age =year(date()) -year(xmembday)
                 yy = year(date())
                 adatae = dateserial(yy,1,1)
                 mm = month(xmembday)
                 dd = day(xmembday)
                   
                  bdate = dateserial(yy,mm,dd)
                  xday = ((bdate - adate)+z)/365
                  if xday >0.5 then
                     age = age + 1
                  end if
            else
                  age = 0
                  membday=""  
            end if
            end if
            ms.close 
     if xmemGender="M" then
         
         sex="男"
      else
         
          sex="女"
      end if

   
       
      

    
            tlnamt = tlnamt + lnamt 
            ttllncnt = ttllncnt + 1
     
        
         select case xmstatus
                case "A"
                     if clsbal > 0 then
                        ttl1 = ttl1 + clsbal
                        cnt1 = cnt1 + 1
                     end if
                     idx ="自動轉帳(ALL)"
                case "B"
                     if clsbal > 0  or (clsbal=0 and lnamt >0)  then
                        ttl2 = ttl2 +clsbal
                        cnt2 = cnt2 + 1
                      end if
                     
                     idx ="破產"
                case "C"
                   if clsbal > 0  or (clsbal=0 and lnamt >0)  then
                       ttl3 = ttl3+ clsbal
                       cnt3 = cnt3 + 1
                     end if 
                     idx ="退社"
                case "D"
                    if clsbal > 0 then
                       ttl4 = ttl4 + clsbal
                       cnt4 = cnt4 + 1 
                    end if 
                     idx ="冷戶"
                case "F"
                    ttl5 = ttl5 +  clsbal
                    cnt5 = cnt5 + 1
                     idx ="特別個案"
                case "H"
                   ttl6 = ttl6 + clsbal
                   cnt6 = cnt6 + 1
                     idx ="暫停銀行"
                case  "J"
                   ttl7 = ttl7 + clsbal
                   cnt7 = cnt7 + 1
                     idx ="新戶"
                case "L"
                 if clsbal > 0 or (clsbal=0 and lnamt>0) then
                   ttl8 = ttl8 + clsbal
                   cnt8 = cnt8 + 1  
                 end if
                     idx ="呆帳"
                case "M"
                  ttl9 = ttl9 + clsbal
                  cnt9 = cnt9 + 1
                     idx ="庫房,銀行"
                case "N"     
                 ttl10 = ttl10 + clsbal
                 cnt10 = cnt10 + 1
                     idx ="正常"            
                case "P"
                if clsbal > 0 then
                   ttl11 = ttl11 + clsbal
                   cnt11 = cnt11 + 1
                    idx ="去世"  
                end if   
                     
                case "T"
                 ttl12 = ttl12 + clsbal
                 cnt12 = cnt12 + 1
                     idx ="庫房"
                case "V"
                     if clsbal> 0 or (clsbal=0 and lnamt>0) then
                        ttl13 = ttl13 + clsbal
                        cnt13 = cnt13 + 1 
                     end if
                     idx ="IVA"
                    
                case "0"
                 ttl14 = ttl14 + clsbal
                 cnt14 = cnt14 + 1
                     idx =" 自動轉帳(股金)"
                case "1"
                 ttl15 = ttl15 + clsbal
                 cnt15 = cnt15 + 1
                     idx ="自動轉帳(股金,利息)"
                case "2"
                 ttl16 = ttl16 + clsbal
                 cnt16 = cnt16 + 1
                     idx ="自動轉帳(股金,本金)"
                case "3"
                 ttl17 = ttl17 + clsbal
                 cnt17 = cnt17 + 1
                     idx ="自動轉帳(利息,本金)"
                case "8"
                 ttl18 = ttl18 + clsbal
                 cnt18 = cnt18 + 1 
                     idx ="終止社籍轉帳"
                case "9"
                 ttl19 = ttl19 + clsbal
                 cnt19 = cnt19 + 1
                     idx ="終止社籍正常"
         end select    
    if  clsbal > 0  and xmstatus<>"V" then
   
%>
     <tr>
          <td width="80"><%=xmemno%></td>
          <td width=180 align="left"><%=memname%></td>   
          <td width="80"><font size="3" face="標楷體"  ><%=memcname%></font></td>
          <td width="80"><%=xmemhkid%></td>
          <td width="80"><%=membday%></td>
          <td width="50" align="center"><%=sex%></td> 
          <td width="80"  align="center"><%=formatnumber(age,0)%></td>  
          <td width="130" align="right"><%=formatnumber(clsbal,2)%></td>
          <td width="130" align="right"><%=formatnumber(lnamt,2)%></td>
          <td width="150" align="center"><font size="3" face="標楷體"  ></font><%=idx%></font></td>
     </tr>

<%
 
        ttlamt = ttlamt + clsbal
        
        end if    
 
          xmemno=rs("memno")
  
        clsbal = 0
        lnamt = 0
        xamt = 0
        
   end if  
 
             select case rs("code")
                      case "0D","D9"

                              lnamt= lnamt + rs("amount")
 
                      case "D8","E0","E1","E2","E3","E6","E7","EC"
                           lnamt = lnamt - rs("amount")
                           
                end select
       select case rs("code")
          case "0A","A1","A2","A3","C0","C1","C3","A0","A4","A7","C5","A8"
               
               clsbal = clsbal + rs("amount")
          case "B0","B1","G0","G1","G3","H0","H1","H3","MF","B3","BF","BE","B8"
                clsbal = clsbal - rs("amount")
         end select
         
    


     rs.movenext
    loop
    rs.close
         set ms=conn.execute("select memname,memcname,membday,memhkid,memGender,membday,mstatus  from memmaster where memno='"&xmemno&"' ")
         if not ms.eof then
             memname = ms("memname")    
             memcname=ms("memcname")
             xmstatus=  ms("mstatus")
             xmembday = ms("membday")
             xmemhkid = ms("memhkid")
             xmemgender =ms("memgender")
             if xmembday<>"" then
                 membday = right("0"&day(xmembday),2)&"/"&right("0"&month(xmembday),2)&"/"&year(xmembday)
                 age =year(date()) -year(xmembday)
                 yy = year(date())
                 adatae = dateserial(yy,1,1)
                 mm = month(xmembday)
                 dd = day(xmembday)
                   
                  bdate = dateserial(yy,mm,dd)
                  xday = ((bdate - adate)+z)/365
                  if xday >0.5 then
                     age = age + 1
                  end if
            else
                  age = 0
                  membday=""  
            end if
            end if
            ms.close 
     if xmemGender="M" then
         
         sex="男"
      else
         
          sex="女"
      end if

   
       
      

         
            tlnamt = tlnamt + lnamt 
            ttllncnt = ttllncnt + 1
        
        
         select case xmstatus
                case "A"
                     if clsbal > 0 then
                        ttl1 = ttl1 + clsbal
                        cnt1 = cnt1 + 1
                     end if
                     idx ="自動轉帳(ALL)"
                case "B"
                     if clsbal > 0  or (clsbal=0 and lnamt >0)  then
                        ttl2 = ttl2 +clsbal
                        cnt2 = cnt2 + 1
                      end if
                     
                     idx ="破產"
                case "C"
                   if clsbal > 0  or (clsbal=0 and lnamt >0)  then
                       ttl3 = ttl3+ clsbal
                       cnt3 = cnt3 + 1
                     end if 
                     idx ="退社"
                case "D"
                    if clsbal > 0 then
                       ttl4 = ttl4 + clsbal
                       cnt4 = cnt4 + 1 
                    end if 
                     idx ="冷戶"
                case "F"
                    ttl5 = ttl5 +  clsbal
                    cnt5 = cnt5 + 1
                     idx ="特別個案"
                case "H"
                   ttl6 = ttl6 + clsbal
                   cnt6 = cnt6 + 1
                     idx ="暫停銀行"
                case  "J"
                   ttl7 = ttl7 + clsbal
                   cnt7 = cnt7 + 1
                     idx ="新戶"
                case "L"
                 if clsbal > 0 or (clsbal=0 and lnamt>0) then
                   ttl8 = ttl8 + clsbal
                   cnt8 = cnt8 + 1  
                 end if
                     idx ="呆帳"
                case "M"
                  ttl9 = ttl9 + clsbal
                  cnt9 = cnt9 + 1
                     idx ="庫房,銀行"
                case "N"     
                 ttl10 = ttl10 + clsbal
                 cnt10 = cnt10 + 1
                     idx ="正常"            
                case "P"
                if clsbal > 0 then
                   ttl11 = ttl11 + clsbal
                   cnt11 = cnt11 + 1
                end if   
                     idx ="去世"
                case "T"
                 ttl12 = ttl12 + clsbal
                 cnt12 = cnt12 + 1
                     idx ="庫房"
                case "V"
                     if clsbal> 0 or (clsbal=0 and lnamt >0 ) then
                        ttl13 = ttl13 + clsbal
                        cnt13 = cnt13 + 1 
                     end if
                     idx ="IVA"
                    
                case "0"
                 ttl14 = ttl14 + clsbal
                 cnt14 = cnt14 + 1
                     idx =" 自動轉帳(股金)"
                case "1"
                 ttl15 = ttl15 + clsbal
                 cnt15 = cnt15 + 1
                     idx ="自動轉帳(股金,利息)"
                case "2"
                 ttl16 = ttl16 + clsbal
                 cnt16 = cnt16 + 1
                     idx ="自動轉帳(股金,本金)"
                case "3"
                 ttl17 = ttl17 + clsbal
                 cnt17 = cnt17 + 1
                     idx ="自動轉帳(利息,本金)"
                case "8"
                 ttl18 = ttl18 + clsbal
                 cnt18 = cnt18 + 1 
                     idx ="終止社籍轉帳"
                case "9"
                 ttl19 = ttl19 + clsbal
                 cnt19 = cnt19 + 1
                     idx ="終止社籍正常"
         end select    
    if clsbal > 0 or(clsbal=0 and lnamt > 0 ) then
   
%>
     <tr>
          <td width="80"><%=xmemno%></td>
          <td width=180 align="left"><%=memname%></td>   
          <td width="80"><font size="3" face="標楷體"  ><%=memcname%></font></td>
          <td width="80"><%=xmemhkid%></td>
          <td width="80"><%=membday%></td>
          <td width="50" align="center"><%=sex%></td> 
          <td width="80"  align="center"><%=formatnumber(age,0)%></td>  
          <td width="130" align="right"><%=formatnumber(clsbal,2)%></td>
          <td width="130" align="right"><%=formatnumber(lnamt,2)%></td>
          <td width="150" align="center"><font size="3" face="標楷體"  ></font><%=idx%></font></td>
     </tr>

<%
 
        ttlamt = ttlamt + clsbal
        
        end if     
 
 %>

     	<tr><td colspan=10><hr></td></tr>
        <tr>
              <td></td>
              <td></td>
              <td></td>
              <td></td>
              <td></td>  
               <td></td>  
             <td></td>             
             <td width="130" align="right"><%=formatnumber(ttlamt,2)%></td>
             <td width="130" align="right"><%=formatnumber(tlnamt,2)%></td>
         </tr>
        <tr><td></td>
             <td></td>
              <td></td> 
              <td></td>
              <td></td>
             <td></td>             
               <td></td>  
             <td width="130" align="right">============</td>
              <td width="130" align="right">============</td>
         </tr>	


</table>
<BR>

<BR>


</center>
</body>
</html>

